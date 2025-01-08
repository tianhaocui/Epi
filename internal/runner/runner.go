package runner

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"reflect"
	"regexp"
	"strings"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"

	"regression_testing/internal/config"
	"regression_testing/internal/model"
)

type Runner struct {
	config        *config.Config
	firstToken    string
	globalHeaders map[string]string
}

func New(cfg *config.Config, _ string) *Runner {
	return &Runner{
		config:        cfg,
		globalHeaders: make(map[string]string),
	}
}

func (r *Runner) Run() ([]model.TestResult, error) {
	startTime := time.Now() // 添加开始时间记录

	f, err := excelize.OpenFile(r.config.ExcelPath)
	if err != nil {
		return nil, fmt.Errorf("无法打开Excel文件: %v", err)
	}
	defer f.Close()

	rows, err := f.GetRows(r.config.SheetName)
	if err != nil {
		return nil, fmt.Errorf("无法读取工作表: %v", err)
	}

	// 跳过第一行
	if len(rows) <= 1 {
		return nil, fmt.Errorf("没有找到测试用例")
	}
	testCases := rows[1:]

	// 初始化全局配置（从第一个有效测试用例获取）
	r.initGlobalConfig(testCases)

	resultChan := make(chan model.TestResult, len(testCases))
	var wg sync.WaitGroup

	jobChan := make(chan struct {
		caseNum int
		row     []string
	}, len(testCases))

	// 启动工作协程
	for i := 0; i < r.config.Concurrent; i++ {
		go r.worker(jobChan, resultChan, &wg)
	}

	// 分发任务
	totalTests := 0
	for i, row := range testCases {
		if _, ok := r.parseRow(row); ok {
			wg.Add(1)
			totalTests++
			rowNum := i + 2
			jobChan <- struct {
				caseNum int
				row     []string
			}{rowNum, row}
		}
	}
	close(jobChan)

	wg.Wait()
	close(resultChan)

	var results []model.TestResult
	failedTests := 0
	for result := range resultChan {
		results = append(results, result)
		if !result.Success {
			failedTests++
		}
	}

	r.sortResults(results)

	// 输出测试汇总
	fmt.Printf("\n测试汇总\n")
	fmt.Printf("总执行时间: %.6fms\n", float64(time.Since(startTime).Microseconds())/1000)
	fmt.Printf("总用例数: %d\n", totalTests)
	if failedTests > 0 {
		fmt.Printf("\033[31m失败用例数: %d\033[0m\n", failedTests)
	} else {
		fmt.Printf("失败用例数: %d\n", failedTests)
	}

	return results, nil
}

func (r *Runner) worker(jobs <-chan struct {
	caseNum int
	row     []string
}, results chan<- model.TestResult, wg *sync.WaitGroup) {
	for job := range jobs {
		testCase, _ := r.parseRow(job.row)
		result := r.executeTest(job.caseNum, testCase)
		results <- result
		wg.Done()
	}
}

func (r *Runner) parseRow(row []string) (model.TestCase, bool) {
	// 1. 基础检查：空行
	if len(row) == 0 {
		return model.TestCase{}, false
	}

	// 2. 检查是否有第二列且是 HTTP 方法
	if len(row) < 2 || !isHTTPMethod(row[1]) {
		return model.TestCase{}, false
	}

	// 3. 检查是否有足够的列数来构建测试用例
	if len(row) < 8 {
		return model.TestCase{}, false
	}

	// 4. 设置请求头
	headers := make(map[string]string)
	// 先添加全局请求头
	for k, v := range r.globalHeaders {
		headers[k] = v
	}
	// 如果当前行有请求头，解析并覆盖全局请求头
	if len(row) >= 7 && row[6] != "" {
		var currentHeaders map[string]string
		if err := json.Unmarshal([]byte(row[6]), &currentHeaders); err == nil {
			for k, v := range currentHeaders {
				headers[k] = v
			}
		}
	}

	// 5. 设置基础 URL
	baseURL := ""
	if len(row) >= 10 && row[9] != "" {
		baseURL = row[9] // 使用当前行的 base-url
	} else {
		baseURL = r.findFirstBaseURL(row[0])
	}
	if baseURL == "" {
		baseURL = r.config.BaseURL // 如果找不到，使用配置中的默认值
	}

	// 6. 设置认证信息
	token := ""
	if len(row) >= 11 && row[10] != "" {
		token = row[10] // 使用当前行的 token
	} else {
		token = r.firstToken // 使用第一个用例的 token
	}
	if token == "" {
		token = r.config.Authorization // 如果都没有，使用配置中的默认值
	}

	// 7. 构建并返回测试用例
	return model.TestCase{
		CaseName:    row[0],
		Method:      row[1],
		Path:        row[2],
		PathParams:  r.parseParams(row[3]),
		QueryParams: r.parseParams(row[4]),
		Body:        row[5],
		Headers:     headers,
		Expected:    row[7],
		StrictMatch: row[8] == "true",
		BaseURL:     baseURL,
		Token:       token,
	}, true
}

func (r *Runner) parseParams(paramStr string) map[string]string {
	params := make(map[string]string)
	if paramStr == "" {
		return params
	}

	pairs := strings.Split(paramStr, "&")
	for _, pair := range pairs {
		kv := strings.Split(pair, "=")
		if len(kv) == 2 {
			params[kv[0]] = kv[1]
		}
	}
	return params
}

func (r *Runner) sortResults(results []model.TestResult) {
	for i := 0; i < len(results)-1; i++ {
		for j := i + 1; j < len(results); j++ {
			if results[i].CaseNumber > results[j].CaseNumber {
				results[i], results[j] = results[j], results[i]
			}
		}
	}
}

// 其他私有方法
func (r *Runner) executeTest(caseNumber int, tc model.TestCase) model.TestResult {
	startTime := time.Now() // 记录开始时间

	// 构建基本结果
	result := model.TestResult{
		CaseNumber:     caseNumber,
		CaseName:       tc.CaseName,
		Method:         tc.Method,
		Path:           tc.Path,
		PathParams:     tc.PathParams,
		QueryParams:    tc.QueryParams,
		RequestBody:    tc.Body,
		ExpectedResult: tc.Expected,
	}

	// 构建 URL
	url := tc.BaseURL + tc.Path
	for k, v := range tc.PathParams {
		url = strings.Replace(url, "{"+k+"}", v, -1)
	}
	if len(tc.QueryParams) > 0 {
		params := make([]string, 0)
		for k, v := range tc.QueryParams {
			params = append(params, k+"="+v)
		}
		url += "?" + strings.Join(params, "&")
	}

	// 创建请求
	req, err := http.NewRequest(tc.Method, url, bytes.NewBufferString(tc.Body))
	if err != nil {
		result.Success = false
		result.Error = fmt.Sprintf("创建请求失败: %v", err)
		result.Curl = r.toCurl(req, tc.Body) // 使用 toCurl 替代 buildCurlCommand
		return result
	}

	// 设置请求头
	req.Header.Set("Content-Type", "application/json")
	if tc.Token != "" {
		req.Header.Set("Authorization", tc.Token)
	}
	// 添加自定义请求头
	for k, v := range tc.Headers {
		req.Header.Set(k, v)
	}

	// 生成 curl 命令
	result.Curl = r.toCurl(req, tc.Body)

	// 执行请求
	client := &http.Client{Timeout: r.config.Timeout}
	resp, err := client.Do(req)
	if err != nil {
		result.Success = false
		result.Error = fmt.Sprintf("执行请求失败: %v", err)
		return result
	}
	defer resp.Body.Close()

	// 读取响应
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		result.Success = false
		result.Error = fmt.Sprintf("读取响应失败: %v", err)
		return result
	}

	result.ActualResult = string(body)
	result.Success = r.validateResponse(result.ActualResult, tc.Expected, tc.StrictMatch)

	// 记录执行时间（毫秒）
	result.ExecutionTime = float64(time.Since(startTime).Microseconds()) / 1000

	return result
}

func (r *Runner) validateResponse(actual, expected string, strictMatch bool) bool {
	var actualMap, expectedMap map[string]interface{}

	// 解析实际结果
	if err := json.Unmarshal([]byte(actual), &actualMap); err != nil {
		return false
	}

	// 解析期望结果
	if err := json.Unmarshal([]byte(expected), &expectedMap); err != nil {
		fmt.Printf("非有效 JSON: %s\n", expected) // 输出非有效 JSON 的信息
		return false
	}

	// 检查 code 字段
	if actualMap["code"] == nil || expectedMap["code"] == nil {
		return false
	}
	if actualMap["code"] != expectedMap["code"] {
		return false
	}

	if strictMatch {
		// 使用 validateMap 验证所有字段
		return r.validateMap(actualMap, expectedMap)
	} else {
		// 非严格匹配时，逐个检查其他字段
		for key, expectedValue := range expectedMap {
			if key == "code" {
				continue // 跳过 code 字段
			}

			actualValue, exists := actualMap[key]
			if !exists {
				return false // 如果实际结果中没有该字段，返回 false
			}

			// 判断是否是正则表达式
			if expectedStr, ok := expectedValue.(string); ok && (strings.HasPrefix(expectedStr, "^") || strings.HasSuffix(expectedStr, "$")) {
				// 使用正则匹配
				if matched, _ := regexp.MatchString(expectedStr, actualValue.(string)); !matched {
					return false
				}
			} else {
				// 直接比较值
				if actualValue != expectedValue {
					return false
				}
			}
		}
	}

	return true
}

// validateValue 验证单个值，支持正则表达式
func (r *Runner) validateValue(actual, expected interface{}) bool {
	// 处理嵌套的map
	if expectedMap, ok := expected.(map[string]interface{}); ok {
		actualMap, ok := actual.(map[string]interface{})
		if !ok {
			return false
		}
		return r.validateMap(actualMap, expectedMap)
	}

	// 处理数组
	if expectedSlice, ok := expected.([]interface{}); ok {
		actualSlice, ok := actual.([]interface{})
		if !ok {
			return false
		}
		return r.validateSlice(actualSlice, expectedSlice)
	}

	// 处理字符串（支持正则表达式）
	if expectedStr, ok := expected.(string); ok {
		actualStr, ok := actual.(string)
		if !ok {
			return false
		}

		// 检查是否是正则表达式（以 ^ 开头或 $ 结尾）
		if strings.HasPrefix(expectedStr, "^") || strings.HasSuffix(expectedStr, "$") {
			matched, err := regexp.MatchString(expectedStr, actualStr)
			if err != nil {
				return false
			}
			return matched
		}

		// 如果不是正则表达式，直接比较值
		return actualStr == expectedStr
	}

	// 其他类型直接比较
	return reflect.DeepEqual(actual, expected)
}

// validateSlice 验证数组
func (r *Runner) validateSlice(actual, expected []interface{}) bool {
	if len(actual) < len(expected) {
		return false
	}

	for i, expectedValue := range expected {
		if !r.validateValue(actual[i], expectedValue) {
			return false
		}
	}
	return true
}

// toCurl 将请求转换为 curl 命令
func (r *Runner) toCurl(req *http.Request, body string) string {
	curl := fmt.Sprintf("curl -X %s", req.Method)

	// 添加请求头
	for key, values := range req.Header {
		curl += fmt.Sprintf(" -H '%s: %s'", key, values[0])
	}

	// 添加请求体
	if body != "" {
		curl += fmt.Sprintf(" -d '%s'", body)
	}

	// 添加URL
	curl += fmt.Sprintf(" '%s'", req.URL.String())

	return curl
}

// validateMap 递归验证map中的所有字段
func (r *Runner) validateMap(actual, expected map[string]interface{}) bool {
	for key, expectedValue := range expected {
		actualValue, exists := actual[key]
		if !exists {
			return false // 如果实际结果中没有该字段，返回 false
		}

		if !r.validateValue(actualValue, expectedValue) {
			return false
		}
	}
	return true
}

// 添加辅助方法来判断是否是HTTP方法
func isHTTPMethod(method string) bool {
	methods := []string{"GET", "POST", "PUT", "DELETE", "PATCH", "HEAD", "OPTIONS"}
	method = strings.ToUpper(strings.TrimSpace(method))
	for _, m := range methods {
		if method == m {
			return true
		}
	}
	return false
}

// 添加 findFirstBaseURL 方法
func (r *Runner) findFirstBaseURL(currentCaseName string) string {
	f, err := excelize.OpenFile(r.config.ExcelPath)
	if err != nil {
		return ""
	}
	defer f.Close()

	rows, err := f.GetRows(r.config.SheetName)
	if err != nil {
		return ""
	}

	// 跳过第一行
	if len(rows) <= 1 {
		return ""
	}
	rows = rows[1:]

	// 找到当前用例的位置
	currentIndex := -1
	for i, row := range rows {
		if len(row) > 0 && row[0] == currentCaseName {
			currentIndex = i
			break
		}
	}

	if currentIndex == -1 {
		return ""
	}

	// 从当前用例向上查找第一个有 base-url 的用例
	for i := currentIndex; i >= 0; i-- {
		row := rows[i]
		if len(row) >= 10 && row[9] != "" {
			return row[9]
		}
	}

	return ""
}

// 新增：初始化全局配置
func (r *Runner) initGlobalConfig(rows [][]string) {
	for _, row := range rows {
		if len(row) < 2 || !isHTTPMethod(row[1]) {
			continue
		}
		// 获取第一个有效测试用例的 token
		if len(row) >= 11 {
			r.firstToken = row[10]
		}
		// 获取第一个有效测试用例的全局 headers（从最后一列）
		if len(row) >= 12 && row[11] != "" {
			var headers map[string]string
			if err := json.Unmarshal([]byte(row[11]), &headers); err == nil {
				r.globalHeaders = headers
			}
		}
		break
	}
}
