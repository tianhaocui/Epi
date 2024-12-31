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

	"github.com/xuri/excelize/v2"

	"regression_testing/internal/config"
	"regression_testing/internal/model"
)

type Runner struct {
	config *config.Config
}

func New(cfg *config.Config) *Runner {
	return &Runner{config: cfg}
}

func (r *Runner) Run() ([]model.TestResult, error) {
	f, err := excelize.OpenFile(r.config.ExcelPath)
	if err != nil {
		return nil, fmt.Errorf("无法打开Excel文件: %v", err)
	}
	defer f.Close()

	rows, err := f.GetRows(r.config.SheetName)
	if err != nil {
		return nil, fmt.Errorf("无法读取工作表: %v", err)
	}

	testCases := rows[r.config.HeaderRow:]
	if len(testCases) == 0 {
		return nil, fmt.Errorf("没有找到测试用例")
	}

	// 读取第一行的baseURL和token作为默认值
	defaultBaseURL := ""
	defaultToken := ""
	if len(testCases[0]) >= 10 { // 确保有足够的列
		defaultBaseURL = testCases[0][8] // 第9列
		defaultToken = testCases[0][9]   // 第10列
	}

	// 更新配置
	r.config.BaseURL = defaultBaseURL
	r.config.Authorization = defaultToken

	resultChan := make(chan model.TestResult, len(testCases))
	var wg sync.WaitGroup

	// 创建工作池
	jobChan := make(chan struct {
		caseNum int
		row     []string
	}, len(testCases))

	// 启动工作协程
	for i := 0; i < r.config.Concurrent; i++ {
		go r.worker(jobChan, resultChan, &wg)
	}

	// 分发任务
	for i, row := range testCases {
		wg.Add(1)
		jobChan <- struct {
			caseNum int
			row     []string
		}{i + 1, row}
	}
	close(jobChan)

	// 等待所有测试完成
	wg.Wait()
	close(resultChan)

	// 收集结果
	var results []model.TestResult
	for result := range resultChan {
		results = append(results, result)
	}

	// 按用例编号排序结果
	r.sortResults(results)
	return results, nil
}

func (r *Runner) worker(jobs <-chan struct {
	caseNum int
	row     []string
}, results chan<- model.TestResult, wg *sync.WaitGroup) {
	for job := range jobs {
		testCase := r.parseRow(job.row)
		result := r.executeTest(job.caseNum, testCase)
		results <- result
		wg.Done()
	}
}

func (r *Runner) parseRow(row []string) model.TestCase {
	baseURL := r.config.BaseURL
	token := r.config.Authorization

	// 如果行中包含baseURL和token列，且不为空，则使用行中的值
	if len(row) >= 10 {
		if row[8] != "" {
			baseURL = row[8]
		}
		if row[9] != "" {
			token = row[9]
		}
	}

	return model.TestCase{
		CaseName:    row[0],
		Method:      row[1],
		Path:        row[2],
		PathParams:  r.parseParams(row[3]),
		QueryParams: r.parseParams(row[4]),
		Body:        row[5],
		Expected:    row[6],
		StrictMatch: row[7] == "true",
		BaseURL:     baseURL,
		Token:       token,
	}
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
	// 使用测试用例中的baseURL
	url := tc.BaseURL + tc.Path

	// 替换路径参数
	for k, v := range tc.PathParams {
		url = strings.Replace(url, "{"+k+"}", v, -1)
	}

	// 添加查询参数
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
		return model.TestResult{
			CaseNumber: caseNumber,
			Success:    false,
			Error:      fmt.Sprintf("创建请求失败: %v", err),
		}
	}

	// 设置请求头
	req.Header.Set("Content-Type", "application/json")
	if tc.Token != "" {
		req.Header.Set("Authorization", tc.Token)
	}
	for key, value := range tc.Headers {
		req.Header.Set(key, value)
	}

	// 生成 curl 命令
	curlCmd := r.toCurl(req, tc.Body)

	// 打印 curl 命令
	fmt.Printf("\n=== 执行测试用例 #%d: %s ===\n", caseNumber, tc.CaseName)
	fmt.Println(curlCmd)

	// 执行请求
	client := &http.Client{Timeout: r.config.Timeout}
	resp, err := client.Do(req)
	if err != nil {
		return model.TestResult{
			CaseNumber: caseNumber,
			Success:    false,
			Error:      fmt.Sprintf("执行请求失败: %v", err),
		}
	}
	defer resp.Body.Close()

	// 读取响应
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return model.TestResult{
			CaseNumber: caseNumber,
			Success:    false,
			Error:      fmt.Sprintf("读取响应失败: %v", err),
		}
	}

	actualResult := string(body)
	success := r.validateResponse(actualResult, tc.Expected, tc.StrictMatch)

	return model.TestResult{
		CaseNumber:     caseNumber,
		CaseName:       tc.CaseName,
		Method:         tc.Method,
		Path:           tc.Path,
		PathParams:     tc.PathParams,
		QueryParams:    tc.QueryParams,
		RequestBody:    tc.Body,
		Success:        success,
		ActualResult:   actualResult,
		ExpectedResult: tc.Expected,
		Curl:           curlCmd,
	}
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
