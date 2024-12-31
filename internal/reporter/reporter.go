package reporter

import (
	"fmt"
	"net/url"
	"path"
	"path/filepath"
	"strings"
	"time"

	"regression_testing/internal/config"
	"regression_testing/internal/model"

	"github.com/xuri/excelize/v2"
)

const (
	// Excel 相关
	defaultSheetNameFormat = "测试报告_%s"
	timeFormat             = "2006-01-02_15-04-05"
	minColumn              = 'A'
	maxColumn              = 'L'
	defaultColumnWidth     = 12

	// 样式相关
	patternType  = "pattern"
	patternValue = 1
	errorBgColor = "FF5900"

	// HTTP 相关
	contentTypeHeader = "Content-Type"
	contentTypeJSON   = "application/json"
)

// 表头定义
var excelHeaders = []string{
	"用例编号", "用例名称", "请求方法", "请求路径", "路径参数",
	"查询参数", "请求体", "期望结果", "实际结果", "测试结果",
	"错误信息", "CURL命令",
}

type Reporter struct {
	config *config.Config
}

func New(cfg *config.Config) *Reporter {
	return &Reporter{config: cfg}
}

func (r *Reporter) GenerateReport(results []model.TestResult, duration time.Duration) error {
	r.printConsoleReport(results, duration)
	return r.generateExcelReport(results, duration)
}

func (r *Reporter) generateExcelReport(results []model.TestResult, duration time.Duration) error {
	f, err := r.initExcelFile()
	if err != nil {
		return wrapError("初始化Excel文件", err)
	}
	defer f.Close()

	sheetName := fmt.Sprintf(defaultSheetNameFormat, time.Now().Format(timeFormat))
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("创建工作表失败: %v", err)
	}

	if err := r.writeHeaders(f, sheetName); err != nil {
		return err
	}

	if err := r.writeTestResults(f, sheetName, results); err != nil {
		return err
	}

	if err := r.writeSummary(f, sheetName, results, duration); err != nil {
		return err
	}

	f.SetActiveSheet(index)
	if err := f.Save(); err != nil {
		return fmt.Errorf("保存Excel报告失败: %v", err)
	}

	absPath, _ := filepath.Abs(r.config.ExcelPath)
	fmt.Printf("\nExcel测试报告已生成: %s (工作表: %s)\n", absPath, sheetName)
	return nil
}

func (r *Reporter) initExcelFile() (*excelize.File, error) {
	f, err := excelize.OpenFile(r.config.ExcelPath)
	if err != nil {
		return nil, fmt.Errorf("无法打开Excel文件: %v", err)
	}
	return f, nil
}

func (r *Reporter) writeHeaders(f *excelize.File, sheetName string) error {
	// 设置所有列的宽度
	for col := minColumn; col <= maxColumn; col++ {
		colName := string(col)
		if err := f.SetColWidth(sheetName, colName, colName, defaultColumnWidth); err != nil {
			return fmt.Errorf("设置列宽度失败: %v", err)
		}
	}

	// 写入表头
	for i, header := range excelHeaders {
		cell := fmt.Sprintf("%c1", 'A'+i)
		f.SetCellValue(sheetName, cell, header)
	}
	return nil
}

func (r *Reporter) writeTestResults(f *excelize.File, sheetName string, results []model.TestResult) error {
	for i, result := range results {
		row := i + 2
		if err := r.writeTestResult(f, sheetName, row, result); err != nil {
			return err
		}
	}
	return nil
}

func (r *Reporter) writeTestResult(f *excelize.File, sheetName string, row int, result model.TestResult) error {
	pathParamsStr, pathErr := formatParams(result.PathParams)
	queryParamsStr, queryErr := formatParams(result.QueryParams)

	// 使用 map 来简化单元格映射
	cellMap := map[string]interface{}{
		"A": result.CaseNumber,
		"B": result.CaseName,
		"C": result.Method,
		"D": result.Path,
		"E": pathParamsStr,
		"F": queryParamsStr,
		"G": result.RequestBody,
		"H": result.ExpectedResult,
		"I": result.ActualResult,
		"J": formatTestResult(result.Success),
		"K": r.formatErrorMessage(result, pathErr, queryErr),
		"L": result.Curl,
	}

	// 写入单元格值
	for col, val := range cellMap {
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", col, row), val)
	}

	return r.setRowStyle(f, sheetName, row, result.Success)
}

func (r *Reporter) formatErrorMessage(result model.TestResult, pathErr, queryErr error) string {
	var errors []string
	if result.Error != "" {
		errors = append(errors, result.Error)
	}
	if pathErr != nil {
		errors = append(errors, fmt.Sprintf("路径参数: %v", pathErr))
	}
	if queryErr != nil {
		errors = append(errors, fmt.Sprintf("查询参数: %v", queryErr))
	}
	return strings.Join(errors, "\n")
}

func (r *Reporter) setRowStyle(f *excelize.File, sheetName string, row int, success bool) error {
	// 只在测试失败时设置样式
	if !success {
		style, err := f.NewStyle(&excelize.Style{
			Fill: excelize.Fill{
				Type:    patternType,
				Color:   []string{errorBgColor},
				Pattern: patternValue,
			},
		})
		if err != nil {
			return err
		}

		// 为整行设置样式
		for col := minColumn; col <= maxColumn; col++ {
			cellName := fmt.Sprintf("%c%d", col, row)
			if err := f.SetCellStyle(sheetName, cellName, cellName, style); err != nil {
				return err
			}
		}
	}
	return nil
}

func (r *Reporter) printConsoleReport(results []model.TestResult, duration time.Duration) {
	fmt.Println("\n=== 测试报告 ===")
	fmt.Printf("总用例数: %d\n", len(results))
	fmt.Printf("总执行时间: %v\n", duration)

	successCount := 0
	for _, result := range results {
		if result.Success {
			successCount++
		}
	}

	fmt.Printf("成功用例数: %d\n", successCount)
	fmt.Printf("失败用例数: %d\n", len(results)-successCount)
	fmt.Println("\n详细结果:")

	for _, result := range results {
		fmt.Printf("\n=== 测试用例 #%d: %s ===\n", result.CaseNumber, result.CaseName)

		// 构建并打印 curl 命令
		curl := r.buildCurlCommand(result)
		fmt.Printf("CURL命令: %s\n\n", curl)

		fmt.Printf("请求方法: %s\n", result.Method)
		fmt.Printf("请求路径: %s\n", result.Path)

		if len(result.PathParams) > 0 {
			fmt.Println("路径参数:")
			for k, v := range result.PathParams {
				fmt.Printf("  %s: %s\n", k, v)
			}
		}

		if len(result.QueryParams) > 0 {
			fmt.Println("查询参数:")
			for k, v := range result.QueryParams {
				fmt.Printf("  %s: %s\n", k, v)
			}
		}

		if result.RequestBody != "" {
			fmt.Printf("请求体: %s\n", result.RequestBody)
		}

		if result.Success {
			fmt.Println("状态: 成功")
			fmt.Printf("响应结果: %s\n", result.ActualResult)
		} else {
			fmt.Println("状态: 失败")
			if result.Error != "" {
				fmt.Printf("错误信息: %s\n", result.Error)
			}
			fmt.Printf("期望结果: %s\n", result.ExpectedResult)
			fmt.Printf("实际结果: %s\n", result.ActualResult)
		}
	}
}

func formatParams(params map[string]string) (string, error) {
	if len(params) == 0 {
		return "", nil
	}
	var pairs []string
	for k, v := range params {
		// 检查参数格式
		if strings.TrimSpace(k) == "" || strings.TrimSpace(v) == "" {
			return "", fmt.Errorf("参数格式化错误: 键或值不能为空")
		}
		// 这里可以添加其他参数格式检查
		pairs = append(pairs, fmt.Sprintf("%s=%s", k, v))
	}
	return strings.Join(pairs, "\n"), nil
}

func formatTestResult(success bool) string {
	if success {
		return "成功"
	}
	return "失败"
}

// 添加新方法用于构建 curl 命令
func (r *Reporter) buildCurlCommand(result model.TestResult) string {
	// 使用 url.Parse 解析基础 URL
	baseURL, err := url.Parse(r.config.BaseURL)
	if err != nil {
		return fmt.Sprintf("无法解析URL: %v", err)
	}

	// 设置路径
	baseURL.Path = path.Join(baseURL.Path, result.Path)

	// 替换路径参数
	urlPath := baseURL.Path
	for k, v := range result.PathParams {
		urlPath = strings.Replace(urlPath, "{"+k+"}", url.PathEscape(v), -1)
	}
	baseURL.Path = urlPath

	// 添加查询参数
	if len(result.QueryParams) > 0 {
		q := baseURL.Query()
		for k, v := range result.QueryParams {
			q.Add(k, v)
		}
		baseURL.RawQuery = q.Encode()
	}

	// 构建 curl 命令
	var curlParts []string
	curlParts = append(curlParts, "curl", "-X", result.Method)

	// 添加请求头
	if r.config.Authorization != "" {
		curlParts = append(curlParts, "-H", fmt.Sprintf("'Authorization: %s'", r.config.Authorization))
	}
	curlParts = append(curlParts, "-H", "'Content-Type: application/json'")

	// 添加请求体
	if result.RequestBody != "" {
		curlParts = append(curlParts, "-d", fmt.Sprintf("'%s'", result.RequestBody))
	}

	// 添加 URL
	curlParts = append(curlParts, fmt.Sprintf("'%s'", baseURL.String()))

	return strings.Join(curlParts, " ")
}

func (r *Reporter) writeSummary(f *excelize.File, sheetName string, results []model.TestResult, duration time.Duration) error {
	startRow := len(results) + 3

	// 定义汇总信息，使用纵向布局
	summaryItems := []struct {
		label   string
		value   string
		isError bool
	}{
		{"测试汇总", "", false},
		{"总用例数", fmt.Sprintf("%d", len(results)), false},
		{"总执行时间", duration.String(), false},
		{"成功用例数", fmt.Sprintf("%d", r.countSuccessfulTests(results)), false},
		{"失败用例数", fmt.Sprintf("%d", len(results)-r.countSuccessfulTests(results)), true}, // 总是标记为错误
	}

	// 创建红色样式
	errorStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    patternType,
			Color:   []string{errorBgColor},
			Pattern: patternValue,
		},
	})
	if err != nil {
		return err
	}

	// 写入汇总信息
	for i, item := range summaryItems {
		row := startRow + i
		// 写入标签和值
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), item.label)
		if item.value != "" {
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), item.value)
		}

		// 如果是失败用例行，设置红色背景
		if item.isError {
			f.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("B%d", row), errorStyle)
		}
	}

	return nil
}

// 提取计算成功测试数量的逻辑
func (r *Reporter) countSuccessfulTests(results []model.TestResult) int {
	successCount := 0
	for _, result := range results {
		if result.Success {
			successCount++
		}
	}
	return successCount
}

// 添加错误包装函数
func wrapError(op string, err error) error {
	return fmt.Errorf("%s: %v", op, err)
}
