package reporter

import (
	"fmt"
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
	patternType    = "pattern"
	patternValue   = 1
	errorBgColor   = "FF5900"
	warningBgColor = "FFEB9C"

	// 时间阈值
	slowTestThreshold = 300 // 300毫秒
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
	// 打开原有的 Excel 文件
	f, err := excelize.OpenFile(r.config.ExcelPath)
	if err != nil {
		return fmt.Errorf("打开Excel文件失败: %v", err)
	}
	defer f.Close()

	// 创建新的工作表
	sheetName := fmt.Sprintf(defaultSheetNameFormat, time.Now().Format(timeFormat))
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("创建工作表失败: %v", err)
	}
	f.SetActiveSheet(index)

	// 设置列宽
	for col := minColumn; col <= maxColumn; col++ {
		colName := string(col)
		f.SetColWidth(sheetName, colName, colName, defaultColumnWidth)
	}

	// 写入表头
	for i, header := range excelHeaders {
		cell := fmt.Sprintf("%c1", minColumn+i)
		f.SetCellValue(sheetName, cell, header)
	}

	// 写入测试结果
	for i, result := range results {
		row := i + 2
		r.writeTestResult(f, sheetName, row, result)
	}

	// 写入汇总信息
	summaryRow := len(results) + 3
	r.writeSummary(f, sheetName, summaryRow, results, duration)

	// 保存文件
	if err := f.Save(); err != nil {
		return fmt.Errorf("保存报告失败: %v", err)
	}

	fmt.Printf("测试报告已保存到工作表: %s\n", sheetName)
	return nil
}

func (r *Reporter) writeTestResult(f *excelize.File, sheet string, row int, result model.TestResult) {
	// 设置错误样式（红色背景）
	errorStyle, _ := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    patternType,
			Pattern: patternValue,
			Color:   []string{errorBgColor},
		},
	})

	// 设置警告样式（黄色背景）
	warningStyle, _ := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    patternType,
			Pattern: patternValue,
			Color:   []string{warningBgColor},
		},
	})

	// 写入测试结果
	cells := []interface{}{
		result.CaseNumber,
		result.CaseName,
		result.Method,
		result.Path,
		formatParams(result.PathParams),
		formatParams(result.QueryParams),
		result.RequestBody,
		result.ExpectedResult,
		result.ActualResult,
		result.Success,
		result.Error,
		result.Curl,
	}

	for i, cell := range cells {
		cellName := fmt.Sprintf("%c%d", minColumn+i, row)
		f.SetCellValue(sheet, cellName, cell)

		// 如果测试失败，设置红色背景
		if !result.Success {
			f.SetCellStyle(sheet, cellName, cellName, errorStyle)
		} else if result.ExecutionTime > slowTestThreshold { // 如果执行时间超过阈值，设置黄色背景
			f.SetCellStyle(sheet, cellName, cellName, warningStyle)
		}
	}
}

func (r *Reporter) writeSummary(f *excelize.File, sheet string, startRow int, results []model.TestResult, duration time.Duration) {
	// 计算统计信息
	totalTests := len(results)
	failedTests := 0
	for _, result := range results {
		if !result.Success {
			failedTests++
		}
	}

	// 写入汇总信息
	f.SetCellValue(sheet, fmt.Sprintf("A%d", startRow), "测试汇总")
	f.SetCellValue(sheet, fmt.Sprintf("A%d", startRow+1), fmt.Sprintf("总执行时间: %.6fms", float64(duration.Microseconds())/1000))
	f.SetCellValue(sheet, fmt.Sprintf("A%d", startRow+2), fmt.Sprintf("总用例数: %d", totalTests))
	f.SetCellValue(sheet, fmt.Sprintf("A%d", startRow+3), fmt.Sprintf("失败用例数: %d", failedTests))
}

func (r *Reporter) printConsoleReport(results []model.TestResult, duration time.Duration) {
	// 计算统计信息
	totalTests := len(results)
	failedTests := 0
	for _, result := range results {
		if !result.Success {
			failedTests++
		}
	}

	// 输出汇总信息
	fmt.Printf("\n测试汇总\n")
	fmt.Printf("总执行时间: %.6fms\n", float64(duration.Microseconds())/1000)
	fmt.Printf("总用例数: %d\n", totalTests)
	if failedTests > 0 {
		fmt.Printf("\033[31m失败用例数: %d\033[0m\n", failedTests)
	} else {
		fmt.Printf("失败用例数: %d\n", failedTests)
	}
}

func formatParams(params map[string]string) string {
	if len(params) == 0 {
		return ""
	}
	var pairs []string
	for k, v := range params {
		pairs = append(pairs, fmt.Sprintf("%s=%s", k, v))
	}
	return strings.Join(pairs, "\n")
}
