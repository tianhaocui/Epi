package reporter

import (
	"fmt"
	"path/filepath"
	"strings"
	"time"

	"regression_testing/internal/config"
	"regression_testing/internal/model"

	"github.com/xuri/excelize/v2"
)

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

func (r *Reporter) generateExcelReport(results []model.TestResult, duration time.Duration) error {
	absPath, err := filepath.Abs(r.config.ExcelPath)
	if err != nil {
		return fmt.Errorf("获取文件路径失败: %v", err)
	}

	f, err := excelize.OpenFile(r.config.ExcelPath)
	if err != nil {
		return fmt.Errorf("无法打开Excel文件: %v", err)
	}
	defer f.Close()

	// 创建新的工作表，使用当前时间作为表名
	sheetName := fmt.Sprintf("测试报告_%s", time.Now().Format("2006-01-02_15-04-05"))
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("创建工作表失败: %v", err)
	}

	// 写入表头
	headers := []string{"用例编号", "用例名称", "请求方法", "请求路径", "路径参数", "查询参数",
		"请求体", "期望结果", "实际结果", "测试结果", "错误信息", "CURL命令"}
	for i, header := range headers {
		cell := fmt.Sprintf("%c1", 'A'+i)
		f.SetCellValue(sheetName, cell, header)
	}

	// 写入测试结果
	for i, result := range results {
		row := i + 2
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), result.CaseNumber)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), result.CaseName)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), result.Method)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), result.Path)
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", row), formatParams(result.PathParams))
		f.SetCellValue(sheetName, fmt.Sprintf("F%d", row), formatParams(result.QueryParams))
		f.SetCellValue(sheetName, fmt.Sprintf("G%d", row), result.RequestBody)
		f.SetCellValue(sheetName, fmt.Sprintf("H%d", row), result.ExpectedResult)
		f.SetCellValue(sheetName, fmt.Sprintf("I%d", row), result.ActualResult)
		f.SetCellValue(sheetName, fmt.Sprintf("J%d", row), formatTestResult(result.Success))
		f.SetCellValue(sheetName, fmt.Sprintf("K%d", row), result.Error)
		f.SetCellValue(sheetName, fmt.Sprintf("L%d", row), result.Curl)
	}

	// 写入汇总信息
	summaryRow := len(results) + 3
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", summaryRow), "测试汇总")
	f.SetCellValue(sheetName, fmt.Sprintf("B%d", summaryRow), fmt.Sprintf("总用例数: %d", len(results)))
	f.SetCellValue(sheetName, fmt.Sprintf("C%d", summaryRow), fmt.Sprintf("总执行时间: %v", duration))

	successCount := 0
	for _, result := range results {
		if result.Success {
			successCount++
		}
	}
	f.SetCellValue(sheetName, fmt.Sprintf("D%d", summaryRow), fmt.Sprintf("成功用例数: %d", successCount))
	f.SetCellValue(sheetName, fmt.Sprintf("E%d", summaryRow), fmt.Sprintf("失败用例数: %d", len(results)-successCount))

	// 设置为活动工作表
	f.SetActiveSheet(index)

	// 保存文件
	if err := f.Save(); err != nil {
		return fmt.Errorf("保存Excel报告失败: %v", err)
	}

	fmt.Printf("\nExcel测试报告已生成: %s (工作表: %s)\n", absPath, sheetName)
	return nil
}

// 添加辅助函数
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

func formatTestResult(success bool) string {
	if success {
		return "成功"
	}
	return "失败"
}

// 添加新方法用于构建 curl 命令
func (r *Reporter) buildCurlCommand(result model.TestResult) string {
	url := r.config.BaseURL + result.Path

	// 替换路径参数
	for k, v := range result.PathParams {
		url = strings.Replace(url, "{"+k+"}", v, -1)
	}

	// 添加查询参数
	if len(result.QueryParams) > 0 {
		params := make([]string, 0)
		for k, v := range result.QueryParams {
			params = append(params, k+"="+v)
		}
		url += "?" + strings.Join(params, "&")
	}

	curl := fmt.Sprintf("curl -X %s", result.Method)

	// 添加认证头
	if r.config.Authorization != "" {
		curl += fmt.Sprintf(" -H 'Authorization: %s'", r.config.Authorization)
	}

	// 添加 Content-Type
	curl += " -H 'Content-Type: application/json'"

	// 添加请求体
	if result.RequestBody != "" {
		curl += fmt.Sprintf(" -d '%s'", result.RequestBody)
	}

	// 添加URL
	curl += fmt.Sprintf(" '%s'", url)

	return curl
}

// 其他私有方法...
