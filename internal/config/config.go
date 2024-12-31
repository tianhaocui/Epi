package config

import (
	"flag"
	"time"
)

type Config struct {
	ExcelPath     string
	BaseURL       string
	Authorization string
	Timeout       time.Duration
	SheetName     string
	HeaderRow     int
	Concurrent    int
}

func Parse() *Config {
	config := &Config{}

	flag.StringVar(&config.ExcelPath, "excel", "test_cases.xlsx", "Excel文件路径")
	flag.StringVar(&config.BaseURL, "base-url", "", "API基础URL")
	flag.StringVar(&config.Authorization, "auth", "", "Authorization头信息")
	flag.DurationVar(&config.Timeout, "timeout", 30*time.Second, "请求超时时间")
	flag.StringVar(&config.SheetName, "sheet", "Sheet1", "Excel工作表名称")
	flag.IntVar(&config.HeaderRow, "header-row", 1, "标题行号（从1开始）")
	flag.IntVar(&config.Concurrent, "concurrent", 5, "并发执行的goroutine数量")

	flag.Parse()

	return config
}
