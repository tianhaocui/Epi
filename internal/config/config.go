package config

import (
	"encoding/json"
	"os"
	"time"
)

// 添加一个辅助结构体来处理 JSON 解析
type jsonConfig struct {
	ExcelPath     string `json:"excel_path"`
	SheetName     string `json:"sheet_name"`
	HeaderRow     int    `json:"header_row"`
	BaseURL       string `json:"base_url"`
	Authorization string `json:"authorization"`
	Timeout       string `json:"timeout"` // 改为 string 类型
	Concurrent    int    `json:"concurrent"`
}

type Config struct {
	ExcelPath     string
	SheetName     string
	HeaderRow     int
	BaseURL       string
	Authorization string
	Timeout       time.Duration
	Concurrent    int
}

func Load() (*Config, error) {
	// 读取配置文件
	data, err := os.ReadFile("config.json")
	if err != nil {
		return nil, err
	}

	// 先解析到临时结构体
	var jsonCfg jsonConfig
	if err := json.Unmarshal(data, &jsonCfg); err != nil {
		return nil, err
	}

	// 解析 timeout 字符串
	timeout, err := time.ParseDuration(jsonCfg.Timeout)
	if err != nil {
		timeout = 30 * time.Second // 默认值
	}

	// 创建最终的配置对象
	cfg := &Config{
		ExcelPath:     jsonCfg.ExcelPath,
		SheetName:     jsonCfg.SheetName,
		HeaderRow:     jsonCfg.HeaderRow,
		BaseURL:       jsonCfg.BaseURL,
		Authorization: jsonCfg.Authorization,
		Timeout:       timeout,
		Concurrent:    jsonCfg.Concurrent,
	}

	// 设置默认值
	if cfg.HeaderRow == 0 {
		cfg.HeaderRow = 1
	}
	if cfg.Concurrent == 0 {
		cfg.Concurrent = 1
	}
	if cfg.SheetName == "" {
		cfg.SheetName = "Sheet1"
	}

	return cfg, nil
}
