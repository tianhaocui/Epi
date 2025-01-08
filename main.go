package main

import (
	"fmt"
	"log"
	"time"

	"regression_testing/internal/config"
	"regression_testing/internal/reporter"
	"regression_testing/internal/runner"
)

func main() {
	cfg, err := config.Load()
	if err != nil {
		log.Fatalf("加载配置失败: %v", err)
	}

	r := runner.New(cfg, "")

	startTime := time.Now()
	results, err := r.Run()
	if err != nil {
		log.Fatalf("执行测试失败: %v", err)
	}

	duration := time.Since(startTime)
	rep := reporter.New(cfg)
	if err := rep.GenerateReport(results, duration); err != nil {
		log.Fatalf("生成报告失败: %v", err)
	}

	for _, result := range results {
		if !result.Success {
			fmt.Println("测试失败")
			return
		}
	}
	fmt.Println("测试通过")
}
