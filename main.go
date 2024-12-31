package main

import (
	"time"

	"regression_testing/internal/config"
	"regression_testing/internal/reporter"
	"regression_testing/internal/runner"
)

func main() {
	cfg := config.Parse()

	testRunner := runner.New(cfg)
	startTime := time.Now()
	results, err := testRunner.Run()
	if err != nil {
		panic(err)
	}

	duration := time.Since(startTime)

	rpt := reporter.New(cfg)
	if err := rpt.GenerateReport(results, duration); err != nil {
		panic(err)
	}
}
