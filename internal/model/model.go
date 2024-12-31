package model

type TestCase struct {
	CaseName    string            // 测试用例名称
	Method      string            // HTTP方法
	Path        string            // 请求路径
	PathParams  map[string]string // 路径参数
	QueryParams map[string]string // 查询参数
	Body        string            // 请求体
	Expected    string            // 期望结果
	StrictMatch bool              // 是否完全匹配
	BaseURL     string            // 基础URL（可选）
	Token       string            // 认证令牌（可选）
}

type TestResult struct {
	CaseNumber     int
	CaseName       string
	Method         string
	Path           string
	PathParams     map[string]string
	QueryParams    map[string]string
	RequestBody    string
	Success        bool
	ActualResult   string
	ExpectedResult string
	Error          string
	Curl           string
}
