## EPI
### excel驱动回归测试工具
* 查看: [cases.xlsx](cases.xlsx)
* 配置: [config.json](config.json)
* 会读取并执行所有case进行调用
* 路径参数: 替换路径中的占位符 ，使用json 
* 查询参数: 会拼接在路径之后;eg: pageNo=1&pageSize=10
* body: request body 使用json
* 完全匹配:
  * FALSE: 代表接口调用结果只需要code与期望结果code一致
  * TRUE: 代表接口调用结果必须和期望结果完全一致
* token: 无特殊标明会使用读到的第一个，如果本行有值 会覆盖
* base-url: 无特殊标明会使用读到的第一个，如果本行有值 会覆盖
* GlobalHeaders: 使用json配置，如果和Headers冲突，会覆盖，只需要配置第一个GlobalHeaders即可
* 方法: 请求方式，GET、POST、PUT等等
* 测试报告:
  * 会在当前文件后面追加sheet方式输出
  * 错误用例会标红
  * 超过配置超时时间的用例会标黄
  * 用例编号=用例的行号



