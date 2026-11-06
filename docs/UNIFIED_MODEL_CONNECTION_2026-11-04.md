# 统一模型连接与双端验收

## 产品行为

- Excel / WPS 均移除“测试用途”和“独立公式 / Agent”选项。只选择 OpenAI 兼容或 Anthropic 兼容协议，填写连接资料，点击“测试连接”。
- 测试收到有效文本回复即成功，可保存并选择。HTTP 200 但返回错误对象、空内容、只有思考内容或截断结果不伪报成功。默认整体时限 120 秒、输出预算 2048 tokens，位于 `capabilities.json/modelConnectionTest`。
- 测试草稿不保存、不初始化 SDK、不调用 MCP、不读取工作簿。修改地址、协议、模型或 Key 后旧测试结果失效，迟到的测试也不会验证另一份草稿。
- 旧请求的 `purpose=agent/formula` 仅作为兼容字段接受，执行同一文本测试；响应 `testedFor=connection`。前端接受有效成功结果，不再检查“所选用途”。协议同步于 Python / TypeScript。
- 模型目录不按 SDK 目录、工具探针或旧 `agentAvailable` 禁用。前端统一按连接 `available` 选择；后台旧能力字段只表示可路由。任务始终带所选连接 ID 进入同一 SDK，不因工具失败换供应商，也不降级到旧 Agent。
- 功能所需工具、图片、运行环境在使用时判断，实际失败按原错误流程报告。连接测试不是工具能力测试，界面不要求用户区分两者。

## 共享实现与边界

共享文本客户端保留历史 `OpenAICompatibleClient` 名称，按保存的协议发送 OpenAI Chat Completions 或 Anthropic Messages；后者提取顶层 system、映射文本/图片/工具结果，将文本、结束状态及 token 用量归一化供既有公式解析使用。Anthropic 使用 `x-api-key` 和版本头；不会把 Anthropic 凭据发到 Chat Completions。

基础地址支持 Anthropic 根地址、`/v1`、已带 `/messages` 的完整端点；OpenAI 支持基础地址或完整 `/chat/completions`。保留供应商自定义路径。不跟随重定向，回环服务绕过系统代理。仅对 OpenAI 明确拒绝 `max_tokens` 的 HTTP 400 适配为 `max_completion_tokens`，不根据模糊错误换模型。

`/function` 两协议共用既有编译/公式解析、安全检查、宿主写入流程；可选的公式默认模型偏好保留，但可选择同一完整连接目录，不是连接类型。内置 `=EB()` 是本地预制规则，本次没有改成远程模型函数。

授权工作表、意图阶段不发送原始数据行、AnalysisPlan 白名单、来源复核、恢复准备、执行一次、回读审阅均保持原有边界。

协议依据：[Anthropic Messages API](https://platform.claude.com/docs/en/api/messages)。

## 验收矩阵

| 项目 | Excel | WPS |
|---|---|---|
| 统一配置/测试入口 | 已实现，共享组件与管理 hook | 已实现，复用同一组件与 hook |
| 旧 Agent 能力为 false 仍能选用原连接 | App 集成测试通过，两个协议目录 ID | 面板集成测试通过，两个协议目录 ID |
| OpenAI / Anthropic 进入 SDK | 双协议 × hostId 参数化测试 | 同一测试覆盖 wps-et |
| 两协议公式生成与响应处理 | 共享后端 HTTP 替身验收通过 | 同一后端入口，未新增 WPS 私有协议 |
| 构建 / 类型检查 | 通过 | 通过 |
| 原生宿主窗格交互与工作簿操作 | 实机待验，本轮不操作工作簿 | 实机待验，本轮不操作工作簿 |

自动化证据：后端首轮全量 2083 passed / 1 failed / 1 skipped；失败是旧“Anthropic 不允许进入文本客户端”断言，替换为协议鉴权断言后，统一连接、SDK 适配与双端路由相关四个文件 145 passed。Excel 全量 1796 passed / 1 timeout / 9 skipped，超时为已有审阅恢复用例；该文件及连接 hook、选择规则、共享配置组件单 worker 复跑 151 passed。WPS 全量 1322 passed / 1 failed / 8 skipped，失败是旧配置文案，更新后管理界面及双协议路由 3 passed。未将定向复跑记成最终版本全量全绿。

最终文本参数兼容补充后，统一连接、文本客户端、连接配置三个文件 57 passed；Python compileall 与 diff 检查通过，日志 `outputs/unified-text-final.log`。

真实上游使用新统一测试函数：DeepSeek `deepseek-v4-flash` 的 OpenAI / Anthropic 两种协议、Kimi `kimi-k3` 的 OpenAI 协议均通过。只发送合成问候；vec 按用户说明排除，没有改任何已保存连接。日志：`outputs/unified-provider-probes.log`。明确复验命令：

```powershell
.\.venv\Scripts\python.exe scripts/check_model_connections.py --label DS --label Kimi --also-deepseek-anthropic
```

脚本只联系显式指定的连接，消耗供应商 API 配额；不会打印 Key、写配置、读工作簿或账本。DeepSeek 的额外 Anthropic 请求只复用项目已有的官方端点映射。

## 原生验收步骤

重启原后端启动脚本并刷新两端加载项。在 Excel、WPS 各检查一次：打开添加连接只有协议选择而无用途选择；分别测试两种协议、保存并选择；模型 ID 与请求一致；普通对话和 `/function` 成功。工具操作失败应显示具体原因且模型继续可选；未经授权不得写入。测试进行中修改草稿，旧请求完成后不得解锁新草稿的保存。

本轮没有重启正在运行的后端或操作真实工作簿，未执行 Git 提交。
