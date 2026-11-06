# OpenAI 兼容模型接入 Claude Agent SDK

## 行为与范围

Excel 与 WPS 共用后端连接目录、SDK 编排、工具权限和连接测试。配置为 `openai_compatible` 的有效地址现在可以进入表格 Agent；不再只允许 DeepSeek 特例。目录的旧 `agentAvailable` 字段表示可路由，不能替代供应商实际工具能力验证。后续同日调整为[统一文本连接测试](UNIFIED_MODEL_CONNECTION_2026-11-04.md)，不再选择用途或用工具能力禁用模型；下文 SDK 工具探针作为开发验收保留。

标准接入：供应商 Chat Completions 基础地址（通常包含 `/v1`）、真实模型 ID、API Key。本地无鉴权接口可留空 Key。接口须提供 SSE 流式响应及函数工具调用；只有聊天能力、非兼容协议或错误的模型/鉴权仍会返回明确失败。不会切换其他连接或降级到另一套 Agent。Anthropic 原生连接和已有 DeepSeek 官方地址映射保留。

## 实现

- `openai_sdk_protocol.py` 只转换协议，没有宿主、文件、凭据或执行权限。映射系统/用户/助手消息、SDK 运行时 system 消息、用户图片、thinking/reasoning_content、JSON Schema 工具、并行调用标识、工具结果与错误、结构化输出以及 SSE。
- `openai_sdk_transport.py` 在每次 SDK 查询期间创建 `127.0.0.1` 随机端口及随机令牌，仅开放 `/v1/messages`。上游 URL、模型 ID、Key 固定绑定该次所选连接，SDK 请求不能覆盖。真实 Key 只由后端发送至已配置上游；CLI 只获得短期回环令牌。没有公共代理入口、全局当前连接或外部图片服务。
- OpenAI 文本逐块转发，工具参数分片暂存。必须完整结束且所有调用名、ID、参数对象都有效，才发布工具调用；截断、未知工具、参数损坏及未完成流均不发布成功终态。
- SDK 的角色模型默认值绑定本次模型；禁用继承的 Bedrock/Vertex/Foundry 路由开关。回环连接绕过系统代理，供应商连接保留既有环境代理。禁止上游重定向，避免凭据跟随其他地址。
- 仅在供应商明确以 HTTP 400 拒绝 `max_tokens` 或 `stream_options` 时分别改用 `max_completion_tokens` 或省略 usage 选项，最多两次；不会基于模糊错误改模型或重放工具执行。
- 取消、超时或来源变化关闭 SDK 流、上游请求和本地监听端口。错误只保留分类和固定说明，供应商返回的原始正文不进入界面或日志。
- 新限额集中在 `config/capabilities.json/openaiSdkTransport`：请求 16 MiB、响应 8 MiB、单次输出 16384 tokens、每查询 128 个 HTTP 请求、上游超时 120 秒。连接测试整体时限从 60 秒调整为 120 秒，用于多轮工具往返及供应商限流重试。

SDK 工具白名单、所选表授权、来源复核、恢复准备、AnalysisPlan 执行一次、回读、确认不重写和放弃恢复均沿用原路径。意图阶段没有新增数据行上传；适配层只能接收 SDK 已有消息。token 用量随上游 usage 转换；未提供 usage 的接口无法提供精确用量，未知模型的费用也不能按 Claude 价格保证美元预算。

## 验证

| 层级 | Excel | WPS |
|---|---|---|
| OpenAI/Anthropic 协议配置说明 | 已实现，共享组件 | 已实现，共享组件 |
| 通用模型选择与请求连接 ID 一致 | 新增 App 集成用例 | 新增面板集成用例 |
| 规划进入同一 SDK，保留 hostId 与只读结构策略 | 共享入口参数化测试 | 共享入口参数化测试 |
| 工具、授权、写入、恢复行为 | 既有全量回归 | 既有全量回归 |
| 原生窗口里实际工作簿写入 | 本轮未执行 | 本轮未执行 |

`server/tests/test_openai_sdk_transport.py` 覆盖协议、分片并行工具、无效/截断响应、鉴权、凭据与模型隔离、取消、限额、参数拒绝适配，以及**真实已安装 SDK/CLI → 回环适配 → 模拟 OpenAI 服务 → MCP 工具 → 结果回传**。这项测试使用合成令牌，不是 SDK 替身。

真实供应商验证只发送合成连接探针，没有读取工作簿：DeepSeek 通过；Kimi 的 SDK/MCP 往返通过，第二轮遇到 429 后由 SDK 重试，约 64 秒完成。vec 按用户指示不再作为验收对象，未改动其配置。

本轮自动化记录：两端构建与 Python compileall 通过。后端全量首跑 2058 passed / 3 failed / 1 skipped，其中一条旧测试仍要求拒绝普通 OpenAI 连接、两条隔离器时限测试受并行负载影响；更新路由断言后，三个相关文件复跑 194 passed（包括 22 条协议/真实 SDK 链路测试）。Excel 全量 1795 passed / 1 timeout / 9 skipped，相关两个文件单独复跑 167 passed；WPS 全量 1321 passed / 1 failed / 8 skipped，失败为旧配置文案断言，更新后配置与通用模型路由 3 条复跑通过。不将分文件复跑写成最终版本全量重跑。日志为 `outputs/openai-*-tests.log`、`outputs/openai-server-full.log`、`outputs/openai-*-recheck.log`。

可显式复验：

```powershell
.\.venv\Scripts\python.exe scripts/check_openai_sdk_connections.py --label DS --label Kimi
```

该脚本会调用指定连接并消耗其 API 配额；仅输出连接名、模型 ID、结果与安全错误说明。独立临时 SDK 目录在退出时清理。没有 `--label` 时测试所有已保存 OpenAI 连接。

原生宿主验收：重启后端、刷新 Excel/WPS；用 OpenAI 协议创建或测试连接，选择该模型发起只读查询。分别核对连接 ID、工具结果、暂停/继续以及保存后重开选择。写入场景应在专用验收工作簿中验证既有授权与恢复流程，本轮不将自动化替身描述为实机通过。

协议参考：[Anthropic 流事件](https://platform.claude.com/docs/en/build-with-claude/streaming)、[OpenAI Chat Completions](https://developers.openai.com/api/reference/resources/chat/subresources/completions/methods/create)。这些格式适配不代表 SDK 官方对全部第三方模型的兼容保证。
