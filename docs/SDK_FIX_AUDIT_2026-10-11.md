# SDK 修复与 Claude 变更复核

工作开始于 2026-10-11，最终复跑于 2026-10-12（Asia/Singapore）。保留既有未提交修改，未提交、推送或发布；未重启现有 8765 服务，未操作真实工作簿。

## 已修复的产品问题

1. **写预览任务提前以回答结束。** 意图模型在 `proceed` / `tool_request` 阶段输出 `planningGoal`，Excel/WPS 将其传入 `PlanRequest`；服务端按同一 turnId 保存并优先采用意图目标。已绑定写预览时，普通回答、本地结果卡和 JSON 回答退路均不能降级目标。两后端共用判据，SDK 错误回执不锁终局，允许同一次 query 继续提交计划。先查后写同样保留目标。
2. **空动作能力产生非法 schema。** 交集为空时省略 `submit_plan`，不发送空 anyOf，也不回落全量动作。有交集时只保留协商动作；旧无 hostId 入口保持原兼容语义。三项反向变异分别抓住省略被拆、收窄被架空和兼容入口被改，源码 SHA256 还原一致。
3. **Python 结果进度事件未贯通。** 补齐生产 SSE 组装、双端共用事件解析、普通后端回调和 WPS 面板消费。Python 模型端回执仍经白名单投影；事件采用严格 `LocalResultCard`，前端只重建引用和结构，不附带矩阵。值经 `POST /api/sdk/tool-bridge/result` 按该轮任务/宿主/工作簿/来源指纹取回。
4. **Excel 保存日志清空结果卡。** 无新结果时保留原卡片；新消息分支也能附结果；多次分析使用最新结果。绑定必须匹配活动的 operationId。WPS 按任务与引用去重，取消或不同任务事件不挂卡。结果卡不会触发计划执行，也不替代写预览。

旧直接调用者如果没有事前 planningGoal，仍保留兼容行为；本轮没有宣称对这些未绑定目标的请求作了强制分类。目标判断仍归模型，没有硬编码业务字段、用户数据或提示词关键词。

## Claude 的 CHANGES_2026-10-11.md 复核

原始改动确实存在，但原文“完成双端自动渲染”的结论超出证据：

- `main.py` 组装 SSE 没有 pythonResult，前端共用解析器也未保留该字段。
- 普通后端未传 on_event，WPS 未接收/展示新事件。
- 原 `persistActivityMessage` 会以 undefined 覆盖原有结果卡；目标消息不存在时也丢失新结果。
- `test_e2e_python_sse.py` 调用协调器并捕获函数回调，没有 HTTP、SSE 解析或宿主面板。保留该文件作为原始证据，但不能称客户端端到端。
- 新 pytest 用例依赖未安装的 pytest-asyncio；现改为仓库已有的 asyncio.run 写法。
- 仓库默认开关被改成 true，与项目默认关闭约定和原有测试冲突。已恢复 false；真实模型诊断和手工验收使用独立配置启用，仍必须通过验证过的 Docker 执行器。

补充测试覆盖真实 HTTP/SSE 序列化、普通后端事件回调、拒绝事件矩阵字段、共用解析器白名单、旧结果保留，以及 Excel/WPS 面板仅从进度事件按引用获取结果。面板测试的网络与宿主仍是替身。

## 真实 SDK 并发证据

诊断脚本：`scripts/diagnose_sdk_http.py`。使用生产 FastAPI 入口、真实 HTTP/SSE、真实 SDK 和用户已配置模型；没有替换 query 或终局处理。只隔离任务账本，宿主是合成数据适配器。

短任务：基线 4.061 秒；两轮持续健康检查分别 4.195 / 3.872 秒完成，同期成功 11 / 10 次健康请求。

Python 连续链：真实模型自主规划 → prepare_python_data → HTTP poll/receipt → 数据登记 → 真实 Docker 分析 → submit_result_plan → 计划返回 → HTTP result 取回。

- 47.427 秒完成，期间成功 115 次健康检查。
- 桥回执和结果查询 HTTP 200；返回单个 writeValues 白名单计划。
- 结果实值核对：金卡 400、银卡 200、未匹配 400、总计 1000；空键定位会员表 A5，重复键定位会员表 A3，未匹配定位订单表 A5。
- 证据：`outputs/sdk-python-http-audit-20261011/report.json`、`synthetic-result.json`；短任务证据在 `outputs/sdk-http-audit-20261011/report.json`。

因此“任何并发 HTTP 请求都会冻死 SDK”不成立。旧脚本观测到的停滞仍可能有其他原因；本次不能证明所有长任务、取消与孤儿进程场景均无问题，也没有用 HTTP 替代桥来绕过产品链路。

## 自动化和构建

- 服务端最终整包：1375 passed / 1 skipped，69.72 秒。日志 `outputs/codex-fix-server-verified-20261011.log`。
- WPS 全量：678 passed / 1 skipped。日志 `outputs/codex-fix-wps-final-20261011.log`。
- Excel 最终全量：1053 passed / 1 skipped，93 个测试文件通过。日志 `outputs/codex-fix-excel-verified-20261011.log`。
- 两端最终 build（含 tsc）均退出 0，日志 `outputs/codex-fix-build-{excel,wps}-final-20261011.log`；服务端 compileall 通过。
- schema 三项反向变异均红，还原 SHA256 一致、基线 15 passed；日志 `outputs/codex-schema-mutations-20261011.log`。

过程失败如实保留：一次 WPS 面板测试从根目录启动未加载 WPS 的 DOM 环境，改用项目 test:wps 后通过；新增回调使旧测试桩签名不匹配，已补参数且保留原断言；一条文件夹写入测试遇到 Windows 文件替换拒绝，单独及整包复跑均通过。账本恢复测试曾使用默认数据库并遇到 malformed，已改为独立临时账本，本轮未修复或迁移用户数据库。

## 生效与下一步

当前运行服务未重启，修复需重启 SDK 服务并刷新插件后生效。已准备 `outputs/sdk-manual-capabilities-20261012.json`，复制当前能力配置且仅显式启用 Python，不改变 Docker 校验、数据授权或写入确认。

退出旧后端后，在项目根目录的 PowerShell 中执行：

```powershell
$env:EXCEL_BRO_CAPABILITIES_PATH = (Resolve-Path 'outputs/sdk-manual-capabilities-20261012.json').Path
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start_claude_sdk_server.ps1
```

不要与 dev:server 同时占用 8765。随后刷新 WPS/Excel 插件，在测试工作簿上复验结果卡、生成副本审阅、确认写回、回读和撤销。本轮真实模型与 Docker 连续分析已取得证据；**真实 Excel/WPS 客户端本轮写回验收未执行**。不将合成宿主、面板替身测试或结果卡展示当作实机写回完成。
