# WPS 规划失败：账本恢复与错误透传（2026-10-12）

## 根因与证据

用户截图在“生成回答或计划”开始约 27ms 后失败，显示 SDK 流式接口不可用。实际监听 8765 的服务健康检查正常，但未知任务的状态请求返回 500。只读 SQLite 检查确认用户任务账本损坏：`planning_tasks` 表及规划记录索引存在 malformed page/btree 错误。数据库损坏发生原因尚未确定，不能据此归因于模型或并发 HTTP。

WPS 初次规划、续跑，以及 Excel 续跑把任意非成功 HTTP 响应都替换成了 SDK 缺失流接口提示。现在只对 404/405 或成功响应无 body 使用该提示；422/500/503 保留服务端错误码、消息，不追加重试或自动回退。服务端规划状态/续跑读取也统一捕获账本错误；SQLITE_CORRUPT/NOTADB 返回 503 `TASK_STATE_CORRUPT`，不向客户端回显路径或 SQLite 原始细节。

## 实际本机恢复

- 停止并核实原 8765 后端（PID 15176 / 36732），未停止 WPS/Vite，未操作工作簿。
- 原库：`C:\Users\jia\AppData\Local\Excel Bro\tasks\state.sqlite3`。
- 完整备份：`C:\Users\jia\AppData\Local\Excel Bro\tasks\recovery-20261012-1005`，包含原库、WAL、SHM及 SHA256 校验记录 `recovery-report.json`。
- 通过 `scripts/recover_task_ledger.py` 离线创建同 schema 新库。所有可读取非规划表逐字段复制并比较，保留 **6 条 office_executions**；其余非规划表均为 0 条。损坏的规划任务/事件/预算扩展仅保留在原库备份，不放回活动库，不重放旧任务。
- 新库 `integrity_check=ok`、外键检查无错误。脚本若遇到非规划表不可读、备份校验失败或本地服务仍监听 8765，会拒绝替换；没有自动重置生产库的启动逻辑。
- 后端已隐藏重启，8765 健康检查 200；未知任务状态恢复为结构化 404 `PLANNING_TASK_NOT_FOUND`。
- 本次人工验收使用 `outputs/sdk-manual-capabilities-20261012.json` 作为 `EXCEL_BRO_CAPABILITIES_PATH`（Python enabled=true），SDK 状态目录沿用 `LocalAppData\ExcelBro\ClaudeAgentSdk`。仓库默认 Python 开关不因此开启。

旧规划不能直接续跑；需要重新发起用户请求。执行登记记录仍保留，但这不等于验证了全部历史执行的可恢复性。

## 实际 8765 服务验证

```powershell
.\.venv\Scripts\python.exe scripts/diagnose_sdk_http.py --base-url http://127.0.0.1:8765 --model <SDK目录ID> --python --out outputs/live-ledger-sdk-20261012
```

`--base-url` 使用现有回环后端及其实际账本；不指定则沿用独立诊断服务/独立账本。此诊断使用真实模型额度、真实 Docker 和合成宿主数据，不向真实工作簿写入。

本次 task：`turn-sdk-http-audit-c049f8af76d14c87945d3917f1845ef3`。

- 3.831s：`bridge.prepare_python_data`，poll/receipt 成功。
- 31.131s：`python.completed`。
- 37.972s：`submit_result_plan`。
- 42.443s：收到 `kind=plan`，单个 `writeValues`，结果取回 200；总耗时 42.451s，期间 102 次并发健康检查。
- 数值：金卡 400 / 银卡 200 / 未匹配 400 / 总计 1000，空编号/重复编号/未匹配明细均有定位。
- 随后从实际任务状态端点读到 `completed`。

证据：`outputs/live-ledger-sdk-20261012/report.json` 与 `synthetic-result.json`。这是实际后端、实际账本、真实 SDK/Docker 的证据，**不是 WPS 客户端读表、点击审阅、确认写回的实机验收**。

## 本轮回归

- Excel api + WPS planning stream：84 passed。
- 账本错误/离线恢复/账本/恢复描述：31 passed。
- main + 账本错误/离线恢复：58 passed（与上一组有 4 条重叠，不合并虚增）。
- 两端 build（均含 tsc --noEmit）成功；Python compileall 成功。
- 未重跑全部历史测试集；未提交、推送或发布。

用户下一步：重新加载 WPS 插件面板，确认同时授权会员表、订单表及写入目标汇总表，重新发送原需求。授权目标不等于同意修改来源表，写入仍须预览并由用户确认。
