# A1 S2 原生恢复收尾（2026-10-26）

本轮完成第二步写入后精确中断、核对并只续第三步的原生观察，并修复两端恢复路径漏报宿主写调用的缺陷。Excel 已实际刷新 WebView；WPS 仅关闭/重开窗格，页面仍复用，不能宣称跨 WebView 或进程重启通过。A1-GATE 仍未通过，A2 未启动；未提交或推送 Git。

## 1. 缺陷与修复

Excel 首轮真实操作已写完第三步，但 host-writes 只有动作 0/1，观测 partial、缺 1 个动作。这是上报遗漏，不能把缺记录读成零调用。负向现场保留在本轮 run 的 s2-excel 目录，operationId 为 execution-a35f651e-c452-4414-afbe-2d78d7def84f。

修复 App.tsx 的 resumeInterruptedOfficeExecutionOwned，以及 WPS 的 wpsExecutionResume.ts：收集恢复执行器 onHostWrite，在账本 finish/suspend 后 finally 中 best-effort 上报，并携带执行所有权。不上移到写入与回执之间；上报失败不改变业务结果。恢复校验、认领、来源边界与先执行再审阅语义未改。

先补失败断言：Excel 2 failed/1 passed，WPS 2 failed；修复后 Excel 专项 3 passed，WPS 恢复套件 24 passed。日志为 outputs/s2-audit-{excel,wps}-{red,green}.log。

## 2. 原生证据

证据根目录：outputs/a1e3-native/runs/20261026T132024-582e07b4/。manifest.json 含隔离实例令牌，不能公开或作为截图证据上传。

| 项目 | Excel R2 | WPS |
| --- | --- | --- |
| 工作簿 | S2-Excel-R2.xlsx | S2-WPS.xlsx |
| taskId | turn-c289119e19dc4ff3a5db0f6922d3190b | task-b8cd0995-5c89-4887-85aa-cd445ccefcee |
| operationId | execution-9a65aae1-3861-43de-a49a-c8b6732343e6 | wps-execution-682a23bf-465c-4f97-88bf-670b2c3843c7 |
| 证据子目录 | s2-excel-r2 | s2-wps |
| 中断时目标 B2:B4 | 111 / 250 / -303 | 111 / 250 / -303 |
| 恢复并确认后目标 | 111 / 250 / 333 | 111 / 250 / 333 |
| 来源始终 | 111 / 250 / 333 | 111 / 250 / 333 |
| 最终账本 | completed，next_action_index=3，in_flight=null | 同左 |
| 观测覆盖 | complete / nonzero，3 个动作，无缺失 | 同左 |
| 宿主调用 | 6 excel_sync_batches，各动作 2 | 3 wps_api_calls，各动作 1 |
| 内容写调用 | 每个动作 mutatingCalls=1 | 每个动作 mutatingCalls=1 |
| 重开证据 | 关闭/打开并右键刷新 WebView | 关闭/打开，页面复用；完整重载未验 |

两端都启用开发面板「下一次第二步中断」，而非凭肉眼关窗抢时间。Excel 双击「核对并继续」；WPS 先核对在途动作，再双击「继续剩余动作」。每端原始 HTTP 日志都是 actions/start 3 次，其中恢复后仅 1 次，actions/reconcile 1 次、最终 complete 1 次。WPS 另有 resume 1 次；Excel 核对入口接管所有权，没有单独 resume 请求。

请求数不能单独证明写入次数。这里同时对照每动作宿主计数、账本推进、保存后独立读取的单元格结果：已完成前缀未重写，恢复仅写第三个动作。两端单位不同，不相加、不互换。

WPS 第二个动作保留原始注入异常观测：failed_after_write / not_applicable，hostWriteCalls=1；后续核对补记该动作已应用、不再次写入。不能把它篡改为 written/verified。第 1、3 个动作是 written/verified。Excel 三个动作均 written/verified。

双视图显示执行前 -101 / -202 / -303；确认后标记消失、临时对比簿关闭、输入框恢复。两端 resumed-host-writes.json 与 confirmed-host-writes.json 一致，确认不重执行。interrupted.xlsx、confirmed.xlsx 为真实 UI 保存后独立读取的证据；各目录附四张 interrupted/resumed/comparison/confirmed.png、operation-http.log 与 summary.json。

收尾发现 summary.json 初版误用 reconcile/claim/finish 作为 URL 后缀，统计为 0；已按原始日志改为 actions/reconcile/resume/complete。原始日志未改，没有把“0”当作实际未发生。

## 3. 检查结果

以下均为修复后实跑，退出码 0；日志前缀 outputs/s2-。

| 检查 | 结果 | 日志 |
| --- | --- | --- |
| test:addin | 1560 passed / 9 skipped | addin-full.log |
| test:wps | 1130 passed / 8 skipped | wps-full.log |
| pytest server/tests | 1842 passed / 1 skipped | server-full.log |
| build:addin / build:wps | 通过（包含 TypeScript） | build-addin.log / build-wps.log |
| lint | 0 error / 36 warnings | lint.log |
| compileall server/app | 通过 | 命令退出码 0 |
| test:a1-e2 | Excel 6 / WPS 188 passed，六份工件自洽 | a1-e2-gate.log |
| git diff --check | 通过 | diff-check.log |

既有提示保留：WPS 包体积提示、门控测试 React act/Node DEP0190 提示；pytest 缓存目录权限提示，测试本身通过。core 本轮未重复运行，不把此前 845 当成本轮数字。文档归档后只补差异检查，不重复无变化的全量。

## 4. 尚未完成与边界

- WPS 完整 WebView/进程退出恢复仍待验；本轮关窗仅复用页面。双击入口也不等于两个独立窗格竞争。
- S3「确认/放弃后重开不重执行」仍单独待验。本轮确认不重写已有证据，未补本轮放弃及确认后的再次重开。
- Excel 实际刷新后固定选表退回当前表；本轮人工重新勾选原来来源/目标两表，未扩大授权。固定选择跨刷新保留未在本轮修复。
- Excel 对比副本列宽不足时显示 #，手动展开 B 列后负数完整。列宽显示问题未修。
- Excel 恢复执行完成后，历史过程中的“正在核对 Office 写入断点”仍可能以未结束警告保留；账本、底部审阅和确认正确。过程收尾展示单独待修。
- 来源过期原生全链、完整故障矩阵等仍按主计划，不能把 S2 子场景写成 A1 整体完成。

仅使用独立测试工作簿；用户既有改动保留；不清理旧账本或其他未跟踪文件。收尾停止本轮隔离 8877，并将两个开发前端恢复默认 8765 连接；8765 用户后端未启动或停止。


## 5. 环境清理时发现的脚本误拒绝

stop 首次拒绝 PID23572；核对后进程路径、命令、创建时刻均属于本轮，只是 JSON 将时间的小数尾零正规化（.4412640Z / .441264Z，部分 PowerShell 版本还解析成 DateTime）。已改按 DateTimeOffset.UtcTicks 比较同一时刻，完整命令/可执行文件及专用模块检查仍保留。新增 test_a1e3_native_stop.py 调用真实 Stop-OwnedServer，替换进程读取与 taskkill：两种等价表示可通过，时间/路径/命令不符均拒绝，没有实际杀进程。之后实际 stop 成功，仅结束本轮测试进程树。

此脚本修复在全量检查之后；另跑原生验收脚本专项及最终 diff-check，不冒充重新跑了整套全量。两个默认前端已恢复，8877已停止；8765当前无监听，未擅自启动日常后端。


清理专项首跑 2 failed/29 passed：一条旧测试要求畸形时间仍报标准身份不匹配，已在解析失败时保留拒绝语义；新 PowerShell 测试捕获预期异常后默认退出码为 1，补显式 exit 0（所有断言后）以表达成功。未放宽旧断言。最终原生脚本专项 31 passed，日志 outputs/s2-native-close-tests.log（首跑）与 s2-native-close-green.log（最终）。这新增 1 条不混入之前服务端全量 1842 的数字。
