# A1 执行恢复证据矩阵（2026-10-24）

本批为 A1 第一项：核对九场景覆盖，并补齐 ①③⑥ 的服务端故障注入证据。不代表 A1 整体完成。执行交互保持先执行、再审阅；恢复中的现场核对不是执行前确认门槛。

## 本批新增证据

`server/tests/test_execution_recovery_a1.py` 共 10 条（Office/WPS 各 5 条），调用真实 FastAPI 路由和隔离 SQLite 账本。每条独立临时库；没有调用模型、Docker 或真实宿主。登记载荷分别沿用 Office 计划协议与 WPS 无值描述协议。

- 在途动作挂起后仍拒绝直接 resume，拒绝前后状态逐字段相同。
- 三动作完成第 0 项、第 1 项回执缺失：账本必须是 next=1/inFlight=1。核对 applied 后只接受第 2 项；not_applied 后接受第 1、2 项。此处是协议动作序列，不是宿主写入计数。
- unknown 不允许补记为完成，保持原状态。
- 两个独立 HTTP 客户端以同一 revision 并发认领，必须一个 200、一个 409；胜者 token 必须轮换。旧 token 携带最新 revision 的完成回执仍被拒绝，胜者回执可成功，排除仅靠旧 revision 拒绝的假证据。
- 完成后两次 GET 返回相同状态且不下发 ownerToken，再认领被拒。该证据不证明 UI 重开时只发 GET。

Excel `excelWriteEndToEnd.test.tsx` 原有回执丢失用例使用 next=1/inFlight=0 的不可能账本状态；修正为核对前 next=0/inFlight=0、核对后 next=1。新增三动作 applied/not_applied/unknown 三条，跑真实 App、执行器、核对函数和替身 Office 宿主，直接断言只写 A3 或 A2/A3、已完成 A1 不被改写、unknown 零写入且不提交补记。首跑 not_applied 夹具错误地以 null/白色填充代表宿主空字符串/无填充，产品正确拒绝，修正夹具后通过。

## 反向变异

命令：`.venv/Scripts/python.exe scripts/probe_a1_execution_recovery_mutations.py`

脚本在独立 Python 子进程中替换生产方法的内存函数，不写生产源码。目标唯一匹配，pytest 必须退出 1 且有指定用例失败；收集错误不计成功。日志在 `outputs/a1-execution-recovery/`，驱动在 `outputs/a1-mutation-driver.log`。生产 task_ledger.py SHA256 前后相同。

| 变异 | Office | WPS | 能证明什么 |
| --- | --- | --- | --- |
| A1_INFLIGHT：删除 claim 的在途检查 | 红 | 红 | 未核对不能直接认领 |
| A1_CURSOR：applied 不推进序号 | 红 | 红 | 已应用动作不能留在待执行序号 |
| A1_CAS：删除 claim 的 revision 比对 | 红 | 绿 | Office 并发判据；WPS 仍被 owner lease 拦截，不能宣称独立钉住 WPS CAS |
| A1_OWNER：完成回执不检查 owner token | 红 | 红 | 最新 revision 不能让旧 owner 冒领 |

## 九场景与两端差距

| 场景 | Office 证据/入口 | WPS 证据/入口 | 尚缺 |
| --- | --- | --- | --- |
| ① 写入成功回执丢失 | App 恢复→reconcileOfficeRangeAction→回执→executePlan；excelWriteEndToEnd 已检查剩余区域写入；本批 API applied | WpsAgentPanel 最后动作核对→补记完成；既有面板测试零宿主写入；本批 API applied | 两端原生宿主故障注入；WPS 未完成后缀续写未实现 |
| ② 运行中断开/重启 | 状态查询与在途检查已有；本批挂起后 claim 拒绝 | owner lease+只读检查已有 test_wps_execution_api | 真正杀进程/重启后两端取证；不能用主动 suspend 冒充重启 |
| ③ 多动作中途断开 | 本批三动作协议序列及三分支真实 App/执行器/替身宿主写入计数 | 本批服务端序列可推进；面板在剩余动作时明确阻止 | WPS 重开后精确续跑未实现；Office 同条链的真实 HTTP+宿主桥注入及两端实机未验 |
| ④ 旧结果迟到 | 暂停/段守卫既有；本批旧 owner+最新 revision 拒绝 | 同上 | 本批未补模型迟到路径变异，不能以回执变异代替 |
| ⑤ 换簿/收窄/改来源 | scope identity、target guard、Office 授权用例已有 | WPS 面板身份/来源冲突拒绝已有 | 统一恢复交错矩阵及各分支变异、实机 |
| ⑥ 双恢复请求 | 本批真实并发 HTTP+SQLite CAS | 本批同路由并发+owner lease；旧 token 拒绝 | 两窗格真实竞争；WPS CAS 单独变异不红 |
| ⑦ 引用过期 | 本地结果解析与桥拒绝已有 | 同协议，结果卡降级已有 | 恢复过程中完整过期→重读重算端到端及变异 |
| ⑧ 完成任务重开 | excelWriteEndToEnd 已验证启动一次 GET、零写入；本批完成后拒 claim | WPS 恢复卡现有已完成只读判断 | 连续两次 UI 重开、无 POST 的独立证据和变异 |
| ⑨ 同名/其他任务串用 | test_task_ledger/test_main 身份拒绝；namespace 守卫 | test_wps_execution_api 跨 namespace、同 operation 不同 task 拒绝 | 两端恢复 UI 完整串用链与独立变异 |

## 必须纠正的主线状态

WPS `WpsAgentPanel.tsx` 的恢复分支明确说明：目标矩阵与确认快照只在会话内，重开后无法精确续跑剩余动作；只核对/补记最后完成动作，不重新生成整份计划冒充恢复。这是**未实现**，不是“实机待验”。服务端登记无值恢复描述，不应为了续跑随手把原始数据矩阵塞进账本。

下一实施点：WPS 剩余动作恢复所需的本机可验证材料与生命周期设计，并检查 Office 三动作恢复全链；必须先明确材料的归属、过期、授权、来源、恢复失败语义，再实现。A2 保留期与 SDK 原生会话持久化尚未开启。

真实 Excel/WPS 的在途断开→只读核对→明确继续均无本批实机证据。用户已验收的双视图、确认、放弃、撤销只计审阅链，不计恢复链。

## 本批验证

core 790 passed；WPS 1023 passed / 3 skipped，构建通过；服务端全量 1736 passed / 1 skipped，compileall 通过。新增 Excel 三分支后定向套件 52 passed；Excel 全量与构建最终结果见 outputs/a1-next-excel-final.log 与 outputs/a1-next-build-final.log。四项变异检出，生产源码未改写。

最终 Excel 全量复跑 1518 passed / 3 skipped，构建通过；lint 0 errors / 37 既有 warnings，diff-check 通过。首次新增用例后的全量中，未改动的 folderScopeSafety「failed selection checks stay in the existing picker」出现一次 pending=false/true 断言失败，原日志保留于 a1-next-excel-final.log；该文件定向复跑 14 passed，随后全量通过（a1-next-excel-recheck.log）。未改该用例，未证明偶发失败根因，不能以复跑通过宣称已修复。
