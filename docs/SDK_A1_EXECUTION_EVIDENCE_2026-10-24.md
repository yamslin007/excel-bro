## 2026-10-25（第四轮）：A1-F1 收尾 —— `ownedRules` 复用同一份类型判据

`ownedRules`（`Add` 回空后的找回、标记失败的重试清理、重载清理）原先对"读不出 `Formula1` 的规则"一律抛错，于是同区域只要有用户自建的色阶/数据条：①`Add` 建好却回空时找不回本次规则、报"未能建立可独立移除的审阅标记"；②`recoverWpsResultMarks` 永远保留记录、每次重开都提示"未清理"。现在把第三轮副本清理的类型判据（`NON_EXPRESSION_RULE_TYPES` + `isKnownNonExpressionRule`）**上移成两个消费者共用的一份**（全文件一处定义），明确可识别的非表达式类型跳过且不修改不删除；空句柄、未知类型、类型读不成数字、表达式规则缺公式仍抛错并保留记录；归属仍只按完整随机公式。先红 3 failed / 35 passed（`outputs/a1-f1-owned-red.log`），改后 39/39 绿；反向两个变异各钉一头：N14（读不出就一律跳过）3 红、N15（回到旧行为一律抛）4 红，还原后 sha256 `b2ecda28…`。本轮 `test:wps` 1097/3、`test:addin` 1538/3、`pytest` 1750/1、两端 build 与 `compileall` 退出 0、lint 0 error / 36 既有 warning、`git diff --check` 退出 0。Excel 端按 `entry.id` 定位、不枚举同区规则，无同类缺陷。A1-F1 范围内已知代码缺陷为 0；仍缺真实窗格的原生证据（不是代码缺陷）。

## 2026-10-25（第三轮）：A1-F1-a 副本清理判据补齐

`clearReviewMarksFromCopySheet` 原先在三种"读不出来"的宿主形态下会说反话：`Count` 缺失被 `Number(Count ?? 0)` 当成 0 条、`Item(i)` 返回空被静默跳过、规则读不出 `Formula1` 一律当色阶/数据条跳过 —— 三种都会在副本仍带本次高亮时返回 `{cleared:0, unconfirmed:0}`，面板于是没有任何提示。现在：读不出条数 / 空句柄 / 类型未知或缺公式 ⇒ 该区域 `unconfirmed`；只有宿主**明确**给出可识别的非表达式类型（色阶 3、数据条 4、前 10 项 5、图标集 6、唯一值 7、空值 9、时间段 10、平均值 11、非空 12、错误 13、无错误 14）才跳过；每个区域最多计一次（按去重后的地址），归属判据仍只有本次完整随机公式。先红 7 failed / 37 passed（`outputs/a1-f1a-red-round3.log`），改后 44/44 绿；M9–M13 五个单点变异全红并逐字节还原（`c7ca2a0d…`），其中 M9 在位时面板用例单独变红（1 failed / 176 passed），证明"真实清理结论 → comparisonNotice → 面板文案"这条链有判别力。本轮实跑 WPS 1091/3、Excel 1538/3、服务端 1750/1，两端 build 退出 0，lint 0 error / 36 既有 warning，`git diff --check` 退出 0。真实窗格仍未点击；`ownedRules`（重试/重开路径）对同区域色阶仍按"读不出来"处理，属同类但未修的过量提示，见计划表。

## 2026-10-25：A1-F1 残留修复（公式留存 + 多轮归属）

标记建到一半失败时不再丢公式：`markWpsResult` 增 `onFormula`，在第一次 `Add` 前回调，面板存进 `currentMarkFormula` 交给双视图，副本按完整随机公式清本次残留；`clearReviewMarksFromCopySheet` 返回 `{cleared, unconfirmed}` 并经 `WpsNativeComparison.reviewMarks` 上报，`unconfirmed>0` 时审阅条如实提示。归属记录由单槽改为 `{v:2, records:[…]}`：新旧两轮并存、各自清理、只处理当前簿（别的簿的记录与其规则都不碰），`v:1` 单条仍能读出并在下次写回时升级。修复前红灯 14 failed / 21 passed（`outputs/a1-f1-red-round2.log`）；M4/M5/M6/M8 四个单点变异全部变红后逐字节还原。本轮 WPS 1081/3、Excel 1538/3、服务端 1750/1，两端 build 与 lint(diff-check) 均退出 0。真实窗格仍未点击。

## 2026-10-24：A1-F1 / A1-F2 收尾核查（未改生产源码）

按七条故障清单复核生产路径，未发现新的产品缺陷；补 2 条行为回归（部分清理失败后重开重试、来源失效只作废该任务），3 个单点变异全部变红并逐字节还原（sha256 与备份一致）。本轮实跑：WPS 1071 passed / 3 skipped、Excel 1538 passed / 3 skipped、服务端 1750 passed / 1 skipped，两端 build（含 tsc）与 compileall 退出码 0。日志 `outputs/a1-round-*.log`、`outputs/a1-f1-mut-*.log`、`outputs/a1-f2-mut-M3.log`。仍是替身宿主 + 生产组件证据，真实窗格未点击；详见主线计划表第 5 节。

## 2026-10-24：A1 独立进程退出与竞争

已补 Office/WPS 各两条独立进程测试：第二动作 start 提交后强杀测试 worker、全新 worker 读取持久在途状态；两个独立 worker 同 revision 认领只有一个成功。WPS 租约未过期仍拒绝核对，过期部分使用测试时钟。未操作真实宿主，未启动网络服务器；不能等同原生窗格崩溃/竞争。详见 [过程与边界](SDK_A1_PROCESS_RECOVERY_2026-10-24.md)。本批无生产代码修改，A1 未整体完成。下一项为恢复材料过期/缺失闭环，A2 不提前开启。

> 更新说明（2026-10-24）：下方为 A1 首批历史证据，其中“WPS 后缀续写未实现/精确续跑未实现”已被后续实现取代；当前已实现并有替身宿主及用户分步恢复反馈，详见 SDK_MAINLINE_PLAN_2026-10-19.md、SDK_WPS_RESUME_2026-10-24.md。连续两次 UI 重开无写入/补记证据、双端同帧双击互斥证据本轮已补；真正杀进程、两窗格竞争等仍未闭环。原批次测试数与变异结果保留为历史，不作为最新总数。

## 2026-10-24：A1 恢复入口并发与重复重开

本批故障回归复现两端同帧双击竞态：WPS 在 targetGuard.check 的 await 之后才取得恢复锁，Excel 恢复入口没有同步互斥。两端现在均在首次 await 前取得本地恢复锁，所有退出路径 finally 释放。服务器 revision/owner 校验保持不变；本地锁不替代跨窗格服务端互斥。

证据：WPS 同帧两次点击、暂停认领 Promise 时只允许一个 claim，完成后仅一个 finish、零宿主内容写入。Excel 原三动作中断用例改为同帧双击，修复前两个审阅分支 finish=2 失败，修复后仍仅续写 B4、start 序号 [0,1,2]、complete 序号 [0,2]、finish 一次。WPS 认领和检查点请求等待期间来源变化/授权收窄四种组合均拒绝后续写入。两端完成后连续两次挂载均不调用写入或补记接口；Excel 清除索引，WPS 保留可清除的只读历史提示，两者展示不同但都不重放。

这些是生产组件/执行器 + 替身宿主自动化证据，双击不是原生实机证据；服务端两客户端竞争由 test_execution_recovery_a1.py 真实 HTTP + SQLite 覆盖。A1 尚缺真正杀进程恢复、引用完整过期链、跨窗格竞争等，不能宣称 A1 全部完成或已进入 A2。用户上一轮“歐克了”作为范围恢复反馈，不扩张为这些故障场景均实机通过。

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
| ⑦ 引用过期 | 2026-10-25 A1-E1：开发专用精确故障入口（`EXCEL_BRO_RESULT_REF_FAULT=1`）在"已登记、未首次写入"处让第一个 Python 引用失效，12+1 条服务端用例与 10 个单点变异；旧引用被真实 `resolve_result_ref` 拒绝且句柄已删 | 同协议；WPS 端起目标与桥逐字一致（6 条）与面板接线（2 条） | 恢复过程中完整过期→SDK 重读重算→新计划只执行一次端的端到端，仍属 A1-E2 |
| ⑧ 完成任务重开 | excelWriteEndToEnd 已验证启动一次 GET、零写入；本批完成后拒 claim | WPS 恢复卡现有已完成只读判断 | 连续两次 UI 重开、无 POST 的独立证据和变异 |
| ⑨ 同名/其他任务串用 | test_task_ledger/test_main 身份拒绝；namespace 守卫 | test_wps_execution_api 跨 namespace、同 operation 不同 task 拒绝 | 两端恢复 UI 完整串用链与独立变异 |

## 必须纠正的主线状态

WPS `WpsAgentPanel.tsx` 的恢复分支明确说明：目标矩阵与确认快照只在会话内，重开后无法精确续跑剩余动作；只核对/补记最后完成动作，不重新生成整份计划冒充恢复。这是**未实现**，不是“实机待验”。服务端登记无值恢复描述，不应为了续跑随手把原始数据矩阵塞进账本。

下一实施点：WPS 剩余动作恢复所需的本机可验证材料与生命周期设计，并检查 Office 三动作恢复全链；必须先明确材料的归属、过期、授权、来源、恢复失败语义，再实现。A2 保留期与 SDK 原生会话持久化尚未开启。

真实 Excel/WPS 的在途断开→只读核对→明确继续均无本批实机证据。用户已验收的双视图、确认、放弃、撤销只计审阅链，不计恢复链。

## 本批验证

core 790 passed；WPS 1023 passed / 3 skipped，构建通过；服务端全量 1736 passed / 1 skipped，compileall 通过。新增 Excel 三分支后定向套件 52 passed；Excel 全量与构建最终结果见 outputs/a1-next-excel-final.log 与 outputs/a1-next-build-final.log。四项变异检出，生产源码未改写。

最终 Excel 全量复跑 1518 passed / 3 skipped，构建通过；lint 0 errors / 37 既有 warnings，diff-check 通过。首次新增用例后的全量中，未改动的 folderScopeSafety「failed selection checks stay in the existing picker」出现一次 pending=false/true 断言失败，原日志保留于 a1-next-excel-final.log；该文件定向复跑 14 passed，随后全量通过（a1-next-excel-recheck.log）。未改该用例，未证明偶发失败根因，不能以复跑通过宣称已修复。

本轮验证：Excel 1535 passed / 3 skipped；WPS 1057 passed / 3 skipped；两端 build（含 tsc）通过并保留 VITE_ENABLE_RECOVERY_PROBE；服务端 1744 passed / 1 skipped，compileall 通过；lint 0 errors / 36 warnings，git diff --check 通过。pytest 另有既有缓存目录无写权限警告，未影响用例结果。日志 outputs/a1-followup-*；修复前红证据 outputs/a1-excel-double-red.log 和 outputs/a1-followup-wps.log。首轮 Excel 连续重开测试误用了不存在的 .input-area，改为现有 .composer textarea 后通过，该失败不计产品缺陷。
