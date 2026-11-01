# 重开之后撤销「未完成执行」已经写入的内容（Excel，2026-10-31）

> **2026-11-01 补记（不改正文）**：本文 §7 写的"真机未验收"是当时的时点结论。此后用户按 §8 在真实 Excel 里
> **点了「↶ 撤销本次执行」并成功**：`execution-6d29e377…` 账本 `undone_at=10-31 10:30:09`，`公式来源证据验收-Excel-B.xlsx`
> 只读核验 `结果1/2/3!B2` 全空、`来源!B5=500`。所以这一条现在的准确状态是"**真机已通过**"。
> 证据台账见 `docs/CONVERSATION_UX_REFACTOR_2026-11-01.md` §8.5。

真机现场（用户提供，不重跑）：

- 文件 `C:\Users\jia\Desktop\测试\公式来源证据验收-Excel-B.xlsx`；
- 授权 `来源 / 结果1 / 结果2 / 结果3`；三个 `writeFormulas`：`结果1!B2==来源!B2*1`、`结果2!B2==来源!B4*2`、
  `结果3!B2==来源!B5*3`；
- 在第二步回读之后、完成回执之前中断（探针）；
- 用户核对在途动作后把 `来源!B5` 由 5 改成 500，再点继续 → 界面正确阻断：
  「来源已变化，不能沿用旧结果续跑。」；
- 磁盘只读核验：`来源!B5=500`、`结果1/2!B2` 仍是公式（缓存值 2 / 8）、`结果3!B2` 空；
- **重开 Excel 后恢复卡仍在、仍阻断，但没有「撤销本次执行」**。截图只能证明入口不可见，
  不能证明材料丢了。

## 1. 根因（读代码核实，不是猜）

1. 撤销栈（`apps/excel-addin/src/hooks/useUndoSnapshot.ts`）**只在** `persisted.status === "completed"`
   的分支里用 `result.undoSnapshot` 重建。未完成 / 被阻断的执行一条都没重建。
2. 材料其实一直在：每个已完成动作的写前快照由 `complete_office_action(checkpoint=…)` 落库，
   `GET /api/office/executions/{id}` 也把它（`checkpoint`）原样返回。**材料有，入口没有。**
3. 处置留痕（`office_executions.abandoned_at` / `undone_at`）此前**没有**投影进状态响应，重开后
   无法判断"这个入口是不是已经被用掉了"。

结论：不是材料丢失，是"**未完成分支从来没把账本里的材料挂进栈**"，加上"续跑被拒"被误当成
"这次执行已经没救了"。两种资格被混成了一种。

## 2. 两种资格（本轮的核心口径）

| | 继续执行（续跑） | 撤销已写入内容 |
| --- | --- | --- |
| 依赖 | 身份 + 状态 + 授权 + 在途材料 + **来源证据仍一致** | **已验证的写前恢复点** + **目标现场仍是这次执行留下的样子** |
| 来源被用户改了 | **必须拒绝**（不能拿旧结果继续写） | **不受影响**（撤销不读来源，也不重新执行计划） |
| 判据出处 | `packages/core/src/pendingExecutionResume.ts` | `packages/core/src/executionUndoRebuild.ts` |

`rebuildExecutionUndo` 的判定顺序刻意是"先身份、再处置、再状态、再材料、最后授权"，每个原因码
只在一处返回点产生：`identity_unavailable` / `identity_mismatch` / `disposition_recorded` /
`already_completed` / `in_flight_unknown` / `nothing_written` / `material_missing` /
`coverage_incomplete` / `plan_mismatch` / `plan_target_never_authorized` /
`plan_target_not_authorized_now`。入参里**没有**任何来源证据字段 —— 这条由单测钉住
（"第二步中断、核对后来源已变化：撤销仍然可用"）。

## 3. 实际改动

| 位置 | 改动 |
| --- | --- |
| `packages/core/src/executionUndoRebuild.ts`（新） | 唯一判据 + `UNDO_RECOVERY_LABEL` + `undoRecoveryDetail`（明说"来源保持你现在的修改""撤销不会重新执行计划"） |
| `packages/core/src/index.ts` | 导出上述类型与函数 |
| `hooks/useUndoSnapshot.ts` | 新增 `ensureUndoSnapshot`（按 `executionId` 去重挂载账本材料）；`pushUndoSnapshot` 改为**同一 executionId 覆盖**；栈改为 **ref + state 同步**（`ensure` 之后同一 tick 就能撤到刚挂进去的那条） |
| `App.tsx` | `officeLedgerStates`（最近一次账本读数，不额外发 GET）＋ `rememberOfficeState`（材料够时挂快照）＋ `officeRecoveryUndoFor`（渲染期按同一读数现算）＋ `undoOfficeRecoveredExecution`（点击时再幂等挂一次，然后走既有 `undoExecutionTo`） |
| `App.tsx` | 账本已处置（`undoneAt` / `abandonedAt`）时恢复卡判为 **blocked**（原文案分开说"撤销过"与"放弃过"），避免已回退的执行还能续跑/补记 |
| `App.tsx` | 消息里那条 `↶ 撤销本次执行` 在**恢复卡接管同一条执行**时让位（只留 `EXECUTION_RECOVERY_ENTRY_HINT`）—— 一个执行只有一个可点入口，且在输入框上方 |
| `components/OfficeRecoveryPanel.tsx` | 正文新增撤销行：`available` 给按钮（`data-recovery-action="undo"`），`blocked` / `done` 只给原因/结论（`data-recovery-undo`） |
| `server/app/{models,task_ledger,main}.py`、`contracts.ts` | 状态响应新增只读 `disposition`（`undoneAt` / `abandonedAt`），三处同形 |
| `excel.ts` | `recoverableActionRange` 导出（撤销目标与执行目标必须同一个函数解析） |

**没有做的事**：没有新增执行前确认；没有改变"计划到达即执行一次、回读后审阅"的流程；没有用当前
单元格反推写前值；没有为"凑一份能撤销的快照"去现读；撤销成功前不写留痕，失败不写成功。

## 4. 自动化证据（替身宿主 + 真实产品代码）

新增 `excelWriteEndToEnd.test.tsx` 一组 5 条（`Excel 未完成执行：重开后从账本材料撤销已写入的内容`）：

1. 中断 → 核对 → 改 `来源!B5=500` → 续跑被拒（零写入）→ 卸载 + `localStorage.clear()` + 重开 →
   从账本材料恢复「↶ 撤销本次执行」→ 撤销后 `结果1/2/3!B2` 全空、来源仍是 500、没有重新执行计划
   （`registerOfficeExecution` / `startOfficeAction` / `completeOfficeAction` 计数不变）、留痕一次且是
   `undone`、入口被消费（不再有按钮）。
2. 重开后目标被人工改过：`assertUndoPostState` 在**任何写入之前**拒绝，被改的格保留用户的值，
   另一格仍是公式，零恢复写入、无 `undone` 留痕、不宣称"已撤销"（入口保留以便改回去重试）。
3. 重开后授权被收窄到只剩来源表：不给按钮，只给 `plan_target_not_authorized_now` 的原因，零写入。
4. 账本已有 `undoneAt`：不再重建入口，也不再摆成"可继续"。
5. 账本没有 `checkpoint`：不给撤销，也不用当前单元格反推。

判据单测 `packages/core/src/executionUndoRebuild.test.ts` 18 条；hook 级 4 条（去重、覆盖、
`ensure` 之后同 tick 撤销、失败保留可重试）。服务端新增
`test_state_response_projects_disposition_so_a_reopened_pane_does_not_offer_it_again`。

## 5. 反向变异（`scripts/mutate_office_undo_rebuild.mjs`，四判据全部成立）

| 变异 | 判别点 | 结果 |
| --- | --- | --- |
| U1 `ensureUndoSnapshot` 变成空操作 | 材料在账本里但没进栈 | 红：撤销根本没跑 |
| U2 把 `needs_reconciliation` 并进 `completed` | "未完成 ⇒ 不可撤销" | 红：连撤销说明都不渲染 |
| U3 状态响应 `disposition=None` | 重开不知道入口已被用掉 | 红（pytest） |
| U4 同一执行的两份快照都留在栈里 | LIFO 恢复两遍 | 红：「放弃变更」被误判"恢复未成功" |

每个变异都记了基线 sha256、失败原文、还原后逐字节一致、还原后复跑绿；脚本结束时工作区干净
（`node scripts/mutate_office_undo_rebuild.mjs` 四个目标全部 `clean`）。

## 6. 两端差异（**不宣称一致**）

| | Excel（本轮） | WPS（未改） |
| --- | --- | --- |
| 写前材料 | 账本 `checkpoint`（服务端持久） | 宿主 `PluginStorage` 的 `WpsUndoRecord` + 本机恢复包（有 TTL） |
| 判据出处 | 共享 `rebuildExecutionUndo` | 宿主适配层 `prepareWpsResume` / `inspectWpsUndoRecord` / `wpsRecoveryVerdictFor` |
| 重开后撤销入口 | **本轮补上**（恢复卡，输入区上方） | 撤销入口挂在结果卡上；恢复包在续跑判定里会被水合回宿主缓存（`ensureWpsRecoveryMaterial`），但**水合后是否让撤销入口回来，本轮没有证据** |
| 处置留痕 | `undone` / `abandoned` 两值（同一契约） | 同一契约、同一端点 |

WPS 侧本轮**只做核实、不改代码**：不重写已验证链路，也不在没有证据的情况下宣称两端一致。
把"重开后撤销未完成执行"接到 WPS 上（复用水合 + 一条与 `rebuildExecutionUndo` 等价的判据）是
后续任务；需要真机证据才算完成。

## 7. 仍未闭环

- **真机未验收**：全部结论都是"真实 App 组件 + 真实服务端契约 + 替身宿主"的自动化证据。
- 恢复入口的**可见位置**只到"真实浏览器布局"（`scripts/check_recovery_entry_viewport.mjs`），
  真实窗格未量。
- 服务端全量 pytest 有 2 条与环境有关的历史失败（`test_python_event_delivery.py`，Docker 引擎
  未运行 ⇒ python 执行返回 `unavailable`），与本轮改动无关。
- 未提交、未推送；A1-GATE 不自动通过，A2 未启动。

## 8. 最短真机验证（只跑这一遍）

先判断旧材料够不够：旧现场（`公式来源证据验收-Excel-B.xlsx`）的账本里已经有 `checkpoint`
（第 1、2 步），`来源!B5=500`、`结果1/2!B2` 仍是本次写入的公式、`结果3!B2` 空 —— 这正是
"材料充分且来源已变"的现场，**可以直接复用，不必重跑 A/B 轮**。

1. 用新版窗格打开 `C:\Users\jia\Desktop\测试\公式来源证据验收-Excel-B.xlsx`（保持 `来源!B5=500`）；
2. 等恢复卡出现，确认它仍说"不能继续"（来源已变化）；
3. 看卡片正文里是否出现 **「↶ 撤销本次执行」** 与那句"来源工作表「来源」保持你现在的修改；
   撤销不会重新执行计划"；
4. 点它 → 期望 `结果1!B2`、`结果2!B2` 变空（`结果3!B2` 本来就是空）、`来源!B5` 仍是 **500**；
5. 再确认卡片上不再有撤销按钮（只剩"已撤销本次执行…"），且没有重新执行计划（结果表里没有新公式）。

若第 3 步看不到按钮：把卡片上的原因原文抄下来 —— 那说明账本里的 `checkpoint` 与进度对不上
（或该执行已被处置），属于"如实阻断"，不是本功能失效；此时不要手工清空目标，把原因反馈回来。
