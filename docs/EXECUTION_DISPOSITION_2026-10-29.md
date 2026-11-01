# 用户处置留痕：放弃 / 撤销必须在服务端留下一行（U4，2026-10-29）

> 这份文件只讲一件事：**"用户放弃了这次执行"和"用户撤销了这次执行"** 从今天起有服务端记录。
> 在此之前，账本里只有进度（`status` / `next_action_index` / `in_flight_action_index`），
> 用户点了「放弃变更」「确认变更」「撤销本次执行」，`office_executions` **一个字都不变**。

## 1. 为什么必须补（真机反证）

2026-10-29 Excel 真机一轮（`execution-c52ad791-6eb8-47c7-99d0-69132b70ef25`）：三步写入 →
第二步后中断 → 约 2.5 小时后只续第三步 → 确认变更。账本这一侧完全自洽
（`status=completed`、`revision=11`、`next_action_index=3`、`in_flight=NULL`、
`office_host_writes` 三行时间对得上，`verification` 三格逐格一致），全部证据见
`outputs/a1-acceptance-20261029-2/`。

但"用户点了确认"这件事 —— **查不到**。实测该行里 `[k for k in row if "abandon" in k or "undone" in k] == []`。
于是"这次执行是被用户主动收掉的，还是根本没跑完"只能靠三方推断：

1. 恢复材料消失了（WPS 才有材料；Excel 侧本来就没有）；
2. `revision` 没有推进；
3. 用户口述 / 截图。

**推断不是记录。** 这就是 `docs/A1_GATE_REMAINING_2026-10-28.md` 的 U4（尚未实现）与
E3（证据待补）是同一条根。

## 2. 做了什么（最小改动，双端同时）

### 2.1 账本

`server/app/task_ledger.py`：

| 项 | 内容 |
| --- | --- |
| 新列 | `abandoned_at REAL`、`undone_at REAL`（`user_version` 11 → **12**，`_REQUIRED_OFFICE_COLUMNS` 同步，旧库走 `ALTER TABLE` 补列） |
| 新方法 | `record_office_disposition(operation_id, disposition) -> (是否存在, 首次时刻)` |
| 只读方法 | `office_disposition(operation_id) -> {"abandonedAt", "undoneAt"} \| None` |

三条语义，全部钉在代码注释里：

- **首次写赢**（`WHERE ... AND <col> IS NULL`）：这一列回答"第一次发生在什么时候"。
  重试、双端各报一次、用户再点一次，都不改写那一刻 —— 与 `reserve_recovery_package_expiry`
  同一口径。
- **不推进 `revision`、不改 `status`、不动游标、不删材料**：处置是"用户做过什么"，不是
  "任务走到哪了"。写进进度会让续跑判据读错。
- **记录不存在就 `(False, None)`**：绝不为了留痕补建一条账本。

`NULL` 的唯一含义是"**没有这条留痕**"：本功能上线前发生过的放弃/撤销、以及从未做过处置的
执行，看起来完全一样。**不许把 `None` 说成"用户没放弃过"。**

### 2.2 端点

| 宿主 | 路径 | 说明 |
| --- | --- | --- |
| Excel | `POST /api/office/executions/{id}/disposition` | 身份必须以 `url:` 开头 |
| WPS | `POST /api/wps/executions/{id}/disposition` | 身份必须以 `wps-path:` 开头 |

体 `{"disposition": "abandoned" | "undone"}`，成功 `204`。两者共用一个实现
（`_record_execution_disposition`），只有前缀不同：**记录不存在或前缀不符一律
404 `EXECUTION_DISPOSITION_NOT_FOUND`** —— 留痕不是跨宿主探测工具，也不许"为了留个痕"
补建账本。非法取值 `422`，不静默归一到某个默认值。

**协议同步（三处同一份取值，一处漂移就会红）**：

| 位置 | 内容 |
| --- | --- |
| `server/app/models.py` | `OfficeExecutionDispositionRequest` / `WpsExecutionDispositionRequest`，均为 `Literal["abandoned", "undone"]` |
| `apps/excel-addin/src/contracts.ts` | **只转发不另写**：从 `@excel-bro/core/executionDisposition` 重新导出类型与标签，并在注释里写明端点与 404/422 语义（与"未完成执行发现"同一口径） |
| `packages/core/src/executionDisposition.ts` | 唯一权威：TS 联合类型 + 归一键 |
| 端点行为 | `test_both_endpoints_accept_both_values` 钉住"两个端点都收这两个值"；`test_invalid_disposition_is_rejected_instead_of_defaulted` 钉住"第三个值一律 422" |

### 2.5 失败与重试：没有"重试"这个功能，只有"重发这一条 POST"这条路

- **记录时机**：只有在**宿主恢复成功之后**才上报 —— Excel 是 `useUndoSnapshot` 的
  `onRestored`（它在 `await undoExecution(snapshot)` 成功之后才被调用，失败分支只会
  `onMessage` 报错），WPS 是 `await undoWpsExecution(...)` 返回之后。
- **失败不回滚结论**：`recordExecutionDisposition` 永不抛错，失败返回 `"failed"`；
  界面文案、单元格状态、审阅标记清理一律不受影响，也**不弹错误**。
- **不自动重试**：上报只发一次。要重试只能重发这一条 POST —— 它不碰宿主，因此
  不可能"再撤一次"，也不可能触发任何写入。
- **重试幂等**：服务端首次写赢（`WHERE <col> IS NULL`），重复上报拿回的是第一次的时刻。

#### 证据口径（2026-10-30 收紧：别把三件事说成一件）

"上报重试不重复撤销"这句话以前被写成一格 ✅，实际它拆开是三层，**只有前两层有客户端用例**：

| 层 | 事实 | 谁证明 | 措辞边界 |
| --- | --- | --- | --- |
| ① 失败不自动重试 | 上报抛错后端点只被调用 **1 次**，宿主写入次数不再增加，界面不弹错 | Excel / WPS 各一条组件用例（`mockRejectedValueOnce` + `toHaveBeenCalledTimes(1)`） | 证的是"**没有**自动重试"，不是"重试是安全的" |
| ② 成功后入口消费 | 栈条目已弹出、结果卡进终态，界面**不存在**第二次撤销的入口 | 同上两条用例的尾部断言（`queryByRole(...)` 为 null） | 证的是"界面上没有第二次撤销的路径"，不是"服务端拒绝了第二次撤销" |
| ③ 重复上报幂等 | 同一 `operationId` 再报一次不改写第一次的时刻，两个取值各写各的列 | 仅服务端 `test_first_write_wins_and_the_two_kinds_are_independent`（直接调账本方法两次） | **没有**客户端用例 —— 因为客户端根本没有重发路径可点 |

所以准确的说法是：**"重试"在这个产品里不是功能，只是"重发同一 POST"这一条窄路；
客户端没有重试实现，因此"用户重试"这个 UI 动作测不出来（它不存在）。**
上面三层合起来说明的只有一件事：即使有人（现在的或以后的代码）重发同一 POST，
也不会二次撤销、也不会把留痕时刻往后改。**别把它写成"重试路径已验收"。**

### 2.6 schema 迁移

`user_version` **11 → 12**；新库在 `CREATE TABLE` 里带这两列，旧库走
`ALTER TABLE office_executions ADD COLUMN {abandoned_at|undone_at} REAL`。
两列同时进 `_REQUIRED_OFFICE_COLUMNS` —— 版本号只是意图，**列才是事实**：
`_schema_is_current` 逐列确认，缺一列就会重跑 bootstrap（2026-10-27 的 500 就是这么来的）。

`test_legacy_ledger_is_migrated_with_the_two_disposition_columns` 把两列**真的删掉**、
版本号写回 11，再用新实例打开：补列成功、`integrity_check = ok`、
且执行进度（`revision` / `status`）一点没变。

### 2.3 共享判据（两端读同一份）

`packages/core/src/executionDisposition.ts`：

- `ExecutionDisposition = "abandoned" | "undone"`，`EXECUTION_DISPOSITION_LABEL` 与审阅卡
  按钮同名（「放弃变更」/「撤销本次执行」），不另起一套叫法；
- `normalizeExecutionDisposition` / `executionDispositionTarget`（只去首尾空白，**不做**
  大小写或前缀归一 —— 账本按完整字符串比身份）；
- `recordExecutionDisposition(send, id, kind)` **永不抛错**，返回
  `"recorded" | "skipped" | "failed"`：没有执行身份就 `skipped`（不去编一个 id），
  端点失败就 `failed`。与 `office_host_writes` 同一套 best-effort 口径 —— 恢复已经发生、
  单元格已经回退，留痕只是观测，**不许**把用户做完的动作改判成失败。

### 2.4 接线（各宿主只加一处）

| 宿主 | 位置 | 判定 |
| --- | --- | --- |
| Excel | `App.tsx` 的 `useUndoSnapshot({ onRestored })` | `resultReview.discard()` 唯一的恢复入口 `restore:` 先写 `revertedByDiscard.current = executionId`；`onRestored` 里 `revertedByDiscard.current === snapshot.executionId` ⇒ `abandoned`，否则 `undone` |
| WPS | `WpsAgentPanel.tsx` 的 `undoExecution` 成功分支 | 直接用已有的 `origin: "discard" \| "undo"` ⇒ `abandoned` / `undone` |

- Excel 用 **ref** 记"谁发起的恢复"，不读 `resultReviewState.phase`：恢复是异步的，等它
  回来时状态可能已被 `settleExternally` 或失败分支改过（"要不要动手的闸门不许读闭包变量"）。
  恢复失败时把 ref 清掉，免得同一执行的另一次真正撤销被贴错标签。
- WPS 两条入口本来就共用同一个恢复实现、只有 `origin` 不同 —— 本轮**没有**新开分支，
  只是把那个已在用的区分写进账本。
- 上报都是 `void` + 内部吞错：React 侧不等它，界面文案一个都不改。

#### 2.4.1 Excel「确认之后撤销」的真实调用链（2026-10-29 核实，纠正 §4.3 的旧结论）

上一轮写"Excel 的 `undone` 不可达"是**错的**。核完代码 + 补了断言之后，这条路径不但存在，
而且已经有用例钉住（`excelWriteEndToEnd.test.tsx`「确认写入 → 界面给出显式撤销范围 →
撤销后逐单元格回到执行前状态」里那段 `undoKind === "undone"`）。

| 环节 | 位置 | 事实 |
| --- | --- | --- |
| 入口 | `App.tsx` 结果卡 `.execution-recovery-footer` 的「↶ 撤销本次执行」 | 按 `message.executedExecutionId ?? message.executedPlanId` 在 `undoStack` 里定位；**这就是值写入执行的撤销入口** |
| 控制器 | `useUndoSnapshot.undoExecutionTo(idOrPlanId)` | 找到索引后 `undoStack.slice(index).reverse()` **从该条到栈顶逆序**撤销 —— 底部那条会连带撤销"之后 N 次"，文案随之变成「↶ 撤销本次执行（含之后 N 次）」 |
| 逐条恢复 | `await undoExecution(snapshot)`（`excel.ts`） | 先 `assertUndoPostState` 在**首个写操作之前**核完现场，冲突即只读失败 |
| 回调 | `onRestored(snapshot)` | **在 `await` 成功之后**才调用（`useUndoSnapshot.ts` 的 for 循环内），失败分支只 `onMessage` |
| 留痕 | `onRestored` 里的 `recordExecutionDisposition(...)` | `revertedByDiscard.current === snapshot.executionId ? "abandoned" : "undone"` |

**支持范围（都要满足，否则按钮不可点）：**

1. 该执行仍在 `undoStack` 里（同一会话、`capabilities.json` 的 `undo.maxEntries` 之内）——
   找不到条目时渲染的是「当前没有可用的撤销记录；这不表示执行已撤销。」，`undoExecutionTo`
   也返回 `false` 而不抛；
2. `postStateCoverage === "complete"`（或走文件夹撤销栈）—— 否则按钮变「↶ 撤销需核对」并
   `disabled`，提示"执行后状态未能完整核对，已停止自动覆盖"；
3. `busy === false`；
4. `targetGuard.check()` 通过（撤销前重新核一次当前目标）。

**身份绑定：** 唯一键是 `executionId`（旧消息没有它时兼容 `planId`）。`undoExecutionTo` 按它定位、
`onRestored` 拿它上报、`markWorkbookExecutionUndone` 也按它标记 —— 三处同一个字符串，
**绝不按"最近一次"或消息位置猜**。

> 顺带澄清一个容易看错的名字：`useUndoSnapshot` 还导出 `undoLastExecution()`
> （= `undoExecutionTo()` 不带参）。它**没有任何界面调用点**（`App.tsx` 只是解构下来没调用，
> 单测直接拿 hook 调它），别把它当成"顶部撤销入口"去找。

**为什么"一个执行只有一个可点入口"是设计而非缺陷：** 顶部 header 的「备份与恢复」
（`aria-label="备份与恢复"`，徽标 = `undoStack.length + folderUndoStack.length`）打开的是
`FolderBackupManager`，里面 `onUndo` 只接**文件夹**撤销（`undoFolderExecutionTo`），另有一段
删表恢复点（`WorksheetCheckpointRecoverySection`），值写入那一格只写了一句
「请在对应的执行结果消息中撤销；这里只汇总当前会话能力。」。**它不撤销值写入执行** ——
这是刻意的：`selectExecutionRecoveryEntry` 保证"全页只有一个可点入口"，在抽屉里再开一个
就会让同一执行出现两处入口（WPS 是另一种取舍：抽屉里保留了卡片同名快捷入口，两处共用
`executionId`，所以留痕只有一条）。两条路径都写 `undone`，不再有"某端不可达"这种说法。


## 3. 判别力（反向变异，四判据）

`scripts/mutate_execution_disposition.py`，三个变异各打一个判别点：

| # | 拆掉什么 | 预期变红 | 判别点 |
| --- | --- | --- | --- |
| D1 | `AND <col> IS NULL` ⇒ 重复上报改写时刻 | `test_first_write_wins_and_the_two_kinds_are_independent` | "第一次发生"是不是真的第一次 |
| D2 | 让处置顺手 `revision + 1` | `test_office_abandoned_is_recorded_and_progress_is_untouched` | 留痕有没有混进进度 |
| D3 | 两种取值映射到同一列 | `test_wps_undone_is_recorded_on_its_own_namespace` | 放弃与撤销会不会互相冒充 |

四判据 = 退出码非 0 / 预期用例名在 FAILED 里 / 失败原文带自家证据 / 还原后逐字节 sha256
一致且复跑全绿。**只卡"退出码非 0"不算**（导入写错也会非 0）。

端上还有三个判别点（`scripts/mutate_excel_disposition_wiring.mjs`，一次一个 SEED）：

| # | 拆掉什么 | 预期变红 | 判别点 |
| --- | --- | --- | --- |
| E1 | Excel `onRestored` 的判定写死成 `"abandoned"` | 「确认写入 → 界面给出显式撤销范围 → 撤销后逐单元格回到执行前状态」 | 确认之后撤销是 `undone`，不许被记成放弃（同一条断言的失败原文：`'abandoned' to be 'undone'`） |
| E2 | 把 `onRestored` 挪到 `await undoExecution(snapshot)` **之前** | 「执行后目标被人工修改，确认与放弃都不覆盖该修改」 | 恢复失败时不许留痕（负向断言 `.not.toHaveBeenCalled()`） |
| W3 | WPS 的 `origin === "discard" ? "abandoned" : "undone"` 写死成 `"abandoned"` | 「撤销本次执行 → 服务端留痕记「撤销」，不是「放弃」」 | 双端同一条判据，两端各一个判别点，不许"一端绿就算两端都对" |

三个都跑成四判据全成立（含 ④b：还原后**复跑**那条用例必须绿 —— 证明红来自变异本身）。
WPS 侧「撤销被指纹复核挡下 ⇒ 不留痕」那条**没有**单独做变异：它的 seed 要跨
`undoExecution` 里三十来行的 try 块搬运，做不到唯一定点；这一格的判别力由
E2 在兄弟宿主上钉的是同一条原理（`onRestored` 只在 `await` 成功之后）+ WPS 侧的可观测断言。

## 4. 边界（本功能不做什么，别读错）

1. **不追溯历史**：这一轮（`execution-c52ad791…`）已经发生的确认动作**永远不会有**处置
   时间戳。E3 对那一轮只能停在"推断"。新记录从下一次放弃/撤销开始。
2. **不记录"确认变更"**。确认 = 清审阅标记，不含恢复，也没有"撤回"这件事；
   本功能的两种取值都不是它。谁要记"确认过"，那是另一件事（`authorizationRevision` 也不是它）。
3. ~~**Excel 的 `undone` 目前不可达**~~ —— **这条 2026-10-29 已证伪，见 §2.4.1**。
   旧结论把结果卡上那颗「↶ 撤销本次执行」错认成"只属于删表检查点"，实际它是**值写入执行**
   的撤销入口（删表检查点是另一套：`WorksheetCheckpointCard` → `restoreWorksheetCheckpoint`
   原地撤销删表，**不经** `useUndoSnapshot`，也不写这一列 —— 它属于结构变更，不是值写入）。
   现状：Excel 两端**都有**真实入口 —— 结果卡「撤销本次执行」写 `undone`，审阅阶段
   「放弃变更」写 `abandoned`；WPS 结果卡 / 备份抽屉两个入口共用同一个 `executionId`。
   四个组合都有用例（Excel 2 条、WPS 3 条），不再是"预留的默认值"。
4. **不代表能查到"用户是否没放弃"**：`NULL` 只是"没有留痕"。
5. **不动发现链路**：处置不参与 `pending_office_executions` 的排序与可见性。放弃过的执行
   如果 `status` 仍是 `needs_reconciliation`，它照样会被发现 —— 那是"要不要清提示"的事，
   由 `discovery_dismissed_at` 管，两列互不冒充。

## 5. 验证

| 层 | 命令 | 结果 |
| --- | --- | --- |
| 服务端 | `python -m pytest server/tests/test_execution_disposition.py` | **9 passed** |
| 服务端全量 | `python -m pytest server/tests --basetemp=<OS 临时目录>` | **1922 passed / 27 skipped / 2 failed**（环境无 Docker：`execute_python_analysis_async` 返回 `unavailable`，与本次改动无关；改动前同一台机器同样这 2 条红） |
| 共享判据 | `npx vitest run src/executionDisposition.test.ts`（`packages/core`） | **6 passed**；core 全量 62 文件 / **917 passed** |
| Excel 接线 | 全量 `npx vitest run`（`apps/excel-addin`） | **122 文件 / 1594 passed / 9 skipped**；本功能 5 条：放弃留 `abandoned`、确认**不**留痕、上报失败：端点只 1 次、界面不留二次撤销入口（**没有**自动重试实现，见 §2.5 口径表）、**确认后撤销留 `undone`**、**恢复失败（冲突/放弃失败）两种取值都不写** |
| WPS 接线 | 全量 `npx vitest run`（`apps/wps-addin`） | **70 文件 / 1222 passed / 8 skipped**；本功能 5 条：撤销留 `undone`、放弃留 `abandoned`、确认不留痕、上报失败：端点只 1 次、入口消失（无自动重试实现）、**撤销被指纹复核挡下不留痕** |
| 类型 / 构建 / lint | `npx tsc --noEmit`（两端）、`npm run build:addin`、`npm run build:wps`、`npm run lint` | 两端 0 错；两个构建退出码 0；lint **0 error**（36 条既存 warning，与改动前同数） |
| 变异（账本） | `python scripts/mutate_execution_disposition.py D1/D2/D3 check` | **四判据全部成立**（基线 8→9 passed 复绿；还原后 sha256 `f1894f70…fcb5` 逐字节一致） |
| 变异（端上接线） | `node scripts/mutate_excel_disposition_wiring.mjs E1/E2/W3 check` | **三条四判据全部成立**（含 ④b 还原后复跑绿；基线 sha256 `App.tsx c82cec0e…`、`useUndoSnapshot.ts 57cd9019…`、`WpsAgentPanel.tsx 961af165…` 逐字节一致） |

**实机未验**（本轮明确不做的两格）：

1. 真实 Excel / 真实 WPS 里点一次放弃/撤销，然后看账本这两列有没有值 —— 没有任何真机证据
   （`discovered` 与 `disposition` 都只能在真实宿主里点一次才知道）；
2. Excel **真实窗格**里结果卡「↶ 撤销本次执行」的可点性只到"真实 App 组件 + 真实样式表 +
   替身宿主"，真实 Excel 上一次都没点过（与恢复入口位置那条同源，见
   `docs/A1_GATE_REMAINING_2026-10-28.md`）。
