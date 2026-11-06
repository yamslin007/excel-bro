# 审阅状态的持久化（2026-11-03）

> 缺陷：执行写完、回读通过，用户还没点「确认全部变更 / 放弃全部变更」时刷新前端或重开
> 任务窗格 —— 待审阅界面消失、输入框回到普通状态，甚至写「本次没有需要你处理的后续操作」，
> 而原表停在一个没人收尾的写入现场上。
>
> 本轮把它改成：**审阅状态是账本里的一个独立事实**，重开时按账本找回审阅，
> 恢复过程一格都不再写。

## 1. 经代码核实的根因（不是推测）

三条，逐条都能指到具体位置：

1. **"完成"的记录根本不在发现范围里。** `server/app/task_ledger.py` 的
   `pending_office_executions` 的 SQL 明确排除了 `completed`（只有
   `running / needs_reconciliation / failed`）。于是"写完了、等用户决定"的那次执行
   在启动发现里**从不出现** —— 界面只能得出"本次没有需要你处理的后续操作"。
2. **账本里没有"待审阅/已确认"这个事实。** 迁就旧结构只有 `status`、`abandoned_at`、
   `undone_at`：`completed` 说明动作跑完了，`abandoned_at`/`undone_at` 说明用户**已经**
   放弃或撤销过。中间那段"写完了但用户还没表态"没有任何列承载。
3. **审阅状态只活在前端内存里。** `ResultReview` 的 `phase`/`receipt` 由 `start()` 在
   执行结束时建立；刷新后 `start()` 不会再被触发。挂载时那段"认领并清理遗留审阅标记"
   （`claimWpsReviewMarks` / Excel 侧的标记认领）清的是**展示层标记**，
   它既不能证明用户确认过，也不该被当成审阅的替代品。

因此"补文案"没用：状态本身丢了。要么把状态持久化，要么每次都靠猜。

## 2. 持久状态：一个列，两个时刻，三种"验证证据"

服务端账本（`office_executions` / `wps_executions` 同构，schema `v14`）：

| 列 | 含义 |
| --- | --- |
| `review_state` | `pending` / `blocked` / `confirmed` / `discarded` / `undone`，或 `NULL`（无法判断） |
| `review_state_at` | 这个审阅状态最后一次变动的时刻（与执行进度、处置留痕各自独立） |

`review_state` 的初值在执行**完成回执**那一条 UPDATE 里一起写入
（`finish_office_execution(..., review_state=...)`）：**同一条语句**写"动作全跑完"与
"结果在等用户决定"，所以不存在"完成了但审阅状态还没落"的中断窗口。

初值只有三种来源，都是证据而非推断（`main._review_state_for_finish`）：

- 结果通过独立验证 ⇒ `pending`；
- 结果明确未通过验证 ⇒ `blocked`（要人看，但**不允许**被当成"确认过的结果"直接接受）；
- 验证证据缺失（`verified is None`）⇒ **不写**，留 `NULL`。

终态集合是 `{discarded, undone}`；`confirmed` 不是终态（确认之后还能撤销），
它只允许迁移到 `undone`。允许迁移的唯一出处是 `packages/core/src/reviewLifecycle.ts`
的 `reviewDecisionAllowed`，服务端 `task_ledger.REVIEW_TERMINAL` 与它同口径。

## 3. 恢复流程（两端同一份判据，各写各的宿主调用）

```
启动/挂载
  └─ 发现（POST/GET /api/{office|wps}/executions/pending）
       └─ 候选带 kind: "resume" | "review" + reviewState
            ├─ resume：原有的"继续剩余 N 步"路径，一个字没改
            └─ review：reviewRestoreVerdict(...)
                 ├─ settled        → 终态，刷新后不再出现
                 ├─ legacy_unknown → 老记录，**不推成待审阅、也不推成已确认**
                 ├─ blocked        → 说清哪件事不成立，保留记录
                 └─ restorable     → adoptCompleted(..., {restored: true})
```

`reviewRestoreVerdict` 的输入是四件可核验的事实：账本身份 vs 当前宿主身份、`status`、
`result_json` 在不在、结果是否 `verified === true`；再加一个**只影响"能不能放弃"**的
材料可读性。

两个刻意的"不作为"：

- **当前选区不是判据。** 刷新之后面板里的勾选是**当下**的选择，不是这次执行登记的那一组。
  拿它去拦恢复，等于每次刷新都因为"选区没恢复"把待审阅弄丢（本轮实测踩到过）。
  恢复一律按账本登记的范围核对与恢复；选区不同只落一句 `limitations.scope` 提示，
  既不扩大也不缩小原范围。
- **材料缺失不等于没有待审阅。** 读不到恢复材料只限制「放弃」与「双视图」，
  不影响"这条审阅还在"。混在一起会让材料丢了的现场被报成"没有待审阅"。

## 4. 怎么保证"恢复审阅"不会重放写入

不是靠"记得别调"，是靠**那条路上没有能写的东西**：

1. 恢复走的是 `ResultReview.adoptCompleted(key, before, receipt, operations, {restored:true})`。
   这个入口的 `operations` 里**没有** `prepare` / `execute` 两个回调 —— 计划提交与执行
   在这条路径上不存在，删掉调用点也删不掉这个事实。
2. 「确认全部变更」只做两件事：`validateResult`（只读核对现场）+ `settle`（把决定写进账本）。
   没有第二个写入点；`confirm()` 里也没有任何"顺手再写一次"的分支。
3. 「放弃全部变更」只走 `restore`，而 `restore` 用的是账本 `checkpoint` 里登记的写前恢复点，
   外加写后现场核对；材料不可读时它自己就拒绝，不退回"按当前单元格反推一个恢复点"。
4. 服务端不接受"确认"去改执行进度：`apply_office_review_decision` 只写
   `review_state` / `review_state_at`，不动 `next_action_index`、不动 `status`、不认领执行权。

验收用例把这三条钉成断言：重开前后 `startOfficeAction` / `completeOfficeAction` /
`finishOfficeExecution` 的调用次数相等，宿主写入次数相等，点完确认之后仍然相等
（Excel `excelWriteEndToEnd.test.tsx`、WPS `WpsAgentPanel.test.tsx` 各一条，且做过反向变异：
把 `review` 分支短路掉，两条用例立刻变红）。

## 5. 中断窗口、重复与并发

- **写不进审阅状态就不算确认。** 前端 `ResultReview` 的 `confirm()`/`discard()`/
  `settleExternally()` 在 `settle` 失败时把决定挂进 `decisionPending`，界面说
  「结果已接受，但审阅状态没有写进执行记录」，并提供重试；**不**谎称已落账。
- **幂等与冲突分开。** `apply_office_review_decision` 在一个 `BEGIN IMMEDIATE` 事务里
  做条件 UPDATE，返回 `applied` / `duplicate` / `conflict` / `missing`。
  重复提交同一个决定 ⇒ `duplicate`（200，不改 `review_state_at`）；
  已终态又收到别的决定 ⇒ `conflict`（409 `EXECUTION_REVIEW_CONFLICT`）；
  记录不存在或跨命名空间 ⇒ 404 `EXECUTION_REVIEW_NOT_FOUND`。
- **旧状态不许覆盖新状态。** 终态不接收任何新决定，所以两个窗格同时操作时，
  后到的旧决定不会把"已放弃"改回"待审阅"。
- **同一执行只恢复一次。** 恢复出来的是**同一个** `executionId`；
  `adoptCompleted` 按 key 覆盖而不是新增，`decisionPending` 的重试也复用同一条记录。

## 6. 历史记录兼容：`NULL` 不等于已确认

迁移（`_SCHEMA_VERSION = 14`）只加列、**不回填**。`review_state IS NULL` 的老记录：

- 不进发现候选（`pending_office_executions` 的 `review` 分支要求 `review_state IN
  ('pending','blocked')`）；
- `reviewRestoreVerdict` 返回 `legacy_unknown`，界面明说
  「这次执行是本能力上线前记录的，无法判断你是否已经处理过它的结果」；
- 既不显示审阅入口，也不声称用户处理过 —— 那是同一份"无法判断"的两种假装。

## 7. 两端各自改了什么

**共享判据（`packages/core`）**

- 新增 `reviewLifecycle.ts`：状态机、迁移规则、恢复核验结论、受阻原因文案。
- `pendingExecutionDiscovery.ts`：候选加 `kind` / `reviewState`（进 `CANDIDATE_KEYS`），
  准入闸做交叉一致性（`resume` 不许 `completed`；`review` 必须 `completed` 且**显式**审阅状态；
  缺性质标记一律拒绝），`review` 与在途动作同档排最前，行文案改成
  「N 个动作已写完，等待你确认或放弃」。
- `resultReview.ts`：`settle` 回调 + `decisionPending` + `retryDecision`；
  `adoptCompleted` 的 `operations` 允许不带 `prepare`/`execute`（新增 `{restored}` 选项，
  恢复时跳过那次会话内的 verified/undoable 自检 —— 那两件事此时由账本证据回答）。

**服务端（`server/app`）**

- `task_ledger.py`：`review_state` / `review_state_at` 两列；`finish_office_execution`
  同事务写入；`apply_office_review_decision`；`pending_office_executions` 扩为
  `status IN (...) OR (status='completed' AND review_state IN ('pending','blocked'))`。
- `models.py`：`PendingExecutionCandidate.kind/reviewState`、
  `OfficeExecutionReviewState`、`OfficeExecutionReviewRequest`、状态响应带 `review`、
  `WpsExecutionResult.verified: bool | None`。
- `main.py`：状态投影带 `review`、候选计算 `review_candidate`、
  `POST /api/office/executions/{id}/review` 与 `POST /api/wps/executions/{id}/review`。

**Excel（Office.js）**

- `contracts.ts` / `api.ts`：`review?` 字段、`sendExecutionReview(namespace, id, decision)`。
- `App.tsx`：`restoreOfficeResultReview(state, isCurrent)`（核验 → `adoptCompleted` →
  不重建颜色标记）；发现与自动摆卡走同一入口；`pendingRecovery` 条件加
  `resultReviewState.phase === "reviewing"`；`isCurrent` 只看工作簿身份。

**WPS（JSAPI）**

- `WpsAgentPanel.tsx`：`restoreWpsResultReview(state, isCurrent)`（`ensureWpsRecoveryMaterial`
  区分 `host` / `package` 为可用，其余只限制放弃；标记给空的 `rules: []`，`clearMarker` 不碰任何东西）；
  `adoptWpsRecoveryCandidate` 在 `completed + reviewNeedsAttention` 时先走恢复；
  `validateResult` 用 `inspectWpsConfirmedResult`，`unknown` 时明说"重开后本机核对记录不再保留，
  不自动确认"。
- `wpsPendingDiscovery.testHost.ts`：替身与服务端同语义（`review` 只认
  `completed + 显式 pending/blocked`）。

## 8. 验证到了哪一档（不得合并宣称）

| 面 | 状态 |
| --- | --- |
| 服务端账本与端点 | `pytest server/tests -q` ⇒ **2039 passed, 1 skipped**（新增 `test_execution_review.py` 13 条） |
| 共享判据 | `packages/core` ⇒ **1193 passed**（新增 `reviewLifecycle.test.ts` 13 条；`pendingExecutionDiscovery` +5、`resultReview` +8） |
| Excel 插件 | `npm run test:addin` ⇒ **1771 passed, 9 skipped**（新增"重开窗格审阅入口回来且一格不再写"） |
| WPS 插件 | `npm run test:wps` ⇒ **1304 passed, 8 skipped**（新增同名验收 + 4 条续跑段用例对齐） |
| 构建 | `npm run build:addin` / `npm run build:wps` 通过；两端 `tsc --noEmit` 干净；`python -m compileall -q server/app` 通过 |
| 反向变异 | 短路 WPS 的 `completed + reviewNeedsAttention` 分支 ⇒ 新验收用例变红（复现"已记录完成/没有剩余动作"的旧症状），还原后转绿 |
| **真实 Excel 窗格** | **未验** |
| **真实 WPS 窗格** | **未验** |
| 双窗格并发 | 只有服务端 CAS 与 409 冲突的自动化证据，**没有真实双窗格实机证据** |

## 9. 宿主限制与"未实现"（说清楚，不拿提示文案遮）

- **双视图重开后不可用。** 执行前快照只活在原窗格的内存里；重开时
  `reviewRestoreVerdict` 给 `limitations.comparison`，`showBefore` 直接拒绝。
  不是"忘了做"，是不拿别的快照顶替。
- **颜色标记两端不同，原因也不同。** Excel 能按计划重新加**条件格式**标记
  （`markExcelResult`：只加条件格式，不动单元格内容与值）；WPS 的标记集给空
  （`rules: []`，`clearMarker` 什么也不碰）——它那套本次标记公式只在上一次窗格内存里。
  两种都不是"标记等于审阅"：审阅仍在账本里，标记只是展示层收尾。
- **"保存为工具"资格不重建**（WPS）。它需要原始计划的完整语义，而账本只留动作描述
  （不存单元格值）。宁可不给这个资格，也不拿残缺材料固化一个工具。
- **openpyxl 文件夹模式没有审阅界面**：该模式不执行"先写后审"的宿主写入，审阅状态
  与恢复流程不适用（不是未实现）。
- **`blocked`（结果未通过独立验证）重开后仍不能确认**：判据在账本，
  重开不会把它升格成 `pending`。
