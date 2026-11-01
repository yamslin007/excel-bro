# Excel 恢复卡两处状态错误（来源复核 / 撤销终态，2026-10-31）

上一轮（`OFFICE_UNDO_AFTER_RESTART_2026-10-31.md`）修好了"重开后没有撤销入口"。本轮修同一张
恢复卡上**两处状态错误**：来源复核结论缺失，以及撤销成功后终态没进卡面结论。

真机现场（用户给定，**不重跑**）：

- 来源变化曾被续跑正确阻断（`来源!B5=500`）；但**重开之后**恢复卡却显示"可继续剩余动作"，
  正文还写"来源和目标已核对"。
- 撤销成功后，**同一张卡**同时显示"已撤销本次执行"（撤销行）与"可继续剩余动作"（结论 +
  按钮），且"继续剩余动作"按钮仍可点。

**当前真实现场（2026-10-31 收尾时点，已撤销完毕）**：`结果1/2/3!B2` 均已为空、`来源!B5=500`。
本轮**不再**要求恢复公式，也**不**要求把来源改回 5；最小真机检查只复用这个已撤销现场。

## 1. 两个根因（读代码核实）

### 根因一：显示判据从不读来源，正文却把"来源已核对"写死

`officeRecoveryVerdictFor`（`App.tsx`）只调 `pendingResumeReadiness`（身份 → 状态 → 授权 →
在途材料），**从不比对来源**，剩余动作 > 0 就直接返回 `resume`；而共享层
`executionRecoveryDetail(resume)` 的正文写死"原计划与当前授权、**来源**和目标已核对"。

于是：发现阶段没读来源 ⇒ 卡面说"来源已核对、可继续"；真实守卫一直在（`executePlan` 里
`startActionIndex > 0` 时 `verifyFormulaResumeSources`），但只在**点续跑之后**才触发 —— 显示与
守卫各说一套。这不是"已经发生过错误写入"，而是**卡面的声称不成立**。

### 根因二：撤销终态只写撤销行，不写卡面结论

撤销成功只 `setOfficeUndoDone`（`App.tsx`），它只影响 `officeRecoveryUndoFor`（撤销行 done），
**不参与** `officeRecoveryVerdictFor`（卡面结论仍按旧账本读数算 `resume`）。账本 `disposition`
又是异步留痕（`recordExecutionDisposition` best-effort），旧读数里没有。于是同卡
"已撤销本次执行"（撤销行）+"可继续剩余动作"（结论 + 按钮）并存，按钮还能点。

## 2. 改动

### 共享层（两端同一份呈现契约）

| 位置 | 改动 |
| --- | --- |
| `packages/core/src/executionRecoveryView.ts` | `ExecutionRecoveryVerdict` 新增终态 `{ kind: "undone" }`；`resume` 增加可选 `source: "matched" \| "not-required"`；新增 `ExecutionRecoverySourceVerdict`（`matched / not-required / unverified / refused`）；新增 `executionRecoveryResumeVerdict(remaining, source)`（**没复核来源就不出 resume**）；`executionRecoveryDetail(resume)` 只在 `source === "matched"` 才说"来源已核对"；`executionRecoveryActions(undone)` 返回 `[]`（一个按钮都不给）；新增 `EXECUTION_RECOVERY_STATE_LABEL.undone = "已撤销本次执行"` 与 `EXECUTION_RECOVERY_UNDONE_DETAIL`；`blocked` 增加可选 `sceneChanged?: true`（"写入过的目标已不在写后状态"是续跑**与撤销两个入口**共同的失效原因，宿主据此把撤销行一并收走） |
| `packages/core/src/executionResumeScene.ts`（新） | `verifyResumeScene`：续跑前核对"这次执行写入过的目标现在还是不是它留下的样子"（`no_evidence / unreadable / changed`）。挡的是"撤销成功但处置留痕上报失败"那一类账本与表不一致的现场 |
| `packages/core/src/index.ts` | 导出上述类型与函数 |

### Excel 适配层

| 位置 | 改动 |
| --- | --- |
| `App.tsx` | `officeSourceVerdictFor(state, {fresh})`：用**同一个** `verifyFormulaResumeSources`（与 `executePlan` 续跑前同一函数）读来源指纹，正面结论按场景键缓存、负面结论不缓存（来源改回基线后重算，不永久锁死）；`officeSceneVerdictFor(state)`：调 `observeOfficeExecutionPostState` + `verifyResumeScene`；`officeRecoveryEvidenceFor`：先现场后来源，只有"还有动作没跑"才读证据；`officeRecoveryVerdictFor(state, evidence)` 改为**终态优先**（本机 ref/state 或账本 `undoneAt` ⇒ `undone`），剩余动作走 `executionRecoveryResumeVerdict` |
| `App.tsx` | `officeRecoveryCardFor` / `officeRecoveryNoticeFor` 增加**必填** `evidence` 参数（所有调用点补上；出错同步分支用 `OFFICE_RECOVERY_NO_EVIDENCE`）；`officeUndoTerminalRef`（撤销成功的**同步**终态镜像，晚一拍的重渲染前也能读到）；`undoOfficeRecoveredExecution` 撤销成功即写 ref + state，并把卡面结论就地改成终态 |
| `App.tsx` | 四个处理器（`reconcileOfficeInFlightAction` / `resumeOfficeRemainingActions` / `backfillOfficeCompletion` / `undoOfficeRecoveredExecution`）在动作前按**最新**账本读数 + 本机终态复核（`officeExecutionTerminalReason`），已撤销 ⇒ 一格不写、不重置游标；`resumeOfficeRemainingActions` 点击时 `officeRecoveryEvidenceFor(state, {fresh:true})` 永远现读来源与现场 |
| `App.tsx` | `officeRecoveryUndoFor` 增加 `sceneChanged` 检查：`entry.card.verdict` 为 `blocked` 且 `sceneChanged` 时，撤销行也返回 `blocked`（"目标已不在写后状态"）—— 撤销成功但处置留痕上报失败后，重开靠写后现场核对发现目标已被撤销清空，既阻断续跑、也收走"再次撤销"入口 |
| `excel.ts` | 抽出 `undoRangePostStateIssue`（撤销前核对与续跑前现场核对共用同一份"哪一代摘要优先"的判据）；新增 `observeOfficeExecutionPostState`（只读逐目标观测） |

**没有做的事**：没有新增执行前确认；没有改变"计划到达即执行一次、回读后审阅"；没有用现读来源
替换登记基线；没有放宽授权或来源阻断；来源变化时续跑的拒绝口径（`FORMULA_SOURCE_CHANGED_MESSAGE`）
一字未改。

## 3. 自动化证据

- `packages/core/src/executionRecoveryView.test.ts`：新增"续跑结论必须被来源复核支撑"5 条（没复核
  不出 resume、matched 才提来源、not-required 不冒称、refused 说清"改回再核对"、旧调用点不带
  source 不被读成"已核对"）＋"撤销是终态"4 条（无动作、badge、正文三件事、进度行口径不变）。
- `packages/core/src/executionResumeScene.test.ts`（新）7 条：没写动作 / 全匹配 / changed / no_evidence /
  unreadable / 三原因分开报 / 坏进度。
- `excelWriteEndToEnd.test.tsx`：
  - 增强「中断 → 核对 → 改来源 → 重开 → 撤销」：重开后断言**没有**"继续剩余动作"、**没有**
    "来源和目标已核对"、有 `来源已变化`；撤销后断言卡面 `data-recovery-kind="undone"`、无继续/补记/
    核对按钮、无"可继续剩余动作"。
  - 新增「来源恢复登记基线后点重新核对（只读）不永久锁死」：改 `来源!B5=500` 重开 ⇒ 阻断 + 只读
    重核；改回 5 点重核 ⇒ "继续剩余动作"回来，零写入。
  - 新增「撤销成功后迟到的发现/重核不把终态覆盖回可继续」：撤销后卡面仍是已撤销、无任何写入按钮、
    再渲染零写入。
  - 修复既有「A1 三动作中途断开（applied）」的替身 mock：`reconcileOfficeActionReceipt` 现在传播
    `request.checkpoint`（真实服务端会把核对累积的写后指纹存进账本），否则"已核对"后账本缺第 1 步
    写后指纹会被 `verifyResumeScene` 误判为 `no_evidence`。
  - 回归⑥（`excelWriteEndToEnd.test.tsx` 描述块「Excel 未完成执行：重开后从账本材料撤销已写入的内容」
    内）：`撤销成功但留痕上报失败 → 本会话保持已撤销且不重复宿主操作；重开后无留痕也不重新开放续跑`。
    触发：`reconcileInFlight` 后 mock `recordExecutionDisposition` 拒绝一次 → 点「撤销」→ 断言留痕只
    `undone` 一次、宿主恢复只写回两个目标格（`结果1!B2`、`结果2!B2`）、无新增登记/领取/执行 → 卸载
    重挂（`listPendingOfficeExecutions` 返回无 `undoneAt` 的候选）→ 断言重开后无「继续剩余动作」「补记」
    「撤销本次执行」按钮、零新写入。判别点：`officeRecoveryUndoFor` 的 `sceneChanged` 检查收走撤销行。
  - 回归⑧（同文件）：`卡片生成后账本处置变成「已撤销」，点原「继续剩余动作」入口 → 处理器按最新处置
    拒绝，零写入`。触发：保存「继续剩余动作」按钮元素（= 旧回调持有者）→ 在点击前把账本
    `disposition.undoneAt` 推进成已撤销 → 点原按钮 → 断言 `claimOfficeExecutionResume` 不增加、零写入、
    卡面收敛 `undone`。判别点：`resumeOfficeRemainingActions` 里点击时的 `officeExecutionTerminalReason`
    最新处置闸（S5 变异钉的正是它）。

## 4. 反向变异（`scripts/mutate_office_recovery_card.mjs`，四判据全部成立）

| 变异 | 判别点 | 结果 |
| --- | --- | --- |
| S1 共享层把 `unverified` 当可继续 | 没复核来源就说"可继续" | 红：`expected 'resume' to be 'blocked'` |
| S2 撤销成功后不重写卡面结论 | 终态被覆盖回 resume | 红：卡面残留 `data-recovery-action="resume"` |
| S3 撤销成功不写同步终态 ref | 卡面结论退化成非 undone | 红：`data-recovery-kind="undone"` 找不到 |
| S4 现场已变却不标 `sceneChanged` | ⑥ 的撤销行守卫失效 | 红：重开后撤销按钮 `data-recovery-action="undo"` 又出现 |
| S5 短路续跑处理器的「最新处置闸」 | ⑧ 的点击守卫失效 | 红：旧回调绕过终态，续跑被放行（claim 增加 / 写入发生） |

每个变异记基线 sha256、失败原文、还原后逐字节一致、还原后复跑绿；脚本结束时工作区干净。

**关于 S3 与 ⑧ 的边界**：S3 钉的是「撤销成功后同一 tick 同步终态 ref」这一件事（错误 #2 的卡面收敛），
它**不是** ⑧ 的证据。⑧ 的守卫是点击处理器的「按最新账本处置拒绝」，判别点是 S5 —— 删掉 S5 那处
`if (terminal)` 闸，回归⑧ 立即红（旧回调不再被最新处置挡住），而 S3 是否在都不影响 ⑧ 的结果。

## 5. 两端差异（只报告事实，不宣称一致）

- **来源复核**：WPS 端**没有这个 bug**。它的 `wpsReconcileVerdict` 的 `resumeAvailable` 来自
  `canResumeWpsExecution` → `prepareWpsResume`，而 `prepareWpsResume` 内部已经
  `assertWpsResumeSources` + `assertWpsFormulaSourceCoverage`（同步读现场）。所以 WPS 给出 `resume`
  时来源**已经核对过**；给 `replan` 时明确说"材料缺失 / 过期 / 来源不符"。Excel 是"出结论时才读
  来源"（本轮补上），WPS 是"结论里已经含来源结论"。

### Excel 正面来源缓存的有效范围与失效条件（核实结论）

`officeSourceVerdictFor` 的场景键含 `sourceDependencyFingerprint`（**登记时的基线指纹**，不是现场
指纹），所以来源现场变化**不会**改变键，已缓存的 `matched` 在发现阶段**不会自动失效**——卡面更新
（自动发现 / 焦点变化，不带 `fresh`）可能短暂显示过时的 `matched`。但这**不会**造成"永久复用 matched"
或真正执行，因为：

1. **`refused` 从不缓存**（`officeSourceVerdictRef.current = null`）：一旦某次现读判出"来源已变化"，
   不会留一份旧的 `matched` 在缓存里被反复命中；
2. **真正要写入的两个入口都强制 `fresh: true`**：`resumeOfficeRemainingActions`（点「继续剩余动作」）
   和 `reverifySelectedOfficeRecovery`（点「重新核对（只读）」）都调
   `officeRecoveryEvidenceFor(state, { fresh: true })`，绕过缓存现读现场；
3. `verifyFormulaResumeSources` 比的是**现场指纹 vs 登记基线**（`plan.sourceDependencyFingerprint`），
   这个基线固定、不随缓存变化，所以"matched → 来源又变 → 再变回"每一步都按当时的现场重新判定。

结论：**卡面更新**可能复用正面缓存（省去每次焦点变化重读整张来源表），但**点击续跑前的强制复核**
永不吃缓存，来源再次变化后不可能因为缓存里的 `matched` 而被放行写入。这正是"卡面更新"与"点击前
强制复核"两件事的边界：前者是显示快照，后者是执行判据。
- **撤销终态**：WPS 端**没有同类问题**，但准确结论不是"材料回收让 resume 失效"。核实后的调用链是：

  1. WPS 的撤销入口挂在**结果卡**（不在恢复卡），撤销成功走 `undoWpsExecution`；
  2. `undoWpsExecution` 把宿主缓存里这条记录的 `record.state` 改成 `"restored"`（`wpsPlanExecutor.ts`
     L2994 `record.state = "restored"; upsertRecord(...)`）；
  3. 之后任何 `resume` 判定都走 `canResumeWpsExecution` → `wpsResumeUnavailableReason` →
     `prepareWpsResume`，而 `prepareWpsResume` 的可续跑状态集合只有
     `["ready", "applying", "partial_write_recoverable"]`（L1423），`"restored"` **不在其中** ⇒ 抛
     "本机续跑材料缺失、过期、被替换或仍需核对；未执行" ⇒ 不可能再算出 `resume`。

  为什么**不能**把材料回收（`retireWpsRecoveryPackage`）当作判据：`ensureWpsRecoveryMaterial` 是
  **宿主缓存优先**（`wpsUndoRecordFor` 按 executionId 仍能找到那条 `restored` 记录，返回 `{ kind: "host" }`
  = "可用"），所以材料回收既**不必然**发生（`retireWpsRecoveryPackage` 是 best-effort、失败就吞掉）、
  也不足以单独保证阻断。真正兜底的是上面第 2、3 步的 `record.state` fail-closed —— 它对"缓存还在 /
  迟到的响应 / 旧的续跑回调"一视同仁，全部从同一个 `prepareWpsResume` 入口挡住，不依赖材料有没有删干净。

  测试支撑：`wpsExecutionResume.test.ts` 新增「撤销成功后 record 变 restored，续跑在拿到执行权前就被
  state 闸拒绝、零写入」——撤销后断言 `record.state === "restored"`、`ensureWpsRecoveryMaterial` 仍是
  `host`（证明材料回收不是判据）、`wpsResumeUnavailableReason` 带"本机续跑材料缺失/过期/被替换"、
  `canResumeWpsExecution === false`、`resumeWpsExecution` 抛错且 `claimWpsExecutionResume` 未被调用、
  撤销后无新增写入。
- WPS 侧本轮**只核实、不改代码**；共享层 `ExecutionRecoveryVerdict` 新增的 `undone` kind 对 WPS
  无影响（WPS 的 `WpsReconcileVerdict` 不会产出 `undone`，且无对 verdict 的穷尽 switch）。

## 6. 仍未闭环

- 真机未验收：全部结论是"真实 App 组件 + 真实服务端契约 + 替身宿主"的自动化证据。
- 回归⑥（处置上报失败时本会话保持已撤销、不重放宿主操作、重开不重新开放续跑）与⑧（持有旧回调
  也不能绕过最新处置状态触发写入）已在自动化里分别由 S4、S5 的反向变异钉住判别力，但**真实 Excel
  上的"旧回调 / 连点 / 重开"未实测**。
- 未提交、未推送；A1-GATE 不自动通过，A2 未启动。

## 7. 最短真机验证（只跑这一遍，复用**已撤销**现场）

当前现场 `公式来源证据验收-Excel-B.xlsx` **已经撤销完毕**：`结果1/2/3!B2` 均为空、`来源!B5=500`。
本轮**不恢复公式、不把来源改回 5、不操作用户工作簿、不动真实账本**，只复用这个现场做一次只读检查：

1. 用新版窗格打开该文件，等恢复卡出现（或触发一次只读重核）；
2. **期望**：重开后**不**显示「继续剩余动作」「补记完成（不重写原表）」「核对在途动作」，也**不**
   显示第二个「↶ 撤销本次执行」入口 —— 只有"已撤销本次执行"终态说明，或按 §1 的阻断口径说明
   "写入过的目标已不在执行后状态"；
3. 若重开后仍出现「继续剩余动作」或「撤销本次执行」按钮，把卡片原因原文与账本处置痕（是否带
   `undoneAt`）一起抄下反馈；**不要**点任何按钮、**不要**手工清空目标或改来源。

这条检查**不**验证"撤销后留痕缺失仍被现场核对阻断"——那需要一个新的、还没撤销的现场才能测，
本轮明确不安排重跑。

## 8. 本轮收尾（2026-10-31，三项交付）

### 8.1 文案修正：已撤销卡不再显示续跑引导

`EXECUTION_RECOVERY_NOTE`（"恢复只使用当前已授权工作表；续跑通过复核后执行剩余动作，再审阅结果。"）
是**写死**的一句续跑引导，两端都无条件渲染，于是已撤销的终态卡上"已撤销本次执行"下面还跟着这句续跑
引导，与终态矛盾。

修法（**不改任何续跑/撤销/授权/来源校验判据**，只派生**呈现说明**）：

- `packages/core/src/executionRecoveryView.ts`：新增 `executionRecoveryNote(verdict)` —— `undone` 返回
  终态说明 `EXECUTION_RECOVERY_UNDONE_NOTE`（"这次执行已撤销，不会再提供补记或审阅入口；来源工作表
  保持你现在的修改。"），其余结论原句照旧。
- Excel `OfficeRecoveryPanel.tsx` / WPS `WpsAgentPanel.tsx`：渲染点从 `EXECUTION_RECOVERY_NOTE` 换成
  `executionRecoveryNote(verdict)`。两端同一份共享函数，措辞不各写一套。

定向测试：

- core `executionRecoveryView.test.ts`：新增「卡片边界说明按结论派生」2 条（undone 不出现续跑引导/
  审阅措辞、其余结论原句照旧）。
- `excelWriteEndToEnd.test.tsx`：新增「已撤销终态不出现续跑引导句，也不给任何写入按钮（说明按状态
  派生）」——撤销前仍有续跑引导、撤销后不出现该句且无继续/补记/撤销按钮。

### 8.2 只读核验保存文件

文件 `C:\Users\jia\Desktop\测试\公式来源证据验收-Excel-B.xlsx`（**只读，未保存、未改写**）：

- 绝对路径：`C:\Users\jia\Desktop\测试\公式来源证据验收-Excel-B.xlsx`
- mtime：`2026-10-31 13:32:26.413 +0800`
- 大小：`11443` 字节
- SHA256：`b480a362972af635ce7cab4372c3f8aa89e9fcb12f6de89ffc9b16b07a62b760`
- 工作表：`来源`、`结果1`、`结果2`、`结果3`
- 单元格读数（`data_only=False`，含公式视图，均无公式）：
  - `结果1!B2` = `None`（无值、无公式）✓
  - `结果2!B2` = `None`（无值、无公式）✓
  - `结果3!B2` = `None`（无值、无公式）✓
  - `来源!B5` = `500` ✓
  - （`来源!B2` = `2`，是来源表既有输入格，非本次核验对象）

磁盘读数与用户所述"结果1/结果2 B2 撤销后为空、来源 B5=500"一致，三张结果表 B2 均无值无公式。

### 8.3 归档与验证

**三类证据分开记**：

- **截图事实**（用户此前展示，本轮不重跑、不复核）：重开后恢复卡显示"已撤销本次执行"，继续/补记/
  重复撤销按钮均消失、只剩「重新核对（只读）」，来源 B5 仍为 500。
- **磁盘读数**（本节 8.2，本轮只读核验）：结果1/2/3!B2 均空、来源!B5=500。
- **自动化证据**（本轮）：core `executionRecoveryView` 25 条、`excelWriteEndToEnd` 97 条、
  WPS `WpsAgentPanel`+`wpsExecutionResume` 286+5skip 全绿；三端 `tsc --noEmit` 0；两端 build 通过。

**未验部分继续单列**：WPS 端撤销终态没有同类 bug（撤销走结果卡 + `record.state="restored"` 闸，见 §5），
但**真实 WPS 未验收**；不把 Excel 通过扩大为双端通过。A1-GATE 不自动通过，A2 不启动。未提交、未推送、
未重启服务，未获取新授权。
