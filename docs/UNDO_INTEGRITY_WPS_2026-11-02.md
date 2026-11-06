# WPS 撤销完整性修复（2026-11-02）

> 一句话：真机上「撤销本次执行」**漏恢复了第二步**（`结果2!B2` 仍是执行后的 8），
> 同时**报了"已撤销并回读确认恢复成功"**、写了 `undone_at` 留痕、把本机恢复材料**提前回收**。
> 判据已修：撤销收尾必须回读**全部**应恢复目标，只有全部回到执行前才允许落留痕、回收材料、
> 说成功。自动化已通过；**真机未复测**（等 B-Retest，见第 6 节）。

## 1. 真机现场（只读取证，未改账本、材料与单元格）

| 事实 | 值 |
| --- | --- |
| 本轮执行 | `wps-execution-281e2090-e221-4590-a8bb-0e75bc111391`（11-02 13:30:03 建，13:31:41 最后更新） |
| 账本状态 | `status=failed`、`next_action_index=2`、`in_flight_action_index=NULL` |
| 检查点 | `{"sheetName":"结果2","range":"B2","type":"writeFormulas","afterFingerprint":"v2:26:a013f678",…}` |
| 宿主写入 | action0 `written` / `readback=verified`；action1 **`failed_after_write`**（`mutatingCalls=1`） |
| 处置留痕 | `undone_at = 13:33:21`、`recovery_package_removed_at = 13:33:21` |
| 现场 | `结果1!B2` 空、**`结果2!B2` 仍是 8**、`结果3!B2` 空、`来源!B5 = 500`（用户改的） |
| 材料 | `execution-recovery/281e2090…/` **已不存在**（撤销那一刻被回收） |

已排除：旧失败执行 `4998b190-…-3d5a0e326041`（11-01 22:38，`next=0`，从未写入）、
A 轮 `304e6eee`（completed）、`fe021515`（11-01 23:18，只 dismiss 过）。

**旧 B 文件保留故障现场**：不清空 `结果2!B2`、不改原账本、不动恢复材料。该执行已经记成
`undone` 且材料已回收，**没有任何自动入口能把它退回去**（新版本也只能"不再误报"，
不会凭空造出证据）。要回退只能人工清那一格，或用一次新的执行覆盖。

## 2. 失败点（三处连成一条链）

1. **进度只按游标算**（`apps/wps-addin/src/wpsRecoveryPackage.ts` 水合）：
   `appliedRows = actionIndex < nextActionIndex ? rowCount : 0`。重开时账本 `next=1 / in_flight=1`，
   于是"写了但完成回执没落"的第二步被判成 `appliedRows=0`。
2. **宿主缓存从不按最新账本重投影**（同文件 `ensureWpsRecoveryMaterial` 的 `existing` 分支）：
   只补 `resumeExpiresAt`。用户点「核对」后账本 `next` 1→2、检查点落库，缓存仍停在重开那一刻
   ⇒ 撤销里两处 `if (action.appliedRows === 0) continue` 继续跳过第二步。
3. **收尾只回读"自己恢复过的那几格"**（`apps/wps-addin/src/wpsPlanExecutor.ts` `undoWpsExecution`）：
   漏掉的那一格没人看，于是 `record.state = "restored"`、报"已撤销…并回读确认恢复成功"，
   面板随后 `retireWpsRecoveryPackage` + `recordExecutionDisposition("undone")`。
   另外成功分支没有清恢复卡，卡面一直挂着「需重新规划」。

## 3. 判据（唯一出处）

- **`apps/wps-addin/src/wpsAppliedActions.ts`（新增，纯函数，无 UI 依赖）**
  - `wpsCheckpointAfterFingerprint(checkpoint, target)`：账本 `checkpoint` 只带定位与指纹，
    归属按 **表名 + 矩形** 认，不按"离谁近"猜。
  - `wpsAppliedActionEvidence(facts, action)`：两条互不冒充的来源 ——
    完成回执已落的动作按登记的目标指纹；`inFlightActionIndex` 且检查点指向本动作的，
    用检查点记下的**真实**写后指纹。两条都不成立 ⇒ `appliedRows = 0`（**没有证据不猜**）。
  - `wpsUndoReadbackTargets(actions)`：撤销收尾要对齐的**全部**值动作（`writeValues` /
    `writeFormulas` / `clearRange`），与 `appliedRows` 无关；格式动作不在此列（它们由
    逐格格式快照自己核对）。
- **`wpsRecoveryPackage.ts`**：水合改用它；新增 `wpsRecordWithLedgerProgress(record, state)`，
  在 `existing` 分支按**最新账本**重投影，且**只补不撤**（执行期已记的"尝试写入"不许被抹回 0，
  否则一次"回执还没送到"的时序就会让撤销悄悄漏掉一格）。
- **`wpsPlanExecutor.ts`**：恢复完再做一次**全目标回读**；任何一块没回到写前、或读不出来
  ⇒ `state = "undo_conflict"` + 抛 `WpsUndoResidualError(regions, unreadable)`，
  **不置 `restored`、不回收材料、不落留痕、不报成功**。
- **`WpsAgentPanel.tsx`**（只 3 处）：导入错误类型；残留走单独文案（区域清单照抄执行器，
  不用 `inspectWpsUndoRecord` 重解释 —— 它只读 `appliedRows>0`，正是漏掉残留的那一侧）；
  撤销成功后若恢复卡属于这次执行 ⇒ `setWpsRecovery(null)`。

三条验收口径（不得合并宣称）：
1. **完整撤销回读**：看每个应恢复目标的实际单元格，不看撤销函数的返回值；
2. **来源修改保留**：`来源!B5` 的 500 一个字符都不动；
3. **成功后才落留痕与回收**：`undone_at` / `recovery_package_removed_at` 的存在必须以
   全目标回读通过为前提；失败时两者都不许出现，材料必须留着。

## 4. 补丁范围与依赖（本次回流）

新增 1：

- `apps/wps-addin/src/wpsAppliedActions.ts`

修改 5（均为按上下文定点应用，**没有整文件覆盖**）：

- `apps/wps-addin/src/wpsRecoveryPackage.ts`（本轮差异 100% 属于本修复）
- `apps/wps-addin/src/wpsPlanExecutor.ts`（同上）
- `apps/wps-addin/src/WpsAgentPanel.tsx`（**只取 3 处**；其余 14 处差异属新 UI，未回流）
- `apps/wps-addin/src/wpsExecutionResume.test.ts`（导入 + 1 个 `describe`／2 条用例）
- `apps/wps-addin/src/WpsAgentPanel.test.tsx`（导入 + 2 条用例；其余 20 处差异属新 UI，未回流）

**依赖**：不依赖新 UI。上面用到的 `wpsRecovery` / `setWpsRecovery` / `wpsRecoveryRef` /
`clearPendingWpsExecution` / `.wps-write-recovery` 复核卡都是合并前既有的；唯一新增依赖是
本批自带的 `wpsAppliedActions.ts` 与既有的 `parseExplicitA1Rect`。所以不需要"先拆最小功能修复"，
这批本身就是最小功能修复。

**未回流（新 UI 批次，另行安排）**：`packages/core/src/recoveryComposer.ts`、
`apps/excel-addin/src/components/RecoveryComposer.*`、`apps/excel-addin/src/App.tsx`、
`types/chat.ts`、`packages/core/src/executionRecoveryView.ts`、`processOutcome.ts`、
`WpsAgentPanel.tsx` 的 14 处输入区改造、以及 `apps/excel-addin/src/excelWriteEndToEnd.test.tsx`
里随 UI 一起新增的 IME 用例（含它自己的身份隔离修正）。

## 5. 验证

- 主树（`E:\codex\excel-bro`，HEAD `3a8dd82c` + 工作区既有未提交改动）：
  - `apps/wps-addin` 全量：**72 files / 1294 passed / 8 skipped**，0 失败；
  - `tsc --noEmit -p apps/wps-addin` 通过；`npm run build:wps`、`npm run build:addin` 通过；
  - 构建产物 `apps/wps-addin/dist/assets/index-Bll8oB8r.js` 内含本修复文案
    （`撤销没有把`、`无法回读核对`、`恢复资料已保留，可以重试`），可据此确认运行版本。
- 新增自动化 4 条：
  1. 四表同形状全链（第二步写入后、回执前中断 → 丢 `PluginStorage` 只从恢复包水合 →
     改 `来源!B5` → 拒绝续跑 → 撤销）：断言 `结果1!B2`、`结果2!B2` 的**公式与计算值都为空**，
     `结果3!B2` 没被碰，`来源!B5` 仍是 500；
  2. 缺证据那一步**不冒充恢复**：报 `WpsUndoResidualError`、`state="undo_conflict"`、
     材料保留、记录仍可撤销、残留单元格内容不被清空；
  3. 面板：残留时**不报"已撤销"、不标 `undone` 卡、不落留痕、不删材料**，并说清是哪一块；
  4. 面板：撤销成功后恢复卡必须消失（前提是真的恢复了）。
- 反向变异 4 次（关掉重投影 / 回读只覆盖已写入 / 撤销后不清卡 / 面板不认残留错误）
  **各只红 1 条**；还原后三个源文件 sha256 与变异前逐字节一致。
- **真机未复测**：本节不含任何真机结论。原 B 文件保持故障现场，不动。

## 6. B-Retest 独立夹具操作清单（本文只列步骤，**没有创建也没有执行**）

前置（一次性）：**新建**一份工作簿另存为 `B-Retest-2026-11-02.xlsx`（已保存身份才支持持久恢复），
建四张表 `来源` / `结果1` / `结果2` / `结果3`；`来源!B5` 填 `8`，三张结果表 `B2` 留空。
**不要复用旧的 B 文件**（那是故障现场，保留原样）。

1. WPS 打开该新文件，面板只勾这四张表（授权上限＝这四张）。
2. 让模型生成三动作公式计划：每张结果表 `B2` 写 `=来源!B5`。计划预览里应看到 3 个公式动作。
3. 打开「恢复验收（开发专用）」→「启用下一次第二步中断」。
4. 执行。预期：第二步写入后中断，结果卡显示「恢复资料已保留」，**不**出现"写入验证通过"。
5. 保存工作簿 → 完全退出 WPS → 重开同一文件 → 面板应自动发现待办并摆出恢复卡。
6. 把 `来源!B5` 改成 `500`（模拟用户自己的修改）。
7. 点「核对结果」。预期：账本 `next` 1→2、检查点落库；随后结论为「来源已变化」，
   续跑入口不可用，**零写入**。
8. 点「撤销本次执行」。逐条验收：
   - `结果1!B2`、`结果2!B2` 的**公式与计算值都为空**（不能"值空了公式还在"）；
   - `结果3!B2` 从未写过：保持空，且撤销没有往它写任何东西；
   - `来源!B5` = **500 保留**；
   - 只出现一次「已撤销…并回读确认恢复成功」，且发生在全目标回读通过之后；
   - 恢复卡消失；「备份与恢复」里该执行不再有撤销入口。
9. 事后只读核对（账本／磁盘）：
   - `office_executions` 该行 `undone_at` 有值、`recovery_package_removed_at` 有值，两者时间晚于撤销动作；
   - `execution-recovery/<operationId>/` 在撤销成功后才消失；
   - `office_host_writes` 中 action1 = `failed_after_write`、`mutatingCalls=1`。
10. 反向用例（安全处理）：同一夹具再来一次，但第 7 步**不做核对**直接撤销。预期：
    - 报「撤销未完成：撤销没有把「结果2!B2」恢复到执行前；恢复资料已保留，可以重试。」
    - **不出现**"已撤销并恢复"、结果卡不标已撤销、`recordExecutionDisposition` 不被调用；
    - 账本无 `undone_at`、材料目录仍在、`结果2!B2` 内容保持原样（不被清空）。

判定：第 8 步全部成立且第 10 步如实失败 ⇒ 本缺陷真机闭环；任何一条不成立都写"未完成"，
不得写"仅待实机"。

## 7. 运行版本与重载

- **更正一条先前判断**：合并后一度记为"开发服务未运行、未触发任何热更新"，这条是错的。
  实测 `127.0.0.1:3987` 有 `node … vite.js`（PID 28864）在监听，且有一个 WPS 窗格
  （`promecefpluginhost.exe`，PID 3460）保持 3 条 established 连接。先前判断失误的原因是
  用 `grep -E ':(3987|3988)\s'` 判监听——`\s` 在 ERE 里不是空白，命令恒无输出。
- 事实：该开发服务的 root 指向主树，合并写入后 Vite 已把本轮新模块纳入模块图，实测
  `GET /src/wpsAppliedActions.ts` = 200、`/src/wpsPlanExecutor.ts` 含「撤销没有把」、
  `/src/WpsAgentPanel.tsx` 含 `WpsUndoResidualError` ⇒ **新代码已经通过热更新进了那个窗格**，
  不是我主动重载，也没有重启服务。
- 影响面：热更新只换 UI/执行器代码，不写单元格、不动账本与恢复材料（旧 B 现场仍是
  `status=failed`、`next=2`、action1 `failed_after_write`、`undone_at`/`removed_at` 均 13:33:21、
  材料目录不存在，与第 1 节逐项一致）。
- B-Retest 前**建议先由用户手动重载该窗格**（关掉任务窗格再打开，或整窗重进），确保窗格内
  是同一份完整 bundle；这一步我没有做，也不会自行做。若改用构建产物，认
  `dist/assets/index-Bll8oB8r.js`（本次构建）。
- 未提交、未推送。
