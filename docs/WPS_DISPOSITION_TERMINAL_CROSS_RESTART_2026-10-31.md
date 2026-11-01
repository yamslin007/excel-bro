# WPS 撤销后跨重启终态识别：核实、修复与真机验收步骤（2026-10-31）

> 本轮是 `docs/OFFICE_UNDO_AFTER_RESTART_2026-10-31.md`（Excel 端）的 WPS 对应核实。
> 用户要求核实的链路：中断 → 重开 → 来源变化阻断续跑 → 撤销已写入动作 → 保留用户修改的来源 →
> 再次重开后**不能续跑、不能重复撤销**。
>
> 本轮分两步：先补 disposition 终态闸（W1–W3），再按用户要求闭环"**disposition 上报失败**"这个
> 组合风险（W4–W5 + 三条生产链路回归）。

## 1. 结论先行（四类状态）

| 状态 | 结论 |
| --- | --- |
| 已实现 | ① 撤销后宿主缓存 `record.state="restored"`；② 服务端 disposition 投影；③ 客户端三层终态判据；④ 水合读 disposition；⑤ **执行入口自身**按已完成动作的写后现场复核（此前已有） |
| 自动化已验 | 同会话撤销阻断、跨重启水合推回 `restored`、留痕失败 + 材料删失败组合场景逐条零写入、检查至写入之间的两处状态变化，均有真实写入计数断言 + 反向变异 |
| 尚未实现 | 真机（完整退出 WPS 再打开）一次都没点过；留痕失败时"是否已撤销"在服务端确实没有事实可读 |
| 真机待验 | 需要用户提供一个**真实的 WPS 未完成执行现场**（当前没有已核实的可复用现场） |

**一句话**：WPS 撤销后跨重启终态识别原先有缺陷（水合按账本游标推回 `partial_write_recoverable`，
重新开放续跑与撤销入口），本轮已修；并进一步证明**即使 disposition 上报失败**，执行入口自身的
现场核对仍然零写入。两层是纵深防御：拆掉任意一层都不会实际写入，拆掉第二层（`assertPriorWritesCurrent`）
会**真的写进原表**（反向变异 W5 的失败原文里出现 `Target!A3`）。

## 2. 根因（读代码核实）

撤销成功后的四层事实：

1. `undoWpsExecution`（`wpsPlanExecutor.ts:2994`）只把宿主缓存的 `record.state` 改成 `"restored"`，
   账本 `status` / `nextActionIndex` 一个字都不动（处置不是进度）。
2. `latestWpsUndo`（同文件 `:2700`）只返回 `verified / partial_write_recoverable / undo_conflict`，
   过滤 `restored` —— 所以**同会话**撤销后按钮消失。
3. `retireWpsRecoveryPackage`（`wpsRecoveryPackage.ts:152`）删本机材料是 **best-effort**。
4. `ensureWpsRecoveryMaterial`（同文件 `:314`）水合时按**账本游标**推 state，原本**不读 disposition**。

于是跨重启后：缓存清空（`restored` 消失）→ 材料若删失败仍在 → 水合推回
`partial_write_recoverable` → 续跑与撤销入口双双重新开放。

**关键补充（本轮核实，纠正上一版文档的一处错误）**：上一版写"执行器 `resumeWpsExecution` 不核对
目标写后现场"是**错的**。执行入口有两条现场核对，且**在 HEAD 就已存在**：

- `prepareWpsResume`（`wpsPlanExecutor.ts:1489`）对每个动作比对「当前现场指纹 vs（已完成 ? 写后 : 写前）」；
- `assertPriorWritesCurrent`（同文件 `:2289`）在**每个动作写入之前**遍历所有 `appliedRows>0` 的动作，
  比对现场是否仍是 `afterFingerprint`，不一致就抛「等待期间已执行区域发生外部变化」。

所以即使 disposition 闸失效，撤销过的执行也走不通续跑 —— 本轮把它变成**可复现的证据**（§4）。

## 3. 改动清单（**严格区分本轮与此前**）

### 3.1 本轮改动

| 位置 | 改动 |
| --- | --- |
| `server/app/models.py` | `WpsExecutionStateResponse` 增加 `disposition: OfficeExecutionDispositionState \| None = None` |
| `server/app/main.py` | `_wps_state_payload` 投影 `disposition`（`undoneAt` / `abandonedAt`），与 `_office_state_payload` 对齐 |
| `apps/excel-addin/src/contracts.ts` | `WpsExecutionState` 增加 `disposition?: OfficeExecutionDispositionState \| null`（`assertWpsExecutionState` 不额外校验，原样保留） |
| `apps/wps-addin/src/wpsExecutionReconcile.ts` | `wpsReconcileVerdict` 加终态判据：`undoneAt`→`undone`；`abandonedAt`→`blocked` |
| `apps/wps-addin/src/wpsExecutionResume.ts` | `wpsResumeUnavailableReason` 与 `resumeWpsExecution` 入口各加一道终态闸 |
| `apps/wps-addin/src/wpsRecoveryPackage.ts` | `ensureWpsRecoveryMaterial` 水合读 `state.disposition`，终态推回 `restored` |
| `apps/wps-addin/src/wpsExecutionResume.test.ts` | 新增 4 条：留痕失败组合场景、检查至写入之间（claim 期间）、动作之间现场被改、撤销后水合不重开 |
| `apps/wps-addin/src/wpsExecutionReconcile.test.ts` | 新增 3 条：撤销过→undone、放弃过→blocked、终态优先 |
| `server/tests/test_execution_disposition.py` | 新增 1 条：WPS 状态响应投影 disposition |
| `scripts/mutate_wps_disposition_terminal.mjs` | W1–W5 反向变异 |

### 3.2 此前已有（HEAD 已存在，**不是本轮新增**）

| 位置 | 事实 | 本轮做了什么 |
| --- | --- | --- |
| `wpsPlanExecutor.ts:1489` `prepareWpsResume` 已完成动作现场核对 | HEAD 已有（`git show HEAD` 第 1153 行同一个错误文案） | 用生产链路回归 + W4 钉住它的独立生效 |
| `wpsPlanExecutor.ts:2289` `assertPriorWritesCurrent` | HEAD 已有（第 1923/1931 行） | 用回归 + W5 钉住"拆掉会实际写入" |
| `undoWpsExecution` 的目标写后现场核对 | HEAD 已有 | 组合场景里证明它挡重复撤销 |
| 处置留痕（U4）本体 | 见 `docs/EXECUTION_DISPOSITION_2026-10-29.md` | 本轮只消费它的 `disposition` 投影 |

## 4. 自动化证据

### 4.1 生产链路回归（新增 4 条，全部用**真实写入计数**判定）

| 用例 | 断言方式 | 结果 |
| --- | --- | --- |
| 留痕失败 + 材料删失败 + 跨重启：旧材料水合后，撤销过的执行续跑与重复撤销都零写入 | `f.writes`（宿主写调用记录）为空 + `claimWpsExecutionResume` 未被调用 + 卡面 `blocked` | 通过 |
| 核对通过到写入之间现场被改（claim await 期间） | `f.writes` 为空 + `startWpsAction` 未被调用 | 通过 |
| 动作之间现场被改（写前检查点篡改更早动作） | `f.writes` 为空 | 通过 |
| 撤销后水合读 disposition → 推回 `restored`、`latestWpsUndo` 为 null | record 状态 + 入口判定 | 通过 |

第一条是用户点名的组合风险，步骤与断言一一对应：

1. 中断并写入两步（动作 0、1 落盘，动作 2 未写；账本 `nextActionIndex=2`）；
2. `await undoWpsExecution(...)` 成功 → 目标回到写前值 `old1/old2/old3`；
3. **刻意不写 disposition、不删材料**（`state.disposition` 保持 `undefined`，后端 `ids()` 仍含该执行）；
4. `f.storage.clear()` 清空宿主缓存 → 模拟完整退出；
5. `ensureWpsRecoveryMaterial` 从旧材料水合 → `kind === "package"`、record 推回
   `partial_write_recoverable`（**如实：无法证明已撤销**）；
6. 卡面 `inspectWpsRecoveryState` → `conflict`；`wpsReconcileVerdict` → `blocked`（给原因，不给按钮）；
7. 用**撤销前捕获的旧 state** 直接调 `resumeWpsExecution`（旧回调入口）→ 抛「已完成或未完成目标已变化」，
   `claimWpsExecutionResume` 未被调用，`f.writes` 为空；
8. 再调 `undoWpsExecution` → 抛「在执行后又被修改」（`undo_conflict`），`f.writes` 为空；
9. 原表仍是 `old1/old2/old3`（既没被重新写入，也没被二次恢复覆盖）。

### 4.2 反向变异（`scripts/mutate_wps_disposition_terminal.mjs`，W1–W5 四判据全部成立）

| 变异 | 拆掉什么 | 预期变红 | 实际证据 |
| --- | --- | --- | --- |
| W1 | `wpsReconcileVerdict` 的 undoneAt 终态判据 | 「撤销过 → 终态 undone」 | `to be 'undone'` |
| W2 | `wpsResumeUnavailableReason` 的 undoneAt 终态闸 | 「撤销后跨重启…仍挡住续跑」 | `已经撤销过` |
| W3 | 水合 disposition 闸（`state.disposition?...` → `false`） | 「撤销后跨重启：水合…撤销入口不重新开放」 | `expected 'partial_write_recoverable' to be 'restored'` |
| W4 | `prepareWpsResume` 已完成动作现场核对 | 「留痕失败 + 材料删失败 + 跨重启…」 | `to not be called at all`（仍零写入，但**越过了执行边界**拿到执行权） |
| W5 | `assertPriorWritesCurrent` | 「动作之间现场被改…」 | **`Target!A3`**（真的写进原表） |

**纵深防御的准确说法**：拆掉 W4 那一层，`assertPriorWritesCurrent` 仍兜住零写入；拆掉 W5 那一层，
`Target!A3` 被写入。所以两层都不能少，且"移除后会导致实际写入"的证据由 W5 提供。

## 5. 诚实标注的边界

**留痕失败 ⇒ 服务端没有终态事实**。`recordExecutionDisposition` 失败时没有本地 fallback，
服务端 `undone_at`/`abandoned_at` 都是 `NULL`。此时：

- 水合**无法**把 record 推回 `restored`（事实就是没有）；
- 但执行入口的现场核对仍然拦（§4.1 已证）：撤销后现场是写前值，账本登记的是写后指纹；
- 卡面判 `blocked` 并给出原因，**不推断撤销成功、不放宽续跑**。

这是边界不是缺陷：不能用"材料还在不在"反推终态（那会把"删失败"误判成"没撤销过"）。

**没有可复用的 WPS 真机现场**。本轮全部结论是"真实 App 组件 + 真实服务端契约 + 替身宿主"，
真实 WPS 里一次都没走过完整退出/重开。`公式来源证据验收-Excel-B.xlsx` 是 **Excel** 的现场，
**不能**当 WPS 验收现场用。

## 6. 验证（准确命令与口径）

### 6.1 客户端

| 命令 | 结果 |
| --- | --- |
| `cd apps/wps-addin && npx vitest run` | **70 文件 passed / 2 skipped；1255 passed / 8 skipped** |
| `cd apps/excel-addin && npx vitest run` | **124 文件 passed / 3 skipped；1637 passed / 9 skipped** |
| `cd packages/core && npx vitest run` | **65 文件；1006 passed** |
| `npx tsc --noEmit`（`apps/excel-addin` / `apps/wps-addin` / `packages/core`） | 三端 0 错 |
| `npm run build:addin` / `npm run build:wps` | 退出码 0 |

### 6.2 服务端（区分定向与全量）

| 范围 | 命令 | 结果 |
| --- | --- | --- |
| 定向① | `.venv/Scripts/python.exe -m pytest server/tests/test_execution_disposition.py -q` | **11 passed** |
| 定向② | `.venv/Scripts/python.exe -m pytest server/tests/test_wps_execution_api.py -q` | **61 passed** |
| 定向合计 | 上面两个文件一起跑（即上一版写的 "72 passed"） | **72 passed**（11 + 61） |
| 全量 | `.venv/Scripts/python.exe -m pytest server/tests --basetemp=<OS 临时目录> -q` | **1948 passed / 27 skipped / 2 failed** |

两条全量失败固定在 `server/tests/test_python_event_delivery.py`
（`test_python_completion_emits_event_with_result_handle`、`test_python_event_includes_only_result_handle_not_values`），
失败原文是 `assert 'unavailable' == 'completed'` —— 本机 Docker 引擎未运行导致 python 执行返回
`unavailable`，与本轮改动无关（改动前同一台机器同样这 2 条红）。

## 7. 最短真机验收步骤（**先确认现场**）

**前提**：目前**没有**已核实的 WPS 现场可用。第一步必须由你先确认/制造一个，否则后面几步无从谈起。

1. **【你必须操作】** 在 WPS 里打开一个测试工作簿，构造三处独立写入（例如 `目标!A1:A3` 三个单元格），
   让一次执行在**第二步回读之后、完成回执之前**中断 —— 或确认你已有一个这样的未完成执行。
   把当时的授权表名、三处目标地址记下来。
2. **【你必须操作】** 重开 WPS 窗格，把授权恢复成同一组表；确认恢复卡出现且给出明确结论
   （不要在这一步点任何写入按钮）。
3. **【你必须操作】** 点「撤销本次执行」，确认目标三处回到执行前、来源表保持你改过的值。
4. **【你必须操作】** **完整退出 WPS**，重新打开同一个工作簿，再次把授权恢复成同一组表。
5. 观察两点（这是本轮的验收目标）：
   - **不续跑**：不出现「继续剩余动作」可点按钮；
   - **不重复撤销**：备份抽屉与结果卡都不出现「↶ 撤销本次执行」。

若第 5 步仍出现可点按钮：把卡片结论原文抄下来反馈。那说明 disposition 留痕没落库且现场恰好与
账本一致（§5 的边界），**不要**把结论写成"已撤销终态安全"，也不要手工清空目标。

自动化已经覆盖（**不需要**真机重跑）：同会话撤销阻断、跨重启水合推回 `restored`、留痕失败组合场景
零写入、检查至写入之间的两处变化、W1–W5 判别力。真机唯一缺的只有"完整退出 WPS 再打开"这一下的
真实宿主行为。

## 8. 未闭环

- 真机未验收（§7 全部步骤一次都没点过），且当前没有可复用的 WPS 现场。
- §5 的"留痕失败 ⇒ 服务端无终态事实"是设计边界，不作修复承诺。
- 未提交、未推送；A1-GATE 不自动通过，A2 未启动。
