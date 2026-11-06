# 恢复交互的输入框整合 —— 隔离树交付 + 主树回流（2026-11-02）

> 隔离树：`E:\codex\excel-bro-ui`（非 git 仓库，普通复制目录）。
> 主树：`E:\codex\excel-bro`（HEAD `3a8dd82c`，工作区大量未提交改动）。
> 第一段**只在隔离树实现**；第二段（本文 §8）已把结果**逐段回流主树**。
> 全程未提交、未推送、未重启 WPS/开发服务、未操作真实工作簿/账本/恢复材料、未启动 A2。

## 1. 目标与达成

把「待核对 / 可继续」要用户点的那一下，从输入框上方那张恢复大卡搬进输入框内部，像澄清问卷一样把普通输入框**置换**掉。沿用已有 `RecoveryComposer`，不另做一套、不只改文案。

- 待核对时：`任务中断 · 第 N 步结果待核对` + 已写入/尚未写入清单（从真实证据派生）+ `另提需求 / 查看详情 / 检查写入结果`。
- 核对成功后同一位置：`检查完成，可以继续` + `下一步：向结果3!B2写入公式` + `前 N 步不会重复执行` + `继续剩余 N 步`。
- 步数、目标、写入状态全部来自共享判据 `recoveryComposerTargets()`（账本游标 `nextActionIndex` + 结论 `outcome` + 动作清单），**不硬编码**；`pending` 的在途动作单独说「是否已写入还没确认」，绝不并入「已写入」。

## 2. 改动文件清单（隔离树 vs 主树）

### 新增
| 文件 | 说明 |
|---|---|
| `packages/core/src/recoveryComposer.ts` | 共享呈现判据唯一出处（`recoveryComposerState` / `recoveryComposerView` / `recoveryComposerTargets` / `recoveryComposerPrimaryLabel`） |
| `packages/core/src/recoveryComposer.test.ts` | 27 条用例 |
| `apps/excel-addin/src/components/RecoveryComposer.tsx` | 组件本体（两端共用），撤销行收进详情、诊断收进详情 |
| `apps/excel-addin/src/components/recoveryComposer.css` | 组件样式（动作行是滚动正文的兄弟，窄窗格三按钮换行） |
| `apps/excel-addin/src/components/RecoveryComposer.test.tsx` | 12 条用例（回流时加到 15 条） |
| `scripts/check_recovery_composer_layout.mjs` | Chromium 布局/截图探针 |

### 改动
| 文件 | 本轮改动 | 主树另有的未提交改动（勿覆盖） |
|---|---|---|
| `packages/core/src/executionRecoveryView.ts` | `refused` 分支补「还有 N 个动作没跑」；`replan` 已知 `reason` 先说 | 无 |
| `packages/core/src/executionRecoveryView.test.ts` | 新增「已知原因排在材料缺失前」用例 | 无 |
| `packages/core/src/formulaWriteback.ts` | `FORMULA_SOURCE_CHANGED_MESSAGE = "来源数据已改变，不能继续原任务。"` | 无 |
| `packages/core/src/processOutcome.ts` | `describeNextStep` 不再写死按钮名（`pendingRecovery`） | 无 |
| `apps/excel-addin/src/components/OfficeRecoveryPanel.tsx` | 新增 `primaryLabel` prop、`officeRecoveryPrimaryHint()` | 无 |
| `apps/excel-addin/src/types/chat.ts` | `officeResume` 新增 `targets` | 无 |
| `apps/excel-addin/src/App.tsx` | 接 `RecoveryComposer`、撤重复大卡、`composerBusyElsewhere` 减 `card.acting` | **主树另有未提交改动（对话定位条 / `onConfirmOriginal` 等），必须逐段合并** |
| `apps/excel-addin/src/excelWriteEndToEnd.test.tsx` | 按钮名族 + composer 节点断言 + IME 测试去时序化 | 主树该位置是不同测试，逐段合并 |
| `apps/wps-addin/src/WpsAgentPanel.tsx` | 接 `RecoveryComposer`、撤大卡、`wpsRecoveryTargets`、备份空态 | **主树另有 296 行未提交改动（撤销完整性等）** |
| `apps/wps-addin/src/WpsAgentPanel.test.tsx` | 按钮名族 + composer 等待条件 | 无 |
| `apps/wps-addin/src/wpsPlanExecutor.ts` | **仅 1 行**：`assertWpsResumeSources` throw 改 `FORMULA_SOURCE_CHANGED_MESSAGE` | 主树另有 216 行未提交改动（公式来源证据/撤销完整性，不同行） |

## 3. 主树撤销完整性修复与隔离树的兼容性

- `apps/wps-addin/src/wpsAppliedActions.ts`：两端**字节一致**（sha256 相同）。
- `apps/wps-addin/src/wpsRecoveryPackage.ts`：两端**字节一致**。
- 这两份是主树撤销完整性修复的载体，已在隔离树中就位 ⇒ **隔离树实现未回退主树的撤销修复**。
- `apps/wps-addin/src/wpsPlanExecutor.ts`：两端仅差 1 行（`FORMULA_SOURCE_CHANGED_MESSAGE` 文案），与主树 216 行改动不同位置、无冲突。

## 4. 验证结果（隔离树，第一段）

| 套件 | 结果 |
|---|---|
| core 全量 | 75 files / 1157 passed |
| Excel `excelWriteEndToEnd.test.tsx` | 100 passed |
| WPS `WpsAgentPanel.test.tsx` | 243 passed / 5 skipped |
| `RecoveryComposer.test.tsx` | 12 passed |

## 5. 测试时序脆弱性修复（隔离树，第一段）

`excelWriteEndToEnd.test.tsx` 的 IME 组合输入测试在**全量跑**时偶发失败（单跑通过）——根因：身份监视的 3s 定时器 / focus / visibilitychange 可能在同一时刻再触发一轮发现，原 mock 只按住第一轮。修法：mock **按住每一次**发现（`pending[]` + `releaseAll()`），使测试对"发现触发几次"不敏感。

## 6. Chromium 真实布局（隔离树，第一段）

`node scripts/check_recovery_composer_layout.mjs` 全档通过（320×640 / 400×720 / 500×760）：主操作只有一颗、有可见文字、动作行是滚动正文的**兄弟**、主操作完整落在视口内；长正文自己滚；「另提需求」后的紧凑待办入口在视口内。

## 7. 证据边界

布局证据是**独立夹具直渲染 `RecoveryComposer`**（文案由共享层 `recoveryComposerView` 派生），**不是**完整 App 状态机在真实宿主里跑出来的截图。真实 Excel/WPS 窗格一次都没点过。

---

## 8. 回流主树（第二段，2026-11-02）

### 8.1 回流方式

- **逐段合并，不整文件覆盖**。`App.tsx` / `WpsAgentPanel.tsx` / `types/chat.ts` /
  `excelWriteEndToEnd.test.tsx` / `WpsAgentPanel.test.tsx` / `executionRecoveryView.test.ts` 均为
  CRLF，`patch` 处理不了 CR（`Hunk FAILED ... different line endings`）⇒ 用
  `.workbuddy/tmp/merge_segments.py`：`difflib.SequenceMatcher` 在**保留行尾的原始行**上对齐，
  只应用"本轮到隔离树的增量"段，主树独有行原样保留；`--check` 先打印将被删的主树行供复核。
- 合并后逐文件核对：`diff_lines=0`（与隔离树一致的部分）、`CR 行数 = 总行数`（换行符未被打散）。
- 主树原本另有改动的文件（`App.tsx`、`WpsAgentPanel.tsx`、`wpsPlanExecutor.ts`）逐项 grep 复核，
  确认**全部仍在**：`ConversationNavigation` ×2、`scrollToMessage` ×2、`onConfirmOriginal`、
  `has-conversation-navigation`、公式来源证据与撤销完整性相关实现。
- **合并后的文件哈希与隔离树不同是预期的**（主树保留了后续修复），没有为追求一致去覆盖主树。

### 8.2 主树落位后与隔离树的差异（本轮新增，三处）

1. **备份抽屉空态**：新增共享判据 `packages/core/src/backupEmptyState.ts`（+ 4 条用例）。
   旧文案「只有确认并完成写入后才会创建」把写前快照说成确认之后的事，顺序反了；撤销之后
   再说「还没有恢复点」会让人以为这次根本没被保护过。现在分两态：
   - `none`：标题「暂无恢复点」+「每次写入前都会先保存完整快照」；
   - `undone`：标题「本次执行已撤销」+「写前快照在第一次写入之前就已保存；撤销完成后它被用掉并回收」。
   `latestExecutionUndone` 只认结果卡上的权威留痕（`verification.status === "undone"`），
   **不按"刚点过撤销按钮"猜**——撤销失败时它不会变成已撤销。
2. **防连点两层**（`RecoveryComposer.tsx`）：点击层 `event.detail > 1` 直接不派发；
   同时主操作带 `key={view.action}`，「检查 → 继续」**换节点**，旧节点的残余事件打不到新按钮上。
   撤销按钮同样只认第一下。新增 3 条组件用例（双击 detail 1/2 只派发一次、detail 0 放行、
   核对→继续节点不同且旧节点点击不派发）+ 1 条端到端用例。
3. **文案常量的断言跟着常量走**：`FORMULA_SOURCE_CHANGED_MESSAGE` 改成人话后，
   `apps/wps-addin/src/wpsExecutionResume.test.ts`（8 处）与
   `apps/excel-addin/src/excelFormulaResume.test.ts`（2 处）原本写死旧串 `"来源已变化"`，
   改为直接引用常量——测试里不再留第二份文案。

### 8.3 回流后自动化结果（主树，顺序执行，不并行套件）

| 项 | 结果 |
|---|---|
| core 全量（`vitest run --root packages/core`） | 76 files / **1160 passed** |
| Excel 全量（`apps/excel-addin`） | 132 files / **1752 passed**，9 skipped |
| WPS 全量（`apps/wps-addin`） | 72 files / **1296 passed**，8 skipped |
| core `tsc --noEmit` | 0 错 |
| `npm run build:addin`（含 tsc + 两次 vite build） | 通过 |
| `npm run build:wps`（含 tsc + vite build） | 通过 |
| `node scripts/check_recovery_composer_layout.mjs` | 三档宽度全过（见 §8.5） |

> 并行跑两端套件会出现 5000ms 超时假失败（CPU 抢占），本轮改为**顺序执行**后稳定复现绿。
> 这不是缺陷，是跑法问题；记录下来免得下次再被误判成回归。

### 8.4 专项验证（对应用户点名的六条）

| 要验证的事 | 用例 |
|---|---|
| 检查零写入 | 端到端「连点『检查』的第二下…」：`contentWrites` 增量为空、`startOfficeAction` 未调用；WPS/Excel 既有「核对只补记」用例同断 |
| 连点不意外续跑 | 同上：双击后 `reconcileOfficeActionReceipt` 只 1 次、`claimOfficeExecutionResume` 0 次；结算后停在「继续剩余 1 步」而不是自动跑 |
| 只续跑剩余动作 | 「在途动作先核对为『已写入』：跳过该步，只写剩余步骤」断言 `writes.map(address)` 只含未跑的区间 |
| 撤销残留不报成功 | 「回归⑥：撤销成功但留痕上报失败 → 本会话保持已撤销且不重复宿主操作；重开后无留痕也不重新开放续跑」 |
| 来源变化拒绝续跑 | 「续跑时来源值变了：拒绝续跑，一格都不写」+ WPS `wpsExecutionResume.test.ts` 8 条 |
| 草稿与 IME 保留 | `RecoveryComposer` 的 `hasDraft` / `choice` 用例 + `excelWriteEndToEnd` 的 IME 组合输入用例（去时序化后稳定） |

### 8.5 布局复验（主树）

`node scripts/check_recovery_composer_layout.mjs`：320×640 / 400×720 / 500×760 三档、reconcile 与 resume 两态全部通过；
截图在 `outputs/probe/recovery-composer-{reconcile,resume}-{320,400,500}x{640,720,760}.png`。
脚本 `cssFiles` 已含新组件样式（漏加会让窄窗格判定失真）。

### 8.6 仍未完成（不得合并宣称）

- **真实 Excel / WPS 窗格零点击**：本轮全部证据是「真实组件 + 真实服务端 + 替身宿主」+ 真实 Chromium 布局，
  属于①②档；③④⑤⑥档**未做**。
- **开发服务侧已提供新代码，窗格是否已加载未核对**（这两件事必须分开说）：
  - `3000`（Excel）/ `3987`（WPS）都在监听。只读 GET 各自模块得到的**就是合并后的内容**：
    `GET https://127.0.0.1:3000/src/App.tsx` 里 `RecoveryComposer` 出现 6 次、`recoveryComposerState` 3 次；
    `GET http://127.0.0.1:3987/src/WpsAgentPanel.tsx` 里 `backupEmptyState` 2 次、`RecoveryComposer` 18 次。
  - 但这只证明 **dev server 下次取模块会给新代码**，**不证明**已经开着的窗格重载了模块图 ——
    本轮没有重启服务、没有重载窗格、没有在窗格里核实过。真正的判据是 §8.7 第 1 步那行标题出现在窗格里。
- 未提交、未推送；未启动 A2。
- 旧探针 `scripts/check_recovery_entry_viewport.mjs` 仍引用上一代选择器（`.office-recovery-panel`），
  本轮未改；要保留就得换到 `.recovery-composer [data-recovery-primary]`。

### 8.7 最短实机验收步骤（下一轮照这个点，Excel / WPS 各一遍）

**夹具**：一张两表工作簿（`来源` + `结果1/2/3`），计划三步：`结果1!B2←=来源!B2*1`、`结果2!B2←=来源!B4*2`、`结果3!B2←=来源!B5*5`（与 B 复测同形）。授权这两张表。

1. 让它跑到**第 3 步写之前**中断（关窗格 / 断执行权），重开窗格 —— 应看到输入框被置换：
   标题「任务中断 · 第 3 步结果待核对」、`已写入：结果1!B2、结果2!B2`、`尚未写入：结果3!B2`、
   底部一颗**文字**按钮「检查写入结果」，且**只有这一颗**（不再有上方大卡）。
2. 点「检查写入结果」：三张表**一格都不变**、过程面板无写入步骤、按钮换成「继续剩余 1 步」。
   这一步**不许**顺手把第 3 步写了。
3. **双击**「检查写入结果」（或点完立刻再点原位）：只能有一次核对请求，不能跨到续跑把表写了。
4. 点「继续剩余 1 步」：只写 `结果3!B2`，`结果1/2!B2` 不重写，随后进审阅。
5. 改 `来源!B5` 为 500 再点「检查」：应提示「来源数据已改变，不能继续原任务。」并**零写入**。
6. 点「撤销本次执行」：三处目标恢复为空；备份抽屉显示「本次执行已撤销」，**不是**「暂无恢复点」。
7. 在有草稿（输入框里先打几个字）或中文输入法组合中时重开窗格：恢复操作区**不得**抢走正在输入的内容，
   只留一行「恢复待办 · 待核对」入口。

以上第 1–7 步在真实宿主里点完之前，本文件所有结论都只到"自动化通过"这一档。

### 8.8 重复状态卡去重（接手完成）

接续未完成补丁：Excel/WPS 共用 `reviewStatusBarSuppressedByRecovery`。
仅当恢复操作区实际接管输入框、审阅处于 failed 且无待清理标记、两处归属同一执行时，
收起上方重复的「本次执行需要核对」卡。结果审阅、其他执行、身份未知、阻断恢复不收起。
失败执行没有成功回执时，展示绑定必须同时匹配审阅任务 key，不能借用上一轮的 operationId。
WPS 续跑失败也绑定对应 resume 审阅任务。消息历史、范围条和执行/授权/恢复判据不变。

验证：Excel 共享判据与 E2E 107 passed；WPS 面板 244 passed / 5 skipped，
补齐续跑身份后定向中断→核对→续跑→确认/放弃 2 passed。
Excel 三步公式用例的挂起响应原为固定身份，现正向使用请求中的真实身份，另保留身份不一致的
反向用例，确认其他执行的卡仍可见。双端构建（含类型检查）及 Chromium 320/400/500 布局通过。
日志位于 outputs/recovery-dedup-*.log。本轮未操作真实宿主、工作簿、账本、恢复材料，未提交推送。
去重效果原生 Excel/WPS 待验；用户此前已确认的检查零写入不要求重复验收。

§8.7 的步骤更正：夹具为四张表，第三条公式应为 `=来源!B5*3`；使用第二步写后回执前中断探针。
界面按实际 outcome 展示（已确认写入可显示「还剩 1 步」，未知才显示该步待核对）。
正常续跑与改来源拒绝续跑是独立分支，不可完成第三步后再当作来源变化阻断用例；
双击检查应从待检查状态开始；IME 保持应测发现晚到，不应要求跨重载保留浏览器组合输入。


### 8.9 结果审阅整合输入框

用户明确授权后，将原输入框上方的结果审阅卡移入 composer-box。两端共用 ResultReviewBar 的 inComposer 分支与启用判据；现有确认、放弃、双视图及标记重试回调不变。移除同一审阅的上方重复卡和等待提示，保留范围条及其他恢复待办 chip；没有其他待办的空 dock 不渲染。原先依赖「上面」位置的恢复指路统一改为「等当前事项处理完后」。

正文默认只显示任务摘要与「请查看表格，选择保留或撤销本次修改」，长说明在查看详情，失败原因直接显示。动作行独立于正文滚动；320px 可换行。审阅结束恢复普通输入，草稿/附件状态不清空；组合输入和澄清回答期间使用原上方入口，避免卸载输入节点。

验证：Excel 120 passed；WPS 面板 244 passed / 5 skipped；双端构建通过。新增组件验证详情折叠、错误可见、处理期间禁用、确认不触发表单提交与标记清理；原 E2E 继续核验确认不重写、放弃恢复、恢复 chip 可达。Chromium 18 档布局通过，含新增 6 档审阅及长错误正文，320px 图已目检。原生 Excel/WPS 本轮未操作，不能宣称实机通过。


### 8.10 多表逐表对比

撤回「多表隐藏双视图按钮」的临时方案。授权内已修改的多张表现在可在输入框选择，对比各自执行前快照；双端共享切换状态机。旧副本经身份/内容/关闭回执保护后才切换；失败不丢句柄。WPS 排列单独放在创建之后，以免排列异常丢掉已建副本。窗口排列时切到对应原工作表。Excel 使用按表分组的缓存值快照，跨表公式不在副本试算。

验证见 DEVELOPMENT 同日章节。选表只是查看，确认和放弃仍作用于整次执行。未改真实工作簿、账本或恢复材料，未提交推送；原生 Excel/WPS 逐表并排效果待验。
