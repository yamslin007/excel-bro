## 补记：澄清必须在输入框边框内

用户截图澄清了“输入框里”的位置要求，原外置回答区不算完成。本轮两端把现有问答区和另提新需求入口一起移入 composer-box，纵向选项加选中圆标，长内容内部滚动，补充文字和发送工具栏保留。提交仍走既有单一处理器；不加自动提交，不改变授权、执行或恢复。

验证：Excel 定向 67 passed，新增内嵌断言后对应提交用例通过；WPS 面板 231 passed / 5 skipped；两端 build 通过。新脚本 check_clarification_inbox_layout.mjs 六档布局通过（两端样式 × 三档视口）。首次脚本截图输出读取 CDP 字段错误，修正 result.data 后全脚本退出 0。Excel/WPS 真机布局仍待验。本轮没有多题协议改动，不虚构问题总数；有多字段时仍完整填写后统一发送。

# 对话交互改造（Excel / WPS）— 2026-11-01（A 批 + B 批 + B 批收尾）

> 本轮只改**对话呈现与操作入口**。授权边界、来源校验、目标现场复核、执行白名单、
> 恢复判据**一行未动**；没有协议改动，没有新增暂停协议，没有服务端改动。

## 1. 范围与状态

### 普通聊天过程精简补记

- 两端共用 `isCompactMessageProcess`：仅正常完成的单步 `message-intent` 且没有工具、目标、事实或 trace 时简洁展示。展开后保留完成说明（无 note 则保留步骤标题）及耗时，不再重复阶段、编号与准备文案；原活动记录未修改。异常及多步骤保持完整展示。
- 普通聊天的过程消息后紧接回答且自身没有正文时，省略过程消息的重复作者标题，回答保留标题。
- `processOutcomeSummary` 在 idle 且无变更/验证信息时返回 null；阻断、审阅、恢复及真实回执仍显示。
- 本轮实跑：core 两份定向测试 32 passed；Excel 过程组件、摘要组件、messageFirstRouting 共 89 passed；WPS 面板 231 passed / 5 skipped；两端 build（含类型检查）通过。日志：`outputs/chat-presentation-*.log`。
- 初次 Excel 回归两条旧断言查找“理解消息”失败，已改验真实完成说明，同时保留不读取工作表断言，新增两端无空摘要、无重复作者、无准备文案检查。未操作真实工作簿；真实窗格视觉效果仍待确认。

| 需求 | 状态 |
| --- | --- |
| ① 消息去重、结果摘要突出到过程区外 | **已完成**（共享派生 + 两端接入） |
| ② 限高执行过程框（160px / 内部滚动 / 结束收起一行 / 不抢滚动） | **已完成** |
| ③ 澄清融入对话：问题在消息流、快捷回答紧贴输入框、点选只选中、发送才提交 | **已完成**（两端；§4） |
| ④ 主操作统一到输入框右下角（发送/暂停/继续/**核对**） | **已完成**（核对与「继续剩余动作」已接上，卡面不再有同名按钮；§3） |
| ⑤ 单张当前操作卡：固定区域收敛为**一个容器**，一次只展开一块 | **已完成**（`ComposerActionDock`；§3） |
| ⑥ 自动跟随、状态文案、浮标 | **已完成**（浮标保持无底色，用 `aria-label`/`title`/`data-mode` 表达新进展） |
| ⑦ 澄清组件只留一套 | **已完成**：`ClarificationFields` 连同它的用例一起删，有意义的断言迁入新组件与契约测试 |
| ⑧ 输入法验收口径 | **已修正**（见 §7-3）：要求是"不发送、不触发暂停/继续"，**不是**"候选不上屏" |

## 2. 用户现在会看到什么

过程区（A 批，未变）：

1. 过程框进行中**最高 160px**，多的部分在框里自己滚；状态栏（步数/耗时/暂停与等待分账/结论 badge）始终在外。
2. 内容超出限高时多一颗「多显示几行」（420px / 结束态 560px，且与 `46vh` 取小），可以「还原高度」。
3. 正常结束**收起成一行**：`已完成 · 10.8s` + 「展开」；**失败/中断/取消/等待回答仍然展开**（失败原因因此始终在折叠区外）。
4. 过程框下面多一段**结果摘要**（改了哪里 / 验证结果 / 下一步），它在折叠区之外，收起也看得见。
5. 用户自己点过「展开/收起」之后，任何自动收起都不再压过他；结束时也不会把正在上翻的视图折叠掉。
6. 收起/展开**不再调用 `scrollIntoView`**：不替用户决定"现在该看哪"。

输入区（B 批）：

7. **只有一颗主按钮**，位置固定在输入框右下角。形态由 `composerPrimaryAction` 一次算出：
   `↑ 发送` / `↑ 打断并补充` / `↑ 发送回答` / `◎ 核对在途动作` / `▶ 继续剩余动作` / `▶ 继续任务` /
   `■ 暂停任务` / `… 正在暂停`。
8. **输入区上方的运行按钮没了**，只剩一行状态说明（等待审阅 / 正在停止 / 不可中断的原因 / 已暂停可从右下角继续）。
   状态留在这里，动作全在右下角。
9. **澄清回答紧贴输入框**：问题本身仍在消息流里（那是上下文），选项/多字段/"将发送：…"预览在输入框上方。
   点选项**只选中**（`aria-pressed` 可读、再点一次取消），按发送才提交；自由输入就是输入框本身，不再有第二个光标。
10. 必填没填完时主按钮禁用且说明"还差哪几项"；提交绑定 `conversationId + messageId + clarificationId`，
    换对话/换任务之后旧题的选中与草稿一起失效（按键身份判定，不靠清空时机）。
11. 显式逃生口「另提新需求（不回答本题）」：**不按文字猜**你想回答还是想换需求。
12. **输入框上方的固定区域一次只展开一块**（审阅结果 / 恢复核对），其余压成一行 chip：
    能切的点一下换主视图，不能切的如实说明（澄清"就在下方"、运行"控制在右下角"）。
13. **恢复卡只剩状态、原因、进度、撤销、补记与一句指路**："下一步「核对在途动作」在输入框右下角"。
    核对**只补记回执**，核对成功后必须**再点一次**才继续 —— 不自动续跑。

## 3. 真实调用链（改动后）

| 交互 | Excel | WPS | 共享层 |
| --- | --- | --- | --- |
| 主按钮形态 | `composerPrimaryAction(...)` | 同一个函数 | `composerPrimaryAction.ts` |
| 主按钮渲染 | `ComposerPrimaryButton` | 同一个组件 | 同上（`key={kind}` + `detail>1` 防双击变续跑） |
| 回车 | `handleComposerKeyDown` → `invokeComposerPrimary(invoke)` | 同样的分派 | 判据同一份 |
| 固定区域当前事项 | `selectCurrentAction(input, pinned)` | 同一个函数 | `currentActionCard.ts` |
| 固定区域容器 | `ComposerActionDock` | 同一个组件 | 同上 |
| 澄清草稿 | `clarificationDraft`（选中/字段）+ `prompt`（自由输入） | `clarificationChoice`（索引）+ `clarificationFieldDrafts` + `prompt` | `clarificationComposer.ts` |
| 能不能提交 | `activeClarificationSubmission.ok` | 同一个函数 | 同上 |

**主按钮的优先级（唯一一份）**：

```
pausing → send-supplement → answer → 【有字 ⇒ send】 → reconcile → recoveryResume → paused(继续分析) → pause
```

- `send`（有字）排在 `reconcile` / `recoveryResume` **之前**：**输入框里有字时，主按钮永远是"发送"**。
  代价是"有字时点不到核对"，所以两端 `composerHint` 会明说"输入框有内容：发送会作为新需求；要核对请先清空输入框"。
  这一条宁可让入口暂时点不到，也不许悄悄把用户正在写的字丢掉。
- `recoveryResume`（写剩余动作）与 `paused`（从检查点继续分析）都会返回 `kind: "continue"` + `invoke: "resume"`，
  但**走不同处理器**；文案由宿主用 `reconcileLabel` / `resumeLabel` 指定，避免"继续"两个字指两件事。

**回车只表达一件事：把这段话发出去。** `invoke` 是 `pause`/`resume`/`reconcile`/`null` 时回车什么也不做
（不 `preventDefault`，换行照旧）。输入法组合中的回车三道判据一起挡：
`nativeEvent.isComposing`、React 侧同一标记、`keyCode === 229` —— **命中就 return，不 `preventDefault`**，
所以候选字照常上屏（见 §7-3）。

## 4. 本轮新增/重写/删除的文件

**共享层（packages/core/src/）**
- `clarificationComposer.ts` + `.test.ts`（18 条）：绑定身份、草稿归属、单/多选、三选一提交（字段 > 选项 > 自由输入）、必填缺项、**空白不算填**。
- `composerPrimaryAction.ts`（`reconcile` / `recoveryResume` / 可定制文案）+ `.test.ts`（**21 条**）。
- `currentActionCard.ts`（`currentActionShortLabel` / `selectCurrentAction(input, pinned)` / `currentActionSwitchTargets`）+ `.test.ts`（**11 条**）。

**Excel 端（apps/excel-addin/src/）**
- 新增 `components/ComposerActionDock.tsx` + `.test.tsx`（**9 条**）+ `composerActionDock.css`。
  chip 行只从 `currentActionSwitchTargets` 取（与 `others` 同源，不另算一份）：可见文案是短名，可访问名与 `title`
  是完整说法（"先处理这一项：核对上次写入留下的在途动作"）—— 一行放得下，读屏也听得懂。
- 新增 `components/ClarificationComposer.test.tsx`（**7 条**）：必填缺项、标签拼法、busy 全禁用、点选只报告 id、
  多选可取消、选项+补充两句都在、单选/多选可访问角色不同。
- `components/OfficeRecoveryPanel.tsx`：主按钮搬走后卡面只剩状态/原因/进度/撤销/补记 + `data-recovery-primary-hint` 指路。
- `components/ResultReviewBar.tsx`：拆出 `resultReviewVisible`（会不会画）与 `resultReviewNeedsDecision`（在等你决定）。
- `App.tsx`：`recoveryPrimaryAction` 选出 `reconcile`/`resume` 并交给主按钮；`invokeComposerPrimary` 分派到
  `runOfficeRecoveryAction`；`dockPinned` 只认仍然存在的那件事；审阅栏按"有无决定"分两处渲染。
- **删除** `components/ClarificationFields.tsx` / `.test.tsx`（生产零引用；契约级断言迁入 `contracts.test.ts`）。
- **删除** `components/CurrentActionSummary.tsx` / `.test.tsx` / `currentActionSummary.css`（被 dock 的 chip 行取代；
  "其余待办逐条点名"的断言迁入 `ComposerActionDock.test.tsx`）。

**WPS 端（apps/wps-addin/src/）**
- `WpsAgentPanel.tsx`：`planningPauseTarget` / `wpsRecoveryPrimary` 决定主按钮形态；
  三张恢复卡从消息流搬进固定区域，卡面 `reconcile`/`resume` 按钮过滤掉，改由图钉指路；
  规划暂停卡动作区两颗按钮一并移除（只留一句"在输入框右下角"）。

## 5. "不可中断"不再靠 UI 进度字段猜

| | 改动前 | 改动后 |
| --- | --- | --- |
| Excel | `resultReviewState.phase === "executing"` | 不变（本来就是阶段枚举） |
| WPS | `writeProgress !== null`（**一行进度文字**） | `hostStage !== null`，取值来自执行器自己的 `update.stage`（`snapshot / writing / verifying / undoing`）与撤销链路 |

为什么必须换：进度文字是**呈现**，文案一改或某条分支忘了设置，按钮就悄悄变得"可以暂停"。
现在暂停判据只认执行器给的阶段。暂停请求在飞时仍是 `…` + `aria-busy` + 禁用，连点不发第二个请求。

## 6. 证据（本轮实跑）

| 项 | 命令 | 结果 |
| --- | --- | --- |
| 共享层全量 | `cd packages/core && npx vitest run` | **71 files / 1095 passed** |
| Excel 全量 | `cd apps/excel-addin && npx vitest run` | **127 passed \| 3 skipped（130 files）／1674 passed \| 9 skipped（1683 tests）** |
| WPS 全量 | `cd apps/wps-addin && npx vitest run` | **70 passed \| 2 skipped（72 files）／1258 passed \| 8 skipped（1266 tests）** |
| 类型检查 | 两端 `npx tsc --noEmit` | 均通过 |
| **构建** | `npm run build:addin` / `npm run build:wps`（各自 `tsc --noEmit && vite build`） | 均通过（addin 另含 `vite.functions.config.ts` 第二轮；只剩既有的 chunk >500kB 提示） |
| 布局（过程框） | `node scripts/check_process_panel_layout.mjs` | 三档视口：过程框 160px、放大 294/331/350px、输入区仍在视口、结束收起且摘要仍在 |
| 布局（恢复入口） | `node scripts/check_recovery_entry_viewport.mjs` | 三档视口：**主按钮**在视口内（320×640 top=593/bottom=624）、追加日志后不动、卡面 0 颗按钮、旧位置复现"看不见"（bottom=-4647.6）；320×640 高卡片：面板限高 294.4px、正文 410→200 自滚 |

**反向变异（不是"只跑一遍绿"）** —— 一次一个 SEED，写回后读回校验 sha256 一致：

| 变异 | 结果 |
| --- | --- |
| M1 把 `hasInput` 那段短路成 `false &&`（有字不再优先发送） | `composerPrimaryAction.test.ts` **2 failed**：`已暂停但有输入 ⇒ 发送补充要求`、`输入框有字 ⇒ 永远发送；既不核对也不续跑` |
| M2 让 dock 同时展开两块 | `ComposerActionDock.test.tsx` **1 failed**：`审阅与恢复同时在场时，只展开审阅那一块` |
| M3 让 `pinned` 失效（永远按默认优先级） | `currentActionCard.test.ts` **2 failed**：`显式选中较低优先级的那件时，它成为主视图…`、`切换候选与 others 同源…` |
| M4 把「有内容就说明冲突」那条短路成 `false &&` | `excelWriteEndToEnd.test.tsx` **1 failed**：`恢复入口只在输入区上方…`，断言 `expected … to contain '输入框有内容：发送会作为新需求'` |

M1 第一次写成了"给 `recoveryResume` 加 `&& !hasInput`"，那是**等价改写**（原本就排在 `send` 之后），
用例全绿 —— 说明变异必须挂在真正决定行为的那个判据上，不然"绿"什么也证明不了。

**因交互变更而更新的既有断言**（逐条说明，没有放宽判据）：

| 文件 | 原断言 | 现在 | 为什么 |
| --- | --- | --- | --- |
| `messageFirstRouting.test.tsx` | `.clarification-custom-input textarea` + 「提交补充」 | `.composer textarea` + 「发送回答」 | 自由输入合并进输入框，提交只剩一处 |
| 同上 | 点选项即提交 | `点选 → aria-pressed=true → 发送回答` | 点选只选中 |
| 同上（暂停用例） | 暂停后「发送」禁用 | 暂停后主按钮是 `▶ 继续任务`（不再是灰掉的 ↑） | 只有真可续跑才给"继续" |
| 同上（输入法用例） | 只断"没发送" | 再断 `fireEvent` 返回 `true`（**没 `preventDefault`**）+ 输入框内容还在 | 口径改为"候选可以上屏，但不发送/不暂停"（§7-3） |
| `WpsAgentPanel` / `wpsPlanningPauseCard` / `wpsPlanningResumeClosure` | `findButton` 按 `textContent` | `findButton` 按 `textContent` **或 `aria-label`** | 主按钮只画字形，动作名在 `aria-label` 上（读屏靠它） |
| 同上（恢复卡） | 卡面有「核对在途动作」两颗按钮 | 卡面 0 颗 + 主按钮 1 颗，且主按钮 `closest(".composer")` 非空 | 入口搬家（§7-1） |
| `wpsPlanningResumeClosure`（暂停轮次） | 等"继续"出现 | 先 `resolvePause` 再等；`pauseRound` 等 `liveSegmentCount === 0` 并断言主按钮形态 | 暂停在飞时主按钮是禁用的 `… 正在暂停` |
| `wpsPlanningPauseCard`（有字用例） | 卡面按钮可点 | 有字时主按钮是「发送」，用原生 setter 清空后「继续分析」回来 | **判据正确**（有字即发送），断言要跟着口径写 |
| `excelWriteEndToEnd`（核对在途） | 找「正在核对…」文案 | 主按钮唯一且 `disabled` | 那句话只画在卡面按钮的 `aria-label` 上，按钮搬走后不再出现 |
| `folderScopeSafety` | 只点 `/^发送$/` | `/^(发送\|发送回答)$/` | 有 pending 问题时发送发的是回答 |
| `TaskExecutionControls.test.tsx` | 控制条上的 ■/▶ | 拆到 `ComposerPrimaryButton.test.tsx` + `TaskExecutionNotice.test.tsx` | 组件搬家，断言跟着搬 |

## 7. 上一轮留白：按结论关闭（不再"两边都算一点"）

1. **核对入口已统一到右下角**（原来那一处有意偏离已关闭）。
   恢复卡保留状态、原因、进度、撤销、补记与一句指路；卡面**不再有**同名主按钮。
   核对成功后**必须再由用户点一次**才续跑，成品里没有任何自动续跑路径。
   点击时复核（`officeRecoveryActionLock`）、终态守卫、身份闸、多候选选择、防重复执行全部沿用原处理器 ——
   这是**入口迁移**，不是恢复判据的改动。
2. **`ClarificationFields` 已删**。全仓核实：除组件本体与它自己的测试，生产引用为零
   （`App.tsx` 的 `updateClarificationFieldValues` 是同名函数的巧合，已顺手改名以免歧义）。
   仍有意义的断言分两处保住：组件行为迁入 `ClarificationComposer.test.tsx`；
   **协议边界**（fields-only 合法、空表单/重复 id/空白标签被拒）迁入 `contracts.test.ts`。
   同类问题一并处理：`CurrentActionSummary` 被 dock 的 chip 行取代后删除，其断言迁入 `ComposerActionDock.test.tsx`。
3. **输入法口径已修正**：中文输入法候选阶段按 Enter **可以正常确认候选字进入输入框**；
   要求是"**不发送消息、不触发暂停/继续**"。代码侧本来就只是 `return`（不 `preventDefault`），
   所以这次改的是**断言与验收口径**，不是行为。
4. **冲突提示与判据不同源（本轮核实发现、已修）**：`hasInput` 把**待发图片**也算"有内容"，
   而"为什么核对入口不在了"那句提示原来另写了一遍 `prompt.trim()` —— 只贴了图片、没打字时
   入口被顶掉，界面上却没有任何说明。现在提示读主按钮自己的输出（`composerPrimary.kind === "send"`），
   与判据同源；有图片时措辞补「并移除待发图片」。
   这句提示**原先两端都没有任何断言**；已补正/反两个方向各一条（有字时给说明、清空后入口回来），
   并用反向变异 M4/M5 证明断言不是恒真。**图片-only 两端都已补测试**：`prepareImageFile` 是附件接缝，
   在 `excelWriteEndToEnd.test.tsx` 与 `WpsAgentPanel.test.tsx` 各用 `vi.mock` 只替掉"解码 + 压缩"一步，
   其余（限额、状态、判据、`removeImage`）都是真的 —— 断言看**主按钮的实际可访问名**（`发送` ↔ `核对在途动作`），
   不只看文案。WPS 侧的变异（`pendingImages.length > 0` → `false`）会让该用例变红，还原后 sha256 逐字节一致。
   WPS 那条用例同时把"唯一主入口"钉住：主按钮是 `.composer .composer-primary`，`[aria-label="核对在途动作"]` 全页恰好 1 个。
5. **提示说法跟着路由走**：暂停之后发出去的是**重新判断**这条要求（不是重放已暂停的计划），
   所以冲突提示在 `paused` 时改说"发送会重新判断这条要求（不重放已暂停的计划）"，其余状态才说"作为新需求"。
6. **撤销 / 补记完成现在分别在哪**（把"卡面 0 按钮"这句话说准）：

   | 动作 | Excel | WPS |
   | --- | --- | --- |
   | 撤销本次执行 | 结果卡「↶ 撤销本次执行」→ `undoExecutionTo` / `undoFolderExecutionTo`（`App.tsx:12660-12680`）；当**恢复卡接管该执行**时消息里让位，改由恢复卡正文的撤销行 → `undoOfficeRecoveredExecution`（`App.tsx:3792`，入口由 `officeRecoveryUndoFor` 决定） | 执行结果卡内「撤销本次执行」→ `undoExecution(id)` → `undoWpsExecution`（`resultUndoRef`，`WpsAgentPanel.tsx:1538`/`1584-1600`）；恢复卡**不画**撤销 |
   | 补记完成（不重写原表） | 恢复卡动作条 `backfill` → `runOfficeRecoveryAction(row,"backfill")` → `backfillOfficeCompletion`（`App.tsx:3356`） | 恢复卡动作条同款按钮（文案来自 `EXECUTION_RECOVERY_ACTION_LABEL.backfill`）→ `runWpsRecoveryAction("backfill")` 分支 |
   | 核对在途 / 继续剩余动作 | 输入框右下角主按钮（卡面只剩指路） | 同（`核对任务状态` / `继续分析`） |

   **`卡面 0 颗按钮` 只对 `reconcile` / `resume` 两个结论成立**：卡面次要动作只保留 `backfill`，
   所以 `completed` 会有「补记完成（不重写原表）」、`blocked`/`insufficient` 会有「重新核对（只读）」+「清除这条恢复提示」。
   准确说法是"**核对/续跑不重复**"，不是"卡上什么都不能点"。
   证据：`excelWriteEndToEnd`「全部动作已写入、只缺最终回执：只给「补记完成（不重写原表）」…」、
   「确认写入 → 显式撤销范围 → 撤销后逐单元格回到执行前」、
   「来源恢复登记基线后：点「重新核对（只读）」按当前现场重新判断」；WPS「撤销按钮归属：执行结果卡内，按 executionId 绑定」。

**仍未完成（不写成"仅待实机"）**：真实 Excel / 真实 WPS 里这套限高、收起一行、右下角主按钮、
固定区域切换、输入法上屏**都还没在真实窗格点过**。布局证据只到"真实构建 + 真实样式表 + 真实 Chromium"，
输入法证据只到 jsdom 的 `isComposing`/`keyCode 229` + 真机待验。

## 8. 两端统一的最短真机清单（每步一个动作；Excel、WPS 各走一遍）

### 8.0 夹具与前置（只做一次）

- 新建并**保存** `对话交互验收.xlsx`，四张表：
  | 表 | 内容 |
  | --- | --- |
  | `来源` | A1:B5 = 10,20,30,40,50 / 1,2,3,4,5 |
  | `结果1` / `结果2` / `结果3` | 各在 A1 写 `target-1/2/3`，**B2 留空** |
- 起服务：`npm run dev:server`、`npm run dev:addin`（或 `npm run dev:wps -- --launch`），窗格用 **DEV 构建**（底部才有「恢复验收（开发专用）」）。
- 第 8 步之前**只勾 `来源`**；第 8 步要勾齐四张。
- 标记：**【展示】**= 不写任何单元格；**【写入】**= 会真的写表，必须在测试簿上做。

### 8.1 步骤

| # | 只做这一个动作 | 看什么 |
| --- | --- | --- |
| 1 | 【展示】发「列出 来源 表 A1:B5 的内容」 | 右下角只有一颗按钮；输入区上方没有第二颗运行按钮（只有一行状态字） |
| 2 | 【展示】等它答完 | 收起成 `已完成 · Xs` + 「展开」；结果摘要仍在折叠区外 |
| 3 | 【展示】点「展开」 | 步骤回来，且**页面不跳**（滚动位置不变） |
| 4 | 【展示】**可选**：若出现「多显示几行」就点它 | 过程框变高（≈420px），框外状态栏仍在。出现条件见 §8.2：必须**先展开**且内容**实际溢出**；短过程不出现 ⇒ **跳过，不算失败** |
| 5 | 【展示】点「还原高度」（仅在上一步点了增高、且按钮存在时才做） | 回到 160px |
| 6 | 【展示】发一句**不读写任何工作表**的普通聊天，原文：`用一句话解释什么是移动平均`，按 Enter | 消息发出、输入框清空。**这句话里不许出现表名、列名、"再补一列"之类**——本步只验"能不能发出去"，任何触表的话都会污染第 9 步之后的现场 |
| 7 | 【展示】微软拼音打「汇总一下」，**候选阶段按 Enter** | 候选字**进输入框**；不发消息、不暂停/继续；再按一次 Enter 才发送 |
| 8 | 【写入】勾齐四张表（`来源`+`结果1`+`结果2`+`结果3`）→ 底部「恢复验收（开发专用）」→「启用下一次第二步中断」→ 原样发 `docs/FORMULA_SOURCE_RESUME_ACCEPTANCE_2026-10-30.md` §4 的任务原文 | 出现「第二步已写入并回读，已在完成回执前中断」；`结果1!B2`/`结果2!B2` 有公式，`结果3!B2` 空。**造出卡就停手**：不点核对/继续/撤销，不关窗，不清理现场，不续跑 |
| 9 | 【展示】看输入区上方那张恢复卡 | 主按钮是 `◎ 核对在途动作`（WPS `核对任务状态`）；卡面**没有**同名按钮，只有一句「下一步在输入框右下角」+ 次要动作（如「清除这条恢复提示」） |
| 10 | 【展示】在输入框打一个字（别发） | 主按钮变 `↑ 发送`，上方写「输入框有内容：发送会作为新需求；要核对在途动作请先清空输入框」（WPS 这句里是「核对任务状态」）；清空后 `◎` 回来 |
| 11 | 【展示】**只贴一张图片**、文字留空 | 主按钮仍是 `↑ 发送`，提示补一句「并移除待发图片」；点图片上的 × → `◎` 回来 |
| 12 | 【展示】发「列出 来源 表 A1:B5 的内容」（恢复卡仍在） | 固定区域出现 chip 行：恢复卡被压成「恢复核对」；点它 → 恢复卡回到主视图；运行那项只显示「控制在右下角」且不是按钮。旧执行**不会**被顺手续跑。chip 需要**同时有两件事**才画（条件见 §8.2）：只有恢复卡一件事时不出现 chip 是正常的，此时只验"主视图就是恢复卡" |
| 13 | 【展示】把窗格拉到最窄（≈320px） | 恢复卡自己滚、输入框不被挤出、右下角按钮可点、chip 行不折成两行 |

### 8.2 三个"可跳过"的准确出现条件（不出现 ≠ 失败）

1. **「多显示几行」**（`TaskProgressTimeline.tsx:414`）：渲染条件是 `expanded && hasLongContent`，
   而 `hasLongContent = processPanelHasOverflow({live, enlarged, contentHeight: body.scrollHeight})`
   —— 在 `useLayoutEffect` 里用 `ResizeObserver` **实测**滚动高度是否超出当前框高；**未展开时强制 `false`**（`257-261`）。
   ⇒ 先展开、且内容真的溢出才有这颗按钮。一句查询这类**短过程不溢出 ⇒ 不出现，不算失败**，第 4/5 步直接跳过。
   同一组件被 WPS 复用（`WpsAgentPanel.tsx:147` import / `9516` 使用），两端判定同源。
2. **固定区域的 chip 行**（`ComposerActionDock.tsx`）：只有 `selection.others.length > 0` 才画 chip；
   `primary === null && others.length === 0` 时**整块不渲染**。⇒ 只有恢复卡一件待办时**不出现 chip 是正常的**，
   第 12 步此时只验"主视图就是恢复卡"。要让 chip 出现，必须当下**同时**存在第二件事（本步就是那次只读查询在跑）。
3. **卡面按钮数量**：`reconcile` / `resume` 两个结论下卡面只剩一句指路（"卡面 0 颗按钮"只对这两个成立）；
   `completed` 会有「补记完成（不重写原表）」，`blocked` / `insufficient` 会有「重新核对（只读）」+「清除这条恢复提示」。
   **不许把"卡面还有次要动作"记成缺陷**，也不许把"核对/续跑不在卡面"记成"按钮不存在"。

### 8.3 澄清：本清单**没有**强制步，但它是真机待验项

追问由模型决定，不能靠改写任务去凑。若要验，发 `按级别填写等级`（不给级别清单）：

- 模型不追问 ⇒ 记「**未覆盖**」，**不能算整轮通过**（这是"本轮真机未覆盖项"，不是"通过"）。
- 模型追问 ⇒ 只验三件：点选**只选中**（输入框不动、不发消息）→ 按发送才提交 → 提交后问题不再重复问。
- 自动化证据只到 `messageFirstRouting.test.tsx` / `WpsAgentPanel.test.tsx`（点选只选中、提交只读 `clarificationSubmission`），
  **不等于真实窗格里被问过一次**。

### 8.4 第 8 步的硬约束与"造卡即停"

- 它用既有确定性入口（`RecoveryProbeControl`「启用下一次第二步中断」），**不是**让用户随手发一条写表需求。
- 授权**恰好四张**（`来源`/`结果1`/`结果2`/`结果3`）；任务原文必须落成**三个独立 `writeFormulas` 单格动作**
  （各写一格 `B2`，1×1），**不合并区域、不加格式动作、不改 `来源` 表**。合并或加动作会让"第二步中断"落在别的动作上。
- 目的是**造出一张恢复卡**；造好就**保留现场**（不关窗、不清空 `B2`、不清理恢复材料、不续跑）。
  核对 / 续跑 / 撤销的完整链路**本轮不重复**，既有证据范围见 §8.5。

### 8.5 既有真机证据：按「宿主＋动作类型＋分支」逐项（2026-11-01 重新核对，纠正上一版）

**上一版把三件事误写成"未验"，这里逐项恢复**：① Excel 公式 A/B 两分支；② Excel「重开之后撤销」；
③ WPS 的放弃恢复。核对方式 = 读既有验收文档 + **只读**核验已保存文件 + **只读**查询账本
（`%LOCALAPPDATA%\Excel Bro\tasks\state.sqlite3` 的**副本**，未改原库、未动工作簿）。

**口径（先立规矩，上一版就是在这里出错）**：

- **界面/文件通过 ≠ 账本留痕通过**；反过来也成立：**账本某列为 `NULL` ≠ 用户没做过** ——
  列可能后来才加、上报是 best-effort、WPS 那几轮早于留痕实现。
- 五列分开写：用户实机操作（含截图）、已保存文件只读核验、账本核验、自动化证据、真正剩余缺口。
- 引用不到的位置写「**待定位**」，不写成「未做」。

#### A. Excel · 公式（3×`writeFormulas`）· A 分支：中断 → 完整退出重开 → 核对 → 续跑只写第三步 → 确认

- **实机**：用户 2026-10-31 09:10–09:14 在真实 Excel 里走完（步骤即 `docs/FORMULA_SOURCE_RESUME_ACCEPTANCE_2026-10-30.md` §5.A）。截图：**待定位**（文档只留结论，未归档图片）。
- **文件只读**（本轮复核）：`C:\Users\jia\Desktop\测试\公式来源证据验收-Excel-A.xlsx`，11789B / mtime `10-31 09:15` / sha256 `aa8a913c…cb0126`；`结果1!B2`=`=来源!B2*1`、`结果2!B2`=`=来源!B4*2`、`结果3!B2`=`=来源!B5*3`（**公式文本**），`来源` A1:B5 = 10…50 / 1…5、`B5=5`。
- **账本**（本轮只读复核）：`execution-9875c898…` status `completed` / `next_action_index=3`；`office_host_writes` 恰好 3 行，`idx0`+`idx1` @`09:10:49`、**`idx2` @`09:14:33`** —— "续跑只写第三步"有账本写序，不是推断；`result_json.verification` = `verified` / `passed=true` / `unverifiedActions=[]`，六项检查里三格 `formulasEqual` 全 True（**公式身份按文本比对**）。
- **自动化**：`excelWriteEndToEnd.test.tsx`「第二步中断 → 重挂载 → 只续第三步」等；`formulaWriteback.test.ts` 续跑来源判据。
- **剩余缺口**：截图待定位；「确认变更」只有界面证据，账本侧可见的只是 `completed + verified`，没有单独的确认留痕。

#### B. Excel · 公式 · B 分支：重开后改来源 → 续跑被拒 → 目标不变 → 撤销

- **实机**：把 `来源!B5` 改成 `500` 后点续跑，界面原文「**来源已变化，不能沿用旧结果续跑。**」；重开 Excel 后先暴露"没有撤销入口"（真机缺陷，`docs/OFFICE_UNDO_AFTER_RESTART_2026-10-31.md` §8），修复后**用户点了「↶ 撤销本次执行」并成功**；随后又暴露卡面同屏"已撤销 + 可继续"（`docs/OFFICE_RECOVERY_CARD_STATUS_2026-10-31.md`）。截图：**待定位**。
- **文件只读**（本轮复核）：`公式来源证据验收-Excel-B.xlsx`，11443B / mtime `10-31 13:32` / sha256 `b480a362…62b760`；`结果1/2/3!B2` **全空**（无值无公式）、`来源!B5=500`。
- **账本**（本轮只读复核）：`execution-6d29e377…` status `failed` / `next_action_index=2`；`error_json` = `{"code":"OFFICE_RESUME_INTERRUPTED","message":"来源已变化，不能沿用旧结果续跑。"}`；`office_host_writes` **只有 `idx0`/`idx1` @`09:31:34`**（被拒后零第三次写入）；**`undone_at` = `10-31 10:30:09`**。
- **自动化**：`excelWriteEndToEnd`「Excel 未完成执行：重开后从账本材料撤销已写入的内容」5 条 + `executionUndoRebuild.test.ts` 18 条 + 变异 U1–U4（四判据）。
- **剩余缺口**：撤销后 Excel 没有 WPS 那种"恢复包回收"（材料就是账本 `checkpoint`），本清单不验；截图待定位。

#### C. Excel · 公式 · 已完成执行上的撤销（同会话 LIFO）

- **实机**：09:27 那次 A 分支完成后，用户在同会话点「撤销本次执行」把三格清空，再把同一夹具用于 B 分支。
- **文件只读**：即 B 条的文件（三格空），时间线上 `09:29:33` 撤销、`09:30:06` 重新执行。
- **账本**（本轮只读复核）：`execution-e73b1188…` status `completed` / next 3；writes `idx0,1`@`09:27:36`、`idx2`@`09:28:09`（第二次印证"只续第三步"）；**`undone_at` = `09:29:33`**。
- **自动化**：同 B + 「确认写入 → 显式撤销范围 → 撤销后逐单元格回到执行前」。
- **剩余缺口**：无（这一格是五列里证据最完整的）。

#### D. WPS · 值（3×`writeValues`）· A 分支：中断 → 完整退出重开 → 核对 → 续跑 → 确认

- **实机**：`5a2d84be…`（10-28 19:42）与 `1a727455…`（10-29 11:41→12:18）两次；结论见 `docs/ABANDON_AFTER_RESUME_ACCEPTANCE_2026-10-28.md` / `docs/A1_GATE_REMAINING_2026-10-28.md` §0。截图：**待定位**。
- **文件只读**（本轮复核）：`outputs/wps-abandon-20261028/WPS放弃验收-20261028.xlsx`：`目标表!B2:B4` 三格有值、`来源表!B2:B4` = 111/222/333。
- **账本**（本轮只读复核）：`5a2d84be…` writes `idx0,1`@`19:42:36`、`idx2`@`19:42:46`；`1a727455…` `idx0,1`@`11:41:26`、`idx2`@`12:18:31` ⇒ **"续跑只写第三步"在 WPS 侧也有账本写序**。
- **自动化**：`wpsPlanningResumeClosure.test.tsx` / `wpsExecutionResume.test.ts` / `WpsAgentPanel.test.tsx`「中断→核对→续跑→确认/放弃」。
- **剩余缺口**：截图待定位；两轮的 `abandoned_at`/`undone_at` 为 `NULL`（口径见 E 条）。

#### E. WPS · 值 · 放弃（跨进程：中断 → 完整退出重开 → 核对 → 续跑 → 放弃/撤销 → 回基线）

- **实机**：`3f00ea0d…`（10-28 23:25 中断 → 23:39 续跑第三步 → 放弃）；`A1_GATE_REMAINING_2026-10-28.md` §0 记为**真机通过**（材料跨会话生效、补记不重写、放弃回基线），用户截图显示 `-101/-202/-303`（截图文件**待定位**）。
- **文件只读**（本轮复核）：`目标表!B2:B4 = -101/-202/-303`（回基线）；`来源表` 仍是 111/222/333（没有被写回）。
- **账本**（本轮只读复核）：`wps-execution-3f00ea0d…` status `completed` / next 3；writes `idx0,1`@`23:25:50`、`idx2`@`23:39:26`（**只续第三步**）；**`abandoned_at` / `undone_at` 均为 `NULL`**，`recovery_package_*` 四列也全空 —— 这一轮早于处置留痕与包回收列的实现，**`NULL` 只说明"这轮没有留痕"，不能反推"用户没放弃过"**；界面与文件证据照旧成立。
- **自动化**：`wpsExecutionResume.test.ts`「撤销成功后 record 变 restored，续跑在拿到执行权前就被 state 闸拒绝、零写入」等。
- **剩余缺口**：处置留痕与恢复包回收的**真机取值**要等新一轮执行（WPS 端上报是 best-effort）。

#### F. WPS · 公式（3×`writeFormulas`）· A 分支：中断 → 完整退出重开 → 核对 → 续跑只写第三步 → 确认

- **实机**：用户 2026-11-01 21:46–22:10 在真实 WPS 里走完，并按六步逐条确认 —— ①三个独立公式写入
  `结果1!B2 =来源!B2*1` / `结果2!B2 =来源!B4*2` / `结果3!B2 =来源!B5*3`；②第二步写入并回读后中断（前两格 2、8 且公式栏为公式，第三格为空）；
  ③保存并**完整退出** WPS，重开同一文件后恢复卡仍可用；④点核对后界面变"可继续剩余动作"，第三格仍空；
  ⑤显式点继续后 `结果3!B2=15` 且公式栏 `=来源!B5*3`，前两格仍 2、8 且公式正确；⑥点确认变更成功
  （原文："已接受当前结果并移除本次审阅标记；保存为工具与撤销本次执行可用"）。截图：**待定位**（全盘未找到 21:40 之后新建的图片；
  `outputs/probe/clarification-inbox.png` 是 11:49 的旧图，与本轮无关）。
- **文件只读**（本轮复核，未改动）：`C:\Users\jia\Desktop\测试\公式来源证据验收-WPS-A.xlsx`，11032B / mtime `11-01 22:15:45` /
  sha256 `ecf3a6ad…8a98bc8`；`结果1/2/3!B2` 公式文本分别 `=来源!B2*1` / `=来源!B4*2` / `=来源!B5*3`，缓存值 `2 / 8 / 15`；
  `来源!A1:B5` = `(10,1)…(50,5)`。**注意**：账本 `workbook_identity` 记的是小写 `…-wps-a.xlsx`，磁盘文件名是大写 `…-WPS-A.xlsx`，
  同一路径大小写归一，不是两个文件。
- **账本**（一致性只读快照 `mode=ro` + `Connection.backup()`，WAL 已并入，未改原库）：本轮唯一执行
  `wps-execution-304e6eee-5ba7-4131-b01a-7c54aaea4e32`（`wps rows today: 1`，不会混入此前登记失败轮或 10-28/10-29 旧轮）；
  `status=completed` / `revision=11` / `next_action_index=3` / `in_flight_action_index=NULL`；`created_at=21:46:09` / `updated_at=22:10:10`；
  `result_json={"actionCount":3,"backfilled":false}`；`checkpoint_json` 指向 `结果3!B2`、`afterFingerprint=v2:26:1562cd4c`（与材料 `actions[2].expectedFingerprint` 一致）；
  **`undone_at` / `abandoned_at` 均 `NULL`**；`recovery_package_*` 除 `expires_at=2026-11-08` 外全为 `NULL`。
- **写序（不是数值推断）**：`office_host_writes` 对本 op **恰好 3 行** —— `idx0` + `idx1` 同在 `21:46:09`
  （`idx0` `written`/`readback=verified`；`idx1` **`failed_after_write`** / `WPS_ACTION_FAILED` = 第二步写进去了但回执失败），
  `idx2` 在 **`22:10:10`**（`written`/`verified`）⇒ **续跑只执行第三步，未重写前两步**。
  依据是 `append_office_host_writes` **只追加、不改状态、不推进 revision**，`recorded_at` 由服务端取 `time.time()`，
  所以 24 分钟的时间差本身就是写序。**第二步回执补记**走 `reconcileWpsActionReceipt`，**不产生新的 host-writes 行** ——
  这正是"回执补记"与"第二次宿主写入"可区分的地方。
- **恢复材料**（只读）：`%LOCALAPPDATA%\Excel Bro\execution-recovery\wps-execution-304e6eee…\material.dat`，1914B，
  `digest=55c63b7b…`；三动作 `expectedFingerprint` = `v2:26:6b460f4f` / `v2:26:a013f678` / `v2:26:1562cd4c`；
  `sourceEvidence` = `来源!A1:B5`（`previewFingerprint = valueFingerprint = v2:146:43024e97`）。
- **账本证据不足（不得补推断）**：`planning_tasks` / `planning_task_events` 里**查不到** `task-a6b33fa7-f25b-45b5-a4ba-82fe39eb6252`，
  `folder_operation_events` 对本 op 计数为 **0** ⇒ 账本**没有**"动作登记 / 开始 / 回执 / 核对"的逐事件序列，
  **"核对（reconcile）"这一步没有独立事件行**；阶段划分只能由 `created_at`、`office_host_writes.recorded_at` 与 `result_json` 推出。
- **自动化**：修复轮 `wpsExecutionResume.test.ts` + `wpsNativeComparison.test.ts` 合计 110 passed；全量 core 1128 / addin 1734 / wps 1283 / server 2026。
- **剩余缺口**：截图待定位；「确认变更」**只有用户界面证据**，账本侧可见的只是 `completed + next_action_index=3`，
  **没有单独的确认留痕**（与 A 条 Excel 同理）；`undone_at=NULL` ⇒ "撤销入口可用"**不等于**本轮实际撤销通过，本轮**没有**撤销动作证据；
  **B 分支仍未验**；续跑期间过程面板文案未切换（见 §8.6）。

#### G. 仍然没做的（不许写成"已通过"）

- **WPS · 公式 · B 分支**：A 已通过（F 条），B 真机仍未走；应按 `FORMULA_SOURCE_RESUME_ACCEPTANCE_2026-10-30.md` §3–5 的 B 段
  改 `来源!B5` 后重开续跑，验"被拒 + 目标不变 + 撤销"。`WPS_ACCEPTANCE_FIXTURE_2026-10-31.md` 的九步清单是两表值复制终态验证，不能当公式 A/B 验收。
- Excel/WPS 的**格式动作**、**删表原地撤销**、并排审阅布局、审阅副本生命周期、WPS「工作目录」服务端链路。

⇒ 所以"**公式精确恢复真机通过**"现在覆盖到 **Excel A/B（含重开后撤销）** 与 **WPS A**；WPS **B 仍缺**。
说"恢复链路真机通过"必须带上宿主＋动作类型＋分支。整个 **A1-GATE 仍未通过**，不要用本节代替闸门结论。

### 8.6 已登记的展示缺陷：续跑进入审阅后，过程面板仍停在"回读第 2 个动作"（**本轮只记录、未修**）

- **复现**：WPS 公式 A 分支（F 条，`wps-execution-304e6eee…`）续跑第三步成功、界面已进入**结果审阅**后，
  上方过程区仍显示 **「回读核对第 2 个动作」/「进度 2/3 步」/「正在回读核对「结果2!B2」…」**，不会切到第 3 个动作、也不会切成审阅态。
- **影响**：状态自相矛盾 —— 结果已经要人审阅，过程却说还在回读上一格。用户按这行字会以为第三步没执行完，
  从而怀疑"继续"没生效（本轮 ④→⑤ 之间正是这种观感）。属**文案/条目累积**问题，不影响写入正确性与账本写序。
- **出处**（只读定位，未改）：`apps/wps-addin/src/wpsActivityLog.ts:126-150` —— `verifying` 分支把条目 id 记成
  `verifying:${index}` 且**续跑开始时不重置**累积，label 为 `回读核对第 ${index + 1} 个动作`，facts 为 `进度 ${index + 1}/${actionCount} 步`；
  行内文案来自 `apps/wps-addin/src/wpsPlanExecutor.ts:2578` 的 `正在回读核对「${action.sheetName}!${action.range}」…`。
- **未做**：本轮**不改产品代码、不重跑写入**（用户明确要求只登记）。

## 9. 复现命令

```bash
# 判据
cd packages/core && npx vitest run src/composerPrimaryAction.test.ts \
  src/currentActionCard.test.ts src/clarificationComposer.test.ts

# 组件与接线（Excel）
cd apps/excel-addin && npx vitest run src/components/ComposerActionDock.test.tsx \
  src/components/ComposerPrimaryButton.test.tsx src/components/ClarificationComposer.test.tsx \
  src/contracts.test.ts src/messageFirstRouting.test.tsx \
  src/excelWriteEndToEnd.test.tsx

# WPS 侧同款入口
cd apps/wps-addin && npx vitest run src/wpsPlanningPauseCard.test.tsx \
  src/wpsPlanningResumeClosure.test.tsx src/WpsAgentPanel.test.tsx

# 布局（真实 Chromium，不启动任何服务）
node scripts/check_process_panel_layout.mjs
node scripts/check_recovery_entry_viewport.mjs
```
