# 审阅副本生命周期 · 交付说明与验收（2026-10-14）

> **2026-10-15 更新**：本文下方记录包含历史停用方案。当前 Excel 兼容关闭已换成 v4 本机协议：获取并复核同一个 Workbook，先 SaveCopyAs 保留恢复文件，再 Close(False)，无 WM_CLOSE 回退。两端审阅卡已移除“重新排列”。最新文件处理、边界与实机步骤见 [开发说明](DEVELOPMENT.md) 的“审阅结束与副本处理（2026-10-15）”。本轮自动化：core 499、Excel 1266、WPS 847、服务端 1455 通过；两端构建通过。真实宿主窗口本轮未操作，关闭实机验收仍待执行。

> 本文件是**本轮交付物**：逐端说明审阅副本「实际创建了什么、能自动关闭什么、能删什么、何时降级、剩余限制」，
> 并把**替身自动化验收**与**真实客户端验收**分开报告。
> 机制细节见 `docs/ARCHITECTURE.md`「2026-10-14：审阅副本的生命周期」；操作步骤见 `docs/DEVELOPMENT.md`
> 「审阅副本生命周期验收（2026-10-14）」。

## 0. 排查结论：是能力缺口，不是文案问题

用户看到的现象是 Excel 兼容模式下「工作簿5…」这类审阅副本在**确认写回或放弃后仍留在屏幕上**。
追根因（不是改提示词能解决的）：

- `apps/excel-addin/src/excelNativeComparison.ts` 的 `closeExcelNativeComparison` 在 `manual` 分支
  **直接 `return false`**，注释写着 `// No API handle: never claim that the window was closed.`；
- `apps/excel-addin/src/hooks/useExcelNativeComparison.ts` 对 `manual` 分支**根本没有调用关闭**；
- Office.js 在兼容路径确实拿不到该副本的窗口句柄（`createWorkbook` 不返回窗口对象，
  `Application.Windows` 在此路径也没有可用的 per-window `Hwnd`）。

但「Office.js 这一条路不支持」**不等于**「所有实现路线都不可能」。服务端已有的本地桥
（`server/app/excel_window_layout.py`）本来就能按身份枚举两个 Excel 窗口 —— 于是走通的路线是
**让本地桥对这个身份已核实的副本窗口投递 `WM_CLOSE`**，由宿主自己决定是否弹保存提示。

## 1. 逐端：实际创建了什么

三种副本，语义完全不同，**不可互相代替**（唯一权威说法：`packages/core/src/reviewCopyLifecycle.ts`
的 `reviewCopyCreationSummary(kind)`）：

| 类型 | 触发路径 | 实际创建对象 | 有磁盘文件吗 | 有窗口句柄吗 |
| --- | --- | --- | --- | --- |
| `excel-native` | Excel Desktop 1.1 原生并排审阅 | `Excel.createWorkbook(base64)` 新建的**未保存**工作簿窗口 | **否** | 有（`Window.close()`） |
| `excel-compat` | Excel 兼容模式（本机生成的 OOXML artifact） | `Excel.createWorkbook(artifact.base64)` 新建的**未保存**工作簿 | **否** | Office.js **无** |
| `wps-native` | WPS 原生并排审阅 | `Worksheet.Copy()` 生成的**未保存**单表副本工作簿 | **否** | `Workbook.Close(false)` |

**身份与所有权判据**（三条都不能省）：

- 副本活动表名 = 本次随机 nonce `copySheetName = "EB_" + 20 hex`（协议侧 `^EB_[0-9a-f]{20}$`）；
- 宿主侧「未保存」标记 + 身份元数据（Excel 由桥读 `EXCEL7` / `OBJID_NATIVEOM`；WPS 是活对象引用，
  另加 `copyIdentity = Name/FullName/Path` 三元组复核）；
- **唯一性**：同一时刻必须「恰好一个未保存 nonce 副本」且「恰好一个已保存原表在场」。

**严禁**把下列当清理依据：「工作簿N」默认名、`EB_` 前缀、当前活动窗口、文件名相似度。
**严禁**把未保存工作簿当磁盘文件上报。

## 2. 逐端：能自动关闭什么

### Excel 原生模式（`excel-native`）
- **关闭前先过第三项判断**：逐格回读副本（`evidence` + 格式摘要）与本次写入后的证据比对。
  - 一致 ⇒ `unchanged`，继续关闭；
  - 不一致 ⇒ `extra_edits`，**不关副本**（用户审阅期间的人工修改不被静默丢弃）；
  - 读不到 ⇒ `edits_unverified`，**同样不关**（不许把"没读到"当成"应该没改"）。
- 关闭通过 Office.js `Window.close()`；返回 `false` ⇒ `close_unconfirmed`（**歧义**信号：被取消 / 被推迟 / 被拒；不是成功关闭，也不推断是保存提示）。
- 关闭前先确认该窗口仍在本会话的窗口集合里；不在 ⇒ `already_closed`（幂等成功）。

### Excel 兼容模式（`excel-compat`）

> **2026-10-15 更新：自动关闭＝未完成缺陷（不是已交付）。** 当前 `COMPAT_COPY_AUTO_CLOSE = false`，
> 兼容副本一律手动关闭；取消/写回之后一个窗口消息都不发（见 `ExcelFormatReviewFlow.test.tsx`）。
> 下面这组闸门是"恢复自动关闭**必须换成工作簿级关闭原语并真机验证**之后**才能满足什么**"，不是当前行为，
> 也不是"把开关改回 `true` 即可"。**根因与未闭环缺陷见 §11。**

- 唯一窄路：本地桥对**身份已核实**的该副本窗口 `PostMessageW(hwnd, WM_CLOSE, 0, 0)`。
- 发窗口前的必过闸门（缺一即 `identity_lost` / 拒绝动作）：
  1. 有**已保存**窗口的当前表带本次 nonce ⇒ 判定 `user_saved`（用户另存/替换过）⇒ **不动它**；
  2. 未保存 nonce 副本 > 1 ⇒ `ambiguous`；
  3. 未保存 nonce 副本 = 0 ⇒ `already_closed`；
  4. 已保存原表窗口 ≠ 1 ⇒ `identity_lost`；
  5. 动作前重新 `describe()` 副本与原表，身份对不上 ⇒ `identity_lost`；
  6. **安全闸**：同一时刻必须存在唯一已保存原表。若副本是最后一个可见工作簿，对它发 `WM_CLOSE`
     会关掉**整个 Excel 实例** —— 因此这条前置条件同时是安全条件。
- 投递后按 `closeWaitSeconds:3` / `closePollIntervalMs:200` 轮询窗口是否消失：
  - 消失 ⇒ `closed`（**这是"确认已关"的唯一来源**：真的是窗口不见了）；
  - 关闭慢、需要几轮轮询才消失 ⇒ 仍然 `closed`（轮询坚持到 deadline，不会看一眼就下结论）；
  - 到点仍可见 ⇒ `close_unconfirmed`（**不推断原因**：保存提示 / 别的模态框挡住 / 单纯慢，在这一层长得一样）。
- **兼容副本不做第三项判断**：Office.js 读不回来，所以这条路**从不宣称**"已确认没有额外修改"，
  只依赖宿主自己对未保存工作簿的关闭提示，并且**不替用户回答**任何提示。
- 明确**不引入**：VBA / 宏 / 任意脚本 / 键鼠注入 / `TerminateProcess` / 通用命令执行入口。

> **兼容副本自动关闭＝未完成缺陷（根因见 §11，当前停用）**
>
> 停用不是"没证明安全"那么轻，而是**机制本身不可靠**：关闭原语是
> `PostMessageW(hwnd, WM_CLOSE, 0, 0)`，作用于**顶层窗口框架**而非**工作簿对象**，
> 且 `request_close` 在投递瞬间只重核 `(class==XLMAIN, pid)`、不重核工作簿身份；
> 顶层窗口句柄与"本次副本工作簿"在 Excel 真实窗口拓扑（SDI / 共享框架 / 句柄回收）
> 下**不是 1:1 绑定**。用户实机见过**原表**的保存提示，正是这条路径打到原表的证据。
>
> 因此：`COMPAT_COPY_AUTO_CLOSE = false`（`apps/excel-addin/src/excelNativeComparison.ts`），
> `closeExcelNativeComparison` 的 `manual` 分支直接返回 `unsupported`，**取消/写回之后
> 一个窗口消息都不发**（不是"发了但失败"，是桥一次都没被调用）—— 当前这条危险路径
> 是停用的。桥实现（`excelWindowLayout.ts` / `excel_window_layout.py`）与收据映射
> （`closeCompatReviewCopy`）保留并单独测试，但**它们今天依赖的窗口原语就是缺陷本身**，
> 不能靠"保留 + 改回开关"当修复。
>
> **恢复前提（不是把常量改回 `true`，而是先换原语再真机取证）**：
> ① 把关闭原语从 `PostMessageW(WM_CLOSE, hwnd)` 换成**工作簿级**的
> `Workbook.Close()`（对 `_window_metadata` 里已取到的 `window.ActiveSheet.Parent`
> 这个 Workbook COM 对象调用）；② 在真实 Excel + 真实本机桥上取证"只关副本、
> 绝不关原表、绝不触发原表保存提示、绝不清除原表撤销"；③ 覆盖"副本是最后一个可见
> 工作簿 / 用户已另存副本 / 多候选"等边界。三项都满足后才把常量改回 `true`。

### WPS 原生模式（`wps-native`）
- **关闭前先过第三项判断**（`wpsCopyContentVerdict`，与写回侧共用同一套判据）：
  - 逐格回读一致 ⇒ `unchanged` ⇒ 调用 `Workbook.Close(false)`；
  - 有额外修改 ⇒ `extra_edits`，**不调用 `Close`**；
  - 读不到 ⇒ `edits_unverified`，**不调用 `Close`**。
- `Close` 不存在 ⇒ `unsupported`；返回 `false` ⇒ `close_unconfirmed`（歧义信号）；抛异常 ⇒ `failed`；成功 ⇒ `closed`。
- 关闭前重核 `copyIdentity`：读失败 ⇒ `already_closed`（用户自己关过）；不一致 ⇒ `user_saved`（那是用户自己的文件，不关不删）。
- **顺序不可颠倒**：身份（①）→ 未另存（②）→ 没有额外修改（③）。前两项通过**不能**跳过第三项。

### 降级路径（**不是**「所有途径都不行」）
| 情形 | 结果 | 呈现 |
| --- | --- | --- |
| 平台无关闭原语 | `unsupported` | 「副本未能自动关闭，请手动关闭」 |
| 旧服务端没有 `/close` 路由（404） | 前端按 `unsupported` | 同上（**不做静默成功**） |
| 桥不可达 / 抛错 | `unsupported` | 同上 |
| 投递失败 | `failed` | 同上，但 detail 说明「原表写入与验证结果不受影响」 |
| 身份无法确认 | `identity_lost` | 交还用户，不自动丢弃 |
| 用户已自己关闭 | `already_closed` | 「此前已由你关闭」（幂等成功，不是错误） |
| **副本里发现了额外人工修改** | `extra_edits` | 「副本里还有未经确认的人工修改，未自动关闭。」 |
| **副本读数拿不到，无法证明没被改** | `edits_unverified` | 「无法确认副本是否被额外修改，未自动关闭。」 |
| 请求已发出、到点窗口仍在 | `close_unconfirmed` | 「关闭尚未确认，请到副本窗口前确认。」（**不推断是保存提示**） |

### 2.1 专项核查一：分清「身份可信 / 未另存 / 没有额外人工修改」

**三项判断，前两项不能代替第三项。**

| # | 判断 | 判据 | 通过后能得出什么 |
| --- | --- | --- | --- |
| ① | 副本身份可信 | 本次 nonce + 未保存标记 + 唯一性 | "这是我创建的那一份" |
| ② | 没被另存/替换 | 身份三元组（Name/FullName/Path）一致 | "它没变成用户自己的磁盘文件" |
| ③ | **没有额外人工修改** | **逐格回读**副本 vs 我们写入后的最终快照 | "它还是我们写入后的样子" |

**①+② 成立推不出 ③**：用户完全可以在并排审阅期间往副本里敲字。唯一入口是
`describeCopyDiscardSafety(verdict)`（`packages/core/src/reviewCopyLifecycle.ts`）：

| verdict | 结果 | 行为 |
| --- | --- | --- |
| `unchanged` | `null` | 允许关闭（可无提示丢弃） |
| `changed` | `extra_edits` | **不关副本**；改动留在副本里，交回用户 |
| `unknown` | `edits_unverified` | **不关副本**；"读不到" ≠ "没改过" |

宿主适配：
- **WPS**：`wpsCopyContentVerdict(staged)`（`apps/wps-addin/src/wpsNativeComparison.ts`）逐格比
  `copyFingerprints` / `copyAfter` / `copyFormatFinal`；`discardWpsNativeComparison` 顺序为
  身份 → 内容 → `Close(false)`。**同一个函数**也被写回侧 `verifyWpsNativeComparison` 使用，
  两条路不会一个说"没改"一个说"改了"。
- **Excel 原生**：`closeExcelNativeComparison`（`apps/excel-addin/src/excelNativeComparison.ts`）在
  `Window.close()` 前用独立一轮 `Excel.run` 回读 `copyEvidence` + 格式摘要；读失败落在 `catch` ⇒ `unknown`。
- **Excel 兼容**：无法回读 ⇒ 不做 ③，也不宣称做了；只靠宿主对未保存工作簿的关闭提示。

**「副本人工修改不导入原表」**是既有性质（写回只提交已审阅的计划）：本次为它补了**两端**的断言 ——
副本被改后写回必须被阻断，且原表一格未写。

### 2.2 专项核查二：关闭请求已发出但窗口仍在 ⇒ 不推断保存提示

- 本机桥只看得到"窗口还在不在"，**没有任何对话框证据**；关得慢、被别的模态框挡住、宿主在弹保存提示，
  在这个层面**无法区分**。因此到点仍在只报 `close_unconfirmed`（「关闭尚未确认」），
  文案里**明确列举**三种可能，而不是断言其中一种。
- 回执**只有三个字段**（`status` / `layoutVersion` / `copyClosed`）—— 连"保存提示"都无处表达。
- 四种结论互不重合且各自可辨：`closed`（成功）/ `close_unconfirmed`（待确认）/ `failed`（失败）/
  `already_closed`（此前已关闭，幂等成功）。清理状态**不改变原表业务结论**：
  `reviewCopyCleanupState` 只为 `verified` / `discarded` 两种业务结果选文案，关闭结果永远插在后面。
- **删掉了两个没人能产生的状态**：`prompt_pending`（改名 `close_unconfirmed`，去掉推断）与
  `user_cancelled`（没有任何适配器能观察到"用户点了取消"；`Window.close() === false` 是歧义信号）。
- 测试覆盖：**关闭延迟**（要几轮轮询才消失 ⇒ 仍算 `closed`）与**非保存对话框阻塞**
  （窗口一直不消失 ⇒ 与"可能是保存提示"**必须得到同一个结论**）。

## 3. 逐端：能删除什么（磁盘）

**结论：当前没有任何磁盘临时文件可删。**

证据（可复核）：`server/app`、两端加载项、`packages/core` 中 **0 处** `tempfile` / `NamedTemporaryFile` /
`mkdtemp` / `TemporaryDirectory`；唯一 `TEMP`/`TMP` 命中在 `python_docker_runner.py` 的 CLI 环境白名单。
三种副本都是**宿主内存里的未保存工作簿**。

因此：`tempFilePath === null` ⇒ 只允许结论 `none` ⇒ 面板只显示「**本次没有需要清理的磁盘文件。**」，
**绝不**出现「已删除临时文件」。

若将来确有临时文件，唯一前置判据是 `validateTempFilePath(root, candidate)`：
只接受本机盘符绝对路径、且必须落在专用临时目录内（**拒绝**相对路径 / UNC / `..` 越界 / 根目录本身 / 控制字符）。
`planTempFileCleanup`：未登记或越界 ⇒ `out_of_scope`；被占用 ⇒ `in_use`（**保留待清理，绝不强杀宿主**）；否则 `deleted`。
**不扫描/不删除用户目录，不按 `EB_` 前缀批量删。**

## 4. 结束时的处理规则（六条，双端一致）

1. **写回并验证成功** ⇒ 尝试关闭本副本（+ 清理**明确属于本流程**的临时文件）⇒ `closed`。
2. **放弃变更** ⇒ 原表零写入 ⇒ **终态**：确认框与三个审阅按钮立刻撤下 ⇒ 尝试关闭/清理本副本 ⇒ `discarded`。
3. **写回失败 / 部分完成 / 结果不确定** ⇒ **默认保留副本**供排查（`RETAIN_MESSAGES`，
   关闭结果强制 `not_attempted`），明说原因，**不自动销毁排障材料**。
4. **用户另存 / 替换过副本**（身份变了、未保存标记消失）⇒ `user_saved`，**不自动丢弃**，交还用户。
5. **原表与用户自己保存的文件永不被关闭或删除**。
6. **关闭失败绝不改写业务结论**：业务结论（标题＋说明）只由 `outcome` 决定 ——「已写回原表」/「已放弃变更」；
   副本关闭结果只作**次要**的独立状态跟在后面。`write_failed`（守卫在写入前拦下）与
   `partial_or_unknown`（执行器跑到一半没确认）必须说成两句**不同**的话，不许互相冒充，也不许套成功模板。

## 5. 界面与布局

- **业务结论与副本状态分开、且由 `outcome` 单向派生**（`packages/core/src/reviewCardPresentation.ts`
  是"哪个阶段说什么、哪条信息放在哪一层"的唯一来源，双端共用）：
  - 审阅中首屏只有五项：标题「请核对审阅副本」/「左侧是审阅副本，右侧是原表。原表尚未修改。」/
    范围行 / 兼容模式必要说明 / 确认框「我已核对副本，确认按此计划修改原表。」；按钮
    「重新排列」「放弃变更」「确认写回原表」（未勾选时写回入口禁用）。
    其余（重启提示、随机副本标识、关闭诊断、磁盘清理）全部下沉到「诊断详情」，技术解释下沉到「兼容模式说明」。
  - 写回成功：标题「已写回原表」＋说明；副本关闭状态**另起一处独立**显示「审阅副本已关闭。」（**不带**业务前缀）。
  - 已放弃：标题「已放弃变更」＋「本次计划未写入原表。」；副本仍打开时另给手动顺序
    （先另存再关），**不**无条件让用户选「不保存」。
  - 写回失败：「写回未完成」（也可能一格未改）/「写回结果不确定」（部分动作已经执行），两档互不冒充。
- **`reviewCardClaimsSuccess()` 是界面套成功模板的唯一许可闸**：只有 `verified` + `verification === 'passed'` 为真。
  文案一律纯文本，`reviewCardMarkdownLeaks()` 是渲染前最后一道闸（`**` 会原样显示星号）。
- **写回闸门只有一份判据且只读实时事实**（`wpsWriteBackBlockReason`，反向变异证据见 §7.1 第 1 条）。
- **防双击**：清理期间按钮禁用；失败给**重试**或手动指引（可重试结果保留现场）。
  写回侧另有"一次写回在跑"的同步闸，重复派发写不了两次。
- 按钮文案随阶段变：审阅中 ⇒「重新排列」「放弃变更」「确认写回原表」；
  写回成功且副本没收尾 ⇒「关闭审阅副本」；写回失败/不确定 ⇒「保留副本并结束审阅」；已放弃 ⇒ **没有**次要动作。
- **布局策略保持上一轮成果**：安全恢复**本流程改过**的布局，保留中途的用户手动调整；
  副本清理**不触碰布局状态机**、不改审阅阶段，也不影响已完成的写回/撤销。

## 6. 协议与共享语义

- 共享状态机：`packages/core/src/reviewCopyLifecycle.ts`
  （`REVIEW_COPY_LIFECYCLE`、`ReviewCopyKind`、`ReviewCopyOwnership`、`ReviewCopyCloseOutcome`(**11 态**)、
  `ReviewCopyContentVerdict`(3 态)、`ReviewCopyBusinessOutcome`(4 态)、`TempFileCleanupOutcome`(5 态)、
  `reviewCopyCleanupState`、`reviewCopyCreationSummary`、`describeCopyDiscardSafety`、
  `validateTempFilePath`、`planTempFileCleanup`）。
  关键不变量：**未知关闭态一律折 `failed`（绝不折 `closed`）**；**写回失败/不确定 ⇒ `not_attempted`**；
  **磁盘清理只在 `closed | already_closed | close_unconfirmed` 才允许**（否则 `deleted` 降级为 `out_of_scope`）；
  **`changed` 与 `unknown` 分成两档**（`extra_edits` / `edits_unverified`），都不允许静默丢弃。
- 限额：`config/capabilities.json` → `reviewCopyLifecycle` =
  **`{ closeWaitSeconds: 3, closePollIntervalMs: 200 }`** —— **只有这两项**，且都是服务端
  `close_copy` 真正读的（关闭后轮询窗口是否消失）。`maxCloseAttempts` / `tempCleanupWaitSeconds` /
  `tempRootLabel` **没有任何消费点，已删除**：本轮**没有**自动重试循环（重试由用户再点一次驱动），
  也**不**声称已实现临时文件清理等待。
- 协议（`layoutVersion: 3`，两端同步）：
  - `contracts.ts` ↔ `server/app/models.py` 双写；
  - 请求 `ExcelReviewCopyCloseRequest` 只收 `workbookPathHash` / `worksheetNameHash` / `copySheetName`，
    `extra="forbid"` —— **绝不收路径、计划、值、窗口句柄**；
  - 回执 `ExcelReviewCopyCloseResult` **只有三个字段** `status` / `layoutVersion` / `copyClosed`；
    刻意**没有**任何原表写回字段 ⇒「关闭失败不能把写回成功改成执行失败」在协议层就无处表达；
  - 路由 `POST /api/office/review-layout/close`（沿用 origin / Host / Excel Origin 守卫）；
  - 前端 `api.closeExcelReviewCopy` 显式投影 + 校验键数与状态白名单；`copyClosed` 与 `status`
    自相矛盾（说关掉了但 `copyClosed=false`）一律按 `failed`。

## 7. 自动化验收结果（替身证据 · 全绿，本轮复跑）

| 项目 | 结果 |
| --- | --- |
| `npm run test:core` | **461 passed / 461**（27 文件；含 `reviewCopyLifecycle.test.ts` **33 例**） |
| `npm run test:addin` | **1239 passed / 3 skipped**（103 passed + 2 skipped 文件） |
| `npm run test:wps` | **833 passed / 3 skipped**（52 passed + 2 skipped 文件） |
| `pytest server/tests/test_excel_window_layout.py` | **94 passed**（含关闭延迟与非保存对话框阻塞两例） |
| `pytest server/tests`（全量） | **1413 passed / 28 skipped / 2 failed** |
| `npm run build:addin` / `npm run build:wps` | 退出码 **0**（均含 `tsc --noEmit`） |
| `scripts/mutate_review_copy_lifecycle.mjs` | **13 / 13 SEED 全 RED**，还原后逐字节 sha256 一致 |
| `scripts/check_excel_layout_guards.py` | **13 项守卫全被抓**（新增 `close_copy` 未打散旧 SEED） |
| `scripts/mutate_excel_format_compat.mjs check` | **18 个 SEED 在位** |
| `git diff --check` | 本任务改动文件**干净** |

- **那 2 个 failed 是环境阻断，不是代码缺陷**：`test_python_event_delivery.py` 两例断言
  `'unavailable' == 'completed'`，因为本机 `docker version` 返回 **rc=1**（客户端在、守护引擎不可用），
  `python_analysis` 按 fail-closed 返回 `unavailable`。与本轮改动无关。
- 变异覆盖的机制（每个 SEED 都是**拆掉被测机制、目标用例必须变红**）：
  S1 写回失败仍保留副本；S2 未知关闭态 ⇒ `closed`；S3 磁盘清理越界；S4 未登记文件声称已删除；
  S5 服务端丢掉审阅上下文校验；S6 `WM_CLOSE` 失败谎报 `closed`；S7 Excel 忽略 `copyClosed`；
  S8 WPS 丢掉身份复核；S9 WPS catch ⇒ `closed`；S10 Excel 钩子在 `blocked` 阶段也关副本；
  **S11 WPS 放弃路径丢掉"副本是否被额外修改"这一闸**；**S12 Excel 原生关闭前同样丢掉这一闸**；
  **S13 把"窗口还在"直接解释成"宿主正在询问是否保存"**。
  （S4 首次**没红** —— 种子只翻 `planTempFileCleanup` 的 `null` 分支，我的断言只看了
  `reviewCopyCleanupState` 的标题；补 `expect(planTempFileCleanup({tempFilePath:null,...})).toBe('none')`
  与正向 `'deleted'` 锚点后才变红。**不红先怀疑断言**。S8 在把 `closeCopyWithoutSaving` 拆成
  `copyOwnershipVerdict` + `requestCopyClose` 之后 SEED 落空（`命中 0 次`）—— apply 的**前置闸**
  如实把它拦成 `APPLY-FAIL` 而不是静默跑绿，这正是"源码重写会打散旧 SEED"该有的表现。）
- `git diff --check` 唯一的非零项是 `apps/excel-addin/src/dataTools.ts:1102: new blank line at EOF.`
  —— 该文件属**上一个会话遗留的未提交改动**，与本轮无关，未改动。

### 7.1 同日第二轮：审阅卡文案与信息层级重构（**不改执行/撤销算法**）

把"哪个阶段说什么、哪条信息放在哪一层"从两端的 JSX 里抽成共享视图模型
`packages/core/src/reviewCardPresentation.ts`（四阶段 = `reviewing | discarded | committed | failed`，
唯一来源是 `ReviewOutcome`），并逐阶段验收界面。复跑结果：

| 项目 | 结果 |
| --- | --- |
| `npm run test:core` | **478 passed / 478**（28 文件；新增 `reviewCardPresentation.test.ts` **17 例**） |
| `npm run test:addin` | **1245 passed / 3 skipped**（103 passed + 2 skipped 文件） |
| `npm run test:wps` | **837 passed / 3 skipped**（52 passed + 2 skipped 文件） |
| `npm run build:addin` / `npm run build:wps` | 退出码 **0**（均含 `tsc --noEmit`） |
| `scripts/mutate_excel_format_compat.mjs check` | **18 个 SEED 在位** |
| `scripts/mutate_review_copy_lifecycle.mjs check` | **13 个 SEED 在位（各命中 1 次）** —— 本轮新加的只读自检 |

**本轮又打散了两个变异 SEED，已修回并各自复跑**（源码重写会打散旧 SEED，这条铁律再次生效）：
- **S10** 原种子写的是 `if (current.phase === "blocked")`，`phase` 已被 `outcome` 取代 ⇒ 落空。
  改为 `current.outcome === "write_failed" || === "partial_or_unknown"`。
- **S13** 原种子含 `加载项**没有**对话框证据`，本轮剥掉 Markdown 记号后 ⇒ 落空。按纯文本重写种子与变异体。
- 两个都复跑过 **apply → 目标用例实跑变红 → restore 逐字节 sha256 一致**：
  S13 `Tests 1 failed | 32 passed (33)`；S10 `Tests 2 failed | 24 passed (26)`。
- **注意 S10 的第一版变异体漏了闭合花括号** ⇒ 整个文件采集失败、vitest 报 `Tests no tests`。
  这不是"目标用例变红"。四判据的第 ②③ 条（有效摘要 + 实跑预期用例）正是用来挡这种假红的。

**本轮新增的三条硬证据（用户点名）**：

1. **放弃后旧按钮与旧回调不可写回**。取放弃**之前**那个闭包（DOM 节点上的 `__reactProps$…onClick`）
   再调一次，原表一格都不许动。这条用例在修复前是**红**的 ——
   `AssertionError: expected 1 to be 0`：写回闸门当时读的是渲染期闭包变量，
   旧闭包里的 `pendingPlan` / `stagedReview` 还是放弃之前的值，"计划还在"是**过期事实**。
   修复：闸门抽成唯一的纯判据 `wpsWriteBackBlockReason`（`no_review | busy | editing_plan | not_acknowledged`），
   **只读每次渲染刷新的实时引用**；授权检查（异步）返回后**再判一次**；
   另加"一次写回在跑"的同步闸（`busy` 是渲染期派生值，同一 tick 的第二次调用看不到它）。
   修完两条断言（红→绿）同时在场：红的是"旧回调写了 1 格"，绿的是"实时事实放行 / 终态拒绝"。
2. **清理失败不恢复审阅状态**。`close_unconfirmed` / `failed` / `unsupported` / `identity_lost` /
   `ambiguous` / `user_saved` / `extra_edits` / `edits_unverified` 八种没收尾结果逐个断言
   确认框与三个审阅按钮**都不许回来**（放弃是终态），副本状态只作**次要**行，
   且磁盘清理行不再出现在放弃场景。
3. **不同结果不混用文案**。两端各有一组显式命名的**四阶段界面**用例
   （`describe("审阅卡四阶段界面（审阅中 / 已放弃 / 写回成功 / 写回失败）")`），
   逐阶段验收首屏与折叠层的分工；四档互相 `not.toContain` 对方的关键句
   （成功不许出现「已放弃变更」/「本次计划未写入原表。」；`write_failed` 与 `partial_or_unknown` 互不冒充；
   取消框与写回入口在终态一律 `queryByText(...) === null`）。

顺带修掉的三处冗余/推断文案（已从生产代码删除）：放弃场景的「原表写入与验证结果不受影响」、
「本次没有需要清理的磁盘文件。」（下沉到诊断层）、`unsupported` 里无条件提示用户选「不保存」
（改为"并自行决定是否保留其中的内容"）。

## 8. 真实客户端验收（**未执行** —— 与替身自动化分开）

**本轮未操作真实 Excel / WPS 窗口。** 「副本被自动关掉、宿主保存提示出现、用户选不保存后窗口消失」
这条链路**必须在真机上跑一遍才算闭环**。分档如实：

| 档 | 状态 |
| --- | --- |
| ① 工件回放（`test:cross-process`） | 未执行（本机 Docker 不可用，消费侧拒绝替身，退出码 2） |
| ② 双端在线连续集成（`test:online`） | 未执行（同上） |
| ③ 真实 Docker | **环境阻断** |
| ④ 真实模型自主规划 | 未执行（SDK 规划期并发 HTTP 冻死 `claude.exe` 流） |
| ⑤ 真实 Excel（含本轮副本关闭） | **未执行** |
| ⑥ 真实 WPS（含本轮副本关闭） | **未执行** |

最短验收步骤已写入 `docs/DEVELOPMENT.md`「审阅副本生命周期验收（2026-10-14）」：
Excel **9 步**（写回后两段陈述 / 关闭未确认不推断保存提示 / 放弃路径原表零写入 /
**副本被人工修改后放弃 ⇒ 不关副本、改动不丢** / **副本被人工修改后写回 ⇒ 阻断且不导入** /
写回失败保留副本 / 另存干扰不自动丢弃 / 同名工作簿 + 切活动窗口不误关 / 清理后仍可撤销），
WPS **4 步**（正常关闭 / **副本被人工修改 ⇒ 放弃与写回都不得静默丢失或导入** / 关闭被拒只报"尚未确认" / 另存不自动丢弃）。
**注意**：Excel 侧需**重启 Python 服务**（旧服务没有 `/close` 路由，会按 `unsupported` 降级为手动关闭）；
WPS 侧只需刷新加载项。

## 9. 剩余限制 / 未闭环（如实）

1. **真实 Excel 2024 与 WPS 的副本关闭链路未实机验证**（见上一节）。当前只能说「已实现且替身自动化可证」。
   - **Excel 兼容副本自动关闭＝未完成缺陷（根因见 §11）**：关闭原语是面向窗口的 `WM_CLOSE`，
     不绑定副本 Workbook 对象，投递瞬间不重核工作簿身份，实机曾打到原表保存提示。当前靠
     `COMPAT_COPY_AUTO_CLOSE = false` 停用危险路径；恢复必须换成工作簿级 `Workbook.Close()`
     并完成真实 Excel 验证，不得仅把开关改回 `true`。
2. **兼容副本仍不能回读单元格**：本轮只解决了「关闭」，跨实例读值仍未实现
   （`docs/ARCHITECTURE.md` 已改成**不再宣称「不能关闭」**，但明说回读未实现）。因此兼容模式**不做**第三项判断，
   只依赖宿主自己的关闭提示。
3. **`extra_edits` / `edits_unverified` 是设计而非缺陷**：副本里有未经确认的人工修改时，
   加载项**故意不关**副本 —— 静默丢弃别人改动比"留个窗口"糟得多。处理完（或手动关闭）后再点一次即可收尾。
4. **`close_unconfirmed` 是设计而非缺陷**：关闭请求发出后窗口仍在，加载项**没有**对话框证据，
   所以只说"尚未确认"并请用户查看窗口；**不**推断是保存提示。用户在窗口上处理后再点一次即可。
5. **Excel 侧依赖本地桥**：桥不可达或服务端过旧 ⇒ `unsupported` ⇒ 降级手动关闭（已有测试覆盖，非「可能」）。
6. **没有自动重试循环**：关闭失败/待确认后由**用户**再点一次（可重试结果会保留审阅卡片/现场）。
   `maxCloseAttempts`、`tempCleanupWaitSeconds`、`tempRootLabel` 三个没有消费点的限额**已从配置删除**，
   不再有任何"看起来已实现"的余地。
7. 未提交、未推送、未发布；未改业务工作簿；未回归既有审阅 / 写回 / 撤销。

## 11. 兼容副本自动关闭：根因与未闭环缺陷（2026-10-15）

> 本轮用户明确要求：把"Excel 兼容副本自动关闭"列为**未完成缺陷**，**不能**把停用开关
> （`COMPAT_COPY_AUTO_CLOSE = false`）当最终交付；先查清"请求关闭副本却弹出原表
> `格式验收.xlsx` 保存提示"的真实根因，再记录关闭对象与工作簿的真实对应关系，恢复自动关闭
> 前必须根因修复 + 真实 Excel 验证，不得仅把开关改回 `true`。

### 11.1 真实根因（为什么不关副本、却弹出原表保存提示）

关闭原语是 `PostMessageW(hwnd, WM_CLOSE, 0, 0)`，落在 `server/app/excel_window_layout.py`
的 `WindowsExcelWindows.request_close`（第 250–264 行）。投递目标是 `XLMAIN` 顶层窗口句柄。
但"顶层窗口句柄"与"本次副本工作簿"在 Excel 里**不是 1:1 绑定**，而这条路径**从头到尾没有
拿到副本的 Workbook COM 对象**——它只调了 `WM_CLOSE`，从没调过 `Workbook.Close()`。三个
具体缺口：

1. **关闭作用于窗口，不作用于文档。** `close_copy`（第 115–168 行）按启发式
   `ActiveSheet.Name == copySheetName（nonce） + unsaved + 同进程`（`describe` 读
   `window.ActiveSheet.Name`）挑出"一个顶层窗口"，再对这个窗口发 `WM_CLOSE`。`WM_CLOSE`
   是 Windows 给窗口**框架**的消息，Excel 自己决定这个框架的 `WM_CLOSE` 会影响哪些工作簿。
   只要框架与实际工作簿的归属不是 1:1（共享框架 / MDI / 同进程内句柄被回收复用 / 用户在投递前
   切换了该框架里的活动工作簿），`WM_CLOSE` 就可能落到原表所在的工作簿 ⇒ 原表保存提示。
   这正是用户实机见过的现象。

2. **投递瞬间没有重核"工作簿身份"。** `close_copy` 在发消息前会 `describe` 复核副本
   （`if native.describe(copy.key) != copy`，第 154–156 行），这一步对"活动表名变了"是严的；
   但真正 `PostMessage` 的 `request_close` 只重核 `class == XLMAIN` 且 `pid` 相同
   （`_is_excel` 仅比 class+pid，第 226–248 行），**不重核工作簿身份**。从 `describe` 复核
   到 `PostMessage` 之间有一个 TOCTOU 窗口：同进程、同类名的窗口句柄一旦被 Excel 回收 /
   重新指派给原表框架，消息就会发到原表。`close_copy` 的复核拦的是"活动表名变了"，
   拦不住"句柄被原表复用"。现有 `test_close_reverifies_identity_immediately_before_posting`
   只覆盖"同一 hwnd 的活动表名被换成 `Unrelated`"——它恰好被 `close_copy` 的 `describe`
   复核拦下；它**没有**覆盖"hwnd 被原表复用、而活动表名仍等于 nonce"或"句柄在 describe 与
   PostMessage 之间易主"这类**真正会打到原表**的路径。

3. **兼容副本仍不能回读单元格**（§9 第 2 条）：这条路从不能做"副本有没有额外人工修改"的
   第三项判断，只能赌宿主自己的保存提示；而宿主提示一旦被误发到原表，就是不可逆的代价。

### 11.2 关闭对象 ↔ 工作簿的真实对应关系（今天实际是什么）

今天**不是**"副本工作簿对象 ↔ 关闭动作"的直接绑定，而是
**"顶层窗口 hwnd ↔ 启发式（活动表名 == nonce + 未保存 + pid）"**。这个映射证明的是
"该框架的活动表是 nonce 表"，**不是**"该框架背后的工作簿对象就是本次审阅副本、且只有它一个"。
因此恢复自动关闭**不能**靠"活动窗口 / 标题 / 文件名 / `EB_` 前缀"——这些在根因里正是不可靠的
来源；必须换成**工作簿级**的关闭原语（见 11.4）。

### 11.3 五项目标 vs 当前窗口原语（不可靠）

| # | 目标 | 当前窗口原语能否可靠满足 |
| --- | --- | --- |
| ① | 写回验证成功后安全关闭本次副本 | 不能保证"只关副本"，不可靠 |
| ② | 放弃后原表零写入 + 安全关副本 + 不向原表发消息 | 当前停用路径已满足（不发任何消息）；但恢复后的窗口原语仍可能误发原表 |
| ③ | 写回失败 / 不确定保留副本 | 写回侧 `not_attempted` 已不发起关闭，满足 |
| ④ | 副本有额外修改 / 无法判断时让用户明确决定 | 兼容模式读不回副本，无法独立判断；窗口原语依赖宿主提示，若误发到原表则违背"让用户决定副本命运" |
| ⑤ | 不关原表 / 不触发原表保存提示 / 不清除原表撤销 | **窗口原语无法保证**（根因）；只有工作簿级原语能保证 |

### 11.4 技术阻断与可行方案

**阻断**：`WM_CLOSE` 窗口原语在 Excel 真实窗口拓扑下无法在代码层证明"只命中副本、绝不命中原表"；
本环境 Docker 隔离引擎不可用，无法在真实 Excel + 真实本机桥上取证（见 MEMORY「环境阻断」）。

**可行方案（可靠且可在对象模型层证明安全）**：

- 不再对顶层窗口发 `WM_CLOSE`，改为经本机桥已有的 `EXCEL7` / `OBJID_NATIVEOM` 访问拿到副本的
  **Workbook COM 对象**——`_window_metadata`（第 321–328 行）里已经取到 `book = sheet.Parent`，
  只是当时只读了元数据就丢了。新增一个关闭方法：对那个 Workbook 对象调用 `book.Close()`。
  - 不带 `SaveChanges` 参数 ⇒ 宿主仍对未保存副本弹"是否保存"，保留用户决定副本命运的权利
    （满足 ④）；
  - 或带 `SaveChanges=False` 强制丢弃（取舍待定，影响"副本有额外修改时要不要先问用户"）。
- 因为 `Close` 作用在**那个特定的 Workbook COM 对象**上，它在对象模型层就**只可能关这一个
  工作簿**，与框架拓扑 / 句柄是否回收无关，也绝不可能关到原表、绝不触发原表保存提示、
  绝不清除原表撤销记录（满足 ①⑤）。这条路径仍能复用 `close_copy` 现有的
  `user_saved` / `ambiguous` / `identity_lost` 闸门做身份前置。
- 切换后仍需真实 Excel 验证：副本是最后一个可见工作簿 / 用户已另存副本 / 多候选 等边界都
  **只动副本**。验证通过前 `COMPAT_COPY_AUTO_CLOSE` 保持 `false`，**绝不**把开关改回 `true`。

> 纪律：本缺陷**不以开关、文案或替身测试宣称已解决**。替身自动化（core / addin / wps / 服务端）
> 全绿只能证明"代码按当前逻辑跑通"，不能证明"实机上只命中副本窗口"。实机验证是闭环的唯一证据。

## 10. 本轮变更文件

**共享语义与配置**
- 新增 `packages/core/src/reviewCopyLifecycle.ts`、`packages/core/src/reviewCopyLifecycle.test.ts`
- `config/capabilities.json`（新 `reviewCopyLifecycle` 块；专项核查后**删掉**三个无消费点的键）

**服务端（本地桥 + 协议 + 路由）**
- `server/app/models.py`（`ExcelReviewCopyCloseRequest` / `ExcelReviewCopyCloseResult`；`status` 用 `close_unconfirmed`）
- `server/app/excel_window_layout.py`（`ExcelWindowLayout.close_copy`、`WindowsExcelWindows.request_close`）
- `server/app/main.py`（`POST /api/office/review-layout/close`）
- `server/tests/test_excel_window_layout.py`（新增 12 个关闭用例：含**关闭延迟**与**非保存对话框阻塞**）

**Excel 加载项**
- `src/contracts.ts`、`src/api.ts`、`src/excelWindowLayout.ts`、`src/excelNativeComparison.ts`
  （关闭前新增第三项判断：独立一轮回读副本证据）
- `src/hooks/useExcelNativeComparison.ts`
- `src/components/ExcelNativeReviewActions.tsx`
- 测试：`excelWindowLayout.test.ts`、`excelNativeComparison.test.ts`（+4 例第三项判断）、
  `components/ExcelNativeReviewActions.test.tsx`、`excelFormatCopyReview.test.ts`、
  `ExcelFormatReviewFlow.test.tsx`、`excelWriteEndToEnd.test.tsx`

**WPS 加载项**
- `src/wpsNativeComparison.ts`（拆出 `copyOwnershipVerdict` / `wpsCopyContentVerdict` / `requestCopyClose`）
- `src/WpsAgentPanel.tsx`
- 测试：`wpsNativeComparison.test.ts`（+5 例）、`WpsAgentPanel.test.tsx`

**脚本与文档**
- 新增 `scripts/mutate_review_copy_lifecycle.mjs`（**13** SEED，只 apply/restore）
- `docs/ARCHITECTURE.md`、`docs/DEVELOPMENT.md`、本文件

**同日第二轮（审阅卡文案与信息层级，见 §7.1）**
- 新增 `packages/core/src/reviewCardPresentation.ts` + `reviewCardPresentation.test.ts`（17 例）
- `packages/core/src/reviewCopyLifecycle.ts`（`copy` 拆出 `statusText` / `tone`，剥掉 Markdown 星号，删两处冗余/推断文案）
- `apps/excel-addin/src/hooks/useExcelNativeComparison.ts`（`ExcelReview['outcome']` 改用 `ReviewOutcome`；
  `save()` 对非 `reviewing` 一律拒；守卫前后分档 `write_failed` / `partial_or_unknown`）
- `apps/excel-addin/src/components/ExcelNativeReviewActions.tsx`（按 `card` 渲染；三段折叠；
  `ExcelNativeReviewActions.test.tsx` +4 例四阶段验收）
- `apps/wps-addin/src/WpsAgentPanel.tsx`（`reviewAcknowledged` 真闸；
  **`wpsWriteBackBlockReason` 唯一闸门 + 实时引用 + 在途闸**；终态撤卡；
  `WpsAgentPanel.test.tsx` +4 例四阶段验收 + 8 态 `it.each` 改为终态断言）
- 其余受影响的 WPS 用例统一补"先勾选再写回"：`WpsAgentPanelFormatReview.test.tsx`、
  `wpsOnlineLink.test.tsx`、`wpsFormatRecoveryLedger.test.tsx`、`wpsContinuousPythonChain.test.tsx`
