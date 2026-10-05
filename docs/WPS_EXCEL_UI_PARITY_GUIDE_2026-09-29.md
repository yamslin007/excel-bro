# WPS / Excel EB 界面对齐：下一批实施与验收指导

- 日期：2026-09-29
- 范围：**任务窗格 UI 与可见交互**。完整动作能力、耐久恢复、规划 SSE 和 `EB()` 的实施仍以 [能力对齐总指南](./WPS_EXCEL_CAPABILITY_PARITY_GUIDE_2026-09-28.md) 为准。
- 性质：待执行清单，**不是完成证明**。不要把自动化绿灯、相似截图或开发版写入闸当作双宿主实机验收。

## 1. 基线与比较方法

截图中的 WPS 是写入完成后的对话，Excel 是普通问候后的过程卡；两张图处于**不同任务状态**，不能据此判断“WPS 缺过程卡”。代码中 WPS 与 Excel 已共用 `TaskProgressTimeline` 和 `MessageCopyButton`。先用同一任务、同一模型连接状态、同一授权范围与同一窗格 CSS 宽度取得成对截图，再判定差异。

当前代码入口：

| 区域 | Excel 基线 | WPS 实现 | 现阶段结论 |
| --- | --- | --- | --- |
| 产品壳与响应式按钮 | `apps/excel-addin/src/App.tsx`、`styles.css` | `apps/wps-addin/src/WpsAgentPanel.tsx` 先加载 Excel 样式，再加载本端 `styles.css` | 同一按钮在窄窗隐藏文字是共用响应式规则；不得仅凭不同宽度截图改为强制显示，需核对实际 CSS 宽度与可用空间 |
| 消息复制 | 共用 `components/MessageCopyButton.tsx` | 同一组件，WPS 保留 WebView 剪贴板降级 | 已有代码与测试；仍需真实 WPS 鼠标、键盘、触屏/非悬停设备验收 |
| 执行过程 | 共用 `components/TaskProgressTimeline.tsx` | 本地读写事件映射到同一组件 | 展示组件共用，不代表事件来源等价；WPS 服务端规划 SSE/补发尚未接线 |
| 写入计划 | `App.tsx` 的 `inline-plan`、详情、假设/警告/动作、编辑入口 | `WpsAgentPanel.tsx` 的 `wps-plan-card`、值样本、警告和 JSON 编辑 | **可见信息与层级需逐项对照**；WPS 不能因只支持少数动作就省略计划风险、前提和授权说明 |
| 验收与撤销 | `verification-card` 根据逐项检查显示通过、未独立验证或失败 | `wps-verification-card` 当前显示动作数、单元格数和目标 | **信息粒度不同**。只能展示执行器真实提供的回读证据，不能在 UI 中虚构 Excel 式检查项或跨关闭恢复 |
| 输入区与历史 | 范围、输入、图片、历史、复制、滚动跟随 | 范围、输入、历史已接；图片入口禁用 | 同功能做同状态对照；图片、工作目录等未开放项必须明确说明，不能做可点击空入口 |

上表所说“需对照”是审查任务，不是断言当前必有缺陷。`apps/wps-addin/src/styles.css` 只应放宿主特有规则与必要窄窗修正；能复用的视觉/交互不要再维护第二份。

## 2. 下一批按风险排序

### A. 先固定同状态 UI 基线，不改业务协议

建立可复现的前端 fixture/测试状态，至少成对覆盖：未配置模型、已配置模型、普通问候、澄清待答/已答、查询结果、规划中、待确认写入计划、计划编辑后重预览、写入中、回读通过、回读失败/状态不明、撤销冲突、历史重开。每个状态记录窗格 CSS 宽度、模型/授权状态和消息顺序；分别检查窄窗、加宽窗、长标题、长消息、滚动到中部时新增进度。**不拿不同任务状态的截图作像素差异结论。**

退出条件：差异账逐项标记为“共用且通过 / 视觉或交互待修 / 宿主能力未开放 / 待实机”，并附代码入口及测试证据。没有实机证据只写“自动化通过”。

### B. 优先补计划和结果的信息层级

1. 对照 Excel 计划卡，把 WPS 当前计划的标题、摘要、前提（`assumptions`）、警告、动作类型及目标、影响范围、值/公式样本、编辑后预览修订、取消与确认清楚分层。保留 WPS 已有的前 5 行 × 前 5 列样本，不用完整矩阵撑满对话。完整值仅留在受控计划/编辑器内，不进过程日志或诊断。
2. “可撤销”必须准确写成当前实现的**会话内恢复**；没有可用恢复点或遇到冲突时不能显示“可撤销”。正式耐久恢复未实现前，不得写“关闭后可撤销”。
3. Excel 验证卡有逐项检查与 `verified` / `executed_unverified` / `failed` 区分；WPS 当前只有简化的结果卡。先盘点 `wpsPlanExecutor.ts` 实际回读证据与失败状态，再决定可显示哪些检查。读不到关键属性、部分写入或回读失败时，不能沿用“写入验证通过”。需要新增可见字段时先定义真实来源、持久化边界和旧历史兼容；若改协议，同时同步 `apps/excel-addin/src/contracts.ts` 与 `server/app/models.py`。
4. 过程消息只在同一位置更新同一条时间线。写入、回读、结果卡的顺序固定；不要额外生成第二张空过程卡，也不要把本地确定性执行标成“Agent”。

退出条件：对同一份受支持的 WPS `AnalysisPlan`，预览→编辑→再预览→确认→回读→撤销的每个状态都有相符文案；失败/未知状态不会显示绿色成功、虚假的检查数或跨关闭保证。

### C. 再补共用控件与响应式细节

- 优先抽取两端语义相同且有稳定 props 的展示组件；复制按钮和时间线已抽取，不再复制其 JSX。不要为“共用”把 Office.js/WPS JSAPI 逻辑塞进组件。
- 比较实际 CSS 宽度相同的头部按钮和模型菜单：文字收缩后图标仍可辨、`title`/`aria-label` 存在、菜单不被窗格裁切；不要为了截图强制显示全部标签。
- 复制要覆盖用户与助手消息，复制成功反馈短暂消失；鼠标悬停、键盘聚焦、非悬停设备均有入口，WPS 剪贴板 API 被拒绝时仍可降级。
- 输入区、范围栏、弹层与历史抽屉检查键盘顺序、Esc/关闭、焦点返回、长文本换行；滚动阅读时不能被新进度强制拉到底部。
- 宿主不支持的图片输入、工作目录、复杂动作保持禁用并说明原因；不打开空弹层，不把状态写成“已完成”。

## 3. 不可突破的边界

- 模型做语义判断；工作表读取、计算、执行和回读由本地确定性工具完成。意图阶段不出站原始数据行，用户选择的工作表不可被 UI 自动扩大。
- 写入必须是白名单 `AnalysisPlan`：预览完整影响、用户明确确认、执行前复核授权/工作簿身份/目标变化；编辑后旧确认失效。不得为 UI 对齐放宽 `config/capabilities.json`、WPS 版本闸或动作白名单。
- 用户可见过程仅含受控元数据与真实回读结果；不显示模型隐藏思维、API Key、未经请求的源表原始行或完整写入矩阵。
- 不借 WPS 图片按钮、宏、VBA、外部程序或磁盘文件后备通道伪造 Excel 能力。`PluginStorage` 不提供跨关闭耐久恢复。
- 保护当前脏工作区；先 `git status`，不要覆盖他人的改动，不要把 `.tmp-audit/` 等临时资料加入提交。当前复制控件相关改动尚未提交，先确认并保留。

## 4. 验收与交付证据

自动化至少包含：同一状态下的两端 DOM 断言；计划/结果卡的正反状态断言；复制图标、可访问名称与 WPS 剪贴板降级；窄窗时按钮不溢出；刷新后的进行中任务显示中断而非成功；凭据与原始数据负向断言。静态截图不替代行为断言，测试替身不替代实机。

运行：

```powershell
npm run test:wps
npm run build:wps
npm run test:addin
npm run build:addin
```

若只改 UI 且未改 Python/协议，可说明为何未跑服务端套件；若修改协议、执行或准备提交，按 `AGENTS.md` 与总指南补跑 `python -m pytest server/tests -q`、`python -m compileall -q server/app`。本机 pytest 临时目录需用工作区内隔离 `--basetemp`；若沙箱整目录触发删除守卫，不可谎称“整套通过”，应分别报告整跑与逐文件结果。**不生成安装包。**

实机对照只能使用隔离测试工作簿，不碰用户真实工作簿。分别确认宿主确实是 Excel 与 WPS，记录完整产品版本、加载项版本、窗格 CSS 宽度、模型/授权状态，并对同一场景截脱敏图；写入场景还须记录预览、真实回读、撤销、重开与失败。未做实机时结论保持“待实机验收”。

交付时给出一张差异表：每项的 Excel 证据、WPS 证据、改动文件、测试节点、实机状态和剩余阻塞；不要只报告测试总数或宣称“全面对齐”。

## 5. 批次 B（计划与结果的信息层级）实施账

实施日期 2026-09-29。只动任务窗格 UI 与宿主无关的共享常量；**未改协议**（`AnalysisPlan` 字段与 `server/app/models.py` 均未变），未放宽 `config/capabilities.json`、WPS 版本闸、动作白名单与预览确认。

### 5.1 口径统一的新做法

计划卡/结果卡的**标题与标签**过去是两端各写一遍中文，改一处就会悄悄漂移（例如恢复能力被写成笼统的“可撤销”）。现在抽到宿主无关层：

- `packages/core/src/planInformationHierarchy.ts`：`PLAN_NOTE_SECTIONS`（`assumptions` / `warnings` 的区块标题）、`SESSION_RECOVERY_LABEL`、`SESSION_RECOVERY_NOTICE_PREFIX`、`VERIFICATION_PASSED_LABEL`。
- Excel `App.tsx` 与 WPS `WpsAgentPanel.tsx` 同时改为引用这些常量。

**只共享口径，不共享证据句子。** Excel 的验证证据是逐项 `checks`（`N/N 项符合预期`），WPS 是整块回读比对（`N 个动作 · M 个单元格已回读核对`）—— 来源不同，强行共用就是假装对齐。

### 5.2 逐项差异表

状态含义：**已实现且自动化通过**＝本轮有对应 DOM/契约测试断言；**待实机验收**＝代码与自动化都成立，但没有真机ProductVersion/窗格宽度截图；**能力阻塞**＝受宿主能力或测试成本限制，本轮不做或做不到。

| # | 项 | Excel 证据 | WPS 证据 | 改动文件 | 测试节点 | 实机 | 阻塞 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | 计划卡“采用的信息” | `App.tsx` 的 `inline-notes` 区块 | `wps-plan-notes` 区块，来源同为 `AnalysisPlan.assumptions` | `WpsAgentPanel.tsx`、`packages/core/src/planInformationHierarchy.ts`、两边 `App.tsx` | WPS：`mirrors the Excel plan card hierarchy…`；跨端：`planInformationHierarchy.test.ts` | 待实机 | — |
| 2 | 计划卡“补充说明” | `inline-notes warning` | 原为无标题 `ul`，现加同级标题 | 同上 + `styles.css` | 同 #1 | 待实机 | — |
| 3 | 无内容时不渲染空区块 | `assumptions.length > 0` / `warnings.length > 0` | 同样条件渲染 | 同上 | WPS：`hides both note blocks when the plan states neither` | 待实机 | — |
| 4 | 编辑后重预览的修订标识 | Excel 编辑走独立计划编辑器，字段外化 | `planPreview.revised` 跟着消息走，渲染“已按你的修改更新；写入前请重新核对” | `WpsAgentPanel.tsx`（类型 + `applyEditedPlan` + `planCard`） | WPS：`marks the preview as edited only after the user rewrote the plan`（含持久化后仍为 `revised: true`） | 待实机 | — |
| 5 | 恢复能力写成会话内 | Excel 侧无此标签 | meta / 过程卡 / 活动时间线 facts 统一为 `会话内可撤销`；结果卡保留“恢复点仅在当前加载项会话内可用；关闭后不能从此卡片撤销。” | `WpsAgentPanel.tsx`、`packages/core` | WPS：`mirrors…` 断言 meta 含该标签；跨端常量断言含“会话内”且源文件不含裸 `可撤销` | 待实机 | — |
| 6 | 失败时不谎报验证通过 | 三态：`verified` / `executed_unverified` / `failed` | 两态：`verified` / `undone`；**写失败不生成结果卡**，走 system 消息 | 未改判定逻辑，只统一 `VERIFICATION_PASSED_LABEL` 常量来源 | WPS：`never reports verification passed when the read-back disagrees`（替身写入不落盘 → 真实回读不一致） | 待实机 | — |
| 7 | 撤销后结果卡改态 | Excel 用 `executedPlanId` 与撤销栈 | `verification.status` 变 `undone`，标题“已撤销并核对恢复” | `WpsAgentPanel.tsx`（本轮仅常量化；逻辑原有） | WPS：`switches the verified card to restored and keeps recovery session-scoped` | 待实机 | — |
| 8 | Excel 验证三态的自动化覆盖 | `App.tsx` 内联三元 | — | `apps/excel-addin/src/planInformationHierarchy.test.ts`（契约级） | 同跨端测试 | 待实机 | **能力/成本阻塞**：`App.tsx` 无廉价整体渲染测试，只能做源码与样式的契约断言，不能声称 DOM 级通过 |
| 9 | WPS 的 partial / failed 结果形态 | Excel 有 | WPS 无（刻意保留：无 evidence 就不画卡） | 不改 | 由 #6 反向断言钉住 | — | **能力阻塞**：等待端回读粒度变化后再评估，本轮不虚构 |
| 10 | “保存为工具 / 验证后保存” | Excel 有 `verifiedPlanIds` 入口 | WPS 无 | 不改 | — | — | **能力阻塞**：属产品能力差异，不是信息层级缺陷；不在本批次冒进开放 |
| 11 | 同状态成对截图基线（批次 A） | — | — | 未做 | — | 未做 | **未完成**：需要真实两端宿主与相同窗格宽度，自动化无法替代 |

### 5.3 本轮自动化的确切结果

| 命令 | 结果 |
| --- | --- |
| `npm run test:wps` | 14 文件 / 149 测试通过（本轮新增 5 条） |
| `npm run build:wps` | `tsc --noEmit` + `vite build` 通过 |
| `npm run test:addin` | 51 文件 / 534 测试通过（本轮新增 `planInformationHierarchy.test.ts`，4 条） |
| `npm run build:addin` | `tsc --noEmit` + 两次 `vite build` 通过 |

未跑服务端套件：本轮没有改动 `server/`、`contracts.ts` 或任何 Python，理由同 §4。

### 5.4 工作区状态

未提交；保留前几轮的既有改动。本轮改动：`apps/excel-addin/src/App.tsx`、`apps/wps-addin/src/WpsAgentPanel.tsx`、`apps/wps-addin/src/styles.css`、`apps/wps-addin/src/WpsAgentPanel.test.tsx`、`packages/core/src/index.ts`；新增 `packages/core/src/planInformationHierarchy.ts`、`apps/excel-addin/src/planInformationHierarchy.test.ts`、`docs/WPS_EXCEL_UI_PARITY_GUIDE_2026-09-29.md`。`.tmp-audit/` 等临时资料不进提交。
