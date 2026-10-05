# 修改前｜修改后 双栏变更对照

> 2026-10-03 退役说明：网页双栏／双网格渲染已按用户要求删除。本文为历史记录，不代表当前实现或待实机验收项目。摘要、回读证据与撤销仍保留；WPS 原生独立副本并排审阅继续使用，Excel 同类原生能力尚待实现。

> 状态：**批次 A / B / C 已落地**（core 共用层 + Excel 侧 + WPS 侧的实际变更）——
> 更准确的说法是：**两端双栏口径已接线并通过自动化，待实机验收**。本文档没有做过真机/真工作簿验证，
> 证据只有单测与端到端 DOM 用例，"完全解决""两端口径已对齐"这类结论目前不成立，下文一律按此口径写。
> 预计变更（C5/C2）与文件夹模式（批次 D）尚未开始。本文档已按代码事实同步到批次 C 之后的行号。
> 起草方式：先读 `AGENTS.md`、`docs/ARCHITECTURE.md`、`docs/DEVELOPMENT.md`，再逐个打开三种模式的
> 预览 / 写前快照 / 写后回读 / 撤销实现，按其**实际能力**写本文档。凡文档承诺的能力，下面都给了源码行号；
> 凡与现有执行器冲突、或当前压根不存在的能力，一律进 §9「冲突清单」，不靠文案伪装成已支持。

---

## 1. 两条数据流：预计变更 与 实际变更

二者**不允许**共用同一个数据来源，也不允许互相顶替。

| | 预计变更 Expected | 实际变更 Actual |
|---|---|---|
| 产生时机 | 预览生成时（用户确认之前） | 本次执行写入**之后**，撤销校验窗口之内 |
| 数据来源 | 用户确认那一刻，对**目标区域**的一次本地读取 + 计划里要写的值 | **同一次执行**的写前快照 + 写入后回读，逐格比对 |
| 表达语义 | 「如果现在执行，会变成这样」 | 「刚刚确实变成了这样」 |
| 失败时的说法 | 没有"成功/失败"，只有"目标区域当前不满足预期，请重新预览" | 回读失败 / 不一致 → 必须显示**失败或未核对**，绝不显示成功 |
| 可否跨次累计 | 否，只对当前计划 | 否，只绑本次 `executionId` |

**硬规则**：实际变更只能来自 §4 列出的回读点。禁止用计划值、用上一次的结果、或用"写了没报错"来充当实际变更。

---

## 2. 共用数据模型（`packages/core/src/changeReview.ts`，新增）

纯数据 + 纯函数，不含任何宿主 API，Excel / WPS / 文件夹三端共用。

```ts
/** 单元格显示给用户的值：null = 空单元格，"" = 空字符串。 */
export type ChangeValue = string | number | boolean | null;

/** 一个单元格在某时刻的样子。value 与 formula 同时容纳，不允许二选一。 */
export type ChangeCell = {
  /** 显示给用户的值：空单元格=null，空字符串="" */
  value: ChangeValue;
  /** 原公式；非公式单元格为 null */
  formula: string | null;
  /** false = 该宿主这一次没采集到计算值，只能展示公式（不得拿公式冒充值） */
  valueCaptured: boolean;
};

export type ChangeKind =
  | "value"        // 普通值变化
  | "formula"      // 公式本身变化（含 公式↔值 互转）
  | "formulaOnly"  // 公式没变，但覆盖行为本身需要警示（见 §3.3）
  | "empty"        // 清空：原非空 → 空
  | "same"         // 无变化（仅在"只看改动"关闭时出现）
  | "not_executed";// 部分写入时，这一行压根没写到

export type ChangeRow = {
  key: string;                 // 稳定行标识，用于虚拟化 key 与滚动同步
  address: string;             // 形如 "C5"
  before: ChangeCell;
  after: ChangeCell;
  kind: ChangeKind;
  changed: boolean;
  /** 修改前是公式，且本次把它覆盖掉了 → 双栏必须突出标记「将覆盖公式 / 已覆盖公式」 */
  overwritesFormula: boolean;
};

export type ChangePhase = "expected" | "actual";

export type RangeChangeReview = {
  sheet: string;
  range: string;
  phase: ChangePhase;
  /**
   * ready=还没写；applied=写后回读与计划一致（唯一允许宣称"写入成功"的状态）；
   * executed_unverified=写确实发生了、前后矩阵也拿到了，但独立验收没全过；
   * readback_failed / readback_mismatch=回读失败或不符；not_executed=这一步压根没写。
   */
  status: "ready" | "applied" | "executed_unverified" | "readback_failed" | "readback_mismatch" | "not_executed";
  rows: ChangeRow[];
  /** 限额内的统计，与 rows 长度可能不等（见 §6.4） */
  changedCount: number;
  totalCount: number;
  truncated: boolean;
};

/** 一份快照：值与公式两个矩阵都带，外加"这次有没有采到计算值"。 */
export type ChangeSnapshot = {
  values: ChangeValue[][];
  formulas: (string | null)[][];
  valuesCaptured: boolean;
};
```

**入参与纯函数**（已落地，见 `packages/core/src/changeReview.ts`）：

- `buildRangeChangeReview({ sheet, range, phase, status, before, after, appliedRows?, limit? })`
  —— `appliedRows` 表示真正写入的行数，超出的行一律判 `not_executed`（既不算改动，也不许显示为已写入）。
- `cellDisplayText(cell)` —— 值未采集时返回 `"显示值未采集"` 且 `unresolved = true`，**不会**把公式原文塞进值栏。
- `canClaimWriteSuccess(status)` —— 只有 `applied` 返回 true；`readback_failed` / `readback_mismatch` / `not_executed` / `ready` 一律 false。
- `changeReviewSummary(review)` —— 摘掉全部单元格内容，只留定位与计数，**是唯一允许进活动日志/历史的形态**。
- `onlyChangedRows(review)` —— 「只看改动」开关；未执行的行保留（它们属于异常提示，不该被过滤掉）。```

**为什么 `ChangeCell` 必须同时装 value 和 formula**：用户在对照里要看的是"值"，但判定写入是否成功时，
公式变更必须核对**公式本身**（`formulas`），不能只看计算值 —— 计算值会因为重算时机而漂移。二者缺一不可。

---

## 3. 纯比较逻辑（同文件，`packages/core`）

### 3.1 空单元格 ≠ 空字符串

- `null` = 空单元格；`""` = 空字符串，**两者必须能互相转成 kind 变化并显示为一次改动**。
- 比较函数**不得**先把 `""` 折叠成 `null`。当前 WPS 侧已有两种 normalization，语义不同：
  - `wpsWorkbookReader.ts:89-95 normalizeCell` —— **保留** `""` 与 `null` 的区别（`return value === "" ? "" : null`）；
  - `wpsPlanExecutor.ts:145-150 canonicalCell` —— **折叠** `"" → null`（仅供指纹计算）。
  - ⇒ 双栏走 **前者**；指纹仍走后者，两条路不许混。

### 3.2 改动的判定

```
changed = before.value !== after.value || before.formula !== after.formula
```
且 `before.valueCaptured === false` 或 `after.valueCaptured === false` 时，
**只能比 `formula`**，并在该单元格标注「显示值未采集」。

### 3.3 公式的处理

- 修改前是公式 → `overwritesFormula = true`，双栏在该行显示「已覆盖公式 / 将覆盖公式」标记。
- 公式单元格的值栏：显示**计算值**；公式原文另起一行小字展示，**不许把公式原文塞进值栏顶替**。
- 值未采集时：值栏显示「显示值未采集」，公式栏照常显示 —— 与该单元格是不是公式无关。

### 3.4 清空（clearRange）

`after` 为「空」，也要渲染成一行 `before →（空）`，`kind: "empty"`。
**不允许**因为 after 是空就判定成"未改动"而不展示。

---

## 4. 三端数据来源（现状接线点）

### 4.1 Excel 加载项（Office.js 当前工作簿模式）

| 数据 | 现状位置 | 是否够用 |
|---|---|---|
| 写前快照 | `excel.ts:3680-3705 captureUndoRange`；load 列表在 `3688`，目前是 `formulas,numberFormat,format/fill/color,format/font/bold,format/font/color`，`3698` 把 `range.formulas` 经 `normalizeValue` 存入。**只存了 formulas，没有 values** | ❌ 需要补 values（见 §9 C1） |
| 写入后回读 | `excel.ts:4049-4092 captureUndoPostState`：写了 `range.load("formulas,numberFormat,…")`（`4069-4071`）后**只留一个 `postFingerprint`**（`4073`），矩阵丢弃 | ❌ 需要在事件点另读（见 §9 C3） |
| 独立验收 | `excel.ts:1830-2267 verifyPlan`：按 `inferredCriteria` **抽样**核对，产出 `VerificationReport{status,passed,checks,unverifiedActions}`；状态判定 `2255-2261`（任一不过=`failed`；零 check 或有 unverified=`executed_unverified`） | 抽样，不等价于全矩阵，二者并存、互不替代 |
| 预览卡片 | `App.tsx:8569-8668 .inline-plan`：标题 + N 步 `actionLabel` 列表 + assumptions/warnings，**一个单元格值都没有** | ❌ 预计变更需新引进 |
| 撤销 | `undoExecution` `excel.ts:4250-…` → 先 `assertUndoPostState(4159-4165)`（要求 `postStateCoverage==="complete"` + workbookIdentity 一致） | ✅ 语义保持不变 |

### 4.2 WPS 加载项

> 状态：**批次 C 已接线、自动化已过、待实机验收**。下表"批次 C 后"一列是改完的行号。

| 数据 | 原始位置 | 批次 C 后 | 是否够用 |
|---|---|---|---|
| 写前快照 | `wpsPlanExecutor.ts:80-86 StoredAction`：`beforeValues` + `beforeFormulas` + `beforeFingerprint`，`540-546` 填充，**完整矩阵常驻**（PluginStorage 会话内） | 不变。`wpsActionChangeReview`（`259`）直接用 `beforeValues` + `toChangeFormulas(beforeFormulas)`（`247`） | ✅ |
| 写入后回读 | `622-627`：`readMatrices` 已读整块 after，`623` 只取 `afterFingerprint`，再 `matricesEqual` 判定；**矩阵随即丢弃** | 同一次读取不再丢：分块读 `713` 与终读 `745` 各留一份到 `afterByIndex`，比对结论记进 `matchByIndex`（`752`）。**没有新增任何宿主调用** | ✅ |
| 交付出口 | 无 | `executeWpsPlan` 第 4 参 `onChangeReview`（`597`）；成功 `760`、失败 `790` 各回调一次 | ✅ 失败也交付（§6.5） |
| 面板接线 | 无 | `WpsAgentPanel.tsx:46` 引入**同一个** `ChangeReviewTable`；`1013` 存进 `changeReviewsRef`（`372`，会话内内存）；成功挂 `verification.executionId`（`1952`）、失败挂 `changeReviewExecutionId`（`1957`） | ✅ |
| 预览卡片 | `WpsAgentPanel.tsx:1355-1397`：**已有待写值矩阵** `action.preview`，前 5 行 5 列截断（`1393`），已区分 `null→"空单元格"` / `""→"空字符串"`（`1388`） | 未动 | ⚠️ 只有"写入后"，缺"修改前"那一栏（预计变更，同 C5） |
| 撤销 | `undoWpsExecution` `664-728`；`680-698` 前置检查，`716-722` 恢复后回到 `beforeFingerprint` 校验 | 语义不变；撤销**成功**时 `1134` 把该 executionId 的明细从内存释放（见下方"撤销与明细生命周期"） | ✅ 语义保持不变 |
| 活动日志 | `wpsActivityLog.ts`，`WpsAgentPanel.tsx:1925` 渲染 | 组件仍从 `excel-addin/src/components/TaskProgressTimeline` 引入（`43`） | ✅ 沿用 |

**撤销与明细生命周期**（批次 C 决定，写在这里是因为它偏离了验收用例 #6 的字面要求）：

- 撤销**成功** → 区域已回到执行前，这份对照描述的现场不再存在，**完整矩阵**从 `changeReviewsRef` 删除
  （`1134`）。**不做**"把明细改标成已还原"：`RangeChangeReview.status` 里没有这个状态，
  硬造一个等于让 UI 显示一份前后完全相同的"变更"。
  - 但"本次曾修改、现已撤销"是**操作记录**，不是单元格状态，不该跟着一起抹掉。做法是先取
    `changeReviewSummary(review)`（只有定位与计数，**没有任何单元格内容**）存进
    `changeReviewSummariesRef`，再删完整矩阵；界面由共用组件 `ChangeReviewUndoNote`
    渲染成一句「本次曾在 Sheet!Range 改动 X / Y 个单元格；现已撤销，逐格明细已按要求释放」。
  - 摘要仍是会话内内存，不进 `ChatMessage`、不进 `WpsUndoRecord`；按 §7 它是唯一允许进日志/历史的形态，
    当前只停在界面上。
- 撤销**失败**（`undo_conflict`）→ 明细**保留**，用户能看到写了什么、为什么不能自动覆盖，符合 #6 后半句。
- 两条路径都满足 §7 第 4 条：**完整矩阵**的生命周期不长于撤销入口。

### 4.3 文件夹模式（openpyxl，服务端）

| 数据 | 现状位置 | 是否够用 |
|---|---|---|
| 预览 | `main.py:2182 folder_execute` / `2192 stream` / `2383 status` / `2478 resume`；全仓 grep `folder_preview\|dry_run` **零命中** | ❌ 无预览端点（见 §9 C2） |
| 写前比较 | `folder_workbooks.py:2729-2744 _matrix_would_change` + 各分支 `action_changed`（`5000-5280`） | ⚠️ 逐格比过，但**只产布尔** |
| 写入后回读 | **无**。`ActionExecutionResult{index,type,sheet,status,changed}`（`models.py:1386-1392`，构造于 `folder_workbooks.py:5399-5408`） | ❌（`changed` 还是写入**前**算的预测，见 §9 C4） |
| 撤销 | 独立 undo_stack + 幂等日志（`4168-4177`），`1907-1914` 有 `len(undo_stack)` 一致性检查 | ✅ 语义保持不变 |

---

## 5. WPS 回读证据边界（必须原样告知用户，不许拔高）

1. **动作白名单**：`config/capabilities.json` 的 `wpsWrite.allowedActions = ["writeValues","writeFormulas","clearRange:contents"]`。
   其余 `clearRange` 变体（`all`/`formats`/`hyperlinks`）在 `wps-et` 里是 `unverified`，**不进计划**，因此也不会有回读证据。
2. **R1C1 铺满公式**：`wpsPlanExecutor.ts:296` 直接抛错拒绝执行 ⇒ 该动作**不会有实际变更**，双栏要显示"未执行"而不是空表。
3. **规模上限**：`wpsWrite.maxCells = 1000`、`chunkRows = 100`；分块写入时 `595-596` 每块都重读一次全区域。
4. **版本闸**：`verifiedVersionPrefixes: []`、`developmentVersionPrefixes: ["12.0"]` ⇒ 目前只在开发调试版本开放
   （`WpsAgentPanel.tsx:1362` 已提示），属"未实机验收"的能力，双栏不得把回读结果表述成"已验证"。
5. **v1 恢复点的空值歧义**：`hasAmbiguousLogicalCell`（`212-217`）会让 v1 记录**拒绝自动覆盖**
   （`683-691`，抛「无法区分空字符串与空单元格」）。
   **这不波及本次执行的"修改前"**：新记录恒定 `version: 2`（`554`），写前快照走 `readMatrices` 的默认 v2
   归一（`540`），`normalizeCell` 保留 `null` / `""` 区别。v1 只影响**旧记录**的撤销，
   那时根本没有本次执行的双栏可展示 —— 不要拿这条去给"修改前可能不准"加免责。
6. **分块写入的"修改后"取自最后一次真实回读**：`chunkRows = 100` 分块时，每块结束都重读整块区域
   （`709`），最后还有一次终读（`736`）。中途抛错时，`afterByIndex` 里留的是**最后一次成功的回读**，
   不是"计划值"。它对应的状态是 `readback_failed`（没走到比对），不是 `applied`。

---

## 6. 双栏交互

### 6.1 布局与同步

- 单张表，结构：`行地址 | 修改前 | 修改后 | 状态`。**左右两栏之间永不换行**，换的是行。
- 「修改前」与「修改后」必须**同一滚动容器内的左右两列**，不允许两个独立滚动区域（会导致错位对照失效）。
- 行 hover / 选中 → 两栏同行高亮联动。
- 键盘：上下键换行，`Tab` 进表格后可逐行聚焦；每行提供「复制此行前后值」按钮（复用现有 `MessageCopyButton`）。

### 6.2 必带控件

| 控件 | 行为 |
|---|---|
| 改动数量 | 「本次共改动 N / M 格」，N 来自 `RangeChangeReview.rows.filter(changed).length` |
| 只看改动 | 开关；关闭时也要能看到"未改动"行，但默认**打开** |
| 同步定位 | 点击某行 → 让宿主选中对应单元格。**注**：现有源码里没有现成的"按地址选中"调用（`excel.ts` 只有 `getSelectedRange()` 读取方向，如 `595/653/674/806`），这是**待实机验证的新增宿主调用**；未通过实机验收前不允许承诺，失败时静默降级（不报错也不假成功） |
| 覆盖已有值 | 任何 `before` 非空的改动行，默认带上「覆盖」标记，不给奖励换入口确认 |
| 将覆盖公式 | `overwritesFormula === true` 的行，显示**独立警示**（不同于普通覆盖） |
| 失败状态 | `readback_failed` / `readback_mismatch` / `not_executed` 时，整张表顶部显示红色/警示条，且不显示任何"成功"措辞 |

### 6.3 窄任务窗（≤430px）

- 不转卡片堆叠，并保持双栏：`行地址` 收窄为行内小字标题，两栏继续左右并排。
- 值文本**不截断到看不出差异**：超长走换行 + 最大高度滚动，禁止 `text-overflow: ellipsis` 吞掉差异。
- 参考现有先例：`apps/wps-addin/src/styles.css:965` 的 `@media (max-width: 430px)` 变体，沿用同一断点值。

### 6.4 限额与虚拟化

- 预览（预计变更）：沿用 WPS 现状 **前 5 行 × 前 5 列**（`WpsAgentPanel.tsx:1393`），三端统一；超出显示提示而非静默截断。
- 明细上限：已加进 `config/capabilities.json` 的 `changeReview` 段
  （`maxRows = 200`、`maxTotalCells = 20000`、`captureMaxCells = 20000`、
  `previewMaxRows = 5`、`previewMaxColumns = 5`），**不放前端硬编码**；
  有测试盯住这些值存在且为正数，防止以后被写死回代码里。
- 超限时：`truncated = true`，顶部一句「已展示 X / Y，其余 Z 项省略」。
- 大范围走虚拟滚动：只渲染可视行，`rows` 受 `maxRows` 限制后仍按窗口切片渲染。

### 6.5 失败后的重试

- **禁止自动重试**。现状已经如此：Excel 续跑只能由用户点按钮触发（`App.tsx:7767`）且必须带
  `recovery:{startActionIndex, undoSnapshot}`（`2006-2007`）；WPS `partial_write_recoverable` 只解锁撤销（`1039`），不自动续跑。
- 失败后必须先展示**实际变更对照**（哪些步写入了、哪些没写），再让用户决定「撤销」还是「重新生成计划」。
- 重试入口不得直接复用上一次的 after 矩阵当作新计划的 before —— 必须重新读取。

**"原计划不能再直接执行一遍"的证据（用户提问点，这里给实证而不是承诺）**

拦住第二次执行的是 `executeWpsPlan` 的**预览指纹复核**，不是"计划已执行"这种显式标记：

| 守卫 | 位置 | 拦什么 |
|---|---|---|
| 计划对象未被改动 | `wpsPlanExecutor.ts:604` `confirmationSnapshot(prepared) !== prepared.confirmationSnapshot` | 计划内容被改过 |
| 源区域仍是预览时的样子 | `619-622` `current.fingerprint !== source.previewFingerprint` | 计算依据变了 |
| **目标区域仍是预览时的样子** | `639-641` `before.fingerprint !== action.previewFingerprint`（保存恢复点之前） | 用户改过 / **上一次已经写过** |
| 执行前再核一次 | `687-689` `current.fingerprint !== action.beforeFingerprint` | 确认到执行之间又变了 |

第一次执行只要真的写了东西，目标区域指纹就不再是预览那一刻的值 ⇒ 第二次点「确认写入」会在
`639-641` 抛「…在预览后发生变化，请重新预览并确认」，**写不进去**，也**不会留下新的恢复点**。
已锁进 `wpsPlanExecutor.test.ts` 两条用例（成功后再跑、回读不一致失败后再跑），并做过变异验证：
把 `639` 的判断改成 `false &&` 后，这两条连同既有的「rejects a changed target after preview」一起红。

⚠️ **仍是缺口（下一批）**：这只是"执行器层面拦得住"，**UI 层面**失败后 `pendingPlan` 没有清空
（`WpsAgentPanel.tsx:1047-1070` 的 catch 只清 `writeProgress`），计划卡片上的「确认写入」按钮
（`1452`）依然可点 —— 点了会被上面那条守卫拒绝并提示重新预览，不会静默重复写入，但体验上是
"让你点、再告诉你不行"，不是"直接不让你点"。补一个显式闸门（失败后置灰 + 引导撤销/重新生成计划）
属下一批。

---

## 7. 隐私边界

1. **不进意图模型**：双栏所需的读取全部是本地确定性读取，不参与 LLM 请求构造；沿用既有"意图阶段不发原始数据行"。
2. **不进聊天历史**：完整前后矩阵存在会话内内存 / 独立 storage key，**不得**挂到 `ChatMessage` 上被持久化。
   参照现有 `conversation.persistedMessagesPerConversation = 60` 的持久化面，双栏明细不在其中。
3. **不出站**：文件夹模式的明细在服务端内存中产生后，只随本次执行结果返回，不写日志文件、不落日志文件外的持久存储。
4. **生命周期**：与实际变更绑定的明细，生命周期**不长于**对应的撤销快照（Excel `undo.maxEntries = 5`；
   WPS PluginStorage 不跨关闭）；撤销入口消失时，明细一并释放，不承诺跨关闭可查。
5. **只读授权区域**：`. RangeChangeReview.range` 始终 ⊆ 本次已写入的目标区域 ⊆ 用户勾选的工作表。补读不扩授权。

---

## 8. 验收用例（先红灯，后修复）

文件格式：`packages/core/src/changeReview.test.ts`（纯逻辑，已落地，15 条，**先红后绿**）+ 两端 DOM/交互测试（批次 B/C 补）。

**红灯记录**（批次 A 实施前）：`Cannot find module './changeReview'`，1 个测试文件 FAIL / 7 个文件 101 测试通过
—— 红灯只来自新增缺口，未连带破坏既有用例。补实现后：**8 文件 / 116 测试全通过**。

**批次 C 的红灯**：新增用例**首跑即绿**（实现与用例同一批落地，没有可记录的红灯）。为排除"必然成立的
空过断言"，补做了一次**变异验证** —— 故意改坏两处生产代码后重跑，恰好炸掉对应的两条：

| 变异 | 炸掉的用例 | 说明 |
|---|---|---|
| 删掉 catch 里的 `emitReviews()`（`790`） | 「回读与计划不一致时照样交付对照…」`expected null not to be null` | 证明"失败也要交付"这条断言真有牙齿 |
| 缺 after 时把 `valuesCaptured` 从 `false` 改成 `true` | 「一个动作都没写时整块标未执行…」`expected false to be true` | 证明"没回读过不许冒充空"不是空过 |
| 把 `639` 的目标区域指纹复核改成 `false &&` | 「执行过一次后再点确认也不会二次写入…」「回读不一致失败后不允许直接重跑」（+ 既有「rejects a changed target after preview」） | 证明"防重复执行"这条不是靠空过 |
| 把撤销成功时"先留摘要"的分支改成 `if (false && …)` | 「…releases them on undo」「keeps the 'changed then undone' note free of cell values…」`expected null not to be null` | 证明撤销后那条摘要真的来自被释放的明细 |

变异已还原。另有一条**误报**值得记：隐私用例最初断言持久化 JSON 不含 `"90"`，
结果被时间戳 `1790658903460` 命中 —— 断言本身太糙，已改成只存在于单元格里的哨兵值
（`单元格私密值-K9Z3`，见 `WpsAgentPanel.test.tsx`「keeps cell values out of the persisted
conversation…」）。这条不是代码问题，记下来防止下次重犯。

下表 #1-#5、#10、以及"显示值未采集"已有纯逻辑覆盖；**#6 撤销、#7 越权拒绝、#9 窄视图、#8 出站隐私的 UI 侧**属宿主层，批次 B（Excel）/ C（WPS）补。

| # | 场景 | 断言要点 |
|---|---|---|
| 1 | 空单元格 vs 空字符串 | `"" → null` 与 `null → ""` 各自判为改动；`before/after` 都是 `null` 判为未改动；不得被 collapse 成同一状态 |
| 2 | 公式 | 公式单元格参与 kind=`formula`；**只有**计算值变而公式没变时不得误判为公式改动；写入成功判定走 `formulas` 比对 |
| 3 | 清空 | `非空 → 空` 必须渲染成一行 `before →（空）`，不许被当成"无变化"过滤掉 |
| 4 | 部分写入 | 分块中断时，已写行有 after、未写行标记 `not_executed`；`changedCount` 只数已写且实际变了的 |
| 5 | 回读不符 | after 与计划不一致 → `status = "readback_mismatch"`，界面不出现任何"成功/已验证"措辞 |
| 6 | 撤销 | 撤销成功后**完整矩阵**释放、只留"曾修改、现已撤销"的摘要（无单元格内容）；撤销失败（`undo_conflict`）时明细保留并给阻断，不得继续说可撤销 |
| 7 | 越权拒绝 | 目标区域超出勾选工作表 → 不产生任何读取与明细（连 before 都不读） |
| 8 | 出站隐私 | 明细不出现在 ChatMessage 持久化快照里、不出现在任何一种请求体里（构造 payload 后断言不含单元格值） |
| 9 | 窄视图 | 430px 下修改前/修改后仍是同一行左右两栏；长文本换行不被 ellipsis 吞掉 |
| 10 | 限额 | 超 `maxRows` 时 `truncated=true` 且提示里带上真实 X/Y 数字，不静默裁掉 |

---

## 9. 冲突清单（指导要求 ↔ 现有执行器能力，如实报告）

> 以下**没有一条**是通过改文案绕过去的。能补的就写清楚要补什么；补不动的就摆明但书。

| # | 冲突 | 现状证据 | 处理 |
|---|---|---|---|
| C1 | 「修改前显示**值**」，但 Excel 写前快照只有 `formulas` | `excel.ts:3688` load 列表、`3698` 存 formulas | ✅ **批次 B/C 全部解决**：Excel 侧 `captureUndoRange` load 列表加 `values`，返回 `{snapshot, beforeValues}`（值单独返回，不进会被持久化的 `UndoRangeSnapshot`）；WPS 侧本就有 `beforeValues`，批次 C 把两端接进**同一个** `ChangeReviewTable` + 同一套 `buildRangeChangeReview`。**两端口径此刻对齐**：都是"普通格显示值、公式格显示计算值并在下方给原公式、覆盖公式单独标记、没采到值就显示'显示值未采集'" |
| C2 | 文件夹模式没有预览端点 | `main.py` 无 dry_run/preview | 若要做「预计变更」，要么新增 preview 端点（**属协议变更**，须同步 `contracts.ts` 与 `server/app/models.py`），要么文件夹模式只出「实际变更」。**本批建议先只做实际变更** |
| C3 | Excel 没有写入后矩阵 | `captureUndoPostState` 只留 `postFingerprint`（`4073`） | ✅ **批次 B 已接线，自动化已过（待实机验收）**：`collectChangeReviews()` 在写入完成、`verifyPlan` 之前，按 `undoRanges`（本次真正写过的区域）补读一次 values+formulas；回读失败标记 `readback_failed` 且不阻断主流程。限额走 `capabilities.changeReview.captureMaxCells` |
| C4 | 文件夹模式 `changed` 是**写入前**的预测 | `_matrix_would_change(2729-2744)`、`5399-5408` | 写动作**完成后**用已加载的 sheet 再读一次目标区域作为 after；现有 `changed` 布尔继续保留，但不能再充当"实际变更"的证据 |
| C5 | Excel 预览卡片完全没有值 | `App.tsx:8569-8668` | 预计变更要在**生成预览**这一事件点读一次目标区域当前值（只读目标区，本地确定性读取，不违反"不发原始行给模型"） |
| C6 | 规模不可控 | Excel `undo.maxTotalCells = 400000`，WPS `wpsWrite.maxCells = 1000` | 双栏一律限额 + 虚拟化；大区域只给摘要与破坏数，不逐格渲染 |
| C7 | WPS R1C1 铺满公式拒绝执行 | `wpsPlanExecutor.ts:296` | 该动作无实际变更可展示，双栏显示「未执行」及原因，不出空白表 |
| C8 | 跨 workspace 复用 React 组件 | WPS 已能从 excel-addin 引组件（`WpsAgentPanel.tsx:43`），生产可行；但 WPS 测试环境存在 React 双实例风险（已记录在 `.workbuddy/memory/MEMORY.md`） | ✅ **批次 C 已接线，自动化已过（待实机验收）**：共用组件依赖只有 `react` + 共用层，**不引任何第三方库**；样式两端各一份（WPS 侧见 `apps/wps-addin/src/styles.css` 末尾，组件内不写死颜色）。新加的 `wpsChangeReviewTable.test.tsx` 用 **WPS 本工作区**的 `react-dom/client` + `act` 挂载 Excel 那个组件，9 条全过 —— 双实例一旦发生，这里会先红。⚠️ 这只能证明"在 jsdom 里没踩双实例"，不能证明真机 WPS 加载项里没有 |
| C9 | 「实际变更」不能被用来偷换「已写入」的结论 | `verifyPlan` 是独立的抽样验收（`excel.ts:1830-2267`） | 两套证据并存、`VerificationReport` 语义不变；`verification.status !== "verified"` 时双栏不得出现"成功/已验证"措辞 |

---

## 10. 分批实施

- **批次 A**（✅ 已完成）— `packages/core/src/changeReview.ts`（模型 + 纯比较）+ `changeReview.test.ts` 15 条红→绿；已挂到 `packages/core/src/index.ts`。不碰三端。
- **批次 B**（✅ 已完成，Excel 侧实际变更）— `captureUndoRange` 补 `values`（C1）、
  `collectChangeReviews` 在事件点补读 after（C3）、新增共用组件
  `apps/excel-addin/src/components/ChangeReviewTable.tsx` 并在 Excel 结果卡渲染。
  前后矩阵只存在 `changeReviewsRef`（按 executionId 索引的会话内内存），
  **不挂 ChatMessage**。新增 6 条执行器用例 + 9 条组件用例。
  **仍未做**：C5（Excel 预览卡的"预计变更"）。
- **批次 C**（✅ 已完成，WPS 侧实际变更）— 复用已有 after 读取（分块 `713` / 终读 `745`，**零新增宿主调用**）
  + `StoredAction.beforeValues`，经 `wpsActionChangeReview`（`259`）进**同一个** `ChangeReviewTable`。
  回读矩阵只活在 `executeWpsPlan` 闭包与 `changeReviewsRef`（`WpsAgentPanel.tsx:372`）里，
  **不进 `WpsUndoRecord`**（那个落 PluginStorage）、**不进 ChatMessage**（那个落 localStorage）。
  撤销成功释放完整矩阵（`1134`）、保留"曾修改、现已撤销"摘要，撤销失败保留明细。
  新增 10 条执行器用例 + 9 条组件用例 + 3 条面板用例，并补做变异验证（见 §8）。
- **批次 C 追加项**（本轮补做，同属"待实机验收"）：
  - **防重复执行的证据** —— 新增 2 条执行器用例把「同一份计划执行过一次 / 回读不一致失败后
    都不能直接再跑」钉死，并写清到底是哪道守卫在拦（见 §6.5）。**UI 层的显式闸门仍是缺口**。
  - **撤销后保留摘要** —— 新增共用组件 `ChangeReviewUndoNote`（`ChangeReviewTable.tsx`），
    WPS 侧接线；新增 1 条面板用例 + 扩展既有撤销用例。Excel 侧尚未接这个组件。
- **批次 D** — 文件夹模式：写后回读补 after（C4）；预计变更按 C2 另行议定。
- **批次 E** — 窄窗/虚拟化/限额（写入 `config/capabilities.json`）、`docs/` 与 `README.md` 同步、三端回归。

每批跑：`npm run test:addin`、`npm run build:addin`、`npm run test:wps`、`npm run build:wps`、
`python -m pytest server/tests -q`、`python -m compileall -q server/app`。
批次 D 若改协议，另加 serv端 pytest 与两侧协议一致性检查。

---

## 11. 本批明确**不做**

- 不改 `AnalysisPlan` 白名单、不动授权边界、不放开任何 `unverified` 动作。
- 不新增脚本 / VBA / 宏 / 外部程序执行能力。
- 不承诺跨关闭可撤销。
- 不把完整前后矩阵写进聊天历史或任何出站请求。
- 不生成安装包、不操作真实工作簿、不提交 Git。
