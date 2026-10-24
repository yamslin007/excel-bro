# 宿主动作能力差距表（Excel / WPS）

日期：2026-10-14（本轮**按实际代码/配置/证据核查后修订**）。
依据：`config/capabilities.json`（`hostCapabilities` + `wpsWrite.allowedActions`）、
`apps/wps-addin/src/wpsPlanExecutor.ts`（执行器实现 + 属性级守卫 + 白名单闸）、
`apps/wps-addin/src/wpsFormatCapability.ts`（已验证属性子集）、
`apps/wps-addin/src/wpsWriteCapabilities.ts`（预检）、
`docs/WPS_FORMAT_PROBE_EVIDENCE_2026-10-13.md`、`docs/WPS_TEXT_WRITE_FIX_2026-10-12.md`、
`docs/EXCEL_FORMAT_COMPAT_2026-10-13.md`、`docs/WPS_W3_A_VALIDATION_2026-09-23.md`。

只记录**动作**层面的能力差距，不记录 UI / 展示差异（那些见
`docs/DUAL_HOST_PLAN_UI_PARITY_HANDOFF_2026-10-14.md`）。

---

## 2026-10-18 当前状态（覆盖下方旧统计）

按 config/capabilities.json 的动作类型去重计数，共 37 类；本轮起点是 6 类有执行器、31 类没有实现，不是“仅待验收”。此前“38 类 / 剩余 28 类”的统计有误，现已更正。目前增加 `setFill`、`setNumberFormat`、`writeTable` 和 `copyRange` 子集，现又接入 `computeValues`，再加入边框子集，现再接入 `sortRange`，合计 **13 类入口有实现，24 类仍未实现**；这不是完成百分比，也不表示 13 类的全部属性均支持。`clearRange` 仍只支持 contents，copyRange 与格式属性限制见下表。

按含变体的配置键计数是 40 项：13 项已支持、27 项未实现；后者是 24 类完整动作加 clearRange 的 all/formats/hyperlinks 三个变体。不能将动作类型数与变体数混用。

新增能力按同一授权、执行一次、回读、结果审阅与恢复路径运行。`writeTable` 是写入矩阵和表头样式，**不是 createTable 表格对象**。WPS 将它展开为 writeValues/setFont/setFill；复制只支持 values/link、转置、跳过空白，展开为值/公式动作。来源完整性与最新值在本机复核，来源值不进入恢复描述。复合操作保留原始计划供保存工具复用，不能把动态复制固化为当次数据。服务端与客户端共同拒绝 all/formulas/formats 复制。条件复制已支持：字段唯一解析，保留表头，链接保持实际来源行号，转置与跳过空白可组合；缺失/重复表头不猜测。条件清空仅限 contents，保留表头与未命中行。

配置 `wpsWrite.actionExpansion` 区分入口动作与实际恢复动作；`maxExpandedActions=500` 不超出恢复协议容量。展开后每个动作有自己的写前/写后回执，配置不得直接扩大恢复动作白名单。

新增自动化：`wpsExtendedFormats.test.ts`、`wpsWriteTable.test.ts`、`wpsCopyRange.test.ts`；服务端 `test_wps_execution_api.py`、`test_planning_capability_narrowing.py`。反向验证：`scripts/check_wps_capability_guards.mjs`。真实 Excel/WPS 的本批新能力尚未实机验收。

### 2026-10-18 继续实施：边框

- 新增 `borderEdges.ts` 和 WPS 边框适配：外框与内部线条分别展开，以同一条物理线的归属检查防止相邻格组合撤销歧义。沿用恢复记录限额，超限整体拒绝。
- 本地快照保留线型、粗细、解析颜色、自动色/主题色模式及色调；服务器仅接收四种 border 属性名和指纹，不接收本地快照。缺关键属性时不写入，未把未知主题色拍平为 RGB。
- 两端审阅规则避开业务边框（含相邻单元格共用边界）：改边框仅叠底色，同时与填充修改重叠时不叠加，直接显示实际结果。用户原有条件格式仍保留。
- 不新增排序、去重、删除表或结构恢复。openpyxl/Office.js 既有边框实现保留；这次新增 WPS 适配与两端标记一致性。真实宿主验收仍未执行。
- 官方接口参考：[WPS Borders](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Borders/obj)；接口声明不代替具体安装版本的读写取证。

### 边框批次最终验证

| 检查 | 实跑结果 | 日志 |
|---|---|---|
| core | 597 passed | outputs/wps-borders-core.log |
| Excel | 1318 passed / 3 skipped | outputs/wps-borders-excel.log |
| WPS | 947 passed / 3 skipped | outputs/wps-borders-wps.log |
| 服务端 | 1511 passed / 1 skipped | outputs/wps-borders-server.log |
| 两端 build（含 tsc） | exit 0 | outputs/wps-borders-build-addin.log / wps-borders-build-wps.log |
| compileall | exit 0 | 同次串行服务端检查 |
| 本轮新增 3 项反向验证 | RED → SHA256 还原 → GREEN | outputs/wps-borders-mutations.log |
| 原审阅 R15 / R17 | 预期业务断言 RED，SHA256 还原，最终全量 GREEN | outputs/wps-borders-review-R15.log / R17.log |

原审阅脚本 17 个 SEED 均在位，本轮只重跑受影响的 R15/R17，不冒称原 17 项全部重新实跑。能力守卫脚本由 14 增至 17 项，本轮新增 3 项已实跑，原 14 项保留上一批证据。针对旧断言和宿主替身的修正：边框缺少快照仍拒绝；面板替身补齐独立条件格式 Interior，断言改为检查活规则而非已删除规则。实机未操作，未提交、推送或发布。

### 2026-10-18 继续实施：计算、条件操作与垂直对齐

- WPS `computeValues` 复用 Excel query_table 与值归一。来源与目标必须已授权，先检查全部授权再读来源；值/公式/数字格式/显示文本变化均使准备结果失效。含表头时同 Excel 加粗及浅绿底；保存工具保留动态计算计划。
- 条件复制与条件清空以本机表头定位，未知/歧义字段明确拒绝。复制链接在过滤和转置后保留真实来源行号。前序数据写入可能改变来源时拒绝，不能使用旧矩阵继续。
- 垂直对齐按逐格旧属性采集、执行、回读、冲突检查及恢复；不将自动化证据写入历史实机 VERIFIED 集合。客户端与服务端新增的恢复字段名是 verticalAlignment。
- 写入等待服务端检查点返回后，重新检查来源与目标内容/格式、工作簿可编辑和工作表保护状态。测试覆盖等待期间外部改值/改格式，必须零写入。
- Excel 条件复制旧分支忽略 copyType/transpose/skipBlanks，现共用 copyRangePlan；真实写入片段分别保存恢复点，跳过格不会被撤销覆盖。all/formulas/formats 条件复制明确拒绝，不降级为值复制。
- 恢复缺陷修复：首个宿主写入前先建立恢复记录；只将真正尝试过的动作归属本次执行。重叠格式第二步中途失败时保存真实部分结果并能逆序恢复；等待期间出现人工修改时不把它误记为自己的写后结果。回读不到目标格式时不得把缺失字段登记为成功证据。

### 本批最终自动化（2026-10-18）

| 检查 | 结果 | 证据 |
|---|---|---|
| 共享层 | 593 passed | outputs/wps-capabilities-next-core.log |
| Excel | 1317 passed / 3 skipped | outputs/wps-capabilities-next-excel.log |
| WPS（变异还原后全量） | 938 passed / 3 skipped | outputs/wps-capabilities-next-wps.log |
| 服务端 | 1510 passed / 1 skipped；compileall exit 0 | outputs/wps-capabilities-next-server.log |
| 两端构建（含各端 TypeScript） | exit 0 | outputs/wps-capabilities-next-excel-build.log、outputs/wps-capabilities-next-build.log |
| 反向验证 | 14 项均业务断言失败、SHA256 还原后通过 | outputs/wps-capabilities-next-guards-final.log；详细 baseline/red/restored 在 outputs/wps-capability-guards/ |

14 项覆盖：OLE 颜色、数字格式恢复、复制来源变化、跳过空白、重叠数据拒绝、动态工具模板、服务端复制变体、计算格式来源、等待期间来源/目标变化、过滤链接行号、垂直恢复、恢复归属、重叠格式部分失败。最初 vertical-restore 的变异确实导致业务异常，但驱动要求 AssertionError，未被计作通过；将恢复正向断言改为 resolves 后完整重跑 14 项，保留首次失败日志，不以它充当通过证据。

Excel 条件复制专项 4 例在有状态宿主中覆盖值/链接、转置、过滤、跳过空白、真实撤销及拒绝不支持模式；所属专项组合 44 passed，见 outputs/excel-filtered-copy.log。恢复新增故障用例先复现再修复，首次失败见 outputs/wps-recovery-defects-before.log。

本轮未运行真实 Excel/WPS、未运行 online/cross-process；已有自动化全绿不替代本批实机验收。未提交、推送、发布。

本轮尚未实现整表/区域结构、排序/去重/筛选视图、自适应尺寸、对象与规则恢复。尤其 deleteWorksheet 仍未实现；没有新增 Delete 调用，也没有结构恢复能力的虚假占位。真实宿主验收未执行。

### 前一批实际验证（2026-10-18）

| 检查 | 本轮结果 | 证据 |
|---|---|---|
| WPS 全量 | 911 passed / 3 skipped | `outputs/wps-capabilities-wps-final.log` |
| Excel 全量 | 1314 passed / 3 skipped | `outputs/wps-capabilities-excel-final.log` |
| 共享层 | 587 passed | `outputs/wps-capabilities-core-final.log` |
| 服务端全量 | 1507 passed / 1 skipped | `outputs/wps-capabilities-server-full.log` |
| 两端构建、服务端 compileall | exit 0 | `outputs/wps-capabilities-excel-build.log`、`outputs/wps-capabilities-wps-build-final.log` |
| 反向验证 | 7 项均触发业务断言失败、SHA256 还原后通过 | `outputs/wps-capabilities-guards-verified.log` 前四项；`outputs/wps-capabilities-guards-final.log` 后三项 |

真实面板调用层新增“保存动态复制工具保留原动作，复用按当前授权重新读取来源”，证明 `originalPlan` 接线而不只证明执行器持有它。新填充/数字格式读取为可选宿主字段：宿主缺这些接口时原有字体动作仍可执行和恢复；新增动作自身仍要求完整元数据。

验证期间脚本一次在文件写入/还原检查处中断；执行器完整哈希与原版本一致，未遗留变异。已收紧脚本，仅在变异实际写入成功后执行恢复检查，避免把写入失败误诊为并发编辑；剩余三项单独重跑通过。不使用旧 red 日志代替本轮结果。

本批没有运行真实 Excel/WPS 工作簿，没有提交、推送或发布。**用户要求的“全部补齐”尚未完成**；下列 25 类未实现动作不能写成“只差实机”。

### 仍需完成的能力组

| 能力组 | WPS 当前缺口 | 完成验收条件 |
|---|---|---|
| 工作表结构 | createWorksheet、deleteWorksheet、activateWorksheet | 当前执行器/恢复器仅面向区域；必须补结构快照、按稳定身份重定位授权、原生副本恢复与对象/引用副作用核对。删除不能用隐藏代替，不能只调 Delete 就宣称可撤销。 |
| 区域结构 | insertRange、deleteRange、mergeCells、unmergeCells | 原值/格式及引用迁移的回读与完整恢复；未授权表不得扩大读取/写入。 |
| 数据整理 | sortRange、removeDuplicates、filterRange、clearFilter | 保持值、公式和格式关联，核对顺序/行集/筛选规则，撤销恢复原状态；不能把仅重排值冒充原生排序。 |
| 基础格式 | autofit；resizeRange.columnWidth | 边框各边独立快照；垂直对齐已有自动化实现，实机待验；列宽已记录真实量化恢复漂移，不能靠容差隐藏。 |
| 规则与视图 | setDataValidation、setConditionalFormat、freezePanes、setHyperlink；clearRange 的 all/formats/hyperlinks | 对应规则/视图/链接状态快照与归属清理，不误删用户原有内容或规则。 |
| 表格对象与可视对象 | createTable、addNamedRange、createChart、createPivotTable、addComment、addNote、addImage、addShape | 原生对象创建、唯一身份、回读、删除或恢复；不是只记录配置。 |
| 复合输出 | splitGroupAggregate；computeValues 的新表输出/多计算动作组合 | 执行器接入受控计算结果、多表输出身份与恢复；已有查询/Python 工具不等于这些动作已实现。 |
| 复制剩余变体 | copyRange.all/formulas/formats | 相对公式迁移、源/目标格式和合并/规则语义保持；当前部分支持不能外推。 |

删除工作表仍是未完成的产品能力。WPS 官方有 [Worksheet.Delete](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Worksheet/member/Delete)，当前缺的是 EB 的完整闭环；不是宿主没有删除 API。不得仅放开白名单。原生 VBA/宏、任意 JS 仍不在已授权实现范围内；受控脚本生成器仍沿自己的动作子集。

## 0. 口径与禁止事项（先读）

三个维度**互相独立**，不得互相推导：

1. **配置声明**：`capabilities.json` 的值（`supported` / `unverified` / 未列出）。
2. **执行器实现**：`wpsPlanExecutor.ts` 里是否**真的**做这件事（含支持的**属性子集**）。
   另有**第二道闸** `wpsWrite.allowedActions`（配置放开），两道闸**缺一不可**
   （`assertAllowedAction`：`WPS_SUPPORTED_ACTION_KEYS` ∈ 且 `allowedActions` 包含）。
3. **实机证据**：真实宿主上是否跑过，附宿主版本与出处。**替身/自动化证据 ≠ 实机证据。**

**禁止**：

- 从 `capabilities.json` 的 `state:"supported"` **直接推导**"实机闭环通过"。该字段只表达
  "按现有证据它可进计划"，**不等于**真机跑过。实机结论以本文件「实机证据」列与出处为准。
- 把**未实现**写成"待实机验收"。未实现 = **未实现**（归因见 §3）。
- 把**未核查的成因**写成"已确认的宿主限制"。
- 靠**放开白名单**或**删除 Excel 能力**来做"表面一致"。
- 把"区域/属性**部分**通过"外推成"整类动作已支持"（见 §2 属性子集）。

> 现状一致性：WPS 侧 **`supported` ⟺ 执行器已实现 ⟺ 白名单已放开**，三者当前逐项一致
> （`wpsPlanExecutor.test.ts` 有一条「白名单与已实现动作集 lockstep」钉住）。
> 故**不存在**"配置标 supported 但执行器不认"的项；反过来说，凡配置标 `unverified` 的项，
> 执行器**确实未实现**，不是"实现了只差实机"。

---

## 1. 逐项能力差距表

列含义：**配置**（Excel 16.0 / WPS 12.0 的 `state`）、**WPS 执行器**（是否实现 + 属性子集）、
**WPS 白名单**（`allowedActions` 是否放开）、**自动化证据**、**实机证据（宿主版本）**、
**剩余差距与验收要求**。

| 动作 | 配置 Excel / WPS | WPS 执行器实现（属性子集） | WPS 白名单 | 自动化证据 | 实机证据（宿主版本） | 剩余差距与验收要求 |
| --- | --- | --- | --- | --- | --- | --- |
| writeValues | supported / supported | ✅ 全量（2D 值矩阵；**禁止**经此注入 `=` 公式） | ✅ | `wpsPlanExecutor.test.ts`（预览→确认→回读→撤销；空串≠null）；`wpsWriteMatrix.test.ts`（文本保真）；`WpsAgentPanel.test.tsx`（面板端到端） | **未执行**（WPS 真实写入链路 2026-10-12 触发过副本写回故障，修复版未复验） | 真实 WPS：副本→确认写回→回读→撤销复验；数字文本保真已在 2026-10-12 修复，待真机确认 |
| writeFormulas | supported / supported | ✅ 全量（**仅 A1 同形矩阵**；`formulaR1C1` 明确拒绝） | ✅ | `wpsPlanExecutor.test.ts`（按 `Formula` 而非 `Value2` 回读；回读不一致不标 verified；R1C1 拒绝） | **未执行** | 真实 WPS：公式文本回读 + 计算值核对 + 撤销复验 |
| clearRange:contents | supported / supported | ✅ `applyTo=contents`；支持按唯一表头字段 filters 选行 | ✅ | `wpsPlanExecutor.test.ts`（同快照安全原语；空单元格渲染成一行改动） | **未执行** | 真实 WPS：清空后回读为空白、撤销恢复复验 |
| clearRange:all | supported / **unverified** | ❌ 未实现（`applyTo !== contents` 即抛错） | ❌ | `wpsPlanExecutor.test.ts`「拒绝结构删除/未授权动作」 | **未执行** | **需开发**：`applyTo=all` 会连格式一起清；须先定回读与撤销口径 |
| clearRange:formats | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：格式清除后的回读/撤销 |
| clearRange:hyperlinks | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：确认实际移除的对象与回读 |
| setFont | supported / supported | ✅ 子集：**`bold` / `italic` / `name` / `size` / `color`**；`underline` **拒绝**（哨兵误用布尔，回读 2≠true） | ✅ | `wpsFormatExecutor.test.ts`（bold 闭环；name/size/color 闭环；混合格式恢复）；`wpsFormatFlow.test.ts`；`WpsAgentPanelFormatReview.test.tsx` | **属性级已确认**：WPS ET `12.0` 单格探测 bold/italic/name/size/color 写→回读→恢复通过（`WPS_FORMAT_PROBE_EVIDENCE_2026-10-13.md`）。**动作闭环实机未执行** | 修 `underline` 枚举哨兵后**先实机探测再实现**；动作闭环（预览→确认→写回→回读→撤销）真机复验 |
| setAlignment | supported / supported | ✅ `horizontal` / `vertical` / `wrapText`；垂直对齐缺恢复属性时拒绝 | ✅ | `wpsFormatExecutor.test.ts`（horizontal/wrapText 闭环）；`wpsVerticalAlignment.test.ts`（逐格垂直对齐回读、冲突和恢复）；`WpsAgentPanelFormatReview.test.tsx` | **属性级已确认**：WPS ET `12.0` 单格探测 `HorizontalAlignment(xlCenter)`/`WrapText` 通过。**动作闭环实机未执行** | 垂直对齐与混合格式恢复的实机验收待执行 |
| resizeRange | supported / supported | ✅ 子集：**仅 `rowHeight`**；`columnWidth` **拒绝**（量化漂移，恢复未通过） | ✅ | `wpsFormatExecutor.test.ts`（rowHeight 闭环；带 `columnWidth` 明确拒绝）；`wpsFormatExecutor.test.ts` 合成写前现场 | **属性级已确认**：WPS ET `12.0` 单格探测 `RowHeight` 通过。`ColumnWidth` **探测失败**（写 12 回读 12.0089…，恢复漂移 ≈0.0089）。**动作闭环实机未执行** | 查明列宽单位/量化与重复设置漂移后再决定是否放开；不得用 epsilon 掩盖 |
| setFill | supported / supported | ✅ 逐格填充颜色与图案快照、回读、属性级撤销；WPS OLE 颜色 | ✅ | wpsExtendedFormats.test.ts | 本批实机未执行 | 主表填充可见、确认不清业务填充、撤销恢复混合原格式 |
| setNumberFormat | supported / supported | ✅ 逐格数字格式快照、回读与恢复 | ✅ | wpsExtendedFormats.test.ts | 本批实机未执行 | 百分比/日期/小数混合格式执行及恢复核对 |
| setBorders | supported / supported | ✅ 子集：明确 A1 区域的外框/内部横纵线；归一线型、颜色、粗细；展开为逐格边界 | ✅ | `wpsBorders.test.ts`：混合主题/自动/RGB 颜色、部分失败恢复、属性冲突、授权、相邻边界归属；双端标记测试 | **未执行**；旧探测 abort 不能算已验证 | 读不到 ThemeColor/TintAndShade/ColorIndex 等必要来源信息则写前拒绝；跨动作从相邻格重复拥有同一物理边界暂拒绝；真实宿主闭环待验 |
| autofit | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行**（`EntireRow/Column.AutoFit` 未执行，避免改整行/列） | **需开发**：自适应列宽无法稳定回读，须先确定可回读判据 |
| setDataValidation | supported / **unverified** | ❌ 未实现 | ❌ | `wpsPlanExecutor.test.ts`（未授权/未实现动作整体拒绝） | **未执行** | **需开发**：规则回读 |
| setConditionalFormat | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：规则回读（Excel 侧含条件格式区域撤销受限，口径另列） |
| freezePanes | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：冻结状态回读 |
| setHyperlink | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行**（`safety.allowHyperlink` 只管 `HYPERLINK()` 公式函数，与本动作无关） | **需开发**：超链接写入/回读（需独立安全评估） |
| createWorksheet | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：新建表及恢复 |
| deleteWorksheet | supported / supported（固定检查点通道） | ✅ 纯删表、本机 XLSX、全表授权；完整归档后原生批量删除 | 独立 worksheetCheckpoint 白名单 | 服务端状态模拟 + 双端面板路由 | **未执行** | 仅打开独立恢复副本，不原地撤销；具体限制见 2026-10-19 更新 |
| insertRange | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：引用迁移 |
| deleteRange | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：引用迁移 |
| mergeCells | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：影响范围与恢复 |
| unmergeCells | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：影响范围与恢复 |
| createTable | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：表格对象身份/回读 |
| addNamedRange | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：名称管理 API 探测 |
| activateWorksheet | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：非写入，但需确认授权边界与恢复语义 |
| writeTable | supported / supported | ✅ 表头与数据写入，加粗/填充表头；不是表格对象 | ✅ | wpsWriteTable.test.ts；共享 writeTableLayout | 本批实机未执行 | 三类实际效果逐项登记、回读、撤销；原公式/格式恢复 |
| copyRange | supported / supported（子集） | ✅ values/link、条件选行、转置、跳过空白；其他变体未实现 | ✅ | wpsCopyRange.test.ts；规划出口子集测试 | 本批实机未执行 | 来源变化零写入、未授权零读取、目标公式保留、撤销与动态工具复用 |
| sortRange | supported / supported（子集） | ✅ 明显 A1 区域、稳定排序、空白恒末位；**含公式 / 数据行格式不一致 / 与表格对象重叠 / 宿主无 `ListObjects` 时写入前拒绝** | ✅ | `tableTransforms.test.ts`（16 条纯判据）；`wpsSortRange.test.ts`（14 条真实网格：执行→回读→撤销、空白末位、表头不动、四条拒绝、预览后变化零写入、改后撤销冲突、未授权）；反向变异 `scripts/check_wps_sort_guards.mjs` 4 项 | **未执行** | 真实 WPS：排序后逐行顺序回读、撤销后逐格恢复；确认 `ListObjects` 在本机版本存在（缺它会被拒绝，不是静默放行） |
| removeDuplicates | supported / **unverified** | ❌ 未实现（`prepareWpsPlan` 明确拒绝并说明原因） | ❌ | `wpsSortRange.test.ts`「removeDuplicates：明确拒绝并说明原因」；共享判据 `duplicateRowPartition` 已实现并测试 | **未执行** | **需开发**：去重是删除整行并上移下方数据，恢复必须先经实机验证 `EntireRow.Insert` 与整块还原；未验证前不提供入口，也不用"压实值+清空尾行"冒充 |
| filterRange | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：筛选规则回读 |
| clearFilter | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：清除筛选规则状态回读 |
| computeValues | supported / supported（子集） | ✅ 现有授权目标表，本地 query_table，展开值和表头样式 | ✅ | `wpsComputedPlan.test.ts` | **未执行** | 来源须在计划最先读取；单次上限来自 wpsWrite.maxCells；不能新建目标表；缺显示/格式证据拒绝 |
| splitGroupAggregate | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：多表输出与动态验收 |
| createChart | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：数据源/位置回读 |
| createPivotTable | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：数据源/位置回读 |
| addComment | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：批注 API |
| addNote | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：备注 API |
| addImage | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：图片插入/位置回读 |
| addShape | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：形状插入/位置回读 |

---

## 2. WPS 写入白名单与属性子集

**白名单**（`config/capabilities.json` → `wpsWrite.allowedActions`，执行器另有
`WPS_SUPPORTED_ACTION_KEYS` 同集）：

`writeValues`、`writeFormulas`、`clearRange:contents`、`setFont`、`setAlignment`、`resizeRange`、`setFill`、`setNumberFormat`、`writeTable`、`copyRange`、`computeValues`、`setBorders`、`sortRange`。

**属性子集**（`wpsFormatCapability.ts`，只依据实机探测结论，不推断）：

- `setFont`：`bold` / `italic` / `name` / `size` / `color`。`underline` **拒绝**。
- `setAlignment`：`horizontal` / `vertical` / `wrapText`。垂直对齐实现证据与实机证据分开记录。
- `resizeRange`：仅 `rowHeight`。`columnWidth` **拒绝**。

> 单格探测**不证明**混合格式恢复正确性；逐格/逐行快照与恢复仍需真机另行验收。

**能力探测（PR1/格式探测入口）**：`wpsFormatProbe.ts` / `WpsFormatProbePanel.tsx` 提供
**只读探测**入口（写哨兵→回读→恢复→回读），**不是**动作可用的证据；探测通过只说明
"该属性可读写可恢复"。

---

## 3. 实机证据（逐项，不外推）

> 只列**有出处**的项；无出处一律"未执行"。这里的"已确认"指**真实宿主上跑过的具体场景**，
> 不覆盖未测属性、未测动作。

**WPS（ET，`Application.Version` 报 `12.0`；本机安装 12.1.0.28488）**

- **加粗 / 水平居中（属性级）**：真实 WPS 上对单格 `A1` 探测 `Font.Bold` 与
  `HorizontalAlignment(xlCenter)`，写哨兵→回读一致→恢复→回读回到原值，**通过**。
  同批通过：`Font.Italic` / `Font.Name` / `Font.Size` / `Font.Color` / `WrapText` / `RowHeight`。
  出处：`docs/WPS_FORMAT_PROBE_EVIDENCE_2026-10-13.md`（原始 JSON
  `outputs/wps-format-probe-20261013/user-reported-051653.json`）。
  **不通过**：`Font.Underline`（哨兵误用布尔）、`ColumnWidth`（恢复漂移）；
  **未测**：`Interior`/`Borders`（因 abort 跳过）、`VerticalAlignment`（哨兵同值跳过）、`AutoFit`（未执行）。
- **审阅写回**：真实 WPS 上触发过"生成右侧副本并排审阅 → 确认写回"链路（2026-10-12 用户实机反馈：
  副本生成时数字文本被转数值）。修复（`wpsWriteMatrix` 临时文本格式后恢复）后的
  "副本 → 写回 → 回读 → 撤销"真机复验**未执行**。出处：`docs/WPS_TEXT_WRITE_FIX_2026-10-12.md`。
- **撤销**：真机撤销仅由上述写回故障间接触发过（`docs/WPS_TASK_LEDGER_RECOVERY_2026-10-12.md` 记录了
  真机"确认写回"无法持久登记现场的故障并已修）；修复后的真机撤销闭环**未执行**。
- **W3 写入/公式恢复/刷新后撤销/异常中断**：`docs/WPS_W3_A_VALIDATION_2026-09-23.md` 明确标注
  **"真实 WPS 12.1.0.28488 上仍需按清单复验"**。

**Excel（真实宿主 `ProPlus2024Volume` `16.0.17932.20076`）**

- **格式动作**：真实 Excel 上跑过——该 LTSC 宿主最高 `ExcelApi 1.17`，旧实现依赖 1.19 的
  `getDisplayedCellProperties` ⇒ **`setFont` 写前被拒**。已补 ≤1.17 兼容路径（自动化 `excelFormatCompatPath.test.ts`
  11 ✓ 覆盖**加粗 / 水平居中** / 混合字号 / 行高 / 条件格式拒绝 / 合并区域拒绝等）。
  **兼容路径的真实 Excel 2024 复验未执行**（步骤见 `docs/EXCEL_FORMAT_COMPAT_2026-10-13.md` §6）。出处：`docs/EXCEL_FORMAT_COMPAT_2026-10-13.md`。
- **双端消息结果卡 / 卡内撤销**：此前已有用户确认的双端结果卡/格式执行/撤销场景证据，
  本轮改动**不推翻**（见 `docs/DUAL_HOST_PLAN_UI_PARITY_HANDOFF_2026-10-14.md`「验证与交付」）。
  本轮**新增/改动**的界面与交互单列"实机未复验"。

---

## 4. 本轮（2026-10-14 续）新增、单列实机待验收

- **WPS「验证后保存为个人工具」确认流程**（本轮新接线，与 Excel 共用 `SaveToolDialog` + `useExecutionApproval`）：
  - 复用 `analyzeToolEligibility` 的固化资格：`writeValues`/`writeTable` 需**固定内容**确认，
    删除/清空/覆盖后需**破坏性**确认；内嵌图片**阻断**。
  - 所有确认项默认不勾选；未勾选不可保存；`eligibility.blocked` 的计划**不出现**勾选框、确认恒禁用。
  - 保存资格绑定**实际已验证的计划版本与执行身份**（`verifiedPlanForSaveRef`）；打开对话框后若
    **计划被替换或执行被撤销**，确认时**重新检查**（`saveToolInvalidationReason`）并禁用确认。
  - 取消不留新工具、不留批准状态；重开/切换候选**复位确认项**；编辑工具计划模板后**重新分析**保存资格。
  - `createTool` 的审批门**未放宽**，不预置确认项为真。
  - 自动化：`WpsAgentPanel.test.tsx`「验证后保存为个人工具」7 条（含未确认拒存/勾选后成功、
    破坏性计划拒/允、取消复位、我的工具重用重新校验+预览、撤销后旧候选失效、编辑后重算资格）。
  - **实机：未验收**（真实 WPS 面板上保存→我的工具→重用→预览→撤销后的失效提示，均未在真机跑过）。

---

## 5. 后续实施顺序（按用户任务，先易后难）

> 区分 **需开发**（要写执行器 + 回读 + 撤销）与 **仅需探测**（先查宿主行为，再决定是否开发）。

1. **基础格式**
   - 已实现待本批实机验收：`setFill`、`setNumberFormat`。
   - **需开发**：`setBorders`、`autofit`。
   - **需先探测**（后开发）：`setFont.underline`（枚举哨兵）、
     `resizeRange.columnWidth`（单位/量化漂移）、`clearRange:all`/`formats`/`hyperlinks`（实际移除对象与回读）。
2. **复制 / 排序 / 筛选 / 去重**：copyRange 的 values/link 子集已实现；其余复制变体以及 `sortRange`、`removeDuplicates`、`filterRange`、`clearFilter` 仍**需开发**。
3. **工作表结构与对象能力**：`createWorksheet`、`deleteWorksheet`、`insertRange`、`deleteRange`、
   `mergeCells`、`unmergeCells`、`createTable`、`addNamedRange`、`activateWorksheet`、
   `splitGroupAggregate`、`createChart`、`createPivotTable`、`addComment`、`addNote`、
   `addImage`、`addShape` —— **全部需开发**（引用迁移、整表恢复、对象身份/位置回读）。`writeTable` 已实现矩阵/表头样式，不属于原生表格对象；删除工作表是用户当前明确提出的下一项，不能用其他能力补齐代替。

每批的验收要求：① 执行器实现 + 属性级守卫；② 值/格式/对象回读与撤销恢复；③ 双端面板/执行器自动化证据；
④ 真实宿主实机验收（含宿主版本）；⑤ 未实现项**不得**写成"待实机"。

---

## 6. 本轮相关（非本表范围）

- 双端计划卡/审阅交互对齐、运行中暂停主题色统一、"编辑计划"命名统一（见交接文档与代码改动）。
- WPS「验证后保存工具 → 我的工具 → 编辑计划模板 → 使用工具 → 当前授权下重新校验与预览」闭环（§4）。


### 2026-10-19 删表能力更新（覆盖上文历史“未实现”状态）

Excel/WPS 均接入纯 `deleteWorksheet` 检查点批处理：固定宿主适配器取得完整本地 XLSX 归档，再逐表删除及读回表名。用户选择的恢复方式是打开归档的独立副本，原表和后续编辑保留；不按旧区域快照重建删表，不声称原地撤销或保留表的公式结果不变。范围执行器的白名单不放开删表，规划上界取两个执行器的动作并集。

限制：Windows 本地服务、原表已保存为本机 XLSX、删除目标在勾选范围内、完整归档所需的额外表已单独取得仅本地备份许可、保留至少一个可见工作表；混合计划、云 URL、含图表表等拒绝。恢复需原表仍打开以定位宿主。完整归档不自动清除。自动化及构建验证见检查点日志；真实 Excel/WPS COM 删除、恢复文件打开和内容完整性验收仍未执行。此更新不代表其余结构或对象动作已实现。
