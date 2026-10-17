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
| clearRange:contents | supported / supported | ✅ 仅 `applyTo=contents` 且无 `filters` | ✅ | `wpsPlanExecutor.test.ts`（同快照安全原语；空单元格渲染成一行改动） | **未执行** | 真实 WPS：清空后回读为空白、撤销恢复复验 |
| clearRange:all | supported / **unverified** | ❌ 未实现（`applyTo !== contents` 即抛错） | ❌ | `wpsPlanExecutor.test.ts`「拒绝结构删除/未授权动作」 | **未执行** | **需开发**：`applyTo=all` 会连格式一起清；须先定回读与撤销口径 |
| clearRange:formats | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：格式清除后的回读/撤销 |
| clearRange:hyperlinks | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：确认实际移除的对象与回读 |
| setFont | supported / supported | ✅ 子集：**`bold` / `italic` / `name` / `size` / `color`**；`underline` **拒绝**（哨兵误用布尔，回读 2≠true） | ✅ | `wpsFormatExecutor.test.ts`（bold 闭环；name/size/color 闭环；混合格式恢复）；`wpsFormatFlow.test.ts`；`WpsAgentPanelFormatReview.test.tsx` | **属性级已确认**：WPS ET `12.0` 单格探测 bold/italic/name/size/color 写→回读→恢复通过（`WPS_FORMAT_PROBE_EVIDENCE_2026-10-13.md`）。**动作闭环实机未执行** | 修 `underline` 枚举哨兵后**先实机探测再实现**；动作闭环（预览→确认→写回→回读→撤销）真机复验 |
| setAlignment | supported / supported | ✅ 子集：**`horizontal` / `wrapText`**；`vertical` **拒绝**（哨兵与原值同值，未验证） | ✅ | `wpsFormatExecutor.test.ts`（horizontal/wrapText 闭环；带 `vertical` 明确拒绝）；`WpsAgentPanelFormatReview.test.tsx` | **属性级已确认**：WPS ET `12.0` 单格探测 `HorizontalAlignment(xlCenter)`/`WrapText` 通过。**动作闭环实机未执行** | 换合法 `xlVAlign` 哨兵先探测 `vertical`；动作闭环真机复验 |
| resizeRange | supported / supported | ✅ 子集：**仅 `rowHeight`**；`columnWidth` **拒绝**（量化漂移，恢复未通过） | ✅ | `wpsFormatExecutor.test.ts`（rowHeight 闭环；带 `columnWidth` 明确拒绝）；`wpsFormatExecutor.test.ts` 合成写前现场 | **属性级已确认**：WPS ET `12.0` 单格探测 `RowHeight` 通过。`ColumnWidth` **探测失败**（写 12 回读 12.0089…，恢复漂移 ≈0.0089）。**动作闭环实机未执行** | 查明列宽单位/量化与重复设置漂移后再决定是否放开；不得用 epsilon 掩盖 |
| setFill | supported / **unverified** | ❌ 未实现（明确拒绝，不静默忽略） | ❌ | `wpsFormatExecutor.test.ts`「setFill/setBorders/autofit 仍拒绝」；`wpsPlanExecutor.test.ts` | **未执行**（探测因基线恢复失败 abort，填充组未测） | **需开发** + 先补实机探测（新独立基线） |
| setNumberFormat | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：数字格式回读/恢复 |
| setBorders | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行**（探测 abort，边框组未测） | **需开发** + 先补实机探测 |
| autofit | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行**（`EntireRow/Column.AutoFit` 未执行，避免改整行/列） | **需开发**：自适应列宽无法稳定回读，须先确定可回读判据 |
| setDataValidation | supported / **unverified** | ❌ 未实现 | ❌ | `wpsPlanExecutor.test.ts`（未授权/未实现动作整体拒绝） | **未执行** | **需开发**：规则回读 |
| setConditionalFormat | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：规则回读（Excel 侧含条件格式区域撤销受限，口径另列） |
| freezePanes | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：冻结状态回读 |
| setHyperlink | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行**（`safety.allowHyperlink` 只管 `HYPERLINK()` 公式函数，与本动作无关） | **需开发**：超链接写入/回读（需独立安全评估） |
| createWorksheet | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：新建表及恢复 |
| deleteWorksheet | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：整表删除恢复/撤销 |
| insertRange | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：引用迁移 |
| deleteRange | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：引用迁移 |
| mergeCells | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：影响范围与恢复 |
| unmergeCells | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：影响范围与恢复 |
| createTable | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：表格对象身份/回读 |
| addNamedRange | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：名称管理 API 探测 |
| activateWorksheet | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：非写入，但需确认授权边界与恢复语义 |
| writeTable | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：表格对象 API 写入/回读 |
| copyRange | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：`copyType`/转置/跳过空白变体逐项 |
| sortRange | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：排序后真实顺序回读 |
| removeDuplicates | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：去重后行集合回读 |
| filterRange | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：筛选规则回读 |
| clearFilter | supported / **unverified** | ❌ 未实现 | ❌ | 同上 | **未执行** | **需开发**：清除筛选规则状态回读 |
| computeValues | supported / **unverified** | ❌ 未实现（执行时求值路径未接入） | ❌ | `wpsPlanExecutor.test.ts` | **未执行** | **需开发**：求值路径接入 WPS 执行器 |
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

`writeValues`、`writeFormulas`、`clearRange:contents`、`setFont`、`setAlignment`、`resizeRange`。

**属性子集**（`wpsFormatCapability.ts`，只依据实机探测结论，不推断）：

- `setFont`：`bold` / `italic` / `name` / `size` / `color`。`underline` **拒绝**。
- `setAlignment`：`horizontal` / `wrapText`。`vertical` **拒绝**。
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

1. **基础格式**（与已有 6 个动作同一批，风险最低）
   - **需开发**：`setFill`、`setNumberFormat`、`setBorders`、`autofit`。
   - **需先探测**（后开发）：`setFont.underline`（枚举哨兵）、`setAlignment.vertical`（换 `xlVAlign` 哨兵）、
     `resizeRange.columnWidth`（单位/量化漂移）、`clearRange:all`/`formats`/`hyperlinks`（实际移除对象与回读）。
2. **复制 / 排序 / 筛选 / 去重**：`copyRange`、`sortRange`、`removeDuplicates`、`filterRange`、`clearFilter`
   —— **全部需开发**（顺序回读、行集合回读、规则回读）。
3. **工作表结构与对象能力**：`createWorksheet`、`deleteWorksheet`、`insertRange`、`deleteRange`、
   `mergeCells`、`unmergeCells`、`writeTable`、`createTable`、`addNamedRange`、`activateWorksheet`、
   `computeValues`、`splitGroupAggregate`、`createChart`、`createPivotTable`、`addComment`、`addNote`、
   `addImage`、`addShape` —— **全部需开发**（引用迁移、整表恢复、对象身份/位置回读，风险最高，最后做）。

每批的验收要求：① 执行器实现 + 属性级守卫；② 值/格式/对象回读与撤销恢复；③ 双端面板/执行器自动化证据；
④ 真实宿主实机验收（含宿主版本）；⑤ 未实现项**不得**写成"待实机"。

---

## 6. 本轮相关（非本表范围）

- 双端计划卡/审阅交互对齐、运行中暂停主题色统一、"编辑计划"命名统一（见交接文档与代码改动）。
- WPS「验证后保存工具 → 我的工具 → 编辑计划模板 → 使用工具 → 当前授权下重新校验与预览」闭环（§4）。
