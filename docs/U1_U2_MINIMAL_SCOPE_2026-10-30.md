# U1 / U2 最小可交付范围与推荐顺序（2026-10-30）

> 依据是**当前代码**，不是文档里的旧描述。每条都给了行号，便于复核。
> 结论先说：**U1 可以切出三片，第一片（`writeFormulas`）是一轮能收的；U2 表面小，实则动的是
> 「谁为数据作证」的信任模型，需要产品/安全点头。**

## 1. 现在到底卡在哪（两处门槛，缺一不可）

**门槛一：登记恢复材料时就已经是 `null`。**
`wpsResumeContext()`（`apps/wps-addin/src/wpsPlanExecutor.ts:1072-1092`）只在两种情形给 context：

1. `isTrustedLiteralPlan(prepared.plan)` 且**每个**动作都是 `writeValues` → `sources: []`（1075-1076）；
2. `sourceDependencySheets` 非空 + 每个动作都是 `writeValues` + `sources` 覆盖全部依赖表
   且与目标表**不重叠** + 源单元格总数 ≤ `taskRecovery.wpsResumeMaxSourceCells`（1078-1082）。

否则返回 `null` ⇒ 登记那一步写成 `...(context && recovery?.taskId ? {resume: {...}} : {})`
（1990 行）⇒ **记录里根本没有 `resume`** ⇒ 续跑时第一句就抛
"原执行未保留精确续跑材料"（1183）。

**门槛二：续跑时逐动作要求 `writeValues`。**
`prepareWpsResume`（1208-1215）对每个动作要求
`a.type === "writeValues" && d.type === "writeValues"`，并比对
`values / beforeValues / beforeFormulas` 三个矩阵与 `beforeFingerprint / previewFingerprint /
expectedFingerprint` 三个指纹。类型不对或矩阵形状不对一律"拒绝续跑"。

两处都是同一条判据的两端：**门槛一决定"有没有材料"，门槛二决定"材料能不能对上"。**

## 2. U1 的三片（按"一轮能不能收"切）

### U1a `writeFormulas`（**已实现，2026-10-30 第十六轮**）

- 现状：被 1075/1079/1208 三条同时排除。可是 `writeFormulas` 的恢复点结构与 `writeValues`
  **同构**（行×列矩阵），差别只在"写的是公式"。
- 最小改动面：
  1. `WpsStoredAction` 增一个"写后公式矩阵"（或让 `values` 字段按 `actionType` 解释，
     但显式字段更好读）；
  2. `wpsResumeContext` 的类型判定放宽到 `writeValues | writeFormulas`；
  3. `prepareWpsResume` 的逐动作分支按 `actionType` 选矩阵比对（`matrixFingerprintV2(values, formulas)`
     已经存在，直接复用）；
  4. 恢复点描述符加 `actionType`。
- 判据（`wpsExecutionResume.test.ts` 扩两条）：公式计划写一半中断 → 只续剩余动作、已写的不重写；
  公式矩阵被改写 → 拒绝续跑。两条都读不同数据源（一个是写入调用次数、一个是记录里的矩阵指纹）。
- 不做的事：不动来源锚点、不动格式、不动结构动作。

> **状态（2026-10-30）**：已实现并双端接线。判据集中在 `packages/core/src/formulaWriteback.ts`
> （公式身份 = 公式文本，`formulaCellSite`/`formulaMatrixSite`/`formulaResumeGate`/`formulaSheetReferences`），
> WPS 走 `prepareWpsResume`/`inspectWpsUndoRecord`，Excel 走 `preflightPlanActions` 的公式跨表授权闸 +
> `classifyOfficeActionState`（本来按公式文本比）。协议两端同步（TS/Python 读同一份
> `formulaActionCases`）。详见 `计划表.md` 第十六轮。

### U1b 格式动作（`setFont` / `setAlignment` / `setFill` / `setNumberFormat` / `setBorders` / `resizeRange` / `sortRange`）

- 现状：恢复点**不是值矩阵**，是"每个格式属性的写前快照"。`WpsStoredAction` 装不下，
  描述符里的 `rowCount/columnCount/cellCount/expectedFingerprint` 语义也全是按值矩阵定义的。
- 最小可交付：先把"格式属性快照 + 格式指纹"做成 `packages/core` 的**纯判据**（两端读同一份），
  再改存储与协议。顺序不能反：先有判据才能写断言。
- 显式不做：`removeDuplicates`（875-880 现在直接拒，理由是 `EntireRow.Insert` 恢复链路没实机取证）、
  以及任何需要行插入/删除的结构动作。

### U1c `copyRange` —— **可能已经支持，但没有任何用例证明**

- `copyRange` 在 939-961 被展开成 `writeValues` + 来源快照，形状上**满足**门槛一；
  244 行的注释则说明它的固定值在 `planHasBakedValues` 眼里"没有本机依据"。
- 所以它到底是"已支持"还是"形状像但被来源锚点挡住"，**现在没有证据**。
- 动作：先补**一条**用例（copyRange 计划中断 → 续跑），绿了就把它从"不支持"里划掉，
  红了再看拒因落在哪一条判据上。不要先写实现。

## 3. U2：缺来源快照的 SDK Python 结果计划

- 卡点很具体：`sourceAnchorRequired = planHasBakedValues(plan) && (requireSourceAnchor ||
  Boolean(plan.localResultRef))`（867-868）—— 带 `localResultRef` 的计划**强制**要来源锚点；
  而 `isTrustedLiteralPlan` 只给"可信编译器封印"的计划开绿灯（1075），Python 结果不是；
  它也没有 `sourceDependencySheets`（结果是算出来的，不在任何表里）⇒ `wpsResumeContext` 返回 `null`。
- 表面看是"补一个 context 变体"，实际要回答的是：**"结果引用"能不能充当来源证据**。
  `packages/core/src/planSourceAnchor.ts` 现在明确写着"不能由模型响应自报豁免"。
- 最小可交付（如果要开）：`sources: []` + `targetBeforeFingerprints` 的新 context 变体，
  **只**对"有 `localResultRef` 且目标写前矩阵齐全"的计划开放；前置条件是先有
  "结果引用 ↔ 计划指纹"的 core 判据。
- 这是**信任模型层面的决定**，不是纯代码量问题 ⇒ 排在 U1 之后，且需要产品/安全确认。

## 4. 推荐顺序

| 顺序 | 项 | 理由 | 一轮能否收 |
| --- | --- | --- | --- |
| 1 | **U1a `writeFormulas`** | 判据现成、改动集中在两个函数、不触碰信任模型 | **已实现（第十六轮）** |
| 2 | **补 `copyRange` 一条用例** | 可能已经在支持范围内，用例先于实现，避免白做 | 能（半天） |
| 3 | **U1b 格式动作** | 必须先出"格式指纹"core 判据，再改存储与协议 | 分两轮 |
| 4 | **U2 结果引用** | 动的是"谁为数据作证"，要产品点头 | 需先决定 |

结构动作（`insertRows` / `removeDuplicates`）不在这条线上：它们的前置是**实机行插入恢复**，
属于真机验收而不是开发排序。
