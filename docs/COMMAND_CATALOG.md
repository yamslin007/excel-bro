# Excel Bro 精确命令清单

> 基准文档，供「基础模式」重构与「精确命令」扩展参考。
> 数据来源：SheetCopilot `xwAPI`（44 个原子操作，源自 SuperUser 真实 Q&A）、
> excel-bro `AnalysisPlan` 白名单（35 个动作）、InstructExcel（ExcelScript API 枚举，作能力面参照）。

## 1. 背景与目的

基础模式（未配置模型时的确定性命令路径）目前只暴露极少数精确命令（地址写值、写公式、清空、编辑已知值），其余 30 余个白名单动作只有「模型规划」能触发。要把精确命令从「基础模式的几个特例」提升为「一层独立、可复用的命令 DSL」，需要先有一份**经过真实需求验证的完整命令基准**。

本清单用 SheetCopilot `xwAPI` 的 44 个原子操作与 excel-bro 现有的 35 个动作做并集对照，标出缺口与可扩展优先级。

## 2. 数据来源说明

| 来源 | 内容 | 对本项目的价值 |
|---|---|---|
| SheetCopilot `xwAPI` | 44 个原子操作，每个带参数 + 语义 + 示例，来自 SuperUser 真实 Q&A 改编的 221 个任务 | **「精确命令」最成熟蓝本**，直接对标 |
| excel-bro `AnalysisPlan` | 35 个白名单动作（`server/app/models.py`） | 当前已实现能力 |
| InstructExcel `APIDesc.csv` | ExcelScript（Office Scripts）完整 API 枚举，类/方法/枚举字段逐条列出 | 太底层，仅作「能力面穷举」参照 |
| excelmanus | Agent 编排层，Excel 操作靠 `run_code` 生成 Python 脚本 | 非白名单，仅借鉴「验证门控 / 高风险审批」，不借鉴命令语法 |

## 3. 能力对照总表

`xwAPI` 44 操作逐行对照 excel-bro 动作，状态列 `✅ 已覆盖` / `🔴🟡🟢 缺口`。

### 3.1 写入 / 复制

| xwAPI 操作 | excel-bro 对应 | 状态 |
|---|---|---|
| `Write`（值/公式） | `writeValues` / `writeFormulas` | ✅ 已覆盖 |
| `CopyPaste` | `copyRange` | ✅ 已覆盖 |
| `CutPaste` | （`copyRange`+`deleteRange` 组合） | 🔴 缺口 |
| `AutoFill`（自动填充） | — | 🔴 缺口 |
| `RemoveDuplicate`（去重） | — | 🔴 缺口 |

### 3.2 行列操作

| xwAPI 操作 | excel-bro 对应 | 状态 |
|---|---|---|
| `InsertRow` / `InsertColumn` | `insertRange` | ✅ 已覆盖 |
| `Delete` | `deleteRange` | ✅ 已覆盖 |
| `MoveRow` / `MoveColumn` | — | 🔴 缺口 |
| `Clear` | `clearRange` | ✅ 已覆盖 |

### 3.3 筛选 / 排序

| xwAPI 操作 | excel-bro 对应 | 状态 |
|---|---|---|
| `Sort` | `sortRange` | ✅ 已覆盖 |
| `Filter` | `filterRange` | ✅ 已覆盖 |
| `DeleteFilter` | `clearFilter` | ✅ 已覆盖 |

### 3.4 格式

| xwAPI 操作 | excel-bro 对应 | 状态 |
|---|---|---|
| `SetFormat`（font/size/color/fill/bold/italic/underline/align 合一） | `setFill`/`setFont`/`setAlignment`/`setNumberFormat`/`setBorders` | ✅ 已覆盖（excel-bro 拆得更细） |
| `DeleteFormat`（仅删格式） | `clearRange(applyTo=formats)` | 🟡 需确认等价 |
| `SetDataType`（数据类型） | — | 🟡 缺口 |
| `SetCellMerge` | `mergeCells` / `unmergeCells` | ✅ 已覆盖 |
| `AutoFit` | `autofit` | ✅ 已覆盖 |
| `ResizeRowColumn` | `resizeRange` | ✅ 已覆盖 |

### 3.5 条件 / 验证 / 锁定

| xwAPI 操作 | excel-bro 对应 | 状态 |
|---|---|---|
| `SetConditionalFormat` | `setConditionalFormat` | ✅ 已覆盖 |
| `SetDataValidation` | `setDataValidation` | ✅ 已覆盖 |
| `SetCellLock`（单元格锁定） | — | 🟢 缺口 |

### 3.6 超链接 / 冻结

| xwAPI 操作 | excel-bro 对应 | 状态 |
|---|---|---|
| `SetHyperlink` | `setHyperlink` | ✅ 已覆盖 |
| `RemoveHyperlink` | — | 🟢 缺口 |
| `FreezePanes` | `freezePanes` | ✅ 已覆盖 |
| `UnfreezePanes` | `freezePanes(rows=0,columns=0)` | 🟢 需确认等价 |

### 3.7 工作表

| xwAPI 操作 | excel-bro 对应 | 状态 |
|---|---|---|
| `CreateSheet` | `createWorksheet` | ✅ 已覆盖 |
| `RemoveSheet` | `deleteWorksheet` | ✅ 已覆盖 |

### 3.8 图表 / 透视表

| xwAPI 操作 | excel-bro 对应 | 状态 |
|---|---|---|
| `CreateChart` | `createChart` | ✅ 已覆盖 |
| `SetChartType/Trendline/Title/Axis/HasAxis/Legend/HasLegend`、`Add/RemoveChartErrorBars`、`Add/RemoveDataLabels`、`SetChartMarker`（12 项细粒度） | `createChart` 一次性完成 | ⚠️ 粒度不同 |
| `CreatePivotTable` | `createPivotTable` | ✅ 已覆盖 |
| `CreateChartFromPivotTable` | — | 🟢 缺口 |

## 4. 缺口清单（按优先级）

| 缺口操作 | 频率 | 精确命令可行性 | 建议 |
|---|---|---|---|
| `RemoveDuplicate` 去重 | 🔴 高频 | ✅ 参数明确（源区域 + 依据列） | 优先补齐 |
| `AutoFill` 自动填充 | 🔴 高频 | ✅ 参数明确（源 + 目标范围） | 优先补齐 |
| `MoveRow` / `MoveColumn` 移动行列 | 🔴 高频 | ✅ 参数明确（源索引 + 目标索引） | 优先补齐 |
| `CutPaste` 剪切粘贴 | 🟡 中频 | ✅ 参数明确（源 + 目标） | 可补齐，或由 move 覆盖 |
| `SetDataType` 数据类型 | 🟡 中频 | ✅ 参数明确（区域 + 类型） | 可补齐 |
| `DeleteFormat` 仅删格式 | 🟡 中频 | ✅ 参数明确 | 先确认 `clearRange(applyTo=formats)` 是否已等价 |
| `RemoveHyperlink` 移除超链接 | 🟢 低频 | ✅ 参数明确（区域） | 可选 |
| `SetCellLock` 单元格锁定 | 🟢 低频 | ✅ 参数明确（区域 + 锁定） | 可选，涉保护需谨慎 |
| `UnfreezePanes` 取消冻结 | 🟢 低频 | ✅ 参数明确 | 先确认 `freezePanes(0,0)` 是否已等价 |
| `CreateChartFromPivotTable` | 🟢 低频 | ❌ 语义较重 | 留模型 |

## 5. excel-bro 独有能力（`xwAPI` 没有，保留）

- `writeTable`（结构化写表）
- `createTable`（Excel 表格对象）
- `addNamedRange`（命名区域）
- `addComment` / `addNote`（批注 / 备注）
- `addImage` / `addShape`（图片 / 形状）
- `splitGroupAggregate`（按字段拆表分组聚合，excel-bro 特有复合动作）
- `activateWorksheet`（切换工作表）

## 6. 精确命令可行性分级（扩展优先级）

### 第一档 ✅ 可直接扩展（参数天然明确，只差解析语法）

- **缺口补齐**：`RemoveDuplicate`、`AutoFill`、`MoveRow`、`MoveColumn`、`CutPaste`、`SetDataType`、`RemoveHyperlink`、`SetCellLock`、`UnfreezePanes`
- **已有动作扩展**：`insertRange`、`deleteRange`、`copyRange`、`clearFilter`、`setNumberFormat`、`setBorders`、`setAlignment`、`mergeCells`、`unmergeCells`、`resizeRange`、`autofit`、`freezePanes`、`createWorksheet`、`deleteWorksheet`、`addNamedRange`、`addComment`、`addNote`、`setHyperlink`、`createTable`

### 第二档 ⚠️ 需字段解析（操作明确，字段名要从表头匹配）

- `sortRange`（「按金额降序」）、`filterRange`（「筛出金额 > 100」）、`setDataValidation`（「B 列设下拉」）、`setFill`（「表头标黄」）、`setFont`（「表头加粗」）

### 第三档 ❌ 建议留模型（本质是业务语义判断）

- `setConditionalFormat`（「标出异常值」）、`createChart`（「画对比图」）、`createPivotTable`（「按月份透视」）、`splitGroupAggregate`、`addImage`、`addShape`、`CreateChartFromPivotTable`

## 7. 结论

- 「全面的精确命令」= 现有 35 动作 + 补齐 10 个缺口。
- 高频必补 5 个：`RemoveDuplicate`、`AutoFill`、`MoveRow`、`MoveColumn`、`CutPaste`。
- 本清单是重构与扩展的基准；每新增一个动作，先在此登记并确认分级，避免回到「散落关键词、隐式调度」的老路。
