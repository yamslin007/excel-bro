# WPS 格式能力探测 · 实机证据（2026-10-13）

> 这是**真实 WPS 宿主**跑出的报告，不是自动化替身结果。
> 原始 JSON：`outputs/wps-format-probe-20261013/user-reported-051653.json`。

## 运行环境

| 项 | 值 |
|---|---|
| 宿主 | WPS ET（`Application.Version` 报告 `12.0`） |
| 工作簿 | `C:\Users\jia\Desktop\tst\格式探测.xlsx` |
| 工作表 | `Sheet1` |
| 探测目标 | `A1`（单个空白单元格） |
| 时间 | 2026-10-13T05:16:53.181Z |

## 原始结论（保留，不篡改）

- **`baselineRestored`: `false`** —— 整轮基线未恢复通过。
- **`aborted`: `true`** —— 因 `ColumnWidth` 恢复失败而提前停止后续写入探测。
- 汇总：共 18 项，`restoreVerified` 8、`failed` 2、`skipped` 8（其中垂直对齐同值跳过 1 项、
  填充/边框因中止跳过 5 项、AutoFit 未执行 2 项）。

## 通过项（8 项，`restore_verified`）

这些项在**本机这个 WPS 版本**上做到了「写哨兵 → 回读一致 → 恢复 → 回读回到原值」：

| 成员 | 哨兵 | 原值 | 映射到的产品动作属性 |
|---|---|---|---|
| Font.Bold | `true` | `false` | `setFont.bold` |
| Font.Italic | `true` | `false` | `setFont.italic` |
| Font.Name | `"黑体"` | `"宋体"` | `setFont.name` |
| Font.Size | `14` | `11` | `setFont.size` |
| Font.Color | `255` | `0` | `setFont.color` |
| HorizontalAlignment | `-4108`（xlCenter） | `1`（xlGeneral） | `setAlignment.horizontal` |
| WrapText | `true` | `false` | `setAlignment.wrapText` |
| RowHeight | `20` | `13.5` | `resizeRange.rowHeight` |

> **注意**：这只是「该属性可读、可写、回读一致、能恢复」，**不等于**对应产品动作已闭环。
> 动作闭环还需要：预览、确认、执行、回读、撤销的端到端实现与验收。

## 失败项（2 项）

### 1. `Font.Underline`（`failed: readback`）

- 原值 `-4142`（`xlNone`），哨兵 `true`，写后回读 `2` ≠ `true`。
- **根因：哨兵类型错误。** `Range.Font.Underline` 在 WPS ET 里是 `XlUnderlineStyle`
  枚举（`xlNone=-4142`、`xlSingle=2`、`xlDouble=-4119` 等），**不是布尔**。
  写 `true` 后宿主把它当作枚举，回读到 `2`（`xlSingle`）。
- **待办**：改用合法枚举哨兵（如 `xlSingle=2` 或 `xlDouble=-4119`），重新验证，
  **不得**用宽松类型比较（`==`、`Boolean()` 归一）掩盖问题。

### 2. `ColumnWidth`（`failed: readback`，且恢复漂移）

- 写 `12` 回读 `12.008889198303223`；原值 `8.44444465637207`，恢复写回后回读
  `8.453332901000977`（≠ 原值，差了约 `0.0089`）。
- **漂移原因尚未查明**：列宽在 WPS 里可能是「字符宽」单位、内部按像素量化存储，
  「量化」目前只能写作**待验证假设**，尚未实测确认单位、量化与重复设置是否产生漂移。
- **待办**：先查明单位、量化与重复设置是否产生漂移。**不得**直接加任意 epsilon 放行；
  若采用宿主精度等价判据，必须提供文档/实测依据、容差边界与多轮无累计漂移证据；
  不能恢复则保持 `ColumnWidth` 动作受限（保持 `resizeRange.columnWidth` 为 `unverified`）。

## 未测项（5 项，因 `aborted` 被跳过）

因 `ColumnWidth` 恢复失败触发「同轮恢复失败即停止」守卫，以下成员**未实际执行**，
**不是失败、也不是不支持**：

- `Interior.Color`、`Interior.Pattern`（填充组）
- `Borders.Item.LineStyle`、`Borders.Item.Weight`、`Borders.Item.ColorIndex`（边框组）

> 这些组需要在**新的独立测试基线**上单独重跑（能力组选择入口已支持），不能把
> 「未测」当成「已验证」，也不能当成「不支持」。

## 哨兵同值跳过（1 项）

- `VerticalAlignment`：原值与哨兵均为 `-4108`（xlCenter），写不出区别，`skipped`。
  需从合法 `xlVAlign` 枚举（如 `top=-4160` / `bottom=-4107`）换哨兵后重新验证。

## 未验证项（2 项，AutoFit）

- `EntireRow.AutoFit`、`EntireColumn.AutoFit`：成员可调用，但**未实际执行**
  （避免改动整行/整列尺寸），`skipped`，**不计入通过**。

## 结论口径（如实，不合并宣称）

- **已验证**：上表 8 项格式属性（本机 WPS 12.0、单格 A1 场景）。
- **已接入产品闭环**（Task #53，仅限已验证属性子集）：
  - `setFont`：`bold` / `italic` / `name` / `size` / `color`（**不含 underline**）。
  - `setAlignment`：`horizontal` / `wrapText`（**不含 vertical**）。
  - `resizeRange`：仅 `rowHeight`（**不含 columnWidth**）。
  - 完整链路：能力声明 → 白名单计划 → 预览 → 确认 → 执行 → 逐格/逐行回读 → 撤销与恢复核对。
  - 逐格/逐行快照（`wpsFormatSnapshot.ts`）保证混合格式恢复，不靠单格探测冒充。
- **未验证（明确拒绝，不静默忽略）**：Underline 枚举哨兵（待修）、ColumnWidth 量化漂移（待查）、
  VerticalAlignment（哨兵同值待换）、Interior/Borders（未测）、AutoFit（未执行）。
- **产品闭环**：setFont/setAlignment/resizeRange 的**已验证属性子集**已接入；其余格式动作
  （setFill/setNumberFormat/setBorders/autofit）仍未接通，保持 fail-closed。
- **Excel 端**：保留独立实机验收任务，与 WPS 端分别报告进度。
