# 撤销按"动作实际影响"拆分恢复（2026-10-14）

> 起因（用户本轮原文，先修撤销、不扩大功能）：
> 1. 按动作实际影响拆分恢复：只改行高/列宽，不清除单元格格式；只改值/公式，不无条件重写全部格式。
> 2. 不能声称 `DIRECT_FORMAT_LOAD_OPTIONS` 已覆盖全部格式。若仍采用 `clear(Formats)`，
>    必须证明所有被清除属性和格式来源均能恢复；对无法完整保存的场景明确限制，不用默认值补齐，
>    **不把显示一致等同于来源语义恢复**。
> 3. 撤销冲突检测必须覆盖实际恢复范围，并采用逐格/逐行状态。用户后续修改要么保留，
>    要么在任何恢复写入前明确报冲突，不能静默覆盖。
> 4. 补四组失败测试（删除线/上下标/图案色/斜线边框保留、后续手改不被覆盖、
>    含条件格式只改行高、非 Normal 样式与直接格式覆盖的恢复）
>    —— **测试替身不能预设 `clear(Formats)` 只删除某个局部对象就代表真实宿主语义**。
>
> 本文记录设计、证据分档与仍未完成的边界。**上一轮的兼容路径设计见
> `docs/EXCEL_FORMAT_COMPAT_2026-10-13.md`，本轮是对它的收窄与纠偏。**

## 1. 病根：旧撤销是"整块重写"，两个后果都是静默的

旧实现里，撤销的执行侧是"该区域该恢复什么就全恢复一遍"：

- 值/公式动作也**逐格采格式**，撤销时按格式矩阵写回 ⇒ 用户在动作之后对**别的属性**做的编辑被覆盖；
- 格式动作**无条件** `clear(Formats)` 后再写回全量属性 ⇒ 除了慢和危险，还会把
  "执行前就有的其它直接格式"和"用户后来的编辑"一起清掉（如果采集清单不全就直接丢）。

两者都属于"显示看起来没错、来源语义已经变了"的那类静默损伤。

## 2. 设计：范围（scope）+ 属性子集（formatFields）**双层收窄**

### 2.1 第一层：范围 —— `undoRestoreScope(action)`

`apps/excel-addin/src/excel.ts` 新增纯函数 `undoRestoreScope`，把"这个动作到底动了哪一类东西"
写死成数据：

| 动作 | formulas | numberFormat | cellFormat | rowHeight | columnWidth |
|---|---|---|---|---|---|
| `writeValues` / `writeFormulas` / `writeTable` / `clearRange` | ✅ | — | — | — | — |
| `setNumberFormat` | — | ✅ | — | — | — |
| `setFill` / `setFont` / `setBorders` / `setAlignment` | — | ✅ | ✅ | — | — |
| `resizeRange` | — | — | — | 仅当给了 `rowHeight` | 仅当给了 `columnWidth` |

两点必须成对记住：

- `cellFormat ⇒ 连带 numberFormat` 是**机制约束**：兼容路径靠 `clear(Formats)` 让"未直接设置"
  的属性回到继承态，而 `clear(Formats)` **连直接数字格式一起清**。不一起恢复就等于把用户的数字格式弄丢。
- `resizeRange` 只碰行高/列宽 ⇒ 它的快照里**没有逐格格式矩阵**（`format.directCells === null`），
  撤销时**一个格式写入都不做**，也就不会因为"区域里有条件格式"被误拒。

采集侧同步门控（`captureRangeFormat`）：`needCellFormat = scope.cellFormat`；
`hasFormatScope = scope.cellFormat || scope.rowHeight || scope.columnWidth` 才采。

### 2.2 第二层：属性子集 —— `formatFields`

`excelFormatActionFields(action)` 给出本动作**实际要改**的属性（如 `setFont{bold}` → `["fontBold"]`），
它同时驱动三处，且三处**读同一份派生**：

1. **采集**记录 `inherited`（哪些属性在执行前是"未直接设置"）；
2. **恢复**只写回这些属性（`directKeysOfFields` / `displayedKeysOfFields`）；
3. **冲突摘要**只比这些属性（`scopeFieldDigests` / `cellFormatFieldDigests`）。

于是：用户改了**别的属性** ⇒ 摘要不变 ⇒ 撤销照常执行且**保留**该编辑；
用户改了**会被写回的属性** ⇒ 摘要变化 ⇒ 在**任何写入之前**抛错并指名字段。

范围与属性两层是**正交**的：范围决定"碰不碰这一类"，属性子集决定"这一类里写哪几个字段"。

## 3. `clear(Formats)` 从"无条件"改成"条件化 + 当前值并回"

兼容路径（宿主 ExcelApi ≤ 1.17）恢复"格式的**来源语义**"（不是显示效果）。难点在于：
Excel **没有**"只取消某一个属性的直接设置"的 API，要让某个属性回到继承态只能整格 `clear`。

因此：

- 只有当本动作**确实把某个"执行前未直接设置"的属性设上过**时才 `clear(Formats)`
  （判据 `saved.directCells.some(cell => cell.inherited.length > 0)`）；
  其余情况**一律不 clear**，按属性写回即可。
- 一旦要 clear，就必须把"**非本动作属性**的**当前**取值"先读回来再并回去
  （`observedDirectCells`，与冲突摘要读**同一份**中间表示）——否则 clear 会带走
  执行前就有的其它直接格式**以及用户在本动作之后的编辑**。
- 本动作管辖、但执行前**未直接设置**的键 ⇒ 显式 `delete`（清空后保持"未直接设置"）。

这条就是"不把显示一致等同于来源语义恢复"落在 clear 路径上的实现：
`clear` + 并回当前值 + 剔除本动作未设置过的键 = 只把本动作改过的那几个属性退回来源语义，
其余一切（含命名样式、数字格式）保持原样。

### 3.1 采集清单 = 宿主可读**全集**，并且如实声明硬边界

`DIRECT_FORMAT_LOAD_OPTIONS` 请求的是"`clear(Formats)` 会重置、而本机 API 读得回来"的全量：
8 条边框（含 `diagonalDown`/`diagonalUp`）、填充 5 项（含 `patternColor`/两个 `tintAndShade`）、
字体 10 项（含 `strikethrough`/`subscript`/`superscript`/`tintAndShade`）、
`shrinkToFit`/`textOrientation`/`readingOrder`/`autoIndent`/`protection`、命名样式 `style`。

**仍然不是"全部格式"**，且这一点写进了类型注释与测试：单元格**内部富文本**（`textRuns`，
ExcelApi 1.18）在 ≤1.17 宿主上**读不出来、也写不回去** ⇒ 本机既无法判断某格是否含富文本，
也无法在富文本被改变时还原。

关于 `clear(Formats)`：官方 Excel 语义说它**不清除**单元格内字符级格式（"Clear Formats 不改
单元格内单个字符的格式"）。**这只是那一步操作的语义依据，不是"含富文本区域可安全执行并撤销"
的整体结论** —— 完整的写入与撤销链路**未经真实宿主验证**；且一旦宿主实际行为与官方语义不符，
本机既发现不了也修复不了。故本节只声明依据，**不给出整体安全结论**。这是宿主 API 的硬边界
—— **不补默认值、不声称已覆盖**。

`style`（单元格命名样式）已确认**可读可写**（`CellPropertiesLoadOptions.style`，ExcelApi 1.9，
与 `format` 同级）⇒ 命名样式可以捕获并写回，`clear` 路径的恢复是**可自证**的。

### 3.1.1 数据验证：`clear(Formats)` 会清除它，写入前 fail-closed（同日第 3 轮补缺口）

官方确认 `clear(Formats)` 会清除区域上的**数据验证**，而数据验证不是格式动作的管辖范围、
本机也读不回逐格验证来源（只能读到 `range.dataValidation.type` 这个 readonly 枚举）。
因此兼容路径下**目标区域含数据验证（`type !== "None"`）时，在写入前抛
`FormatRangeUnrecoverableError` 拒绝**，与条件格式 / 合并单元格同口径 —— 而不是采集/恢复数据验证
（那属于"扩新能力"，超出本轮范围）。判据读 `Excel.DataValidation.type`（ExcelApi 1.8，≤1.17 可读），
`"None"` = 无验证放行；其余（`WholeNumber`/`List`/`MixedCriteria`…）= 拒绝。

### 3.2 「当前值」必须是**恢复那一刻**的现场（同日第 2 轮修复的真实缺陷）

`clear(Formats)` 之后要把"非本动作属性"的当前值补回去，否则会连带清掉别的属性。
这个"当前"原来是在**复核阶段一次性读回**、按动作下标存进 Map 再在恢复时取用的。
撤销是**反序**执行的，于是同区域多个格式动作会踩到这个：

> 计划 = `B2:E5 加粗` + `B2:E5 水平居中`。反序撤销：先恢复**居中**（把 `horizontalAlignment`
> 从直接格式里删掉，回到继承态）；再恢复**加粗**时，它取到的"当前值"是**执行后那次**
> 的快照 —— 那份里 `horizontalAlignment` 还是 `Center` —— 于是又被合并写回。
> **结果：撤销"看起来执行了"（加粗确实掉了），但居中残留。**

修法：把读取挪进恢复循环，**每个动作恢复前重新读一次此刻的现场**
（`readCurrentDirectCells`，与 `readUndoGuardState` 同一套原语 `DIRECT_CELL_PROPERTIES_LOAD` +
`readDirectCellFormat`，表示不会漂移）；`assertUndoPostState` 不再回传那份一次性 Map。

这个缺陷是**组件点击链路**（`components/ExcelFormatReviewFlow.test.tsx` 的同区域组合用例）
先抓出来的 —— 模块级的单动作用例都测不到它（单动作只读一次，缓存与现场恰好相同）。

判别点：`src/excelUndoScope.test.ts` 新增 describe ⑤「先加粗后居中，撤销后两个属性都回到
执行前」；反向变异 **C17**（把 `readCurrentDirectCells` 改成按 `RequestContext` 缓存首次读取）
让它变红。

## 4. 冲突检测：覆盖实际恢复范围，逐格/逐行，写入前报错

判据按强度由强到弱三级退回（**不是为了放行，而是为了"升级不把旧恢复点判成被改过"**）：

1. `scopeDigest`（第三代，新记录主判据）：载荷**恰好**等于 `scope ∩ formatFields`，逐格/逐行；
   字段名集合本身也进摘要（范围收窄/放宽都不是"现场没变"）；
2. `extendedFormatDigest`（第二代，旧记录）：固定的全属性集合（第三代的超集）；
3. `postFingerprint`（第一代，最旧）：只看 值/公式/数字格式/填充色/加粗/字色。

冲突文案指名到**具体属性**（`cellFormatFieldDigests` + `FORMAT_FIELD_LABELS`），例如
`「数据」A1 在本次撤销将要写回的字段（单元格格式（加粗））上又发生了变化，已停止撤销以免覆盖后续编辑。`
—— "单元格格式对不上"对用户没有可操作性，"（加粗）"才有。

所有检查都在**首个写操作之前**完成；不满足条件时整次撤销保持**只读失败**，原表零写入。

## 5. 测试替身保真（用户第 4 条点名的红线）

旧替身不保真会制造**假绿**，本轮逐条对齐真实宿主语义：

| 真实语义 | 替身修正 |
|---|---|
| 写值**不改格式** | `writeCell` 保留 `format`/`style`/`styleName`（空值只清内容不删格式） |
| `clear(Formats)` 清**整格全部格式 + 数字格式 + 命名样式** | 一并 `delete style`/`format`，`styleName = undefined`，并记一次 `kind:"format"` 写 |
| `getCellProperties` 只返回**被请求**的字段 | 逐字段按 `loadOptions` 门控（`style` 单独按 `loadOptions.style` 返回） |
| 边框 `getItem` 同时接受数值与 `"EdgeTop"` 等字符串 | `borderSideKey` 双向映射（修掉"所有边都写进 top"的真缺陷） |
| `CellStyle` 需覆盖 `CellPropertiesFormatLoadOptions` 可读全集 | 补齐删除线/上下标/图案色/tint/收缩适应/文字方向/阅读顺序/自动缩进/保护 |
| 显示值 ≠ 直接设置 | 新增 `cellDisplayedStyle()`（硬默认 ← 样式背景 ← 直接设置），与 `cellStyle()`（只含直接设置）**分开** |
| 条件格式规则是工作表状态 | `conditionalFormatCount` 做成表状态 + 读写口（常量永远等于 1，什么都证明不了） |

## 6. 自动化证据快照（2026-10-14 01:1x）

| 证据 | 结果 |
|---|---|
| `apps/excel-addin` `npx tsc --noEmit` | 仅剩既有交付物缺失 `src/crossProcessChain.test.tsx(39,22) TS2307`（`config/cross_process_chain.json`） |
| Excel 格式相关套件（6 个 + 新增 1 个） | **139 passed / 8 files**（含新增 `src/excelUndoScope.test.ts` **9 passed**） |
| Excel 全量 `vitest run` | **1157 passed / 1 skipped**，1 个文件收集期失败 = 上述缺失工件（与改动前一致） |
| WPS 全量 `vitest run`（回归） | **796 passed / 3 skipped**，0 失败 |
| `packages/core` 全量 | **423 passed / 26 files** |
| 反向变异 `scripts/run_excel_format_compat_mutations.mjs` | **16 个变异全部让目标用例变红**；5 个受监视文件**逐字节还原**（sha256）；复跑全绿 |
| `vite build`（Excel 端 taskpane + functions） | `✓ built`（376.90 kB taskpane；functions 2.06 kB） |
| `npm run build:wps`（WPS 端） | `✓ built`（589.87 kB，仅有既有的 >500 kB 分块提示） |

> 注意：`npm run build:addin` 的完整脚本是 `tsc --noEmit && vite build && vite build -c vite.functions.config.ts`，
> 会在第一步就因**既有缺失工件** `config/cross_process_chain.json` 中止；上表的两条 `vite build`
> 是**绕过那一步 tsc** 单独跑的，用来证明"打包本身没坏"。该缺失工件不是本轮引入的。

新增套件 `src/excelUndoScope.test.ts`（ExcelApi ≤ 1.17 兼容宿主）九条，逐一对应本轮要求：

- ① 原有删除线 + 上标 + 图案色 + 斜线边框，加粗后撤销**逐项保留**（另有一条证明旁格不受牵连）；
- ② 加粗后手改 `textOrientation`/`shrinkToFit`/`locked`/`formulaHidden` ⇒ 撤销**保留**；
  加粗后手改**另一个受管属性**（字色）⇒ 撤销**保留用户的新字色**（这条才是"属性子集"的判别点）；
  反例锚点：手改**加粗本身** ⇒ **写入前**报冲突、原文指名"加粗"、原表零写入、用户改动保留；
- ③ 含条件格式（1 条规则）+ 直接字体色的区域**只改行高** ⇒ 撤销**零格式写入**、
  `saved.format.directCells === null`、字体色与规则数都不变（另有一条只改列宽）；
- ④ 非 Normal 命名样式（`Good`/`Neutral`）+ 直接填充覆盖 ⇒ 撤销后样式名与填充都还原、加粗消失。

反向变异清单（`scripts/mutate_excel_format_compat.mjs`，共 16 个）：C1–C4、C8–C12 沿用上一轮；
C5/C6 改为**采集/恢复门控**（本轮机制）；C7 改为"clear 后不并回当前值"；
新增 C13（恢复不按属性子集收窄）、C14（冲突文案不指名属性）、C15（采集清单漏删除线）、
C16（命名样式不写回）、**C17（补回的当前值用一次性快照 ⇒ 同区域多动作撤销残留）**。
同日第 2 轮复跑：**C1–C17 全部 RED**，五个文件逐字节还原，复跑全绿。

同日第 3 轮（数据验证 fail-closed + 组合撤销补齐）：新增 **C18**（区域含数据验证不再拒绝 ⇒
`clear(Formats)` 会静默清掉验证）；`excelUndoScope.test.ts` 补 3 条组合撤销用例（部分重叠不同属性、
字号 12→14→16 连续覆盖、恢复中途失败如实抛错），13 passed；`excelFormatCompatPath.test.ts` 补
2 条数据验证用例（含验证拒绝零写入 / 无验证放行），13 passed。变异脚本 `MUTATIONS.length` 动态输出，
`UNDO_SCOPE_SPEC.minTests` 10→13、`COMPAT_SPEC.minTests` 11→13。

## 7. 实机验收材料（本轮指向）

- 夹具与步骤：`outputs/format-acceptance-2026-10-14/`（`格式验收.xlsx` +
  `验收步骤与核对表.md`）；生成/复位脚本 `scripts/make_format_acceptance_fixture.py`。
- 第二区 G1:L5 专为本轮四类场景构造（G2 删除线/上标/图案色/斜线边框、G3 后续手改、
  G4:L4 条件格式、G5 命名样式 + 直接填充覆盖）。
- 环境快照（写文档时实测）：Excel 运行中且**加载项刚被真实拉取**
  （`2026-10-13T17:16:05Z` / `_host_Info=Excel$Win32$16.01$zh-CN`）、3000 与 8765 LISTENING；
  WPS `et.exe` 运行中但 **3889 未监听**（做 WPS 侧前需先起 WPS 前端服务）。
- 上一轮那份 `outputs/format-acceptance-2026-10-13/格式验收.xlsx` 仍被宿主占用（写入
  `PermissionError`）⇒ 本轮**另起目录**，旧文件不动；**不为了写文件去结束用户正在用的宿主进程**。

## 8. 仍未完成（不得合并宣称）

- **Excel 2024 实机核对未执行**：夹具与步骤已备好（第 7 节），需用户在实际宿主机上按表跑完并贴回。
- ~~`npm run build:addin` 的 `tsc --noEmit` 步骤被既有缺失工件 `config/cross_process_chain.json` 挡住~~
  **同日第 2 轮已修**：消费侧 `crossProcessChain.test.tsx` 去掉静态 `import`，改运行时按需读取 +
  `EB_CROSS_PROCESS=1` 门控（与 `onlineLink.test.tsx` 同一编排）⇒ `tsc --noEmit` 0 错、
  `build:addin` 退出码 0。**但跨进程工件回放本身仍未执行**（Docker 不可用 ⇒ `npm run test:cross-process`
  退出码 2，且明确拒绝用旧工件/替身顶替）。
- 「显示值路径」（ExcelApi ≥ 1.19）本轮同样按 `formatFields` 收窄，但它恢复的是**显示值**
  （受本机 API 表达力限制，无法区分"显示一致"与"来源一致"）。这条**口径差异是刻意的、已声明的**，
  不等于"来源语义已恢复" —— 该路径的实机核对同样未执行。
- WPS 侧格式动作的**实机**闭环仍未验收（自动化 = 替身证据）；本轮只确认 WPS 自动化未回归。
