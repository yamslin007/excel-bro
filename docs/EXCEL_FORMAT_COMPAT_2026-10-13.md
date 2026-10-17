# Excel 2024（批量许可永久版）格式动作兼容路径（2026-10-13）

> 起因：Excel 实机失败根因已核实 —— 当前宿主为 **ProPlus2024Volume / 16.0.17932.20076**。
> `apps/excel-addin/src/excel.ts` 的 `captureRangeFormat` 强制依赖 **ExcelApi 1.19 /
> `Range.getDisplayedCellProperties`**，而该宿主不支持，于是 `setFont` 在**写之前**就被拒，
> 格式动作在这台机器上完全不可用。
>
> 本文记录根因、修复设计、证据分档，以及**恢复 Excel 实机验收所需的重启与重新生成计划步骤**。

## 1. 根因：不是"该升级 Office"，而是宿主结构性拿不到 1.19

按微软 Excel JavaScript API **要求集**表，**volume-licensed perpetual Office on Windows（LTSC）**
**不支持 ExcelApi 1.18 / 1.19 / 1.20**，最高只到 **1.17**（Office 2024 = Version 2302）。
用户的 `ProPlus2024Volume 16.0.17932` 正属此类。

因此 `getDisplayedCellProperties` 在这台宿主上是**结构性不可用**，不是配置问题。
把"升级 Office"当作唯一修复方案是错的：正确方向是补一条 **≤ 1.17 的兼容路径**。

## 2. 两条逐格格式路径（由宿主能力决定）

`apps/excel-addin/src/excelFormatCapability.ts` 把"这台宿主到底能做什么"抽成**纯函数**
（`resolveExcelFormatCapability`），判定用要求集而不是"方法是否存在"：

| 路径 | 依赖 | 逐格读取 | 写进快照 |
| --- | --- | --- | --- |
| `displayed` | ExcelApi **1.19** | `getDisplayedCellProperties`（**实际显示**的格式：直接设置 + 表格样式 + 命名样式 + 条件格式的结果） | `UndoRangeFormat.cells` |
| `direct`（1.9–1.17） | ExcelApi **1.9** | `getCellProperties`（**只返回直接设置在单元格上**的属性，缺席即"未直接设置"） | `UndoRangeFormat.directCells` + `inheritedFields` |
| `unavailable` | 连 1.9 都没有 | —— | 格式动作**执行前检查**即被拒（值/公式动作不受影响） |

### 2.1 兼容路径的恢复语义（本轮核心设计）

`direct` 路径的撤销**两步走**（`restoreRangeFormat`）：

1. `range.clear(Excel.ClearApplyTo.formats)` —— 清掉目标区域的**全部直接格式**，
   所有属性退回"继承样式 / 条件格式 / 默认"；
2. `setCellProperties` —— 写回捕获到的**直接格式集合**。

结果：执行前**直接设置过**的属性逐格精确还原，执行前**未直接设置**的属性回到继承态。
**没有把显示效果写成显式值** —— 恢复的是"格式的来源语义"，而不是"看起来一样"。

代价与配套：
- `clear(Formats)` 会重置区域内**全部**直接格式（含数字格式），所以采集必须把会被重置的
  属性**一网打尽**（`DIRECT_FORMAT_LOAD_OPTIONS` 额外含 `shrinkToFit` / `textOrientation` /
  `readingOrder` / `protection`）；
- 撤销顺序必须改为"**先恢复格式、后写 `range.numberFormat`**"，否则数字格式会被 `clear` 清没。

### 2.2 无法可靠恢复的场景：**明确拒绝**，不填默认值冒充

兼容路径下，目标区域含**条件格式**（`conditionalFormats.getCount().value > 0`）或
**合并单元格**（`getMergedAreasOrNullObject()` 非空）时，显示状态无法在执行前确定，
且条件格式/合并区域格式不逐格成立 ⇒ 抛 `FormatRangeUnrecoverableError`，
在**写入前**抛出并**指名**缺哪个 API / 哪类区域。

判据只在动作**确实会改逐格属性**时才查（`needsCellCapture`）：纯 `resizeRange`（行高/列宽）
不该因为"区域里有条件格式"被误拒。

"未采集 ≠ 默认值"在两条路径上都成立：`cells` 的每个字段都来自宿主返回值，
`directCells` 只装宿主**确认直接设置过**的字段；**绝不补 `defaultCellStyle()` 之类的默认值**。

### 2.3 能力检测前移

- **预检**（`preflightPlanActions`）拦 `path === "unavailable"`：在执行前检查阶段就点名
  "缺 ExcelApi 1.19 / 1.9，影响第 N 步"；
- **预览面板**（`ExcelNativeReviewActions.tsx`）用 `excelFormatCapabilityNotice` 给出**非阻断**提示：
  走兼容路径 / 被拒绝 + 缺失 API + 受影响步骤。

两条都不等用户确认后才报一句笼统错误。

### 2.4 执行与副本审阅同源

`export async function applyFormatAction(context, range, action)` 是六个格式动作**唯一**的写入实现：
执行器（`applyAction` 的格式分支）与 `stageExcelNativeComparison` 共用它。
否则"副本试算的样子"与"写回原表的样子"会分叉。

副本审阅的格式证据（≤ 1.17 可用）：
- `assertCopyFormatApplied` 在副本上回读**逐格直接格式**（不用显示值），不符即
  "第 N 步 X 在副本中的回读结果与计划不符（…），未开放写回原表。"；
- `copyFormatEvidence`（`formatEvidenceDigest`）让"副本里的格式被改过"也能在写回前被发现；
- `compareExcelNativeCommittedResult` 现在把**格式**一并比对。

**手动兼容副本**（`manualReviewFile.ts` 的 `projectFormatAction`）是 ≤1.17 宿主上实际走的
副本路径：格式动作按与实际写入逐字对齐的语义投影进 OOXML，并按**实际修改属性**标记变化的格。

## 3. 本轮发现并修复的**真实缺陷**

为省一次 `getCellProperties`，曾把"只有格式动作才逐格采集"当优化（`modifiedFields.length > 0`）。
后果两条，都被全量回归逮住：

1. **显示值路径直接崩溃**：`cellProperties` 变 `null`，但 `if (displayedAvailable)` 分支仍
   `cellProperties!.value` ⇒ `Cannot read properties of null`。命中 `excelNativeComparison.test.ts`
   两条真值用例。
2. **兼容路径有损恢复**：值/公式动作不再采 `directCells` ⇒ 撤销退回**区域级单值**分支，
   把区域里**没被这次动作改过**的混合格式拍平。

修复：逐格采集对**所有可恢复动作**都做（不是格式动作的特权），并补两条回归守卫
（`mixedFormatUndo.test.ts` + `excelFormatCompatPath.test.ts`）；变异 **C5 / C6** 专门证明
拆掉它必须变红。

## 4. 证据分档（本轮）

| 分档 | 状态 | 证据 |
| --- | --- | --- |
| 兼容路径闭环（**ExcelApi 1.19 = false** 宿主，真实 `executePlan` / `undoExecution`） | **已通过** | `src/excelFormatCompatPath.test.ts` **11 ✓**（加粗 / 水平居中 / 混合字号 / 行高 / 条件格式拒绝 / 合并区域拒绝 / 继承≠默认值 / 值·公式·数字格式不受影响 / 值动作逐格捕获） |
| 能力判定纯函数 | **已通过** | `src/excelFormatCapability.test.ts` **20 ✓**（三路径 / 受影响属性 / 前移提示 / 期望值 / 取值口径） |
| 格式动作副本审阅（真实 `stageExcelNativeComparison` + 有状态宿主） | **已通过** | `src/excelFormatCopyReview.test.ts` **5 ✓**（副本真实试算、原表零写入、提交后格式对照 different→matched、副本格式漂移拦截） |
| 兼容副本 OOXML 投影 | **已通过** | `src/manualReviewFile.test.ts` **10 ✓**（加粗/居中/填充/数字格式/行高列宽投影、幂等不标记） |
| 预览阶段能力提示（真实面板组件） | **已通过** | `src/components/ExcelNativeReviewActions.test.tsx` **19 ✓**（≤1.17 提示含缺失 API 与受影响步骤且非阻断；1.19 无提示；缺 1.9 警示） |
| 混合格式按粒度恢复 | **已通过** | `src/mixedFormatUndo.test.ts` **21 ✓**（含本轮新增的值动作回归守卫） |
| Excel 端全量回归 | **已通过** | `npx vitest run` → **1141 ✓ / 1 skip**，0 个真实测试失败；唯一 FAIL 是 `crossProcessChain.test.tsx` 的**收集期**错误（`config/cross_process_chain.json` 交付物缺失，需 Docker 重建） |
| 类型检查 / 打包 | **部分** | `vite build` 退出码 0；`tsc --noEmit` 仅剩上述**同一交付物缺失**的 1 条 TS2307，无本轮引入的类型错 |
| 服务端协议一致性 | **已通过** | `server/tests/test_protocol_consistency.py` **61 ✓**（`contracts.ts` 扩展与 `models.py` 一致） |
| 反向变异 | **已通过** | `node scripts/run_excel_format_compat_mutations.mjs` → **12 个变异各自变红**，5 个文件逐字节还原，复跑全绿 |
| **真实 Excel 2024 实机** | **未完成** | 见第 6 节步骤；本轮**不宣称实机通过** |
| 真实模型自主规划 / 工件回放 / 双端在线 | **未完成 / 环境阻断** | 本机 Docker 不可用，`npm run test:cross-process` / `test:online` 拒绝用替身顶替通过 |

## 5. 复跑命令

```bash
# 兼容路径闭环（ExcelApi 1.19 = false 宿主）
cd apps/excel-addin && npx vitest run src/excelFormatCompatPath.test.ts
# 能力判定纯函数 / 副本审阅 / 兼容副本投影 / 面板提示
cd apps/excel-addin && npx vitest run src/excelFormatCapability.test.ts src/excelFormatCopyReview.test.ts \
  src/manualReviewFile.test.ts src/mixedFormatUndo.test.ts src/components/ExcelNativeReviewActions.test.tsx
# 反向变异（12 变异 + 逐字节还原 + 复跑全绿）
node scripts/run_excel_format_compat_mutations.mjs
# 服务端协议一致性
.venv/Scripts/python.exe -m pytest server/tests/test_protocol_consistency.py -q --basetemp=.tmp-pytest/btm
```

## 6. 恢复真实 Excel 验收：重启与重新生成计划

本轮改动全部落在 **Excel 前端代码**（`excel.ts` / `excelFormatCapability.ts` / `contracts.ts` /
`excelNativeComparison.ts` / `manualReviewFile.ts` / `ExcelNativeReviewActions.tsx`），
**未改** `config/capabilities.json` 与 `manifest.xml`，也**未改服务端**。
因此：加载项与后端都要加载新代码，且**必须重新生成计划**（旧预览的格式快照是旧代码采集的）。

1. **完全退出 Excel**（不是关文档）：宿主会缓存加载项资源，只关任务窗格不生效。
2. 重新侧载加载项：`npm run stop:excel` → `npm run start:excel`。
3. 确保任务窗格拿到**新**的 bundle：`npm run dev:addin`（用 dev server 时），
   或 `npm run build:addin`（manifest 指向构建产物时）。
   - 实机确认"宿主到底拉没拉新代码"的唯一证据是 `apps/excel-addin/request-log.txt`。
4. 重启后端（本轮验收用普通后端）：`npm run dev:server`（`http://127.0.0.1:8765`）。
   - **不要**同时起 SDK 后端；SDK 后端只能走 `scripts/start_claude_sdk_server.ps1`。
   - 若用的是受控能力表（`scripts/start_manual_python_acceptance.ps1` 逐字复制 `capabilities.json`），
     本轮未改该文件，无需重跑；改了才需要。
5. 在 Excel 里重新打开**已保存**的 `.xlsx` 测试工作簿（未保存的工作簿没有 `url:` 身份，
   工具桥与副本审阅都拿不到）。
6. **重新发一句请求**，让模型重新生成计划；再走
   「预览 → 生成副本并排审阅 → 确认写回原表 → 回读 → 撤销」。旧对话里的预览按钮不适用。

### 验收关注点（与自动化口径对齐）

- 预览阶段应看到**非阻断**提示：`缺失 ExcelApi 1.19（Range.getDisplayedCellProperties）；受影响动作：第 N 步…`
  （这是"能力检测前移"的预期表现，不是错误）。
- 计划里放 **`setFont` 加粗**、**`setAlignment` 水平居中**、**混合字号的 `setFont{size}`**、
  **`resizeRange` 行高** 各一步；
- 确认前原表**零写入**；
- 写回后：逐格加粗/居中落位；**各格字号各自变成动作值**；行高按**整行**生效；
  区域内的值、公式与数字格式**一字未动**；
- 撤销后：**直接设置过的**属性逐格回到原值；**原本未直接设置**的属性仍是"未直接设置"
  （不是被写成某个显式默认值）；行高按行还原；
- 若目标区域含**条件格式**或**合并单元格**：界面应**在写入前**明确拒绝并说明原因，
  原表零写入 —— 这是**预期行为**，不是故障。

## 7. 红线（本轮未越过）

- **没有删除任何恢复守卫来"放行"**：不可靠恢复的场景一律继续 fail-closed。
- 不填默认值冒充采集成功；`cells` / `directCells` 只装宿主真的返回过的字段。
- 不把"升级 Office"当唯一修复方案，也不要求用户换宿主。
- 不把显示效果写回后声称"原始格式语义已恢复"。
- 本轮**不宣称实机通过**，未提交、未推送、未发布。
