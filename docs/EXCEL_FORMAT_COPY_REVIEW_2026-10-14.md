# Excel 格式副本审阅：组合动作入口修复

用户实机计划是在同一 `B2:E5` 区域先加粗、再水平居中。虽然格式动作已经加入副本分派，
`nativeComparisonPlanIssue` 仍无条件拒绝重叠区域，导致 `App` 隐藏“生成副本并排审阅”，
改显“执行到 Excel”。本次允许格式动作之间按顺序重叠；值/公式/清空内容动作的重叠仍拒绝。

原生模式每个格式动作执行后立即验证，整份计划执行完后保存最终证据，支持后续动作覆盖
前一步字号等属性。最终证据同时包含逐格格式、数字格式、逐行行高和逐列列宽；审阅期间
修改副本尺寸会阻断写回，提交后对照也能识别尺寸不符。变化单元格按地址去重。

原生与 OOXML 兼容副本都保留格式动作的实际填充色，不用绿色标记覆盖审阅结果。
值写入动作仍沿用原有绿色标记。

## 证据边界

- `excelFormatCopyReview.test.ts`：真实审阅与执行模块，有状态 Office.js 替身。
- `components/ExcelFormatReviewFlow.test.tsx`：真实审阅组件、hook、生成、确认写回及撤销，
  固定计划，Office.js 与窗口排列桥为替身；不是整个 App、真实模型或真实账本证据。
  两个用例分别模拟原生窗口 API 和 ExcelApi 1.17 / 无 Desktop 1.1 的兼容模式。
- 兼容用例核对实际生成的 OOXML 样式内容，不使用替身复制出来的工作簿冒充已解析文件。
- 真实 Excel 2024 本次尚未验收。兼容模式只生成可打开的独立文件，不能自动回读或关闭
  该副本；必须由用户核对并勾选确认。不能将它宣称为原生副本自动回读。
- WPS 产品代码未改；此前用户报告的 WPS 实机通过结论保留。Excel 撤销兼容实现由另一
  任务并行维护，本记录不替代其完整恢复边界验收。

## 本机复跑

```powershell
npm exec --workspace @excel-bro/addin -- vitest run src/excelFormatCopyReview.test.ts src/excelNativeComparison.test.ts src/manualReviewFile.test.ts src/components/ExcelNativeReviewActions.test.tsx src/components/ExcelFormatReviewFlow.test.tsx
```

实机使用已保存的独立测试工作簿，刷新加载项后重新生成“B2:E5 加粗、水平居中，不改值和公式”的计划。
应出现“生成副本并排审阅”；点击后原表不变，独立副本显示组合格式。兼容模式核对副本、
勾选确认，再“保存变更到原表”。写回后核对值与公式，最后撤销并核对原有混合字号。
旧计划和旧审阅会话不得沿用。

## 本轮验证与并行任务交接

- 副本相关 5 文件：75 passed（含新组件原生/兼容两条点击链路）。
- WPS 格式流程、组合边界、面板回归：24 passed。未修改 WPS 产品代码。
- 三条反向变异（格式重叠、副本格式核对、填充色保真）均触发断言失败，
  `excelNativeComparison.ts` 逐字节还原；后续全量中的相关文件通过。
  日志：`outputs/codex-excel-copy-mutation-*.log`。
- Excel 全量：1145 passed / 3 failed / 1 skipped；文件 96 passed / 4 failed / 1 skipped。
  完整日志：`outputs/codex-excel-copy-full.log`。这是并行修改工作区的运行结果，不是冻结树验收。
- 三个失败：`excelFormatCompatPath.test.ts` 与 `mixedFormatUndo.test.ts` 的值动作撤销
  保留对齐断言（期望 Left、得到 undefined）；`excelWriteEndToEnd.test.tsx` 的人工改值后
  撤销冲突提示断言。另有 `crossProcessChain.test.tsx` 因缺少生成物
  `config/cross_process_chain.json` 在收集期失败。
- 标准 `build:addin` 退出非零：缺少上述 JSON（TS2307）；另有正在并行维护的
  `officeTestHost.ts:386` 默认样式未补齐新增 `CellStyle` 字段（TS2740）。
  不能将“副本相关测试通过”写成“全量/构建通过”。

按用户确认的文件分工，本轮未修改 `excel.ts`、`officeTestHost.ts` 或撤销测试；这些失败项
交给正在维护恢复路径的任务处理，再合并复跑。没有提交、推送或发布，也没有操作真实工作簿。

---

# 第二轮（同日）：清掉回归与构建阻断，并把比对副本补成交付必需环节

上一节末尾那份 `outputs/codex-excel-copy-full.log` 是**并行工作区的中间态快照**（它里面
已经出现本轮改写后的测试名），不能当作当前结果。本轮在**当前代码**上重跑，逐项定位。

## 一、当前状态复现（不是旧日志）

| 上次失败项 | 当前结果 | 性质 |
| --- | --- | --- |
| `excelFormatCompatPath.test.ts`（值动作撤销对齐期望 Left/undefined） | **11 passed** | 已被撤销双层收窄修复 |
| `mixedFormatUndo.test.ts`（值动作撤销格式保留） | **21 passed** | 同上（替身 `writeCell` 曾丢 `style`/`format`，已保真） |
| `excelWriteEndToEnd.test.tsx`（人工改值后撤销冲突提示） | **25 passed** | 文案已改为**指名字段**，断言同步 |
| `crossProcessChain.test.tsx` 收集失败（缺 `config/cross_process_chain.json`） | 已修（见下） | **环境依赖 + 测试编排缺陷** |
| `officeTestHost.ts:386` TS2740（新 `CellStyle` 字段默认样式未补齐） | 不复现 | 已修 |
| 缺 JSON 连带 TS2307，`build:addin` 失败 | **已修** | 同上 |

`npm run test:addin`：**1160 passed / 3 skipped（1163），文件 100 passed / 2 skipped（102），退出码 0**
（1160 = 修复前的 1157 + 本轮新增 2 条行高点击用例 + 1 条撤销范围回归用例）。
两个 skipped 文件 = `onlineLink.test.tsx` + `crossProcessChain.test.tsx`，都是**环境门控**。
`npx tsc --noEmit`（addin）退出码 0。`npm run build:addin`（`tsc --noEmit` + 两次 `vite build`）退出码 0。

## 二、`crossProcessChain.test.tsx`：为什么是环境 + 编排，而不是产品

该消费侧用例**静态 `import` 一份只在真实生产侧跑完才写出的工件**
（`server/tests/test_cross_process_chain_real.py` → `config/cross_process_chain.json`，
由**真实 ASGI 服务端 + 真实 Docker 隔离容器**产出）。本机 `docker version` 客户端在、
**守护进程未运行**（`npipe:////./pipe/dockerDesktopLinuxEngine` 连不上）⇒ 生产侧
`pytestmark` 整体 skip ⇒ 不产出工件。

修法（**不补假工件、不用旧工件、不改 skip 通过**）：

1. 消费侧去掉静态 `import`，改为**运行时按需读取** + `EB_CROSS_PROCESS=1` 显式门控，
   与 `onlineLink.test.tsx`（`EB_ONLINE=1`）**同一编排**；缺工件却又开了门控 ⇒ 直接抛错
   拒绝对回放。这一步同时消掉默认 `vitest run` 的收集失败与 `tsc` 的 TS2307。
2. `scripts/run_cross_process_chain.mjs` 在**确证工件已产出之后**才带上 `EB_CROSS_PROCESS=1`
   跑消费侧，并把 cwd 收到 `apps/excel-addin`（否则 `@excel-bro/core` 别名解析不到）。

实测：`npm run test:cross-process` → 生产侧 4 条用例 skip（Docker 阻断）→ 脚本
**退出码 2** 并打印「连续链路未完成：生产侧环境阻断…本命令不会用旧工件或 mock 顶替通过」。
**跨进程工件回放本轮未执行（环境阻断），照实记为未完成。**

## 三、把比对副本补成交付必需环节

三条链路（同区域加粗＋水平居中、混合字号统一修改、**行高修改**）在组件点击层要求全覆盖。
前两条已在 `components/ExcelFormatReviewFlow.test.tsx`；**行高这条此前只到模块集成层**
（`excelFormatCopyReview.test.ts` 的「行高证据」用例），本轮补进组件点击链路：

- 新增 `行高动作：副本内按整行生效、确认才写回、撤销把每一行行高逐行还原`
  （`describe.each(["native","manual"])` 展开 → 该文件 **4 用例**）。
- 断言到**行级**语义：`A1:B2` 行高 44 ⇒ 第 1/2 行都是 44、**第 3 行不动**、未声明的
  **列宽保持默认 8.43**；确认前 `host.originalWrites() === []`；撤销**逐行**还原为 20/30
  （不是统一刷一个值）；值始终 `[["甲",100],["乙",200],["丙",300]]`。

逐条对照用户九项要求（现状）：

| 要求 | 落点 |
| --- | --- |
| 确认前原表零写入 | `excelFormatCopyReview.test.ts`、`ExcelFormatReviewFlow.test.tsx` 均断言 |
| 同区域/部分重叠按顺序生效 | `excelFormatCopyReview.test.ts`（同区域加粗+居中）、`wpsFormatCombinationBoundary.test.ts` |
| 比对用**最终状态**、行高按整行 | 「字号依次覆盖按最终状态核对」+「行高证据」+ 新增行高点击用例 |
| 值/公式/未声明格式不变 | `excelNativeComparison.test.ts`（保留公式/字面 `=`）、`excelFormatCompatPath.test.ts` |
| 撤销恢复原始混合格式；用户后续修改保留/拒绝 | `excelUndoScope.test.ts`（含 C13 判别点） |
| 放弃副本不改原表 | `excelNativeComparison.test.ts`（discard never writes original） |
| 原生模式检测副本被改 | `excelFormatCopyReview.test.ts`（副本格式被改 ⇒ 写回前 `副本中的格式已变化`） |
| 兼容模式**不宣称**原生自动比对通过 | 组件原文「加载项不能自动回读或关闭兼容副本…必须勾选确认」+ 测试「never claims automated copy readback」 |
| 不以直接执行原表代替副本验收 | 两模式都先 `createWorkbook`（`createWorkbookCalls()).toHaveLength(1)`）再写回 |

## 四、撤销恢复按实际影响范围（复核，未改）

`undoRestoreScope` + `excelFormatActionFields` 双层收窄、`clear(Formats)` 条件化并把
非本动作属性的当前值并回、冲突检测逐格/逐行且**首个写操作之前**完成 —— 均已落地。
替身 `clear(Formats)` 清 **`style` + `format` + `styleName`**（含数字格式与命名样式），
**没有**简化成"只删局部对象"来掩盖风险。

反向变异复跑：`node scripts/run_excel_format_compat_mutations.mjs` →
**C1–C16 全部 RED（合格）**、五个文件**逐字节还原**、还原后复跑全绿，退出码 0。

## 四（附）组件点击链路抓出的一个真实缺陷（本轮已修）

把「行高」补进组件点击链路、并把该文件跑进全量之后，`ExcelFormatReviewFlow.test.tsx` 的
**同区域组合**用例（`B2:E5` 加粗 + 水平居中）在**当前基线**上稳定变红：
`expected 'Center' to be undefined` —— **撤销后居中残留**。

- **定位**：`undoExecution` 反序恢复时，「`clear(Formats)` 之后要补回的非本动作属性当前值」
  取自**复核阶段一次性读回**、按动作下标缓存的 Map。较早的动作（加粗）处理时拿到的仍是
  **执行后**那份快照，里面还带着 `horizontalAlignment: "Center"` ⇒ 又被合并写回。
  模块级用例都是单动作（只读一次，缓存与现场恰好相同），所以**测不到**。
- **修复**：读取挪进恢复循环，每个动作恢复前重新读**此刻**现场（新函数
  `readCurrentDirectCells`，与 `readUndoGuardState` 同一套原语）；`assertUndoPostState`
  不再回传那份一次性 Map。
- **回归用例**：`src/excelUndoScope.test.ts` 新增 describe ⑤（**10 ✓**，第 3 轮补齐到 **13 ✓**：
  部分重叠不同属性、字号 12→14→16 连续覆盖、恢复中途失败如实抛错）。
- **判别力**：反向变异 **C17**（该读取改成按 `RequestContext` 缓存首次结果）⇒
  该用例变红（`1 failed | 9 passed`）；17 个变异全部 RED、五文件逐字节还原、复跑全绿。
  第 3 轮补 **C18**（区域含数据验证不再拒绝 ⇒ `clear(Formats)` 静默清掉验证）与
  `excelFormatCompatPath.test.ts` 2 条数据验证用例（13 ✓），共 **18 个变异**。

设计细节见 `docs/EXCEL_UNDO_SCOPE_2026-10-14.md` 第 3.2 节。

## 五、本轮验证快照（证据分档）

| 档 | 结果 |
| --- | --- |
| ① 产品修复 | 撤销范围/属性双层收窄、`clear` 条件化 + 当前值并回、冲突指名属性、替身保真（承接上一节）+ **本节新修：`clear` 补回的"当前值"改为恢复那一刻现场** |
| ② 模块集成 | 针对性 9 文件 **144 passed**（`excelUndoScope` 10、`excelFormatCompatPath` 11、`mixedFormatUndo` 21、`excelWriteEndToEnd` 25、`excelFormatCopyReview` 10、`excelNativeComparison` 34、`manualReviewFile` 10、`ExcelNativeReviewActions` 19、`ExcelFormatReviewFlow` 4） |
| ③ 组件点击（全部模拟范围） | `components/ExcelFormatReviewFlow.test.tsx` **4 passed**（native + manual × 加粗居中 / 行高；另有混合字号断言）—— 真实组件 + 真实 hook + 真实生成/写回/撤销，Office.js 与窗口排列桥、固定计划为替身；**不是**整个 App / 真实模型 / 真实账本 |
| ④ 真实服务端 / 模型 | 服务端 pytest 见第六节；**真实模型自主规划未执行** |
| ⑤ 真实 Excel | **未验收**（本机 Excel 2024 未操作）；材料见 `outputs/format-acceptance-2026-10-14/` |
| ⑥ WPS 独立回归 / 实机 | 回归 `npm run test:wps` **796 passed / 3 skipped（54 文件）退出码 0**、`build:wps` 退出码 0、`tsc --noEmit` 0 错；**WPS 格式动作实机未验收** |

其余：`npm run test:addin` **1160 passed / 3 skipped（102 文件）退出码 0**；`npm run test:core` **423 passed（26 文件）**；
addin `tsc --noEmit` 0 错、`build:addin` 退出码 0；反向变异 **17/17 RED**（含 C17）。

## 六、服务端 pytest（本轮未改服务端代码，按约定照实跑）

`.\venv\Scripts\python.exe -m pytest server/tests -v --tb=line -p no:cacheprovider --basetemp=<干净目录>`
→ **2 failed / 1397 passed / 28 skipped，87.82s，退出码 1**。完整日志：`outputs/server-pytest-2026-10-14.txt`。

两类问题，**都不是本轮改动引入的**（本轮未改 `server/`）：

1. **首次用默认 `--basetemp=.tmp-pytest` 跑出 376 errors** —— 逐条定位为
   `PermissionError [WinError 5]: 'C:\Users\jia\AppData\Local\Temp\pytest-of-jia'`，
   即 **pytest 遗留临时目录被占用**（环境阻断）。同文件换干净 basetemp 后 47 passed。
   用干净 basetemp 的全量结果即上面那行。
2. **2 条失败**：`test_python_event_delivery.py::test_python_completion_emits_event_with_result_handle`
   与 `::test_python_event_includes_only_result_handle_not_values`，均
   `assert 'unavailable' == 'completed'`。原因：这两条**直接调真实 `execute_python_analysis_async`**
   （装置未开启隔离执行）却**没有** `test_python_docker_real.py` 那样的
   `pytestmark = skipif(not _docker_available())` 守卫 ⇒ 本机 Docker 不可用时返回 `unavailable`。
   **环境暴露的测试编排缺口，不是产品缺陷**；本记录不擅自把它们改成 skip 了事，
   交回给服务端测试的维护者按既有 `skipif` 约定处理。

另：中间有两次全量跑到 36 分钟被环境挂住（`docker` CLI 探测），已停止；`-v` 落盘那一次 1.5 分钟跑完，
说明是**环境抖动**而非用例挂死（单跑 `test_managed_recovery.py` 5 passed / 6.58s 亦佐证）。

未提交、未推送、未发布；未改业务工作簿；未扩展新格式能力。

