# /function 支持文件夹多工作簿（跨文件公式）— 实施任务书

> 交接说明：本文档是完整实施方案，含背景定论、既有代码事实（文件:行号）、分步改动清单与验收方法。按方案执行即可，方案外的扩展不要做。

## 一、背景与定论

/function 是"当前工作簿"短链：`sendMessage` 命中 `/function ` 前缀即短路（apps/excel-addin/src/App.tsx:2543），绕开整个 planner 链路，且**不做任何 sourceMode 检查**。它从头到尾只读当前 Excel 打开的工作簿：

- 阶段一 `handleFunctionCommand`（App.tsx:2373）用 `captureSelectionContext()`（excel.ts:478）取活动表，`suggestWriteTarget`（App.tsx:355）建议写入格；
- 阶段二 `confirmFunctionTarget`（App.tsx:2427）用 `captureWorkbook([sheet])`（excel.ts:555）抓 headers/columns/sampleRows，`loadDictionaryForFormula`（App.tsx:371）只扫**本簿**名字含"字典/映射"的表；
- 协议层 `GenerateFormulaRequest`（server/app/rule_generator.py:26；前端镜像 contracts.ts:9）只有 description/activeCell/headers/columns/sampleRows/dictionary/modelId，没有多工作簿字段。

后果：用户在 folder 模式勾选 A.xlsx sheet1 + B.xlsx sheet2 后用 /function，模型从未收到 B 的任何信息，生成的公式必然"没考虑 B 表"。**这是设计边界缺失，不是代码 bug，不是模型问题。**

## 二、目标

把用户**已勾选的子表**注入公式生成上下文，让模型生成含真外部引用 `[B.xlsx]Sheet2!A1` 的公式。

关键前提：勾选态天然存在，无需假设——`useScopeSelection.ts` 的 `folderCatalog`（files 含 id/name/relativePath/worksheets）+ `folderSheetKeys`（元素为 `fileId\0sheetName`，utils.ts:91）+ `folderSelections()`（useScopeSelection.ts:123）。后端已有按勾选读表结构的能力：`create_folder_snapshot`（server/app/folder_workbooks.py:277，openpyxl 读 headers/dataRows）。

写入目标**保持现状**：当前打开的工作簿。这样保留 Excel JS 真机试算能力（`previewFormulaFirstCell`，excel.ts:499），不破坏"先预览 + 执行后强验收"边界。

## 三、实施步骤

### 1. 协议扩展：GenerateFormulaRequest 加 extraSheets

前后端同步加可选字段（默认空，向后兼容）：

```
extraSheets: Array<{
  sourceFile: string      // 文件名，如 "B.xlsx"（用于外部引用语法）
  sourcePath: string      // 相对路径（展示用，区分同名文件）
  sheetName: string
  headers: string[]
  columns: string[]       // 列字母，与 headers 一一对应
  sampleRows: string[][]  // 前 5 行
}>
```

- 前端：apps/excel-addin/src/contracts.ts（GenerateFormulaRequest，约 :9）
- 后端：server/app/rule_generator.py:26

### 2. 前端：folder 勾选 → extraSheets

改 `confirmFunctionTarget`（App.tsx:2427）：

- 现有本簿逻辑（headers/columns/sampleRows/dictionary）不变，作为主表上下文；
- 新增：当 `sourceMode === "folder"` 且 `folderSheetKeys.length > 0`：
  - 用 `folderSelections()` + `folderCatalog` 解析每张勾选表的 `{fileName（relativePath 末段）, relativePath, sheetName}`；
  - 优先复用 `workbook` state 里已有的 folder snapshot（`confirmSheetSelection`，App.tsx:1455 存的，sheet 带 sourceFileId/sourceFile），直接取各表 headers/dataRows 前 5 行，零新增请求；无 snapshot 时调 `createFolderSnapshot`（api.ts:412）现取；
  - 列字母用既有 `columnLetterFromIndex` 推导；
  - 组装 extraSheets 传入 `generateFormula`（api.ts:453）。
- 字典表扫描逻辑不变；勾选表天然覆盖"映射表在另一个文件"的场景。

### 3. 后端：提示词注入 + 跨簿引用规则

`_build_formula_generation_prompt`（rule_generator.py:45）：

- 主表段落后，每个 extraSheet 追加一段：
  ```
  外部工作簿「B.xlsx」工作表「Sheet2」（可在公式中用 [B.xlsx]Sheet2! 引用）：
  列头→列字母映射：A列=…、B列=…
  样本行（前几行数据）：…
  ```
- system prompt（rule_generator.py:51-67）增加规则：
  - 引用外部表必须用完整语法 `[文件名]表名!区域`，文件名严格使用注入的名字，禁止编造未提供的文件/表；
  - 区域用绝对引用（`$A$2:$B$7` 风格），与现有字典表规则（rule_generator.py:61）一致。

### 4. 黑名单白名单化（重点，三处同步 + 测试）

现状：`[file.xlsx]Sheet!A1` 语法被三处硬禁——前端 `formulaSafety.ts:28`（EXTERNAL_WORKBOOK_REF_PATTERN）、后端 `safety.py:41`、`models.py:194` writeFormulas 校验器。既有测试 formulaSafety.test.ts:72 断言 `'[report.xlsx]Sheet1'!A1 → "EXTERNAL_REF"`。

改为"仅放行会话内勾选文件"：

- `server/app/safety.py`：`dangerous_formula`（:54）增加可选参数 `allowed_external: set[str] | None`；命中外部引用 pattern 时，方括号内文件名 ∈ allowed_external 则放行，否则仍返回 "EXTERNAL_REF"。UNC 等其他规则一律不动。
- `server/app/rule_generator.py:229-238`：生成结果过黑名单时，把 `request.extraSheets` 的 sourceFile 集合作为 allowed_external 传入。
- `server/app/models.py:194`：writeFormulas 校验器维持现状即可（本期 folder execute 链路不改，见"范围外"）。
- 前端 `apps/excel-addin/src/formulaSafety.ts`：镜像同一白名单逻辑，`assertSafeFormula`（:70）增加可选白名单参数。两处调用点传入勾选文件名集合：试算（excel.ts:512）与 writeFormulas（excel.ts:2850 附近）。
- 测试：formulaSafety.test.ts 保留无白名单拒绝用例，新增白名单放行用例；server/tests 的 safety 测试同步。

### 5. 试算与预览的降级展示（验收机制不动）

- 试算零改动：B.xlsx 已在 Excel 打开 → 真算出值；未打开 → #REF!，**不阻断**，原样展示。
- 预览卡（MessageItem.tsx:454-560）新增小块：解析公式中的 `[X.xlsx]` 引用，列出「本公式引用外部工作簿：B.xlsx › Sheet2」；当试算结果为 #REF!/#NAME? 且公式含外部引用时，显示提示「外部工作簿未在 Excel 中打开，打开后自动重算」。
- 验收保持 formulasR1C1Equal 公式串比对，不新增计算值检查（与项目既有边界一致）。

## 四、范围外（后续单独立项，本次不做）

- 公式经 `/api/folders/execute` 写回 folder 源文件（openpyxl 只写不算，验收降级需单独设计）；
- 路径化外部引用（B 未打开时 Excel 需完整路径解析）；本期以"双开"为前提，预览卡文案说明；
- 输出结果文件作为写入目标。

## 五、改动文件清单

| 文件 | 改动 |
|---|---|
| apps/excel-addin/src/contracts.ts | GenerateFormulaRequest 加 extraSheets |
| apps/excel-addin/src/App.tsx | confirmFunctionTarget 组装 extraSheets（~:2427）；白名单透传 |
| apps/excel-addin/src/formulaSafety.ts | 白名单化 + assertSafeFormula 可选参数 |
| apps/excel-addin/src/excel.ts | 试算/写入两处 assertSafeFormula 调用带白名单（:512、:2850 附近） |
| apps/excel-addin/src/MessageItem.tsx | 预览卡外部引用提示块（~:454-560） |
| apps/excel-addin/src/formulaSafety.test.ts | 白名单用例 |
| server/app/rule_generator.py | 请求模型 + 提示词注入 + system 规则 + 黑名单放行（:26/:45/:229） |
| server/app/safety.py | dangerous_formula 加 allowed_external（:41/:54） |
| server/tests/（safety 相关） | 白名单用例 |

## 六、验收

1. 前端单测：formulaSafety.test.ts 全绿（含新白名单用例）。
2. 后端单测：`cd server && python -m pytest tests -k "safety or rule_generator" --basetemp=<项目内临时目录>`。**必须加 --basetemp 指向项目内目录**，否则 tmp_path fixture 在这台 Windows 上报 WinError 5。
3. 端到端手测：
   - folder 模式勾选 A.xlsx sheet1 + B.xlsx sheet2 → 输入 `/function 把B表的单价按编号匹配过来`；
   - 确认阶段二请求体 extraSheets 含 B 的列头/样本行；
   - 生成公式含 `[B.xlsx]Sheet2!` 引用；
   - B 在 Excel 打开时首格试算有真实值；B 未打开时预览卡显示外部引用提示且不阻断；
   - 确认写入后验收通过（公式串比对），打开 B 后单元格自动算出值。
4. 回归：workbook 模式 /function 行为与之前完全一致（extraSheets 为空的路径）。
