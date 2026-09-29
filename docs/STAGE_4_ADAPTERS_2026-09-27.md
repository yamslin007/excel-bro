# 阶段 4 进度记录（一）：两端 HostAdapter 实现(2026-09-27)

> 阶段 4 的退出条件是参数化列映射在三条通道全绿 + 两宿主实机记录。
> **本记录只覆盖第一步：写出两端适配器实现。** 映射计算、确认绑定、
> 空串负向比较修正、指纹版本化均未开始。

## 1. 改动

| 文件 | 行数 | 内容 |
|---|---|---|
| `apps/wps-addin/src/wpsHostAdapter.ts` | 190 | `createWpsHostAdapter` / `createWpsKeyValueStore` / `resolveWpsCapability` / `createWpsPrimitiveOps` |
| `apps/wps-addin/src/wpsHostAdapter.test.ts` | 118 | 11 条 |
| `apps/excel-addin/src/officeHostAdapter.ts` | 120 | `createOfficeHostAdapter` / `createOfficeKeyValueStore` / `resolveOfficeCapability` / `createOfficePrimitiveOps` |
| `apps/excel-addin/src/officeHostAdapter.test.ts` | 52 | 4 条 |
| `packages/core/src/ports.ts` | −8 | 删掉 `readSheets` 与 `ReadBudget`,理由见 §3.1 |

## 2. 两处诚实的降级,不是偷懒

### 2.1 WPS 白名单外的动作报 `unverified`,不报 `unsupported`

`capabilities.json` 的 `wpsWrite.allowedActions` 只记录**哪些动作已验收**,
**不记录宿主 API 是否具备**。所以白名单外的动作只能报 `unverified`。

报 `unsupported` 等于断言"WPS 没有这个 API",拿不出证据就不该下这个结论。
两者行为相同（方案 §4.2:都不进计划）,区别只在给用户的说明,谎报无收益。
真要区分,前提是逐个动作在真实宿主上探测过 API 存在性 —— 那属于阶段 5
逐动作验收的一部分。

### 2.2 Office 侧全部报 `supported`,但不宣称"Office 无需能力矩阵"

`executeAction` 覆盖 `ExcelAction` 全部成员,且这些动作本就在 Office 上
开发并长期使用。现阶段没有"某动作在某 Office 版本上行为不同"的证据,
硬编码一份未经验收的版本表只会制造假精确。

所以 `resolveOfficeCapability` 目前只对空动作类型报 `unsupported`。
真需要按版本收窄时,在此处按 `Office.context.diagnostics.version` 加判断 ——
位置已经留好,不是以后要重新找地方。

## 3. 三个实现判断

### 3.1 从 `ports.ts` 删掉 `readSheets`

阶段 3 我把它定义成 `readSheets(sheets, budget): Promise<unknown>`。
**返回 `unknown` 的接口成员约束不了任何东西** —— 两端实现返回什么都合法,
类型检查形同虚设。而真正的快照类型 `WorkbookSnapshot` 仍在 app 侧的
`contracts.ts`,没法在 L0 引用。

阶段 4 的参数化列映射只需要读已授权的源区域,`HostPrimitiveOps.readValues`
已覆盖。两端现有的 `captureWorkbook` / `listWpsWorksheets` 暂不经本接口。
等 `contracts.ts` 拆分后再定义这个成员。`ReadBudget` 随之删除（零使用点）。

### 3.2 `PluginStorage.getItem` 的非字符串折成 `null`,**不用 `String()`**

阶段 3 的接口注释写的是"须做字符串强制",实现时改了主意:

```ts
if (typeof raw !== "string" || raw === "") return null;
```

`String(undefined)` 得到 `"undefined"` —— 一个能通过"非空字符串"检查、
却毫无意义的值,会被当成真实存过的数据传给 `JSON.parse` 或直接展示。
折成 `null` 让调用方走"没存过"的既有分支。测试里把 `42` / `{a:1}` /
`undefined` 三种都锁住了。

`removeItem` 在 WPS 侧是**可选成员**（`wpsHost.ts:46`）,缺失时退化为写空串;
`getItem` 对空串也返回 `null`,使"已删除"与"存过空串"在读侧一致。

### 3.3 Office 的 `localStorage` 缺失时,读降级但写必须抛错

读不到走默认值是可接受的降级（隐私模式、部分 Office 容器）。
写入静默失败不行 —— 用户会以为设置保存成功了。

## 4. 回归

| 套件 | 阶段 3 | 本次 |
|---|---|---|
| `npm run test:core` | 4 files,28 passed | 4 files,**28 passed** |
| `npm run test:addin` | 46 files,474 passed | **47 files,478 passed** |
| `npm run test:wps` | 8 files,60 passed | **9 files,71 passed** |
| `npm run test:server` | 457 passed,3 xfailed | **457 passed,3 xfailed** |
| `npm run lint` | 0 errors,11 warnings | **0 errors,11 warnings** |
| `npm run build:addin` | 100 modules | **100 modules** |
| `npm run build:wps` | 64 modules | **64 modules** |

**模块数未变,这是本次最需要说明的一点**:两个适配器目前**没有任何消费者**。
入口文件都没 import 它们,所以不进依赖图,也不进产物。它们现在是
"已实现、已测试、未接线"的状态。接线要等映射计算落地,否则只是把
现有调用链改绕一圈而不产生任何行为收益。

一次未复现的服务端偶发（`test_a7_a9_write_and_undo_round_trip[baseline-D-to-E]`）
本轮未再出现,原因仍未查明,记录见
[阶段 3 端口记录 §5.1](./STAGE_3_PORTS_2026-09-27.md)。

---

# 阶段 4 进度记录（二）：参数化列映射计算层

## 6. 改动

| 文件 | 行数 | 内容 |
|---|---|---|
| `server/app/column_mapping.py` | 185 | `build_column_mapping_actions` / `map_column_values` + 四个错误类 |
| `server/tests/test_column_mapping.py` | 128 | 15 条单元用例 |
| `server/tests/test_column_mapping_baseline.py` | −18 | 移除三处预期失败标记 |

### 6.1 先纠正一条上一轮的错误结论

上一轮我曾报告"模块已创建、三条 XFAIL 已转 XPASS、`exit=0`"。
**这条是错的**：`server/app/column_mapping.py` 当时并不在磁盘上,
`git status` 只显示测试文件被改过,重跑得到的是 `6 passed, 3 xfailed`。

顺带解答那句"`strict=True` 却 `exit=0` 值得核实":本轮模块真正落地后
重跑,得到 `3 failed`,失败原因写的是 `[XPASS(strict)]`。**strict 机制正常**,
上一轮的 `exit=0` 是在能力尚不存在时跑出来的 XFAIL,不是 XPASS。

### 6.2 三处"拒绝而非兜底"

| 情形 | 处理 | 理由 |
|---|---|---|
| 来源值不在映射表内 | 抛 `UnmappedCategoryError` | 猜错的分类在单元格层面看不出来,会静默流入下游统计 |
| 源/目标在同一表上重叠 | 抛 `OverlappingRangeError` | 逐格落地时读写先后决定结果,而这个顺序三条通道不一致,没有可承诺的语义 |
| 形状不符 | 抛 `ShapeMismatchError` | 截断或补齐都会写出用户没要求的格子 |

未命中值**全部收集后一次抛出**,不是遇到第一个就停。映射表是人工维护的,
一次只报一个会逼调用方反复试错。

### 6.3 空白与未匹配是两件事

空白 = 源本来没有值（`None` 或纯空白字符串）→ 走 `blank_label`。
未匹配 = 源有值但映射表没覆盖 → 是调用方映射表不完整,拒绝。

**显式映射项优先于空白策略**:调用方若把 `""` 写进 `mapping`,
说明它要区分"空串"与"真空格",这个意图应当被尊重。

## 7. 本次回归

| 套件 | 上一轮 | 本次 |
|---|---|---|
| `npm run test:core` | 4 files,28 passed | 4 files,28 passed |
| `npm run test:addin` | 47 files,478 passed | 47 files,478 passed |
| `npm run test:wps` | 9 files,71 passed | 9 files,71 passed |
| `npm run test:server` | 457 passed,3 xfailed | **475 passed,0 xfailed** |
| `npm run lint` | 0 errors,11 warnings | 0 errors,11 warnings |
| `npm run build:addin` | 100 modules | 100 modules |
| `npm run build:wps` | 64 modules | 64 modules |

服务端 +18 条 = 3 条原 XFAIL 转正 + 15 条新单元用例。

**又一次需要说明的"没有消费者"**:`column_mapping.py` 目前只被测试 import,
生产代码里没有任何调用点。计划生成路径（`planner.py` / `intent.py`）尚未改为
走这个入口。它现在与两个适配器同处一个状态:已实现、已测试、未接线。

三条通道共用同一实现这件事,**现阶段只由"它是纯函数、无宿主依赖"保证**,
不是由"三端都在调它"证明的。真正的跨宿主贯通仍需接线 + 实机记录。

---

# 阶段 4 进度记录（三）：映射入口接线到规划路径

## 8. 改动

| 文件 | 行数 | 内容 |
|---|---|---|
| `server/app/excel_agent.py` | +215 | `map_column_values` 工具 + `_snapshot_region` / `_mapping_from_pairs` / `_map_column_plan` |
| `server/tests/test_map_column_tool.py` | 210 | 21 条 |
| `server/tests/test_planner.py` | ±6 | 两处工具集断言改为"只读工具已撤下" |

`column_mapping.py` 不再是孤岛 —— 生产代码现在有了调用点。

## 9. 分工：模型给规则，服务端算值

模型只决定**哪一列、按哪张映射表、空白叫什么**；值矩阵由
`build_column_mapping_actions` 在服务端逐格算出,模型不经手。

理由不是"模型算不准",而是**这类错误在单元格层面看不出来**:逐格抄写会漏行,
也会把没列在规则里的来源值自行归类。用户看到的是一列看起来合理的标签。

工具是**终局工具**（返回 `PlanResponse`，直接结束 agent 循环，
`excel_agent.py:1410` 的既有分支），不是"算完再交给模型 submit_plan"——
后者等于把算好的矩阵再经模型转抄一遍，白做。

`PLAN_ONLY_TOOLS` 刻意不含它:那是 `submit_plan` 校验失败后的定向修复通道,
此时只该暴露 `submit_plan`。

## 10. 三处拒绝

### 10.1 快照越界:拒绝,不补 `None`

`_read_range` 对快照外的格子补 `None`。那对只读查询是合适的降级 ——
模型看到空值会继续追问。但映射结果**要落盘**:补出来的 `None` 会被映射成
空白标签,写进实际上没有数据的行,用户看到凭空多出来的一列标签。

行边界按**实际捕获量**（`len(headers) + len(dataRows)`）判定,不是 `rowCount`
—— 后者是宿主报告的总行数,快照截断时它仍然很大。列边界按**声明宽度**
`columnCount` 判定:`dataRows` 的行可能因尾部空值被裁短,此时区域内缺失的项
是真实空格,折成 `None` 是正确的。

### 10.2 映射表重复键:报错,不后者覆盖前者

同一来源值给了两个目标标签,是需求本身有歧义。静默取一个等于替用户做了
没有依据的选择。目标标签相同的重复项则接受（无歧义）。

### 10.3 `blankLabel` 为空:拒绝

空标签会让空白单元格写入空串,而空串与"未写入"在回读侧不可区分
（正是 §1.1 要修的那个问题）。工具描述里明确要求模型没问清就先问用户。

## 11. 一处需要说明的取舍

产出的是 `writeValues`,而系统提示第 112 行写着"不要把算出来的结果塞进
writeValues"。这不是自相矛盾:那条针对的是**聚合查询结果**（源数据一变就过期,
应改用 `computeValues` 在执行时重算）;这里写入的是用户给定的标签表的查表结果,
映射表本身就是固定输入。

但它**确实仍是一张快照**。所以计划的 `warnings` 里明确写出"源区域之后再改动,
目标列不会自动跟着变",不靠"它其实不算结果"这种自我辩解绕过去。

## 12. 两端对等性：这条路径在 WPS 上根本走不通

先纠正本节的初稿。我原本写的是:产出的 `writeValues` 已在
`wpsWrite.allowedActions` 白名单内,因此"不需要新增动作类型即可在两个宿主上执行"。
**这条推理是错的**,它只检查了"动作类型能不能执行",没检查"值是从哪来的"。

`_snapshot_region` 从 `PlanRequest.workbook` 的 `dataRows` 里取源值。而 WPS
送上来的快照,`dataRows` **恒为空**:

```ts
// apps/wps-addin/src/wpsReadonlyAgent.ts:126 —— 全插件唯一的 dataRows 赋值处
worksheets: sheets.map((sheet) => ({ ...sheet, dataRows: [], truncated: sheet.rowCount > 0 }))
```

配 `planningReadPolicy: "none"` —— 这是刻意的隐私/体积取舍,不是缺陷。
用 WPS 形态的 `PlanRequest` 实测,工具直接拒绝:

```
「表1」的快照只到第 1 行，区域 B2:B4 要到第 4 行，超出已捕获范围（该表快照已截断，请缩小区域或分批处理）
```

**§10.1 那个"拒绝而非补 None"的决定在这里救了一次**:如果它像 `_read_range`
一样补 `None`,WPS 上会静默写出一整列"未分类",而用户看到的是一列长得很合理的标签。
拒绝让架构缺口在测试阶段就炸出来,而不是在用户的表里。

结论:映射计算放服务端,对 Excel 与文件夹通道成立（前者快照带值,后者服务端直接读文件）,
对 WPS **永远不成立**。修法见记录（四）。

## 13. 本次回归

| 套件 | 上一轮 | 本次 |
|---|---|---|
| `npm run test:core` | 28 passed | 28 passed |
| `npm run test:addin` | 478 passed | 478 passed |
| `npm run test:wps` | 71 passed | 71 passed |
| `npm run test:server` | 475 passed | **496 passed** |
| `npm run lint` | 0 errors,11 warnings | 0 errors,11 warnings |
| `npm run build:addin` | 100 modules | 100 modules |
| `npm run build:wps` | 64 modules | 64 modules |

### 13.1 改了两条既有断言,说明理由

`test_planner.py` 两处原本写 `_tool_names(...) == {"submit_plan", "submit_answer"}`。
新增终局工具后它们失败。这**不是真实回归**:两条用例检查的是
"`dataResults` 已带入时只读工具被撤下",同文件 `:1211` 的同类断言本来就用
`<=`。改为断言"只读工具集与实际工具集不相交" + "两个终局工具在内",
今后新增终局工具不会再假报失败。

---

# 阶段 4 进度记录（四）：映射计算下沉到 `packages/core`

## 14. 改动

| 文件 | 行数 | 内容 |
|---|---|---|
| `packages/core/src/columnMapping.ts` | 214 | Python 侧的 TS 镜像 |
| `config/column_mapping_cases.json` | 20 条用例 | 跨语言共享夹具 |
| `packages/core/src/columnMapping.test.ts` | 89 | 加载夹具，21 条 |
| `server/tests/test_column_mapping_shared_cases.py` | 96 | 加载同一夹具，21 条 |
| `packages/core/src/index.ts` | +14 | L2 导出 |

## 15. 为什么是两份实现，而不是一份

记录（三）把计算放在服务端。对 WPS 不成立（§12）。但也不能反过来"只放客户端":
文件夹通道由服务端直接读写 xlsx，客户端根本不在场。

所以两份实现是必需的，不是妥协:

| 通道 | 源值在哪 | 谁算 |
|---|---|---|
| Office | 快照带 `dataRows` 上传 | 服务端（已有） |
| 文件夹 | 服务端直接读文件 | 服务端（已有） |
| WPS | 只在客户端，`planningReadPolicy: "none"` | 客户端（本次） |

## 16. 一致性靠夹具，不靠人工对读源码

`config/column_mapping_cases.json` 列 20 条输入 + 期望输出/期望错误，
两侧各自加载、断言同一批值。仓库里既有的 `test_safety_mirror.py` 用的是
**抓源码文本**的办法 —— 那对"清单"（公式黑名单）有效，对"逐格输出的逻辑"无效。

**验过夹具确实会红**，不是摆设:把 `stringifyKey` 里的两行布尔分支删掉，
TS 侧 2 条立刻失败（`数字 0 与 false 不是空白`、`布尔真按 Python str() 形态键化`），
Python 侧照常全绿 —— 正是"单边修改一侧变红"该有的样子。已还原。

### 16.1 夹具锁住的一处真实语言差异

`String(true)` 在 JS 是 `"true"`，Python 的 `str(True)` 是 `"True"`。
映射表的键是用户写的中文标签居多，但源列若是布尔（WPS 读出的复选框列就是），
两端会算出**不同的未命中集合**:一端报"来源值 true 不在映射表内"，另一端正常映射。
`stringifyKey` 按 Python 形态键化，理由是服务端那份是先落地的、且 openpyxl
读出的值也走 `str()`。

## 17. 尚未完成

**core 里这个函数目前没有调用点。** 与记录（一）的两个适配器同一状态:
已实现、已测试、未接线。WPS 侧要真正产出映射计划，还需要:

1. 一条"模型给规则 → 客户端算值"的回传路径（现在模型只能调服务端工具）;
2. 决定服务端工具与客户端计算如何共存，不能两处各写一遍分支判断。

在此之前，§12 描述的缺口**仍然存在** —— WPS 上让模型走
`map_column_values` 依旧会被拒绝。这是诚实的当前状态，不是"已修复"。

## 18. 本次回归

| 套件 | 上一轮 | 本次 |
|---|---|---|
| `npm run test:core` | 28 passed | **49 passed**（5 files） |
| `npm run test:server` | 496 passed | **517 passed** |
| `packages/core` `tsc --noEmit` | 0 | 0 |

---

# 阶段 4 进度记录（五）：WPS 本地计算路径贯通，§12 的缺口补上

## 19. 改动

| 文件 | 行数 | 内容 |
|---|---|---|
| `server/app/column_mapping.py` | +48 | 抽出 `validate_column_mapping_shape`（不需要源值的那部分校验） |
| `server/app/excel_agent.py` | +60 | `_map_column_plan` 按读策略分叉；新增终局响应分支 |
| `server/app/models.py` | ±14 | `ColumnMappingPair` 去掉 `from`/`to` 别名，见 §21 |
| `server/tests/test_map_column_tool.py` | +120 | +11 条 |
| `server/tests/test_no_business_hardcoding.py` | ±8 | 「客户」改否定环视，见 §22 |
| `apps/wps-addin/src/wpsColumnMappingPlan.ts` | 112 | 规则 + 已读工作表 → 计划 |
| `apps/wps-addin/src/wpsColumnMappingPlan.test.ts` | 148 | 12 条 |
| `apps/wps-addin/src/wpsReadonlyAgent.ts` | +60 | `local_plan_rule` 分支 + `readMappingSource` |
| `apps/wps-addin/src/WpsAgentPanel.tsx` | +6 | 本端算出的计划跳过服务端校验，见 §23 |
| `apps/excel-addin/src/contracts.ts` | ±10 | 跟随字段名改动 |

记录（四）§17 写的「core 里这个函数没有调用点」到此结束：`build:wps` 从
64 模块变成 **69 模块**，映射计算真正进了 WPS 产物。

## 20. 分叉条件是读策略，不是宿主

`_map_column_plan` 现在按 `PlanRequest.planningReadPolicy` 分两条：

| 读策略 | 源值在哪 | 服务端返回 | 谁算 |
|---|---|---|---|
| `"ifNeeded"` | 快照带值 / 服务端能读文件 | `PlanResponse`（含 actions） | 服务端（原样不动） |
| `"none"` | 只在客户端 | `LocalPlanRuleResponse`（只有规则） | 客户端 |

**刻意不按宿主分叉**：服务端不认识宿主，也不该认识（方案的既定决策）。
`"none"` 就是「这个客户端自己读值、不上传」的宿主无关信号，任何送 `"none"`
的客户端都走同一条路。Excel 两条通道送的都是 `"ifNeeded"`，行为完全未变
（`test_read_policy_if_needed_still_computes_on_the_server` 锁住）。

服务端**没有值也仍要拦**该拦的：范围解析、形状不符、同表重叠三类不需要值就能
判定，由新抽出的 `validate_column_mapping_shape` 在下发规则前执行。留到客户端
才发现等于把错误推到模型离开循环之后，那时它已无法修正参数。

## 21. 一处自己找出来的契约坑：别名只在一条序列化路径上出现

`ColumnMappingPair` 起初照模型工具参数写成 `Field(alias="from")`
（`from` 在 Python 是关键字，只能靠别名）。**这会坏在序列化上**：

```
bare  : ['source', 'target']     # SSE 里的 model_dump(mode="json")
alias : ['from', 'to']           # response_model（FastAPI 默认 by_alias=True）
```

本仓库两条路径都在用。同一个模型吐出不同键名，客户端的契约校验只在其中一条上
通过 —— 而这种失败只在跑到那条路径时才出现。响应模型因此不跟随工具参数的
命名，统一用无别名的 `source`/`target`，两侧字段名一致。

`test_wire_keys_are_identical_on_both_serialization_paths` 直接断言
`model_dump() == model_dump(by_alias=True)`，把「无别名」钉住。

## 22. 「客户端」不是行业字段

`test_no_business_hardcoding.py` 的 `BUSINESS_TERMS` 含「客户」，而本轮注释里
大量出现「客户端」（与服务端相对的软件术语），撞了。

改成 `"客户(?!端)"` 否定环视，**没有从清单里删掉这一项** —— 删了会放过真正的
客户字段。已验过 `客户名称` 仍判违规、`客户端读取源值` 不判。

## 23. 本轮最实质的一处发现：算好的值差点从校验接口漏出去

接线跑通、测试全绿之后复查 UI 路径，发现 `WpsAgentPanel.tsx:554` 对每份计划都
调 `validatePlan`，而它会把整份计划 POST 到 `/api/plans/validate`。

**那等于把刚才刻意留在本地的值上传了**：`writeValues` 的验收条件里带着完整的
值矩阵（`models.py` 的 `RangeEqualsCriterion.expected`），而单射映射可逆 ——
标签列加映射表足以还原源列。整条「源值不上传」的设计会在最后一步被自己抵消。

此前那条「源值不上传」的用例只检查了**规划请求**，给了假的安全感 —— 它看不到
后续这次 POST。这是本轮最需要记下的一点：验证点选错位置时，绿灯毫无意义。

跳过服务端校验在这条路径上不丢东西，三点都核实过：

1. WPS 的执行验证是**自洽的** —— 写入前后指纹比对 + 回读
   （`wpsPlanExecutor.ts:347` 起），全程不读 `acceptanceCriteria`
   （grep 确认 WPS 侧生产代码零引用）；
2. 授权与存在性由 `prepareWpsPlan` 对着**活宿主**查（版本闸、只读态、保护、
   合并、重叠、单元格上限），比服务端拿名字列表比更严；
3. 区域已在服务端下发规则时规范化过。

判定用 `WpsTurnResult.computedLocally` 显式标记，不靠翻计划内容猜。
**验过这道闸确实会红**：把三元判断改成 `false`（即恢复原来的无条件校验），
`WpsAgentPanel.test.tsx` 立刻 1 条失败，已还原。

## 24. 尚未覆盖的一处不对等

共享夹具锁的是**纯函数**那一层。两端各自的「从已读数据里切出源区域」
（服务端 `_snapshot_region`、客户端 `sliceMappingSource`）**不在夹具覆盖范围内**，
只有各自语言的单元测试。两者的越界策略目前一致（都拒绝、都不补 `None`），
行内缺项也都折成 `null`，但这是**人工对读得出的结论，不是测试保证的**。

真要保证，需要把「捕获原点 + 值矩阵 + 请求区域 → 切片结果」也做成共享夹具。
本轮没做，记在这里。

> **后记（记录六）：上面那句"两者的越界策略目前一致"是错的。**
> 夹具一跑就发现两侧各漏一个方向。详见 §28。这正是"人工对读"这种验证方式
> 本身的问题 —— 不是这次读得不够仔细。

## 25. 本次回归

| 套件 | 上一轮 | 本次 |
|---|---|---|
| `npm run test:core` | 49 passed | 49 passed |
| `npm run test:addin` | 478 passed | 478 passed |
| `npm run test:wps` | 71 passed | **88 passed**（10 files） |
| `npm run test:server` | 517 passed | **528 passed** |
| `npm run lint` | 0 errors,11 warnings | 0 errors,11 warnings |
| `npm run build:addin` | 100 modules | 100 modules |
| `npm run build:wps` | 64 modules | **69 modules** |

## 26. 仍未做的：两宿主实机记录

上面全部是自动化测试。**这条路径一次都没在真实 WPS 上跑过。**
M1-M11 仍未开始，§12 那个「架构缺口」现在有了代码与测试，但
「两个平台呈现相同、安全且可恢复的行为」这条验收标准要实机才能算过。

---

# 阶段 4 进度记录（六）：切片逻辑的跨语言夹具，顺带抓出两个真实缺陷

## 27. 改动

| 文件 | 行数 | 内容 |
|---|---|---|
| `config/mapping_slice_cases.json` | 13 条用例 | 「原点 + 值矩阵 + 请求区域 → 切片结果」共享夹具 |
| `apps/wps-addin/src/wpsMappingSlice.shared.test.ts` | 88 | 加载夹具，14 条 |
| `server/tests/test_mapping_slice_shared_cases.py` | 110 | 加载同一夹具，14 条 |
| `apps/wps-addin/src/wpsColumnMappingPlan.ts` | +30 | `MappingSourceSheet` 加 `columnCount`；补右侧列越界拒绝；行检查拆两支 |
| `server/app/excel_agent.py` | +10 | 补左侧列早于原点的拒绝 |
| `apps/wps-addin/src/wpsReadonlyAgent.ts` | +1 | `readMappingSource` 带出 `columnCount` |
| `apps/wps-addin/src/wpsColumnMappingPlan.test.ts` | ±10 | 跟随字段新增，加一条右侧越界用例 |

## 28. §24 的结论是错的：两侧各漏一个方向，且互为镜像

上一轮我写"两者的越界策略目前一致（都拒绝、都不补 `None`）"。夹具跑起来立刻
证伪。实测（`_snapshot_region`，原点 `B2`、`columnCount: 2`）：

```
A2:A4 -> [[None], [None], [None]]     ← 列在原点左侧，静默补 None
C2:D4 -> ERR 已用区域只到第 C 列       ← 列在右侧溢出，拒绝
```

客户端 `sliceMappingSource` **恰好相反**：拦左侧（`rect.startColumn <
sheet.startColumn`），不拦右侧 —— 它压根没有右边界可比，`MappingSourceSheet`
当时没有 `columnCount`，而按各行实际长度判定会把"尾部空值被裁短"误判成越界。

| 方向 | 服务端（改前） | 客户端（改前） |
|---|---|---|
| 行早于原点 | 拒绝 | 拒绝（与行溢出合在一条） |
| 行超出捕获量 | 拒绝 | 拒绝 |
| 列早于原点 | **静默补 `None`** | 拒绝 |
| 列超出声明宽度 | 拒绝 | **静默补 `null`** |

后果与 §10.1 要防的是同一件事：补出来的空值被映射成空白标签，写进实际上没有
数据的格子。两侧都改为四个方向全拒。客户端新增 `columnCount` 字段（取
`WpsSheetReadResult.columnCount`，即宿主报告的声明宽度），与服务端按
`columnCount` 判列、按实际捕获量判行的口径对齐。

**这不是"上一轮读得不够仔细"能解释的**。§16 已经论证过夹具优于对读源码，
但只把它用在了纯函数层；切片层留了个"人工对读结论"，一轮之后就证明这个结论
是错的。验证手段选错，绿灯与结论都不可信 —— 与 §23 那条"验证点选错位置时，
绿灯毫无意义"是同一条教训的第二次出现。

## 29. 夹具的两处设计取舍

**错误只按种类共享，文案两侧各自声明。** 服务端说「快照/已捕获」，客户端说
「已读入」—— 同一件事在两侧措辞本就不同（一个是上传的快照，一个是本地读入的
内存），强行统一没有收益。夹具只给 `error: "columnBeyondCaptured"` 这类种类名，
各侧自己维护"种类 → 消息片段"表，沿用 `column_mapping_cases.json` 的既有形态。

**`rowCount` 刻意填成远大于实际捕获量的值。** Python 侧装载夹具时
`rowCount = startRow + len(values) + 9999`，把 §10.1 那条"行边界按实际捕获量
判定而不是 `rowCount`"一并锁住：若哪天改回按 `rowCount` 判，越界用例会全部
变成通过。

**`MAX_READ_ROWS` / `MAX_READ_COLUMNS` 不进夹具。** 那是服务端的读预算，
约束的是模型单次工具调用的请求规模；WPS 侧数据已在内存里，没有对应的请求
需要约束。这是**合理的单侧差异**，不是缺口。

## 30. 验过两侧都会红

- 客户端：把右侧列越界判断改成 `false && ...`（即恢复改前行为），TS 侧
  `列超出声明宽度：拒绝，不补 null` 立刻失败，Python 侧照常全绿。
- 服务端：把 `min_column < origin_column` 改成 `False and ...`，Python 侧
  `列早于捕获原点：拒绝，不补 null` 立刻失败，TS 侧照常全绿。

均已还原。这是"单边修改一侧变红"该有的样子。

## 31. 本次回归

| 套件 | 上一轮 | 本次 |
|---|---|---|
| `npm run test:core` | 49 passed | 49 passed |
| `npm run test:addin` | 478 passed | 478 passed |
| `npm run test:wps` | 88 passed（10 files） | **103 passed（11 files）** |
| `npm run test:server` | 528 passed | **542 passed** |
| `npm run lint` | 0 errors,11 warnings | 0 errors,11 warnings |
| `npm run build:wps` | 69 modules | 69 modules |
| `apps/wps-addin` `tsc --noEmit` | 0 | 0 |

WPS +15 = 14 条夹具用例 + 1 条右侧越界单元用例；服务端 +14 = 同批夹具用例。

---

## 5. 阶段 4 剩余工作

| 项 | 状态 |
|---|---|
| 两端 `HostAdapter` 实现 | 记录（一）完成 |
| 适配器接线到入口 | 未开始 |
| 参数化列映射的本地确定性计算 | **记录（二）（四）完成，两侧均未接线** |
| 映射入口接入计划生成路径（服务端） | **记录（三）完成** |
| 映射入口接入 WPS 客户端路径 | **记录（五）完成（仅自动化测试，§26）** |
| 切片逻辑的跨语言夹具 | **记录（六）完成，顺带修掉两个镜像缺陷（§28）** |
| §3.5 确认绑定（A11/A12） | 未开始 |
| §1.1 空串负向比较修正 | 未开始 |
| §6.3 指纹版本化（v2 前缀 + 双读） | 未开始 |
| 移除阶段 G 预期失败标记 | **记录（二）完成** |
| 两宿主实机记录 M1-M11 | 未开始（**此前记录的"本机无宿主环境"有误**,已核实 Office16 与 Kingsoft 均在本机） |

## 32. 后续更新：WPS 开发写入通道与空串验收（2026-09-27）

本节更新上表的两项旧状态，不改写前面各轮的历史记录：

- 实机 `Application.Version="12.0"` 与安装目录产品版本 `12.1.0.28488` 不同。
  原 `verifiedVersionPrefixes: ["12.1"]` 没有对应实机证据，已清空；仅开发调试
  加载项通过 `developmentVersionPrefixes: ["12.0"]` 临时进入预检、计划预览和确认执行。
  生产构建仍拒绝写入，未把 `12.0` 宣称为正式验收版本。
- `wpsWriteVersionGate.ts` 成为能力探测、适配器和执行器共用的版本规则；执行时
  再核对宿主版本没有从预览时变化。只读、授权、保护、合并、区域、动作白名单、
  单元格上限、快照与撤销检查均保留。
- §1.1 空串负向比较已修：`writeValues` 保留 `""`，独立回读直接区分 `""` 与
  `null`。宿主折叠后会报错并留下 `partial_write_recoverable` 记录，不再静默给出
  `verified`；旧撤销指纹仍按旧算法折叠，§6.3 的 v2 双读尚未开始。
- 自动化：WPS 108 passed，Excel addin 478 passed，服务端 542 passed；两端构建
  与 Python 编译通过。实机 M2-M10 仍待在独立测试工作簿执行，不能据单测宣布验收。

## 33. 后续更新：WPS 确认绑定与撤销指纹 v2（2026-09-27）

本节继续更新 §5 旧待办，不把阶段 4 宣布完成：

- `prepareWpsPlan` 捕获每个目标区域的写前 v2 指纹、最终计划与动作快照、授权表集合、
  工作簿身份及预览修订。确认执行前逐项复核；预览后改目标或改计划，均在建立恢复点
  与写入前拒绝，要求重新预览。任务窗格还检查当前选择的工作表集合没有变化。
- 本地列映射结果在内存中携带源区域值快照；计划准备时核对计算时源值，确认执行
  时核对预览时源区域（含公式）的指纹。源数据仍不发送到 `/api/plans/validate`。
- 新 WPS 恢复点写 `version: 2` 与 `v2:` 类型化指纹，公式优先，`null` 与 `""`
  明确区分；固定输入/输出向量锁住算法。保留原 `PluginStorage` 键与 v1 算法双读。
  v1 恢复点在无空值歧义时仍可撤销；涉及空字符串或空单元格时标为冲突，保留
  记录并拒绝自动覆盖。
- 新增负向测试先确认旧实现会错误通过，再实现修复。WPS 定向 22 条测试和构建已通过。
- 尚未正式接线的 WPS `HostAdapter` 原语曾只按工作表名复核；同名但路径变化的
  工作簿仍能读到值。新增负向测试先复现，再改为共用 `assertWpsAuthorization`，
  让原语每次读取或写入前均核对工作簿与已授权工作表身份。

仍未闭合：WPS 可编辑完整预览及编辑后重新确认（A5/A11 的 UI 部分）、Office 与
文件夹通道的 A11/A12 对照、适配器正式接线、两宿主实机 M1-M11。单测通过不等于
目标 WPS 版本可安全正式开放写入。

## 34. 后续更新：WPS 可编辑预览（2026-09-27）

§33 列出的 WPS UI 缺口继续收窄：计划卡现显示各动作前 5 行、前 5 列的待写值；
用户可像 Excel 计划编辑器一样修改完整 JSON。编辑态隐藏确认写入；应用修改时
先跑白名单结构校验，服务端计划再跑服务端校验，本地映射计划为避免源值外传
继续只做本地校验；两者都重新执行 WPS 宿主授权、区域、矩阵、源/目标快照预检，
产生新的预览修订。UI 回归覆盖「编辑后未写入、重新预览后才可确认、最终写入修订值」
和「本地映射编辑不调用服务端校验」。

当前只展示前 5×5 单元格，完整矩阵通过编辑器可查；任务窗格真实宽度的可用性
还要实机验证。阶段 4 仍未过闸：Office/文件夹 A11/A12 对照、适配器接线、
两宿主实机 M1-M11 均未完成。

本轮最终回归：core 49、Excel addin 478、WPS 115、server 542 全部通过；
两端构建、Python 编译与 `git diff --check` 通过；lint 0 error、11 条既有 warning。
pytest 仍有本机 `.pytest_cache` 写入权限 warning，不影响 542 条测试结果。
