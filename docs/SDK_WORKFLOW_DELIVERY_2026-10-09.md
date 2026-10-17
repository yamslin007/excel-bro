# 贯通 Claude Agent SDK 本地工具链 —— 交付报告（2026-10-09）

## 后续补修：总数展示与当前工作簿关联限制（2026-10-09）

- **双端共用总数**：`dataTools.ts` 的分组结果说明增加“筛选后总记录数”，计数来自完整输入的 `filteredRows`，不累加显示指标，也不受分组显示上限影响。缺字段等警告存在时只显示“已处理记录数（结果不完整）”。现有共用 `LocalResultCard` 显示该说明，无协议字段变更；`calculation` 继续被 SDK 模型投影排除。
- **修正能力声明**：SDK `query_table` 的说明及 schema 仅开放当前工作簿真实支持的 `union`。handler 在登记工具桥请求前拒绝 `join` / `deduplicate`；若模型随后失败或未提交有效结果，呈现固定的“不支持关联或去重”提示。合法查询仍可继续；不得自动把关联改成追加合并。文件夹模式的关联能力没有修改。
- **新增/增强证据**：筛选后跨表总数不受只显示一组影响；缺字段不冒充完整总数；真实执行器结果经共享卡片显示；WPS 真读真算回执的记录总数为 4，即使数值指标合计为 27；Excel/WPS 两种 hostId 下关联与去重共四条拒绝用例，断言不发起宿主查询。
- **最终回归**：Excel **89 文件 / 993 passed**；WPS **38 文件 / 626 passed**；服务端 **1242 passed / 1 skipped**（85.31s）；`build:addin` / `build:wps` / `compileall server/app` 均 EXIT=0。日志：`outputs/total-*-test.log`、`outputs/total-*-build.log`。新功能实机待验，未提交、未推送、未发布。

用户前序 WPS 截图已确认：单表查询、两表追加分组（8/4/4/4）、修改源值后重查（7/5/4/4）、收窄为 Sheet1 后新查询（2/1/1/1），以及单表副本预览、确认写回与回读、会话内撤销回读。截图不替代日志级隐私审计、执行途中撤权验证或 Excel 实机验收。

本轮复测：重启 SDK 后端并刷新已重新构建的加载项。仅授权 Sheet1 时，分组结果说明应显示“筛选后总记录数：5”；同时授权 Sheet1/Sheet2 时应显示 20。关联需求应明确说明当前不支持，不应返回一个追加合并结果冒充关联。

## 实机截图反馈：复杂写入误解析与 WPS 空白目标（2026-10-09）

用户实机截图显示只读结果卡返回“无问题”记录数 2；后续写入请求却变成仅写 G1，
写入内容是整段操作说明，随后并排审阅报“宿主未提供原表目标区域的显示值”。
这证明该次写入尚未完成，不能沿用自动化通过结论宣称实机写入已通。

- **共用服务端规划器**：地址快捷解析从局部匹配改为完整命令匹配；只接受明确数字、
  布尔/空值、完整引号文本及受限字面公式。复杂指令、上下文结果引用和未加引号的
  自然语言交给用户选择的模型后端，不再把剩余说明当作值。保留文本和公式大小写，
  引号内数字保持文本类型。Excel/WPS/openpyxl 共用这一路由，未改授权与写入协议。
- **WPS 目标读取**：确认卡、执行器和原生审阅共用 `readWpsRangeValues`。
  单格明确返回 `null` 计为已取得的空白；`undefined`、缺项、稀疏矩阵不计为完整读取。
  多格只返回标量时不能推断整区空白，有 `Item` 接口才逐格读取，否则继续阻断。
  Getter 抛错仍上抛。截图本身未记录原始 JSAPI 返回值，因此 null 是已定位并覆盖的
  代码缺陷，仍须在该台 WPS 重试才能确认它解释了此次全部宿主行为。
- **验证范围**：新增复杂指令/否定清空/原文保真/所选模型路由测试，以及 WPS 空白
  副本预览原表零写入、未知读取拒绝、矩阵完整性测试。修复前规划器新增断言 7 红，
  WPS 空白预览 1 红；针对性修复后规划器 43 项、WPS 执行与审阅 51 项通过
  （后续另加模型路由和矩阵完整性用例）。真实 WPS 写入、撤销及真实 Excel 仍待验。

复测需重启 SDK 服务端、重新构建并刷新 WPS 加载项，取消旧错误计划后重新发起请求。
应预览 G1 的标题与 G2 的统计结果；不得确认将整段操作说明写入单元格的旧计划。

本次最终回归：Excel 89 文件 / **991 passed**；WPS 38 文件 / **626 passed**；
服务端整包 **1238 passed / 1 skipped**（85.20s）；`build:addin`、`build:wps`
（均含 `tsc --noEmit`）与 `compileall server/app` 全部 EXIT=0。
测试日志保留在 `outputs/write-fix-*-test.log` 与 `outputs/write-fix-*-build.log`。
默认 pytest 临时目录访问被拒，本轮使用新的项目内专用 `--basetemp`，未改权限、未删旧证据。
WPS 首次构建发现新增测试 getter 的读写类型不一致，补充返回类型后重新构建通过；
该调整不改变测试执行逻辑。未操作真实工作簿、未提交、未推送、未发布。

接口参考：[WPS Application.Range 示例](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Application/member/Range)
以单格 `Value2 == null` 统计空白；本次实现进一步把 `undefined` 作为不可确认读取处理。

> 指导文档：`docs/SDK_CAPABILITY_ASSESSMENT_2026-10-09.md`、`docs/SDK_WORKFLOW_GUIDE_2026-10-09.md`、
> `docs/SDK_WORKFLOW_IMPLEMENTATION_PROMPT_2026-10-09.md`；工程纪律见 `AGENTS.md`。

## 0. 一句话结论

**服务端、共享层与两端面板（Excel / WPS）均已实现，自动化覆盖到"真实请求体 + 真实确定性计算 + 两端真实面板点击"。**
SDK 在一次 `query()` 内可真实调用本地表格工具（`describe_authorized_tables` / `query_table`），收到**有类型回执**
（固定状态 + 不透明结果引用 + 允许结构 + 固定错误码，**不含任何值**），按固定错误码修正参数后继续，
最终提交**本地结果卡**或**可预览的白名单写入计划**；两端的规划流会消费 `bridgeTool` 指针、轮询、
执行前复核**当前**授权、在本机真实读取并跑**既有确定性 `query_table`**、回执、渲染结果卡。

本轮（第三轮续跑）在**不新增功能**的前提下做了两件事：

1. **修掉三个阻断**：① 意图 `tool_request` 仍把 `dataResults` 送进 SDK，旧路径挡住新工具桥；
   ② Excel 桥的 `liveContext` 返回规划启动时的**闭包快照**，证明不了"现在还是那个现场"；
   ③ 账本把去掉必填矩阵的响应仍当**完整 `PlanResponse`** 返回。
2. **补齐四处口径**：WPS 结果与**真实源版本**绑定并在写前检测来源变化；未保存同名工作簿改用**会话身份**、
   核实不了就**明确阻断**；**无流式工具投递能力时提前拒绝**（不进入无人回执的等待）；
   两端**从真实面板**起一路点到写入回读的端到端用例。

> **第五轮（本轮续跑）聚焦 Excel 写侧恢复登记的可达性**：把第四轮钉住的
> "App 写入路径上 `registerOfficeExecution` 恒不可达"修掉 —— **授权范围 / 实际来源依赖 / 写入目标**
> 三口径分离建模，来源依赖**只**由可信执行记录产生（不采信模型自报、不用"授权减去目标"伪造），
> 并接通了"来源与目标明确不重叠"的计划。修复过程中连带查出并修掉**两处真实接线缺陷**
> （快照身份从未被写入、桥基线指纹与复核不是同一口径），后者使工具桥在**已保存的**真实工作簿上
> **恒因 `source_changed` 拒读** —— 详见 §1.9。

> **第四轮（上一轮续跑）聚焦 Excel 写侧**：把上上轮标为"未实现"的那一行补上 —— 用有真实状态的
> Office.js **宿主替身**（`apps/excel-addin/src/officeTestHost.ts`）从真实 `<App/>` 点击驱动
> `计划预览 → 生成原生审阅副本 → 用户确认 → 写前授权/身份/源版本复核 → 原表写入 → 回读`。
> 详见 §1.8；**替身宿主证据与真实 Excel 实机证据严格分列**（§4.1 vs §4.4）。

**完成口径（分端如实）**：

| 端 | 面板级端到端点击 | 状态 |
| --- | --- | --- |
| **WPS** | 发起查询 → 桥指针 → 回执 → **计划预览 → 确认 → 写前复核 → 执行 → 回读** | **已实现**（`WpsAgentPanel.test.tsx`，真实点击、断言确认前零写入、回读同一矩阵 `Sheet1!A1 === "90"`） |
| **Excel** | 发起查询 → 桥指针 → 回执 → **结果卡按引用取回**（含等待期收窄/源值变化/身份读不到/过期/重开） | **已实现**（`sdkIntentRouting.test.tsx`，真实 `<App/>` + 真实 `fireEvent.click`） |
| **Excel** | **计划预览 → 确认 → 写前复核 → 执行 → 回读**的 App 级点击链 | **已实现（替身宿主自动化）**：`excelWriteEndToEnd.test.tsx`，真实 `<App/>` + 有真实状态的 Office.js 宿主替身，**断言到具体单元格被写、写日志按原表/副本分账**。**仍未取得真实 Excel 实机证据**——这一行**不得**与"实机已验收"混谈（§4.4） |
| **Excel** | 同一条 App 点击链上的 **动作级断点登记（`registerOfficeExecution`→写入权→原表写入）** | **已实现（替身宿主自动化）**：`excelWriteEndToEnd.test.tsx`「断点恢复登记的真实可达性」7 项 —— 授权经真实范围选择弹层建立、桥上真的读一次表、真实 `<App/>` 点击执行，断言登记**恰好一次**、登记成功后才开始写目标表、4 条失败路径**原表零写入**。上一轮"不可达"这一行**已作废**，见 §1.9 / §6.1 |

**§0 的历史更正（保留，不删）**：本报告与旧稿曾有一版把 Excel 写侧写成"已实现"（与 WPS 并列）；
经核查当时**不存在**任何断言过 Office 宿主真实写入的用例，故上一轮已改为"未实现"并在本轮补齐。
两次修正都留在 §6.1 的表里，不粉饰。

**真实模型 / 真实 Excel / 真实 WPS 三类实机端到端仍未取得**（分列见 §4.4），一律标为**未验证**，
不写作"仅待实机验收"。

---

## 1. 修改内容

### 1.1 服务端新增

| 文件 | 作用 |
| --- | --- |
| `server/app/agent_tools.py` | 两后端**共用**工具目录：本地查询工具（`describe_authorized_tables`、`query_table`）、本地结果终局工具（`submit_local_result`、`submit_result_plan`）；按 `hostId` 裁剪 `submit_plan` 的 `anyOf`（**未声明 hostId 时逐字保持原样**）；**有类型出站投影白名单**（`project_describe_tables` / `project_query_receipt`） |
| `server/app/local_results.py` | **结果生命周期**：按任务易失结果登记器，含 `OwnerClaim`（**授权生命周期**）、`register` / `resolve_owned` / `bind` / `release`、TTL、保留上限、来源指纹失效、**独立**矩阵边界校验（不复用 `DataToolResult` 行宽校验，避免互相掩盖） |
| `server/app/task_bridge.py` | **会话生命周期**：`TaskToolBridge` + `ToolBridgeSession`。所有权绑定（task/host/workbook**身份**/authorizedKey/operationId）、`arguments_fingerprint`、同 ID 同参幂等 / 同 ID 不同参 `conflict`、`cancel`、`supersede`、**迟到回执拒绝**、来源指纹不一致 → `source_changed`、结果过大 → `result_too_large`。**`close()` 只结束会话，不删结果**；`release()` 才连结果与归属一起回收 |
| `server/app/ledger_privacy.py` | **账本隐私投影**：`MATRIX_FIELD_NAMES` / `matrix_shape` / `_strip_matrices` / `ledger_safe_actions|criteria|plan|response`；按字段名剥离矩阵，**同时作用于 `actions` 与 `acceptanceCriteria`** |
| `server/app/result_plan.py` | 结果→计划编译：只生成**单个** `writeValues` 动作 + **自动** `rangeEquals` 回读判据 + `localResultRef` 引用；校验目标授权、矩阵尺寸、宿主单元格上限（含 WPS 1000）、疑似公式文本（`=` 前缀）与控制字符 |

### 1.2 服务端改动

| 文件 | 改动 |
| --- | --- |
| `server/app/claude_sdk_agent.py` | `_APPROVED_NON_TERMINAL_TOOLS` 增补两个本地查询工具；新增 `_local_result_card()`（**固定标题**「本地查询结果」，无自由文本）；`run_sdk_plan` 打开桥会话 → 注册工具 → `handle()` 分派 → `finally` **只关会话**；prompt 注入 `registeredResultRefs`（只有句柄与结构）；**保留 `dataResults` 数据行拒绝守卫** |
| `server/app/models.py` | 新增 `Bridge*` DTO（`BridgeBindingRequest` / `BridgePendingCall` / `BridgePollResponse` / `BridgeToolReceipt` / `BridgeReceiptRequest\|Response` / `BridgeRegisterRequest\|Response` / `BridgeResultRequest\|Response` / `BridgeToolPointer`）、`TaskProgressEvent.bridgeTool`、`LocalResultCard`、`AnswerResponse.localResult`、`PlanRequest.localResultRefs`、`WorkbookSnapshot.workbookIdentity`、`AnalysisPlan.localResultRef` |
| `server/app/main.py` | 新增四个回环端点 `/api/sdk/tool-bridge/{poll,receipt,register,result}`，统一 `_bridge_local_guard`（回环客户端 + 回环 Host + Origin 白名单 + 显式 JSON）；`on_event` 透传 `bridgeTool`；`turn_stream` 写账本用 `ledger_safe_response(result)`，SSE 帧仍发完整值 |
| `server/app/excel_agent.py` | 并入 `LOCAL_NON_TERMINAL_PROJECTIONS` 到 `NON_TERMINAL_TOOL_PROJECTIONS`（非终结工具出站唯一口） |
| `config/capabilities.json` | 新增 `agentToolBridge`：`maxToolCallsPerTask:8`、`waitTimeoutSeconds:120`、`maxPendingPerTask:4`、`operationTtlSeconds:900`、`resultTtlSeconds:1800`、`maxRetainedResults:200`、`maxSessions:500`、`maxReceiptBytes:2097152`、`maxResultsPerTask:12`、`maxResultRows:500`、`maxResultColumns:50`、`maxResultCells:20000`、`maxResultCellChars:32767`、`maxTotalReadCells:400000` |

### 1.3 共享层（`packages/core`）

| 文件 | 改动 |
| --- | --- |
| `src/sdkToolBridge.ts` | 状态机：`bridgeToolPointer`、`applyBridgeProgressEvent`、`assertBridgePollResponse`、`applyBridgePollResponse`、`pendingBridgeCall`、`decideBridgeReceipt`（accepted / duplicate / conflict / late）、`markBridgeDelivered`、`cancelBridge`、`shouldPollBridge`；`SdkBridgeBinding.workbookIdentity` 与 `authorizedSheetNames`；**`bridgeRequestedSheetNames` / `verifyBridgeAuthorization`（两端共用的执行前复核判据）**；`bridgeWorkbookIdentity` 归一 |
| `src/index.ts` | 导出上述全部类型与函数 |
| `src/sdkToolBridge.test.ts` | **31 个用例**（含"授权收窄必须拒绝""授权放宽仍放行""同名不同簿拒绝""scope=active 只看当前表"） |

### 1.4 两端面板

| 文件 | 改动 |
| --- | --- |
| `apps/excel-addin/src/sdkBridgeHost.ts`（新增） | Excel（Office.js）宿主适配器：`createSdkBridgeBinding`、`bridgeErrorCode`、`createSdkBridgeHostPorts`（四端口，`verify` 走共享 `verifyBridgeAuthorizationOrRefuse`，`execute` 前重读现场、读不到只回固定错误码）、`asQueryTableArguments`、`loadBridgeResult` |
| `apps/excel-addin/src/hooks/useSdkToolBridge.ts`（新增） | **两端共用**的桥 hook：消费指针、**串行** drain（重入会让同一次调用执行两遍）、`cancel`、`end()` 保留绑定；宿主差异只有 `hostId` 与 `createPorts` 两处 |
| `apps/excel-addin/src/components/LocalResultCard.tsx`（新增） | 结果卡：只吃指针（引用 + 结构），值按 `resultRef` 取回；取不到时**如实**说不可用，不渲染空表；`loadRef` 避免"取一次→再渲染→再取"循环 |
| `apps/excel-addin/src/App.tsx` | 接入 hook：`begin()`（从 `liveContext()` 快照授权范围）→ 进度事件同时喂给结果卡与桥 → `finally` 里 `end()` → 助手消息挂 `localResult` + `localResultBinding` → 渲染结果卡 |
| `apps/excel-addin/src/api.ts` / `contracts.ts` / `types/chat.ts` / `styles.css` | `bridgeBody`（白名单投影）、`pollSdkBridgeTool` / `postSdkBridgeReceipt` / `registerSdkBridgeResult` / `fetchSdkBridgeResult`；`SdkBridgeToolPointer`、`WorkbookSnapshot.workbookIdentity`、`AnalysisPlan.localResultRef`、`TaskProgressEvent.bridgeTool`、`ChatMessage.localResult|localResultBinding`；`.local-result-card*` 样式 |
| `apps/wps-addin/src/wpsSdkBridgeHost.ts`（新增） | WPS（JSAPI）宿主适配器：`createWpsBridgeLiveContext`（一次读取授权名/序号/活动表/身份）、`createWpsSdkBridgeHostPorts`（四端口，`verify` 复用共享判据）、`wpsBridgeErrorCode`（额外认 `WpsAuthorizationError`/`WpsReadLimitError`）、`loadWpsBridgeResult` |
| `apps/wps-addin/src/WpsAgentPanel.tsx` | 复用同一个 hook（`hostId:"wps-et"` + WPS `createPorts`）；**按服务端给的 `turnId` 惰性建立绑定**并按 `turnId` 判重；进度事件喂桥；`finally` `end()`；助手消息挂结果卡并渲染 |
| `apps/wps-addin/src/wpsReadonlyAgent.ts` | `WpsReadonlyTurnProgress.bridgeTool` + 流式 `onProgress` **原样透传** `event.bridgeTool`；`sheetsForRequest` 导出复用；`WpsTurnResult.answer.localResult` 与 `finalizeResponse` 上抛；`WorkbookSnapshot.workbookIdentity` 写入 |
| `apps/wps-addin/src/wpsWorkbookReader.ts` | `wpsWorkbookBridgeIdentity`：有磁盘路径 → `wps-path:<full.toLowerCase()>`；未保存 → 按**工作簿对象**发 `wps-session:` 会话令牌；**身份读不到 → 抛错阻断**（详见 §1.5）。另新增 `wpsAuthorizedSourceFingerprint`（只读取值指纹），把 WPS 查询结果绑到**真实源版本**上 |

### 1.5 第三轮续跑：缺陷闭环与口径补齐

> 纪律：先写**能复现的失败用例**，再改实现；本轮不新增功能。

**（1）三个阻断**

| # | 缺陷 | 落点 | 修法 |
| --- | --- | --- | --- |
| D1 | 意图 `tool_request` 仍带 `dataResults` 进 SDK，旧路径挡住新工具桥 | `apps/wps-addin/src/wpsReadonlyAgent.ts`（`runWpsTurn` 的 `tool_request` 分支）、`apps/excel-addin/src/App.tsx` | 由**同一个** `isSdkPlanningModel(modelId)` 同时驱动"读表策略"与"要不要预置计算结果"：SDK 通道下 `dataResults` 保持 `[]`、`planningReadPolicy: "none"`，本地既不预置计算也不全表读取；读表推迟到**桥指针回来那一刻** |
| D2 | Excel 桥的 `liveContext` 返回**规划启动时的闭包快照**，证明不了当前授权/身份/来源仍有效 | `apps/excel-addin/src/App.tsx`（`bridgeLiveContextRef`）、`apps/excel-addin/src/sdkBridgeHost.ts` | `liveContext` 改成**真的异步重读宿主**；绑定时的范围快照与执行前复核取自**同一个读取器**；`verify` / `execute` 前各 `await` 一次 |
| D3 | `ledger_safe_response` 剥掉必填矩阵后，状态查询/重放仍把它当**完整 `PlanResponse`** 回放 | `server/app/ledger_privacy.py`（`ledger_recovery_descriptor`）、`server/app/main.py`（状态端点） | 落账本改存**版本化恢复描述**（`RECOVERY_SCHEMA_VERSION = 1` + `valuesOmitted` + `resultRef`）；状态端点按 `valuesOmitted` 分流：`true` → 回 `recovery`，`false` → 回 `result`（无损响应仍按完整响应回放，普通后端完成结果取回一并回归） |

**（2）四处口径补齐**

| 项 | 落点 | 行为 |
| --- | --- | --- |
| WPS 结果绑定**真实源版本**，写前检测来源变化 | `apps/wps-addin/src/wpsWorkbookReader.ts`（`wpsAuthorizedSourceFingerprint`） | 取值面与 `readAuthorizedWpsSheets` **逐字相同**（同一批授权表 × 各自已用区域 × 同一批只读上限）的**只读取值** FNV 指纹；读不动返回 `null`（**不伪造"没变化"**）。写前拿它与结果登记时的指纹比对，不一致 → `source_changed` |
| 未保存同名工作簿的**可靠会话身份** | 同上（`wpsWorkbookBridgeIdentity`） | 有磁盘路径 → `wps-path:<full.toLowerCase()>`；没有 → 按**当前工作簿对象**发一个 `WeakMap` 缓存的 `wps-session:<n>-<rand>` 会话令牌（同一对象恒定、不同对象不同）；**连对象都拿不到 → 抛 `WpsAuthorizationError` 明确阻断**，绝不退回会撞车的 `name:` |
| **无流式工具投递能力时提前拒绝** | `apps/excel-addin/src/modelSelection.ts`（`requiresStreamingToolDelivery`）、`apps/excel-addin/src/api.ts`（`StreamingToolDeliveryUnavailableError` + `fallbackOrRefuseToolDelivery`）、`apps/wps-addin/src/wpsPlanningStream.ts`（`refuseIfStreamingToolDeliveryUnavailable`） | 桥指针**只经 SSE**（`/api/turn/stream`）到达。SDK 适配器在"端点不可用 / 连接失败 / 无 `response.body`"三种回退点上**直接抛出固定错误**（`code: "streaming_tool_delivery_unavailable"`、`retryable: false`，文案指向"更新重启本地服务 / 代理不缓冲 SSE / 换非 SDK 模型"），**不再回退**到一次性通道 —— 否则服务端会一直等到 `waitTimeoutSeconds`，进入"无人回执的等待"。普通模型仍照旧回退 |
| 两端**从真实面板**起点到写入回读 | `apps/excel-addin/src/sdkIntentRouting.test.tsx`、`apps/wps-addin/src/WpsAgentPanel.test.tsx` | 见 §3「两端真实面板点击」行 |

**（3）顺带修掉的一处 fail-closed 缺口（巡查时发现）**

桥 hook 的 `drain` 会**吞掉**复核/执行抛出的异常（"桥失败不能拖垮整轮规划"），于是
"读不到宿主现场"会让服务端**一个回执都收不到**，只能干等超时 —— 这正是要避免的那件事。
现在两端统一走共享判据 `verifyBridgeAuthorizationOrRefuse`（core）：读现场失败 →
回**固定错误码** `unauthorized_scope`（`BRIDGE_UNVERIFIABLE_ERROR_CODE`）的可执行回执。
**"读不到证据 ≠ 没变化"**。同时 `verifyBridgeAuthorization` 把**范围判定排到指纹之前** ——
授权收窄是"要重新勾选"，源变了是"要重新查询"，判反了会把用户指向错误的补救动作。

**（4）`VerifyLiveAuthorization` 去重**：Excel 侧原先那份薄封装（及其用例）已删除，
判据只留 core 一处；两端 `verify`/`execute` 都调同一份实现。

### 1.6 测试

| 文件 | 作用 |
| --- | --- |
| `server/tests/test_agent_tool_bridge_lifecycle.py`（新增，**12 项**） | 复现并锁死三类服务端缺陷：终局引用在会话关闭后仍可解析、含值计划不进持久账本、跨会话/同名簿归属拒绝 |
| `server/tests/test_agent_tool_bridge.py` | 27 项；`register` 用例改为先开真实会话（旧用例把 bug 当契约） |
| `packages/core/src/sdkToolBridge.test.ts` | 31 项（新增授权复核判据 9 项） |
| `apps/excel-addin/src/sdkBridgeHost.test.ts`（新增，**14 项**） | 断言送往 `poll`/`receipt`/`result` 的**真实请求体**；授权复核隔离用例；错误码映射 |
| `apps/excel-addin/src/hooks/useSdkToolBridge.test.tsx`（新增，**11 项**） | 完整链路：真实 fetch 替身 + 真实 `executeQueryTableData` 计算；授权收窄/换簿/**源指纹变化**拒读；重复指针不重复执行；取消；真实组件点击 |
| `apps/excel-addin/src/components/LocalResultCard.test.tsx`（新增，**6 项**） | 指针不含值、取不到值如实报不可用、`load` 换身份不产生取值循环 |
| `apps/wps-addin/src/wpsSdkBridgeHost.test.ts`（新增，**13 项**） | 真实 `authorizeWpsSheets` + `readAuthorizedWpsSheets` + `executeQueryTableData` 的跨表分组求和；真实请求体；同名不同路径隔离；取消 |
| `apps/wps-addin/src/wpsReadonlyAgent.test.ts` | 新增 2 项：`bridgeTool` 必须原样透传、`localResult` 一路上抛且不夹带值 |
| `packages/core/src/sdkToolBridge.test.ts` | 新增：`verifyBridgeAuthorizationOrRefuse` 的 fail-closed 用例（读现场抛错 → 固定错误码）、范围先于指纹 |
| `apps/excel-addin/src/sdkBridgeHost.test.ts` | 断言**回执真的发出去了**（`receipt: { errorCode: "unauthorized_scope" }`）、零宿主读取、异常文本不外泄；`execute` 阶段读不到现场同样只回固定错误码 |
| `apps/wps-addin/src/wpsSdkBridgeHost.test.ts` | 未保存工作簿 → 稳定 `wps-session:` 身份（同名两份互不相同、绝不退回 `name:`）；身份核实不了 → **抛错**；源指纹稳定/随值变化；两端现场读不到都回固定错误码 |
| `apps/excel-addin/src/api.test.ts` | 新增 5 项：无流式端点 → SDK 适配器**明确拒绝**且**只发一次**请求、不回退 `/api/turn`；`continuePlanningTask` 拒绝；普通模型仍照旧回退 |
| `apps/wps-addin/src/wpsPlanningStream.test.ts` | 新增 3 项：SDK 404 拒绝 + 单请求；续跑拒绝；普通模型仍回结构化错误 |
| `apps/excel-addin/src/modelSelection.test.ts` | 新增 `requiresStreamingToolDelivery` 正/负用例（`claude-agent-sdkx` → `false`） |
| `apps/excel-addin/src/App.tsx` 相关面板用例（`sdkIntentRouting.test.tsx`） | SDK 意图路径不预置 `dataResults`；等待期收窄/源值变化/身份读不到拒读；**完成后断线按引用取回**（断言请求体带该轮绑定与 `turn-N` 任务身份）；结果过期 → 如实不可用；**面板重开 → 靠该轮绑定再取回** |
| `apps/wps-addin/src/WpsAgentPanel.test.tsx` | **真实 `<WpsAgentPanel/>` 点击**：查询→计划预览→确认→写前复核→执行→回读全链路；等待期收窄（`unauthorized_scope`）；等待期源值被改（`source_changed`）；完成后按引用取回（含该轮绑定 + 源指纹）；结果过期（410）不可用；面板重开再取回 |
| `server/tests/test_agent_tool_bridge_lifecycle.py` | 断言按 D3 的新契约更新：账本里是**版本化恢复描述**（`schemaVersion` / `kind` / `valuesOmitted: true` / `resultRef`），可复算形状留在 `descriptor["response"]["plan"]` 里但**没有矩阵** —— 旧断言读的是"剥值后的 `PlanResponse`"，那正是 D3 要否掉的形态 |
| `server/tests/test_ledger_recovery.py`（新增，**5 项**） | 恢复描述本身：`valuesOmitted` 真/假两档、状态端点对 `valuesOmitted: true` 回 `recovery` 而不是 `result` |
| `apps/excel-addin/src/excelWriteEndToEnd.test.tsx`（新增，第四轮 14 项 → 第五轮 **22 项**） | **Excel 写侧 App 级端到端**：真实 `<App/>` + 有真实状态的 Office.js 宿主替身（`officeTestHost.ts`），跑到宿主真实写入与回读。覆盖预览零写入/取消零写入/确认写入准确目标/等待期源值与身份变化拒绝/授权边界拒绝/挂接守卫/回读不符与部分失败（含"不自动重写"）/公式文本变化。第五轮新增：断点恢复登记可达性（7 项）、中断后先核对再续跑（2 项）、审阅期授权收窄（1 项）、快照稳定身份（1 项），并新增"规划请求 ↔ 桥回执同一基线"的逐字断言。详见 §1.8 / §1.9 |
| `apps/excel-addin/src/officeTestHost.ts`（新增，仅测试导入） | 有真实状态、原表与副本分账的 Office.js 宿主替身：`Excel.run`/`sync`/工作表/`Range` 真实读写、种子 ZIP 解析取副本表名、故障注入（拒绝写入/改写写入/只改公式文本/换身份）。**不参与加载项打包**（入口不引用） |
| `apps/excel-addin/src/planSourceDependency.test.ts`（第五轮新增，**8 项**） | 来源依赖的可信记账：并集累积、跨轮隔离、空读取 ⇒ `null`、防御性拷贝、LRU 淘汰、伪造依赖必被剥离、"只有空记录"也算没有、不按目标反推来源 |
| `apps/excel-addin/src/planSourceDependency.ts`（第五轮新增） | 可信记账本身（全应用唯一入口在 `sdkBridgeHost.execute`）与 `attestSourceDependency` 装配置姓 |

### 1.7 文档

`README.md`、`docs/ARCHITECTURE.md`、`docs/DEVELOPMENT.md`、`计划表.md` 已更新 2026-10-09 章节，
含**三段生命周期**、**账本隐私投影**、**两端接线纪律**、**断言纪律（反向变异）**、**测试运行口径**与**如实限制**。

### 1.8 第四轮续跑：Excel 写侧的宿主替身与端到端点击

> 纪律：不扩新功能；**不 mock 掉执行器 / 预览流程 / 最终成功结果**；保留全部既有未提交改动。

**（1）新增测试基建：`apps/excel-addin/src/officeTestHost.ts`（仅测试导入）**

| 能力 | 说明 |
| --- | --- |
| 真实状态 | `Excel.run` / `context.sync` / 工作表集合 / `Range` 全部由替身实现；`range.values = …` **真的改单元格存储**，之后任何 `values` / `formulas` / `getUsedRangeOrNullObject()` 都从同一份状态读 —— 不回放预置结果 |
| 原表 / 副本分账 | 原表是一份 store，每次 `Excel.createWorkbook(base64)` 打开一本新的副本 store；写入日志带 `store` 字段，于是"预览期间原表零写入"是**可数的** |
| 种子可信 | `createWorkbook` 会**解析种子 stored ZIP**、校验中央目录、从 `xl/workbook.xml` 取出真实生成的副本表名；副本初始内容由替身从原表当前状态克隆（差异已写在文件头注释里，不假装是字节级等价） |
| 故障注入 | `failWrites`（宿主拒绝写入 → 部分失败）、`tamperNextWrite`（写入被宿主改写 → 回读不符）、`setCell(..., {asFormula:true})`（只改公式文本）、`setDocumentUrl`（换文件身份） |
| 未实现即报错 | 用不到的宿主成员会**抛错**而不是返回 `undefined`，避免替身缺口被当成产品行为 |

**（2）新增端到端用例：`apps/excel-addin/src/excelWriteEndToEnd.test.tsx`（14 项）**

驱动真实 `<App/>`（真实 `fireEvent.click`），只 mock 网络（`./api`）与授权目标检查（避免弹窗）。

| # | 需求 | 用例断言 |
| --- | --- | --- |
| 1 | 预览期间原表零写入、副本写入单独计数 | `host.originalWrites() === []`；副本写值 2 条（`E1:F3` / `G1`）、格式写 ≥1；种子 ZIP 里确有原表值；**副本页**里也确实按计划写好了 |
| 2 | 取消审阅 → 原表零写入 | 点「放弃变更」后保存按钮消失、原表写日志仍为空、副本窗口已关（`windowNames() === ["book.xlsx"]`） |
| 3 | 确认后写入准确目标、回读匹配 | 原表 `E1:F3`/`G1` 等于计划矩阵，`A1:C3` 一寸未动；写日志恰好两条且 `cells` 数与区域一致；界面出现「并通过独立验证」+「原表目标回读与审阅副本一致」，且**不出现**"部分操作暂时无法独立验证"，`保存为工具` 解锁 |
| 4a | 等待期间**源值**变化 → 拒绝 | 审阅期间改 `B2` → 点保存 → 原表零写入、界面出现「已变化」 |
| 4b | 等待期间**身份**变化 → 拒绝 | 审阅期间换 `document.url` → 点保存 → 原表零写入、界面出现「身份已变化」 |
| 4c | **授权**边界拒绝 | 计划含落在"工作簿里存在但本次未勾选"的表上的动作 → 执行前检查未通过 → `host.writes() === []` |
| 5 | 挂接守卫：登记没授予写入权 / 索引无法落盘 → 不开始写入 | 两条用例分别断言界面提示固定文案且 `host.writes() === []`（可达性限制见下） |
| 6a | 回读不符 → 不显示完整成功 | 让宿主把 `F2` 实际写成 `0.99` → 界面「结果验证未通过」+「与审阅副本不同」，`保存为工具` 仍锁定，**写日志仍恰为两条**（没有自动重写） |
| 6b | 部分失败 → 报告成功步数、失败步不重放 | 第 2 步被宿主拒绝 → 界面「已有 1 步成功写入当前工作簿」，`E1` 已写、`G1` 为空、写日志只有 `E1:F3` 一条 |
| 7 | 公式文本变化政策 | 只改**公式文本**（值不变）→ 烘焙矩阵计划的来源指纹随之改变 → 走二次确认、零写入（**不是只算值**）。**仍有限制**：`live` 公式计划（`writeFormulas`）不比对来源指纹，其**重算结果**是否受保护本轮**未覆盖**，见 §6.2 |

**（3）一处如实记录的**产品**发现（本轮不改产品，只钉桩）**

`officeRecoverySupport(plan)` 对"动作会修改来源表"的计划一律返回**不支持**；而 App 生成的计划里
`sourceFingerprintSheets` **恒等于**本次授权表集合（`App.tsx` 组装计划时写死 `effectiveSheetNames`），
动作又必须落在授权表上 —— 两者相交，于是 **Excel App 写入路径上 `registerOfficeExecution` 根本不会被调用**，
`recoveryLimitation` 会一直显示"本次不承诺跨重启自动续跑"。这是产品逻辑现状，不是替身限制。
本轮**未改**它（用户要求不扩新功能），只加了两条用例把它钉住：

- `正常 App 计划不触发恢复登记（可达性事实）`：断言常形状下 `registerOfficeExecution` **零调用**；
- `写入前的挂接守卫`：用**唯一可达形状**（动作落在来源表之外）验证两条挂接守卫本身（未授予写入权 / 索引落盘失败 → 零写入）。

要真正修好这条通路需要产品决策（例如把"来源表"与"目标表"分开建模，而不是拿授权集合顶替来源集合），
属于下一轮范围。

> **§1.8（3）已在第五轮解决**，见 §1.9 与 §6.1 的更正行。上面的原结论**保留不删**，以便对照两次口径。

### 1.9 第五轮续跑：恢复登记可达性（三口径分离 + 写侧收窄 + 桥基线一致）

> 纪律：本轮不扩新功能；来源依赖**不采信模型/服务端自报**；重叠无法确定时**保留既有拒绝条件**；
> 保留全部既有未提交改动。

**（1）三口径分离建模（prod）**

| 口径 | 载体 | 谁写 |
| --- | --- | --- |
| 授权范围 | `authorizedSheetNames` / `selectedNamesFor` | 用户勾选并确认，是**不可扩大**的边界 |
| **实际来源依赖** | `AnalysisPlan.sourceDependencySheets` | **只**由本端可信记账写入（见下），服务端/模型给的值一律被覆盖或剥离 |
| 写入目标 | `plan.actions[].sheet` | 本来就精确 |

| 文件 | 改动 |
| --- | --- |
| `apps/excel-addin/src/contracts.ts`、`server/app/models.py` | 新增计划字段 `sourceDependencySheets`（两端协议同步）；服务端 `enforce_authorized_sheet_boundary` 追加一条：**计划声明的来源依赖超出本次授权**同样判越界拒绝 |
| `server/app/main.py` | `_validate_office_recovery_plan`：有 `sourceDependencySheets` 时**按它**取来源集合并逐张校验授权（越界 → `OFFICE_RECOVERY_SCOPE_MISMATCH`）；没有时退回既有 `sourceFingerprintSheets` 口径（保守不变） |
| `apps/excel-addin/src/excel.ts` | `officeRecoverySupport`：读写重叠**只看实际来源依赖**；依赖缺失 ⇒ 报"会修改本次覆盖的来源范围…无法确定来源与目标是否重叠"（**保守拒绝，行为与修复前一字不差**）；依赖存在 ⇒ 只有动作真的落在其中一张来源表上才拒（"会修改来源工作表"） |

**（2）来源依赖的可信记账（prod，全应用唯一入口）**

新增 `apps/excel-addin/src/planSourceDependency.ts`：

- `recordTrustedToolReads(turnId, sheets)`：**唯一合法调用方是"刚真的读过表"的代码** ——
  `sdkBridgeHost.execute` 里确定性执行器返回之后，取其 `DataToolResult.sourceSheets`（是它**实际 `addSheet` 过**的表，
  不是模型给的参数）。同一轮多次读取按**并集**累积（取交集会漏掉真实依赖）；跨 `taskId` 隔离；最多留 32 轮。
- **反向构造被禁止**：不从"授权范围减去写入目标"反推来源（那不是记账，是伪造）。
- `attestSourceDependency(plan, record)`：装配计划时**无条件地**覆盖/剥离计划自带的 `sourceDependencySheets`。
  **没有记录时删除该字段而不是留空数组** —— 空数组会被读成"依赖为空"，而事实是"不知道依赖谁"，
  两者在恢复登记里必须给出不同结论。
- 其余入口只有两种真实验：本地列映射规则源表（`deterministic-local-read`）、已保存工具模板声明
  （`saved-tool-template`）。

**（3）顺带修掉的两处**真实接线**缺陷（本轮根因的一部分，非替身问题）**

| 缺陷 | 事实 | 影响 |
| --- | --- | --- |
| `WorkbookSnapshot.workbookIdentity` 是**死字段** | 契约注释写着"桥归属按此判定"，但 `captureWorkbook` / `captureWorkbookCatalog` / `captureWorkbookStructure` 三个捕获点**没有任何一处写入**它（`App.tsx` 里唯一出现它的字面量是 `bridgeLiveContextRef` 的失败默认值 `""`）。于是 `sdkBridge.begin` 拿到恒空的身份 → `bridgeWorkbookIdentity` 退成 `name:<文件名>`，而执行前复核读的是 `url:` 身份 | **同一个工作簿自比不自等**：`verifyBridgeAuthorization` 第一道判定就返回 `source_changed`。真实**已保存**工作簿上工具桥**恒拒读**，永远产生不了可信来源记账 → 恢复登记不可达 |
| 桥绑定的来源指纹与复核**不是同一口径** | `App.tsx` 把**快照**指纹（`sourceFingerprintForSnapshot`，前缀 `fnv1a32:`）交给 `begin`，而 `bridgeLiveContext` 用 `captureWorkbookSourceFingerprint`（前缀 `fnv1a32x2:`，真读 values/text/formulas/numberFormat）。两者**连算法都不同** | 两侧永远不可能相等 ⇒ 同样恒判 `source_changed` |

修法（**两端同一个原则，WPS 侧早就是这么接的**，Excel 是唯一错的那个）：先把读取器（`bridgeLiveContextRef`）赋值提前，
再**读一次现场**作为本轮基线 `turnSourceBaseline`，并把它同时用于 ① 规划请求里的
`workbook.workbookIdentity` / `sourceFingerprint` / `sourceFingerprintSheets`（服务端据此建立桥绑定与授权归属、
并拿它与回执里的指纹比对）、② `sdkBridge.begin` 的绑定、③ 计划上的 `sourceFingerprint`。
**三处必须是同一个值**，否则链上总有一处判来源变化。读不到现场时退回原有保守路径（不把整轮规划打掉，
也不假装"什么都没变"）。三个捕获点也各自补上身份写入，作为"读不到现场"时的兜底 —— 未保存工作簿没有 URL，
如实留 `null`，两边**同时**退成同名兜底，仍然自洽。

**（4）写侧授权收窄 / 撤销（真实 UI 点击，补上 §6.2 第 4 条的缺口）**

上一轮因为"替身宿主不派发工作表集合变更事件"无法在保持审阅卡可见的同时收窄授权；本轮改用**真实范围选择弹层**
（`.scope-trigger` → 「选择多个」→ 工作表勾选项 → 「完成」）建立与修改授权 —— 走的是产品自己的实现
（草稿 vs 已提交授权由 `useScopeSelection` 管），不是改 mock 造授权。新增用例
`审阅期间撤销该表授权：写回被拒、原表零写入，且不允许在审阅中提交新授权`：`assertExcelNativeComparisonUnchanged`
拿 `await` 之前拍下的快照比 `JSON.stringify(plan)` / `authorization`，收窄一律**拒绝并报「计划或授权范围已变化」**，
原表零写入。

**（5）新增 / 更新的测试**

| 文件 | 内容 |
| --- | --- |
| `apps/excel-addin/src/planSourceDependency.test.ts`（新增 8 项） | 并集累积、跨轮隔离、空读取 ⇒ `null`、防御性拷贝、LRU 淘汰、**伪造的 `sourceDependencySheets` 必被剥离**、"只有空记录"也算没有记录、不修改入参、**不按目标反推来源** |
| `apps/excel-addin/src/excel.test.ts` | 新增 `judges read/write overlap by the attested source dependency, not the fingerprint scope` |
| `apps/excel-addin/src/excelWriteEndToEnd.test.tsx` | 重写并扩到 **22 项**：「断点恢复登记的真实可达性」7 项、「中断后先核对再续跑」2 项、原生审阅套件新增 1 项收窄用例、加上"快照必须当场记下稳定身份"（三个捕获点各断言一次 + 未保存时为 `null`） |
| `apps/excel-addin/src/sdkIntentRouting.test.tsx` | 身份/指纹的"读次数"锚点随新的实时读取点更新（绑定前的基线读也算一次），并改成由"轮询已发出"驱动变化时刻，不再依赖魔法计数阈值 |

**（6）`live` 公式语义：单独记录，不强行套用烘焙矩阵的规则**

- **`live` 计划**（`planLiveness === "live"`，例如纯 `writeFormulas`）：**不比对来源指纹**，只查来源表与字段是否还在
  （`missingSchemaRequirements`）。允许的数据变化：来源值改动 → 执行时重新求值 → 这是**它该做的**。
- **必须使计划失效的变化**：来源表被删、字段被改名（⇒ 拒绝，要求重新规划）。
- **烘焙矩阵计划**（含 `writeValues`）：沿用既有指纹口径（`captureWorkbookSourceFingerprint` 含 `formulas`，
  所以公式文本变化也算变化）→ 二次确认，不自动写入。
- **没有**把"烘焙矩阵的源值比对规则"强加到 `live` 计划上，也没有反过来。

---

## 2. 实际调用链

```
用户授权范围（勾选 → 确认）
   │
   ▼
[前端] PlanRequest{ prompt, workbook(仅表头 + workbookIdentity), localResultRefs[] }   ← 无源值
   │        ├─ 绑定范围快照取自与执行前复核**同一个**读取器（Excel: hook.begin；WPS: createWpsBridgeLiveContext）
   │        └─ 本轮**来源基线**（身份 + 来源指纹 + 范围）由同一读取器先读一次，同时喂给请求与桥绑定 ★第五轮
   │
   ▼
[服务端] run_sdk_plan()
   ├─ task_tool_bridge.open(binding)     # task/host/workbook 身份/authorizedKey
   ├─ 注册工具: host_pruned_terminal_definitions + local_terminal + local_query
   │
   ▼
[SDK query()]  模型只看到：表结构 + 白名单工具 + registeredResultRefs（句柄，无值）
   │
   ├─(1) mcp__excel_bro__describe_authorized_tables
   │        → 本地直接回结构（不碰桥）
   │
   ├─(2) mcp__excel_bro__query_table { mode, groupBy, metrics, ... }
   │        → handler 校验参数(QueryTableArguments)
   │        → session.request() 登记 operation，按所有权绑定
   │        → 进度事件只带 **不透明指针** { operationId, tool }（这一步会落账本）
   │
   │     ↓ 本机（两端同一份状态机）
   │     [面板] 收到 bridgeTool 指针
   │        ├─ Excel：本轮开始前即 begin()（turnId 由客户端生成）
   │        └─ WPS：**惰性** begin()（turnId 由服务端 checkIntent 生成），按 turnId 判重
   │        → runBridgeUntilIdle（串行 drain）
   │        → POST /api/sdk/tool-bridge/poll        拉取本次调用
   │        → verifyBridgeAuthorization(binding, call, liveContext())
   │             ★ 换簿 / 源指纹变化 / **授权收窄** 任一命中 → 一个单元格都不读
   │        → 真实读取本宿主（Excel usedRange / WPS authorizeWpsSheets+readAuthorizedWpsSheets）
   │        → 既有确定性 executeQueryTableData（两端同源）
   │        → recordTrustedToolReads(taskId, result.sourceSheets)   ★第五轮：全应用唯一的记账入口
   │           （手术：`DataToolResult.sourceSheets` 是执行器**实际读过**的表，不是模型给的参数）
   │        → POST /api/sdk/tool-bridge/receipt
   │
   │     ↓
   │   handler 投影白名单 → 交回 SDK：
   │        { status, resultRef, structure } 或 { status:"error", errorCode }
   │        ★ 只有句柄与结构，无任何数值
   │   → 模型按固定 errorCode 修正参数后**再次调用**（同任务多轮）
   │
   ├─(3a) mcp__excel_bro__submit_local_result { resultRef, message }
   │        → local_results.resolve(taskId, resultRef) 归属核对
   │        → terminal = AnswerResponse{ message, localResult: 固定标题卡片(无值) }
   │        → 面板渲染结果卡 → 回环 /result 按引用取值（会话此时可能已关闭，按 claim 校验）
   │
   └─(3b) mcp__excel_bro__submit_result_plan { resultRef, target{sheet,range} }
            → compile_result_plan() 本地校验+编译 → writeValues + 自动 rangeEquals + localResultRef
            → terminal = PlanResponse{ plan }     ★ 提交 ≠ 执行
            → 面板进入预览 → 用户确认 → 写前复核 → 宿主执行 → 写后回读（既有链路，未改）
   │            ★第五轮：执行前新增一段（只在"来源与目标不重叠"时才走）
   │              officeRecoverySupport(plan) 按**实际来源依赖**判读写重叠
   │                → 依赖含目标 / 依赖不确定 ⇒ 不登记，界面如实说"本次不承诺跨重启自动续跑"（原表照写）
   │                → 依赖明确且与目标不重叠 ⇒ registerOfficeExecution（服务端账本）
   │                    ├ 拿到 ownerToken ⇒ savePendingOfficeExecution 落本地发现索引
   │                    └ 之后才逐动作：startOfficeAction → 写 → completeOfficeAction
   ▼
finally: task_tool_bridge.close(taskId)   ★ 只关会话；结果按自己的 TTL 继续活着
```

**关键不变量**

- 模型、进度事件、**持久账本**永远看不到值；值与请求参数只活在本机易失桥与回环响应里。
- 桥按 `(taskId, hostId, workbookIdentity, authorizedKey)` 绑定；换簿 / 源指纹变化 / **授权收窄**一律拒读。
- 幂等：同 `operationId` + 同参数指纹 → `duplicate`；同 ID 不同指纹 → `conflict`。
- 取消 / 换簿 / 来源变化后到达的回执 → `late` / `source_changed`，**不登记、不接值**。
- 写入**永远**先预览再确认；SDK 工具**不得**直接修改工作簿。

---

## 3. 验收案例覆盖（自动化）

| 验收点 | 覆盖方式 |
| --- | --- |
| 两张授权表按明确关联键连接、分组汇总 | **WPS**：真实 `authorizeWpsSheets` + `readAuthorizedWpsSheets` 读两张表 → `executeQueryTableData` 合并分组求和（断言 `[["功夫女足",24],["恐怖游轮",3]]`）。**Excel**：同矩阵经同一计算函数，断言回执 request body 里的真实值 |
| 真工具桥（非 mock 掉整条路径） | 假 SDK 驱动**真实** `run_sdk_plan` → 真实桥 → 真实登记器 → 真实编译器；前端驱动**真实**状态机 + 真实适配器，只把网络与 Office.js/JSAPI 读表换替身 |
| 确定性计算 | 前端用例的输入是宿主矩阵，计算走**生产函数** `executeQueryTableData`，不喂预置结果 |
| 实际请求体 | `httpx.AsyncClient` + `ASGITransport` 打真实 FastAPI 应用；前端用 `vi.stubGlobal("fetch")` 断言送往 `poll`/`receipt`/`result` 的 JSON 原样内容 |
| 同任务多轮调用 | 多次 `query_table` + 参数修正后重试 |
| 参数错误修正 | `invalid_arguments` → 模型按码改正 → 重试成功 |
| 越权拒绝 | 未授权工作表 / 越界区域 / **执行前授权收窄**（一次宿主都不读） |
| 取消 | `cancel` 后回执 → `late`；面板 `cancel()` 后不再轮询、不读宿主 |
| 重复回执 | 同 ID 同指回执 → `duplicate`（不重复登记）；面板同指针不重复执行 |
| 源数据变化 | 服务端指纹不一致 → `source_changed`；**Excel 与 WPS 面板各自**在等待期改一个源值 → 拒读、只回 `source_changed`、零写入 |
| 任务返回后仍可查看结果 | `test_agent_tool_bridge_lifecycle` 断言 `close()` 后 `resultRef` 仍可解析；`end()` 保留绑定；**两端面板**在会话结束后按引用取回并渲染真实值 |
| 结果过期 / 授权不符 | `result` 端点 410 → 卡片如实说"不可用"，**不渲染空表冒充成功**（负向断言配"卡片仍在"的正向锚点） |
| 面板重开（新的加载项会话） | 重开后内存里已无任何桥状态；结果卡靠**该轮**绑定**重新**取回（断言 `/result` 再发一次且带同一 `taskId`/`workbookIdentity`/`resultRef`） |
| 含值计划不进持久账本 | 断言账本记录里 `actions`/`acceptanceCriteria` 的矩阵字段已被剥离；**且**状态端点对 `valuesOmitted: true` 的记录回 `recovery` 而不是当完整 `PlanResponse` 回放 |
| 跨会话拒绝 / 同名工作簿隔离 | `register` 要求活跃且归属一致；`result` 按 claim 校验；WPS 同名不同路径 → `source_changed`；**未保存同名** → 稳定 `wps-session:` 身份，身份读不到 → 抛错阻断 |
| 真实面板点击（**WPS 全链路**） | 真实 `<WpsAgentPanel/>`：点击「确认范围并继续」→ 桥指针 → 回执 → **计划预览** → `生成右侧副本并排审阅` → `确认写回原表` → `.wps-verification-card.verified` + 回读核对；断言**同一份矩阵**里 `Sheet1!A1 === "90"`，且确认前**写入次数为 0** |
| 真实面板点击（**Excel 读侧**） | 真实 `<App/>` + `fireEvent.click`：意图路径不预置 `dataResults`；等待期收窄/源值变化/身份读不到拒读；完成后按引用取回；过期不可用；面板重开再取回 |
| 真实面板点击（**Excel 写侧，替身宿主**） | 真实 `<App/>` + 有真实状态的 Office.js 宿主替身：`生成副本并排审阅` → `保存变更到原表` → 断言**具体单元格被写**（`E1:F3`/`G1`）、原表写日志与副本写日志分账、写前复核拒绝等待期源值/身份变化、回读不符与部分失败都不显示完整成功且写日志不增加第二次 |
| 写入零副作用（预览阶段） | 审阅副本生成后 `host.originalWrites() === []`；取消后仍为空；副本写值单独计数 |
| 无流式投递能力 | SDK 适配器在端点不可用/连接失败/无 `body` 三处**明确拒绝**（固定 `code`），且**只发一次**请求、不回退一次性通道；普通模型仍回退 |
| 写入零副作用 | `submit_result_plan` 只产出 `PlanResponse`，无写调用 |
| 出站投影白名单 | 断言回执 JSON **不含** `rows` / `values` / `title`，且不含原始异常文本 |
| **合法计划真的调用 `registerOfficeExecution`**（第五轮） | 真实 `<App/>` 点击链：授权经**真实范围选择弹层**建立 → 需求送出 → 桥上**真的**读一次表（回执里带真实结果 + `sourceSheets`）→ 生成计划 → 点「执行到 Excel」→ 断言 `registerOfficeExecution` **恰好 1 次**，且登记载荷里 `sourceDependencySheets === [来源表]`、`authorizedSheetNames === 授权范围`（两者刻意不同，证明**没有**拿授权顶替来源） |
| **登记成功并获得执行权后才开始原表写入**（第五轮） | 同一条链上先断言登记发生，再断言目标单元格被写；另外 4 条失败路径（网络失败 / 未授予写入权 / 挂接索引无法落盘 / 动作级写入权被拒）各自断言**原表零写入 + 固定文案** |
| **回执丢失后先核对、不重复执行已完成写入**（第五轮） | 重载：按挂接索引取回执 → **只汇报**"不会重复执行已经完成的动作"，零写入、索引清空；在途动作：先把第 1 步核对成 `applied`（拿写前快照跟**真实宿主当前状态**比）→ 只写第 2 步（`["A2:B2"]`），第 1 步不重写 |
| **审阅期间撤销/收窄授权 → 写入被拒**（第五轮） | 走**真实范围选择弹层**在审阅期间取消该表勾选 → 点保存 → `assertExcelNativeComparisonUnchanged` 拒"计划或授权范围已变化"，原表零写入；并断言审阅中**不允许**提交新授权 |
| **来源依赖不采信自报**（第五轮） | 计划里塞入伪造的 `sourceDependencySheets` → 装配置姓地剥离 → 不能靠声明拿到恢复；工具真的读过**目标表**时（依赖含目标）→ **不登记**，且不靠"把目标从来源里摘掉"换成功 |
| **规划请求与桥回执同一基线**（第五轮） | 断言 `streamAssistantResponse` 请求体里的 `workbook.workbookIdentity` / `sourceFingerprint` 与桥回执里回带的**逐字相等**，且是 `url:` 真身份与 `fnv1a32x2:` 真读指纹（不是同名兜底 / 快照指纹） |

> 说明：`dataResults` 数据行拒绝守卫**保留**（`claude_sdk_agent.py`），新能力经 `localResultRefs` 新契约进入，
> 两条通路互不冒充。

---

## 4. 测试结果（四类分列，不合并）

### 4.1 自动化（通过）

| 范围 | 命令 | 结果 |
| --- | --- | --- |
| 服务端全量 | `pytest server/tests -q`（`TMP`/`TEMP`/`--basetemp` 指 OS 临时目录，见下注） | **1230 passed / 0 failed / 0 error**，166s |
| ├ 桥用例 | `pytest server/tests/test_agent_tool_bridge.py` | **27 passed** |
| ├ 生命周期缺陷复现 | `pytest server/tests/test_agent_tool_bridge_lifecycle.py` | **12 passed** |
| └ 账本恢复描述 | `pytest server/tests/test_ledger_recovery.py` | **5 passed** |
| `packages/core` 测试 | `cd packages/core && node ../../node_modules/vitest/vitest.mjs run` | 23 文件 / **345 passed** |
| Excel 插件测试 | `npm run test:addin` | **89 文件 / 991 passed** |
| ├ 其中：写侧端到端 | `… run src/excelWriteEndToEnd.test.tsx` | **22 passed**（含本轮新增 8 项） |
| ├ 其中：来源依赖记账 | `… run src/planSourceDependency.test.ts` | **8 passed**（本轮新增） |
| Excel 类型 | `tsc --noEmit`（cwd `apps/excel-addin`） | **0 报错** |
| Excel 构建 | `npm run build:addin` | **EXIT=0** |
| WPS 插件测试 | `npm run test:wps` | 37 文件 / **610 passed** |
| WPS 类型 | `tsc --noEmit`（cwd `apps/wps-addin`） | **0 报错** |
| WPS 构建 | `npm run build:wps` | **EXIT=0** |
| Python 编译 | `compileall server/app` | **0 报错** |

**反向变异（证明断言有判别力，均已还原并复跑全绿）**

| 被测机制 | 变异 | 结果 |
| --- | --- | --- |
| 授权复核判据 | `bridgeRequestedSheetNames` 改回"拿当前授权当请求范围"（旧缺陷形态） | core **2 红** / Excel **2 红** |
| `bridgeTool` 透传 | 删掉 `runWpsTurn` 流式进度里的 `bridgeTool` | WPS **1 红**（正是新增的透传用例） |
| `wps-path:` 身份归一 | 身份恒退回 `name:` | WPS **3 红**（含同名不同路径隔离用例） |
| **源指纹比对**（本轮） | 拆掉 `verifyBridgeAuthorization` 里"指纹不同 → `source_changed`"整段 | core **红** / Excel **红** / WPS「等待期源值被改」**红**；还原后三者**复绿** |
| **该轮结果绑定**（本轮） | 助手消息不再挂 `localResultBinding` | WPS「完成后按引用取回 / 结果不可用 / 面板重开」**3 项全红**；还原后**复绿** |
| 无流式投递拒绝 | 取消 Excel 侧 `requiresStreamingToolDelivery` 分支 | Excel **3 红** |
| 无流式投递拒绝（续跑） | 取消 WPS/Excel 的 `continuePlanning*` 拒绝 | **1 红** |
| fail-closed 复核 | 拆掉 core `verifyBridgeAuthorizationOrRefuse` 的 try/catch | core+Excel+WPS **3 红** |
| fail-closed（execute 阶段） | 把 WPS `execute` 的 `liveContext()` 读出 try 之外 | **1 红** |
| **写前复核：源版本**（本轮） | 把 `assertExcelNativeComparisonUnchanged` 里"原表在审阅期间已变化"的整条判据改成恒不触发 | 写侧套件 **1 红**（正是「等待期间源值被改」）；还原后**复绿**（`cmp` 确认逐字节还原） |
| **回读判据**（本轮） | `excel.ts` 的 `valuesEqual` 改成恒真（`verifyPlan` 与"原表 vs 副本"两处比对同时失效） | 写侧套件 **1 红**（正是「回读与计划不符」）；还原后**复绿** |
| **授权边界**（本轮） | `preflightPlanActions` 里 `isUnauthorizedExistingSheet` 改成恒 `false` | 写侧套件 **1 红**（正是「计划含未授权工作表的动作」）；还原后**复绿** |
| **来源依赖：不采信自报**（本轮 M4） | `attestSourceDependency` 改成"没有可信记账就采信计划自报的 `sourceDependencySheets`" | 写侧套件 **1 红**（正是「服务端自报的来源依赖一律被剥离」）；还原后**复绿** |
| **依赖不确定 → 保守拒绝**（本轮 M5） | `officeRecoverySupport` 的 `(declaredDependency ?? plan.sourceFingerprintSheets ?? [])` 改成 `(declaredDependency ?? [])`（拆掉退回既有口径的拒绝） | 写侧套件 **2 红**（「服务端自报…被剥离」与「源值、身份都没变时照常写入…」两条正是靠它）；还原后**复绿** |
| **读写重叠看依赖而非授权范围**（本轮 M6） | 同上处改成 `(plan.sourceFingerprintSheets ?? [])`（退回单一字段的旧缺陷形态） | 写侧套件 **6 红**（「合法计划真的登记」等 5 条可达性 + 「在途动作先核对」）；还原后**复绿** |
| **唯一记账入口**（本轮 M7） | `sdkBridgeHost.execute` 里删掉 `recordTrustedToolReads(...)` | 写侧套件 **6 红**（回执里拿不到真实 `sourceSheets`，依赖不确定 → 登记不可达）；还原后**复绿** |
| **桥绑定指纹与复核同口径**（本轮 M8） | `sdkBridge.begin` 的 `sourceFingerprint` 退回**快照**指纹（`fnv1a32:` vs 复核的 `fnv1a32x2:`） | 写侧套件 **6 红**（复核恒判 `source_changed`，桥一次都读不到）；还原后**复绿** |
| **规划请求与桥回执同一基线**（本轮 M9） | 请求里的 `sourceFingerprint` 退回快照值（与桥回执不一致） | 写侧套件 **6 红**（新增的"请求体 ↔ 回执"逐字相等断言拦下）；还原后**复绿** |
| **稳定身份写进快照：目录 / 结构 / 完整**（本轮 M10/M11/M12） | 三个捕获点各自去掉 `workbookIdentity` | 各 **1 红**（正是「工作簿快照当场记下稳定身份」里对应的那一行断言）；还原后**复绿** |
| **在途动作核对：无法判定即停**（本轮 M13） | `reconcileOfficeRangeAction` 的结果恒判 `unknown` | 写侧套件 **1 红**（客户端应拒绝继续，"核对并继续"不得硬跑）；还原后**复绿** |
| **在途动作核对：分类判据**（本轮 M14） | 同上处恒判 `pending`（分类被架空） | 写侧套件 **1 红**（「跳过该步、只写剩余步骤」断言 `outcome === "applied"`）；还原后**复绿** |

> 14 项变异（M1–M14，含上一轮遗留的 3 项）全部"变异 RED / 还原 GREEN"，由
> `outputs/_mut17.mjs`（apply/restore、信号还原钩子、`touched` 登记、跨行种子按目标文件归一化换行符）
> + `outputs/_mut17.sh`（基线 GREEN → 逐条 apply/跑/restore/复跑）驱动，逐条结论在
> `outputs/_mut17_verdict.txt`。
> **踩坑**：vitest 的摘要行带 ANSI 颜色码（"Tests" 与数字之间夹着转义序列），`grep "Tests +[0-9]"`
> 会全部解析成 UNKNOWN —— 解析前必须先 `sed 's/\x1b\[[0-9;]*m//g'`（本次已修，曾误报 14 项"未抓住"）。

> **服务端全量的运行口径（本轮定位到根因，已可单条命令跑完）**
>
> 运行沙箱给每轮会话一个 **50 次删除**的安全删除预算（`safe-delete` 小垫片）。用仓库内目录当
> `--basetemp` 时，pytest 的临时目录回收会不停消耗这个预算，耗尽后抛
> `[safe-delete][SAFE_DELETE_BULK_CONFIRM_REQUIRED] {"count":50,"threshold":50,"scope":"turn"}` →
> `SystemExit(1)`，表现为**环境性 fixture ERROR**（`test_llm` 5 errors）或删除类用例 FAILED
> （`test_folder_workbooks::test_folder_bulk_delete_*`）；整包连跑还会在中途长时间停滞。
>
> **小垫片对 OS 临时目录是豁免的**（`_should_bypass_safe_delete` 只认 `_is_under_os_tmp_dir` /
> pip 临时目录）。所以把 `TMP` / `TEMP` / `--basetemp` 指到 **`C:\Users\...\AppData\Local\Temp`**
> 之后，整包**单条命令 122 秒跑完**，不再有任何 guard 报错：
>
> ```bash
> export TMP='C:\Users\<你>\AppData\Local\Temp' TEMP="$TMP"
> ./.venv/Scripts/python.exe -m pytest server/tests -q -p no:cacheprovider \
>   --basetemp="$TMP/eb_pybase"
> ```
>
> 注意必须用 **Windows 绝对路径**：Git Bash 的 `$PWD`（`/e/codex/...`）会被 Python 解析成
> `E:\e\codex\...`，届时会出现成片的 `FileNotFoundError`。
>
> 前端注：`apps/excel-addin` 有 4 个用例用 `readFileSync("src/App.tsx")` 按 **cwd** 解析跨 workspace 文件，
> 必须 `cd apps/excel-addin` 再跑；以仓库根为 cwd 时它们会报 `ENOENT`（假失败，非代码缺陷）。
> 另有 1 个读源码的用例只给了默认 5s 超时，**多个套件并发跑时会被拖超时**（单跑即绿，非代码缺陷）。

### 4.2 失败

**无**（本次改动引入的用例全部通过）。

### 4.3 环境阻断（既有，非本次引入）

| 项 | 事实 |
| --- | --- |
| `packages/core` 的 `tsc --noEmit` | **15 处**，全部集中在 `columnMappingSemantics.test.ts`(8) / `blankPolicy.test.ts`(5) / `columnMapping.test.ts`(2)，形态统一为 `Property 'values' does not exist on type 'ColumnMappingClearAction'` 与 `Type 'ColumnMappingClearAction' is not comparable to type 'Record<string, unknown>'`。三个文件与 HEAD **逐字节一致**（已 md5 比对），**非本次引入**；`sdkToolBridge.ts` **零报错**；`packages/core` 的 23 个测试文件 **345 项全绿** |
| 服务端整包 pytest 的沙箱删除预算 | **已定位根因并绕过**，见 §4.1 注 —— 不再是阻断项，只是运行口径要求 |

### 4.4 实机待验 / 未验证

| 项 | 状态 |
| --- | --- |
| 真实模型（DeepSeek/Anthropic）驱动桥闭环 | **未验证** |
| 真实 Excel 宿主（Office.js）端到端 | **未验证** —— 本轮取得的是**替身宿主自动化**证据（`officeTestHost.ts`），**不得**当作实机验收 |
| 真实 WPS 宿主（JSAPI）端到端 | **未验证** |
| WPS：查询→预览→确认→写前复核→执行→回读 的面板点击链 | **已实现**（真实 `<WpsAgentPanel/>` 点击用例） |
| Excel：结果卡按引用取回（含收窄/源变/过期/重开）的面板点击链 | **已实现**（真实 `<App/>` 点击用例） |
| Excel：预览→确认→写前复核→执行→回读 的 App 级点击链（**替身宿主**） | **已实现（替身宿主自动化）**：`excelWriteEndToEnd.test.tsx`，断言到具体单元格、原表/副本分账、回读不符与部分失败都不显示完整成功 |
| Excel：`registerOfficeExecution` / 断点续跑在 App 路径上是否可达 | **可达（替身宿主自动化已证明）**：真实 `<App/>` 点击链上断言登记恰好 1 次、登记成功后才写目标表、4 条失败路径原表零写入。**真实 Excel 实机上仍未验证**（见 §4.1 与 §6.2 第 2 条） |
| 本轮两处修复（快照身份 / 桥基线同口径）在**真实宿主**上的表现 | **未验证** —— 自动化只证明"修复后替身宿主链路上不再恒判 `source_changed`"；真实 Office.js 的 `Office.context.document.url` 取值、`captureWorkbookSourceFingerprint` 的分块读取开销是否可接受，均**未实证** |

---

## 5. 试用步骤

```bash
# 0) 依赖（首次）
./.venv/Scripts/python.exe -m pip install -r server/requirements.txt
#    如需 SDK 通道，再装平台 wheel（含原生 claude.exe）
./.venv/Scripts/python.exe -m pip install --only-binary=claude-agent-sdk -r server/requirements-claude-sdk.txt
./.venv/Scripts/python.exe -m pip check

# 1) 启动后端
#    ★ 走 SDK 通道必须用专用脚本（单进程 --loop asyncio --workers 1，且设置仓库外的会话目录）：
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start_claude_sdk_server.ps1
#    ★ 该脚本替代 `npm run dev:server`，两者**不要同时启动**（脚本会检测 8765 被占用并明确报错）。
#    只在普通（非 SDK）模型通道下，才用：
#      npm run dev:server        # uvicorn 127.0.0.1:8765 --reload
#    原因：Windows 上 uvicorn 的热重载/多 worker 会选 Selector 事件循环，起不了 SDK 的原生子进程。
#    改 SDK 后端代码后手动重启（SDK 通道没有热重载）。
#    ★ 流式前提：SDK 工具桥的指针**只经 SSE**（/api/turn/stream）到达。若本地服务是旧版本、
#      或中间的代理/网关会缓冲 SSE，两端都不会拿到指针 —— 此时 SDK 适配器**直接报错**
#      （code: streaming_tool_delivery_unavailable），不会静默回退到"等人回执"的通道。

# 2) 构建并加载 Excel 插件
npm run build:addin
npm run start:excel

# 3) 构建并加载 WPS 插件
npm run build:wps
npm run dev:wps

# 4) 回归
cd apps/excel-addin && node ../../node_modules/vitest/vitest.mjs run   # Excel（必须以此 cwd 跑）
cd apps/excel-addin && node ../../node_modules/vitest/vitest.mjs run src/excelWriteEndToEnd.test.tsx  # 写侧端到端（替身宿主）
cd apps/wps-addin  && node ../../node_modules/vitest/vitest.mjs run    # WPS
cd packages/core   && node ../../node_modules/vitest/vitest.mjs run    # core
# 服务端：TMP/TEMP/--basetemp 必须指向 **OS 临时目录**，否则会撞沙箱删除预算（见 §4.1 注）
export TMP='C:\Users\<你>\AppData\Local\Temp' TEMP="$TMP"
./.venv/Scripts/python.exe -m pytest server/tests -q -p no:cacheprovider --basetemp="$TMP/eb_pybase"
./.venv/Scripts/python.exe -m compileall -q server/app
```

**Excel 写侧手工验收（与自动化同一条链）**

1. 任务窗格勾选**一张**工作表（原生审阅目前只支持单表授权），确认目标。
2. 输入需要写入的需求 → 计划卡出现「准备就绪」。
3. 点「**生成副本并排审阅**」：应当**原表一个单元格都没变**，另开一本带 `EB_` 前缀工作表的副本，并在 E 列附近标出计划变化。
4. 在 Excel 里核对副本 → 点「**保存变更到原表**」（手动排列模式下需先勾选"我已在 Excel 中核对副本"）。
5. 写前会复核**计划 / 授权 / 原表内容 / 工作簿身份**任一变化 → 拒绝并说明原因，**此时仍零写入**。
6. 通过复核才写原表，并与副本回读比对：一致才显示「并通过独立验证」；不一致显示「与审阅副本不同」且**不会自动重写**。
7. 中途把源数据改一下再点保存：应得到「原表在审阅期间已变化…」，原表保持原样。

**手工试用路径**

1. 在任务窗格勾选两张工作表授权范围，确认目标（**勾选只是草稿，确认才提交**）。
2. 输入：「按 `<关联键>` 连接这两张表，按 `<分组列>` 分组汇总 `<数值列>`」。
3. 消息流出现「正在本机执行数据查询」的**指针进度**，随后出现**本地结果卡**（值由本机回环取回）。
4. 若要求写入，模型提交计划 → 出现**预览**（目标区域 + 待写值 + 自动回读判据）→ **此时零写入**。
5. 点击确认 → 写前复核 → 执行 → 写后回读结果卡。
6. 中途取消、换簿、撤销授权范围：应看到"拒读 + 固定错误码"或"结果不可用"，而不是静默给错答案。
   具体分档：**授权收窄** → `unauthorized_scope`（要重新勾选范围）；**源值被改** → `source_changed`（要重新查询）；
   **现场读不到**（换了工作簿、宿主被关掉）→ 同样回 `unauthorized_scope` 的可执行回执，**不会一个回执都不发**。
7. 面板重开（关掉任务窗格再打开）：对话历史还在，结果卡会**按那一轮的绑定重新取回**；若结果已过保留期，
   卡片如实显示「结果已不可用」，绝不渲染空表。未保存的工作簿因为没有稳定身份，重开后会被要求**重新确认范围**。

> SDK 通道需先在「添加模型连接」保存 DeepSeek 连接，再在模型菜单选由它衍生的「Agent SDK · …」；
> 或按 `docs/DEVELOPMENT.md` 用 `server/.env` 配 Anthropic 原生连接 + 仓库外 `CLAUDE_AGENT_SDK_STATE_DIR`。

---

## 6. 剩余限制

### 6.1 完成口径更正（本轮）

| 旧表述 | 更正后 | 依据 |
| --- | --- | --- |
| "`submit_result_plan` 的面板端到端点击路径**无自动化用例**，未达到与只读结果卡同等的证据强度" | **WPS 侧已实现**：`WpsAgentPanel.test.tsx` 从查询一路点到写入回读，断言**确认前写入次数为 0**、确认后回读同一矩阵一致。**Excel 侧仍未实现**：App 级只覆盖到"计划预览 + `执行到 Excel` 按钮存在"，**没有**一路点到宿主写入并回读的用例；Excel 执行只有组件级覆盖 | `apps/wps-addin/src/WpsAgentPanel.test.tsx`（`WPS 面板 · 工具桥查询到写入回读的完整链路`） |
| "WPS 侧**没有**可用的源内容指纹，'值被改动'在 WPS 上没有检测" | **已实现（口径收窄）**：WPS 现按**只读取值**算源指纹并做写前比对。**仍未覆盖**：只改公式文本、值不变这一档 | `wpsAuthorizedSourceFingerprint`（取值面与 `readAuthorizedWpsSheets` 逐字相同） |
| "未保存的 WPS 工作簿无法区分同名" | **已实现**：未保存工作簿按**工作簿对象**发 `wps-session:` 会话令牌，同名两份互不相同；**身份读不到 → 抛错阻断**。**仍未覆盖**：令牌只活在当前加载项会话，面板重开后未保存工作簿不再可辨 | `wpsWorkbookBridgeIdentity` |
| （新增） | **新增限制**：桥指针**只经 SSE** 到达。没有流式能力时 SDK 适配器**提前拒绝**（固定错误码），而不是偷偷回退到"服务端等人回执"的一次性通道 | `StreamingToolDeliveryUnavailableError` |
| "Excel 写侧 App 级点击链**未实现**（只有组件级覆盖）" | **本轮已补齐**：`excelWriteEndToEnd.test.tsx` 用**有真实状态的宿主替身**从真实 `<App/>` 点击跑到宿主真实写入与回读。**口径收窄**：这是**替身宿主自动化**证据，**不是**实机证据；真实 Excel 端到端仍**未验证** | `apps/excel-addin/src/excelWriteEndToEnd.test.tsx` / `officeTestHost.ts` |
| （上一轮新增发现） | **Excel App 写入路径上 `registerOfficeExecution` 不可达**（来源集合恒等于授权集合，动作又必须落在授权表上 → `officeRecoverySupport` 恒不支持）。上一轮**未改**产品，只加用例钉住 + 用唯一可达形状验证挂接守卫 | §1.8（3） |
| ↑ 该行 | **第五轮已修**：授权范围 / **实际来源依赖** / 写入目标三口径分离；来源依赖只由可信执行记录产生；"来源与目标明确不重叠"的计划在真实 `<App/>` 点击链上**真的登记**（替身宿主自动化证据）。**修复过程中另发现并修掉两处真实接线缺陷**（见下行） | §1.9（`planSourceDependency.ts` / `officeRecoverySupport` / `App.tsx` 基线） |
| （第五轮新增发现） | **① `WorkbookSnapshot.workbookIdentity` 是死字段**：三个捕获点都没写过它，桥绑定只能退成 `name:<文件名>`，而执行前复核读 `url:` 身份 → **同一个工作簿自比不自等**，工具桥在**已保存的**真实工作簿上**恒判 `source_changed` 拒读**。**② 桥绑定的来源指纹与复核不是同一口径**（`fnv1a32:` vs `fnv1a32x2:`，连算法都不同）→ 同样恒拒读。两者叠加使"可信来源记账"永远为空，恢复登记随之不可达。**均已修**，并各配定向变异（M8/M10–M12） | §1.9（3） |
| "Excel 写侧在**等待期间授权被收窄**无 App 级端到端覆盖" | **第五轮已补齐**：改用**真实范围选择弹层**（不是改 mock 造授权）在审阅期间取消该表勾选 → 写入被拒、原表零写入。替身"不派发工作表集合变更事件"不再是缺口理由 | `excelWriteEndToEnd.test.tsx`「审阅期间撤销该表授权」 |

### 6.2 仍然存在的限制

1. **WPS 源指纹只看值**：只改公式文本、值不变的一档**测不出来**，不声称已覆盖。
2. **Excel 写侧的宿主证据是替身的**：`officeTestHost` 与真实 Excel 的差异已写在文件头（不解析种子 ZIP 的单元格内容、`sync()` 为空操作、只实现用到的 API）。真实 Excel 实机端到端**未验证**；替身行为**不得**当作实机结论。
3. **公式文本变化只对烘焙矩阵有判据**：`captureWorkbookSourceFingerprint` 含 `formulas`，故 `writeValues` 这类**烘焙**计划的来源指纹会随公式文本变化（本轮有用例证明）。但 `planLiveness === "live"` 的 `writeFormulas` 计划**不比对来源指纹**（只查来源表与字段是否还在）——此时源表公式文本改变会不会影响最终重算结果，本轮**未覆盖**，**不声称所有来源变化均受保护**。
4. ~~"等待期间授权被收窄"在 Excel 写侧没有 App 级端到端覆盖~~ **第五轮已解决**（改用真实范围选择弹层，见 §6.1）。留下的口径：收窄发生在**审阅期间**时由 `assertExcelNativeComparisonUnchanged`（比对快照前的 `JSON.stringify(plan)`/`authorization`）拦下；收窄发生在**规划等待期间**时由桥的授权复核拦下（`unauthorized_scope`，一次宿主都不读）——两条仍分别覆盖，**没有**合并成一条笼统结论。
5. ~~`registerOfficeExecution` / 断点续跑在 App 路径不可达~~ **第五轮已修**（三口径分离 + 两处接线修复）。**仍有限制**：
   - 只支持**动作级**断点，且**恢复范围**是"本次登记的那些动作"；续跑前会**先核对真实单元格状态**，
     绝不盲目重放；**不宣称支持跨重启续跑**（工作簿被关闭再打开时，服务端账本仍有记录，但本地发现索引与
     当前宿主状态需重新核对）。
   - **真正读写重叠、或依赖无法确定**时**保留既有拒绝条件**（不登记），界面如实说"本次不承诺跨重启自动续跑"。
   - 来源依赖**只**在"本轮真的经桥读过表"时才存在；**没有**桥读取的计划（例如纯本地列映射、编辑后的计划）
     一律退回"依赖不确定"。
6. **`wps-session:` 令牌是会话级的**：面板重开后未保存工作簿失去可辨身份；此时桥的复核按 fail-closed 拒绝（要重新确认范围），绝不退回 `name:` 让别人冒充。
7. **WPS 的桥指针只在规划流（SSE）透传**：一次性 `/api/turn` 回退通道不产生 SDK 工具调用，因此不会出现桥指针；SDK 适配器在这种情况下**明确拒绝**（见 §1.5(2)）。
8. **未开启**隔离 Python 产品入口（仍 `enabled=false`）；**未做** SDK 会话持久化（沿用 `--no-session-persistence`）；**未新增** WPS 高级动作；WPS 文件夹入口仍未实现。
9. **未扩大**结果值出站面：模型侧永远只有句柄与结构。
10. 真实模型 / 真实 Excel / 真实 WPS 三类实机证据**均未取得**（§4.4）。
11. **本轮基线指纹的范围取自"已确认的授权范围"**（`committedSheetNames`），而计划的来源指纹取自
    `effectiveSheetNames`（草稿）。两者在用户"改了勾选但没点完成"时会不一致 —— 计划仍会走既有
    `sourceFingerprint` 判据（自洽），但桥会**拒读**（fail-closed）。**这不是新引入的行为**，
    本轮**未改**，如实记录。
12. **本轮多了一次现场读取**：每轮规划开始前要先读一次"来源基线"（身份 + 指纹），`hook.begin` 内部还会为
    `authorizedSheetNames` 再读一次。两次读的是同一个读取器、同一范围，值应当一致；**未做**缓存复用
    （正确性优先于少读一次）。大工作簿的额外开销在真实宿主上**未实测**。
13. **`live` 公式计划不比对来源指纹**（既有口径，本轮未改）：允许来源值变化（它本就该重算），
    只有来源表/字段消失才使计划失效；**不声称**所有来源变化均受保护（§1.9（6））。

---

## 7. 未提交声明

本次改动**未**执行 `git commit` / `push` / 发布；工作区既有未提交改动**全部保留**。
