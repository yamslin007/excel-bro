# 隔离 Python 分析：源码核查结论与分阶段实施计划

日期：2026-10-06
性质：**只做源码核查与实施规划，未写任何实现代码**（本轮未改动运行时代码、配置、依赖）。
依据：`AGENTS.md`、`docs/ARCHITECTURE.md`、`docs/DEVELOPMENT.md`、`计划表.md`、
`docs/SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md`（完整）、
`docs/DEEPSEEK_ISOLATED_PYTHON_PROMPT_2026-10-06.md`。

> 补充事实：用户本次贴给 DeepSeek 的提示词，其前一版本**已经在仓库里**
> —— `docs/DEEPSEEK_ISOLATED_PYTHON_PROMPT_2026-10-06.md`（设计文档 §10 引用了它）。
> 本次贴出的是扩写版。两份都定位为"交给编码代理实施"，不是生产 system prompt。

---

## 一、源码根因核查（结论均带 `文件:行号`）

### 1.1 执行器：零实现（不是"待实机验收"）

对 `server/app` 全量检索 `subprocess` / `sandbox` / `JobObject` / `docker` / `wsl` / `isolat`
（`grep -rni`），**只有 3 处命中，全部与执行无关**：

- `claude_sdk_agent.py:101` 注释（说明没有 subprocess transport）
- `excel_window_layout.py:4` 注释
- `task_ledger.py:147` `sqlite3.connect(..., isolation_level=None)`（SQLite 参数，非沙箱）

**结论**：隔离执行器、run_python_analysis 工具、协调器、datasetHandle/resultHandle、双端入口，
全部**未实现**。设计文档 §10 的"运行时未实现"属实。这必须写成"未实现"，
**不能**写成"已实现，仅待实机验收"。

### 1.2 SDK 工具注册：无通配符，但**非终结结果被整体拒绝**（关键阻塞点）

| 项 | 位置 | 现状 |
|---|---|---|
| 工具定义源 | `excel_agent.py:1298` `AGENT_TOOLS`、`:1302` `AGENT_TOOLS_WITH_DATA = _terminal_tools()`、`:1303` `PLAN_ONLY_TOOLS` | 注册进 MCP 的是**纯终结工具集** |
| 注册 | `claude_sdk_agent.py:224` 遍历 → `:249` `sdk.tool()` → `:251` `create_sdk_mcp_server` | — |
| 权限 | `claude_sdk_agent.py:250` `allowed.append(f"mcp__excel_bro__{name}")`、`:79` `allowed_tools=allowed`、**全名精确匹配**；`:74` `tools=[]`、`:75` `disallowed_tools`、`:77` `strict_mcp_config=True`、`:80` `permission_mode="dontAsk"` | 满足"仅显式注册、不开通配符" |
| 终结/非终结判定 | `excel_agent.py:1444` `_execute_tool` 返回 `AssistantResponse \| dict`；返回 `dict` 即非终结（`:1448` `select_plan_actions`） | — |
| **SDK 侧非终结返回** | `claude_sdk_agent.py:239-240`（已复核原文） | `if isinstance(result, dict): return {"content":[{"type":"text","text":"非终结工具结果不可用。"}], "is_error": True}` |

**这条是实施中最硬的阻塞点**：`run_python_analysis` 按设计**天生是非终结工具**
（模型调用 → 拿结果 → 继续推理）。但 SDK 后端现在**一条非终结结果都不返回**。
所以在 SDK 上接这个工具，不是"加个工具"就完事，必须新增
**"已批准非终结工具的有类型投影白名单"**——这正是设计文档 §5 那句
"只有新增/已批准工具的有类型投影能返回"的具体落点。不做这一步，
工具在 SDK 后端永远只会拿到「非终结工具结果不可用」。

### 1.3 普通后端存在"任意 dict 直接出站"（隐私风险点，已复核）

`excel_agent.py:2402`（原文已复核）：

```python
tool_content = json.dumps(result, ensure_ascii=False)
if cache_key is not None:
    cache[cache_key] = tool_content
```

非终结工具的任意 dict 被 `json.dumps` 后直接进模型消息流，并写进缓存。
**新增 run_python_analysis 若不先改这里**，等于顺手给模型开一个任意 dict 出口，
且出口内容包含 Python 的 stdout / 结果 / 异常——正是设计文档 §4 禁止的
"嵌入数据的代码与输出默认不能出站"。

对照：SDK 侧已经堵住了（`claude_sdk_agent.py:254-269` 手工构造 header-only payload，
注释明确 `Never serialize the request's snapshots, lastResult or dataResults`）。

### 1.4 planningReadPolicy / dataResults 守卫现状

- 定义：`models.py:2033-2034`（默认 `ifNeeded`）、`DataToolResult` `models.py:1864`
- SDK 强制收紧：`claude_sdk_agent.py:209-213`（已复核）——
  `if request.planningReadPolicy != "none" or request.dataResults: raise ModelNotAvailableError(...)`
- 普通后端：`excel_agent.py:1449`（`none` 禁用只读工具）、`:1851`（lastResult 条件剥离）、`:1898-1902`（选工具集）
- **既有违反点**：`excel_agent.py:1853` 无条件 `json.dumps` 全部 `dataResults`；
  `:270-313` `_read_range` 返回 `values` 矩阵，再经 `:2402` 回灌模型

实施时**不得**为接 Python 而删这些守卫（用户明确要求）。

### 1.5 已有可复用的机制（好事，别重造）

| 需求 | 现成机制 | 位置 |
|---|---|---|
| 源指纹 / 授权版本 | `sourceFingerprint` + `sourceFingerprintSheets` | `models.py:116`、`:965`；计算 `folder_workbooks.py:2209/:2219`；核对 `folder_workbooks.py:4553-4561`、`main.py:1190`、`models.py:1379-1384` |
| 目标变化检测 | `checkAuthorizedTarget` → `same / file_changed / scope_changed / unavailable` | `packages/core/src/authorizedTarget.ts:13`；`useWorkbookContext.ts:261` watch；WPS `wpsAuthorizedTarget.ts:27-30` + `authorizationKeyOf`（`WpsAgentPanel.tsx:506`）→ `authorizationRevision+1`（`:1134/:1754`） |
| 授权草稿 vs 已提交 | `selectedSheetNames`（草稿）vs `committedSheetNames`（已提交） | `useScopeSelection.ts:89` / `:96`；确认 `:168/:182`、取消 `:194` |
| 每次读前核对 | `assertWpsAuthorization` | `wpsWorkbookReader.ts:183` |
| AnalysisPlan 白名单（约 35 类动作） | `ExcelAction` 联合类型 | `models.py:671-705+`；`AnalysisPlan` `models.py:961` |
| 计划校验 | `/api/plans/validate` | `main.py:3161`；`models.py:1371/:1379`；能力收窄 `excel_agent.py:991`；文件夹 `folder_workbooks.py:2787` |
| 写入守卫（revision / ownerToken / 指纹回读） | `_wps_evidence` 写前/写后指纹回读；owner 租约 | `main.py:1346`、`:1298-1300`、`:1256`、`:1173`；`models.py:2102-2162`、`:2249-2323` |
| 进度事件 | `TaskProgressEvent` + SSE | `models.py:2542` / `contracts.ts:2409`；SSE `api.ts:302`、`wpsPlanningStream.ts:145` |
| 进度展示 | Excel `useActivityProgress` + `TaskProgressTimeline`；WPS `beginActivity` + `wpsActivityLog` + `progressCard` | `App.tsx:8422`；`WpsAgentPanel.tsx:1921/4662`、`wpsActivityLog.ts:15-66/69/76/103` |
| 取消 | AbortController 全链路 | WPS：`:1174`、`:2191/:2742/:3520`、`wpsWorkbookReader.ts:306/338` 抛 AbortError；Excel：`App.tsx:1048/1051/2394`、`api.ts:263/372` |
| 恢复记录白名单投影 | `sanitizeTaskResumeRecord`（前端） | `packages/core/src/taskResume.ts:299`，白名单 `:321-339`，条件挂 turnId/operationId/requestId `:334-339` |
| 账本落盘不含矩阵 | `models.py:2165-2188` 无矩阵字段、`extra="forbid"` `:2176/:2202`；`main.py:2136` 只写 `descriptor.model_dump()`；旧记录兜底 `main.py:1332-1343` | — |
| 隐私出站测试基建 | 替换 `globalThis.fetch` 捕获真实 body | `wpsLocalPlanOutbound.test.tsx:431-434`（断言 `:523/626/633`）；Excel `localMappingOutbound.test.ts:64-67` |
| 宿主写入计数探针 | 数 `Value2`/`Formula` setter | `wpsLocalPlanOutbound.test.tsx:352`；`WpsAgentPanel.test.tsx:3545`；Excel `excelNativeComparison.test.ts:30`（`:94/:108`）、`excel.test.ts:1185` |
| 默认关闭的能力开关写法 | `capability_bool(段, 键, default=False)` | `"safety": {"allowHyperlink": false}`；消费 `safety.py:51`；宿主开关 `capabilities.json:71/118/165` 的 `enabled` |
| 宿主能力矩阵 | `hostCapabilities` supported/unverified；前后端由 `config/action_capability_cases.json` 锁定 | `capabilities.json:69-208`；`core/src/actionCapability.ts` ↔ `server/app/action_capability.py` |

### 1.6 尚不存在的东西（要新建）

- **`datasetHandle` / `resultHandle`**：源码中**完全不存在**，仅出现在提示词文档里。
  需新建，但应建在既有 `sourceFingerprint` + `authorizationRevision` 之上。
- **可信协调器**：无。
- **有类型出站投影层**：无（现状是 `json.dumps` 任意 dict）。
- **`pythonExecution` 配置段**：无。`config/capabilities.json` 顶层现有
  `version / authorizedTarget / ... / agent / formula / safety / llm / turns / progress / taskRecovery` 等段；
  `agent` 段（`:256-269`）含 `maxTurns / claudeSdkMaxBudgetUsd / maxTotalTokens /
  maxWallTimeSeconds / maxConsecutiveErrors / maxCacheHitsPerQuery / maxToolPathLoopLength /
  warningTurn / maxReadRows(200) / maxReadColumns(30) / ...`。新限额建议新建 `pythonExecution` 段。

### 1.7 协议同步机制（新增类型必须走的清单）

协议同步**靠测试而非代码生成**：`server/tests/test_protocol_consistency.py:385`
比对后端 `ExcelAction` 的 type Literal 集合与前端样本 `VALID_ACTION_SAMPLES`（`:69`）的 key；
`:345` 让 Pydantic 实解析样本；`:524` 要求每个动作被文件夹执行器实现或显式 unsupported；
`:548` 反查陈旧项。前端运行时校验在 `contracts.ts:1722` `allowedActionTypes`、
`:1799` `assertAssistantResponse`、`:2063/:2075` `assertEditablePlan`。

新增动作需同改：`models.py` → `contracts.ts` → `VALID_ACTION_SAMPLES` → `excel.ts` →
`folder_workbooks.py`（清单 `:741-747`）。
新增结果类型需同步：`PlanResponse`（`models.py:2483`）/ `AssistantResponse`（`contracts.ts:1562`）/
`TaskProgressEvent`（`models.py:2542`、`contracts.ts:2409`）。

### 1.8 日志风险点

- `excel_agent.py:2026/2108/2543` 只记计数，无原始值（安全）
- **风险**：`formula_tools_integration.py:46/95/107` 打印参数与内容片段 ——
  Python 的 stdout/异常若走同类日志路径会泄数据，实施时需一并纳入出口策略

---

## 二、环境核查：本机**当前没有可验证的隔离运行时**（环境阻断）

| 项 | 实测结果 |
|---|---|
| OS | Windows 11 专业版，Build 26100（24H2） |
| Docker | CLI 在 PATH，但**守护进程未运行**：`failed to connect ... npipe:////./pipe/dockerDesktopLinuxEngine` |
| Docker 后端 | 报的是 `dockerDesktopLinuxEngine` → **依赖 WSL2** |
| WSL | `wsl.exe` **被本沙箱的程序黑名单阻断**，无法在本会话启动或验证 |
| Hyper-V 特性 | `Get-WindowsOptionalFeature` 返回 unknown（需管理员权限，未取得） |

**结论**：设计文档 §4 要求"先核对目标 Windows 部署能力，选择可实际提供隔离的方案并实测"。
在**当前这台机器、当前这个会话**里，这条**做不到**——没有可启动、可验证的隔离运行时。

因此：

- 阶段 1（隔离原型）在本机**环境阻断**，不能靠 mock 或配置文字蒙混过去。
- 要推进，必须先解决以下**任一**前置（都需要用户决策/授权）：
  1. 启动 Docker Desktop（需 WSL2 可用；本会话 wsl.exe 被黑名单阻断）
  2. 采用 Windows 沙箱 / Job Object + 受限令牌等 OS 原生方案并逐一实测
  3. 换一台具备可用隔离运行时的机器做阶段 1 的取证

这些都属于"安装或启用新运行环境"，按用户要求**必须先申请授权**。

---

## 三、分阶段实施计划

> 每阶段都有**退出条件**：不通过就停在当前阶段，不接真实数据、不启用功能。

### 阶段 0：契约与配置骨架（默认关闭）

- 目标：把类型、限额、开关先定下来，且**默认关闭**，此时功能对 UI 完全不可见。
- 要改：
  - `config/capabilities.json` 新建 `pythonExecution` 段（代码长度、累计读取/运行额度、
    CPU/内存/墙钟/scratch 上限、输出大小、默认 `enabled: false`）
  - `apps/excel-addin/src/contracts.ts` ↔ `server/app/models.py` 同步新增联合类型
  - 同步 `server/tests/test_protocol_consistency.py` 的 `VALID_ACTION_SAMPLES` / 结果类型样本
- 验收：新增的类型一致性测试通过；`enabled: false` 时两端入口不渲染。
- 退出条件：协议一致性测试绿；开关默认关闭且无 UI 入口。

### 阶段 1：隔离原型（**本机当前环境阻断**）

- 目标：纯合成输入，先证明"越权读文件/凭据、联网、回环、外部程序、逃逸"被**环境**拒绝。
- 前置：解决上文"环境核查"里的阻断（需授权）。
- 验收（设计文档 §9 表，逐条取证）：
  - 越权读宿主目录/仓库/工作簿/配置/账本 → 被拒
  - 互联网 / 局域网 / **回环** → 被拒
  - 外部程序、pip、派生进程 → 被拒
  - 不继承 API Key 与宿主环境变量
  - CPU/内存/线程/进程数/墙钟/scratch 上限可验证
  - 超时、取消、失联 → **整个执行树**终止，迟到结果被拒
  - 运行器缺失 / 配置损坏 → 失败关闭，**零裸执行**
- 不可替代证据：真实集成测试。cwd、venv、临时目录、AST 检查、禁词、monkeypatch、
  mock SDK、配置截图**一律不算**。
- 退出条件：上表全项取证通过。未通过 → **停止，不进入阶段 2**。

### 阶段 2：可信协调器 + 句柄 + 有类型出站投影

- 目标：任务内句柄、授权绑定、输入输出验证、本地投递、生命周期、模型投影。
- 要改（在既有指纹机制上建，不重造）：
  - 新建 `datasetHandle`：绑定 任务 + 工作簿/工作表身份 + 授权版本 + 范围 + `sourceFingerprint`
    （复用 `models.py:116/:965`）。**每次访问都核对**，禁跨任务、禁扩范围。
  - 新建 `resultHandle`：随机、任务内、短期，随取消/授权变化/会话结束失效。
  - **修 `excel_agent.py:2402`**：把"任意 dict → `json.dumps` 出站"改成
    **已批准工具的有类型投影**；未登记工具不得直出。
  - **修 `claude_sdk_agent.py:239-241`**：新增"已批准非终结工具"白名单与有类型投影，
    否则 `run_python_analysis` 在 SDK 上永远拿不到结果。
  - 一并纳入 `formula_tools_integration.py:46/95/107` 的日志出口。
- 验收：隐私用例**先断言本地哨兵存在**，再捕获真实 `fetch` body / 模型工具消息 / SSE /
  日志 / 存储 / 账本 / SDK 出口。stdout、错误、摘要、嵌入数据的代码**分别**测。
- 退出条件：四类出口攻击分别被拒；无哨兵泄漏。

### 阶段 3：run_python_analysis 工具接入（普通后端 + SDK 同语义）

- 目标：两端共用工具协议、授权与结果语义。
- 要改：`excel_agent.py` 新增工具（入参禁路径/命令行/pip 指令/宿主对象/密钥）；
  SDK 侧走阶段 2 新增的非终结白名单；`AGENT_TOOLS_WITH_DATA`（`excel_agent.py:1302`）保持终结集语义，
  Python 工具**单独建模**，不改现有终结工具行为。
- 不改：现有模型连接配置、不自动切供应商、不要求用户重填配置。
- 退出条件：普通后端与 SDK 后端对同一工具产生**语义一致**的结果与错误类别；
  执行失败只返回安全错误类别，原始 stderr/traceback 留受控本地诊断。

### 阶段 4：结果 → 白名单 AnalysisPlan 编译 + 本地预览

- 目标：Python 输出作为**不可信数据**验证后编译成既有白名单计划。
- 校验项：类型、尺寸、坐标、目标授权、**公式与文本语义**（防字符串被隐式当公式）、
  宿主动作能力。不支持的动作**明确拒绝**，不为 WPS 假造支持。
- 结果矩阵**留本地**，只用 `resultHandle` 引用；不进恢复描述、不进 HTTP 规划请求、
  不进 SDK 工具文本。
- 复用既有链路：WPS `planCard`（`WpsAgentPanel.tsx:4425`，含"修改前→修改后"）→
  副本 `:4649` → `confirmPendingPlan`（`:2978`）→ 验证卡 `:5383`；
  Excel `useExecutionApproval.ts:16` → `App.tsx:5377/5251` → 验证卡 `:8670`。
- 退出条件：无合法动作时**显示本地分析结果**，不伪造写入支持。

### 阶段 5：双端接线（Excel / WPS / 文件夹）+ 进度·取消·恢复语义

- 入口流程：确认目标与选表 → **隔离就绪检查** → 当前授权范围读取 → Python 分析 →
  本地结果/副本预览 → 用户确认 → 既有白名单执行与回读。
- 目标变化确认必须在**读取/分析之前**（复用 `checkAuthorizedTarget` `core/src/authorizedTarget.ts:13`），
  不是失败后才提醒。
- 进度：扩 `TaskProgressEvent` 的 stage 枚举与 `wpsActivityLog.ts:15-66` 的 stage→phase 映射，
  **不另起机制**。分清 排队 / 分析 / 结果可审阅 / 等待写入确认。
- 取消：复用既有 AbortController 链路；先停止数据交付与发布，再终止执行树，
  迟到响应**不得**生成可执行计划。
- 恢复：只存身份摘要、授权版本、范围、状态、指纹、进度；**不存**值矩阵、原始输出、
  带数据字面量的代码、模型转录。保持 `no-session-persistence`。
  重启后内存结果丢失 → 明确显示"**重新读取并分析**"，**不冒充精确续跑**。
- 分析状态与工作簿写入状态**分开**；未知写入先核对，**不能**把 `unknown` 变成 `not_written` 重放。
- 授权动作与澄清答案分开保存；已回答信息按 `questionId` 保留/替换，不循环追问。
- 纯分析任务**不强加**写入确认，但**不跳过**数据授权。
- 退出条件：两端实际点击完整链路，证明 确认前新增宿主写入为 0；取消/授权变化/现场冲突时 0；
  确认后仅执行批准计划且独立回读。

### 阶段 6：证据、反向变异、回归、文档

- 反向变异（**独立**拆守卫，各自必须变红，还原后复跑）：
  隔离失败中止 / 出站投影 / 句柄授权绑定 / 预览确认 / 取消结果发布门。
  **变异红灯不能代替真实沙箱验证**（阶段 1 的证据是另一回事）。
- 回归（AGENTS.md）：`npm run test:addin`、`npm run build:addin`、`npm run test:wps`、
  `npm run build:wps`、`python -m pytest server/tests -q`、`python -m compileall -q server/app`；
  改 core 另加 `npm run test:core`。
- 文档：同步 `docs/ARCHITECTURE.md`、`docs/DEVELOPMENT.md`、根目录 `README.md`、`计划表.md`；
  新增启动依赖、接口、配置、环境变量一并记录。
- 报告**分六类**：源码根因 / 实际实现 / 测试通过 / 失败 / 环境阻断 /
  未实现与实机待验（按 Excel/WPS/文件夹 × 普通/SDK 分列）。

---

## 四、风险登记

| 风险 | 说明 | 处置 |
|---|---|---|
| **隔离运行时缺失** | 本机无可用隔离运行时，阶段 1 环境阻断 | 必须先申请授权解决前置，不达标不进阶段 2 |
| **SDK 非终结结果被拒** | `claude_sdk_agent.py:239-241` 使得新工具无法回结果 | 阶段 2 必须新增已批准非终结工具的有类型投影白名单 |
| **任意 dict 出站** | `excel_agent.py:2402` 直 `json.dumps` | 阶段 2 一并改造，否则新工具自带一个任意出口 |
| **为接功能删隐私守卫** | `planningReadPolicy` / `dataResults` 可能被顺手删掉 | 用户明令禁止；阶段 2 只做加法 |
| **把"摘要"当安全** | 数值聚合、stdout、异常、图像都可能编码原数据 | 出口策略统一管，不按名称放行 |
| **重造轮子** | 指纹、授权检测、写入守卫、进度、取消均已存在 | 阶段 4/5 一律复用既有链路 |
| **假完成** | 用开放 Bash / 宿主裸执行 / mock 凑证据 | 阶段 1 与阶段 6 的证据清单里已逐条排除替代品 |
| **单端绿灯** | 只做 Excel 就宣称完成 | 阶段 5 要求两端实际点击链路 |

---

## 五、未实现与实机待验清单（当前状态）

### 未实现（**不得**写成"仅待实机验收"）

| 项 | 范围 |
|---|---|
| 隔离执行器 | 全部（server 侧零实现） |
| 可信协调器 | 全部 |
| `run_python_analysis` 工具 | 普通后端 + SDK 均未实现 |
| `datasetHandle` / `resultHandle` | 全部 |
| 有类型出站投影层 | 全部（现状是任意 dict 直出） |
| 结果 → AnalysisPlan 编译 | 全部 |
| 双端入口（Excel 当前工作簿 / WPS / openpyxl 文件夹） | 全部三端 |
| `pythonExecution` 配置段与默认关闭开关 | 全部 |

### 实机待验（实现之后才存在的项目）

- 真实 Windows 宿主上 Excel 与 WPS 的完整点击链路（确认前 0 写入、取消 0 写入、确认后独立回读）
- 真实隔离环境下的越权/联网/逃逸取证（依赖隔离运行时）
- 开放真实模型连接后的端到端行为（需付费授权）
- 获准的独立测试工作簿上的验收（**不得**操作真实客户工作簿）

### 平台未验证

- 本会话：Docker 守护进程未运行；`wsl.exe` 被沙箱程序黑名单阻断；Hyper-V 特性状态未取得（需管理员）。
  → **本平台未验证任何隔离方案**。

---

## 六、本轮交付边界

- 未修改任何运行时代码、配置或依赖；未安装环境；未提交 Git；未打包；未操作真实工作簿。
- 未删除或 stash 任何既有未提交改动。
- 产出只有本文档。

---

## 七、落地记录（2026-10-06 批次：失败关闭骨架 + 出站投影层）

本批把计划中**不依赖隔离运行时**的部分先行落地，全部按失败关闭纪律实现。隔离执行器本体仍然**未实现**，功能总开关保持默认关闭。

### 已实现

| 项 | 文件 | 说明 |
|---|---|---|
| `pythonExecution` 配置段 | `config/capabilities.json` | `enabled: false`（显式默认关闭）+ 9 项资源限额集中在配置 |
| 可信协调器（失败关闭） | `server/app/python_analysis.py` | 无任何执行能力；`execute_python_analysis` 只返回固定不可用回执（disabled / runner_missing / rejected）；回执字段白名单投影 `RECEIPT_KEYS`；工具 schema（code + datasetHandles，additionalProperties=false） |
| 工具分发分支 | `server/app/excel_agent.py` `_execute_tool` | `run_python_analysis` → 协调器，永不触达宿主 |
| **非终结工具有类型出站闸** | `server/app/excel_agent.py` | `_non_terminal_tool_content` 唯一出站口 + `NON_TERMINAL_TOOL_PROJECTIONS` 显式登记表（select_plan_actions / get_workbook_context / find_fields / read_range / run_python_analysis）；未登记工具一律拒绝下发，替代原先的任意 dict 直 `json.dumps` |
| 工具表条件注入 | `server/app/excel_agent.py` `model_tools()` | 总开关开启时在传给模型的列表追加；**不替换规范列表**（`current_tool_mode`/修复降级/只读批次降级的身份判断保持原样）；plan_only 模式不注入 |
| SDK 批准集 | `server/app/claude_sdk_agent.py` | `_APPROVED_NON_TERMINAL_TOOLS = {run_python_analysis}`；批准项经出站闸投影下发，其余 dict 结果维持"非终结工具结果不可用"拒绝；开关开启时注册进 MCP 工具表与 allowed_tools |
| 测试 | `server/tests/test_python_analysis.py` | 25 条：默认关闭、失败关闭（宿主执行入口哨兵）、源码级执行词扫描、入参校验、白名单投影、出站闸、普通后端 e2e（真实主循环 + 哨兵）、SDK 注册/回执/拒绝 |

### 反向变异证据（全部 RED(OK) 后还原复绿）

| 变异 | 拆掉的机制 | 抓到的测试 |
|---|---|---|
| M1 | 回执字段白名单改为原样放行 | 2（投影白名单 ×2） |
| M2a | SDK 批准集改为无条件拒绝 | 1（SDK 投影回执） |
| M2b | SDK 批准集改为无条件放行 | 1（SDK 拒绝行为保持） |
| M3 | 回执经白名单键 `resultHandle` 回显模型代码 | 3（失败关闭哨兵 + e2e + SDK） |
| M4 | 拆掉工具表注入的默认关闭闸 | 1（默认不供给） |

### 回归与环境

- `compileall -q server/app` 通过；`test_python_analysis` + `test_claude_sdk_agent` + `test_protocol_consistency` 106 passed；全量 `server/tests` 见当日验收记录。
- 前端未改动（本批未新增协议字段，`contracts.ts` 无需变更；前端不读 `pythonExecution` 段）。
- 隔离运行时探测（只读）：Docker 客户端在但守护进程未运行（WSL2 后端）；`wsl.exe` 被沙箱程序黑名单硬阻断（不可重试）；Hyper-V 管理服务未发现 → **本平台未验证任何隔离方案**。

### 截止本批仍未实现（不得写成"仅待实机验收"）

- 隔离执行器本体（容器/沙箱/Job Object 任一形态）、健康检查与资源限额 enforcement
- `datasetHandle` 登记与核对（dataResults → 任务内句柄）、`resultHandle` 结果存储与过期
- 分析结果 → `AnalysisPlan` 编译、双端入口 UI（Excel 当前工作簿 / WPS / openpyxl 文件夹）
- 即：`run_python_analysis` 当前对模型永远返回"隔离 Python 不可用"，不会运行任何代码

---

## 八、落地记录（2026-10-06 第二批：隐私/入参骨架修复）

源码核查发现三处真实缺陷并全部修复（先红后绿，测试与生产代码均为本批产出）。**本批只补隐私闭环与实际入参校验，不接隔离执行器、不安装环境**；`pythonExecution.enabled` 保持 `false`。

### 修复内容

| 缺陷 | 根因 | 修复 |
|---|---|---|
| ① Python 参数进模型历史与持久检查点 | `assistant_message = dict(message)` 完整携带 `code` 进入 `messages` → `save_planning_checkpoint` 整份落 SQLite `checkpoint_json` | 构建点单次投影 `_withhold_python_code_in_assistant_message`（`code`→固定占位，确定性字段级处理，tool_call_id/名称/其余参数不变，幂等）；同轮执行参数走 `_capture_raw_python_arguments` 受控内存（按 call id，pop 即弃）；恢复轮拿到即占位 → 执行时返回 `code_not_preserved` 固定回执并要求重新提供，不冒充续跑 |
| ② 回执自由文本直通 | `project_python_analysis_receipt` 只按 `RECEIPT_KEYS` 筛键、值任意放行 | 有类型回执：status/reasonCategory 封闭枚举 + 组合白名单 `_ALLOWED_RECEIPT_COMBINATIONS`；`message` 按可信状态组合再生固定文案（不读输入）；未登记字段、伪造 `runId`/`resultHandle` 一律 ValueError；投影失败在出站闸与主循环出口失败关闭为 `PYTHON_PROJECTION_FAILED_CONTENT`（不 `str(error)`，只对 `run_python_analysis` 生效，其他工具保持原异常路径） |
| ③ schema 声明未落实到入参校验 | `_validate_arguments` 接受任意额外字段 | 可信侧严格校验：未知键拒绝、code/datasetHandles 类型/空白/超限拒绝、句柄数量与单项长度限额进 `capabilities.pythonExecution`（`maxDatasetHandles=8`、`maxDatasetHandleChars=256`）；拒绝回执不回显入参；双后端（普通 + SDK）同一校验 |

出口审计结论（写入测试注释与本文档）：`_tool_step`/`_tool_argument_facts` 不覆盖 `run_python_analysis`（事件无参数泄露）；`_tool_cache_key` 仅限只读工具；logger 只记指标——事件/日志/缓存出口无参数路径。**SDK CLI 托管的内部转录无法投影**：服务端只能收口工具结果、header-only prompt 与检查点，`claude_sdk_agent.py` 如实记录该限制，不宣称双后端同等隐私闭环。形状合法 ≠ 授权合法：datasetHandle 登记核对仍未实现，字符串校验不冒充授权检查。

### 测试与反向变异

- 测试从 25 条扩到 **45 条**（`test_python_analysis.py` 重写）：三缺陷复现（每轮模型请求快照扫描、真实账本 SQLite 全列扫描、恢复轮 `code_not_preserved`）、回执再生/伪造拒绝/双出口失败关闭、真实入口额外字段拒绝 + 19 例参数化非法入参 + capabilities 边界 + SDK 同一校验。45/45 全绿。
- 4 项独立反向变异（脚本 `scripts/mutate_python_privacy_1006.mjs`，全部 RED(OK) 后还原复绿，基线先证 GREEN）：

| 变异 | 拆掉的机制 | 变红测试 |
|---|---|---|
| M1 | `_withhold_python_code_in_assistant_message` 改 no-op | 3：模型请求快照 / SQLite 持久化 / 恢复轮 |
| M2 | `project_python_analysis_receipt` 改原样放行 | 3：message 再生 / 未知组合 / 伪造字段 |
| M3 | `_validate_arguments` 未知键检查短路 | 3：额外字段拒绝 / 不回显 / SDK 同一校验（边界限额用例保持绿，判别点唯一） |
| M4 | `_non_terminal_tool_content` 异常出口保护拆除 | 3：出站闸 / 真实主循环 / SDK 出口（放行用例保持绿） |

### 回归（AGENTS.md 全量，2026-10-06）

| 命令 | 结果 |
|---|---|
| `npm run test:addin` | 776 passed |
| `npm run build:addin` | 通过 |
| `npm run test:wps` | 528 passed + **1 failed（既有，单列）** |
| `npm run build:wps` | 通过 |
| `pytest server/tests -q` | **844 passed** |
| `python -m compileall -q server/app` | 通过 |

既有失败单列：`wpsNativeComparison.test.ts` 1 条（显示值漂移用例期望写回循环守卫文案，实际触发更早的 `assertOriginalUnchanged` 守卫；用例判别文案选错，与本批无关且 git 无改动，定性见 [CAP-02 文档](CAP02_WPS_PLANNING_RESUME_2026-10-06.md)）。无环境阻断。

### 截止本批仍未实现（不得写成"仅待实机验收"）

- 隔离执行器本体、健康检查、资源限额 enforcement（当前只有**校验**限额，没有执行，也就没有 enforcement 对象）
- `datasetHandle` 登记与核对、`resultHandle` 结果存储与过期
- 分析结果 → `AnalysisPlan` 编译、双端入口 UI
- SDK 检查点恢复（当前接入未实现检查点恢复；跨后端续跑被 `ModelNotAvailableError` 挡住）
- 即：本批后工具对模型返回的仍是"未启用 / 不可用 / 未通过校验 / 代码未保存"类**固定回执**，永不运行任何代码

---

## 九、落地记录（2026-10-06 第三批：三条隐私绕过路径补修）

源码核查用合成数据复现了三条上一批未收口的绕过路径，全部先红后绿修复。**本批不接隔离执行器、不安装环境**；`pythonExecution.enabled` 保持 `false`。不再沿用"隐私闭环已完成"的汇报结论，只记录实际路径与证据。

### 缺陷①：被拒绝的参数仍进入消息与检查点

- **根因**：`withheld_python_arguments` 只替换顶层 `code` 键，其余参数（未知 `apiKey`/`path`/`command`、`datasetHandles` 内嵌对象、畸形 JSON）原样保留，即使入口判定 `invalid_request`，这些敏感原值仍进入下一轮模型请求与 SQLite 检查点。
- **修复**：改为可信侧最小白名单投影——解析失败/非对象只留 `{"code": marker}`；解析为对象时只保留 `code`（固定占位）+ 形状合法的 `datasetHandles`（`_project_safe_dataset_handles`，只做形状校验、不冒充授权），其余键全部丢弃。入参校验与历史投影是两道独立边界，被拒绝 ≠ 可以原样落盘/进历史。

### 缺陷②：旧检查点仍携带原始代码

- **根因**：恢复入口直接 `dict(item)` 复制 `resume_checkpoint.messages`、`dict(...)` 复制 `pendingAssistantMessage`，新投影只处理本轮构建的 `assistant_message`，旧历史里的 code 与 pending 里的原始 code 被当成本轮执行参数。
- **修复**：新增 `_project_python_code_in_checkpoint_messages`，在恢复入口对 `messages` 与 `pendingAssistantMessage` 统一投影（复用最小白名单）；区分"本轮新响应"与"恢复响应"两条路径；恢复轮拿到占位 → 协调器判 `code_not_preserved`，明确要求重新提供，不冒充断点续跑。不含 Python 的检查点恢复是恒等操作，不受影响。

### 缺陷③：执行入口异常仍原样回显

- **根因**：投影函数抛异常的路径已收口，但 `_execute_tool`/协调器抛异常时，普通后端 `except ... as error` 写 `state.last_error_message = str(error)` + `{"error": str(error)}`，SDK `handle` 写 `f"...{str(error)[:300]}"`。
- **修复**：新增 `execution_error_receipt()` / `python_error_response_content()` 统一安全回执（`reasonCategory=execution_error`，固定文案）；普通后端工具循环与 SDK `handle` 对 `run_python_analysis` 异常同源映射为固定安全内容，原始异常不进模型消息、`last_error_message`、事件 trace、SDK 工具结果。`asyncio.CancelledError` 继承 `BaseException`，不在常规 `except ValueError` 范围，继续传播不被吞。

### 测试与反向变异

- 测试从 45 条扩到 **56 条**（`test_python_analysis.py`）：缺陷①（额外字段/嵌套句柄/SQLite/多工具批次/畸形 JSON）、缺陷②（旧 messages/旧 pending/无 Python 恢复边界）、缺陷③（普通后端/SDK/取消传播）。全部先证明哨兵存在再走真实链路。
- 4 项独立反向变异（脚本 `scripts/mutate_python_privacy_oct06_third.mjs`，全部 RED(OK) 后还原复绿）：

| 变异 | 拆掉的机制 | 变红测试 |
|---|---|---|
| M1 | `withheld_python_arguments` 退回只删 code 键 | 4：额外字段 / 嵌套句柄 / SQLite / 多工具批次 |
| M2 | `_project_python_code_in_checkpoint_messages` 改恒等 | 2：旧 messages / 旧 pending |
| M3 | 普通后端执行异常安全映射分支短路 | 1：执行入口异常（loop） |
| M4 | SDK 执行异常安全映射分支短路 | 1：执行入口异常（SDK） |

### 回归（AGENTS.md 全量，2026-10-06 第三批）

| 命令 | 结果 |
|---|---|
| `npm run test:addin` | 776 passed |
| `npm run build:addin` | 通过 |
| `npm run test:wps` | 528 passed + **1 failed（既有，单列）** |
| `npm run build:wps` | 通过 |
| `pytest server/tests -q` | **855 passed** |
| `python -m compileall -q server/app` | 通过 |

既有失败单列同 §八：`wpsNativeComparison.test.ts` 1 条，与本批无关、git 无改动，未越界修。无环境阻断。

### 截止本批仍未实现（不得写成"仅待实机验收"）

- 隔离执行器本体、健康检查、资源限额 enforcement（当前只有**校验**限额）
- `datasetHandle` 登记与核对、`resultHandle` 结果存储与过期
- 分析结果 → `AnalysisPlan` 编译、双端入口 UI
- SDK 检查点恢复（当前接入未实现检查点恢复）
- SDK CLI 托管的内部转录仍无法投影，服务端只能收口工具结果/prompt/检查点——不宣称双后端同等隐私闭环
- 即：工具对模型返回的仍是"未启用 / 不可用 / 未通过校验 / 代码未保存 / 执行错误"类**固定回执**，永不运行任何代码

## 十、落地记录（2026-10-06 第四批：两处隐私遗漏收口）

源码核查用合成数据复现了两处上一批未收口的隐私遗漏，全部先红后绿修复。**本批只修这两项，不实作执行器、不启用 Python、不安装环境、不扩大工具权限**；`pythonExecution.enabled` 保持 `false`。

### 遗漏 A：未登记句柄仍是自由文本出口

- **根因**：`_project_safe_dataset_handles` 只做字符串类型/非空/长度形状校验后**保留原值**，于是"形状合法"的模型自报句柄（含长得像正常 `dataset://` 句柄的字符串）仍进入下一轮模型请求与真实 SQLite `checkpoint_json`。"形状合法"不能证明它是可信侧签发的句柄。
- **修复**：`_project_safe_dataset_handles` 改为**一律返回空列表**（不读取输入内容，也不做正则/前缀/UUID/长度猜测冒充授权），历史/检查点投影对 `datasetHandles` 的唯一合法形态是空列表——固定、不含原值的兼容表示，清楚表达「未验证」。入参校验（`_validate_arguments`）与历史投影仍是两道独立边界：非法请求仍拒绝、恢复轮仍判 `code_not_preserved`、本批新响应的受控内存参数通道不受影响。

### 遗漏 B：旧检查点错误文本仍会重放

- **根因**：`_project_python_code_in_checkpoint_messages` 只处理 assistant 工具调用参数，没清理旧 `run_python_analysis` 工具错误回执、`loopState.lastErrorMessage`、Python 相关 `lastWarning`；旧 code 虽已替换，但错误哨兵仍进入恢复后模型请求并再次保存到 SQLite。
- **修复**：恢复入口在内容进入模型请求/工具处理/事件/再次保存**之前**统一收口：
  - 对能确认属于 `run_python_analysis` 的旧工具回执（按 `name` 或 `tool_call_id` 关联，不匹配自然语言错误文案），用可信固定回执（`code_not_preserved`）**整体重建**，不读旧 `error`/`message`/`stdout`/`traceback`，也不只删某个敏感键；
  - 涉及 Python 的检查点同步清空 `loopState.lastErrorMessage`/`lastWarning` 自由文本字段（这些字段无 tool_call_id，无法逐条确认归属，fail-closed 清空，不猜测安全后重放）；
  - 无 `name`、且 `tool_call_id` 对不上任何 assistant 调用的自由文本回执 → 抛 `ValueError`（固定「需要重新生成预览」文案）失败关闭，拒绝恢复；
  - 不含 Python 的合法检查点维持既有恢复行为；不原地改动调用方传入的原检查点对象；非 Python 工具回执与 tool_call_id 配对保留。

### 测试与反向变异

- 测试从 56 条扩到 **67 条**（`test_python_analysis.py`）：未登记句柄（形状合法 `dataset://` 哨兵/模型出口/SQLite/被拒请求/恢复现场）、旧错误上下文（工具回执/lastErrorMessage/lastWarning/已完成/pending/混合工具/无 name 由 tool_call_id 确认/归属不明拒绝/不含 Python 不受影响）。全部先证明哨兵存在再走真实链路（真实 `FolderOperationLedger` register→checkpoint→pause→SQLite 原始行读回）。
- 4 项独立反向变异（每次只拆一个守卫，RED 后还原复绿）：

| 变异 | 拆掉的机制 | 变红测试 |
|---|---|---|
| M1 | `_project_safe_dataset_handles` 退回保留原值 | 5：未登记句柄（投影/模型请求/被拒/SQLite/恢复） |
| M2 | 跳过旧 Python 工具回执重建 | 3：旧错误上下文/混合工具/无 name |
| M3 | 跳过 loopState 错误/警告清空 | 1：旧错误上下文 SQLite 持久化 |
| M4 | 绕过归属不明恢复拒绝 | 1：归属不明拒绝 |

### 回归（2026-10-06 第四批）

| 命令 | 结果 |
|---|---|
| `npm run test:addin` | 776 passed |
| `npm run test:wps` | 528 passed + **1 failed（既有，单列）** |
| `pytest server/tests/test_python_analysis.py` | **67 passed** |
| `pytest`（agent/loop/planner/claude_sdk/main 定向） | 173 + 36 + 50 passed |
| `python -m compileall -q server/app` | 通过 |

**环境阻断（如实记录，不当作通过）**：本会话沙箱的 safe-delete 批量守卫会拦截 `test_folder_workbooks.py` 里的 `shutil.rmtree`/`unlink` 清理操作（报 `SAFE_DELETE_BULK_REJECTED`，指向 `E:\codex\excel-bro\.tmp-pytest`），导致该文件 14 条与目录清理相关的既有用例在本环境无法跑绿；这是沙箱对测试自身删除行为的拦截，与本批改动（`python_analysis.py`/`excel_agent.py`，不触碰 `folder_workbooks.py`）无关。全量 `pytest server/tests -q` 因此在本会话未能给出干净总计数（第三批在干净环境为 855 passed）。`npm run build:addin`/`build:wps` 已在上一轮通过，本批仅后端改动，未重跑前端构建。

既有失败单列同 §八：`wpsNativeComparison.test.ts` 1 条，与本批无关、git 无改动，未越界修。

### 截止本批仍未实现（不得写成"仅待实机验收"）

- 隔离执行器本体、健康检查、资源限额 enforcement（当前只有**校验**限额）
- `datasetHandle` 登记与核对（本批只收口"未登记句柄不保留原值"，**不新增完整句柄登记系统**）、`resultHandle` 结果存储与过期
- 分析结果 → `AnalysisPlan` 编译、双端入口 UI
- SDK 检查点恢复（当前接入未实现检查点恢复）
- SDK CLI 托管的内部转录仍无法投影，服务端只能收口工具结果/prompt/检查点——不宣称双后端同等隐私闭环
- 即：工具对模型返回的仍是"未启用 / 不可用 / 未通过校验 / 代码未保存 / 执行错误"类**固定回执**，永不运行任何代码

## 十一、落地记录（2026-10-06 第五批：两个恢复分支补修）

源码核查用合成数据复现了两个恢复分支漏洞，先红后绿修复。**本批只补这两项，不实作执行器、不启用 Python、不安装环境、不扩大权限**；`pythonExecution.enabled` 保持 `false`。

### 分支①：pending-only Python 的 loopState 漏清理

- **根因**：恢复入口只执行 `_scrub_python_loop_state(state, messages)`，其判定 `_checkpoint_involves_python` 只查 `messages`。当 Python 调用只存在于 `pendingAssistantMessage`（messages 只有 system/user）时，`lastErrorMessage`/`lastWarning` 未被清理，首次 `model_response_saved` 检查点就把它们写进真实 SQLite。
- **修复**：`_checkpoint_involves_python` 与 `_scrub_python_loop_state` 增加 `pending_assistant_message` 参数，判定同时考虑 messages（assistant 调用名 + tool 回执 name）与 pending（assistant 调用名）；恢复入口把 pending 现场读取提前到 scrub 之前，一并传入。不含 Python 的现场仍保留既有恢复语义，不一律清空所有工具的 `lastErrorMessage`/`lastWarning`；不原地改动检查点对象。

### 分支②：带 name 的孤立工具回执绕过归属检查

- **根因**：`_project_python_code_in_checkpoint_messages` 只拒绝「无 name 且 tool_call_id 对不上」的回执。回执填入 `name="get_workbook_context"` 但 `tool_call_id` 孤立时，既不是 `run_python_analysis`、`name` 又非空，于是敏感 content 被原样保留并送入模型/SQLite。
- **修复**：新增 `_build_tool_call_association`，一次构建 assistant 调用 ID→工具名 的可信关联（纳入 messages 与 pending，真正冲突即拒绝）。每个 `role=tool` 回执必须能关联到对应调用：孤立 ID 拒绝；有 `name` 时必须与调用名一致（不一致拒绝）；无 `name` 但 ID 可关联时由调用推导归属；归属为 Python 的回执固定重建，非 Python 回执保持内容与顺序；messages/pending 一致镜像（同 id 同 name）不误判冲突。恢复入口改为一次构建关联表后传给投影函数，替换原先「messages / pending 各自投影、各自构建关联」的两段式。

### 测试与反向变异

- 测试从 67 条扩到 **74 条**（`test_python_analysis.py`）：pending-only 持久化（每次 save_planning_checkpoint 后立即读 SQLite 原始行逐边界取证，含恢复后立即暂停路径）、带 name 孤立回执拒绝、拒绝不发起模型请求/不落盘、名称冲突拒绝、合法 Python 回执固定重建、messages/pending 一致镜像不误拒、真正关联冲突拒绝。
- 3 项独立反向变异（每次只拆一个判据，`scripts/mutate_python_resume_branches.mjs`，RED 后还原复绿）：

| 变异 | 拆掉的机制 | 变红测试 |
|---|---|---|
| M1 | `_checkpoint_involves_python` 忽略 pending | 1：pending-only loopState 持久化 |
| M2 | 名称非空就放行孤立回执 | 2：带 name 孤立回执拒绝 + 拒绝不落盘 |
| M3 | 跳过调用 ID 与工具名一致性检查 | 1：名称冲突拒绝 |

### 回归（2026-10-06 第五批）

| 命令 | 结果 |
|---|---|
| `npm run test:addin` | 776 passed |
| `npm run test:wps` | 528 passed + **1 failed（既有，单列）** |
| `pytest server/tests -q` | **873 passed** |
| `python -m compileall -q server/app` | 通过 |

既有失败单列同 §八：`wpsNativeComparison.test.ts` 1 条（显示值漂移用例期望写回循环守卫文案、实际触发更早的 `assertOriginalUnchanged` 守卫；用例判别文案选错，与本批无关、git 无改动，未越界修）。无环境阻断。

### 截止本批仍未实现（不得写成"仅待实机验收"）

- 隔离执行器本体、健康检查、资源限额 enforcement（当前只有**校验**限额）
- `datasetHandle` 登记与核对（本批只收口"未登记句柄投影为空列表"，**不新增完整句柄登记系统**）、`resultHandle` 结果存储与过期
- 分析结果 → `AnalysisPlan` 编译、双端入口 UI
- SDK 检查点恢复（当前接入未实现检查点恢复）
- SDK CLI 托管的内部转录仍无法投影，服务端只能收口工具结果/prompt/检查点——不宣称双后端同等隐私闭环
- 即：工具对模型返回的仍是"未启用 / 不可用 / 未通过校验 / 代码未保存 / 执行错误"类**固定回执**，永不运行任何代码
