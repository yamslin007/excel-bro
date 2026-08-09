# Excel Bro 响应慢问题分析（cc-check）

> 分析人：Claude Code
> 分析日期：2026-07-30
> 目的：定位任务窗格「响应非常慢」的根因，供后续（codex）继续排查与优化。

## 结论先行

**不是「思维链太长」的问题。** 这个项目并没有让模型做长链推理（CoT）。
真正的慢来自 **架构性的多次串行模型往返 + 每次都携带一个巨大的工具 schema + 全程非流式响应**。

## 一轮对话的实际模型调用链

用户点一次「发送」，后端会 **串行** 调用模型多次：

| 阶段 | 文件 | 调用次数 | 说明 |
| --- | --- | --- | --- |
| 意图判断 | `server/app/intent.py:158-170` | 1~2 次 | 结构校验失败会再修复 1 次；`max_tokens=1200` |
| Agent 规划循环 | `server/app/excel_agent.py:389-395` | 1~8 次 | `MAX_AGENT_TURNS=8`（来自 `config/capabilities.json` 的 `agent.maxTurns`），每轮一次往返 |

即一个写入操作，最坏需要 **3~10 次串行的模型请求**，每次都要等对方完整生成完才继续下一步。
客户端 `server/app/llm/client.py` 没有 `stream=True`，所以每次都是阻塞式等待整个响应。
用户端体感就是长时间「转圈」。

## 三个真正的瓶颈（按影响排序）

### 1. 每轮 Agent 都重发一个 ~10400 token 的工具 schema（最致命）

`server/app/excel_agent.py:227` 把整个 `AnalysisPlan` 的 JSON Schema 内联展开，塞进 `submit_plan` 工具定义。

实测数据：

```
AnalysisPlan 内联 schema = 36413 字符 ≈ 10400 token
```

原因：

- `ExcelAction` 是 **36 种动作** 的大联合类型（`server/app/models.py` 中约 36 个 `*Action` 类）。
- `excel_agent.py` 的 `_inline_json_schema()` 会把 Pydantic 的 `$defs` 引用 **全部展开且不去重**，进一步放大体积。
- 这个约 1 万 token 的 schema 在 Agent 循环里 **每一轮都完整重发一次**。8 轮 = 8 万+ input token 仅用于传 schema。
- 很多模型对超大 tool schema 的首 token 延迟明显更差；本地小模型（Ollama / LM Studio）尤其严重。

复现测量命令（在仓库根目录运行）：

```bash
python -c "
import sys, json
sys.path.insert(0, 'server')
from app.models import AnalysisPlan
from app.excel_agent import _inline_json_schema
s = _inline_json_schema(AnalysisPlan.model_json_schema())
txt = json.dumps(s, ensure_ascii=False)
print('AnalysisPlan 内联 schema 字符数:', len(txt))
print('粗略估算 token 数(~/3.5):', int(len(txt)/3.5))
"
```

### 2. Agent 调用没有设置 `max_tokens`

`server/app/excel_agent.py:391` 的 `chat_completions()` 调用没有传 `max_tokens`，输出无上限。
生成一个大的 `AnalysisPlan` JSON 时输出 token 多，而输出是逐 token 生成的，直接拉长每轮耗时。
（对比：`intent.py` 有 `max_tokens=1200`。）

### 3. 全程非流式，用户看不到任何进度

`server/app/llm/client.py:46-94` 用一次性 `response.json()`，没有 `stream=True`。
即使总耗时不变，多次串行 + 无流式，主观上就是「卡死很久没反应」。

## 如何用数据确认是哪一环最慢

项目自带诊断，别靠猜：

- 前端 `apps/excel-addin/src/diagnostics.ts` 已分阶段记录耗时：
  `intent_model` / `planning_model` / `local_query` / `folder_query` / `execution` / `verification` / `saved_tool`。
  设置页可导出 JSON（`exportDiagnosticReport()`）。
- 跑一次慢操作，导出诊断，看是 `intent_model` 慢，还是 `planning_model` 占大头（大概率是后者）。
- 后端 `/api/diagnostics` 记录本地 API 路径、状态码和耗时。
- 若使用本地网关（Ollama / LM Studio），本地小模型处理 1 万 token 的 schema 会特别慢，此时瓶颈是「模型算力 + schema 体积」双重叠加。

## 优化方向（改动量从小到大）

1. **给 Agent 加 `max_tokens`** — 一行改动，先止血。`excel_agent.py:391`。
2. **精简 `submit_plan` schema（收益最大）** — 不必把 36 种动作全展开成一个巨型 schema。可选方案：
   - 让模型先只声明 `kind`，再按需要二次获取该动作的详细 schema；
   - 按意图裁剪，只暴露与当前需求相关的动作子集；
   - `_inline_json_schema` 改为保留 `$defs` 引用（若目标模型支持 `$ref`），避免展开去重问题。
   - 目标：把 ~1 万 token 压到几百。
3. **开启流式（`stream=True`）** — 总耗时未必降，但体感大幅改善。需要 `client.py` 支持 SSE 解析。
4. **减少 Agent 轮数 / 简单查询短路** — 简单查询（`query_table` 已在前端确定性计算）其实用不着走完整 Agent 循环。

建议实施顺序：先做 1，再做 2，效果最明显。改动前先导出一次诊断 JSON 作为基线，改动后对比。

## 关键文件索引

- 模型客户端（非流式）：`server/app/llm/client.py`
- 意图判断（1~2 次调用）：`server/app/intent.py`
- Agent 循环（1~8 次调用 + 巨型 schema）：`server/app/excel_agent.py`
- 巨型 schema 来源（36 种动作）：`server/app/models.py`
- Agent 轮数上限配置：`config/capabilities.json`（`agent.maxTurns = 8`）
- 规划入口：`server/app/planner.py:1834`（`create_plan`）
- 前端分阶段诊断：`apps/excel-addin/src/diagnostics.ts`

## 改造进度（Claude Code，2026-07-30）

> **给 codex 的快速索引**：下表汇总本轮全部改进。三个主题——先降单次开销（#1/#2/D）、
> 再让失败不白等（A/B）、最后改体感与防中断（C/E）。详细说明见表下分批小节。
> 最终验证：后端 129 项、前端 24 项测试全绿；`npm run build` clean。

| 编号 | 改进 | 主题 | 关键文件 | 测试 | 状态 |
| --- | --- | --- | --- | --- | --- |
| #1 | Agent 加 `max_tokens=4000`（输出封顶） | 降单次开销 | `config/capabilities.json`(`agent.maxOutputTokens`)、`excel_agent.py` | `test_planner.py` | ✅ |
| #2 | 精简 `submit_plan` schema（10403→5764 token，-45%） | 降单次开销 | `excel_agent.py`(`_strip_metadata`/`_plan_tool_schema`) | `test_planner.py` ×2 | ✅ |
| A | 过载/限流指数退避重试（429/5xx/超时/网络抖动） | 失败不白等 | `config/capabilities.json`(`llm.maxRetries` 等)、`capabilities.py`(`capability_float`)、`llm/client.py`、`llm/errors.py`(`retry_after`) | `test_llm.py` ×4 | ✅ |
| B | Agent 只读工具结果缓存（跨重试存活） | 失败不白等 | `turn_state.py`(`TurnState.tool_cache`)、`excel_agent.py`、`planner.py`、`main.py` | `test_planner.py` ×2 | ✅ |
| C | SSE 流式分步进度（`/api/turn/stream`） | 改体感 | `main.py`(`_map_turn_error`/新端点)、`excel_agent.py`(`on_event`)、`planner.py`、`models.py`+`contracts.ts`(`TurnStepEvent`)、`api.ts`(`streamAssistantResponse`)、`App.tsx` | `test_main.py` ×4 + `api.test.ts` ×4 | ✅ |
| D | 数据已在手时砍掉 3 个只读工具（稳健版提速） | 降单次开销 | `excel_agent.py`(`_readonly_tools`/`_terminal_tools`/`AGENT_TOOLS_WITH_DATA`) | `test_planner.py` ×2 | ✅ |
| E | 模型超时纳入 `llm.timeoutSeconds`（可配置、防中断） | 防中断 | `config/capabilities.json`、`capabilities.py`(`model_timeout_seconds`)、`planner.py`、`intent.py` | `test_llm.py` ×1 | ✅ |

**回归命令**（本机 `%TEMP%` 有权限问题，必须指定 `--basetemp`）：
```bash
python -m pytest server/tests --basetemp=.tmp/pytest -p no:cacheprovider -q   # 后端 129 项
cd apps/excel-addin && npm run build && npx vitest run                        # 前端构建 + 79 项
```

架构文档同步位置：`docs/ARCHITECTURE.md` §4.2（流式与失败不浪费）、§6.1（模型调用分层）、§7（协议 `TurnStepEvent`）、§12（配置与端点）。

### 已完成

**1. 给 Agent 调用加 `max_tokens`（瓶颈 #2）**
- `config/capabilities.json` 的 `agent` 节点新增 `maxOutputTokens: 4000`（遵循项目「限额统一配置」原则）。
- `server/app/excel_agent.py` 读取 `MAX_OUTPUT_TOKENS` 并传入 `chat_completions`，Agent 输出不再无上限。

**2. 精简 `submit_plan` 工具 schema（瓶颈 #1，收益最大）**
- 新增 `_strip_metadata()`：剥离 Pydantic 自动生成的 `title`/`description` 注解。
  **关键坑**：`AnalysisPlan` 自身有名为 `title` 的业务字段，早期实现误删了它会导致必填校验失败、触发修复循环。已修正为「只剥离标量注解，进入 `properties`/`$defs` 容器时保留键名」。
- 新增 `_plan_tool_schema()`：在内联 schema 基础上移除 `acceptanceCriteria`（由 `AnalysisPlan.add_deterministic_acceptance_criteria` 验证器自动推断；模型即便仍主动提交，`default_factory` 也照常接受）。
- **实测：submit_plan schema 从 10403 token → 5764 token，降 ~45%。** 动作字段级结构完整保留，判别式 (`type`) 结构不变。
- 回归测试：`server/tests/test_planner.py` 新增
  `test_submit_plan_schema_is_compacted_but_preserves_action_fields` 和
  `test_plan_without_acceptance_criteria_still_validates`。

**验证**：`python -m pytest server/tests` 全部 116 项通过（注意本机 `%TEMP%` 目录有权限问题，需用 `--basetemp=.tmp/pytest` 指定可写临时目录）。

### 已完成（三件套改造，2026-07-30 续）

主题：「不要糟蹋用户已经等掉的时间」。实施顺序 #A → #B → #C。

**A. 过载/限流自动重试（止血）— 呼应瓶颈 #3 的失败浪费**
- 用户实测撞到上游 `HTTP 429 engine_overloaded_error`。`errors.py` 早已把 429/5xx 标 `retryable=true`，但**没有任何真正的重试逻辑**，一失败就废掉整轮。
- `config/capabilities.json` 新增 `llm` 节点（`maxRetries: 2` / `retryBaseDelaySeconds: 1.0` / `retryMaxDelaySeconds: 20.0`），遵循「限额统一配置」原则。
- `server/app/capabilities.py` 新增 `capability_float`；`client.py` 用 `_capability_nonneg_int`（允许 0 以关闭重试）。
- `server/app/llm/client.py`：`chat_completions` 外层加指数退避重试（`retry_base_delay * 2**attempt` + 抖动，封顶 `retry_max_delay`）；`LLMHTTPStatusError` 按 `.retryable` 判定（429/5xx），Timeout/Connect/Transport 也重试，`LLMResponseError`（解析失败）和 4xx 非 429（如 401）**不重试**。响应带 `Retry-After` 数值头时优先采用（`errors.py` 的 `LLMHTTPStatusError` 新增 `retry_after` 字段）。
- 测试（`server/tests/test_llm.py`）：两次 429 后成功、超上限抛原异常、401 不重试、`Retry-After` 控制延时。

**B. Agent 工具结果缓存（省功）— 消除「失败后重头开始」**
- 根因两层：`run_excel_agent` 的 `messages` 是局部变量，异常即销毁；`record_completion` 只在整轮成功后写缓存，失败重试从第 1 轮起。
- `get_workbook_context` / `find_fields` / `read_range` 是对同一份 `PlanRequest` 快照的**确定性只读计算**，天然可缓存。
- `server/app/turn_state.py`：`TurnState` 新增 `tool_cache` 字段（按 `turn_id` 存活，受现有 TTL/`maxActive` 约束）。
- `excel_agent.py`：只读工具结果按 `(name, sorted-args)` 缓存复用；`create_plan`（`planner.py`）透传 `tool_cache`，`main.py` 在 `completion_lock` 内取 `state.tool_cache` 传入。
- **注**：不是「断点续跑」（不保存 `messages`），模型仍从第 1 轮重新决策，但每步查数据近乎瞬时返回。语义安全：缓存的是事实读取，不是模型判断。
- 测试（`server/tests/test_planner.py`）：同 `turn_id` 两次调用第二次命中缓存（spy 断言底层不重复执行）、不同 arguments 不串缓存。

**C. 流式分步进度（体感）— 瓶颈 #3 的最终落地**
- 采用 **SSE 分步事件**（非 token 级流式，避开模型网关流式兼容风险）。
- 事件协议：`step`（对齐 `advanceActivity(title, detail, completedStep?)`）/ `result`（完整 `AssistantResponse`）/ `error`（HTTP 已 200，模型错误序列化为流内事件，镜像 `service_error` 的 `{code,message,retryable}` + `status`）。
- 后端：`main.py` 抽 `_map_turn_error` 供 JSON/流式端点共用；新增 `POST /api/turn/stream`（仅接受 plan 请求），用 `asyncio.Queue` + `create_task(create_plan(..., on_event=...))` 边跑边产 step，缓存命中直接产 result。`excel_agent.run_excel_agent` 加 `on_event` 回调（默认 None → 行为不变）。`models.py` + `contracts.ts` 双端新增 `TurnStepEvent`。
- 前端：`api.ts` 新增 `streamAssistantResponse(request, {onStep})`（`fetch` + `getReader()` + `TextDecoder` 按 `\n\n` 切帧，跨帧边界安全）；旧后端 404/405 或网络错误在收到事件前**降级**回退 `createAssistantResponse`。`App.tsx` 改用流式，`onStep → advanceActivity`。
- 测试：后端流式产 step+result、error 事件（HTTP 仍 200）、缓存路径不发 step、拒绝 intent 请求；前端 `api.test.ts` 分片解析、`onStep` 调用、`result` 返回、`error` 抛 `ApiRequestError`、404 降级。

**验证**：后端 `python -m pytest server/tests --basetemp=.tmp/pytest` 全部 126 项通过；前端 `npm run build`（tsc + vite）clean，`api.test.ts` + `contracts.test.ts` 共 24 项通过。

### 已完成（真实耗时优化，2026-07-30 续）

主题：直接缩短单次真实耗时，降低长耗时被超时中断的概率。

**D. 数据已在手时砍掉只读工具（稳健版提速）**
- 纯查询链路的浪费：意图阶段返回 `tool_request(query_table)` → 前端确定性算出结果塞进 `PlanRequest.dataResults` → 但 `run_excel_agent` 仍跑完整 Agent 循环，每轮携带 ~5764 token 的 `submit_plan` schema **和 3 个用不上的只读工具**（`get_workbook_context`/`find_fields`/`read_range`），最坏 8 轮。
- `excel_agent.py`：把 `_tools()` 拆成 `_readonly_tools()` + `_terminal_tools()`，产出两个常量 `AGENT_TOOLS`（完整）与 `AGENT_TOOLS_WITH_DATA`（仅 `submit_answer`+`submit_plan`）。`run_excel_agent` 在 `request.dataResults` 非空时选后者。
- 效果：数据已在手时，每轮不再携带只读工具定义，模型也不会多绕几轮重复读取（系统提示词早已要求「优先直接使用 dataResults」，现在从工具层面强制）。写入意图仍保留 `submit_plan`，稳健无损。
- 测试（`test_planner.py`）：`test_agent_drops_readonly_tools_when_data_results_present`（断言 tools 不含 3 个只读工具、含 submit_answer/submit_plan）、`test_agent_keeps_readonly_tools_without_data_results`（无 dataResults 时 5 个工具齐全）。

**E. 模型超时纳入统一配置（防中断安全网）**
- `AI_TIMEOUT_SECONDS=60` 原先硬编码在 `planner.py` 和 `intent.py`。
- `config/capabilities.json` 的 `llm` 节点新增 `timeoutSeconds: 60.0`；`capabilities.py` 新增 `model_timeout_seconds()`（env `AI_TIMEOUT_SECONDS` 优先且必须为正数，否则回退配置默认，非法值不中断请求）。`planner.py`/`intent.py` 改为调用它（意图阶段仍夹在 10~60 秒）。
- 测试（`test_llm.py`）：`test_model_timeout_prefers_env_override_then_config`（覆盖 env 优先、非法/非正值回退）。

**验证**：后端 `python -m pytest server/tests --basetemp=.tmp/pytest` 全部 129 项通过。

### 待办（后续可选，收益递减）

**激进版纯回答短路（本次未做，稳健版已覆盖大部分收益）**
- 检测到纯查询（无写入意图）时，首轮只给 `submit_answer`（连 `submit_plan` 大 schema 也去掉），一轮出答案；模型不肯回答再回退完整工具集。提速更大但边缘场景多一次回退往返。

**进一步压缩 schema**
- 若目标模型支持 `$ref`，可保留 `$defs` 引用避免动作间重复定义（当前 `_inline_json_schema` 为兼容严格模型而全展开）。
- 或两阶段：模型先声明需要的动作类型，再按需下发该动作子 schema。此方案加往返和出错率，谨慎评估。

---

## 安装包卸载器 Bug（2026-07-30 实机排查记录）

### 问题现象

在本机（计算机名 `AC`）卸载后，以下残留持续存在，导致 Excel 重启后仍尝试加载插件（日志可见 `GET /index.html` 404）：

- `ExcelBro.exe` 进程和 8765 端口仍在监听
- `%LOCALAPPDATA%\Programs\Excel Bro` 安装目录未删
- `%LOCALAPPDATA%\Excel Bro` 数据目录未删
- `HKCU\...\Wef\TrustedCatalogs\{41f62f5c...}` 仍指向 `\\AC\ExcelBroAddins`
- SMB 共享 `ExcelBroAddins` 仍存在
- 两张 `CN=Excel Bro localhost` 根证书仍在用户证书存储

### 根因

**`packaging/uninstall.ps1:48` 的设计缺陷**：删除 SMB 共享失败时直接 `throw`，导致脚本中断，后续的注册表清理、证书删除、进程终止、目录删除全部跳过。

```powershell
# 现状：UAC 弹窗取消 → $shareProcess.ExitCode -ne 0 → throw → 后续清理全跳过
if ($shareProcess.ExitCode -ne 0) {
    throw "无法移除 Excel Bro 本机加载项目录。"
}
```

### 另一个发现：两套加载并存

安装包（TrustedCatalogs）和开发旁加载（`npm run start:excel`）会**同时写入不同的注册表位置**：

| 来源 | 注册表位置 |
|---|---|
| 安装包 | `HKCU\...\Wef\TrustedCatalogs\{41f62f5c...}` |
| `npm run start:excel` | `HKCU\...\Wef\Developer\9c758d40...`（子键，含 `UseDirectDebugger` 等值） |

安装包的卸载器只清 TrustedCatalogs，**不清 Developer 子键**，导致开发旁加载残留。

### 修复建议（`packaging/uninstall.ps1`）

1. **把 `throw` 改为警告继续**，或把「删 SMB 共享」挪到最后，确保本地清理（进程/注册表/证书/目录）先执行完。
2. **补充清理 Developer 子键**：
   ```powershell
   Remove-Item -LiteralPath "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer\$manifestId" `
       -Recurse -Force -ErrorAction SilentlyContinue
   ```
3. **可选**：卸载失败时提示用户重试（因为 UAC 取消是最常见原因），而不是留下半清理状态。

### 关键文件

- 卸载器：`packaging/uninstall.ps1:30-49`（SMB 删除 + throw）
- 安装器：`packaging/installer.py:232`（删除旧 Developer 值，但只删 value 不删 subkey）

---

## 前端体验改造：过程透明 + 运行中打断/转向（Claude Code，2026-08-02）

> 主题：让用户「看得清模型在想什么」并「随时能插手」。两块都是**纯前端**改动，
> 不动后端协议。最终验证：前端 `tsc --noEmit` clean、`npx vitest run` 100 项全绿。

| 编号 | 改进 | 主题 | 关键文件 | 状态 |
| --- | --- | --- | --- | --- |
| F | 需求确认结果分层展示（模型的理解 + 锁定指令，可展开原始明细） | 过程透明 | `App.tsx`(`describeIntentDecision`/`describeQueryArguments`/`ActivityStep`)、`styles.css` | ✅ |
| G | 运行中打断 + 硬转向（带新话重跑），模型请求全程可 abort | 随时打断 | `api.ts`(3 个模型请求加 `signal`)、`App.tsx`(`turnAbortRef`/`steerTurn`/`stopTurn`/`isAbortError`)、`styles.css` | ✅ |

**回归命令**：
```bash
cd apps/excel-addin && npx tsc --noEmit && npx vitest run   # 前端类型检查 + 100 项
```

### F. 需求确认结果分层展示（过程透明）

- 背景：用户反馈「执行过程能看到的太少，无法判断大模型是否偏离轨迹」。选定方案 =
  **只用现有后端数据 + 分层展示**（不改 `/api/turn`）。
- `IntentCheckResponse` 里本就返回 `summary`（模型如何理解需求）、`confirmedPrompt`
  （锁定要做什么）、以及 clarification/tool_request 的结构化信息，此前**没有surfaced**。
- `ActivityStep` 新增可选 `note`（第一层，默认可见：模型的真实理解）和 `detail`
  （第二层，收在「查看详情」`<details>` 里：锁定指令、追问选项、本地查询参数翻译）。
  `advanceActivity` 增加第 4 参 `stepInsight?: {note, detail}` 透传。
- 新增 `describeIntentDecision(intent)`：按 proceed / clarification / tool_request 三分支
  产出 `{label, note, detail}`；`describeQueryArguments(args)` 把 `QueryTableArguments`
  翻成人话（方式/字段/分组/指标/筛选/合并/上限）。
- 注入点：`checkIntent` 返回后、`continueIntentDecision` 前各调一次 `advanceActivity`
  （`sendMessage` 主路径 + `resolveClarification` 澄清路径两处）。
- 渲染：实时活动卡与持久化 `activityLog` 两处步骤列表都改成
  `label + note + 查看详情`；`styles.css` 加 `.activity-step-body/-note/-detail`。
- **边界**：只 surface 模型的**结论**（理解/锁定），不是分步 CoT。要展示真实推理链需后端
  在 `/api/turn` 返回，本轮未做。

### G. 运行中打断 + 硬转向（随时打断）

- 背景：用户羡慕 Claude Code「随时可打断」。诊断出真正打不断的原因 =
  **`checkIntent` 是裸 `fetch` 没有 `signal`**，而大部分等待时间花在模型请求上。
  旧的 `queryAbortRef` 只包住本地表格扫描一小段，作用域太窄。
- 正确入手点是 **turn 边界**，不是那个旧按钮。新增 `turnAbortRef`：每个 turn 开始
  （`sendMessage` 模型阶段 / `resolveClarification`）新建 `AbortController` 存入，
  `finally` 里只在「仍是自己这个」时清空（被新 turn 覆盖则不动）。
- `api.ts`：`checkIntent`、`createAssistantResponse`、`streamAssistantResponse`
  （`StreamHooks.signal`）三个模型请求都加 `signal` 透传给 `fetch`；流式回退也带上。
- `App.tsx`：`turnAbortRef.signal` 接到两处 `checkIntent`、自纠重试的 `checkIntent`、
  `streamAssistantResponse`、以及本地取数（`executeQueryTableTool` 复用 turn controller）。
- 新增 `isAbortError(reason)`：识别 `DOMException(AbortError)` 与
  `DataToolExecutionError(code=CANCELLED)`。所有相关 catch 在打断时**静默收尾**、不弹红字，
  且 `executeRequestedDataTool` 打断时**跳过自纠重试**。
- 交互（**不用 Esc**，因 Office iframe 里 Esc 不可靠且语义是「关闭」）：
  - **硬转向**：运行中输入框打字 + Enter → `steerTurn(text)` 暂存新话到
    `pendingSteerRef`、abort 当前 turn；一个 `useEffect([busy])` 在 busy 回落后带新话
    重跑一个 turn（= abort + 带增补上下文重发，对齐 Claude Code）。
  - **纯停止**：输入框留空 → `stopTurn()` 只 abort、不重跑。
  - 发送按钮 busy 时变形：有字=`↑`(加脉冲，转向)、无字=`■`(停止)；活动卡「取消本地查询」
    升级为 turn 级「停止」，busy 全程可见。文案随 busy 切换。
- **边界**：只覆盖**模型请求 + 本地读数**阶段（用户痛点所在）。**Excel 写入执行阶段未接打断**
  ——写到一半中断涉及回滚，风险高，暂缓；写入本有「先预览再确认」兜底。若要执行阶段也可中途停，
  需单独设计回滚逻辑。

### 关键文件（F/G）

- 分层展示与打断主体：`apps/excel-addin/src/App.tsx`
- 模型请求 `signal` 透传：`apps/excel-addin/src/api.ts`
- 按钮/步骤样式：`apps/excel-addin/src/styles.css`
- 意图返回结构（surface 的数据来源）：`apps/excel-addin/src/contracts.ts`（`IntentCheckResponse`）
