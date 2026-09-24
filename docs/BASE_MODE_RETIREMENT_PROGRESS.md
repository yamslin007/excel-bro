# 基础模式退出产品入口 — 实施进度记录

> 依据：`docs/BASE_MODE_RETIREMENT_GUIDE.md`
> 记录时间：2026-09-24 18:40（GMT+8）
> 项目根：`E:\codex\excel-bro`

---

## 〇、一句话结论

**阶段 A / B / C 全部完成。** 六条验证命令全部通过（后端 444 项、Excel 464 项、WPS 38 项测试全绿，两端 `tsc` 与 `vite build` 通过，`compileall` 无错误）。

**唯一未完成的是真实宿主点击验收** —— 这无法由自动化替代，清单见第 5.5 节。

---

## 一、目标回顾

从 Excel 和 WPS 的**模型列表、默认选择、首次引导**中移除"基础模式"，改为完全由用户配置外部模型：

- 通过服务地址 + 模型 ID + API Key 配置外部模型
- 未配置时显示"添加模型"，并**保留用户已输入的任务**
- 模型调用失败时**明确反馈**，不静默降级、不擅自切换供应商

### 必须保留（红线）

| 保留项 | 状态 |
| --- | --- |
| 确定性读取 / 筛选 / 统计 / 计算 / 已固化工具 | ✅ 已保留（`_local_address_plan`、`_local_data_tool_response`、`_local_bulk_worksheet_deletion_plan`） |
| 白名单 `AnalysisPlan`、执行前预览、用户确认 | ✅ 未触碰 |
| 工作表授权边界、备份、撤销、checkpoint、幂等 | ✅ 未触碰 |
| 历史 `provider=local` 结果的正常展示 | ✅ 未全局替换（`contracts.ts` 的 `provider: "model" \| "local" \| "system"` 保持原样；消息作者标签由"基础模式"改为"本地工具"） |
| WPS 现有只读边界 | ✅ 未触碰（`wpsWrite.enabled` 仍为 `false`） |

### 架构红线（EB 后台）

| 红线 | 状态 |
| --- | --- |
| 后台承担配置 / 凭据管理 / 请求转发 / 任务协调，不是本地 AI | ✅ 未改为 WebView 直连供应商 |
| 不删除后台 | ✅ 未删除 |
| API Key 不写入浏览器持久存储 / 日志 / 聊天记录 | ✅ 未新增任何前端持久化 |

---

## 二、阶段总览

| 阶段 | 内容 | 状态 |
| --- | --- | --- |
| **A** | 状态与兼容基础：模型目录四态、选择优先级、旧 `local` 迁移、无模型协议、稳定错误码 | ✅ **已完成** |
| **B** | 清理 Excel / WPS UI 中的基础模式入口 | ✅ **已完成** |
| **C** | 更新 README 与 docs；跑六条全量验证命令 | ✅ **已完成** |

---

## 三、阶段 A 已完成明细

### 3.1 后端协议与错误语义

| 文件 | 改动 |
| --- | --- |
| `server/app/llm/config.py` | 新增异常 `ModelNotConfiguredError` / `ModelNotAvailableError`；新增 `LEGACY_LOCAL_MODEL_ID = "local"`；`selected_model_config()` 由"返回 `ModelConnection \| None`（None=基础模式）"改为"**返回 `ModelConnection`，不可用时抛稳定错误**"；另立 `resolve_model_config()` 保留"可返回 None"的查询语义（供健康检查、公式回退使用）；`model_catalog()` 删除 `local` 占位项、无真实连接时 `defaultModelId = None` / `models = []`；`model_status()` 无配置时返回 `mode: "model"`（原 `"local"`）；`set_formula_model_id()` 把 `local` 归一化为空串 |
| `server/app/llm/__init__.py` | 导出新增 `LEGACY_LOCAL_MODEL_ID`、`ModelNotConfiguredError`、`ModelNotAvailableError`、`resolve_model_config` |
| `server/app/planner.py` | 删除 `_local_analysis()`（那个返回"基础模式不支持此操作"的兜底）；`_local_address_plan()` 提到模型解析**之前**（关键：让已固化的确定性地址命令不依赖模型）；其余请求改由 `selected_model_config()` 解析，缺少连接时抛 `ModelNotConfiguredError`，指定失效连接/旧 `local` 时抛 `ModelNotAvailableError` |
| `server/app/intent.py` | `check_intent()` 在 `config is None` 时不再静默返回 `IntentProceedResponse`，改为抛 `ModelNotConfiguredError` |
| `server/app/main.py` | `_map_turn_error()` 新增 `MODEL_NOT_CONFIGURED`(409) / `MODEL_NOT_AVAILABLE`(409) 两个稳定错误码；`/api/intent` 端点接入同一套 `_map_turn_error`（此前只走泛化的 `INVALID_REQUEST` 422）；流式 `_planning_task_stream`、续跑 `continue/stream` 接入同一语义 |
| `server/app/models.py` | **协议收紧反向修正**：`IntentCheckRequest.modelId` 与 `PlanRequest.modelId` 的 `min_length=1` 放开为"缺省/空串 = 本次未选择模型"。理由：删掉 `local` 后，"未配置模型"成了一个**合法状态**，若仍要求 `min_length=1`，这个状态根本无法通过 FastAPI 校验，会被拦成 422 而非给出 `MODEL_NOT_CONFIGURED` |

### 3.2 前端状态与选择

| 文件 | 改动 |
| --- | --- |
| `apps/excel-addin/src/api.ts` | `ServiceHealth.mode` 收紧为 `"model"`；`ModelOption.provider` 收紧为 `"model"`；`ModelCatalog.defaultModelId` 改为 `string \| null`；新增 `ModelCatalogStatus = "idle" \| "loading" \| "ready" \| "failed"` |
| `apps/excel-addin/src/modelSelection.ts` | 重写：新增 `LEGACY_LOCAL_MODEL_ID`、`NO_MODEL_SELECTED`、`isRealModelId()`、`chooseAvailableModel()`（5 级优先级 + 过滤 `local`/不可用，无可用返回 `""`）、`shouldPersistSelection()`、`resolveDisplayedModelId()` |
| `apps/excel-addin/src/hooks/useServiceHealth.ts` | 初始 `modelOptions` 改为 `[]`（不再注入虚拟 `local`）；新增 `modelCatalogStatus` / `modelCatalogError` / `defaultCatalogModelId` state；新增 `loadModelCatalog()` / `refreshServiceState()` / `retryModelCatalog()`；**目录读取失败时不清空 `modelOptions`**，只置 `failed` + error |
| `apps/excel-addin/src/hooks/useModelManagement.ts` | `openConnectionCreator()` 不再依赖 `openSettings()` 成功（**表单始终可打开**）；`refreshModelsAfterSettings()` 不再 `?? "local"`；新增 `migrateModelSelection()` |
| `apps/excel-addin/src/helpCommand.ts` | 改为 `MODEL_REQUIRED_HELP_TEXT`（含"添加模型"/"服务地址"/"API Key"/确定性命令/执行预览，**不含**"基础模式"/"离线"/"零延迟"/"本地 AI"）；`BASE_MODE_HELP_TEXT` 仅留作 deprecated 别名 |

### 3.3 关键设计决策

1. **"目录失败 ≠ 没有模型"**：`loadModelCatalog` 失败时**不清空**已有列表，只置 `failed` + error，并提供 `retryModelCatalog()`。否则后台临时不可用会被误判成"用户从没配过"。
2. **界面显示 = 请求发送**：收敛到单一函数 `chooseAvailableModel()`（Excel/WPS 共用），删掉 `?? modelOptions[0]` 这类"回退到列表首项"的展示逻辑；App.tsx 中所有 `modelId:` 请求字段统一改用解析后的 `effectiveModelId`（不再直接读原始存储值）。
3. **确定性能力与模型解耦**：`_local_address_plan()` 排在模型解析之前 —— 没配置模型时，"清空 A1:B2""在 A5 填 4"依然可用。这既满足"保留已固化工具"，也避免它们被当成"基础模式回归"。
4. **`local` 的三种处理分层**：产品入口层清掉；能力层保留；协议层（`AnswerResponse.provider`、历史结果）保留。

---

## 四、阶段 B 已完成明细（Excel + WPS UI）

### 4.1 Excel（`apps/excel-addin/src/App.tsx`、`styles.css`）

- `effectiveModelId` / `selectedModel` 派生值上移到 `slashModelCommands` 之前（原本 `useMemo` 引用下方才声明的变量，是 9 个类型错误里的 2 个）。
- 删除 `handleSelectModel` 中的 `local` 分支与 `BASE_MODE_HELP_SHOWN_KEY`；改为 `isRealModelId()` 守卫，非法 id 不再写回选择存储。
- `/help` 文案改用 `MODEL_REQUIRED_HELP_TEXT`；斜杠命令描述由"查看基础模式支持的命令（离线功能说明）"改为"查看使用说明与模型配置步骤"。
- 顶部状态栏的 `isBaseMode` 图标与帮助按钮改为 `modelNotConfigured` 驱动；文案"查看基础模式帮助"→"查看模型配置说明"。
- 模型菜单删除 `option.id === "local"` 的图标/描述分支；勾选状态改用 `effectiveModelId`（原为 `selectedModelId || "local"`）。
- 消息作者标签对 `provider === "local"` 由"基础模式"改为"本地工具"。
- 连接卡片"当前使用"标记、断点恢复文案与备用模型回退判定统一改用 `effectiveModelId`。
- CSS：`base-mode-icon` → `model-status-icon`，`base-mode-help-button` → `model-help-button`；删除 `.brand-copy span.base-mode`、`.model-option-copy small`（基础模式描述）等失效规则。

### 4.2 WPS（`apps/wps-addin/src/`）

| 文件 | 改动 |
| --- | --- |
| `WpsAgentPanel.tsx` | 初始模型列表由"注入 `local`+`基础模式`"改为**空数组**；新增 `catalogStatus` / `catalogError`，四态分别表达（正在读取 / 无法读取+重试 / 尚未配置+添加模型 / `模型：<名称>`）；`effectiveModelId` 改用共用的 `chooseAvailableModel()`，删除 `?? models[0]` 回退与 `selectedModelId \|\| defaultModelId`；新增迁移 effect（仅 `ready` 时执行）；提交前守卫：无可用模型时提示去「添加模型」并**保留输入框文本**、不发起查询；模型菜单增加加载中/失败/未配置说明与「重试读取模型配置」 |
| `WpsModelSettingsPanel.tsx` | 公式模型下拉删除 `models.filter((model) => model.id !== "local")` 黑名单过滤（源头已无 `local`） |
| `styles.css` | 新增 `.eb-model-menu-note`、`.eb-model-retry` |

---

## 五、阶段 C 已完成明细（文档 + 验证命令）

### 5.1 文档更新

| 文件 | 改动 |
| --- | --- |
| `README.md` | 5 处："当前闭环"第 4 步、"没有配置模型时"段落、模型规划段落结尾、新安装引导段落、多连接段落结尾；统一改为"未配置时引导添加模型 / 不退回规则模式 / 显示与请求一致" |
| `docs/ARCHITECTURE.md` | 新增「模型配置与选择语义」一节（目录四态、选择优先级、稳定错误码、确定性命令与模型解耦、`provider=local` 语义）；`planner.py` 行注由"基础模式与最终规划"改为"确定性地址命令与最终规划" |
| `docs/DEVELOPMENT.md` | 模型配置章节由"未配置时只显示基础模式和首次添加引导"改为四态说明 + 显示/请求一致性说明；"从哪里开始改"表格与测试清单中的"基础模式"改为"确定性地址命令" |
| `docs/COMMAND_CATALOG.md` | 引语与背景段落中的"基础模式"改为"确定性地址命令" |

**未改动（属于历史记录，不应回改）**：`Issue.md`、`计划表.md`、`docs/WORK_LOG_2026-08-18.md`、`docs/HELP_COMMAND_IMPLEMENTATION.md`、`docs/MESSAGE_AUTHOR_ENHANCEMENT.md`、`docs/MESSAGE_TEXT_SIZE_ENHANCEMENT.md`、`docs/DESIGN_REVIEW_2026-08-16.md`，以及两份任务书 `BASE_MODE_RETIREMENT_GUIDE.md` / `BASE_MODE_RETIREMENT_PROMPTS.md`。

### 5.2 六条验证命令（2026-09-24 18:25–18:33 实测）

| 命令 | 结果 |
| --- | --- |
| `npm run test:addin` | ✅ 46 files / **464 passed** |
| `npm run build:addin` | ✅ 通过（`tsc --noEmit` + 两次 vite build） |
| `npm run test:wps` | ✅ 6 files / **38 passed** |
| `npm run build:wps` | ✅ 通过 |
| `python -m pytest server/tests -q` | ✅ **444 passed**（73.6s） |
| `python -m compileall -q server/app` | ✅ 无输出（exit 0） |

### 5.3 新增的行为测试（非源码字符串匹配）

**后端 `server/tests/test_planner.py`**

- 语义请求无模型 → 抛 `ModelNotConfiguredError`，文案含"添加模型/服务地址/API Key"
- 一条连接都没有 → `ModelNotConfiguredError`；指定已删除连接 → `ModelNotAvailableError`
- 已配置模型时携带旧 `local` → `ModelNotAvailableError`（刻意不装 mock transport：若真走到模型调用会变成 `LLMConnectError` 而失败）
- 旧 `local` 且无任何配置 → 等同未配置
- 确定性地址命令零模型可用；**已配置模型时仍走本地快路径**（用"一旦被调用即断言失败"的 `explode` 桩替换 `selected_model_config`，证明确实不消耗模型解析）
- 清空/写值/写公式三条字面命令的产出计划断言

**前端 `apps/wps-addin/src/WpsAgentPanel.test.tsx`**（+5 项）

- 目录加载成功但为空 → 显示"尚未配置模型" + "添加模型"，且不含"基础模式"
- 无模型提交任务 → 出现引导文案，**输入框仍保留原文本**，且 `runWpsReadonlyTurn` 未被调用
- 存储为旧 `local` + 目录两条真实连接 → 请求 `modelId` 等于解析出的默认连接（不是 `local`，也不是列表首项），且旧记录被写回为新连接
- 目录读取失败 → 显示"无法读取模型配置"、**不叠加**"尚未配置模型"；点"重试"后恢复为真实连接，`listModels` 被调用两次
- 后台不可读时点"＋ 添加" → 仍进入可填写表单，"添加并使用"可点，并显示"仍可填写配置"

**`apps/excel-addin/src/modelSelection.test.ts`**（14 项，已有）：5 级优先级、空选择、**绝不解析到 `local`**、跳过不可用、`isRealModelId`、`shouldPersistSelection`、`resolveDisplayedModelId`。

**`apps/excel-addin/src/helpCommand.test.ts`**（4 项，已有）：含模型配置引导、保留确定性命令、**不含**"基础模式/离线/零延迟/本地 AI"。

### 5.4 沙箱适配（复现用）

```powershell
# 后端：系统 TEMP 被拒，必须把 basetemp 指到工作区内
$env:TEMP="E:\codex\excel-bro\.tmp-pytest"; $env:TMP=$env:TEMP
.\.venv\Scripts\python.exe -m pytest server/tests -q `
  --basetemp=E:\codex\excel-bro\.tmp-pytest\bt -p no:cacheprovider

# Excel 前端（workspaces 提升到根 node_modules）
cd apps\excel-addin; node ..\..\node_modules\vitest\vitest.mjs run

# WPS 前端（本工作区自带 node_modules）
cd apps\wps-addin; .\node_modules\.bin\vitest.cmd run
```

---

## 六、未完成项

### 6.1 ⛔→✅ 原阻塞项已解除

上一版记录的 11 个类型错误（App.tsx 9 个、WpsAgentPanel.tsx 2 个）**已全部修复**，两端 `tsc --noEmit` 与 `vite build` 均已通过。

### 6.2 未做的真实验收（如实列出，**不能用构建成功代替**）

- **真实 Excel 中的点击验收**：完全未做
- **真实 WPS 中的点击验收**：完全未做
- **文件夹模式下的选择与无模型行为**：未在真实宿主验证
- **窄侧栏（约 350px）与模型设置弹层**：未在真实宿主验证

以下场景**只在自动化测试中覆盖，没有任何一次真人点击验证**：

1. 首次启动、目录尚在加载时的界面表现（Excel 与 WPS 各一次）
2. 后台已停止时打开"添加模型"表单，然后恢复后台再测试/保存（草稿是否还在、旧离线标记是否锁死按钮）
3. 配置一条连接 → 删除它 → 再删除最后一条的界面流转
4. 旧 `local` 记录（可手工把 `localStorage` 的模型项改回 `local`）的实际迁移
5. 模型调用失败（无效 Key / 供应商断开）时的用户可见反馈
6. 断点恢复 / 续跑：暂停任务 → 顶部切换模型 → 用新模型接管同一 `turnId` 继续
7. Excel 明确写入需求仍须预览 + 确认；WPS 明确写入需求仍被拒绝

### 6.3 已知遗留（非本次范围，供决策）

1. **确定性地址命令的正则宽松度**：`_local_address_plan` 的区域正则未锚定整句，因此像"清空 B2 到 B10 的数据"只会匹配到 `B2`，与用户意图不符。这是既有的解析器限制；本次改造把它从"仅无模型时生效"扩大为"始终先于模型生效"，命中面因此变大。建议后续单独收敛该正则（例如要求整句匹配），不要在本次改动里顺手改。
2. **WPS 测试不依赖 `@testing-library/react`**：WPS 工作区没有该依赖，根目录那份会带进 React 19.2.x，而 WPS 用的是 19.1.1，形成双实例触发 Invalid hook call。现改用本工作区的 `react-dom/client` + `react` 的 `act` 手写交互测试（约 60 行辅助函数）。若今后 WPS 需要更多交互测试，建议给 `apps/wps-addin` 显式加上 `@testing-library/react` 依赖，与本地 React 版本对齐。
3. `BASE_MODE_HELP_TEXT` 作为 deprecated 别名保留在 `helpCommand.ts`，等确认外部无引用后可删除。

---

## 七、变更文件清单

### 后端

```
server/app/llm/config.py        server/app/llm/__init__.py
server/app/planner.py           server/app/intent.py
server/app/main.py              server/app/models.py
server/tests/test_planner.py    server/tests/test_main.py
```

### 前端（Excel）

```
apps/excel-addin/src/App.tsx
apps/excel-addin/src/api.ts                      apps/excel-addin/src/api.test.ts
apps/excel-addin/src/modelSelection.ts           apps/excel-addin/src/modelSelection.test.ts
apps/excel-addin/src/helpCommand.ts              apps/excel-addin/src/helpCommand.test.ts
apps/excel-addin/src/hooks/useServiceHealth.ts   apps/excel-addin/src/hooks/useServiceHealth.test.ts
apps/excel-addin/src/hooks/useModelManagement.ts apps/excel-addin/src/hooks/useModelManagement.test.ts
apps/excel-addin/src/styles.css
```

### 前端（WPS，整个目录尚未纳入 Git）

```
apps/wps-addin/src/WpsAgentPanel.tsx             apps/wps-addin/src/WpsAgentPanel.test.tsx
apps/wps-addin/src/WpsModelSettingsPanel.tsx     apps/wps-addin/src/styles.css
```

### 文档

```
README.md  docs/ARCHITECTURE.md  docs/DEVELOPMENT.md  docs/COMMAND_CATALOG.md
```

### 改动前就存在的用户改动（**非本次改造引入，已被保留**）

```
M README.md            M config/capabilities.json   M docs/ARCHITECTURE.md
M docs/DEVELOPMENT.md  M package.json               M server/tests/test_safety.py
?? apps/wps-addin/     ?? docs/BASE_MODE_RETIREMENT_GUIDE.md 等 5 个 docs
```

---

## 八、Git 状态

**未提交任何改动、未推送、未打包安装程序**（遵守任务书要求）。

---

## 九、备注

- 本次实施期间工作区存在并行编辑。`server/app/planner.py` 与 `server/tests/test_planner.py` 的最终形态以**当前工作区实际代码**为准；本文档已按实测结果更新。
- 上一轮遗留的失效引用已同步修正（改 `planner.py` 后 `planner_module.resolve_model_config` 不再存在，旧测试辅助函数会以 `AttributeError` 失败）。
- 沙箱内 `bash` 的 coreutils 不可用（`ls`/`cat`/`wc` 均 not found），文件与命令操作请走 PowerShell 或专用读写工具。
