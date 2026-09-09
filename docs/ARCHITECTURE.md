# Excel Bro 架构文档

## 系统架构

### 整体架构

- 前端：React + TypeScript + Office.js
- 后端：Python FastAPI
- 通信：HTTP REST API
- 存储：localStorage（前端）+ JSON 文件（后端规则）

```text
┌───────────────────────────────────────────────────────┐
│                      Excel 任务窗格                     │
│                                                       │
│  React App.tsx                                        │
│    ├── 13 个 Custom Hooks                              │
│    ├── UI 组件                                         │
│    └── 工具模块                                        │
│                                                       │
│  Office.js ──> 当前 Excel 工作簿                       │
└──────────────────────┬────────────────────────────────┘
                       │ HTTPS / REST
┌──────────────────────▼────────────────────────────────┐
│                 FastAPI 本地服务                       │
│                                                       │
│  /api/turn  /api/intent  /api/plan                    │
│  /api/models  /api/settings/model                     │
│  /api/folder  /health                                 │
│                                                       │
│  Python 服务模块                                       │
│    ├── models.py / contracts.py                       │
│    ├── intent.py / planner.py / excel_agent.py        │
│    ├── folder_workbooks.py / folder_data.py           │
│    └── capabilities.py / safety.py                    │
└──────────────────────┬────────────────────────────────┘
                       │ OpenAI-compatible API
┌──────────────────────▼────────────────────────────────┐
│                  外部模型服务                          │
└───────────────────────────────────────────────────────┘
```

### 前端架构

#### 核心组件层级

```text
App.tsx (5946 行) - 业务协调层
├── 13 个 Custom Hooks - 状态管理层
│   ├── 状态管理类 (6 个)
│   │   ├── useConversation
│   │   ├── useUIState
│   │   ├── useExecutionApproval
│   │   ├── useScopeSelection
│   │   ├── useActivityProgress
│   │   └── useUndoSnapshot
│   ├── 资源管理类 (4 个)
│   │   ├── useImageAttachments
│   │   ├── useModelManagement
│   │   ├── useToolManagement
│   │   └── useWorkbookContext
│   └── 交互辅助类 (3 个)
│       ├── useSlashCommands
│       ├── useCopyFeedback
│       └── useServiceHealth
├── UI 组件（实际接线）
│   ├── PetCompanion / RuleManager / SlashCommandAutocomplete / ThemePanel
│   └── 其余界面（消息流、范围浮层、设置/工具/历史抽屉、输入区等）
│       内联渲染在 App.tsx
│   （曾把上述界面抽成独立展示组件文件，但从未被引入，已作为死代码删除）
└── 工具模块
    ├── conversation.ts
    ├── format.ts
    ├── intent.ts
    ├── range.ts
    ├── utils.ts
    ├── excel.ts
    ├── dataTools.ts
    └── storage.ts
```

#### Hooks 架构

Hook 之间不直接调用彼此的内部状态，主要通过 props 或回调通信：

```text
useServiceHealth ──refreshServiceHealth──> useModelManagement

useConversation ──setMessages──> App.tsx

useScopeSelection ──workbook/snapshot──> App.tsx

useActivityProgress ──onPersistLog──> App.tsx

useUndoSnapshot ──onAfterUndo/onMessage──> App.tsx
```

#### 数据流向

```text
[用户交互]
    ↓
[App.tsx 协调层]
    ↓
[Hooks 状态管理层]
    ↓
[API / Office.js]
    ↓
[FastAPI / 当前工作簿 / localStorage]
```

### 后端架构

`server/app/` 的主要模块：

```text
server/app/
├── main.py                 # FastAPI 应用、路由、CORS
├── models.py               # Pydantic 请求/响应模型
├── intent.py               # 意图判断和澄清
├── structural_intent.py    # 可确定识别的高风险工作表结构命令
├── planner.py              # 基础模式与最终规划
├── excel_agent.py          # 模型工具循环
├── agent_loop_protection.py # Agent 预算、循环检测和安全降级
├── turn_state.py           # turnId 进程内热状态、TTL 和缓存
├── task_ledger.py          # SQLite 持久任务、规划现场、事件、预算和幂等状态
├── folder_workbooks.py     # 文件夹执行
├── folder_data.py          # 文件夹确定性查询
├── capabilities.py         # 共享能力配置
├── safety.py               # 安全与协议校验
├── rule_generator.py       # /function 公式生成主流程
├── formula_tools/          # /function 工具调用架构（schema/编译器/注册表/Prompt）
├── formula_tools_integration.py  # 工具调用路径与主流程的衔接
└── llm/                    # 模型调用、错误和重试
```

### Agent 循环保护

excel_agent.py 的每轮模型调用都由 LoopState 追踪工具路径、缓存命中、连续结构
错误、累计 token、墙上时间和快照截断。循环开始前按固定顺序检查四层预算闸：
最大轮次、连续错误、累计 token、墙上时间。检测层还会识别重复缓存查询以及
2 至 4 步的重复工具路径，并把一次性警告附加到工具结果；接近预算尾部时，
根据发散、收敛失败或循环状态向最后一条用户消息追加收尾提示。
一批只读工具成功返回后，后续轮次只暴露 submit_plan 与 submit_answer，要求
Agent 基于已有上下文直接收敛，避免重复读取工作簿耗尽 token 预算。
结构化澄清选项一经用户点选即视为最终语义，不再把确认结果送回意图模型重复
提问。前端通过 `PlanRequest.planningReadPolicy="none"` 标记该路径；规划 Agent
只获得 `submit_plan` 与 `submit_answer`，并依据已授权工作表的准确表头生成
运行时计划。`splitGroupAggregate` 等动作仍在执行阶段读取完整数据，因此不受
200 行模型快照限制。
澄清选项通过 resolutionMode 区分 append 与 replace。append 只补充缺失参数，
并明确最终确认高于原始背景；replace 的 resolution 必须是完整最终需求，前端会
彻底丢弃与之冲突的原始指令。计划提交格式失败后，历史工具参数会压缩为最小占位，
下一轮只保留 submit_plan，并使用 agent.repairMaxOutputTokens 限制输出；生产环境
连续两次结构错误即停止，避免完整规划反复消耗 token。自由文本提交失败时也只在
历史中保留最小占位。splitGroupAggregate 的 splitBy 会自动成为输出列；模型
若又把同名字段放入 groupBy，服务端会在校验前确定性移除并在计划警告中披露，
而不是消耗一轮模型调用修复不改变语义的协议冗余。模型把采用 default_factory
的顶层集合字段提交为 null 时，同样按省略处理并使用模型默认值。

当前输入若自身包含完整动作和目标，必须作为新任务覆盖最近对话和上一轮结构化
意图；只有明显省略式追问才承接旧任务。“删除除某表外的所有工作表”由前后端
共同使用的确定性语法识别，规划器直接生成 `deleteWorksheet` 白名单动作，不再
允许模型复用上一轮拆分或汇总计划。
拆分聚合执行后，两条执行通道都会根据本次实际生成的工作表名称建立动态
`rangeEquals` 验收条件，并重新读取写入范围核对表头、分组字段、指标和占比；
超过 500 行的结果会分块验收，不以“工作表存在”代替具体效果验证。
数值比较采用与 Office.js 相同的浮点容差，只忽略 Excel 保存重读产生的机器精度
末位差异；可见数值变化、计数错误和占比口径错误仍会导致验收失败。
文件夹执行在原子替换成功后、独立验收开始前登记服务端恢复记录，因此验收
`failed` 或 `executed_unverified` 不会丢失撤销入口。恢复记录只保存服务端生成的
目标路径、备份路径与写入前后内容指纹；`/api/folders/undo` 按 LIFO 级联恢复，
已有文件从自动备份原子替换，本次新建文件恢复为不存在。恢复前先验证所有当前
文件仍等于对应写入后指纹、备份仍等于写入前指纹；任何一项变化都在磁盘写入前
整体拒绝，避免撤销覆盖后续人工编辑。文件夹备份集中位于所选工作目录的
`.excel-bro-backups/executions/<undo_id>/`，同一次执行的多个文件共享一个
manifest 并保留原相对路径。manifest 中每个目标使用 `recoveryAction` 区分从副本
恢复的 `restore` 与撤销时移除本次新建文件的 `remove`；旧 manifest 缺少该字段时
根据 `backupRelativePath` 兼容推导。manifest v2 记录 preparing、prepared、applying、
ready、recovering、restored、rolled_back 与 needs_attention 生命周期以及文件级前后
指纹；旧 v1 记录保守降级为只读历史，不冒充可直接撤销。快速撤销栈淘汰和撤销成功只
改变资格或历史状态，不直接删除磁盘副本。

清理先保护当前撤销链、执行/恢复/导出中的记录、主动保留和异常资料，并为每个原文件
保留最后一份可核对副本；7 天是自动副本最短保留期，1 GB 是软上限。自动清理和手动
删除复用同一事务：验证 manifest、登记根目录、整条路径和 Windows reparse point 后，
先将完整执行目录原子移动到 `.trash/<operation_id>/`，再递归清理；占用导致的残留纳入
待清理空间并可在下一次“清理旧备份”安全重试。删除确认同时绑定 manifest revision 与
级联失效 ID，影响变化时拒绝旧确认。撤销自身的回滚救援副本集中保存在
`operations/<operation_id>/rollback/`，仅在恢复无法完整回滚时保留。

备份列表为已登记目录签发短期 `backupManagementToken`；打开、取出、保留、删除、清理
和旧版归拢均不接收任意目标路径，也不扩大工作表数据授权。取出副本写入固定 exports，
使用稳定 operationId 和原子清单实现断线幂等重试。打开接口只解析列表登记的
`.excel-bro-backups`，目录不存在时不签发令牌，也不创建空目录。文件夹保存临时结果前
及备份复制前检查磁盘余量，无法保留配置的安全余量时在替换原文件前停止。

Office.js 执行快照按真实 executionId 绑定，成功、动作失败和验证异常都返回已经捕获的
恢复资料。快照记录写后指纹、工作簿身份、新建表状态与覆盖完整度；撤销前逐项复核，
覆盖不完整、工作簿切换或后续人工修改时停止直接写回。

文件夹动作结果可携带 `changed`。执行器对能够确定比较的写值、写公式、写表格、
排序、清空内容或超链接、清除筛选、数字格式和冻结窗格进行本地前后状态比较；
全部动作均未改变目标时不进入文件保存与恢复点创建。暂时无法可靠比较的动作保守地
视为已变化，避免为了节省空间而漏掉恢复点。

达到 token、轮次、连续错误等硬限制后不再抛出裸异常，而是返回 provider 为 system 的
AnswerResponse，其中包含停止原因、压缩后的工具调用记录和针对截断、格式错误
或循环的建议。流式规划另受 `taskRecovery` 的时间预算控制：首次 60 秒，在完整模型响应
或完整单次工具调用边界持久保存现场并暂停；用户通过带 revision 和稳定 requestId 的明确操作继续，
每次增加 60 秒且累计不超过 300 秒。继续不会提高 token、轮次或错误上限。
OpenAI-compatible 适配器统一返回 payload 与 usage；缺失 usage 的服务会得到零值
token 统计。

### 任务进度与可审查判断轨迹

规划与两条执行通道统一使用 `TaskProgressEvent`。事件包含 `operationId`、单调递增
`sequence`、阶段、步骤、状态、可选计数、目标、`toolName`、`facts` 和 `trace`；
前后端分别在 `contracts.ts` 与 `models.py` 校验。`toolName` 由本地工具执行器填写，
`facts` 只允许工具级白名单摘要；前端仍会限制数量和长度、隐藏疑似凭据并把越界
绝对路径降级为文件名。`trace` 只记录可核对的理解、证据、假设、决策、冲突和
纠偏，不传递模型隐藏思维、提示词或只读工具返回的原始数据行。

`/api/turn/stream` 继续兼容旧 `step` 事件，但载荷已使用统一协议。Office.js
执行器直接向活动 Hook 推送预检、撤销数据捕获、逐动作、拆分子表和验证事件。
文件夹执行通过 `/api/folders/execute/stream` 推送预检、逐动作、逐子表、集中备份、
逐文件原子保存和验证事件。客户端断线时用同一 `operationId` 和最后收到的
`afterSequence` 重连。`FolderOperationLedger` 把请求指纹、单调事件、终态结果和错误
持久保存到本机应用数据目录的 SQLite；内存记录只是热缓存。完成/失败记录按
`taskRecovery.operationRetentionDays` 保留完整载荷，之后压缩成永久幂等墓碑，因此旧
`operationId` 不会在 TTL、容量淘汰或服务重启后变成一份新的写入。旧的非流式写入
入口统一拒绝执行。状态通过 `GET /api/folders/execute/status` 查询。

前端在本地保存当前待挂接的文件夹操作身份；任务窗格刷新后先查询持久状态，再用同一
操作标识重连。已完成结果可以跨服务重启重放；重启时仍是 `running` 或记录证据不足的
任务转换为 `needs_reconciliation`，不会自动重做 Excel 写入。用户已登记的工作目录由
持久索引定位，启动只扫描这些目录的受管恢复元数据，不扫描整盘，也不自动扩展授权。
执行、撤销和独立恢复共享规范化目录的进程内及跨进程排他锁；独立恢复还必须核对
manifest revision、生命周期、路径边界和写前/写后文件指纹。

这是 [任务断点恢复指导 v0.5](TASK_CHECKPOINT_RECOVERY_GUIDE.md) 的分阶段实现：阶段 A
提供持久身份、事件补收、安全挂起和恢复到执行前；B1 仅对单个
`splitGroupAggregate` 保存完整暂存产物和逐文件发布 checkpoint，重启后核对指纹并跳过
已发布文件；B2 保存流式规划的 messages、LoopState、工具缓存、开放工具模式、修复状态和
完整待处理模型响应，可在安全边界暂停并显式继续。阶段 C 进一步按单个工具调用保存完成
状态，允许用户在完整模型调用边界显式选择已配置备用连接，并为已保存工作簿中的
`writeValues`、`writeFormulas`（非 R1C1 铺满）、`setFill`、`setFont` 和
`setNumberFormat` 提供动作级续跑。任意 DAG、结构动作和复杂对象续跑仍未开放。

规划连接中断后，持久状态仍为 `running`、进程内没有活动执行者且存在完整 checkpoint 时，
对外解释为 `interrupted`。用户可明确选择能力匹配的已配置模型接管同一 `turnId`；状态查询和
继续接口使用同一资格结果，并在账本 CAS 领取与进程内活动登记之间保持互斥。前端续跑复用
原活动消息及 `operationId`，以客户端最后实际归并的事件序号请求补发；服务端最新序号只用于
校验上界，不能替代客户端游标。事件补发只恢复同一条时间线，不代表重新执行已完成工具。

Office 动作恢复由服务端 SQLite 账本作为权威现场：确认执行后、首个写入前登记计划、原始
授权和稳定工作簿 URL 身份；每一步在 `context.sync()` 前提交目标范围的写前快照，写入并
读取实际状态后再提交累计检查点。刷新或异常退出后，前端本地只保存 operationId 等发现
索引；继续前重新查询服务端、核对当前工作簿身份，并对在途范围做三态判断。只有完全匹配
预期结果时补记为已完成，完全匹配写前状态时重试，其他状态保持挂起。写入权通过 revision
和轮换 ownerToken 约束，状态查询不返回令牌。

活动 Hook 从 `started` 起就把步骤加入有序时间线，按 `operationId + sequence` 隔离
并去重事件，再按 `stepId` 原位更新。每段活动在消息流中拥有独立消息和稳定
`persistenceKey`；实时进度与完成日志更新同一个消息节点，禁止再通过“最后一条助手
消息”推断日志归属。用户消息会排在本轮过程消息之前，需求确认、计划和结果会追加在
过程消息之后，因而不会在完成时重排或跳位。旧记录中附着在普通助手消息上的
`activityLog` 继续兼容显示，刷新时会丢弃没有完成日志的临时活动占位。

实时卡和历史消息共用默认展开的单层时间线。完成后保存包含工具、动作、目标、安全
事实和判断轨迹的完整步骤快照；活动级 `outcome` 与最后一步状态分离，需求分析步骤
可以完成而整段活动显示“等待确认”。网络中断只记录连接状态并保留最后进度，不能冒充
失败、完成或已取消；验证失败使用 warning/conflict 表达，不把“文件已经写入”错误
显示成“任务未执行”。消息流在用户接近底部时自动跟随新增步骤；用户主动向上阅读后
不抢夺滚动位置，而是显示返回最新进度的提示。

### /function 公式生成链路

`/function` 短链优先走「本地确定性跨表公式」（前端 `crossTableFormula.ts`，
不调用 AI）。覆盖不了时后端按两级架构生成：

1. **工具调用路径**（`formula_tools/` + `formula_tools_integration.py`）：
   模型只返回结构化函数调用 JSON（`FormulaToolCallResponse`），由代码编译器
   递归编译成公式文本，保证语法正确、字面量不丢失；编译结果与主流程一样经过
   `dangerous_formula` 安全检查。
2. **端到端兜底**（`rule_generator.py` 原有逻辑）：工具调用路径任何一步失败
   （模型未配置、JSON 无效、编译失败、安全检查拦截）都自动回退到原有
   「模型直接生成公式文本」方式，用户无感知。

协议不变：请求仍为 `GenerateFormulaRequest`，响应仍为
`GenerateFormulaResponse`（现代版 + 兼容版两条公式）。工具调用路径只生成一版
公式：`compatFormula` 与 `modernFormula` 相同，但 `compatExplanation` 会附加
「⚠️ 此公式使用了 Excel 365+ 函数，可能不兼容旧版本」警示，避免承诺虚假兼容；
真正的兼容版转换（XLOOKUP → VLOOKUP 等）留待后续实现。

### 通信协议

主要 API：

- `POST /api/turn`：统一轮次入口
- `POST /api/turn/stream`：SSE 流式规划
- `GET /api/turn/{turnId}/status`：查询持久规划状态、预算和续跑能力
- `POST /api/turn/continue/stream`：按 revision 与稳定 requestId 显式继续暂停的规划
- `POST /api/intent`：兼容意图入口
- `POST /api/plan`：兼容规划入口
- `GET /api/models`：模型目录
- `GET /api/settings/model`：模型配置
- `PUT /api/settings/model`：更新模型配置
- `POST /api/settings/model/connections`：新增连接
- `DELETE /api/settings/model/connections`：删除连接
- `POST /api/formulas/generate`：/function 短链生成原生公式（可携带文件夹模式勾选的外部工作簿工作表上下文 `extraSheets`）
- `POST /api/folders/select`：选择并扫描文件夹
- `POST /api/folders/refresh`：刷新文件夹文件列表（保持 sessionId 不变）
- `POST /api/folders/execute`：兼容的非流式文件夹执行入口
- `POST /api/folders/execute/stream`：带 operationId 幂等保护的 SSE 文件夹执行入口
- `GET /api/folders/execute/status`：查询可恢复的文件夹执行状态和最终结果
- `POST /api/folders/execute/resume/stream`：继续经核验的 B1 拆分任务剩余发布
- `GET /api/folders/backups`：列出当前工作目录的受管理备份
- `POST /api/folders/backups/export`：按 operationId 幂等取出副本到固定 `exports/` 目录
- `POST /api/folders/backups/delete`：按修订号和影响摘要复核后事务删除指定受管备份
- `POST /api/folders/backups/cleanup`：安全清理过期、孤立和超软上限备份
- `POST /api/folders/backups/legacy/collect`：确认后归拢旧版散落备份
- `GET /api/diagnostics`：诊断信息

前端协议定义在 `apps/excel-addin/src/contracts.ts`，后端协议镜像在 `server/app/models.py`。
AnswerResponse.provider 支持 model、local 和用于预算熔断降级的 system；
PlanResponse.provider 仍只允许 model 或 local。

## 关键设计决策

### 为什么选择 Custom Hooks？

- 状态和副作用能按业务域拆分，避免 `App.tsx` 继续膨胀。
- Hook 与 React 渲染模型天然集成，无需引入额外运行时。
- 单一 Hook 可独立测试，测试成本低。
- 组件只消费清晰的返回接口，降低耦合。

### 为什么不使用状态管理库？

- 当前状态主要是局部 UI、对话、工具和模型配置，跨组件共享需求有限。
- 引入 Redux/Zustand 等库会增加包体积、样板代码和学习成本。
- Custom Hooks + props 回调已经能覆盖现有协作场景。
- 后续如出现复杂全局状态，再评估轻量状态库。

### 为什么使用 localStorage？

- 本地优先，无需账户和跨设备同步。
- 保存对话、工具、模型选择和 UI 偏好已经足够。
- 用户将 Excel Bro 任务窗格拉宽到 380px 以上时，前端保存实际宽度；下次打开或重新显示窗格时，在 `TaskPaneApi 1.1` 可用的 Excel 中通过 `Office.extensionLifeCycle.taskpane.setWidth()` 恢复。旧版 Excel 不支持该 API 时继续使用手动拖动引导，不伪装成已恢复。
- 常用工作目录的扫描目录与工作表勾选会本地保存；首次产品引导的“以后再设置”也使用版本化 UI 偏好记录。重新使用目录时仍由后端扫描和校验，缓存不能代替文件访问检查。
- 设置工作目录只建立候选文件范围，不等于授权其中全部工作表；授权边界仍由用户的工作表勾选确定，首次设置不会自动勾选。
- 对 API Key 只保存到后端本地配置，不进入前端 localStorage。
- Office.js 待恢复任务在 localStorage 中只保存 operationId、对话定位和稳定工作簿身份；
  完整计划、授权、写前快照和动作检查点仅由服务端持久账本提供。
- 数据量较小，读写频率可接受。

## 性能考虑

### Bundle 大小

- 当前前端为单任务窗格应用，主要逻辑在 `taskpane` bundle。
- 优化策略：
  - 按需拆包
  - 避免重复依赖
  - 工具函数模块化，减少重复内联
  - 未来可评估 `React.lazy`

### 渲染性能

- 关键指标：任务窗格加载、规划流式渲染、本地查询进度。
- 优化点：
  - 保持 `useCallback` 依赖完整，减少无意义重建。
  - 对长消息列表按需渲染。
  - 进度更新使用函数式 state，避免闭包过期。
  - 清理定时器和事件监听，防止后台开销。

## 安全考虑

### 公式黑名单

- 本地 `formulaSafety.ts` 拦截危险公式。
- 服务端仍通过协议和动作白名单限制最终写入。

### 工作表授权边界

- 计划编辑器调用 `/api/plans/validate` 时，通过 `PlanValidateRequest` 同时发送本次授权工作表和当前现有工作表名称。
- Office.js 执行器在最终预检中以授权集合检查所有现有来源和目标表；新输出表只有在执行前不存在时才可越出授权名称集合。
- 文件夹执行器以会话内已选“文件 + 工作表”引用作为授权集合，并拒绝读取或覆盖输出工作簿中的未授权现有表。
- “仅保留某表、删除其他工作表”不会自动把未选表加入授权集合；前端先切换到多选范围并要求用户勾选全部涉及工作表，授权完整后才生成删除预览。
- 文件夹快照中的动作目标使用“文件 › 工作表”限定名；确定性结构命令可用 `sourceSheet` 解析用户输入的原始表名，但一次只能作用于一个已选文件，避免同名表歧义或跨文件误删。
- 授权检查必须发生在任何写入、备份或文件替换之前。

### API Key 管理

- 完整 Key 只保存在后端本地配置。
- 前端只接收掩码状态和连接 ID。
- 日志和诊断不记录完整 Key、提示词和原始数据行。

### Origin 校验

- FastAPI 只监听 `127.0.0.1`。
- CORS 只允许本地任务窗格来源。
- 当前工作簿模式通过 Office.js，不暴露任意文件访问。

## 扩展性

### 如何添加新功能

1. 判断属于语义决策还是确定性计算。
2. 协议层同步修改 `contracts.ts` 和 `models.py`。
3. 前端执行层实现 Office.js，后端决定是否支持 openpyxl。
4. 补充动作说明、风险判断、验收条件和测试。
5. 更新 `docs/ARCHITECTURE.md` 和 `docs/HOOKS_GUIDE.md`。

### 如何添加新 Hook

1. 在 `apps/excel-addin/src/hooks/` 新建文件。
2. 定义选项和返回类型。
3. 使用 `useCallback` / `useEffect` 并完整声明依赖。
4. 添加副作用 cleanup。
5. 编写单元测试。
6. 运行 `npm run test:addin` 和 `npm run build:addin`。
7. 更新 Hook 文档。

### 如何添加新命令

1. 在 `useSlashCommands` 或相关命令目录中注册命令。
2. 定义命令值、标签和描述。
3. 在 `App.tsx` 或对应组件中处理命令结果。
4. 补充命令过滤和自动补全测试。

## 参考资料

- P1 重构总结：`docs/P1_REFACTOR_SUMMARY.md`
- Hooks 使用指南：`docs/HOOKS_GUIDE.md`
- 设计审查报告：`docs/DESIGN_REVIEW_FINAL_SUMMARY.md`
