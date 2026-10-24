## 2026-10-24：用户 Excel 实机验收反馈

用户在 Python验收-Excel.xlsx 实机确认：双视图对比、放弃变更并恢复、确认变更、确认后撤销均正常。提供的四张截图均显示“已固定 3 个工作表”，本轮未再出现范围丢失。这是用户实机反馈，限定于本次 Excel 场景；不扩展为 WPS 或所有异常交错均已验收，也不据此宣称此前模式切回 auto 的根因已完整复现。

## 2026-10-24：后台扫描不得改写固定工作表选择

Excel 执行/撤销收尾会调用 scan。范围 Hook 的扫描回调改为读取当前模式，避免被旧任务保留的 auto 闭包把后来固定的多表草稿覆盖成活动表。固定选择在目录刷新中原样保留；缺失表仍由准入/目标守卫拒绝，不自动收窄、不自动扩大。WPS 同簿固定模式本就保留选择，本轮不改宿主适配。

新增回归先证明旧扫描回调会收窄勾选，再验证修复；真实 App 覆盖固定三表执行和确认后仍为三表。用户截图中“模式本身变成 auto”的完整实机交错尚未复现，不能把这一处已证明的勾选覆盖缺陷当成全部现场根因。未执行真实窗格验收。

## 2026-10-24：恢复 Excel 执行后双视图窗口排列

执行前值快照副本必须保留窗口桥约定的 EB_ 加 20 位十六进制身份；中文前缀会被 prepare/close 协议拒绝，导致只打开副本而无法并排或自动关闭。现恢复该身份格式，不放宽服务端识别与关闭边界。排列失败通过审阅状态展示具体原因，保留副本清理句柄以及确认/放弃入口，不重放主表写入。WPS 仍使用既有原生 CompareSideBySideWith 路径。

回归覆盖真实 Excel 面板生成种子、桥协议名称、调用排列以及 prepare/arrange 失败提示；双端构建与回归另行记录于 outputs/dual-layout-*。真实 Excel/WPS 窗口排列仍需实机验证，生成文件不能替代并排成功证据。

## 2026-10-24：双视图使用本次修改区域的执行前值快照

Excel 不再为执行后双视图调用整表 capture。执行前仅按计划的动作矩形读取 values、numberFormat、直接格式与行列尺寸；对比画布中动作间的空隙留白，不读取或导出空隙、其他区域或依赖表。副本只写固定值，不导出公式，不复制图表、名称、合并及条件格式效果。画布仍受既有限额约束；无法读取可靠快照仍明确失败。主表执行、回读、确认和恢复链不变。

WPS 保留宿主原生单表副本适配，但本次修改区域按执行前固定值重建，不重新写回公式；clearRange 的执行前对比同样恢复旧值，不再次清空。副本其他区域仅是宿主背景，不作为执行前证据。列宽和无法采集格式等既有宿主限制仍明确提示；不宣称复制整簿或全部复杂对象。两端共享审阅条说明对比范围、公式不重算及对象限制。

Excel 审阅期间的输入框提示改为“等待审阅结果”，其他不可暂停的处理中状态不再误称暂停失效。双端自动化覆盖值快照、公式不重算、原表不重复写入与恢复；Excel 检查区域外数据不进入副本，WPS 检查清空前值。真实双端窗格仍待验。

## 2026-10-24：可选双视图复制限制不再阻止主表执行

Excel captureExcelBeforeResult 曾复用完整副本的 capture 检查；工作簿名称、复杂对象和已有跨表公式导致可选素材采集失败，阻断自动执行。现用显式 ComparisonUnavailableError 区分复制能力限制；仅这类限制保存为不可用结果，写入仍经过既有执行器的授权、来源、恢复快照和回读检查。未知宿主错误、保护、授权与计划错误继续抛出，不作宽泛忽略。点击双视图时明确说明不可用，不伪造不完整副本、不重放执行；确认与放弃仍由已验证回执与恢复点处理。

WPS captureWpsBeforeResult 已将可选对比限制与动作恢复分离，本轮未改其行为。新增 Excel 真实 App/执行器加替身宿主用例覆盖名称、跨表公式、确认不重放、对比失败后撤销、未知读取错误及格式恢复快照失败零写入。替身 getSpecialCellsOrNullObject(Formulas) 现在从真实单元格公式状态返回区域，不再恒为空。

真实 Excel/WPS 窗格仍待验证。本次不认定截图中具体是哪一种复杂对象，也不修改用户工作簿。

## 2026-10-23：续跑计划接入执行后审阅

用户确认当前交互为计划就绪后执行并展示结果，再选择确认变更、放弃变更或双视图；不新增“确认执行 / 取消计划”门槛。本节覆盖此前本轮误提的执行前确认方案。确认变更只清除审阅标记，不重放写入；放弃变更使用既有恢复点。

Excel 首次 appendMessage 会登记自动执行候选，续跑 applyPlanningContinuationResult 只更新消息，导致计划落到旧副本入口。两条路径现共用 enqueueProducedPlan，普通计划与本地规则计划都登记到当前回复段的执行后审阅链；按消息身份去重，加载历史不触发。保留工作簿身份、授权、来源复核、白名单执行、回读和恢复边界。WPS 保持原有自动执行路径，本次误加的批准状态已撤回。

新增 excelWriteEndToEnd 回归：暂停后继续交付计划，无需再次点击执行即可写入并回读，默认不建副本，结果有放弃与双视图入口，确认变更不重复写入。真实 Office.js/WPS 窗格仍待验，自动化仅为真实组件/执行器加替身宿主。

## 2026-10-23：Excel 续跑结果落在当前回复段

真实服务只读记录 turn-4152c91387e24ef19fcbaef4cd7cf1ce 显示 submit_result_plan 已成功，恢复描述为汇总表 A1:G9 的白名单 writeValues 计划；用户反馈最新 Python 结果卡后没有预览入口。Excel applyPlanningContinuationResult 仍把计划写入旧暂停消息，和新回复段的进度分离。现捕获同 turnId 活动段的 persistenceKey，计划/回答落到该段；同任务旧暂停入口退役，历史进度不改写。无匹配活动段才使用原消息。WPS 既有路径在末尾 append 计划预览，本次不改执行器。

测试 messageFirstRouting.test.tsx 新增继续后预览与最新过程卡同属一条消息、仅一个计划的断言。临时退回旧目标消息时用例失败，还原后通过。只读服务事实不等于已验收真实窗格新版本；重新加载后的 Office.js 实机仍待验。没有绕过预览确认，没有改用户工作表。

## 2026-10-23：进度展示语义修订

两端共用 TaskProgressTimeline：有后续段的历史卡标明“本段已结束 · 已接续”；接续段保持接续标题，不再因 awaiting_input 改名为需求分析。普通 awaiting_input 只表示等待下一步操作；已收到 preflight-plan 完成证据时才提示预览已生成、待确认写入。closedByFinish 中的步骤以中断符号展示，真实 warning 保持警告；不修改历史事件或终态。planning.model 明确标作包含工具调用的整体阶段，耗时不可与工具耗时相加，原始事件顺序保持。

Python 同次分析的 analysis_completed 与 pythonResult 在同一个 python.analysis:<run_id> 完成事件内投递；无代码工件的兼容分支仍发送独立结果事件。结果引用、代码正文不进账本的边界不变。旧的持久化重复事件不按标题合并，避免误合并真实的多次分析。暂停详情的 resolved 表示本段已接续，不宣称整个任务或写入已完成。

双端自动化与服务端回归日志：outputs/progress-presentation-*.log。真实 Excel/WPS 窗格尚待验收；仅调整展示和事件身份，不改变写入授权或自动执行策略。

## 2026-10-23：发送与任务控制分离

发送按钮固定为 ↑，空输入/空表单提交不恢复任务。Excel/WPS 共用 TaskExecutionControls，放在输入框上方、表单之外；■ 暂停，暂停请求处理中显示禁用的 …，暂停后 ▶ 由用户另一次点击才恢复。按钮始终 type=button，暂停/继续使用不同 React key，不复用为提交按钮；忽略双击的第二次点击。输入文字的发送与既有语义续跑仍走原意图路径。

本次覆盖旧约定“空输入点发送即可继续”。暂停后的自动核对仍只查询状态；这次调整不新增写入能力、不改变授权和预览确认边界。组件原生 click 用例验证同步渲染时不会连带 form submit 或 resume；双端面板验证空提交无动作、明确 ▶ 才恢复。真实宿主验收尚未执行。

## 2026-10-23：暂停回执归属与历史段计时修订

两端暂停请求增加独立的回执代数：新一轮运行或新一次暂停会使旧回执失效，迟到成功、失败、超时不能覆盖当前输入框提示。Excel 续跑的活动收尾也检查原 controller 的归属；这是防守性保护，不宣称已复现实机交错。

共享 TaskProgressTimeline 的非活动段固定显示结束时的完整计时快照，不再每秒重绘。这修订此前“暂停后历史卡的等待进度应继续增长”的展示约定；底层等待区间仍保留供恢复时累计。冻结显示不代表后台已停止，服务端结论仍由暂停回执和状态查询决定，pauseRequested 只表示请求被接受。

双端回归覆盖输入框/卡片两入口各三轮暂停继续、迟到回执不污染下一轮、8 秒超时及超时后的迟到成功。Excel 新段位置额外比较对应的最新用户消息。服务端经 ASGI HTTP 路由检查重复暂停、清理后 interrupted、active=false 和未知任务 404。

证据：outputs/pause-followup-*.log。真实 Excel/WPS 窗格尚未验收；不把组件替身测试记为实机通过。未扩展写入/回读中途暂停。

## 2026-10-23：继续复用已完成分析，暂停到原任务

同日晚些：暂停/继续的**收尾按归属**再收一处。`resumePlanningTaskFromMessage` 的 `finally` 里，
`turnAbortRef` 与 `sdkBridge.end()` 原本就按 `ownsTurn` 判，但 `setStatus("idle")` 与续跑并发闸
是无条件写的 —— 一次被取代的执行只要晚于新一轮收尾，就会把新一轮的运行态按回 idle
（暂停按钮消失、可中止对象还在）并放掉闸。现统一为：controller / 运行态 / 并发闸 / 桥
各自只收"自己那一代"；并发闸从布尔改成记持有者（`AbortController | null`）。WPS 侧同口径
（`settleExecution` 读 `planExecutionTeardown`）。按钮形态由共享判据 `pauseButtonState` /
`executionPauseState` 决定：**没有可中止对象就不许画 ■**，`…` 表示不可暂停并给原因，▶ 表示
"已暂停且输入为空"。暂停请求的三种回执分开表达（本端已停止 / 服务端请求已接受 / 未确认），
`pauseRequested:true` 不等于"已停下"。注意两种既有口径不要当缺陷改：收尾后的"等待进度"
继续走属于"后台可能仍在执行"的分桶（`finishTimelineAt`），`is-live` 归零也不等于
`finishActivity` 已生效。真实 Excel / WPS 窗格仍未验收。

同日补修（字段理解阶段）：`/api/turn` 的意图分支、`/api/intent/message` 与兼容 `/api/intent` 在尚未产生规划身份时，将只读模型调用绑定到 HTTP 请求生命周期。`until_disconnected` 在 FastAPI 消费请求体后监听断连，取消并等待调用清理；迟到成功结果不登记决定、不发布。此机制不用于写入或持久规划。共享 `checkIntent` 对响应头及正文等待均响应取消；Excel 字段意图和异步结构读取、WPS 字段意图执行层也显式中断等待，WPS 不把中断步骤补成 completed。卡片与输入框暂停共用原处理入口。

SDK 的 Python 完成回执使用 resultHandle，现与查询回执的 resultRef 一样记录到 after_tool 检查点；恢复初始化继承已完成工具与结果引用，before_model 不再覆盖为空。模型仍是新的调用，接续提示明确优先复用有效引用、只处理剩余目标；引用过期/归属或来源核对不通过仍拒绝复用，不恢复原模型会话，也不绕过预览确认。

两端的恢复游标由 planningReplay.ts 综合同一任务的本地日志/暂停记录/在途游标，不能用服务端最新序号冒充已消费位置。轮询按身份和序号去重；接回进度的本地步骤收到响应即 completed。WPS 在 completed 页不提前以 awaiting_input 收尾；计划校验有独立完成回执。成功取回结果后退役该任务所有旧暂停入口。

POST /api/turn/{turnId}/pause 返回 PlanningPauseAck（turnId、pauseRequested）。它请求取消原规划子任务，保持活跃互斥闸直到清理结束；重复请求不再次取消清理。是否已停下、能否继续仍查询 status。两端输入框暂停均调用此接口，并中断流/轮询/结果校验等待；迟到回包不再提交界面。暂停后的自动状态核对只报告事实，不自动取回或处理完成结果。写入与回读中途暂停仍未开放，本接口不作用于已确认的写入动作。

# Excel Bro 架构文档

## 2026-10-22：执行确认与跨任务回复段修正

本节修订此前“容器即将启动即视为 running”的口径。`docker_worker.py` 仅在隔离闸通过、代码编译成功后，于执行入口向 stderr 发出固定 V1 控制帧；不含正文、数据、路径或密钥。Docker 适配器通过有界收集器识别完整首帧，并只触发一次 `on_runner_start`。创建 CLI 进程、写入 stdin 或容器启动尝试均不代表代码已执行。控制帧不进入模型、账本或工作簿。旧不可变镜像无首帧时，只能在合法成功回包后确认执行；失败/取消且无确认时不附加 executedVersion，双端显示“未确认执行”。使用实时确认需要重建隔离镜像并按既有验证流程更新不可变摘要，不能改用可变 tag 或绕过隔离检查。

内部 Python 工具回执新增封闭组合 cancelled/cancelled、cancelled/authorization_changed，沿既有投影白名单产生固定文案，不发布取消后的结果。取消与执行失败分开；没有确认的失败落 not_started，HTTP/SSE 字段未增加，两端与 models.py 的语义注释同步。

回复段接链与活动身份切换分别处理：继续已 superseded/closed 的父段也补 nextSegmentId，保留父段状态、日志与耗时；其它任务不接入这条链。Excel/WPS 仍共用 openReplySegment。WPS 的 updateActivity/endActivity 通过 updateWpsActivityMessage 统一写回，冻结段在归并/收尾前被拒绝，活动段还校验传入任务身份。

## 2026-10-21：暂停 / 继续的步骤身份、终态与只读补发（Excel 与 WPS 同口径）

**问题不是文案，是四层各写一套。** 事件状态在 `wpsProgressFromTaskEvent` 里被丢掉、
活动日志靠"下一步出现"猜上一步完成、桥/隔离分析/文件夹执行三条生产入口只发
`started + progress` 从不发终态、恢复时只修最后一行的计时 —— 合起来就是"过程卡
永远停在进行中"、"继续之后编号重来"、"已完成步骤又被拉回运行态"。

### 一、事件与步骤身份

- **完整保留结构化状态**：`wpsProgressFromTaskEvent` / Excel 的 `applyProgressEvent`
  一律把事件自己的 `status` 交给共享层，不再由本地按"有没有下一步"推断。
- **三层身份**：`operationId`（任务 / 桥调用）、`stepId`（步骤）、`requestId`（继续
  动作幂等键）。桥的步骤 id 是 `bridge.{tool}:{operationId 前 12 位}`
  （`agent_tool_bridge.py` 的 `bridge_step_id`），隔离分析的终局步骤是
  `python.completed:{py_xxxx}` —— **同工具连续两次调用是两个步骤，同一次调用的重放
  仍是同一个步骤**（序号只前进）。分隔符只能用 `:`：`models.py` 的
  `^[A-Za-z0-9_.:-]+$` 会拒 `#`。
- **终态必须真发**：桥（`completed` / `failed` / `cancelled` / 超时）、隔离分析
  （每次执行一条自己的终局）、文件夹执行（`folder.prepare-files`、各拆分表的
  `.sheets`、`folder.resume.publish`）全部补齐"这一步结束了"。执行异常时先把
  **被打断的那一步**关成 `failed`（标题"这一步没有完成"），再报整体
  `folder.failed` —— 顺序反过来，前端就只能靠猜。
- **父子阶段**：模型规划里覆盖多个子工具的阶段由阶段自己的事件宣告完成，不靠
  "出现下一步"假装完成。

### 二、暂停与继续

- **三条入口、两个落点**：点按钮 / 空输入点发送 → 直接进暂停卡；输入框里打
  「继续」「接着做」→ **仍先走意图理解**（`understandMessage` → `kind:
  task_recovery`），再落到同一条 `resumePlanningFromPause`。本地不写关键词分类：
  那是语义判断，归模型。
- **继续 = 新开一个回复段，不是改旧卡**：同一任务仍然只有一个有效执行实例
  （`operationId` / `stepId` / `eventSequence` / 结果引用 / 写入回执全部不变），
  但展示上**允许多个历史回复段**。继续后必须在这条用户消息之后**立刻**出现新的
  助手回复段（`openReplySegment` 把旧段降为 `superseded` 并接 `nextSegmentId`），
  旧段保留它停止时的状态，并提供「查看后续进展」的轻量定位入口。
  **旧约定「永远接回原来那张过程卡」已作废**：它把「恢复正确」寄托在"界面上只有
  一张卡"上，于是同一任务的历史被反复改写，用户也分不清哪条反馈是最新的。
- 段身份是一条独立的展示身份（`ReplySegment`：`id` / `taskId` / `status` /
  `previousSegmentId` / `nextSegmentId` / `continuesTask`），判据只有一份：
  `packages/core/src/replySegments.ts`。四条不许破的约束：
  1. **多段但只有一个活动段**（`segmentAcceptsProgress` 只认活动段，任务相同的
     旧段也不接受进度）；
  2. **先更新共享任务状态，再投影到段**（段切换不改步骤身份、事件游标、结果引用、
     写入回执）；
  3. **旧订阅抢不回活动段**（继续时把旧段冻结，迟到事件只进权威记录）；
  4. **历史段保留当时状态**，迟到结果在最新段里说明，不回头改写旧段。
- **段身份 = 承载它的消息 id**：`previousSegmentId` / `nextSegmentId` 与消息 id
  同一个键，否则「查看后续进展」要在两套身份之间翻译。
- 继续时**不继承上一段时间线**（`previousLog: null`）：带过来就等于把两段合成一段，
  新段会从第 1 步重复展示旧步骤。上一段完成了几步只用于一句文案，不复用时间线。
- **三条路径**：
  1. 原任务 `running` → **只读补发**（见下）接回事件流，不重发需求；
  2. `completed` → 取回 `result` 走同一条 `applyTurnResult` / 首条结果处理链；
  3. 有检查点 → 带 `expectedRevision` / `requestId` / `afterSequence` 继续。
- **重复点击合并**：`resumingPauseRef`（WPS）/ `planningResumeInFlightRef`（Excel）
  是**同步**闸，同一帧内两次点击只产生一次恢复；超时重试沿用同一个 `requestId`。
- **没有检查点的阶段**如实说不能继续，不假装重试成功。

### 三、只读补发端点

`GET /api/turn/{turn_id}/events?afterSequence=N` → `{turnId, status, active,
lastSequence, events}`。**不调模型、不写账本、不创建任务**。未知任务 404
`PLANNING_TASK_NOT_FOUND`，负数 422。这条通道存在的唯一理由：原任务在跑时
「继续」会被 409 挡住（再发一次就是重复规划、重复读表），但用户按下继续之后
什么都不发生同样是错的。两端轮询口径一致：`1200ms` 间隔、`90s` 上限，超限
如实说"仍在运行"，**不重发需求**。

### 四、计时与状态口径

- 步骤自己的终态 ≠ 活动级 `outcome`：`finishTimeline` 的第 7 参
  `currentStatus` 由调用方给 —— "带着待确认结果收尾"的活动里，这一步可能真做完了。
- `closedByFinish`：只记**被结束活动关掉、且没有结论**（`warning`）的步骤。恢复时
  只重开它们；`completed` / `failed` / `cancelled` 一律原样保留（"已结束的步骤不得
  被迟到的进行中事件重新激活"）。`closedByFinish` **缺省**（老日志）时才退回按
  `warning` 猜 —— 写成 `?? []` 会让"新日志明确为空"和"老日志没这字段"混成一件事。
- 恢复时**每一行**按自己已累计的耗时重新定点（`startedAt = now - 已累计`），不是只修
  最后一行。暂停时间 / 断线等待时间与真正执行时间分桶记（`timelineDurations`），
  后台仍在执行的那段不许算成暂停。
- 判据只有一份：`packages/core/src/activityTimeline.ts`。两端只按引用转发，
  变异证据见 `scripts/run_activity_resume_mutations.mjs`（T1–T6）。

## 2026-10-18：WPS 边框闭环

`setBorders` 在准备阶段按明确 A1 区域展开为单格外沿，内部线条归上/左单元格；物理边界键发现跨动作相邻侧归属冲突时整体拒绝。白名单、操作数/单元格限额及授权继续生效，保存工具保留原始计划而非展开后的内部动作。

边框本地快照含 LineStyle / Weight / Color / 颜色来源模式 / ThemeColor / TintAndShade；四个属性为 borderTop / borderBottom / borderLeft / borderRight。执行前必需字段读不到即拒绝；写后逐边回读；撤销仅恢复声明的边界并复核，主题色/自动色按模式恢复。协议同步 contracts.ts、models.py 与服务端校验，只传属性名与指纹，不传边框快照或单元格值。部分写入失败仍保留恢复记录。

审阅标记由共享 `reviewMarkStyleForPlan` 按区域重叠决定：边框动作使用底色，填充动作使用边框，同时修改二者则不覆盖实际格式。Excel/WPS 共用判据。原有用户条件格式不会因本轮被删除；显示优先级仍须真实宿主核对。

这是开发实现和自动化证据，不代表 WPS 任意版本均支持所有边框成员。缺少颜色来源或属性回读的版本写前拒绝；真实宿主验收见 DEVELOPMENT.md。


## 2026-10-18：WPS 范围能力扩展（仍未全量补齐）

WPS 新增填充、数字格式、垂直对齐、写表、复制子集与 computeValues。共用 writeTableValues/表头配色保持 Excel 既有命名区域起点能力；WPS 使用显式 A1 布局。复制纯函数 copyRangePlan 仅生成白名单动作，来源从当前授权的本机范围读取，不调用模型。WPS 的执行计划在准备期展开，恢复账本登记每个真实效果，confirm 不重复执行。

originalPlan 仅保存复合操作语义供工具复用，与展开后的执行计划同时纳入不可变校验；保存动态复制时不固化当次值。copyTypes 与 actionExpansion 进入能力配置，恢复协议仍只接受实际原语；客户端新增 fillColor/fillPattern/numberFormat 格式字段，server/models 与 main 白名单同步。填充和字体使用 WPS OLE RGB（红色位于低字节）；修正原字体转换红蓝颠倒。

新增格式字段逐格采集；含未知目标字段写前拒绝。撤销只恢复声明字段，执行前副本也只恢复本次改动字段，避免无关 NumberFormat 被重写。copyRange 当前限 values/link，可条件选行、转置、跳过空白；不支持任意前序数据写入后依赖变化的同计划复制，也不伪装支持相对公式或完整格式复制。

computeValues 在 WPS 准备期通过既有本地 query_table 求值，展开 writeValues + 可选表头 setFont/setFill。必须是计划首个动作，来源和既有目标均在本次授权范围；值、公式及归一依赖的数字格式/显示文本都保留本机证据并在首次写入前重新核对。条件复制/清空同样保留源区域证据，缺少唯一表头字段不当成零命中。原始计算计划保留供工具复用，不能固定旧结果。

WPS 检查点异步返回后再次检查目标内容/格式及可编辑状态；新增恢复字段 verticalAlignment 同步客户端协议、server/models 与 main 属性闸。Excel 条件复制共用 copyRangePlan，逐实际写入片段采集恢复点、生成动态回读判据；跳过空白格不写不恢复。无条件复制仍沿既有宿主原语，本轮不宣称补齐它的所有恢复限制。

格式与值的恢复记录在首个宿主写操作前建立；中途异常只归属实际尝试过的动作。重叠格式动作在自身同步写入后刷新前序动作的最终快照，保留部分成功的恢复证据；异步等待期间出现人工改动时不把它吸收进本次写后快照，后续动作和撤销均拒绝覆盖。目标格式回读缺失不能当作有效写后证据。

deleteWorksheet、排序/筛选/去重、结构与对象操作仍未实现，详见 [能力差距表](HOST_ACTION_CAPABILITY_GAP_2026-10-14.md)。本批新增能力的真实宿主验收未执行。受控 JS 与文件夹执行保持既有边界。


## 2026-10-18：SDK 连接配置收口（第二批）

Excel/WPS 使用同一连接管理 Hook 与协议控件。托管连接新增 `protocol`（`openai_compatible` / `anthropic`），旧文件缺字段按 OpenAI 兼容处理；编辑请求省略字段时保留旧协议、空密钥保留原密钥。显式 Anthropic 地址按供应商配置原样使用，不猜测自定义网关地址。DeepSeek 官方 OpenAI 地址仍沿已有受控映射，模型与密钥不变。

模型目录和设置返回 `agentAvailable`、`agentUnavailableReason`、`formulaAvailable`。`available` 仍表示连接已配置，不等于 SDK 协议可用，更不等于实测连通。Agent 可用性核对协议、密钥、仓库外运行目录；目录不会联网探测。双端菜单、共享选择器和恢复候选排除明确不可用的 Agent 连接。已有不兼容选择保留在本地，但显示为需要选择兼容连接，不静默切换供应商。原 `claude-agent-sdk:<连接ID>` 仍可解析，前端迁移到同一 `connection:<连接ID>`，目录不重复展示。

连接测试新增明确用途：`POST /api/settings/model/connections/test` 接收 `purpose`，返回 `testedFor`。SDK Agent 测试通过真实 SDK 注册唯一 `connection_probe` MCP 工具，必须收到正确随机令牌的回执；只有文本 OK 不算通过。没有表格、文件、网络或命令工具，不发送工作簿数据；仍禁用 SDK 会话持久化。独立公式测试仅走 OpenAI 接口，不能用于验证 SDK；旧客户端省略 purpose 时保留公式测试语义，新客户端拒绝用途不匹配或缺少 testedFor 的成功回执。修改协议/用途/内容后，旧测试立即失效。限额在 `modelConnectionTest`（60 秒 / 3 轮）；密钥缺失在启动 SDK 前拒绝，不借用缓存登录。

独立公式只能选择 OpenAI 兼容连接；它与表格 Agent 的 SDK 编排分开。两端均已接线，工作簿/文件夹共用同一服务配置，不新增宿主操作能力。协议同步 `contracts.ts` 与 `models.py`。SDK 持久会话、原生检查点恢复、旧循环模块物理拆分尚未实施，不能把本轮连接收口当成这些能力已完成。


## 2026-10-18：Agent 唯一入口与结构化本地工具

当前生产入口为：消息 → SDK 理解/澄清 → SDK 工具规划 → 授权本地工具/隔离 Python → 受控结果与 AnalysisPlan → 双端既有执行、回读、结果审阅、恢复。所有自然语言需求进入 SDK，确定性工具不在模型前截断需求。三个阶段目前仍是独立 SDK 调用，共用结构化任务上下文；不等于已实现持久 SDK 会话。

`agent_runtime.sdk_agent_connection` 是连接适配的唯一判据。显式 SDK 连接直接使用；既有 DeepSeek 官方 OpenAI 端点按已验证映射转换，模型、凭据和视觉标志不变。未知网关明确返回 `MODEL_NOT_AVAILABLE`，不拼接猜测 `/anthropic` 地址、不回退普通模型、不换供应商。普通传输仍用于独立公式与连接测试，不再是表格 Agent 路线。

`planner.create_plan` 不再调用 `run_excel_agent`，`intent.check_intent` 不再提前解析自然语言删表。旧 `excel_agent.py` 暂留共享工具实现、提示与历史循环测试，生产规划入口没有调用旧循环；模块尚未完成物理拆分。旧检查点不能当作 SDK 会话继续，明确拒绝而不重新执行。

新增 `plan_tools.py`：

- `plan_delete_other_worksheets(keepSheets[])`：只对当前授权的单个文件生成删表计划；多个保留表为数组，名称在授权元数据内唯一解析；不同文件同名表、未授权名称、无变化均返回明确错误。
- `plan_range_edit(sheet, range, operation, value)`：按结构化参数生成写值、公式、清空内容计划；保留字面常量类型，拒绝公式冒充常量、无效区域与超限矩阵。字面写入限额来自 `agentPlanTools.maxLiteralCells`。

这些工具不读值、不写工作簿，只输出白名单计划；按真实宿主能力目录裁剪工具及操作参数，提交再次校验。WPS 未支持的删表没有被放开。批量文件、复杂数据加工继续使用既有工具，不能靠本工具猜测跨文件目标。

澄清是 Agent 通用决策能力：需要关联键、重复项处理、覆盖目标等决定时使用已有澄清协议；保留原需求和用户回答。答案不扩大授权，范围变化仍走数据范围入口。SDK 意图上下文新增上一轮结构化意图、结果标题/表头/来源、工具失败的允许字段，不发送结果行、原始行或宿主错误全文。

两端共用 `modelSelection.ts`：所有有效 Agent 连接走结构-only 规划与 SSE 工具投递，不再根据 ID 前缀选择本地预计算。服务端把旧调用的读策略收窄为 `none`，仅投影元数据；任意 `dataResults` 仍拒绝，计算值须经受控工具桥及有归属结果引用。SSE 缺失时明确失败，不回退到无法投递工具的请求。

SDK 已列入基础依赖，`npm run dev:server` 统一启动 Windows 单进程、无热重载服务，使用仓库外 SDK 目录。未新增 HTTP DTO；`contracts.ts` 与 `models.py` 沿用已有协议。本轮不改结果审阅与宿主恢复算法。历史段落中的“两后端”以本节为当前状态。

## 2026-10-18：限定工作表范围的受控 JS

用户明确授权脚本操作本次勾选的多张工作表；未授权宿主文件、网络、进程或其他工作簿访问。当前实现为 `packages/core/src/controlledScript.ts` 的 **JS 子集解释器 → 白名单 AnalysisPlan**，不是原生 VBA / 宏 / 任意 JS 运行器。Acorn 仅解析语法树；解释器没有 `eval`、动态函数、宿主对象或全局对象，也没有任意方法调用。脚本不读取单元格或调用模型，唯一输出接口为 `emit(action)`。

支持 `const`、有限数组的 `for…of`、`if`、数组/对象、基本运算和自身属性读取；`authorizedSheets` 为当前勾选的表名列表。当前动作集是写常量、清空内容、字体、对齐、行高。公式、跨表读取、列宽与结构动作尚未开放给脚本；写值拒绝公式前缀，不能通过公式访问未勾选表。源码、步数、嵌套、动作数、目标格数和输出大小受 `config/capabilities.json.controlledScript` 限制。

共享 `ControlledScriptEditor` 接在 Excel 计划编辑器、WPS 工具模板/计划编辑器。生成只替换 JSON 草稿，丢弃旧结果引用和验收结论；不写入工作簿、不创建恢复点。保存沿原校验路径，执行前仍由各端执行器复核当前工作簿与授权。原生宏无法提供“只能访问勾选工作表”的权限隔离，本轮不调用 `Application.Run`、不修改宏安全设置、不注入 VBProject；这些能力为**未实现**，不能写成已开放。

WPS 的执行前快照现在按每个动作的授权表采集，允许多张勾选表执行与撤销；可选双视图仍只支持单张写入表，多表点击对比时明确拒绝，不影响已执行结果与撤销。Excel 文件夹执行沿既有白名单计划路径，未新增文件夹脚本运行器；WPS 工作目录脚本入口未开放。当前证据是解释器、UI 接线与有状态宿主替身执行/撤销测试；真实 Excel/WPS 的脚本操作仍待验收。SDK 尚未新增自主编写/运行脚本工具，不把手动编辑入口当成 SDK 脚本闭环。

## 2026-10-17：隔离分析完整性与结果尺寸回执

当前审阅顺序以用户最新授权为准：主表执行一次并回读、标记结果，可选双视图展示执行前状态，确认只清标记、放弃真实恢复。下文早期“先复制再写回”为历史方案。

普通后端与 SDK 共用 `python_analysis_tool_definition`，要求分析代码在过滤/去重/关联前保留用户要求的异常及来源行位置，输出汇总、总计和明细；这属于模型行为指导，不是确定性业务完整性证明。

`project_python_analysis_receipt` 现在从可信协调器的 `resultCard` 提取 `resultShape: {rowCount, columnCount}` 给模型。仅成功状态、结果引用匹配且尺寸为正整数时投影，禁止布尔值或字符串尺寸；不透传卡片、表头、值、来源名或任意自由文本。旧回执无卡片时不猜尺寸。模型据此计算目标区域；业务列名已经位于矩阵首行时使用 `includeHeader=false`，不为猜测范围补空行。

这是后端模型工具回执的内部结构，不新增前后端 HTTP/SSE 字段，`contracts.ts` / `models.py` 无协议改动。隐私测试在 `test_python_result_shape.py`，普通后端与 SDK 仍走相同投影入口。通道成功与业务事实验收分开记录，见 [主线状态](SDK_CAPABILITY_ASSESSMENT_2026-10-09.md)。

## 2026-10-10：双端在线连续集成 + 两后端终局结果工具共用（最新）

**本轮只补接线，不扩新功能。** 目标是让**同一次任务**在两端各跑完一条**在线**链路：
真实面板发起 → 真实 HTTP → 真实规划流（SSE）→ `bridgeTool` 指针 → 面板自己的桥读
**有状态宿主替身** → 真实回执 → 服务端登记数据集 → `run_python_analysis` → **真实
Docker 隔离执行器** → 结果登记与引用 → 白名单计划编译 → 面板预览（确认前零写入）
→ 用户确认 → 产品执行器写回 → 产品回读 → 用户撤销 → 核对恢复。

### 缺口一（已修）：普通后端没有终局结果工具

`submit_local_result` / `submit_result_plan` 原先**只注册在 SDK 后端**；普通
（OpenAI 兼容）后端的 `_execute_tool` 没有对应分支，模型调用只会拿到
`{"error": "未知工具：submit_result_plan"}`。抽出**共用语义** `server/app/result_submission.py`
（`resolve_result_ref` 归属校验、`local_result_card`、`build_local_result_answer`、
`build_result_plan_response`、`ResultSubmissionError` 固定码），两个适配器**只注册/分派/返回**，
不再有两套实现。普通后端经 `planning_tool_definitions(base)` 注册（`PLAN_ONLY_TOOLS` 原样返回），
拒绝一律抛 `ValueError`，由既有循环变成模型可见的纠正消息，不静默成功。

### 缺口二（已修）：WPS 并排审阅副本把"授权张数"当成"写入张数"

`stageWpsNativeComparison` 原先要求 `authorization.sheets.length === 1` 且来源也在
同一张表上。两表关联汇总（授权 3 张：两张只读来源 + 一张写入目标）因此**整片被挡在
确认写回之外**，而同一份计划在 Excel 端走的是另一条已存在的显式执行入口 —— 两端分叉。
现在判据改为**"写入只落在一张已授权表上"**，副本只复制**写入目标**那一张表
（`WpsNativeComparison.writeSheet`）。安全性质**没有放宽**：副本仍然只有一张表，绝不
把来源表或任何未写入/未授权的表复制出去；跨表**写入**依旧拒绝；来源的核对发生在
**原表**上（`assertOriginalUnchanged` 会用全部授权序号重读一遍），不需要进副本。

### 一条命令与照实分区

`npm run test:online`（`scripts/run_online_link.mjs`）起**测试专用服务端**
（`server/tests/online_entry.py`：只绑回环、专用端口、临时能力表打开隔离 Python、
`EXCEL_BRO_CONFIG_DIR` 隔离模型目录以免开发者真实连接混入），就绪检查通过后依次跑
`apps/excel-addin/src/onlineLink.test.tsx` 与 `apps/wps-addin/src/wpsOnlineLink.test.tsx`。
两端共用 `config/online_link_scenario.json` 的**同一份数据 / 规则 / 期望**，并都做
"改宿主里键 7 的第一次出现（金卡→银卡）后重跑"的**因果验证**（基线 金卡 400 /
银卡 200 / 未匹配 400 → 改后 金卡 300 / 银卡 300 / 未匹配 400），改动只发生在 TS 宿主
替身上。Docker 不可用时命令打印环境阻断并**非零退出**。

**允许的替身只有两类**：① 固定模型的工具调用序列（只在 `online_entry.py` 注入，生产
没有任何旁路开关）；② 有状态宿主替身（Excel `officeTestHost` / WPS 注入
`getWpsApplication()`）。poll/receipt/result/规划流**不摆 fetch 桩**；结果矩阵、异常
说明、最终计划都不是预制常量；不从旧工件回放；测试不直接给 `range.values` / `Value2`
赋值代替产品写回。

**照实分区（不得合并宣称）**：

| 项 | 状态 |
|---|---|
| ① 工件回放集成（`npm run test:cross-process`、`continuousPythonChain` / `wpsContinuousPythonChain`） | **通过**（保留，如实命名：面板消费服务端算好的结果，面板自己的回执未参与那次计算） |
| ② 双端在线连续集成（`npm run test:online`） | **通过**（真实服务端 + 真实 Docker；模型与宿主对象为替身） |
| ③ 真实 Docker 隔离执行器 | **通过**（本轮在线链路里的计算由它完成） |
| ④ 真实模型自主规划 | **未完成**（本轮用固定工具调用序列） |
| ⑤ 真实 Excel / 真实 WPS 客户端 | **未完成**（本轮用有状态宿主替身） |

### 2026-10-09 实机反馈补修：本地总数与查询能力声明

共用 `dataTools.ts` 在分组结果的 `calculation` 中附上筛选后记录数，取自处理完整输入的计数器，独立于指标类型、分组显示上限。读取有缺失警告时只称“已处理记录数（结果不完整）”。Excel/WPS 共用本地结果卡，无新协议字段；`calculation` 继续不进入模型白名单投影。

SDK `query_table` 的工具 schema 仅声明 `combine.mode=union`，纠正先前错误宣称的 join/deduplicate 能力。handler 防御性拒绝这两种模式，不投递宿主；模型终止且没有有效终局时提供固定能力限制说明。不得自动将关联改成追加合并；文件夹关联实现不受影响。

## 2026-10-09：贯通 Claude Agent SDK 本地工具链（本机工具桥，双端已接线）

**目标与边界**：授权范围确定后，SDK 能在**一次 `query()` 内**真实调用本地表格工具、接收有类型回执、按固定错误码修正参数并继续，最终提交本地结果卡或可预览的白名单写入计划。本轮不开启隔离 Python 产品入口、不做 SDK 会话持久化、不新增 WPS 高级动作、不扩大结果值出站；选表授权、意图阶段无原始行、精确 MCP 白名单、写入先预览再确认全部保持不变。

**为什么需要工具桥**：SDK 在 Python 后端运行，`query_table` 要读的却是任务窗格里**打开且可能未保存**的工作簿 —— Python 的 MCP handler 读不到它。因此新增**易失任务工具桥**：SDK handler 提交查询请求 → 桥登记 operation 并按所有权绑定 → 拥有本任务的宿主适配器（Office.js / WPS）在本机执行**现有确定性 `query_table`** → 经本机回环端点回执 → handler 投影后有类型地交回 SDK，同一次 `query()` 继续。

**落点与职责**：

| 区域 | 文件 | 职责 |
|---|---|---|
| 共享工具目录/投影 | `server/app/agent_tools.py` | 两后端共用的本地查询工具定义（`describe_authorized_tables`、`query_table`）、本地结果终局工具（`submit_local_result`、`submit_result_plan`）、按宿主裁剪 `submit_plan`、**有类型出站投影白名单** |
| 结果登记（结果生命周期） | `server/app/local_results.py` | 值矩阵只留可信侧的按任务登记器；句柄不透明、按**归属 claim** 核对、TTL 失效、来源变更失效、保留上限。**独立于会话生命周期** —— 会话结束不删结果 |
| 任务工具桥（会话生命周期） | `server/app/task_bridge.py` | 会话与所有权绑定、幂等（同 ID 同参幂等／同 ID 不同参拒绝）、取消与迟到回执拒绝、来源指纹不一致→`source_changed`、重授权即 `supersede`；`close()` 只结束会话 |
| 账本隐私投影 | `server/app/ledger_privacy.py` | 按字段名剥离矩阵（`values`/`rows`/`headers`/`formulas`/`expected`…），作用于 `actions` 与 `acceptanceCriteria`；持久账本只存去值投影，回环 SSE 仍交付完整值 |
| 结果→计划编译 | `server/app/result_plan.py` | 只生成 `writeValues` 单动作 + 自动 `rangeEquals` 回读 + `localResultRef` 引用；校验矩阵类型/尺寸、目标授权、WPS 格数上限、疑似公式文本 |
| SDK 适配 | `server/app/claude_sdk_agent.py` | 注册本地工具、经桥等待回执、终局装配 `AnswerResponse.localResult` / `PlanResponse`，`finally` 只关会话 |
| 本机回环端点 | `server/app/main.py` | `/api/sdk/tool-bridge/{poll,receipt,register,result}`；回环 + Host + Origin + JSON 严格守卫。`register` 要求**活跃且归属一致**的会话；`result` 只按**授权归属**核对（会话可已关闭） |
| 共享前端状态机与复核判据 | `packages/core/src/sdkToolBridge.ts` | 指针→轮询、待回执跟踪、发请求前判定 accepted/duplicate/conflict/late；`bridgeRequestedSheetNames` / `verifyBridgeAuthorization`（**两端同一份判据**） |
| 双端宿主适配器 | `apps/excel-addin/src/sdkBridgeHost.ts`、`apps/wps-addin/src/wpsSdkBridgeHost.ts` | 组装 `poll`/`verify`/`execute`/`receipt` 四端口：执行前重读宿主现场 → 授权复核 → 真实读取（Office.js `usedRange` / JSAPI `authorizeWpsSheets`+`readAuthorizedWpsSheets`）→ **既有 `executeQueryTableData`** 计算 |
| 双端面板接线 | `apps/excel-addin/src/hooks/useSdkToolBridge.ts` + `components/LocalResultCard.tsx`；`apps/wps-addin/src/WpsAgentPanel.tsx` | 消费 `bridgeTool` 指针、串行 drain（重入会让同一次调用被执行两次）、取消即停、`end()` 保留绑定供终局取值；结果卡按 `resultRef` 取回值渲染 |
| 协议 | `server/app/models.py` ↔ `apps/excel-addin/src/contracts.ts` | `Bridge*` DTO、`PlanRequest.localResultRefs`、`AnswerResponse.localResult`、`TaskProgressEvent.bridgeTool`、`WorkbookSnapshot.workbookIdentity`、`AnalysisPlan.localResultRef`、**`AnalysisPlan.sourceDependencySheets`（第五轮新增）** |
| **实际来源依赖**（第五轮） | `apps/excel-addin/src/planSourceDependency.ts` | 授权范围 / **实际来源依赖** / 写入目标三口径分离。记账**唯一入口**是 `sdkBridgeHost.execute` 里确定性执行器返回之后（取 `DataToolResult.sourceSheets`）；计划自带的 `sourceDependencySheets` 一律被 `attestSourceDependency` 覆盖或剥离；**没有记录就删字段**（不是空数组） |

**三段生命周期分别建模**（混在一起是三类真实缺陷的来源）：

| 生命周期 | 载体 | 结束时机 | 关键纪律 |
|---|---|---|---|
| 授权 | `OwnerClaim` + `authorizationRevision` | 重授权 / 换簿 / 授权范围变化 | 重授权只推进版本并作废旧结果；不编造澄清答案 |
| 会话 | `ToolBridgeSession` | 一轮规划结束（`close()`） | **只结束会话，不删结果**；否则终局 `resultRef` 取不到值 |
| 结果 | `RegisteredResult`（`resultRef`） | 自己的 TTL / 保留上限 / 来源变化 | 跨轮按**授权归属**取值；有值计划**不进**持久账本 |

**执行前复核**（`verifyBridgeAuthorization`）三类变化一律拒绝读宿主：换簿（工作簿身份不同）、源数据变化（指纹不同）、**授权收窄**（绑定时的授权范围不再被当前授权完整覆盖）。最后一条最容易漏：只看"当前授权非空"会放行一个少读一张表、静默算出错答案的查询。绑定时的范围快照取自**与复核同一个读取器**，避免两份对不上。

> **第五轮的口径收紧**：绑定时的范围快照取自同一个读取器这条纪律，同样适用于**身份与来源指纹**。
> 曾经 Excel 侧只把范围接对了、身份与指纹仍从快照取 —— 身份是死字段（三个捕获点都没写过它，只配到
> `name:<文件名>` 兜底，而复核读 `url:`），指纹连算法都不同（快照 `fnv1a32:` vs 复核 `fnv1a32x2:`），
> 于是**已保存的**真实工作簿上工具桥恒判 `source_changed` 拒读。现在 Excel 侧先读一次现场作为**本轮基线**，
> 并把它同时喂给规划请求、桥绑定与计划指纹（WPS 侧本来就是这么接的）。

**通道分离**：进度事件只带**不透明指针** `bridgeTool{operationId,tool}`（会落账本）；请求参数只活在本机易失桥里，由宿主适配器经回环 `poll` 拉取；值矩阵只在回环 `receipt` 上行与 `result` 下行出现。**模型可见回执**只含固定 `status`、不透明 `resultRef`、允许的结构（headers/rowCount/columnCount/complete/truncated）与固定 `errorCode`；`rows`、`title`、`calculation`、`warnings`、`scannedRows` 与原始异常**一律不出站**。

**两端接线**：`apps/excel-addin` 与 `apps/wps-addin` 都消费 `bridgeTool` 指针 → 轮询 → 执行前复核当前授权 → 真实读取 → **既有确定性 `query_table`** → 回执 → 结果卡。计算两端**同源**（都调 `executeQueryTableData`），只有"怎么读宿主"各走本宿主；绑定、状态机、判据、结果卡共用一份实现。WPS 的 `turnId` 由服务端 `checkIntent` 生成，客户端本轮开始前拿不到，故绑定**惰性**建立并按 `turnId` 判重（重建会清空 `delivered`，导致同一次调用被执行两次）。

**结果卡**：只读任务以 `AnswerResponse.localResult`（引用 + 结构，**不含值**）结束；本地渲染器经回环 `result` 端点按引用取回值，取不到时**如实**说"不可用"，不渲染空表冒充成功。`title` 用**固定标签**，不回显可能来自源数据的自由文本。

**写入仍走老链路**：`submit_result_plan` 只做**本地校验与编译**，产出 `writeValues` 白名单计划后**进入预览**；用户确认、写前复核、宿主执行、写后回读一律不变。SDK 工具**不得**直接修改工作簿。

**"展示结果"不等于"完成任务"（2026-10-11，缺陷 3）**：曾出现真实 SDK 模型（`turn-f493dba5ae61470d99a54dd7a933e53d`）读完表、跑完隔离分析、拿到 40×7 结果后**直接 `submit_local_result` 收尾**，`submit_result_plan` 一次都没调 ⇒ 面板 `pendingPlan` 空 ⇒ 确认写回卡根本不渲染。根因**不是**"模型不肯调工具"：全仓并无"本轮规划目标是否达成"的判据，`AnswerResponse` 被无条件当终局（`excel_agent.py` 主循环 `return parsed`），而 SDK 侧 `terminal` 一旦非空即**硬锁后续所有工具**（`claude_sdk_agent.py`）—— 模型即使察觉也**不可能**补救。

修法落在**共享层**（`result_submission.py`，普通后端与 SDK 走同一判据）：

- 意图阶段模型返回 `planningGoal`，两端传入 `PlanRequest`；服务端按同一 `turnId` 保存并优先采用意图决定。目标仍由模型理解，运行时不按行业字段或关键词推断。
- 已绑定 `write_preview` 的任务，`submit_local_result`、`submit_answer` 和 JSON 回答退路均不能直接结束规划；终局工具重新声明 `read_only` 也不能降级目标。普通后端在本轮循环内纠正；SDK 工具错误不置终局，允许同一次 query 继续提交计划。
- `read_only` 可正常展示结果；写入仍必须是白名单计划，经预览和用户确认。旧的直接调用方若不提供目标，保留原兼容语义，不能宣称这些调用也有事前目标约束。
- `planningGoal` 是任务完成判据，不是写入授权；工作表授权、来源校验和宿主预检继续独立执行。
- 测试：`test_bound_planning_goal.py` 覆盖双后端提前回答、终局降级、SDK 文本退路及同轮修正。

同时修掉一处**真实次生缺陷**：SDK 的 `submit_result_plan` 在编译被拒（尺寸不符 / 未授权 / 来源已变）时回执**缺 `is_error`**，看起来像成功 —— 模型会据此收尾并谎称完成。

**空白只有一种写法（2026-10-11）**：分析产出的 `""` 在**编译层**（`result_plan._blank_cell_of_write`）折成 `None`（写入意义上的空白 = 空单元格）。断言侧也必须同口径比较（`_matrix_of` / `foldBlank`），否则会把这次**正确的折叠**误判成"回读与计划不一致"。**只折结果矩阵**：来源数据的 `memberRows` / `orderRows` 一概不动。

**能力收窄为空集时不得产出空 `anyOf`（2026-10-11，缺陷 3 的次生缺陷）**：修复缺陷 3 后真实 SDK 驱动仍然每次 502 `MODEL_HTTP_ERROR`，探针取到 SDK 原始异常才看清：网关（`https://api.deepseek.com/anthropic`）直接拒绝**整轮**请求 —— `API Error: 400 Invalid schema for function 'mcp__excel_bro__submit_plan': [] has less than 1 item`。触发点是 `agent_tools.host_pruned_terminal_definitions` 的 `{"anyOf": [variants[name] for name in selected]}`：`selected = 服务端上界 ∩ 客户端自报`，而客户端漏报 `selfReportedActions`、`hostVersion` 不匹配、宿主开关关闭或宿主不可识别时这个交集都是**空集**，于是产出**空 `anyOf`**。

空 `anyOf` 是非法 JSON Schema，会让网关拒绝整轮请求。当前修法是：协商动作集合为空时省略 `submit_plan`，保留回答、澄清等工具；有动作时只暴露交集。不能用回落全量动作掩盖协商失败。提交校验和执行预检继续独立存在。`_mutate_action_schema_legality.py` 三项变异均红，还原后 15 项通过。

**本轮未完成 / 宿主限制（如实记录）**：

- WPS 侧**没有可用的源内容指纹**：复核只按工作簿身份 + 授权工作表集合判变化（`sourceFingerprint` 置 `null` = 不比对），不声称已覆盖"内容层面"的改动检测。
- 未保存的 WPS 工作簿拿不到磁盘路径，身份退回 `name:<工作簿名>`：**两个同名未保存工作簿无法区分**（宿主不给稳定标识）。
- ~~`submit_result_plan` 的"预览→确认→写前复核→执行回读"在面板上的端到端点击路径无自动化用例~~ —— 2026-10-10 已由双端在线连续集成补上（见文首最新一节：Excel 与 WPS 各有一条真实点击路径，含撤销；但模型与宿主对象仍是替身）。
- **2026-10-11 并发诊断更正**：旧 A/B 脚本把 SDK 与健康检查放在不同进程，且替换了结果处理；日志只能证明曾停滞，不能证明“任何 HTTP 都冻死”或“环境问题而非代码”。`scripts/diagnose_sdk_http.py` 通过生产 FastAPI/SSE 路径、真实 SDK、真实模型复验：两轮短任务分别成功处理 11 / 10 次并发健康检查；真实 Python 工具链在 47.427 秒完成，115 次健康检查成功，桥回执和结果查询均 200，最终调用 `submit_result_plan`。宿主为合成数据替身，未操作实际工作簿。报告位于 `outputs/sdk-python-http-audit-20261011/report.json`。不据此声称所有并发或取消场景均无问题。
- 真实模型 + 真实 Docker + 真实 HTTP 的合成宿主链路已验证；本轮没有新增真实 Excel/WPS 客户端点击验收。仓库默认 Python 开关恢复 false；真实诊断通过独立能力配置启用，不修改全局默认值。验证不代表发布。



## 2026-10-08：双后端任务上下文与本地澄清整合

Excel/WPS 从已有任务记录投影同源 `taskContext`（原需求与明确提交的业务决定），普通意图/规划与 SDK 意图/规划都接收；同题后答替换先答，新需求不继承旧规则，显式替换同时重置恢复记录。上下文不含本地源值/目标矩阵/授权，不等于 CLI 会话或写入确认。TS/Python 严格协议同步，限额集中 taskContext；SDK 仍独立 query、禁止数据行及任意工具，无精确检查点恢复。

共享本地补全卡沿用完整无歧义的显式等级，只在新目标/排序缺口时要求确认；非法输入逐字段提示，提交前阻断。补全后仍重读授权源列、生成预览、独立确认写回。WPS 文件夹与隔离 Python 产品接线不在本批；自动化与实机分列。需重启后端并刷新两端，见 [实施及试用清单](./AGENT_TASK_INTEGRATION_2026-10-08.md)。

## 2026-10-08：手动规划暂停与原任务恢复

**交互补修（覆盖首批手动核对方案）**：共享过程卡与输入区统一称“暂停”，已有 `turnId` 时自动查询一次原任务，不自动 POST 继续。原需求留在用户消息/任务上下文，不回填普通发送框，也不覆盖新草稿。可续检查点由用户明确继续；仍运行/查询失败保留核对入口，无检查点明确需要重新开始。WPS 旧收尾不清除新操作进度，Excel 查询/继续有同步并发闸；不新增轮询、协议字段、模型调用或写入权限。无任务身份时不能编造续跑现场；双端真实宿主时序待验。

双端手动暂停保留原任务身份，状态尚未核实只给查询入口；服务端确认检查点后才由用户继续。中断接收不代表服务端立即暂停，运行中/未知状态不得重发需求。WPS 普通继续省略模型切换字段，事件适配保留 `turnId/sequence`，历史只记客户端实际游标；未知游标从零补收。继续/取回结果复核原授权，不扩展写入、文件夹或 Python 能力。SDK 无可用检查点时不能精确续跑，双宿主实机尚待验收；见 [实施记录](./MANUAL_PLANNING_PAUSE_2026-10-08.md)。

## 2026-10-08：统一数据范围入口与选中即校验（最新）

正式入口的 `TargetScopeNotice` 仅有“检查数据范围”按钮，不提供清除提示的“暂不操作”或独立重试。提醒未处理时保留授权阻断与待续分析；打开选择器不是授权，确认前仍不能读取字段或写入。文件/表真实变化展示原范围与当前范围（仅名称，不展示私有身份）；无法核实仍不声称已变化。本补修不改变守卫触发条件，无范围入口的独立 WPS 兼容回退保持原行为。

选择器摘要仅显示简短状态与必要的具体限制，移除重复“重新检查”按钮及授权说明；原“刷新”同时重读目录和重新校验，同名称/会话也重查。Excel/WPS 刷新只调整自动草稿与候选显示，不提交授权、不静默删除手动缺失项。WPS 当前表两处显示共用元数据快照；Excel 表格分析禁止将新增手动草稿借发送自动授权。未新增持续轮询；确认/发送仍复核真实宿主。

实际 Excel/WPS 窗格不再叠加第二个“原授权 / 当前文件”选择弹窗。正常追问复用有效授权；目标变化或证据不足时，在数据范围入口旁显示非模态状态，打开现有选择器后隐藏重复提示。只有这个选择器提交授权；业务补全仍不夹带授权。独立嵌入且没有范围入口的 WPS 面板保留旧确认组件作为兼容回退，不用于正式宿主入口。

两端复用 `useScopeAdmission`：勾选自动检查元数据，检查中/不可用项禁止提交；只读、受保护或宿主写入能力未知时可授权分析，但不冒充可写。确认前重新检查，并与可见目录的私有文件身份及上一次检查结果绑定；换文件、删表、超时、畸形结果与迟到响应不能静默缩小选择或授予新范围。Office 仅读取名称/稳定 ID、文档身份及受支持的 readOnly/protection 属性；WPS 仅检查身份、表目录/尺寸和保护状态，不读取选区值。功能限额集中在 `scopeAdmission`。

文件夹模式新增本机 `POST /api/folders/admission`，仅接受已登记 sessionId/fileId 与选中表名；检查文件存在、根目录边界、读取/写回访问、工作表存在与保护状态，返回封闭枚举和内容/选择指纹，不发布快照授权、不创建备份、不执行模型或写入。XML 只解析元数据属性，文本不进入业务数据对象；内容哈希会读取文件字节，不等于提取/上传数据行。该预检并非永久权限，执行前来源指纹、授权及写入能力仍独立复核；openpyxl 不自动执行宿主保护规则，因此已有受保护目标的写入在预检中明确拒绝，拆分汇总对源表的只读使用不因此禁止。

`confirmScope(..., {continuePendingAnalysis:true})` 只允许续接由分析入口登记的一次新请求，保留输入/答案；未登记的旧写入不得排队或重放。WPS 文件夹入口仍未实现。详见 [实施与实机清单](./SINGLE_SCOPE_SELECTOR_2026-10-08.md)。此节覆盖此前独立目标弹窗的正式入口描述；模型协议未扩权，TS/Python 的文件夹预检协议同步更新。需重启后端并刷新两端。

## 2026-10-08：当前表确认、未知身份与真实变化分离

双端共用授权交互新增本地 `confirmation_required`，不再把一致弱证据的显式同意伪装成 `unavailable`。首次需要确认只显示一份文件/表范围；临时缺失有限后台重试，耗尽后只提供重新检查，身份冲突立即阻断，不能用第三份证据缺失掩盖前两份冲突。真实文件/表变化仍独立选择目标。已有确认不因每条消息重新询问，重检查不重放旧写入。状态仅在内存，不改模型/服务端协议，写入审批仍独立。自动化与实机分列，见 [实施记录](./TARGET_CONFIRMATION_STATES_2026-10-08.md)。

## 2026-10-08：Excel 文件夹范围安全修复

文件夹快照对全部选中文件做读取前后内容指纹检查，成功后才发布授权；失败不把旧行绑定到新文件指纹，执行前仍独立复核。澄清使用共享 `scopeFingerprint`，本地绑定绝对目录根路径与已选表名，不把目录路径加进模型/API 协议。同根/同会话/已选目标均存在的候选刷新保留原确认，但不授予草稿授权；模式切换不凭保存的勾选自动授权。目录文件转当前工作簿时先只读名称目录，再仅读明确选择的表。

本批不新增 WPS 文件夹入口；共享后端与纯函数可复用，WPS 文件夹模式仍未实现。前后指纹不是外部进程锁或原子文件系统快照；自动化与实机证据分列，见 [修复记录](./FOLDER_SCOPE_SAFETY_FIX_2026-10-08.md)。

## 2026-10-08：产品身份与内部阶段分离

消息理解、字段意图与规划提示共用 `assistant_identity.py` 的产品身份：对用户是 Excel Bro AI 表格助手，不自称路由器或处理环节。底层模型只在用户询问时引用服务端选中连接的 `model` 配置标识；不把 SDK/兼容协议等同于模型，不保证网关背后的厂商或实际模型。普通后端与 SDK 同源；身份提示不增加工具、数据访问或写入权限。

双端普通聊天成功完成后过程卡默认收起，可手动展开；运行中、失败、警告和等待确认保留现有可见性。`ActivityLog.defaultCollapsed` 只是本地展示标记，不是服务端协议或授权信息。详见 [实施记录](./ASSISTANT_IDENTITY_2026-10-08.md)。

## 2026-10-08：消息理解先于字段读取

双宿主的新消息先调用独立 `/api/intent/message`，只发送当前消息、限长对话和图片数量；不自动附带工作簿、字段、priorResult 或 priorIntent。`conversation` 直接回复，`needs_workbook_context` 才复核目标授权、读取字段并进入既有字段意图/规划。第一阶段失败或截断不默认读表。普通模型/SDK 同源提示和协议，SDK 权限未扩大。Excel 启动和范围刷新只读目录元数据；文件夹发送也先消息理解，用户主动文件扫描保持不变。

授权前允许的只是这次**无工作簿上下文**的理解请求，不授予数据访问；表头/数据行/字段意图/规划/写入仍必须在授权复核后。后台元数据监测和既有业务补全/执行守卫保留，写入审批独立。当前接口和进度顺序覆盖此前“所有意图模型都在授权后”的描述；旧批次测试是历史证据。详情见 [实施记录](./MESSAGE_FIRST_ROUTING_2026-10-08.md)，需重启后端并刷新两端。

## 2026-10-08：已选范围与后续追问使用同一授权

Excel/WPS 的范围确认接入共享 `useAuthorizedTarget.confirmScope`：点击时只重读文件与表身份元数据和宿主证据，核实后一起提交已选范围、授权基线及当前会话的分析同意；不读表头/数据行、不调用模型、不重放旧请求或写入。每条追问仍走 `checkForAnalysis`，比对原授权与当前元数据、宿主证据。WPS 不再因普通消息开始而撤销弱包装证据的显式同意；同意只绑定当前已确认的文件/表范围，检测到变化、证据冲突、撤销、关闭或重载后失效。写入执行器不消费这个分析许可，白名单计划、预览和独立写入确认不变。文件夹模式与模型协议未改。

无法核实身份与确实换文件采用不同提示；不能用同名、同表头、隐藏弹窗或直接放行来解决重复确认。实现、自动化证据与双宿主实机步骤见 [修复记录](./SCOPE_AUTHORIZATION_FOLLOWUP_FIX_2026-10-08.md)。

## 2026-10-07：易失结果登记与双端本地预览准备（最新状态，产品仍关闭）

内部 `run_registered_result_internal` 把真实容器的严格矩阵结果交给同一易失授权登记器，发布随机 resultHandle、原始 JSON 字节 SHA256 和矩阵尺寸，不返回原始值。结果绑定原授权及有序数据句柄/指纹；替换源快照删除依赖结果，撤销/重授权/任务 TTL/服务重启失效。数量和编码总字节限额在 `pythonResult`。没有持久结果库或模型工具成功通道。

新增本机 `/api/python/data/result` 与 `/result/discard`，复用回环、Host、Origin/hostId、总开关、严格 JSON 和有界请求守卫，固定错误文案、no-store。协议同步 core/Excel contracts/models.py，不扩展 PlanRequest。core 在本机读取并核对实际响应字节指纹，读取前后重读源值；结果过期不冒充文件变化。仅接受有界、规则、有限安全标量矩阵；共享编译器只生成已有 `writeValues` + `rangeEquals`，目标由可信用户流程给定，限定已登记授权表的精确 A1 范围。拒绝疑似公式/宿主文本转义前缀，不让输出携带坐标/脚本/动作。

Office/WPS 已有只读预览准备器：核对原数据会话绑定的宿主、源快照、计划不可变及目标原值/公式；Office 调用本地 preflight 并按授权 sheet ID 读取目标，WPS 复用 prepareWpsPlan。均不调用执行器、原生副本或持久账本。Office 完整可写性检查/原生副本、WPS 原生副本及两端确认写回仍需产品接线；不得直接把含结果矩阵的计划传给旧远端 plan_validate、聊天历史或持久执行/撤销记录。详细双端差距、证据与下一批门槛见 [实施记录](./PYTHON_RESULT_PREVIEW_2026-10-07.md)。`pythonExecution.enabled=false`；SDK 内部转录限制、文件夹适配及安装包未实现。

以下为前一阶段快照，当前分层状态以上节为准。

## 2026-10-07：内核门禁与授权数据接线（前一阶段，产品仍关闭）

worker 在生成代码之前安装不可撤销的 Linux/amd64 seccomp 附加过滤，拒绝 exec/派生进程/线程/socket，安装失败不执行。真实新测试 5 项通过，原有 15 项保持通过。独立 `/api/python/data/{scope,snapshot,revoke}` 仅接收已允许的本机加载项请求；易失登记绑定任务/文件 hash/授权版本/精确区域，随机句柄、撤销/TTL/大小限制，不进模型、历史或账本。两端使用同一 core 客户端，各自复用宿主强身份与只读原语；新协议未扩展 PlanRequest。

本批仅接通内部登记数据→容器分析。产品工具、结果库/AnalysisPlan 编译、双端 UI/生命周期、文件夹入口及安装包未接线；`pythonExecution.enabled=false`。SDK 内部转录限制仍存在。新镜像 RepoDigest 和当前能力证据见 [授权数据实施记录](./PYTHON_AUTHORIZED_DATA_2026-10-07.md)。下节是此前运行时阶段快照，其中“禁 exec/句柄未实现”不再作为当前状态。

## 2026-10-07：Docker 本机运行时验收（产品入口仍关闭）

用户启动 Docker Desktop 并授权构建后，本机 Linux 引擎、专用镜像构建及 15 项真实合成数据测试已跑通。当前不再处于“本机引擎不可用”的阻断状态；旧记录仅代表此前阶段，不据此推断 WSL 或代理策略的因果。实际镜像身份与证据见 [本机验收记录](./DOCKER_PYTHON_REAL_VALIDATION_2026-10-07.md)。

新增可信构建入口 `scripts/build_python_sandbox.py`：仅把 `.dockerignore`、`docker/Dockerfile`、`app/docker_worker.py` 三个明确文件打包为 stdin tar；不遍历/发送仓库。使用临时无凭据 Docker CLI 配置，显式固定本机端点与 daemon-bound `default`/`docker` builder，拒绝其它 driver、多节点或远端 endpoint；基础镜像先解析真实 RepoDigest 再构建，前后核对引擎身份。构建镜像不启用产品功能，不改用户全局 context，不用本地 image ID 冒充 RepoDigest。

运行器新增固定容器线程环境（OpenBLAS/OpenMP/MKL/NumExpr 各为 1），使 numpy/pandas 在既有单进程限额内可用，不扩大 PID 配额；输出读取用 `read1` 及时交付短输出，不等满 4096 字节或 EOF。已用真实管道先红后绿及容器内运行/取消验证。

`pythonExecution.enabled=false` 保持不变。完全禁止容器内 exec、datasetHandle 授权登记、结果句柄、结果编译到 `AnalysisPlan`、Excel/WPS/文件夹入口及安装包仍未实现；SDK 内部转录投影限制仍存在。15 项通过不代表完整产品 Python 或全部威胁边界已经完成。

## 2026-10-06：隔离 Python 分析的设计授权（历史阶段记录）

以下“当前状态/环境阻断”描述为当时快照；2026-10-07 构建及真实验证的最新状态以本节之前的本机运行时验收为准。

用户批准的例外是模型可编写 Python，在经过验证的隔离执行器中分析授权数据，不是开放 SDK Bash/Read 或宿主任意脚本。模型选用普通后端或 SDK 不改变数据/写入授权；共用受控工具与结果协议，Excel/WPS 通过各自宿主适配读取当前授权范围，文件夹模式通过明确的文件/工作表 ID 读取。不默认复制整表给模型或远程服务。

计划新增 `run_python_analysis`：输入任务内数据句柄与代码，隔离工作进程无网络、宿主目录/工作簿句柄、凭据和外部程序权限；有资源配额、取消及整个执行树终止。隔离缺失则拒绝，不能用 cwd、venv、AST 检查或禁词当沙箱，也不能退回同用户裸执行。输出矩阵只在本地，模型出口由可信投影器控制；代码产生的摘要、stdout、错误均不能默认发送。结果经本地适配编译到既有 `AnalysisPlan`、预览、确认、回读链路，Python 本身不写工作簿。

当前状态（2026-10-07 八批）：失败关闭骨架与**隐私/入参闭环已落地并补修多轮绕过路径**，并已新增 **Docker 运行时适配器（代码落地并修复剩余缺陷，真实隔离实测受环境阻断）**。受控工具 `run_python_analysis` 由 `python_analysis.py` 失败关闭协调器承载，出站走 `NON_TERMINAL_TOOL_PROJECTIONS` 有类型投影。多批形成的隐私边界：① 历史投影是可信侧最小白名单（`code`→固定占位 + `datasetHandles`→空列表，句柄登记未实现故不保留模型自报的任何句柄原值）；② 恢复现场统一投影与归属核对（孤立 ID/名称冲突/关联歧义失败关闭拒绝恢复）；③ 执行路径异常统一安全映射（普通后端与 SDK 同源固定回执，不回显原始异常）；④ loopState 收口（涉及 Python 时清空错误/警告自由文本）。**运行时层**：新增三层分离（协调器 `python_analysis.py` → 可替换运行时适配器 `python_docker_runner.py` → 容器内固定工作程序 `docker_worker.py`）。Docker 是第一种运行时实现，不是项目强制启动依赖；适配器可替换（未来可加 AppContainer 等）。固定启动链由可信代码决定（`--network none`/只读根/非 root/`--cap-drop ALL`/禁提权/CPU·内存·swap·pids 限额/`--log-driver none`），模型只提供 code 经 stdin（绝不进命令行、不用 pickle、不拼 shell、不挂载用户目录/仓库/工作簿/凭据/socket）。**剩余缺陷已修复并各配红灯测试与独立反向变异（第八批 B1–B6）**：① **输出运行期有界**（不是"有界临时文件"）——stdout/stderr 由读取线程并发收集、分通道限字节，任一通道触顶即**中途终止执行**并进入容器终止核实，协议结果/纯 stdout/stderr 预算分离，保留 `--log-driver none`；② **清理可核实 + 真实取消**（不是"清理闭环已完成"）——绑定可信侧运行身份与精确容器身份，`rm -f` 后 `container inspect` 区分"确认已消失"/"无法查询"，清理失败或状态未知**不冒充成功**（不发布结果、不复用沙箱），解决启动/取消竞态，取消后到达的迟到结果被拒收，只清理本次运行拥有的容器（禁止全局 prune/宽前缀删除）；③ **核实并锁定真实本地端点**（不是"剔除变量即绑定本机"，也不是"判 scheme 前缀"）——`docker context inspect` + `docker info` 实测端点，只接受**经结构校验的本机形态**（Windows 命名管道 `//./pipe/<name>` 或扩展长度前缀 `//?/pipe/<name>`；Unix socket 本机绝对路径），拒绝 `tcp://`/`ssh://`/`fd`/未知、远端主机段命名管道（如 `npipe:////remote-host/pipe/...`）、带远端 authority 的 unix socket、空/畸形地址，**不把任意主机别名改写为 local**；整轮锁定到已核实端点，不改用户全局 context，不把 `.OSType=linux` 当本机身份证明；端点解析取结构化字段 `Endpoints.docker.Host`（不能只取第一个 URL 匹配，否则会误取 otel 端点的 `unix://`）；执行路径不按对象名信任 `VerifiedEndpoint`，入口与 CLI 最内层各做一次结构校验；**运行期端点核实 ≠ builder 核实**（buildx 远端 builder 由构建流程单独核实，不因 runner 的 `-H` 锁定而视为安全）；④ **收紧配置/镜像/结果校验**——非 root 校验拒绝 `0:1000`/`root:1000`/`00:0` 等，移除错误的 `maxProcesses + 1` 接线，CPU 速率与累计 CPU 时间分开实现（不可实现则失败关闭），PID 限额不再冒充"禁止 exec"，结果字节/类型/深度/元素数边界 + 拒 NaN/Infinity/畸形 UTF-8/多包协议，区分合法 `result=null` 与缺失 `result`，绑定已核实不可变镜像身份（`imageDigest=null` 时失败关闭，不静默降级为可变 tag）；⑤ **最小化构建上下文**——`server/.dockerignore` 白名单，排除 `.env`/`model-connections.json`/凭据/测试/缓存/账本/工作簿/无关源码，核实真正进入上下文的文件集；⑥ **修真实验收测试的假通过路径**——网络隔离配**可达性正向对照**（不接受"Python 还活着"当对照）、宿主哨兵先真创建再核对、凭据先注入合成哨兵再验证未被继承、资源限额测试须证明探针已启动并给出限额证据、容器查询失败不得返回空列表而冒充"无残留"、真实测试中的实现异常必须失败而非被宽 `except` 吞成 skip。**边界如实记录**：普通后端模型/持久化/异常出口已审查收口；**SDK CLI 托管内部转录无法投影**，不宣称双后端同等隐私闭环；Docker 日志不作为代码/数据/输出的持久化出口（`--log-driver none` 为环境强制手段，非仅文档承诺）。**真实隔离实测受环境阻断**：本机 Docker 客户端存在但无可用 Linux 容器引擎（Docker Desktop 走 WSL2 后端，`wsl.exe` 被代理黑名单拦截）——mock 层已全绿，真实用例整体 skip（明确标注环境阻断，未跑不等于通过）。**代码层面的修复与判别力证据不等于安全闭环完成**：隔离执行器真实隔离证据、句柄登记、结果转 `AnalysisPlan`、双端入口、镜像构建摘要均未完成，`pythonExecution.enabled` 保持 `false`。下面 SDK"受限表结构通道"说明仍是当前实现状态。完整威胁边界、分阶段实施、双端与真实隔离验收见 [隔离 Python 设计](./SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md)，实施交接见 [DeepSeek 提示词](./DEEPSEEK_ISOLATED_PYTHON_PROMPT_2026-10-06.md)。

## 项目原则：新增功能必须双宿主对齐（2026-10-05）

Excel/WPS 同等用户任务能力是新功能的完成条件，不只是未来目标。新增功能从设计起同时列出两端入口、语义、授权/隐私、执行前确认、结果核验、异常与恢复、窄窗交互及测试证据。共享语义/协议/确定性计算优先复用，宿主差异通过适配层处理，不要求调用同一个 API，也不允许为了表面一致扩大授权或降低验证。双端自动化和实机证据分别报告；一端未实现时不能标整体完成。确有宿主限制或用户明确批准分批时，计划表保留具体未对齐项和验收条件，并区分未实现、宿主限制、实机待验。

审阅窗格比例已双端接线：**唯一参照是整个显示器可用工作区**，默认 **副本 35% / 原表+AI 65% / AI 30%（占整个可用宽度）**，比例/手动优先共享，尺寸获取与宽度恢复由宿主适配。两端自动化已通过；真实宿主截图验收仍独立待验，不能称所有版本已验收。

## 2026-10-14：审阅布局改为「按整个工作区 35/35/30」（修真实比例缺陷）

**根因（真实缺陷，不是文案问题）**：AI 比例一直按**整个工作区宽度**在 `packages/core` 计算，但两端喂进去的"宿主宽度"口径不一致 —— Excel 传的是**原表窗口内容区宽度**（`clientWidthCssPx`，是原表窗口自己的宽度，不是工作区），WPS 传的是 `Application.Width`（含两张表的主窗口）。于是同一份 `reviewAiWidthFraction`（当时 0.4）在两端代表**不同含义**，Excel 侧实际得到的是"占原表窗口若干比例"的 AI 窗格 —— 用户截图的 40/24/36 与"AI 约占右侧窗口 60%（36/60）、原表只剩 40%（24/60）"一致，说明 AI 拿走的是右侧窗口的大头，而不是"整个可用宽度的 30%"。

**修法**：一个外侧比例 + 一个内侧比例，且内侧比例只有一个参照。
- 外侧：`reviewPaneSizing.reviewCopyWidthFraction`（默认 0.35），服务端 `position_pair` 取 `1 − 该值` = 0.65；删除 `excelWindowLayout.originalWidthFraction`。
- 内侧：`reviewPaneSizing.reviewAiWidthFraction`（默认 0.3），`reviewPaneTargetWidth` / `reviewLayoutPlan` 都作用于**整个工作区宽度**。`reviewLayoutPlan(W)` = 副本 `round(0.35W)`、原表 `W − 副本`、AI `clamp(0.3W, 320, ⌊原表/2⌋)`、原表可见表格 = 原表 − AI；`degraded` 在原表可见宽 < 280 或 AI < 320 时为真。
- 参照来源：Excel 本地桥 `measure` 现在返回两个字段 —— `clientWidthCssPx`（诊断用）与 **`workAreaWidthCssPx`（AI 比例唯一参照）**，后者由 `work_area_width_css_px` 用 per-monitor 线程 DPI 上下文读 `MonitorFromWindow(..., MONITOR_DEFAULTTONEAREST)` + `GetMonitorInfo().Work`，再按 `GetDpiForWindow` 归一到 96-DPI 逻辑像素；WPS 仍是 `Application.Width`（磅 → ×96/72）。前端只接受**两个字段齐备且合法**的测量回执，缺工作区宽度宁可降级为手动，也不拿客户区宽度顶替。

**用户入口与降级**：Excel 与 WPS 都在副本审阅卡上提供「重新排列」次按钮，**只有用户主动点击**才重跑默认布局。`rearm` 是一个只由该点击递增的计数器，作为 AI 窗格自适应状态机的额外依赖，用来清掉"用户已手动拖过"的标记并重套默认比例；消息更新、流式输出、回读、保存工具**都不改** `rearm`，因此不会反复抢回布局。宿主不能由代码设置 AI 窗格宽度时（Excel 缺 `TaskPaneApi 1.1`、WPS 的 `Window` 没有可写宽度成员且 `TaskPane.Width` 被静默忽略），文案必须**分开说**："窗口已排列"（由本地桥/原生比较的真实几何或回执核对过）与"AI 比例尚未确认，请拖到整个可用宽度约 30%"（真实回读结论），不得宣称已按 35/35/30 全部自动完成。退出审阅时只恢复**本次流程自动改过**的宽度：`useWpsReviewPaneRatio` 的恢复被 `!manual` 门控，用户中途拖过的宽度不被当成我们的临时布局抹掉。

**两端自动控制的真实范围**：Excel 由 Python 本地桥经 Win32 精确设置两窗矩形（外侧 35/65 可读回核对）；WPS 只能用 `Windows.CompareSideBySideWith` / `ResetPositionsSideBySide` / `SyncScrollingSideBySide` 重排两张表窗口，**不能**设置精确宽度。排列失败只写反馈，不改审阅阶段，也不影响已完成的写回/撤销。

## 2026-10-14：审阅副本的生命周期（身份、关闭、临时文件、业务结果分离）

**用户看到的问题**：Excel 兼容模式生成「工作簿5…」这类审阅副本，确认写回或放弃后副本仍留在屏幕上。排查结论是**能力缺口**不是文案问题：`closeExcelNativeComparison` 在 manual 分支直接 `return false`（注释写着「No API handle」），而前端对 manual 分支干脆不调用关闭。

**三种副本，各自的身份与所有权**（`packages/core/src/reviewCopyLifecycle.ts` 的 `reviewCopyCreationSummary` 是唯一权威说法）：

| 类型 | 实际创建了什么 | 谁关它 / 怎么关 | 有磁盘文件吗 |
| --- | --- | --- | --- |
| `excel-native` | `Excel.createWorkbook(base64)` 新建的**未保存**工作簿窗口（Desktop 1.1） | Office.js `Window.close()`；宿主自行决定是否弹保存提示 | 否 |
| `excel-compat` | `Excel.createWorkbook(artifact.base64)` 新建的**未保存**工作簿（本机生成的 OOXML） | Office.js 在兼容宿主**没有窗口句柄** ⇒ 改用本机桥对该窗口发 `WM_CLOSE` | 否 |
| `wps-native` | `Worksheet.Copy()` 生成的**未保存**单表副本工作簿 | WPS `Workbook.Close(false)` | 否 |

**副本身份判据（不得用文件名/前缀/活动窗口）**：本次随机生成的 `copySheetName = EB_<20hex>` nonce（protocol 侧 `^EB_[0-9a-f]{20}$`）是副本活动表名，再叠宿主侧身份元数据与唯一性。Excel 侧服务端要求「未保存 + 活动表名 == nonce + 唯一 + 同一时刻唯一已保存原表在场」；WPS 侧持有的是活对象引用，另加 `copyIdentity`（Name/FullName/Path 三元组）复核。"工作簿N"、`EB_` 前缀、活动窗口、文件名相似都**不是**判据。

**无提示丢弃副本前必须过三项判断，且前两项推不出第三项**（`describeCopyDiscardSafety` 是第三项的唯一判定入口）：
① **副本身份可信**（nonce + 未保存标记 + 唯一性）；② **没被另存/替换**（否则那是用户自己的文件）；
③ **副本里没有额外的人工修改**。①+② 只说明"这是我们的副本、它没变成用户的磁盘文件"，**不说明**"它还是我们写入后的样子"——用户在并排审阅期间完全可以往副本里敲字。只有适配器**逐格读回并核对**过才允许说 `unchanged`；读不到一律 `unknown`（不许把"没读到"当成"应该没改"）。③ 不通过 ⇒ 返回 `extra_edits`（确认被改过）或 `edits_unverified`（读不到），**不关副本**、把决定权交回用户；这两种结果同时禁止声称磁盘文件已清理。

**结束时的处理规则**（业务结论与副本清理结果**永远分开陈述**）：
1. 写回并**验证**后 ⇒ `closed`：主卡标题即业务结论「已写回原表」；副本关闭状态**另起一处独立**显示（「审阅副本已关闭。」），不带业务前缀。
2. 放弃变更 ⇒ `discarded`：**终态** —— 确认框与三个审阅按钮立刻撤下，主卡只说「已放弃变更」＋「本次计划未写入原表。」。副本仍打开时另给手动处理顺序（**不**无条件让用户选「不保存」：副本里可能有用户自己要留的修改）。
3. 写回失败 / 部分完成 / 结果不确定 ⇒ **默认保留**副本供核对（`not_attempted` + `RETAIN_MESSAGES`），不自动关闭、不删除排查材料；面板只结束会话并说明原因。
4. 用户另存 / 替换过副本（身份变了 / 未保存标记消失）⇒ `user_saved`，**不自动丢弃**，明确交还用户决定。
5. 原表与用户自己保存的文件**永不**被关闭或删除。
6. 关闭失败**绝不**改写业务结论：业务结论（标题＋说明）只由 `outcome` 决定，副本关闭结果只作**次要**的独立状态跟在后面，不冲淡「计划已取消 / 已写回」这一主结论。`write_failed`（守卫在写入前拦下 ⇒ 原表可能一格未改）与 `partial_or_unknown`（执行器跑到一半没确认 ⇒ 结果未知）必须说成**两句不同**的话，绝不互相冒充。

**审阅卡的信息分层（`packages/core/src/reviewCardPresentation.ts`，两端共用）**。这是"哪个阶段说什么、哪条信息放在哪一层"的唯一来源；两端只负责把手上的**事实**喂进来（**不**入此层：执行算法、撤销算法、关闭/清理判据）。

| 阶段（唯一来源 `outcome`） | 首屏（主卡） | 折叠层 |
| --- | --- | --- |
| `reviewing` 审阅中 | 标题 / 说明 / 范围 / 兼容模式必要说明 / 确认框 / 三个按钮（重新排列 · 放弃变更 · 确认写回原表） | 「兼容模式说明」（格式恢复、为何读不回副本）、「诊断详情」（开发环境重启提示、排列与同步滚动指引、随机副本标识、关闭诊断、磁盘清理结果） |
| `discarded` 已放弃 | 标题 ＋ 说明（＋副本仍打开时的手动处理顺序） | —— |
| `verified` 写回成功 | 标题 ＋ 说明 ＋ 实际写回与回读证据；副本关闭状态**独立**显示 | 「兼容模式说明」「诊断详情」 |
| `write_failed` / `partial_or_unknown` 写回失败 | 标题 ＋ 说明（两档**不同**的实话）＋实际原因 | 同上 |

四条纪律：
- `outcome` 是阶段的**唯一**来源，界面想过什么不算数；标题与说明**按 outcome 取值**（不是按 phase），否则 `write_failed` 与 `partial_or_unknown` 会拿到同一个标题。
- 标题 ＋ 说明**就是**业务结论，主卡不再另起一行重复；左右排列说明只出现在审阅中那一句里，不重复念。
- `reviewCardClaimsSuccess()` 是界面套成功模板的**唯一**许可闸：只有 `verified` + `verification === 'passed'` 为真。
- 文案一律**纯文本**（这些串直接进 DOM 文本节点）：`reviewCardMarkdownLeaks()` 是渲染前最后一道闸，`**粗体**` 会原样把星号显示给用户。

**写回入口的唯一闸门（`wpsWriteBackBlockReason`）**。判据只有一份，且**只读实时事实**：`confirmPendingPlan` 是渲染期闭包，旧闭包会一直看到它诞生那一刻的 `pendingPlan` / `stagedReview` / `reviewAcknowledged` —— 放弃之后拿着旧闭包再调一次，闭包里的"计划还在"是**过期事实**。所以闸门读的是每次渲染刷新的实时引用，并在授权检查（异步）返回后**再判一次**；另有一次写回在跑的同步闸（`busy` 是渲染期派生值，同一 tick 的第二次调用还看不到它）。这条路径已有反向变异用例钉住：把闸门改回闭包变量 ⇒ 旧回调会把已放弃的计划写进原表（`expected 1 to be 0`）。
- **异步授权检查回来后的身份复核**（`wpsWriteBackStillCurrent` → `reviewWriteBackTargetChanged`，`packages/core/src/reviewWriteBackGate.ts`）：请求在发起点记下身份三元组——**同一次审阅 / 同一计划版本 / 同一原工作簿**，检查回来后逐项按**对象引用**比对（内容相同但换过对象也算换了），任一侧缺席一律 fail-closed。只判"当前有审阅"会把**新副本**的确认状态借给旧请求：用户核对过的那一份从来没被写、没核对过的那一份反而被写。这条复合路径目前被 `busy`（含 `targetGuard.checking`）挡在界面之外，属纵深防御；判据本身由反向变异钉住（S16/S17）。Excel 端的等价防护是 `save()` 的 `messageId` 归属校验 + 副本 nonce 绑定（兼容模式）+ `assertExcelNativeComparisonUnchanged` 在 `Excel.run` 之后逐字复核计划/授权/工作簿身份/原表快照，并有反向变异 S18 钉住"授权检查在途期间换掉现场 ⇒ 旧请求不写回"。

**关闭副本（当前实现）**：见本文“审阅结束与副本处理（2026-10-15）”。兼容路径已换为同一 Workbook 先归档再关闭，旧 WM_CLOSE 路径不使用；仍需单独实机验收。

**「窗口还在」不等于「宿主在问是否保存」**：关闭请求发出后窗口仍存在，只能说明"关闭尚未确认"。本机桥**没有对话框证据**，关得慢、被别的模态框挡住、宿主在弹保存提示，在这个层面**长得一模一样**。所以一律报 `close_unconfirmed`（「关闭尚未确认，请到副本窗口前确认」）并提示用户查看该窗口，**不推断**原因（关闭回执不包含任何保存对话框证据）。同理由此删掉了两个"看起来有、其实没人会产生"的状态：`prompt_pending`（改名为不推断的 `close_unconfirmed`）与 `user_cancelled`（没有任何适配器能观察到"用户点了取消"；Excel `Window.close()` 返回 `false` 是**歧义**信号，一律按 `close_unconfirmed` 报）。

**磁盘文件**：生成审阅时是未保存工作簿；Excel 兼容路径结束审阅前新增本机恢复归档。恢复文件保留并展示路径，不属于自动删除的临时文件。

**关闭协议**：v4 请求必须携带 closeVersion，回执增加 recoveryCopyPath；原表业务结论仍与关闭结果独立。排列接口保留 v3。

**界面**：`ExcelNativeReviewActions` 与 `WpsAgentPanel` 都渲染 `reviewCopyCleanupState` 产出的「业务结果 → 副本清理结果 → 磁盘文件」三段；可重试的结果（`close_unconfirmed` / `extra_edits` / `edits_unverified` / `failed` / `identity_lost`）**保留审阅卡片**（Excel）或审阅现场（WPS）以便再点一次，已确认收尾的结果才释放会话。`blocked` 阶段按钮改为「保留副本并结束审阅」，`committed` 阶段为「关闭审阅副本」。清理副本**与撤销无关**：不动原表的撤销快照/恢复记录（Excel 端 `excelWriteEndToEnd` / `ExcelFormatReviewFlow`、WPS 端 `wpsNativeComparison` 都有「清理后仍可撤销」的断言）。

**未消费的限额不当成已实现能力**：`config/capabilities.json.reviewCopyLifecycle` 只留服务端真正读的两项 `closeWaitSeconds` / `closePollIntervalMs`（关闭后轮询窗口是否消失）。此前的 `maxCloseAttempts` / `tempCleanupWaitSeconds` / `tempRootLabel` **没有任何消费点**，已删除 —— 本轮**不**引入自动重试循环（重试由用户再点一次驱动），也**不**声称已实现临时文件清理等待。

**验收状态**：替身自动化已证明并在本轮复跑全绿 —— core **455/455**（27 文件，含 `reviewCopyLifecycle.test.ts` 27 例）、Excel 加载项 **1235 通过 / 3 跳过**（105 文件含 2 跳过）、WPS **827 通过 / 3 跳过**（54 文件含 2 跳过）、服务端窗口布局模块 **92 通过**、服务端全量 **1411 通过 / 28 跳过 / 2 失败**（失败两例为 `test_python_event_delivery.py`，本机 Docker 不可用 ⇒ `unavailable`，属环境阻断非代码缺陷）；两端 `npm run build`（含 `tsc --noEmit`）退出码 0；`scripts/mutate_review_copy_lifecycle.mjs` 10 个 SEED 全红且逐字节 sha256 还原，`scripts/check_excel_layout_guards.py` 13 项守卫全被抓，`mutate_excel_format_compat.mjs check` 18 个 SEED 在位。**真实客户端验收未执行**：Excel 2024 与 WPS 里「副本被自动关掉、宿主保存提示出现、用户选不保存后窗口消失」这条链路必须在真机上跑一遍才算闭环；最短步骤见 `docs/DEVELOPMENT.md`「审阅副本生命周期验收（2026-10-14）」。当前只能说「已实现且替身自动化可证」，不能说「实机已验证」。

## 2026-10-05：WPS 审阅窗格比例对齐

`packages/core/src/reviewPaneSizing.ts` 提供双端比例目标、实际宽度确认与手动优先状态机。相关限额从 `excelWindowLayout` 移到 `config/capabilities.json.reviewPaneSizing`，默认 40%、320 CSS px 建议下限、1500ms 检查、200ms 防抖、8px 容差。至少给表格留半窗口是共享产品上限，在 Excel 同时满足 Office 原生上限，不宣称 WPS API 有相同硬性限制。Excel 原有接口与普通偏好恢复不变。

`useWpsReviewPaneRatio` 仅在本次 `stagedReview` 存在时启动。`wpsReviewPaneSizing` 获取已存在的 EB 窗格 ID，不创建窗格；要求右停靠、可见、正宽度以及稳定 Application.Hwnd。Workbooks 集合中原表完整身份和未保存副本身份必须各唯一，当前文档须为其中之一；只读 Name/Path/FullName 等身份，不读 Range/Cells，也不调用 Activate。主窗口 `Application.Width` 是磅，按 96/72 换算逻辑像素，不能用两张表任一子窗口或 iframe/screen 当作主窗口。WPS [Application 文档](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Application/obj) 声明宽度单位；[TaskPane.Width 文档](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/TaskPane/member/Width) 未明确单位，因此用进入审阅时 native Width/真实 WebView viewport 对进行初始校准，不在 resize 尚未落定时反复校准而制造尺寸漂移。实际 viewport 到达目标才清除手动指引，不把 setter 返回当成功。DPR 改变即停止自动路径并提示手动/重新生成审阅；初始非标准 WebView 缩放、跨显示器等必须实机核实，不能凭替身声明通用支持。

普通宽度仅在内存中捕获，退出/放弃/确认后或卸载时尽力恢复；窗格替换、隐藏、宿主切换时不得误调别的窗格。Ribbon 与面板是不同 WebView，使用 `wpsReviewPaneLease` 的短期所有权标记防止 Ribbon 把审阅窄宽重新迁移到默认 1080。PluginStorage 仅保存 paneId/随机 owner/过期时间，6 秒 TTL，可见或隐藏期间续期，手动优先后只续期不调宽；退出按 owner 清除，支持没有 removeItem 的宿主；失效标记不能永久阻止普通初始化。没有路径、计划、矩阵、公式、请求模型/后台/执行账本，也不放宽现有保存与原表核对。

缺少 Width/Hwnd/Workbooks/右停靠现有窗格等证据，原表/副本关闭或身份模糊、API 抛错时停止自动并显示手动路径，绝不改用屏幕大小猜测。界面指引嵌入现有副本审阅状态卡，不另起全局操作卡。WPS 主窗口剩余空间继续由原生比较承载两表，未改变左右关系、同步滚动或审阅确认。openpyxl 文件夹模式无原生任务窗格，不引入桌面窗口操作。

## 2026-10-05：审阅 AI 窗格按右侧窗口 40% 自适应

替代下面 2026-10-04 的固定 360px 行为。外层副本左 40% / 原表与 AI 右 60% 不变；右侧内部 AI 目标是原表 XLMAIN 客户区宽度的 40%，不是屏幕或任务窗格 iframe 的 40%。新增只读 `/api/office/review-layout/measure`，复用 prepare 的严格三字段身份摘要/随机副本 nonce 请求模型与回环、Host、Excel Origin、JSON 守卫。必须唯一匹配已保存原表和本次未保存 nonce 副本，测量前后复核双方身份；仅返回 `clientWidthCssPx`，不返回句柄、路径、计划、值矩阵或标题，不落盘或领取执行权。Win32 GetClientRect 在显式 per-monitor 线程 DPI 上读取，再用 GetDpiForWindow 转成 96-DPI 逻辑像素，finally 恢复线程 DPI；最小化/隐藏/身份变化安全失败。TS/Python 同步增加测量回执，不改既有版本 3 排列与同步回执。

`useReviewPaneRatio` 从活跃副本 session 的原表身份、表名和副本 nonce 取测量范围。TaskPaneApi 1.1 可用才请求宽度；最小建议 320px，极窄时半窗口上限优先（例如 500px 原窗口只能请求 250px）。只读尺寸检查默认 1500ms，resize/focus 防抖 200ms，不重叠请求，页面隐藏暂停原生检查，离开审阅/卸载停止；不是滚动轮询，也不驱动任意窗口。只在真实 host 宽度变化时 setWidth，避免自身 resize 回环；窗格单独改变且不匹配程序请求时，判为用户手动调整，本份审阅停止自动调宽。窗口证据丢失/API 抛错停止自动路径；setWidth 返回不等于成功，实际 iframe 宽度才清除指引。审阅中不覆盖普通宽度偏好，退出后沿用普通恢复逻辑。未获得 session 前不先跳到固定 360px。

WPS 原生比较、openpyxl 文件夹模式、数据授权、执行与保存确认不变；不自动执行或回写任何单元格。真实 Excel 的 DPI/多显示器/WebView 缩放、拖动时序仍待实测；自动化只证明替身与代码行为。

## 2026-10-05：可见边框贴齐与原生同步滚动

外层副本左 40% / 原表与 AI 右 60% 保持。`windowGapDip` 默认改为 0；在物理 DPI 上读取 DWM `DWMWA_EXTENDED_FRAME_BOUNDS`，用它与 GetWindowRect 的差补偿不可见拖拽边框，定位后再次核对可见矩形。不修改系统边框样式，也不以原始 Win32 矩形接近当成视觉间隙已消除；异常或不可获取可见尺寸时回退手动。原有 6 DIP 的说明是此前版本行为。

本地桥接现在不仅摆放窗口：对同 PID、唯一原表与本次新建未保存 nonce 副本，使用原生对象模型 `Windows.CompareSideBySideWith`，显式设置 `SyncScrollingSideBySide=true` 并回读，初始行列视口对齐到原表。两条前端副本路径都消费此桥接。不运行 VBA/宏/脚本，不读取 Range/Cells，不创建 Application、不打开/另存文件，不安装滚动钩子或后台轮询。只允许两份已核实窗口；有第三个应用窗口时不替换潜在无关比较。COM 包装对象 Hwnd、原表路径/表名摘要、副本 nonce/未保存状态、进程与截止时限在变更前核对；宽度移动前和移动后再次核对现场。COM 指针及线程初始化按作用域释放，不把对象带出回调。

TS/Python 回执同步升级为 `layoutVersion:3`，在真实几何回执中新增 `synchronousScrolling: enabled | different_instance | unavailable`。旧版仅排列回执不可冒充联动，前端提示重启后台。仅 Compare 返回真且同步属性回读为真、行列视口相同，才显示已启用；跨实例或证据不足仅显示贴齐排列及手动指引。跨实例自动同步滚动未实现，不是待验项；不为它放宽既有隐私、身份或写入边界。窗口/开关回执不是单元格验证，也不自动勾选审阅。WPS 原生比较和文件夹模式不变。

真实宿主的双向滚动、DWM 时序、多显示器/DPI 与任务窗格留存仍需独立测试工作簿验收；自动化不证明已操作真实 Excel。后台修改需按原启动方式重启、刷新加载项并生成新副本。

## 2026-10-04：审阅时临时收窄 AI 窗格

外层副本左 40% / 原表与 AI 右 60% 不变。`useUIState` 接收原生审阅创建中或存在的状态；进入审阅时，若 `TaskPaneApi 1.1` 可用且当前窗格过宽，仅请求一次配置中的 `excelWindowLayout.reviewTaskPaneWidthCssPx`（默认 360 CSS px）。审阅期间不在焦点/可见性恢复时套用普通宽度，也不将临时 resize 覆盖普通聊天宽度偏好；退出审阅后恢复普通偏好。已经更窄的窗格不强制加宽，用户后续手动拖动不被反复重置。

`setWidth` 可能被宿主静默忽略；只有真实窗格宽度测量能清除“向右拖动 AI 左边界”的提示，不能把请求成功当成尺寸验收。无法使用接口时保留手动路径，不使用任意脚本或扩大窗口桥接权限。此变更仅 Excel 前端，WPS/文件夹模式、窗口几何协议、数据授权与保存确认不变。

## 2026-10-04：窗口排列版本与真实回执

兼容模式的 prepare 回执新增 `layoutVersion: 2`；前端拒绝没有此版本的旧凭据，不向旧服务发起排列。arrange 请求仍只投影一次性 token，回执包含 `status`、`layoutVersion`、`copySide: left`、`originalSide: right`。回执由 Win32 方法在实际矩形、左右关系及双方进程身份核对后产生，HTTP 路由不得只凭调用返回就构造成功。版本/顺序不符不显示「已自动排列」；旧服务有专门重启提示，兼容副本仍可手动审阅，保存确认门槛不变。无新增数据行、单元格、路径或句柄出站/落盘，WPS 与文件夹模式不接此版本协议。

本机观察到运行服务 12:19:33 启动，布局源文件 12:25:46 修改，服务命令不含 reload；运行中的 OpenAPI 仍只定义 `status: arranged`。因此刷新前端不能更新该后台，代码更新后须停止旧进程并按原命令重启、刷新加载项和重新生成副本。版本守卫避免旧服务泛化成功回执被新界面解释为图 2 布局。没有自动重启/终止用户进程，也不补认已有副本。

## 2026-10-04：窄窗格与规划进度交接

任务窗格单列显式使用 `minmax(0, 1fr)`，避免 CSS Grid 的自动最小列宽被内部头部和长文本撑开后遭外层 `overflow: hidden` 裁切。共享头部采用自然 flex 换行，品牌与图标在实际放得下时保持同一行，不再在 ≤700 CSS px 时强制分行；极窄时允许换行，不隐藏品牌或操作。澄清与判断轨迹的网格和长字段也约束在可用宽度内。

Excel `analyzeConfirmedIntent` 的本地步骤只表示「准备规划请求」，范围复核及请求整理完成后明确结束，再接服务端规划事件；它不表示尚在运行的整段「生成预览」。活动结束时，未收到结束确认的其他步骤收敛为警告（失败/取消则沿用对应状态），不保留假运行图标，也不因后续步骤完成而冒充这些步骤成功。活动归属、事件去重、授权与写入协议不变。

隔离 Edge 渲染检查使用真实 Excel App 或 WPS 面板的头部与输入区，以及生产 `TaskProgressTimeline` / `LocalMappingResolution` 组件，覆盖 320/350/400/430/500/620/700/900 CSS px、长模型名/范围/源值与整格/拆分两个表单状态。尺寸证据不等于真实 Excel/WPS WebView 验收，宿主截图仍待重测。

## 2026-10-03：网页前后对比退役

按用户要求移除 Excel 的 `review.html` 对话框与 WPS 的窗格全幅网页双网格，并删除专用组件、传输状态、样式、构建入口及 core 网格投影模块。侧栏保留变更摘要（定位、计数、真实回读状态）及已撤销的历史说明；执行器的前后证据采集、授权核对、验证与恢复能力不删除。

WPS 独立副本原生并排、确认写回与放弃变更保持不变。Excel 原生单表副本审阅已接线（范围和限制见下一节），不等于完成两端全部原生能力对齐。专注对话窗口与只读宿主能力探针不是网页前后对比，继续保留。旧双网格指导文件仅为历史记录，不代表当前产品入口。

## 2026-10-03：Excel 原生单表副本并排审阅

**2026-10-04 兼容路径补充**：`excelNativeReviewMode` 先检测基础 `ExcelApi 1.13` / `createWorkbook`，再区分支持 `ExcelApiDesktop 1.1` 的 Office 窗口模式和不支持它的兼容模式（内部仍名为 manual，如本机 Office 2024 Volume）。兼容模式使用 `manualReviewFile.ts` 将授权表快照与白名单计划直接投影为本地 OOXML：包含真实公式、常量、基础格式及绿色计划变化标记，要求 Excel 打开后重算，不伪造公式缓存结果；列宽按点到字符宽度近似转换。副本数据只通过 `createWorkbook` 交给本机 Excel，不调用不支持的 Office 窗口 API，不向原表写入，不新增单元格数据出站或持久存储。范围、动作、复杂对象与公式依赖限制和 Office 窗口模式相同，未支持的计划不能绕过限制。

兼容模式通过本机桥尝试左右排列，仍需人工核对副本后确认；不能自动回读副本计算值，写回只执行已审阅计划。关闭改为工作簿级归档后关闭，详见本文 2026-10-15 更新。

### 兼容副本的本机窗口排列桥接（2026-10-04）

复用现有本地 FastAPI 进程，新增 `/api/office/review-layout/prepare`、`/arrange`，两端契约同步。只接受回环客户端、回环 Host、Excel 3000 端口的精确 Origin 和 JSON；前端拒绝向远程服务发送布局请求。prepare 载荷严格限定为 Windows 完整路径规范化后的 SHA-256、授权表名 SHA-256、本次随机 `EB_<20hex>` 表名；arrange 只接收一次性凭据。前端显式投影字段，服务端 extra=forbid，不能传计划、值矩阵、完整路径、窗口句柄或命令。原表身份从既有 Office URL 推导，仅支持已保存本机磁盘/UNC 地址，不修改既有授权及持久身份协议。

Windows `XLMAIN`/`EXCEL7` 可访问性接口只读取窗口关联的活动表名和工作簿 Path/FullName 等身份元数据，验证进程为 Excel.exe。不读取 Range/Cells，不创建 COM Application，不运行外部程序/VBA/脚本，不修改工作簿数据；身份扫描不激活窗口，版本 3 比较阶段仅可激活已核实原窗口、调整本次两窗视口及同步开关。prepare 唯一匹配原表并记录所有既有 Excel HWND/PID；createWorkbook 后 arrange 只接受新窗口、未保存副本和本次 nonce 三者同时成立的唯一候选。已有“工作簿1”、仅同名或同路径的多窗口都不放行；移动前再次核对双方身份及 HWND/PID，过期或模糊证据不移动。

Win32 仅在原表所在显示器工作区排列本次两窗：**左侧审阅副本 35%，右侧原表与 AI 窗格 65%**（见 2026-10-14 一节；此前为 40/60，是 AI 窗格被挤出原表的根因），先扣除配置间隔再分配宽度（版本 3 默认可见间隔 0 DIP，并补偿不可见拖拽边框），高度填满工作区。不修改 Excel 自带的系统边框样式或单元格格式；任务窗格的审阅临时宽度由前端单独处理；外侧尺寸由 `reviewPaneSizing.reviewCopyWidthFraction`（服务端取 `1 − 该值`）与 `windowGapDip` 配置，间隔按窗口 DPI 换算，负坐标显示器同样按工作区计算。跨 Excel 进程也可尝试；DPI 处理后核对实际矩形，不一致或失败则尽力恢复原窗口位置，禁止报成功。限额放在 `config/capabilities.json` 的 `excelWindowLayout`：凭据 120 秒、最多 20 项、等待新窗口 12 秒、轮询 250ms、请求超时 20 秒、矩形误差 4px。凭据绑定 Origin、单次使用，只有内存元数据，不落执行账本。COM 元数据调用可能因 Excel 忙而阻塞，前端请求超时并回退，服务端检查等待截止后不得发起迟到的移动；不是所有原生 COM 调用都可强制取消。服务重启会丢弃凭据，必须重新生成副本。窗口布局成功不能自动勾选审阅、放宽写回；版本 3 的同步滚动独立回执只在原生比较和开关回读成功时为 enabled，跨实例自动同步未实现。**关闭**走同一条本机桥的独立协议 `/api/office/review-layout/close`（版本 3）：身份判据与排列相同，但回执只回答「副本关没关」，**不含**任何原表写回结论；兼容副本的**单元格回读**仍未实现（见 2026-10-14 生命周期一节）。

仅 Windows 交互桌面且与 Excel 同用户、可访问的权限级别适用；非 Windows、云文件、不支持的身份、依赖缺失、后台服务会话、权限隔离均回退手动。文件夹 openpyxl 路径不启动窗口布局，WPS 保持既有原生比较通道不变。真实 Office 2024 COM 身份接口、跨实例排列、DPI/多显示器与聊天留存仍待实机验收。

`excelNativeComparison.ts` 用 `Excel.createWorkbook(base64)` 打开本机生成的单表 OOXML 数据种子，不用整本工作簿导出、不用后端生成文件。`Application.windows` 中必须唯一发现本次新增窗口，且其活动表匹配一次性随机名称，才允许通过 `Window.activeWorksheet` 写副本；不能识别时明确失败，不用当前工作簿代替副本。窗口并排使用 `compareCurrentSideBySideWith` / `resetPositionsSideBySide`，聊天仍在原加载项窗格。原生窗口发现、实际左右顺序、任务窗格留存、关闭提示仍需 Excel 真机验收；`createWorkbook` 若打开到另一 Excel 实例，本路径会安全报错。

仅支持一张已授权表上的互不重叠 `writeValues`、矩阵 `writeFormulas`、`clearRange(contents)`。`config/capabilities.json` 的 `excelNativeComparison` 限制已用/合并显示范围 20,000 格、OOXML 8 MiB。保留常用单元格格式、边框、行高列宽以及宿主确认的真实公式；文字常量即使以 `=` 开头也保持种子内的文本类型，不回放整块 formulas 矩阵。复杂对象（表格、合并、图表、形状、透视表、名称、条件格式、数据验证）明确拒绝；跨表、外部、结构化、INDIRECT/OFFSET 引用不能隔离试算，亦拒绝。不是全保真工作簿复制，打印配置、隐藏状态等未对齐。未支持的复杂计划保留既有显式执行流程，不自动从原生审阅失败降级为写原表。

右侧按白名单计划试写、只重算副本工作表、实际回读并标记变更。公式文本回读不等于语义验证，仍由原执行器验收条件核实计算结果；写回后另以 `compareExcelNativeCommittedResult` 比较原表目标的真实公式及计算值与副本，不同或证据不可用时分别收敛为失败/未核实，变更摘要同步降级，不用计划值冒充回读。`useExcelNativeComparison` 仅在内存中按消息 ID 持有一个副本会话；快照、值和 OOXML 不挂到聊天历史，不新增网络/存储通道。**这不是对旧 Office 执行账本数据边界的重构**，确认写回仍复用既有登记/执行/恢复接口。

确认写回先核对工作簿身份、授权、计划、原表全范围数据/公式/基础格式、副本全已用范围数据/公式/数字格式；再在 `executePlan.beforeFirstWrite` 中于持久登记回调之后、首次原表写入之前复核一次。仅执行已审阅的计划，副本手动编辑不导入。失败或部分写入禁止直接重复保存，沿用原执行证据和恢复点处理；关闭失败执行的副本不宣称“原表未写入”，也不撤销既有写入。原生审阅遇到过期来源不提供旧结果直写旁路。放弃只关闭已识别的副本；`Window.close()` 返回 false 时保持会话，不冒充已关闭。写回完成但副本关闭失败，不重复写入。写回不等于保存原文件到磁盘。

## 系统架构

### 整体架构

- 前端：React + TypeScript；Excel 宿主使用 Office.js，WPS 宿主已通过 W0/W1 接入读取、W2 Agent 核心验收，并已实现 W3 受控写入闭环
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
                       │ 可选后端：OpenAI-compatible API / Claude Agent SDK
┌──────────────────────▼────────────────────────────────┐
│                  外部模型服务                          │
└───────────────────────────────────────────────────────┘
```

### 可选 Claude Agent SDK 后端

`/api/models` 有两种显式 SDK 选项：三项环境配置 `CLAUDE_AGENT_SDK_MODEL`、
`ANTHROPIC_API_KEY`、仓库外绝对路径 `CLAUDE_AGENT_SDK_STATE_DIR` 同时存在时列出
`claude-agent-sdk`；已有托管 DeepSeek 连接使用官方 `https://api.deepseek.com`
或 `/v1` 端点、带 Key，且配置了会话目录时，另列出按连接 ID 绑定的
`claude-agent-sdk:<id>`。后者复用该连接的模型 ID 与 Key，只把 SDK 请求明确指向
DeepSeek 官方 `https://api.deepseek.com/anthropic` 兼容端点；不接受任意地址、
HTTP、其它路径或按相似域名推断。连接改址、删掉或清掉 Key 后衍生选项即失效。
原有 OpenAI-compatible 连接及默认选择保持不变；用户在 Excel/WPS 共用模型菜单中显式选 SDK，
`intent.py` / `planner.py` 才分发给 `claude_sdk_agent.py`，不会在请求失败时自动换供应商。
SDK 依赖本地 Python 包与原生 Claude Code 可执行文件，子进程在专用会话目录启动；
Windows SDK 服务由 `scripts/start_claude_sdk_server.ps1` 单进程启动、不启用热重载。
当前 Uvicorn 的 reload/多 worker 模式会安装不支持子进程的 Selector 事件循环；
适配器在 SDK 调用前检测这一组合并明确报不可用。启动脚本同时拒绝已占用的 8765 端口，
与普通 `npm run dev:server` 二选一。CLI 启动/退出异常只返回固定 SDK 提示，
不将原始 stderr、密钥或输入内容透传给前端。
该目录仍用于 CLI 配置；适配器添加 `--no-session-persistence`，本次 SDK 请求不保存会话转录。
旧版可能已有会话资料，不能放进仓库、项目日志或前端存储。
SDK 子进程显式收到选中连接的 base URL 和 Key，避免继承外部环境中的另一服务商地址；
DeepSeek 的 Claude Code 内部模型档位也固定到该连接的模型 ID。SDK 的美元预算参数
不等同于 DeepSeek 的实际账单上限，费用仍以服务商账单为准。

首批 SDK 接入是受限的表结构通道：意图阶段只发送当前需求和结构摘要，
规划阶段只发送工作簿名、已授权表名、范围、尺寸、表头，不发送
`dataRows`、`displayRows`、`dataResults`、`lastResult`、历史对话或 `sourceFile` 路径。
要求数据行的规划仍显式报不可用，不静默转到旧连接。用户主动添加的 PNG/JPEG/WebP
图片是独立输入：意图阶段仅发送数量，规划阶段通过 SDK 异步 user 消息的原生 image/base64
content block 发送；不创建临时图片文件，不开放 Read/Bash。只有明确声明视觉能力的连接才能接收。
SDK 内置文件、命令、网络和代理工具禁用；仅注册 Excel Bro 的终局白名单工具，
SDK 提示显式使用 MCP 全名提交终局结果。若服务商按共用提示返回结构化 JSON，
适配器先解析，再经同一个 `submit_plan`/`submit_answer`/`submit_clarification` handler 复核；
纯文字不能代替计划，JSON 计划也不能绕过宿主能力校验。
`submit_plan` 仍由 `AnalysisPlan`、宿主能力交集和原有预览确认链路校验。
SDK 没有工作簿句柄，也不能直接执行写入。原有文件夹模式与 Office/WPS 都经过
同一分发器，但文件夹请求若需要数据结果亦明确拒绝，不扩大用户选表授权。
预算上限在 `config/capabilities.json` 的 `agent.claudeSdkMaxBudgetUsd`。

#### 两条后端的检查点差异（CAP-02 第 6 条）

两条后端产出**同一联合类型** `AssistantResponse`（plan / answer / clarification），
并共用**同一份工具契约**：`run_sdk_plan` 从 `excel_agent.AGENT_TOOLS_WITH_DATA`
取定义注册进 MCP，终局也走同一批 `_execute_tool` 校验
（`test_sdk_backend_shares_the_blank_policy_contract_not_a_forked_schema` 断言逐字同一份）。

但**检查点能力不对等**，必须在 `planner.py` 显式说明而不是静默降级：

| 维度 | 普通模型后端 | SDK 后端 |
|---|---|---|
| 检查点现场 | `excel_agent` 循环里维持 `loopState` / `messages` / `toolCache` | **当前接入未实现检查点恢复** —— 本次接入把对话状态交给 CLI 子进程托管，服务端拿不到那份现场 |
| 数据行 | 支持 `planningReadPolicy=ifNeeded` | 仅表结构（`none`） |
| 跨后端续跑 | 可被 SDK 请求方引用（会被拒绝） | **显式 `ModelNotAvailableError`** |

> **口径**：表格里写的是"**当前接入**未实现检查点恢复"，说的是**这次接入方式**
> （CLI 子进程托管对话状态、服务端读不到现场）这一事实，**不是**对 Claude Agent SDK
> 本身能力的推断。2026-10-19 已核实当前安装的 `claude-agent-sdk 0.2.159` **确实提供**
> `resume` / `session_id` / `fork_session` / `resume_session_at` / `resume_drops_turn`
> 与 `get_session_messages`，也支持 `session_store` 镜像；当前适配器仍显式使用
> `--no-session-persistence`（`server/app/claude_sdk_agent.py`），因此**产品侧仍未接入**
> 会话恢复。是否启用取决于转录中会落入哪些内容的隐私评审，不能把"SDK 有接口"当成
> "现在可以打开"。将来换接入方式时改这张表，而不是改历史结论。

`planner.py` 在分发到 SDK **之前**检查 `resume_checkpoint`：

```python
if config.adapter == "claude_agent_sdk":
    if resume_checkpoint is not None:
        raise ModelNotAvailableError(
            "Claude Agent SDK 不能恢复另一规划后端的检查点，请重新发起预览。"
        )
```

拒绝的理由不是"做不到"，而是**静默重新规划会制造假成功**：用户点的是"从上次进度继续"，
拿到的却是一次从头开始的新规划 —— 可能重复调用模型与工具，界面还显示"继续成功"。
拒绝发生在启动 CLI 之前，一次模型调用都不发出
（`test_sdk_refuses_another_backends_checkpoint_instead_of_replanning` 断言 `fake.calls == []`）。

#### 规划阶段的结构化澄清

两种模型后端共用 `submit_clarification` 终结工具，返回 `kind: clarification`，
复用 `IntentClarificationResponse`，不把追问混入普通 `answer`，也不解析助手文字猜测问题。
`IntentClarification` 新增可选 `fields`（最多 4 个、唯一 id、label、placeholder）；
至少有一个填写字段或至少两个选项。旧选项卡兼容不变。Excel/WPS 共用无 hooks 的
`ClarificationFields`，各字段独立受控输入，全部非空后才可提交；用户答案按标签组合，
连同原需求经原有授权/表头复核再规划，不直接执行。取消后不生成预览，不读额外数据。
本地未覆盖源值仍走仅面板内存的 `LocalMappingResolution`，不借这条新出口上传源值。

非流式 turn 收到规划澄清时进入 `awaiting_clarification`，相同请求可缓存问题，
用户补充后可以继续。Excel 流式规划调用的请求指纹不可修改，因此规划澄清后的回答
启动新的 turn（保留完整原需求和授权范围），不复用旧持久规划任务。续跑记录新增
`originalPrompt`，旧记录缺少原需求时明确要求重发，不用问题摘要冒充原任务。
自动化验证与两端宿主实机验收分开报告。
Excel 的填写卡保留原规划读取策略，SDK 的首轮与自由文本续答均在本端显式选用
`planningReadPolicy: none` 并剥离数据行；不因“不是预设选项”恢复上传快照。

### WPS 宿主接入状态

apps/wps-addin 使用独立的 WPS Ribbon 与任务窗格入口；功能区采用独立的
“Excel Bro”选项卡，不向 WPS 内置“开始”选项卡插入分组。点击入口打开任务窗格后，
尝试切回内置“开始”选项卡；宿主不支持切换时仍保留已打开的任务窗格。W0 已在 WPS 表格个人版
12.1.0.28488 实测显示功能区、任务窗格并读取当前工作簿、活动表和选区。W1 增加
明确工作表授权、来源绑定、Value2 标准化、限额、分块进度、取消和块边界复核。

WPS 入口不会伪造 Office 全局对象，也不会把打开工作簿透明降级为磁盘文件读取。
W2 已把 WPS 读取器接入现有 query_table 确定性计算与模型连接服务。意图阶段只发送
授权表的字段结构；完整值留在本地计算，回答阶段只发送紧凑结果。WPS 响应出口过去拒绝
所有 `AnalysisPlan`，这项旧限制已在 W3 移除；查询仍走确定性本地计算，写入计划则必须通过
服务端规范化校验和 WPS 本地白名单校验，并在用户确认后执行。

W3 开发验收通道当前开放 `writeValues`、同形 A1 公式矩阵 `writeFormulas` 与 `clearRange(applyTo=contents)`。执行器只允许已授权工作表、
明确 A1 矩形区域、本地已保存且非只读的 `.xlsx`，并拒绝保护表、合并区域、公式注入、
重叠目标和超限写入。执行前完整快照暂存于 WPS `PluginStorage`，写入按行分块，
每块更新恢复点，完成后独立回读；公式动作比对 `Formula` 矩阵而非计算后的 `Value2`，
撤销前必须核对执行后指纹，检测到后续修改时拒绝覆盖。`PluginStorage` 关闭加载项后不持久化，
会话内恢复点不能承诺跨关闭恢复；跨关闭的写入耐久恢复见下「WPS 写入耐久恢复（CAP-03）」。
`#EB_RULES` 与 `#EB_RULES_BACKUP` 是内部系统表：WPS 枚举不会展示，授权入口和执行前复核也会拒绝。
整表删除、行列结构删除、格式、R1C1 铺满公式、拆表等动作仍默认拒绝，未知能力不降级执行。
目标版本的真实写入、异常中断和公式恢复仍须在独立测试工作簿完成验收，详见
[W3-A 验证记录](./WPS_W3_A_VALIDATION_2026-09-23.md)。

**WPS 写入耐久恢复（CAP-03，2026-10-02）**：会话内 `PluginStorage` 恢复点之外，首次原表写入
之前把执行身份、授权与**恢复描述**持久登记到服务端 SQLite 账本（`office_executions` 表，复用
`FolderOperationLedger` 的 revision CAS + ownerToken 轮换 + 动作级检查点），走独立的
`/api/wps/executions/*` 路由 —— **不把 WPS 冒充 Office 接入 `/api/office/executions/*`**。
身份用 `wps-path:` 前缀 + 已保存工作簿的小写磁盘完整路径，与 Office 的 `url:` 前缀是两个
命名空间，账本按完整字符串比较，跨前缀重放必然拒绝。

登记载荷**不是完整计划**：只发 `WpsRecoveryDescriptor`（动作类型 / 目标表与区域 / 目标形状 /
写前指纹 `previewFingerprint` / 写后逻辑指纹 `expectedFingerprint` / `applyTo` / `formulaR1C1`）。此前直接上传 `AnalysisPlan` 会把 `writeValues` 的目标
值矩阵带出站并落进 `plan_json` —— 恢复功能不得为"能精确续跑"放宽这条约束。三重闸：类型层
（`registerWpsExecution` 只收 descriptor）、服务端（`WpsRecovery*Descriptor` 用
`extra="forbid"`，矩阵字段直接 422）、客户端入站（`assertWpsExecutionState` 拒收含矩阵的状态）。

动作级（不是分块）写前提交写前指纹、写后提交写后指纹，完成/失败分别提交
`/complete`/`/suspend`。本地只存发现索引（operationId / 身份 / taskId），服务端也只存恢复元数据，
不存完整现场矩阵。写前证据、写后检查点、核对回执、完成结果同样使用严格字段白名单，
完成结果只收计数；异常请求使用固定码与固定说明，不上传宿主错误中的单元格值。
请求的 runtime 投影也丢弃额外顶层 `plan`，不能只依靠 TypeScript 类型约束。

恢复执行权保护由 `taskRecovery.wpsOwnerLeaseSeconds` 配置（当前 120 秒）。仍在 `running`
且近期有账本活动时，`ownerActive=true`、`canResume=false`；领取与核对在同一 SQLite
事务中复核执行权 + revision，不能抢占仍活跃的执行者。状态 GET 和重复登记均不泄露已有
ownerToken；Office 路由先核对 `url:` 命名空间，不能绕过 WPS 守卫。该时间保护不是续跑
许可证，过期后仍必须核对现场，未知/冲突均拒绝操作。允许用户点击只读「重新核对」，不轮询。

重载后读索引 → 查服务端 → 身份核对 → **五态核对**（`completed` 补记完成 /
`reconcile` 核对在途动作 / `replan` 重新读取并生成预览 / `blocked` 人工核对 /
`insufficient` 挂起）。按钮只由 `wpsReconcileActions` 给出，判定不给就不画；`blocked` 与
`insufficient` 一个写入或补记按钮都没有。恢复卡挂在消息流顶部。判定逐条要正面证明：
此前 `running`/`failed` 且无在途动作就返回"可以继续"（不看 `canResume`、身份、剩余动作、现场）
已被收紧。

现场证据来自 `inspectWpsRecoveryState`：先校验**全部**动作是否在当前已确认授权内，
再仅回读登记目标矩形；已完成动作比对写后指纹，剩余动作比对写前指纹，在途动作单独归类。
不依赖 PluginStorage 的整次执行状态，不补齐未采集的值；缺写后指纹的旧描述仅判证据不足。
点击按钮再次查询 + 回读，补记完成在领取后还要再次回读；全部动作完成只提交回执，宿主写入为零。
已有完成记录只显示历史事实，不宣称当前单元格仍未被修改。重规划前须明确勾选确认当前范围，
旧执行先领取处置权再挂起；规划失败保留发现索引，新预览仍要单独确认写回。

校验指纹不是加密，也不是公式计算正确性的证据。公式跨关闭仅核对逻辑公式文本，
**计算值跨关闭分离核对未实现**。旧版含矩阵的持久记录拒绝解释/下发/接管，错误不回显
敏感内容；本轮不自动删除历史账本，需要另行决定存量迁移/清理。

**精确续跑剩余动作未实现**（不是"待实机验收"）：`WpsPreparedPlan` 是内存对象（确认快照、
值矩阵、来源快照、授权索引、宿主版本），重载后重新 `prepareWpsPlan` 会读到"写后"现场，
无法重建首次执行的确认快照；值矩阵不落盘也不出站。PluginStorage 恢复点里虽有目标矩阵，
但跨加载项关闭即失效，且不含授权索引与来源指纹 —— 拿它续跑等于绕过写前的授权与源区核对。
因此剩余动作非空只能走 `replan`，面板明说"无法精确续跑剩余动作 …… 那是新的一轮"。


**恢复点的入口挂在"产生它的那次执行结果卡"内**，按 `executionId` 与执行绑定 ——
不按消息位置、不按计划标题、也不按"最近一次回复"推断。旧实现把一张撤销卡无条件挂在
消息流末尾，于是**上一轮**写入留下的恢复点看起来像本轮唯一能做的下一步。
「备份与恢复」抽屉只保留一个快捷入口。
每张卡如实区分四种状态：**可撤销 / 撤销冲突 / 已撤销 / 恢复点已不可用**；
只有当前会话里活跃的那条恢复点本人可点 —— 新执行取代旧恢复点、面板重开
（加载项会话已结束）或撤销完成后，旧卡不再保留可点按钮，卡片降级为状态说明。
确有恢复点的结果**不止"回读完全通过"**：写入成功但计算值与副本不符或无法核实、
以及可恢复的部分写入也会画卡（`outcome` 分别记为 `readback_passed` / `unverified` /
`partial`），文案与色调都不冒充"验证通过"；**执行器没给证据才不画卡**。
归属判定与文案在 `apps/wps-addin/src/wpsExecutionRecovery.ts`（纯函数），
渲染与 live 恢复点在 `WpsAgentPanel.tsx`。

**任务中断卡的"原表写入状态"是一条独立的生命周期，必须如实推进**（2026-10-02）。
恢复记录（`@excel-bro/core` 的 `taskResume`）在写路径上按四个时点推进：

| 时点 | `phase` | `writeState` | `resumability` |
| --- | --- | --- | --- |
| 预览已生成、等用户确认 | `preview_ready` | `not_written` | `replan` |
| **第一次宿主写入之前** | `executing` | `unknown` | `reconcile` |
| 写入与回读完成 | `executing` | `confirmed_result` | `unavailable`（终态） |
| 部分写入 / 回读不符 | `executing` | `needs_review` | `reconcile` |
| 用户放弃变更 | `preview_ready` | `not_written` | `unavailable`（终态） |

**顺序是关键**：登记必须早于第一次宿主写入。否则写入途中断线 / 重载时记录里仍是
`not_written`，卡片会对用户说"原表未写入"—— 而原表可能已被改过。宁可先说"未知"。
面板重开时 `unknown` 与 `needs_review` 都收敛为"先核对"，中断原因保留为
"写入结果无法确认"，**不降级成笼统的"连接中断"**；已完成/已放弃的任务是终态，
不再出现"继续"入口。规划身份（`turnId`）随记录一起保存，供后续状态查询定位该任务。

### 规划前能力交集（双宿主对齐 P0）

两端加载项在发规划请求时会自报宿主身份：`hostId`（`excel-office` / `wps-et`）、
`hostVersion` 与 `selfReportedActions`（本端能执行的动作清单）。服务端据此计算可规划动作：

```text
客户端自报 A  →  ∩  ← 服务端按 config/capabilities.json 的可信上界 B
                │
           A ∩ B ⊆ A（客户端只能收窄，不能扩大）
```

- 可信上界来自 `config/capabilities.json` 的 `hostCapabilities`，按
  `hosts.<hostId>.versions.<versionPrefix>.actions.<actionType>` 三层寻址；
  版本前缀按"相等或后接 `.`"匹配（`12.0` 命中 `12.0.1`，不命中 `12.01`），
  重叠前缀取**最长匹配**。
- 动作能力是三态：`supported`（已实机验收，可进计划）、`unverified`（可能支持但未验收，
  **不进计划**）、`unsupported`（宿主 API 不具备，不进计划）。后两者行为相同，区别只在说明文案。
- 请求**未声明** `hostId` 时保持原行为（兼容 `/api/plan` 与文件夹通道）；
  声明了但**无法识别**的宿主按上空集处理并拒绝整份计划 —— 不给"我不认识这个 id"留绕过口子。
- `clearRange` 按 `applyTo` 分键（`clearRange:contents` 等），父动作出现不等于全部变体放开。
- 这是第二道防线。执行预检仍是第三道，不因为有了规划前收窄而删除。

同一份判定在 TS 与 Python 两侧各有一份实现，由跨语言共享夹具
`config/action_capability_cases.json` 对齐；单边改动实现会立刻让另一侧变红。
Office 侧的自报清单 `apps/excel-addin/src/officeActions.ts` 由
`officeActions.test.ts` 对着 `contracts.ts` 的 `ExcelAction` 联合体与 `excel.ts`
的 `executeAction` 执行分支双向锁定，防止漏报（功能静默消失）或虚报（确认后才报错）。
具体阶段与验收矩阵见
[WPS 个人版适配指导](./WPS_PERSONAL_EDITION_ADAPTATION_GUIDE.md)。

WPS 与 Excel 任务窗格共用 `apps/excel-addin/src/styles.css` 作为产品视觉基线，并保持相同的
头部、模型菜单、新对话、工具、备份、格仔、更多菜单、消息流、数据范围入口和输入区结构。

**输入区的「上拉」是共享实现**（`apps/excel-addin/src/hooks/useComposerResize.ts`）：
拖动 `.composer-resize-handle` 调高度、内容变多时自动长高，长到上限之后**允许滚动查看**。
把手放在 `.composer-input-region` 的上内边距里、位于 `.composer-box` **之外** ——
挪进输入框内部会盖住文本区、抢走光标。WPS 侧曾用 CSS 原生的 `resize: vertical` 代替，
在 WPS 的 WebView 里既拖不动，又与继承来的 `overflow-y: hidden` 叠加，把超出高度的几行
直接切掉；那段覆盖已删除，`min-height` 与 hook 的 `minHeight` 保持同一个值。
跨端一致性由 `composerResizeParity.test.ts`（契约）与 `useComposerResize.test.tsx`
（行为）把关。
2026-09-30 用户明确修正 WPS 双视图需求：原表与**独立复制的工作簿**使用 WPS 自带
「并排比较」，对话留在任务窗格。WPS 计划预览后先复制已授权的单张工作表，
只在未保存副本上试写、回读和标色；「确认写回原表」重新核对两边状态后才调用
既有 `executeWpsPlan`，放弃则 `Close(false)`（按钮为「放弃变更」）。
**侧栏只负责状态、摘要、风险与决策**：原生表格负责比较，不重新引入窄 HTML 双网格，
也不默认把逐格长公式铺成主画面。成功结果不再展示旧的双网格。
副本数据只活在当前会话内存；多表计划不通过复制整本工作簿变相扩大授权。
此原生路径**已接线且自动化通过，待 WPS 真机验收左右顺序、并排窗口、底色及任务窗格留存**；
首次真机点击曾因把已保存工作簿的 JSAPI 包装对象用 `===` 判身份而在复制前误拒。
并排审阅现用已保存文件的名称/完整路径/目录、授权表及计划快照复核原表；
新建副本另绑定创建时的名称/路径与写后快照。该修复不放宽未保存工作簿的澄清身份判据。
详见[WPS 原生并排审阅](./WPS_NATIVE_SIDE_BY_SIDE_REVIEW_2026-09-30.md)。

**审阅侧栏的信息层级**（2026-09-30）。状态说明独占 `.wps-native-review-status`，
操作区 `.wps-native-review-actions` **只放按钮**（「重新排列」「放弃变更」「确认写回原表」）。
两者混进同一个单行 flex 容器会把按钮压到逐字折行 —— 那是真机截图里按钮竖排的直接
原因。按钮 ≥44px 高且 `white-space: nowrap`（不靠缩小字号或截断文案塞进一行）；
≥560 CSS px 水平排列，更窄时 `flex-basis: 100%` 整行纵排。完整公式与计划默认收进
`<details>`，用等宽字 + 横向滚动，不再用 `overflow-wrap: anywhere` 在窄列里任意折断。
**默认收起不等于删除信息**：授权范围、覆盖风险、未知类别策略与实际写入类型仍须可查。

**"预计"与"副本回读"必须严格区分**。前者只能来自可信的本地确定性预期，必须标明
*尚未写入*；后者必须来自同一次副本试写的真实回读，不得拿计划值冒充。公式试算与
固定值回退用**不同的真实状态文案**，判据是计划里真实的 `action.type`（不是文案猜测）；
试算或回读失败时不显示"已回读通过"，也不给可写回的主按钮。WPS 没有 Excel 式逐项
`checks`，不得借 UI 伪造验证证据。
Excel 独立 Web 审阅窗口及此前的 WPS Web 宽窗探索见
[宽视图指导](./WIDE_CHANGE_REVIEW_GUIDE_2026-09-29.md)，不应再被当成 WPS 当前目标形态。
宿主差异只保留在读取、写入、恢复和功能可用性适配器中；WPS 尚不支持的工作目录
或结构动作必须明确禁用或拒绝，不能用外观相同的空入口伪装已实现。WPS 的“跟随当前”范围在
每次发送或运行工具时重新解析活动工作表；手动多表范围只会收缩到仍存在的已选表，不会自动
扩大。固化查询直接复用本地确定性执行器，并再次检查保存的数据源是否包含在本轮授权中；
对话历史保存在 WPS 加载项自己的本地存储中。模型连接仍由本机 FastAPI 配置统一提供。
WPS 的本地读取、计算及写入回调会整理为 Excel 共用的 `TaskProgressTimeline` 数据形态，
记录步骤、阶段、动作类型、已授权目标、行数和终态，不把原始单元格值写进过程日志。
过程消息保存在对话历史中；刷新时尚无终态的步骤显示为中断，不冒充继续运行或成功。
消息流仅在用户靠近底部时自动跟随更新；上翻阅读时显示“查看最新进度”，不强制拉回底部。
Excel/WPS 共用的消息自动跟随状态绑定当前对话 ID：新建或切换对话时清除上一段对话的未读进度提示并恢复跟随；内容已位于底部或不足一屏时，不显示该提示。三点提示按钮的样式保持不变。
共享滚动 Hook 的实现放在 `useMessageAutoFollowRuntime.tsx`，由原 `.ts` 入口转发；Vite 为实现生成 Fast Refresh 签名，Hook 顺序变化时不得复用不兼容的组件状态，旧窗格缓存的入口地址仍可加载。
这只对齐了当前 WPS 本地可见过程；服务端规划流的逐事件补发仍须单独接线与验收。

写入能力按共享核心与宿主适配器分层：`apps/excel-addin/src/planCore.ts` 提供两端共用的
A1 矩形边界、二维写入矩阵形状和清空内容值矩阵。`AnalysisPlan` 与服务端
`/api/plans/validate` 仍是同一计划协议；Excel 的 Office.js 执行器与 WPS 的 JSAPI
执行器只负责宿主对象、授权复核、执行前恢复点、实际写入和回读。新增共同写入动作应
先在共享计划语义中实现和测试，再分别接入两个宿主适配器；不能复制一份业务规则。
宿主版本验证与动作可用性独立于共享语义，未经验收的版本不因核心复用而自动开放写入。
当前开发调试加载项对实机报告的 `Application.Version="12.0"` 提供临时验收通道，
让独立测试工作簿能够验证写入、回读和撤销；这不是把 `12.0` 认定为产品版本或
正式已验收版本。先前把安装目录的 `12.1` 误放进正式名单，现已清空
`verifiedVersionPrefixes`；生产构建在真实版本信号和实机验收完成前拒绝 WPS 写入。
预检、计划准备
和确认执行共用同一版本规则，工作簿、授权、动作白名单和恢复点检查不因临时通道放宽。
WPS 写入计划中的空字符串会原样保留；独立回读若得到 `null`，不能再折叠后报告
`verified`。新恢复点使用带 `v2:` 前缀的类型化指纹，区分空字符串与空单元格；
旧 v1 记录仍用旧算法核对，无空值歧义时可以撤销，有歧义时保留恢复资料并拒绝
自动覆盖。预览绑定最终计划、授权表、工作簿身份和目标写前指纹；本地列映射
还绑定源列在计算与预览时的值，执行前复核源和目标，变化时必须重新预览。
WPS 预览显示每个动作前 5 行、前 5 列的待写值；完整计划可按与 Excel 一致的
JSON 方式编辑。应用修改会重新校验计划和实时授权、生成新预览修订，旧确认不能
直接用于修改后的矩阵；本地映射计划仍不把可逆源值送去服务端校验。
本地列映射在同表纯文本且公式可安全限长表示时先生成 `writeFormulas`：共享
确定性映射仍是预期值权威，未知新组合返回 `#N/A`；WPS 只在未保存的右侧副本
试写并核对公式文本及计算值。试算不符会关闭该副本，改用 `writeValues` 重新生成
待审阅副本，不会因公式试错写原表。Excel 当前工作簿与文件夹通道尚无等价的
写入前隔离试算，继续用固定值，不冒充公式能力对齐。
这仍不能代替 WPS 实机写入、保存重读、撤销和异常中断验收。

意图协议增加 `conversation` 结果：问候、能力咨询等非工作簿任务由模型直接回复，
不进入查询工具、规划或澄清卡。只有关键歧义会返回 `clarification`；明确查询继续由
本地 `query_table` 计算，写入仍必须先生成白名单计划并由用户预览确认。WPS 在意图判断
前仅读取已授权表的候选表头和尺寸，完整数据只在模型请求确定性查询后读取；
结构摘要不会包含原始数据行。澄清卡提供推荐选项和自由文本补充。

列映射的匹配顺序是**先整格、后拆分**：来源单元格整体命中映射表就用它；没命中且
规则声明了分隔符（`delimiters`）时，只按已声明的分隔符拆分（多个候选按在格中
**出现位置最靠前**者切分，不按声明序——声明序表达优先级，不是切分方式），拆出的
项逐项精确匹配后按 `valueRank` 取最高级（未声明 rank 的项视为最低）。规则的
`delimiters` 与 `valueRank` 必须**成对**给出：只声明分隔符表达不了"取哪一个"，
服务端会判定规则不完整并返回澄清，而不是按 `mapping` 顺序硬排一个——那个顺序是
模型的声明顺序，不代表用户说的优先级。该校验排在 `planningReadPolicy` 分叉**之前**，
两条通道（服务端算 / 客户端算）用的是同一份 `@excel-bro/core` 实现，规则不全时
两侧都算不出来。

**`rank` 是优先级，不是等级名里的数字**（2026-09-30 真机回归）。core 的约定是
**rank 越大越优先**，而业务等级常见的命名（名字里最小/最"0"的那个最严重）与之**反向**。
模型若照抄字面数字（把最严重的填成 rank 0），**不会报任何错**：服务端 `_rank_entries`
原样保留、core 按 rank 取最大、公式再忠实编码本地预期 —— 整条链路**一致地**给出那个
**较低**的等级，用户只会看到结果莫名变化，排查时又处处"自洽"。所以：
- 极性本身由构造保证，且已被夹具 `column_mapping_cases.json` 的"反过来声明则取另一个"
  用例跨语言钉住（`packages/core` 与 `server/app/column_mapping.py` 各跑一遍）；
- `rank` 与"拆分后取哪一个"的**方向说明**必须留在 `map_column_values` 的 schema 里
  （参数级 `rank` 说明 + 工具级说明 + 系统提示三处），并用**抽象**等级名举例 ——
  生产提示词里不得出现任何具体业务等级字面量，那条负向断言在
  `server/tests/test_map_column_tool.py`。

算不出正确答案时的归因刻意分两类，不得互相冒充：**归因确定**——已按声明的分隔符
拆开、拆出的项里有映射表没覆盖的类别，抛 `UnmappedCategoryError`，补救动作只有
「补充类别映射」；**归因不确定**——整格没命中、又像是用了某种没被声明为可分分隔符
的写法，抛 `NeedsClarificationError`，补救动作有两个可能（补分隔符 / 补映射），
必须让用户选。界面统一用一张澄清卡（复用既有需求澄清卡的状态机），但选项**按归因
收窄**：归因确定时不给「指定其他分隔符」，否则会让用户去改一个本来没问题的设置。
可选项只由 `reason` 推导（`columnMappingClarificationCard`），不直接信服务端下发的
`resolvedBy`。澄清内容**不含来源格原文**，只有归因、缺失类别与"无法处理的格数"——
原文属于本端数据，要展示的阶段界面有自己的源值；错误文案里也不嵌入整格原文。

**两条通道都要在读取源值之前分流澄清**。Office 侧在 `App.tsx` 的 `local_plan_rule`
分支查 `rule.clarification`；WPS 侧在 `runWpsTurn` 的同名分支做同样判断，且两处都必须
排在读取映射源列**之前**——服务端已判定"算不出正确答案"时再读一遍源值，等于白取一次
用户数据只为了说一句"缺映射"。

另有一种**只有本端读到源值后才能发现**的缺口：规则结构完整，但实际来源值未命中映射、
整格含未声明的分隔形式，或**源区域真的存在空白格而空白怎么处理还没定**。Excel 与 WPS
均把 core 的 `UnmappedCategoryError` / `NeedsClarificationError` / `BlankPolicyNeededError`
接到共用的 `LocalMappingResolution`，不再把它们作为终止系统错误。三类缺口按 `reason`
分成三组互斥的选项，**不许互相冒充**：`unmapped` 补类别、`separator_or_mapping` 让用户
在分隔符与映射之间选、`blank` 只给空白的四种处理。空白那一支必须发生在真实读取之后 ——
没读到空白就先问"空白怎么办"是刻意消灭的行为（历史上 `blankLabel` 是必填字段，模型为了
填满它只能在读源值之前提问，用户被迫回答一个可能不存在的问题）。
卡片在当前面板内存展示真实未覆盖值，允许用户选择逐项精确映射，或明确新增分隔符、
补齐拆分项映射、排列**全部目标标签**的优先顺序并勾选确认。候选分隔符只是建议；
不能自动拆分、用 `blankLabel` 冒充未知值，或暗用映射声明顺序决定“最高”。
卡片的持久化消息只含数量、范围和状态；原始来源值与补充规则不放进澄清消息，
目标矩阵不随澄清或规划请求出站。**生成预览后的 Excel `plan` 仍沿用现有计划历史
持久化策略**，不能把“澄清卡不存矩阵”误写成“整个聊天历史不存矩阵”。Excel 重开时
因内存草稿已消失而使待答卡失效。WPS 作答前复核授权身份、
真实表头和宿主对象；Excel 复核工作簿身份及勾选范围，没有可验证身份时不展示
无法安全续答的假卡片，而是要求先保存工作簿并重新发送需求；通过后均**重新读取**授权源列，
在本地生成新的白名单 `AnalysisPlan` 预览，仍须用户确认才写入。重读若又出现缺口，
在原卡继续询问，不沿用旧快照；不调用 `/api/plans/validate`（可逆目标矩阵不能出站）。
服务端预先返回 `rule.clarification` 的语义缺参路径仍使用通用需求澄清卡，
不能与这里“本地读取后发现的来源值缺口”混称同一个闭环。

**空白怎么处理是四条互不相同的路**，不能塞进一个字符串：

| 策略 | 落地方式 | 与相邻策略的差别 |
|---|---|---|
| `keep` | 该位置**从写入范围里排除**（`splitTargetActions` 拆成连续段） | 不是"读回原值再写回"——后者是一次真实写入，会改公式与格式 |
| `clear` | `clearRange` + `applyTo: "contents"` | 得到真空单元格，不是常量空字符串 |
| `value`（含 `""`） | 写一个用户指定的值 | 写空串时 `ISBLANK()` 为假，与 `clear` 不是一回事 |
| `undecided` | 尚未决定 | 只有真的读到空白格才升级成一次澄清；静默折成默认标签是被禁止的 |

协议上用结构化 `BlankPolicy` 表达，`contracts.ts` 与 `models.py` 同步；旧 `blankLabel`
非空串等价于 `{kind:"value"}`，语义无损。跳过位置过多时按
`capabilities.columnMapping.maxSplitActions` 拒绝并说明原因，不降级成写空值。WPS 的公式
通道只在目标是一整块连续写入时才启用——公式会覆盖区域内每一格，表达不了"这一格别动"。

**授权动作与澄清答案是两种状态，不得互相冒充**：点"重新确认数据范围"只证明"你现在授权
这几张表"，不产生答案。实现上把重授权落到 `continued` 终态（无真实答案时原样保留卡片），
说明文案走 `REAUTHORIZATION_NOTICE` 进消息流；只有用户真的填过答案才推进到 `resolved`，
且答案带 `questionId`，改答案**替换**旧确认而不是追加——追加会让模型同时看到两条互相
冲突的确认。

这条判据必须落在**真实取值范围**上，不能靠调用方贴的标签自证。WPS 侧的结构扫描
（`readAuthorizedWpsSheetSchemas`）与数据读取共用 `readChunkRange`，所以该函数
**按目的构造不同的取值面**：结构扫描只取**第 1 行 × 前 `structureColumns` 列**
（表头行），报告的行列数仍是整块已用区域 —— 授权边界与上限判定不受读取范围影响；
数据读取才取完整授权矩形。"澄清档没读源值"因此是**可被証伪的事实**：
`wpsWorkbookReader.test.ts` 按替身宿主真实的 `Value2` 访问坐标记账，断言结构扫描
取过值的格子里**没有一格属于数据行**，而"算得档"必须真的读到源列（对照锚点）。

**宿主调用不得多传自有实参**。`purpose` 只在 `readChunkRange` **内部**决定取值面的
构造与断言，**绝不传给宿主** —— 真宿主对多余实参不宽容，曾因假设"真宿主会忽略
多余实参"而把它塞进 `sheet.Range(first, last, purpose)`，2026-09-30 真机实测
WPS ET 直接抛 `Too many parameters.`，结构扫描第一步就失败。替身是 JS 函数，
多传的实参会被**静默忽略**，所以本地全绿、真机爆炸。防线已内置：替身的
`Range` 按 rest 参数检查实参个数，多传即抛同一错误（反向变异实测 12 条红），
今后任何"往宿主调用里塞自有标签"的改动都会在本地测试当场暴露。

同一个口径也适用于出站面，但两者是**独立的证据面**：读进内存 ≠ 发出去。
本地映射计划在 `fetch` 请求体层面另有一条负向断言——源列值、映射源列地址、
规则本体与算好的目标矩阵都不出现在任何请求体里（`wpsLocalPlanOutbound.test.tsx`）。

**待答澄清绑定授权身份，不绑定索引**。索引只是"第几张表"：重排、插入、删除或
替换之后同一个索引会指向**另一张表**，用户对着旧表回答的澄清会被拿去算新表。
宿主没有跨会话稳定的工作表 ID，因此 WPS 侧保守地把**工作簿名 + 每个已授权表的
索引与名称（含顺序）**算成身份指纹，再加上结构扫描实际读到的**表头指纹**；
任一项变化就作废待答澄清、把原需求退回输入框，要求重新授权、读取与预览
（`wpsClarificationScope.ts` + `WpsAgentPanel` 的 `clarificationScopeStillValid`）。
代价是"表改名"会让澄清卡失效——这是必须付的代价，不得为了"改名不断对话"
牺牲授权边界。指纹只含表头与表名，不含单元格数据值。
这两条是可序列化指纹的判据；**同会话内**还有第三条不可序列化的对象身份判据
（见下），三条并列，任一条判"不同"即失效。

**两条判据必须读不同的数据源，否则后一条只是前一条的副本。** 这是踩过的坑：
"表头指纹"曾一度喂的是 `listWpsWorksheets().sheets.map((s) => [s.index, s.name])`
——**工作表元数据**，与身份指纹同源。后果是只禁用身份判据时，表头判据顶上来把
用例撑绿，造成"身份判据没必要"的假象（实测：仅禁用身份判据，5 条定向用例只红了
1 条）。现在表头指纹只取自 `runWpsTurn` 结构扫描真实读到的第 1 行单元格值
（`WpsTurnResult.headerRows`），复核时经 `readAuthorizedWpsSheetSchemas` 重读。
两侧的独立必要性由 `wpsClarificationIdentityNecessity.test.tsx` 用**单判据变异**
分别钉住：只禁身份 → 4 条红；只禁表头 → 1 条红。不得同时禁用两条来充当证据。

**第三条判据：同会话内存对象身份。** 指纹是**可序列化**的，只负责"跨会话还能证明的
那部分"；它**结构上**就看不见两类碰撞：① 两个**不同目录的同名工作簿**（指纹只带
`workbook.Name`——路径含用户名等信息，不进指纹、**不进 localStorage、不进请求体、
不进日志与错误文案**）；② **同索引、同名、同表头**删除后重建的工作表（复制粘贴一张
表即可，无需任何"看起来危险"的操作）。这两类只能靠**不可序列化的内存对象引用**证明：
它只在当前会话活着，恰好覆盖"指纹看不见、但同会话内还能对得上"的区间。建卡时
`captureWpsObjectIdentity` 记录工作簿对象引用与每张已授权表的对象引用，作答前
`verifyWpsObjectIdentity` 逐项复核（`wpsObjectIdentity.ts`）。**会话内内存 ref，不落盘。**

**放行需要正面证明，不是"没发现不同"。** 三层判据：① 工作簿对象引用；② 路径/完整名
（**只在内存比对**，未保存工作簿两边都为空时**不因此判相同**）；③ 逐张表的对象引用。
**只要引用对不上就判 different，不存在 fall-through 到 same**——名称、表头、尺寸、
甚至稳定 ID 都能在"删除后重建""整表复制"里保持一致，拿它们反向补证等于把
fail-closed 拧回 fail-open（这是实现中实际踩到并纠正过的错：初版判据在引用不等时用
`CodeName` 补证，结果"重建后 CodeName 相同"被判成 same）。稳定 ID 与尺寸只作**加固**：
在引用**已经相等**时再核一遍，防宿主对象池复用；结构上无法证明身份时**一律失效**。

**稳定 ID 只作加固，不作放行依据。** `Worksheet.CodeName` 文档写明"只读、改名不影响"，
但**未说明唯一性**，也**未说明保存重开后是否保留**，且本项目**未在桌面 ET 实机验证**；
`Worksheet.SId` 属 **WebOffice JSAPI**（promise-based），桌面 ET 宿主**未必提供**。
两者都按可选字段探测式读取（`readSheetStableId`，逐候选 `try/catch`），读不到一律回退
`undefined`，**不得凭空造 ID**；捕获身份失败（拿不到工作簿、序号越界、表不存在）时
**抛错**并把 ref 置空，复核随即 fail-closed，不静默返回空身份。

**体验代价（显式接受，不得悄悄优化掉）**：若真宿主每次 `Item()` 返回的是**新包装对象**
而非同一引用，合法续答也会被拒、要求重新授权与重新预览。这是 fail-closed 的必然代价，
属**实机待验项**；不得为了让用例更好看而放宽判据。反向风险——"未保存工作簿的合法续答
被误杀"——已由定向用例保护：同会话内引用可证明时应正常作答。
另：不得为了验身份重读数据行，也不得扩大用户勾选的工作表范围（仍用 `turnSheetIndexes`）。

**作答前只能有三种结果，且不得互相冒充**（2026-09-30 真机回归后修正；
`wpsClarificationVerdict.ts` 是唯一判定入口）：

| 结果 | 条件 | 处理 |
| --- | --- | --- |
| `same` | 对象身份被**正面证明**相同 | 继续回答，复用该轮授权 |
| `changed` | 有**正面变化证据**（路径/表名变了、索引越界、工作簿关掉了、或真的读到了不同的表头） | 旧卡 `invalidated`，提示"数据范围已经变化" |
| `unknown` | **证明不了**（引用换了但无其它证据、读到失败、面板刚重开） | 旧卡 `unconfirmed`：**不谎称数据已变化**，保留原需求与用户刚填的回答，请他在数据范围里重新确认原工作簿与授权工作表；确认后作为**新一轮**重新读取结构、重新规划与预览 |

**"引用不相等"属于 `unknown`，不属于 `changed`。** 宿主可能对**同一张表**返回**新的包装
对象**（用户只改了一格又撤回就会触发）。把它判成"变了"会对着范围根本没变的用户说
"数据范围已经变化" —— 那是假话，且他无从纠正。真正的变化靠**正面证据**判（路径、表名、
索引、表头），不靠"引用对不上"反推。

**读到失败（`null`）不算变化证据。** 指纹比对时任一侧读不到就记"没有可比信息"；
把读失败当成变化是这套判据此前的另一个假话来源。

**面板重开**由独立的读取路径判据保障：`loadWpsChatHistory` 把 `pending` 收敛为
`unconfirmed`（不是 `invalidated`）—— 重开本身没有给出任何变化证据。

三条判据的独立必要性由**单判据反向变异**钉住（`wpsObjectIdentity.test.ts` 19 条 +
`wpsClarificationCollision.test.tsx` 8 条 + `wpsClarificationVerdict.test.ts` 11 条）：
只禁对象身份 → 2 条红，且**正是**"不同目录同名工作簿"与"同索引同名同表头重建"两条；
只禁表头 → 1 条红；只禁重开失效 → 1 条红；把 `unknown` 退回旧的 `changed` → **8 条红**；
复原后全绿。不得同时禁用多条来充当证据。

**`unknown` 不是放行。** 它同样挡住作答，只是不撒谎：用户的原需求与回答被原样放回输入框，
重新确认范围后走**全新**一轮（重新读取、规划、预览），不复用旧卡的授权，
也仍然要经过写入前的预览与确认。

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
├── planner.py              # 确定性地址命令与最终规划
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

2026-10-03 普通规划通道改为按需加载动作 schema：首轮保留列映射、回答、澄清和
`select_plan_actions`（需要数据的兼容入口另保留只读工具），不携带全部动作参数。
模型选择当前宿主允许的动作类型后，下一轮只提供这些类型的完整 `submit_plan` 参数；
完整 AnalysisPlan、宿主能力、授权与安全校验不裁剪。选择结果随规划检查点保存与恢复，
这不是具体行业关键词路由，也不执行任何写入。SDK 当前仍使用原来的终局工具集合。
Excel 当前工作簿的新规划默认 `planningReadPolicy=none`，与 WPS 本地规则通道一致；
无本轮查询结果的结构规划不夹带上轮结果值（客户端与模型载荷分别剥离 lastResult）。
文件夹及显式 ifNeeded 请求仍保留请求快照读取能力，不冒充宿主实时读取。
预算停止返回 `answer.outcome=failed` 与结构化 `stopReason`；两端主路径/Excel续跑消费失败，
服务端热状态与持久规划账本也记 failed，不再因返回 answer 而假报完成。
老回答缺 outcome 仍按正常回答解释。上游 finish_reason=length/content_filter 时拒绝
整份输出（包括看似合法的工具调用）；length 单独停止为 output_length，不缩额重试，
content_filter 沿用有限修复。只记录结束类型及用量计数，不记录原文。
两种后端预算口径仍不同，不把 SDK 成功解释为 token 更少或底层模型更强。

excel_agent.py 的每轮模型调用都由 LoopState 追踪工具路径、缓存命中、连续结构
错误、累计 token、墙上时间和快照截断。循环开始前按固定顺序检查四层预算闸：
最大轮次、连续错误、可选累计 token、墙上时间。累计 token 默认不触发停止，
仅当 agent.maxTotalTokens 配置为正整数时启用；计数与检查点恢复不受影响。
检测层还会识别重复缓存查询以及
2 至 4 步的重复工具路径，并把一次性警告附加到工具结果；接近预算尾部时，
根据发散、收敛失败或循环状态向最后一条用户消息追加收尾提示。
一批只读工具成功返回后，后续轮次只暴露终局工具与按需动作选择工具，要求
Agent 基于已有上下文直接收敛，避免重复读取工作簿耗尽 token 预算。
结构化澄清选项一经用户点选即视为最终语义，不再把确认结果送回意图模型重复
提问。前端通过 `PlanRequest.planningReadPolicy="none"` 标记该路径；规划 Agent
只获得映射规则、澄清、回答与按需计划工具，并依据已授权工作表的准确表头生成
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

达到 token、轮次、连续错误等硬限制后不再抛出裸异常，而是返回 provider 为 system、
outcome 为 failed 的 AnswerResponse，其中包含停止原因、压缩后的工具调用记录和针对截断、格式错误
或循环的建议。流式规划另受 `taskRecovery` 的时间控制，2026-10-10 起**提醒与暂停分离**：
`planningReminderSeconds`（60 秒）只是**非阻塞**的耗时提醒——超过它只发一条限频进度
事件，任务继续、界面不显示"继续"按钮；真正的停止阈值是 `planningHardWallTimeSeconds`
（1800 秒），按**累计活跃执行时间**计算，到达后由循环在启动下一次模型/工具调用之前收口
并如实说明（`pauseType=budget_exhausted`，不承诺"必然可以继续"）。每个可恢复边界仍照常
持久保存检查点。用户通过带 revision 和稳定 requestId 的明确操作接续现场；续跑不提高
token、轮次或错误上限。Claude Agent SDK 后端以同一总时限作单次请求上限，没有 60 秒停顿，
但**没有检查点恢复能力**。
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

### 任务中断后继续（第一阶段）

上面的机制覆盖的是**执行**与**流式规划预算**的中断。第一阶段补的是它们之外、
用户最容易撞上的那一类：**需求已经说清楚、授权也勾好了，任务却因为断网、
服务重启、任务窗格重载而停在半路**，而界面上只剩一句"已失效，请重新发送需求"。
用户既没改需求也没改范围，却被迫从头再来。

阶段、阻塞原因、能否安全继续**是三件不同的事**，分开表达（共享纯语义层
`packages/core/src/taskResume.ts`）：

| 维度 | 类型 | 说明 |
| --- | --- | --- |
| `phase` | `awaiting_clarification` / `answer_submitted` / `reading_source` / `planning` / `preview_ready` / `executing` | 任务走到哪一步 |
| `blocker` | `network` / `user_input` / `authorization` / `source_changed` / `execution_unknown` / `storage` | **为什么停**（与阶段正交） |
| `resumability` | `checkpoint` / `replan` / `reauthorize` / `reconcile` / `unavailable` | 下一步**能安全做什么** |
| `writeState` | `not_written` / `confirmed_result` / `needs_review` / `unknown` | 原表写入状态 |

把"网络中断 / 等待回答 / 需要重新授权 / 执行结果不确定"一律归为"澄清卡已失效"，
会导致两种假话：范围根本没变却被告知"数据已经变化"，或者写入结果未知时被诱导重跑。
`taskResumeDecision()` 是唯一判定入口，**只有 `resumability === "checkpoint"` 才允许说
"从上次保存的进度继续"**；`replan` 一律说明会"重新生成预览"，不冒充精确续跑。

身份分层（复用既有字段，不新造并行账本）：

- `taskId` —— 用户任务的稳定身份，由加载项在一次**新的用户请求**时生成并随恢复记录
  持久化。同一任务内的重新规划、重新授权、重新生成预览都沿用它。
- `turnId` —— 一次规划尝试，仍由服务端 `planning_tasks` 权威持有；重新规划可以换
  `turnId`，但 `taskId` 不变。
- `authorizationRevision` —— 用户**显式确认过**的数据授权版本。它不是许可凭据：
  恢复时必须重新读取真实宿主元数据并重新确认。
- `operationId` / `executionId` —— 真实执行身份，沿用既有文件夹/Office 执行恢复，
  重试、查询、重连**不得**新建。
- `requestId` —— 一次继续动作的幂等身份，沿用既有 `PlanningContinueRequest`。

**持久化走显式白名单。** `sanitizeTaskResumeRecord()` 只重建显式列出的字段，
任何多余键（`plan`、`values`、目标矩阵、`sourceSnapshots`、宿主对象引用、完整路径）
都会被丢弃 —— "原始数据行 / 完整源值 / 目标矩阵不落盘"因此是**结构性**保证，
而不是靠调用方自觉。`originalPrompt` 为空时消毒返回 `null`，调用方据此明确反馈
"无法恢复"，不用问题摘要冒充完整任务。记录里也**不存在**"永久有效授权"这样的字段。

两端接线：

- **WPS**（`WpsAgentPanel.tsx`）：澄清卡与恢复卡是**同一张卡**的不同状态。面板重开时
  待答澄清仍按既有判据收敛为 `unconfirmed`（重开本身没有给出任何变化证据，
  不许说"范围已变"），但卡片不再只剩一句失效说明，而是渲染中断卡：**中断原因 /
  已保留的信息 / 原表写入状态**，加「确认此范围并继续」与「取消任务」。
  「确认此范围并继续」做的是一次**显式重新授权** —— 重新读取当前工作簿、工作表与
  表头，捕获新的对象身份，推进一版 `authorizationRevision`，再用"原需求 + 用户已确认
  的信息"发一轮**全新**请求。旧读取快照、旧计划与旧确认资格就此废弃，因此
  **重新授权之后不会再被同一个旧对象引用检查卡住**。若确认时读到的身份与建卡时
  不同（双方都读到才算证据），先明确告知"任务将应用于新的数据范围"，不悄悄转移授权。
  已提交的回答会写进恢复记录：续答那一轮断线时卡片转为可继续，继续时把原需求与
  确认内容一起重新发出，**不让用户重复回答同一个问题**。「取消任务」是终态
  （`resumability: unavailable`），无论怎样重开都不会再被自动恢复。
- **Excel**：`useScopeSelection` 把**草稿**与**已提交授权**分开，新增
  `committedSheetNames` / `committedScopeMode` / `scopeDraftDirty` /
  `confirmScopeSelection()` / `commitAuthorizedScope()` / `cancelScopeEdit()`。
  勾选、取消勾选、全选/清空只改草稿；「确认范围」才提交，
  且**提交时按当前真实工作簿清单复核**草稿（弹层开了很久时，已经删掉的表不会被
  提交成授权）。「取消范围编辑」把草稿放回已提交授权，一字不改 —— 用户点开只是
  想看一眼别的表时，不该把正在等待回答的任务打掉。
  「范围」包含**口径**（跟随当前 / 固定多表）：`committedScopeMode` 与
  `committedSheetNames` 同批写、同批退，守卫只读已提交口径；口径差异也算草稿差异
  （见文末 2026-10-19（五续））。
- **WPS 范围面板**（`WpsProbePanel.tsx`）同样改成草稿/已提交两态：勾选只改草稿，
  底部动作是「确认范围」与「放弃修改」（面板重开时草稿回到已提交范围）。
  授权版本只在**已提交范围**真正变化时递增。

第一阶段**支持**：等待澄清、已提交答案尚未生成预览、读取或本地计算中断、规划中断、
已生成预览但需要重新读取或重新授权这五类中断在两端保留任务与已确认信息，
并给出绑定 `taskId` 与原消息的继续入口。

第一阶段**不支持**（明确记录，不冒充已实现）：WPS 写入中断后的自动续写；
Excel 侧的图形化中断卡（共享语义层已就绪，Excel 本轮只落地了范围确认闭环与
`planningResume` / `officeResume` / `folderResume` 既有入口）；服务端
`planning_tasks` 尚未冗余保存 `taskId`（`taskId` 当前只在加载项侧作为任务稳定身份，
服务端仍以 `turnId` 为规划权威身份）；跨设备/跨浏览器恢复；豁免预览与写回确认的
"一键重做"。**恢复预览不等于确认写回**：继续之后仍然要经过白名单 `AnalysisPlan`、
预览和用户确认才会写入。

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

### 规划暂停在面板重开后的恢复（CAP-02 第 5 条）

与上面「任务中断后继续」是**两条独立通路**，刻意不合并：

| | 任务中断后继续（第一阶段） | 规划暂停恢复（CAP-02 §5） |
|---|---|---|
| 处理对象 | 需求/授权已就绪但停在半路（断网、重启、重载） | 规划**主动暂停**等用户决定（预算/模型故障/被打断） |
| 权威索引 | 恢复记录（`taskId` + 原需求） | 历史消息里的 `planningPause.turnId` |
| 服务端现场 | `planning_tasks` / `office_executions` | `planning_tasks`（状态查询端点） |
| 是否碰过单元格 | 可能已写入（需核对 `writeState`） | **一个都没碰过** |

合并会让"只是规划暂停"也走写入核对，反过来"写入待核对"被当成可续跑的规划 ——
两类事实混在一张卡上，用户分不清到底写没写。

三层分工：

| 层 | 位置 | 职责 |
|---|---|---|
| 共享层（纯函数） | `packages/core/src/planningStream.ts` → `planningPausedEntriesFromHistory` | 从历史挑暂停条目：只挑 `planningPause`、同 `turnId` 取**最后一条**、载荷取不到不挑、空白 `turnId` 跳过 |
| 宿主适配层（查询） | `apps/wps-addin/src/wpsReadonlyAgent.ts` → `reinspectWpsPlanningPause` | 拿 `turnId` 查真实状态；**只查询不继续**；`payload` 优先用服务端状态，历史快照仅供显示 |
| 界面层 | `apps/wps-addin/src/WpsAgentPanel.tsx` 挂载核对 effect + `planningPauseReload` | 核对中不画操作区；核对通过才重建活跃暂停；读不到进度给保守说明 |

四条纪律（都有反向变异用例）：

1. **重开只核对，不自动续跑。** 自动续跑会在用户完全没点过任何东西的情况下触发
   一次模型调用；有副作用的事必须由用户点按钮发起。
2. **按钮只由服务端结构化状态推导。** 唯一允许产出按钮的地方是
   `wpsPlanningResumeView`（→ 共享层 `planningResumeDecision`），它直接吃
   `getWpsPlanningTaskStatus` 的返回。历史快照能推出按钮，但**不许**用它推。
3. **`revision` / `lastSequence` 一律采信服务端。** 历史快照里的值只用于显示。
4. **查不到时不下结论。** 既不说"能继续"，也不说"任务没了"，只给保守说明。

两个判据**分开写**：`canContinue` 决定"服务端是否放行"，`action` 决定"放行时走哪条
通路"。合并成一个条件会把"放行但不是精确续跑通路"误报成"该任务当前不能精确续跑"，
那是替服务端下一个它没下的结论 —— 面板手里并没有这个证据。

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

### 模型配置与选择语义

模型入口只面对用户自行配置的外部 API 模型，模型列表里不再有虚拟的“基础模式”
占位项。`/api/models` 在没有任何真实连接时返回 `models=[]`、`defaultModelId=null`；
前端把目录状态分为四种并分别表达：加载中、加载失败、加载成功但为空、已配置真实连接。
读取失败不清空已加载的列表，也不据此覆盖本机保存的选择。
WPS 新增模型连接时先打开可填写的表单，再异步读取本机共用的已保存连接；
检查失败与确实为空分开显示，且不会清除输入内容。同一模型 ID（即使服务地址也相同）
可以保存多条连接；目录按保存顺序显示 `(1)`、`(2)` 后缀，后缀不写入模型 ID。
每条连接具有独立的 `connection:<id>` 和 API Key；模型调用只能按此 ID 解析凭据，
不能按模型 ID 合并或选取第一条连接。

选择规则集中在前端 `modelSelection.ts`（Excel 与 WPS 共用），按“本次明确指定 →
本机保存且仍有效 → 服务端声明的默认 → 目录首个可用真实连接 → 无可用”的顺序解析。
界面展示的模型与请求携带的 `modelId` 都取自这一次解析结果，因此不存在“显示首项、
实际请求 `local`”这类不一致。旧的 `local` 记录会在目录成功加载后迁移到有效连接；
删除连接后重新计算选择，删除最后一个连接回到未配置状态。

服务端按入口契约决定是否需要模型，缺模型时抛出稳定错误并由 `/api/turn` 系列入口
映射为 `MODEL_NOT_CONFIGURED`（一条连接都没有，引导「添加模型」）或
`MODEL_NOT_AVAILABLE`（指定的连接已删除、失效或不在允许列表）。普通 Agent 请求
不会静默降级为规则模式，也不会擅自改用其他供应商；模型调用失败保持原任务状态并
返回明确错误。健康检查、模型设置、既有结果查看、确定性工具和恢复资料访问不受
模型校验拦截。

`_local_address_plan()`（清空区域、已知地址写值/写公式）与
`_local_bulk_worksheet_deletion_plan()` 只做字面解析，排在模型解析之前，因此不依赖
模型——它们是保留下来的确定性能力，不是基础模式入口的回归。`provider="local"`
始终表示“结果来自本地确定性工具”，与用户选择了哪个模型无关，历史消息照常展示。

### 通信协议

主要 API：

- `POST /api/turn`：统一轮次入口；规划完成可返回 `plan`、`answer` 或不含单元格值的 `local_plan_rule`。后者由客户端在已授权工作表内本地计算计划（Excel 与 WPS 共用 `packages/core` 的同一套确定性实现，仅宿主读取方式不同），仍须预览确认后写入。服务端只在 `planningReadPolicy: "none"`（客户端声明源值不外发）时下发规则；此模式下客户端送出的规划快照会剥掉数据行，服务端据此判越界、产规则，但拿不到任何源值
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

### 普通模型输出额度与截断（2026-10-03）

- 意图、规划、格式修复分别从 `intentContext.maxOutputTokens` / `agent.maxOutputTokens` /
  `agent.repairMaxOutputTokens` 读取 `null | 正整数`。默认 null：兼容客户端省略
  `max_tokens`，不以极小额度提前切断推理及最终结构化输出，也不硬编码供应商/业务类别。
- 单次额度与循环预算独立：agent.maxTotalTokens 现默认 null，关闭普通通道累计
  token 停止（此前 131072 已取消）。LoopConfig 默认同样为 None，避免备用调用路径
  暗中恢复旧 20000 上限；正整数可以重新开启，非法数值不可默默当作无限。
  开启时在下一模型轮前检查，不是正在生成时的硬费用封顶。用量与检查点计数继续保留，
  超过旧额度的已保存模型响应可以续跑，不通过清空计数假装任务获得预算。
  8 轮、连续结构错误、重复工具警告及授权/预览守卫仍保留，不承诺无限任务运行。
  模型 HTTP 默认超时 180 秒，现有环境覆盖优先；规划整体硬时限仍 300 秒。
- `finish_reason=length` 不等于格式错误：意图不采纳哪怕能解析的残片，走既有安全路由
  兜底；规划（包括恢复检查点中的 `_incomplete=length`）终止为
  `answer.outcome=failed, stopReason=output_length`，不执行工具、不缩小额度重试。
  Python 与两宿主共享 TS 契约同步。真正的格式错误仍执行有限结构修复。
- 默认仍受模型/网关输出与上下文容量限制，不宣称无限输出；SDK 独立预算机制未改。
  诊断仅记 finish_reason、输入/输出/总 token、可用时的 reasoning_tokens 和请求额度模式，
  不输出思考正文、用户文本或矩阵。reasoning_tokens 仅诊断，不再叠加进 total_tokens。
- 自动化证据：真实 HTTP 请求体覆盖意图与规划无 max_tokens、非空 messages/tools，
  本地敏感哨兵存在但不出站；35000 输出 token 的响应后仍能进入下一轮提交；
  截断的合法工具/JSON 不得接受。真实模型成功率与真宿主仍需实测。

- P1 重构总结：`docs/P1_REFACTOR_SUMMARY.md`
- Hooks 使用指南：`docs/HOOKS_GUIDE.md`
- 设计审查报告：`docs/DESIGN_REVIEW_FINAL_SUMMARY.md`

### 澄清续答与重新授权（2026-10-05）

目标确认前移：两端在需求分析入口调用共享 `checkForAnalysis`，已有强身份证据且授权一致时直接进入分析；弱包装对象没有有效显式同意、身份证据缺失或目标变化时，先暂停并显示独立目标选择卡。确认前不读取表头/数据行、不请求意图或规划模型、不创建需求分析活动。确认期间再次采集证据与元数据，变化仍拒绝。业务补全表单只做规则填写和既有授权复核，不授予新目标同意。WPS 已确认范围的显式前置同意只许可重新读取/规划；有匹配的私有目标键、实际表头与授权指纹时，可继续语义澄清及本地计算，但不把原生对象 unknown 改成强身份相同，执行器不消费此许可。2026-10-08 起普通追问不清除弱证据同意；检测到路径/表身份/授权集合冲突、撤销或会话结束仍失效。写入强身份、快照和独立确认不变。文件夹模式、模型协议不变。详见 `docs/AUTHORIZATION_BEFORE_ANALYSIS_2026-10-05.md`；原内联勾选方案仅保留历史记录。

授权目标前置确认：两宿主共享内存基线、空闲元数据监测和操作入口检查。变化先弹目标/工作表选择卡，确认时重新核对；旧预览不能移植到新授权。此层不替代原有强身份及快照守卫；文件夹模式不变。实现范围、未实现项与宿主验收见 `docs/AUTHORIZED_TARGET_PREFLIGHT_2026-10-05.md`。

回答与授权继续分离，但不能先拦截授权、再丢弃答案：WPS 作答先存真实确认，下一轮携带原 taskId、原需求与全部答案。本地规则仅在内存按原消息定位；unknown 时显式重新授权后直接重读计算并生成新预览，不能重新启动原需求而反复问同一道题。Excel 身份不足同样保留本地规则并给出确认入口，读取新结构后才能继续。正面的路径、表身份或授权集合冲突仍拒绝旧规则，预览不能代替确认写入。详见 `docs/CLARIFICATION_LOOP_FIX_2026-10-05.md`。


## 2026-10-11：Python 进度结果卡补全

Claude 初版新增 `pythonResult` 回调，但 HTTP 事件组装、双端共用解析器、普通后端回调及 WPS 消费尚未接通。当前补齐这些边界，服务端字段采用 `LocalResultCard` 严格模型；前端重建元数据白名单，值仍按该轮绑定从回环端点获取。结果展示不改变任务目标，不执行计划。Excel 保存活动日志时保留原有卡片，补新建消息分支并使用该活动任务身份校验绑定；WPS 对结果引用去重并拒收取消/不同任务事件。函数回调测试不称作真实 SSE/客户端端到端；新增 HTTP SSE 与双端面板替身测试另列。

## 审阅结束与副本处理（2026-10-15，本轮范围）

- Excel、WPS 审阅卡移除“重新排列”；生成副本时的首次自动排列保留。
- 放弃计划后立即进入终态，旧写回入口失效；确认写回成功后才收尾副本。写回失败或结果不确定时保留副本。
- Excel 兼容路径启用 v4 本机关闭协议。先核对唯一副本与原表身份，再获取 Workbook 对象，直接检查该对象的未保存状态、唯一工作表 nonce 与 XLSX 格式；不在核对后重新获取关闭对象，不使用 WM_CLOSE、ActiveWorkbook、全局 DisplayAlerts 或关闭整个 Application。
- 同一 Workbook 先 SaveCopyAs 到应用数据目录 `review-copies/<随机目录>/review-copy.xlsx`，检查文件为可读 XLSX 后，再复核对象身份并 Close(False)。归档失败、身份不符时不关；关闭失败仍保留已生成的恢复文件。恢复文件包含副本中的额外人工修改，**不会写回原表**。
- 默认恢复目录为 `%LOCALAPPDATA%\Excel Bro\review-copies`；配置 `EXCEL_BRO_CONFIG_DIR` 时位于该目录下。结果卡显示实际路径；文件仅在本机保留，不上传模型、不自动删除，用户确认不需要后自行删除。原表文件和用户自行另存的副本均不删除、不自动关闭。
- 请求新增必填 `closeVersion: 4`，旧服务在校验时拒绝；返回 `status / layoutVersion:4 / copyClosed / recoveryCopyPath`。排列接口仍是 v3。升级后需重启 Python 服务并刷新加载项。
- Excel 原生与 WPS 保留既有内容核对后关闭逻辑：额外修改或读数未知时保留副本；没有生成磁盘文件时不宣称已删除文件。关闭副本不清除原表撤销记录。
- 实现与自动化证据已补齐；本轮未操作真实 Excel/WPS 窗口，自动化通过不表示实机关闭已验收。

最短实机核对：在专用测试工作簿生成副本，分别点“放弃变更”和确认写回；确认只关闭副本、原表保持打开且无原表保存提示，写回后仍可撤销。兼容副本中另改一格后放弃，打开卡片列出的恢复文件核对该修改仍在。另存副本后再放弃应保留已另存的工作簿。归档目录不可写时应保留副本并报告未关闭。


## 2026-10-18：WPS 排序（sortRange）的语义与守卫

**分层**：顺序语义放共享层 `packages/core/src/tableTransforms.ts`（纯函数，无宿主依赖）；
宿主读写与影响面核对放 WPS 适配层 `apps/wps-addin/src/wpsPlanExecutor.ts`
（`prepareSortRangeAction` / `assertSortableRange`）。两端读同一份 `sortedMatrix`，
避免"同一份计划在 Excel 与 WPS 排出不同顺序"。

**执行模型**：排序**不是**新增一条宿主调用，而是把它折算成"目标矩阵 = 当前值的稳定重排"，
复用既有的值写入、回读、恢复点与撤销链路：

1. `prepare`：读该区域的值/公式 → 跑守卫 → 用 `sortedMatrix` 算出目标矩阵 → 记入
   `WpsPreparedAction.values`，恢复基线仍来自同一次读取的 `beforeValues`/`beforeFingerprint`。
2. `execute`：沿用值动作的写前复核（区域指纹必须仍等于预览时）、分块写入、逐块回读，
   最终整块回读必须与目标矩阵**逐格相等**才算通过。
3. 恢复点：`beforeValues`/`beforeFormulas` 是整块写前现场；撤销按属性/整块写回，
   冲突判定用 `afterFingerprint`。执行后用户又改过区域 ⇒ 撤销直接报冲突，不覆盖。

**为什么必须守卫**（写入前拒绝，不静默丢内容/格式）：

| 场景 | 后果 | 处理 |
|---|---|---|
| 区域含公式 | 重排会改变相对引用的指向（`=B2*2` 搬到第 5 行仍指 B2） | 拒绝，并说明"请先转成值或缩小区域" |
| 数据行内容级格式不一致（字体/填充/数字格式/对齐/换行） | 只重排值会让格式与数据脱钩 | 拒绝；边框与行高属于网格/物理行属性，本就不随行移动，不计入 |
| 与表格对象（ListObject）重叠 | 破坏表头、汇总行与结构化引用 | 拒绝 |
| 宿主没有 `ListObjects` | 无法证明"不在表格里" | **按"无法确认"拒绝**，不当作"没有表格对象" |
| 区域只有表头 | 没有可排序的数据行 | 拒绝 |

**空白语义**：空白（`null` / `undefined` / 空串）在排序中**永远排最后**，与升/降序无关；
这一条在比较器里先于方向翻转处理，否则降序会把空行顶到第一屏。

**审阅与标记**：排序走值动作路径，双栏对照与临时标记读的是真实的写前/写后矩阵，
不伪造逐格差异；范围与变更摘要由既有 `wpsActionChangeReview` 给出。

**去重（removeDuplicates）为什么不在这批**：它是**删除整行并把下方数据上移**的结构变更，
恢复必须先经实机验证 `EntireRow.Insert` 与整块还原。共享层的判据
（`duplicateRowPartition` / `deduplicatedMatrix`）已实现并有测试，但执行入口**不开放**，
`prepareWpsPlan` 明确拒绝并说明原因；不用"压实值后清空尾行"冒充去重。


## 2026-10-19：纯删表计划与完整本地检查点

用户选择的恢复语义为：**打开执行前恢复副本，保留当前原表及后来编辑**。因此这条路径不使用旧的单元格快照重建被删除工作表，也不执行原地整簿覆盖。

- SDK 通过现有 `plan_delete_other_worksheets` 做语义规划；共享 `worksheetDeletion.ts` 只识别已生成的纯 `deleteWorksheet` 白名单计划，不按用户文本另建快捷分支。名称不明确仍走现有澄清。
- Excel `runPlan` 与 WPS `executeStructuralDeletion` 在执行前重新核对身份、勾选范围及异步检查前后的一致性。新生成的纯删表计划按已授权的直接执行流程进入固定本地适配器；混合删表计划不落入旧范围恢复路径。文件夹模式仍用原来的文件夹执行/备份链路。
- `POST /api/worksheet-checkpoints/{execute,confirm-backup,status,open}` 只接受回环调用、明确允许的加载项 Origin、JSON 和严格模型。`execute` 在无需补充备份许可时执行删除；`confirm-backup` 仅续行已绑定计划；`status` 不写宿主；`open` 只打开新的恢复副本。生成计划不等于运行任意代码。
- 固定原生适配器 `worksheet_checkpoint.py` 附着已经打开的工作簿，保留同一 Workbook/Worksheet 对象核对并调用。Excel 用已验证窗口取得对象；WPS 需 `ket.Application` COM 和前端提供的 Hwnd 一致。没有 `Application.Run`、宏、生成脚本或宿主 Python 回退。
- 完整归档会包含所有工作表。若存在未勾选/隐藏表，先返回 `needs_backup_consent` 和确切表名，不备份、不删除；用户可授予仅本地备份许可，不改变分析/删除授权范围。当前仅支持本机已保存、可写、未保护结构的 `.xlsx`，至少保留一个可见工作表。云 URL、未保存簿、XLS/XLSM、图表表及混合修改计划均拒绝。
- `SaveCopyAs` 将含当前未保存编辑的工作簿保存到应用数据目录 `worksheet-checkpoints/<operationId>/before.xlsx`，不保存原文件。校验 ZIP、大小/条目限制、完整表名顺序、SHA256 并持久登记后才能首删。容量限制来自 `worksheetCheckpoint` 配置。
- 删除期间暂时关闭宿主事件及提示，并在 finally 恢复；按保留的 Worksheet 对象删除，逐项核对剩余表名。未完成/回执不明为 `partial_or_unknown`，不自动续删。相同 operationId 只查已有结果；本机线程锁和 OS 文件锁防止多个服务进程并行执行/误认在途状态。
- 恢复先复核原始归档哈希，再复制出新的 `recovered-<随机值>.xlsx` 并打开，关闭链接更新和宏/事件，随后恢复应用设置。不调用原表的 Save、Close、Quit，不覆盖当前原表，不把恢复文件当临时文件自动删除。当前需原表仍打开以定位同一宿主实例。
- 前端刷新只按**本机保存的归属**查询状态（见下文「检查点的归属」）；不会重放删除。成功回执必须含合法结果及可用恢复副本才能显示完成。恢复不是原地撤销；删除可能影响剩余表中的跨表公式，不能声称剩余表计算结果不变。

配置分工：WPS `wpsWrite.allowedActions` 仍是范围执行器白名单；`worksheetCheckpoint.allowedActions` 仅 `deleteWorksheet`，属于固定原生适配器。规划能力取二者并集，不能把删表塞进范围执行器。另补齐已有 `sortRange` 的服务端/TS 恢复登记类型，仍按值指纹校验、不接收格式证据或原始矩阵。

证据层级：新增自动化覆盖真实面板路由、完整归档及部分失败状态模拟；实际 Excel/WPS 的 COM 接口可用性和恢复打开效果尚待实机验收，不能以 `supported` 或构建成功替代实机证据。


### 2026-10-19：隐藏表导致“已全选却无法删表”的修复

实机只读元数据证据：`评分表.xlsx` 有 8 张可见表，以及非常隐藏的 `#EB_RULES` / `#EB_RULES_BACKUP`；范围面板过滤后只列出 8 张。两次失败检查点的 authorizedSheets 均为 Sheet1–Sheet8，deletedSheets 为空、recoveryAvailable=false。服务端却要求全簿 10 张表授权，造成无法从面板满足的条件；没有据此判定 WPS Delete 接口不可用。

- 不按保留名称豁免权限，不把隐藏表自动加入模型/删除范围。服务端只读表名等元数据，说明缺少的备份许可。用户点击“允许仅本地备份并继续”后才 SaveCopyAs；“暂不删除”不续行。
- `confirm-backup` 只收 operationId 与确切 backupSheets，复用持久化的原计划/原授权/原文件；不接受客户端替换动作或授权。额外许可在本地检查点记录中留证，备份内容不上传模型。
- 等待期间整个工作表名顺序必须不变，备份许可必须与当时缺少的名称完全一致；原来选中的删除目标仍只能来自 authorizedSheets。异步许可前两端重新校验当前宿主、授权与会话。查询或重新加载不恢复写权限，也不自动续删。
- 共享卡片将等待许可显示为决策状态；没有恢复文件的旧失败记录折叠展示。清理 UI 不删除恢复记录。新的确认接口沿用回环/Origin/严格 JSON 模型限制与线程、进程双锁。
- 本轮未在业务工作簿执行备份或删除；自动化不能替代真实宿主删除验收。


## 2026-10-19：检查点的归属、消息内结果卡与恢复入口

**问题**：检查点按宿主级 `operationId` 列表加载，没有会话、工作簿与原请求的绑定；两端又在**全部聊天消息之后**统一渲染这份列表。于是删除完成卡、失败记录与暂缓请求长期贴在聊天底部，换会话/换工作簿/刷新都清不掉。

**归属（唯一来源：`packages/core/src/worksheetCheckpoint.ts`）**

- 每条操作在**发起那一刻**绑定 `{ operationId, conversationId, workbookIdentity, requestMessageId }`，先落盘再派发。落盘顺序是"响应丢失也不重复删除"与"刷新后回挂原消息"的共同前提。
- 结果卡按 `requestMessageId + conversationId` 渲染在**对应请求消息**内；不再有会话末尾的统一列表。切换会话不会把回执写到新会话，刷新也不会让旧记录重新冒泡。
- 水合只查询**当前会话 + 当前工作簿**的绑定，且同一 `operationId` 每次挂载最多查一次；`execute` 在同一次挂载内按 `conversationId:planId` 去重。两者合起来保证"刷新只查状态、不重放删除、不重复出卡"。
- 旧版只存 `operationId` 的记录解析为**归属未知**：不归属当前会话、不被当前工作簿认领；恢复能力保留在「备份与恢复」里，按服务端保存的请求元数据（工作簿路径 + host）列出。

**展示**

- 复用 `ExecutionResultCard` 与 `.execution-recovery-footer` 按钮样式：完成只报「已删除 N 张工作表」，删除清单/剩余表/诊断默认折叠（`ExecutionResultDetails`）。
- 恢复入口的按钮文案与普通结果卡逐字一致（「↶ 撤销本次执行」），旁边说明「打开执行前副本，当前原表保留」。打开后只追加一句「已打开恢复副本。」，不写"已撤销/已还原"，不调用单元格撤销算法，不新增执行过程卡。
  > 2026-10-19（三续）起口径变更：这颗按钮改成**在原工作簿原地撤销**，说明句换成「在原工作簿撤回本次删除，保留后续编辑」；"打开执行前副本"只留在「备份与恢复」。见下文「删表撤销改为原地撤销」。
- 备份授权仍在原请求处展示；`confirm-backup` 只接受当时的确切表名，且重授权不改变分析/删除范围，也不恢复已结束的授权现场。
- 「备份与恢复」抽屉新增「工作表删除恢复点」区段（Excel `FolderBackupManager` 注入、WPS 备份抽屉直接渲染），只读列表 + 打开副本。

**协议**

- `WorksheetDeleteCheckpointRequest` 新增可选 `conversationId` / `requestMessageId`（缺失即"归属未知"，旧记录仍可读）。
- 新增 `POST /api/worksheet-checkpoints/list`：`{ workbookPath, host }` → `{ checkpoints: WorksheetCheckpointSummary[] }`。服务端按 `canonical_path` 归一后匹配**同一工作簿**、同一 `host`、同一 `Origin`，记录 `createdAt` 与计划标题；损坏记录跳过而不影响其余记录。列表只读，写入锁被占用时退化为不落盘的读取。


## 2026-10-19（续）：删表撤销入口与"旧计划退役不得吃掉回执"

**现象**：Excel 删除成功后撤销入口消失，聊天里只剩「原目标计划：…（目标已重新授权，需重新生成预览）」；WPS 删除成功显示的是「打开执行前恢复副本」，与普通写入结果卡的「↶ 撤销本次执行」不是同一个名字。

**Excel 的真实根因**（不是样式，也不是归属丢失）：删表让已授权工作表变少 ⇒ 下一轮目标守卫判 `scope_changed` ⇒ 用户按提示在弹层重新确认范围 ⇒ `adopt()` 对**所有**会话消息执行 `retireTargetMessage`（`plan: undefined`）。而结果卡的渲染当时挂在 `message.plan && isWorksheetDeletion(message.plan)` 上，于是"计划预览退役"连带把已完成执行回执藏掉了。检查点绑定、`operationId`、`recoveryAvailable` 与应用层授权过滤都没有参与这个消失过程；`requestMessageId` 也一直有效。

**判据**：`requestMessageId` 是回执的唯一归属。结果卡的渲染**只读** `entryForMessage(message.id, conversationId)`，不再以 `message.plan` 为前提。这样：

- 退役带走的是"还能再执行的预览"，回执与恢复入口留在原请求位置；
- 未执行的删除请求没有回执（没有 entry），退役后就是干干净净的一条历史，不会伪造结果卡；
- 水合（刷新）也走同一条判据，不重复出卡、不重放删除。

**交互统一**：两端结果卡右下角都用 `WORKSHEET_CHECKPOINT_UNDO_LABEL`（「↶ 撤销本次执行」），与普通写入结果卡同名同位置（`.execution-recovery-footer`），旁边固定一句 `WORKSHEET_CHECKPOINT_RECOVERY_HINT`（「打开执行前副本，当前原表保留」）。点击只调检查点 `open`：不调 `execute`、不覆盖原表、不新增执行过程卡；在途禁用、失败保留重试。「备份与恢复」列表仍用 `WORKSHEET_CHECKPOINT_RECOVERY_LABEL`（一行一条记录，"本次"在那里没有唯一所指）。

> 2026-10-19（三续）起：按钮文案不变，动作改为原地撤销（`restore`），说明句改为 `WORKSHEET_CHECKPOINT_UNDO_HINT`。见下文「删表撤销改为在原地恢复」。

**边界不变**：`recoveryAvailable` 仍是显示撤销入口的唯一许可 —— 未执行、失败且没有恢复档案时没有按钮；`partial_or_unknown` 且确有档案时保留入口，同时保留"未全部完成"的原话。
## 2026-10-19（再续）：删表后的卡片顺序、状态收尾与"预期收窄"

**现象三件**：① WPS 先显示删除结果与撤销按钮，后显示执行过程；② Excel 删表完成后同一次计划仍挂着「准备就绪」「验证后保存」，看着像还没执行；③ 两端删表成功后都会弹「工作表范围发生变化 / 尚未写入」，与刚完成的结果直接冲突。

**统一顺序**（两端同构，同一批组件）：`TaskProgressTimeline`（完成后默认折叠）→ `ExecutionResultCard` / `WorksheetCheckpointCard`（"工作表删除已完成 · 已删除 N 张工作表"，明细折叠）→ 卡内 `.execution-recovery-footer`（左侧 `WORKSHEET_CHECKPOINT_RECOVERY_HINT`，右侧 `WORKSHEET_CHECKPOINT_UNDO_LABEL`）。撤销按钮属于结果卡，因此必然在过程之后；卡下面不再重复画一遍过程，聊天底部也不再有常驻检查点列表。WPS 侧把结果卡的渲染位置从 `message.verification` 之后、时间线**之前**，移到时间线**之后** —— 只统一按钮文字而不统一排列，等于没统一。

**旧状态收尾（Excel 特有）**：计划卡的「准备就绪」「验证后保存」原来只判 `message.plan` 存在。删表成功后计划仍在（按要求保留在折叠明细里），于是界面像"还在等执行"。改为按**检查点是否已有回执**收尾：`entryForMessage(...)?.result` 存在即视为本次已结束，不再显示待执行文案与重复执行入口；计划、执行记录、恢复点一个都不删。WPS 侧本来就用 `executedPlanCard` 承担"执行后仍可复核"，无需改。

**范围收缩为什么会被误判**：删表本来就会让表集合变小。若删完什么都不做，下一轮目标守卫读到的表比基线少，判 `scope_changed` → 弹「尚未写入」→ 用户重新确认 → `adopt()` 对全部消息 `retireTargetMessage`，结果卡随之消失（上一轮已修，但提示本身仍是错的）。

**"预期"必须被证明**，判据落在共享层 `packages/core/src/worksheetDeletion.ts` 的 `reconcileWorksheetDeletionScope`。放过的条件是**同时**成立：

1. 前后工作簿身份存在且相同（`identity_missing` / `workbook_changed` 一律拒绝）；
2. 本次执行确有回执（`receipt === null` 拒绝 —— 没有证据就不是预期）；
3. 回执的删除清单与剩余表清单各自唯一（含系统表时两侧同口径过滤）；
4. 回执的剩余表与真实目录**逐名相同**（少一张、多一张都算 `receipt_mismatch`）；
5. 删除清单全部在授权内、且确实已不在目录里；
6. 幸存数 = 授权数 − 删除数，且幸存不为空（`nothing_authorized_left`）。

通过后返回的 `scope` 保持**原授权顺序、只减不增**。端点侧只做两件事：用真实目录再读一次目标，交给 `useAuthorizedTarget.bindVerifiedShrink` 再卡一遍"同簿 + 只减不增 + 不引入未知表"，通过才推进基线；随后把已提交范围对齐到同一份集合（Excel `setSelectedSheetNames`、WPS `onAuthorizeTarget`）。任一步不成立就原样返回，守卫照常按范围变化拦截 —— **没有全局关守卫，也没有无条件清警告**。

**一个真实的坑（同步记账 vs 渲染期状态）**：`confirmBackup` 的顺序是「`patchEntry` 写回执 → 立刻 `await onCompleted()`」。`patchEntry` 里 `receipts` 是 ref，同步落账；`entries` 是 state，要等下一轮渲染。收尾那一步若从 `entriesRef` 取回执，拿到的必然是 `null` —— 于是"有证据的收窄"被误判成"没证据"，下一轮照样弹提示。`receiptForMessage` 因此先查 `bindings` + `receipts`（同步记账），再回落 `entriesRef`。这条行为由反向变异钉住：去掉同步分支，正向用例立刻变红并复现「工作表范围发生变化」。

**测试替身只保留一份目录事实**：替身宿主的 `Worksheet.delete()` 是空实现（真实删除走 COM），所以"删完之后少一张表"必须显式表达。新增 `TestOfficeHost.sheetNames()` / `removeSheet(name)`，删表用例调 `host.removeSheet(TARGET)`；`targetFor` 替身读同一份目录。这样产品 `scan()` 与授权读数永远看到同一份事实，测试里不再另抄一份"少了哪几张表"的声明。

**验收分档**：自动化侧已闭环（core 632 / Excel 1353 / WPS 965 全绿，两端 `tsc --noEmit` 与 build 通过，关键断言均有反向变异证据）。真实 Excel / WPS 宿主的顺序与提示仍需在有专用工作簿的机器上实机确认 —— 自动化证据不等于实机证据。


## 2026-10-19（三续）：删表撤销改为「在原工作簿原地撤销」

**需求变更**：用户改变了此前"撤销时打开恢复副本"的交互选择 —— 点「↶ 撤销本次执行」后要在**当前原工作簿**撤回本次删除，继续留在原文件里工作，不跳到 `recovered-xxx.xlsx`。完整检查点继续保留，「打开执行前副本」降级为「备份与恢复」里的备用入口。**这是恢复能力改造，不是改按钮文案，也不是"打开副本后再切回原窗口"。**

**为什么两端能共用一条实现**：删表本来就走服务端**固定 COM 适配器**（`host="excel"` 定位 `EXCEL7` 文档窗口的保留对象；`host="wps-et"` 取 `ket.Application` 并用 `Hwnd` 校验），两端拿到的是同一个已打开工作簿对象的引用。因此原地撤销既不需要宿主原生 Undo 栈，也**不假定**它能撤销程序删表：只要在同一个对象上调用对象模型。UI 侧只提供 `operationId` 与授权现场，不开放脚本、宏或系统命令。

**恢复算法**（`server/app/worksheet_restore.py`；入参是纯宿主对象，可被替身驱动）

1. `restore_session` 先挂到原簿，关 `EnableEvents` / `DisplayAlerts`、`AutomationSecurity = 3`（禁宏），再用 `Workbooks.Open(archive, 0, True)` **只读**打开归档，`finally` 里 `Close(SaveChanges=False)` 并逐层还原设置。原簿全程不关闭、不保存、不替换 —— 后续编辑因此得以保留。
2. 逐表 `Worksheet.Copy(Before, After)` 整表复制（保留表格对象/图表/批注/合并/条件格式/数据验证/行高列宽），**不是**搬 `Value`。位置由 `_anchor` 决定：贴"归档顺序里前方最近的存活表"之后，没有就贴 `before`，再没有就追加到末尾。
3. **位置与名字是硬判据**：复制完立刻比对 `sheet_names(book)` 与预期列表，不等就抛错 —— 宁可判失败，也不谎报成功。
4. 可见性按归档写回（`Visible` 不会随 `Copy` 过来），不符即抛错。
5. 副本指纹：`UsedRange.Formula` 归一化矩阵 + `sha256` 与归档逐表比对；超 `worksheetCheckpoint.maxRestoreCells` 直接判失败（不静默跳过核对）。
6. `#REF!` 修复：只对"归档公式引用了本次被删表"**且**"现场公式仍含 `#REF!`"的格子写回归档原文；工作簿级 `Names.Item(name).RefersTo` 同样按归档写回。删除后变成 `#REF!` 的公式**不保证**仅靠恢复同名表就自动修好，所以必须有执行前/执行后证据 —— 修复条数写进回执（`repairedFormulas` / `repairedNames`）。
7. 每放回一张就落一次账（`_record_restore_progress`）。部分成功时能如实说出"已恢复哪些 / 未恢复哪些"，重试跳过 `already_restored`，**不重复造表**。同名表已存在 ⇒ `RestoreConflict`，一张都不动。

**冲突与撤销顺序**：`_newer_undeleted` 查同一 `origin` + 同一 `host` + 同一工作簿里**更晚**且仍未撤销的删除记录 ⇒ 拒绝，不拿旧检查点覆盖后续操作。已 `completed` 的记录再次请求直接返回（幂等，不碰宿主）。

**协议**：新增 `POST /api/worksheet-checkpoints/restore` → `WorksheetRestoreResult`（`status` ∈ `completed` / `partial` / `failed` / `conflict`，附 `restoredSheets` / `failedSheets` / `unverifiedSheets` / `repairedFormulas` / `repairedNames` / `limitations`）。`WorksheetDeleteCheckpointResult.restore` 与 `WorksheetCheckpointSummary.restoreStatus` 让结果卡与「备份与恢复」都能区分"已撤销删除"。能力开关 `worksheetCheckpoint.allowedActions` 增加 `restoreWorksheet`，限额新增 `maxRestoreCells`。客户端 `assertWorksheetRestoreResult` 自证覆盖关系：`completed` 必须覆盖全部删除表且无失败/未核对，`partial` 必须同时给出已恢复与失败清单。

**界面**（两端同构，同一批组件）：`.execution-recovery-footer` 按钮仍是 `WORKSHEET_CHECKPOINT_UNDO_LABEL`，说明句换成 `WORKSHEET_CHECKPOINT_UNDO_HINT`（「在原工作簿撤回本次删除，保留后续编辑」）。成功后**只更新同一张卡**：「已撤销本次删除」+「已恢复 N 张工作表」+ 折叠明细里的「已放回：…」，按钮收起（终态），不新增执行过程卡、不新增消息行。冲突/部分/失败：给原文、保留重试入口（`conflict` 不给重试）、**不写"已撤销"**；`recoveryAvailable` 为真时补一句指路「备份与恢复」的副本。删除未完整完成或没有归档时不给入口，只给 `undoReason`。

**授权侧**：共享判据 `packages/core/src/worksheetRestore.ts` 的 `reconcileWorksheetRestoreScope`，与删表时的 `reconcileWorksheetDeletionScope` 对称。撤销把目录变多，不处理就会让下一轮守卫误报"目标已变化"。放过的条件是**同时**成立：前后身份存在且相同；已提交授权里的表都还在目录里；回执说放回的表真在目录里；目录里多出来的**只能是这次放回的**（其他新增 ⇒ `unexpected_sheets`）。通过后 `authorized` **原样返回** —— 撤销不把表加进授权。端点随后用真实目录再核一遍（`useAuthorizedTarget.bindVerifiedRestore`：同 `fileKey`、授权名单前后完全相同、新目录名只能是"原来就有"或"本次 restored"、restored 必须真在目录里），通过才推进基线并对齐已提交范围（Excel `setSelectedSheetNames`、WPS `onAuthorizeTarget`）；任一步不成立就原样返回，守卫照常拦截 —— 不全局关守卫。

**仍然不支持 / 待验**

- 真实 Excel / WPS **没有实机证据**。自动化用替身宿主 + 真实 openpyxl 读写归档，证明的是算法与协议，不是 COM 行为。
- "撤销后保留当前工作表、选区和滚动位置"未做任何保证：实现不主动切换活动表，但也未核对。
- 撤销只修**仍为 `#REF!`** 的格子，不覆盖"删除后用户又改过的关联公式"；这类情况以"修复条数 + 指纹比对"如实反映，不做猜测式覆盖。
- 整簿覆盖、关闭重开、替换磁盘文件三条路径都没有实现，也不允许作为"撤销"出现。


### 2026-10-19：Excel 原生定位与范围重复确认修复

Excel 检查点不再先 `describe()` 再第二次取得工作簿：一次 NativeOM 读取同时取得路径身份与 Workbook 对象，验证唯一窗口、进程身份、路径后保留同一对象，供删除/原地恢复使用。Windows `AccessibleObjectFromWindow` 可暂时返回 E_FAIL；只重试只读发现过程，`worksheetCheckpoint.nativeReadAttempts=3`、`nativeReadRetryDelayMs=150` 控制上限和间隔。持续不可读取、明确身份变化、重复候选分别拒绝；不通过 ActiveWorkbook 猜目标，不重试已经开始的写入。WPS 定位分支未变。

Excel 从文件选择转为当前工作簿模式时，`adoptCurrentWorkbookScope` 在同一次状态更新提交 `committedSheetNames`。此前依赖后续 effect 提交授权，守卫启用后的首次检查会读到旧/空名单，导致已选范围又报 scope_changed。草稿勾选仍不改变授权。新增真实 useScopeSelection + useAuthorizedTarget 组合测试在旧代码上复现该提示，修复后通过。

证据边界：真实 Excel 仅做身份元数据只读检查，旧路径复现 E_FAIL，最终路径连续三次定位成功；未在业务工作簿执行删除、恢复、保存。用户反馈 WPS 已正常，Excel 完整删除→撤销仍待更新服务和加载项后的实机验收。

### 2026-10-19（五续）：数据范围「授权同步」——草稿口径不是授权口径

现象：勾选并「确认」了工作表，之后仍反复弹「工作表范围发生变化」。本轮定位到三条彼此独立的成因，
各自留下判别点（变异 A1–A6；`scripts/mutate_scope_authorization.mjs` +
`scripts/run_scope_authorization_mutations.mjs`，六个 SEED 全部 RED(OK) 且还原后逐字节一致）。

1. **守卫按草稿口径读"当前范围"**。`inspect` 原来读 `workbookScopeMode` —— 那是弹层里的**草稿**
   口径（点一下「跟随当前」就变）。于是"点开看看"之后，后台检查改用活动表去比已提交的多表范围，
   弹「工作表范围发生变化」。现在 `useScopeSelection` 分出 `committedScopeMode`（已提交口径），
   守卫只读它；与 WPS 的 `scopeMode`（已提交）/ `draftMode`（草稿）对齐，两端同源。
2. **「取消」只回退名字、不回退口径**。草稿口径留在原地，触发器与守卫都改按另一套口径说话。
   `cancelScopeEdit` 现在把 `workbookScopeMode` 也退回 `committedScopeMode`。
3. **"已提交"曾被被动 effect 抄写**。旧实现是
   `useEffect(() => { if (selectionConfirmed) setCommittedSheetNames(selectedSheetNames) })`，
   于是**任何**把 `selectionConfirmed` 置真的路径（重读宿主、发送前那一句）都会悄悄改写授权；
   而守卫基线由另一条路推进，中间那一帧就是"新范围 + 旧基线"，也就是假的范围变化。
   现在只有 `commitAuthorizedScope(范围, 口径)` 能写已提交；重读宿主只刷新草稿（首读播种一次）。
   删表收窄（`bindVerifiedShrink`）、原地撤销（`bindVerifiedRestore`）、范围确认三处都在**同一次**
   状态更新里提交，不留中间态。
4. **读数与判断之间的代数隔离**。`useAuthorizedTarget.check()` 在 `await` 之后才读 `bound.current`；
   而删表收尾 / 原地撤销 / 从文件入口采用当前工作簿都**不经过** `read()` 直接 bind，
   于是"拿旧读数比新基线"是常态。现在每次读数记下 `bindGeneration`，回来时代数变了就作废重读
   （最多作废 2 次，之后按"读不出来"处理）—— 既不用旧读数制造假变化，也不放行旧请求。
5. **口径差异算草稿差异**。`scopeDraftDirty` 现在同时比名字与口径（WPS 侧一直是这条判据）；
   发送闸从"仅手动模式"放宽到"草稿 ≠ 已提交"，因为**读取**是按草稿口径走的：
   切到「跟随当前」不确认就发送，会把当时的活动表当授权表读出去。

回归（真实 App / 真实守卫 / 真实范围 Hook，只替身网络与目标读数）：①多表确认后后台检查与下一条
需求都不再要求确认；②从文件入口采用当前工作簿后不重复确认；③固定多表授权后切换活动表不误报；
④草稿改口径后「取消」不改授权；⑤删表收窄与原地撤销后不误报（既有用例）；⑥确认期间在飞的旧读数
不污染新基线；⑦真正换簿 / 外部删表 / 回执与目录不一致仍被拦下；⑧「跟随当前」口径下切表仍要先确认；
⑨只是看了看另一套口径：后台不误报、发送也不静默改读。

证据边界：本轮结论全部是**替身宿主 + 真实 React 链路**的自动化证据。真实 Excel / WPS 下
2s 轮询与真实目标读数延迟的组合、"切表即提示"是否仍显得频繁，**未做实机验收**
（步骤见 `docs/DEVELOPMENT.md`「数据范围授权同步实机验收（未执行）」）。

### 2026-10-19（六续）：「暂时无法核实当前目标」——授权集合的身份不含点选顺序

现象：已确认多表范围，发送任务仍被拦在「暂时无法核实当前目标」（`TargetScopeNotice` 对
`reason === "unavailable"` 的文案），且重新确认后依旧复现。

根因（**一个主题、两处顺序敏感**）：同一组已授权工作表**只因顺序不同**被当成两个目标。

1. 链路上的三个事实本来就不同序：用户点选顺序（Sheet11 → Sheet1 → Sheet10）成为守卫基线
   `bound.selected`；`App.tsx` 的 `adopt` 按**工作表目录**顺序取出名字写进已提交；
   下一次 `inspect` 按那些名字映射回来，于是变成目录顺序（['1','10','11']）。
2. `checkAuthorizedTarget` 比较 `selected` 时**先排序** ⇒ 判 `same`，所以范围提示不弹 ——
   现场看到的不是"范围变化"，而是下面这条。
3. `officeAnalysisTargetEvidence` 的键是 `JSON.stringify([identity, target.selected])`，
   **没有**统一顺序 ⇒ 上面两份范围产出**两个不同的键**。`checkForAnalysis` 对
   `boundEvidence / evidence / repeated` 三个键做冲突判断（`new Set(keys).size > 1` ⇒
   `pause(..., "unavailable")`），于是完全相同的授权被判成"证据冲突"。
   每次重新确认都会重演，所以表现为"反复"。

修复（**判据沉到共享层，两端共用同一份**）：

- `packages/core/src/authorizedTarget.ts` 新增三个共享出口：
  `canonicalTargetSelection()`（授权集合的规范形）、`authorizedTargetEvidenceKey()`（身份 +
  规范形）、`selectedTargetSheetNames()`（已授权名字按 `selected` 顺序）。
  `checkAuthorizedTarget` 也改用它，不再自己抄一份 `[...x].sort()`。
- 规范形**只消除顺序**：返回新数组（绝不原地 `sort` 调用方的数组）、**不去重**、不合并成员。
  重复 ID、缺失表、增删成员照旧是 `scope_changed`。
- Excel `officeAnalysisTargetEvidence` 与 WPS `wpsConfirmationTargetKey` 都改走
  `authorizedTargetEvidenceKey`，两端口径一致；WPS 的工作表身份（索引 / 名称 / 稳定 ID）
  仍逐个参与，一张变了就换键。
- **没有**删除证据冲突判断，**没有**对 `unavailable` 直接放行，也没有改删除/原地恢复算法。

为什么 `App.tsx` 的 `adopt` 没有再顺手改成"保留点选顺序"：`adopt` 写出来的名字顺序在
规范形之后已不参与任何身份判断，再改一处顺序只会留下第二套"看起来一样"的顺序规则。
需要它保持点选顺序的调用点（如执行前复核 `App.tsx` 里对 `authorization` 的逐字比对）
维持现状，本轮未触碰。

回归（真实 App / 真实守卫 / 真实适配器，只替身宿主对象与网络）：
A. 活动表是最后一张，按非目录顺序确认三张表后**首条**需求不再要求确认目标；
B. 相同 ID 集合不同顺序 ⇒ 证据键一致（Excel 与 WPS 各一条）；
C. 增删成员、换文件、替换文档引用仍被拦截（键不同 / 证据为 `null`）；
D. 真实适配器的 `adopt → bind → checkForAnalysis` 完整链（`useAuthorizedTarget.test.tsx`，
   真实 `inspectOfficeAuthorizedTarget` + 真实 `officeAnalysisTargetEvidence`）——
   替身若自己给一个固定证据键，这条判据就会被替身抹掉，所以这里刻意不替换。
反向变异：`node scripts/mutate_scope_identity.mjs I1..I4` + 驱动
`node scripts/run_scope_identity_mutations.mjs`（四判据 + 还原 sha256 一致，
四个 SEED 全部 `RED(OK)`，日志落 `.tmp-mutations/summary-scope-identity.json`）。

证据边界：以上全部是**替身宿主 + 真实 React/守卫/适配器链路**的自动化证据。真实 Excel /
WPS 下端到端是否仍有此提示**未做实机验收**（步骤见 `docs/DEVELOPMENT.md`「已确认多表范围
实机验收（未执行）」）。

### 2026-10-19（七续）：第一批 SDK 业务闭环的实测结果与证据边界

本批不改写上面第六轮的身份判据；它只做两件事：**核对**身份阻断已在生产代码与真实 App
链路中收口，以及把第一批合成案例跑到可复核的程度。

**身份链路的补充证据（不是重复修复）**：`onlineLink.test.tsx` 原先给
`officeAnalysisTargetEvidence` 注入固定键 `online-target-1`。那会让"同一组表两个键"
的缺陷在在线集成里永远抓不到，也违反"替身不许给固定证据键"。现在该方法走真实实现
（`fileKey`/`fileRef` 取宿主现场），范围选择按**非目录顺序**（订单表 → 汇总表 → 会员表）
点击。

**但这条在线链路不是顺序判据的独立证据**：实测应用 I1（证据键退回直接
`JSON.stringify(selected)`）后，在线用例**仍然通过**。原因是 `inspectOfficeAuthorizedTarget`
替身在确认后拿到的名字已经是同一顺序，链路里不会出现"基线按点选顺序、已提交按目录顺序"
的两份数据。顺序判据仍由真实 App 用例 A（真实 `adopt → bind → checkForAnalysis`）与
`I1–I4` 变异证明；在线改造的价值是去掉固定键、强制真实 `fileKey`/`fileRef` 核对，并把
"确认后开放保存工具"纳入断言。不要用这条在线用例冒充顺序缺陷的回归证据。

**真实模型 / SDK / Docker 合成案例**（`scripts/diagnose_sdk_http.py --python`）：

| 运行 | 结果 | 证据 |
| --- | --- | --- |
| 基线（首次，任务描述对"空编号"范围有歧义） | 汇总与重复/未匹配正确；**会员表空编号明细缺失**（写成"共 0 条"） | `outputs/sdk-batch1-real-baseline/` |
| 基线（收窄歧义后重跑） | 金卡 400 / 银卡 200 / 未匹配 400 / 总计 1000；会员表第 5 行空编号、第 3 行重复键 7、订单表第 5 行 404 均可定位 | `outputs/sdk-batch1-real-baseline3/` |
| 因果变化（仅改会员表 C2 金卡→银卡） | 金卡 300 / 银卡 300 / 未匹配 400 / 总计 1000；三类异常定位不变 | `outputs/sdk-batch1-real-causal/` |

事后业务核对是独立脚本 `scripts/check_sdk_batch1_result.py`，只读工件、不参与产品运行；
它必须能在首次失败工件上报错（只报"缺少会员表空编号明细"），在修复后的工件上通过。
运行环境噪声：SDK 子进程退出时被本机安全策略拦下 `reg.exe`，不影响已落盘的
`report.json` / `synthetic-result.json`；另有一次 `MODEL_CONNECTION_ERROR` 是传输波动，
失败工件保留在 `outputs/sdk-batch1-real-baseline2/`。

**替身宿主集成**：`npm run test:online` 两端各 1 passed（真实服务端、真实 Docker、真实
面板与桥；固定模型工具序列 + 有状态宿主替身），并新增"确认后保存工具入口存在"的断言。
`npm run test:cross-process` 生产侧 4 passed、消费侧 2 passed，消费用例改为当前
"计划到达即执行一次 → 确认只清标记 → 撤销"契约，不再等待已取消的「执行到 Excel」入口。

**未完成**：真实 Excel / WPS 实机仍未执行；因此"真实模型自主规划 + 真实宿主写回/撤销"
这条完整证据链仍缺最后一环，不能用替身证据代替。

### 2026-10-19（八续）：已安装 SDK 的会话接口核实（仅设计，不启用）

`claude-agent-sdk 0.2.159` 的 `ClaudeAgentOptions` 提供 `resume`、`session_id`、
`fork_session`、`resume_session_at`、`resume_drops_turn`，包内另有
`get_session_messages` 与 `session_store`。这与"当前接入未实现恢复"并不矛盾：当前
适配器显式添加 `--no-session-persistence`，服务端也拿不到 CLI 的对话现场。

恢复能力的四项语义必须分开设计，不能互相冒充：

1. **查询已有任务状态**：只读，回答"做到哪一步"，不推进任何写入；
2. **恢复 SDK 规划会话**：重新建立上下文，但必须先复核任务、工作簿、授权、计划版本与来源；
3. **继续未完成执行**：只对"写前登记过、现场可核对"的剩余动作生效，已完成步骤防重放；
4. **撤销已完成执行**：走既有原地恢复/检查点通道，不是"再跑一次"。

隐私前置条件未满足前不打开持久化：转录里不得出现原始数据行、携带数据的代码、密钥或
工具原始输出。若安全恢复必须依赖尚未取得的持久化授权，只提交可审阅设计，由用户决定。

**补充事实（2026-10-19 晚，只读 `--help` 实测）**：捆绑 CLI `2.1.281` 对
`--no-session-persistence` 的说明原文是 `sessions will not be saved to disk and cannot be
resumed (only works with --print)` —— 该开关与 `resume` 互斥。因此"打开会话恢复"不是
删掉一个 extra_arg，而是先定转录落盘范围、清理与保留期。

同时核实到一件更紧的事：**SDK 规划任务今天根本没有"继续"能力**。
`server/app/planner.py:41` 收了 `on_checkpoint` 却不传给 `run_sdk_plan`，
账本 `checkpoint_json` 恒空 ⇒ `_planning_pause_payload` 的 `canContinue` 恒假
（`server/app/main.py:1787`）⇒ 出错即 `finish_planning_task(status="failed")`
（`:2005`）。客户端 `App.tsx:2810-3075` 那套续跑 UI 对 SDK 任务走不到。

完整设计（状态转换表、九种故障场景、最小持久记录与隐私边界、改动清单、验收矩阵）见
`docs/SDK_RECOVERY_DESIGN_2026-10-19.md`；本文件不再重复表述，避免两份描述漂移。

## 2026-10-20：规划前桥准备与检查点持久化失败

WPS 首次规划在意图返回 taskId 后调用并等待 `preparePlanningBridge`，成功后才发送规划请求。
准备阶段仅读取身份与授权元数据。首次 `bridgeTool` 到达时同步固定来源指纹，
校验本轮身份/授权后才允许轮询，不重建桥或清空 delivered；失败则停止并提示。
进度回调只服务已经准备好的本轮桥，不再在事件里异步初始化。
两端共享的 `useSdkToolBridge.begin` 在读取授权失败、空授权或初始化失效时返回 false；
新初始化期间不把旧绑定视为活桥。已完成结果的绑定仍保留供按引用读取。

检查点状态文件的原子替换遇到 `PermissionError` 时，只重试替换本身，不重放删除或 Copy。
限额位于 `worksheetCheckpoint.stateWriteAttempts`（默认 4）及 `stateWriteRetryDelayMs`
（默认 75 ms）。耗尽后停止并返回固定提示，保留最后成功提交的状态和已有归档。
每张恢复表在 Copy 前提交内部字段 `restorePendingSheet`，核对后提交进度并清除此字段。
若复制后回执未能持久化，状态查询/列表将其投影为 `restore.unverifiedSheets`。
重试先核对原工作簿身份与目录：待核对表已存在则停止，不按同名猜测归属；确认不存在后，
先提交该观察再允许补恢复。不会自动覆盖用户后来创建的同名表。

以上为实现及故障注入边界，未替代真实 SDK、Excel、WPS 的 A1 验收。

## 2026-10-20：暂停任务的自然语言核对与原任务归属

消息理解协议增加可选 `recoveryTasks`（当前会话至多六个未结束规划任务 ID）及
`task_recovery { turnId, provider }` 响应。模型结合原有对话理解恢复意图；本地不按
“继续”等关键词重建需求。服务端解析、客户端 API 与面板均拒绝未提供过的任务 ID；
带新图片的消息不能借恢复分支忽略图片。无候选时不新增请求字段，不增加原始行数据传输。

Excel 将恢复意图接回原消息的状态核对入口。WPS 在开始新需求前保留原任务的内存归属
（会话、原提示、授权索引、目标对象、任务 ID、游标与续跑请求身份），普通聊天不清除
暂停上下文，新任务不移除旧卡的核对入口。该上下文不新增持久化；重开窗格只能先查询
服务端，不能用当前工作簿冒充旧任务目标。服务端确认有可续检查点后才显示继续按钮，
点击时仍走现场身份、授权、来源证据及执行账本校验。

自然语言恢复先查询，不自动 POST 继续请求，也不新建规划。WPS 查询得到已完成的纯删表
计划时，不进入会立即写入的结构删除通道；显示已生成方案、未执行，后续删除仍需重新
进入授权与执行检查。其他结果沿用既有结果接回流程。运行中状态更新原卡，不追加互相
矛盾的“请重新发送 / 不要重复发送”提示。

自动化覆盖模型协议、真实 HTTP 请求投影、双端面板恢复路由、旧卡保留、非法任务 ID、
重开 WPS 窗格与只查不删边界；真实 SDK + Excel/WPS 在途中断续跑仍待实机验收。

### 2026-10-20 补充：规划任务建立之前的暂停

上节只覆盖已经拿到规划身份的暂停，没有覆盖消息理解、字段读取、字段意图判断阶段。
这些阶段没有检查点，旧代码只追加系统提示，因此模型拿不到恢复候选，“继续”仍会
成为新的空需求。不能用“没有检查点”代替对原需求的保留。

双端共用 `usePrePlanningPause` 与 `PrePlanningPauseCard`：只在本窗格内存保留原提示、
附件、会话、原授权快照和口径，并关联原暂停消息。模型只收到 `preplanning-*` 引用；
引用不传给服务端状态/继续接口。选择该引用或点击“继续分析”时，先核对原工作簿、
已授权集合及口径，再从原需求重新分析，明确不声称从检查点续跑。现场变化、不可读、
核对期间切换会话或新请求抢占都不放行。引用在发起前同步消费，避免双击重复发起。
候选按消息顺序排列；新请求不追加重复的原需求气泡，原暂停行更新为已接回。

消息/字段模型返回后再次检查取消状态；Excel 字段读取后的取消也保留原需求。
已经提交规划的任务仍走原任务核对链，不能回落到本机重分析。共享内存不新增磁盘转录，
重开窗格不承诺恢复这类尚无服务端身份的早期请求。当前早期恢复实现限于两端当前
工作簿模式；文件夹模式没有同等可验证宿主快照，不自动重放，原需求仍留在对话。

两端均已实现并有面板用例；真实宿主早期暂停仍待复测。独立真实 SDK 合成文本验证：
前文含“没有检查点”，输入“继续”后返回对应本机引用，未读取工作簿。该证据只证明
语义路由，不代替真实宿主全链路。

### 2026-10-20 补充：早期恢复请求的同步交接

用户实机显示“已接回原需求”后只结束两次消息理解。原因是 `restart` 同步消费内存
引用后立即调用旧渲染闭包，双端候选列表仍从渲染期 `entries` 读取，再次把已消费的
引用提供给模型。模型再次选择该引用时，恢复方法找不到原条目，直接返回。
原测试根据提示词是否为“继续”固定第二次模型响应，未覆盖这个交错。

共享 Hook 新增 `pendingEntries()`，调度读取实时、未被领取的候选；渲染仍使用
`entries`。明确选定原需求后的 `restartingAnalysis` 不提供任何恢复候选，避免重分析
再次跳转到已消费任务或其他暂停任务。双端对早期恢复同样校验 ID 确实已提供。
工作簿与授权复核、字段重读及后续模型判断保持原流程，不新增语义关键词判断。

双端面板用例覆盖输入“继续”与按钮入口；模型替身只要收到恢复候选就再次选择，
要求最终确实进入原需求分析。真实 SDK 两次合成请求分别返回 `task_recovery` 和
`needs_workbook_context`；此项没有读写宿主，不能作为 Excel/WPS 实机完成证据。

### 2026-10-21：暂停交互收敛到输入框

双端普通暂停默认只显示一行状态，原核对、模型切换和诊断卡放入 `PauseDetails`
展开项。已确定不能继续的状态展开说明，授权变化仍由范围守卫提示。
输入为空且没有附件时，右下角按钮可继续当前会话最近一条未结束的暂停任务；
输入了内容仍走模型语义判断，新消息不会被本地关键词规则改成续跑。没有暂停任务
时空发送仍禁用，写入/回读阶段的中断限制保持不变。

消息理解协议 `task_recovery` 新增 `action: inspect | continue`，旧响应省略时只查询。
明确继续会先查状态，确认可续后在同一次用户操作内完成原有证据核验和续跑；后台
自动检查与仅询问进度仍只读，不因检查点可用而自动继续。运行中的原任务不新建、
不重放；完成、失败、未知状态沿原分支处理。早期本机暂停没有规划身份，明确继续
仍复核原授权后重新分析原需求；仅查看状态不启动重分析。

WPS 查询路径返回本次核实的继续上下文供显式操作使用，不从尚未重绘的 state
读取；随后沿既有继续入口再次读取服务端状态、授权与来源证据。空发送按钮只表达
继续意愿，不扩大工作簿授权、不绕过执行账本。协议更新需要后端和双端同时加载。

### 2026-10-21：续跑接回原步骤时间线

WPS 首次进度将 task/turn ID 与 sequence 绑定到原 activityLog；检查点继续在请求前
重新激活原卡，流事件按 stepId 更新原行，新 stepId 才追加。重复序号及异任务事件
不合并。结果取回也复用该任务的过程卡；再次暂停保持相同关联。旧历史缺少绑定时，
仅在该暂停消息之前、同一用户轮次内查找过程卡，不能借用另一任务的步骤。

Excel 与 WPS 复用 resumeActivityLog，保留完成行、累计耗时与原编号，恢复末行计时
时排除暂停间隔；不再为每次继续插入一个带 requestId 的编号步骤。语义判断选中
原任务后移除本次“理解消息”的临时过程卡。文本继续仍由模型判断，不限制为固定
词语；仅查询状态不继续。早期无检查点的重新分析仍按真实重读过程展示，不伪装成
SDK 断点执行。授权、来源复核与宿主写入路径未改。

> **2026-10-22 修订**：本节「接回原卡」的展示约定已被下文的**多回复段**取代。
> 保留的只有身份与时序那部分（task/turn ID、sequence、stepId 归并）：任务身份不变，
> 但展示上继续一定会**新开一段**，旧段冻结为历史。

### 2026-10-22：同一任务、多个回复段（过程卡不再唯一）

**作废的旧约定**：一任务绑定一张可变过程卡、继续后更新上方旧消息、"界面上只有一张
卡"被当成恢复正确的判据。它带来三个真实问题：继续后最新消息下面没有反馈、过程卡
被反复改写导致对话割裂、以及一张卡越滚越长。

**新约定**（判据 `packages/core/src/replySegments.ts` + 两端 `resumeSegmentRef`）：

| 维度 | 旧 | 新 |
| --- | --- | --- |
| 任务身份 | 一个 | **一个**（`operationId`/`stepId`/序号/结果引用/写入回执不变） |
| 回复段 | 一个 | **多个**（旧段 `superseded`、新段 `live`，`nextSegmentId` 成链） |
| 继续后进度落点 | 旧卡 | **用户消息之后的新回复段** |
| 旧卡 | 继续被改写 | **冻结当时状态** + 「查看后续进展」定位入口 |
| 步骤编号 | 接着旧卡数 | 新段**不重排**编号，只说明"继续（此前 N 步已完成）" |

四条不变量：①多段但只有一个活动段；②先更新共享任务状态再投影到段；③旧订阅抢不回
活动段（段切换不改步骤身份/事件游标/结果引用/回执）；④历史段保留当时状态、迟到
结果进权威记录并在最新段说明。

**测试口径也跟着改**：原来断言"只有一张过程卡"的用例改成行为判据（段数 +1、旧段
逐步文本逐字不变、活动段数 ≤ 1、新段位置在最新用户消息之后），**没有删除也没有
放宽断言**。

### 2026-10-22：Python 代码内容块（11 态 + 正文不进账本）

代码块的状态必须**有真实事件依据**（`packages/core/src/pythonCodeArtifact.ts`，
11 态：正在生成/已生成/正在校验/正在隔离运行/分析完成/正在整理结果/已暂停/**未开始
执行**/生成失败/校验未通过/执行失败）。当前 SDK 通道一次送达整段代码，因此第一条状态是
**"代码已生成"+"正在校验"，不说"正在逐字生成"**；`summarizing` 只有在上游真的
进入整理阶段时才允许出现，界面不得自行演绎。

> **这 11 态里有 3 态目前没有生产发出点**（`generating` / `summarizing` /
> `generation_failed` 只存在于 `server/app/models.py` 的枚举里），能由生产链路真正产生的是
> 8 态（`generated` / `validating` / `running` / `analysis_completed` / `paused` /
> `not_started` / `validation_failed` / `execution_failed`）。**定义了状态 ≠ 生产能产生
> 并验证它**；不要因为枚举齐全就宣称这 3 态已验证，也不要为了让状态表好看去造发出点。

> **2026-10-22 二次修订（缺陷修复）**：`not_started` 是这一批新增的第 11 态。
> 旧实现把"功能关闭 / 运行器不可用 / 没有可信绑定 / 授权拒绝 / 启动前取消"都走成
> `validating → running(executedVersion=1) → execution_failed`，而执行器**一次都没被
> 调用**（复现证据：`outputs/acceptance-2026-10-22/independent-review/python-status.json`，
> `runner_calls: 0`）。这不只是文案错：它宣称一段从未运行的代码"跑过一版"，还把用户
> 引向"重试执行失败"。现在的次序是：
>
> | 场景 | 事件序列 | 带 `executedVersion` |
> | --- | --- | --- |
> | 正常执行 | `generated → validating → running → analysis_completed` | 只有后两态带 |
> | 功能关闭 / 运行器缺失 / 无绑定 / 授权拒绝 | `generated → not_started` | 不 |
> | 入参非法（如未知字段） | `generated → validation_failed` | 不 |
> | 启动前取消 | `generated → validating → not_started`（说明"取消发生在启动之前"） | 不 |
> | 运行中取消 | `generated → validating → running → paused` | 是（它真跑过） |
> | 启动后异常 | `generated → validating → running → execution_failed` | 是 |
>
> **`running` 只有一个证据源**：`on_runner_start` —— 由
> `run_isolated_analysis_internal` 在配置/端点/镜像全部通过、**即将调用
> `run_python_isolated`** 时同步回调。执行前的确定性检查拆在
> `prepare_python_execution`（不碰执行器）里，真正起容器的是
> `run_prepared_execution`。取消与异常路径都必须留终态，不许停在 running。

**增量判重按事件身份，不按文本**（`applyPythonCodeEvent`）：判据是
`sequence`（迟到/重放整帧丢弃）→ `attemptId`（重新生成 = 重建工件）→ `version`
（内容版本）。**不同事件的相同片段必须原样保留**：`f(g(1` + `)` + `)` 拼出来是
`f(g(1))`，旧实现按 `code.endsWith(delta)` 去重会丢掉第二个右括号，把用户看到的代码
改坏（复现证据：`outputs/acceptance-2026-10-22/independent-review/code-delta.json`）。
缺片（版本跳过）记 `gap`：`complete` 压成 `false`、"复制全部"关闭，补齐（收到中间片或
收到完整快照）后自动恢复 —— **不许丢片后仍宣告完整**。完整快照整体替换正文。
> 验证层级：当前生产供应商**只返回完整代码**，增量路径的证据来自单元/业务用例（含
> 连续括号、连续换行、重复缩进、相同内容不同身份、同事件重放、乱序、缺片、快照覆盖、
> 旧尝试迟到），**没有**做过真实供应商流式验收。

**正文不进持久账本**（既有硬边界：`test_python_analysis.py` 断言 SQLite 行内不得出现
代码 sentinel）。因此事件只带 `codeId` / `attemptId` / 阶段 / 版本 / 字符数，正文登记
在进程内 `server/app/python_code_artifacts.py`（24 条上限、30 分钟 TTL），由界面按
`codeId` + **任务身份**经回环端点 `POST /api/sdk/python-code` 取回。取不回来（过期/
淘汰/任务对不上/服务未连通）一律 404 且不区分原因，界面直说"正文没有保留"。

**只读**：代码块没有任何执行入口（本仓库只有"生成白名单计划 → 预览 → 确认 → 执行"
一条写入路径）。"已运行"只在 `executedVersion === version` 时出现；未完成草稿不提供
复制全文；代码一律作为文本节点渲染。

### 2026-10-21：早期分析也接回原过程卡

步骤编号没有第 7 步特例。早期暂停引用现在保存原 activity ID 与已完成的消息理解结果；
恢复时沿用该卡，未完成的消息理解可以重试，已经完成的消息理解不再重新调用。
工作簿结构仍重新核验，未完成的字段分析重新调用模型；这是阶段接续，不是模型调用内部续算。
稳定的本机阶段 ID 避免核验过程重新追加已完成行，后续新阶段继续追加。
暂停操作放在原时间线标题栏，移除早期暂停的灰色说明卡；没有历史时间线时保留简洁入口。

> **2026-10-22 修订（缺陷修复）**：本条原文「恢复时沿用该卡」**已作废**。早期分析
> （理解消息 / 读表结构 / 字段意图）的继续现在也**新开一个回复段**
> （`resumeEarlyAnalysisActivity` → `openReplySegmentMessage`），旧卡保留暂停那一刻的
> 内容与耗时并给出「查看后续进展」。
> 任务侧行为不变：未完成的消息理解可以重试、**已完成的不再重新调用**、工作簿结构按现场
> 重新核验、未完成的字段分析重新调用模型。
> 这一段**没有**服务端任务身份（尚未形成规划检查点），因此段上**不写** `taskId`：不伪造
> 后端任务、也不宣称 SDK 内部断点续算，归属如实写成"本机按原需求重跑"。
> 三条入口（输入"继续"、空输入点发送、点卡片上的继续）都经 `restartEarlyAnalysis` 落到
> 同一条路径，行为一致。

当前实现边界：分析阶段可接收暂停，但取消接收不保证服务端计算已停止；已有任务仍先查询状态。
写入与回读期间目前禁用中断，尚未实现“当前动作完成并核对后暂停后续动作”。因此不能宣称
所有执行阶段均可暂停。原生 SDK 会话续算、跨窗格恢复早期引用也不在本次交付内。

### 2026-10-21：WPS 继续分析复用已确认授权

WPS 的 checkOriginalPlanningTarget 原先将对象身份 unknown 与 different 合并拦截，
一律声称范围已变化；原分析入口已接受的显式范围确认没有被续跑分支消费。宿主为
同一工作簿/工作表返回新 JS 包装对象时，可因此误拦继续分析。

继续检查点允许在 unknown 时复用已有分析确认：原任务证据键、当前证据键与已确认键
必须一致，且再次通过共享 checkForAnalysis；证据包括完整路径、表索引、表名及可用
稳定 ID。无确认、读不到或键变化仍拒绝。different 仍拒绝；unknown 保留原任务并
如实要求核实，不再声称范围已变化。此规则仅用于分析续跑；取回已完成结果仍保持
原有严格身份检查，来源指纹、服务端续跑证据、写入校验均未放宽。Excel 没有这条
WPS 原生对象引用附加闸门，沿用共享授权守卫并做回归。

### 2026-10-21：读取失败不撤销已确认的工作表授权

共享 useAuthorizedTarget 原先在 pause 的所有分支清空 analysisConsent，导致持续
读取失败耗尽重试后，同一范围恢复可读仍需重复确认。现在 unavailable 默认保留
原确认键；放行仍需重新读取完整一致证据并匹配原键，核实前停止分析。明确的文件/
范围变化、显式重新确认要求、以及多次读取之间的证据冲突仍清空确认键。冲突不能
因界面也使用 unavailable 文案而被当成暂时读取失败。

双端范围提示增加“重新检查”，未知状态不要求重新选择同一范围。recheck 只恢复
事先登记为分析的回调，不能把 check() 留下的写入回调重新分类成分析后自动执行。
勾选选择器现有草稿/确认交互未改变；本次解决已确认范围在宿主读取失败后重复失效。


### 2026-10-21：勾选即授权（覆盖此前草稿/确认交互）

用户明确要求 Excel / WPS 同步取消额外范围确认。现在每次勾选、取消勾选、全选、清空和模式选择
都由共享 useImmediateScopeAuthorization 调度自动核验。目录刷新与仅打开选择器不请求新授权；
收起保留选择。界面选择与已核验授权仍分开存储，仅作为异步核验期间的状态，不再要求用户提交草稿。

每次显式选择递增代数；旧核验跨 await 后必须通过 isCurrent 才能 adopt/bind。选择期间阻止读写和
后台范围提示，串行收尾接续最后一次选择；失败保留勾选并停止读写，刷新可重试，不自动重放旧请求。
共享守卫允许显式空范围撤销授权并清除分析确认；固定模式的空选择不回退到活动表。临时宿主读取失败
沿用前节的已确认键保留规则，真正的文件/表身份变化仍拦截。原有执行、来源指纹、回读及撤销校验保留。

Excel 与 WPS 工作目录两级勾选同样自动核验/提交。Excel 目录会话重建只复用本窗格已核实的同根同表
授权，不能再次拿旧 sessionId 撤销新会话授权；重建失败、根变化或表缺失均停止。WPS 工作目录服务端
查询/执行能力没有在本次扩展。两端是否完整实机通过仍须单独验收，自动化不能代替宿主证据。

界面补充：两端授权说明统一移到滚动容器外的固定标题下；底部完成按钮显示“确认”，仅关闭选择面板，
不触发授权提交或任务执行。目录模式保留底部已选计数，核验状态随顶部说明展示；列表按内容高度排列。

### 2026-10-22：运行中任务恢复读表回执与时间线收尾

Excel 与 WPS 的 running 恢复分支现在同时接回事件时间线和本机工具桥。接回前重新核验
原文件、已选工作表与来源指纹；状态查询跨 await 后必须再次检查宿主，不能只用查询前证据。
桥使用原 turnId，逐次执行仍读取当前授权。接回先查询服务端待处理调用，补偿暂停前已接收到
指针、但回执未完成的情况；只消费权威 pending 调用，不根据历史事件自行重放工具。
补发事件同时交给进度、结果引用和读表桥，工具回执沿已有幂等路径提交。状态补发 GET 本身
仍只读；本机工具执行不拥有写表权限。暂停后的自动核对不接回工具或轮询进度。

WPS 点击继续时若任务已完成，转入既有结果取回链；不重新规划。规划完成的结果只进入
正常结果/预览流程，删除计划继续保留专门的写入闸门。继续中的暂停详情显示“正在继续”，
同步 ref 防止重复恢复；在核对状态期间再次暂停也走中断处理，不显示为网络失败。

共享时间线对终态补发冻结步骤耗时，终态修正可以更新结论，但不能把步骤重新激活或重新计时。
WPS 本机消息理解完成即关闭对应步骤，早期用户暂停保持可恢复状态。任务仍在进行时，最后
一步完成不会让整张卡提前显示已完成。原卡/编号保留，未完成行继续接收事件。

前台计时保存统一起点和 pause/waiting 区间；本机分析停止计为暂停，已知服务端任务的断开
计为等待进度，因为后台可能仍在运行。两端共用组件展示前台耗时、暂停和等待时间，不把
它们当作服务端 CPU 时间。终态步骤及已经结束的卡冻结计时。

本节只补齐分析恢复与展示，不实现写入/回读中途暂停、SDK 会话持久化或跨窗格早期恢复。
双端实现与组件回归已完成；真实 Excel/WPS 原生窗格验收仍待执行。


### 2026-10-22：模型额度与长度限制提示

模型失败的 JSON / SSE 服务错误增加稳定 code：MODEL_QUOTA_EXHAUSTED、MODEL_CONTEXT_LIMIT、
MODEL_OUTPUT_LIMIT、MODEL_BUDGET_LIMIT、MODEL_TURN_LIMIT；错误信封结构不变，两端契约与
server/app/models.py 同步。HTTP 402 或供应商错误元数据明确说明余额/额度耗尽时才归类为
额度不足；普通 429 继续表示临时限流，不推断账户余额。硬限制 retryable=false，不自动重试。

SDK 适配器保留最后一次 AssistantMessage 的错误类型，以及失败 ResultMessage / ResultError
中的状态、错误码、预算/轮数终止原因和输出截断标记。正常回答不作为错误分类依据，供应商
原始错误内容只用于识别，不进入上述限制提示。SDK 规划步骤在模型错误时发送 failed 终态。
意图与连接测试入口同样保留额度错误码；不会用泛化的连接失败掩盖额度不足。

Excel 与 WPS 共用 ApiRequestError 和 packages/core/src/modelUsageNotice.ts，将已知 code
转换为中文原因与操作建议，并保留请求以便处理额度后重试。提醒说明模型调用受限；不能据此
宣称工作表未写入、写入失败或已经撤销。写入与回读结论由本地执行记录及结果卡提供。本轮不
新增余额查询接口、不猜测剩余 token、不自动切模型，不改变预览、写入确认及原有恢复边界。
