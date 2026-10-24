# SDK 任务恢复与会话恢复设计（2026-10-19）

本文件是**唯一**的恢复设计文档。它替代此前散落在 `CAP02_WPS_PLANNING_RESUME_2026-10-06.md`、
`CONTINUOUS_CHAIN_ACCEPTANCE_2026-10-10.md`、`SDK_CAPABILITY_ASSESSMENT_2026-10-09.md`
里的"恢复"设想。与本文冲突的旧描述以本文 + 生产代码 + 有效配置为准。

本轮**不改生产行为、不打开持久化开关、不重启服务、不操作业务工作簿**。本文只回答四件事：
现在到底有什么、推荐怎么接、怎么验证、还差什么。

---

## 0. 结论摘要（先看这段）

1. **应用侧已经有一套完整的恢复机制，但它只覆盖"执行"这一层，规划层在 SDK 后端上是空的。**
   `planning_tasks` / `planning_task_events` / `office_executions` 三张表、CAS 续跑、
   owner token 轮换、动作级在途标记、真实宿主核对，全部已实现并有测试；
   缺的是**规划层能产出一个 SDK 可用的检查点**。
2. **今天 SDK 规划任务一旦断线/出错就是 `failed`，没有任何"继续"。**
   原因链是三处：`run_sdk_plan` 不接 `on_checkpoint`（`server/app/planner.py:41`）⇒
   账本永远没有 checkpoint ⇒ `_planning_pause_payload` 里 `canContinue` 恒假
   （`server/app/main.py:1787-1793`）⇒ 出错分支走 `finish_planning_task(status="failed")`
   （`server/app/main.py:2005-2012`）。客户端那一整套续跑 UI（`App.tsx:2810-3075`）因此
   **对 SDK 任务永远走不到**。
3. **推荐方案 A 为主干**：用应用自己的账本重建规划上下文，把"继续分析"做成
   **一次新的规划**（可能重读、重算、重新取结果），写入阶段仍由执行账本防重放。
   **方案 B（SDK 原生 `resume`）作为最后阶段的可选加速器**，前置条件是隐私与生命周期评审；
   在 B 落地前，**任何"恢复了模型会话"的说法都是不成立的**。
4. **最关键的风险不是技术，是语义**：`office_executions.plan_json` 已经把
   **含单元格值的完整计划**写进磁盘且**没有任何清理路径**（`server/app/task_ledger.py:1611`、
   全文无 `DELETE FROM office_executions`），而规划账本反而是去值的
   （`server/app/ledger_privacy.py:154`）。打开 SDK 转录之前，必须先把这个不对称讲清楚。
5. **真正需要用户决定的只有三件事 + 一个附带问**（第 10 节）：转录是否落盘/落哪/留多久
   （连带：磁盘上已有的 30 个历史转录怎么处置）；执行账本含值计划是否纳入保留期；
   "继续分析 = 重新规划"这个产品措辞能不能接受。

---

## 1. 当前代码事实（带位置）

### 1.1 SDK 适配器：能规划，但不留痕、不可续

| 事实 | 位置 | 说明 |
| --- | --- | --- |
| 显式关闭会话持久化 | `server/app/claude_sdk_agent.py:111` | `extra_args={"no-session-persistence": None}` |
| 捆绑 CLI 版本 | `.venv/Lib/site-packages/claude_agent_sdk/_cli_version.py` | `2.1.281`；SDK 包 `0.2.159` |
| CLI 对该开关的语义（只读 `--help` 实测） | 同上 CLI | 原文：`Disable session persistence - sessions will not be saved to disk and cannot be resumed (only works with --print)`。**"cannot be resumed" 是 CLI 自己写的**，不是推断 |
| `ResultMessage` 只取文本 | `server/app/claude_sdk_agent.py:125-191` | `_result()` 只读 `final.result`；`session_id` 既未使用也未持久化 |
| 无检查点回调 | `server/app/planner.py:41` | `create_plan` 收了 `on_checkpoint` 却**不传给** `run_sdk_plan` —— 这是一个**死参数** |
| 跨后端续跑直接拒绝 | `server/app/planner.py:29-33` | `resume_checkpoint is not None` ⇒ `ModelNotAvailableError("Claude Agent SDK 不能恢复另一规划后端的检查点…")` |
| 会话目录被强制隔离 | `server/app/claude_sdk_agent.py:61-71` | `CLAUDE_AGENT_SDK_STATE_DIR` 必须是**仓库外**绝对路径，`mode=0o700`，并作为 `CLAUDE_CONFIG_DIR` + `cwd`；默认值 `%LOCALAPPDATA%\ExcelBro\ClaudeAgentSdk`（`scripts/start_claude_sdk_server.ps1:23`） |

**只读探测结果（2026-10-19 晚，实测，仅计数不读内容）**：

| 观测 | 结果 | 含义 |
| --- | --- | --- |
| 会话目录现状 | `projects/` 下 30 个 `.jsonl`，合计约 2 MB；最新一个停在 **2026-10-05 18:52** | 10-19 的真实模型/SDK 运行**没有**产生新转录 ⇒ `--no-session-persistence` **当前确实生效** |
| 目录内最新文件 | `.last-cleanup` 为 2026-10-19 09:49（CLI 自己写的），但**没有** 10-19 的 `.jsonl` | 同上，不是"没跑过"而是"跑了没落盘" |
| 16/30 个转录含 `writeValues` 字样 | 但 **0** 个含字面值矩阵（`"values":[[`）、**0** 个含 `"rows":[[`、**0** 个含 `sk-` | 那 16 个是**工具定义**里的字段名，不是写入值；早期转录**未证实**含单元格值，也未证实干净 |
| 清理代码 | 全仓无任何删除 `CLAUDE_AGENT_SDK_STATE_DIR` 内容的路径（`_state_dir()` 只 `mkdir`） | 这 30 个旧转录**永久留存**；A2 必须处理 |

> 结论：**不能写"SDK 目录里没有转录"**。准确表述是"当前开关有效、但历史上已留下 30 个旧转录，
> 且应用没有清理路径"。

**由此可判定的两件事**：
- 当前"打开持久化"不是改一个布尔值。CLI 明确写了持久化关闭时**不可 resume**，
  所以要么走 SDK 的 `session_store`（由 SDK 物化到临时 config dir），要么放弃 `--print` 语义。
  两条都要单独取证，不能靠"SDK 有 `resume` 参数"推断。
- `on_checkpoint` 是死参数这件事本身是一条**可测的判别点**：给它塞一个必须被调用的
  计数器，SDK 路径下计数恒 0。第 9 节的自动化用例 M1 就钉这一条。

### 1.2 规划账本：完整，但 SDK 侧喂不进去

| 事实 | 位置 |
| --- | --- |
| 表结构（`running/waiting_user/completed/failed/cancelled`、`checkpoint_json`、`result_json`、`revision`、`last_sequence`、`active_elapsed_seconds`） | `server/app/task_ledger.py:354-386` |
| 事件序号连续 + executor_id 校验 | `server/app/task_ledger.py` 的 `append_planning_event` |
| 续跑 CAS（`expected_revision` + 稳定 `requestId`） | `server/app/task_ledger.py` 的 `claim_planning_continuation` |
| 预算扩展幂等表 `(turn_id, request_id)` | `server/app/task_ledger.py:387-396` |
| 成功完成写**版本化恢复描述** | `server/app/main.py:1964-1970` → `ledger_privacy.ledger_recovery_descriptor` |
| 重连先重放事件，再从 `after_sequence` 增量 | `server/app/main.py:1864-1866` |
| 断流时**取消子任务但不改状态** | `server/app/main.py:2045-2060`（`finally` 里 cancel + await） |
| 状态端点把"不在活跃集合里的 running"判成 `interrupted` | `server/app/main.py:2172-2178` |
| 完成态按 `valuesOmitted` 分流：完整 `result` 或**只给** `recovery` | `server/app/main.py:2163-2171` |

**`AgentBudgetPause` 在生产路径不可达**：唯一抛出点在 `server/app/excel_agent.py:2580`，
位于 `run_excel_agent(...)`（`:2280`）内；而 `run_excel_agent` **没有任何生产调用点**
（只剩 `server/tests/test_agent_integration.py` 直接调用它，以及 `.tmp/runtime-smoke`
里的打包快照）。所以 `pause_planning_task` 现实上只剩 `MODEL_*` 可重试错误一条入口，
而那条入口要求 `current.checkpoint is not None` —— SDK 永远为假 ⇒ **直接 failed**。
结论：**"预算暂停/继续"这套 UI 对 SDK 任务没有机会出现**，不能当成已验证能力。

### 1.3 执行账本：这一层是真的能用

| 能力 | 位置 | 关键语义 |
| --- | --- | --- |
| 执行登记 | `server/app/task_ledger.py:1611` | `identity_prefix` 命名空间：Office `url:` / WPS `wps-path:`；同 ID 异指纹/异工作簿/异任务一律拒绝 |
| 动作开始/完成 | `begin_office_action` / `complete_office_action` | 写入前落 `in_flight_action_index` + `before_snapshot_json` |
| 挂起 | `suspend_office_execution` | 有在途动作 ⇒ `needs_reconciliation` |
| 恢复认领 | `server/app/task_ledger.py:1831` | `revision` CAS + `in_flight_action_index IS NULL` + **owner token 轮换**（旧任务窗格从此无法提交回执） |
| 在途核对 | `reconcile_office_action` | `applied` 必须带完整检查点；`not_applied` 退回原序号 |
| 收尾 | `finish_office_execution` | `next_action_index` 必须等于动作数 |
| WPS 单写者 | `_assert_wps_idle` + `taskRecovery.wpsOwnerLeaseSeconds=120` | 同机 WPS 路径同一时刻只允许一个持有者 |
| 真实宿主核对（Excel） | `apps/excel-addin/src/excel.ts:5450` | 读真实区域，与"期望后态 / 写前快照"比对，返回 `applied` / `not_applied` / `unknown`；身份不一致直接抛错 |

### 1.4 客户端：状态机已经写好了，只是喂不进去

- 规划续跑：`App.tsx:2810`（先 `targetGuard.check()`）→ `getPlanningTaskStatus`
  → 按 `completed/failed/running/canContinue` 分流 → 授权集合与来源指纹比对
  （`App.tsx:2978-3010`）→ `continuePlanningTask`（带 `expectedRevision` + `requestId`
  + `afterSequence`，`App.tsx:3052-3075`）。
- **"自动核对只查状态，检查点可用也不自动 POST"** 写在 `App.tsx:2921-2928` 的注释与分支里，
  这是场景 ⑧ 的现有实现。
- 完成态回放**绝不**把 `recovery` 当完整计划渲染（`App.tsx:2842-2860`）。
- 执行恢复：`App.tsx:2409-2532` —— 先 `getOfficeExecutionStatus`，在途则
  `reconcileOfficeRangeAction`，否则 `claimOfficeExecutionResume`，再以
  `recovery.startActionIndex` 继续执行。
- 历史落盘脱敏只覆盖 `title` 与 `message.text`
  （`apps/excel-addin/src/conversation.ts:96`），注释明确写了**不处理** `plan` /
  `resultContext` / `activityLog`。`resultContext.rows` 因此**会进 localStorage**。

### 1.5 易失数据与它们的真实寿命

| 易失物 | 寿命 | 位置 |
| --- | --- | --- |
| 本地结果引用（`res_<32hex>`） | `agentToolBridge.resultTtlSeconds=1800`，且全局上限 `maxRetainedResults=200`；**进程重启即全丢**（内存注册表） | `server/app/local_results.py:251` |
| 结果引用的归属 | `OwnerClaim` 绑 `host_id` + 工作簿名 + 工作簿身份 + 授权键，`resolve_owned` 跨轮仍校验 | `server/app/local_results.py:336` |
| 规划轮内状态 | `TurnRegistry` 内存，`turns.ttlSeconds=86400`、`maxActive=500` | `server/app/turn_state.py` |
| 工具桥操作 | `operationTtlSeconds=900` | `config/capabilities.json` |
| Python 数据范围 | `pythonData.scopeTtlSeconds=900` | 同上 |
| 结果引用失效后的表现 | `404 BRIDGE_RESULT_UNAVAILABLE`「本地结果不存在或已过期，请重新查询。」 | `server/app/main.py:428-440` |

**这条决定了方案 A 的边界**：结果引用不可持久化，所以"继续分析"在服务重启后
**必须允许重读重算**，而不能假装引用还在。设计上要求：引用不可用时**不伪造**，
显式降级为"重新读取来源并重算"，并把这件事写进用户可见文案。

### 1.6 保留期与清理（现状是不对称的）

| 表 | 有清理吗 | 依据 |
| --- | --- | --- |
| `folder_operations` | 有，`taskRecovery.operationRetentionDays=30`，**只留墓碑** | `server/app/main.py:184-189` → `task_ledger.py:862` |
| `planning_tasks` / `planning_task_events` | **没有** | 全仓无 `DELETE FROM planning_tasks` |
| `office_executions` | **没有**，且 `plan_json` 是**含值**的完整计划（`TEXT NOT NULL`，无投影） | `server/app/task_ledger.py:397-420`、`:1630` |

账本物理位置：`%LOCALAPPDATA%\Excel Bro\tasks\state.sqlite3`
（`server/app/app_data.py:30`；macOS 为 `~/Library/Application Support/Excel Bro`）。

---

## 2. 四种用户操作必须分开，别互相冒充

| 操作 | 允许做什么 | 明确禁止 | 现有落点 |
| --- | --- | --- | --- |
| **查询任务状态** | 只读账本与活跃集合，回答"做到哪一步" | 不调模型、不写表、不重放工具、不自动重试、不自动 POST 继续 | `GET /api/turn/{turn_id}/status`（`main.py:2144`）、`GET /api/office/executions/{id}` |
| **继续分析** | 重新核实工作簿身份 / 授权范围 / 来源指纹 / 计划版本 / 结果引用有效性；必要时**重读、重算、重新取结果**；产出一份新计划 | 不重放已完成的写入；不因为"客户端没收到响应"就再执行；不宣称恢复了模型会话 | `POST /api/turn/continue/stream`（`main.py:2190`）→ `App.tsx:3052` |
| **继续执行** | 只处理经核实**尚未完成且可安全继续**的剩余动作；在途动作必须先核对现场 | 现场不明却直接写；靠"没收到成功响应"判失败；换 owner 后旧窗格再提交回执 | `claim_office_resume` + `executePlan(startActionIndex)`（`App.tsx:2474-2532`） |
| **撤销本次执行** | 走既有恢复记录，在**原工作簿原地**撤回 | 改成打开副本；与"继续任务"混成一个按钮；把撤销当"再跑一次" | `worksheet_restore.py` + 两端审阅卡的「↶ 撤销本次执行」 |

**命名纪律**（写进代码注释与用户文案）：
- 重新发起一次规划 **不叫**"SDK 原生会话恢复"，叫"**按当前现场重新分析**"。
- 把剩余动作跑完 **不叫**"断点续跑 SDK 会话"，叫"**继续未完成的写入**"。
- `resume` / `continue` / `reconcile` / `undo` 四个词在协议、日志、文案里各指一件事，
  不许混用。

---

## 3. 方案 A vs 方案 B

### 方案 A：复用应用账本，重建必要的规划上下文

**做法**：把"续跑规划所需的最小事实"写进 `planning_tasks.checkpoint_json`，
形成**应用自有**的 planning checkpoint：

```
{
  "schemaVersion": 1,
  "kind": "planning_context",
  "originalPrompt": "<用户原始需求，不含数据行>",
  "confirmations": [ { "questionId", "answerLabel", "answeredAt" } ],
  "resolvedTarget": { "identity", "evidenceKey", "sheetNames" },   // 只留身份与成员，不留数据
  "toolOutcomes": [ { "tool", "status", "resultRef"|null, "structureOnly": true } ],
  "authorizationRevision": 3,
  "sourceFingerprint": "<sha256>",
  "sourceFingerprintSheets": ["会员表","订单表","汇总表"]
}
```

续跑 = 用这份 checkpoint **重新组装一次 `PlanRequest`** 交给 `run_sdk_plan`。
`resultRef` 只是**提示**（可能已过期）：取不到就按 1.5 节降级为重读重算。

**一个已存在、可直接复用的关键点**：`POST /api/turn/continue/stream` 今天**已经**在用
`claimed.request` 重建 `PlanRequest`（`main.py:2324-2326`：`PlanRequest.model_validate(
claimed.request).model_copy(update={"modelId": claimed.current_model_id})`）。
也就是说"重建请求"这一步**不需要新写**；A0 真正要补的只有三件：
① 让 `checkpoint_json` 非空（否则 `canContinue` 恒假，连入口都进不去）；
② `planner.py` 不再对 `kind="planning_context"` 一律抛错；
③ 把已确认的澄清与已完成工具结果作为**提示**拼进规划上下文，避免重复澄清、无谓重读。

- **优点**：完全复用已验证的账本、CAS、事件回放与客户端状态机；不引入新的隐私面；
  与"隐私前不打开转录"的现行边界一致；Excel/WPS 天然同构。
- **缺点**：不是"真续跑"——模型的思考过程、已消耗的上下文不会保留，续跑后可能重走
  已完成的推理；成本不可控（第 10 节的措辞决策由此而来）。
- **必须诚实的地方**：这是"**重新规划 + 复用已核实的结果引用**"，不是会话恢复。

### 方案 B：SDK 原生会话恢复

**做法**：不再传 `--no-session-persistence`；用 `session_store` 或 CLI 默认目录保存转录，
续跑时 `resume=session_id`（可选 `fork_session` 保留原会话）。

**已核实的硬约束**：
- CLI `2.1.281` 自己写明：持久化关闭 ⇒ **sessions cannot be resumed**。所以 B 的前置
  动作**不是**"删掉那个 extra_arg"，而是先把"转录写到哪、谁能读、留多久"定下来。
- SDK `0.2.159` 提供 `materialize_resume_session`：把 store 里的会话**物化到临时
  `CLAUDE_CONFIG_DIR`** 并复制认证文件，结束后清理。这条是目前唯一"不把转录长期留在
  用户主目录"的可用路径，但它是 SDK 私有流程，需要夹具实验确认（第 9 节 M4）。
- 转录内容由 CLI 决定，**应用无法逐字段裁剪**。模型看到的每个工具结果都会进转录，
  而本项目的工具结果**就是数据行**（`query_table` / `submit_local_result`）。
  这与 1.4 节"历史落盘要脱敏"是同一条红线，但转录的泄漏面更大（原始工具输出）。

**结论**：B 不是"能不能做"，是"**允不允许**"。它必须建立在三个前提上：
① 工具结果进 SDK 上下文之前就已经是**去值投影**（今天是：SDK 只能拿到结构 + 引用，
`planningReadPolicy="none"` —— 所以这一条**恰好已经成立**，见 `planner.py:38-41`）；
② 转录目录由应用控制并纳入清理；③ 隐私用例能证明转录无机密、无数据行。

### 分阶段推荐

| 阶段 | 内容 | 为什么这个顺序 |
| --- | --- | --- |
| **A0** | 让 SDK 规划任务**有** checkpoint（应用自有上下文），"继续分析"变成真的重新规划 | 补的是"今天完全没有"的能力；不碰隐私开关；改动集中在 `planner.py` + `main.py` + 共享层 |
| **A1** | 执行恢复的场景补齐 + 双端实机验收（实现已存在，缺证据与边界用例） | 已有的东西先取证，避免在没证据的层上继续叠加 |
| **A2** | 易失引用与账本生命周期的收口（保留期、清理、损坏记录、旧版本迁移） | A0/A1 都会把记录留得更久，必须先有清理，否则是在放大已有问题 |
| **B** | 原生会话恢复（先夹具实验 M4，再决定是否打开） | 依赖 A2 的清理与 A0 的语义分层；隐私评审不过就不做 |

**不推荐**：跳过 A0 直接做 B。B 只能续"模型上下文"，解决不了"SDK 任务出错即 failed、
`canContinue` 恒假"这个当前最痛的点；而且 B 一上来就要动隐私开关。

---

## 4. 状态转换表

### 4.1 规划任务状态（可信事实 = `planning_tasks` 行 + `_active_planning_tasks` 集合）

| 状态 | 可信事实持有者 | 进入事件 | 退出事件 | 刷新/重连后如何重建 | 用户可点 | 禁止动作 |
| --- | --- | --- | --- | --- | --- | --- |
| `running`（活跃） | 账本 + 内存活跃集合（两者同时成立） | 新任务 `register_planning_task`；或续跑认领成功 | 完成 / 暂停 / 失败 / 断流 | 活跃集合命中 ⇒ `status:"running"`、`canContinue:false` | 只能"查看进度"，不能继续 | 并发第二个 SSE；自动重试 |
| `running`（**已中断**） | 账本 `status='running'` **且** 不在活跃集合 | 服务重启 / 流被丢弃 / 进程退出 | 用户点继续（认领成功）或任务被判 failed | 状态端点返回 `interrupted`；`pauseType="interrupted"`（`main.py:2174`） | 继续（若 `checkpoint` 非空）、切备用模型 | **认定成功或失败**；盲目重试 |
| `waiting_user`（澄清中） | 账本 + 澄清记录（`confirmations`） | 模型要澄清 | 用户回答后重新规划 / 放弃 | 事件回放 + `_planning_pause_payload` | 继续、换模型 | 把"等待用户"当"卡死"自动重试 |
| `waiting_user`（预算暂停，**当前不可达**） | 账本 `active_elapsed_seconds` | 预算耗尽 | 继续（追加时长） | `pauseType="budget"` / `budget_exhausted` | 继续（受 `planningMaxExtensions=4` 限） | 承诺"按一下就能跑完" |
| `completed`（值未省略） | `result_json` 里的完整 `response` | `finish_planning_task(completed)` | 终态 | 状态端点回 `result` | 展示既有结果、保存工具、撤销 | **自动重跑**；再次执行写入 |
| `completed`（值已省略） | `result_json` 里的恢复描述（`valuesOmitted:true` + `resultRef`） | 同上 | 终态 | 只回 `recovery`，**不回** `result` | 展示"账本未存值"、按引用看结果（可能已过期）、重新运行 | 把恢复描述当完整计划渲染（`App.tsx:2842`） |
| `failed` | `error_json` | 不可重试错误 / 无 checkpoint 的可重试错误 | 终态 | 状态端点回 `error` | 重新发起 | 从失败态"继续" |
| `cancelled` | `error_json` | 用户取消 | 终态 | 同上 | 重新发起 | 续跑 |
| 不可继续（原因 A：无 checkpoint） | `checkpoint_json IS NULL` | — | — | `canContinue:false` + `reason` | 重新发起 | 给"继续"按钮 |
| 不可继续（原因 B：总时长耗尽，**当前不可达**） | `active_elapsed_seconds >= planningHardWallTimeSeconds(1800)` | — | — | `pauseType="budget_exhausted"`（`main.py:1755-1760`） | 重新发起 | 与原因 A 糊成同一句文案（`App.tsx:2904-2905` 已分开） |

> 两个"不可达"的含义：`AgentBudgetPause` 只能由 `run_excel_agent`（`excel_agent.py:2580`）
> 抛出，而它已无生产调用点。这两行是**契约占位**，A0 落地后应重新评估是否恢复预算暂停，
> 或者把它们从用户可见面撤掉 —— 留着一个永远不会出现的"继续"入口同样是在说谎。

### 4.2 单次执行状态（可信事实 = `office_executions` 行）

| 状态 | 可信事实 | 进入事件 | 退出事件 | 重建方式 | 用户可点 | 禁止动作 |
| --- | --- | --- | --- | --- | --- | --- |
| `running`（无在途） | 账本 `in_flight_action_index IS NULL` | `register` / `claim_resume` / 动作完成后 | 下一个动作开始 | `GET status` | 继续（CAS 认领） | 两个窗格同时写（owner token 会互踢） |
| `running`（**有在途**） | `in_flight_action_index != NULL` + `before_snapshot_json` | `begin_office_action` | `complete` / `suspend` / `reconcile` | `GET status` 显示在途序号 | 先核对（`reconcile`），不能直接 `claim_resume` | 直接写下一个动作；凭"没收到成功响应"判失败 |
| `needs_reconciliation` | 同上 | `suspend` 且有在途 | `reconcile(applied/not_applied)` | 同上 | 核对待续跑状态 | 认领后不核对就续写（`claim_office_resume` 的 `in_flight IS NULL` 条件会拒） |
| `completed` | `next_action_index == len(actions)` | `finish_office_execution` | 终态 | `GET status` | 撤销、保存工具 | 重放任何动作 |
| `failed` | `error_json` | 不可继续的错误 | 终态 | 同上 | 撤销已写部分、重来 | 从失败态静默续跑 |
| 部分完成 | **不是独立状态**：由 `next_action_index > 0 且 < len(actions)` + `status='running'` 表达 | — | — | 由两者组合判断 | 继续剩余动作 | 把"部分完成"写成"已完成" |

### 4.3 结果审阅状态（可信事实 = 客户端审阅卡 + 宿主实际状态）

| 状态 | 可信事实 | 进入 | 退出 | 重建 | 可点 | 禁止 |
| --- | --- | --- | --- | --- | --- | --- |
| 已执行待回读 | 客户端回读结果（尚未落盘） | 计划到达后**自动执行一次**完成 | 回读成功 / 失败 | 无法重建 ⇒ 视为"现场不明" | — | 在回读未完成时允许确认 |
| 待确认（标记已置） | 宿主单元格标记 | 回读+标记完成 | 确认 / 撤销 / 关闭 | 重开工作簿时标记仍在 ⇒ 由标记恢复（`recoverExcelResultMarks`） | 确认变更（只清标记）、双视图、保存工具、撤销 | **确认时再次执行** |
| 已确认 | 标记已清 | 「确认变更」 | 终态 | 标记不在 ⇒ 视为已确认 | 保存工具、撤销 | 把确认当"再次写入" |
| 已撤销 | 撤销记录 + 单元格已回原值 | 「↶ 撤销本次执行」 | 终态 | 由撤销记录判断 | — | 把撤销实现成开副本 |
| 现场不明 | 无 | 回读失败 / 换簿 / 宿主断开 | 核实成功 | — | 核对 | 声明成功或失败 |

---

## 5. 恢复决策表（九种故障场景）

| # | 场景 | 处理方式 | 判据（谁来判） | 测试落点 |
| --- | --- | --- | --- | --- |
| ① | 写入成功但响应丢失，用户点"继续" | 先核对在途动作的**真实宿主状态**，判 `applied` 则补记检查点并跳到下一动作；`not_applied` 则退回原序号重做；`unknown` 保持挂起并如实告知 | `reconcileOfficeRangeAction`（读真实区域 vs 写前快照）。**不读客户端"是否收到响应"** | 单测：`excel.ts` 核对三分支；集成：在途注入后 `claim_resume` 必须被拒（`in_flight IS NULL`） |
| ② | `running` 但宿主断开/服务重启 | 不判成功也不判失败：规划层 ⇒ `interrupted` + `canContinue` 视 checkpoint；执行层 ⇒ 有在途则 `needs_reconciliation`，无在途则可 CAS 认领 | 服务端：活跃集合 / `in_flight_action_index`；客户端：`targetGuard` + 身份比对 | 集成：停服务后 `GET status` 形态；`claim_office_resume` 在在途时返回 409 |
| ③ | 多动作执行到一半中断 | 逐动作记账：`next_action_index` 之前的**已验证完成**；`in_flight` 那一刻**结果未知**；之后的**确定未执行**。三类分别可写/须核对/未动 | `next_action_index` + `in_flight_action_index` 两个字段的组合，不引入第三份状态 | 单测：3 动作中途挂起 ⇒ 断言 `next=1, in_flight=1`；恢复后只跑 1..2 |
| ④ | 暂停/取消后旧模型或工具结果迟到 | 迟到结果只能写事件（且必须过 `append_planning_event` 的序号/executor 校验），**不能**形成新计划或新写入；已取消任务的 SSE 连接不得再被续上 | 账本序号 + `executor_id` + 任务状态；执行层则 owner token 轮换 | 单测：取消后到达的事件被拒（序号冲突/状态不允许）；执行层：旧 token 提交回执失败 |
| ⑤ | 恢复期间换簿 / 收窄授权 / 改来源或计划 | 一律拒绝过期现场，给固定错误码；不"尽力继续"。授权集合比对用共享层规范形；来源用指纹 | `canonicalTargetSelection` / `authorizedTargetEvidenceKey` / `sourceFingerprint`；执行层 `workbook_identity` 全串比较 | 已有：`run_scope_identity_mutations.mjs` I1–I4；新增：恢复期间改来源 ⇒ 客户端在 POST 前拦；服务端 `request_fingerprint` 不符 ⇒ 409 |
| ⑥ | 双击继续 / 两个客户端同时恢复 | 服务端 CAS：规划层 `expected_revision` + 稳定 `requestId` + 活跃集合二次检查；执行层 `revision` CAS + owner token 轮换 | `claim_planning_continuation` / `claim_office_resume` | 集成：并发两个 continue ⇒ 一个 200 一个 409 `PLANNING_TASK_STILL_RUNNING`；执行层旧 token 失效 |
| ⑦ | 结果引用过期 / 服务重启丢失 | **不伪造可用**：`404 BRIDGE_RESULT_UNAVAILABLE` ⇒ 客户端显式降级为"重新读取来源并重算"；已完成的账本记录仍按 `valuesOmitted` 如实展示 | `local_results.resolve` 返回 None；账本 `valuesOmitted` | 单测：TTL 过期（`resultTtlSeconds=1800`）后取结果 404；集成：重启后取旧 ref ⇒ 404 且 UI 不崩 |
| ⑧ | 已完成任务再打开展示既有结果 | 只查状态、只回放既有结果；**不自动重跑**；`valuesOmitted` 的无值记录如实说"账本未存值" | `GET status` 的 `completed` 分支 + `App.tsx:2921` 的"自动核对只查状态" | 已有断言：`App.tsx:2921-2928`；新增：连续两次打开只发一次 GET，无 POST |
| ⑨ | 同名文件/同名表/旧会话/其他任务回执串用 | 身份不用文件名/路径当唯一判据：Office 用 `url:` 命名空间 + `Office.context.document` 引用；WPS 用 `wps-path:` + 创建时登记的 `name+fullName+path`；回执校验 `task_id` 绑定 | `register_office_execution` 的三重拒绝（异指纹/异工作簿/异任务）；`resolve_owned` 的 claim 校验 | 单测：同名不同簿 ⇒ 拒绝；`task_id` 不符 ⇒ 拒绝；跨前缀（`url:` vs `wps-path:`）⇒ 拒绝 |

### 5.1 operationId 规则（防"每次新 ID 绕过幂等"）

- **恢复旧操作必须复用旧 `operationId`**，并附 `expectedRevision`。
  `operationId` 只在"**用户明确发起一次新任务**"时生成。
- 生成规则保持现状：客户端一次执行会话一个 ID，写进 `officeOperationRecovery` 落盘；
  重开面板时**先读旧 ID 再决定**，不先造新 ID。
- 服务端用三重绑定兜底：`request_fingerprint`（同 ID 异请求 ⇒ 拒绝）、
  `workbook_identity`（异簿 ⇒ 拒绝）、`task_id`（异任务 ⇒ 拒绝）。
- **禁止**：为了"重试"而换 ID；用新 ID 去重放已完成的动作。
- 规划层同理：`requestId` 在一次"继续"里必须稳定（重连重发用同一个），
  换 `requestId` 就等于把幂等让给用户手速。

---

## 6. 最小持久记录与隐私边界

### 6.1 规划层（`planning_tasks` + 事件）

| 字段 | 用途 | 来源 | 敏感 | 位置 | 保留 | 删除条件 | 恢复时如何复核 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `turn_id` | 任务身份 | 客户端生成 | 否 | `state.sqlite3` | 跟随记录 | 记录删除 | 与客户端消息绑定一致 |
| `request_json` | 重建 PlanRequest 的骨架（需求文本、目标身份、口径） | `PlanRequest` | **中**（含用户需求原文，可能含业务内容） | 同上 | 同上 | 同上 | 重新核对目标身份与授权成员 |
| `request_fingerprint` | 判"同 ID 异请求" | 规范化后的请求 | 否 | 同上 | 同上 | 同上 | 恢复时**必须**重算比对 |
| `checkpoint_json` | **A0 新增**：规划上下文（见 3 节结构） | `on_checkpoint` | **低**（只允许白名单字段，**无值**） | 同上 | 同上 | 同上 | `schemaVersion` + 成员/身份/指纹逐项比对 |
| `result_json` | 完成后的恢复描述（`valuesOmitted` + `resultRef` + 去值 response） | `ledger_recovery_descriptor` | **中**（`summary`/`title` 是模型自由文本，可能含源自数据的表述） | 同上 | 同上 | 同上 | `valuesOmitted` 决定能否当完整响应 |
| 事件 `payload_json` | 进度回放 | `on_event` | **中**（`detail`/`facts`/`target` 是自由文本） | 同上 | 同上 | 同上 | 只用于展示，不驱动写入 |
| `error_json` | 失败原因 | 错误映射 | **低**（已是中文固定串；`str(error)` 需再次确认） | 同上 | 同上 | 同上 | 只展示 |
| `executor_id` | 判"这份账本属不属于当前进程" | 启动时生成 | 否 | 同上 | 同上 | 同上 | 不一致 ⇒ 视为中断 |

**红线（规划层）**：原始数据行、`values`/`rows`/`headers`/`formulas`/`expected` 矩阵、
带数据的代码、密钥、工具原始输出，**一律不进**。已有实现的判别点就是
`ledger_privacy.MATRIX_FIELD_NAMES`（按**键名**判定，新增类型忘记登记也不会漏值）。
A0 的 checkpoint 必须复用同一套键名判定，**不许另写一份**。

### 6.2 执行层（`office_executions`）

| 字段 | 用途 | 敏感 | 现状问题 |
| --- | --- | --- | --- |
| `plan_json` | 断点续跑必须知道"还要写什么" | **高：含 `values` 矩阵 = 单元格数据** | **全量落盘、无投影、无清理**（`task_ledger.py:1630`） |
| `authorization_json` | 续跑前复核授权 | 低（表名清单） | 无清理 |
| `before_snapshot_json` | 撤销与核对回原值 | **高：含原值** | 同上 |
| `checkpoint_json` | 已应用动作的累积快照 | **高：含值/格式** | 同上 |
| `result_json` / `error_json` | 结果与失败 | 中 | 无清理 |
| `owner_token` / `revision` | 单写者 + CAS | 否 | 无 |

**这是本设计里最需要用户拍板的地方**：执行账本**必须**含值（否则撤销与核对无法成立），
但它今天**永久留在磁盘上**。三个可选处置（第 10 节选项 B）。

### 6.3 SDK 转录（方案 B 才有）

| 问题 | 现状 | 若启用必须做到 |
| --- | --- | --- |
| 会不会有数据行 | 现状：SDK 只能拿到**结构 + 结果引用**（`planningReadPolicy="none"`，`planner.py:38-41`），工具结果不含值。实测 30 个历史转录里 **0** 个含字面值矩阵 | 保持这条不变；任何"给 SDK 直接看数据行"的改动都必须先过隐私评审 |
| 会不会有密钥 | 密钥只在 `env`（`claude_sdk_agent.py:79-96`），不进 prompt；实测 30 个历史转录 **0** 个含 `sk-` | 转录里不得出现 `ANTHROPIC_API_KEY`/`AUTH_TOKEN`；需用例断言 |
| 会不会有工具原始输出 | `_result` 注释已声明"Never expose raw CLI stderr"；但转录会保留 `tool_use`/`tool_result` 原文 | 这是 B 的最大未知项，需 M4 逐文件确认（不是抽样） |
| 落在哪 | `CLAUDE_AGENT_SDK_STATE_DIR`（仓库外、`0700`） | 若用 `session_store`，需确认 `materialize_resume_session` 的临时目录已清理 |
| 留多久 | 无（当前关闭），但**历史 30 个转录永久留存** | 必须显式定：按天数 / 按任务数 / 会话结束即删；并补一条清理历史转录的路径 |

### 6.4 推荐默认方案

1. 规划层 checkpoint **只存白名单字段**，复用 `ledger_privacy` 的键名判定。
2. 执行层保留含值现场（功能必需），但**纳入保留期**：终态记录超过
   `operationRetentionDays` ⇒ 清 `plan_json` / `before_snapshot_json` / `checkpoint_json`
   的值部分，**保留** `operation_id` + `request_fingerprint` + `status`（防 ID 复用绕过）。
3. SDK 转录**默认仍关闭**；打开前先做 M4 夹具实验，用真实一次规划证明"转录里没有数据值"。
   另外必须补一条**历史转录清理**路径（现状 30 个旧转录永久留存，见 1.1 节探测），
   清理前逐文件确认为宜，不做"按时间一刀切"的静默删除。
4. 清理失败不阻塞启动（现状已是 `try/except + logger.exception`，`main.py:185-189`），
   但**必须**在诊断信息里可见"上次清理失败"，不能让用户以为已经清干净。
5. 损坏记录：`_encode`/`_decode` 失败按"记录不可用"处理，**不重建猜测值**；
   旧版本（`schemaVersion` 不认识）⇒ 视为不可续，提示重新发起。

---

## 7. Excel 与 WPS：共用层 / 适配层分工

### 7.1 分工原则

**共用层只放"判据"，适配层只放"怎么读、怎么写"。** 任何"两端各写一份看起来一样"的实现
都不算完成——同一件事有两个真值来源，迟早在某一端漂移。

| 关注点 | 归属 | 现状 | 缺口 |
| --- | --- | --- | --- |
| 四操作语义、"继续/核对/重授权"优先级 | 共享层 `packages/core/src/taskResume.ts` | **已有**（`taskResumeDecision` 两端共用） | 文案与"到达即执行一次"契约需核对（A1） |
| 授权集合身份（含规范形） | 共享层 `authorizedTarget.ts` | **已有**，两端 `export {}` 转发 | 无 |
| 规划检查点结构（A0） | **新增** 共享层 `taskRecoveryContext.ts` | 不存在 | 必须两端同一份，不许各写 |
| 协议字段与错误码 | `contracts.ts` ↔ `server/app/models.py` | 已有，需同步新增字段 | A0 新增 `kind`/`schemaVersion` 必须两边同时改 |
| 幂等（`operationId` 复用、`requestId` 稳定） | 判据在共享层，CAS 唯一执行点在服务端 | 服务端已有 | 客户端生成器需要把"复用旧 ID"写成不变式 |
| 宿主对象定位 / 元数据读取 | **适配层** | Excel：`Office.context.document.url` + `Excel.run`；WPS：`wpsReadWorkbook()` + 已保存路径 | WPS 端 `ActiveWorkbook` 类读数已统一改走绑定（记忆第 10 条缺陷） |
| 执行 / 回读 / 撤销 | **适配层** | Excel `excel.ts`；WPS `wpsExecutionLedger.ts` + 执行器 | 见 7.3 |
| 在途核对 | **适配层** | Excel `reconcileOfficeRangeAction`（返回 `applied/not_applied/unknown`）；WPS `wpsExecutionReconcile`（返回 `reconcile/blocked/insufficient/completed`） | **两套判定形状不同**（7.3 第 1 条） |
| 单写者租约 | 服务端 | `wpsOwnerLeaseSeconds=120`（仅 WPS 路径） | Excel 无租约，靠 `revision` + owner token |

### 7.2 已有可复用能力（两端口径一致的部分）

1. **恢复四语义的文案与优先级**：`taskResumeDecision` 的 `unavailable > reconcile >
   reauthorize > checkpoint > replan` 顺序两端一致。
2. **恢复记录白名单**：`sanitizeTaskResumeRecord` 显式枚举字段，多余字段被丢弃——
   这是"不许把整份计划塞进客户端恢复记录"的执行点。
3. **执行账本的协议面**：`/api/office/executions/*` 与 `/api/wps/executions/*` 走**同一套
   语义与同一组动作级端点**（`register / status / actions/start / actions/complete /
   suspend / resume / actions/reconcile / complete`），只在路径前缀和 `identity_prefix` 上区分宿主。
4. **身份不被文件名冒充**：Excel 用 `url:`，WPS 用 `wps-path:`；`register_office_execution`
   按完整字符串比较，跨前缀必然拒绝。
5. **审阅卡四阶段**：`ReviewOutcome` 唯一来源，两端同一份（`reviewCardPresentation.ts`）。

### 7.3 需要新增的实现 / 已证实差异

| # | 事项 | 事实 | 处理 |
| --- | --- | --- | --- |
| 1 | 在途核对的判定形状不一致 | Excel 直接给 `applied/not_applied/unknown`；WPS 先算 `localInspection`（`completed/not_started/conflict/unknown`）再给结论 | **不合并成一套**（两者输入不同），但把"什么情况必须 `insufficient`"写成共享判据表，避免一端放宽 |
| 2 | 租约只在 WPS 侧 | `_assert_wps_idle` + `wpsOwnerLeaseSeconds` 只作用于 `wps-path:` | Excel 侧靠 owner token 轮换；文档必须写清"Excel 没有跨进程租约"这个真实差异 |
| 3 | 检查点体积上限 | 两端都受 `taskRecovery.maxCheckpointBytes=64MB` 约束 | A0 新增的规划上下文要单独设更小的上限（`maxPlanningContextBytes`） |
| 4 | 规划检查点（A0） | 两端都没有 | 共享层定义 + 两端转发；服务端只认 `kind="planning_context"` |
| 5 | 宿主"现场可读性" | Excel 在未保存工作簿上无稳定身份（`url` 可能为空）⇒ 不能启用断点续跑；WPS 需已保存到磁盘 | 这是**已证实的宿主限制**，恢复入口在"未保存工作簿"上必须不出现 |

### 7.4 尚未核实的行为（不得写成已支持）

- WPS 端 `localInspection` 的 `completed/not_started` 判定在**公式、合并单元格、跨表引用**
  上的准确性——现有证据只覆盖 `writeValues` 类动作。
- Excel 端 `reconcileOfficeRangeAction` 在**格式类动作**（`setBorders`/`resizeRange`）上的
  判定：这些动作的"后态"不是简单矩阵相等，可能被判 `unknown`。
- 两端在**撤销后再恢复**的交互（撤销是否让 `office_executions` 记录失效）。
- 断网/宿主进程重启后 `WpsApplication` 对象引用是否仍有效（现有绑定只保证同一会话内）。

### 7.5 实机验收方法（两端各做，不互相代替）

| 步骤 | Excel | WPS |
| --- | --- | --- |
| 1 | 打开已保存工作簿，勾选 3 张表并确认范围 | 同（必须已保存到磁盘） |
| 2 | 发送需求，计划到达后**自动执行一次** | 同 |
| 3 | 执行到第 2 个动作时强杀任务窗格（或断网） | 同 |
| 4 | 重开面板：应显示"需要核对"，不自动写 | 同 |
| 5 | 点核对：应读真实区域给出 `applied`/`not_applied`，并补记/退回 | 同 |
| 6 | 继续：只应执行剩余动作（用写入次数断言） | 同 |
| 7 | 撤销：原地回原值，不开副本 | 同 |
| 8 | 换一个同名文件：应被身份校验拒绝 | 同 |
| 9 | 已完成任务重开：只展示，不重跑 | 同 |

## 8. 改动清单（分批，不一次做完）

### 批次 A0：让 SDK 规划任务有可续现场

> 状态：**已实施**（自动化证据齐全，真实 Excel / WPS 宿主未验收）。下面是设计时的清单；
> 实际落地与它的差异单列在 §8.0.1，**以 8.0.1 为准**。

| 文件 | 改动 | 完成条件 |
| --- | --- | --- |
| `packages/core/src/taskRecoveryContext.ts`（新增，共享层） | planning checkpoint 的**唯一**结构定义 + 白名单投影 + `schemaVersion`；两端只 `export {}` 转发 | 两端读到同一份；新增字段必须在这里登记 |
| `server/app/planner.py` | `on_checkpoint` 真正传给 `run_sdk_plan`；`resume_checkpoint` 不再一律抛错，而是识别 `kind="planning_context"` 并据此重建请求；未知 kind 仍拒绝 | 跨后端 checkpoint 仍被拒（`kind` 判别点） |
| `server/app/claude_sdk_agent.py` | `run_sdk_plan(..., on_checkpoint=None)`；在**工具边界**（每个工具调用前后）回报 checkpoint；`_result` 保留 `session_id` 到事件（**不落盘**） | 检查点只在白名单边界产生；`session_id` 不进账本 |
| `server/app/main.py` | `on_checkpoint` 已就绪，只需确认 `active_elapsed_seconds` 与 `approved_wall_time_seconds` 的语义不变 | 继续后预算不被重复追加（`planning_budget_extensions` 主键幂等） |
| `config/capabilities.json` | `taskRecovery` 新增 `maxPlanningContextBytes`；不放宽任何既有值 | 超限拒绝且错误可读 |

**完成条件（硬）**：M1/M2/M3 三条用例全绿；`npm run test:core`、`test:addin`、`test:wps`、
`build:addin`、`build:wps`、服务端 `pytest` 全绿；SDK 任务断流后 `canContinue=true`
且继续后**没有**任何写入。

#### 8.0.1 实做结果（与上面清单的差异，以本节为准）

| 设计项 | 实际落地 | 差异原因 |
| --- | --- | --- |
| 共享层 `taskRecoveryContext.ts` | `packages/core/src/planningContext.ts`（同名模块未起，结构定义放这里）+ `packages/core/src/planningContext.fixtures.json` | 结构与**夹具**分开：26 条夹具被 TS 与 Python（`test_planning_context.py`）**同时**读取，判据只有一份 |
| 两端各自"转发"契约 | 两端**不**再校验检查点；`continuation` 由服务端算好，客户端只照抄 `available/code/message/stage` | 检查点在服务端账本里，客户端看不见全貌；本地再判一遍等于写第二份判据 |
| `resume_checkpoint` 改造 | 保留"非 `planning_context` 一律拒绝"，新增 9 条契约拒绝码 + 4 条服务端闸门码 | 拒绝原因必须可区分，否则"没现场"和"现场坏了"共用一句话 |
| `config/capabilities.json:maxPlanningContextBytes` | **未加**，沿用既有 `taskRecovery.maxCheckpointBytes`（64MB） | 检查点字段是固定白名单 + 结果引用上限，体积不由外部输入撑大；单独上限留给 A2 与清理一起做 |
| 预算暂停 | 不实现；`AgentBudgetPause` 生产仍不可达 | 与本批无关，且不得把它写成"SDK 已支持预算暂停" |

**A0 新支持的阶段**：只有 `before_model`（模型调用前）与 `after_tool`（工具返回后）两个
白名单边界；**进入宿主写入阶段之后的任何中断都不进 A0**——那类情况一律回
`PLANNING_CONTEXT_HAS_EXECUTION` / `TARGET_CHANGED` / `SCOPE_CHANGED` / `SOURCE_CHANGED`，
先走执行账本与宿主核对。

#### 8.0.2 A0 补充修复：现场证据与续跑桥（2026-10-20，以本节为准）

A0 首版有两处"看起来已经做了、实际没接线"的地方。它们的共同点是**界面上完全正常**
（按钮能点、进度在动），只是什么都没真的发生 —— 这种缺口比崩溃危险得多。

**缺口一：继续时"没证据也能过"**

| 环节 | 首版缺陷 | 现在 |
| --- | --- | --- |
| 协议 | `workbookIdentity` / `authorizedSheetNames` / `sourceFingerprint` 都是可选字段 | 三项证据**必填**；`sourceVerification` 是三态字面量；自相矛盾（声明已核实却不给指纹、未核实却给了指纹）由 `model_validator` 拒绝 ⇒ 422 |
| 服务端判定 | `_reverify_continuation_target` 只在字段**非 None** 时比对，三项全省略 ⇒ 返回 `None` ⇒ 放行 | 身份为空 ⇒ `EVIDENCE_MISSING`；身份不同 ⇒ `TARGET_CHANGED`；授权成员为空 ⇒ `EVIDENCE_MISSING`；收窄/扩大都 ⇒ `SCOPE_CHANGED`；`failed` ⇒ `EVIDENCE_UNREADABLE`；`not_collected` 且原任务依赖来源 ⇒ `SOURCE_UNVERIFIED`；`verified` 与记录指纹不同 ⇒ `SOURCE_CHANGED` |
| 闸门位置 | —— | 全部发生在 `claim_planning_continuation`（领取续跑权）**之前**、模型调用**之前**；被拒的请求不产生执行登记、不改 revision |
| 客户端（Excel） | 只在 `sourceMode === "workbook"` 且本地有指纹时才比对 | 继续前先 `buildPlanningContinuationEvidence(...)`，读不到就停下，**不发请求** |
| 客户端（WPS） | 宿主读取抛错时用空对象继续 | `createWpsBridgeLiveContext` 失败即停下；三态由共享层判定，"读取失败"不退化成"尚未采集" |

**来源核实三态**（共享层唯一实现 `packages/core/src/planningContinuationEvidence.ts`）：

- `verified`：本轮**重读**了依赖来源，`sourceFingerprint` 必填并与记录比对；
- `not_collected`：`sourceRequired === false` 时才成立（原任务既无来源指纹也无结果引用），
  **且不调用** reader；
- `failed`：reader 抛错或返回空。它与 `not_collected` 是两件**必须分开到达服务端**的事。

服务端不信客户端自称的 `sourceRequired`：它自己从检查点推
（`planning_context.source_verification_required(record)` —— 有来源指纹 / 有指纹表 /
有 `resultRefs` 任一即视为依赖来源）。**合法的纯结构任务不会被堵死**：
`test_structural_task_may_continue_without_reading_source` 是它的正向锚点。

**缺口二：续跑时读表桥是聋的**

`end()` 之后 `binding` 仍然保留（结果卡要按引用取值），所以"看 `binding()?.taskId`
就跳过重建"的写法会把读表指针交给一条已经收尾的桥 —— 服务端一路等到超时。修法：

| 事项 | 落地 |
| --- | --- |
| 桥的"活没活" | `useSdkToolBridge` 分离 `liveRef`（可服务）与 `epochRef`（本地代数），新增 `ensureLive()` / `isLive()`；`begin`/`cancel`/`end` 都 `epoch+1` |
| 一次继续只初始化一次 | `ensureLive` 在 `liveRef && binding.taskId === options.taskId` 时只更新读取器、**不重建**（重建会清空 `delivered`，同一次调用会被执行两遍） |
| 迟到事件隔离 | `serviceBridgeOnce` / `runBridgeUntilIdle` 收 `isStale()`，在 poll 返回后、每个 call 前、verify 后三处提前退出；返回 `stale` 而不是写回状态 |
| 指针不丢 | WPS 抽出 `wpsProgressFromTaskEvent`（首次规划与续跑**共用**一份映射），把 `bridgeTool` / `pythonResult` 一并带出；两端 `onProgress` 都接 `handleProgressEvent` |
| 收尾 | 两端继续入口的 `finally` 都 `sdkBridge.end()`；取消/断线/失败同样走这条路 |
| 不写工作簿 | 桥只执行读表与确定性计算，产物走 `receipt` 回执；写入仍只经既有授权/执行/审阅链路 |

**仍然没做的事（不许写成"已完成"）**：真实 Excel / 真实 WPS 的续跑按钮一次都没点过；
"真实 SDK 在续跑中真的发出读表请求"没有实机证据。自动化证据只到"真实 App 组件 +
真实服务端 + 替身宿主 + 真实桥状态机"。

#### 8.0.3 桥的**生命周期竞态**（2026-10-20 第三批，以本节为准）

上一节修的是"桥没被重新立起来"。这一批修的是立起来之后**轮次之间**的三处交错 ——
共同点还是那句：界面完全正常，只是有一个请求永远不会发出，或者发错了轮次。

| # | 复现时序 | 根因 | 现在的行为 |
| --- | --- | --- | --- |
| **R1 轮次接续** | 旧轮次 poll 在飞 → `end()` → `ensureLive()` 立新桥 → 新轮次唯一一个 `bridgeTool` 事件 → 旧 poll 返回 | 新事件到达时 `runningRef` 仍被旧 drain 占着，只能记进 `queuedRef`；旧 drain 发现 stale 后**直接 return**，`finally` 只把 `runningRef` 放开，账记着却没人兑现 ⇒ 新桥是活的但 poll/execute/receipt 全为 0，必须再等一个进度事件 | stale 分支不再 return：只要此刻确有一条**活着且排过队**的桥，就回到循环顶部用**最新**的 options/binding/epoch 继续服务。串行性不变（同一个 drain），`operationId` 去重不变，没有新增轮询或重试 |
| **R2 迟到初始化** | `begin()` 进入 `await liveContext()` → `cancel()`/`end()` → 读取返回 | 初始化跨过 await 后无条件写 binding + `liveRef=true`，把已经取消的桥重新点亮；两次 `begin` 交错时，先发起、后返回的那个覆盖新轮次 | `begin` 在开头记下自己的代数，`await` 之后核验"我还是当前代吗"，不是就**一个字段都不写**并返回 `false`。`ensureLive` 的返回值改为**这一次 `begin` 的结果**，不再拿全局 `liveRef` 冒充成功 |
| **R3 迟到回执** | 复核通过 → `execute` 开始读表 → 轮次失效（取消/结束/被替换）→ 读表返回 | `serviceBridgeOnce` 在 poll 返回后、每个 call 前、verify 后都查 `isStale`，唯独 `execute` 与 `receipt` 之间没有 ⇒ 旧轮次的读表结果以"本轮结果"的身份回执进新桥 | 读表返回后、发回执前再查一次：失效则**不发回执、不推进 `delivered`、不清 pending**，整份状态由调用方按 `stale` 丢弃 |

**回执已经在途的那一半不假装能撤回**：本地取消改变不了"服务端可能已经收到"这一事实。
这一批只拦"还没发出"的回执；已经发出的如实记下 outcome、不提交本地状态，交给既有的归属、
幂等与状态核对。**没有**通过清空 `delivered`、吞错误或放松服务端校验来绕过竞态。

**判据（每处一条定点用例 + 一条反向变异，全部用可控 Promise 定序，不靠 sleep）**

| 用例 | 钉的东西 | 变异（撤回修复后必须变红） |
| --- | --- | --- |
| `旧轮次在途时到达的指针，旧轮次退出后必须被接上` | 不依赖第二个事件也能完成一次读表并回执；两次 poll 分属两轮 | R1 |
| `初始化期间 cancel / end()`、`两次初始化交错`、`ensureLive …返回 false` | 迟到初始化不作废状态、不覆盖新轮次、不冒充成功 | R2 |
| `读表返回后才发现本轮已失效：一条回执都不许发` | 读表发生、回执为 0、`delivered` 不推进 | R3 |
| `回执已在途时才失效：如实记录 outcome，但不假装本地能撤回` | 反面锚点：别把在途回执一起吞掉 | （R3 不应把它变红） |

驱动：`scripts/mutate_bridge_lifecycle_races.mjs` + `scripts/run_bridge_lifecycle_race_mutations.mjs`
（基线先绿、点名用例、还原后逐字节 sha256、复跑全绿四条判据）。

**这一批仍然没做的事**：真实 SDK / 真实 Excel / 真实 WPS 一次都没跑；三处竞态的复现与验证
全部在"真实 Hook 挂载 + 真实共享状态机 + 替身宿主端口 + 替身网络"上完成。A1 的实机部分
照样未验收。


#### 8.0.4 续跑收尾的**两处调用链缺口**（2026-10-20 第四批，以本节为准）

第三批修的是"桥立起来之后轮次之间的交错"。这一批修的是**收尾路径**上的两处：
一是旧轮次**异常**退出（不是正常返回）时队列同样没人接；二是桥初始化返回 `false`
（取消/收尾/被取代）时调用方**不看返回值**、照样发请求。

| # | 复现时序 | 根因 | 现在的行为 |
| --- | --- | --- | --- |
| **E1 异常退出的接续** | 旧轮次 poll 在飞 → `end()` + `ensureLive()` 立新桥 → 新轮次唯一指针排队 → 旧 poll **抛 AbortError / fetch reject** | `drain` 的外层 `catch` 只记账，`finally` 只放开 `runningRef`；正常返回的 stale 分支会接续队列，异常路径**绕过**了那套判断 ⇒ 新桥三项计数（轮询/读取/回执）全为 0，与"继续分析之后一动不动"完全一样 | `runBridgeUntilIdle` 的调用点包一层 try/catch 记 `roundFailed`，正常返回、失效返回、抛异常三条路走**同一个**出口（`!queuedRef → return`）。但"接续 ≠ 重试"：`roundFailed && !isStale()` 时直接 return，留待下一条真指针先问服务端，绝不擅自重放一次状态未知的读表 |
| **E2 桥没立起来仍发请求** | `App.tsx` 续跑 `await ensureLive(...)` 不看返回值就 `continuePlanningTask`；首次 `await begin(...)` 同样；`WpsAgentPanel.tsx` 的 `bridgeReady` 只用于"要不要转发进度"，`resumeWpsPlanning` 照样调用 | 桥初始化被取消/收尾/被取代时返回 `false`，请求照样发出去；服务端一发出 `bridgeTool` 指针就交给不存在的桥，模型只能等超时，而界面看起来续跑"成功"了 | 两端消费 `Promise<boolean>`：`false` 时**不发请求**，取消按取消收尾、真实失败如实提示、被取代的旧调用一字不写。同时收尾 `sdkBridge.end()` 加**归属核对**（`turnAbortRef === 本轮` / `controllerRef === 本轮`），旧调用的 `finally` 不再无条件 `end()` 关掉新轮次的桥 |

**判据（三条异常退出定点用例 + 三处调用层证据 + 三条反向变异，全部可控 Promise 定序，不靠 sleep）**

| 用例 | 钉的东西 | 变异（撤回修复后必须变红） |
| --- | --- | --- |
| `旧轮次轮询抛 AbortError / 普通网络错误 / 回执处抛错`（三条） | 旧轮次**异常**退出后，排队的新指针仍被接上；不依赖第二个事件；旧异常不写成新失败、不回执未知状态 | E1 |
| `同一代里回执失败：不自动重试、不擅自重放` | 反面锚点：接续不得退化成自动重试；下一条真指针才重新轮询 | （E1 不应把它变红） |
| Excel `首次规划 begin=false ⇒ 不发请求` / `继续 ensureLive=false ⇒ 不发请求`（`messageFirstRouting.test.tsx`） | 调用层真的消费返回值；请求计数为 0；界面如实说明 | E2 |
| WPS `ensureLive=false ⇒ 不发续跑请求`（`wpsPlanningPauseCard.test.tsx`） | 同上，WPS 侧 | E3 |

驱动：`scripts/mutate_bridge_resume_tail_gaps.mjs`（E1/E2/E3，基线先绿、点名用例、还原后逐字节 sha256、复跑全绿）。

**这一批仍然没做的事**：真实 SDK / 真实 Excel / 真实 WPS 一次都没跑；两处缺口的复现与验证
全部在"真实 Hook 挂载 + 真实面板组件 + 替身宿主 + 替身网络"上完成。A1 的实机部分照样
未验收。


### 服务端测试失败的真实原因（2026-10-20 核实，撤销上一轮的"COM 环境"归因）

`test_worksheet_restore.py` 整包偶发 `partial_or_unknown` / `partial`，上一轮曾写成
"真实 Excel COM 环境耦合"——**证据不足，撤回**。本轮用只读探针
（`outputs/probe/eb_raise_probe.py`，`sys.monitoring` RAISE + `_restore_failure`/`restore` 包装，
不改任何产品代码）在并发压力下复现并抓到真身：

- **根因**：`worksheet_checkpoint._write` 的 `temp.replace(folder/state.json)`（第 66 行）
  在 Windows 上偶发 `PermissionError(13, 拒绝访问)`，触发点在 `_record_restore_progress`
  （每放回一张表就落盘一次）或 `_execute_locked` 的多轮落盘。被 `except Exception` 吞成
  `partial_or_unknown`（删除阶段）或 `_restore_failure` 的 `partial`（撤销阶段）。
- **证据边界（本轮更正）**：夹具使用合成宿主，不是 COM 失败证据。
  探针只证明 `os.replace` 抛出 PermissionError；没有证明锁来自索引器、防病毒或某个具体进程。
  原子替换本身合理，不代表产品已正确处理失败，不能据此排除持久化韧性缺陷。
- **另一处独立噪声**：并发整包偶发 `test_python_docker_defects2.py::test_worker_payload_rejects_excessive_elements`
  抛 `MemoryError` —— 该用例 `list(range(10_000_000))` 在四进程并发时同时撑爆内存，是
  **测试资源问题**，与被测代码无关。

原“无需改产品代码”结论撤回。2026-10-20 本轮增加有限的原子替换重试，以及 Copy 前持久化
在途表名；回执失败后查询状态为待核对，不重复复制同名表。详见 ARCHITECTURE 的同日章节。
探针为只读诊断工具，`outputs/` 已在 `.gitignore`。

#### 8.0.5 首次规划桥准备与落盘故障收口（2026-10-20）

第四批的“WPS 首次请求前就检查桥”表述不准确：此前只有 `bridgeTool` 回调内的迟到检查。
本轮在意图给出 taskId 后、规划请求前，等待 `preparePlanningBridge`。
准备阶段只读身份/授权元数据；首次受控工具指针到达时同步固定来源指纹，不重建桥、不清 delivered。
此后仍按既有来源/授权复核服务工具。澄清或纯回答不因建桥提前读取数据行。
共享 begin 对空授权/不可读现场返回 false，两端停止发请求。

持久化故障注入使用合成工作簿与真实状态文件、HTTP 路由，不涉及真实 COM。
暂时替换失败只重试 IO；持续失败保留最后提交状态；恢复的在途标记在复制前落盘，
复制后无回执时重载仍可识别未知结果。重试若发现同名表已出现则停止，保留副本核对入口。
这不是自动接管未知表，也不是 SDK 原生会话恢复。A1 的真实 SDK/Excel/WPS 证据仍未取得。


### 批次 A1：执行恢复取证（实现不改或小改）

| 事项 | 完成条件 |
| --- | --- |
| 九场景的 ①③⑥ 用例补齐（现有覆盖偏 ②⑤） | 每个场景一条独立用例，**反向变异**能变红 |
| 双端实机验收：在途断开 → 核对 → 续跑 | 真实 Excel / 真实 WPS 各一次，工件与截图落盘 |
| `taskResumeDecision` 文案与"到达即执行一次"契约的一致性核对 | 若冲突，改文案并加断言（WPS 侧 `WpsAgentPanel.tsx:1057`） |

### 批次 A2：生命周期与清理

| 文件 | 改动 | 完成条件 |
| --- | --- | --- |
| `server/app/task_ledger.py` | 新增 `compact_execution_payloads(retention_days)`：终态执行记录去值 + 保留身份列；纳入 `compact_terminal_operations` 的调用链 | 清理后 `operation_id` 仍在，指纹仍能拒重放 |
| `server/app/main.py` | 启动时同时压规划与执行两侧 | 清理失败可见 |
| 文档 | 第 6.4 节落成 `docs/DEVELOPMENT.md` 的可执行步骤 | — |

### 批次 B：原生会话恢复（需用户先决定第 10 节选项 A）

先做 M4 夹具实验（不改生产），再决定下一步。

---

## 9. 验收矩阵

### 9.1 自动化 / 单元 / 集成

| ID | 判据 | 落点 | 反向变异（必须变红） |
| --- | --- | --- | --- |
| M1 | SDK 路径下 `on_checkpoint` **必须**被调用 ≥1 次 | `server/tests/`（替身 `run_sdk_plan`） | 把 `planner.py` 的透传删掉 ⇒ 计数 0 ⇒ 红 |
| M2 | 继续规划**不产生任何写入**：`executePlan` 调用计数为 0 | `apps/excel-addin/src/*test*` | 让继续路径顺手触发执行 ⇒ 红 |
| M3 | `resume_checkpoint` 的 `kind` 不认识 ⇒ 固定错误码 | 服务端单测 | 删 `kind` 判断 ⇒ 接受未知现场 ⇒ 红 |
| M4 | （夹具，不改生产）一次真实规划的转录里**不含**数据值/密钥：把 store 里会话 dump 出来做正则 + 值比对 | `scripts/probe_sdk_transcript.py`（新增，只读） | 若发现值 ⇒ 结论是"B 不可用"，不是"改判据" |
| M5 | 并发两个 `continue` ⇒ 一个 200、一个 409 | 服务端集成 | 删 `expected_revision` ⇒ 两个都过 ⇒ 红 |
| M6 | 在途动作时 `claim_resume` ⇒ 409 | 服务端集成 | 删 `in_flight IS NULL` 条件 ⇒ 红 |
| M7 | 结果引用过期 ⇒ 404 且 UI 降级为"重读重算" | 客户端 + 服务端 | 让客户端吞掉 404 ⇒ 红 |
| M8 | 已完成任务打开两次 ⇒ 只有 GET、没有 POST | 客户端 | 让 `awaitingStatus` 分支自动 POST ⇒ 红 |
| M9 | 清理后旧 `operation_id` 仍拒重放 | 服务端 | 删墓碑/身份列 ⇒ 红 |

### 9.2 故障注入

| 注入点 | 手段 | 期望 |
| --- | --- | --- |
| SSE 中途断流 | 客户端 `AbortController` / 服务端取消生成器 | 账本仍 `running` ⇒ 状态端点 `interrupted` ⇒ 可继续 |
| 进程重启 | 杀掉服务再起 | 同上；执行层在途 ⇒ `needs_reconciliation` |
| 宿主拒写 | 替身宿主在指定动作抛错 | `failed`，已写部分可撤销 |
| 响应丢失 | 替身宿主写入成功后不返回 | 恢复时核对为 `applied`，**不重复写**（断言写入次数） |
| 迟到事件 | 取消后手动追加事件 | 被序号/状态校验拒绝 |
| 换簿 | 恢复到一半切工作簿 | 客户端 `targetGuard` 拦截；服务端身份比对拦截 |

### 9.3 双端实机（不得用一端代替另一端）

| 项 | Excel | WPS |
| --- | --- | --- |
| 断流 → 继续分析（A0） | 自动化：待验 / 实机：**未执行** | 自动化：待验 / 实机：**未执行** |
| 在途断开 → 核对 → 续跑（A1） | 待验 | 待验 |
| 结果引用过期降级 | 待验 | 待验 |
| 撤销本次执行（既有） | 用户已反馈正常 | 用户已反馈正常 |
| 已完成任务重开不重跑 | 待验 | 待验 |

> 自动化一栏不给"已通过"：A0 的**端到端**证据目前只到"真实 App 组件 + 真实服务端 +
> 替身宿主"，真实 Excel / 真实 WPS 的按钮一次都没点过。替身证据 ≠ 实机证据。


---

## 10. 需要用户决定的事项（三条 + 一个附带问）

**A. SDK 会话转录是否允许落盘？**
- A1 不允许（**推荐**）：保持 `--no-session-persistence`，只做方案 A。代价是"继续分析"
  会重走推理，成本不可控。
- A2 允许，但只经 `session_store` 物化到临时目录、任务结束即删：需要先过 M4。
- A3 允许长期落盘到 `CLAUDE_AGENT_SDK_STATE_DIR`：需要明确天数、谁能读、如何一键清除。

**A 附带一问**：磁盘上已有 **30 个历史转录**（2026-10-01～10-05，约 2 MB，应用无清理路径）。
要不要在本批一并删除？三个选项：立刻删 / 逐文件检查后删 / 先留着但加清理入口。

**B. 执行账本（含写入值与写前快照）的保留期怎么办？**
- B1 与 `folder_operations` 一致：30 天后**去值**、保留身份列（**推荐**）。
- B2 更短（如 7 天）：撤销窗口变短，需在 UI 说明"超过 N 天不能再撤销"。
- B3 保持现状（永久保留）：最简单，但要接受磁盘上长期留有历史单元格内容。

**C. "继续分析"的产品措辞能不能接受？**
- C1 叫「**按当前现场重新分析**」，并说明"会重新读取来源；已完成的写入不会重做"（**推荐**）。
- C2 仍叫「继续」，但副文案必须写明"可能重新计算"。
- C3 在 A 不做的前提下不做"继续分析"，只有"重新发送需求"。

**C 已定（A0 实施时按此落地）**：按钮叫「**继续分析**」；说明只写
「将按当前授权重新读取并分析；已确认的口径会保留，写入前仍需预览与你的确认。」
**刻意不写**"已完成的写入不会重复执行"——本端**证明不了**这件事，只有查明执行账本后
（`PLANNING_CONTEXT_HAS_EXECUTION` 一类的核对结果）才允许出现。谁想在文案里加这句，
必须同时给出"从账本查到的依据"，否则就是第二个"看起来一样"的假承诺。


---

## 11. 仍未实现 / 不得宣称已交付

- **SDK 原生会话恢复未接入**：适配器仍 `--no-session-persistence`；`session_id` 未持久化；
  CLI `2.1.281` 明确"持久化关闭 ⇒ 不可 resume"。本批所有"继续"都是**应用侧重新规划**，
  不是恢复模型会话，**不得**写成"从模型断点原样续跑"。
- **A0 已实施，但只覆盖两个白名单边界**：`before_model` / `after_tool`。进入宿主写入阶段
  之后的中断**仍拒绝**（`PLANNING_CONTEXT_HAS_EXECUTION` 等 4 条闸门码），必须先走执行
  账本与宿主核对。**"有检查点"不等于"可以继续"**：`checkpoint_json` 非空只是必要条件。
- **A0 的真实宿主证据为空**：真实 Excel / 真实 WPS 上一次"断流 → 查询 → 继续分析"都没跑。
  现有证据 = 服务端集成用例（真实 HTTP 路由 + 真实账本 + sqlite 直写坏现场）、
  两端组件用例、9 条反向变异。
- **`AgentBudgetPause` 路径不可达**（旧循环已不在生产路径）——"预算暂停/继续"这套
  UI 对 SDK 任务**实际没有机会出现**；不要把它当成已验证能力。本批**未**实现预算暂停。
- **执行账本无保留期、无清理**：`plan_json` 永久含值。
- **WPS 工作目录的服务端查询/执行链路未接通**（选择与授权已通），与恢复无关但同样未交付。
- **真实 Excel / 真实 WPS 的恢复入口全部未实机验收**：本文所有判据目前只有替身宿主、
  单测与服务端集成证据。
- **九场景的两端覆盖不对等**：WPS 侧已成组
  （`wpsExecutionReconcile` / `wpsExecutionRecovery` / `wpsRecoveryInspection` /
  `wpsRecoveryPrivacy` / `wpsTaskResume`），Excel 侧只有 `officeOperationRecovery`
  与跨进程链路。两端证据**不能互相代替**；逐场景需按 9.1 的 M 编号重新对齐后才算覆盖。
  本轮 A0 只补了"断流 → 继续分析"一条，并把 A1 的 ①③⑥ 与"响应丢失、迟到结果"留在原地。
- **A0 的一条已知残余**：WPS 暂停卡的 `hasCheckpoint` 那一行（`continuation ? … : pauseType != null`）
  目前**没有**判别力用例——`blocked` 分支在它之前返回，变异它测试仍绿。已记入残余，
  不当作已覆盖。
- **转录隐私（M4）从未验证过**，所以"打开持久化也安全"这句话目前**没有任何证据**。
- **SDK 会话目录里有 30 个历史转录（2026-10-01～10-05，约 2 MB）且无清理路径**。
  当前开关有效（10-19 与 10-19 A0 的真实运行都没有新增转录），但不能据此宣称"目录是干净的"。
  本轮**未删除**这些转录（用户未批准）。
- **`test:cross-process` / `test:online` 未在本轮重新取证**：A0 改动不触碰跨进程链路与在线
  链路，但"没跑"就是没跑，不推断它们仍然绿。
