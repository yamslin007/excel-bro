# CAP-02：WPS 流式规划与检查点续跑

> 2026-10-08 补修：普通继续不传模型切换字段；手动暂停新增“核对状态 → 明确继续”，双端同步。游标仅用客户端实际收到的位置，旧历史从零补收；下文服务端 `lastSequence` 仅是当前上界，不是已收游标。详见 [实施记录](./MANUAL_PLANNING_PAUSE_2026-10-08.md)，实机尚待验收。

> 第一批（2026-10-06 上午）：第 5 条「面板刷新后恢复可恢复的规划任务」已完成。
> 第二批（2026-10-06）：续跑结果闭环、幂等身份、真实事件游标、断线恢复、授权复核、真实链路测试。
> 边界：本轮**不扩展写入动作、不做文件夹模式、不推进 CAP-03 精确动作续跑**。

## 用户的六条要求

1. 规划事件增量展示。
2. 用户可操作的规划暂停卡。
3. 网络中断后查询原任务状态。
4. 按服务端检查点继续，而不是重复发送原需求。
5. **面板刷新后恢复可恢复的规划任务。**（本批完成）
6. 普通模型后端和 Agent SDK 路径消费同一结构化结果；后端确有差异时明确说明，不能静默丢结果。

## 第 5 条：刷新后恢复，只核对不续跑

### 为什么"重开"本身不是任何结论的证据

面板重开只证明"内存丢了"。落盘历史里那份 `planningPause` 是**当时的快照**，
`revision` / 游标 / 成因都可能已经过期。直接照它画按钮，用户点下去会撞上
服务端 CAS（409），或者更糟 —— 在服务端已推进的状态上做一次基于旧现场的继续。

因此刷新路径的纪律是：**重开只核对，不自动续跑**。自动续跑会在用户完全没点过
任何东西的情况下触发一次模型调用；有副作用的事必须由用户点按钮发起。

### 三层分工

| 层 | 文件 | 职责 |
|---|---|---|
| 共享层（纯函数） | `packages/core/src/planningStream.ts` → `planningPausedEntriesFromHistory` | 从历史里挑出暂停条目：只挑 `planningPause`、同 `turnId` 取最后一条、载荷取不到不挑、空白 `turnId` 跳过 |
| 宿主适配层（查询） | `apps/wps-addin/src/wpsReadonlyAgent.ts` → `reinspectWpsPlanningPause` | 拿 `turnId` 查服务端真实状态，**只查询不继续**；载荷优先用服务端真实状态，历史快照仅供显示 |
| 界面层 | `apps/wps-addin/src/WpsAgentPanel.tsx` 挂载核对 effect + `planningPauseReload` state | 核对中不画按钮；核对通过才重建活跃暂停；读不到进度给保守说明 |

### 关键判据

- **按钮只由服务端结构化状态推导**。唯一允许产出按钮的地方是
  `wpsPlanningResumeView`（→ 共享层 `planningResumeDecision`），它直接吃
  `getWpsPlanningTaskStatus` 的返回。历史快照能推出按钮，但**不许**用它推。
- **`revision` 采信服务端；`afterSequence` 采信客户端实际接收位置**。服务端 `lastSequence` 用作上界，不能冒充已收游标；未知位置从零补收。
- **两个守卫分开写**：`canContinue` 决定"服务端是否放行"，`action` 决定
  "放行时走哪条通路"。合并成一个条件会把"放行但非精确续跑通路"误报成
  "该任务当前不能精确续跑" —— 那是替服务端下一个它没下的结论。
- **核对中不画操作区**：状态文字照给（用户知道任务还在），但按钮不出现。
  否则用户点到的是基于过期 revision 的动作。
- **读不到进度不下结论**：既不说"能继续"，也不说"任务没了"，只给保守说明。

### 与 CAP-03 的关系

两条**独立通路**，刻意不合并：

- CAP-03 处理"**写入**留下未完成现场"（`pendingWpsExecution` → `office_executions`），
  关心单元格有没有被写过；
- CAP-02 这一段处理"**规划**停在等用户决定"（历史 `planningPause` → 服务端规划账本
  `turnId`），**一个单元格都没碰过**。

合并会让"只是规划暂停"也走写入核对，反过来"写入待核对"被当成可续跑的规划 ——
两类事实混在一张卡上，用户分不清到底写没写。

## 第 6 条：双后端消费同一结构化结果

### 相同

两条后端最终都产出 `AssistantResponse` 联合类型（plan / answer / clarification），
并且**共用同一份工具契约**：`claude_sdk_agent.run_sdk_plan` 从
`excel_agent.AGENT_TOOLS_WITH_DATA` 取工具定义注册进 MCP，
`sdk_schema == shared_schema` 有逐字断言（`test_sdk_backend_shares_the_blank_policy_contract_not_a_forked_schema`）。
终局提交也走同一批 `_execute_tool` 校验，不接受裸文本或未校验计划。

### 差异（明确说明，不静默丢结果）

> **口径（第二批修正）**：下表写的是"**当前接入**未实现检查点恢复"，描述的是
> **本次接入方式**这一事实，**不是**对 Claude Agent SDK 本身能力的推断。
> SDK 是否另有会话/恢复机制属于**未核实**，文档不得写成"SDK 永远不支持检查点"。
> 换接入方式或补做现场托管时，改这张表，而不是改历史结论。

| 维度 | 普通模型后端 | Claude Agent SDK 后端 |
|---|---|---|
| 检查点 | 有：在 `excel_agent` 循环里维持 `loopState` / `messages` / `toolCache`，支持 `resume_checkpoint` | **当前接入未实现检查点恢复**：本次接入把对话状态交给 CLI 子进程托管，服务端读不到那份现场 |
| 数据行 | 支持 `planningReadPolicy=ifNeeded` | 仅表结构（`none`），带数据行请求显式拒绝 |
| 跨后端续跑 | — | **显式拒绝**，见下 |

跨后端续跑在 `server/app/planner.py` 里以 `ModelNotAvailableError` 挡住：

```python
if config.adapter == "claude_agent_sdk":
    if resume_checkpoint is not None:
        raise ModelNotAvailableError(
            "Claude Agent SDK 不能恢复另一规划后端的检查点，请重新发起预览。"
        )
```

**为什么必须显式拒绝而不是静默重新规划**：用户点的是"从上次进度继续"，
静默重新规划会给他一次从头开始的新规划 —— 可能重复调用模型与工具，界面还显示
"继续成功"。这是最坏的一种假成功。拒绝发生在**启动 CLI 之前**，
一次模型调用都不许发出去（`test_sdk_refuses_another_backends_checkpoint_instead_of_replanning`
断言 `fake.calls == []`）。

## 证据

| 套件 | 条数 | 结果 |
|---|---|---|
| `packages/core/src/planningStream.test.ts` | 30 | 全绿（含 5 条新增刷新筛选） |
| `apps/wps-addin/src/wpsReadonlyAgent.test.ts` | 35 | 全绿（含 5 条新增核对器） |
| `apps/wps-addin/src/wpsPlanningPauseCard.test.tsx` | 10 | 全绿（含 4 条新增刷新恢复） |
| `apps/wps-addin/src/wpsPlanningStream.test.ts` | 29 | 全绿 |
| `server/tests/test_claude_sdk_agent.py` | 36 | 全绿（含 1 条新增跨后端拒绝） |

**反向变异（唯一证据）**：`.tmp-cap02-reload-mutation.mjs`，9 个变异点全部
`RED(OK)` —— 拆掉任一守卫，对应用例必红；还原后复跑全绿。覆盖：面板 4 处、
共享层 3 处、核对器 2 处。

跨后端拒绝的变异：删掉 `planner.py` 里的守卫 → `test_claude_sdk_agent.py`
报 `LLMResponseError`（掉进"未提交有效受控结果"）→ 1 failed；还原后 36 passed。

## 第一批：未实现 / 边界（其中三条已在第二批修复）

- **第 1、2 条**（增量事件展示、用户可操作暂停卡）：暂停卡已能出现并能点"继续"，
  但面板内**逐步增量渲染**仍走一次性出口 —— 事件帧进了消费链，界面不逐帧刷新。
- ~~第 3 条（断线后查询原任务状态）~~ → **第二批已修**（同一会话内断线也先查状态）。
- 第 4 条（按检查点继续）已有 `resumeWpsPlanning` 与按钮链路，
  但**精确续跑剩余动作未实现**（属 CAP-03，本轮不推进）。
- ~~SDK 后端无检查点~~ → 口径已改为"**当前接入未实现检查点恢复**"，见上表。
- 未跑：真实 WPS 宿主、真实模型、付费请求、真实工作簿。

---

# 第二批：续跑补修（2026-10-06）

起因是逐项**复现**出来的五个真实缺陷，不是"补几个用例"。每条都先证明旧行为错在哪，
再改，再补真实链路测试，最后用**反向变异**证明守卫真的有判别力。

## 第 1 条：续跑结果闭环 —— 结果不许被丢掉

**复现到的缺陷**：点"继续"后服务端回的是一整个 `AssistantResponse`
（`plan` / `clarification` / `local_plan_rule` 都可能），面板却只回一句
"请重新发送需求"。用户要的东西其实算好了，界面却让他从头再说一遍。

**修法**：把 `runTurn` 里那套"结果怎么落地"的逻辑抽成**唯一一条处理链**
`applyTurnResult(result, context)`，三个入口共用：

| 入口 | 来源 | 走的是同一条链 |
|---|---|---|
| 正常提问 | `runWpsReadonlyTurn` 返回 | `applyTurnResult` |
| 断线后取回 | `inspectWpsPlanningTask().result` | `applyTurnResult` |
| 点"继续" | `resumeWpsPlanning().outcome` | `applyTurnResult` |

链内分流（与原来一致，不新造语义）：

- `plan` → 进**预览**，等用户确认；**不自动写原表**。
- `clarification`（含 `localMapping`）→ 落成**可填写的澄清卡**。
  `local_plan_rule` 走的是本地映射澄清卡，**不是** answer —— 把它当 answer
  渲染就是一次假成功（用户会以为已经算完了）。
- `answer` → 正常回答。
- 又 `planning_pause` → 落一张**新**卡，不复用旧卡。

**证据**：三条用例分别断言宿主写入为 0（`writeCountingHost()` 直接数
`Value2` / `Formula` setter），并且断言界面上出现的是预览卡 / 澄清卡，
而不是"请重新发送需求"。

## 第 2 条：幂等身份 —— 一次继续 = 一个 requestId

**复现到的缺陷**：`requestId` 在每次调用时新生成。超时重试因此变成"另一次继续"：
服务端按 `(turn_id, request_id)` 判重，换新 id 就是**每重试一次多扣一份预算、
多跑一轮规划**。

**修法**：`requestId` 在 `resumePlanningFromPause` **入口**生成并存进
`resumeRequestRef`（与 `turnId` 绑定）。整个"查状态 → 继续 → 超时重试"期间固定不变；
只有用户在这张卡上**重新点击**（上一次已结束）才换新身份。

**重复点击**由三道**互相独立**的闸共同挡住：① 按钮 `disabled={busy}`（DOM 层，
用户根本点不到）；② `if (busy) return;`；③ `resumingPauseRef`（同步 ref，不依赖
渲染时序）。这是**纵深防御**：单拆任意一道，其余仍独立生效（变异时已验证：
拆一道不红，拆三道才红）。

## 第 3 条：真实事件游标 —— 去重、身份隔离、补发都接进消费链

**复现到的缺陷**：消费链里有一个**本地自增**的 `legacySequence`，绕开了共享层的
去重闸与身份闸；续跑时 `afterSequence` 直接取 `status.lastSequence`（服务端最新序号），
**会把客户端没收到的那一段永久跳过**。

**修法**：

| 位置 | 改动 |
|---|---|
| `consumeWpsPlanningStream` | 用 `newPlanningCursor(turnId)` 取代本地计数器；每一帧 `step` / `progress` 都过 `applyPlanningEvent`，被拒（身份不符 / 陈旧重复）**就地丢弃** |
| 帧身份 | `eventTurnId` 取 `raw.operationId`，缺失才回退到本次 `turnId` |
| `continueWpsPlanning` / `streamWpsPlanning` | 新增 `afterSequence` 入口 |
| `resumeWpsPlanning` | `afterSequence: input.afterSequence ?? status.lastSequence` —— **客户端游标优先** |

## 第 4 条：断线恢复 —— 先查，再决定能不能重发

**复现到的缺陷**：catch 块只说"可以重新发送需求"，压根没查过原任务还在不在。
而重发 = 一次**全新规划**：重复调用模型、重复读取，服务端那条可能已经跑完、正在跑、
或停在可续跑的检查点上。

**修法**：新增 `reconcileInterruptedTurn(turnId, ...)`，用本轮已经用过的 `turnId`
去查真实状态，五态分治（判据与 `planningResumeDecision` 同一份，不另立一套）：

| 状态 | 动作 | 允许重发？ |
|---|---|---|
| `completed` 且有结果 | `applyTurnResult` 取回结果 | 否（结果其实算好了） |
| `running` | 明说仍在处理中 | 否（会与在跑的那一轮打架） |
| 可继续（`waiting_user` / 模型故障暂停） | 落暂停卡 | 否 |
| 有任务但不可继续 | 如实说明 | 否 |
| 服务端明确 404 / `PLANNING_TASK_NOT_FOUND` | 提示可重发 | **是**（唯一放行口） |

**查询本身失败**（后端没起）→ 保守：既不重发，也不谎称任务不存在，只说"稍后重试核对"。
区分这两件事的判据是 `error.status === 404 || error.code === "PLANNING_TASK_NOT_FOUND"`，
不是"读不到"。

## 第 5 条：授权复核 —— 续跑前重新对一遍

续跑会把"原任务的结果"落到**当前**工作簿上。等待网络响应这段时间里，用户完全可能
切了工作簿 / 改了勾选范围 —— 那等于用 A 表的意图去写 B 表。

**修法**：在"查完状态、发出继续之前"插一道 `targetGuard.check()`，与"重新预览"
走同一个判据，不另立一套。不通过就停，已停在原表之外，**0 写入**。

另一条边界：**授权复核 ≠ 澄清答案**。被拒时只往消息流里报一句事实，
**不产生**任何澄清卡 / 答案卡 —— 否则就是把"授权确认"伪装成了"用户回答了问题"。

## 第二批证据（分四类报告，不合并成一句"全绿"）

### 通过（自动化，本机）

| 套件 | 条数 | 结果 |
|---|---|---|
| `apps/wps-addin/src/wpsPlanningResumeClosure.test.tsx`（新增） | 10 | 全绿 |
| `apps/wps-addin/src/wpsPlanningStream.test.ts` | 33 | 全绿（含 4 条新增游标） |
| `apps/wps-addin/src/wpsReadonlyAgent.test.ts` | 38 | 全绿（含 3 条新增 requestId / 游标） |
| WPS 全量套件 | 34 文件 / 529 | 528 通过，1 失败（见下） |
| Excel 端全量（`npm run test:addin`） | 72 文件 / 776 | 全通过 |
| 服务端（`pytest server/tests`） | 799 | 全通过 |

**反向变异（唯一证据）**：`.tmp-cap02-resume-mutation.mjs`，
**10 个变异点全部 `RED(OK)`，0 `GREEN(NG)`** —— 基线先绿，逐条拆守卫，
每条都必须变红，拆完自动还原并复跑。

| # | 变异点 | 结果 |
|---|---|---|
| 1 | 续跑 plan 分支重新变成"请重新发送需求" | RED |
| 2 | 每次点击都生成新的 requestId | RED |
| 3 | 同时拆掉 `disabled` / `busy` / ref 三道并发闸 | RED |
| 4 | 断线不查状态、直接建议重发 | RED |
| 5 | 服务端说没有任务时也拒绝重发（反向锚点） | RED |
| 6 | 忽略客户端游标，直接用服务端最新序号 | RED |
| 7 | 消费链忽略 `afterSequence` | RED |
| 8 | 消费链不做 sequence 去重 | RED |
| 9 | 消费链不做任务身份隔离 | RED |
| 10 | 本地映射澄清当 answer 渲染 | RED |

### 失败

- `apps/wps-addin/src/wpsNativeComparison.test.ts`：1 条失败（详见上文"已知失败"），
  **与本批无关**，本批未修。这是 WPS 全量套件里**唯一**的失败项。
- 变异脚本第 3 条在"只拆一道闸"时曾是 `GREEN(NG)` —— 这不是实现缺陷，
  是**判别点选错**：三道闸互为冗余。已改为同时拆三道，现为 `RED(OK)`。
  这个中间结论**如实保留**，不粉饰成"一开始就是红的"。
- `npm run build:wps` 曾因两条 `Tuple type '[]' of length '0' has no element at index '0'`
  失败（新增用例里 `vi.fn(async () => …)` 没声明参数，`mock.calls[0][0]` 被推成空元组）。
  已给 mock 补上入参类型，现在 `tsc --noEmit` 与 `build:wps` 均通过。
  这条记在这里是因为它**真实发生过**，不是环境噪声。

### 环境阻断

- 无最终阻断。中途踩到一次：`--basetemp` 用 bash 的 `$PWD`（`/e/codex/...`）会被
  Python 解析成 `E:\e\codex\...`，导致 264 个 `FileNotFoundError`。改用 Windows
  绝对路径后 `799 passed`。这是**本机环境**问题，与代码无关。

### 实机待验（**不是**"已实现但未验证"）

- 真实 WPS 宿主上的**连点时序**（三道并发闸在真实渲染时序下是否仍各自生效）。
- 真实网络断线后的**事件补发**（`afterSequence` 补发在真实 SSE 重连下的表现）。
- 续跑过程中**切换工作簿 / 改勾选范围**的真实宿主拦截。
- SDK 路径的真实宿主验收。
- **精确续跑剩余动作**：属 CAP-03，**未实现**，不列在这里。

## 第二批仍未实现 / 边界

- 面板内**逐步增量渲染**事件帧：帧已经进了消费链（去重、身份隔离、补发都生效），
  但界面仍是一次性出口，不逐帧刷新。
- **SDK 后端当前接入未实现检查点恢复**（口径见上表），只能重新规划。
- 未跑：真实 WPS 宿主、真实模型、付费请求、真实工作簿。

## 已知失败（与本批无关，且不在本批范围内）

`apps/wps-addin/src/wpsNativeComparison.test.ts` 1 条失败：
"公式文本未变但计算显示值漂移时，两边都不能冒充仍是已确认现场"，
期望文案 `原表在并排审阅后发生变化`，实际得到
`原表或授权范围在预览后发生变化，请重新预览并确认。`。

该用例与实现文件在 git 里**均无改动**（与 HEAD `9ad8953` 一致），
属既有缺陷：显示值漂移走的是 `assertOriginalUnchanged`（`wpsNativeComparison.ts:122`）
这条更早的守卫，而用例期望的是写回循环里的那句话（`:298`）。
两处都是合法守卫，只是用例选错了期望文案。**本批未修**，避免越界改动。
