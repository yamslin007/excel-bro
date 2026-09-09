# 规划任务切换模型后从断点继续：专项修复指导

版本：v1.0  
日期：2026-09-09  
代码基线：`a8a7467 feat: add stage C checkpoint recovery`  
状态：代码与自动化测试已实施，等待真实 Excel 验收

## 1. 目标

修复以下用户可见问题：

- 规划任务中断后，用户切换模型并点击“继续”，界面看起来从第 1 步重新开始。
- 顶部模型菜单已经切换，但原任务并没有由新模型接管。
- 续跑时出现第二张执行过程卡，旧步骤与新步骤被分开显示。
- 断线后事件游标不准确，可能重复展示已收到的步骤，或错误描述为重新执行。

目标行为：

> 同一个 `turnId`、同一个 checkpoint、同一条执行时间线；用户明确选择的新模型只接管尚未完成的模型节点，已经提交的工具结果、授权范围、累计预算和判断轨迹全部保留。

本修复不扩大阶段 C 的能力范围，不实现任意 DAG、并行子任务或不具备完整 checkpoint 的强行续跑。

## 2. 已确认根因

### 2.1 顶部模型选择与续跑请求脱节

顶部模型菜单只更新全局 `selectedModelId`。规划恢复卡的普通“继续”按钮调用：

```text
resumePlanningTaskFromMessage(message)
```

只有恢复卡内单独的“使用备用模型继续”按钮才传入 `fallbackModelId`。因此用户先在顶部切换模型，再点击普通“继续”时，新选择不会进入 `/api/turn/continue/stream`。

### 2.2 `interrupted` 状态不能切换模型

后端当前只允许以下状态切换模型：

```text
status = waiting_user
error.code = MODEL_*
error.retryable = true
```

浏览器连接或 SSE 流意外断开后，持久任务通常仍为 `running`，但内存中已无活动执行者；状态查询把它解释为 `interrupted`。此时虽然存在完整 checkpoint，也能继续原任务，但模型切换校验会返回 `MODEL_SWITCH_NOT_ALLOWED`。

### 2.3 续跑无条件创建新的过程卡

`resumePlanningTaskFromMessage()` 调用普通 `beginActivity()`。该函数每次都会追加新的 `activityPending` 消息，并生成新的 `persistenceKey`。它没有复用原任务的活动消息，也没有把 `turnId` 作为初始 `operationId`。

结果是：即使后端实际使用原 `turnId`，用户仍会看到第二张从本地“正在继续”步骤开始的卡片，误以为任务已重启。

### 2.4 断线现场没有保存真实客户端事件位置

网络异常分支把 `lastSequence` 固定写为 `0`。续跑时使用消息里的 `lastSequence` 作为 `afterSequence`，所以服务端会从头补发事件。

从头补发本身不应导致业务重跑，但在“新建空过程卡”的前提下，补发的旧事件会被展示成新任务从第 1 步执行。

注意：不能简单把 `afterSequence` 改为服务端最新的 `taskStatus.lastSequence`。服务端序号是当前上限，客户端断线前可能漏收其中一部分；盲目从服务端最新序号开始会丢失事件。正确做法是持久保存“客户端最后实际收到的序号”，并用服务端最新序号校验其合法性。

### 2.5 回归测试覆盖了 API，没有覆盖真实 UI 接线

现有后端测试验证了 `model_error` 后使用备用模型继续，也验证了 checkpoint 被传回 Agent。但没有覆盖：

- `interrupted` 状态切换模型；
- 顶部当前模型与恢复按钮的联动；
- 原活动卡复用；
- 断线后按客户端事件游标补收；
- 页面刷新后同一卡片恢复。

## 3. 必须保持的正确性边界

实施时必须同时满足：

1. 续跑只能使用原 `turnId`，不能调用 `/api/turn/stream` 创建新任务。
2. 服务端 checkpoint 是执行现场的权威来源；前端消息只负责发现和展示。
3. 模型只能在完整模型调用边界切换，不恢复半份模型响应。
4. 用户必须明确选择目标模型；不能后台静默改用另一个连接。
5. 原工作表授权、来源指纹和图片能力要求必须重新核对，不能扩大数据范围。
6. 已完成工具调用不得重做；累计 token、轮次、错误次数和有效运行时间不得清零。
7. 活跃任务不能被第二个继续请求接管。
8. 重复点击必须依靠稳定 `requestId` 和 revision 幂等处理。
9. 进度补发只重建界面，不代表重新执行业务节点。
10. 不把提示词、原始数据行、API Key 或隐藏思维写入进度事件和诊断信息。

## 4. 目标状态机

| 持久状态 | 内存活动执行者 | checkpoint | 允许继续当前模型 | 允许显式切换模型 | 处理方式 |
| --- | --- | --- | --- | --- | --- |
| `running` | 有 | 任意 | 否 | 否 | 返回任务仍在运行，前端只重连进度 |
| `running` | 无 | 有 | 是 | 是 | 对外解释为 `interrupted`，CAS 领取原任务 |
| `running` | 无 | 无 | 否 | 否 | 标记需处理或失败，不能伪造续跑 |
| `waiting_user` + budget | 无 | 有 | 是 | 否 | 用户追加时间后继续当前模型 |
| `waiting_user` + retryable `MODEL_*` | 无 | 有 | 是 | 是 | 重试当前模型或显式选择备用连接 |
| `completed` | 无 | 任意 | 不需要 | 不需要 | 重放原结果，不再调用模型 |
| `failed` / `cancelled` | 无 | 任意 | 否 | 否 | 显示终态，不创建新任务 |

允许 `interrupted` 切换模型的前提必须全部成立：

- 原记录仍为 `running`；
- `turnId` 不在 `_active_planning_tasks`；
- 存在完整 checkpoint；
- 目标连接已配置且不是当前连接；
- 图片任务的目标连接支持图片；
- revision、requestId 和硬预算检查通过。

## 5. 目标交互

### 5.1 恢复卡按钮语义

恢复卡应根据当前选择显示明确动作：

- 当前选择与原模型相同：`继续使用「模型 A」`
- 当前选择是允许的备用连接：`使用「模型 B」从断点继续`
- 原任务仍在后台运行：`重新连接进度`
- 只需找回已完成结果：`查看已完成结果`
- 没有安全 checkpoint：按钮禁用，并说明不能续跑的原因

用户在顶部模型菜单完成选择后，再点击带模型名称的继续按钮，可视为一次明确模型选择。请求中仍必须携带：

```json
{
  "modelId": "connection:...",
  "confirmModelSwitch": true
}
```

不能仅根据服务端推荐顺序自动切换。

恢复卡内可保留备用模型下拉框，但它与顶部模型菜单必须使用同一个选中状态或明确显示“仅用于本次续跑”，避免出现两个相互矛盾的模型选择。

### 5.2 时间线展示

正确示例：

```text
执行过程 · 8 步 · 已恢复
✓ 1 已识别工作表字段结构
✓ 2 已理解需求与授权范围
...
! 8 生成操作方案时连接中断
↻ 已切换到模型 B，从保存的规划现场继续
● 8 正在生成操作方案
```

禁止展示：

```text
旧卡：执行过程 · 8 步 · 连接中断
新卡：执行过程 · 2 步 · 正在执行
```

“已切换模型”是恢复分段标记，不应被计为重新完成前 7 个业务步骤。

## 6. 前端改造方案

### 6.1 保存原活动消息身份

在本地 `ChatMessage.planningResume` 中增加：

```ts
activityMessageId?: string;
clientLastSequence?: number;
```

这属于前端对话持久化字段，不是服务端执行协议。旧历史没有这些字段时，按以下顺序寻找原活动消息：

1. 使用 `activityMessageId`；
2. 查找 `activityLog.operationId === turnId` 的消息；
3. 都找不到时才创建兼容占位，并明确标记为“已找回原任务”，不能伪装成新任务。

初次启动规划时就应把 `turnId` 传给活动 Hook 作为 `operationId`，不要等第一条 SSE 事件到达后再绑定。

### 6.2 新增“恢复活动”能力

在 `useActivityProgress` 增加独立的恢复入口，例如：

```text
resumeActivity(previousLog, options)
```

职责：

- 复用原 `persistenceKey`；
- 复用原 `operationId/turnId`；
- 载入原 `steps` 和 `traces`；
- 保存客户端最后收到的 `sequence`；
- 累加原 `totalMs`，不把时间重新从 0 计算；
- 增加一条受控的“正在恢复/已切换模型”分段标记；
- 后续相同 `stepId` 原位更新，不追加重复步骤。

实时渲染时应二选一：

- `activityPending` 且当前活动匹配消息：显示实时 Timeline；
- 否则显示 `activityLog`。

不能同时渲染实时 Timeline 和同一消息的历史 `activityLog`。

### 6.3 保存真实客户端事件游标

`useActivityProgress` 应暴露或持久保存最后成功归并的 `sequence`。出现断线时：

```text
planningResume.clientLastSequence = activity.sequence
```

恢复前读取服务端状态后计算：

```text
afterSequence = clamp(clientLastSequence, 0, taskStatus.lastSequence)
```

规则：

- 客户端序号小于服务端序号：补发缺失事件；
- 两者相同：只接收新事件；
- 客户端序号异常大于服务端序号：降级到安全值并记录脱敏诊断；
- 旧记录没有客户端序号：允许从 0 补发，但必须归并到旧步骤，不能生成第二张卡。

### 6.4 续跑模型选择

`resumePlanningTaskFromMessage()` 应先查询任务状态，再确定本次模型：

```text
requested model = 恢复卡本次显式选择
               ?? 顶部当前选择（与原模型不同且在服务端允许列表中）
               ?? 原任务当前模型
```

如果顶部选择与原模型不同，但不在服务端返回的允许列表中，应停在恢复卡并说明原因，不得回退为新任务，也不得静默重试原模型。

继续请求固定使用状态查询返回的：

- `turnId`
- `revision`
- 可切换连接列表

事件游标使用客户端已接收位置，并受服务端 `lastSequence` 上限约束。

### 6.5 禁止回退到新任务入口

恢复路径发生以下错误时，只更新原恢复卡：

- 状态查询失败；
- revision 冲突；
- 模型不允许切换；
- 来源指纹变化；
- checkpoint 不存在；
- 达到预算上限。

任何恢复错误都不得调用普通 `sendMessage()`、`analyzeConfirmedIntent()` 或 `/api/turn/stream` 重新提交原需求。

## 7. 后端改造方案

### 7.1 统一判断“是否允许模型切换”

把分散在 `_planning_pause_payload()` 和 `/api/turn/continue/stream` 中的判断收敛为一个确定性函数，输入至少包括：

- 持久记录；
- 当前 `turnId` 是否有活动执行者；
- 是否存在 checkpoint；
- 任务是否需要图片能力；
- 当前已配置连接。

输出：

- `can_switch_model`；
- 允许的连接摘要；
- 不允许时的稳定原因码。

状态查询和继续接口必须调用同一个判断函数，避免界面显示“可以切换”而执行接口又拒绝。

### 7.2 允许安全的 interrupted 接管

`/api/turn/continue/stream` 的模型切换条件扩展为：

```text
retryable model_error
OR
persisted running + no active executor + complete checkpoint
```

这不是允许任意运行中任务抢占。接口进入时仍必须先确认 `_active_planning_tasks` 中没有原任务，并通过账本 revision CAS 领取。

### 7.3 收紧领取与活动登记的竞态窗口

当前“检查活动集合”和“账本 CAS 领取”之间存在时间窗口。实施时应确保：

1. 同一进程内，活动检查、CAS 领取和活动登记形成一个受控领取过程；
2. SQLite revision 保持跨请求最终仲裁；
3. 相同 `requestId` 重试只能复用相同目标模型；
4. 不同请求同时继续时，最多一个进入模型循环；
5. 失败领取不能写入模型切换记录。

不要只依赖 React 的 `busy` 或进程内集合保证唯一执行者。

### 7.4 模型切换事件

领取成功后追加一条单调事件：

```text
stepId: planning.model.switch.<count>
status: completed
title: 已切换到用户选择的模型
summary: 保留原任务范围、已完成工具结果和累计预算，从完整检查点继续。
```

事件只能记录连接 ID、模型展示名等受控摘要，不能记录 Base URL、Key、完整请求或 checkpoint。

随后使用：

```text
PlanRequest.modelId = claimed.current_model_id
resume_checkpoint = claimed.checkpoint
```

继续 Agent 循环，不重新调用意图识别，也不重新构建工作表授权。

### 7.5 断线后的状态语义

SSE 客户端断开后：

- 先取消仍在途的模型请求；
- 等待子任务真正停止后再释放活动执行者标记；
- 保留最近一次完整 checkpoint；
- 状态查询把“持久 running、内存无活动执行者”解释为 `interrupted`；
- 不把半份模型响应保存成可恢复结果。

## 8. 协议与兼容性

首期可以复用现有服务端协议字段：

- `pauseType: budget | model_error | interrupted`
- `currentModelId`
- `canSwitchModel`
- `availableModelConnections`
- `modelId`
- `confirmModelSwitch`

因此原则上不需要增加新的前后端 API 字段。若实施中改变这些字段含义，必须同步：

- `apps/excel-addin/src/contracts.ts`
- `server/app/models.py`
- API 解析测试和后端模型测试

`activityMessageId` 与 `clientLastSequence` 只进入前端本地 `ChatMessage`，旧对话应保持兼容。

## 9. 分步实施顺序

### 第一步：修复后端状态机

- 抽取统一切换资格判断；
- 允许安全 `interrupted` 状态显式切换；
- 保持 active fence、revision、requestId 和预算约束；
- 增加后端单元测试。

退出条件：同一 `turnId` 从 interrupted checkpoint 由备用模型完成，账本只有一条任务记录并产生一条模型切换记录。

### 第二步：修复前端请求接线

- 顶部模型选择与本次续跑目标建立明确映射；
- 按钮显示具体模型名称；
- 续跑只调用 `/api/turn/continue/stream`；
- 保存真实客户端事件位置。

退出条件：切换顶部模型后点击继续，请求携带原 `turnId` 和明确的 `modelId`，不出现新的 `/api/turn/stream`。

### 第三步：复用原时间线

- 保存/寻找 `activityMessageId`；
- 给 Hook 增加恢复旧日志能力；
- 合并补发事件；
- 避免实时与历史 Timeline 重复渲染。

退出条件：UI 只有一张过程卡，已完成步骤不消失、不重复，恢复步骤在原位置继续。

### 第四步：真实故障验收

- Vite 页面刷新；
- 本地服务重启；
- 模型连接超时；
- SSE 网络中断；
- 顶部切换模型后继续；
- 连续点击继续；
- 工作簿数据变化；
- 图片任务切换到不支持图片的模型。

## 10. 自动化测试矩阵

### 10.1 后端

1. `model_error` + checkpoint：重试当前模型。
2. `model_error` + checkpoint：切换允许的备用模型。
3. `interrupted` + checkpoint + 无活动执行者：切换备用模型成功。
4. `interrupted` + 有活动执行者：返回 `PLANNING_TASK_STILL_RUNNING`。
5. `interrupted` + 无 checkpoint：拒绝继续。
6. 目标模型与当前模型相同：按继续当前模型处理，不新增切换记录。
7. 目标模型未配置或能力不匹配：拒绝，原记录不变化。
8. 重复相同 requestId：不重复增加预算、不重复记录切换。
9. 相同 requestId 改用另一模型：拒绝。
10. 两个不同 requestId 同 revision 并发：只有一个领取成功。
11. 续跑后 `resume_checkpoint` 含原工具缓存，已完成工具不再调用。
12. 切换模型后累计预算和 LoopState 不清零。

### 10.2 前端

1. 普通暂停继续复用原活动消息。
2. 网络断线保存最后实际收到的 sequence。
3. 从 0 补发旧事件时按 stepId 更新，不生成重复步骤。
4. 顶部选择允许的备用模型后，继续请求携带该模型和确认字段。
5. 顶部选择不允许的模型时，不发送继续请求并显示原因。
6. 服务端返回 completed 时只重放结果，不启动新活动。
7. revision 冲突后刷新状态，不创建新 turn。
8. activityPending 时不同时渲染历史 Timeline。
9. 旧消息没有 activityMessageId 时能按 operationId 找回。
10. 页面刷新后仍使用原 persistenceKey。

### 10.3 网络请求断言

针对一次恢复操作记录请求序列，必须是：

```text
GET  /api/turn/{原 turnId}/status
POST /api/turn/continue/stream
```

期间不得出现：

```text
POST /api/turn/stream
POST /api/intent
```

## 11. 真实 Excel 验收脚本

1. 打开一份隔离测试工作簿并选择一张工作表。
2. 提交一个至少需要一次字段读取和一次模型规划的任务。
3. 等时间线已完成多个步骤后，让模型连接中断。
4. 确认界面保留原步骤，并显示可以继续。
5. 在顶部切换到另一个已配置、能力匹配的模型。
6. 确认按钮文字变为“使用「新模型」从断点继续”。
7. 点击继续。
8. 确认没有出现第二张从第 1 步开始的过程卡。
9. 确认时间线出现一次模型切换标记，并从原未完成步骤继续。
10. 确认后端账本只有原 `turnId`，`modelSwitchCount` 增加 1。
11. 确认已完成的读取工具调用次数没有增加。
12. 若最终生成写入计划，仍必须预览并由用户确认，不能因恢复自动写入 Excel。

## 12. 验收标准

全部满足才算修复完成：

- 切换模型后仍使用原 `turnId` 和原 checkpoint。
- 原任务记录的 `currentModelId` 更新，模型切换历史增加一次。
- 不产生语义相同的新规划任务。
- 已完成工具结果不重复计算。
- UI 始终只有一张该任务的执行过程卡。
- 步骤序号单调、旧步骤不重复、缺失事件能补回。
- 页面刷新和 SSE 断线后行为一致。
- 活跃任务、无 checkpoint、来源变化、能力不匹配和预算耗尽均安全拒绝。
- 不扩大工作表授权，不泄露恢复现场或模型密钥。
- 自动化测试、前端构建和 Python 编译全部通过。

## 13. 需要修改的主要文件

| 文件 | 主要职责 |
| --- | --- |
| `apps/excel-addin/src/App.tsx` | 续跑入口、模型选择、原活动消息挂接、错误降级 |
| `apps/excel-addin/src/types/chat.ts` | 本地恢复消息增加活动身份和客户端事件游标 |
| `apps/excel-addin/src/hooks/useActivityProgress.ts` | 从历史日志恢复同一时间线并归并补发事件 |
| `apps/excel-addin/src/hooks/useActivityProgress.test.ts` | 时间线恢复、去重和 sequence 测试 |
| `apps/excel-addin/src/api.test.ts` | 继续请求中的 turnId、afterSequence、modelId 和确认字段 |
| `server/app/main.py` | interrupted 切换资格、领取流程和切换事件 |
| `server/tests/test_main.py` | 状态机、模型切换、并发与 checkpoint 续跑测试 |
| `docs/ARCHITECTURE.md` | 完成后同步规划恢复和同一时间线语义 |
| `docs/DEVELOPMENT.md` | 增加真实故障调试与验收说明 |

## 14. 提交前检查

```powershell
npm run test:addin
npm run build:addin
python -m pytest server/tests -q --basetemp .pytest-tmp-planning-resume
python -m compileall -q server/app
```

测试临时目录应使用新的、隔离的路径，避免 Windows 上旧 pytest basetemp 被占用影响结果。

## 15. 实施记录

2026-09-09 已完成：

- 顶部所选模型与恢复请求联动，按钮明确显示接管模型；
- `interrupted` 且具有完整 checkpoint、无活动执行器时允许显式模型切换；
- 账本 CAS 领取和进程内活动登记使用同一互斥区间；
- 续跑复用原活动消息和 `turnId`，不再新建空过程卡；
- 本地活动日志持久保存客户端最后归并的 sequence，并据此补收事件；
- 增加中断换模型和同时间线事件去重的自动化回归测试。

尚需按第 11 章在真实 Excel 中完成断线、顶部换模型、刷新后继续和单卡时间线验收。
