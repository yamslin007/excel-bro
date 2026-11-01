# A1-E3：来源变化后的终止 + 双端缺口矩阵（2026-10-25）

> 范围：A1-E3 这一批。**没有**启动 A2、**没有**开启 SDK 原生会话持久化、**没有**提交 Git。
> 本文只写有证据的事；"未做"一律写"未做"，不写"仅待实机"。

---

## 0. 结论摘要

1. **修掉一个真实缺陷**：`source_changed` 被判出来之后，本轮规划**没有停**——桥继续把后面的
   工具调用派给宿主，SDK 循环继续跑，模型一次次重读、一次次被拒，直到
   `MODEL_TURN_LIMIT`。现在它是**本次规划的终态**，且终止是代码级的（不靠提示词）。
2. **新增 20 条自动化用例**：服务端 16（停止与守卫 14 + 并发认领 2），前端 4（Excel 面板 1、SSE 帧 1、
   Excel 授权改名 1、WPS 面板 1）。
3. **9 个单点变异全部变红并逐字节还原**（S1–S4 / D1 / W1–W2 / C1–C2）。
4. **双端缺口矩阵 12 行**（§3）。Excel 侧发现 1 处真实缺口（未保存同名簿的**读路径**身份回退，写路径安全），
   其余为"已实现 + 缺专测"或"原生待验"。
5. **原生证据这一批一行都没有**：真实 Excel / 真实 WPS 窗格里一次都没点过"继续分析"、
   没看过一次来源变化停止。A1-GATE **不通过**。

---

## 1. 缺陷：`source_changed` 之后的无效重试

### 1.1 形状（真实日志）

第七轮 B 路的第 1 轮**真实 SDK 实验**（2026-10-25）：同一轮里 6 次受控读表，**全部**被 `source_changed` 拒绝，紧接着
`MODEL_TURN_LIMIT` 收尾。用户看到的是"任务失败"，而真正发生的是"前提早就没了，但通路没关"。

### 1.2 层：不是模型，不是提示词，是桥 + SDK 工具循环

| 位置 | 旧行为 | 后果 |
| --- | --- | --- |
| `task_bridge.ToolBridgeSession.deliver()` | 把拒绝回执 `set_result` 交回模型，但**session 仍可用** | 下一次 `request()` 照样 `on_request` → **真的又派给宿主读表** |
| `claude_sdk_agent.run_sdk_plan.handle()` | 只在 `receipt["status"] == "ok"` 时记账，`source_changed` 仅记一条 warning | 模型收不到"停"，继续按自己的理解重试 |
| `_result()` 的流循环 | 不检查任何停止信号 | 即使桥停了，也要等模型自己不再发调用 |

因此"在提示词里写不要重试"从来不是解：真正的通路没关。

### 1.3 修法：一处判据 + 两段式停止

- **唯一判据/唯一文案**：`server/app/planning_source_guard.py`
  （`SOURCE_CHANGED_RECEIPT_CODE`、`SOURCE_CHANGED_TURN_CODE`、`SOURCE_CHANGED_STOP_MESSAGE`、
  `is_source_changed_receipt`、`SourceChangedDuringPlanning`）。桥、SDK 循环、HTTP 错误映射、
  两端面板全部读这一份，不各写一句。
- **桥的两段式顺序**（`deliver()` 尾部）：先 `pending.future.set_result(receipt)` 把
  `source_changed` 交回模型，**再** `stop_for_source_change()`。顺序反了，模型收到的是
  `cancelled`，看不出真正原因。
- **停止态回同一固定码**：`request()` 在 `cancel_reason == source_changed` 时仍回
  `source_changed`（不是 `cancelled`），且**不再 `on_request`** —— 不产生新的宿主读表。
- **连带作废**：`TaskToolBridge.stop_for_source_change(task_id)` = session 停止 +
  `drop_python_state`（数据集登记、协调器绑定、已登记结果与归属 claim 一起作废）。
- **流式循环真停**：`_result(..., stop_event=...)` 在处理**每条消息前**检查并 `break`，
  `finally` 里 `aclose()` 关流，然后抛 `SourceChangedDuringPlanning`。
- **工具入口双闸**：`handle()` 开头 `if source_stop.is_set(): return is_error + STOP_MESSAGE`
  （连 `describe_tables` 也拒）。
- **对外是一个明确终态**：`_map_turn_error` → `(409, "SOURCE_CHANGED", msg, retryable=False)`；
  `/api/turn` 与 `/api/turn/stream` 的 except 元组都收这个类型。

### 1.4 明确**没有**做的三件事（各有正向锚点用例）

| 没做 | 为什么 | 锚点 |
| --- | --- | --- |
| 把所有工具错误都变成终态 | 会废掉 A1-E2 已验证的"重算一次就好" | `test_retryable_timeout_does_not_stop_planning`、`test_successful_reread_does_not_stop_planning`（并断言 `host_polls == 2/3`，即两次/三次都真到了宿主） |
| 放宽来源核对 / 自动接受旧结果 | 那是把缺陷当功能 | 既有 A1-E2 A/B 两路 + `wpsSdkBridgeHost.test.ts:391` |
| 自动重新规划写入 / 加"执行前确认"按钮 | 加确认不是终止，且会让写入路径多一个前置条件 | `test_new_turn_after_stop_reads_the_current_source_again`（用户**明确**重新发起才重新读算） |

### 1.5 分档

- ② **双端在线集成：通过**（真实服务端 + 真实面板组件 + 替身宿主 + 真实 SSE 帧）。
- ④ **真实模型规划：本批未重跑**（A1-E2 那一轮的模型证据不能挪用到本批的停止行为）。
- ⑤/⑥ **真实 Excel / 真实 WPS：未做**。

---

## 2. 本轮证据清单

### 2.0 新增的开发入口（默认关闭）

`POST /api/dev/source-changed-fault/arm`（+ `/cancel`、`GET`），开关 `EXCEL_BRO_SOURCE_CHANGED_FAULT=1`：
让**被绑定的那一次任务**的下一次工具调用直接得到生产那份 `source_changed` 并走生产停止链路（不派发给宿主）。
存在的唯一理由：`source_changed` 的两个真实产生点都要求「模型在来源变化后再读一次」，
真实宿主里无法确定性触发 —— 靠运气不叫验收。开关关闭时端点是 404、桥的行为与原来逐字节一致
（`test_dev_fault_is_off_by_default` + 变异 D1 钉住）。

**2026-10-25 第二批：作用域收紧（此前是真缺陷）**

原来的启用是"下一次工具调用，不论哪本簿"，实测可复现：

```
arm({}) -> armed hostBound=False workbookBound=False
另一本簿 -> source_changed dispatched=0   # 被抢走
stolen = True      phase after other = consumed
```

（探针 `outputs/verify-a1e3-id/fault_scope_probe.py`，前后输出在同目录
`fault-scope-before.txt` / `fault-scope-after.txt`。）

被抢的后果不是"多注入一次"这么轻：**被抢的那一轮真的停在半路**，而要验的那本簿
一次都看不到停止。

现在的契约（三个变化，缺一不可）：

| 变化 | 以前 | 现在 | 为什么 |
| --- | --- | --- | --- |
| 启用时的绑定 | 可留空（"下一次调用"） | `hostId` + `workbookName` **必填**，`workbookIdentity` 可选 | 留空等于把归属交给"谁先说话" |
| 任务归属 | 无（按调用命中） | 服务端在**开桥那一刻**绑定（`TaskToolBridge.open` → `bind_task`），此后只认那个 `taskId` | 那一刻 taskId / 宿主 / 工作簿都已确定，用户与模型都无从指定 |
| 状态 | `off/idle/armed/consumed/expired` | 增加 `bound`，并多一个 `taskBound` 字段 | 能看出"已绑给某一次任务"，还是"还在等人来绑" |

单次 + 120 秒过期 + 可取消这三条没变；开关关闭时仍是 404、`take()` 恒 False。

**注入能证明什么、不能证明什么**：它证明"守卫一旦给出 `source_changed`，这条响应链真的停、
旧入口真的消失、宿主写入为 0"；它**不能**代替"真实来源变化被检测出来"（那要求模型在改数
之后再读一次，见原生验收包 §4）。

### 2.1 新增

| 编号 | 用例 | 文件 | 关键断言 |
| --- | --- | --- | --- |
| S1 | `test_bridge_stops_dispatching_after_source_changed` | `server/tests/test_source_changed_stop.py` | 桥停止后 `source_stopped is True`，且宿主**没有**再被要求读表 |
| S2 | `test_late_host_receipt_after_source_stop_is_refused` | 同上 | 停止后重投迟到回执 → `BridgeError("cancelled")`，不进账本 |
| S3 | `test_registered_result_is_released_after_source_stop` | 同上 | `local_results.resolve(...) is None`（归属已释放） |
| S4 | `test_sdk_planning_stops_on_first_source_changed` | 同上 | 真流式逐条 yield，第 1 次 `source_changed` 即停，`len(tool_results) == 1` |
| S5 | `test_in_flight_tool_calls_after_the_stop_are_refused_locally` | 同上 | 停止后同批再发 2 条调用：本地拒（3 条结果，后 2 条 `is_error`） |
| S6 | `test_sdk_turn_reports_stopped_step_and_never_claims_success` | 同上 | 有 `planning.source_changed / failed` 步骤；不出现成功模板 |
| S7 | `test_retryable_timeout_does_not_stop_planning` | 同上 | 无停止事件，`host_polls == 2` |
| S8 | `test_successful_reread_does_not_stop_planning` | 同上 | 无停止事件，`host_polls == 3` |
| S9 | `test_turn_stream_reports_source_changed_as_stopped_not_limit` | 同上 | SSE 帧码是 `SOURCE_CHANGED`，**不是** `MODEL_TURN_LIMIT` |
| S10 | `test_another_paused_turn_keeps_its_continue_entry` | 同上 | 另一条独立暂停任务仍有 `canContinue` |
| S11 | `test_new_turn_after_stop_reads_the_current_source_again` | 同上 | 用户重新发起 → 按当前数据重读重算，正常返回 |
| C1 | `test_only_one_concurrent_office_resume_claim_wins` | `server/tests/test_recovery_claim_concurrency.py` | 8 线程同 revision：成功数**恰好 1**；旧 `owner_token` 开动作被拒；赢家只能开一次 |
| C2 | `test_only_one_concurrent_planning_continuation_wins` | 同上 | 8 线程不同 `request_id`：成功数恰好 1；扩展记录只有 1 条 |
| E1 | Excel 面板：来源变化后的停止 | `apps/excel-addin/src/sdkIntentRouting.test.tsx` | 原文照转、活动卡 `data-status="failed"`、无"继续"、无 `.inline-plan`、`executeQueryTableTool` 未调用；重发正常完成 |
| E2 | SSE 帧原样（409 / `SOURCE_CHANGED` / retryable=false） | `apps/excel-addin/src/api.test.ts` | 真 SSE 帧断言四字段 |
| E3 | 已授权表改名（张数与顺序不变，只有名字变） | `apps/excel-addin/src/sdkBridgeHost.test.ts` | `unauthorized_scope`，`execute` 未调用 |
| D1 | `test_dev_fault_is_off_by_default` | `server/tests/test_source_changed_stop.py` | 开关关闭时 `phase == "off"`、`take()` 恒 False、`arm()` 抛错；桥的行为与生产一致 |
| D2 | `test_dev_fault_turns_the_next_call_into_a_real_stop` | 同上 | 开启后下一次调用直接得到 `source_changed`、`dispatched == []`（没派给宿主）、单次（第二个任务不受影响） |
| G1 | `test_dev_fault_status_is_readable_and_writes_still_need_json` | `server/tests/test_source_changed_stop.py` | GET 状态 200（真 HTTP）；写方法缺 JSON 声明仍 403 |
| W1 | WPS 面板：来源变化后的停止 | `apps/wps-addin/src/WpsAgentPanel.test.tsx` | `STOP_MESSAGE` 原文、`data-status="failed"`、无"继续分析"、无暂停卡、`host.writes() === 0`、重发正常完成 |

### 2.2 复用（未重跑昂贵实验，只在新用例里引用其结论）

| 场景 | 已有证据 |
| --- | --- |
| 换簿 / 同名簿（服务端） | `test_agent_tool_bridge_lifecycle.py:416 / :437` |
| 同轮重新授权、换簿后旧结果作废 | 同文件 `:504 / :514 / :529` |
| A1-E2 来源变化 A 路（算完才变） | `WpsAgentPanel.test.tsx:6788`、Excel `excelWriteEndToEnd.test.tsx` |
| A1-E2 来源变化 B 路（模型确实再读） | `WpsAgentPanel.test.tsx` 中"A1-E2 B 路（WPS）"两条 |
| 中断 → 重开 → 只续剩余动作 → 审阅 | `apps/excel-addin/src/excelWriteEndToEnd.test.tsx:620`（双击"核对并继续"、`startOfficeAction` 索引 `[0,1,2]`、续跑只写 `B4`、`从第 3 步继续并完成`） |
| WPS 只续第三个动作 | `wpsExecutionResume.test.ts:38 / :106` |
| 回执丢失只补记不重写 | `wpsExecutionReconcile.test.ts:75 / :109 / :158 / :172`、`test_task_ledger.py:592` |
| 审阅标记归属（重开只清自己的） | `wpsResultReviewMarkers.test.ts:143 / :164 / :179 / :459 / :488 / :515 / :360` |

---

## 3. 双端缺口矩阵

状态标记：**✅ 已实现且自动化通过**（均为替身宿主级）｜**⚠️ 已实现、缺端到端专测**｜
**❌ 未实现**｜**⏳ 原生待验**（真实宿主里一次都没跑过）｜**🚫 宿主限制**。

> 全表共同前提：`✅` 一律只到"真实服务端 + 真实面板组件 + 替身宿主"。真实 Excel / WPS
> 里的同一场景都是 `⏳`。

| # | 场景 | Excel | WPS | 尚缺 |
| --- | --- | --- | --- | --- |
| 1 | **换工作簿**（不同路径） | 生产入口 `sdkBridgeHost.verify` → `verifyBridgeAuthorization`。`sdkBridgeHost.test.ts:122`「同名但不同工作簿必须拒绝」✅ | `wpsSdkBridgeHost.test.ts:391`「换簿（同名不同路径）：拒读，错误码 `source_changed`」✅ | 两端 ⏳：真实窗格里真的开两个簿再切 |
| 2 | **同名工作簿**（同名的另一个簿 / 未保存） | **2026-10-25 第二批已修**：新增 `currentOfficeBridgeWorkbookIdentity()` —— 已保存用 `url:`，未保存按**宿主的工作簿对象**发一个会话内令牌（`unsaved:<n>-<rand>`，WeakMap 记在 document 对象上）。规划请求、桥绑定、执行前复核、续跑证据四处读同一份，两个同名未保存簿不再三项全等。回归 `apps/excel-addin/src/bridgeWorkbookIdentity.test.ts`（红→绿，见 §4 的 M1）✅ | `wpsSdkBridgeHost.test.ts:203`「未保存的工作簿用会话内对象身份，绝不退回会撞车的 `name:`」+ 本轮新增 `wpsSdkBridgeHost.test.ts`「窗格重开后的未保存工作簿是另一个会话身份」✅（`wpsSdkBridgeHost.ts:213` 的 `wps-session:`） | 两端 ⏳：真实窗格里开两个未保存同名簿 |
| 3 | **工作表身份变化**（改名 / 删表 / 重排） | 生产入口同 #1。`sdkBridgeHost.test.ts:138`（收窄也拒）+ **本轮 E3**（改名也拒，张数顺序都不变）✅ | `wpsSdkBridgeHost.test.ts:295 / :368`、`WpsAgentPanel.test.tsx:3063 / :5017`（改名 → 待答澄清失效 / `unauthorized_scope` 零写入）✅ | 删表后的**原地撤销**另属 `worksheet_restore`（见 `run_worksheet_restore_mutations.mjs` R1–R8，原生未验收）；重排场景两端都缺专测 |
| 4 | **授权收窄** | `sdkBridgeHost.test.ts:138 / :329`（`unauthorized_scope` 且**不读宿主**）✅ | `wpsSdkBridgeHost.test.ts:295 / :368`、`WpsAgentPanel.test.tsx:5017` ✅ | ⏳ 两端：数据范围面板上真的取消一个勾选（面板交互 + 视觉） |
| 5 | **旧模型结果迟到** | `test_result_ref_is_rejected_from_another_session`、`test_workbook_change_invalidates_retained_results`、`sdkIntentRouting` 的"过期结果"链（`resultExpiryRecoveryChain.test.tsx`）✅ | A1-E2 的"A1-E1 结果引用过期验收入口"、"SDK response provenance cannot replace a trusted bridge"2 条 ✅ | ⏳ 真实模型真的吐出旧引用（需真实 SDK + 独立端口） |
| 6 | **旧桥事件迟到** | 桥按 `operationId + requestFingerprint` 判 `duplicate / conflict / late`；停止态一律 `cancelled`（**S2**）✅ | 同判据（共享 `packages/core/src/sdkToolBridge.ts`）；`WpsAgentPanel.test.tsx`「不把刷新后的在途活动当仍在运行」✅ | ⏳ 两端：真实窗格里把进度事件拖到收尾之后再让它到 |
| 7 | **旧执行者迟到 / 被替代** | `test_resume_claim_is_compare_and_swap_and_fences_old_executor`（旧执行者 commit 被"过期执行者"拒）+ **C1**（旧 `owner_token` 开动作被拒）✅ | `wpsExecutionResume.test.ts:147`「lost claim / concurrent owner never writes」（409 owner active → 写入 `[]`）✅ | ⏳ 真实双窗格 |
| 8 | **同时恢复 / 快速重复点击** | `excelWriteEndToEnd.test.tsx:620` 连点两次"核对并继续"：`registerOfficeExecution` 1 次、`startOfficeAction` 索引 `[0,1,2]`、续跑只写剩余 1 格 ✅；**C1/C2** 提供权威侧真并发证据 ✅ | `wpsExecutionResume.test.ts:147` + `claimWpsExecutionResume` CAS ✅；**C2** 同 | ⏳ 两端：真实双击 / 真实两个面板 |
| 9 | **写入后回执丢失** | `excelWriteEndToEnd.test.tsx:1359`「重载后按挂接索引核对回执：只汇报，不重复执行已完成动作」（6 个写接口一个都没再调）✅、`:1354`「已完成前缀但恢复快照丢失：拒绝续写」（生成不了只含后缀的撤销点）✅ | `wpsExecutionReconcile.test.ts:75 / :109 / :158 / :172 / :200`（分类正确 + 剩余 0 只补记）+ `test_task_ledger.py:592`（in-flight 未核对不许认领；补记后 `next_action_index=1`、旧 token 失效）✅；**原生 2026-10-28 ✅**：`wps-execution-1fd938fd…` 第二步 `failed_after_write` 后补记，`office_host_writes` 里 `actionIndex=1` 仍只有原失败那一刻的 1 行（见 `A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md` 同日条） | ⏳ Excel：真实断开（关任务窗格 / 断网）后再回执。WPS 原生还缺"补记时刻"的独立记录——账本无逐次修订历史 |
| 10 | **第二步中断后只续第三步** | `excelWriteEndToEnd.test.tsx:620`（`从第 3 步继续并完成`、`completeOfficeAction` 索引 `[0,2]`）✅；**原生 2026-10-29 ✅（账本级）**：`execution-c52ad791…` 三步写入，`office_host_writes` 0/1 号同秒（16:03:44）、2 号在 18:37:11（相隔 2h33m27s），`next_action_index=3`、`in_flight=NULL`、`checkpoint.uncertainActionIndexes=[]`、回读 B2/B3/B4 逐格一致（`outputs/a1-acceptance-20261029-2/EVIDENCE.md`） | `wpsExecutionResume.test.ts:38`「resumes only third」、`:106`（只写 `Target!A3`、`startWpsAction` 1 次）✅；**原生 2026-10-28 ✅**：同一执行 15:49:06 只新增 `actionIndex=2` 一行（1 次变更调用 + 回读 verified），B4 检查点 `-303 → 333` 两个指纹可重算命中 | ⏳ 两端：**中断是开发探针制造的**，不是"回执丢失"（#9 仍空）；Excel 的"核对零写入"只是账本推断（核对不写 `office_host_writes`，中间没有第三个写入时刻），计数口径仍是端上自报 |
| 11 | **异常退出后恢复** | `test_abrupt_child_process_exit_keeps_committed_event_for_reconciliation`（服务端）+ Excel 重开链 ✅ | `wpsExecutionResume` 的重开夹具 + `wpsRecoveryInspection.test.ts` ✅；**原生 2026-10-28 夜间（正常退出重开）✅**：`wps-execution-3f00ea0d…` 完整退出 WPS 后重开，恢复材料跨会话可用、补记不重写、只续第三步（`ABANDON_AFTER_RESUME_ACCEPTANCE_2026-10-28.md` §16） | ⏳ **崩溃/被杀进程**（Excel / WPS 被任务管理器结束）——正常退出版不算；两端都还没有 |
| 12 | **完成 / 确认 / 放弃 / 撤销后重开** | `excelWriteEndToEnd.test.tsx:620` 的 `confirm` / `discard` 两分支重开后各自的终态与写次数 ✅ | `wpsResultReviewMarkers.test.ts:143 / :488 / :515`；`wpsExecutionResume.test.ts:166`（第二次中断保留原年龄与回滚内容）；`:360`「空句柄无法找回时仍能 confirm/discard，**审阅与比较不会重执行**」✅；**原生 2026-10-28 两半都有证据**：①确认后**保存并完整退出重开**，结果仍为 `111/222/333`、该执行 `status=completed` 不再出现在恢复候选（`pending_office_executions` 只取 `running/needs_reconciliation/failed`），账本侧只见证"无新写入、无新执行"（"确认"本身无服务端记录）；②**续跑后放弃**（同会话）：三个动作全部回到 `-101/-202/-303`、恢复材料目录被收走、账本 `revision` 未再推进（`ABANDON_AFTER_RESUME_ACCEPTANCE_2026-10-28.md` §10）；③**跨进程版**（同日 23:25→23:39，`wps-execution-3f00ea0d…`）：第二步中断→保存→**完整退出 WPS**→重开→核对→续跑→放弃，三个动作回 `-101/-202/-303`，保存后与原始副本**逐格 0 差异**、条件格式规则 0 条，重开不再出现恢复卡（同文档 §16）；④**收尾轮**（19:38/19:42）：放弃之后**原对话直接发只读查询**正常出结果（`turn-68fd1b…`），确认之后同样能继续（`turn-4ec317…`，但那一轮账本**没有中断记录**，只算"普通确认执行后"），同 §15 | ⏳ 两端：真实**撤销**（结果卡「撤销本次执行」）未验；「**续跑后**确认 + 随后对话」未验（跨进程的放弃版已过，见上）；「确认后重开」的单元格值仍只有截图（放弃版的磁盘值已有：逐格 0 差异） |

### 3.1 真正的缺口现在只剩"原生没跑过"

> **2026-10-28 更新**：WPS 的 #9 / #10 已有原生证据（用户跑完"中断 → 补记 → 只续第三步 → 确认变更"，
> 账本与宿主观测只读核验通过）。#12 的 WPS 两半也补上了：确认后**保存并完整退出重开**
> 结果保持、以及**续跑后放弃**回到三个动作的执行前值（同会话，非跨进程）。
> Excel 同类路径仍是 `⏳`（#9 / #10 / #12 都没有原生证据）。
> 同日还修掉两个真机收尾缺陷（恢复完成后 `pendingPlan` 不收尾 ⇒ 新任务被永久挡住；
> 撤销说明的"改动 2 / 3"把中断阶段统计说成整次执行、且把"回读失败"算成改动），
> 见 `ABANDON_AFTER_RESUME_ACCEPTANCE_2026-10-28.md` §11。**A1-GATE 仍不通过**：
> 真实 Excel 一端、崩溃版恢复、真实撤销都还空着。
>
> **2026-10-29 更新（Excel 真机一轮）**：Excel 的 #10 拿到**账本级原生证据**
> （`execution-c52ad791…`，中断由开发探针制造）；#11 只是**窗格重开**（`EXCEL.EXE` 未退出），
> 不算跨进程；#12 的"确认"分支只有"终态由账本支撑 + 界面口述"，**"确认"这个动作本身服务端
> 无记录**（U4 缺口实证，本轮已实现 `abandoned_at` / `undone_at`，见
> `docs/EXECUTION_DISPOSITION_2026-10-29.md`，**不追溯**这一轮）。
> 同轮还确认两件不能宣称的事：**夹具从未保存**（mtime 早于执行、`汇总表 B2:B4` 仍是
> `-101/-202/-303`），"保存落盘"一档不成立；**"本机索引丢失后由后端发现"证据不足**
> （账本无发现记录、`discovery_dismissed_at` 仍为 `NULL`），不得因恢复卡出现了就标通过。
> 证据分档见 `outputs/a1-acceptance-20261029-2/EVIDENCE.md`。**A1-GATE 仍不通过**：
> 崩溃版恢复、真实撤销、S1、Excel 的"放弃"分支仍空着。
>
> **2026-10-28 收尾轮补充**：#12 的"收尾之后还能不能接着聊"这一问，在 WPS 上分开记两笔 ——
> **放弃之后可以继续对话 = 真机通过**（`5a2d84be` 中断+续跑 → `turn-68fd1b…` 正常出结果）；
> **确认之后可以继续对话 = 真机通过，但那一轮账本没有中断记录**（`7afd3dc6` 三步同一毫秒写完、
> `revision=8`），所以只能记成"普通确认执行后"，"续跑后确认 + 随后对话"仍缺。
> 同日另记一个**未修复**的独立待查项（明确查 `B2:B4` 却提示"结果不完整（扫描已达上限）"）。
> 全部剩余项按"尚未实现 / 自动化已验但真机未验 / 证据待补"三分类见
> `A1_GATE_REMAINING_2026-10-28.md`。
>
> **2026-10-28 夜间再补一笔**：#11 与 #12 的 WPS **跨进程（正常退出重开）**版通过
> （`wps-execution-3f00ea0d…`：第二步中断 → 保存 → 完整退出 → 重开 → 核对 → 续跑 → 放弃；
> 保存后与原始副本逐格 0 差异、条件格式规则 0 条；重开不再出现恢复卡）。
> 这是第一次拿到**落盘后的**单元格证据，此前每一轮都只能靠截图。
> 但重开首次出现「证据不足」、点一次「重新核对（只读）」才恢复，**未归因**，单列为待查项
> （同文档 §6）。**崩溃/被杀进程**仍未验，不得用正常退出版代替。

第 2 行 Excel 未保存同名簿的**读路径**已修（§2.0 与新的一节 §7）。写路径从来是安全的：
服务端 `OfficeExecutionRegisterRequest` / `WpsExecutionRegisterRequest` 要求
`url:` / `wps-path:` 前缀，未保存的会话内身份一律拒登记。

未保存工作簿的身份**只在一个加载会话内有意义**，这是宿主给不出更多东西：
- 窗格重开 = 新 JS 环境 = 新 document 对象 = 新令牌 ⇒ 服务端比对检查点里的身份判
  `PLANNING_CONTEXT_TARGET_CHANGED`（"当前工作簿与原任务的不是同一份…请按当前工作簿
  重新发起需求"）。措辞对"重开的同一份未保存簿"略重，但方向正确、动作正确；
- 同一会话内可以正常继续分析（身份、授权、来源指纹三项照旧重读核对）；
- **没有**为拿身份写 `document.settings`，也没有要求用户先保存才能读表。

---

## 4. 反向变异（唯一判别点，逐字节还原）

| 变异 | 改了什么 | 预期红点 | 实测 | 还原 |
| --- | --- | --- | --- | --- |
| S1 | 桥不在 `source_changed` 时停 | `session.source_stopped is True` | 红（1 条） | `task_bridge.py` d68015ee… |
| S2 | SDK 工具分支不触发停止 | `SourceChangedDuringPlanning` | 红 | `claude_sdk_agent.py` 93daaa29… |
| S3 | 流式循环不看停止信号 | `SourceChangedDuringPlanning` | 红 | 同上 |
| S4 | 太松：任何工具错误都当终态 | `assert not isinstance(caught.value, SourceChangedDuringPlanning)` | 红 | 同上 |
| D1 | 验收开关默认打开（生产注入口失控） | `assert True is False`（`enabled()`） | 红（1 条） | `dev_source_changed_fault.py` 75fc5d59… |
| W1 | 面板不原样透出服务端停止原因 | `应显示来源变化导致的停止` | 红（1 条） | `WpsAgentPanel.tsx` 6647c078… |
| W2 | 把 `SOURCE_CHANGED` 收尾成 `awaiting_input`（文案不变，只改终态） | `expected 'awaiting_input' to be 'failed'` | 红（1 条） | 同上 |
| C1 | `claim_office_resume` 去掉 revision 条件 | `并发认领出现了 8 个执行权` | 红（1 条） | `task_ledger.py` 17eec2cb… |
| C2 | 规划侧**两层**版本判据一起拆（事务内预检 + 条件更新） | `并发继续出现了 8 个执行权` | 红（1 条） | 同上 |
| M1 | 未保存簿的身份退回 `name:`（＝修复前行为） | `expected { ok: true } to deeply equal { ok: false, errorCode: 'source_changed' }` | 红（3 条） | `apps/excel-addin/src/excel.ts` 逐字还原，复跑 5 passed；日志 `outputs/verify-a1e3-id/identity-red.log` |
| F1 | 故障入口只认"绑过任务"，不认绑给**哪个**任务 | `assert dev_source_changed_fault.status()['phase'] == 'armed'`（被另一本簿抢走） | 红（3 条：抢走 / 同名不同簿 / 开桥绑定） | `server/app/dev_source_changed_fault.py` 逐字还原，复跑 18 passed |

**C2 第一次只拆了一层，用例仍然绿**——记录在案：规划侧的预检与条件更新是冗余层，
只拆一层不能证明判据有判别力；必须一次拆全。这正是"变异要有唯一判别点，冗余层一次拆全"
的实例，不是"用例空过"。

驱动脚本：`outputs/mut-a1e3/mutate.py`（S、D）、`mutate_wps_panel.py`（W）、`mutate_claim.py`（C），
每次运行前后各核一次 sha256，日志在同目录 `mut-*.log`。

---

## 5. 未闭环（进下一批前先看）

### 5.1 Excel

- 原生未验：来源变化停止的真机表现（消息、活动卡、有没有写入）、双击"核对并继续"在真机上的
  效果（替身里已经证明只执行一次）、真实断开后的回执补记、真实撤销。
- 原生未验：**两个未保存同名簿**（各开一个、都叫「工作簿1」）在真实任务窗格里不会互相读到对方。
  自动化只到"真实 `bridgeLiveContext` 读取器 + 真实复核判据 + 两个宿主文档对象"。
- 数据范围面板**标题仍在滚动容器内**（只改了 WPS），属既有已知差异。

### 5.2 WPS

- **宿主写调用计数已接入（2026-10-25 第三批，见 §8.2）**。此前的"三个写调用点未打点"已不成立；
  现在 WPS 的口径是**逐次写 API 调用**（分块写每块一次），与 Excel 的"刷新批次"**不互相折算**。
  仍属未验的是：真实 WPS 窗格里这些调用次数是否与替身一致（`⏳`）。
- 原生未验：来源变化停止后**旧继续入口真的消失**、另一条独立暂停任务**真的还在**、
  真实窗格里改名/收窄/换簿、真实双窗格并发恢复。
- 并排审阅布局（35/35/30）与审阅副本生命周期的真机截图仍未执行。

### 5.3 共同

- A1-GATE 条件（原生证据）不满足。
- 真实模型规划本批未重跑；本批停止行为只到"服务端 + 替身宿主 + 真实面板组件"。
- Excel 的宿主写调用数只到"**刷新批次**"这一档（Office.js 一次 `sync` 可能承载多个写操作），
  要精确到"写操作个数"需要逐个写入点打点，本批没做。**WPS 反过来**：它记的是逐次写 API 调用，
  两者的数字放在同一张表里比较没有意义（§8.1.4）。

---

## 6. 下一步

1. 按 `docs/A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md` 做**最小**原生验收（Excel 与 WPS 各 3 个场景，
   每场景只需点几下、只需保存两样东西）。写入次数改成"导出一次 `/host-writes`"。
2. 原生结果回来后，把 §3 的 `⏳` 逐行改成 `✅/❌`，A1-GATE 才谈得上通过。
3. WPS 宿主写调用计数（§5.2 第一条）：在三个写调用点累加 + 面板补报，判据直接复用
   `@excel-bro/core/hostWriteAudit`，不需要新的口径讨论。

---

## 7. 「实际写入次数」的口径（2026-10-25 第二批）

### 7.1 旧口径错在哪

验收包过去写的是"从 arm 到结束不应出现 `POST /executions/{id}/actions/start`，出现一次就是缺陷"，
以及"`actions/start` 恰好 3 次"。这两条都把**请求**当成了**事实**：

| 现象 | 真实含义 |
| --- | --- |
| `actions/start` 出现 1 次 | 有一个执行者**申请**写第 N 个动作；可能被账本拒（409），也可能认领后没碰宿主就失败 |
| 出现 3 次 | 三次申请，不保证三次写入 |
| 一次都没有 | 确实没有申请过 —— 这一条仍然成立 |

所以现在把这四件事分开说，并且**只用能观测到的那一档下结论**：

| 档位 | 来源 | 能证明 | 不能证明 |
| --- | --- | --- | --- |
| 请求次数 | 服务端访问日志 / `GET .../host-writes` 之外的请求行 | 有人申请过 | 写入是否发生 |
| 认领结果 | 账本：409/422 与 `ownerToken` 是否仍然有效 | 谁拿到了写这次动作的权利 | 是否真的写了 |
| **宿主写调用数** | 端上执行器在**真的调用宿主写 API 处**累加，best-effort 补报 | 向宿主发起了几次写调用 | 单元格最终值对不对 |
| 结果与回读 | 端上 `outcome` / `readback` + 账本检查点 + 整份计划回读 | 写完没有、回读是否通过 | 未来不会被改 |

### 7.2 新增的只读取证端点（两端同形）

```
POST /api/office/executions/{operationId}/host-writes     # 端上补报（best-effort）
GET  /api/office/executions/{operationId}/host-writes     # 读汇总（人来判读）
POST /api/wps/executions/{operationId}/host-writes
GET  /api/wps/executions/{operationId}/host-writes
```

- 只存**计数与状态**：`actionIndex / actionType / hostWriteCalls / mutatingCalls / outcome /
  readback / errorCode`，**没有**单元格值、公式正文、文件名或工具参数；
- 追加式（`office_host_writes` 表），同动作多条时按**最后一条**算；查询不改任何执行状态；
- `ownerToken` 不一致 ⇒ 409 `HOST_WRITES_REJECTED`；未知或另一端的 operationId ⇒ 404；
- 上报是**观测**：调用方在账本提交（`finish`/`suspend`）**之后**才发，失败只少一条证据，
  不会给"写入 → 账本提交"之间新增任何中断窗口。

`outcome` 四态的含义（`packages/core/src/hostWriteAudit.ts` 是唯一口径）：

| outcome | 含义 | 计入 `hostWriteCalls` |
| --- | --- | --- |
| `written` | 宿主调用已返回 | 是 |
| `failed_before_write` | 还没进宿主写 API 就失败 | **否**（这就是"认领后写前失败不算已写入"） |
| `failed_after_write` | 进过写 API 但随后失败（可能写了一半） | 是 |
| `unknown` | 没拿到确定的宿主结果 | 是（单独计数，不并进"已写入"） |

### 7.3 双端状态

| 端 | 宿主写调用数 | 结果/回读 | 证据 |
| --- | --- | --- | --- |
| Excel | ✅ 已实现：动作窗口内**刷新批次**数（包住本次 `Excel.run` 的 `context.sync`，按动作取差值） | ✅ `written` + 恢复模式 `readback: verified`；失败按"窗口内刷新 0 次 ⇒ `failed_before_write`，否则 `failed_after_write`" | `apps/excel-addin/src/excel.ts`（`onHostWrite`）、`App.tsx`（`flushHostWriteAudits`）、`server/tests/test_host_write_audit.py` |
| WPS | ✅ **第三批已实现**：**逐次写 API 调用**数（值/公式分块写每块一次；格式动作按 setter 各一次；边框恢复一条边一次），辅助调用单列 `mutatingCalls` | ✅ 同上四态；写前失败 `hostWriteCalls === 0`，写 API 抛错记 `failed_after_write` | `apps/wps-addin/src/wpsHostWriteCounter.ts`、`wpsWriteMatrix.ts`、`wpsBorders.ts`、`wpsPlanExecutor.ts`、`WpsAgentPanel.tsx`、`apps/wps-addin/src/wpsHostWriteAudit.test.ts`（10 条） |

**替身宿主 ≠ 实机**：Excel 的计数是"真实执行器 + jsdom 替身上下文"级证据；真实窗格里
`context.sync` 的次数与分块策略是否与替身一致，要等原生验收。WPS 同理 —— 计数点在真实
赋值语句上，但 WPS JSAPI 在真机上是否把这些赋值合并成更少的 API 调用，只能在真机上看。

### 7.4 "零写入"仍然不成立（第三批修正）

上面这张表说的是"能采到调用数"，**不是**"没采到就等于 0"。上报是 best-effort，
`entries = []` 与"明确 0 次调用"是两句话 —— 第三批加了 `observation` 判据来分开它们，
见 §8.1。旧验收包里"空 host-writes 即证明零写入"的写法已经删掉。


---

## 8. 第三批：三个问题（2026-10-25）

本批只做三件事，不扩到 A2：**①「没收到观测」不能冒充「明确零写入」 ② WPS 写调用观测
（此前确实未实现）③ 开发故障入口的同名工作簿串用（真缺陷）**。
三条都已实现并有自动化闸门；**均未实机验收**，A1-GATE 仍不通过。

### 8.1 问题一：没有观测 ≠ 明确零写入（`HostWritesObservation`）

#### 8.1.1 旧写法错在哪

上报是 best-effort：面板在账本 `finish`/`suspend` 之后才补报，一次都到不了是完全正常的。
所以 `reportedActions=0 / hostWriteCalls=0 / entries=[]` 同时对应两件**完全不同**的事：

| 实际情况 | 旧读法 | 后果 |
| --- | --- | --- |
| 一条上报都没到 | "零写入" | 把"没测"当成"没发生"，验收结论无效 |
| 上报到了、每次都是 0 次调用 | "零写入" | 这才是真的 |
| 只到了一部分 | "零写入" | 漏掉的那几个动作可能写过 |
| 有 `unknown` / `failed_after_write` | "零写入" | 可能写了一半 |

#### 8.1.2 可信执行边界：谁能算"可能写过"

写入路径只有一个入口：`actions/start` 认领 → 宿主写 API。所以"可能碰过宿主"的动作集合由
账本决定，不是由上报决定：

```
attempted = max(next_action_index, in_flight_action_index + 1)   # 再按计划动作数截断
```

没认领过的动作**不可能**碰过宿主。这就是"明确零"的执行边界：只有边界内**每个**动作都有观测，
"零"才有依据。实现：`server/app/main.py` 的 `_host_write_boundary()`。

#### 8.1.3 判据：`observation` × `callEvidence`

`GET /api/{office|wps}/executions/{id}/host-writes` 新增 `observation` 字段：

| `observation` | `callEvidence` | 能说的话 |
| --- | --- | --- |
| `not_collected` | `none` | 没收到观测 —— 对宿主一无所知（另看 `attemptedActions` 是否 > 0） |
| `partial` | `nonzero` | 已看到非零调用（覆盖不全，同看 `missingActions`） |
| `partial` | `unknown` | 只覆盖一部分 ⇒ **不能说零** |
| `complete` | `zero` | **明确观测到零写调用**（唯一一种） |
| `complete` | `unknown` | 有条目结果未知 ⇒ 不能说零 |

配套字段：`attemptedActions` / `reportedActions` / `missingActions` / `unexpectedReports` /
`unit` / `lastReportedAt`。**`unexpectedReports`（落在边界外的上报）本身就是矛盾证据** ——
账本说没认领过，端上却说写过，只能记 `unknown`。

展示层唯一口径：`packages/core/src/hostWriteAudit.ts` 的 `describeHostWriteObservation()`，
`provesZeroWrites` **只在 `complete + zero` 且自洽时为真**。这一层还做一次**自洽复核**：
`zero` 却带着 `missingActions > 0` 或 `unexpectedReports > 0` 时**拒绝升级**成"零写入"。
那不是第二份覆盖推断（覆盖判定仍只有服务端一份），而是拒绝把矛盾的线上数据当更强的结论——
这条是第三批的变异 H4 钉住的（见 §8.4），它第一次跑就抓到一个真漏洞：`partial + zero`
当时会被读成"明确零写入"。

#### 8.1.4 两端计量单位不同，不许互相折算

```
unit = { office: "excel_sync_batches", wps: "wps_api_calls" }
```

- **Excel = 刷新批次**：Office.js 把排队的写操作按批提交，一次 `context.sync` 可能承载多个
  写操作 ⇒ 一个动作写 2 个单元格也可能是 1 个批次；
- **WPS = 写 API 调用**：逐次赋值调用，分块写每块一次。

同一个动作在两端的数字天生不同。验收表里只写数字不写单位 = 下次没人能判断那是什么。

#### 8.1.5 遥测不得影响业务执行

- 上报在账本提交**之后**（`finish`/`suspend` 尾部 `flushHostWriteAudits()`），
  `void ... .catch(() => {})`：失败只少一条证据；
- 不改执行状态、不推进 `revision`、不触发重写、不改变已确认结果；
- 端上观测回调抛错被 `emitHostWrite` 吞掉（用例「观测回调抛错不影响执行」钉住：
  `record.state === "verified"` 且单元格真的写进去了）。

#### 8.1.6 回归（`server/tests/test_host_write_observation.py`，10 条全绿）

| 用例 | 钉住的事实 |
| --- | --- |
| `test_no_report_is_not_zero_writes` | 空上报 ⇒ `not_collected` + `none`，但 `attemptedActions=1`（边界是真的） |
| `test_no_action_claimed_is_a_ledger_fact_not_an_observation` | "账本说没认领" ≠ "观测到零" |
| `test_partial_coverage_can_never_claim_zero` | 覆盖 1/2 ⇒ `partial` + `unknown` |
| `test_full_coverage_with_zero_calls_is_the_only_zero` | 唯一一种 `complete` + `zero` |
| `test_nonzero_calls_are_reported_even_when_coverage_is_incomplete` | 非零与覆盖不全**分别记录**，不许互相抵消 |
| `test_unknown_outcome_blocks_the_zero_verdict` | 覆盖全但有条目 `unknown` ⇒ 不是零 |
| `test_report_outside_the_ledger_boundary_is_flagged_not_counted_as_zero` | 边界外上报 ⇒ `unexpectedReports=1` + `unknown` |
| `test_lost_half_of_the_reports_does_not_turn_into_zero` | 后一半丢了不把已有数字抹成 0 |
| `test_unknown_operation_never_returns_a_verified_looking_zero` | 记录不存在/跨命名空间 ⇒ 404 `EXECUTION_NOT_FOUND`，没有"读到 0"这回事 |
| `test_wps_namespace_declares_its_own_unit` | WPS 的单位是 `wps_api_calls` |

#### 8.1.7 规划期就终止（没有 executionId）怎么验

`source_changed` 停在规划阶段时**根本没有执行记录**，那时没有任何 `executionId` 可查 ——
所以不许拿一个"看起来像"的 ID 去查空记录（那恰好会返回 404，很容易被误读成"零写入"）。
这种情况下的验证方法在验收包 §1.3 重写成三条可操作的事实：

1. 面板与 SSE 帧给出 `SOURCE_CHANGED`（本次规划的终态）；
2. **没有** `POST /api/office/executions/register`（连执行记录都没建）—— 服务端访问日志里看；
3. 工作簿单元格与开始时逐格一致（人在 Excel/WPS 里看，或对比另存的副本）。

只有第 1 条成立、第 2 条不成立（有 register 但没读完）时，才去查 `host-writes`，并按
§8.1.3 的口径读 —— 而不是按"空数组"读。

### 8.2 问题二：WPS 写调用观测（此前 ❌ 未实现）

#### 8.2.1 计数点在真实的宿主赋值处

| 位置 | 计数 | 语义 |
| --- | --- | --- |
| `wpsWriteMatrix.ts` `range[property] = ...` 前 | `countHostWrite(true)` | 计划写（1 块 1 次） |
| `wpsWriteMatrix.ts` 每个 `cell.NumberFormat = "@"` 与恢复赋值前 | `countHostWrite(false)` | **辅助**调用（保文本原样，只进 `hostWriteCalls`） |
| `wpsBorders.ts` `restoreWpsBorder` 每条边 | `countHostWrite(true)` | 计划写 |
| `wpsPlanExecutor.applyFormatAction` 每个 setter | `countHostWrite(true)` | 计划写（setFill 2 次 / setFont 5 次 / setAlignment 3 次 / RowHeight 1 次） |

**不是**拿 `actions/start` 或动作完成事件替代：那些是申请与回执，不是宿主调用。

#### 8.2.2 作用域把辅助写天然排除

`wpsHostWriteCounter.ts` 是**作用域式**计数：`beginHostWriteScope()` / `endHostWriteScope()` /
`countHostWrite()` / `hostWriteScopeActive()`。没有活动作用域时 `countHostWrite` 直接返回。

因此**审阅副本、撤销（`wpsUndoRecordFor` 的恢复写）、清标记、`PluginStorage`** 全都不计 ——
它们不在动作窗口里。用例「撤销与辅助写不进观测」直接调 `writeWpsMatrix` 验证：
作用域外一次都不产生条目，而撤销记录仍然是 `verified`。

**如实记一条边界**：撤销/放弃的宿主写现在**既不进计划写调用数，也没有单独的"撤销写调用数"** ——
它自己的账是撤销点状态（`state` / `appliedRows` / 失败清单），不是本模块的计数。
所以"放弃变更后 `hostWriteCalls` 没变"是**正确**的，不能读成"没写"。要真正按次数核对撤销，
需要再给撤销路径单独开一个计数命名空间；本轮没有做，也没有拿它代替任何现有判据。

#### 8.2.3 两个计数不能混

- `hostWriteCalls`：动作窗口内**所有**宿主写调用（含辅助）；
- `mutatingCalls`：只数承载**计划内容**的调用。

混成一个数就会把"为了保文本原样改了两格格式"当成"写了两次"，重放判据直接失真。
变异 H6 把辅助调用改成 `countHostWrite(true)`，用例立刻从 `mutatingCalls=1` 变红成 3。

#### 8.2.4 覆盖清单（都有用例）

| 场景 | 期望 | 用例 |
| --- | --- | --- |
| 一个动作只写一块 | `hostWriteCalls=1`、`mutatingCalls=1`、`written` + `verified` | 一个动作只写一块时… |
| 120 行分块 | 块数 = 调用次数（2），不是单元格数（120） | 分块写… |
| 保文本格式 | 5 calls / 1 mutating | 文本保真用的 NumberFormat… |
| **认领被拒** | `failed_before_write` + 0 calls + `not_applicable` | 认领被拒… |
| 预览后目标就变了（执行边界之外） | **一条条目都没有**（缺席 ≠ 记零） | 预览后目标就变了… |
| 写 API 抛错 | `failed_after_write` + 1 call + `WPS_ACTION_FAILED` | 宿主写 API 抛错… |
| 撤销/辅助写 | 不计 | 撤销与辅助写不进观测 |
| 观测回调抛错 | 执行照常 `verified` | 观测回调抛错不影响执行 |
| 第二步写后回执丢失→只续第三步 | 只有 `[2]` 一条，前两步一次都没重记 | 恢复只写剩余动作时… |
| 作用域嵌套 | 退出恢复外层计数 | 作用域可嵌套… |

"确认变更不增加计划写调用"由账本侧保证（确认只清标记、不重执行；见主线计划既有结论），
本轮没有新增 WPS 侧用例来重复它。

#### 8.2.5 隐私

条目只有 `actionIndex / actionType / hostWriteCalls / mutatingCalls / outcome / readback /
errorCode`。**没有**单元格值、公式正文、完整工具参数、文件名或密钥。`actionType` 是白名单
里的名字，不含参数。

### 8.3 问题三：开发故障入口的同名工作簿串用（真缺陷，先红后修）

#### 8.3.1 改造前的实证

`arm` 只要求 `hostId + workbookName`，`workbookIdentity` 可选 ⇒ 两本同名簿落进同一个键，
**谁先开桥谁把这次单发拿走**（探针 `outputs/verify-a1e3-round3/fault_identity_probe.py`，
前后输出 `fault-identity-before.txt` / `-after.txt`）：

```
改造前
  arm(只有名字)                 -> 被接受 phase=armed
  另一本同名簿（乙）            -> 回执=error 派发=0 phase=consumed taskBound=False
  stolen = True
  目标簿（甲）                  -> 回执=ok 派发=1 phase=consumed
```

被抢的后果不是"多注入一次"：**被抢的那一轮真的停在半路**，而要验的那一本一次都看不到停止。

#### 8.3.2 修法：一个共用判据 `server/app/dev_fault_target.py`

| 规则 | 实现 |
| --- | --- |
| 可核验身份 = `url:` / `wps-path:` / `wps-session:` / `unsaved:` | `TRUSTED_IDENTITY_PREFIXES`、`is_trusted_identity()` |
| `name:<名字>` 回退值**不是**身份 | 同上（它本身就是按名字认簿） |
| 显式给了身份 ⇒ 用它 | `resolve_identity()` 返回 `("...", "declared")`，不可核验则抛 `IdentityUnavailable` |
| 没给 ⇒ 只接受"服务端只见过这一个名字下的一个身份" | `("...", "observed")`；这是**证据**，不是猜测 |
| 零个或多个 ⇒ 拒绝并说明原因 | `IdentityUnavailable`，消息里点明缺"身份"、怎么取（`seenTargets`） |
| 任务绑定只认身份相等 | `bind_task()`：`is_trusted_identity(target)` 为假 ⇒ 不绑；`target != armed.workbook_identity` ⇒ 不绑 |

只有可信路径才登记目标：`TaskToolBridge.open()` 末尾（那时宿主/名字/身份都已确定，且**只在
开发开关打开时**登记），并**只读身份**，不读任何业务行。`taskId` 仍由开桥路径绑定，不由
`arm` 指定。

#### 8.3.3 改造后

```
  arm(只有名字)                 -> 被拒绝: 缺少可核验的工作簿身份…
  seenTargets = ['url:https://example.com/甲/测试.xlsx']
  arm(只有名字, 但服务端只见过甲) -> phase=armed identityBound=True source=observed
  另一本同名簿（乙）            -> 回执=ok 派发=1 phase=armed      ← 不再被抢
  目标簿（甲）                  -> 回执=error 派发=0 phase=consumed
```

`status()` 的四个分支都带 `identityBound` / `identitySource` / `seenTargets`，
让人能从返回值里**抄**到该填的身份，而不是猜路径格式。

#### 8.3.4 E1 入口同步修复

`dev_result_ref_fault` 有同一处判据（只认名字、`fallback_workbook_identity` 回退）。
本批把它改成同一个 `dev_fault_target`，属于"同一身份边界内修复"，**没有**做无关重构：
`_Arm.identity` 直接返回可核验身份，`bind_task` 同样要求身份相等，`status()` 同样带
`identityBound` / `identitySource`。

#### 8.3.5 回归（`server/tests/test_source_changed_stop.py` 追加 6 条，与既有合计 37 passed）

| 用例 | 钉住的事实 |
| --- | --- |
| `test_dev_fault_refuses_to_arm_without_a_workbook_identity` | 只给名字 ⇒ 拒绝，理由里出现"身份"，`phase` 仍是 `idle` |
| `test_dev_fault_uses_the_only_identity_the_server_has_seen` | 只见过一个 ⇒ `observed` 补全（是证据不是猜） |
| `test_dev_fault_refuses_when_the_name_carries_two_identities` | 同名两身份 ⇒ 拒绝；显式指定后能启用（两条路互不替代） |
| `test_dev_fault_rejects_an_unverifiable_identity` | `name:` 前缀不算身份 |
| `test_dev_fault_does_not_bind_a_task_that_only_carries_a_name` | 任务侧只报名字 ⇒ 不绑、`phase` 仍 `armed`、那次调用照常派给宿主 |
| `test_dev_fault_from_a_reopened_unsaved_workbook_session_never_binds` | 窗格重开换了会话身份 ⇒ 旧启用不绑新会话 |
| `test_dev_fault_cancel_and_expiry_never_leave_a_half_armed_entry` | 取消/过期不留半启用 |

`test_result_ref_fault.py` 与 `test_python_coordinator.py` 同步改用可核验身份
（`url:https://example.com/测试.xlsx`），否则它们测的是"旧契约"。

### 8.4 反向变异（8 个单点，逐字节还原）

驱动：`node scripts/mutate_a1e3_host_writes.mjs <SEED>[ restore]` +
`node scripts/run_a1e3_host_writes_mutations.mjs`。判据五条 + 一条"变异确实改了文件"：
退出码非 0、有效摘要、**预期那条用例真的跑了并红了**、失败原文含**它自己那一处**的证据、
还原后 sha256 与变异前一致。

| 变异 | 改了什么 | 预期红点 | 实测 |
| --- | --- | --- | --- |
| H1 | 空上报直接判 `complete` + `zero` | `test_no_report_is_not_zero_writes` | 红（`'complete' == 'not_collected'`） |
| H2 | 忽略"落在执行边界之外"的上报 | `test_report_outside_the_ledger_boundary…` | 红（`'zero' == 'unknown'`） |
| H4 | 展示层照字面读 `zero`，不核对覆盖自洽 | core `矛盾的线上数据不许被升级` | 红（`expected true to be false`） |
| H5 | WPS 计划写不计数 | `一个动作只写一块时…` | 红（`"hostWriteCalls": 0`） |
| H6 | WPS 辅助改格式也算计划写 | `文本保真用的 NumberFormat…` | 红（`expected 3 to be 1`） |
| H7 | 服务端没见过身份时按名字补一个 | `test_dev_fault_refuses_to_arm_without…` | 红（`DID NOT RAISE`） |
| H8 | 同名多身份不再拒绝，取最近那个 | `test_dev_fault_refuses_when_the_name_carries_two…` | 红（`DID NOT RAISE`） |
| H9 | 任务侧只报名字也照样绑 | `test_dev_fault_does_not_bind_a_task_that_only_carries_a_name` | 红（`'bound' == 'armed'`） |

全部 `RED(OK)`，日志 `.tmp-mutations/H*.log`，汇总
`.tmp-mutations/summary-a1e3-host-writes.json`、`.tmp-mutations/run-a1e3-host-writes.txt`。

两次"判据串写错"记录在案（**不是**变异无效，是取证方式错了）：

1. H5 第一轮 `-t` 直接吃了整条用例名，名字里的**字面 `+`**（"written + 回读"）被当成量词 ⇒
   整个文件 10 条全 `skipped`、退出码 0。那是"没跑"，不是"没红"。改用无元字符的前缀。
2. H6/H9 第一轮的 `errorContains` 猜错了实际文案（H6 实际是 `expected 3 to be 1`；
   H9 更早一步就红在 `'bound' == 'armed'`）。宁可改证据串，也不放宽到"只要非 0 就算红"。

### 8.5 分档（不得合并宣称）

| 档 | 状态 |
| --- | --- |
| ① 工件回放 | ✅ 变异日志与探针输出都在仓库内 |
| ② 双端在线集成（真实服务端 + 真实执行器/面板 + 替身宿主） | ✅ Excel 既有；WPS 观测本批新增 |
| ③ 真实 Docker | 与本批无关 |
| ④ 真实模型规划 | 本批**未重跑**（不重复昂贵实验） |
| ⑤ 真实 Excel | ❌ 未做 |
| ⑥ 真实 WPS | ❌ 未做 |

**替身证据 ≠ 实机证据**：WPS 的写调用次数与 Excel 的刷新批次数在真机上是否与替身一致，
本轮一次都没看过。A1-GATE 不通过。
