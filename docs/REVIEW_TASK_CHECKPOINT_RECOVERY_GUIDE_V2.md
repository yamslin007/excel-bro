# 评审意见 v2：任务检查点、挂起与断点恢复开发指导

被评审文档：[`TASK_CHECKPOINT_RECOVERY_GUIDE.md`](TASK_CHECKPOINT_RECOVERY_GUIDE.md) v0.1（2026-09-08）
前一版评审：[`REVIEW_TASK_CHECKPOINT_RECOVERY_GUIDE.md`](REVIEW_TASK_CHECKPOINT_RECOVERY_GUIDE.md) v0.1（保留，不再作为结论依据）
本版日期：2026-09-08
评审基线：`ac1d896`
本轮未修改运行代码，未执行工作簿写入，未提交 Git

行号为核对当日位置，随开发会变化，请以符号名定位。

---

## 0. 本版为何存在

v1 评审的结论是"取消断点续跑、以 M1–M5 替代"。经用户对照代码核查，该结论所依赖的四项事实判断有误、一项推论混淆了时间点。**v1 的核心结论撤回。**

v1 的部分论证引用了未经直接核对的符号（其中若干符号在仓库中根本不存在，例如 `_undo_stacks`、`undo_folder_operation`、`_check_files_in_use`、`asyncio.to_thread`，grep 零命中）。本版所有代码断言均为直接阅读源文件确认，未核实项一律显式标注。

### 撤回与修正清单

| v1 中的判断 | 状态 | 本版结论 |
| --- | --- | --- |
| 2.1 「R2 目标只有 60 秒，接续收益无法覆盖成本」 | **撤回** | 60 秒是轮次边界阈值，非单次调用上界；且重跑成本不止时间 |
| 2.2 「执行时求值与检查点天然冲突」（阻塞） | **降级** | 不是架构矛盾，是待补齐的数据版本规则 |
| 2.3 「不引入 SQLite」（手段当结论） | **改写** | 真正要求是「单一权威记录」，存储选型可后定 |
| 2.3 「单机单进程，不存在多执行者竞争，lease 可整段删除」 | **撤回** | 跨 operation 执行确实并行且零互斥；可简化，不可删除 |
| 2.4 「§9 / R4 应删除」 | **保留但改置** | 不删除，移入第三阶段按收益逐项开放 |
| 4 「仓库不存在应用数据目录解析」 | **事实错误** | `packaging/runtime.py:37` 已有，缺的是开发/安装环境统一 |
| 4 「无 fsync」 | **表述错误** | manifest 有完整 fsync；工作簿本体没有。区分比概括更重要 |
| 5.1 「只缺扫描和按钮，数百行即可」 | **撤回** | 现有恢复入口依赖会话内存 + `ready` 状态，中断场景进不去 |

---

## 1. 总体结论

**采纳「缩小首期范围、补齐安全协议、分阶段实现真实续跑」的路线。** 原指导文件方向成立，但铺得过大，实施前需补齐四项前置设计（第 3 章）。

首期不追求完整现场重建，但**最终目标保留**跳过已完成节点、从断点继续。第一阶段的中断识别与安全核对是通往该目标的必经步骤，不是它的替代品。

存储选型（SQLite 与否）暂不决定。选定某种存储格式不等于解决一致性——先明确权威记录归属与恢复协议，再定载体。

---

## 2. 已确认的现状

以下均为直接阅读源文件确认。

### 2.1 应用数据目录：已有基础，缺环境统一

`packaging/runtime.py:37` `_data_root()`：macOS 返回 `~/Library/Application Support/Excel Bro`，其余平台返回 `%LOCALAPPDATA%/Excel Bro`。`main()`（`:43` 起）执行 `mkdir(parents=True, exist_ok=True)`、设置 `EXCEL_BRO_CONFIG_DIR` 与 `EXCEL_BRO_CAPABILITIES_PATH`、`os.chdir(data_root)`，并将日志写入 `data_root/excel-bro.log`。

**缺口**：该逻辑只在打包运行入口。开发环境直接启动 uvicorn 时不经过 `runtime.py`，`_data_root()` 不被调用，工作目录也不切换。新增任务存储必须统一两种环境的路径解析行为，否则开发期与安装期的恢复资料位置不一致。

### 2.2 落盘耐久性：元数据已具备，数据文件未具备

| 对象 | 落盘路径 | 耐久处理 |
| --- | --- | --- |
| manifest.json | `_atomic_write_manifest`（`folder_workbooks.py:568`） | 写入 → `flush()` → `os.fsync(fileno())` → `os.replace()`，含 `PermissionError` 三次重试（`:573-584`） |
| 工作簿本体 | `_save_workbooks_atomically`（`folder_workbooks.py:3386`） | `workbook.save(临时文件)` → `os.replace()`，**无 flush/fsync** |

**这个组合本身是风险点，比"有没有 fsync"更值得写入设计文档**：断电后可能出现 manifest 已记录 `applied`、而工作簿内容尚未落盘的**方向性不一致**。恢复协议必须假设"元数据比数据更新"这一偏向，不能假设两者同步。

### 2.3 并发面：跨 operation 真并行，零文件互斥

已确认的执行链：

- `_folder_operations_lock`（`main.py:608`）是 `asyncio.Lock`，**只保护注册表字典的读写**（`:918`、`:1000`），不保护执行体。
- `folder_execute_stream` 在锁内创建 `_FolderOperationState` 并 `asyncio.create_task(_run_folder_operation(...))`（`:952-959`）。
- `_run_folder_operation`（`:667`）通过 `run_in_threadpool(execute_folder_plan, request, on_progress)`（`:681`）把整个同步执行体投入线程池。**多个 operation 的执行因此在不同线程真并行。**
- 单个 operation 内部串行：`execute_folder_plan`（`folder_workbooks.py:5032`）是同步函数，用 `_OPEN_WRITE_WORKBOOKS` contextvar 管理打开的工作簿；contextvar 恰好提供线程隔离。**不存在单操作内并发写。**

互斥现状：

- 全仓库仅两个 `asyncio.Lock`：`turn_state.py:41`（`completion_lock`）与 `main.py:608`。
- **无 `threading.Lock`、无文件锁库（`msvcrt` / `portalocker` 零命中）、无按目标路径的占用标记或锁字典。**
- `operationId` 由前端生成（`api.ts` 约 `:517-518`，`Date.now()` 36 进制 + 随机）。请求指纹（`_folder_request_fingerprint`，`main.py` 约 `:611-618`）只在**同一 operationId** 下比对；同 operationId 不同内容会被拒（`:925-927`，`retryable=False`）。
- **两个不同 operationId 指向同一批文件时，不被任何机制拦截。** 多窗格、重复点击（若前端重新生成 ID）、恢复操作与原执行并存，都落在这个窗口里。
- `folderMaxPendingOperations`（100）是注册表容量上限，淘汰时按 `created_at` 排序移除 `status != "running"` 的条目（`:932-944`），不是并发度限制。

**结论**：分布式租约可以简化——不需要跨节点的 lease 续约与执行世代。但**三层互斥都需要**：任务级互斥、目标文件级互斥、恢复操作与原执行的排他控制。这三层都不为零。

### 2.4 现有恢复入口的真实前置条件

`_restore_folder_undo_impl`（`folder_workbooks.py:3761`）的门槛，逐条：

1. 从 `_session(request.sessionId)` 取 `session.undo_stack` —— **会话内存**，服务重启即失效。`_register_folder_undo`（`:3739`）还按 `undo.maxEntries`（5）截断（`:3755-3757`）。
2. 每条 record 要求 `manifest["version"] == BACKUP_MANIFEST_VERSION`（现行值 2，约 `:214-215`），否则拒绝（`:3777-3780`）。
3. 要求 `manifest["status"] == "ready"`（`:3781-3782`）。**中断留下的恰是 `applying` / `needs_attention`，直接被拒。**
4. 逐目标核对当前文件指纹等于 `post_fingerprint`（`:3795`），不等则拒绝以避免覆盖后续编辑。
5. 备份文件必须存在（`:3801-3804`）且其指纹等于 `prior_fingerprint`（`:3805-3811`）。

第 4、5 条的指纹核对逻辑质量高，**应当复用**。但入口整体建立在「会话内存 + `ready` 状态」之上，**重启后进不去，部分写入时也进不去**。

因此第一阶段需要的是一条**以 manifest 为起点的独立恢复路径**，而不是给现有入口加扫描器。v1 的"数百行"估计不成立。

### 2.5 规划阶段的时间上界

`excel_agent.py` 约 `:955-959`：`async with OpenAICompatibleClient(...)` 内 `for turn_index in range(loop_config.max_turns)`，循环顶部才刷新 `state.elapsed_time = time.time() - start_time`，随后 `should_stop`。

**`maxWallTimeSeconds`（60）是轮次边界的检查阈值，不是单次调用的时间上界。** 进行中的模型请求无法被该机制打断；`chat_completions` 的内部重试也在同一轮次内完成，其间不经过任何边界检查。

`llm/client.py` 的具体重试次数与退避参数**本版未独立核实**，故不引用数字。但机制层面的结论不依赖具体值：单轮实际耗时可显著超过 60 秒。

同时，重跑成本不止时间：**重复的模型调用费用、重复的用户澄清对话、用户等待**都计入。第二项尤其不能用秒数衡量——让用户重走一遍澄清是产品体验损失，不是性能损失。

### 2.6 其余已确认事实

- 仓库**无任何 SQLite / 持久化数据库使用**（仅 `test_installer_packaging.py` 提及路径）。
- `TurnRegistry`（`turn_state.py`）全内存 `OrderedDict`，TTL 3600 秒、`maxActive` 500。
- `run_excel_agent` 每次新建 `LoopState` 与 messages，不持久化循环现场。
- `_prune_folder_operations`（`main.py:723`）按 `created_at`（非 `finished_at`）计龄，只清理 `status != "running"`。淘汰即该 operation 全部事件消失，`afterSequence` 补收（`:962`，`min(afterSequence, len(events))`）将返回空而非缺口报错。
- `executeFolderPlan`（`api.ts` 约 `:521-646`）`attempt < 2`（最多重连一次），且前置条件为**错误不是 `ApiRequestError`**（约 `:625-635`）——即 `retryable = true` 的 `ApiRequestError` 也不会自动重试。
- `App.tsx` / `useActivityProgress.ts` 未见任何对执行状态接口的调用；重载后无自动挂接（已确认，非推断）。
- 关键限额：`agent.maxTurns` 8、`agent.maxWallTimeSeconds` 60.0、`agent.maxTotalTokens` 20000、`llm.timeoutSeconds` 60.0、`progress.folderOperationTtlSeconds` 3600、`folderMaxPendingOperations` 100、`folder.maxFiles` 100、`undo.maxEntries` 5、`folderBackupRetentionDays` 7。
- `计划表.md`（末次更新 2026-08-30）未出现断点恢复 / checkpoint / 任务持久化条目。

---

## 3. 实施前必须补齐的四项

原指导文件缺这四项，未补齐不具备实施条件。

### 3.1 收益评估（决定检查点粒度）

统计三项，用数据决定粒度而非预设：

- **中断频率**：按阶段（规划 / 计算 / 发布）分别统计异常退出与用户中断次数。
- **重跑耗时**：规划阶段实际耗时分布（重点是 P95 与最大值，因 2.5 已证明 60 秒不是上界）；文件夹计算与发布阶段按文件数与行数分档。
- **重复模型调用成本**：token 与费用，含调用已发出但结果未落地的部分。usage 不确定时记为未知，不计零（原文档 §10.2 此点正确，保留）。

粒度决策规则：只为"频率 × 重跑成本"显著的边界建检查点。

### 3.2 数据版本规则（原 2.2 降级后的正确形态）

区分两个时间点，这是执行时求值与检查点得以共存的关键：

| 时间点 | 规则 |
| --- | --- |
| 预览时 | 不把样本计算结果固定成最终写入值（保持活脚本语义） |
| 真正执行时 | 读取当时数据，**并记录本次执行实际使用的输入版本** |

中断后的判定：

1. 输入版本仍符合要求 **且** 产物完整 → 可以继续。
2. 输入发生变化 → 按依赖失效传播规则重算，或重新确认。
3. 无法判定输入版本 → 挂起，不猜测。

需明确列出：每类动作的"输入版本"如何定义（内容指纹 / 结构指纹 / 快照 ID）、变化后哪些下游节点必须重算、哪些变化需要重新预览确认。

### 3.3 状态冲突处理（原 2.3 改写后的正确形态）

要求不是"不用 SQLite"，而是：

- **明确单一权威记录**：谁的记录决定"是否可以再次写入"。考虑到 2.2 的方向性不一致（元数据可能比数据新），权威记录的结论必须能被实际文件内容否决。
- **明确核对顺序**：任务记录 → 备份记录 → 实际文件，三者不一致时按什么顺序收敛，谁负责判断。
- **明确证据分级**：沿用原文档 §7.3 的分级（匹配写后态 / 匹配写前态 / 部分应用 / 两者都不匹配 / 仅知同名存在）。最后一档必须判为证据不足。

存储载体在此之后再选。

### 3.4 并发互斥（原 2.3 撤回后的真实需求）

按 2.3 的现状，需要三层，均为本机进程内，不需要分布式租约：

1. **任务级互斥**：同一逻辑任务不允许两个执行体。当前 `operationId` 的复用检查只覆盖同 ID，需要补按任务身份的判定。
2. **目标文件级互斥**：按目标文件路径的排他控制。当前完全缺失，且写入执行在线程池中真并行。多文件需按稳定顺序获取以避免互相等待（原文档 §7.4 此点正确）。
3. **恢复与原执行的排他**：恢复操作不得与仍在运行的原执行并存。这一层在引入恢复入口后才出现，现在没有。

注意：软件内部互斥拦不住用户在另一个 Excel 实例改文件，占用检查、写前条件核对、写后验证仍需保留。

---

## 4. 三阶段范围

### 阶段一：中断识别与安全核对

**目标**：不误报成功、不重复写入、能重新找到任务与备份。

| 项 | 内容 |
| --- | --- |
| A1 | 以 manifest 为起点的独立恢复路径，能读取 `applying` / `needs_attention` 状态（现有入口按 2.4 第 3 条拒绝这些状态） |
| A2 | 启动时扫描遗留记录；恢复资料的定位不依赖会话内存 |
| A3 | 按 3.3 的证据分级给出核对结论；证据不足即挂起 |
| A4 | 界面出口：核对实际结果 / 恢复到执行前 / 放弃并保留资料。三者是不同操作，不合并为含糊的"恢复"按钮 |
| A5 | 三态分离表达：任务结束 / 写入成功 / 业务验收成功可分别显示（含"已写入，验证未通过"） |
| A6 | 前端重载后重新挂接，且不得回退到无幂等的新写入请求（修正 2.6 中已确认的缺口） |
| A7 | 3.4 的第 1、2 层互斥 |

**退出条件**：对每个写入窗口能说明恢复行为；未知写入不会被自动重跑；旧请求不被当作新任务；被淘汰的 operation 返回明确过期状态而非空事件。

**不宣称**：完成本阶段不等于支持断点续跑。

### 阶段二：受限的真实续跑（文件夹模式优先）

**目标**：重建执行现场，跳过已完成节点，从断点继续。

只对同时满足三条的节点开放：**输入版本明确、产物已持久化、效果可核对。** 不满足的节点明确标注不支持并走阶段一的出口。

依赖：3.1 的收益数据（定粒度）、3.2 的版本规则（定能否复用）、3.3 的权威记录（定谁说了算）、3.4 第 3 层互斥（恢复排他）。

**退出条件**：强制中断后结果等价、不重复输出、保持执行前恢复能力。验收须包含真实子进程退出与重启，不能只抛可被 `except` 捕获的异常。

### 阶段三：按收益逐项扩展

- 规划上下文恢复（依 3.1 对规划阶段的实测结论决定是否做、做到哪一层）
- 备用模型接续
- Office.js 动作级恢复

关于 Office.js：跨重启续跑难度最大，暂缓的判断成立。补一条实施时的提醒——原文档 §9 自列的限制（未保存工作簿无法跨重启可靠识别、图表/透视表缺可靠核对、宿主不可用不得由后端接管）意味着能力表中相当比例的行会是"不支持"。因此这一项应以"逐动作按实测收益开放"推进，先建能力表框架、再逐个填，避免一次性投入到多数行为负结论的表格上。

---

## 5. 仍需注意的点

- **原文档 §12.1 的目录布局需重写**：`%LOCALAPPDATA%/Excel Bro` 有基础（2.1）但不能写成"沿用现有解析"，须说明开发/安装环境统一是新增工作。SQLite 相关部分待 3.3 结论后重写。
- **原文档 §4.2 步骤 3 的"刷新并校验"**：对工作簿本体，该能力当前不存在（2.2）。要么在该步实现 flush/fsync，要么在第 4 章就地声明不具备断电耐久性，不要只在 §12.3 末尾免责。
- **原文档 §8.2 的部分发布策略自相矛盾**：前文说"多文件发布不是整体原子事务""不能一边回滚旧文件一边从后续节点继续"，后文又建议"首期对所有目标状态确定的任务允许继续发布"。部分发布场景下 C09 无法验收，需二选一。
- **原文档 §1 前端重连描述偏乐观**：应改为"仅对非 `ApiRequestError` 的网络类异常重连一次；`ApiRequestError` 即使标记 `retryable` 也不自动重试"（2.6）。
- **原文档 §8.3 / C26 的 manifest 版本表述**：v2 是现行版本而非遗留版本，应改为"v1 为遗留版本，v2 为现行版本且不含任务现场"。
- **事件淘汰窗口未描述**：应在 §1 或 §13 补明，并与 §7.2 的"旧恢复请求返回明确结果或已过期"连起来。

---

## 6. 尚无法证明的保证

不应进入任何对外能力声明：

1. **断电耐久性**。工作簿本体落盘无 flush/fsync；且与 manifest 的耐久处理不对称（2.2）。
2. **多文件部分发布后的自动继续**。见第 5 章 §8.2 矛盾。
3. **Office.js 未知写入的自动接管**。原文档 §7.4 已写"首期不允许"，建议提升为硬约束而非首期限制。
4. **跨 operation 的写入互斥**。当前完全缺失（2.3），在 3.4 落地前不成立。

---

## 7. 待测量与待核实项

| 项 | 说明 |
| --- | --- |
| 规划阶段实际耗时分布 | 2.5 已证明 60 秒不是上界，但真实分布未知。这是 3.1 与阶段三取舍的关键输入 |
| `llm/client.py` 重试参数 | 本版未独立核实重试次数与退避策略；影响单轮耗时上界的具体数值 |
| 中断频率 | 按阶段分别统计，当前无数据 |
| `os.replace` 在目标平台的实际耐久语义 | 需结合平台行为与故障测试，不能一概描述为断电可靠 |
| 强制结束进程与断电行为 | 本轮未做实验 |

---

## 8. 核对方法说明

本版所有代码断言均由直接阅读源文件得出，已核实位置：

- `packaging/runtime.py`：`_bundle_root`（`:13`）、`_serve_frontend`（`:22`）、`_data_root`（`:37`）、`main`（`:43`）
- `server/app/main.py`：`_folder_operations_lock`（`:608`）、`_folder_request_fingerprint`（约 `:611`）、`_run_folder_operation`（`:667`，`run_in_threadpool` 于 `:681`）、`_prune_folder_operations`（`:723`）、`folder_execute`（`:890`）、`folder_execute_stream`（`:907`，锁 `:918`，同 ID 冲突 `:925-927`，容量淘汰 `:932-944`，建任务 `:952-959`，补收 `:962`）、`folder_execute_status`（`:998`）
- `server/app/folder_workbooks.py`：`BACKUP_MANIFEST_VERSION`（约 `:214`）、`_tree_has_reparse_point`（`:550`）、`_atomic_write_manifest`（`:568`）、`_set_manifest_status`（`:587`）、`_save_workbooks_atomically`（`:3386`）、`_register_folder_undo`（`:3739`）、`_restore_folder_undo_impl`（`:3761`）、`_execute_folder_plan`（`:4058`）、`execute_folder_plan`（`:5032`）
- `server/app/excel_agent.py`：messages 构造（约 `:938`）、工具集选择（约 `:949`）、主循环与 `should_stop`（约 `:955-959`）
- `server/app/turn_state.py`：`completion_lock`（`:41`）
- `config/capabilities.json`：2.6 所列限额
- `apps/excel-addin/src/api.ts`：`executeFolderPlan`（约 `:521-646`）

grep 零命中确认：`_undo_stacks`、`UndoCheckpoint`、`undo_folder_operation`、`_record_undo_checkpoint`、`_check_files_in_use`、`asyncio.to_thread`、`run_in_executor`、`threading.Lock`、`msvcrt`、`portalocker`、SQLite 使用。

**规模参考**：`server/app/` 26 文件 / 13,872 行（`folder_workbooks.py` 5,046 行）；`apps/excel-addin/src/` 115 文件 / 36,470 行（`App.tsx` 7,989 行、`excel.ts` 4,510 行）。

---

## 9. 版本记录

- v2，2026-09-08：撤回 v1 关于"取消断点续跑"的结论。修正应用数据目录、fsync 对象区分、规划阶段时间上界、并发互斥现状、恢复入口前置条件五项事实判断。改为采纳三阶段路线（中断识别与安全核对 → 受限真实续跑 → 按收益扩展），并列出实施前必须补齐的四项前置设计。存储选型不在本版决定。
- v1，2026-09-08：见 `REVIEW_TASK_CHECKPOINT_RECOVERY_GUIDE.md`。结论已撤回，文件保留以备追溯。
