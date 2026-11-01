# A1-E3 原生验收包（2026-10-25）

> 这是给**人**照做的清单：每个场景都写清"初始数据 / 步骤 / 预期界面 / 预期单元格 /
> 写入次数怎么取 / 要保存什么 / 失败怎么保留现场"。
> **不重做**已经验收过的普通成功流程（新建计划 → 执行 → 确认写回那条老路）。
> 本包只覆盖三件必须看真机的事：**来源变化后的停止**、**中断恢复只续剩余动作 + 连点一次**、
> **确认/放弃后重开不重执行**。

---

## 0.0 当前脚本入口（2026-10-26 修订，替代第十一轮启动命令）

每次 `prepare` 创建新 run 目录，不覆盖已有工作簿、日志和证据。SDK 运行目录在仓库外
临时目录；配置与账本在 run 内隔离。仅复制已有连接供测试使用，不在证据中导出密钥。
`manifest.json` 含本地实例令牌，不上传、不作为公开证据。

```powershell
# 项目根目录；从别处调用时使用脚本绝对路径。
.\scripts\a1e3_native.ps1 prepare
# 在另一个终端启动所验宿主的开发前端（前台服务，Ctrl+C 结束）。
.\scripts\a1e3_native.ps1 frontend -HostId excel-office
# WPS 改用：.\scripts\a1e3_native.ps1 frontend -HostId wps-et
```

前端同时设置 `VITE_API_BASE_URL`，WPS 另设代理目标。已有端口被占用时明确拒绝，
不自动杀掉占用进程。先按正常开发方式停止原项目 Vite，再启动本轮前端并重新加载窗格。

在生成的 `a1e3-验收-副本.xlsx` 中固定勾选「来源」「目标」，选择可用 SDK 连接，
发一次只读请求，让真实宿主桥向测试实例登记身份。若刚重启后端，WPS 可通过
「管理模型连接」重新读取模型目录后返回；不需要改连接或密钥。

```powershell
.\scripts\a1e3_native.ps1 status
.\scripts\a1e3_native.ps1 arm -HostId excel-office
# WPS 使用 -HostId wps-et。目标不唯一时必须再指定 -Name / -Identity。
# 不带筛选时：真实终端列出候选并等待选择；非交互调用非零退出且不启用。
```

有效期 120 秒内发送：
「必须重新用工作簿工具读取『来源』B2:B4，将每个值乘以2，写入『目标』B2:B4。
只改目标，来源保持不变。如果工具报告来源变化，立即停止，不要重试。」
本场景是故障停止响应链验收；不替代真正修改来源后的检测验收。

```powershell
.\scripts\a1e3_native.ps1 evidence -HostId excel-office
# WPS 使用 -HostId wps-et；需要中止故障时执行 cancel。
.\scripts\a1e3_native.ps1 stop
```

默认使用 `latest.json` 指向的 manifest；旧轮可以显式 `-Manifest <绝对路径>`。
`-Base/-Port` 不一致即拒绝。`stop` 比对 PID、创建时间、可执行文件、完整命令及专用
启动模块；不一致不停止。取消请求明确发送 JSON `{}`。

取证仅认本轮 `attemptId → boundTaskId → consumedTaskId`：test-only ASGI 包装器委托
真实 fault 函数，只记录其真实返回值，不改故障判据；普通 `server.app.main:app` 没有
该审计端点。CLI 校验实例令牌、runId、账本任务身份、宿主和时间，只查询该任务执行记录，
不会导出“最近八条”猜归属。`consumed` 后生产状态会清掉 `taskBound`，不把它当失败。

缺失账本、未绑定/未触发、取消、错实例、接口错误均不能当作“未进入执行”。只有同一
任务已 `failed + SOURCE_CHANGED` 且没有执行登记，才记录「未进入执行」；有执行时
原样保留服务端观测，仅 `complete + zero + 已知单位` 可称零调用。每次导出独立目录，
失败证据标 `insufficient`；CLI 结果不替代 UI/单元格观察，`nativeUiAndWorkbook` 保持 pending。

详情和本轮实际结果见 [工具修复与原生记录](SDK_A1_NATIVE_TOOLING_REPAIR_2026-10-26.md)。

---

## 0. 准备（一次就好）

### 0.1 一份独立测试簿（脚本生成，不动你的正式文件）

```powershell
.\scripts\a1e3_native.ps1 prepare
```

生成 `outputs/a1e3-native/runs/<runId>/a1e3-验收.xlsx`（母本）和同目录 `a1e3-验收-副本.xlsx`（验收打开用），
内容固定：

```
来源!B1 = 金额   B2:B4 = 111 / 250 / 333
目标!B1 = 金额   B2:B4 = -101 / -202 / -303（执行前初值）
汇总             空
```

**每个场景都从 `a1e3-验收-副本.xlsx` 重新打开**（场景之间不要接着用同一份，否则上一次的
写入会混进来）。每场景复制母本为独立新文件；`prepare` 创建新 run，不覆盖旧证据，已有隔离端口占用时会拒绝。

### 0.2 后端（脚本起独立测试服务端，不碰你的 8765）

`prepare` 已经帮你起好了测试服务端（独立端口 8877 + 独立配置/状态目录 + 复制已保存模型连接）。
它带 `EXCEL_BRO_SOURCE_CHANGED_FAULT=1`。**不要**再 `npm run dev:server`——那会占用 8765，
与本脚本隔离测试服务端的设计冲突。

> 开关含义：把"被绑定的那一次任务的下一次工具调用"直接变成一次**真实发生**的
> `source_changed` 终止（回执、停桥、丢弃状态全走生产代码，只是不派发给宿主读表）。
> 不设它，S1 就只能靠"希望模型在来源变化后再读一次"——那不叫验收，所以 S1 只在设了它时做。

### 0.3 Excel 端

```powershell
# 终端 2（关键：让加载项连测试服务端，不是 8765）
.\scripts\a1e3_native.ps1 frontend -HostId excel-office   # https://localhost:3000

# 终端 3
npm run start:excel          # 打开测试副本，加载项只旁加载到它
```

任务窗格里：选好模型 → 数据范围里勾选 `来源`、`目标` → 确认（S1 用不到 `汇总`）。

### 0.4 WPS 端

```powershell
# 关键：让 WPS 的 /api 代理指向测试服务端，不是 8765
.\scripts\a1e3_native.ps1 frontend -HostId wps-et   # http://127.0.0.1:3889
```

**完全重启 WPS**（已开的 WPS 不会热加载新登记）→ 功能区「Excel Bro」→「打开 Excel Bro」→
打开同一份 `a1e3-验收-副本.xlsx` → 勾选 `来源`、`目标` → 确认。

> Excel 端与 WPS 端**不要同时对同一份文件**跑同一个场景。

### 0.5 怎么确认窗格真的连上了测试服务端（不是 8765）

发需求前先看测试服务端日志 本轮 run 目录的服务端日志：窗格里发任何请求，这里都会出现
对应 `GET/POST /api/...` 行。**如果一条都没有、而你确定点了发送，那就是窗格还连在别处**
（多半是 8765 或没起前端），先解决，不要带着"连错服务端"往下跑——那样故障注入会打到错误的实例，
要么不触发、要么污染别的验收。

---

## 1. 场景 S1：来源变化 → 本次规划停止（本批唯一的新行为）

**只在 `EXCEL_BRO_SOURCE_CHANGED_FAULT=1` 时做。**

### 步骤

1. 发需求**之前**，先让服务端"见过这本簿"，再启用一次（120 秒内有效、单次、
   **必须能唯一确定是哪一本**）：

```powershell
# ① 先在面板里发一条只读需求，例如「来源表有多少行」，让它把桥立起来
#    （开桥那一刻服务端才登记宿主 + 名字 + 身份；不先做这一步，seenTargets 是空的）

# ② 问服务端"它见过哪几本簿"（脚本会打印，含 identity，照抄即可）
.\scripts\a1e3_native.ps1 status

# ③ 启用：identity 抄自上一步 seenTargets 里的 workbookIdentity，不要手写。
#    - 只见过一本：.\scripts\a1e3_native.ps1 arm -HostId excel-office
#      （服务端会补全 identitySource=observed）
#    - 同名多本：必须显式给 -Identity（值抄 seenTargets），否则脚本拒绝
.\scripts\a1e3_native.ps1 arm -HostId excel-office [-Identity "<从 seenTargets 抄>"]
```

   预期返回里 `phase` 是 `armed`、`identityBound` 是 `true`、`identitySource` 是
   `declared`（你显式给的）或 `observed`（服务端替你补的）。
   - Excel 的 `hostId` 是 `excel-office`，WPS 是 `wps-et`；
   - 可核验身份的前缀只有四种：`url:`（已保存）/ `unsaved:`（未保存）/
     `wps-path:`、`wps-session:`（WPS）—— **不要手写**，抄 `status` 打印的原文；
   - 服务端一个身份都没见过（没做 ①）或多个同名簿被见过 ⇒ 脚本会**停下**并说明原因。

   > **不能**只传名字、不能传空：`arm` 要求"能唯一确定是哪一本簿"。
   > 这是对的：留空等于把这次启用交给"谁先说话"，并发任务或另一本同名簿会把它抢走，
   > 被抢的那一轮莫名停住，而要验的簿一次都看不到停止（第三批已修，见
   > `SDK_A1_E3_STOP_AND_GAPS_2026-10-25.md` §8.3）。

2. 立刻在面板里发 S1 需求：
   「读取来源表 B2:B4，将每个值乘以 2，写入目标表 B2:B4，来源表保持不变。」
3. 等它结束（几秒到一两分钟）。

> 发完需求后可以用 `.\scripts\a1e3_native.ps1 status` 看它有没有被绑上：
> `phase` 变成 `bound`、`taskBound` 是 `true` = 已经绑给**这一次任务**（开桥那一刻绑的）。
> 如果这一轮模型**一个工具都没调**（直接回答），`phase` 会停在 `bound`（绑过但没命中）——
> 执行 `.\scripts\a1e3_native.ps1 cancel` 再重来，免得它落到 120 秒内的下一次需求上。

### 预期界面

- 对话里有一条系统消息，**原文**是：
  > 来源已变化，当前任务已停止：没有继续读取来源、没有继续计算，也没有写入。请按最新数据重新发起需求。
- 本次执行过程卡以**失败**收尾 —— 不是"已完成"，也不是"暂停 / 等待继续"。
- 输入框上方**没有**继续入口；对话里**没有**「继续分析」按钮。
- 没有计划预览卡、没有「确认变更」。

### 预期单元格

- `汇总` 整张表**一格未动**（全空）。
- `来源` 与验收开始时**完全一致**（111 / 250 / 333）。

### 写入次数怎么取（S1）

**这条不要再数 `actions/start` 的行数** —— 那只是"申请过写"的次数，不是写入次数。
S1 要证明的是"宿主一次都没被写过"，按下面三条**依次**做，不要跳过前两条直接查 `host-writes`：

#### 1.3.1 先确认"这次根本没建执行记录"

S1 的正常结局是**规划期就停**，那时**没有任何 operationId**。所以第一件事是看请求：

- 服务端窗口里**不应出现** `POST /api/office/executions/register`（WPS 是 `/api/wps/...`）；
- 也不应出现任何 `POST .../actions/start`。

只要没有 `register`，就不存在执行记录，也就**不存在"写入次数"这个问题**。这时：

> **不要**拿一个看起来像的 ID（甚至自己编一个）去 `GET .../host-writes`。
> 那个端点对不存在的记录一律返回 **404 `EXECUTION_NOT_FOUND`**，把 404 当"零写入"就是
> 把"查不到"当成"没发生"。正确写法是记"**未进入执行，无 executionId**"。

#### 1.3.2 再看界面终态

- 面板与 SSE 帧给出 `SOURCE_CHANGED`（不是 `MODEL_TURN_LIMIT`）；
- 本次执行过程卡以**失败**收尾，输入框上方没有继续入口。

#### 1.3.3 只有在"真建了执行记录"时才查宿主观测

如果日志里**确实**出现了 `register`（例如停止发生在读取之后、执行之前），直接跑：

```powershell
.\scripts\a1e3_native.ps1 evidence
```

脚本会**自动**从独立账本里发现 executionId，逐条读生产端点 `/api/{office,wps}/executions/{id}/host-writes`，
把服务端算好的 `observation` 原样落进 `outputs/a1e3-native/evidence/evidence.json`，并打印人话结论。
**不要**自己 `Invoke-RestMethod` 拼 `host-writes`、也不要手填 executionId。

**读 `observation`，不要读 `entries` 的长度**：

| `observation.observation` | `observation.callEvidence` | 你能写进验收表的结论 |
| --- | --- | --- |
| `not_collected` | `none` | **未采集**（上报一条都没到）—— 不能写 0 |
| `partial` | `unknown` | 覆盖不全 ⇒ **未知** |
| `partial` / `complete` | `nonzero` | **确实向宿主发过写调用**（覆盖不全时再加"还缺 N 个动作") |
| `complete` | `zero` | **明确零写调用**（唯一允许写"0 次"的一格） |
| `complete` | `unknown` | 有动作没拿到确定结果 ⇒ **未知** |

同时记下 `observation.unit`：Excel 是 `excel_sync_batches`（**刷新批次**），WPS 是
`wps_api_calls`（**写 API 调用**）。**两端的数字不能互相比较，也不能相加。**

想让人话，直接用 `packages/core/src/hostWriteAudit.ts` 的
`describeHostWriteObservation(observation).label`，界面与验收记录读同一句，不要各写一句。

#### 1.3.4 辅助观察（仍然有用，但不是判据）

- 从 arm 到结束**不应出现** `POST .../actions/start`：出现说明有人申请过写；
- `GET .../executions/<operationId>`（有记录时）的 `nextActionIndex` 应为 0。

### 要保存（S1）

1. 面板截图（要能看见那条系统消息与失败态）；
2. PowerShell 从 arm 到结束的整段输出（直接全选复制到 txt；这是"有没有 register / start"的唯一凭据）；
3. `S1-<端>-host-writes.json`（有 executionId 才存；没有就写"未进入执行，无 executionId"）；
4. 工作簿另存为 `S1-<Excel|WPS>-结果.xlsx`。

### 失败（不符合预期）时怎么保留现场

不要撤销、不要重发、不要关面板：先截图，再复制 PowerShell 输出，然后才关闭。

### 这条验收**证明**什么 / **不**证明什么

**两种验收分别记录，不许合并成一句结论。**

| 记录项 | 怎么验 | 能证明 |
| --- | --- | --- |
| **A. 停止响应链**（本场景 S1） | 用 `/arm` 注入确定的 `source_changed` | 守卫一旦给出 `source_changed`，本次规划**真的停**、旧继续入口**真的消失**、宿主**一次没被写过** |
| **B. 真实来源变化被检测** | §4 的手改数据（不可靠触发） | 真实指纹不符也能判出来（要求模型在改数之后再读一次） |

S1 属于 **A**。它能证明的是"停止这件事本身是可靠的"，**不能**代替 B：
`source_changed` 在真实宿主上的两个产生点都要求「模型在来源变化之后又读了一次」，
那一步是否发生不由我们控制。所以：

- S1 通过**不等于**"真实来源变化检测通过"；
- S1 通过**不等于**"没有误报"（那是 A1-E2 已验的另一条路：写前复核会给出
  「数据来源在预览后发生了变化」而不是停止）；
- 只有在 §4 真的看到了相同停止消息，才能在这一行记 B 的证据；没看到就写"未触发"，不写"未通过"。

---

## 2. 场景 S2：精确中断 → 重开 → 连点恢复

当前实测见 [S2 双端记录](SDK_A1_S2_NATIVE_RECOVERY_2026-10-26.md)。不要用“第二步刚出现就关窗”抢时序，也不要假定一个区域复制一定被模型拆成三个动作。

1. 用独立已保存测试副本，固定勾选来源、目标；来源 B2:B4 为 111/250/333，目标为 -101/-202/-303。S1 故障保持未启用。
2. 展开「恢复验收（开发专用）」并启用下一次第二步中断。
3. 发需求：「按顺序执行三个独立动作：把来源 B2 的值复制到目标 B2，再把来源 B3 的值复制到目标 B3，再把来源 B4 的值复制到目标 B4。只复制值，不修改来源，不合并动作。执行后供我审阅。」核实计划确为三个动作。
4. 中断后应是目标 111/250/-303；保存截图、工作簿与账本/host-writes。第二步已落表但完成回执未交付，第三步未执行。
5. 关闭重开窗格。单独记录是否真的重载 WebView：WPS 隐藏再打开可能复用页面，不算进程重启；Excel 本轮还执行了 WebView 刷新。重载后若选择丢失，只恢复原两表授权。
6. Excel 快速连点「核对并继续」。WPS 先点「核对在途动作」，核对成功后快速连点「继续剩余动作」。
7. 目标应为 111/250/333，来源不变；底部接回双视图、放弃、确认。打开双视图核实执行前值，然后确认并保存。

四档证据分别保存：请求日志、认领/账本、宿主调用观测、真实 UI 与保存后单元格。不要用最终值或请求次数单独证明无重放。

- 同一 operation 的 actions/start 共 3 次，恢复后只有第三步一次；actions/reconcile 一次，最终 complete 一次。WPS 另有 resume 一次。
- 从本轮 8877 取得该 operation 的 `/api/{office,wps}/executions/<operationId>/host-writes`，与该任务账本身份对应。S1 的故障 evidence 导出命令不能冒充 S2 取证。
- 两端 observation 必须 complete/nonzero，动作索引 0/1/2 各有记录。Excel 单位 excel_sync_batches，WPS 单位 wps_api_calls，不混算。
- 本轮 Excel 每动作 hostWriteCalls=2、mutatingCalls=1；WPS 每动作各为 1。结合执行窗口和日志核实前缀未重写。
- WPS 第 2 步故障可能保留 failed_after_write/not_applicable；核对只补账不改写原始异常观测，不能强求全部 written/verified。已完成账本与最终回读另列证据。
- 确认前后 host-writes 应完全一致；确认只清审阅标记。撤销/放弃写调用不在计划调用计数中，不能用计数不变证明它没有写。

每端保存 interrupted/confirmed.xlsx、interrupted/resumed/confirmed-host-writes.json（已取得部分）、operation-http.log、账本摘要及中断/恢复/对比/确认截图。未知或缺失记录按未知，不能按 0。

---

## 3. 场景 S3：确认 / 放弃后重开，不重执行

### 步骤（确认）

1. 发：「在目标表 A1 写入 完成」（单动作，很快）。
2. 审阅条出现后点「确认变更」。
3. 关闭并重开任务窗格。

预期界面：对话里保留「已确认变更」；**没有**「确认变更 / 放弃变更」按钮。
预期单元格：`目标!A1` 仍是「完成」。
写入次数：`actions/start` 仍然只有 **1** 次；**并且**宿主观测里 `entries` 仍是 1 条
（`actionIndex = 0`、`outcome = "written"`）。"确认变更"只是接受结果并清标记，
**不得**新增 `hostWriteCalls`；新增了就是这个动作被重执行了。

> 读法：`observation.observation` 应为 `complete`，`reportedActions` 与 `totals` 与
> 确认**之前**那次完全一致（确认前先存一份 json，确认后再取一份，两份逐字段比）。
> 撤销/放弃是**另一件**操作：它会真的调用宿主恢复写入，但**不计进计划写调用** ——
> 所以"放弃后 `hostWriteCalls` 不变"是正常的，不要把撤销的写当成重放（判据见
> `SDK_A1_E3_STOP_AND_GAPS_2026-10-25.md` §8.2.2）。

### 步骤（放弃）

4. 从副本重新打开文件，发同一条需求。
5. 审阅条出现后点「放弃变更」。
6. 关闭并重开任务窗格。

预期界面：不再有审阅按钮；对话里是「已放弃并恢复」。
预期单元格：`目标!A1` 回到执行前的空值。

---

## 4. 可选（不可靠）：真实指纹路径

想在真机上顺便看一眼真实 `source_changed` 可以这样做：面板出现读取进展后**立刻**把
`来源!B2` 从 111 改成 250，然后等。

- 模型确实又读了一次 → 你会看到与 S1 **相同**的停止消息；
- 模型没有再次读取 → 你会看到「数据来源在预览后发生了变化」（写前复核拦下，A1-E2 已验收）。

两种都不算缺陷，但**不能拿这条代替 S1**：它不可确定性触发。

---

## 5. 记录表（照这个填，不要写推测）

| 场景 | 端 | 界面符合 | 单元格符合 | start 请求次数 | 观测覆盖 `observation` | 调用数 + 单位 | 证据文件名 | 不符合的具体文字或值 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| S1 | Excel | | | | | | | |
| S1 | WPS | | | | | | | |
| S2 | Excel | 恢复/对比/确认正常 | 111/250/333 | 3 | complete/nonzero | 6 excel_sync_batches | 本轮 s2-excel-r2 | WebView 实际刷新；过程警告/列宽见 S2 记录 |
| S2 | WPS | 恢复/对比/确认正常 | 111/250/333 | 3 | complete/nonzero | 3 wps_api_calls | 本轮 s2-wps | 窗格复用；完整 WebView 重载仍待验 |
| S3 | Excel | 用户截图：确认/放弃后重开符合 | 确认A1完成；放弃A1空，B列负值保留 | 未采集 | 未采集 | 未核实 | 用户对话截图；见末尾补记 | 无调用审计，不推断无重写 |
| S3 | WPS | 用户确认已重开；两路径符合 | 确认A1完成；放弃A1空，B列负值保留 | 未单独采集 | complete/nonzero（两轮各自） | 每轮3 wps_api_calls，其中mutatingCalls=1 | s3-reopen-20261027；ledger-diagnosis-20261027/s3-live-host-writes.json | 真实账本正常；完整进程重启未证实 |

**"观测覆盖"一列填这几个词之一**（与 `observation.observation` / `callEvidence` 一一对应）：

| 填什么 | 对应 | 含义 |
| --- | --- | --- |
| `未采集` | `not_collected` + `none` | 一条上报都没到（**不是 0**） |
| `部分` | `partial` + `unknown` | 覆盖不全 ⇒ 未知 |
| `部分+非零` | `partial` + `nonzero` | 已经看到调用，但还缺动作 |
| `完整+零` | `complete` + `zero` | **唯一允许写"0 次"的一格** |
| `完整+未知` | `complete` + `unknown` | 有动作没拿到确定结果 |

**"调用数 + 单位"一列必须带单位**：`2 excel_sync_batches` / `4 wps_api_calls`。
只写 "2" 下次没人知道那是什么。没采到就写 `未知` 或 `未采集`，**不要留空、不要写 0**。

> 填表纪律：**"没有记录"要写"未采集/无记录"，不能写 0**。一列里只有数字而不写来源，
> 下次没人能判断那是"确实 0 次"还是"没测"。

### 5.1 每一档证据能证明什么（不许越级）

| 你手上有的 | 能写下的结论 | 不能写 |
| --- | --- | --- |
| 服务端日志里的请求行 | "有人申请过写第 N 个动作"（`register` / `actions/start` 各几次） | "写了几次" |
| 认领结果（200 / 409 / 422） | "谁拿到了写这次动作的权利" | "真的写了" |
| `entries` 里某动作 `outcome=written` + `readback=verified` | "这个动作的宿主调用已返回、回读等于计划值" | "整份计划都对"（要看全部动作） |
| `observation = complete + zero` | "边界内每个动作都有观测，且都没调过宿主写 API" | "宿主绝对没变"（别的进程也可能改） |
| `observation = not_collected / partial / unknown` | **未知** | 任何形式的"0 次" |
| 单元格截图 / 另存的工作簿 | "最终值是这样" | "中途没写过又改回来" |

### 5.2 故障注入与真实来源变化，分开记

- S1 用 `/arm` 注入：证明的是**停止响应链**（停桥、旧入口消失、没有向宿主发起写调用——
  这一句要按 `observation` 判读，能写"零"的只有 `complete + zero` 那一格）；
- §4 手改数据：才可能证明**真实来源变化被检测**，但不可靠触发，没看到就写"未触发"。

两件事各写一行，不要用一条"通过"把它们合并。

证据放 `outputs/a1e3-native/<场景>-<端>/`。

---

## 6. 需要新增什么开发入口

**需要 1 个，且已实现**（默认关闭）：

| 入口 | 开关 | 作用 |
| --- | --- | --- |
| `POST /api/dev/source-changed-fault/arm`（+ `/cancel`、`GET`） | `EXCEL_BRO_SOURCE_CHANGED_FAULT=1` | 让**被绑定的那一次任务**的下一次工具调用直接得到生产那份 `source_changed` 并走生产停止链路 |

Body **必须**带 `hostId` + `workbookName`，`workbookIdentity` 强烈建议显式给：任务归属由服务端在
**开桥那一刻**绑定，之后只认那一个 `taskId`。启用时的目标也必须**能唯一确定是哪一本簿**：

| 你传的 | 服务端行为 |
| --- | --- |
| `workbookIdentity` 是 `url:` / `unsaved:` / `wps-path:` / `wps-session:` | 用它（`identitySource = "declared"`） |
| 不传，但这个宿主 + 名字下**只见过一个**可核验身份 | 用那个（`identitySource = "observed"`，有证据才补） |
| 不传，且**一个都没见过** | **拒绝**：`缺少可核验的工作簿身份…`（先发一条只读需求，再看 `seenTargets`） |
| 不传，且同名下见过**多个**身份 | **拒绝**：`工作簿名「…」在服务端见过 N 个不同的身份…`（必须显式指定） |
| 传 `name:xxx` 或任何非上述前缀 | **拒绝**：`给定的工作簿身份不在可核验范围…` |

留空 Body / 只给名字被拒是**有意的**：那等于把这次启用交给"谁先说话"，并发任务或另一本
同名簿会把它抢走，被抢的那一轮莫名停住，而要验的簿一次都看不到停止（第三批已修，见
`SDK_A1_E3_STOP_AND_GAPS_2026-10-25.md` §8.3）。

`GET` 同一个 URL 会返回 `seenTargets`（最近见过的工作簿，含**该抄的身份字符串**）、
`identityBound`、`identitySource`。别猜路径格式，从这里抄。

没有它，S1 在真实宿主上无法确定性触发（模型是否再读一次不可控）。开关关闭时它是 404，
生产行为与原来逐字节一致（`test_dev_fault_is_off_by_default` 钉住）。

---

## 7. 这一包**不**覆盖

- 两个窗格**同时**恢复（并发认领）—— 需要两个窗格，且服务端真并发证据已有
  （`server/tests/test_recovery_claim_concurrency.py`），原生双窗格另立一包；
- 真实撤销、真实删表恢复（`worksheet_restore` 的 R1–R8）；
- **未保存同名工作簿**：身份已改成"会话内对象身份"（自动化已验），但真实窗格里
  "各开一个、都叫「工作簿1」"还要人看一眼。想验就让两个未保存簿同名，各开一个任务窗格，
  分别在两个窗格里问同一句「来源表有多少行」，**看读到的是不是自己那份**。
- 并排审阅布局比例（35/35/30）的截图验收。
- **两个开发故障入口的"同名簿互抢"**：自动化已验（`test_source_changed_stop.py` 与
  变异 H7–H9），但"两本同名簿 + 两个任务窗格"的原生场景不在本包内 —— 它需要同时开两本簿，
  与本包"一次只开一份副本"的前提冲突。
- **观测覆盖判据的实机表现**：`complete` / `partial` / `not_collected` 这三种结果在真实宿主上
  各出现一次，才算真正看过——本包只要求你**如实记录是哪一种**，不要求凑齐三种。

### 7.1 本包**明确不要**做的事

- 不要为了"拿到一个数字"去编一个 `executionId` 或复用一个旧 ID（不存在的记录一律 404，
  把 404 读成"零写入"就是本次修正要消灭的那种错误）；
- 不要把 Excel 的 `excel_sync_batches` 和 WPS 的 `wps_api_calls` 比大小或相加；
- 不要把 S1（故障注入）的通过写成"真实来源变化检测通过"；
- 不要在验收过程中改服务端代码或放宽任何断言；不符合预期就把现场留下来。


### 2026-10-27：用户完成 S3 确认 / 放弃后的窗格重开复验

用户提交六张真实 WPS 截图，并明确回复「我已重開」。确认路径：目标!A1 保留「完成」，黄色审阅标记移除，重开后显示「恢复点已不可用」；这是会话内撤销的既定边界。放弃路径：A1 恢复为空，B2:B4 保持 -101/-202/-303，重开后显示已撤销并恢复，没有本轮待恢复入口。S3 的这两条 WPS 可见行为已有原生证据；用户先前 Excel 截图也显示确认后保留「完成」、放弃后恢复为空。

证据归档：outputs/a1e3-native/s3-reopen-20261027/（六张 WPS 截图与 evidence.json）。本次重开由用户操作；不据此认定 WPS 进程或 WebView 已销毁重建。

重复写入审计未闭环：默认本机账本 C:/Users/jia/AppData/Local/Excel Bro/tasks/state.sqlite3 使用 mode=ro 查询 office_executions 时返回 sqlite3.DatabaseError: database disk image is malformed；四个隔离验收账本均未匹配 S3 执行。未修改、重建或删除任何账本，也未操作正在运行的后端。单元格值相同不能证明未重复写入，调用次数与 observation 均记为未采集。后续须先定位本次服务实际账本，并对数据库损坏单独诊断，保存原文件及 WAL 后再决定修复。

A1-GATE 保持未通过，A2 未启动。下一步优先闭环账本可读性与 S3 不重复写入证据，再补 WPS 完整重载和原生故障矩阵。此轮只归档证据及更新文档，未改生产代码、未提交或推送 Git。


### 2026-10-27：S3 执行账本排查（旧结论已撤回，见后续更正）

已通过 8765 后端 PID 26288 的打开文件列表确认默认账本即实际使用账本。稳定复制三文件后，副本合并 WAL 查询报损坏；忽略 WAL 的旧主文件 quick_check=ok，但只有截至 10 月 12 日的记录。主文件只有 44 页，校验和全部通过的 WAL 末次提交要求 563 页，基础页仍缺 505 页。WAL 中包含 WPS S3 簿名，不可直接删除或回退旧主文件。旧文件替换/截断等成因未证实。

证据与恢复边界：outputs/ledger-diagnosis-20261027/REPORT.md、wal-analysis.json、snapshot-104703/diagnosis.json。未停止服务或修改原账本。S3 可见行为符合仍保留；不重复写入审计未闭环。后续优先查找同代完整备份，再决定副本抢救或健康隔离账本复采。A1-GATE 未通过，A2 未启动。


### 2026-10-27：账本损坏结论撤回，真实账本及 S3 WPS 审计通过

Codex MSIX 将默认账本读取重定向到自己的 LocalCache 旧主文件，前轮把旧主文件和真实 WAL 混读，误判生产损坏。通过真实后端文件句柄取得的主文件为 2289664 字节，副本合并 WAL 后 integrity_check=ok、foreign_key_check 空；130 条执行、129 条规划、68 条宿主写观测。原数据库无需修复；未改写、重建或删 WAL。后端短暂停住用于保持句柄位置，已恢复，真实审计接口 GET 随后成功；撤回暂停新写入建议。

WPS S3 确认 c4244afe… 与放弃 d1545de6… 两轮均 completed，动作推进 1/1；每轮 hostWriteCalls=3 wps_api_calls、mutatingCalls=1、readback=verified。服务端均返回 complete/nonzero、无缺失/越界观测；记录内未见确认/重开导致重复计划内容写入。放弃/审阅辅助写不属于该计数。Excel S3 此轮未匹配到记录，不借用 WPS 计数。

最终证据：outputs/ledger-diagnosis-20261027/REPORT.md、handle-snapshot-111950/verification.json、s3-live-host-writes.json。旧505页/损坏结论及相应阻塞项已被本条取代；旧错误样本保留且标记 superseded。A1-GATE 仍未通过，A2 未启动；原生矩阵与完整重载等缺口不因本条被标完。


### 2026-10-28：WPS「中断 → 核对补记 → 只续第三步 → 确认变更」原生路径只读核验

用户报告：原执行 `wps-execution-1fd938fd-b20f-4042-9f3b-66b10f5bffdc`（`WPS恢复验收.xlsx`，
目标表 `B2/B3/B4`）经**重新发现 → 第二步补记 → 第三步续跑**后结果为 `111/222/333`，并已点击
「确认变更」。本轮**只读**核对该执行的账本、动作观测与宿主写入计数，**未**操作工作簿、账本、
WAL、恢复材料，未重启前后端。取证方式：复制默认账本三文件到临时目录后只读查询，原文件
sha256 前后一致（`state.sqlite3 eeb62c2c…66a8`、`-wal f793330f…6283`）。

**账本终态**：`status=completed`、`next_action_index=3`、`in_flight_action_index=NULL`、
`revision=11`、`updated_at=2026-10-28 15:49:06`、`result={actionCount:3, backfilled:false}`。
授权集合 `["目标表","来源表"]`（与收窄后的两张一致）。`checkpoint_json` = `目标表!B4`
`before=v2:19:6b546b99` → `after=v2:18:fbf8e469`。

**指纹重算（独立于数值口述）**：按 `matrixFingerprintV2`（FNV-1a）对本机实值重算，六个指纹
全部命中——写前 `B2=-101`、`B3=-202`、`B4=-303`，写后 `111` / `222` / `333`。也就是说登记的
计划与回读的检查点都指这三格，**不是**拿最终数值反推的。

**三条断言的证据与缺口**：

| 断言 | 支持证据 | 明确缺失 |
| --- | --- | --- |
| 第二步核对只补记 | ①补记路径 `reconcileWpsActionReceipt` 只有一个 POST，全链路无宿主写入调用点（`api.ts:1319`）；②`office_host_writes` 中 `actionIndex=1` **只有 1 行**，`recorded_at=10-27 22:44:39`（原 `failed_after_write` 那一刻），15:49 无任何新行；③终态 `next=2`、`in_flight=NULL` 与"补记 applied"一致 | 补记时写入的 B3 检查点已被第三步覆盖（`checkpoint_json` 单槽），账本不再保留 `v2:18:39ce3816` 那一条；补记发生的**具体时刻**没有记录（`updated_at` 只留最后一次） |
| 续跑只执行第三步 | ①`office_host_writes` 仅 `actionIndex=2` 在 **10-28 15:49:06** 新增 1 行：`hostWriteCalls=1`、`mutatingCalls=1`、`outcome=written`、`readback=verified`；②`actionIndex=0/1` 在该时刻之后无新行；③B4 检查点 before/after 指纹可重算命中 | 计数是**端上自报**（不是宿主生成的审计日志），端上少报一次写入账本看不出来；`payload_json` 无逐格写入内容，只能证明"一次变更调用 + 回读指纹相符" |
| 确认变更没有再次执行写入 | ①观测表最后一行是 15:49:06，此后全表无新行；②`updated_at=15:49:06`、`revision=11` 未再推进 ⇒ 确认没走 `finish`/`reconcile`/`claim` 任何会 bump 的路径；③同一 `workbook_identity` 下**没有新建执行**（最后一条 `created_at=10-27 22:44:38`）；④确认路径 = `validateResult`（只读 `inspectWpsUndoRecord`）+ 清审阅标记，不重放 `writeValues`（`WpsAgentPanel.tsx:6617`） | 「用户点了确认变更」本身**无服务端记录**（纯前端/宿主动作），只有用户口述；清审阅标记属于宿主格式写入，**不在**该计数表内（与 2026-10-27 S3 条同一口径）；无 WPS 侧独立日志或录屏 |

**revision 算术（推断，不是证据）**：按最小必要链（register 1 → begin0 2 → complete0 3 →
begin1 4 → suspend1 5 → claim 6 → reconcile1 7 → begin2 8 → complete2 9 → finish 10）应得 10，
实际 11，多 1 步。与终态相容的单步推进只有"再一次 `claim_office_resume`"（换 `owner_token`，
不触碰宿主）；`begin` 必跟 `complete`/`suspend`（+2）、`reconcile` 的前置 `next=in_flight=actionIndex`
已不成立、`finish` 只允许 `running→completed`。**账本没有逐次修订历史**，所以这条只说明"账本
与该推断相容"，**不能**当成"没有第二次写入"的独立证据。

**顺带发现（只读观察，未处理）**：
- 工作簿磁盘态 `B2/B3/B4` 仍是 `111/222/-303`（文件最后保存 10-27 22:46，宿主 `~$` 锁文件在），
  `333` 只在宿主内存里 ⇒「完成」不等于「已保存」；用户未保存前，磁盘不能作为任何判据。
- 该执行的持久恢复材料 `execution-recovery/wps-execution-1fd938fd…/material.dat`（1726 B，
  sha256 `a00cee38…122d0`）仍在盘上，到期 `2026-11-03`。完成路径没有删除它，本轮也未删。

证据归档：`outputs/a1e3-native/s2-resume-20261028/`（`README.md`、`execution.json`、
`host-writes.json`、`verification.json`）。**A1-GATE 未通过，A2 未启动**：本条只闭环
"WPS 第二步中断后只续第三步且补记不重写"这一项，Excel 同类路径与"确认后重开"仍未验。

### 2026-10-28：WPS「续跑后放弃变更」原生验收 + 两个收尾缺陷修复

同一轮的第二个真机场景：**同会话内** 第二步中断 → 核对在途动作 → 继续第三步 → **放弃变更**。
详细记录见 `ABANDON_AFTER_RESUME_ACCEPTANCE_2026-10-28.md` §10–§14，
证据目录 `outputs/a1e3-native/abandon-20261028/`。

**取证方式已改进**：本轮不再"复制账本三文件"，改用 `sqlite3` **一致只读快照**
（`file:…?mode=ro` + `backup(out, pages=512)`）—— WAL 模式下分次复制主库/WAL/SHM
会得到不一致视图。快照 `integrity_check=ok`、`user_version=11`，源文件 sha256 前后一致。

**两次尝试必须分开**：`wps-execution-9e152fde…`（`WPS放弃验收-20261028.xlsx`，本轮放弃）
与 `wps-execution-1fa82782…`（`…- 副本.xlsx`，误点「确认变更」，**只能算确认路径**）。
识别按**完整工作簿路径 + task_id + 执行记录**，不按"最新一条"。

**关键事实**：计划是**三个独立 `writeValues`**（`目标表!B2/B3/B4`，常量 111/222/333，
`resumeContext.sources=[]`），**不是来源复制** —— 素材里的 `copyRange` 只是任务文本。
中断点由 `office_host_writes` 钉住：`idx1 = failed_after_write / WPS_ACTION_FAILED @18:05:48.681`；
续跑只新增 `idx2 = written / verified @18:05:56.658`，`idx0/idx1` 在 18:05:56 无新行。

**放弃的可得证据与缺口**：材料目录已被收走、账本 `updated_at` 停在 `18:05:56.645`/`revision=11`、
磁盘两个 `.xlsx` 字节相同且 mtime 仍是 `16:24:44`（整轮从未落盘）⇒ 单元格最终值只能依赖用户截图。
**放弃无服务端记录**，删除前的材料内容也没有留存 —— 这两项一律写"缺失"，不用终态反推。
**本轮不是跨进程恢复测试**：中断后没有完整退出 WPS。

**同日修掉的两个真机收尾缺陷**（详见同文档 §11）：
1. `runWpsRecoveryAction` 的三条终态分支没有收尾 `pendingPlan` ⇒ 恢复完成后新任务被
   「请先确认或取消当前写入计划」永久挡住（用户靠「＋ → 放弃计划并新建」绕行）。
   修法：执行开始时记归属绑定，三条终态统一 `settleRecoveredPlan`，归属判据收敛到
   `@excel-bro/core` 的 `planSettlesExecution`（执行身份 + 计划按引用 + 消息 id，fail-closed）。
2. 撤销说明「本次曾在 … 改动 2 / 3 个单元格」把**中断那一刻**的阶段统计说成整次执行，
   且分子里混进了"回读失败被当成改成空"的假改动。修法：`changeReviewUndoFacts` 拆成
   已回读确认 / 未通过回读核对 / 未执行三类，文案标注「这次执行中断时」。

**回归（先红后绿，皆有反向变异）**：3 条面板级面板 + 5 条 core 归属判据 + 3 条撤销说明；
删掉收尾调用 ⇒ 新任务用例真机同款报错；去掉引用比较 ⇒ 归属用例红；让 `readback_failed`
累加改动数 ⇒ 撤销说明用例红。

**A1-GATE 仍未通过，A2 未启动**：Excel 一端的 #9/#10/#12 全是 `⏳`；
真实撤销、跨进程版"续跑后放弃"、"确认后重开"的单元格值（只有截图）都还空着。

### 2026-10-28 收尾轮：确认 / 放弃之后能不能继续对话（同会话）

前一条修完 `pendingPlan` 不收尾之后，要验的只有一句话：**收尾之后再发消息，闸门放不放行**。
做法是在同一窗格、同一对话里直接发一条只读查询。

两次真机各一对任务（写入 + 紧随其后的只读查询），账本逐条可核实
（证据目录 `outputs/a1e3-native/abandon-20261028-followup/`，快照同样走 `backup` 一致视图）：

| | 确认轮 | 放弃轮 |
| --- | --- | --- |
| 执行 | `wps-execution-7afd3dc6…f9c1e050`（19:38:53） | `wps-execution-5a2d84be…f7e1e1`（19:42:36） |
| 宿主写入 | idx0/1/2 **同一毫秒** `written/verified`，`revision=8` | idx1 `failed_after_write` @19:42:36.585 → idx2 `written/verified` @19:42:46.350，`revision=11` |
| 恢复材料 | 仍在（1658 B） | 已回收（唯一删除点在撤销成功分支） |
| 随后查询 | `turn-4ec317…` `completed` / `answer`；`complete=false`、`sourceSheets=[目标表,来源表]` | `turn-68fd1b…` `completed` / `answer`；`complete=true`、`sourceSheets=[目标表]` |
| 用户看到 | `111/222/333` | `-101/-202/-303`，撤销摘要标注"执行中断时" |

**记账口径（本轮最重要的结论）**：

- **放弃之后可以继续对话 = 真机通过**（中断 + 续跑 + 放弃，随后查询正常出结果）；
- **确认之后可以继续对话 = 真机通过，但只覆盖"普通确认执行"**：`7afd3dc6` 在账本里
  **没有中断记录**（三步同一毫秒、`revision=8`、无 `failed_after_write`），
  按 `1fd938fd` / `9e152fde` / `1fa82782` / `5a2d84be` 四条已确认的形状反推，它是一次跑完的普通执行。
  `settleRecoveredPlan` 只在三条**恢复终态分支**上调用，普通执行走不到 ⇒ 这一轮不能用来证明该修复。
- **"续跑后确认 + 随后对话"仍缺**：前半有 `1fd938fd`（15:49 续跑 + 确认 + 保存重开），
  后半只有 `WpsAgentPanel.test.tsx` 的自动化用例。不把两半拼成一句"通过"。

**"没被挡下"的取证方式**：闸门在端上，账本没有字段。可核实的是三条同时成立 ——
①两条 read 的 `planning_tasks` 都 `status=completed`、`resultKind=answer` 且有 `resultRef`；
②`planning_task_events` 里各只有一次 `query_table`，没有混进写入类动作；
③此后该工作簿没有再产生新的 `office_executions`。若闸门当时仍拦着，①不会成立。

**保留**：`1fd938fd`（`…0f5bffdc`）"确认后保存重开"的通过记录未被本轮覆盖：
`status=completed`、`revision=11`、`updated_at=15:49:06.522`、材料 1726 B 仍在。

**同日记录的未修复项**：确认后那次查询的结果卡带出
「结果不完整（扫描已达上限），仅展示已取回部分。」，而放弃后那次没有；两次的差别只有
`complete` 与 `sourceSheets`。已定位到 `dataTools.ts` 的 `complete = warnings.length === 0`
与「未找到字段」告警、以及 `LocalResultCard.tsx:67` 只写了"扫描已达上限"这一种原因，
但 `query_table` 的实际参数不落账，**根因未证实**。作为独立待查项写入
`A1_GATE_REMAINING_2026-10-28.md` §4，不因另一轮未复现而关闭。

**A1-GATE 仍未通过，A2 未启动**。剩余项按三分类（尚未实现 / 自动化已验但真机未验 / 证据待补）
列在 `A1_GATE_REMAINING_2026-10-28.md`；最值得做的下一项是
**WPS 跨进程「中断 → 完整退出 → 重开 → 续跑 → 放弃」**（一次操作补 #11 与 #12 两格）。

### 2026-10-28 夜间：WPS 跨进程（**正常退出重开**）「中断 → 重开 → 续跑 → 放弃」

上一小节末尾把"跨进程版"列为下一项，本节即其结果。**注意边界**：这是用户**手动完整退出 WPS**
后再打开的跨进程恢复，**不是**杀进程 / 崩溃恢复 —— 两者代码路径相同但触发器不同，验收分档不能合并。

**精确身份**（不按"最新记录"挑）

| 项 | 值 |
|---|---|
| 执行 ID | `wps-execution-3f00ea0d-6a93-4b44-87d8-0779df926ef1` |
| 任务 ID | `task-da15be01-9682-49a9-8327-a51f32b40493` |
| turn ID | `turn-6352251f00b04fae9b5e33ddce294b44`（23:25:42 → 23:25:49） |
| 工作簿身份 | `wps-path:e:/codex/excel-bro/outputs/wps-abandon-20261028/wps放弃验收-20261028.xlsx` |

**用户实际做了什么**：第二步中断 → 保存成 `111 / 222 / -303` → **完整退出 WPS** → 重开同一文件 →
首次显示「证据不足」→ 点「重新核对（只读）」后恢复正常 → 核对第二步 → 续跑至 `111 / 222 / 333` →
放弃变更 → 恢复 `-101 / -202 / -303` → 保存重开，三个负数保持，未再出现恢复卡。

**账本侧（一致快照，`mode=ro` + `backup API`）**

| 断言 | 结果 | 依据 |
|---|---|---|
| 只一条执行、属于本任务与本簿 | ✅ | `task_id` + 完整路径 identity |
| 第二步 `failed_after_write` | ✅ | idx0 / idx1 同落 `23:25:50.403`，`WPS_ACTION_FAILED`、`readback=not_applicable` |
| 第三步与第一步相隔 816 秒（**不证明进程更换**，只排除同一次连跑） | ✅ | idx2 `written/verified @23:39:26.371` vs idx0 `23:25:50.403` |
| 前两步没有重复写入的**回执**（账本层，非宿主审计） | ✅ | `office_host_writes` 共 3 行，0/1/2 各 1 行；`backfilled=false` |
| 终态 `completed` / `next=3` / 无在途 / 未 dismiss | ✅ | `execution.json` |
| 恢复材料已清退 | ✅ | 目标目录不存在，`execution-recovery` 根 mtime `23:45:58` |

**磁盘侧（本轮第一次拿到"保存后"的值）**

- `WPS放弃验收-20261028.xlsx`：11022 B / `db31f17b83ee84b5…da1748` / mtime `23:56:41.911`
- 目标表 `B2:B4 = -101 / -202 / -303`；来源表 `B2:B4 = 111 / 222 / 333`；未授权表 `B2 = 987654`
- 与原始副本（8121 B / `6442c0ce…` / `16:24:44.718`）逐格比对 **0 处差异**（值 / 数字格式 / 字体），
  **条件格式规则 0 条**、填充色一致 ⇒ 写回负值与**审阅标记清理**都落到了磁盘。

**四档证据分开记（不许越级）**

1. 账本证据 —— 动作序号、状态、耗时间隔、材料是否在；
2. 磁盘取值 —— 保存后的真实单元格值 / 条件格式 / 填充；
3. 截图 —— 界面**当时那样显示**，不证次数与时序（8 张，来自 `%TEMP%\QQ_*.png`；文件 mtime 是落盘时间，
   复制/移动会改它，**不等于拍摄时刻**，只能当"不晚于"的近似）；
4. 用户操作反馈 —— "完整退出 / 重开 / 点了重新核对"属口述，账本无痕。

**账本不能证明的 6 条**：不记进程；「证据不足」无服务端记录、端上 `localInspection` 取值不落盘；
「重新核对（只读）」是纯读、无痕；`hostWriteCalls` 是端上自报而非宿主审计；`before_snapshot` 已清空；
材料删除前的内容未留存。

**记入待查项（未修复）**：首次重开显示「证据不足」，点「重新核对（只读）」后即恢复。
该文案唯一对应 `wpsExecutionReconcile.ts:75`（`localInspection` 为 `null`/`unknown`），据此可排除另外 6 支；
但 `unknown` 本身有 6 条候选原因（授权抛错 / 描述不符 / 表未授权 / 形状不符 / `Value2`·`Formula` 为
undefined / 格式快照不全，另有面板 `try` 吞错），**每条都无证据**。因此只记现象，
**不归因于加载时机、授权或材料丢失**。要闭这条需要 DEV 原因码 + 两种复现条件（首次重开必现、
二次必不复现）先坐实；拿到原因码之前不动 `wpsExecutionReconcile.ts` / `wpsRecoveryInspection.ts` /
`WpsAgentPanel.tsx` 三处。

**证据目录**：`outputs/a1e3-native/cross-process-abandon-20261028/`
（`README.md`、`verification.json` 11 条断言、`execution.json`、`host-writes.json`、`family.json`、
`turns.json`、`recovery-materials.json`、`workbooks.json`、`disk-values.json`、`screenshots.json`、
`snapshot-meta.json`、`screenshots/` 8 张）；
文档侧见 `ABANDON_AFTER_RESUME_ACCEPTANCE_2026-10-28.md` §16 与 `A1_GATE_REMAINING_2026-10-28.md` §6。

**A1-GATE 不因本轮单项成功而整体通过，A2 未启动。**
