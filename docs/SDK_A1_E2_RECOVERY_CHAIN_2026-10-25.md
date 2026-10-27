# A1-E2：真实 SDK 的「结果引用过期恢复」全链（2026-10-25）

承接 A1-E1。E1 只做到"**能让**某一次任务的第一个 Python 结果引用在登记后、首次写入前
精确失效"；E2 要回答的是真实 SDK 在收到 `result_ref_invalid` 之后**会不会**重新算、
新引用能不能编译出新计划、新计划在两端真实面板上**是不是只执行一次**。

## 1. 链路与唯一的替身

```
面板等价请求 POST /api/turn/stream（真实模型 claude-agent-sdk:model-bad94777…）
  → 桥轮询 /api/sdk/tool-bridge/poll → 驱动按指针真读宿主替身 → /receipt
  → prepare_python_data：登记输入数据集（句柄 dataset://）
  → run_python_analysis：真实 Docker 隔离执行 → 结果登记 R1
  → 模型用 R1 提交结果/计划 → ★A1-E1 在此消费：R1 删除 + 固定码 result_ref_invalid
  → 模型重算 → 结果登记 R2 → 用 R2 提交 → 服务端 compile_result_plan 编译白名单计划
  → 消费侧（真实 App / 真实面板）：执行一次 → 回读 → 审阅（确认 / 放弃）
```

替身只有两处：**宿主读表**（脚本里的矩阵）与**网络往返**（消费侧 fetch → 回环端点）。
模型、服务端、结果登记器、隔离执行容器、计划编译、两端执行器与审阅状态机都是生产代码。

驱动：`scripts/drive_real_sdk_result_expiry.py`（`--scenario recovery|source-change|pause`，
`--host wps-et|excel-office`）。一条命令复跑：`npm run test:a1-e2`。

## 2. 主场景（恢复成功）的证据

两份工件（两端各跑一轮，互不代替）：

| 宿主 | turnId | 旧引用 | 新引用 | 计划动作 |
| --- | --- | --- | --- | --- |
| `wps-et` | `turn-a1e2-recovery-1792899738` | `res_f4ad9571…` → **404** | `res_80869007…` → 可读 | 1（`汇总表!A1:G11`） |
| `excel-office` | `turn-a1e2-recovery-1792899871` | `res_bcf8656f…` → **404** | `res_012fb64a…` → 可读 | 1（`汇总表!A1:H10`） |

可判定的六个事实（每条都能在工件里复核）：

1. **故障确实在那一刻失效**：`fault.timeline` = `armed`（0.0s）→ `bound`（taskBound=true，
   0.62s）→ `consumed`（35.93s；Excel 那一轮是 45.85s）。`consumed` **只能**由
   `resolve_result_ref` 里的 `take_expiry` 交回 True 触发 —— 即"模型那一次确实把旧引用
   交给了真实读取路径"，而该路径 raise 的就是固定码 `result_ref_invalid`。
2. **旧引用已删**：`probes[0]` 404；`repeatOldRef` 再取一次仍 404（不因缓存/重试复活）。
3. **新引用可读**：`probes[1]` 200，且 `plan.localResultRef` 就是它 —— 新计划建立在
   **重算**结果上，不是旧引用。
4. **重算而不是重放读表**：`bridge.prepareCount = 1`、`bridge.refusals = []`；服务端账本
   这 13 条步骤是 `planning.model → bridge.prepare_python_data → python.analysis:py_823cf126cb4d → python.analysis:py_2831f4d71eac → planning.submit → planning.model`。
   两次 `python.analysis` 的步骤哈希不同（不同代码/不同执行），各自登记了新引用。
5. **重试有界**：整轮只有 2 次分析、1 次提交，没有循环重试。
6. **规划期零宿主写入**：桥只承载 `query_table` / `prepare_python_data` 两个读工具，
   写回只能来自编译后的计划动作（且需用户确认）。旧引用被拒那一轮 `planActionCount=0`。

**没有**证明的东西（照实写）：过期后模型**没有重新读表**，它复用了仍在有效期内的
**输入数据集**重算。所以"重读"这一环在真实链路上**没有发生**；发生的是"重算 + 新登记"。
这与"复用已失效引用或缓存正文"是两件事（新 `py_` 步骤 + 新 `res_` 句柄），但也不能
把本轮说成"重新读表重算"。

## 3. 三个反向分支

| 分支 | 结果 | 证据 / 说明 |
| --- | --- | --- |
| 重复旧引用 | **通过** | `repeatOldRef.ok = false`（404）；同一 `(task, ref)` 只失效一次，`_consumed` 之后不再影响别的引用 |
| 用户暂停 | **通过** | 在故障 `consumed`（25.93s）那一刻发 `POST /api/turn/{id}/pause`，`pauseRequested=true`；`kind=null`、`planActionCount=0`，账本只到 `python.analysis` **第一次** completed 就停（7 条，没有第二次分析、没有 `planning.submit`）；登记引用只有被删的那 1 个（404）；检查点 `stage=after_tool`、`resumability=continue` ⇒ 停在"可继续分析"处，**没有**产生新计划、**没有**写入 |
| 恢复期间来源变化 | **A 路 + B 路都已闭环（B 路见 §10）** | 第一次 `prepare` 交值成功后由**确定性屏障**改 `订单表!B2` 100 → 250（指纹 `eb-src:3c5c3ddd…` → `eb-src:3b0bd325…`）。A 路：模型**没有再次读取**（`prepareCount=1`、`refusals=[]`），计划照常编译（`planSourceBasis=old`，数值含旧汇总 400/1000、不含 550/1150）；**旧结果在这条路径上原本会被写进单元格（真实缺陷，已修，见 §3.2）**，修完两端写前复核各自拒绝、**写入 0 次**。B 路：2026-10-25 后续轮次用另一条测试需求让模型**真的再读一次**，被生产共享核心判据以 `source_changed` 拒交值，随后按旧来源编译出的计划在两端真实面板上仍是 **0 写入**（§10）。 |

## 3.1 「恢复期间来源变化」这一分支怎么做的（2026-10-25 补）

**约束**：不换来源表、不改提示词去逼模型重读；必须覆盖"模型**没有**再次 prepare"这一档。
所以时序只由**确定性屏障**决定，模型读不读完全交给它自己：

1. 桥把第一次 `prepare_python_data` 的**值交出去**；
2. 同一个线程紧接着改宿主：`订单表!B2` 100 → 250（身份、表名、授权、目标表都不变）；
3. 之后模型要是再读，桥只回 `source_changed`（与面板同一条判据）；
4. 模型不读，就照旧用旧值算、拿新 `resultRef`、编译计划。

**本轮实测（两端各跑一轮真实 SDK）**：

| 端 | turnId | 桥交值 | 被拒 | 计划 | 计划依据 | 计划数值 | 规划期写入 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| WPS 宿主 | `turn-a1e2-source-change-1792901938` | **1** 次 | **0** 次 | 1 个动作（`汇总表!A1:F10`） | `old` | 含 400/1000，**不含** 550/1150 | **0** |
| Excel 宿主 | `turn-a1e2-source-change-1792901998` | **1** 次 | **0** 次 | 1 个动作（`汇总表!A1:H9`） | `old` | 含 400/1000，**不含** 550/1150 | **0** |

- `planSourceBasis` 不是猜模型怎么算的：它只从"桥**实际交出去过**哪些值"推 ——
  变更之后**再没有**一次成功的交值 ⇒ `old`（模型手里只有旧那一版，计划里的固定值
  只可能来自它）。被拒的那几次只回了错误码，一个值都没给，**不算交值**。
- 两个工件的 `sourceChange` 都记着：改的哪一格、改在**第几次桥调用之后**、
  改动前后的来源指纹（`eb-src:3c5c3ddd…` → `eb-src:3b0bd325…`）。
  `npm run test:a1-e2` 会校验"改的格子确实在指纹覆盖范围内、变更晚于首次交值"。

**B 路（模型确实再 prepare）本轮没有发生** —— 两轮都是 `prepareCount=1`、`refusals=[]`。
它的判据当时不是没有证据，而是**不在这一层**：模型一旦再读，桥按既有策略拒交值，
两端各有一条面板用例钉住"只回 `source_changed`、零写入"
（Excel `sdkIntentRouting.test.tsx`「等待期间源值变化（身份不变）」、
WPS `WpsAgentPanel.test.tsx`「等待期间源值被改（工作簿身份不变）」）。
**这两条是"桥自己拒"的单元证据，不能当"真实模型再读"的端到端证据。**

**2026-10-25 后续轮次已补 B 路**：另写了一条自然需要两次读取的测试需求，两端各拿到
一轮有效实验 —— 模型**真的**在同一条任务里第二次调用 `prepare_python_data`，被生产共享
判据以 `source_changed` 拒交值，随后仍按旧来源提交计划，在两端真实面板上仍是
**0 写入**。完整证据、反向对照与变异见 **§10**；§9 末尾"B 路证据仍缺"那一条已随之
改写（未完成的部分只剩原生窗格验收）。

## 3.2 这一分支挖出的真实缺陷（已修）

**症状**：模型按旧来源算完、计划编译成功之后，若用户在"模型读完数据"到"计划到达"
之间改了来源表，**旧结果会被直接写进目标区**，界面上没有任何提示。

**失败回归（先红后修，两端各一条）**：

| 端 | 用例 | 修前 |
| --- | --- | --- |
| Excel | `resultExpiryRecoveryChain.test.tsx`「来源在分析期间变化（模型没有再次读取）：自动执行写入 0 次」 | **失败：新增写入 1 次** |
| WPS | `WpsAgentPanel.test.tsx`「计划按旧来源算完、来源在放行前被改：写入 0 次，目标区保持空」 | **失败：新增写入 1 次** |

**根因（两端的守卫位置不同，缺陷同源）**：

- Excel 的写前复核一直在（`excel.ts` 的 `executePlan`：含烘焙矩阵的计划比
  `sourceFingerprint`），但计划的指纹是**计划到达面板那一刻现读**的
  （`captureWorkbookSourceFingerprint`）。用户若在分析期间改了源表，这枚指纹记的
  就是**改后**的现场，写前再读一次当然相等 —— 自比自等，等于没有这道闸。
- WPS 压根没有计划级的来源锚点：写前只按 `prepared.sources` 逐区域复核，而由服务端
  编译的计划**一个 source 都没有**（数据是在隔离容器里读的），这一层对它完全无效。

**修法（两端读同一份判据）**：

- 新增 `packages/core/src/planSourceAnchor.ts`：`planSourceAnchor` /
  `comparePlanSourceAnchor`（三态 `unchanged | changed | unverifiable`）/
  `sourceAnchorFromBinding` / `stampPlanSourceAnchor` / `sameSheetSet` 与两句固定文案。
- **锚点改成"模型读到的那一版"**：Excel 用本轮发规划请求时的 `turnSourceBaseline`
  （身份/指纹/覆盖表三处同源），WPS 用桥绑定的指纹（`pinSourceFingerprint` 在第一个
  受控读取时固定）。只有基线缺失、或覆盖表与本次授权集合不同（用户中途改授权，
  另有闸门）时才退回原来的现读口径。
- WPS 补上计划级写前复核：`executeWpsPlan` 里 `assertPlanSourceCurrent()` 按锚点重算
  一次授权范围的内容指纹再比；Excel 的烘焙分支改读同一份三态判据。
- 拒绝时的文案：`数据来源在预览后发生了变化，这份计划写入的是预览时算出的固定值，未写入；
  请按当前数据重新生成。`（Excel 侧仍是既有的 `StaleSourceError` 长文案 + 可坚持执行）。

**修后证据**：两端回归转绿；`npm run test:a1-e2` 退出码 0（Excel `Tests 4 passed`、
WPS `Tests 183 passed`）。

**反向变异（判据必须能红，两个单点各拆一处）**：

| 变异 | 退出码 | 失败原文里的那处证据 | 还原 |
| --- | --- | --- | --- |
| 去掉 WPS 的 `assertPlanSourceCurrent()` | 1（`1 failed / 182 skipped`） | 「应判来源变化并拒绝写入（**当前新增写入 1 次**）」 | 逐字节一致 |
| Excel 把 `anchorUsable` 强制为 `false`（退回"到达时现读"） | 1（`1 failed / 3 skipped`） | 「面板没有按"来源变化"拦下这份计划；**新增写入 1 次**」 | 逐字节一致 |

两次变异都是"改字节 → 读回校验 → 跑单条用例 → `finally` 还原 → 逐字节比对"，
还原后文件哈希与变异前一致；失败原文里带的是**这条机制自己**的观测值（写入次数），
不是"退出码非 0"这种弱判据。

**仍然存在、必须写明的差异**：Excel 在 `StaleSourceError` 之后仍提供**显式**的
「仍要执行旧结果」入口（`acknowledgeStaleSource`，既有烘焙计划策略，本轮没有改动），
自动路径不会走它；WPS 侧**没有**这个入口（直接拒绝）。是否要在恢复链上撤掉这个入口，
是产品决定，本轮不擅自改。

**2026-10-25 收口更新**：SDK 固定值计划缺少本轮桥来源锚点时拒绝写入，不再退回计划到达时现读。
响应里自带指纹不能替代本轮桥证据；跨任务、缺失或不完整锚点均不能进入自动写入。
本地规则、受控计算及文件夹模式继续使用各自的确定性来源复核，不能拿某个本地动作的快照
豁免整份 SDK 固定值计划。具体修复、测试宿主变化和证据边界见 §9。

## 4. 消费侧（两端各一条）

| 端 | 用例 | 断言 |
| --- | --- | --- |
| Excel | `apps/excel-addin/src/resultExpiryRecoveryChain.test.tsx`（`EB_A1_E2=1` 门控） | 工件自洽（旧 404 / 新可读 / 计划引用新引用 / `armed→bound→consumed`）；真实 `<App/>` 上执行**一次**（写入条数 = 计划动作数）→ 逐格回读 = 计划值 → 来源表一格未变 → 「确认变更」后**写入条数不变** → 「撤销本次执行」后目标区逐格回到空表；`loadBridgeResult(旧引用)` 返回 `null`（面板拿不到正文） |
| WPS | `apps/wps-addin/src/WpsAgentPanel.test.tsx` 的 `A1-E2 消费侧` 三条（`EB_A1_E2=1` 门控） | 同一份工件自洽断言；真实 `WpsAgentPanel` 执行一次 → 逐格回读 → 「确认变更」写入条数不变、显示"已确认变更"；「放弃变更」由**产品自己**声明已放弃（`已放弃`）且目标区回到执行前、来源表未变 |

两端各自消费**自己那一轮**的工件（`EB_A1_E2_ARTIFACT` 指到 Excel 宿主那一份），
不以一端代替另一端。`scripts/run_a1_e2_chain.mjs` 会先校验两份工件自洽（引用 2 个、
旧引用 404、计划引用新引用、宿主身份对得上、时序 `armed→bound→consumed`），
工件缺失或消费侧被**跳过**都判非零退出。

替身补充说明（两处，都是"替身缺能力会伪装成产品缺陷"的坑，写在这里免得下次重踩）：
- WPS 面板替身的 `PluginStorage` 必须是真的（撤销记录存在这里、确认前还要读回来核对，
  恒返回 `null` 会让「确认变更」永远停在"无法核对"）；
- `Range().Formula` 必须**可写**，且"空公式"要落成**空单元格**（`null`）：撤销/放弃是靠
  `setRangeFormulas(restorationMatrix(...))` 一次把值恢复回去的，替身把空公式记成空串，
  恢复后按值核对就会差一格。

门控用例的两条纪律（本轮各踩一次，都只在这两条上暴露）：

1. **工件不许在 `describe.skipIf` 的工厂作用域里读**。`skipIf` 的工厂在"跳过"时仍会被调用
   （收集用例阶段就会跑），工厂里 `const artifact = readArtifact()` 会让"没有工件"这种正常
   开发/CI 状态把整套用例带崩 —— 那是环境缺失污染构建，不是产品缺陷。实测：临时移走
   `outputs/a1-e2-recovery.json` 后 `npm run test:wps -- src/WpsAgentPanel.test.tsx`
   **退出码 1**。改成 `artifactOf()` 在 `it` 内部惰性读，同一条件 **179 passed / 3 skipped、退出码 0**。
2. **门控用例的回归只在门控打开时暴露**。把工件读取挪进 `runRecoveredPlan()` 时，
   `runTurnMock.mockResolvedValue({ kind: "plan", plan })` 里的 `plan` 还指着工厂作用域里
   已经不存在的那个名字：默认套件（3 条被 skip）**照样全绿 1105 passed**，
   `npm run test:a1-e2` 却两条 `ReferenceError: plan is not defined`。
   所以：**改过门控用例之后必须跑一次开门控的命令**（`npm run test:a1-e2`），
   默认套件的绿不能替它作证。
3. **改完已有 CRLF 文件要核对行尾**。本仓库 `core.autocrlf=true`，而按文本整体写回文件
   的编辑方式会把整份文件落成 LF（实测 `WpsAgentPanel.test.tsx` 变成 6741 个裸 LF），
   于是 `git diff --check` 虽然还是 0、Git 却提示 "LF will be replaced by CRLF"。
   改完必须数一遍 `CRLF` / 裸 `LF`，归一化之后再重跑 lint 与相关套件。

## 5. lint 36 → 37 的来源（本轮定位）

结论：**是本批（A1-E1）引入的**，且是规则的**误判**，不是依赖写错。

- 第 37 条落在 `apps/wps-addin/src/WpsAgentPanel.tsx` 的 `resultRefFaultTargetOfPanel`
  这个 `useMemo` 上，文案是 "unnecessary dependency"。
- 用最小样例复现：`useMemo(() => { try { return f(); } catch { return null; } }, [dep])`
  —— 函数体只有 try/catch 时，`react-hooks/exhaustive-deps` 把函数体当成"没用到任何
  依赖"，于是每个依赖都被判成多余。同一文件里 3 条更早的同类告警（`authorizationKey`）
  也是这个形状，属于既有噪声。
- 处理：把依赖真正用于判断（读现场后核对"读到的还是面板刚观察过的那本簿吗"，不一致
  就不给目标、按钮置灰），而不是让它当"缓存失效开关"。改完 **37 → 36**，未放宽任何断言、
  未加 `eslint-disable`。

## 6. 本轮全量检查（真实退出码）

| 检查 | 结果 |
| --- | --- |
| `npm run lint` | 0（0 error / **36** warning，与修前同数：新增代码 0 告警） |
| `npm run test:addin` | 0（**1551 passed / 7 skipped**；7 = 4 条 A1-E2 门控〔3 条主场景 + 1 条来源变化〕+ 3 条既有跳过） |
| `npm run build:addin` | 0（`tsc --noEmit` + `vite build`） |
| `npm run test:wps` | 0（**1106 passed / 6 skipped**；新增的来源变化用例**常开**，不吃门控） |
| `npm run build:wps` | 0（`tsc --noEmit` + `vite build`） |
| `pytest server/tests` | 0（**1763 passed / 1 skipped**，109.6s，`--basetemp` 在 OS 临时目录） |
| `compileall server/app scripts` | 0 |
| `git diff --check` | 0 |
| `npm run test:a1-e2` | 0（Excel `Test Files 1 passed` / `Tests 4 passed`；WPS `Tests 183 passed`；另校验**四份**工件：两份主场景 + 两份来源变化） |

上表是**修完门控用例之后**整套重跑的结果（见 §4 的两条纪律：那次 `plan` 的 `ReferenceError`
只有在 `EB_A1_E2=1` 时才现形，默认套件一直是绿的）。重跑顺序与实测耗时：
lint + `test:addin` + `build:addin` 共 80s，`test:wps` + `build:wps` 共 58s，
`test:a1-e2` 35s，pytest 约 110s。`npm run test:a1-e2` 在**四份工件都在**时，
Excel 端 `Tests 4 passed`、WPS 端 `Tests 183 passed`。

行尾归一化（§4 第 3 条）之后，**受影响的检查又重跑了一轮**，结果与上表一致：
`lint` 0（0 error / 36 warning）、`test:wps` 0、`build:wps` 0、`test:a1-e2` 0、
`test:addin` 0、`build:addin` 0、`git diff --check` 0。

§3.1/§3.2 的来源变化分支（含缺陷修复）落地之后，**又整套重跑了一轮**：
`lint` 0（36 warning）、`test:addin` 0（1551/7 skip）、`build:addin` 0、
`test:wps` 0（1106/6 skip）、`build:wps` 0、`test:a1-e2` 0（4/183）、
`pytest` 0（1763/1）、`compileall` 0、`git diff --check` 0。

一处**环境噪声**（记下来免得被当成回归）：把 lint + 两端 test + 两端 build 串在**同一条命令**
里跑时，`vite build` 的 `emptyDir` 会被环境的 safe-delete 批量闸挡下
（`SAFE_DELETE_BULK_CONFIRM_REQUIRED {"count":50,"threshold":50,"scope":"turn"}`，
计数**跨进程**累计，前面的测试已经用掉额度）。同一轮的 `tsc --noEmit` 已通过，
`npm run build:wps` **单独重跑退出码 0**。这不是产品缺陷，与 pytest 侧那条
safe-delete 噪声同源。

**B 路（§10）落地之后又整套重跑了一轮**，数字见 §10.6：`lint` 0（36 warning）、
`test:addin` 0（**1552**/9 skip）、`build:addin` 0、`test:wps` 0（**1118**/8 skip）、
`build:wps` 0、`test:core` 0（834）、`pytest` 0（**1770**/1）、`compileall` 0、
`git diff --check` 0、`test:a1-e2` 0（**六份**工件、Excel 6 / WPS 187）。日志前缀
`outputs/a1e2b-`。上表那轮的数字（1551 / 1106 / 1763）是**更早一轮**的实测，
两轮之间新增的用例是 B 路的 7 条服务端判定用例 + 两端各 2 条门控用例，不改已有断言。

## 7. 两端最短原生操作步骤（A1-E2，不使用 computer-use）

前置：服务端以 `EXCEL_BRO_RESULT_REF_FAULT=1` 启动（开发构建的面板才会显示该入口）。

1. 打开**独立测试工作簿**（Excel、WPS 各一次；不要用当前工作簿），勾选来源表 + 目标表并「确认」；
2. 面板底部「结果引用过期（开发专用）」→「启用下一次引用过期」，确认状态是"已启用（一次性）"；
3. 发一条需要 Python 分析并写入目标表的请求，等它跑完；
4. 验收要点（逐条看，不要合并）：
   - 旧结果引用被拒（界面出现"结果引用无效或已过期，请重新查询"或等价固定文案）；
   - **目标表在旧引用被拒那一刻一个单元格都没变**；
   - 模型在同一任务里重新算出一份新结果，通过校验后**自动执行，再审阅结果**；
   - 执行只发生**一次**（目标区 = 计划值，逐格核对），随后进入审阅；
   - 「确认变更」后目标区不再变化；「放弃变更」后目标区回到执行前；
   - 全程来源表未被修改。
5. **来源变化这一支（§3.1）的最短复验**：发请求前把目标表窗口并排打开；
   在面板显示"正在读取/正在分析"**之后、写入之前**，手动改一个来源金额
   （例如订单表 B2 的 100 改成 250），然后等结果：
   - 已算完的计划**不得写入**（目标区保持原样），界面出现
     "数据来源在预览后发生了变化…未写入；请按当前数据重新生成"；
   - 若面板给的是 Excel 的「仍要执行旧结果」入口，**不要点它**（那是有意的显式覆盖，
     不属于自动路径）；
   - 记下这次看到的顺序（读取 → 改源 → 计划到达 → 拒绝 → 0 写入）。
6. 把界面上看到的顺序（拒绝 → 重算 → 执行 → 审阅）与 A1-E1 面板状态（已触发）记在一起。

## 8. 不宣称

- **「恢复期间来源变化」的 A 路 + B 路都已闭环，但都是"真实组件 + 替身宿主"**。
  A 路：模型没有再次读取，旧结果由写前复核拦下；B 路：模型确实再读一次、被
  `source_changed` 拒交值、随后的旧计划仍是 0 写入（§10）。两条都**没有**在真实
  Excel / 真实 WPS 窗格里跑过，原生实机仍算未做。
- 本轮**新增了一条产品行为**（两端写前来源锚点 + WPS 计划级写前复核，§3.2），
  它只在"真实组件 + 真实服务端 + 宿主替身"上验过；**真实 Excel / 真实 WPS 窗格里
  一次都没跑过**，原生实机仍算未做。
- Excel 的「仍要执行旧结果」显式覆盖入口**仍然存在**，自动路径不会走它；WPS 侧没有这个
  入口。这条差异是既有产品策略，是否撤掉属于产品决定，本轮没有擅自改。
- B 路证明的是"再读一次会被拒"与"迟到的旧计划不会写入"。它**不**证明"模型一定会回执
  给用户"：两端模型都只在计划里写了一句"来源核对未完成"，界面文案的实际呈现未在原生
  窗格核过。
- A1-E2 完成 **≠** A1 整体验收，A1-GATE 仍不达标。

## 9. 2026-10-25 Codex 接手：来源锚点收口

本轮检查的是 DeepSeek 已留下的部分实现，保留既有未提交修改。未调用 computer-use、
未操作用户工作簿、未启停用户 8765 服务、未提交或推送 Git。

### 生产路径补齐

1. `stampPlanSourceAnchor(plan, null)` 必须清除响应里的指纹与来源列表。旧实现原样返回
   响应，导致“本轮没有绑定”仍可能沿用响应自带指纹。新增回归先失败再修复。
2. WPS SDK 固定值入口由可信面板传入 `requireSourceAnchor`；本地规则/本地映射仍走
   原有逐区域来源检查。带 `localResultRef` 的固定值计划即使漏传选项仍必须有锚点。
   普通固定值本地操作不能因 SDK 闸门误伤，SDK 也不能借无关本地快照绕过。
3. `sourceAnchorRequired` 加入不可变准备快照；准备后篡改标志必须拒绝。
   面板的编辑与回退重准备路径保留该标志。
4. WPS 写前只复核锚点所含的已授权来源表，不读取未授权表，也不把额外授权目标表
   混入来源指纹。按原锚点顺序复算：同一来源集合换顺序时指纹不同，不能误报来源变化。
5. 双端仍先执行一次再审阅；确认不再次执行。Excel 的既有显式旧结果覆盖入口未改。

### 回归与替身边界

- 接手时 WPS 全量 118 failed / 988 passed / 6 skipped：SDK 来源闸误伤本地计划，
  同时旧面板替身直接返回固定值计划，没有模拟真实桥绑定。
- 共用缺绑定回归先红（1 failed / 13 passed）：`outputs/codex-anchor-red.log`。
- 新增跨任务桥拒绝、响应伪造指纹拒绝、缺锚点/不完整锚点、未授权来源、现场不可读、
  来源子集、不可篡改执行分类、SDK 不借本地快照豁免、结果引用强制检查等回归。
- 来源顺序回归先红（1 failed / 44 passed）：`outputs/codex-anchor-order-red.log`；
  修正后该文件 45 条随全量通过。
- 普通面板替身现在对模拟 SDK 固定值响应建立真实前端桥绑定并投递受控读进度；
  专门的来源变化/桥故障用例关闭该便捷适配，继续自行控制真实桥事件与屏障。
  不给计划硬塞指纹，不放宽来源检查，不把替身当成真实模型交值证据。
- 测试宿主补齐 UsedRange、Item 与 Range(firstCell,lastCell) 读法。门控宿主也必须
  支持该读法，否则来源核对合法拒绝；门控初跑两条失败保留在 `codex-anchor-e2.log`，
  补齐替身后开门控重跑通过。仍消费既有四份真实链工件，本轮没有重新运行真实 SDK/Docker。
- WPS 构建首跑类型失败（新增用例缺 scopeDock 参数、替身 Range 可选类型），修正后
  构建通过；失败日志 `codex-anchor-build-wps.log` 保留。没有通过扩超时或放宽断言解决。

### 本轮实际检查

| 检查 | 结果 | outputs 日志 |
| --- | --- | --- |
| test:core | 834 passed，exit 0 | codex-anchor-core-final.log |
| test:addin | 1552 passed / 7 skipped，exit 0 | codex-anchor-excel-final.log |
| test:wps（最后修正后） | 1118 passed / 6 skipped，exit 0 | codex-anchor-wps-complete.log |
| test:a1-e2（开门控） | Excel 4 passed / WPS 185 passed，exit 0 | codex-anchor-e2-final.log |
| build:addin / build:wps（含 tsc） | exit 0 | codex-anchor-build-excel.log / codex-anchor-build-wps-complete.log |
| pytest server/tests | 1763 passed / 1 skipped，exit 0 | codex-anchor-server-final.log |

服务端测试有既有 pytest 缓存目录写权限警告（不影响测试退出码）；WPS 构建仍有大 chunk
警告；面板测试仍有 React act 警告。最终 lint 为 0 error / 36 warnings，exit 0
（`outputs/codex-anchor-lint-final.log`）；compileall server/app scripts 为 exit 0；
git diff --check 为 exit 0（`outputs/codex-anchor-diff-check.log`）。diff-check 仍提示
三个既有 Excel 测试文件的 LF 将被 Git 转为 CRLF，本轮未为此改写无关文件。

### 后续仍需完成

- ~~B 路真实模型第二次 prepare 的 `source_changed` 拒绝证据仍缺~~ → **2026-10-25
  后续轮次已补，见 §10**（两端各一轮有效实验；不能拿 A 路写前拒绝或模拟桥用例代替
  这条约束仍然有效）。
- 原生 Excel/WPS 的过期恢复、改源拒写与实际写入次数证据仍按 §7 由用户验证。
  最终值相同不能证明没有重复写入。
- A1-E3 故障矩阵、原生竞争/异常退出证据仍未齐；A1-GATE 不标完成，A2 未启动。

## 10. B 路：模型**确实**再读一次（2026-10-25 后续轮次）

A 路（§3.1）证明的是"模型不再读时，旧结果只能被写前复核拦下"。它答不了一个问题：
**模型再读一次的时候，那次读取走的是不是生产桥的授权/来源复核**。B 路要的就是这个：
同一条任务里模型自己发起第二次 `prepare_python_data`，被 `source_changed` **拒交值**
（一个值都不给），随后即使它仍按旧来源提交计划，也不能绕到写入。

### 10.1 只改驱动、测试与文档，生产代码不加新行为

1. **判定不再复刻一份**。驱动原先在 Python 里写了"指纹不等就回 `source_changed`"，
   那是第二份判据。新增 `scripts/bridge_authorization_gate.mjs`：stdin 收
   `{binding, call, live}`，直接 `import` `packages/core/src/sdkToolBridge.ts` 的
   `verifyBridgeAuthorization`，stdout 回 `{ok:true}` / `{ok:false,errorCode}` ——
   与真实面板执行前复核同一份代码。入口自身出错（Node 版本太老、输入形状不对）**退出码 2
   并中止实验**，绝不当成"拒绝了一次读取"混进工件。
2. **新场景** `source-change-reread`：提示词只在既有需求后**追加**一段用户要求
   （"提交写入预览前再重新读取一次这两张来源表核对金额；这是一次新的读取，不要复用
   之前登记的数据集……被拒绝就按第一次读到的数据生成预览并注明来源核对未完成"）。
   来源表、授权范围、写入目标、故障开关一概不变；读不读、读几次仍由模型自己决定。
3. **分支判定** `second_read_evidence()`：`none`（没来源变更记录）/ `A`（变更后没有新的
   `prepare`）/ `B`（再读了、且被拒）/ `B-unrefused`（再读了但**没**被拒 —— 算缺陷，
   不算通过）。`plan_source_basis()` 只从"桥**实际交出去过**哪些值"推计划依据，
   被拒的那几次只回了错误码、一个值都没给，**不算交值**。
4. **修掉驱动自身一个口径错位**：`reads_at_prepare.append(len(served))` 记的是"第几次
   成功交值"，却拿去和"第几次桥调用"（`index`）比。B 路只要中间多一次被拒调用，两者就会
   错开一格，屏障会改成"变更早于首读"的错误结论。改成按调用序号记录。
5. **`repeatOldRef` 只在启用引用过期时校验**：B 路刻意不启过期，两个引用本来就都有效，
   拿"旧引用再取一次可读"当异常是驱动措辞错误（Excel 轮日志里那行"异常"即此，工件字段
   保留但链路校验不读它）。

### 10.2 真实实验：每端 1 轮有效，第 1 轮如实归档

| 轮次 | 端 | turnId | 桥结果 | 覆盖 B 路？ | 工件 |
| --- | --- | --- | --- | --- | --- |
| 第 1 轮 | WPS 宿主 | `turn-a1e2-source-change-reread-1792926801` | 5 次 `prepare_python_data` + 1 次 `query_table` **全部** `source_changed`；最后 `MODEL_TURN_LIMIT`，`kind=null`、`refs=0` | **部分**：拒绝侧覆盖到了（守卫在连续重试下没有松口），但没有计划 ⇒ "迟到旧计划不得写入"没东西可验 | `outputs/a1-e2-source-change-reread-round1-turntlimit.json` / `.log` |
| 第 2 轮 | WPS 宿主 | `turn-a1e2-source-change-reread-1792926973` | 第 2 次 `prepare` 被拒；`kind=plan`、1 个动作 | **是** | `outputs/a1-e2-source-change-reread.json` / `.log` |
| 第 2 轮 | Excel 宿主 | `turn-a1e2-source-change-reread-1792927052` | 第 2 次 `prepare` 被拒；`kind=plan`、1 个动作 | **是** | `outputs/a1-e2-source-change-reread-excel.json` / `.log` |

第 1 轮失败的原因**不是**守卫太紧，而是需求没写"被拒后不要反复重试"：模型连着重试到
轮数上限。第 2 轮的提示词补了这一句，两端各一次就拿到有效轮次。真实实验共 **3 轮**
（每端不超过 2 轮），没有无上限重试碰运气；第 1 轮的工件与日志**保留**，不删不覆盖。

### 10.3 两端证据

| 项 | WPS 宿主 | Excel 宿主 |
| --- | --- | --- |
| 首次受控读取（**交值**） | #1 `op_84cd9f344bb3e3399a4ed021d0c53980` → `bridge.prepare_python_data:84cd9f344bb3`，`completed` | #1 `op_053bfed62c994ce457160668f5cab949` → `bridge.prepare_python_data:053bfed62c99`，`completed` |
| 改源屏障 | 首次交值**之后**改 `订单表!B2` 100 → 250；指纹 `eb-src:3c5c3ddd…` → `eb-src:3b0bd325…` | 同（同一份夹具、同一格、同一对指纹） |
| 第二次真实工具调用 | #2 `op_521c69d416dd4306119e64e47af3b361` → `bridge.prepare_python_data:521c69d416dd`（37.36s） | #2 `op_2426ce6c93c2806528ea0ec97e29c670` → `bridge.prepare_python_data:2426ce6c93c2`（66.40s） |
| 裁决从哪来 | 生产 `verifyBridgeAuthorization`：`gate = {ok:false, errorCode:"source_changed"}` | 同 |
| 拒绝码 / 是否交值 | `source_changed`；`served=false`（一个值都没给） | 同 |
| 来源范围 | `会员表!A1:C6` + `订单表!A1:B5`，两次**完全相同**（`sameScope=true`） | 同 |
| 服务端账本（11 / 15 步） | `planning.model` → `prepare#1 completed` → `prepare#2 warning` → `python.analysis:py_4c27cf27c7ae` → `planning.submit` → `planning.model completed` | `planning.model` → `prepare#1 completed` → `python.analysis:py_10a5f9e1597c` → `prepare#2 warning` → `python.analysis:py_35d2000c976c` → `planning.submit` → `planning.model completed` |
| 计划 | `汇总表!A1:G11`，1 个动作，固定值 400 / 200 / 400 / 1000 | `汇总表!A1:G10`，1 个动作，固定值 400 / 200 / 400 / 1000 |
| 计划依据 | `old` | `old` |
| 规划期宿主写入 | **0** | **0** |

这六条让它站得住（每条都能在工件里复核）：

1. **第二次调用是模型自己发起的**，不是驱动造的：独立 `operationId` 由服务端下发、
   独立 `stepId`、独立账本条目；用户需求要的是"新的读取"，驱动只负责应答。
2. **裁决不是驱动拼的字符串**：`gate` 字段就是生产判据的返回值，驱动对同一份 TS 直接执行。
3. **账本把这次调用记成 `warning`**（"没做成"），而第一次是 `completed` —— 两条读的是
   不同数据源（桥自报 vs 服务端账本），互为锚点。
4. **改源晚于首次交值**：`afterServedCall=1` 与首次交值的调用序号相等，
   链路校验会挡住"变更早于首读"的工件（那等于模型读到的本来就是新值）。
5. **两次读的是同一批来源**：`sameScope=true` ⇒ 两次之间唯一的变量是"数据变了"。
6. **两端模型被拒后仍提交了计划**（WPS 的分析在拒绝**之后**才跑，Excel 是算完一次、
   被拒之后**又算了一次**），计划数值仍是旧汇总 400/200/400/1000、不含新汇总 550/1150；
   两份计划的文案里都写着"第二次读取被拒绝（`source_changed`），来源核对未完成" ——
   这是拒绝回执**真的送达
   模型**的证据，不只是本地记了一笔。于是"迟到的旧计划"在真实链路上**确实出现了**，
   而不是"模型自己收手了所以没有东西可验"。

### 10.4 写入 0 次：两端真实面板，消费侧各两条用例

| 层 | 证据 | 结果 |
| --- | --- | --- |
| 规划期 | 两端工件 `hostWritesDuringPlanning = 0`；桥只承载 `query_table` / `prepare_python_data` 两个读工具 | 0 |
| Excel 面板 | `resultExpiryRecoveryChain.test.tsx`「迟到的旧计划：真实 `<App/>` 上写前复核拒绝，写入 0 次，目标区保持空」 | 写入 **0** 次；目标区（计划范围 `汇总表!A1:G10`）逐格为空、`会员表!A1:C6` 逐格等于原夹具、被改的那一格仍是 250；界面出现"数据来源在预览后发生了变化" |
| WPS 面板 | `WpsAgentPanel.test.tsx`「A1-E2 B 路（WPS）：模型再次读取被 source_changed 拒绝」：`bridgeRoundTripHost` + `armBridgeTurn` 投递这条任务，`host.mutate("订单表",2,2,250)` 只由用例执行，先等界面出现"数据来源在预览后发生了变化"再查 `host.writes()` | 写入 **0** 次（新增写入 0）；`汇总表` 1..12 行 × 1..8 列逐格 `null`；`会员表!B2` 仍是原值、`订单表!B2` 是 250；终态 `host.writes()===0` |

WPS 端那条用例先构造**真实桥往返**（首读交值 → 改源 → 第二次 `prepare` 由面板按生产
判据回 `source_changed`），再把该工件里的计划作为"迟到的计划"投进去 —— 也就是说它
验的是面板自己的写前复核，不是"用例手写了一个拒绝"。Excel 端同理走真实 `<App/>` 的
`executePlan`。两端都不吃默认跳过：`npm run test:a1-e2` 里 Excel `Tests 6 passed`、
WPS `Tests 187 passed`。

### 10.5 反向对照与变异（这些断言必须能红）

| 手段 | 命令 | 结果 |
| --- | --- | --- |
| 拿 A 路工件跑 B 路用例（反向对照） | `EB_A1_E2_REREAD_ARTIFACT=outputs/a1-e2-source-change-excel.json` | 两端都红：`expected 'source-change' to be 'source-change-reread'` —— 说明这组断言不是"随便一份工件都能过" |
| 变异 B1：去掉 WPS `assertPlanSourceCurrent()` | `python outputs/mut-a1e2b/mutate.py B1` | 退出码 1，`1 failed / 185 skipped`，失败原文「应判来源变化并拒绝写入（**当前新增写入 1 次**）」 |
| 变异 B2：Excel 把来源闸整块跳过（`if (false && (anchor || …))`） | `python outputs/mut-a1e2b/mutate.py B2` | 退出码 1，`1 failed / 4 skipped`，失败原文「面板没有按"来源变化"拦下这份迟到的旧计划；**新增写入 1 次**」 |

两次变异都是"命中唯一（`check` 输出 `hits=1`）→ 改字节 → 读回校验 → 跑单条用例 →
`finally` 按原 EOL 写回 → 逐字节比 sha256"。还原后 `wpsPlanExecutor.ts` =
`33d65970…`、`excel.ts` = `e4ecf7f9…`，与变异前一致（复核命令：
`sha256sum apps/wps-addin/src/wpsPlanExecutor.ts apps/excel-addin/src/excel.ts`）。
失败原文里带的是**这条机制自己**的观测值（写入次数），不是"退出码非 0"这种弱判据。

### 10.6 本轮检查（真实退出码，日志在 `outputs/a1e2b-*.log`）

| 检查 | 结果 |
| --- | --- |
| `npm run lint` | 0（0 error / **36** warning，与上轮同数） |
| `npm run test:addin` | 0（**1552 passed / 9 skipped**；9 = 6 条 A1-E2 门控 + 3 条既有跳过） |
| `npm run build:addin` | 0（`tsc --noEmit` + `vite build`） |
| `npm run test:wps` | 0（**1118 passed / 8 skipped**） |
| `npm run build:wps` | 0 |
| `npm run test:core` | 0（**834 passed**） |
| `pytest server/tests` | 0（**1770 passed / 1 skipped**，122.5s，`--basetemp` 在 OS 临时目录） |
| `compileall server/app scripts` | 0 |
| `git diff --check` | 0（仍是那 3 条既有 Excel 测试文件的 LF→CRLF 提示） |
| `npm run test:a1-e2` | 0（**六份**工件全部自洽：两份主场景 + 两份 A 路 + 两份 B 路；Excel `Tests 6 passed`、WPS `Tests 187 passed`） |

服务端新增 7 条 `server/tests/test_a1_e2_driver_branch.py`，钉的是**判定本身**而不是
跑一遍模型：无变更 ⇒ `none`；没再读 ⇒ `A`；再读且被拒 ⇒ `B`（带 `operationId`/
`stepId`/`code`/`sameScope`）；再读**没**被拒 ⇒ `B-unrefused`（缺陷不算通过）；来源范围
不同 ⇒ `sameScope=false`；被拒的读取不算交值 ⇒ `planSourceBasis="old"`；
`afterServedCall` 是**调用序号**不是交值计数（口径写错会给出相反结论）。

### 10.7 仍然缺什么（逐端，不许合并宣称）

- **Excel**：B 路只到"真实 SDK + 真实服务端 + 真实 Docker + 真实 `<App/>` 组件 + 替身宿主"。
  **真实 Excel 窗格里一次都没点过**，"来源在预览后变化 → 拒写"这条提示实际长什么样、
  是否被用户看见，都还没有原生证据。Excel 的「仍要执行旧结果」显式入口仍在。
- **WPS**：同上，只到"真实 `WpsAgentPanel` 组件 + 替身宿主"；真实 WPS 窗格未验。
- **模型行为**：第 1 轮说明模型在没被明确要求时会连着重试到轮数上限。守卫在重试下没有
  松口（每次都 `source_changed`），但任务会以 `MODEL_TURN_LIMIT` 失败。这是既有产品
  行为，本轮**没有**改（改它属于产品决定，不在"补证据"范围内）。
- **A1-E3 故障矩阵**、原生竞争/异常退出证据仍未齐；**A1-GATE 不标完成，A2 未启动**。
