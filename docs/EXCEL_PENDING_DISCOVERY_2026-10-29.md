# Excel「未完成执行」发现入口（CAP-03 Excel 侧）

- 日期：2026-10-29
- 性质：实现 + **替身自动化**证据；**真机未验收**
- 范围：只做「找到入口 + 说清楚能不能续跑」，不改既有写入路径、不扩授权
- 关联：`docs/A1_GATE_REMAINING_2026-10-28.md`、`docs/WPS_RECOVERY_DIAGNOSTICS_2026-10-29.md`、
  `packages/core/src/pendingExecutionDiscovery.ts`（上一轮的共享发现判据）

## 0. 一句话

Excel 端此前**只有一条 `localStorage` 索引**把用户带回断点：索引不在（换机器、清缓存、
窗格重开、换端口）就再也看不到这次执行，而账本里的进度和写了一半的动作都还在 ——
与 WPS 2026-10-27 那个现场同源。本轮补上**后端发现**，并把「发现到了 ≠ 可以续跑」
交给共享判据 `pendingResumeReadiness` 判，而不是由界面顺手给一张"能继续"的卡。
**第十轮追加**：真机截图暴露"能发现 ≠ 交互和 WPS 一致"——恢复卡还是一个「核对并继续」
把补记与续跑合并。本轮把结论/措辞/按钮集合搬到共享层（`executionRecoveryView`），
Excel 端拆成「核对在途动作」（只补记，零单元格写入）/「继续剩余动作」（只跑没跑过的）/
「补记完成（不重写原表）」，见 §5.1。
**第十一轮追加**：真机又暴露第二个问题 —— 卡片就绪了，**入口却被滚出视口**。核对成功后
追加的日志把消息区滚到底，而更新成「继续剩余动作」的卡还挂在旧消息下面。本轮把**可点击的
那一份**搬到输入区上方的固定区域（`.office-recovery-panel`），消息里只留**没有按钮**的执行记录，
并用真实浏览器布局量出"入口在视口内、追加日志后位置不变、旧位置复现看不见"，见 §5.2 / §4.5。

**验收分档（不得合并宣称）**：
① 工件/替身自动化 **已通过**；② 双端在线集成 **未做**；③ 真实 Docker **不适用**；
④ 真实模型规划 **未做**；⑤ 真实 Excel **未做**；⑥ 真实 WPS **未做**（本轮不碰 WPS 判定）。

## 1. 改造前的实际缺口（逐条核过代码，不是猜）

| # | 事实 | 位置 |
| --- | --- | --- |
| G1 | 唯一入口是本机单条索引 `excel-bro.pending-office-execution.v1`，**没有后端发现** | `apps/excel-addin/src/officeOperationRecovery.ts` |
| G2 | 挂载路径只比 `pending.workbookIdentity === identity`，**不比授权、不比在途材料** | `App.tsx` 挂载效应（`officeRecoveryAttachStartedRef`） |
| G3 | 后端接口**早就存在且与 WPS 同构**（`GET /api/office/executions/pending?workbookIdentity=…`），只是 Excel 面板没调 | `server/app/main.py:3127`、`api.ts` `listPendingOfficeExecutions` |
| G4 | 多候选（同一 URL 多条未完成）没有任何选择入口 | — |
| G5 | 用户改正授权后，卡面**不会**跟着改判（只在点续跑时被执行器拦） | — |

## 2. 实现

### 2.1 共享判据（新，宿主无关）

`packages/core/src/pendingExecutionResume.ts`

```ts
pendingResumeReadiness(input: PendingResumeInput): PendingResumeReadiness
pendingResumeTargetSheets(plan): string[]
```

判定顺序固定：**身份 → 状态 → 登记授权 → 用户确认过的范围 → 工作表是否还在 → 在途快照**。
八个原因码，每个只在**一处**返回点产生（界面按码给文案，不按字符串猜）：
`identity_unavailable`、`identity_mismatch`、`completed`、`not_resumable`、
`plan_target_never_authorized`、`plan_target_not_authorized_now`、`plan_target_missing`、
`in_flight_without_snapshot`。

**关键设计取舍（本轮最要紧的一条）**：`currentAuthorizedSheets` 传的是
**用户在本会话里真的确认过的范围**，不是"宿主的当前勾选"。原因：Excel 窗格重新挂载时
`applyWorkbookSnapshotSelection` 的首读播种会把已提交范围写成**当前活动表**，那不是用户手势；
拿它比对，真机上就会变成"我什么都没动，界面却说我把汇总表取消了"。
`null` = 没有可比对的手势 ⇒ **不比对**（注意：不是"一张都没授权"）。

### 2.2 Excel 端发现（薄包装，镜像 WPS）

`apps/excel-addin/src/officeRecoveryDiscovery.ts`

- `currentOfficeRecoveryIdentity()`：只认 `url:` 前缀，未保存 / 读失败 ⇒ `null`，**一个请求都不发**。
  未保存簿的 `unsaved:` 会话内对象身份**不用于发现**（账本登记的永远是 `url:`）。
- `discoverOfficePendingExecutions()`：身份在本机读、发现走后端；响应再过一次准入闸
  （`assertPendingExecutionDiscovery`），并校验 `namespace === "office"`（WPS 候选不许串台）。
- 三种结论：`no_identity` / `unavailable`（**不是**"没有未完成任务"）/ `resolved`。

### 2.3 多候选显式选择

`apps/excel-addin/src/components/OfficeRecoveryChoicesCard.tsx`：住在输入区上方（不是消息里），
每条候选一个「核对这一次」；`truncated` 给截断说明；失败给 `role="status"` 文案。
排序仍由共享层 `orderPendingExecutionCandidates` 决定（有在途动作排最前、写前失败排最后），
界面**不**挑"最新的一条"。

### 2.4 已选执行重新核验 + 异步结果防覆盖

`App.tsx`：

- `officeResumeReadinessFor(state)`：唯一填参点，身份/状态/授权/在途材料都从这里进共享判据。
- `reverifySelectedOfficeRecovery(isCurrent)`：只读 `GET` 状态 + 同一判据重算，只对**用户已选的那一条**。
- `discoverOfficeRecovery(identity, isCurrent)`：缓存优先（索引没人接手时才读）→ 后端兜底 →
  `choose` 只给入口 → `unavailable/mismatch` 挂提示**不清卡**（清卡等于把入口弄丢）→
  `status === "completed"` 直接 `settled`（清索引 + 保留挂载效应找回的完成回执，不改写成阻断卡）。
- **代数闸**：`officeDiscoveryGenerationRef`，依赖（身份 / 范围内容 / 口径 / 确认版本）变化即作废旧代数，
  迟到的响应 `isCurrent()` 为假 ⇒ 既不摆卡也不换对象。

### 2.5 只读边界

发现与重核**全程只读**：不发 `claim`、不 `reconcile`、不推进 revision、不执行动作、不写单元格。
写入资格仍要在用户点续跑时按**既有**路径重新取得（`registerOfficeExecution` 及其后续不变）。

## 3. 自动化抓到的真实缺陷（本轮修掉，勿复活）

首版把"用户确认过的范围"用 `useEffect([committedSheetNames, committedScopeMode])` 同步进 ref。
两条都会让真机上"用户已经改正、界面还说旧结论"：

1. **值不变 ⇒ effect 不重跑**：用户把范围收成与首读播种相同的一份（很常见：播种本来就是活动表），
   `committedSheetNames` 内容相同，effect 不触发，ref 永远是 `null` ⇒ 不做比对 ⇒ 仍判"可继续"。
2. **时序**：守卫那条路在 `await` 里 `adopt` 提交，渲染与 effect 可能**早于**"这次确认是用户手势"
   的标记落下来 ⇒ effect 提前 return，之后没有新的渲染把它补回来（ref 变化不引起渲染）。

**修法**：确认成功处**直接写** ref（`recordConfirmedScope`，唯一写入点），并用一个自增版本号
（`officeScopeConfirmedRevision`）进发现依赖来**触发**重核 —— 不靠"值变了"来触发。
`scripts/mutate_office_recovery_discovery.mjs` 的 D1 就是拆掉这一处，用来钉死它。

## 4. 证据

### 4.1 自动化（替身宿主 / 真实 App 组件 / 真实服务端判据）

| 层 | 结果 |
| --- | --- |
| `packages/core` | **911 通过 / 61 文件**（含 `pendingExecutionResume.test.ts` 12 条、`pendingExecutionDiscovery.test.ts`、`executionRecoveryView.test.ts` 13 条 —— 第十一轮加了 5 条入口唯一性/进度句） |
| `apps/excel-addin` | **1590 通过 / 9 跳过 / 122 文件**（含 `officeRecoveryDiscovery.test.ts` 8 条、`Excel 未完成执行的发现入口` 13 条 —— 第十一轮加了入口位置/换簿原地改判/核对中三态） |
| `apps/wps-addin` | **1217 通过 / 8 跳过 / 70 文件**（本轮未改 WPS 判定） |
| 服务端 | **1913 通过 / 27 跳过 / 2 失败**；两条失败在 `test_python_event_delivery.py`（期望 `completed`、实际 `unavailable`），本机 **Docker 引擎未启动** ⇒ `runner_missing`，属环境阻断（`docs/CHECKPOINT_2026-10-17.md` 起已记录） |
| 构建 | `npm run build:addin`、`npm run build:wps` 退出码 0；`apps/excel-addin` `tsc --noEmit` 0 错 |
| lint | 新增/改动文件 `eslint` **0 error**（`App.tsx` 有 8 条既存的 hooks 依赖告警，非本轮引入） |

面板回归十条（`excelWriteEndToEnd.test.tsx`，真实 `<App/>` + 替身宿主 + 网络替身）：

1. 多条候选**只给选择入口**：两个「核对这一次」、`getOfficeExecutionStatus` 未调用、`host.writes()` 为空、
   六个写入类接口全未调用。
2. 点选身份不符的一条：只给阻断 + 「重新核对（只读）」，**无**续跑按钮，且只查了被点那一条。
3. 单条自动摆卡 → 用户确认收窄授权 → 卡面变红（点名「汇总表」）、无续跑按钮、重核的是同一条。
4. **迟到的发现结果不许落盘**：第一次请求在飞时授权已变，release 旧响应 ⇒ 零状态查询、零候选、零写入。
5. **窗格重开 + 本机索引不在**（探针在第 2 步回读后中断 → `clearPendingOfficeExecution` → 重新挂载）：
   靠后端发现摆出「待核对在途动作」→ 点「核对在途动作」（**零单元格写入**，只补记回执）→
   卡片换成「可继续剩余动作」→ 点「继续剩余动作」只写第 3 步（写入只有 `B4`，第 2 步不重放）
   → 审阅收尾「从第 3 步继续并完成」、无「执行未完成」。这一条就是真机最小验收的自动化孪生体。
6. **`confirmedScope=null`（从没确认过范围）**：首读播种是活动表、计划要写汇总表，判定**不**比对 ⇒
   卡可继续、不出现"你把表取消了"、`.scope-trigger` 文案一字不变、发现补回挂接记录、零写入。
7. **授权被撤销（清空并确认）**：卡面立刻阻断、只给只读重核、再点重核仍是阻断；随后刷新一轮目录
   （草稿会重新播种）也**不会**让卡片回到"可继续"、零写入。
8. **换簿后返回**：`host.setDocumentUrl(另一本)` ⇒ 改判「当前不是产生该断点的工作簿」、无续跑按钮、
   **不清**本机索引；切回来 ⇒ 卡面恢复、每次状态查询都是同一条 `operationId`、全程零写入。
9. **续跑不信任卡面**：卡片摆好之后把状态改成"在途但没快照" ⇒ 点下去立刻以
   「在途动作缺少写前快照」停下，`startOfficeAction` / `registerOfficeExecution` 均未调用、零写入。
10. **DEV 面板「模拟本机索引丢失」**：只清 `excel-bro.pending-office-execution.v1`，
    探针键原样保留 —— 真机不需要人进 DevTools 删 localStorage。

**"零写入"的适用范围（不许一句话盖全部）**：第 1–4、6–10 条判的是**发现与重核**阶段
（`host.writes()` 为空，且 `registerOfficeExecution` / `startOfficeAction` / `completeOfficeAction` /
`reconcileOfficeActionReceipt` / `finishOfficeExecution` / `suspendOfficeExecution` 均未调用）。
**第 5 条不是零写入**：它真的续跑，判据是"只写 `汇总表!B4` 这一格"——
`contentWrites(host.originalWrites()).slice(before.length)` 必须**恰好**等于 `["B4"]`，
`completeOfficeAction` 的 actionIndex 恰好 `[0, 2]`（第 2 步在中断前已完成，不重放），
续跑前后 `B2/B3` 保持 111/222。

### 4.2 反向变异（`scripts/mutate_office_recovery_discovery.mjs`）

| 变异 | 拆掉什么 | 预期红的用例 | 四判据 |
| --- | --- | --- | --- |
| D1 | `recordConfirmedScope` 不记"用户确认过的那一份"（版本照旧推进） | 第 3 条 | 全成立 |
| D2 | 代数闸恒真（`isCurrent = () => generation >= 0`） | 第 4 条 | 全成立 |
| D3 | 多候选直接采用排序第一条 | 第 1 条 | 全成立 |
| D4 | 清空授权那条提交路径不再记账（只 `commitAuthorizedScope`，不 `recordConfirmedScope([])`） | 第 7 条 | 全成立 |
| ~~D5~~ | ~~续跑不再复核写前快照~~ | ~~第 9 条~~ | **2026-10-29 第十轮退役**，见下 |

四判据 = ① 退出码非 0；② 预期用例名出现在 `FAIL` 里；③ 失败原文带它**自己那处**证据；
④ 还原后逐字节 `sha256` 与基线一致（基线 `08f85f0f…d27716`）。
四个变异一个只打一个判别点：D1 只动"比对哪一份名单"、D2 只动"旧响应能不能落地"、
D3 只动"要不要用户选"、D4 只动"清空算不算一次确认"。**D1–D4 于第十轮全部重跑，
四判据仍逐条成立**（本轮改过 `App.tsx`，SEED 必须重新定点）。

**D5 为什么退役**：它打的 `if (!state.beforeSnapshot)` 已经不在续跑路径上了 ——
"在途动作"现在是**独立结论**（`reconcile`），复核落在 `officeRecoveryVerdictFor` 与
`pendingResumeReadiness`（`in_flight_without_snapshot`），续跑侧只剩"按最新读数重新判定"
这道闸。留着旧 SEED 只会 `SEED 未命中`（退出码 3），那等于假装有证据。等价的保护改由
下面 S4 承担。

### 4.3 反向变异（`scripts/mutate_office_recovery_split.mjs`：核对与续跑真的分开了吗）

| 变异 | 拆掉什么 | 预期红的用例 | 四判据 |
| --- | --- | --- | --- |
| S1 | 核对之后顺手把剩余动作也跑掉（把两步合回一次点击，按钮名不变） | 「A1 三动作中途断开：核对只补记回执（零写入）…」 | 全成立（证据 `A3:B3`） |
| S2 | 续跑忽略服务端进度，从第 1 个动作重来 | 「在途动作先核对为「已写入」：跳过该步…」 | 全成立（证据 `A1:B1`） |
| S3 | 卡片不再把 `backfill` 算成本端能执行的动作 | 「全部动作已写入、只缺最终回执：只给「补记完成（不重写原表）」…」 | 全成立（证据 `补记完成（不重写原表）`） |
| S4 | 续跑前不再按最新读数复核结论 | 「续跑不信任卡面：状态里出现没有写前快照的在途动作时立刻停下」 | 全成立（证据 `在途动作缺少写前快照`） |

D4 是**本轮实测发现并修掉的真缺陷**：`confirmSheetSelection` 的"空授权"分支直接
`commitAuthorizedScope([], "manual")` 就 `return true`，没有 `recordConfirmedScope([])` ——
用户清空授权后卡片仍写着"可继续"，只有点下去才失败。D1 修的是"确认收窄"那条路，
一个修好不代表另一个修好（两条提交路径不同），这正是分成两个变异的原因。

### 4.4 反向变异（`scripts/mutate_office_recovery_placement.mjs`：入口真的搬出来了吗）

| 变异 | 拆掉什么 | 预期红的用例 | 实际 |
| --- | --- | --- | --- |
| **P1**（两步结构搬迁：先占位、再把面板插到消息里） | 入口搬回**消息流**（全页仍只有一张可点卡） | `恢复入口只在输入区上方…` 的**位置**断言（`panel.closest(".message-stream")` 必须为 null）；**唯一性**断言应保持绿 | 红 1 条，失败原文 `expected <section …(2)>…(3)</section> to be null`；唯一性断言（按钮数 1、记录无按钮）在同一轮里通过 |
| **P2**（往只读记录里补一个 `继续剩余动作`） | "历史消息里没有按钮" | 同一条用例的**唯一性**断言 | 红 1 条，`expected [ …(2) ] to have a length of 1 but got 2` |
| **P3**（选择器 `rows.length - 1 → 0`，取第一条） | "新的执行覆盖旧的" | core `取最后一条带卡的记录…` + `只挑一条…` | 红 2 条；App 侧全绿（那里只有一条卡，选谁都一样 ⇒ 判别点必须在 core） |

四判据逐条成立：退出码非 0、预期用例名在 FAIL 里、失败原文带**它自己**那处证据、
还原后逐字节 sha256 与基线一致（`App.tsx 4c3ea8dd…`、`OfficeRecoveryPanel.tsx c21e644b…`、
`executionRecoveryView.ts e17f1a88…`）。

**布局脚本的灵敏度也单独证过**：把 P1 应用上，`check_recovery_entry_viewport.mjs` 立刻在
`320x640: 恢复入口不许住在滚动容器里` 判红；还原后同一脚本四档全绿。

### 4.5 真实浏览器布局证据（`scripts/check_recovery_entry_viewport.mjs`）

本机**没有 Playwright**，Chrome 又是 GUI 子系统程序（Windows 上 `--dump-dom` 的 stdout
拿不到），所以脚本走**原始 CDP**：`--remote-debugging-port` + WebSocket +
`Emulation.setDeviceMetricsOverride` 固定视口 + `Runtime.evaluate` 取矩形。
页面是**真实 `<App/>`**（esbuild 打包真实源码）+ **真实 `styles.css`**（源文件，非副本），
`localStorage` 里种 48 条历史 + 一条带恢复卡的消息（卡挂在第 2 条，后面还有一堆日志 ——
正是真机形态）。**不发请求、不碰工作簿。**

| 档位 | 结果 |
| --- | --- |
| 320x640 | 入口按钮 `top=474.0 / bottom=506.0`（视口 640）；追问日志后 `top` 仍是 474.0；输入框 `top=544.0`（在入口之下、未被压） |
| 400x720 | `top=554.0 / bottom=586.0`；追问后不变；输入框 `top=624.0` |
| 500x760 | `top=594.0 / bottom=626.0`；追问后不变；输入框 `top=664.0` |
| 320x640 + 长原因（阻断，正文 410px 撑满） | 面板限高生效 `height=294.4`（`min(46vh, 320px)`），正文 `410 → 200` 自己滚，动作按钮 `bottom=506.0`，输入框 `top=544.0` |
| **灵敏度对照**（同一张卡塞回第一条消息 + 滚到底） | 按钮 `bottom ≈ -4596.6 / -4477.0 / -3977.0` ⇒ **看不见**，脚本能复现真机缺陷 |

`--shots` 会写出 `outputs/probe/recovery-entry-320x640.png` 等图（含 `-tall-` 那张），
用于人工核对版式；数值断言与图互不替代。

**不算数的部分**：这是 Chrome 的布局引擎，不是 Excel/WPS 的宿主窗格；
`--dump-dom` 那条路在 Windows 上不可用（GUI 子系统没有 stdout），所以脚本才走 CDP。

## 5. 双端差距（**不宣称两端统一**）

| 维度 | Excel（本轮后） | WPS（现状） | 差距性质 |
| --- | --- | --- | --- |
| 后端发现 | `officeRecoveryDiscovery.ts` → `GET /api/office/executions/pending` | `wpsRecoveryDiscovery.ts` → `GET /api/wps/executions/pending` | **已对齐**（判据同一份共享模块） |
| 身份命名空间 | `url:`（未保存簿不发现） | `wps-path:`（小写完整磁盘路径） | 宿主差异，非缺口 |
| 多候选显式选择 | `OfficeRecoveryChoicesCard` | 面板既有选择入口 | 外观各自实现，判据共享 |
| 可续跑判定 | **接入** `pendingResumeReadiness`（含"用户确认过的授权"） | 判定仍在 `apps/wps-addin/src/WpsAgentPanel.tsx` 的 `wpsRecoveryVerdictFor`：身份判定 + **现场核对**（`inspectWpsRecoveryStateDetail`，含 `authorized_sheet_not_allowed` / `action_sheet_not_authorized`）+ **材料级** `prepareWpsResume` | **未统一（口径不同，不是缺口）**：WPS 用"读现场 + 材料指纹"，Excel 用"名单集合 + 快照存在性"，两者不可互相顶替 |
| 授权变化后卡面改判 | **有**：`reverifySelectedOfficeRecovery`（只读、同一条、同一判据） | **有**：`reverifySelectedWpsRecovery`（2026-10-28 修，`scripts/mutate_wps_recovery_reverify.mjs` R1/R2 四判据已钉；用例「授权从三张表收回成原两张后，同一条已选执行重新核验并改判」） | **已对齐**（实现各一份，判据来源不同） |
| 迟到的响应 | 代数闸 `officeDiscoveryGenerationRef` | 代数闸（同一模式） | **已对齐** |

> **本表 2026-10-29 更正一处**：本文早前版本把"授权收窄后卡面不改判"写成 **WPS 侧缺口**，那是
> **错的**，也与 `计划表.md` 2026-10-28（第二轮）的记录自相矛盾。事实是：WPS 早在 2026-10-28 就修好了
> 这条（用户取消多勾的一张表 ⇒ `authorizationKey` 变化 ⇒ 执行 `reverifySelectedWpsRecovery` ⇒
> 卡面重新改判），并且有反向变异钉住（R1 = 回到"只保护不重算"必须红；R2 = 迟到响应不许覆盖新现场）。
> 当时的缺陷现场是"`discoverWpsRecovery` 在 `selected` 分支直接 `return`"，已经拆掉。
> 两端真正的差别只在**判据来源**（上一行），不在"有没有重核"。

**为什么不在这一轮把 WPS 也切到 `pendingResumeReadiness`**：WPS 的判定必须叠一层真实宿主材料校验
（`prepareWpsResume`），那需要真 `WpsApplication`；在自动化里用替身把它抽象掉，等于把已经
在真机上验证过的材料级判据换成一条没在真机验证过的路径。这属于"重写已通过的 WPS 链路"，
本轮明确不做。所以只交付 **Excel 侧**，两端差异逐条记在上表。

## 5.1 恢复交互：本轮对齐了什么、还差什么（2026-10-29 第十轮）

上一轮只做到"能发现"，卡面还是 Excel 自己的一套：一个「核对并继续」把**补记回执**与
**执行剩余动作**合成一次点击，正文还说两遍同一件事。真机截图暴露的正是这一点。本轮把
**结论与措辞**搬到共享层（`@excel-bro/core/executionRecoveryView`），两端读同一份。

### 已对齐（同一份判据 / 同一份文案 / 同一种信息层级）

| 项 | 共享出处 | Excel | WPS |
| --- | --- | --- | --- |
| 六种恢复结论（已完成 / 待核对在途动作 / 可继续剩余动作 / 需重新规划 / 需要人工核对 / 证据不足） | `ExecutionRecoveryVerdict` | `officeRecoveryVerdictFor` 产出 | `wpsReconcileVerdict` 产出 |
| 结论标题（badge） | `EXECUTION_RECOVERY_STATE_LABEL` / `executionRecoveryBadge` | 同 | 同 |
| 一句正文（第几步已写入 / 还剩几步没跑） | `executionRecoveryDetail` | 同 | 同 |
| 动作集合（按钮只可能来自这里） | `executionRecoveryActions` | 过滤后使用 | 直接使用 |
| 按钮文案 | `EXECUTION_RECOVERY_ACTION_LABEL` | 同 | 同 |
| 卡片标题 / 固定边界说明 | `EXECUTION_RECOVERY_TITLE` / `EXECUTION_RECOVERY_NOTE` | 同 | 同 |
| **核对 ≠ 续跑** | 两种互斥结论 + 两个按钮 | 「核对在途动作」只补记回执（零单元格写入）、「继续剩余动作」只跑没跑过的动作 | 同（2026-10-28 起已是这个形状） |
| 信息层级 | 标题 + badge + 一句正文 + 动作区 + 折叠的「执行记录」 | 同 | 同 |

**"分开"落在执行路径上，不是换了个按钮名**：`reconcileOfficeInFlightAction` 走
"读现场 → 提交回执 → 按新结论重画卡片"，全程不调用 `executePlan`；`resumeOfficeRemainingActions`
从 `nextActionIndex` 起跑；`backfillOfficeCompletion` 只 claim + `complete`（`result` 写
`{backfilled:true, actionCount}`），也不调用 `executePlan`。反向变异 S1/S2/S3 分别钉住
"核对不写单元格 / 续跑只写剩余 / 只剩回执时不执行动作"。

### 仍有宿主差异（**不宣称两端完全一致**）

| 差异 | Excel | WPS | 为什么保留 |
| --- | --- | --- | --- |
| 在途动作的现场结论何时可知 | 发现阶段只读服务端状态 ⇒ `outcome: "pending"`（"可能在停止前已经写入，也可能没有"）；点核对后服务端回执给出 applied / not_applied | 判定当场回读宿主 ⇒ 直接给 applied / not_applied | Excel 的发现链路刻意不读单元格（§2.5 只读边界）；把读数挪进发现会让每次挂载都扫一遍目标区域 |
| 「重新读取并规划」 | **不提供**（结论会落到 `blocked` 并说明原因） | 有（`replan`，重走一轮任务） | Excel 没有"用原需求就地重规划"的入口，要另起一轮对话；给假按钮更糟 |
| 「清除这条恢复提示」 | **不提供** | 有（`dismiss`，写服务端处置元数据） | `/api/office/executions` 没有处置端点；只清本地会让提示重开又回来 |
| 兜底入口 | 无可执行动作时给「重新核对（只读）」 | 给 `dismiss` + 「重新核对（只读）」 | 同上 |
| 执行记录详情 | 折叠行显示"已完成 N / M 步 · 在途第 k 步 / 剩余 n 步" | 折叠 `<details>` 放 DEV 原因码 | 各自暴露自己能确定的字段 |
| 入口**位置** | 输入区上方固定区域（`.office-recovery-panel`），消息里只留只读记录 | 消息流**末尾**（`messages.map` 之后）的 `.wps-write-recovery` | 见 §5.2：WPS 的末尾位置天然躲开了"自动跟随把卡顶出视口"，Excel 的"跟着某条消息"躲不开。两端都**没有**在实机上量过视口 |

> 上一轮"发现入口接线完成"的说法**不覆盖**交互对齐：本轮之前，Excel 的卡面确实还是一个按钮
> 合并两步。真机 ①–⑥ 仍未跑（§8），所以本文件不宣称恢复交互已在真实 Excel 上验收。

## 5.2 恢复入口的**位置**：为什么卡片不在消息里了（2026-10-29 第十一轮）

### 真机现象（用户截图 + 文字，两条合起来才是事实）

点「核对在途动作」→ 聊天区自动滚到底 → 显示"已核对在途动作，未写单元格"→ 更新成
「继续剩余动作」的卡**还留在上面那条旧消息的位置**，视口里看不到下一步入口；
`汇总表` 仍是 111/222/-303（核对确实没写单元格）。
**截图只证明入口不在视口**，不证明卡片状态没更新 —— 所以本轮同时查了状态与布局。

### 根因（两处叠加，缺一不成）

1. **卡片住在消息流里**：恢复卡原来是 `message.officeResume` 的一部分，画在**产生它的
   那条消息**下面（`App.tsx` 消息渲染分支）。
2. **消息区会自动跟随**：`useMessageAutoFollowRuntime.tsx` 的 `useLayoutEffect` 在
   `signal`/`conversationId` 变化且"贴底"时执行 `container.scrollTop = container.scrollHeight`；
   核对成功会 `beginActivity` 追加日志 ⇒ 信号变化 ⇒ 滚到底 ⇒ 旧消息里的卡被顶出视口。

两者单独都不致命：卡不在流里，滚动就与它无关；消息区不自动跟随，卡就还在原地。
**本轮的修法是去掉第 1 条，而不是给第 2 条打补丁** —— "再滚一次""定时器纠正"
`scrollIntoView` 都会打断用户上翻历史，属于用症状掩盖问题。

### 改动

| 位置 | 改动 |
| --- | --- |
| `packages/core/src/executionRecoveryView.ts` | 新增 `selectExecutionRecoveryEntry(rows, recoveryOf)`：**只挑一条**（最后一条带卡的记录；`operationId` 空白跳过；没有就返回 `null`）。新增 `executionRecoveryProgress`（"已完成 N / M 步 · 在途第 k 步 / 剩余 n 步"）与 `EXECUTION_RECOVERY_ENTRY_HINT` |
| `apps/excel-addin/src/components/OfficeRecoveryPanel.tsx`（新） | `OfficeRecoveryPanel`：**唯一可点入口**。结构三层 —— 标题 + badge / 可滚动的正文 / 动作区，动作区是正文的**兄弟**（正文再长也只自己滚）；`data-operation-id` / `data-recovery-kind` / `data-recovery-acting` 供判据读取。`OfficeRecoveryRecord`：历史消息里的**只读**记录（无按钮、无 badge、非 `role="status"`），仅当它就是当前入口时多一句「操作入口在输入框上方」 |
| `apps/excel-addin/src/App.tsx` | 恢复卡整块移出消息；输入区上方（`OfficeRecoveryChoicesCard` 之后、`.composer-input-region` 之前）渲染 `OfficeRecoveryPanel`；消息里改渲染 `OfficeRecoveryRecord`；`officeRecoveryEntry` 由共享选择器派生（**不存 ref**，换会话/换簿自然跟着变） |
| `apps/excel-addin/src/styles.css` | 新增 `.office-recovery-panel`（column flex + `max-height: min(46vh, 320px)`）、`-body`（`flex:1; min-height:0; overflow-y:auto`）、`-actions`（兄弟层，`flex:0 0 auto`）、`.office-recovery-record`（只读记录）；删除已无用的 `.office-resume-card` |
| WPS | **未改**。WPS 的 `.wps-write-recovery` 画在消息流**末尾**（`messages.map` 之后），追加日志只会把它继续往下推、自动跟随正好落在它上面，天然没有这个问题。按要求不为形式一致重写它已验证的链路 |

**状态一致更新**：核对中（`acting: "reconcile"` → 按钮变「正在核对…」且禁用）、阻断、
证据不足、完成（无动作时只给「重新核对（只读）」）都是**同一张卡就地换内容**，
判据是 `data-recovery-kind` / `data-recovery-acting` 与按钮文案；不是"再插一张新卡"。

### 边界**没有**被这次搬迁改变

核对只补记回执（零单元格写入）、继续只跑剩余动作、确认变更不重放、不自动继续、
不扩大授权、不新增执行前确认 —— 这些判据仍在 `excelWriteEndToEnd.test.tsx` 的原用例里，
本轮一条都没有放宽（`A1 三动作中途断开` 三种 outcome 仍绿；`mutate_office_recovery_split.mjs`
的 S1–S4 仍四判据成立）。

### 剩余限制

- 布局证据来自**真实 Chrome 布局引擎**（`scripts/check_recovery_entry_viewport.mjs`），
  **不是** Excel/WPS 的 WebView2/CEF 窗格：宿主窗格的字号、缩放、滚动条宽度可能让数值不同。
  比值判据（在视口内 / 位置不变 / 不压输入框）应当稳健，但**实机截图仍未取**。
- `max-height` 上限按 46vh/320px 取；若将来卡片塞进更多内容，需要重新量（脚本里已有
  320x640 + 长原因的档位，`--shots` 出图）。

## 6. 宿主限制（真实存在，不当缺陷）

1. **未保存的 Excel 工作簿不参与发现**：`document.url` 为空 ⇒ 没有 `url:` 身份。界面不摆候选、
   不报错 —— "识别不到就什么都不说"。已保存后自动参与（身份观察会重跑发现）。
2. **服务不可达 ≠ 没有未完成任务**：显式说"暂时读不到"，卡片保留。
3. **`limit` 截断**：候选被截断时给截断说明，不假装"就这么多"。
4. **发现只能找回入口**：执行权（`ownerActive`）仍由服务端决定；发现在飞期间用户点了别的簿/改了范围，
   旧结论作废（代数闸）。

## 7. 真机验收（**只跑一条最小链路**，一次做完再回传）

前提：`npm run build:addin`；服务端已启动；Excel 已加载加载项；**DEV 构建**（或
`VITE_ENABLE_RECOVERY_PROBE=true`），面板底部能看到「恢复验收（开发专用）」。

**不碰的三件事**：不手动暂停（用探针的确定性中断）、不手删 `localStorage`（用面板上的
「模拟本机索引丢失」）、不停真实服务（服务不可达那条分支已由自动化覆盖）。

### 7.0 启动（三个终端，都在仓库根目录 `E:\codex\excel-bro`）

```
# ① 后端：127.0.0.1:8765（端口被占用时直接报错，不会静默换端口）
npm run dev:server

# ② 前端任务窗格：vite → https://localhost:3000（DEV 构建，DEV 面板才可见）
npm run dev:addin

# ③ 把加载项装进桌面 Excel（读 manifest.xml，打开 Excel 并加载窗格）
npm run start:excel
# 收工：npm run stop:excel
```

`dev:server` 的等价命令（同一份配置）：
`.venv\Scripts\python.exe -m uvicorn server.app.main:app --host 127.0.0.1 --port 8765 --loop asyncio --workers 1`。

### 7.1 夹具（可重复）

```
.venv/Scripts/python.exe scripts/make_excel_recovery_fixture.py
```

产物 `outputs/excel-recovery-fixture-20261029/Excel恢复夹具-20261029.xlsx`：
`综合票房`（B2:B4 = 111/222/333）、`汇总表`（B2:B4 = -101/-202/-303）、`备用表`。
脚本同时打印**本次验收要用的那一句话**与逐格期望。**已存在的夹具不覆盖**（上一次验收
可能改过单元格），需要重建时显式 `--force`。

### 7.2 步骤（一条链路：重开发现 → 核对 → 只续剩余动作 → 审阅收尾）

| 步 | 操作 | 期望（判据） |
| --- | --- | --- |
| S0 | 打开夹具 → 授权 `综合票房` + `汇总表` → 展开「恢复验收（开发专用）」→ 点「启用下一次第二步中断」 | 状态行变成「已启用：仅作用于当前工作簿的下一次执行」 |
| S1 | 发夹具打印的那一句（"把「综合票房」的 B2、B3、B4 三格值分别写到「汇总表」的 B2、B3、B4"） | 计划是 **3 个独立单格 `writeValues`**；若模型给了别的形状，探针会 fail-closed 并明说"需要三个独立单格值写入"，此时重发一句即可（不要改夹具） |
| S2 | **什么都不用点**：计划到达后按既定流程在授权范围内校验一次并**自动执行一次**（没有"执行前确认"这一步；这张卡上也不该出现「执行到 Excel」——那个按钮只在计划不满足自动执行条件时才出现，正常现场**看不到**它） | 面板报"恢复验收：第二步已写入并回读，已在完成回执前中断"；`汇总表` = 111 / 222 / **-303**（第 3 步未执行）；状态行「验收中断已触发」 |
| S3 | 点「模拟本机索引丢失」（**不要**去 DevTools 删 localStorage） | 状态行「已清空本机挂接记录（账本未改）」；此时不要求界面有任何变化 |
| S4 | 关掉任务窗格，重新打开 | **靠后端发现摆出恢复卡**：标题「上次写入的恢复核对」+ badge「待核对在途动作」，正文写清"第 2 个动作可能已写入"与"还有 1 个动作没执行"；按钮**只有**「核对在途动作」一个（Excel 没有服务端的"处置/清除"端点，所以**不画**「清除这条恢复提示」——WPS 才有；画一个点了只在本窗口消失、重开又回来的按钮等于骗人），**没有**「继续剩余动作」；**不得**显示"没有未完成任务"；此时单元格仍是 111/222/-303，`汇总表!B4` 仍是 -303 |
| S4a | 点「核对在途动作」 | **零单元格写入**（`汇总表` 仍是 111/222/-303，一格都没变）；badge 换成「可继续剩余动作」，正文写"可继续剩余 1 个动作"，按钮变成「继续剩余动作」 |
| S5 | 点「继续剩余动作」 | 只写第 3 步：`汇总表!B4` 变成 333，B2/B3 保持 111/222（第 2 步不重放）；对话出现"Office 任务已从第 3 步继续并完成" |
| S6 | 审阅收尾：结果卡看逐格回读 →「确认变更」 | 显示已确认变更、审阅条不再出现"执行未完成"；`汇总表` = 111/222/333；撤销入口只在结果卡内。**「确认变更」只清审阅标记，不再写一次**（第二次写入若发生就是缺陷） |

任何一格不符就**停在该步**回传（附截图 + 该步之前的账本 operationId 尾部即可），不要继续往后跑。

**两条流程口径（照既有契约，不新增步骤）**：

1. **没有"执行前确认"这一步**。计划在授权范围内通过计划校验、来源复核与恢复准备后**执行一次**，
   回读并展示结果供审阅；`确认变更` 只接受结果、清掉审阅标记，**不再执行**
   （`docs/ARCHITECTURE.md:152`：「不得增加执行前确认按钮」）。所以 S2 不需要点任何按钮，
   S6 的「确认变更」也不产生第二次写入。给这条链路加"执行前确认"就是回归。
   **例外与自检**：`执行到 Excel` 那个按钮仍在代码里，但只在计划**不满足**自动执行条件时渲染
   （例如工作目录来源、删除表、或计划本身存在审阅阻断原因）。如果 S1 之后
   看见了它，说明计划形状或来源复核没成立 —— 点它不算"照流程走"，应先让模型重给三格写入。
2. **"零写入"只用于发现与重核阶段**（S3、S4 以及 §7.3 里那些分支）。S5 是**真的续写**：
   它必须写出 `汇总表!B4 = 333` 这一格，判据是"只写这一格"（B2/B3 保持 111/222、第 2 步不重放）。

### 7.3 其余分支：**已自动化，不要求真机手点**

| 分支 | 自动化判据 |
| --- | --- |
| 授权收窄 / 撤销后卡面阻断 | 第 3 条、第 7 条（D1/D4 变异钉住） |
| 从没确认过范围（不误判成撤销） | 第 6 条 |
| 换到另一本工作簿再切回 | 第 8 条 |
| 续跑前重新核验（无写前快照立刻停） | 第 9 条（`mutate_office_recovery_split.mjs` S4 钉住） |
| 服务不可达说"暂时读不到"、不清卡 | `officeRecoveryDiscovery.test.ts` 与 `pendingExecutionDiscovery.test.ts` |
| 多条候选不自动选、不查状态 | 第 1 条、第 3 条 |
| 迟到响应不覆盖新现场 | 第 4 条（D2 变异钉住） |

## 8. 未做 / 未验 / 下一步

- **真机进度（2026-10-29 更新，按证据分档，别合并宣称）**：
  - **S4 → S4a → S5 → S6 有证据**：用户重开后看到「发现未完成执行」→ 核对后仍是 `111/222/-303`
    → 继续后 `111/222/333` → 确认后恢复卡消失。账本侧支撑到"三步各写一次、2 号动作晚 2h33m、
    `next=3`、`in_flight=NULL`、逐格回读一致"（`execution-c52ad791…`，
    `outputs/a1-acceptance-20261029-2/EVIDENCE.md`）。
  - **S0–S3 没有证据**：不知道有没有真的点「模拟本机索引丢失」、也没记探针启用那一刻的界面。
  - **「靠后端发现」这一格仍是"证据不足"**：账本没有"发现"记录，`discovery_dismissed_at` 仍是
    `NULL`（从未清过提示），界面只有口述。**不得因为恢复卡出现了就标通过**。
  - **磁盘终值不成立**：夹具 `.xlsx` 的 mtime 仍是 14:22（早于 16:03 的执行）、`汇总表 B2:B4`
    仍是 `-101/-202/-303` ⇒ 内存已改、**从未保存**。"保存落盘"一档不标通过，也不要求用户补做。
- **入口位置**：真实浏览器布局已量化（`scripts/check_recovery_entry_viewport.mjs`，320/400/500 +
  高卡片档）；真实 Excel 窗格**这一次有口述**（"固定入口可见"）但**截图字节未落盘**，
  所以只算"口述通过"，WPS 侧仍未量过。
- **恢复交互对齐只到自动化**（§5.1）：Excel 的「核对在途动作 / 继续剩余动作 / 补记完成」
  三分支有真实 `<App/>` + 替身宿主 + 反向变异证据，但**真实 Excel 窗格里一次都没点过**；
  真实 WPS 侧本轮未改动，也没有重跑真机。所以"两端恢复交互已一致"这句话现在**不能说**，
  只能说"结论、措辞、按钮集合与信息层级读同一份共享层；执行路径各自实现"。
- **WPS 侧未接入** `pendingResumeReadiness`（判定口径不同，见 §5；WPS 的授权变化重核
  在 2026-10-28 已修，本轮**更正**了早前那处错误结论）。本轮不重写 WPS 链路。
- 桌面/在线共享链接（OneDrive/SharePoint）身份归一化、`unsaved:` → 保存后自动重挂接：未做。
- **A2 不启动**；未提交、未推送。
