# A1 验收剩余项（历史批次＋2026-11-01 校正）

> 当前排期以 [主线计划 §0](SDK_MAINLINE_PLAN_2026-10-19.md#0-当前收口队列2026-11-01优先于下方历史批次状态) 为准。下方 10-28/29 收盘描述保留历史时点；公式已实现、Excel 公式 A/B 与撤销已有原生证据，**WPS 公式 A 也已于 2026-11-01 真机通过**（写序证据齐全，见 UX 台账 §8.5 F），不能再整体标未做。WPS 公式 **B 仍标未验**。

> 这份文件的唯一用途：把 A1 还没闭环的东西**按"为什么没闭环"分开写**，避免三类东西混成一句
> "待实机"。
>
> 三类定义（**不许互换**）：
>
> - **尚未实现**：代码/协议里根本没有这条能力。再多的实机操作也验不出来。
> - **自动化已验但真机未验**：判据有、用例会红会绿、替身宿主跑过，但真实 Excel / 真实 WPS
>   里一次都没点过。
> - **证据待补**：真机上确实发生过，但某一环没有可核实的记录（或只有截图/口述），
>   差的是**取证**，不是功能。
>
> **A1-GATE 结论：不通过。A2 未启动。** 上一行的判定依据是
> `SDK_MAINLINE_PLAN_2026-10-19.md` §4 第 4 行的验收条件（九场景逐项结论明确；原生与自动化
> 证据分别齐全；支持范围明确；关键未闭环项解决或获批分批交付；相关全量检查通过），
> 目前补过三轮：#12 上「收尾之后还能继续对话」的真机记录、**#11/#12 的跨进程（正常退出重开）版**、
> 以及 2026-10-29 的 **Excel 端「中断 → 重开发现 → 核对 → 只续第三步 → 确认」**。
> 第三轮把 Excel 的 #10 与"恢复入口位置"两格从空变成有证据，其余档位结论未变：
> 崩溃版恢复、真实撤销、S1、以及"本机索引丢失后由后端发现"仍然空着（后者**有界面口述但证据不足**，
> 见 §0.1）。本轮另补了一项**尚未实现**的能力：U4 处置留痕（见 §1.1）。

---

## 0.1 2026-10-29 第三轮补了什么（Excel，真机）

用 `outputs/excel-recovery-fixture-20261029/` 的三格夹具，唯一一条记录
`execution-c52ad791-6eb8-47c7-99d0-69132b70ef25`（`task_id turn-2c05a9ce…`）。
完整证据分档见 **`outputs/a1-acceptance-20261029-2/EVIDENCE.md`**（含账本快照
`ledger-snapshot.json` 与只读磁盘读数）。

| 项 | 之前 | 之后 |
| --- | --- | --- |
| Excel「第二步中断后只续第三步」（§2.1 #10） | ⏳ 一格全空 | **账本通过**：`office_host_writes` 0/1 同秒（16:03:44）、2 在 18:37:11，相隔 2h33m27s；`next_action_index=3`、`in_flight=NULL`、`checkpoint.uncertainActionIndexes=[]` |
| Excel「核对只补记、零单元格写入」 | ⏳ | **账本推断通过**（非直接记录）：核对不写 `office_host_writes`，且中间没有第三个写入时刻 |
| Excel「恢复入口在固定位置、没被日志挤走」 | 只到真实浏览器布局 | **口述通过**（用户报告点了核对后入口仍可见）；**截图字节未落盘**，没有第二份可核实记录 |
| Excel「确认变更后恢复卡消失」 | ⏳ | **账本通过**（`status=completed`、`in_flight=NULL`、`next=3`）+ 口述 |
| **保存落盘后的终值** | 每轮都卡在"只有截图" | **未发生**：夹具 `.xlsx` mtime 仍是 14:22（早于 16:03 的执行）、sha256 `95eec1f1…`，`汇总表 B2:B4` 仍是 `-101/-202/-303`；`EXCEL.EXE` 仍在运行 ⇒ 内存已改、从未保存。**不标通过** |
| **「本机索引丢失后由后端发现」** | ⏳ | **未通过（证据不足）**：账本没有"发现"记录，`discovery_dismissed_at` 仍是 `NULL`（从未清过提示）；只有界面口述。**不得因为恢复卡出现了就标通过** |
| 用户点了「确认变更」在服务端有没有记录 | 没查过 | **没有**（U4 缺口实证）⇒ 本轮已实现 U4（§1.1），但**不追溯**这一轮 |

---

## 0. 2026-10-28 两轮补了什么

| 项 | 之前 | 之后 |
| --- | --- | --- |
| 「放弃变更之后能不能继续对话」（WPS，同会话） | 只有自动化 | **真机通过**（`5a2d84be` 中断+续跑 → `turn-68fd1b…` 正常出结果） |
| 「确认变更之后能不能继续对话」（WPS，同会话） | 只有自动化 | **真机通过，但只覆盖"普通确认执行"**（`7afd3dc6` 三步同一毫秒写完 → `turn-4ec317…` 正常出结果，账本无中断记录） |
| 「续跑后确认 + 随后对话」 | 无 | **仍缺**（前半有 `1fd938fd` 实机，后半只有自动化；同一轮真机没有） |
| 「查询 `B2:B4` 却提示扫描不完整」 | 未记 | **记为独立待查项**（§4，未修复） |

证据：`outputs/a1e3-native/abandon-20261028-followup/`（`README.md` 里有四类证据的能力边界表），
记账口径见 `ABANDON_AFTER_RESUME_ACCEPTANCE_2026-10-28.md` §15。

**同日夜间还做了一轮跨进程（正常退出重开）**，`wps-execution-3f00ea0d…`，
证据目录 `outputs/a1e3-native/cross-process-abandon-20261028/`，记录见同文档 §16。它带来三处变化：

| 项 | 之前 | 之后 |
| --- | --- | --- |
| WPS 跨进程「中断→退出→重开→核对→续跑→放弃」 | 一次都没做（本文件 §5 原推荐项） | **真机通过**（材料跨会话生效、补记不重写、放弃回基线） |
| 「落盘后的最终值」 | 每轮都卡在"文件从未保存，只有截图" | **首次有磁盘证据**：与原始副本逐格 0 差异 + 条件格式规则 0 条 |
| 首次重开的「证据不足」 | 未记 | **独立待查项**（§6）：已加 28 个原因码，真机跑过 1 次但**未复现**（首条为成功路径、无码） |

## 0.2 三类剩余项总表（2026-10-30 重整）

> 这一节替代"剩下的都是实机"。**三类不许混**：①代码里没有 ②要人拍板 ③代码有、自动化绿、
> 只差在真实宿主里跑一次。谁把它们写成一句"待实机"，就等于把"没写"说成"没验"。

### A. 尚未实现（现在写也验不出来）

| # | 项 | 说明 | 为什么本轮没做 |
| --- | --- | --- | --- |
| U1 | 跨进程动作恢复范围 | writeFormulas 双端已实现（Excel A/B 原生已过；**WPS A 于 11-01 真机已过、B 未验**）；copyRange 的 values 展开已有只续第三格用例 | 格式与结构恢复仍需实现；各子集分开结论，见主线 §0 |
| U2 | 缺来源快照的 SDK Python 结果计划的精确续写 | 同上，明确列为不支持 | 没有写前矩阵就无法构造撤销点；与 U1 同一根 |
| U7 | ~~旧"先复制再写入"死代码（`WpsAgentPanel.tsx` 的 `stagePendingPlan`）~~ | **2026-10-30 已删**（46 行 + 3 处孤立 import；见 §1.3）。`stageWpsNativeComparison` / `fixedValueFallbackFromFormula` **保留** —— 它们有自己的直测，删不删是另一个决定 | 已闭环，不再排轮 |

### B. 待产品/隐私决定（不是测出来的，得有人拍板）

| # | 项 | 卡在哪 | 现状 |
| --- | --- | --- | --- |
| U3 | SDK 原生会话持久化（`--no-session-persistence` 仍开） | 「谁能读历史转录」的隐私边界未定；它与 `resume` 互斥 | **2026-10-30 已核实**（§1.4、`docs/SDK_SESSION_PRIVACY_2026-10-30.md`）：本机 30 份转录、612 行、**全部早于开关（2026-10-07 `a8ae64cb`）**，产品代码不读它们，但目录 ACL 不是产品保证。**仍不开**，等三个决定 |
| U6 | Excel 侧的撤销说明卡与恢复收尾卡（`ChangeReviewUndoNote` / `settleRecoveredPlan`） | Excel 面板没有 `pendingPlan` 闸门，要不要加是产品决定 | **2026-10-30 第十七轮复核**（§1.5）：`pendingPlan` 不是"缺的闸门"而是**架构差异**（Excel 计划到达即执行）；原记两处"用户可见缺口"（撤销范围说明 / 恢复后旧卡标已执行）**均已用 Excel 自己的方式闭环**（`App.tsx:11721` 撤销范围文案 + `officeResume: undefined` 等价收尾），已补证据，未搬 WPS 结构 |
| — | Excel「备份与恢复」抽屉里要不要再放一个**值写入**撤销快捷入口 | 会破坏"一个执行只有一个可点入口"（`selectExecutionRecoveryEntry`） | 刻意没加；WPS 是另一种取舍（两处共用 `executionId`） |

### C. 自动化已验、真机待验（代码齐、用例绿，只差在真实宿主里点一次）

| # | 项 | 自动化证据 | 真机缺口 |
| --- | --- | --- | --- |
| C1 | 处置留痕 | 两端已有实现与自动化 | Excel 10-31 两条 undone_at 已有取值；**WPS 仍是待验**——11-01 公式 A 轮 `undone_at=NULL`，但该轮只到「确认变更」、用户没点撤销，既不算留痕通过也不否定用户操作。不追溯旧轮 |
| C2 | **U5 材料回收留痕**（本轮新增） | 12 条 pytest + R1–R5 变异 + 服务端全量 | 真实 WPS 里撤销一次后，`recovery_package_*` 四列有没有值 |
| C3 | 保存落盘后的终值 | 既有只读核验 | Excel 公式 A/B 已有保存文件证据；**WPS 公式 A 于 11-01 补上保存文件证据**（`公式来源证据验收-WPS-A.xlsx`，三格公式文本 + 缓存值 2/8/15）；WPS 值放弃回基线也已有。10-29 Excel 未保存仅是那一轮事实 |
| C4 | Excel 结果卡/重开撤销入口 | 组件自动化及用户操作 | 10-31 已点击成功，见 UX 台账 §8.5 B/C；截图待定位，不再标从未点击 |
| C5 | S1 来源失效停止 | 注入脚本与账本 | 10-26 双端注入原生已做，见 SDK_A1_NATIVE_TOOLING_REPAIR；不扩大为完整手动改源交错矩阵 |
| C6 | 崩溃/被杀进程版恢复 | 正常退出版已过（`3f00ea0d`） | 杀进程版空着，不能用正常退出代替 |
| C7 | 双窗格/换簿/同名簿/授权收窄视觉/并排布局/删表撤销/格式动作 | 各有一批面板级用例 | 见 §2.1–§2.4 逐格 |

**"后端发现"（C 类的特殊一格，属"证据待补"）**：行为有界面口述，账本无"发现"记录、
`discovery_dismissed_at` 仍是 `NULL`。取证方式见 §5.1 —— **不要**为了让提示出现去写只读发现、
**更不要删恢复材料**去模拟索引丢失。

## 1. 尚未实现（验不出来，得先写）

| # | 项 | 现状 | 为什么不能靠"实机再跑一次"解决 |
| --- | --- | --- | --- |
| U1 | 跨进程动作恢复范围 | writeFormulas 双端已实现（Excel A/B 原生已过；**WPS A 于 11-01 真机已过、B 未验**）；copyRange 的 values 展开已有只续第三格用例 | 格式与结构恢复仍需实现；各子集分开结论，见主线 §0 |
| U2 | **缺少来源快照的 SDK Python 结果计划**的精确续写 | 同上，明确列为不支持。卡点：`localResultRef` 的计划**强制**要来源锚点（867-868），而它没有 `sourceDependencySheets` ⇒ 材料为 `null` | 表面是"补个 context 变体"，实际是"结果引用能不能充当来源证据"—— 动的是信任模型，要产品点头 |
| U3 | **SDK 原生会话持久化**（`--no-session-persistence` 仍开） | 未接入；与 `resume` 互斥，隐私边界未定 | 需要一个"谁能读历史转录"的决定，不是测出来的 |
| U4 | ~~**放弃/撤销的服务端留痕**~~ | **2026-10-29 已实现**（见 §1.1）：`abandoned_at` / `undone_at` + 两端 `disposition` 端点 + 共享判据。**Excel undone_at 已有原生账本证据；WPS 新版留痕未验**，且**不追溯**已有记录 | 原样只读记录：`None` = "没有留痕"，不是"用户没放弃过" |
| U5 | ~~**删除前的恢复材料留痕**（哈希/快照）~~ | **2026-10-30 已实现**（见 §1.2）：DELETE 端点改成「取指纹 → 落账 → 才删」，删成功后再记一次「删除确认」；`office_executions` 加 `recovery_package_digest` / `recovery_package_bytes` / `recovery_package_retired_at` / `recovery_package_removed_at`（`user_version` 12 → 13）；落账失败**不删**（503，材料留到期）。**`retired_at` 只证明"记下了删前摘要"，`removed_at` 才证明"文件确已删除"** —— 不许只凭前者宣称删除成功。**实机未验**，且**不追溯**上线前删掉的材料。详见 `docs/RECOVERY_PACKAGE_RETIREMENT_2026-10-30.md` | 原样只读记录：`NULL` = 没有留痕，不是"材料还在" |
| U6 | **Excel 侧的撤销说明卡与恢复收尾卡** | ~~`ChangeReviewUndoNote` / `settleRecoveredPlan` 只有 WPS 面板渲染；Excel 面板没有 `pendingPlan` 闸门，也没有这张说明~~ **2026-10-30 第十七轮已闭环**（§1.5）：撤销范围说明用 `App.tsx:11721` 文案、恢复收尾用 `officeResume: undefined` 等价接线，均已有测试 | 不是"没验"，是没接线；要不要加是产品决定，本轮不擅自加 |
| U7 | ~~旧"先复制再写入"死代码（`WpsAgentPanel.tsx` 的 `stagePendingPlan`）~~ | **2026-10-30 已删**（见 §1.3）。附带删掉 5 条**永远不会失败**的负向断言，保留有直测的助手函数 | 已闭环 |

### 1.1 U4 本轮实现（2026-10-29）

设计、边界与验证见 **`docs/EXECUTION_DISPOSITION_2026-10-29.md`**。要点：

- `office_executions` 加 `abandoned_at` / `undone_at`（`user_version` 11 → 12），
  **首次写赢**、**不推进 `revision`**、不改 `status`、不动游标、不删材料；
- `POST /api/office/executions/{id}/disposition` 与 `POST /api/wps/executions/{id}/disposition`，
  身份前缀不符或不存在的记录一律 404（不补建账本）；
- 判据唯一出处 `packages/core/src/executionDisposition.ts`（两端只按引用转发）；
- 接线各一处：Excel `useUndoSnapshot.onRestored`（放弃由 `restore:` 预置 ref 判定）、
  WPS `undoExecution` 已有的 `origin`。

**三条边界，别读错**：①不追溯（`execution-c52ad791…` 那一轮的确认永远没有时间戳，E3 对它仍是推断）；
②不记「确认变更」；③~~Excel 的 `undone` 不可达~~ —— **2026-10-29 已证伪**：Excel 结果卡上的
「↶ 撤销本次执行」就是**值写入执行**的撤销入口（走 `useUndoSnapshot.undoExecutionTo` →
`await undoExecution` → `onRestored`），确认之后撤销写 `undone` 有用例钉住；删表检查点是
**另一套**（`restoreWorksheetCheckpoint`，原地撤销删表，不写这一列）。两端现在**四个组合都有入口**
（Excel 2 条、WPS 3 条用例）。完整调用链、支持范围与阻断条件见该文档 §2.4.1。

### 1.2 U5 本轮实现（2026-10-30）

设计、边界与验证见 **`docs/RECOVERY_PACKAGE_RETIREMENT_2026-10-30.md`**。要点：

- `DELETE /api/execution-recovery/{id}` 的顺序改成**核身份 → 取指纹 → 落账 → 才删 → 再记删除确认**；
- **两个不同的事实，不许混成一句"删掉了"**：
  `retired_at`（+ 指纹、字节数）= 删前记下"要删的是这一份"，**只是记录**；
  `removed_at` = `remove()` 报告真的删掉了之后才写，**才是删除确认**。
  于是"落账成功、删除失败"在账本里读得出来（前者有值、后者为空）—— 读这一行的人
  **不许只凭 `retired_at` 宣称材料已删**。
- `office_executions` 加 `recovery_package_digest` / `recovery_package_bytes` /
  `recovery_package_retired_at` / `recovery_package_removed_at`（`user_version` 12 → 13，
  四列同时进 `_REQUIRED_OFFICE_COLUMNS`，旧库逐个 `ALTER`）；两组时刻都**首次写赢**；
  **只存指纹，不存材料正文**；
- 落账失败 ⇒ **不删**（503 `RECOVERY_PACKAGE_RETIREMENT_UNRECORDED`）；
  删了但确认记不下 ⇒ 如实 503 `RECOVERY_PACKAGE_REMOVAL_UNRECORDED`（不假装 204，也不假装没删）；
- 指纹取的是**文件字节**的 sha256（不用 header 里那个内容摘要）：终态回收时材料可能已经过期
  或损坏，那正是最需要留痕的时刻；
- **与 U4 的口径正好相反**：处置留痕必须等宿主恢复成功**之后**写，这一条必须在删除**之前**写；
  前者失败可以吞，后者失败**不许删**。两条不许互相照抄。
- 判别力：`scripts/mutate_recovery_package_retirement.py` **R1–R5 四判据全部成立**
  （R1 删前不取指纹 / R2 落账失败也照删 / R3 四列从必需列名单拿掉 / **R4 删除失败也宣称删掉了** /
  **R5 删除确认不首次写赢**）；新增用例 **12 条**、相邻既有 **100 条**、**服务端全量 1934 passed**。

**边界**：不追溯（上线前删掉的材料永远没有这条留痕）、`NULL` = 没有留痕（不是"材料还在"）、
`removed_at` 为空 = **没有确认删除**（不是"删除失败"这一种解释）、指纹不是"材料可用"的结论、
不新增对外读接口、只覆盖 WPS（Excel 侧没有本机恢复材料）。

### 1.3 U7 本轮实现（2026-10-30）

详见 `docs/DEAD_CODE_U7_2026-10-30.md`。一句话：**只删"没人调的那段流程"**。

- `stagePendingPlan` 46 行整段删除（0 调用点、0 测试引用），连带 3 处孤立 import
  （`stageWpsNativeComparison` / `WpsFormulaTrialError` / `fixedValueFallbackFromFormula`）。
- 同时删掉 **5 条永远不会失败**的 `expect(stageWpsNativeComparison).not.toHaveBeenCalled()`：
  唯一调用者是死代码，所以这些断言删除前就已经恒真。同处正向锚点（`host.writes()` 不变、
  `stageWpsBeforeComparison` 调用一次）全部保留。
- **保留**`stageWpsNativeComparison` 与 `fixedValueFallbackFromFormula` 本体：它们有各自的直测，
  删不删是"删能力"，得另有结论。代价是它们现在"生产侧无调用者"—— 这是新事实，别读成"已支持"。
- 验证：`tsc --noEmit` exit 0、`npm run build` exit 0、`vitest run` **70 文件 / 1222 用例全过**
  （2 文件 / 8 用例跳过）。注意 `tsconfig` 没开 `noUnusedLocals`，**孤立 import 不会让 tsc 变红**，
  所以必须 grep 复核（四符号在面板内均归零）。

### 1.4 U3 核实（2026-10-30，仍不开）

完整方案见 `docs/SDK_SESSION_PRIVACY_2026-10-30.md`。只记三条会改变判断的事实：

1. `%LOCALAPPDATA%\ExcelBro\ClaudeAgentSdk` 里有 **30 份转录 / 612 行 / 1.44 MiB**，
   时间跨度 **2026-10-01 → 10-05**；`--no-session-persistence` 是 2026-10-07 的 `a8ae64cb` 引入的
   ⇒ **这些转录全部早于开关**，不是现在的代码写出来的。
2. **产品服务端不读它们**：`_state_dir()` 只把它当 `CLAUDE_CONFIG_DIR` + `cwd`（`claude_sdk_agent.py:173`）。
   能读的是"本机同用户 / SYSTEM / Administrators"。
3. **目录 ACL 不是产品保证**：`claude_sdk_agent.py:81` 的 `mode=0o700` 在 Windows 上不生效
   （只影响只读位）。现在的三条 ACE 是"本来就在"的，换台机器未必一样。
   **含单元格值的记录**：`map_column_values` / `submit_plan` 的入参会带值 ⇒ 转录里可能有工作表数据。

**未做**（按要求）：不新增持久化、不改开关、**不删**这 30 份文件。

### 1.5 U6 复核：Excel 到底缺什么（2026-10-30，本轮不加）

**先纠正一个说法**：Excel 没有 `pendingPlan` **不是缺闸门**，是架构差异 —— Excel 的计划到达即
自动执行一次（"执行一次 → 确认只清标记"），压根不存在"待确认计划在位"这个状态。
把 WPS 的 8 处 `pendingPlan` 闸门搬过来，等于搬实现结构，不改用户行为。

真正缺的是两处**用户看得见**的东西（**2026-10-30 第十七轮复核：均已用 Excel 自己的方式闭环，见下**）：

| # | 缺的行为 | 证据（可复核） | 第十七轮复核结论 |
| --- | --- | --- | --- |
| 1 | 结果卡上有「↶ 撤销本次执行」，但**没有一句说明撤销范围**（本次/含之后 N 次、哪些不回退） | `ChangeReviewUndoNote` 定义在 **Excel 包**（`apps/excel-addin/src/components/ChangeReviewPresentation.tsx:31`），却只被 **WPS** 渲染（`WpsAgentPanel.tsx:9055` import 自 excel-addin）；Excel 自己的 `App.tsx` 无渲染点 | **已闭环（另一种方式）**：Excel `App.tsx:11721` 渲染 `↶ 撤销本次执行（含之后 N 次）` + `11742`「只能从最近一次往前逐级撤销」；虽未用 `ChangeReviewUndoNote` 组件，但"撤销范围说明"这个用户可见行为已存在，`excelWriteEndToEnd`/`crossProcessChain` 等测试有断言 |
| 2 | 恢复完成后，消息里那张旧卡**不标"已执行/不可续跑"** | `settleRecoveredPlan`（`WpsAgentPanel.tsx:3137`，调用点 3261/3328/3365）只有 WPS 有；Excel 用 `executedPlanId` / `executedExecutionId` 标记，没有等价的收尾 | **已闭环（等价接线）**：Excel `updateFolderResumeMessage` 把 `officeResume: undefined`（`App.tsx:3055`）+ `executedPlanId`/`executedExecutionId` 达到等价效果——旧卡恢复入口消失、标记已执行；测试多处断言「继续剩余动作」收尾后消失 |

**刻意不加、别再当缺口**：Excel「备份与恢复」抽屉里没有值写入撤销入口 —— 这会破坏
"一个执行只有一个可点入口"（`selectExecutionRecoveryEntry`），是取舍不是遗漏。

### 1.6 U1 / U2 的最小可交付范围与顺序（2026-10-30）

详见 `docs/U1_U2_MINIMAL_SCOPE_2026-10-30.md`。结论：

| 顺序 | 项 | 为什么这个顺序 | 一轮能否收 |
| --- | --- | --- | --- |
| 1 | **U1a `writeFormulas`** | 恢复点结构与 `writeValues` 同构、判据现成、不碰信任模型 | **已实现（第十六轮，双端）** |
| 2 | **补一条 `copyRange` 续跑用例** | 它展开成 `writeValues` + 来源快照，形状上可能已经满足；**用例先于实现**，避免白做 | 能（半天） |
| 3 | **U1b 格式动作** | 恢复点是属性快照不是值矩阵；必须先出"格式指纹"core 判据 | 分两轮 |
| 4 | **U2 结果引用** | 要回答"结果引用能不能充当来源证据"，等于改信任模型 | 需先决定 |

`insertRows` / `removeDuplicates` 不在这条线上：前置是**实机行插入恢复**，属真机验收不属开发排序。

## 2. 自动化已验但真机未验

### 2.1 双端缺口矩阵里仍是 `⏳` 的格子（`SDK_A1_E3_STOP_AND_GAPS_2026-10-25.md` §3）

| # | 场景 | Excel | WPS |
| --- | --- | --- | --- |
| 1 | 换工作簿 | ⏳ 真实窗格开两个簿再切 | ⏳ 同 |
| 2 | 同名工作簿（含未保存） | ⏳ 两个同名未保存簿 | ⏳ 同 |
| 3 | 工作表改名/删表/重排 | ⏳ 面板上真的取消勾选；重排无专测 | ⏳ 同；删表撤销见 §2.4 |
| 4 | 授权收窄 | ⏳ 数据范围面板真实交互 + 视觉 | ⏳ 同（面板视觉一直没验收过） |
| 5 | 旧模型结果迟到 | ⏳ 真实模型吐旧引用 | ⏳ 同 |
| 6 | 旧桥事件迟到 | ⏳ 真机把进度事件拖到收尾之后 | ⏳ 同 |
| 7 | 旧执行者迟到/被替代 | ⏳ 真实双窗格 | ⏳ 真实双窗格 |
| 8 | 同时恢复/快速重复点击 | ⏳ 真实双击 | ⏳ 真实双击（本轮**没有**连点） |
| 9 | 写入后回执丢失 | ⏳ 真实断开（关窗格/断网）后再回执；本轮的"中断"是**开发探针**制造的，不算回执丢失 | ✅ 原生（`1fd938fd` 补记不重写）；仍缺"补记时刻"的独立记录 |
| 10 | 第二步中断后只续第三步 | **✅ 真机（账本级）**`c52ad791`：0/1 号写同秒、2 号晚 2h33m27s、`next=3`、`in_flight=NULL`；核对零写入是**推断**（核对不写 `office_host_writes`），计数口径仍是端上自报 | ✅ 原生（`1fd938fd` 15:49、`5a2d84be` 19:42 两次）；计数仍是端上自报 |
| 11 | 异常退出后恢复 | ⏳ 真杀进程；本轮只是**窗格重开**（`EXCEL.EXE` 未退出），不能当跨进程证据 | **正常退出版 ✅**（`3f00ea0d`：完整退出 WPS 后重开，材料跨会话可用、续跑只写第三步）；**崩溃/被杀 ⏳** |
| 12 | 完成/确认/放弃/撤销后重开 | **部分 ✅**：确认后终态由账本支撑（`completed` / `in_flight=NULL` / `next=3`），恢复卡消失是口述；**"确认"这个动作本身服务端无留痕**（U4 已补、不追溯）；放弃/撤销两分支仍 ⏳ | 确认后**保存重开**✅（`1fd938fd`）、**同会话**放弃✅（两次）、**跨进程放弃 ✅**（`3f00ea0d`，重开后无恢复卡）；**真实撤销 ⏳** |

### 2.2 九场景（`SDK_A1_EXECUTION_EVIDENCE_2026-10-24.md`）里"原生宿"仍是空的

①回执丢失 ②运行中断开/重启 ③多动作中途断开 ④旧结果迟到 ⑤换簿/收窄/改来源
⑥双恢复请求 ⑦引用过期 ⑧完成任务重开 ⑨同名/其他任务串用 —— 服务端与面板级证据齐全，
**真实宿主里一次都没做过**。其中 ③ 在 WPS 上已有原生证据（同会话一次 + 跨进程一次 `3f00ea0d`）；
② 只做了**正常退出**版，**崩溃/被杀进程**仍空着；⑥ 完全没碰。

### 2.3 原生验收包 S1–S3（`A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md`）

| 场景 | 状态 |
| --- | --- |
| **S1 来源变化 → 本次规划停止** | **一次都没跑过**。这是本批唯一的新行为，`EXCEL_BRO_SOURCE_CHANGED_FAULT` 的脚本、绑定、证据采集都做好了，但没有原生观察。真实模型侧只有 A1-E2 的 B 路那种"第二次读被拒"的间接证据，不能替代 S1。 |
| **S2 精确中断 → 重开 → 连点恢复** | 中断+续跑有原生（同一会话两次、跨进程一次 `3f00ea0d`）；**"连点"没做**（本轮与 `3f00ea0d` 都是单次点击） |
| **S3 确认/放弃后重开** | 确认侧有（`1fd938fd` 保存重开）；放弃侧有 **跨进程版**（`3f00ea0d`：重开后无恢复卡、三个负数保持） |

### 2.4 与 A1 有交叉但不在 A1 内的既有 `⏳`

- 删表「原地撤销」（`run_worksheet_restore_mutations.mjs` R1–R8）——实机未验收；
- 数据范围面板的视觉/同步（`run_scope_authorization_mutations.mjs` A1–A6、`run_scope_identity_mutations.mjs` I1–I4）——
  静态 CSS 断言不算视觉验收；Excel 端面板标题仍在滚动容器内；
- 并排审阅布局（35/35/30）截图未执行；
- Excel 2024 / WPS 格式动作实机闭环未执行。

## 3. 证据待补（行为有了，记录缺了）

| # | 缺什么 | 现状能证明到哪一步 | 补法 |
| --- | --- | --- | --- |
| E1 | **"续跑后确认 + 随后对话"同一轮** | 前半 `1fd938fd`、后半只有 `WpsAgentPanel.test.tsx` 的 `it.each("确认变更")` | 见 §3.1 |
| E2 | **"确认后重开"的单元格值** | 放弃版已有磁盘证据（`3f00ea0d`，11022 B / `db31f17b…`，逐格 0 差异）；**确认版仍只有截图**（`1fd938fd` 那轮文件未保存；2026-10-29 Excel 轮同样**未保存**：mtime 早于执行、`B2:B4` 仍是 `-101/-202/-303`，见 `outputs/a1-acceptance-20261029-2/EVIDENCE.md` §2） | 确认版验收时**另存一次**再看磁盘值，或明确写"截图即终值"。**本轮不要求用户补做** |
| E3 | **放弃的服务端记录** | 材料消失（唯一删除点在撤销成功分支）+ `revision` 不推进 + 截图 | **U4 已实现**（`docs/EXECUTION_DISPOSITION_2026-10-29.md`）：下一次放弃/撤销起，账本会写 `abandoned_at` / `undone_at`。**不追溯**历史轮，所以本条对已发生的几轮仍是"推断" |
| E4 | **补记时刻的独立记录** | 账本无逐次修订历史，只能看 `updated_at` 与 `office_host_writes.recorded_at` | 需要给账本加事件表（A2 范围） |
| E5 | **宿主真正写了几次** | `hostWriteCalls` / `mutatingCalls` 是端上自报 | 宿主侧独立日志或 WPS 宏层计数 |
| E6 | **本轮截图** | 会话内贴图，**图片字节没有落盘**（已核对会话转录中无 `image/png|jpeg` 数据块；2026-10-29 Excel 轮同样：`outputs/` 与仓库根在 18:30 后无新 png，`clipboard-images` 最新是 12:18） | 按 `…followup/screenshots/README.md` 的 8 个文件名放进目录。**不要求用户重跑或补图**，缺就记缺 |
| E7 | **`query_table` 的实际参数** | `planning_task_events` 只记 `tool` 名 + `operationId` | 见 §4 的验证步骤 |

### 3.1 E1 的补法（约 3 分钟，不改代码）

目标：让"续跑后确认 + 紧接着发消息"落在**同一条执行**上。

1. 打开 `outputs/wps-abandon-20261028/WPS放弃验收-20261028.xlsx`（先确认 `目标表 B2:B4` 是 `111/222/333`
   或先放弃回 `-101/-202/-303`，两者都行，只要知道基线）；
2. 勾 `目标表`+`来源表` → 发三步写入任务 → **用开发探针「启用下一次第二步中断」**（与 §10.2 同一入口）；
3. 等界面出现「待核对在途动作」→ 点「核对在途动作」→ 点「继续剩余动作」；
4. **点「确认变更」**（不要点放弃）；
5. **在同一个对话里**直接发："查询目标表 B2:B4 的当前值，不修改任何单元格"；
6. 取证：`python scripts/snapshot_ledger_20261028b.py` → `python scripts/export_followup_evidence_20261028.py`
   （脚本里把 `CONFIRM` 换成新的 `operationId`），核对三件事——
   ①新执行的 `office_host_writes` 有 `idx1 failed_after_write` 且 `idx2` 显著靠后（= 真的续跑过）；
   ②`revision` 应为 11 而不是 8；③随后那条 `planning_tasks` 是 `completed` 且有 `localResult`。

三条都成立 ⇒ 才可以把"**续跑后**确认 + 随后对话"记为真机通过，并把 E1 从本节移走。
缺任何一条就还是"证据待补"，**不要**用"普通确认执行"那一轮顶替。

## 4. 独立待查项：明确查询 `B2:B4`，结果卡却说「扫描不完整」

**状态：2026-11-01 已定位 + 已修（自动化通过，**真机未复测**）。**

### 4.0 结论与修法（2026-11-01 补）

真机再一次复现（`turn-b018aec5a1cc4f50ab1d8caa16fcb551`，WPS，19:04）：
用户只说「只读取「来源」工作表的 A1:B5」，结果卡 `sourceSheets` 却是
`["来源","结果1","结果2","结果3"]`、`complete=false`，界面照旧写「扫描已达上限」。
这一次能确定两件事：

1. **表名在"模型产出的 arguments"这一层就没了**，不是桥或执行器丢的 ——
   证据是执行器把四张表都 `addSheet` 了：旧代码里 `scope: "selected"` 等于"全部已授权表"，
   只有 `range` 带表名或 `scope: "active"` 才会收窄。桥只做形状闸（原样转发），
   所以问题在上游：授权集合被当成了本次查询范围。
2. **「扫描已达上限」是写死的一句文案**，与真实原因无关：任何 `complete=false` 都这么写。

修法两处，互不替代：

* 新增 `sheets`（本次要查的表）+ `packages/core/src/queryTarget.ts` 的 `resolveQueryTarget`：
  多张可查表却没绑定时**拒绝**并列出可选表名，不再退成"全都查"；
  `range`/`sheets`/`scope:"active"` 三种绑定的优先级也只在这里定。
* `complete` 拆成 `complete` + `incomplete[]` + `truncated`/`truncation`：
  原因由产生它的那段代码写（缺字段 / 没有数据行），卡片按原因渲染，截断单独一句；
  模型回执带同一份事实。读失败与扫描超限仍然整条报错（不产出部分结果），
  所以它们**不会**出现在卡片上 —— 这正是"不能统一说成扫描达上限"的正解。

**参数不落账**这一条也补了：进度事件的桥指针现在带 `BridgeQueryBindings`
（模式/scope/表名/区域/表头口径，不含任何值），下次真机可以直接看参数。

**未做**：真机复测（本次没有打开真实工作簿）。所以本节只能记为"已修、待复测"，
不得写成"真机通过"。

### 4.1 事实（原始记录，保留）

| 轮 | 结果卡 | `complete` | `sourceSheets` | 界面 |
| --- | --- | --- | --- | --- |
| 19:39:56（确认后） | `res_15c9d722c89b513a2596d6d38a0273b6`，`rowCount=3` | **false** | `["目标表","来源表"]` | 出「结果不完整（扫描已达上限），仅展示已取回部分。」 |
| 19:43:31（放弃后） | `res_fff90058508135bbbf84a77682abe1d3`，`rowCount=3` | true | `["目标表"]` | 无该提示 |

两次的 `rowCount` 都是 3、都是 `B2:B4`、都在同一本簿上，差别只有 `sourceSheets`。
模型在第一次的回答里把它解释成"你只要了前 3 行，所以完整标记为否"——**这个解释与代码不符**。

### 4.2 代码定位（两端共用同一份判据）

- `apps/excel-addin/src/dataTools.ts:480`：`sourceSheets` 在**扫描每张表时先登记名字**，
  与这张表是否真的产出记录无关；
- `dataTools.ts:487-488` + `117-146`：`recordsFromSheet` 对**缺请求字段**的表往 `warnings` 里推
  `「<表名>」未找到字段：…` 并**整表跳过**；
- `dataTools.ts:567 / 611 / 626`（以及 legacy 路径 `810 / 854 / 870`）：`complete: warnings.length === 0`；
- 卡片文案只覆盖一种原因：`apps/excel-addin/src/components/LocalResultCard.tsx:67`
  ——「结果不完整（**扫描已达上限**），仅展示已取回部分。」

即：**`complete` 把"扫描触顶""某张表缺字段""筛选后无记录"三种完全不同的原因压成一个布尔**，
而文案只说了其中一种。19:39 那次最符合"扫了两张表、`来源表` 没有请求的列 ⇒ 一条 warning ⇒
`complete=false`"，但**这只是与结构一致的假设**。

### 4.3 尚未证实的部分（不要写进结论）

- `query_table` 的**实际参数没有落账**（`planning_task_events` 只记 `tool` 与 `operationId`），
  所以"模型当时传了 `scope=workbook` 还是 `scope=active`""请求了哪个字段"都是反推；
- 不能排除"那一次 `limit` 被调小 / 触发了别的 warning"这类同样能置 `complete=false` 的原因；
- 因此**不能**写"根因已定位"。

### 4.4 验证步骤（下次要查时照这个做，别凭印象）

1. 在真实 WPS 里对一本**多表已授权**的簿发一条按区域的只读查询，**同时**开服务端 debug 日志，
   记下 `query_table` 的完整 `arguments`（`scope` / `fields` / `limit` / `filters`）；
2. 看返回的 `localResult.complete` 与 `sourceSheets` 是否复现 19:39 那种组合；
3. 单测级确认：`executeQueryTableData` 传入两张表、其中一张缺请求字段 ⇒ `complete=false`
   且 `warnings` 里出现「未找到字段」（这条今天就能写，不依赖真机）；
4. 结论只有两种情况：**"扫了不该扫的表"**（收窄 `sheetsForRequest` 的入参）或
   **"文案把原因说错了"**（把 `complete` 拆成 `reason` 或按 warning 分类给不同文案）。
   两种修法互不替代，先定哪一种再动手。

### 4.5 已知的相邻风险（同一处代码）

`HEADER_SCAN_ROWS` 只用于收集 `availableFields`（`dataTools.ts:482`），若表头不在前几行，
`availableFields` 会缺字段 —— 这与"缺字段 ⇒ `complete=false`"叠在一起会更容易误报不完整。
本条不展开，只登记。

## 5. 下一项最值得做的工作

### 5.0 原推荐已完成（留档，别再跑）

**2026-10-29 部分完成**：Q0–Q3 + 「确认」分支已由用户真机走通（`execution-c52ad791…`，
证据 `outputs/a1-acceptance-20261029-2/`）。与计划的差别有三处，别按原表当成"全过"：
①第三步续完后用户点的是**确认**而不是**放弃**（Q4 的放弃分支在 Excel 上仍空）；
②**没有保存**（Q5 的磁盘取证因此不成立，磁盘仍是 `-101/-202/-303`）；
③中断是**开发探针**制造的，不是"回执丢失"（§2.1 #9 仍是 ⏳）。
**不要求用户为此重跑** —— 已发生的就按已发生的记，缺的照实标 ⏳。

WPS 跨进程「中断 → 完整退出 → 重开 → 核对 → 续跑 → 放弃」于 2026-10-28 夜间完成并通过
（`wps-execution-3f00ea0d…`，证据 `outputs/a1e3-native/cross-process-abandon-20261028/`）。
原计划 P0–P6 的判定点全部命中，只有 P3 的预期形态有出入：重开后**没有**直接摆出"待核对在途动作"，
而是先给「证据不足」、点一次「重新核对（只读）」才恢复（已记入 §6）。该偏差不影响 P4–P6 的结论。

### 推荐：**Excel 端「两步后中断 → 核对 → 只续第三步 → 放弃变更」**

理由：

1. Excel 的 #9 / #10 / #12 三格**全空**，是矩阵里最大的一块空白；WPS 侧同型路径已经四轮通过
   （`1fd938fd`、`9e152fde`、`5a2d84be`、`3f00ea0d`），可以直接照现在的判据抄作业；
2. Excel 的自动化底子最厚：`excelWriteEndToEnd.test.tsx:620` 已有"双击只核对并继续、
   `startOfficeAction` 索引 `[0,1,2]`、续跑只写 `B4`、确认/放弃两分支重开"三条成熟用例，
   真机只需要验证"同一套判据在真实 Office.js 里也成立"；
3. 顺带能验掉一个只有替身证据的修复：`currentOfficeBridgeWorkbookIdentity()`
   在真实宿主里对**未保存同名簿**给出的会话内身份（`bridgeWorkbookIdentity.test.ts` 至今没有真机样本）。

步骤（Excel，判定点不成立就停）：

| 步 | 操作 | 判定点 | 留什么证据 |
| --- | --- | --- | --- |
| Q0 | 用 `scripts/make_a1e3_native_fixture.py` 同款三格簿（或复制 `WPS放弃验收` 改名为 `Excel放弃验收.xlsx`），记 `目标表 B2:B4`、`来源表`、`未授权表 B2`、sha256 | 与基线一致 | 磁盘 sha256 + 三格截图 |
| Q1 | Excel 端加载项连独立测试服务端 → 数据范围勾 `目标表`+`来源表` → 发三步写入 | 账本出现**新** `operationId`；`workbookIdentity` 是 `url:` 前缀 | 新 operationId + identity 原文 |
| Q2 | 用开发探针在第二步写入后中断（Excel 侧同一入口） | `next_action_index=1`、`in_flight_action_index=1`、`idx1 failed_after_write` | 账本行 |
| Q3 | 点「核对并继续」（**连点两次**，看是否只认领一次） | `registerOfficeExecution` 只 1 次；`startOfficeAction` 索引 `[0,1,2]` | 端上计数 + 账本 |
| Q4 | 续跑完成后**不点确认**，点「放弃变更」 | `目标表 B2:B4` 回基线；`来源表`/`未授权表` 不变 | 逐格截图 |
| Q5 | **保存**再取证 | 磁盘三处基线；与原始素材逐格比对 0 差异；条件格式规则 0 条 | 快照 JSON + `disk-values` 同款比对 |

Q5 的"保存后再取证"是这一轮学到的东西：不保存就永远只能用截图当终值。

### 第二选择：**S1（来源变化 → 本次规划停止）原生**

本批唯一的新行为，也是"原生与自动化证据分别齐全"最刺眼的一格。要先有一次只读需求把桥立起来，
再按 `A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md` §1 的 `arm` 流程绑定，成本高于 Excel 那一项，
但它是 A1-E3 的立项理由，不能一直空着。

### 第三选择：**WPS 崩溃/被杀进程版恢复**

跨进程的**正常退出**版已经通过（§5.0）。崩溃版要验的是"宿主没来得及刷盘时，恢复材料与
账本是否还自洽"，与正常退出不是一回事。做法：同样在第二步中断，然后用任务管理器**结束 WPS 进程**，
重开后走同一套核对。**不得**用正常退出的结论代替，也不得把"没崩"当成"崩溃也能恢复"。

### 5.1 两条取证纪律（2026-10-30 定，省掉一整轮重跑）

**① 真实撤销优先复用"仍有有效恢复点的已完成执行"，不默认从中断重来。**

理由：撤销（U4 的 `undone` / U5 的 `recovery_package_removed_at`）要的只是"一次确认过的执行 +
一份还在材料里的恢复点"。为此重新制造一次中断，等于把已经通过的 #10/#11/#12 三格再跑一遍
（耗时间且不加证据）。做法：挑一份**已确认、恢复点在有效期内、单元格仍是执行后值**的执行，
直接在结果卡点「↶ 撤销本次执行」，然后只读账本这两组列。
只有"找不到可用恢复点"时才退回"执行→确认→撤销"的最短路径。

**② 「后端发现」的取证用开发诊断或请求记录，不去动账本、也不删材料。**

- **不要**为了"让发现提示出现"而让只读发现在账本里写记录 —— 那会把"发现过"变成自证；
- **不要**删除恢复材料来模拟"本机索引丢失" —— 那是破坏取证对象，而且删材料本身现在会留痕（U5），
  等于先污染再取证；
- 正确做法：用**开发诊断入口**（DEV 面板/地址栏 hash 的诊断视图）看到端上发出的发现请求与
  服务端返回，配合**请求记录**（服务端访问日志中的 `GET /api/office/executions/...` 与
  发现端点）作为"这一次真的问了"的证据；再看 `discovery_dismissed_at` 是否仍为 `NULL`
  以区分"没问过"和"问过但被清过提示"。

### 明确**不要**再重复的（已经通过，重跑只是耗时间）

- 普通三步写入 + 确认变更（`7afd3dc6`、`1fd938fd`）；
- 同会话"中断 → 核对 → 续跑 → 放弃"（`9e152fde`、`5a2d84be`）；
- **跨进程（正常退出重开）"中断 → 重新核对 → 续跑 → 放弃"（`3f00ea0d`）**；
- **Excel「中断 → 重开发现 → 核对 → 只续第三步 → 确认」（`c52ad791`）**：账本三格逐格一致、
  `next=3`、`in_flight=NULL`、`office_host_writes` 时间对得上；**不重跑**（断点仍是开发探针造的，
  重跑一次也变不成"回执丢失"）；
- **Excel 恢复入口位置**（固定入口没被日志顶出视口）：2026-10-29 用户口述通过；
  截图字节**未落盘**，所以它只算"口述通过"—— 但**不要求再点一次**，缺的是截图归档，不是行为；
- "放弃之后能不能继续对话"（`5a2d84be` + `turn-68fd1b…`）、"普通确认之后能不能继续对话"
  （`7afd3dc6` + `turn-4ec317…`）；
- 只读取证本身（快照脚本 + 导出脚本已可复跑，不需要再让人点一遍）。

## 6. 第二项待查：首次重开显示「证据不足」，点一次只读重试就恢复

**状态（2026-10-29 更新）：诊断能力已交付，根因仍未确定、未归因、未修复。** 与 §4 是两件事，别合并。

这一节原来写的是"拿到原因码之前不动三处"。**原因码那一步已经做完了**（只做诊断，不动判定）：
`inspectWpsRecoveryState` 的每一条提前返回都有了独立码，`localInspection` 的取值、阶段、动作序号、
表数量与 ID 尾部会在 **DEV** 下留一条记录。完整码表、字段边界、获取方式与判读表见
[WPS 恢复核对 DEV 诊断](WPS_RECOVERY_DIAGNOSTICS_2026-10-29.md)。
**这仍然不能算根因修复**：真机只跑过 **1 次、且是成功路径**（2026-10-29 03:42:31Z，
`#1 discovery verdict=reconcile outcome=completed phase=values …`，无原因码），
**28 个码一个都还没在真机上触发**；下面 6.3 的六条候选原因依旧全部无证据。

### 6.1 事实（可核实到哪一步）

2026-10-28 夜间跨进程轮（`wps-execution-3f00ea0d…`）：完整退出 WPS 后重开同一文件，
面板先给出 badge「**证据不足**」，正文「**无法读取或证明当前目标现场，不能补记或重新规划写入。**」
（截图 02 / 03，23:32:07 与 23:32:17）。用户点「重新核对（只读）」后，同一张卡回到
「待核对在途动作」（截图 04，23:39:06），随后补记 → 续跑 → 放弃全部正常。

账本侧：这段时间**没有任何记录**。`office_host_writes` 在 23:25:50.403 之后直到 23:39:26.371
才出现下一行；`office_executions.revision` 只数状态写次数，也不记"只读核对"跑过几次。

### 6.2 能确定到哪一支（且只到这一支）

那段正文在代码里**只有一处**：`apps/wps-addin/src/wpsExecutionReconcile.ts:75` ——
`input.localInspection == null || input.localInspection === "unknown"`。
因此可以**排除**同一函数里的其余分支：

| 被排除的原因 | 对应文案（未出现） |
| --- | --- |
| 服务端没有这次执行的记录 | 「服务端没有这次执行的记录，无法判断写入现场。」（:52） |
| 当前工作簿与记录身份不一致 | 「当前工作簿与这次写入记录的身份不一致…」（:57） |
| 动作总数/游标非法 | 「动作总数或进度游标无效…」（:69） |
| 在途动作与游标不一致 | 「在途动作与进度游标不一致…」（:80） |
| 服务端说需要核对却没给在途动作 | 「服务端标记需要核对，但没有给出在途动作位置…」（:87） |
| 有在途动作但本机恢复点读不到现场 | 「有在途动作，但本机恢复点读不到现场…」（:98） |

也就是说：**当时身份比对是通过的、服务端状态是拿到的**，缺的是"当前现场读出来/证明得了"。

### 6.3 `localInspection` 怎么会是 `unknown`

`localInspection` 由 `inspectWpsRecoveryState(app, state, indexes)` 给出
（`apps/wps-addin/src/wpsRecoveryInspection.ts`，面板调用点 `WpsAgentPanel.tsx:2549`）。
它在下列任一情况返回 `"unknown"`（每条都是**候选**，不是结论）：

1. `authorizeWpsSheets(app, selectedIndexes)` 抛错 —— 例如刚重开时授权工作表索引还没解析出来；
2. 身份或动作描述与登记不一致（`descriptor.actions.length !== totalActions` 等）；
3. 登记里某张表不在**当前授权**集合内（重开后授权范围还没恢复成两张表）；
4. 任一动作为目标区域形状不符、或 `Value2` / `Formula` 读回 `undefined`（宿主还没就绪）；
5. 格式动作的属性名不可识别 / 快照不完整；
6. 面板侧 `wpsRecoveryVerdictFor` 的 `try` 抛错被吞掉（`WpsAgentPanel.tsx:2552`）。

**这六条当时没有任何一条有证据。** 2026-10-29 起它们各自对应独立原因码（可选其一，不是合并）：

| 上面第几条 | 现在会产生的原因码 |
| --- | --- |
| ① 授权抛错 | `authorize_no_selection`（已选表是空集）/ `authorize_threw` |
| ② 身份或描述与登记不一致 | `workbook_identity_mismatch` / `workbook_identity_unreadable` / `descriptor_empty` / `descriptor_action_count` / `action_index_mismatch` / `action_range_*` / `action_cell_count_mismatch` / `action_fingerprint_missing` / `cell_budget_exceeded` |
| ③ 登记表不在当前授权 | `authorized_sheet_not_allowed` / `action_sheet_not_allowed` / `action_sheet_not_authorized` |
| ④ 区域形状/读值为空 | `host_range_unreadable` / `host_range_shape_mismatch` / `value_read_undefined` / `value_shape_unexpected` / `workbook_unreadable` |
| ⑤ 格式属性/快照 | `format_fields_undeclared` / `format_fields_unrecognized` / `format_snapshot_incomplete` |
| ⑥ 面板 `try` 吞错 | `panel_threw`（另有 `inspection_threw` 覆盖核验函数内部的异常） |

**注意：加码不等于有证据。** 2026-10-29 03:42 那次重开，首条记录是 `outcome=completed`、
无原因码（详见 [诊断文档](WPS_RECOVERY_DIAGNOSTICS_2026-10-29.md) §5.1）——即**那次没有走进本节的
失败分支**，也**没有**第二条 `recheck` 可对照，**本轮未复现**。所以本节状态仍是"未归因"。
尤其不能写成"加载时机问题""授权没恢复""材料丢了"——
材料这一支由证据目录 `verification.json` 的 X6（材料已随撤销收尾回收，续跑本身成功）
反向说明当时并不缺材料；
而"加载时机"与"授权"两说都没有日志支持。

**再复现时的一条硬约束**：诊断缓冲区是**内存态**，页面重建（重开 / 刷新 / 换窗格）即清零。
「首次重开那一条」只能在**当次载入里**抄；一旦重载就拿不回来，不能靠事后补看。

### 6.4 要闭这条，需要补什么（下次照做）

1. ~~给每个 `return "unknown"` 加 DEV 原因码~~ —— **2026-10-29 已完成**：
   `inspectWpsRecoveryStateDetail` 带 `code` / `phase` / `actionIndex` /
   `selectedSheetCount` / `authorizedSheetCount`，面板只在 DEV 记录（非 DEV 一行不留）。
2. 复现条件仍是两条，**分别取一次**：①重开后**立刻**看（不等面板刷新）；②重开后等 5–10 秒再看。
   若只有①出现，方向就是"宿主/授权未就绪"；两条都出现，方向在状态或描述本身。
   本次不要求"首次必现、二次必不复现" —— 没复现就如实写"未复现"。
3. 每一步都要**同时抄下两条记录**（首条 `discovery` 与点一次「重新核对（只读）」后的首条 `recheck`），
   靠 `op=` / `task=` 尾部确认是同一次执行。只给一条无法对照。
4. 拿到真实码之后再决定改哪：是"重开后先等一步再判"，还是"`unknown` 时自动重试一次"，
   或者"文案要区分'读不到现场'与'读到了但不一致'"。**在拿到真实码之前不改这三处**，也不加自动重试。

## 7. 本轮不做的事

- 不改真实工作簿、账本、恢复材料；不重启服务；不提交、不推送；
- 不为了让某格变绿而降低判定点（尤其"中断点"与"跨进程"这两处，拿同会话冒充跨进程就等于没测）；
- 不把 U4/U5 这类"没有记录"写成"已验证不存在"；
- 不把 §4 的待查项因为"这轮没复现"标为已修复。

## 8. WPS 公式计划「写入前恢复登记失败」（2026-11-01）

**状态：已定位 + 已修（自动化通过，**真机未复测**）。**

真机：三个独立 `writeFormulas` 计划已生成，写入前登记被拒，界面只有一句

> 无法持久登记本次写入现场，已取消写入；原表未被修改。续跑来源越界、形状不符或不支持该动作组合。

场景：来源 `A1:B5`（只读）；`结果1!B2 =来源!B2*1`、`结果2!B2 =来源!B4*2`、`结果3!B2 =来源!B5*3`；目标原为空。

### 8.0 根因（只读定位，不是从文案猜的）

那句话是**服务端**拼好的：`main.py::_validate_wps_recovery_descriptor` 里
`WPS_RESUME_CONTEXT_INVALID` 的 `message`。界面只是把它接在"无法持久登记…"后面。

真正命中的是**动作组合**那一条：

```python
valid = all(a.type == "writeValues" for a in request.descriptor.actions)   # 旧
```

即"带来源证据的描述里只允许值写入"。而端上从公式恢复落地那天起就把
`writeFormulas` 算作可带来源证据的动作（`wpsPlanExecutor.isResumableValueAction`），
两端口径漂移 ⇒ **凡是公式计划，一律登记不了**。

判别证据（不是推理）：同一份载荷**只把三个动作的 `type` 换成 `writeValues`**，
其余（来源表名、`A1:B5`、5 行 2 列、授权集合、两个指纹、`version`）一字不改，
登记立刻 200。见 `server/tests/test_wps_formula_resume.py::test_action_type_is_the_only_input_that_decides_this_branch`。
所以不是"来源越界"，不是"形状不符"，也不是协议字段缺失 —— 端上产出的
`resumeContext.version === 1` + 一块来源证据，服务端本来就认（只是被那个 `all` 一并否掉）。

另外两个叠加缺陷（同一处代码）：

1. **四条拒绝原因合并成一个布尔**：来源越界 / 形状不符 / 动作组合 / 配额，任一条不成立都报同一句，
   排查只能靠猜。
2. **零写入却留下在途核对待办**：写入前预置的恢复标记是
   `executing + execution_unknown + reconcile + unknown`（当时唯一诚实的说法），
   中止时没有回退 ⇒ 中断卡写"写入结果无法确认，需要先核对"，审阅条写"本次执行需要核对"。
   现场其实**一个单元格都没动**。

### 8.1 修法

| 位置 | 改动 |
| --- | --- |
| `main._WPS_RESUME_SOURCE_ACTION_TYPES` | 新增：可带来源证据的动作 = `{writeValues, writeFormulas}`。与 TS `isResumableValueAction` **同一份判据**，漂移由 `test_resume_source_action_types_match_the_ts_criterion` 钉住 |
| `main._validate_wps_recovery_descriptor` | 拆开 `valid`：`WPS_RESUME_ACTION_UNSUPPORTED` / `WPS_RESUME_SOURCE_RANGE_INVALID` / `WPS_RESUME_SOURCE_SHAPE_MISMATCH` / `WPS_RESUME_SOURCE_OUT_OF_SCOPE`(403) / `WPS_RESUME_SOURCE_TARGET_OVERLAP` / `WPS_RESUME_SOURCE_LIMIT`。判据一个字没放宽，只是不再合并陈述 |
| `WpsAgentPanel` 四个写前出口（身份缺失 / 登记失败 / 无写入权 / 索引存不下） | 统一走 `abandonBeforeFirstWrite`：恢复标记退回 `WPS_NOT_EXECUTED_PATCH`（= `not_written` + 保留需求后重新规划）、计划预览不再是"待执行"、原因经 `zeroWriteAbortRef` 交给审阅通道 |
| 审阅通道 `execute` | 写前中止抛 `ReviewNotStartedError` ⇒ 结论 `not_started`（条上写**未执行**）。与既有 `noActionStarted` 那条**不合并**：那条读账本现场，登记失败时账本是空的，条件恒假 |

不新增任何"执行前确认"；不跳过登记、不删来源证据、不放宽授权。

### 8.2 断言（红→绿）

- `server/tests/test_wps_formula_resume.py`（11 条，新建）：四表三公式同形状**必须登记**；
  六种拒绝各报各码；v2 字面值计划的两条边界不放宽；区域解析兜底闸直接测（HTTP 层被 pattern 挡在前面）。
- `apps/wps-addin/src/wpsFormulaResumeRegistration.test.ts`（4 条，新建）：同一份计划的前端载荷
  （`version 1` + `来源!A1:B5` 5×2 + 三个 `writeFormulas` + 无值矩阵）**以及**
  "登记成功 → 前两步写入并回读 → 第二步回执前中断 → 第三步一格没写"。
- `WpsAgentPanel.test.tsx`：`登记失败（零写入）：结论是「未执行」，不留虚构的在途核对入口`、
  `服务端拒绝原因原样到达界面`；原"登记失败零写入"负向用例保留。
- 既有 `test_wps_resume_context_rejects_unsafe_metadata[formula]` 是**把缺陷写成了断言**（它要求
  `writeFormulas` + 来源证据必须 422），改为 `[format]`：要守的是"格式动作没有来源值"那半边。

### 8.3 未做

真机复测（本轮没有打开真实工作簿），所以只能记"已修、待复测"。**→ 复测已于同日 21:46–22:10 完成，见 §9。**

---

## 9. 上条修复的真机复测：WPS 公式 A 分支通过（2026-11-01 21:46–22:10）

**结论：WPS 公式 A 真机操作链路通过**（登记 → 前两步写入并回读 → 第二步回执前中断 → 完整退出重开 → 核对 → 续跑只写第三步 → 确认变更）。
逐项五列台账在 `docs/CONVERSATION_UX_REFACTOR_2026-11-01.md` §8.5 F，此处只留索引与判据。

| 维度 | 证据 |
| --- | --- |
| executionId | `wps-execution-304e6eee-5ba7-4131-b01a-7c54aaea4e32`（当天唯一一条 WPS 执行，`task-a6b33fa7…`） |
| 用户实机 | 六步逐条确认（三格公式、第三步为空、重开恢复卡可用、核对后"可继续剩余动作"、继续后 15、确认变更成功）；**截图待定位** |
| 保存文件（只读） | `C:\Users\jia\Desktop\测试\公式来源证据验收-WPS-A.xlsx`，11032B / mtime `22:15:45` / sha256 `ecf3a6ad…8a98bc8`；`结果1/2/3!B2` = `=来源!B2*1` / `=来源!B4*2` / `=来源!B5*3`，缓存值 `2/8/15` |
| 账本 | `status=completed` / `revision=11` / `next_action_index=3`；`created_at=21:46:09`、`updated_at=22:10:10`；`result_json={"actionCount":3,"backfilled":false}`；`checkpoint` 指向 `结果3!B2` |
| 写序（关键） | `office_host_writes` 恰好 3 行：`idx0`+`idx1` @`21:46:09`（`idx1`=`failed_after_write`/`WPS_ACTION_FAILED`）、`idx2` @`22:10:10` ⇒ **续跑只执行第三步**。该表只追加、`recorded_at` 由服务端取时，时间差即写序；第二步**回执补记**不产生新行，故可与"第二次宿主写入"区分 |
| 自动化 | 修复轮 `wpsExecutionResume` + `wpsNativeComparison` 110 passed；全量 core 1128 / addin 1734 / wps 1283 / server 2026 |

**不能据此宣称的东西**：

- 账本里**查不到** `planning_tasks` / `planning_task_events` 的该 `task_id`，`folder_operation_events` 计数 0 ⇒ 没有"登记 / 开始 / 回执 / 核对"的逐事件序列，**"核对"这一步在账本里没有独立事件行**（证据不足，不补推断）。
- 「确认变更」只有用户界面证据，账本只有 `completed + next_action_index=3`，**没有独立确认留痕**。
- `undone_at=NULL`、`recovery_package_digest/bytes/retired_at/removed_at` 全 `NULL` ⇒ **"撤销入口可用"不等于本轮撤销通过**，本轮未做撤销，C1/C2 两格仍未验。
- **WPS 公式 B 仍未验**；A1-GATE **不通过**、A2 未启动。

**新登记的展示缺陷（只记录、未改代码）**：续跑进入结果审阅后，过程面板仍显示「回读核对第 2 个动作 / 进度 2/3 步 / 正在回读核对「结果2!B2」…」。
出处 `apps/wps-addin/src/wpsActivityLog.ts:126-150`（`verifying:<index>` 条目累积、续跑开始不重置）与 `wpsPlanExecutor.ts:2578`。复现与影响见 UX 台账 §8.6。
