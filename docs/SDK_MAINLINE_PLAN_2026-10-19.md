# SDK 主线计划与验收表

更新日期：2026-11-01（现有证据对齐；未新增原生操作）。本文是当前唯一主线排期；保留原文件名以兼容已有链接。

**当前位置：A0 规划恢复已实现但未完整验收；A1 执行恢复正在收尾；A2 尚未开始。**
正常分析、写入和审阅已有双端实现及用户实机反馈，不能因此把全部异常恢复能力标为完成。

计划整理后的第一批修复已执行：A1-F1 已修复并完成双端自动化；A1-F1/F2 仍待原生复验。2026-10-24 收尾核查轮
逐项复核了七条故障清单，**未发现新的产品缺陷**（生产源码未改），补了 2 条回归、3 个单点变异全部变红，
并重跑了双端全量与构建（数字见第 5 节末尾）。原计划逐字保存在
[历史记录](SDK_MAINLINE_HISTORY_THROUGH_2026-10-24.md)；历史中的“全部未实机”“WPS 后缀续写未实现”等结论已被后续子场景证据部分取代，不再用于当前排期。

## 0. 当前收口队列（2026-11-01，优先于下方历史批次状态）

本轮只读代码与已有验收记录，不重新读取真实工作簿或账本。既有事实引用 [真机台账 §8.5](CONVERSATION_UX_REFACTOR_2026-11-01.md#85-既有真机证据按宿主动作类型分支逐项2026-11-01-重新核对纠正上一版)；不能把历史记录的数字称作本轮实测。

| 项目 | 最新结论 | 下一步 |
| --- | --- | --- |
| Excel 公式 A/B、同会话及重开撤销 | 10-31 已有实机、保存文件及账本证据；B 的 undone_at 有值 | 保留通过范围，截图待定位；确认接受结果无单独留痕，不要求重跑整个分支 |
| WPS 值动作跨进程续跑/放弃 | 10-28/29 已有原生与写序证据 | 旧轮 NULL 不否定用户操作；新版 disposition 与材料回收仍需新证据 |
| writeFormulas 精确续写 | 双端已实现；Excel A/B 真机通过；**WPS A 于 11-01 21:46–22:10 真机通过**（`wps-execution-304e6eee…`，写序 3 行分两批，续跑只写第三步）；**WPS B 未验** | 只补 **WPS 公式 B**（改来源 → 重开续跑被拒 → 目标不变 → 撤销），按公式验收文档 §3–5 B 段；两表值复制终态夹具不能替代 |
| copyRange 精确续写 | 已有 values 子集自动化：展开成 writeValues，前两步中断后仅写 Target!A3 | 本轮定向复跑 1 passed / 65 未选中；不外推格式复制、link 或原生能力 |
| S1 注入来源失效停止 | 10-26 双端已有原生观察及保存副本证据 | 保留注入场景范围，不代替全部手动改源/交错矩阵 |
| 格式/结构精确恢复、缺来源快照 Python 结果恢复 | 尚未闭环 | 继续设计/实现，不能写成只差实机；不得放宽来源守卫 |
| 崩溃恢复、跨窗格竞争、引用过期完整原生链 | 仍有缺口 | 按场景与现有工件核对，不用正常退出/双击/替身宿主冒充 |
| SDK 原生 session resume | 未接入，no-session-persistence 仍开启 | S 阶段隐私决定；应用检查点恢复不等于 SDK 原生会话恢复 |

**下一批优先：WPS 公式 B → 新版终态处置与材料回收证据 → 剩余故障矩阵逐格收口 → A1-GATE 复判。A2 仍未启动。** WPS 公式 A 已于 11-01 真机通过（见上表与台账 §8.5 F），但 **A1-GATE 不因此整体通过**；后两项不是“一次真机就能全部完成”。

## 1. 固定交付约定

- 在用户明确授权的任务和所选工作表内，白名单 AnalysisPlan 通过校验、来源复核、恢复准备后执行一次，再回读并供审阅。不得增加执行前确认按钮。
- 确认变更只接受结果并清除标记，不再执行；放弃变更使用已验证恢复点。查询或预览任务不擅自写入。
- 读取状态、继续规划、继续执行、重新规划、撤销是不同操作；不能以重新跑整份任务冒充精确续写。
- 不扩大选表授权，不跳过工作簿身份、来源、目标或执行权检查。恢复材料不足时明确拒绝或进入受限重新规划，不降低守卫。
- Excel、WPS 分别记录实现和证据。共用语义优先放共享层；宿主差异留在适配层。
- 2026-10-26 用户已追加授权使用 computer-use 和操作工作簿。本轮仅使用独立生成的验收副本；原生观察和替身宿主测试分别记录。

## 2. 当前阶段表

| 阶段 | Excel | WPS | 当前完成判定 / 剩余差距 |
| --- | --- | --- | --- |
| 正常 SDK 业务链 | 已实现；用户反馈自动写入、双视图、确认、放弃、撤销正常 | 已实现；用户反馈写入、审阅及重新计算结果正常 | 成功路径已获部分原生证据；不覆盖所有动作与故障场景 |
| A0 规划暂停与继续 | 应用检查点、只读核对、续跑桥已实现 | 同类能力已实现；来源变化后的旧继续入口已补收尾 | 未完整验收。需真实 SDK 继续时读表/工具回执链证据；不能再说用户从未点过继续 |
| A1 在途写入核对与续写 | 受支持区域写入已实现；用户恢复审阅反馈正常 | 有条件的 writeValues/writeFormulas 后缀续写已实现；值动作已有原生反馈，公式原生未验 | 未整体完成。原生实际写入次数、过期全链、跨窗格竞争等尚缺 |
| A2 账本生命周期 | 未实施本阶段收口 | 未实施本阶段收口 | A1 达标后实施；保留期、终态去值、拒重放身份、清理与迁移 |
| S：SDK 原生会话恢复 | 未接入 | 未接入 | 先做转录隐私夹具实验和生命周期设计，再由用户决定启用 |
| B：SDK 自主受控脚本工具 | 受限解释器已有，未注册为 SDK 自主工具 | 同一共享解释器；SDK 自主工具未接入 | 复用解释器生成白名单计划，不开放任意脚本执行 |
| C：WPS 能力与工作目录 | 作为能力对齐参考，已有能力需分项核实 | 目录选择/授权已有；服务端查询、计算、执行、恢复链未接通 | 未实现部分不能写成实机待验；筛选、去重、结构/对象能力先更新准确差距表 |
| D：架构与交付清理 | 待实施 | 待实施 | SDK 共享依赖抽离、安装包 CLI 生命周期、确认无调用后的旧路径清理 |

原计划把“原生会话恢复”和“自主受控脚本”都叫 B，本文把前者记为 S，仅消除编号歧义，不改变优先级。
执行顺序：A0/A1 收口 → A2 → S 隐私实验与启用决策 → B → C → D。未获原生会话启用决策时保持关闭，不自行删配置开关。

## 3. A1 当前验收矩阵

“自动化”包括真实生产组件/执行器与替身宿主；“原生反馈”是用户截图及操作反馈。两者不能合并冒充一条完整原生调用日志。

| 场景 | Excel 当前证据 | WPS 当前证据 | 剩余工作 |
| --- | --- | --- | --- |
| 写入成功但回执丢失 | 实际执行器与面板自动化验证 applied / not_applied / unknown；用户恢复审阅正常 | 在途核对、只补记不重写、最终回执补记已有自动化和用户分步反馈 | 2026-10-26 S2 已补同链 ID、调用与回执；WPS 完整 WebView 重载仍待验 |
| 第二步中断，只续第三步 | 2026-10-26 原生精确中断、WebView 刷新、双击恢复通过；完整宿主调用记录 | 同日原生核对、双击续写与完整宿主记录；关窗复用页面 | 前缀未重写已有同链证据；WPS 完整重载/原生跨窗格竞争仍待验，见 S2 记录 |
| 运行中异常退出 | 独立 worker 强杀后恢复 SQLite 状态已有 | 同类测试；租约过期用测试时钟 | 原生宿主/WebView 异常退出证据；现有 worker 不实际写单元格 |
| 并发恢复 / 旧执行者 | 同帧双击互斥、真实 API/SQLite、独立进程唯一认领已有；**8 线程同 revision 并发认领成功数恰好 1**（`test_recovery_claim_concurrency.py`，C1/C2 变异各 8 成功即红） | 同类证据；旧 owner 被拒；CAS 独立变异仍有租约冗余层；同一份并发认领证据覆盖规划侧（两层版本判据一次拆全才红） | 两窗格原生竞争 |
| 来源变化 / 换簿 / 授权收窄 | 相关守卫与回归已有；**`source_changed` 是本次规划的终态**（桥停止派发、SDK 循环 break、409 `SOURCE_CHANGED`），面板原文显示且无继续入口（第八轮） | 同类守卫已有；用户实测来源 222→250 后拒绝旧规划，新任务计算并写入 500；停止行为同上（第八轮，含 WPS 面板用例） | 双端恢复等待期间完整交错矩阵；2026-10-26 已完成真实 Excel 注入停止链观察；WPS 注入停止也已完成原生观察。真正手动改源的原生检测仍待验；来源变化不是引用过期 |
| 旧结果迟到 / 任务串用 | 桥代次、段归属、旧 owner 拒绝等已有 | 同类守卫已有 | 两端模型迟到、同名工作簿/其他任务串用的完整恢复链及独立变异 |
| 结果引用过期 | 结果可读→404→连续重开不复用；真实 API 拒绝旧引用 | 同类证据 | 真实 SDK 产出与双端替身宿主消费已串联（A1-E2）；**B 路（模型确实再读一次被 `source_changed` 拒交值）已补真实模型证据**（§ 2026-10-25 第七轮）；原生同链写入次数、故障交错仍缺 |
| 恢复材料缺失 / 过期 | 缺少前缀恢复快照时在读写前拒绝 | 真实本机材料过期/删除测试；重新规划→新执行→确认→重开全链自动化 | Excel 没有同款过期重新规划入口；不能强加 WPS 30 分钟机制；任意新计划不保证跳过前缀 |
| 完成后连续重开 | 自动化无重复写入；清发现索引 | 自动化无重复写入/补记；允许只读历史提示 | 终态/确认/撤销与重开组合证据补齐并归档 |
| 完整退出宿主（浏览器索引 + 宿主缓存都丢） | 同一发现入口已接线（`/api/office/executions/pending`），面板级回归未补 | 自动化：经后端发现找回入口、显示待核对、零自动动作；自动化无重复写入/补记 | 双端真实宿主**完整退出**后重开一次；真实窗格里点一次「核对在途动作」 |

### 恢复能力的明确限制

- WPS 精确续写当前限于：已登记 resumeContext、本机恢复材料在有效期内（当前 30 分钟）、来源与目标表分离、完整来源快照、目标不重叠的 writeValues/writeFormulas；copyRange 仅限满足同样证据条件的值复制展开。
- WPS writeFormulas 精确续写已实现；值复制 copyRange 经 writeValues 展开已有定向用例。格式/结构动作及缺少来源快照的 SDK Python 结果计划仍未闭环，不标为只差实机。
- Excel 恢复依赖既有计划和检查点；已完成前缀缺少恢复快照时必须拒绝，不能生成只覆盖后缀的撤销点。
- 如需分批交付 A1，必须明确列出双端动作范围与未实现项，并取得用户同意；未完成项不得靠改状态表消失。

## 4. 下一批工作队列

| 顺序 / ID | 工作 | 状态 | 验收条件 |
| --- | --- | --- | --- |
| 1 / A1-F1 | WPS 审阅标记 Formula1 空对象故障 | 代码完成；自动化通过（含收尾轮 2 条补测 + 3 个变异）；**原生待复验** | 宿主返回 null/无法建立规则时不崩溃；不删除用户规则；已写结果、审阅与清理状态如实；补 WPS 回归及 Excel 非回归 |
| 1 / A1-F1-a | 标记部分失败后仍能清掉**本次**残留（含“无法确认时明确提示”） | 2026-10-25 列为待修复 → 已修复（含第三轮副本判据、第四轮 `ownedRules` 同判据）；自动化通过；**原生待复验** | 失败时仍交出完整随机公式；双视图副本只按它清；`unconfirmed>0` 必须提示；读不出条数、空句柄、表达式缺公式、类型未知一律 unconfirmed，只有宿主**明确**给出非表达式类型才跳过（**一份**白名单，副本清理与归属扫描共用）；找回与重载清理遇到同区域色阶不得阻断；每个区域最多计一次；不按前缀/颜色/区域批量删 |
| 1 / A1-F1-b | 归属记录单槽被新审阅覆盖 | 2026-10-25 列为待修复 → 已修复；自动化通过；**原生待复验** | 新旧两轮记录并存、各自清理；兼容 v1；不同簿不串用；存储无单元格值 |
| 1 / A1-F2 | 来源失效旧暂停入口及审阅控制栏 | 代码完成；自动化通过（含“独立暂停任务不被一起作废”回归）；**原生待复验** | 来源变化拒绝后不再允许旧任务续跑；审阅时显示等待审阅；确认后不复活失效任务；独立暂停任务保留 |
| 2 / A1-E1 | 引用过期的精确故障入口 | 2026-10-25 已实现；自动化通过（服务端 12 条 + 真实登记路径 1 条；共享判据 14 条、控件 13 条、WPS 目标 6 条、WPS 面板接线 2 条；10 个单点变异）；**原生待复验** | 入口仅开发模式、用户显式单次启用、绑定具体任务与结果引用；生效位置 = 「已登记、尚未首次写入」；旧引用在任何宿主写入前被拒且句柄已删（无缓存正文可复用）；不改生产 TTL、不清空全部结果、不借用第二步写入中断；两端接入，判据在共享层；不加执行前确认 |
| 2 / A1-E2 | 真实 SDK 过期恢复全链 | 2026-10-25 已串联真实链（真实模型 + 真实服务端 + 真实登记器 + 真实 Docker）：两端各一份工件，旧引用 404 / 新引用可读 / 计划用新引用编译；消费侧两端各 3 条门控用例（执行一次 → 回读 → 确认不重执行 → 放弃/撤销回执行前）；`npm run test:a1-e2` 一条命令复跑（门控用例修完已整套重跑：lint / 双端 test+build / pytest / compileall / `git diff --check` / `test:a1-e2` 退出码全 0）。**「恢复期间来源变化」A 路（模型没有再次读取）已闭环**：确定性屏障在首次交值后改源，两轮真实 SDK 都是 `prepareCount=1 / refusals=[] / planSourceBasis=old`，并由此挖出并修掉一个真实缺陷（计划指纹取"到达那一刻"⇒ 旧结果被静默写入；WPS 更本没有计划级写前复核）；**B 路（模型确实再 prepare）已补真实模型证据**（第七轮：两端各一轮有效实验，第二次真实工具调用被生产共享判据拒 `source_changed`，随后的旧计划在两端真实面板上 0 写入）；本轮补齐缺锚点拒写、响应指纹清除、WPS 本地与 SDK 来源分类及顺序复核，双端全量与开门控用例通过；**原生实机未做** | 保持来源不变以单独验证过期；旧引用拒绝；SDK 重新读算并生成新引用；新计划执行一次并回读；确认不重写、放弃可恢复；两端分别留日志 |
| 3 / A1-E3 | 补齐故障矩阵与原生写入计数 | 自动化部分完成；2026-10-26 验收工具缺陷已修，Excel/WPS S1 注入停止均已取得原生观察与精确账本证据；S2 续写漏报观测已修并实机复测，WPS 完整重载待验 | 12 行双端缺口矩阵 + `source_changed` 终态修复 + 并发认领真并发证据 + 19 个单点变异（见 [A1-E3](SDK_A1_E3_STOP_AND_GAPS_2026-10-25.md)）；第九轮修掉**未保存同名簿的读路径身份撞车**与**验收入口可被别的簿抢走**，并把"实际写入次数"改成四档口径；**第十轮**补齐三处：①「没有观测」不再冒充「明确零写入」（`HostWritesObservation`，只有 `complete+zero` 才算零）②**WPS 宿主写调用计数已实现**（作用域式，计划写与辅助写分开）③故障入口的目标收敛到**可核验工作簿身份**（`dev_fault_target`，两个开发入口共用）；**第十一轮**把 S1 的准备/故障绑定/证据采集做成脚本（`scripts/a1e3_native.*` + `make_a1e3_native_fixture.py`），用户只负责在窗格里发需求点按钮；2026-10-26 另修验收脚本的实例/任务归属、观测口径、进程与目录隔离；当前原生剩余见本轮记录，不把脚本准备当作验收通过。**原生待完成项**（[原生验收包](A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md) S1–S3：来源变化停止、双击只续剩余、确认/放弃后重开） |
| 4 / A1-GATE | A0/A1 阶段验收 | 未达标（2026-10-27 先补持久恢复材料链、再补**发现入口**（索引降级为缓存 + 后端只读发现），两批均仅自动化；原生未做） | 九场景逐项结论明确；原生与自动化证据分别齐全；支持范围明确；关键未闭环项解决或获批分批交付；相关全量检查通过 |
| 5 / A2 | 生命周期收口 | 未开始，依赖 A1-GATE | 清理后旧 operationId 仍拒重放；终态去值与保留身份；清理失败可见；损坏/旧版本记录行为明确；历史转录清理有入口和删前确认 |

E2 的来源变化反向分支单独验：若来源被改，不能拿旧计划续写；重新发起新任务成功不算同一任务恢复成功。
不得要求用户再次重复已经验收的普通写入/确认流程来代替缺失的过期故障注入。

## 5. 最近反馈与验证基线

| 项目 | 最近可用证据 | 限制 |
| --- | --- | --- |
| WPS 当前来源重新计算 | 用户截图：来源 111 / 250 / 333，目标 222 / 500 / 666，黄色标记、审阅按钮、确认后清标记均可见 | 是重新发起后的成功路径，不是 resultRef 过期恢复；未据截图确认写入次数 |
| Excel/WPS 审阅恢复 | 用户此前确认双视图、放弃、确认、撤销正常；WPS 后续补记完成进入审阅 | 不扩张为所有动作、原生崩溃或跨窗格竞争通过 |
| 暂停残留修复 | 上一轮 Excel 1538 passed / 3 skipped；WPS 1062 passed / 3 skipped；两端 build 含 tsc 通过；定向 diff-check 通过 | 日志 outputs/pause-terminal-*；本次仅引用，未重跑；真实窗格尚未复验 |
| 服务端 | 最近已记录全量 1748 passed / 1 skipped；之后结果引用相关定向 63 passed | 非暂停修复批次的全量结果，不当作本轮全量重跑 |

### 2026-10-24 第一批：A1-F1 标记空句柄修复

- 已复现代码缺陷：WPS `FormatConditions.Add` 返回 null 时，创建和清理路径都直接访问 `Formula1`；`Item` 返回 null 或 `Count` 非法时，重载清理误判规则不存在并清除归属记录。原生截图的实际宿主返回序列尚无日志，不能宣称已证明唯一实机根因。
- 修复：Add 前保存本次随机公式和区域；空返回时只按完整随机公式找回唯一规则，不重试 Add；创建后抛错也能清理本次规则。集合不可读时保留待清理记录；绝不按前缀、颜色或区域清空用户规则。
- 本次新增 7 条测试：5 条故障回归在修复前全部失败、修复后通过；另有 2 条使用真实共享审阅控制器及 WPS 标记适配器的确认/放弃回归，比较可用且 execute 始终一次。宿主和业务执行/恢复为替身，不等于原生 Office/WPS 验收。
- 本轮全量：WPS 1069 passed / 3 skipped；Excel 1538 passed / 3 skipped；两端 build（包含 tsc）退出码 0。WPS 构建仍有包体积提示，测试有 React act 提示，未隐藏。后端未改，本轮没有重跑后端全量，也没有提交 Git。
- 日志：`outputs/marks-null-red.log`（5 failed / 10 passed）、`outputs/marks-null-green.log`（当时 15 passed）、`outputs/marks-null-wps.log`（包含最终 17 条标记测试）、`outputs/marks-null-addin.log`、`outputs/marks-null-build-wps.log`、`outputs/marks-null-build-addin.log`。

原生复验（用户操作，不使用 computer-use）：

1. 加载新前端，在独立测试工作簿保留原固定选表范围。用“来源表 B2:B4 乘 2 写入目标表 B2:B4，来源表不修改”的既有案例运行一次；当前来源 111 / 250 / 333 时目标应为 222 / 500 / 666。
2. 检查没有 `Formula1` 空对象报错，目标区域出现审阅标记；打开双视图，再确认变更，值不变、标记消失。另一次独立执行可点放弃变更，检查恢复到该次执行前状态，不重跑已确认任务。
3. 若仍有“标记状态待核对”，保留提示并反馈；不要手动删除全部条件格式。自动化覆盖空返回/暂不可读，但实机是否出现这类返回仍需现场证据。
4. A1-F2 另测：暂停后改变来源，再继续应拒绝旧任务且撤下旧 ▶；按新来源发起新任务后进入审阅，确认后旧暂停入口不复活。不得以普通成功写入替代这条复验。

第一步代码工作已完成，原生复验仍待用户；A1-E1（结果引用过期精确故障入口）的实现与自动化证据已完成（见 2026-10-25 小节），**A1-E1 完成不等于 A1-E2 已验证**：真实 SDK 的"过期 → 重新读算 → 新引用 → 只写一次 → 确认不重写/放弃可恢复"全链仍未串联验证，A1 阶段总体验收尚未完成。

### 2026-10-24 收尾核查轮：A1-F1 / A1-F2（未改生产源码）

按七条故障清单逐项对照**生产路径**后，结论是**没有发现新的产品缺陷**，因此本轮没有改生产源码；补了 2 条行为回归，并用 3 个单点变异证明断言有判别力。

| 清单项 | 生产路径上的判据 | 本轮证据 |
| --- | --- | --- |
| Add 空返回但实际创建成功 | 空返回只按**完整随机公式**在集合里找回唯一规则，绝不重试 Add | 变异 M1（退回“单次 Add 直接取句柄”）→ 2 条红 |
| Add 创建后抛错 | 创建**前**先写归属记录；已验证句柄进 `handles`，catch 直接按句柄删 | 既有 `throwAfterAdd` 用例仍绿 |
| 集合暂不可读、后续恢复可读 | `ownedRules` 对 Count/Item/Formula1 三处都按“读取失败 ≠ 规则不存在”处理，一抛错就保留记录 | `nullItem` / `invalidCount` 两态绿 |
| 部分清理失败、重开后重试 | catch 只把**真正没清掉**的区域写回记录（不是“处理过的全部”） | 新增用例：1 处未清 → 宿主恢复 → 重开只清它并保留用户规则；变异 M2（写回 `tracked`）→ 2 条红 |
| 同区域用户规则 / 其他任务标记 | 归属判据只有完整随机公式；`deleteOwnedRules` 先全量校验再删 | 用户规则与另一轮 `EB_REVIEW_` 规则均保留 |
| 标记异常后确认 / 放弃 / 双视图 | 标记失败降级为 notice，仍进 `reviewing`；确认只接受结果不重执行 | 面板级「颜色标记未完整建立」+ 控制器级 confirm/discard |
| 来源失效旧任务 / 独立暂停任务 | 只按 `turnId` 作废该条；落盘 `resolved` 使重开不复活 | 新增「只作废该任务」用例；变异 M3（去掉 `turnId` 过滤）→ 1 条红 |

本轮实跑（不是引用上一轮）：

- WPS `npm run test:wps`：**1071 passed / 3 skipped**（66 files passed / 2 skipped），退出码 0，含本批新增 2 条；`outputs/a1-round-wps-full.log`。`npm run build:wps`（`tsc --noEmit` + vite）退出码 0，仅既有 >500 kB chunk 提示；`outputs/a1-round-build-wps.log`。
- Excel `npm run test:addin`：**1538 passed / 3 skipped**，退出码 0；`npm run build:addin` 退出码 0；`outputs/a1-round-addin-full.log`、`outputs/a1-round-build-addin.log`。
- 服务端：`python -m pytest server/tests -q`（venv，`--basetemp` 在 OS 临时目录）**1750 passed / 1 skipped**，退出码 0，112s；`compileall -q server/app` 退出码 0；`outputs/a1-round-pytest-full.log`、`a1-round-compileall.log`。仅有既有的 `.pytest_cache` 无写权限 warning。
- 变异：M1/M2/M3 各只改一处、命中唯一；运行后**逐字节还原**并与备份 sha256 比对一致（markers `770ff269…`、panel `cad381b9…`）。日志 `outputs/a1-f1-mut-M1.log`、`a1-f1-mut-M2.log`、`a1-f2-mut-M3.log`；驱动 `outputs/mut-backup-a1/mutate.py`。
- WPS 测试仍有 40 处 React `act` 警告，未隐藏；不据此改断言。

仍缺的原生证据（自动化不能替代）：

- 真实 WPS 里 `FormatConditions.Add` 是否真的返回 null、`Item` 是否真给空句柄：现在测的是**替身构造的返回序列**，实机序列没有日志。
- 真实窗格一次完整的「标记失败 → 审阅 → 确认/放弃 → 双视图」点击；面板级仍是替身宿主。
- 已知残留（本轮**未**改，也不夸大）：①标记失败但残留规则时拿不到本次公式，执行前副本会带着残留高亮；②归属记录是单槽，上一轮失败记录未清就开新审阅时旧记录会被覆盖，旧规则失去自动清理线索；③Excel / WPS 真实窗格一次都没点过。①② 已在 2026-10-25 单独修复（见下一节），③ 仍未闭环。

原生复验步骤（用户操作，不使用 computer-use；独立测试工作簿，保持原固定选表）：

1. 跑一次既有案例「来源表 B2:B4 乘 2 写入目标表 B2:B4，来源表不改」（来源 111/250/333 → 目标 222/500/666）：确认无 `Formula1` 空对象报错、目标区域有审阅标记，点「双视图对比」确认执行前副本不黄，再「确认变更」——值不变、标记消失。
2. 同一工作簿再跑一次，这次点「放弃变更」：确认主表回到该次执行前、标记撤下，且没有第二次写入。
3. A1-F2：先让一条任务暂停 → 修改来源数据 → 点继续：应拒绝旧任务并撤下旧 ▶；随后按新来源发新任务进入审阅，确认后旧暂停入口不复活。
4. 若出现「标记状态待核对」，保留提示并反馈现场（不要手动删除全部条件格式）；重开窗格应能自动清理该处。

### 2026-10-25：A1-F1 两处残留修复（F1-a / F1-b）

上一节记录的 ①② 两条残留本轮修掉了，**生产路径确实改了**（`wpsResultReviewMarkers.ts`、`wpsNativeComparison.ts`、`WpsAgentPanel.tsx` 两处审阅接线）；其余已通过的修复没有重做。

**修了什么**

- **F1-a 本次公式不再丢**：`markWpsResult(application, plan, onFormula?)` 在生成公式后、**第一次 Add 之前**回调公式。面板（两处审阅入口）把它存进 `currentMarkFormula`，`showBefore` 照旧传给 `stageWpsBeforeComparison` —— 标记建到一半失败时，双视图副本仍能按**完整随机公式**清掉本次残留。归属判据仍然只有完整公式；不按前缀/颜色/区域删。
- **F1-a 清理结果如实上报**：`clearReviewMarksFromCopySheet` 返回 `{cleared, unconfirmed}`，`WpsNativeComparison.reviewMarks` 带出；`unconfirmed > 0` 时面板在审阅条上说「执行前副本的审阅标记未能核实清理（N 处区域）…请以原表与回读结果为准」。没有公式时照旧不清（宁可副本高亮，也不误删）。
- **F1-b 记录从单槽改成多条**：同一 key 里改成 `{v: 2, records: [...]}`，每条 = 簿身份 + 本次随机公式 + 表名/区域；`upsertRecord` 只动**同一条公式**，`removeRecord` 只删本次，确认/放弃不再 `forget` 整个记录。`recoverWpsResultMarks` 只处理**当前簿**的记录，另一簿的记录与规则原样保留（连被复制过来的那条也不碰），两条各清各的。
- **兼容 v1**：`readStored` 把旧的 `{v: 1, workbookFullName, formula, ranges}` 当一条记录读出来，下一次写回时升级成 v2；旧记录一条都不丢。记录解析**只接受**四个字段，业务单元格值既读不进也写不出。
- 顺手删掉已无调用点的 `forgetWpsResultMarks`（清空整个 key 是这次要消灭的行为）。

**先红后绿（可复现）**

- 修复前：`outputs/a1-f1-red-round2.log` —— **14 failed / 21 passed**，失败点正是缺公式（`seen` 为空）、没有清理报告（`reviewMarks` undefined）、单槽只剩 1 条记录、`records` 形状不存在。
- 修复后：`wpsResultReviewMarkers.test.ts` 26、`wpsBeforeComparison.test.ts` 9、`WpsAgentPanel.test.tsx` 176，全绿。
- 单点变异（各只改一处、命中唯一、还原后 sha256 与备份一致）：M4 面板不交公式 → 1 红；M5 登记退回单槽 → 3 红；M6 副本读不出来不报告 → 2 红；M8 恢复不看簿身份 → 2 红。日志 `outputs/a1-f1-mut-M4..M8.log`，驱动 `outputs/mut-backup-a1r2/mutate.py`。
- 覆盖清单：标记部分失败后开双视图（面板级 + 适配器级）、旧记录未清理时开新审阅、两轮记录重载后分别清理（含一处清不掉只留它自己的记录）、用户规则与其他任务标记不受影响、确认/放弃各只执行一次且残留记录不被丢弃。

**本轮实跑**

- WPS `npm run test:wps`：**1081 passed / 3 skipped**，退出码 0（`outputs/a1-r2-wps-full.log`）；`npm run build:wps`（tsc + vite）退出码 0（`outputs/a1-r2-build-wps-round2.log`；首跑因 `readonly` 与 `satisfies` 冲突 TS4104 退出 2，改类型后通过 —— 记在这里，不当成没发生）。
- Excel `npm run test:addin`：**1538 passed / 3 skipped**，退出码 0；`build:addin` 退出码 0（`outputs/a1-r2-addin-full.log`、`a1-r2-build-addin.log`）。
- 服务端 `pytest server/tests -q`（venv，basetemp 在 OS 临时目录）：**1750 passed / 1 skipped**，退出码 0（`outputs/a1-r2-pytest-full.log`）。
- `lint` 0 error / 36 既有 warning；`git diff --check` 退出码 0。

**A1-F1 原生验收（单独记录，自动化不能替代）**

1. 独立测试簿跑既有案例（来源 111/250/333 → 目标 222/500/666），确认标记与审阅正常、双视图副本不带本次高亮。
2. 若宿主真的出现 `FormatConditions.Add` 空返回/删不掉：界面应显示「颜色标记未完整建立…」；此时点「双视图对比」，副本不应带本次黄色高亮；若显示「未能核实清理」，把该提示与现场一起反馈（**不要**手动删全部条件格式）。
3. 上一次审阅留有标记没清掉时，直接发起**第二次**任务并确认：新任务清标记只清自己那一条，旧的高亮仍在；重开窗格时旧的应被自动清理并给出提示。
4. A1-F2 保持原步骤（暂停→改来源→继续应被拒且撤下旧 ▶）。

仍未闭环：真实 WPS/Excel 窗格一次都没点过（F1-a/F1-b 的证据是生产组件 + 替身宿主）；重开时的清理提示只能在真实宿主上看到文案效果。

### 2026-10-25（第三轮）：A1-F1-a 副本清理判据补齐

上一轮把"本次随机公式"交到副本侧了，但**副本侧怎么读集合**还有一处会说反话：读不出来时默认成"没有"。本轮只动 `clearReviewMarksFromCopySheet`（拆出 `clearOneCopyRange` + 非表达式类型白名单），多记录存储与公式提前传递的修复**原样保留**，没有扩大重构。

**三个判据漏洞（旧行为 → 现在）**

| 宿主读数 | 旧行为 | 现在 |
| --- | --- | --- |
| `Count` 缺失 / null / 非数字 | `Number(Count ?? 0)` ⇒ 当成 0 条，静默报"已清理" | 读不出条数 ⇒ 该区域 `unconfirmed`，一条都不删 |
| `Item(i)` 返回 null/undefined | `continue` 静默跳过（"这条不存在"） | 该区域 `unconfirmed`；能核实的那条照删 |
| 规则读不出 `Formula1` | 一律当色阶/数据条跳过 | 只有宿主**明确**给出可识别的非表达式类型（色阶 3 / 数据条 4 / 前 10 项 5 / 图标集 6 / 唯一值 7 / 空值 9 / 时间段 10 / 平均值 11 / 非空 12 / 错误 13 / 无错误 14）才跳过；类型未知、类型不是数字、或 `xlExpression`(2) 读不出公式 ⇒ 保留规则 + `unconfirmed` |
| 计数口径 | 抛错与空句柄都可能重复计 | 每个**区域**最多计一次（按 `new Set(addresses)` 去重；同区域读不出多条仍算 1 处） |

归属判据没变：只认本次**完整随机公式**，前缀/颜色/区域一律不构成归属。

**先红后绿（可复现）**

- 先补 9 条回归：`outputs/a1-f1a-red-round3.log` —— **7 failed / 37 passed**，失败点全是 `expected { cleared: 0, unconfirmed: 0 } to deeply equal { cleared: 0, unconfirmed: 1 }`（"读不出来"被当成"没有"）。改完 `outputs/a1-f1a-green-round3.log` 44/44 绿。
- 单点变异（各只改一处、命中唯一、还原后逐字节 sha256 = `c7ca2a0d…`）：**M9** Count 当 0 → 3 红；**M10** 空句柄跳过 → 3 红；**M11** 缺公式一律跳过 → 3 红；**M12** 凡缺公式都报未核实（反向：假警报）→ 2 红；**M13** 重复地址不去重 → 1 红。日志 `outputs/a1-f1a-mut-M9*.log`…；驱动 `outputs/mut-backup-a1r3/mutate.py`。
- 面板路径有判别力：M9 在位时面板用例「副本 Count 读不出来：真实清理结论经 comparisonNotice 显示在面板」**单独变红**（1 failed / 176 passed，`a1-f1a-mut-M9-panel.log`）。该用例让结论由**真实**清理函数算出（不是手写 `reviewMarks` 字面量），面板只负责显示 —— 只有 `stageWpsBeforeComparison` 这一步因面板测试整体替身而由适配器自己的用例覆盖。

**本轮实跑**

- WPS `npm run test:wps`：**1091 passed / 3 skipped**（66 files 通过 / 2 跳过），退出码 0（比上轮 +10 条）；`outputs/a1-f1a-wps-full.log`。
- Excel `npm run test:addin`：**1538 passed / 3 skipped**，退出码 0；`outputs/a1-f1a-addin-full.log`。
- 服务端 `pytest server/tests -q`（venv，basetemp 在 OS 临时目录）：**1750 passed / 1 skipped**，退出码 0，105s；`outputs/a1-f1a-pytest-full.log`。
- `build:wps` / `build:addin` 均退出码 0；`lint` 0 error / 36 既有 warning；`git diff --check` 退出码 0。

**A1-F1-a 追加的原生复验步骤（并入上面 1–4 步）**

5. 在测试簿的**目标区域**上手工加一条「色阶」条件格式，再跑一次审阅并点「双视图对比」：**不应**出现「未能核实清理」（这是本轮防的假警报）；色阶规则本身必须还在。
6. 若宿主偶发读不出条数：应出现「未能核实清理（N 处区域…）」，且副本上一处高亮都没被多删（用户规则与色阶都在）。

**已知但本轮未改（同类判据在别的路径上，需单独一轮）**：`ownedRules`（标记失败的重试与 `recoverWpsResultMarks`）对"同区域里的色阶规则"仍按"读不出来"抛错，于是那一处记录会被永久保留、每次重开都提示"未清理"。方向上是**过量提示**（不会谎称干净），但和本轮修的副本侧不对称。要改必须连同它自己的失败回归一起做 —— 放宽它会直接影响"已清理"这句话的资格。

### 2026-10-25（第四轮）：A1-F1 收尾 —— `ownedRules` 复用同一份类型判据

上一节点名的最后一处同类缺陷本轮修掉了。改动只在 `wpsResultReviewMarkers.ts`：把第三轮为副本清理写的 `NON_EXPRESSION_RULE_TYPES` + `isKnownNonExpressionRule` **上移**成两个消费者共用的一份判据（全文件仅此一处定义），`ownedRules` 遇到"读不出公式"的规则时：

- 宿主**明确**给出可识别的非表达式类型（色阶 3 / 数据条 4 / 前 10 项 5 / 图标集 6 / 唯一值 7 / 空值 9 / 时间段 10 / 平均值 11 / 非空 12 / 错误 13 / 无错误 14）⇒ **跳过，不修改、不删除**；
- 空句柄、未知类型、类型读不成数字、`xlExpression`(2) 读不出公式 ⇒ 仍抛错，保留归属记录，不宣称清理完成；
- 归属判据没动：仍只按**完整随机公式**认领，前缀/颜色/区域一律不构成归属。

受影响的三个调用点同步受益：`Add` 回空后的找回、标记失败的 catch 清理、`recoverWpsResultMarks` 的重载清理。

**先红后绿**

- 先补 6 条回归：`outputs/a1-f1-owned-red.log` —— **3 failed / 35 passed**，红的三条都报 `WPS 审阅标记句柄暂不可读取…；1 处标记状态待核对` 或 `上次审阅遗留的 1 处标记未清理`，即色阶把整段挡住。改完 `outputs/a1-f1-owned-green.log` 39/39 绿（含新增的"两个入口分类一致"锚点）。
- 两个方向各一个单点变异（驱动 `outputs/mut-backup-a1r4/mutate.py`，还原后 sha256 = `b2ecda28…`）：
  - **N14 太松**（读不出公式就一律跳过，不认类型）→ **3 红**：未知类型用例、文本类型用例、等价性锚点 —— "没找到"冒充"清掉了"会被抓到；
  - **N15 太紧**（回到旧行为，凡读不出公式就抛）→ **4 红**：色阶找回、色阶重载清理、多任务多簿、等价性锚点 —— 用色阶挡住本次规则的旧缺陷会被抓到。

**本轮实跑（AGENTS.md 清单逐项）**

- `npm run test:wps`：**1097 passed / 3 skipped**，退出码 0（上轮 1091，+6）；`outputs/a1-f1-owned-wps-full.log`。
- `npm run test:addin`：**1538 passed / 3 skipped**，退出码 0；`outputs/a1-f1-owned-addin-full.log`。
- `npm run build:wps` / `npm run build:addin`：均退出码 0。
- `python -m pytest server/tests -q`（venv，basetemp 在 OS 临时目录）：**1750 passed / 1 skipped**，退出码 0，100s；`python -m compileall -q server/app` 退出码 0。
- `npm run lint`：0 error / 36 既有 warning；`git diff --check` 退出码 0。

**Excel 端是否有同类缺陷**：没有。Excel 的条件格式规则**有可读 id**，`clearExcelResultMarks` 只按记录里的 `entry.id` 定位，`find` 之后才校验 `type === 'Custom'`+公式，从不枚举同区其它规则，因此色阶不会阻断任何一步。两端机制不同（WPS 无 id，只能按公式枚举），不存在"需要同步一份白名单"的第二个副本。

**A1-F1 范围内的已知代码缺陷：0 条**（副本清理、`ownedRules` 两条路径都已按同一判据修完并有反向变异钉住）。仍缺的只有**原生证据**（不是代码缺陷）：真实 WPS 窗格里「Add 回空 + 同区色阶」与「重开清理」这两步一次都没点过。

**追加的原生复验步骤（并入上面 1–6 步）**

7. 在目标区域先放一条色阶，再跑一次审阅（制造 `Add` 空返回的情形更好）：应正常出现标记与审阅，**不得**报「未能建立可独立移除的审阅标记」；色阶与其它规则必须原样保留。
8. 上述状态下手动关掉再重开窗格：应清掉本次残留标记并给出「已清理上次审阅遗留的标记」，色阶仍在；**同一状态再重开一次应不再提示**（记录已清）。

当前源码核查：planner.py 透传 on_checkpoint/resume_context；SDK 继承 completedTools/resultRefs，但仍使用 --no-session-persistence；这是应用检查点恢复，不是原生 SDK 会话恢复。
仓库 config/capabilities.json 的 pythonExecution.enabled 仍为 false；原生验收可能使用单独能力配置，本文不推断运行中服务的实际开关。

### 2026-10-25（第五轮）：A1-E2 真实 SDK 过期恢复全链

**做了什么**（详细证据另见 `docs/SDK_A1_E2_RECOVERY_CHAIN_2026-10-25.md`）：

- 新驱动 `scripts/drive_real_sdk_result_expiry.py`：真实模型 + 真实服务端 + 真实登记器 +
  真实 Docker，串起 `读表 → 重算 → 新引用 → 计划编译`，途中由 A1-E1 入口把旧引用精确失效；
  场景 `recovery | source-change | pause`，可用 `--host` 分别跑 `wps-et` / `excel-office`。
- 新一条命令 `npm run test:a1-e2`（`scripts/run_a1_e2_chain.mjs`）：先校验两份工件自洽
  （旧引用 404、计划引用新引用、宿主身份、`armed→bound→consumed`），再跑两端消费侧用例；
  工件缺失或被**跳过**都判非零退出。
- 消费侧新增：Excel `resultExpiryRecoveryChain.test.tsx`（3 条）、WPS
  `WpsAgentPanel.test.tsx` 的 `A1-E2 消费侧` 三条；均以 `EB_A1_E2=1` 门控，默认跳过。

**结论（分开写，不合并）**：

| 项 | 结论 |
| --- | --- |
| 旧引用在真实读取路径上被拒 | **通过**（`consumed` 只可能由 `resolve_result_ref` 触发；旧引用 404，重复取仍 404） |
| 模型收到固定码后重新算并登记新引用 | **通过**（两次 `python.analysis`、两个新引用，`planning.submit` 用新引用） |
| 新计划在两端真实面板上只执行一次 | **通过**（写入条数 = 计划动作数；逐格回读 = 计划值；来源表未变） |
| 确认变更不重执行 / 放弃回到执行前 | **通过**（两端各自断言写入条数不变；放弃后产品自述"已放弃"且目标区回空） |
| 恢复期间来源变化被拒 | **未取到证据**：模型未重读宿主（`prepareCount=1`、`refusals=[]`），桥的 `source_changed` 闸没被走到，该轮仍编译出计划。**不算通过** |
| 用户暂停 | **通过**（`pauseRequested=true`，停在第一次分析之后，无新计划、无写入，检查点 `resumability=continue`） |
| 原生实机（真实 Excel / WPS 窗格） | **未做** |

**lint 36 → 37 的来源**：本批（A1-E1）引入，且是 `react-hooks/exhaustive-deps` 对
"函数体只有 try/catch 的 useMemo"的误判（用最小样例复现）；已让依赖真正参与判断，
37 → 36，未加 `eslint-disable`、未放宽断言。

**一个必须记住的驱动缺陷（本轮踩过并已修）**：请求里**必须**带 `workbookIdentity`
（以及 `sourceFingerprint`）。服务端按 `host + 名字 + 身份` 三者全等判桥归属，少了它就
回落到 `name:<名字>`：轮询永远返回空 `calls`、回执永远 `stale_binding`，症状是每个桥
调用都按 `timeout` 收场（步骤状态 `warning`），而界面上看不出原因。

**门控用例的两条纪律（收尾时各踩一次）**：① 工件不许在 `describe.skipIf` 的工厂里读
（跳过的收集阶段也会执行工厂，缺工件会把整套用例带崩，实测退出码 1；改惰性读后缺工件
179 passed / 3 skipped）；② **门控用例的回归只在门控打开时现形** —— 惰性读取重构后
`runTurnMock.mockResolvedValue({ kind: "plan", plan })` 里的 `plan` 已不在作用域，
默认套件 1105 passed 全绿，`npm run test:a1-e2` 却两条 `ReferenceError`。修完后整套检查
（lint / 双端 test+build / pytest / compileall / `git diff --check` / `test:a1-e2`）已重跑，
退出码全 0，见 `docs/SDK_A1_E2_RECOVERY_CHAIN_2026-10-25.md` §6。

### 2026-10-25（第六轮）：A1-E2「恢复期间来源变化」+ 写前来源锚点缺陷

**约束（用户给的，执行时逐条守住）**：不换来源表逼模型 `prepare`、不改提示词诱导工具顺序、
必须覆盖"模型没有再次 prepare"这一档；屏障控时序不用 sleep；模型工具调用次数 / 实际来源读取 /
计划编译 / 宿主实际写入**分开记录**；旧来源结果被拒时写入必须是 0；两端分别验证；
真实 SDK 本轮运行与历史工件消费分开记；改门控用例后必须跑 `test:a1-e2`。

**做了什么**：

- 驱动 `scripts/drive_real_sdk_result_expiry.py`：来源变化分支改成**确定性屏障** ——
  第一次 `prepare_python_data` **交值成功之后**改 `订单表!B2` 100 → 250，并记下
  `sourceChange{afterServedCall, before, after, fingerprintBefore, fingerprintAfter}`；
  新增 `planSourceBasis`（只从"桥实际交过哪一版值"推：变更后再无成功交值 ⇒ `old`）
  与 `planNumbers`（人可读旁证）。不换表、不碰提示词。
- `scripts/run_a1_e2_chain.mjs`：新增两份来源变化工件的自洽校验（改的那格必须在指纹
  覆盖范围内、变更晚于首次交值、`planSourceBasis` 可判定、规划期 0 写入、宿主对得上），
  缺一份即非零退出。
- **先写失败回归再修**：Excel `resultExpiryRecoveryChain.test.tsx` 与 WPS
  `WpsAgentPanel.test.tsx` 各一条"计划按旧来源算完、来源在放行前被改" ⇒ 修前**都是
  新增写入 1 次**。
- 修复（两端同一份判据 `packages/core/src/planSourceAnchor.ts`）：计划锚点改成
  **模型读到的那一版**（Excel = `turnSourceBaseline`，WPS = 桥绑定指纹），
  WPS 补计划级写前复核（`executeWpsPlan` 的 `assertPlanSourceCurrent`），
  Excel 烘焙分支改读同一份三态判据。

**结论（分开写）**：

| 项 | 结论 |
| --- | --- |
| A 路：模型没有再次 prepare | **通过**（两轮真实 SDK：`prepareCount=1`、`refusals=[]`、`planSourceBasis=old`、计划数值含 400/1000 不含 550/1150、规划期 0 写入） |
| A 路：旧结果能否绕过写前复核写入 | **修前能（真实缺陷）**；修后两端各自拒绝、**写入 0 次**，两端的失败回归均转绿 |
| B 路：模型确实再 prepare | **本轮未发生**（没有新增真实模型证据）；判据由两端既有面板用例覆盖（`source_changed` + 零写入），如实标注。→ **下一轮（第七轮）已补真实模型证据** |
| 原生实机（真实 Excel / WPS 窗格） | **未做** |

**需要产品决定的一条**：Excel 在 `StaleSourceError` 之后仍有显式「仍要执行旧结果」入口
（既有烘焙计划策略），自动路径不会走它；WPS 没有这个入口。是否在恢复链上撤掉，留给下一轮。

### 2026-10-25（第七轮）：A1-E2 B 路 —— 模型**确实**再读一次被拒

证据全文见 [A1-E2 §10](SDK_A1_E2_RECOVERY_CHAIN_2026-10-25.md#10-b-路模型确实再读一次2026-10-25-后续轮次)。

**这轮补的是什么**：A 路只证明"模型不再读时，旧结果只能靠写前复核拦下"。B 路问的是
另一件事 —— **模型再读一次时，那一次读取走的是不是生产桥的授权/来源复核**。

**只改驱动、测试、文档，生产代码没有新增行为**：

- 新增 `scripts/bridge_authorization_gate.mjs`：Python 驱动通过它**直接 import**
  `packages/core/src/sdkToolBridge.ts` 的 `verifyBridgeAuthorization`，不再在驱动里
  复刻第二份"指纹不等就拒"的判据。入口自身执行失败一律中止实验（不当成"拒绝了一次读取"）。
- 驱动新增场景 `source-change-reread` 与 `second_read_evidence()` / `plan_source_basis()`：
  分支判定只认"变更之后还有没有新的 `prepare_python_data` 调用、回执是什么码"，
  被拒的读取**不算交值**。
- 修掉驱动自身一处口径错位（`afterServedCall` 记的是调用序号，不是"第几次成功交值"）。
- 两端消费侧各补 2 条门控用例；服务端补 7 条判定用例（`test_a1_e2_driver_branch.py`）；
  链路脚本新增两份 B 路工件的自洽校验（缺一份即非零退出）。

**真实实验：3 轮，每端 1 轮有效**

| 轮次 | 端 | 结果 | 覆盖 B 路 |
| --- | --- | --- | --- |
| 1 | WPS | 6 次受控读全被 `source_changed` 拒，最后 `MODEL_TURN_LIMIT`，`kind=null` | 部分（只覆盖拒绝侧，没有计划） |
| 2 | WPS | 第 2 次 `prepare` 被拒；计划 1 个动作；`old` | **是** |
| 2 | Excel | 同上（模型在被拒后又重算了一次再提交） | **是** |

第 1 轮失败原因是需求没写"被拒后不要反复重试"（模型连着重试到轮数上限），
不是守卫太紧；工件与日志保留在 `outputs/a1-e2-source-change-reread-round1-turntlimit.*`。

**结论（分开写）**：

| 项 | 结论 |
| --- | --- |
| B 路：模型确实再读一次 | **通过**（独立 `operationId`/`stepId`、独立账本条目；两次来源范围完全相同 `sameScope=true`） |
| 第二次读取被生产判据拒交值 | **通过**（`gate={ok:false,errorCode:"source_changed"}`，`served=false`，账本记 `warning`） |
| 拒绝回执真的送达模型 | **通过**（两份计划的文案里都写着"第二次读取被拒绝（source_changed）…写入的是第一次读取的结果"） |
| 迟到的旧计划不得写入 | **通过**（两端真实面板各一条用例：写入 **0** 次、目标区逐格为空；规划期 `hostWritesDuringPlanning=0`） |
| 反向对照 / 变异 | **成立**（拿 A 路工件跑 B 路用例两端变红；变异 B1/B2 各 1 红，失败原文带"新增写入 1 次"，逐字节还原） |
| 原生实机（真实 Excel / WPS 窗格） | **未做** |

**逐端状态**

| 端 | 已实现 | 未完成 | 原生待验 |
| --- | --- | --- | --- |
| Excel | B 路全链（真实模型 + 真实服务端 + 真实 Docker + 真实 `<App/>` + 替身宿主）；写前来源锚点复核；2 条门控用例 | 真实窗格未点过 | 改源后拒写提示的实际呈现；是否有人点到「仍要执行旧结果」 |
| WPS | 同上（真实 `WpsAgentPanel` + `bridgeRoundTripHost`）；计划级 `assertPlanSourceCurrent` | 真实窗格未点过 | 同上；提示文案效果 |

**一行必须记住的行为**：模型在没被明确要求时会在被拒后连续重试到 `MODEL_TURN_LIMIT`。
守卫每次都拒（没有松口），但任务会失败。改它属于产品决定，本轮没动。

**本轮检查**：lint 0（36 warning）、`test:addin` 1552/9 skip、`test:wps` 1118/8 skip、
`test:core` 834、`pytest` 1770/1、双端 build 与 `compileall` 退出码 0、
`git diff --check` 0、`npm run test:a1-e2` 0（六份工件，Excel 6 / WPS 187）。
日志 `outputs/a1e2b-*`、`outputs/a1-e2-source-change-reread*.log`。

**A1-GATE 仍不达标，A2 未启动；未提交/推送 Git。**

### 2026-10-25（第八轮）：A1-E3 —— `source_changed` 终态 + 双端缺口矩阵

证据全文 [A1-E3](SDK_A1_E3_STOP_AND_GAPS_2026-10-25.md)；原生清单
[A1-E3 原生验收包](A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md)。

**修掉一个真实缺陷（不是补文案）**。第七轮 B 路第 1 轮日志的形状（6 次受控读全被
`source_changed` 拒 → `MODEL_TURN_LIMIT`）本轮定位到了层：

- `ToolBridgeSession.deliver()` 把拒绝回执交回模型后，**session 仍然可用**，下一次
  `request()` 照样 `on_request` —— 真的又派给宿主读表；
- `run_sdk_plan.handle()` 只在 `status == "ok"` 时处理，`source_changed` 只记一条 warning；
- `_result()` 的流循环不看任何停止信号。

所以"提示词里写不要重试"从来不可能有效。修法是**代码级终止**：一处判据
（`planning_source_guard.py`）+ 两段式停止（先交回回执再停桥）+ 停止态回同一固定码且不再
`on_request` + `drop_python_state` + 流循环 `break` 并 `aclose()` + `handle()` 双闸 +
对外 409 `SOURCE_CHANGED`（`retryable=False`）。

**明确没做**（各有正向锚点用例）：没把 `result_ref_invalid` / `timeout` 一起当终态；没放宽
来源核对；没自动重新规划写入；没加执行前确认按钮。

**新增 20 条自动化用例 / 9 个单点变异**（S1–S4、D1、W1–W2、C1–C2，全部变红并逐字节还原）。
其中两条是"反面"证据：`timeout` 不触发停止、成功重读不触发停止（并断言两次/三次都真到了宿主）。

**并发认领**：`server/tests/test_recovery_claim_concurrency.py` 用 8 线程同一 revision 抢
office/planning 两种认领，成功数必须**恰好 1**；拆掉版本条件（规划侧两层要一次拆全）立刻
出现 8 个执行权 —— 这条证明用例真的在测并发，不是把顺序调用当并发。

**双端缺口矩阵 12 行**（§3 文档）：`✅` 只到"真实服务端 + 真实面板组件 + 替身宿主"；
`⏳` 是没在真实宿主跑过；唯一真实缺口是 Excel **未保存同名簿的读路径身份**（回退
`name:<名>`，写路径因服务端要求 `url:` 前缀而安全）。

**新增 1 个开发入口（默认关闭）**：`POST /api/dev/source-changed-fault/arm`，开关
`EXCEL_BRO_SOURCE_CHANGED_FAULT=1`。存在的唯一理由：`source_changed` 的两个真实产生点都
要求"模型在来源变化后再读一次"，真实宿主无法确定性触发 —— 靠运气不叫验收。关闭时端点 404、
桥行为逐字节不变（`test_dev_fault_is_off_by_default` + 变异 D1）。顺带修掉守卫一处
"GET 也要求 JSON 声明 ⇒ 状态永远 403"的坑（只放松无体方法，写方法仍要求）。

**本轮检查（真实结果）**：

| 检查 | 结果 |
| --- | --- |
| `npm run test:addin` | 1555 passed / 9 skipped（exit 0） |
| `npm run test:wps` | 1119 passed / 8 skipped（exit 0） |
| `npm run test:core` | 834 passed（exit 0） |
| `pytest server/tests` | **1786 passed / 1 skipped**（exit 0） |
| `compileall server/app` | exit 0 |
| `npm run lint` | 0 error / 36 warning（exit 0） |
| `build:addin` / `build:wps` | 均 exit 0（WPS 第一次因沙箱批量删除守卫失败，单独重跑通过；`tsc --noEmit` 两轮都通过） |
| `git diff --check` | exit 0 |
| `npm run test:a1-e2` | exit 0（Excel 6 / WPS 188 通过；六份工件自洽校验通过）—— 桥改过，本轮重跑 |

**A1-GATE 仍不达标**（原生证据这一批一行都没有），**A2 未启动**，**未提交/推送 Git**。

### 2026-10-25（第九轮）：A1-E3 收尾 —— 身份碰撞 / 写入次数口径 / 验收入口作用域

证据全文 [A1-E3 §2.0、§3、§4、§7](SDK_A1_E3_STOP_AND_GAPS_2026-10-25.md)；原生清单
[A1-E3 原生验收包](A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md)。

**本轮只做三件事，都是"已确认的缺口"，不扩到 A2。**

1. **Excel 未保存同名簿的读路径身份撞车（真缺陷，已修）**

   `currentOfficeWorkbookIdentity()` 在 `document.url` 为空时返回 `undefined`，桥的身份于是
   退回 `name:<名字>`；两个都叫「工作簿1」的未保存簿在复核的三项（host + name + identity）上
   全等 ⇒ 授权给 A 的任务可能读到 B 的数据。

   修法是新增 `currentOfficeBridgeWorkbookIdentity()`：已保存仍用 `url:`（与写入账本、恢复
   检查点同一口径），**未保存**按**宿主的工作簿对象**发一个会话内令牌（`unsaved:<n>-<rand>`，
   WeakMap 记在 document 对象上，思路与 WPS 的 `wps-session:` 一致）。规划请求、桥绑定、
   执行前复核、续跑证据四处读同一份。

   边界如实记：它只能证明"同一次加载里是同一个对象"。窗格重开 = 新对象 = 新令牌 ⇒ 服务端
   比对检查点身份判 `PLANNING_CONTEXT_TARGET_CHANGED`，必须重新发起。**没有**为此往
   `document.settings` 写东西，也不要求用户先保存才能读表。写路径从来是安全的（服务端
   `register` 要求 `url:` / `wps-path:` 前缀）。

   红→绿：`apps/excel-addin/src/bridgeWorkbookIdentity.test.ts`（5 条）。拆掉机制后
   `expected { ok: true } to deeply equal { ok: false, errorCode: 'source_changed' }`（3 红），
   逐字还原后 5 passed（变异 M1）。

2. **「实际写入次数」的口径（旧口径不成立）**

   旧验收说明拿 `actions/start` 的**请求行数**当写入次数：重复请求会被拒、认领后可能没碰
   宿主就失败、一次动作还可能分块写多次。现在分四档并只用能观测到的下结论：请求次数 /
   认领结果 / **宿主写调用数** / 结果与回读。

   - 唯一口径 `packages/core/src/hostWriteAudit.ts`（四态 `outcome`，`unknown` 不并进"已写入"）；
   - Excel 执行器在**真的调用宿主写 API 的地方**计数（包住本次 `Excel.run` 的 `context.sync`，
     按动作取差值 ⇒ 覆盖分块写，也覆盖"还没写就失败 = 0 次"）；
   - 端上在账本 `finish`/`suspend` **之后** best-effort 补报；服务端只存计数与状态，
     新端点 `POST/GET /api/office|wps/executions/{id}/host-writes` 读汇总（追加式，
     同动作取最后一条，别人的 `ownerToken` 409，另一端的 ID 404，上报不改执行状态）；
   - **WPS 当时未实现**（三个写调用点未打点，见 A1-E3 §5.2；**第三批已实现，见第十轮**）：
     那一轮 WPS 只能到"请求 + 认领 + 回读"三档。

3. **`source-changed` 验收入口的作用域（真缺陷，已修）**

   原来 `arm` 可留空 Body，命中点是"下一次工具调用"——探针实测另一本簿会把它抢走
   （`stolen = True`），被抢的那一轮真的停在半路。现在：`hostId` + `workbookName` 必填、
   `workbookIdentity` 可选；任务归属由服务端在**开桥那一刻**绑定（`TaskToolBridge.open` →
   `bind_task`），此后只认那一个 `taskId`；状态多一个 `bound` 与 `taskBound`。
   拆掉任务绑定后 3 条用例变红（变异 F1），逐字还原后 18 passed。

本轮：core 834；Excel 1560 / 9 skip；WPS 1120 / 8 skip；pytest 1795 / 1 skip；
lint 0 error / 36 warning；双端构建通过；`git diff --check` 干净。
**A1-GATE 仍不达标**（Excel 的读路径身份、写入计数只到"真实执行器 + 替身宿主"；
真实 Excel / WPS 窗格里一个场景都还没跑）；**A2 未启动**；**未提交/推送 Git**。

### 2026-10-25（第十轮）：A1-E3 第三批 —— 观测覆盖 / WPS 写调用 / 故障目标身份

证据全文 [A1-E3 §8](SDK_A1_E3_STOP_AND_GAPS_2026-10-25.md)；验收清单
[A1-E3 原生验收包](A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md)（§1 已按新口径重写）。

**本轮只做三个问题，都是"已确认的缺口/缺陷"，不扩到 A2，也不提前让用户做原生 S1。**

1. **「没有观测」不能冒充「明确零写入」（真缺口，已修）**

   `host-writes` 是 best-effort 补报，所以 `entries = []` 与"明确 0 次调用"是两句话。
   新增 `HostWritesObservation`（`observation` / `callEvidence` / `attemptedActions` /
   `reportedActions` / `missingActions` / `unexpectedReports` / `unit` / `lastReportedAt`）：
   **只有 `complete + zero` 才算"明确零"**，`not_collected` / `partial` / 含 `unknown`
   一律不许；记录不存在 ⇒ 404，不返回"看起来已验证的零"。

   可信执行边界由账本给（`_host_write_boundary` = `max(next_action_index, in_flight + 1)`
   再按计划动作数截断）：没认领过的动作不可能碰过宿主。展示层唯一口径
   `describeHostWriteObservation()` 还做一次**自洽复核** —— `zero` 却带 `missingActions`
   或 `unexpectedReports` 时拒绝升级。这条复核第一次写就抓到一个真漏洞
   （`partial + zero` 曾被读成"明确零写入"），由变异 H4 钉住。

   两端计量单位不同、不许折算：Excel = `excel_sync_batches`（刷新批次），
   WPS = `wps_api_calls`（写 API 调用）。

2. **WPS 写调用观测（此前 ❌ 未实现，非宿主限制）**

   `wpsHostWriteCounter.ts` 作用域式计数，计数点落在**真实赋值处**：`wpsWriteMatrix`
   （计划写 1 次 + 保文本用的 `NumberFormat` 辅助若干）、`wpsBorders.restoreWpsBorder`
   （一条边一次）、`applyFormatAction` 各 setter（各一次）。作用域只在动作窗口内 ⇒
   撤销、审阅副本、清标记、`PluginStorage` 天然不计。
   `hostWriteCalls`（含辅助）与 `mutatingCalls`（只数承载计划内容）**不许混**。
   面板在 `finish`/`suspend` 之后 best-effort 补报，失败只少一条证据。

   十处场景有用例：单块 1 次 / 120 行分块 2 次 / 辅助 5 calls 对 1 mutating /
   认领被拒 `failed_before_write` 且 0 次 / 执行边界之外**一条条目都没有** /
   写 API 抛错 `failed_after_write` / 撤销不计 / 回调抛错不影响执行 /
   回执丢失只续第三步 / 作用域嵌套。

3. **开发故障入口的同名工作簿串用（真缺陷，先红后修）**

   两个开发入口（`dev_source_changed_fault`、`dev_result_ref_fault`）都允许只报
   宿主 + 名字 ⇒ 两本同名簿落进同一个键，谁先开桥谁把这次单发抢走。探针实证改造前
   `arm(只有名字)` 被接受、另一本同名簿（乙）先开桥就把这次抢走（`stolen = True`），
   目标簿（甲）照常读表、看不到任何停止；被抢的那一轮**真的停在半路**。

   修法：共用 `server/app/dev_fault_target.py` —— 只认 `url:` / `wps-path:` /
   `wps-session:` / `unsaved:`；`name:` 回退值不算身份；显式给了用它（`declared`），
   没给只在"服务端只见过这一个身份"时补全（`observed`，是证据不是猜），零个或多个
   ⇒ **拒绝**并说明原因（顺带给出 `seenTargets` 让人抄）；`bind_task()` 也要求可核验
   身份相等，任务侧只报名字时不绑。`taskId` 仍由开桥路径绑定，不由 `arm` 指定。
   E1 入口在同一身份边界内同步修复（属同一处判据，不是无关重构）。

   回归：`test_source_changed_stop.py` 追加 6 条（与既有合计 37 passed），
   `test_result_ref_fault.py` / `test_python_coordinator.py` 同步改用可核验身份。

**反向变异：8 个单点全部变红并逐字节还原**（H1/H2 空上报与边界外上报、H4 展示层自洽、
H5 计划写计数、H6 辅助写混入、H7 没身份按名字认簿、H8 同名多身份、
H9 任务侧只报名字也绑）。驱动 `scripts/mutate_a1e3_host_writes.mjs` +
`scripts/run_a1e3_host_writes_mutations.mjs`（四判据 + "变异确实改了文件" + sha256 还原）。
其中两次是**取证方式写错**（`-t` 吃了字面 `+` 导致整文件 skip；`errorContains` 猜错文案），
已记录在案，没有放宽成"只要非 0 就算红"。

本轮：core 845；Excel 1560 / 9 skip；WPS 1130 / 8 skip；pytest 1812 / 1 skip；
lint 0 error / 36 warning；双端构建通过；`compileall server/app` 与 `git diff --check` 均 exit 0；
门控 `npm run test:a1-e2` 通过（消费侧 Excel 6 / WPS 188）。
**A1-GATE 仍不达标**：WPS 写调用次数与 Excel 刷新批次数只在"真实执行器 + 替身宿主"里看过，
真实 Excel / WPS 窗格里一个场景都还没跑；**A2 未启动**；`--no-session-persistence` 未动；
**未提交/推送 Git**。

### 2026-10-25（第十一轮）：A1-E3 S1 原生验收的脚本化准备（不扩功能、不启动 A2）

本轮只做「把准备、故障绑定、证据采集变成可执行脚本」，让用户只负责在真实窗格里
发需求、点按钮、截图。**脚本准备完成 ≠ 原生验收完成**，A1-GATE 保持未通过。

改动（全部新增/最小，不碰生产判据）：

1. `scripts/make_a1e3_native_fixture.py`：生成独立测试工作簿（`outputs/a1e3-native/`
   `a1e3-验收.xlsx` 母本 + `-副本.xlsx`），三表 `来源`/`目标`/`汇总`；来源 B2:B4 = 111/250/333，
   目标 B2:B4 = -101/-202/-303。
2. `scripts/a1e3_native.py`：服务端交互 CLI（`status`/`arm`/`cancel`/`evidence`）。
   - 身份**只从 `seenTargets` 抄**，不手写路径格式；多候选/零候选/身份不在 seen 一律停下，
     **不默认取第一条**；
   - `evidence` 只读生产端点 `GET .../host-writes`，把服务端的 `observation` 原样落盘，
     **不重写判据**；账本 SQLite 只读模式只用于**发现** turnId/executionId（定位，非判据）；
   - 失败路径（服务不可达 / 无 seen / 多候选 / 身份错 / arm 未 armed / 过期 / 未触发 /
     无 executionId）都停下并说清，不判"通过"。
3. `scripts/a1e3_native.ps1`：编排入口（`prepare`/`status`/`arm`/`cancel`/`evidence`/`stop`）。
   - `prepare` 起**独立端口（默认 8877）**测试服务端 + 独立配置目录（复制
     `server/model-connections.json` 复用已保存连接、密钥不外传）+ 独立状态目录；
   - **不碰用户 8765**：端口占用只读检查后拒绝启动，`stop` 只杀本次记录的进程树；
   - 健康检查确认测试服务端真的可用后才说"就绪"。
4. `apps/wps-addin/vite.config.ts`：`server.proxy["/api"].target` 改为可用
   `EB_TEST_API_TARGET` 覆盖（默认仍 `http://127.0.0.1:8765`，零生产影响），让真实 WPS
   窗格能连测试服务端；Excel 端用既有 `VITE_API_BASE_URL` 覆盖。
5. `server/tests/test_a1e3_native_cli.py`：CLI 失败路径 9 条（无 seen / 多候选拒绝 /
   身份不在 seen / arm 未 armed / 过期 / 成功 / 服务不可达 / observation 文案 / 无 execution）。

验证：`test_a1e3_native_cli.py` 9 passed；`test_host_write_observation.py` +
`test_source_changed_stop.py` 合计 44 passed（无回归）；`build:wps` exit 0；
`compileall`/`py_compile` 通过；`git diff --check` 干净。脚本端到端已用独立端口 8877
起过真实 `main:app`（`EXCEL_BRO_SOURCE_CHANGED_FAULT=1`），确认状态目录隔离到
`outputs/a1e3-native/server-config/`（用户 `%LOCALAPPDATA%/Excel Bro` 未被污染）。

**如实记录**：本轮验证期间，用户原本在 8765 上的后端（探测时 PID 11288）**退出了**。
本脚本从未对该进程执行任何 kill（所有清理只针对自己记录的 PID 与后台 task）；但其退出
时间落在本轮工作窗口内，原因（用户自行关闭或崩溃）**未证实**，已在下文与最终报告里标注，
不视为本轮脚本造成，也不据此判定验收状态。

### 2026-10-27（第十二轮）：发现入口 —— 浏览器索引降级为缓存（不推进 A1-GATE）

**现场**：三步常量写入在第二步回执前中断，保存并**完整退出 WPS** 后重开同一文件：账本
`needs_reconciliation / nextActionIndex=1 / inFlightActionIndex=1 / ownerActive=false`、原表仍是
111/222/-303、授权仍是两张表，但**恢复入口消失**，面板显示的是另一条历史任务的"已撤销"。
只读取证结论：进度、材料、现场都在，缺的是"重新发现未完成执行"的入口。

**已实现**：浏览器 `pending-wps-execution.v1` 从唯一入口降级为**缓存**；权威发现改为
`GET /api/{wps|office}/executions/pending`（按当前工作簿稳定身份精确查找，WPS 用
`wps-path:`、Excel 用 `url:`，命名空间互不冒充）。响应只带最小元数据；多候选只给选择入口；
"清除提示"只写处置时间戳；发现全程只读。共享判据在
`packages/core/src/pendingExecutionDiscovery.ts`，宿主身份读取在
`apps/wps-addin/src/wpsRecoveryDiscovery.ts`。**Excel 侧同一入口与本轮一起接线**（同一份协议、
同一份共享判据、同一组服务端用例），面板级回归只在 WPS 侧补齐。

**证据**：共享判据 13 条、服务端 16 条（先红 16 failed → 16 passed）、账本迁移 9/10/11 三档、
WPS 面板新增 10 条（A 完整丢失缓存 / B 完整恢复链 / C 身份与竞态）；反向变异 D1–D6 逐条四判据
成立并逐字节还原。**测试隔离同时补上**：两端 `setupFiles` 装网络闸门，未替身的
`fetch`/XHR/WebSocket/beacon 一律令用例失败（本轮据此抓出 8 个真实出口并封住，含
`/host-writes`、`/execution-recovery/{id}` DELETE、`/worksheet-checkpoints/list`、
`/api/turn/{id}/events`、`/__wps_probe`）。

**没做的事**：真实 Excel / WPS 窗格里没有点过"待核对在途动作"；真实宿主完整退出后的自动发现
没有原生复核。**A1-GATE 仍未通过、A2 未启动**；本轮只解决发现能力，不把自动化写成原生验收。

## 6. 后续阶段的交付边界

- **A2**：按恢复设计完善执行账本保留期（既定目标含 30 天后 plan_json 去值）、规划账本清理、历史转录清理、迁移与损坏记录处置。清理不得让旧执行 ID 可以重新执行，不静默删除用户历史转录。
- **S**：先做 session_store / materialize_resume_session 的只读夹具实验，证明转录内容符合隐私要求，明确目录与清理；未完成用户决策前保持原生持久化关闭。
- **B**：SDK 工具复用 controlledScript.ts。只生成白名单计划，沿用校验、执行、回读和撤销；解决 TS 解释器与 Python 服务端接线，不另造一份行为漂移的解释器。
- **C**：先核对 capabilities 与生产可达入口，再分批补 WPS 筛选/去重、工作目录和结构对象能力。每批同时设计授权、审阅、写入、回读、恢复。
- **D**：逐步抽离 SDK 对 excel_agent.py 的共享依赖；确认无生产调用才移除旧循环/副本路径；验证安装包 CLI 启停、依赖和错误反馈，不改用户连接与密钥。

## 7. 证据索引

- [恢复设计及九场景](SDK_RECOVERY_DESIGN_2026-10-19.md)：设计与历史残余，当前实施状态以本文覆盖。
- [A1 执行恢复证据](SDK_A1_EXECUTION_EVIDENCE_2026-10-24.md)：保留早期矩阵，不能直接把其中历史未实现项当最新结论。
- [精确中断验收入口](SDK_PRECISE_INTERRUPT_2026-10-24.md)
- [WPS 受限续写](SDK_WPS_RESUME_2026-10-24.md)
- [恢复完成接回审阅](SDK_RECOVERY_REVIEW_2026-10-24.md)
- [WPS 最终回执](SDK_WPS_FINAL_RECEIPT_2026-10-24.md)
- [材料缺失与过期](SDK_A1_MATERIAL_RECOVERY_2026-10-24.md)
- [结果引用过期](SDK_A1_RESULT_EXPIRY_2026-10-24.md)
- [独立进程退出与竞争](SDK_A1_PROCESS_RECOVERY_2026-10-24.md)
- [A1-E3 来源变化终止与双端缺口矩阵](SDK_A1_E3_STOP_AND_GAPS_2026-10-25.md)
- [A1-E3 原生验收包](A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md)
- [整理前完整历史](SDK_MAINLINE_HISTORY_THROUGH_2026-10-24.md)

### 2026-10-25：Codex 补齐 A1-E2 来源锚点收口

接手发现缺绑定时响应指纹残留、WPS SDK 闸门误伤本地计划、复核来源范围/顺序不一致及
测试宿主缺少桥和 Range 读法。已修复，并新增生产入口与执行器回归。
SDK 固定值计划没有本轮可信桥锚点时拒写；本地入口继续独立来源复核，不能互相冒充。
不改变先执行再审阅，也不增加执行前按钮。

本轮：core 834；Excel 1552 / 7 skip；WPS 1118 / 6 skip；pytest 1763 / 1 skip；
双端构建通过；开门控 A1-E2 Excel 4 / WPS 185 通过。详细失败与修复记录、日志以及
替身边界见 [A1-E2 证据 §9](SDK_A1_E2_RECOVERY_CHAIN_2026-10-25.md#9-2026-10-25-codex-接手来源锚点收口)。

接下来的队列仍按顺序：
1. ~~A1-E2 B 路真实 SDK 再次读取后的拒绝证据~~ → **已补**（见 2026-10-25 第七轮与
   [A1-E2 §10](SDK_A1_E2_RECOVERY_CHAIN_2026-10-25.md#10-b-路模型确实再读一次2026-10-25-后续轮次)）；
   剩下的是**双端原生过期恢复/改源拒写日志**（真实窗格，需用户按步骤操作）。
2. ~~A1-E3 逐项补齐身份、迟到、并发、异常退出与动作调用次数证据~~ → 自动化部分已完成（第八轮：`source_changed` 终态修复 + 12 行双端缺口矩阵 + 并发认领真并发证据 + 9 个单点变异）；**剩下的是原生操作**（`docs/A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md` 的 S1–S3）。
3. 对照 A1-GATE 判断完成或明确分批范围；尚不启动 A2、SDK 原生会话持久化。

本轮没有操作原生窗格或用户工作簿，没有新跑真实模型，不能扩大为 A1 整体通过；
未提交/推送 Git。既有其它改动保留。

收尾检查：lint 0 error / 36 warnings，compileall server/app scripts 与 git diff --check
均 exit 0。日志见 A1-E2 §9；原生验收结论保持未完成。


### 2026-10-26：接手修复 S1 验收工具

原第十一轮存在真缺陷：裸 arm 只列目标却返回成功；取消请求未发 JSON；消费后误要求
仍 taskBound；证据抓最近记录且忽略宿主；缺账本被当作未执行；partial+zero 被说成零调用；
PID 文件可能误停复用 PID；Base/Port 可不一致；夹具、日志被固定目录覆盖；工作目录依赖。
现已修复，专用包装器审计真实绑定/消费，只读关联任务与账本，不改生产 API/协议。

已发现并修正本轮实现问题：SDK 目录一度放入仓库内，模型因此不可用；现放仓库外临时目录，
启动即检查，并补失败回归。第一次导出被当前沙箱拒写，授权后重跑成功，未把失败输出当证据。
服务端测试收集前隔离配置与 SDK 目录；已知 A1 测试默认账本再作逐测试隔离，未清理用户旧账本。

专项 30 passed；服务端第一轮全量 1838 passed/1 skipped；Excel 1560/9 skipped、WPS
1130/8 skipped，双端构建通过。Excel 与 WPS/服务端并跑时曾有3条5秒超时，保留原日志，
其他全量结束后单独跑 Excel 全绿，没有扩大超时或修改断言。最终检查数字在本轮证据文档更新。

Excel S1：真实 Office.js 窗格、真实 SDK、真实本机服务，故障与账本同一 taskId，
failed/SOURCE_CHANGED、execution 空；界面失败收尾、原文停止提示且无继续入口。该项不等于
真实手动改源检测，也不等于 A1-GATE。WPS 同场景已补齐：同任务 SOURCE_CHANGED/无执行登记、界面失败收尾无继续入口。双端测试簿保存后独立读取均确认来源111/250/333、目标-101/-202/-303不变。S2/S3及真正手动改源检测仍未在本轮完成。
A1-GATE 未通过，A2 未启动；未提交/推送。


本轮原生证据索引：[S1双端记录](SDK_A1_NATIVE_TOOLING_REPAIR_2026-10-26.md)，包含两条
精确taskId、不可覆盖JSON、原生停止截图与保存后的测试工作簿。S1“故障注入停止响应链”
已取得真实双端证据；不能扩张为来源变化检测、S2/S3或A1-GATE全部完成。


2026-10-26 收尾：双端S1注入停止已完成，来源/目标初值在真实UI保存后独立读取全数不变。
最终服务端1842/1 skip、专项30、core845、Excel1560/9 skip、WPS1130/8 skip，双端build、
lint（0 error/36 warning）、compileall、A1-E2门控（Excel6/WPS188）、脚本语法/差异检查通过。
8877已停、3000/3889恢复默认前端；8765未操作。下一队列为S2/S3原生证据，A1-GATE仍未通过。


### 2026-10-26：S2 原生恢复与写调用观测收尾

详见 [S2 双端原生记录](SDK_A1_S2_NATIVE_RECOVERY_2026-10-26.md)。Excel 首轮实机发现恢复写入未上报，观测 partial；已给两端恢复路径补 onHostWrite 收集，在账本收尾后 best-effort 上报，并通过先红后绿回归。

修复后 Excel 实际 WebView 刷新再双击恢复，WPS 关闭重开（页面复用）后核对、双击续写：两端前两步未重写，只续第三步，来源保持111/250/333、目标由111/250/-303变111/250/333。两端完整观测各3动作；Excel6刷新批次、WPS3API调用，单位不混算。双视图执行前值正确，确认清标记且不新增计划写调用。原始日志、截图、保存后工作簿与同任务账本已归档。

修复后全量 Excel1560/9skip、WPS1130/8skip、pytest1842/1skip；双端build、lint（0error/36warning）、compileall、A1-E2门控（Excel6/WPS188）及差异检查通过。

当前剩余队列：S3确认/放弃后重开不重执行；WPS真正WebView/进程重启恢复；完整故障矩阵及原生过期/真实改源等证据。另记录Excel刷新后固定选择丢失、对比副本列宽、恢复历史步骤警告三项现场边界，本轮未修，不以账本完成掩盖展示问题。A1-GATE仍未通过，A2未启动。既有改动保留，未提交/推送Git。


### 2026-10-26：S3 数据范围确认入口修复

原生反馈：Excel-S3-放弃.xlsx 从「来源」切到「目标」，在数据范围面板点「确认」后仍提示范围变化。根因是两端页脚确认只关闭面板，未调用授权核验；勾选/模式操作才会提交，未改变勾选时旧授权仍在。

Excel 与 WPS 的工作簿页脚现在通过共享 useImmediateScopeAuthorization 发起一次明确核验，仅最新请求成功后关闭面板。失败、过期回执或卸载后的回执不关闭面板；不自动续跑或写入。工作目录已有勾选即提交的快照保持不重复创建，仍有待提交选择时先完成核验再关闭。X 保持关闭语义。

回归先红后绿：Excel 真实 App 组件 + 替身工作簿覆盖「切表 → 只点页脚确认 → 提示撤下 → 下一条需求可执行 → 再切表仍拦截」；WPS 覆盖同入口及失败保留旧授权/重试；共享 hook 覆盖迟到、失败、卸载。首轮全量发现目录快照多建及旧同步关闭断言，已修正目录分支与等待语义；一条中断验收用例并跑时超时，未扩超时，专项和最终全量均通过，原失败日志保留。

最终自动化：Excel 1563 passed / 9 skipped；WPS 1132 passed / 8 skipped；lint 0 errors / 36 warnings。日志：outputs/scope-confirm-addin-full-final.log、scope-confirm-wps-full-final.log、scope-confirm-lint-final.log；初始失败：scope-confirm-excel-red.log、scope-confirm-wps-red.log；两端最终构建日志：scope-confirm-build-addin-final.log、scope-confirm-build-wps-final.log。

用户复验：重新加载 Excel Bro 窗格，保持「目标」活动表 → 打开数据范围 → 点确认 → 原范围提示应消失 → 重新发送保留的需求。结果写入 A1 后点「放弃变更」，应恢复 A1 为空，B2:B4 保持 -101/-202/-303。本轮未代用户操作工作簿，修复后的原生入口仍待复验；S3 与 A1-GATE 不据自动化标为通过，A2 未启动。未提交或推送 Git。


### 2026-10-27：S3 字面值写入与写前拒绝状态修复

用户真实 WPS 反馈：写入「完成」到目标!A1，在复核阶段被当成缺来源锚点的计算结果拒绝，随后误显示写入结果不确定及恢复记录缺失。两条根因分别修复：可信范围编辑编译器返回字面值来源，双端使用绑定动作的共享内存封印；WPS 来源前置检查 + 共享 not_started 状态阻止写前拒绝误入执行恢复。

Excel / WPS：代码已实现，自动化通过，本轮修复后原生待复验。旧计算结果及缺锚点计算结果仍拒写；不增加执行前确认，不将历史计划自动升级为可信字面值。WPS 自动化验证缺锚点拒写后重开也没有恢复卡。

实跑：core851，Excel1564/9skip，WPS1134/8skip，pytest1844/1skip；双端build、lint（0error/36warning）、A1-E2门控通过。详见 [修复记录](SDK_S3_LITERAL_WRITE_FIX_2026-10-27.md)。A1-GATE仍未通过，A2未启动。未操作用户工作簿、未重启8765、未提交/推送。


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

### 2026-10-27：完整退出宿主后的本机恢复包（材料丢失去根因）

真实现场：第二步写入并回读后、完成回执前中断，完整退出 WPS 重开，账本能核对第二步并显示
「已补记写入，未重复写入」，但随后提示「找不到原执行的本机恢复记录」，第三步续不了。

两个问题分开定责：①材料丢——`prepareWpsResume` 只读宿主 `PluginStorage`，退出后为空且没有第二个来源；
②资格过期——`wpsResumeMaxAgeSeconds=1800` 是当时唯一的窗口期。**没有**把①归因于②，也**没有**只延长 TTL。

改动：恢复材料改存本机后端应用数据目录（`application_data_root()/execution-recovery/<op>/material.dat`，
首行元数据 + 原样正文，原子写 + `fsync` + 只重试替换）；共享协议 `packages/core/src/executionRecoveryPackage.ts`
与 `server/app/models.py` 同步，双端共读 `config/execution_recovery_package_cases.json`。WPS 侧新增
写前落盘硬闸（失败即零宿主写入）、跨进程水合（`ensureWpsRecoveryMaterial`）与终态回收
（`DELETE /api/execution-recovery/{id}`，只在撤销收尾调用）。Excel 既有 before-image 持久化复用未改。

窗口期分离：仅宿主缓存按 1800s；有持久材料按服务端算定的 `expiresAt`（`unfinishedRecoveryPackageRetentionDays=7`），
首次落盘算定，重复保存同正文、查询、打开窗格都不延长；配额不足只清已过期的，不淘汰未结束任务的包。

本轮证据（自动化）：core 859 passed；Excel 1564 passed / 9 skipped；WPS 1139 passed / 8 skipped；
两端 `tsc --noEmit` 退出码 0；恢复包服务端 19 passed（`--basetemp` 在 OS 临时目录）。
新增失败回归先于修复：清空 `PluginStorage` 后 `canResumeWpsExecution` 原为 `false`（`找不到原执行的本机恢复记录`），
修复后为 `true` 且只写第三步。反向变异 `scripts/mutate_wps_recovery_material.mjs` 两个 SEED 各打掉一条机制，
四判据齐全，还原后 `sha256` 与基线逐字节一致。

**原生未做**：真实 WPS「完整退出 → 重启后端 → 重开同一簿 → 核对第二步 → 仅续写第三步」一次都没跑过，
因此 **A1-GATE 仍未通过，A2 未启动**。本机恢复包只覆盖 `writeValues`；格式/公式/结构/删表不承诺跨进程精确续跑。
设计与边界见 [本机恢复包](SDK_WPS_RECOVERY_PACKAGE_2026-10-27.md)。

### 2026-10-27：持久恢复包的有效期判据与"保存不可续期"（第二轮）

第一轮把材料搬到跨进程可读的位置后，留下两个已确认缺陷，本轮都先写失败回归再改生产代码：

1. **有效期内仍被 30 分钟拦住**：`prepareWpsResume` 前段已按 `resumeExpiresAt` 判过，后段综合条件里
   仍有无条件的 `age > wpsResumeMaxAgeSeconds*1000` —— 同一份材料被两套期限同时裁决。收敛为唯一判据
   `wpsResumeWindow(record, now)`：有持久期限只认它、字段非法明确拒绝、无该字段才用旧缓存窗口、
   `now >= deadline` 即过期（与服务端 `deadline <= moment` 同口径）。身份/授权/来源/现场/计划指纹/
   在途/执行权校验一条未删 —— 有效期内不等于免检。
2. **过期后重存又拿满一个保留期**：`execution_recovery.save()` 对"同正文但已过期"继续生成新的
   `savedAt`/`expiresAt` 覆盖文件。改为过期即 `RecoveryPackageExpired`、损坏即 `RecoveryPackageCorrupt`
   （不再被 `except (OSError, RecoveryPackageCorrupt)` 当成"没有旧文件"），并在账本新增
   `office_executions.recovery_package_expires_at`（**首次写赢**，只存时间戳），`PUT` 先用它登记/读取
   首次期限 —— 文件被过期清理后再 PUT 也拿不到新期限，不靠"删了再建"实现幂等。
3. 顺带修掉同一纪律下的第三个缺口：面板**点击前复核**另算一套（不看本机材料），材料过期时
   "需重新规划"的按钮点下去被判成"可以继续"后直接返回 —— 按钮看似失效。现在三处（挂载/刷新/
   点击前复核）共用 `wpsRecoveryVerdictFor`。

自动化证据：WPS 1145 passed / 8 skipped、Excel 1564 passed / 9 skipped、core 860 passed、
服务端 1870 passed / 1 skipped、恢复包专项 26 passed、续跑专项 33 passed；双端 `tsc --noEmit` 0、
`compileall` 0、lint 0 error（36 既有 warning）、`git diff --check` 0。反向变异 M3（插回无条件 30 分钟）
打红「a durable package older than the 30-minute cache window…」，M4（过期就返回同正文）
打红「test_expired_material_cannot_be_re_saved_into_a_new_window」，还原后 `sha256` 与基线逐字节一致。

**仍无原生证据**：真实 WPS「完整退出 → 重启后端 → 重开同一簿 → 核对第二步 → 仅续写第三步」一次都没跑，
面板那几条"完整退出"是替身宿主模拟的。**A1-GATE 仍未通过，A2 未启动**；本机恢复包只覆盖 `writeValues`。
细节见 [本机恢复包](SDK_WPS_RECOVERY_PACKAGE_2026-10-27.md) §10–§15。

### 2026-10-27：恢复包的期限解析与"单条坏材料"（第三轮）

两个独立缺陷，都先写失败回归再改实现：

1. **无时区 `expiresAt` → `TypeError`**：`datetime.fromisoformat("2035-01-01T00:00:00")` 不报错，
   返回 naive 时间，与带 UTC 的 `now` 比较直接抛 `TypeError: can't compare offset-naive and
   offset-aware datetimes`。当时**三处各写了一份**解析（`_deadline()`、`load()`、`_prune_locked()`），
   所以读回、同正文重存、清理三个入口全炸；接口层实测是 503（`TypeError` 走不进状态映射表）。
   收敛为一个解析口 `_instant()`：必须是 `str`、非空、可解析、且 `tzinfo`/`utcoffset()` 都有值；
   无时区**不猜本机时区、不自动补 UTC**，一律按损坏处理（`load`/`save` → `RecoveryPackageCorrupt`，
   账本传入的首次期限无效 → 422），另加 `_aware()` 守内部 `now` 与 `default_expiry(moment)` 的调用约定。
2. **单条坏材料拖垮整批清理**：期限比较在 `for` 里裸奔 ⇒ 一条坏包让同目录其它包全不处理；
   非法文本又被当成"已过期"直接删掉（连排查依据都没了）。改为逐条 `try/except`，计数协议扩为
   `{removed, failed, unconfirmed}`：期限读不出 ⇒ 保留 + 记 `unconfirmed`，既不删也不重算期限。

自动化：恢复包专项 26 → **41 passed**（先红 13 failed / 2 passed）；相关 API/账本 103 passed；
WPS 1145/8、Excel 1564/9、core 860、服务端 1870 → **1885 passed / 1 skipped**；双端
`build` 0/0、`compileall` 0、lint 0 error（36 既有 warning）、`git diff --check` 0。
反向变异 `M5`（去掉时区校验）打红四条新增用例、失败原文即 `TypeError`，还原后
`execution_recovery.py=97e69f8d…` 与基线逐字节一致。

**原生未做**：没有打开真实 WPS/Excel。**A1-GATE 仍未通过，A2 未启动。**
细节见 [本机恢复包](SDK_WPS_RECOVERY_PACKAGE_2026-10-27.md) §16–§20。

## 2026-10-27：三步明确常量被来源闸拒绝

用户原生截图：目标 B2/B3/B4 仍为 -101/-202/-303，写入111/222/333的三步计划被
“缺少计算来源”拒绝。可信单区域编译器和内存封印原先都只支持单动作。
新增 plan_literal_edits 并保留三个独立动作；双端共用完整动作序列封印，计算结果
来源闸继续执行。自动化又抓到无来源常量未生成持久材料，增加恢复来源证据 v2：
只从活体可信常量计划产生，sources为空；v1继续要求非空来源。跨缓存丢失恢复后
经原账本和目标核对，只续第三步；撤销还原原始三步快照。

两端：代码已实现、自动化验证；真实 Excel/WPS 本轮未运行。WPS 完整退出恢复
仍待用户按新版入口复验，未把它标成完成。A1-GATE 未通过，A2 未启动。
实现、实跑日志与用户下一步见 [有序常量修复](SDK_ORDERED_LITERAL_WRITES_2026-10-27.md)。

本轮实跑：Excel 1566/9、WPS 1148/8、core 866、服务端1901/1；双端 build、
compileall、diff-check 退出码0；lint 0 error/36 warning；A1-E2 门控Excel6/WPS194。
最后新增v2 HTTP参数后相关三文件161 passed。真实宿主未运行，未提交/推送Git。

## 2026-10-27：旧账本缺列把恢复包 PUT 打成 500

真机诊断链：execution register 200 → 恢复包 `OPTIONS` 200 → **恢复包 `PUT` 500**
（`no such column: recovery_package_expires_at`）→ suspend 200 → `GET` 404。浏览器只报
`Failed to fetch`。

根因是**判据读意图不读事实**：`_schema_is_current` 只比 `PRAGMA user_version`，版本号被写到 10
而列没建时永久跳过补列。修法与纪律见
[本机恢复包](SDK_WPS_RECOVERY_PACKAGE_2026-10-27.md) §21–§22 与 `ARCHITECTURE.md`
「版本号是意图，列是事实」。

客户端同轮补的是**"失败之后敢不敢说没写过"**：`noActionStarted` 四条判据（正面证据 / 账本游标 /
在途动作 / 执行权）缺一条都保留本机待办；`ReviewNotStartedError` 是唯一允许把审阅置为
`not_started` 的口。第一版多加的 `status !== "running"` 被 M6 证明抓不到任何变异（服务端
`ownerActive` 与它同源），已撤回 —— **冗余判据不是防线**。

自动化：服务端专项 72 → **74 passed**；core 审阅 28；WPS 面板 194/5（199）；全量 core **867**、
WPS **1153/8**、Excel **1566/9**、服务端 1885 → **1907 passed / 1 skipped**；双端 `tsc` 0/0、
`build` 0/0、`compileall` 0、lint 0 error（36 既有 warning）、`git diff --check` 0。
反向变异 M6–M10 各自**只**打红一条用例、还原后 `WpsAgentPanel.tsx=07682414…`、
`wpsExecutionRecovery.ts=61b7273b…`、`task_ledger.py=124b4ec5…` 与基线逐字节一致。

升级只需重启服务进程（自动补列，保留任务/进度/检查点/首次期限）；不重建账本、不删 WAL。
**原生未做**：没有重启运行实例、没有打开真实 Excel/WPS。**A1-GATE 仍未通过，A2 未启动。**
# 2026-10-27：开发中断旧绑定收尾

WPS 用户原生截图显示三个动作已执行但探针仍等待第二步回读；后端只读记录确认本次完成，
不能计入“第二步中断→完整退出→只续第三步”验收。已确认并修复双端执行结束时未释放本次
未触发绑定、控件到期不刷新两项遗漏；新增真实 WPS 面板中断及写前失败收尾回归。
此处原生根因仍缺窗格绑定 operationId，不能把自动化通过当作现场根因完全闭环。
A1-GATE 保持未通过、A2 未启动。细节见 SDK_PRECISE_INTERRUPT_2026-10-24.md 新增章节。

# 2026-10-27：完整退出后的发现入口（索引降级为缓存）

WPS 三步写入在第二步回执前中断，保存并**完整退出**后重开：账本 `needs_reconciliation/1/1`、
原表 111/222/-303、授权两张表都在，**恢复入口没了**。根因形状 = 唯一入口是浏览器
`localStorage` 的单条索引（`pending-wps-execution.v1`），完整退出即消失；账本记录、恢复材料、
原表现场都不参与发现。既有"完整退出"用例在清宿主缓存后又手工写回索引，恰好绕过了缺口。

修法：索引降级为**缓存**，权威发现改为 `GET /api/{wps|office}/executions/pending`（按当前宿主 +
当前工作簿稳定身份精确查找）。响应只带最小元数据；多候选只给选择入口；"清除提示"只写处置
时间戳（`discovery_dismissed_at`，账本 `user_version` 10 → 11）；发现全程只读。共享判据
`packages/core/src/pendingExecutionDiscovery.ts`，WPS 宿主身份读取在
`apps/wps-addin/src/wpsRecoveryDiscovery.ts`，Excel 侧同轮接线。

自动化：共享判据 13 条（先红后绿）、服务端 16 条（先红 16 failed → 16 passed）、账本迁移
9/10/11 三档、WPS 面板新增 10 条（完整丢失缓存 / 完整恢复链 / 身份与竞态）；全量 core **881**、
WPS **1168/8**、Excel **1566/9**、服务端 **1926 passed / 1 skipped**；双端 `tsc`/`build` 0/0、
`compileall` 0、lint 0 error（36 既有 warning）、`git diff --check` 0。反向变异
`scripts/mutate_wps_recovery_discovery.mjs` D1–D6 逐条四判据成立，还原后
`WpsAgentPanel.tsx=aa0b20c8…` 与基线逐字节一致。

测试隔离同时补上：两端 `setupFiles` 装 `packages/core/src/testNetworkGate.ts`，未替身的
`fetch`/XHR/WebSocket/beacon 直接令用例失败（无允许列表），据此抓出并封住 8 个真实出口。

**原生未做**：真实 Excel/WPS 窗格里没有点过"核对在途动作"，完整退出后的自动发现没有原生复核。
**A1-GATE 仍未通过，A2 未启动。** 细节见 [本机恢复包](SDK_WPS_RECOVERY_PACKAGE_2026-10-27.md) §28。
