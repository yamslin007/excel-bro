# WPS 恢复核对：DEV 诊断能力（2026-10-29）

> **这份交付只加"看得见原因"，没有修任何根因。**
> 2026-10-28 夜间真机出现过一次：完整退出 WPS 后重开，面板先显示「证据不足」，
> 点一次「重新核对（只读）」就恢复正常。当时界面只有这四个字，`unknown` 的十几条产生路径
> 一条都分不出来。本轮把每一条都给了**原因码**，并且只让它在 DEV 出现。
> **在拿到真机的真实原因码之前，不得宣称根因已修复，也不得改判定逻辑。**

## 1. 这一层是什么、不是什么

| | |
| --- | --- |
| 是 | 对**已有**现场核对过程的一次旁注：结论 + 停在哪一步 + 为什么 + 是第几个动作 + 表数量 + ID 尾部 |
| 是 | 只在 DEV 记录；非 DEV 一行都不留、不打日志、不发请求 |
| **不是** | 判据的一部分。`wpsReconcileVerdict` 的输入输出一个字节都没变 |
| **不是** | 自动重试、自动放行、把 `unknown` 降级成"可以继续" |
| **不是** | 修复。原因码只回答"当时卡在哪"，不回答"为什么会卡" |

边界原样保留：`unknown` 仍然判 `insufficient`（「证据不足」），仍然只给
「清除这条恢复提示」+「重新核对（只读）」，**不出现**补记、续跑、重新规划之外的任何写入入口；
诊断本身不产生 claim、不推进 `revision`、不写单元格。

## 2. 原因码全表（每个码对应唯一一处返回点）

`apps/wps-addin/src/wpsRecoveryInspection.ts` 的 `inspectWpsRecoveryStateDetail` 里，
每个提前返回都带一个码。加码时**不要复用"看起来差不多"的码** —— 复用等于把现场又糊回去。

| 阶段 `phase` | 原因码 | 触发点 | 它说明什么 | 它**不**说明什么 |
| --- | --- | --- | --- | --- |
| `authorize` | `authorize_no_selection` | 本次核验拿到的已选表是空集 | 授权范围当时是空的（重开后还没恢复/还没选中） | 是"慢"还是"丢"，要看重试那条 |
| `authorize` | `authorize_threw` | `authorizeWpsSheets` 抛错 | 授权解析本身失败（序号失效/系统表/无工作簿） | 具体是哪一种（`errorName` 只到类名） |
| `identity` | `workbook_identity_unreadable` | 当前 `FullName` 拼不出 `wps-path:` | 宿主还没给出可用的完整路径（未保存/后缀不合规） | 是不是同一本簿 |
| `identity` | `workbook_identity_mismatch` | 身份与账本记录不同 | 确实不是同一本簿 | 谁改了名/路径 |
| `descriptor` | `descriptor_empty` | 描述里没有动作 | 服务端给的描述不完整 | 服务端为什么这么给 |
| `descriptor` | `descriptor_action_count` | 动作数与 `totalActions` 不符 | 描述与游标不同源 | 哪一侧错 |
| `sheets` | `authorized_sheet_not_allowed` | 登记过的授权表不在当前授权集合里 | 当前授权比登记时窄 | 是用户收窄还是没恢复 |
| `actions` | `action_sheet_not_allowed` | 动作目标表不在当前授权集合 | 同上，具体到某个动作 | 同上 |
| `actions` | `action_sheet_not_authorized` | 动作目标表不在**登记**授权表名里 | 描述自身越界 | 同上 |
| `actions` | `action_index_mismatch` | 动作序号与数组下标不符 | 描述错位 | 谁写的描述 |
| `actions` | `action_range_unparsable` | 目标地址解析不出矩形 | 地址字符串异常 | 表实际内容 |
| `actions` | `action_range_shape_mismatch` | 登记的 `rowCount/columnCount` 与地址不符 | 描述内部不一致 | 同上 |
| `actions` | `action_cell_count_mismatch` | `cellCount ≠ row×col` | 同上 | 同上 |
| `actions` | `action_fingerprint_missing` | 旧描述没有 `expectedFingerprint` | 历史记录不可核对 | 现场到底写没写 |
| `actions` | `cell_budget_exceeded` | 目标总格数超 `maxCells`（1000） | 规模越界，拒绝核对 | 现场是否已写 |
| `read` | `workbook_unreadable` | 授权成功后拿不到工作簿对象 | 工作簿在核对中途消失/不可访问 | 为什么消失 |
| `read` | `host_range_unreadable` | `sheet.Range(address)` 返回空 | 宿主没交出区域对象 | 区域是否真的为空 |
| `read` | `host_range_shape_mismatch` | 宿主交出的行列数与登记不符 | 宿主的区域对象与登记不是一回事 | 哪边对 |
| `format` | `format_fields_undeclared` | 格式动作没声明属性 | 描述不足以核对格式 | 现场格式 |
| `format` | `format_fields_unrecognized` | 声明里有不认识的属性名 | 版本/描述不匹配 | 现场格式 |
| `format` | `format_snapshot_incomplete` | 有格读不到，快照不完整 | 宿主没把所有格交出来 | 缺的是哪一格（不记单元格值） |
| `values` | `value_read_undefined` | `Value2` / `Formula` 读回 `undefined` | 宿主读取接口没就绪/没暴露 | 值是多少（**刻意不记**） |
| `values` | `value_shape_unexpected` | 读回形状不是登记的矩阵 | 返回结构不对 | 同上 |
| — | `in_flight_unexpected` | 在途动作既非写后也非写前（结论 `conflict`） | 疑似人工修改或部分写入 | 谁改的 |
| — | `completed_action_mismatch` | 非在途动作与它该在的一侧不符（结论 `conflict`） | 同上 | 同上 |
| — | `inspection_threw` | 核验函数自身抛错（外层 `catch`） | 在上面那个 `phase` 断的 | 异常正文（只留白名单 `name`） |
| — | `panel_threw` | 面板调用层抛错（`inspectWpsRecoveryStateDetail` 之外） | 同样带 `phase=authorize` 与 `name` | 同上 |
| — | `state_missing` | 面板根本没拿到服务端状态 | 谈不上核对现场 | 服务端为什么没给 |

## 3. 记录什么、绝不记录什么

一条记录的字段是**封闭集合**（测试 `wpsRecoveryDiagnostics.test.ts` 里逐字段断言）：

| 记 | 不记（结构性禁止，不靠自觉） |
| --- | --- |
| `seq` / `at` / `stage` / `verdict` / `outcome` / `code` / `phase` | 单元格值、公式、选区内容 |
| `operationIdTail` / `taskIdTail`（各 6 位） | 完整执行身份、`ownerToken`、任务正文 |
| `revision` / `nextActionIndex` / `totalActions` / `inFlightActionIndex` | 工作表名、工作簿名、完整磁盘路径 |
| `selectedSheetCount` / `authorizedSheetCount`（**数量**） | 恢复包正文、计划正文、模型密钥 |
| `actionIndex` | 异常 `message`（WPS 的异常消息常整段带出路径） |
| `errorName`（白名单正则 `^[A-Za-z_$][A-Za-z0-9_$]{0,40}$`，不合格一律丢弃） | 端上自报的写入次数 |

`stage` 有五个取值，用来把"哪一次核对"分开：
`discovery`（首次发现）/ `select`（点选候选）/ `recheck`（**点「重新核对（只读）」**）/
`reverify`（已选卡在工作簿/授权变化后重算）/ `pre-action`（点击恢复动作前的复核）。

## 4. 怎么拿到诊断（两种，都要 DEV 构建）

1. **面板里看**：恢复卡底部多了一个默认收起的「开发诊断（仅 DEV）」，展开就是这段文本，
   可直接选中复制 / 截图。它**不是按钮**，不进动作区，也不改变按钮集合。
2. **dev server 控制台看**：用 `VITE_ENABLE_RECOVERY_DIAG=true` 启动 dev server
   （例如 `VITE_ENABLE_RECOVERY_DIAG=true npm run dev:wps`），每条记录会 POST 到
   `/__wps_probe`，dev server 打印 `[wps-bridge] {"stage":"recovery",…}`。
   默认**关**着：测试与日常开发不会因此往外发请求。

报告格式（`wpsRecoveryDiagnosticReport()`，时间为 UTC）：

```
DEV 恢复诊断（共 2 条，显示最近 2 条；时间为 UTC；只含原因码，不含表名/单元格值/路径）
#1 23:32:07.600 discovery verdict=insufficient outcome=unknown code=authorize_no_selection phase=authorize op=…926ef1 task=…b40493 rev=3 cursor=1/2 selected=0 authorized=0
#2 23:39:06.877 recheck verdict=reconcile outcome=not_started code=… op=…926ef1 task=…b40493 rev=3 cursor=1/2 selected=2 authorized=2
```

两条记录的 `op=` / `task=` 相同才是一次对照；`seq` 递增说明先后。

## 5. 下一次真机需要怎么走（最短步骤）

1. 用 DEV 构建（`npm run dev:wps`，需要控制台取证就加 `VITE_ENABLE_RECOVERY_DIAG=true`）。
2. 复现 2026-10-28 夜间的现场：三步写入 → 第二步中断 → **保存** → **完整退出 WPS** → 重开同一文件。
3. **重开后立刻**（不等面板刷新、不点任何按钮）看恢复卡：
   展开「开发诊断」，把**首条 `discovery` 记录**整段抄下来/截图。
4. **再点一次「重新核对（只读）」**，展开诊断，抄下**首条 `recheck` 记录**。
5. 两条一起给出即可。若第一次没出现「证据不足」，如实记"未复现"，不要为了复现去反复重开。

判读（**条件式**，不是结论）：

| 首次 `discovery` | 只读重试 `recheck` | 方向 |
| --- | --- | --- |
| `authorize_no_selection` / `authorize_threw`（`selected=0`） | 正常 | 重开后授权范围/工作表索引尚未就绪 |
| `workbook_identity_unreadable` | 正常 | 宿主 `FullName` 尚未就绪 |
| `value_read_undefined` / `host_range_unreadable` | 正常 | 宿主读取接口尚未就绪 |
| `format_snapshot_incomplete` | 正常 | 格式快照采集不完整 |
| `authorized_sheet_not_allowed` / `action_sheet_not_allowed` | 正常 | 当前授权集合与登记不一致 |
| `descriptor_*` / `action_*`（形状、序号、指纹） | **同一个码** | 与加载时机无关，问题在描述/服务端数据本身 |
| 两次同一个码 | 同 | 不是"冷启动一次就好"，方向在这些码本身 |
| `inspection_threw` + `error=…` | 正常 | 宿主抛错，按 `error` 名字继续查 |

### 5.1 首条真机记录（2026-10-29 03:42:31.989Z）

按 §5 走了一遍，用户给出的是**首条**，也就是首次发现那一条
（截图 `outputs/a1e3-native/recovery-diagnostics-20261029/screenshots/2026-10-29-1142-first-discovery.png`，
sha256 `68f0f710…bec0cab`；同名文件为剪贴板落盘副本，mtime 只代表复制时刻）：

```
#1 03:42:31.989 discovery verdict=reconcile outcome=completed phase=values
   op=…899412 task=…92ea13 rev=5 cursor=1/3 inFlight=1 selected=2 authorized=2
```

对应的执行（SQLite `mode=ro` + `backup` 一致快照核对，只读）：

| 项 | 值 |
| --- | --- |
| operationId | `wps-execution-1a727455-85bd-47dc-8531-7f3c9b899412` |
| taskId | `task-884bc2e6-62c6-40b7-a7c7-3bc50192ea13` |
| 工作簿 | `wps-path:e:\codex\excel-bro\outputs\wps-abandon-20261028\wps放弃验收-20261028.xlsx` |
| 授权 | 2 张（目标表、来源表） |
| 计划 | 3 个 `writeValues`：B2 / B3 / B4（每动作 1 格） |

时间线（UTC；账本 `created_at` / `office_host_writes` / `updated_at`）：

- `03:41:26.447` 建执行；`03:41:26.918` 动作 0 `written/verified`、动作 1 `failed_after_write`（同一批落盘）
- `03:42:31.989` **本截图这条发现记录**（账本当时 rev=5、`next=1`、`inFlight=1`）
- `04:18:31.228` 动作 2 `written/verified`；账本终态 `completed`、rev=11、`inFlight=null`

**这一条证明了什么**

1. 首次发现这条路**读到了现场**：`phase=values`（一路走到读值阶段）、**没有 `code`**（没有提前返回）、
   `outcome=completed`。也就是说这一次载入**没有**走到「证据不足」。
2. `selected=2 authorized=2`：授权范围读到两张表，`authorizeWpsSheets` 未抛错。
3. `verdict=reconcile` + `outcome=completed` 是**有在途动作时的正常分支**：`wpsExecutionReconcile.ts`
   在 `inFlight != null` 且 `localInspection === "completed"` 时给 `applied`，卡片因此显示
   「第 2 个动作的现场与写后一致，可补记为已写入，不重复写」。

**这一条不能证明什么**

- **不是「证据不足」的样本**：`verdict≠insufficient`、`outcome≠unknown`、`code` 为空。
- 报告写 `共 1 条`：这次载入里**没有**第二次「重新核对（只读）」记录，§5 要求的**两条对照没拿到**。
- 不能反推"上一次显示证据不足是因为 XX"——那一条的码没有留存。
- **缓冲区是内存态**：页面重建（重开 / 刷新 / 换窗格）即清零。所以「首次重开那一条」只能在
  **当次载入**里抄；一旦重载就拿不回来了。这是这套诊断的**固有限制**，不是记录逻辑的缺陷。

## 6. 适用范围（不宣称两端都已修复）

- 这套判定与文案**只存在于 WPS 端**：`wpsReconcileVerdict` / `inspectWpsRecoveryState` /
  「证据不足」四个字都是 WPS 专有；Excel 端没有 `localInspection` 这条判定路径，
  **本次诊断不覆盖 Excel**，也没有对 Excel 做任何改动。
- 两端共用的 `packages/core` 没有改动。
- 跨进程放弃的既有验收结论（`docs/ABANDON_AFTER_RESUME_ACCEPTANCE_2026-10-28.md` §16）不变。

## 7. 测试与反向变异（本轮证据）

`npm run test:wps`：**1217 passed / 8 skipped**（新增 34 条）。关键用例：

- `wpsRecoveryDiagnostics.test.ts`：DEV 门控（非 DEV 一条不留、不打日志、不发请求）、
  缓冲区上限与序号连续、字段封闭集合、异常 `name` 白名单、报告可对照、遥测默认关闭。
- `wpsRecoveryInspection.test.ts`：**每个原因码一条**，且正向锚点（正常结论不带码）同时断言。
- `WpsAgentPanel.test.tsx`「首次发现与只读重试各留一条诊断」：一次 insufficient → 点只读重试 → 恢复，
  两次记录 `stage` 不同、`code` 可辨、ID 尾部相同、**宿主零写入**、诊断里没有表名/路径/单元格值。

反向变异（`scripts/mutate_wps_recovery_diagnostics.mjs`，一个个 SEED，还原后逐字节 sha256 一致）：

| 变异 | 判据 | 结果 |
| --- | --- | --- |
| M1 把"已选表为空"并回 `authorize_threw` | 「已选表为空 → authorize_no_selection」 | 红（`expected 'authorize_threw' to be 'authorize_no_selection'`） |
| M2 去掉 DEV 门控 | 「非 DEV 一条都不留」 | 红（缓冲区里出现了 1 条） |
| M3 去掉异常 `name` 白名单 | 「异常只留白名单 name」 | 红（`expected 'C:\机密目录\客户名单.xlsx' to be undefined`） |
| M4 把只读重试记成 `discovery` | 「各留一条诊断」 | 红（找不到 `recheck` 记录） |

日志与还原 sha256：`outputs/a1e3-native/recovery-diagnostics-20261029/`。

## 8. 未完成（照写，不含糊）

- **根因仍未确定**：真机只拿到 **1 条成功路径**记录（§5.1），**原因码一个都还没在真机上触发过**；
  「首次重开证据不足、只读重试后恢复」**本轮未复现**，也没有可对照的第二条。
- 事件字段 `outcome` 记的是**现场结论**（`completed|not_started|conflict|unknown`），与
  `wpsReconcileVerdict` 自己带的 `outcome`（`applied|not_applied`）**同名不同义**；后者目前只出现在
  卡片文案里，诊断行里没有。改名（`outcome`→`site`）并补记 `applied/not_applied` 未做 ——
  要同时动 `wpsRecoveryDiagnostics.test.ts` 的字段集合断言与四组变异脚本，留作下一轮。
- 没有改任何判定逻辑、没有加自动重试、没有放宽 `unknown`；`wpsExecutionReconcile.ts` 未改。
- Excel 端无对应路径，验证也不覆盖 Excel。
- A1-GATE 仍未通过，A2 未启动。
