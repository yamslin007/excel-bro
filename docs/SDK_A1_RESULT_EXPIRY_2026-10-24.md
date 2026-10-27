# A1 结果引用过期验证（2026-10-24）

本批为结果引用 resultRef 生命周期证据，与 WPS 本机恢复材料过期分开。只添加测试和文档，无产品源码修改；保持先执行、再审阅，不新增执行前确认。

## 已取得证据

- 服务端真实 FastAPI ASGI + 结果登记表：Office/WPS 各一条。相同授权下先读取成功，单独令结果超过配置 TTL，再连续读取两次均为 404/BRIDGE_RESULT_UNAVAILABLE，错误体不泄露结果内容；resolve_result_ref 拒绝旧引用。登记新结果后取得新引用，可生成带新值的白名单计划，旧引用仍为 404 且不能解析。
- Excel/WPS 真实面板：首次取得结果并渲染，随后端点改为 404，连续卸载重挂载两次均显示不可用与重新查询提示，无空表冒充成功、无缓存旧值、无自动模型调用。Excel 无执行登记请求；WPS 无登记且宿主内容写入为零。现有 410 测试保留，新增测试补上生产实际使用的 404。
- 反向变异：独立 Python 进程将 LocalResultRegistry._prune_locked 改为 no-op，两端新增用例均在“期望 404，实际 200”处失败。未改生产源码。正常服务端相关套件随后 63 passed。

## 不可扩大解释

前端网络为替身；服务端新结果由测试登记，不是模型主动重读重算。两层证据没有串成同一个真实 SDK 恢复任务。新引用生成计划也不等于宿主实际执行该计划。完整的恢复期间引用过期→SDK重读重算→新计划执行→审阅仍未闭环；A1 不标整体完成，A2 未开启。

过期只改变测试结果的 created_at，保留有效授权，不修改生产 TTL。首跑漏传 bind 的 workbook_identity，补全为既有测试使用的 None 后通过；该失败为夹具错误，不计产品缺陷。没有 computer-use、没有操作用户工作簿。

## 复现与日志

服务端：pytest server/tests/test_agent_tool_bridge.py server/tests/test_normal_backend_result_tools.py -q。
前端：Excel sdkIntentRouting.test.tsx；WPS WpsAgentPanel.test.tsx，新增测试名均为“结果先可读后过期404：连续重开不复用缓存或自动重算”。
日志 outputs/a1-result-*；反向变异日志 a1-result-mutation.log。

---

# A1-E1 落地：结果引用过期的**精确故障入口**（2026-10-25）

上一节用的是"把测试结果的 `created_at` 改到 TTL 之外"来制造过期 —— 那只能证明
**读取漏斗**会拒，无法在真实 SDK 任务里**按需、单次、精确定点**地制造一次过期。
本节把这件事做成产品里的一个**开发验收入口**（`EXCEL_BRO_RESULT_REF_FAULT=1`），
供 A1-E2 在真实 SDK 链路上串联验证。

## 落在哪一层，为什么

结果正文只在本机易失登记器 `local_results` 里，模型与界面都只有不透明句柄
`res_<32hex>`。要让"旧引用被拒"成为**真实读取路径**上的事实（而不是 mock 或一句
UI 文案），故障只能发生在引用被解析的那一刻：

```
用户显式启用 (arm)                       ← POST /api/dev/result-ref-fault/arm
  → 该工作簿的下一次规划任务开桥 (bind_task)   ← task_bridge.open() 的可信代码
  → 该任务第一条 Python 结果登记 (bind_python_result) ← python_coordinator 的可信代码
  → 该句柄第一次被 resolve_result_ref 使用 (take_expiry) → 从登记器删除 + 固定码拒绝
```

**"已登记、尚未首次写入"就是上面第 3 步到第 4 步之间**：结果已经登记（`register`
成功、句柄可查），而任何宿主写入都还没发生 —— 因为计划（`compile_result_plan`）
是写入的唯一入口，它拿不到 item 就编译不出来。

刻意不做什么：

- 不改 `agentToolBridge.resultTtlSeconds`，不缩短生产结果寿命；
- 不清空全部结果、不整任务丢弃 —— 只动被绑定的那一个句柄；
- 不借用"第二步写入后中断"入口（`recoveryProbe` 的状态完全独立）；
- 请求体里**没有** `taskId` / `resultRef`（`extra="forbid"`）：模型无从指定别的任务
  或别的引用，绑定只由服务端在两条可信路径上完成；
- 失效后**删掉句柄**，所以没有"缓存正文"或重试能把一个已过期的引用救回来。

## 双端接入

| 位置 | 文件 | 说明 |
| --- | --- | --- |
| 服务端状态机 | `server/app/dev_result_ref_fault.py` | 开关、arm/cancel/status、三处绑定；关闭时所有写操作是空操作 |
| 端点 | `server/app/main.py` | `POST arm` / `POST cancel` / `GET status`；未开启一律 404 `DEV_FAULT_UNAVAILABLE` |
| 读取漏斗 | `server/app/result_submission.py::resolve_result_ref` | 命中时 `discard` + 固定 `result_ref_invalid`（两个后端共用这一条） |
| 任务绑定 | `server/app/task_bridge.py::open` | 真实开桥 = 这次规划任务开始 |
| 句柄绑定 | `server/app/python_coordinator.py` | `register` 成功后绑定"第一条" |
| 共享判据 | `packages/core/src/resultRefFault.ts` | 可见性、请求字段、六种 phase 文案、按钮可按性（**唯一一份**） |
| 共享控件 | `packages/core/src/ResultRefFaultControl.ts` | 展开才读状态；收起时不播报（`role="status"` 是面板"还有没有进度提示"的判据） |
| Excel 适配 | `apps/excel-addin/src/App.tsx` + `api.ts` | 工作簿名 = `workbook.name`，身份 = `currentOfficeWorkbookIdentity()`（与桥同一串） |
| WPS 适配 | `apps/wps-addin/src/wpsResultRefFaultTarget.ts` + `WpsAgentPanel.tsx` | 名字 = `listWpsWorksheets().name`（`Workbook.Name`，**不是**路径 basename），身份 = `wpsWorkbookBridgeIdentity()`（与桥同一串） |

两端都不加执行前确认：仍然是"授权任务校验通过后执行一次，再审阅"。

## 自动化证据（本轮实跑）

| 套件 | 结果 | 覆盖 |
| --- | --- | --- |
| `server/tests/test_result_ref_fault.py` | 12 passed | 普通路径不受影响、单次触发、第二条不受影响、查询结果不是目标、同簿不同任务隔离、跨簿不串用、默认关闭、端点 404、请求不接受 taskId/resultRef、TTL 过期、取消、零宿主写入 |
| `server/tests/test_python_coordinator.py::test_armed_fault_expires_the_real_python_registration` | 1 passed | **真实登记路径**：真实 `task_tool_bridge.open` + 真实 `execute_python_analysis` → `local_results.register`，证明绑定在生产代码里而不是测试自己调的 |
| `packages/core/src/resultRefFault.test.ts` | 14 passed | 可见性、目标（空名字不给目标，不退化成全局开关）、六种 phase 文案互不冒充、按钮可按性 |
| `apps/excel-addin/src/resultRefFaultControl.test.tsx` | 13 passed | 非开发模式整块不渲染、挂载不打请求、请求只有三项、绑定后锁死、失败不当状态播报、收起不播报 |
| `apps/wps-addin/src/wpsResultRefFaultTarget.test.ts` | 6 passed | 目标与 `createWpsBridgeLiveContext` **逐字相同**（名字/身份/未保存令牌/失败返回 null） |
| `apps/wps-addin/src/WpsAgentPanel.test.tsx`（新 2 条） | 179 passed | 面板确实接入；读不到现场时入口可见但不可启用 |

判别力：10 个单点变异全部在**预期用例**上变红并逐字节还原
（`outputs/a1-e1-mutations.log`、驱动 `outputs/mut-backup-a1e1/mutate.py`）。

| 变异 | 改了什么 | 变红的用例 |
| --- | --- | --- |
| S1 | `enabled()` 恒 True | 默认关闭 / 端点 404 |
| S2 | `bind_task` 不认工作簿身份 | 跨簿不串用 |
| S3 | `bind_python_result` 覆盖成"每条都失效" | 单次触发 |
| S4 | 失效后不 `discard` 句柄 | 单次触发、零宿主写入、真实登记路径 |
| S5 | 删掉 `python_coordinator` 的绑定行 | 真实登记路径（**只有它红**） |
| S6 | 删掉 `task_bridge.open` 的绑定行 | 6 条依赖"开桥即绑定"的用例 |
| C1 | WPS 目标用完整路径当工作簿名 | WPS 目标 3 条 |
| C2 | 空工作簿名也给目标 | 共享判据"不退化成全局开关" |
| C3 | 控件不理会 `visible` | 非开发模式整块不渲染 |
| C4 | 收起时也常驻状态播报 | 收起时不播报 |

本轮全量：Excel **1551 passed / 3 skipped**、WPS **1105 passed / 3 skipped**、
服务端 **1763 passed / 1 skipped**；两端 build（含 `tsc --noEmit`）退出码 0；
`compileall` 0；lint 0 error / 37 warning（新文件 0 告警）；`git diff --check` 0。

## A1-E1 ≠ A1-E2（必须分开记录）

本节只证明"**能让某一次任务的第一个 Python 结果引用在登记后、首次写入前精确失效，
并且旧引用被真实读取路径拒绝、句柄已删、零宿主写入**"。

**没有**证明：真实 SDK 在收到 `result_ref_invalid` 之后会不会**重新读表重算**、
重算出的新引用能否编译出新计划、新计划是否**只执行一次**并回读、确认是否不重写、
放弃能否恢复。这条链是 A1-E2，仍需在真实 Excel / WPS 上串联验证。

## 两端操作（原生复验用，不使用 computer-use）

前置：服务端以 `EXCEL_BRO_RESULT_REF_FAULT=1` 启动，前端用开发构建（开发验收模式）。

1. 打开**独立测试工作簿**（Excel 或 WPS 各一次），保持固定选表范围；
2. 展开面板底部「结果引用过期（开发专用）」，点「启用下一次引用过期」
   —— 状态应显示"已启用：…（一次性，N 秒内有效）"；
3. 发一条需要 Python 分析的请求（例如"读取来源表 B2:B4 乘 2"），等模型拿到结果、
   准备写入预览；
4. 预期：模型这一轮的**结果引用被拒**（固定码 `result_ref_invalid`，界面提示"结果
   引用无效或已过期，请重新查询"），**目标表一个单元格都没被写**；
5. 记下这一次的行为（是否重新读算、是否给出新结果）—— 这就是 A1-E2 要验收的内容；
6. 面板状态应变为"已触发：旧引用被拒绝并已删除；后续任务不受影响"。同一任务后续的
   第二条 Python 结果、以及之后的新任务都**不受影响**；若未触发，10 分钟后自动作废
   （可「取消验收」）。
