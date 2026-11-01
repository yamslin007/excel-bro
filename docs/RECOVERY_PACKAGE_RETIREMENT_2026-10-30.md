# 删除前的恢复材料留痕（U5，2026-10-30）

> 一句话：恢复材料被删掉的那一刻，账本里要留下"删的是哪一份"。
> 在此之前 `retireWpsRecoveryPackage` 直接删 —— 材料消失之后，"当时那份材料是不是这一份"
> 只能靠推断，而终态回收恰恰是**唯一**合法的删除点。

## 1. 为什么补（缺口原文）

`docs/A1_GATE_REMAINING_2026-10-28.md` §1 U5：

| 项 | 现状 | 为什么不能靠"实机再跑一次"解决 |
| --- | --- | --- |
| 删除前的恢复材料留痕（哈希/快照） | `retireWpsRecoveryPackage` 直接删；删前不存证 | 删掉之后无法回答"当时那份材料是不是这一份" |

E3 那一格的证据链有一环是"材料消失了"（`docs/SDK_A1_E3_STOP_AND_GAPS_2026-10-25.md`）。
"消失"是可观测的，"消失的是哪一份"不是 —— 差的就是这一行。

## 2. 做了什么

### 2.1 账本（`server/app/task_ledger.py`）

| 项 | 内容 |
| --- | --- |
| 新列 | `recovery_package_digest TEXT`、`recovery_package_bytes INTEGER`、`recovery_package_retired_at REAL`、**`recovery_package_removed_at REAL`**（`_SCHEMA_VERSION` 12 → **13**，四列同时进 `_REQUIRED_OFFICE_COLUMNS`，旧库逐个 `ALTER TABLE`） |
| 新方法 | `record_recovery_package_retirement(operation_id, digest, size)`：删**之前**记"要删的是哪一份" |
| 新方法 | `mark_recovery_package_removed(operation_id)`：`remove()` 报告**真的删掉了**之后才记 |
| 只读方法 | `recovery_package_retirement(operation_id) -> {"digest", "bytes", "retiredAt", "removedAt"} \| None` |

### 2.1.1 两个事实，不许混成一句"删掉了"

这是本功能**最容易读错**的地方，代码、测试、文档三处都按这一条写：

| 列 | 含义 | 能据此宣称什么 |
| --- | --- | --- |
| `recovery_package_digest` / `_bytes` / `retired_at` | 删除**之前**记下的"那一份的身份"与时刻 | 只能说"我记下了要删的是这一份"，**不能**说删掉了 |
| `recovery_package_removed_at` | `remove()` 返回 True 之后才写 | **才能**说"那一份确实被删掉了" |

于是"落账成功、删除失败"在账本里**读得出来**（前者有值、后者为空）——
这正是不能只凭 `retired_at` 宣称删除成功的原因。
两组时刻都**首次写赢**（`WHERE … IS NULL`）：重试、二次回收都不改写第一次。

**不存材料正文**：恢复包里含单元格原值，账本里一个值都不许有
（`test_recovery_material_never_enters_the_task_ledger` 钉着这条，本功能的用例又加了一条正向锚点）。

### 2.2 取指纹（`server/app/execution_recovery.py`）

`raw_digest(operation_id) -> (sha256, bytes) | None`：

- 摘要取的是**文件字节**，不是 header 里那个"内容摘要" —— 终态回收时材料可能已经**过期或损坏**
  （那恰恰是最需要留痕的时刻），header 不可信、也不一定读得出；被删掉的对象本来就是这堆字节；
- **不**校验协议、不解析正文、不判期限、不碰 mtime：它只回答"是哪一份"，不回答"能不能用"
  （可用性是 `load` / `status` 的事）；
- 读失败返回 `None` ⇒ 调用方**拒绝删除**，绝不拿猜出来的摘要去留痕。

### 2.3 端点顺序（`server/app/main.py`）

`DELETE /api/execution-recovery/{operation_id}`：

```
核身份（_wps_record）→ 取指纹 → 落账（retired）→ 删（remove）→ 删成功才记「删除确认」（removed）
```

| 情形 | 结果 | 账本 |
| --- | --- | --- |
| 正常删除 | 204 | `retired_at` + `removed_at` 都有 |
| 落账失败 | **503 `RECOVERY_PACKAGE_RETIREMENT_UNRECORDED`，材料不删** | 什么都没有（材料还在） |
| 落账成功、删除失败 | 507 `RECOVERY_PACKAGE_PERSISTENCE`（或同类） | 只有 `retired_at`，`removed_at` 为空 |
| 删了但确认记不下 | **503 `RECOVERY_PACKAGE_REMOVAL_UNRECORDED`**（如实说"删了但没记下确认"） | 只有 `retired_at`，`removed_at` 为空 |
| 材料本来就不在 | 204 | 不留痕（不能瞎记一个指纹） |
| 身份不符 / 没登记 | 409 | 不建记录、不留痕 |
| 删除失败后重试 | 第二次 204 | `retired_at` 与指纹保留第一次，只补 `removed_at` |

### 2.4 与 U4 的口径**正好相反**，不许互相照抄

| | 处置留痕（U4） | 材料回收留痕（U5） |
| --- | --- | --- |
| 写在哪一步 | 宿主恢复**成功之后** | 删除**之前**（+ 删成功再记确认） |
| 为什么 | 没恢复就不许说"撤销了" | 删完就无从取证 |
| 失败怎么办 | 吞掉（best-effort，不影响用户结论） | **不删**（503，材料留到期） |
| 冻结的是 | 第一次"用户处置"的时刻 | 第一次"材料被交出"的指纹与两个时刻 |

## 3. 判别力（反向变异，四判据）

`scripts/mutate_recovery_package_retirement.py`，五点各打一个判别点：

| # | 拆掉什么 | 预期变红 | 判别点 |
| --- | --- | --- | --- |
| R1 | 删前不再取指纹（`material = None`） | `test_delete_records_the_digest_of_the_bytes_it_deleted` | 留痕到底存不存在（连带红 7 条） |
| R2 | 落账失败也照删（去掉那条 `raise`） | `test_ledger_failure_refuses_to_delete_and_keeps_the_material` | 失败时材料是否还在 |
| R3 | 四列从 `_REQUIRED_OFFICE_COLUMNS` 里拿掉 | `test_legacy_ledger_is_migrated_with_the_three_retirement_columns` | "版本号只是意图，列才是事实" |
| R4 | **删除失败也写"删除确认"** | `test_removal_failure_records_the_digest_but_no_deletion_confirmation` | 有没有凭 `retired_at` 宣称删除成功（连带红重试用例） |
| R5 | 删除确认**不**"首次写赢" | `test_first_digest_and_first_removal_survive_a_second_retirement` | 二次回收会不会改写第一次的删除时刻 |

五条都跑成四判据全成立（退出码非 0 / 预期用例名在 FAILED / 失败原文带自家证据 /
还原后逐字节 sha256 一致且复跑 12 passed 全绿）。基线：
`main.py be700642…`、`task_ledger.py cedd8258…`。

> R3 的判据必须把 `user_version` 调到**当前**版本才有意义：版本落后时 bootstrap 反正会重跑、
> 列会被补回来，那条用例就等于没在测这份名单。这也是 2026-10-27 真机 500 的真实形态
> （版本号看着是新的、列却没补）。

## 4. 验证

| 层 | 命令 | 结果 |
| --- | --- | --- |
| 新增用例 | `python -m pytest server/tests/test_recovery_package_retirement.py` | **12 passed** |
| 相邻既有 | `test_execution_recovery_package.py` + `test_execution_disposition.py` + `test_task_ledger.py` + `test_execution_recovery_a1.py` | **100 passed** |
| **服务端全量** | `python -m pytest server/tests --basetemp=<OS 临时目录>` | **1934 passed / 27 skipped / 2 failed**（仍是环境无 Docker 的 `execute_python_analysis_async` → `unavailable`，与本改动无关，改动前同红） |
| 编译 | `python -m compileall -q server/app` | 退出码 0 |
| 变异 | `mutate_recovery_package_retirement.py R1–R5 check` | **五条四判据全部成立** |

用例覆盖（12 条，对应用户点名的每一格）：删前指纹即删除时字节 · 两个事实分得开 ·
删除失败后重试（保留首次指纹与时刻、只补确认）· 落账后删除失败 · 重复删幂等 ·
重新落材料再删（首次摘要与首次删除时刻都保留）· 材料不存在 · 未登记执行 · **身份不符（跨命名空间）** ·
留痕里没有单元格值 · 旧库补四列 · 指纹非法与记录缺失。

**diff 检查**：本轮 `git diff` 只落在这三处 —— `task_ledger.py`（四列 + 两个方法 + 迁移）、
`main.py`（DELETE 端点顺序与两种失败码）、`execution_recovery.py`（只读 `raw_digest`）；
顺手修掉上一轮遗留的一处**行粘连**（`def append_office_host_writes(` 的参数被挤到同一行），
修完复跑 `test_task_ledger.py` + 本文件 **41 passed**，两处变异 R3/R5 用新基线重跑仍全成立。

## 5. 边界（别读错）

1. **不追溯**：本功能上线前删掉的恢复材料永远没有这条留痕。`NULL` 一律读作"没有留痕"，
   不是"材料还在"，也不是"没被删过"。
2. **`removed_at` 为空 ≠ 删除失败**：它只表示"没有确认到删掉"，可能是删除报错、也可能是
   探测到有材料但删到时已不在。要判原因看那次请求的错误码。
3. **指纹不是"材料可用"的结论**：它只说明"当时磁盘上那堆字节是什么"。材料是否过期/损坏，
   仍由 `status` / `load` 回答。
4. **不改变删除的触发条件**：仍然只在"任务真的结束"（原地撤销成功）之后调，执行成功时
   **不许**调 —— 那时审阅、核对、放弃都还要用这份材料。
5. **不新增对外读接口**：留痕留在 `office_executions` 四列里，取证时直接读账本；
   要不要开一个只读端点由后续验收口径决定，本轮不擅自加。
6. **只覆盖 WPS 侧**：恢复包本身就是 WPS 的能力（Excel 侧没有本机恢复材料），
   所以这条留痕天然只有 WPS 会写。
