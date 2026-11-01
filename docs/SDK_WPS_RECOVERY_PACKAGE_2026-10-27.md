# 完整退出宿主后仍可核对并恢复原执行：本机恢复包（2026-10-27）

## 1. 真实缺陷（先复现，再修）

真实 WPS 现场：计划里三个独立的 `writeValues` 动作，第二步已写入并回读，在**完成回执前**精确中断。

1. 完整退出 WPS，重新打开同一个工作簿；
2. 服务端账本仍然找到这个任务，核对第二步后显示「已补记写入，未重复写入」；
3. 接着提示「**找不到原执行的本机恢复记录**」，第三步无法继续。

两个问题必须分开处理，不能互相顶替：

| 问题 | 根因 | 本批处理 |
| --- | --- | --- |
| 本机恢复材料丢失 | `prepareWpsResume` 只从宿主 `PluginStorage` 读撤销记录；完整退出后那里是空的，也没有任何其它材料来源 | 材料改存**本机后端应用数据目录**，宿主缓存降级为缓存/索引 |
| 恢复资格过期 | 只有 `taskRecovery.wpsResumeMaxAgeSeconds`（1800s）这一个窗口期 | 持久材料自带 `expiresAt`（默认 7 天）；两个窗口期分开，查询与打开窗格都不延长（第一轮仍留了两套期限并存的缺口，第二轮收成唯一判据，见 §10） |

**没有**把"材料丢失"归因于 30 分钟，**没有**只延长 TTL 就说修好，**没有**断言 `PluginStorage` 在所有环境都不持久 —— 它只是在这条退出路径上不可靠，因此不再是唯一来源。

## 2. 材料放哪、放什么

- 目录：`application_data_root()/execution-recovery/<operationId>/material.dat`
  （`EXCEL_BRO_CONFIG_DIR` → 否则 `%LOCALAPPDATA%\ExcelBro`）。**不放** Temp，**不放**工作簿旁边，
  **不上云**，**不进模型上下文**。
- 文件结构：第 1 行元数据 JSON（`headerVersion/operationId/digest/size/createdAt/savedAt/expiresAt`）
  + `\n` + **原样字节正文**。用"一行元数据 + 原样正文"而不是把正文塞进 JSON 字符串，
  是为了让读回的内容与客户端写下的内容**逐字节相同**（客户端据此做落盘读回校验）。
- 正文 = `ExecutionRecoveryPackage`（`packages/core/src/executionRecoveryPackage.ts`）：
  原计划动作顺序与每步目标值/写前值/公式、计划版本（`plan:v1:N:hex8`）、`executionId`/`taskId`、
  宿主与工作簿身份（`wps-path:...`）、原始授权范围、来源版本证据、写前/写后指纹、
  创建时间、格式版本与完整性信息。

**分工不许混淆**：

- 恢复包 = **不可变原始材料**。写一次，之后只读（同 `operationId` 换正文一律拒绝）。
- 账本（`office_executions`）= **可变执行进度**：`nextActionIndex`、在途动作、租约、`revision`。
  恢复时游标只认账本，不采信包里的"当时进度"。

## 3. 正确性约束与落点

| 约束 | 落点 |
| --- | --- |
| 先落盘后写表；首次宿主写入前原子落盘 + 读回校验 | `wpsPlanExecutor.executeWpsPlan` 的 `recovery.persistMaterial`（在 `try` 内、`record.state="applying"` 之前）；`persistWpsRecoveryPackage` 读回**逐字节**比对 |
| 落盘失败必须明确停止且零宿主写入 | 落盘抛错归一进执行器 `catch` ⇒ `state="restored"`、`error` 保留原始原因、`appliedRows` 全 0 |
| 不在执行成功时删材料 | 材料只在**终态回收**（撤销成功、原表回到执行前）时删；确认后仍要能撤销，所以不删 |
| 崩溃窗口 | 材料已存未写 / 已写未回执 / 回执已提交但最终回执失败 / 写完未审阅 / 进程在保存或清理中退出：由"材料在盘上 + 账本游标 + 宿主现场三态核对"共同判定，任一侧缺失即停 |
| 恢复前重新核对身份/授权/来源/已完现场/在途/剩余 | 复用既有 `prepareWpsResume`（水合后原样工作），不另写一套"从包里读"的并行判定 |
| 已确认完成的前两步不得重写 | 续跑只执行 `position >= startActionIndex` 的后缀；用例用宿主写调用记录 + 动作身份证明 |
| 在途动作能证明与预期写后状态一致时只补记 | 既有 `inspectWpsRecoveryState` 三态（`applied`/`not_applied`/`unknown`） |
| 不凭账本推断工作簿一定保存过 | 水合时 `appliedRows`/`afterFingerprint` 只由账本游标**描述**，"到底写没写"仍读宿主现场 |
| 单执行者 | 账本 `claimWpsExecutionResume` 的 `revision`/`owner_token` CAS；领取失败一律不写 |
| 身份不能退回按文件名匹配 | `wps-path:<FullName>` 精确串，且包/账本/现场三方一致 |
| 原值不进出站/日志/遥测/异常 | 包正文只走本机接口与本机文件；异常消息只说状态（缺失/过期/损坏/冲突/配额） |

## 4. 双端状态

| 能力 | Excel | WPS |
| --- | --- | --- |
| 登记执行时的 before-image 持久化 | 既有实现：服务端 `office_executions.before_snapshot_json` / `checkpoint_json` / `result_json.undoSnapshot`，本批复用未改 | 本批新增 `execution-recovery` 恢复包 |
| 本机发现索引 | 既有 `officeOperationRecovery.ts`（只存不含快照的 `operationId`） | `PluginStorage` 撤销记录 + `localStorage` 待恢复索引 |
| 完整退出宿主后取回材料 | 既有路径已满足（材料在服务端账本） | 本批实现：`ensureWpsRecoveryMaterial` → 水合回宿主缓存 |
| 跨进程格式版本/白名单校验 | 共享协议 | 共享协议（同一份 `config/execution_recovery_package_cases.json` 双端共读） |
| 终态回收 | 未改（沿用既有账本生命周期） | `DELETE /api/execution-recovery/{id}`，只在撤销收尾调用 |
| 精确恢复的动作范围 | 既有范围 | **仅 `writeValues`**；格式/公式/结构/删表**不承诺**跨进程精确续跑 |
| 原生宿主验收 | 未做（本批） | **未做** |

## 5. 保留策略（与执行权、Python 临时结果分开）

| 材料 | 生命周期 | 位置 |
| --- | --- | --- |
| 执行权租约 | 短时独立管理（`wpsOwnerLeaseSeconds=120`） | 账本 |
| Python 临时结果引用 | 原有独立缓存策略未改 | 账本 + 结果引用表 |
| 未完成任务恢复包 | `unfinishedRecoveryPackageRetentionDays=7`（`config/capabilities.json`） | 本机应用数据目录 |
| 审阅与撤销材料 | 独立生命周期（撤销点不被本批改动） | 宿主缓存 + 账本 |

- `expiresAt` 在**首次落盘**时算定；同正文重复保存不改 `expiresAt`，查询与打开窗格都不延长。
- 容量：`recoveryPackageMaxBytes`（单包 32MiB）、`recoveryPackageTotalMaxBytes`（合计 256MiB）、
  `recoveryPackageMaxOperations`（200）。不足时先清理**已过期**的，仍不足就明确报 507 并拒绝新增，
  **绝不**淘汰尚未结束任务的恢复包。
- 状态：`missing` / `expired` / `corrupt` / `conflict` / `quota` / 落盘失败各有独立 code，
  界面据此说"缺失""过期""材料损坏"，不合并成一句"无法恢复"。
- 旧任务材料真的丢了：如实说能力不足，**不伪造** before-image，不用空记录冒充"已安全恢复"。

## 6. 与 A2 的关系

本批只做"材料跨进程可用"，不碰账本的保留期与去值：

- A2 仍负责：清理后旧 `operationId` 拒重放、终态去值、清理失败可见、损坏/旧版本记录行为、
  历史转录清理入口。
- 本批的恢复包**已经是去值化之外的独立存储**（不在账本里、不在事件里），A2 去值不得把它当成
  账本字段来清理；两者的保留期要各自显式配置，不能共用一套淘汰逻辑。

## 7. 反例（反向变异，两条机制各一个判别点）

`node scripts/mutate_wps_recovery_material.mjs`

| 变异 | 打掉什么 | 必须变红的用例 | 失败原文里的自家证据 |
| --- | --- | --- | --- |
| M1 | `ensureWpsRecoveryMaterial` 退回"只信宿主缓存" | 「full process restart loses PluginStorage but resumes only third from durable material」+ 面板「完整退出 WPS 清空宿主缓存后重开…」 | `expected 'none' to be 'package'` / `完整退出后必须能从本机恢复包续跑` |
| M2 | 拆掉执行器的写前落盘 | 同一用例的"中断后后端确实有材料"锚点 + 「zero host writes」 | `expected [] to include 'wps-resume-0001'` / `promise resolved instead of rejecting` |

四判据都核过：退出码 1、红的是预期那几条、失败原文带自家证据、还原后 `sha256` 与基线逐字节相同
（`wpsRecoveryPackage.ts=3730a7fa…`、`wpsPlanExecutor.ts=232e487d…`）。

## 8. 命令、退出码与证据路径

日志目录：`outputs/recovery-package-20261027/`（整段落盘，未用 `| tail` 截断取证件）。

| 检查 | 命令 | 退出码 | 结果 | 日志 |
| --- | --- | --- | --- | --- |
| 恢复包服务端 | `.venv/Scripts/python.exe -m pytest server/tests/test_execution_recovery_package.py -q -p no:cacheprovider --basetemp="C:/Users/jia/AppData/Local/Temp/eb-pytest-recovery"` | 0 | 19 passed | `server-recovery.log` |
| WPS 续跑 | `npx vitest run src/wpsExecutionResume.test.ts`（`apps/wps-addin`） | 0 | 29 passed | `wps-resume-green.log` |
| 面板·完整退出后重开 | `npx vitest run src/WpsAgentPanel.test.tsx -t "完整退出 WPS"` | 0 | 2 passed | `panel-restart-green.log` |
| WPS 全量 | `npx vitest run` | 0 | 1141 passed / 8 skipped | `wps-full.log` |
| Excel 全量 | `npx vitest run` | 0 | 1564 passed / 9 skipped | `excel-full.log` |
| core 全量 | `npx vitest run` | 0 | 859 passed | `core-full.log` |
| 双端类型检查 | `npx tsc --noEmit` | 0 / 0 | 两端均 0 | `tsc-wps.log`、`tsc-excel.log` |
| lint | `npm run lint` | 0 | 0 error / 36 既有 warning | `lint.log` |
| 文件夹模式（openpyxl 路径） | `pytest server/tests/test_folder_workbooks.py server/tests/test_folder_data.py` | 0 | 88 passed → 未受影响 | — |
| 反向变异 M1 | 见 §7 | 变异后 1 / 还原后 0 | 预期用例红 | `mutation-M1-red.log`、`mutation-M1-restored.log` |
| 反向变异 M2 | 见 §7 | 变异后 1 / 还原后 0 | 预期两条红 | `mutation-M2-red.log`、`mutation-M2-zero-writes-red.log`、`mutation-M2-restored.log` |

修复前的红灯（同一用例、修复前代码）：`expected false to be true`，抛错原文
「找不到原执行的本机恢复记录」——即用户在现场看到的那句话。

## 9. 最短原生验收步骤（**未执行**）

准备：独立测试工作簿（不要用正在用的那份）、隔离的后端实例（独立 `EXCEL_BRO_CONFIG_DIR` 与端口）。

1. 目标 A1/A2/A3 设为 `-101 / -202 / -303`；用「把来源三格分别写到目标三格」的固定值计划，确认来源三表已授权。
2. 执行这三步，在**第二步写入并回读之后、完成回执之前**中断（开发模式精确故障入口）。
3. 确认现场：A1=`-101` 以外的第一步写入值、A2=第二步值、A3 仍是 `-303`；保存工作簿。
4. **完整退出 WPS**（不是关侧栏、不是关窗格），重启本地后端，再打开同一工作簿。
5. 打开窗格：应显示可核对、可继续；先核对，界面只能说"第二步已写入，未重复写入"，不得重写前两步。
6. 点继续：只写第三步，目标应为 `第一步值 / 第二步值 / 333`；宿主写调用次数与账本只推进 `actionIndex=2`。
7. 点确认变更：值不变、不再执行；再点撤销可回到 `-101/-202/-303`。
8. 另起一轮同样中断，第 6 步后点放弃变更：应回到 `-101/-202/-303`，并且本机恢复材料被回收（目录消失）。

没有第 4 步的完整退出，就不算验证本缺陷；没有 1–8 的原生证据，A1-GATE 不通过。

---

# 第二轮（同一日）：有效期判据与"保存不可续期"

第一轮把材料搬到了跨进程可读的位置，但留了两个**已确认**缺陷：有效期内仍被旧缓存期限拦住、
过期材料重新保存又拿满一个保留期。两条都先写失败回归，再改生产代码。

## 10. 问题 1：持久包仍被旧的 30 分钟校验拦截

**根因**：`prepareWpsResume` 前一段已经用 `record.resumeExpiresAt` 判过有效期，后一段的综合条件里
仍留着无条件的
`now - Date.parse(record.createdAt) > capabilities.taskRecovery.wpsResumeMaxAgeSeconds * 1000`
（`apps/wps-addin/src/wpsPlanExecutor.ts`）。于是**同一份材料被两套期限同时裁决**：持久包写着 7 天，
缓存窗口说 30 分钟，后者无条件生效 —— 7 天内、超过 30 分钟一律"已过期"。

**修改**（`wpsPlanExecutor.ts`）：

- 新增**唯一**有效期判据 `wpsResumeWindow(record, now)`：
  - 无记录 ⇒ `找不到原执行的本机恢复记录。`
  - `createdAt` 不可解析或晚于 `now` ⇒ 拒绝（未来时间不放行）；
  - **有** `resumeExpiresAt` ⇒ 以它为唯一期限；**字段存在但不可解析 ⇒ 明确拒绝**，不退回另一种期限；
  - **无** `resumeExpiresAt`（旧宿主缓存）⇒ 沿用 `wpsResumeMaxAgeSeconds`。
- 到期边界统一为 `now >= deadline` 即过期，与服务端 `deadline <= moment` 同一口径（`now == expiresAt` 算过期）。
- `prepareWpsResume` 的综合条件**删掉** `age > …` 与重复的 `sameJsonEvidence`，改为先过
  `wpsResumeWindow` 再走其余判据。
- **有效期不等于免检**：身份、授权、来源证据、目标现场、计划指纹、在途动作、执行权/`revision`
  校验一条未删 —— 有效期内仍要逐项核对现场。
- `packages/core/src/executionRecoveryPackage.ts` 的 `executionRecoveryPackageVerdict` 同步把边界
  从 `now > expires` 改成 `now >= expires`（原实现会放过"恰好到期"那一瞬）。

## 11. 问题 2：相同材料过期后重新保存会重新获得保留期

**根因**：`server/app/execution_recovery.py` 的 `save()` 只在"同正文且 `deadline > now`"时直接返回；
同正文**已过期**时继续往下走，生成新的 `savedAt`/`expiresAt` 并覆盖文件 —— 等于每次重复 PUT 都能
把保留期刷新一次，与"相同 `operationId` 幂等保存不刷新有效期"直接冲突。

**修改**：

| 位置 | 行为 |
| --- | --- |
| `execution_recovery.save()` | 同正文且未过期 ⇒ 幂等返回，`savedAt`/`expiresAt` 原样（不重写文件）；同正文**已过期** ⇒ 抛 `RecoveryPackageExpired`；正文不同 ⇒ 仍 `RecoveryPackageConflict`；文件读不出/头损坏 ⇒ `RecoveryPackageCorrupt`，**不再**被 `except (OSError, RecoveryPackageCorrupt)` 当成"没有旧文件"而覆盖 |
| `default_expiry()` / `_deadline()` | 期限只在一处算与读；`registered` 期限同样受 `deadline <= moment` 约束 |
| `task_ledger.reserve_recovery_package_expiry()` | 新列 `office_executions.recovery_package_expires_at`（含建表与迁移）。**首次写赢**：只有第一次登记能写入，之后只读；执行不存在返回 `None` |
| `PUT /api/execution-recovery/{id}` | 先向账本登记/读取**首次**期限，再用它调 `save()`。于是**文件被过期清理删掉之后再 PUT，也拿不到新期限**（不是"删了再建"绕开幂等） |
| `DELETE /api/execution-recovery/{id}` | 204 幂等；先核账本身份再删 |
| `_quota_exceeded` | 语义改为 `count+1 > maxOperations or used+incoming > totalMaxBytes`（旧参数把"旧文件还在"重复扣了一次） |

账本只多一列**时间戳**，不放单元格正文 —— 材料仍在恢复包里，账本去值口径不变；"不得通过删除过期
文件再新建来实现幂等"因此也不成立：资格由账本那一列决定，删文件不重置它。

## 12. 问题 3（本轮发现并按同一原则修）：面板"点了没反应"

第一轮把"本机材料是否可用"接进了**挂载**与**刷新**两处的结论，却漏了**点击前的那次复核**：
`runWpsRecoveryAction` 自己又算了一套（只看身份/现场，不看材料）。于是持久包过期时界面显示
"需重新规划"，点下去按另一套条件判成"可以继续"，判定不含该动作就直接 `return` —— 按钮看似失效。

**修改**：新增 `wpsRecoveryVerdictFor(state)` 作为**唯一**计算点（身份 + 现场 + 本机材料，
`resumeAvailable` 与 `resumeUnavailableReason` 出自同一次判断），挂载、刷新、点击前复核三处
全部改用它。这仍是"同一份判据、不许有两套期限"的同一条纪律。

## 13. 先红后绿与变异证据（M3/M4）

用例与变异体一一对应，**变异就是把旧行为原样放回**，所以"变异后变红"同时证明了用例在旧代码下是红的：

| 变异 | 打掉什么 | 判别点 | 变红用例 | 失败原文 |
| --- | --- | --- | --- | --- |
| M3 | 在 `wpsResumeWindow` 之后插回无条件 30 分钟检查 | 有效期判据的**来源** | `a durable package older than the 30-minute cache window still resumes only the third action` | `expected false to be true` |
| M4 | `save()` 的"已过期就抛"改回"未过期就返回同正文" | 重复保存**不能续期** | `test_expired_material_cannot_be_re_saved_into_a_new_window` | `DID NOT RAISE <class '…RecoveryPackageExpired'>` |

M3 同时打红 `durable expiry boundary equals the server's, and an expired package writes nothing`
（它先断言 `deadline-1` 可恢复 —— 30 分钟窗口下这一步就不成立）。M4 只打红一条：
`test_pruned_material_cannot_be_re_saved_past_the_first_expiry` 走的是"显式传入已过期的首次期限"，
变异体没有改那条路径 —— 这正是"一个变异一个判别点"，不靠多红几条撑场面。

四判据：退出码 1 / 红的是上面这些用例名 / 失败原文带自家证据 / 还原后 `sha256` 逐字节一致
（`wpsPlanExecutor.ts=5c6b2240…`、`execution_recovery.py=d569b954…` —— 后者是当时的基线，
第三轮加入 `_instant()` 后变为 `97e69f8d…`，见 §20）。

先红后绿的另外两处（不是变异，是用例契约随语义更新）：

- 面板「真实过期材料重新规划全链」原来靠把宿主缓存 `createdAt` 改成 2000 年模拟过期。新语义下
  持久包的 `expiresAt` 才是权威，改缓存时间点**既不延长也不缩短**那份材料，所以改为
  `backendRecovery().forceExpire(...)`（本机后端真的读不出正文，对应线上 410）。改之前该用例红在
  `expired recovery entry`（界面给的是"可继续剩余动作"）。
- 面板「点「重新读取并规划」」原来依赖**上一条用例残留**的 `validatePlan` 替身实现，单独跑会拿到
  `undefined` 并走进 `waitForPlanning` 的 `reading 'then'`。现在用例自己 stub，用例之间不许靠泄漏接续。

## 14. 第二轮实跑命令与退出码

日志目录：`outputs/recovery-package-20261027/`（整段落盘）。

| 检查 | 命令 | 退出码 | 结果 | 日志 |
| --- | --- | --- | --- | --- |
| WPS 续跑专项 | `npx vitest run src/wpsExecutionResume.test.ts` | 0 | 33 passed（原 29） | `wps-resume-green.log` |
| WPS 面板 | `npx vitest run src/WpsAgentPanel.test.tsx` | 0 | 187 passed / 5 skipped | `panel-full-round2.log` |
| WPS 全量 | `npx vitest run` | 0 | 1145 passed / 8 skipped（连跑两次） | `wps-full-round2.log` |
| Excel 全量 | `npx vitest run` | 0 | 1564 passed / 9 skipped | `excel-full-round2.log` |
| core 全量 | `npm run test:core` | 0 | 860 passed | `core-full.log` |
| 服务端全量 | `pytest server/tests -q --basetemp=<OS 临时目录>` | 0 | 1870 passed / 1 skipped | `server-full.log` |
| 恢复包专项 | `pytest server/tests/test_execution_recovery_package.py -q` | 0 | 26 passed（原 19） | `server-recovery-green.log` |
| 双端类型检查 | `npx tsc --noEmit` | 0 / 0 | 两端均 0 | `tsc-wps-round2.log`、`tsc-excel.log` |
| 双端构建 | `npm run build:wps` / `npm run build:addin` | 0 / 0 | 均通过 | `build-wps.log`、`build-excel.log` |
| compileall | `python -m compileall -q server/app` | 0 | 通过 | `compileall.log` |
| lint | `npm run lint` | 0 | 0 error / 36 既有 warning | `lint.log` |
| `git diff --check` | — | 0 | 仅 11 行 LF/CRLF 提示（工作副本 LF，非本轮引入） | `diff-check.log` |
| 反向变异 M3 | `node scripts/mutate_wps_recovery_material.mjs M3` | 变异后 1 / 还原后 0 | 预期用例红 | `mutation-M3-red.log`、`mutation-M3-restored.log` |
| 反向变异 M4 | 同上 `M4` | 变异后 1 / 还原后 0 | 预期用例红 | `mutation-M4-red.log`、`mutation-M4-restored.log` |

`--basetemp` 一律放在 **OS 临时目录**（`_is_under_os_tmp_dir` 豁免），不放仓库内：环境里的
`Path.unlink` 包装器会在同一回合累计删除 50 个受管路径后 `SystemExit(1)`，计数跨进程，越跑越假红。

## 15. 本轮仍然**没有**做的事

- 真实 WPS 一次都没跑：上面所有"完整退出"都是**替身宿主**（清空 `PluginStorage` + 换新应用对象 +
  同一份本机后端材料）模拟的，**不等于**真实 WPS 已验收。
- 真实 Excel 同样未跑；Excel 侧走既有 before-image 路径，本批未改。
- 只覆盖 `writeValues`；格式/公式/结构/删表不承诺跨进程精确续跑。
- 未启用 SDK 原生会话持久化，未新增执行前确认按钮，未扩大勾选工作表范围，未操作任何用户工作簿。
- **A1-GATE 仍未通过，A2 未启动。**


---

# 第三轮（同一日）：无时区到期时间，与"单条坏材料拖垮整批清理"

前两轮解决了"材料在哪"和"期限谁说了算"。这一轮是两个**独立**的缺陷：一个把
"读得出但没法用"的时间当成可比较的时刻，一个让一条坏材料连累整批清理。

## 16. 问题 1：无时区 `expiresAt` → `TypeError`

**复现（独立临时目录，不碰真实数据）**：用真实 `save()` 存一份合法包，只把
`material.dat` **首行元数据**的 `expiresAt` 改成 `2035-01-01T00:00:00`，正文、摘要、
长度一字不动，然后分别调 `load()`、同正文 `save()`、`prune_expired()`。
三个入口都抛：

```
TypeError: can't compare offset-naive and offset-aware datetimes
```

**根因**：`datetime.fromisoformat("2035-01-01T00:00:00")` **不报错**，返回的是
**naive**（无时区）时间；随后与 `datetime.now(timezone.utc)` 比较就炸。而且当时
**三处各写了一份** `fromisoformat(...replace("Z","+00:00"))`：
`_deadline()`（保存路径）、`load()`、`_prune_locked()`。所以"只修一处"必然还有两处抛。

接口层实测：修复前 `GET /api/execution-recovery/{id}` 返回 **503** —— `TypeError` 不是
`RecoveryPackageError`，走不进那张状态映射表，于是既不是 409 也不是 500 的固定文案，
而是通用错误处理。这不是"正常恢复全坏"：正常生成的 UTC（`...Z`）一直没问题，
受影响的只有"字段被改成无时区"这一条形态。

**修改**（`server/app/execution_recovery.py`）：

| 改动 | 内容 |
| --- | --- |
| 新增 `_instant(value)` | **读取、保存、清理共用的唯一解析口**。只接受 `str`、非空、能被 `fromisoformat` 解析、且 `tzinfo` 与 `utcoffset()` 都非空 —— 缺任何一条返回 `None`（= 无法确定）。`Z`、`+00:00`、`+08:00` 等合法偏移都按**同一时刻**比较；无时区**不猜本机时区、也不自动补 UTC**（猜错方向会让材料提前失效或永不失效，都比"明确说损坏"糟） |
| `_deadline(header)` | 只调 `_instant(header.get("expiresAt"))`；缺失 / `null` / 错误类型 / 非法文本 / 无时区一律 `None` |
| `load()` | 用 `_deadline()`；`None` ⇒ `RecoveryPackageCorrupt("恢复包期限字段损坏。")`，不再自己解析、不再漏 `TypeError` |
| `save()` | 账本传入的首次期限改走 `_instant()`；无效 ⇒ 既有 `RecoveryPackageError("恢复包保留期限字段无效。")`（HTTP 422），不落盘 |
| `_prune_locked()` | 用 `_deadline()`；`None` ⇒ **不删**、计入 `unconfirmed`，见 §17 |
| 新增 `_aware(moment, what)` | 内部时刻的调用约定：`load(now=)`、`prune_expired(now=)`、`default_expiry(moment)` 传 naive 一律显式拒绝（`RecoveryPackageError`），堵住"只修文件读取、漏掉另一个比较入口" |
| `status()` | 先验 `now` 再进 `try`：调用方写错不该被折成"材料损坏"这种关于材料的结论 |

## 17. 问题 2：单条损坏材料会中断整批清理

`_prune_locked()` 原来有两个问题：

1. **一条坏材料拖垮整批**：期限比较在 `for` 循环里裸奔，`TypeError` 直接冒到
   `prune_expired()` 外面 —— 同目录其它包一个都没处理，而配额清理恰恰是最需要它跑完的时候。
2. **"读不出期限"被当成"已经过期"**：非法文本走 `except ValueError: deadline = None`，
   随后 `if deadline is not None and deadline > moment: continue` 落空 ⇒ **删掉**。
   删掉就没有排查依据了，而且这等于把"不确定"当成"过期"。

**修改**：逐条 `try/except`（`OSError` → `failed`，其它异常 → `unconfirmed` 且**不删**）；
期限读不出 ⇒ `unconfirmed += 1` 并 `continue`；计数协议变为
`{"removed": 已确认过期并删除, "failed": 清理动作失败, "unconfirmed": 材料在但期限无法确定}`。
三个计数**不许互相顶替**：`unconfirmed` 既不算清理成功，也不算清理失败；损坏材料原样保留
（不重写、不生成新期限）。`save()` 的配额路径仍调 `_prune_locked()` 并忽略计数（行为不变）。

## 18. 先红后绿

新增 15 条用例（A–D），**先跑红，再改实现**：修复前 13 failed / 2 passed，两条通过的是
本来就该成立的对照（`Z`/偏移等价、恰好到期即过期）。

| 组 | 用例 | 修复前 |
| --- | --- | --- |
| A | `test_naive_expiry_reads_as_corrupt_not_a_type_error` | `TypeError` @ `execution_recovery.py:370` |
| A | `test_naive_expiry_refuses_resave_and_leaves_bytes_identical` | `TypeError` @ `:316` |
| B | `test_prune_clears_only_confirmed_expiry_and_keeps_undecidable_material` | `TypeError` @ `:226`（整批中断） |
| C | `test_expiry_offsets_compare_as_the_same_instant` | 通过（对照） |
| C | `test_exact_expiry_instant_is_already_expired` | 通过（对照） |
| C | `test_unusable_expiry_values_are_corrupt_and_never_pruned[×7]` | 7 条全红（`TypeError` 或"被当成过期删掉"） |
| C | `test_registered_expiry_must_carry_an_explicit_timezone` | 红（naive 被接受） |
| C | `test_naive_now_is_refused_instead_of_guessed` | 红（`TypeError` 而非明确拒绝） |
| D | `test_http_naive_expiry_maps_to_corrupt_without_echoing_material` | `assert 503 == 409` |

修复后同一文件 **41 passed**（原 26）。第二轮语义的回归一并重跑：WPS 续跑专项 33 passed
（>30 分钟仍可恢复、到期不可恢复、删除重存不可续期），未受影响。

## 19. 反向变异 M5

`node scripts/mutate_wps_recovery_material.mjs M5` —— 去掉 `_instant()` 的时区校验
（= 恢复旧解析行为，naive 照收）。

| 判据 | 结果 |
| --- | --- |
| 退出码 | 变异后 1 / 还原后 0 |
| 红的用例名 | `test_naive_expiry_reads_as_corrupt_not_a_type_error`、`test_naive_expiry_refuses_resave_and_leaves_bytes_identical`、`test_unusable_expiry_values_are_corrupt_and_never_pruned[2035-01-01T00:00:00]`、`test_http_naive_expiry_maps_to_corrupt_without_echoing_material`（4 failed / 8 passed） |
| 失败原文 | `TypeError: can't compare offset-naive and offset-aware datetimes` |
| 还原后哈希 | `execution_recovery.py=97e69f8d…`，与基线逐字节一致；随后重跑该文件 41 passed |

**`test_prune_clears_only_confirmed_expiry_and_keeps_undecidable_material` 在 M5 下仍然绿** ——
这是有意的：它钉的是"整批不中断 + 不删未确认材料"，由 §17 的逐条容错保证；`TypeError` 被
`except Exception` 兜住后照样记 `unconfirmed`、照样不删。也就是说"能不能解析"（M5 打红）
与"一条坏了会不会连累整批"（B 组用例）是**两条机制**，各自有判别点，谁也不替谁作证。

## 20. 第三轮实跑与边界

日志仍在 `outputs/recovery-package-20261027/`（新文件名，未覆盖旧证据）：

| 检查 | 退出码 | 结果 | 日志 |
| --- | --- | --- | --- |
| 恢复包专项（先红） | 1 | 13 failed / 2 passed | `r3-red-before-fix.log` |
| 恢复包专项（修复后） | 0 | 41 passed（原 26） | `r3-recovery-green.log` |
| 相关 API / 账本 | 0 | 103 passed（`test_execution_recovery_a1`、`test_wps_execution_api`、`test_task_ledger`、`test_ledger_recovery`、`test_recovery_claim_concurrency`） | `r3-related-api-ledger.log` |
| WPS 续跑专项 | 0 | 33 passed | `r3-wps-resume-green.log` |
| WPS 全量 | 0 | 1145 passed / 8 skipped | `r3-wps-full.log` |
| Excel 全量 | 0 | 1564 passed / 9 skipped | `r3-excel-full.log` |
| core 全量 | 0 | 860 passed | `r3-core-full.log` |
| 服务端全量 | 0 | 1885 passed / 1 skipped（原 1870） | `r3-server-full.log` |
| 双端构建 | 0 / 0 | `build:wps`、`build:addin` 均过 | `r3-build-wps.log`、`r3-build-excel.log` |
| compileall / lint / `git diff --check` | 0 / 0 / 0 | lint 0 error / 36 既有 warning；diff 仅 11 行 LF/CRLF 提示 | `r3-compileall.log`、`r3-lint.log`、`r3-diff-check.log` |
| 反向变异 M5 | 变异后 1 / 还原后 0 | 见 §19 | `r3-mutation-M5-red.log`、`r3-mutation-M5-restored.log` |

边界（本轮**没有**动）：`execution_recovery.py=97e69f8d…` 是第三轮的新基线，前两轮文档里
记的 `d569b954…` 属于上一版文件，不再适用。首次期限语义（只认账本、不可续期、`now >= expiresAt`）
一条未改；身份、选表授权、来源复核、现场核对、动作指纹、执行权判据一律未放宽；仍只覆盖
`writeValues`；没有执行前确认、没有 SDK 会话持久化、没有 A2。

**客户端侧的边界记录（本轮不改）**：`packages/core/src/executionRecoveryPackage.ts` 用
`Date.parse(expiresAt)`，只拒绝 `NaN`；按 ES 规范无偏移字符串会被当作**本机时区**。这条
路径拿到的 `expiresAt` 只来自服务端响应，而无时区的头在服务端已经被判损坏、根本出不来，
所以当前没有可达入口；记在这里是为了**下一次改动时不要忘了这个不对称**，而不是本轮顺手改协议。

**原生验收照旧未做**：本轮全部是自动化证据，没有打开真实 WPS/Excel。
**A1-GATE 仍未通过，A2 未启动。**

---

# 第四轮（2026-10-27 真机）：旧账本缺列 → PUT 500 → `Failed to fetch`

## 21. 现场与根因

真实 8765 服务的 `/api/diagnostics` 依次给出：execution register 200 → 恢复包 `OPTIONS` 200 →
**恢复包 `PUT` 500** → suspend 200 → `GET` 404。

服务端原文：

```
sqlite3.OperationalError: no such column: recovery_package_expires_at
```

链条：第三轮给账本加了列 `office_executions.recovery_package_expires_at`，同时把
`_SCHEMA_VERSION` 提到 10。但迁移的**唯一入口**是 `_run_schema_bootstrap`，而它的第一句是
`if _schema_is_current(path): return` —— 旧库的 `user_version` 如果是 9 就会迁移，可**如果版本号
被写成 10、列却没建**（第三轮之前就存在过的库、或迁移中途失败），检查只看版本号，于是永久跳过
补列。之后每一次 `PUT` 都在 `reserve_recovery_package_expiry` 上撞缺列。

浏览器把这层 500 折成 **`Failed to fetch`**，因为异常穿出路由、响应没有 CORS 头。用户看到的是
"网络不通"，而实际上是数据库少一列 —— 这也是"错误必须变成带 CORS 的结构化响应"的原因。

## 22. 修法：版本号**不等于**列存在

`_schema_is_current` 现在同时要求两件事：

```python
row = connection.execute("PRAGMA user_version").fetchone()
columns = {str(column[1]) for column in connection.execute("PRAGMA table_info(office_executions)")}
return bool(row) and int(row[0]) >= _SCHEMA_VERSION and "recovery_package_expires_at" in columns
```

迁移本身用守卫式补列，不重建表：

```python
if "recovery_package_expires_at" not in office_columns:
    connection.execute("ALTER TABLE office_executions ADD COLUMN recovery_package_expires_at TEXT")
connection.execute("PRAGMA user_version = %d" % _SCHEMA_VERSION)
```

两条合起来才闭合：检查认列（不再被版本号骗），迁移只补列（不会因为"版本号旧但列已在"而重复
`ADD COLUMN` 报错）。**版本号是意图，列是事实**，判据必须读事实。

服务端 PUT 的期限登记也包进了同一个错误映射：

```python
try:
    reserved = await run_in_threadpool(
        _folder_operation_ledger.reserve_recovery_package_expiry, operation_id, expiry)
except Exception as error:
    raise _recovery_package_error(error) from error
```

账本异常 → `_task_state_error` → **503 `TASK_STATE_UNAVAILABLE`**，带 CORS 头、不带数据库原文。

## 23. 客户端：失败之后能不能说"未执行"

PUT 500 之后写前落盘必然失败。失败要收尾，但**"收尾"不等于"当成没写过"**。执行器的 `catch`
会把 `record.state` 归一为 `restored`（一个动作都没 apply 时）或 `partial_write_recoverable`
（有 `appliedRows > 0`），面板据此分类证据。`restored` 是"未开始写入"**唯一**的正面证据。

`noActionStarted` 是四条判据的合取，缺一条都只能保留本机待办与恢复入口：

| 判据 | 打掉它会发生什么 |
| --- | --- |
| `evidence.kind === "not_started"` | 读失败 / 回执丢失 / 找不到记录 / 中间态都算"未执行" ⇒ 对着可能已改过的原表说假话 |
| `nextActionIndex === 0` | 账本已推进第二步也说没动过 ⇒ 重开后续跑会重放已完成动作 |
| `inFlightActionIndex == null` | 第 0 步正在写、回执没回来也说没动过 |
| `ownerActive === false` | 挂起失败（账本仍 `running`、本端仍握执行权）也说这次结束了 |

**一处真实教训**：我第一版在这四条之外又加了一条 `status !== "running"`。M6 拆掉它时四条用例
**全绿** —— 因为服务端 `_wps_state_payload` 用 `status == "running" && 租约未过期` 算
`ownerActive`，二者在真实数据里同步。多写的那一层抓不到任何变异，是装饰而不是防线，已撤回；
测试里 `ownerActive:true, status:"running"` 才是"挂起失败"的真实组合。

`packages/core/src/resultReview.ts` 相应新增 `ReviewNotStartedError`：只有宿主**正面证明**没开始
才把审阅置为 `not_started`，普通执行错误仍是 `failed`。缺这个区分，`not_started` 会变成"任何
失败都算没写"的通用挡箭牌。

## 24. 反向对照与变异

四条用例，一条打一处；另外补一条"旧库迁移不抢写锁"。

| 变异 | 判别点 | 打红的用例（唯一） | 还原 sha256 |
| --- | --- | --- | --- |
| M6 | `ownerActive === false` | 写前恢复准备失败但…挂起失败：账本仍 running… | `WpsAgentPanel.tsx=07682414…` |
| M7 | `nextActionIndex === 0` | …账本已推进到第二步… | 同上 |
| M8 | `inFlightActionIndex == null` | …第 0 步在途、回执还没回来… | 同上 |
| M9 | 证据读失败 ≠ 没写过 | 宿主恢复点读不出来时不得当成「没写过」 | `wpsExecutionRecovery.ts=61b7273b…` |
| M10 | 启动路径要不要写锁 | `test_read_is_not_serialized_behind_concurrent_write_lock` | `task_ledger.py=124b4ec5…` |

M6–M9 的失败原文一律是 `expected null not to be null` —— 即 `pending-wps-execution` 被清掉了，
正是这条机制管辖的那个键。M10 的失败原文是
`sqlite3.OperationalError: database is locked`（`PRAGMA user_version = 10` 拿不到写锁）。

M10 那次同时红了**新建的**一条与**既有的** F4 用例。它们读同一机制，属于重复证据；
按"冗余层一次拆全"删掉新建的那条，把增量（新列 `recovery_package_expiry` 也必须不抢写锁）
并进既有 F4 用例，让判别点保持唯一。

脚本：`scripts/mutate_wps_recovery_clear_gate.mjs`（一次一个 SEED、归一化 LF、按原 EOL 写回、
写回读回校验、`SIGINT/TERM/HUP` + `uncaughtException` 还原钩子）。

## 25. 旧库迁移：只补列，不碰执行本身

两条用例钉住"补列是纯 schema 动作"：

- 有进度的执行（`next_action_index=1`、`checkpoint={"cells":[["new"]]}`、`before_snapshot=NULL`）
  drop 列 + `user_version=9` 之后重启，`office_execution()` 与迁移前**整个 dataclass 相等**，
  `PRAGMA integrity_check = ok`；期限仍是"未登记"，随后第一次 `reserve` 生效、第二次不许推后。
- 版本号落后但列已在（`user_version=9`）：只能升级版本号，已登记的 `2035-01-01T00:00:00Z` 原样
  保留，重启后的 PUT 也推不动它。

"不抢写锁"沿用既有 F4 用例：另一个连接 `BEGIN IMMEDIATE` 时，新实例的 `get()` 与
`recovery_package_expiry()` 都要在 0.5s 内返回。早退（`_run_schema_bootstrap` 首句
`_schema_is_current` 为真即返回）就是靠这条守住的 —— 拆掉它，`PRAGMA journal_mode = WAL`
与 `ALTER TABLE` 会去抢写锁，直接 `database is locked`。

## 26. 升级步骤与边界

**升级步骤（用户侧最短操作）**：只重启 8765 的服务进程，不需要重建账本、不需要删 WAL、
不需要重新发任务。

1. 停掉当前 `dev:server` 进程；
2. 用同一份工作区重新启动（`npm run dev:server`）；
3. 新进程在第一次实例化账本时读到 `user_version=9` / 缺列，自动 `ALTER TABLE ... ADD COLUMN`
   并写下 `user_version=10`；
4. 打开诊断端点确认 `schema` 已到 10 —— 这一步只是核实，不是前提。

已存在的任务、动作进度、检查点、首次期限都保留（§25 的用例就是钉这个）。原先失败的 PUT 重发即可。

**没做的事**：没有操作任何运行实例（没有重启 8765、没有调用它的 API、没有改用户的库文件）；
没有删/重建账本、没有删 WAL、没有清空历史或恢复索引；没有打开真实 Excel / WPS。
**A1-GATE 仍未通过，A2 未启动。**

## 27. 第四轮实跑

日志在 `outputs/recovery-package-20261027/`，全部 `r4-` 前缀（未覆盖前三轮证据）：

| 检查 | 退出码 | 结果 | 日志 |
| --- | --- | --- | --- |
| 服务端专项（账本 + 恢复包） | 0 | 74 passed（原 72） | `r4-scoped-server.log` |
| core 审阅专项 | 0 | 28 passed | `r4-scoped-core.log` |
| WPS 面板专项 | 0 | 194 passed / 5 skipped（199） | `r4-scoped-panel.log` |
| core 全量 | 0 | 867 passed | `r4-core-full.log` |
| WPS 全量 | 0 | 1153 passed / 8 skipped | `r4-wps-full.log` |
| Excel 全量 | 0 | 1566 passed / 9 skipped | `r4-excel-full.log` |
| 服务端全量 | 0 | 1907 passed / 1 skipped（原 1885） | `r4-server-full.log` |
| 双端 `tsc --noEmit` | 0 / 0 | 无输出 | `r4-tsc-wps.log`、`r4-tsc-excel.log` |
| 双端构建 | 0 / 0 | `build:wps`、`build:addin` | `r4-build-wps.log`、`r4-build-excel.log` |
| compileall / lint / `git diff --check` | 0 / 0 / 0 | lint 0 error / 36 既有 warning；diff 仅 LF/CRLF 提示 | `r4-compileall.log`、`r4-lint.log`、`r4-diff-check.log` |
| 反向变异 M6–M10 | 变异后均 1，还原后均 0 | 见 §24 | `r4-mutation-M6..M10.log` |

## 28. 第五轮：完整退出后的**发现入口**（本轮主题）

### 28.1 现场（只读取证，不推断）

用户在 `WPS恢复验收.xlsx` 里执行三个独立常量写入（B2=111、B3=222、B4=333），在**第二步写入
并回读之后、完成回执之前**中断。目标值当时 111/222/-303，界面显示"待核对在途动作"。保存、
**完整退出 WPS**、重开同一个文件后：

| 事实 | 来源 | 结果 |
| --- | --- | --- |
| 账本状态 | `GET /api/wps/executions/{id}` 只读查询 | `status=needs_reconciliation`、`nextActionIndex=1`、`inFlightActionIndex=1`、`ownerActive=false` |
| 原表现场 | 人工读单元格 | 仍是 111 / 222 / -303（与账本"第二步在途"一致） |
| 授权 | 只读查询 | 仍是两张表 |
| 界面 | 窗格截图 | **恢复入口没了**，面板显示的是另一条历史单动作任务的"已撤销" |

结论：**账本进度、恢复材料、工作簿现场都还在，缺的只是"重新发现未完成执行"的入口。**
三个曾被怀疑的因素（浏览器索引丢失 / 前端 origin 改变 / 挂载与工作簿切换时机）中，
**只有"索引丢失"被证实为缺口本身**；另外两个既没有证据也没有被证伪 —— 因此它们不是本轮的
根因结论，只是本轮的**验收条件**（切换工作簿、元数据晚到都必须能重新发现）。

### 28.2 根因的形状：唯一入口是缓存

`pending-wps-execution.v1` 是 `localStorage` 里的**单条**索引，完整退出 WPS 时随宿主缓存一起
消失。旧的面板发现只有一条路：

```
loadPendingWpsExecution()  →  null  →  return（"没有未完成任务"）
```

账本里的记录、本机恢复包里的不可变材料、原表的写入现场**都不参与发现**，所以它们都在也没用。
既有"完整退出"用例之所以是绿的，是因为它在清 `PluginStorage` 之后**又手工写回了那条索引** ——
恰好绕过了这个缺口。

### 28.3 修法：索引降级为缓存，发现走后端

| 层 | 文件 | 内容 |
| --- | --- | --- |
| 共享判据 | `packages/core/src/pendingExecutionDiscovery.ts`（新） | 准入闸（白名单 `extra="forbid"` 语义）、排序、选择规则、文案；限额读 `config/capabilities.json` |
| 协议 | `apps/excel-addin/src/contracts.ts`、`server/app/models.py` | `PendingExecutionCandidate` / `PendingExecutionsResponse`；`schemaVersion=1`、`namespace`、`workbookIdentity`、`limit`、`truncated`、`candidates` |
| 服务端 | `server/app/main.py`、`server/app/task_ledger.py` | `GET /api/office/executions/pending`、`GET /api/wps/executions/pending`、`POST /api/wps/executions/{id}/dismiss-discovery`；账本 `pending_office_executions()` / `dismiss_office_execution_discovery()`；`user_version` 10 → **11**，新增 `discovery_dismissed_at` |
| 客户端 | `apps/excel-addin/src/api.ts` | `listPendingExecutions` / `listPendingWpsExecutions` / `listPendingOfficeExecutions` / `dismissWpsExecutionDiscovery`（404 视为已处置） |
| WPS 适配层 | `apps/wps-addin/src/wpsRecoveryDiscovery.ts`（新） | 读宿主身份（与登记同一函数 `wpsWorkbookIdentity`）、发请求、二次过闸、交给共享选择规则 |
| WPS 面板 | `apps/wps-addin/src/WpsAgentPanel.tsx` | 缓存优先、后端兜底；多候选卡；"暂时读不到"卡；身份观察（间隔/焦点/可见性）；代数闸 |

**硬边界（写进代码注释与用例）**：

1. **身份必须完整相等**。WPS 用 `wps-path:<小写磁盘完整路径>.xlsx`，Excel 用 `url:<归一化 url>`；
   两种命名空间互不冒充。**禁止**用文件名、活动表名、表结构、相同单元格值、"最近一条"推断归属；
   未保存簿 / 身份不可读 / 身份不合法时**一个请求都不发**。
2. **发现只恢复入口，不恢复写入资格**。认领、材料水合、身份/授权/来源/目标现场复核仍走既有
   那一条路径；本轮没有复制第二份恢复判据（`wpsRecoveryVerdictFor` 仍是唯一计算点）。
3. **响应只带最小元数据**：`operationId`、`taskId`、`status`、`createdAt`/`updatedAt`、
   进度三元组、三个布尔。**没有** `ownerToken`、单元格值、计划正文、恢复包正文，也没有
   别的执行记录。多带一个字段即被准入闸拒绝（因为服务端那侧是 `extra="forbid"`）。
4. **多候选不自动选**。排序把"有在途动作"放最前、把"写前失败"放最后（避免一条一个动作都
   没推进的记录挡住真正写了半截的那条），但**一条都不自动采用**；截断时即便只剩一条也要用户选。
   "已完成"一律不进候选；另一条任务不会被顺手清掉。
5. **"清除恢复提示"语义很小**：只写 `discovery_dismissed_at` 时间戳 + 清本地缓存。
   不改 `status`、不推进 `revision`、不动游标、不删材料；再次出现需要账本真有新变动
   （`dismissed_at < updated_at`）。
6. **发现全过程只读**：不 claim、不补记、不 start/finish、不写单元格。重开先出现"待核对在途
   动作"，用户显式核对之后才可能"可继续剩余动作"，之后只写第三步。

### 28.4 先红后绿

| 层 | 用例 | 先红 | 后绿 |
| --- | --- | --- | --- |
| 共享判据 | `packages/core/src/pendingExecutionDiscovery.test.ts`（13 条） | 实现文件未落地 ⇒ 找不到模块 | 13 passed |
| 服务端 | `server/tests/test_pending_execution_discovery.py`（16 条） | 16 failed | 16 passed |
| 服务端迁移 | `test_task_ledger.py::test_existing_ledger_without_discovery_column_migrates_without_touching_the_execution`（9/10/11 三档） | 缺列时读 pending 抛 `no such column` | 全部通过，执行行逐字段未变 |
| WPS 面板 | `WpsAgentPanel.test.tsx` 新增 10 条（见 §28.5） | 未接线时只能靠手工塞 `operationId`，删掉那条索引后 0 条能过 | 208 passed / 5 skipped |

### 28.5 面板回归覆盖（A / B / C）

**A 完整丢失前端缓存**：`storage.clear()` + `localStorage.removeItem(PENDING_KEY)` + 换新宿主
对象 + 后端账本留下真机形状（`needs_reconciliation/1/1`）→ 面板**经真实发现路径**找到原执行、
显示"待核对在途动作"、`listPendingWpsMock` 用当前工作簿身份查了一次，而
claim / reconcile / start / complete / finish / suspend **全部为 0**，宿主写入 0，账本快照逐字段未变，
卡面不含单元格内容。

**B 完整恢复链**（确认/放弃各一条）：显式点"核对在途动作" → `reconcileWpsActionReceipt` 收到
`actionIndex=1, outcome=applied`，**原表零写入**；随后"继续剩余动作" → `startWpsAction` 只出现
`actionIndex=2`，宿主只写 `Target!A3`；确认后值 `[10,20,30]` 且不再执行；放弃后回到
`["old1","old2","old3"]`。全程 `registerWpsExecution` / `runTurn` 均为 0（不是重新规划）。

**C 身份与竞态**：同名不同路径互不发现；先开另一簿再切回（靠焦点触发的身份观察重跑）仍能发现；
换簿但选表序号相同 ⇒ 旧卡必须撤掉、不认领；身份不可读（未保存）⇒ 一个请求都不发、保存后再发现
（元数据晚到）；旧查询迟到（扣住响应再放行）⇒ 不许画到新簿上；多候选不自动选/不自动查状态；
服务不可达 ⇒ "暂时读不到"、恢复可用后重新核对；材料损坏 ⇒ 按既有判据拒绝续写（只允许重新规划）。

### 28.6 反向变异 D1–D6（每个只打一个判别点）

`node scripts/mutate_wps_recovery_discovery.mjs D1 check`（…D6），四判据：退出码非 0、预期用例名
出现在 FAIL、失败原文带**它自己那处**证据、还原后逐字节 sha256 一致。

| 变异 | 拆掉的东西 | 必红的用例 | 失败原文证据 |
| --- | --- | --- | --- |
| D1 | 后端发现调用 | 完整退出 WPS 丢掉浏览器索引与宿主缓存后… | 索引丢了也必须经后端发现重新摆出待核对入口 |
| D2 | 缓存不校验归属 | 打开另一个工作簿时：属于原簿的本地索引不在这里摆卡 | 应当按当前簿的身份去发现 |
| D3 | 多候选自动选第一条 | 工作簿有多条未完成写入时只给选择入口… | 两条候选都必须摆出来给用户选 |
| D4 | 查不到当成没有 | 服务不可达时说"暂时读不到"… | 查不到必须如实说读不到 |
| D5 | 迟到响应的代数闸 | 旧查询迟到不覆盖新工作簿… | 换到新簿后迟到的旧响应不许画出恢复卡 |
| D6 | 身份观察只剩挂载一次 | 先打开另一个工作簿再切回原簿… | 切回原簿后必须重新发现，而不是停在"没有未完成任务" |

六条全部"四判据成立"，每次运行后 `WpsAgentPanel.tsx` 的 sha256 都回到
`aa0b20c8…63baa`（脚本自己打印基线，`check` 结束再比一次）。

### 28.7 测试隔离：未替身的网络出口一律失败

诊断期间发现**测试曾经真的在打用户那台 8765**：面板只替掉了 `api.ts` 的一部分函数，其余仍走
真实实现，而基地址是 `import.meta.env.VITE_API_BASE_URL ?? "http://127.0.0.1:8765"`。于是"跑一遍
全量测试"会留下 `POST /api/wps/executions/<测试合成 id>/host-writes` 这样的请求。

现在两端 `setupFiles` 都装 `packages/core/src/testNetworkGate.ts`：

- 任何未被 `vi.stubGlobal("fetch", …)` 或明确替身接管的 `fetch` / `XMLHttpRequest` /
  `WebSocket` / `sendBeacon` **立刻抛错**，错误原文带"测试网络闸门"；
- 调用方 `catch` 掉也没用：`afterEach` 会因为有未处理的请求让这条用例失败；
- **没有允许列表** —— 允许某个真实地址就是给下一次泄漏开口子；
- 那些"故意要跑真实实现"的接缝（桥轮询、暂停请求、计划校验）用
  `guardNetworkSeam(real)`：**只有本用例自己装了 fetch 替身才真的调用**，否则按"连不上"失败。

装上闸门后一次性抓出并封住的真实出口：`/executions/{id}/host-writes`、`/execution-recovery/{id}`
（DELETE）、`/worksheet-checkpoints/list`、`/api/turn/{id}/events`、`/api/turn/{id}/pause`、
`/api/plans/validate`、`/api/sdk/tool-bridge/poll`、`/__wps_probe`（DEV 遥测）。

### 28.8 本轮边界（不许读成"实机验收通过"）

- **已自动化到**：真实面板组件 + 真实共享判据 + 真实服务端端点（独立 SQLite）+ 替身宿主 +
  真实 Chrome/jsdom 渲染。
- **未做**：真实 Excel / WPS 窗格里点过"待核对在途动作"一次都没有；真实宿主完整退出后的
  自动发现没在原生环境复核过。
- **A1-GATE 仍未通过，A2 未启动。** 本轮只解决"发现能力"，不是 A1 整体完成。

## 29. 第六轮：授权改正了，结论没跟着改（2026-10-28）

§28 那条 `source: "selected"` 保护写过头了：它挡住了"别的候选改判这一条"，也顺手挡住了
"对**这一条**重新算"。真机现场因此卡住 —— 用户误选三张表 → 卡面报"需重新规划：工作簿、宿主版本
或授权已变化" → 用户回面板取消那张不该勾的表（三张 → 两张，`authorizationKey` 确实变了、发现
effect 确实重跑）→ 卡面仍停在旧阻断，只能靠"重新规划"（等于丢掉这次执行）脱身。

### 29.1 改了什么

| 位置 | 改动 |
| --- | --- |
| `WpsAgentPanel.tsx` `reverifySelectedWpsRecovery` | 新增：`operationId`/`taskId` 取自当前卡，只 `GET` 状态，再用 `wpsRecoveryVerdictFor` 同一份判定重算；不认领、不补记、不执行、不写宿主 |
| `WpsAgentPanel.tsx` `discoverWpsRecovery` | selected 分支拆成三种现场：身份一致 ⇒ 重算同一条；身份已变 ⇒ 不动（交给换簿）；身份不可读 ⇒ 不动、不猜 |
| `WpsAgentPanel.tsx` `wpsRecoveryRef` | 类型补 `taskId`（"还是不是同一条"要按它判，`operationId` 单独不够） |
| `wpsPlanExecutor.ts` `assertWpsResumeSources` | 三种不匹配各自点名（身份 / 授权集合 / 宿主版本），判据顺序与 fail-closed 不变，只报数量与版本，不打印路径与表名 |

### 29.2 判据没变，变的只是"在哪一步问"

重新核验与首次摆卡、点击前复核**共用** `wpsRecoveryVerdictFor`
（`wpsIdentityVerdict` + `inspectWpsRecoveryState` + `recoveryMaterialReason` +
`wpsResumeUnavailableReason`）。面板里没有第二套"宽松版能不能续跑"，所以不存在"刷新后比点击时更宽松"。

### 29.3 反向变异（`scripts/mutate_wps_recovery_reverify.mjs`）

- `R1` 把 selected 分支改回直接 `return` ⇒ 「授权从三张表收回成原两张后…重新核验并改判」必红，
  失败原文带 `收回第三张表后必须对同一条执行重新核验…`；
- `R2` 把迟到写回的三道闸（两处 `isCurrent()` + `operationId/taskId` 比对）**一次拆全** ⇒
  「连续切换授权时迟到的核验结果不许覆盖新现场」必红，失败原文带 `先发那次的迟到结论不许翻案`。
  只拆一层时另两层会先把变异体挡住、用例照样绿 —— 冗余层必须一起拆，否则会得出"写回条件不重要"。

两条都是四判据（退出码 / 预期用例名 / 自家证据原文 / 还原后逐字节 sha256），
`WpsAgentPanel.tsx` 基线 `97c0ba01…2dc0cd`，跑完与基线一致。

顺带修了 `mutate_wps_recovery_discovery.mjs` 的 `D2` 定点：上一节删掉"身份为 null 也采纳索引"
后那行 SEED 已不存在，脚本退 3（`SEED 出现 0 次`）。重新定点后 D1–D6 六条全部四判据成立。

### 29.4 仍未做

真实 WPS 窗格里"取消第三张表 → 卡面改判 → 点继续 → 只写第三步 → 回读 → 审阅"一次都没跑过。
本轮证据是真实面板组件 + 真实服务端（隔离 SQLite）+ 替身宿主 + jsdom。**A1-GATE 仍未通过，A2 未启动。**
