# WPS 格式动作恢复账本链路修复（2026-10-13）

> 起因：真实 WPS 实机验收在「确认写回」时失败，界面原文
> **「无法持久登记本次写入现场，已取消写入；原表未被修改。WPS 恢复描述的动作定位或形状无效」**。
> 本文记录根因、修复、证据分档，以及**恢复实机验收所需的重启与重新生成计划步骤**。

## 1. 根因：格式动作没有可用的现场证据（四层同时断）

| 层 | 位置 | 断点 |
| --- | --- | --- |
| 出站契约 | `apps/excel-addin/src/contracts.ts` `assertWpsRecoveryDescriptor` | 动作类型硬编码 `["writeValues","writeFormulas","clearRange"]` ⇒ **格式动作在 fetch 之前**就被拒，界面报「动作定位或形状无效」 |
| 模型层 | `server/app/models.py` `WpsActionEvidence.type` | `Literal` 只允许三个值动作 ⇒ 写前检查点被 422 |
| 服务端闸 | `server/app/main.py` `_WPS_RECOVERABLE_ACTION_TYPES` | 10 项里错误放行了端上没接通的 `setFill`/`setNumberFormat`/`setBorders`/`autofit`，而属性级校验完全缺失 |
| 现场证据 | `apps/wps-addin/src/wpsExecutionLedger.ts` | 恢复描述按**值矩阵指纹**生成。格式动作**不改值**，值矩阵指纹在写前/写后**完全一样**，对格式改动没有区分力 |

四层里任何一层单独放宽都不够：只加三个枚举名，登记出来的会是一份"永远达不到"的记录，
服务端随后仍会以「证据与登记的写前状态不一致」拒掉写前检查点。

## 2. 本轮修复

### 2.1 格式证据指纹（`f1:`）—— 三种现场可区分

`apps/wps-addin/src/wpsPlanExecutor.ts`：

- `wpsFormatEvidenceFingerprint(type, fields, snapshot)` / `formatEvidenceFingerprint(action, snapshot)`
  —— 由**现场快照**在**已声明属性**上算（行高 `rowHeights`、其余逐格）。
- `formatExpectedFingerprint(action, rowCount, columnCount)` —— 由**声明值 + 区域形状**算写后目标。
- 三个函数共用一套归一化，保证「回读现场指纹 === 登记目标指纹」。
- 只含枚举数字/布尔/字符串，**不含任何单元格内容**；快照本身不出站。

### 2.2 属性级合成写前现场 `composedFormatSnapshot(actions, index)`

同区域多个格式动作按顺序执行时，后一个动作的"写前"是前一个动作执行后的状态。
拿预览**原始**快照当基准会同时坏掉写前复核与账本登记，所以：

- 只看前序动作**声明过**的属性；
- 行高按**整行**判相交（`A2:A3` 与 `C3:C4` 在第 3 行上确实互相影响，列不相交也要叠加）；
- 其余属性按矩形真正相交的格。

三处调用同一份合成结果：执行器写前复核、副本审阅写前复核、账本 `previewFingerprint`。

### 2.3 与登记端命名空间隔离

- 格式动作：指纹必须是 `f1:`，且必须声明非空 `formatFields`（属性名）。
- 值动作：不得声明格式属性，指纹不得是 `f1:`。

### 2.4 只读现场核对认识格式动作

`apps/wps-addin/src/wpsRecoveryInspection.ts` 原先一律按值矩阵指纹核对现场 —— 格式动作不改值，
`inspectWpsRecoveryState` 永远判不出「已完成 / 未开始」，中断后的**对账与续跑整条走不通**。
现在按登记声明的属性名采集当前格式快照再算 `f1:` 指纹，且：

- 属性名不可识别 ⇒ `unknown`（不用剩余属性凑哈希）；
- 未声明属性 / 快照不完整 ⇒ `unknown`；
- 全程**零 setter**（用例逐次取证）。

### 2.5 服务端第三道闸

`server/app/main.py` `_validate_wps_format_evidence`：值动作不得声明格式属性；格式动作必须声明
非空且属该动作白名单内的属性；两个指纹都必须是 `f1:`。`models.py` 的 `WpsRecoveryActionDescriptor`
新增 `formatFields` 并逐名白名单校验，`WpsActionEvidence.type` 扩到 6 个已开放动作。

**旧恢复记录**：值动作记录不受影响（`formatFields` 默认空列表）；格式动作此前从未登记成功过，
因此不存在需要迁移的旧格式记录。旧记录若不符合新协议，`_wps_record` 一律 fail-closed 拒绝下发。

## 3. 证据分档（本轮）

| 分档 | 状态 | 证据 |
| --- | --- | --- |
| 模块集成 | **已通过** | `wpsFormatExecutor.test.ts` 14 ✓、`wpsFormatCombinationBoundary.test.ts` 12 ✓、`wpsRecoveryInspection.test.ts` 20 ✓ |
| **跨进程链路**（真实面板 + 真实出站校验 + 真实 HTTP + 真实 ASGI 服务端 + 真实 SQLite 账本；仅替身 WPS 宿主、规划固定） | **已通过** | `npm run test:wps-format-ledger` → 2/2 ✓（`outputs/wps-format-ledger-run2.log`） |
| 服务端账本 API | **已通过** | `server/tests/test_wps_execution_api.py` 47 ✓（含格式动作登记/检查点/回读/对账/续跑/白名单同步/快照走私拒绝） |
| 反向变异 | **已通过** | `node scripts/run_wps_format_ledger_mutations.mjs` → 5 个变异各自变红，4 个文件逐字节还原，复跑全绿（`outputs/wps-format-ledger-mutations.log`） |
| 两端回归 | **已通过** | WPS 796 ✓ / Excel 1096 ✓；`tsc --noEmit` 两端 0 错；`build:wps` / `build:addin` 退出码 0 |
| 服务端全量 | **部分** | 1397 ✓、28 skip、2 ✗。两项失败在 `test_python_event_delivery.py`，原因是本机 **Docker 隔离引擎不可用**（`runner_missing`），属环境阻断，与本轮改动无关 |
| 真实模型自主规划 | **未完成** | 跨进程用例里规划是固定的；本轮不构成该证明 |
| 真实 WPS 客户端 | **未完成** | 见第 5 节步骤，需人工恢复实机验收 |
| 真实 Excel 客户端 | **未完成** | 本轮未涉及 |
| 工件回放 / 双端在线 | **环境阻断** | `npm run test:cross-process` 与 `npm run test:online` 都因 Docker 不可用而**拒绝**用替身顶替通过（退出码 2） |

## 4. 复跑命令

```bash
node scripts/mutate_wps_format_ledger.mjs check          # 变异前置状态闸
npm run test:wps-format-ledger                            # 跨进程链路（自带临时服务端）
node scripts/run_wps_format_ledger_mutations.mjs          # 反向变异（5 变异 + 逐字节还原）
npx vitest run --config vite.config.ts src/wpsFormatExecutor.test.ts   # 在 apps/wps-addin 下
.venv/Scripts/python.exe -m pytest server/tests/test_wps_execution_api.py -q --basetemp=.tmp-wps-api
```

## 5. 恢复真实 WPS 验收：重启与重新生成计划

本轮改动落在**两端的前端代码**（`contracts.ts`、`wpsPlanExecutor.ts`、`wpsNativeComparison.ts`、
`wpsExecutionLedger.ts`、`wpsRecoveryInspection.ts`）与**服务端**（`main.py`、`models.py`），
因此两边都必须重启，且**必须重新生成计划**（旧预览快照是旧代码采集的，不适用于新的登记口径）。

1. **完全退出 WPS**（不是关文档）。WPS 会缓存加载项资源，只关窗格不生效。
2. **重新构建 WPS 加载项**：`npm run build:wps`（本轮已构建通过）。
3. **重启后端服务**（普通后端，WPS 验收用这一条）：
   `npm run dev:server`（`http://127.0.0.1:8765`）。
   - **不要**同时起 SDK 后端。若要用 SDK 后端，只能走 `scripts/start_claude_sdk_server.ps1`，
     二者不能并存。
4. **确认服务端能力表与仓库一致**：本轮未改 `config/capabilities.json`
   （`wpsWrite.allowedActions` 已含 `setFont`/`setAlignment`/`resizeRange`，`maxCells` 1000）。
   如果这次验收用的是受控能力表（`scripts/start_manual_python_acceptance.ps1` 会逐字复制），
   重新跑一次那个脚本以刷新副本。
5. **重新打开 WPS → 打开专用测试工作簿**（必须是**已保存**的 `.xlsx`；未保存没有 `wps-path:` 身份）。
6. **重新发送请求**，重新走一遍「预览 → 生成右侧副本并排审阅 → 确认写回原表 → 回读 → 撤销」。
   旧对话里的预览按钮不适用 —— 计划的格式快照必须由新代码重新采集。
7. 若界面提示上次的恢复点需要核对，按面板提示核对后**放弃旧恢复点**再重新生成计划；
   失败过一次的登记不会在服务端留下记录（判据见 `test_wps_format_descriptor_refuses_snapshot_smuggling`）。

**验收关注点**（与自动化口径对齐）：`setFont` + `setAlignment` 同区域 + 一个行高动作；
确认前原表零写入；写回后格式落位、行高按**整行**生效、区域内外值/公式一字未动；撤销按属性级恢复。

## 6. 红线（本轮未越过）

- 不上传完整计划、原始矩阵或格式快照；出站的只有定位、属性名与 `f1:` 指纹。
- 登记失败即中止，**不跳过登记、不吞错继续写**。
- 未验证属性（`underline` / `vertical` / `columnWidth`）继续拒绝，未新增任何容差。
- 端上未接通的动作（`setFill` / `setNumberFormat` / `setBorders` / `autofit`）继续 fail-closed。
- 本轮**不宣称实机通过**，未提交、未推送、未发布。
