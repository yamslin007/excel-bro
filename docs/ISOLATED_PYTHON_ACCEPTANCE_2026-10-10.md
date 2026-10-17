# 隔离 Python 全链路 · D/E/F 最终验收报告（要求 5 + 6，2026-10-10）

> 本报告覆盖用户点名的**要求 5（一个通用合成案例贯穿全部流程）**与
> **要求 6（最终验收）**。要求 1–4 的完成记录见 `.workbuddy/memory/2026-10-10.md`
> 与各自套件注释。**自动化证据 / 真实隔离容器证据 / 真实模型 / 真实 Excel / 真实 WPS 五类分开报告，互不代替。**

---

## 0. 一句话结论

**一个通用合成案例已贯穿「两张来源表 → 清洗/关联 → 汇总 + 异常明细 → 结果卡 →
预览 → 确认写回 → 回读 → 撤销」全链**，且双端（Excel Office.js / WPS JSAPI）用
**同一套规则、同一批脏数据、同一套期望值**。全链关键守卫经**反向变异**证明有判别力
（拆掉变红、还原变绿、无残留）。仓库要求的全量测试与两端生产构建**全部通过**；
本轮 Docker 引擎可用，**真实隔离容器执行已实测通过（15 + 151 项）**，与自动化测试分列。

---

## 1. 通用合成案例（要求 5）

### 1.1 数据（用户点名的全部缺陷类型）

`apps/excel-addin/src/syntheticResultEndToEnd.test.tsx` 与
`apps/wps-addin/src/WpsAgentPanel.test.tsx`（WPS 端新增 describe）用**同一批数据**：

| 表 | 行 | 会员编号 | 说明 |
| --- | --- | --- | --- |
| 会员表 | 2 | `"  7  "` | 键含**前后空格**（trim 后 = `7`） |
| 会员表 | 3 | `"7"` | **重复键**（保留第一次出现） |
| 会员表 | 4 | `"007"` | **前导零**编号（≠ `7`） |
| 会员表 | 5 | `""` | **空键**（不参与关联） |
| 会员表 | 6 | `"9"` | 正常键 |
| 订单表 | 5 | `"404"` | **未匹配**（会员表里没有） |

### 1.2 先定死的清洗/关联规则（不由实现决定）

1. 键一律 `trim()` 后比对；但**形态保留**：`007` 与 `7` 是**不同的**键。
2. 空键（trim 后为空）**不参与关联**，计入"空键"明细。
3. 重复键**保留第一次出现**，后续计入"重复键"明细，不静默合并。
4. 关联为**左连接**（订单为左表）：订单里没有对应会员的记录计入"未匹配"明细。

### 1.3 手算期望值（独立于实现）

- 分组键 = trim 后的会员编号，按订单表聚合：`7`→100、`007`→200、`9`→300、`404`→400
- **键集合** `["007","404","7","9"]`；**共 4 笔**；**合计 1000**
- 异常明细：**空键 1**、**重复键 1**、**未匹配 1（键 404）**

### 1.4 引擎语义（先钉死，不是假设）

产品查询引擎 `executeQueryTableData`（`apps/excel-addin/src/dataTools.ts`，两端同源）：

- **必需字段是全部请求字段的并集**：某表缺任一必需字段 ⇒ **整张跳过** +
  `「表名」未找到字段：…` 警告，且 `complete=false`。故字段集不同的表**必须分开查询**。
- **按原始值分组**（`JSON.stringify` 做键），**不代做**去空格/去重键/跨表关联 ——
  清洗/关联是**分析步骤本身**，由本案例显式实现。

> 诚实边界：案例**不使用**引擎的隐式"清洗"，而是把规则显式定死并实现；用例第 0 段
> 把上述引擎语义断言为前置事实，避免"拿实现当期望"。

### 1.5 各段覆盖

| 段 | 内容 | 关键断言 |
| --- | --- | --- |
| 第 0 段 | 引擎语义钉死 | 并集缺字段整张跳过 + `complete=false`；原始值分组得 5 组 |
| 第 1 段 | 清洗/关联规则 + 汇总 | 手算期望逐项；**换表名/字段名（甲表/乙表、编号/款项）后仍成立** |
| 第 2 段 | 真实 `<App/>` 点击链 | 确认后**只写目标表**（来源逐格未动）、写日志恰一次、逐单元格回读、撤销逐格回执行前 |
| 第 3 段 | 真实结果转计划 | `collectPythonData` → **白名单编译** → `prepareOfficePythonResult` → 写前复核；来源被改 ⇒ 拒绝零写入；目标不在声明来源 ⇒ 编译拒绝 |
| WPS 段 | 桥真读真算 → 结果计划 → 预览 → 确认 → 写回 → 回读 | 收据里 headers/rows 与 Excel 端**同一套期望**；异常明细到达预览 |

---

## 2. 真实容器执行（**与自动化分列**）

本轮 Docker 引擎**可用**（此前为环境阻断），已实测：

| 项 | 结果 |
| --- | --- |
| `verified_endpoint()` | `npipe:////./pipe/docker_engine`（本机 npipe） |
| `verify_engine()` | `True` |
| 沙箱镜像 | `excel-bro-python-sandbox@sha256:812d1cce…`（**按 digest 固定**） |
| `server/tests/test_python_docker_real.py` | **15 passed**（真实容器：网络隔离正/负对照、文件哨兵、限额、清理身份核对） |
| docker defects/mutations/runner/worker 套件 | **151 passed** |
| **合成案例真实容器复算** | `run_isolated_analysis_internal` 在真实容器内跑清洗/汇总 ⇒ `{"keys":["007","404","7","9"],"total_count":4,"total_amount":1000,"unmatched":["404"]}` —— 与手算**逐项一致** |

> 说明：上表第 4/5 行是真实容器证据；第 6 行是"合成案例规则在真实容器里复算"的
> **结果一致性**证据。它**不**替代"真实 Excel/WPS 客户端 + 真实模型"的实机链路（见 §5）。

---

## 3. 反向变异验证（拆掉必红）

`scripts/mutate_synthetic_case_guards.mjs` + `scripts/run_synthetic_case_mutations.mjs`：

| 变异 | 拆掉的守卫 | 结果 |
| --- | --- | --- |
| S1 | 白名单编译不再要求"目标表在声明来源里" | **1 failed**（负向用例）→ RED(OK) |
| S2 | 预览不再核对来源快照与 `scope.sources` 一致 | **2 failed**（core 结果流，两端各一）→ RED(OK) |
| **S3+S3b** | 一次拆全"写前复核 + `captureSources` 指纹比对"**冗余层** | **1 failed** → RED(OK) |
| S4 | 缺字段不再整张跳过 | **1 failed**（第 0 段）→ RED(OK) |

> **S3 单拆是 GREEN(NG)**：`captureSources()` 自身也会因来源值变化抛错 —— 那是
> **冗余层**，属**判别点选错**；必须与 S3b 同时拆才算一次完整变异。已固化进驱动。
> 驱动末尾断言源码无 `变异S` 残留（`pythonResult.ts` / `pythonData.ts` /
> `officePythonResult.ts` / `dataTools.ts`）。

要求 1–4 的变异见 `scripts/run_office_undo_mutations.mjs`（M1–M6b，六组全 RED(OK)）。

---

## 4. 全量测试与构建（要求 6）

| 项 | 命令 | 结果 |
| --- | --- | --- |
| core | `vitest run`（packages/core） | **389 passed / 25 files** |
| Excel 端 | `vitest run`（apps/excel-addin） | **1040 passed / 91 files** |
| WPS 端 | `vitest run`（apps/wps-addin） | **675 passed / 39 files** |
| 服务端 | `pytest server/tests`（`.tmp-pytest` basetemp） | **1302 passed / 1 skipped**（161s） |
| Excel 构建 | `npm run build:addin` | **✓ built**（`tsc` 通过 + vite 产物） |
| WPS 构建 | `npm run build:wps` | **✓ built**（`tsc --noEmit` 通过 + vite 产物） |
| 两端 tsc | `tsc --noEmit -p apps/*/tsconfig.json` | **EXIT=0 / EXIT=0** |
| lint | `npm run lint` | **25 problems（1 error + 24 warnings）**，**0 项由本轮引入**（详见下） |

> 服务端唯一 skip 的原因：`server/tests/test_python_sandbox_build.py:47`
> `symlink creation unavailable`（Windows 非管理员无法建符号链接）—— 与隔离执行
> 本身无关，非环境阻断。本轮 basetemp 走仓库内 `.tmp-pytest`，规避了 OS 临时目录
> safe-delete 阈值问题。

**lint 唯一 error 为既存问题，非本轮引入**：
`apps/wps-addin/src/wpsScopeDraft.test.tsx:48:5` `react-hooks/rules-of-hooks` ——
该行是 `vi.mock("./WpsAgentPanel", () => ({ default: ({…}) => <div…/> }))` 里 mock 的
**默认导出箭头函数**（被 eslint 误当作 React 组件），`git show HEAD:` 版本**同一位置
同一报错**（HEAD 版在第 38 行、结构逐字相同）。本轮对该文件是**新增测试用例**
（+213/−20），未触碰该 mock 结构。24 个 warning 全为既有 `exhaustive-deps` 提示，
分布见 `outputs/lint-dfe.log`。

---

## 5. 实机验收（最短步骤 + 期望结果）—— **实机待验**

> **交付顺序（按要求 6）**：**先**交付可重复的**完整自动化案例**（§1 测试文件 + §3 变异脚本
> 驱动，一条命令即可复跑全链），**再**给出测试工作簿、最短实机步骤与期望结果（本节）。
> 自动化复跑命令：
> `npx vitest run apps/excel-addin/src/syntheticResultEndToEnd.test.tsx` 与
> `node scripts/run_synthetic_case_mutations.mjs`（后者自带基线/还原/残留校验）。

以下为**真实 Excel / WPS 客户端**上的人工验收步骤；本报告**不声称**已实机通过。

### 5.1 测试工作簿

新建 `合成案例.xlsx`，三张表：

- `会员表`（A1:C6）：表头 `会员编号|会员|等级`；数据行按 §1.1。
- `订单表`（A1:B5）：表头 `会员编号|金额`；数据行按 §1.1。
- `汇总表`：空表。

### 5.2 步骤

1. 打开工作簿 → 在加载项里点范围选择 → **勾选「会员表」「订单表」「汇总表」** → 完成。
2. 发送：「按会员编号汇总订单金额，结果写到汇总表」。
3. 依次确认：数据登记 → 分析 → 结果卡 → 写入预览（看到三条异常明细）→ 确认写回。
4. 观察回读核对结果 → 点撤销。

### 5.3 期望结果

- `汇总表!A1:C5` =
  `[["会员编号","订单数","金额合计"],["7",1,100],["007",1,200],["9",1,300],["404",1,400]]`
- 预览/结果卡出现三条异常明细：空键 1、重复键 1、未匹配 1。
- 确认前 `会员表`/`订单表` **一个单元格都没变**；撤销后 `汇总表` 逐格回空。

---

## 6. 已实现 / 未实现 / 宿主限制

### 已实现（自动化 + 真实容器）

- 数据登记桥、隔离容器执行、结果登记、白名单计划编译、结果预览、确认写回、回读、撤销。
- 通用合成案例全链（双端同源语义）；全链关键守卫反向变异验证。

### 未实现（如实列出）

- **精确续跑剩余动作**：登记的是 `WpsRecoveryDescriptor`，剩余非空只能 `replan`。
- **SDK 后端检查点恢复**：当前接入**未实现**（普通后端有检查点）。
- **同表不重叠的跨重启断点恢复**：**不承诺**（当 `officeRecoverySupport` 不支持时
  `registerOfficeExecution` 不调用；支持时才登记，见下条）。
- **原生审阅与恢复登记互斥**：**已澄清（旧结论作废）**。曾据"审阅要求唯一授权表、
  `officeRecoverySupport` 又对动作改来源表一律不支持"推断 App 路径恒不可达 ——
  该结论**在第五轮已修**：授权范围 / **实际来源依赖**（`plan.sourceDependencySheets`，
  只来自 `sdkBridgeHost.execute` 的 `recordTrustedToolReads`）/ 写入目标**三者分开建模**
  后，`officeRecoverySupport` 只看**实际来源依赖**判读写重叠，正常计划（动作落在目标表、
  与来源不重叠）**会**走到 `registerOfficeExecution`。证据：
  `apps/excel-addin/src/excelWriteEndToEnd.test.tsx`「断点恢复登记的真实可达性」7 项
  —— 真实 `<App/>` 点击链上断言登记**恰好 1 次**、登记成功后才写目标表、4 条失败路径
  原表零写入。**限制保留**：真实 Excel 实机仍未验证；来源依赖**缺失**时按既有保守条件
  退回指纹范围（依赖不确定就说不确定，不用"授权范围减写入目标"伪造来源）。
  本轮（续 6）**已复核可达性并删除全部"恒不可达"表述**，详见
  `docs/CONTINUOUS_CHAIN_ACCEPTANCE_2026-10-10.md` §4（含"`WorkbookSnapshot.workbookIdentity`
  是死字段"亦已作废）。

### 宿主限制

- **WPS 图片通道**始终不可用（不渲染 disabled 图片按钮）。
- **跨表公式重算**：首批结果写入用**定值矩阵**，**不覆盖**公式重算语义。
- **隔离执行**：仅已验证 Docker 隔离执行器；**运行时缺失即 fail-closed，绝不回退宿主裸执行**。

### 实机待验

- 真实 Excel 客户端 + 真实模型 + 真实 WPS 客户端上的端到端人工验收（§5）。
  本报告未运行实机，**不标"已验证"**。
- **分环节口径以 `docs/CONTINUOUS_CHAIN_ACCEPTANCE_2026-10-10.md` §5 为准**：
  真实模型（自主规划）、真实 Excel、真实 WPS 三项写 **未完成**，**不**统一写成"仅待实机" ——
  后两者在**自动化层面也未闭环**（用的是替身宿主，替身证据 ≠ 实机证据）。

---

## 7. 未提交声明

本轮全部改动**保留在工作区，未 commit / 未 push / 未发布**。变异脚本与驱动的
`restore-all` 已确认源码无 `变异S` / `变异M` 残留。
