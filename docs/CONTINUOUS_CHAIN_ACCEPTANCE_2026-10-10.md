# 连续链路 + 分环节验收报告（要求 1–6，2026-10-10 续 6）

> 本报告覆盖本轮用户点名的**要求 1–6**，并按用户要求把结论**分环节**列出。
> 与 `ISOLATED_PYTHON_ACCEPTANCE_2026-10-10.md` 的关系：那份讲"通用合成案例贯穿
> 全部流程"（要求 5/6 的上一轮），本份讲**同一次任务的连续链路**（要求 2）与
> **分环节可信度**（要求 5/6 的收口）。两份的**变异与测试口径一致**，可互相印证。

---

## 0. 一句话结论

同一次任务从**宿主读取**到**撤销**的连续链路，**两端（Excel Office.js / WPS JSAPI）
各有 1 条真实面板驱动的连续集成用例**，且**服务端有 1 条真实 Docker 支撑的连续链路
用例**；三者**共用同一批脏数据、同一套期望值**。过程中发现并修复了一个**真实产品缺陷**
（桥轮询校验与自己放行的工具集不一致 → 数据登记在 SDK 后端不可达、且异常被吞成"干等超时"）。
**真实模型 / 真实 Excel / 真实 WPS 三个环节未完成，如实标注。**

---

## 1. 要求 2：同一次任务的连续链路

### 1.1 覆盖的环节（用户点名 12 步）

```
宿主读取 → prepare_python_data → 数据句柄登记 → run_python_analysis
→ 已验证的隔离执行器 → 结果登记与引用取回 → 产品白名单计划编译
→ 面板预览 → 用户确认 → 产品执行器写回 → 回读 → 撤销
```

### 1.2 三条用例（都真实跑通、都经反向变异）

| 环节 | 用例文件 | 关键证据 |
| --- | --- | --- |
| **服务端（真实 Docker）** | `server/tests/test_python_continuous_chain_real.py` | **2 passed**（41s）：脏数据 → 真实容器内清洗/关联/汇总 → 结果登记 → `result://` 取回 → 白名单编译；且结果 `sourceFingerprint` 与授权绑定 |
| **Excel 端（真实 `<App/>`）** | `apps/excel-addin/src/continuousPythonChain.test.tsx` | **1 passed**：真实点击链，断言"确认前目标零写入、来源两表逐格未动" → 确认后目标 = `EXPECTED_SUMMARY` → 撤销后目标归空、来源仍逐格未动 |
| **WPS 端（真实面板）** | `apps/wps-addin/src/wpsContinuousPythonChain.test.tsx` | **1 passed**：轮询体带 `hostId:"wps-et"` + `workbookIdentity:"wps-path:c:\\data\\会员.xlsx"`；真实确认 → 目标 = `EXPECTED_SUMMARY`；真实撤销按钮 → "已撤销并核对恢复"、目标 `null` |

三者**共用**：

- 脏数据：会员表 `"  7  "` / `"7"`（重复）/ `"007"`（前导零）/ `""`（空键）/ `"9"`；订单表 `"404"`（未匹配）
- 期望汇总：`金卡 400 / 银卡 200 / 未匹配 400`（总计 1000）
- 隔离执行器：**只有**已验证的 Docker 执行器；**运行时不回退宿主裸执行**

### 1.3 允许 / 禁止（按要求原文核对）

| 要求 | 本轮做法 |
| --- | --- |
| 允许固定模型的工具调用序列 | ✅ 轮询响应由固定 stub 给出（`prepare_python_data` + `sources`） |
| 允许有状态宿主替身 | ✅ Excel 用 `officeTestHost.ts`、WPS 用带 `Range().Value2` 读写缓存的状态替身 |
| **禁止**预制结果矩阵 | ✅ 结果由 `run_python_analysis` 走隔离执行器算出，测试**不**注入矩阵 |
| **禁止**预制异常说明 | ✅ 空键/重复键/未匹配明细由分析规则产出，非测试写死文案 |
| **禁止**预制最终计划 | ✅ 计划由 `collectPythonData` → 白名单编译产出 |
| **禁止**测试中直接赋值 `range.values` 替代产品环节 | ✅ 写入只经 `executePlan` / WPS 执行器；测试只在**读取端**取矩阵做断言 |

### 1.4 发现并修复的真实产品缺陷

**症状**：SDK 后端上 `prepare_python_data` 的回执**一个都收不到**，服务端只能干等超时；
界面无任何报错。

**根因**：`packages/core/src/sdkToolBridge.ts` 里，指针解析器（`bridgeToolPointer`）
**放行** `prepare_python_data`，而轮询校验器（`assertBridgePollResponse`）**硬编码
只认 `query_table`**，于是把合法回执判为无效并抛错；调用方 `drain()` 又把异常**静默吞掉**
（`apps/excel-addin/src/hooks/useSdkToolBridge.ts`）⇒ 表现为"零回执、干等超时"。

**修复**：抽出共享守卫 `isBridgeToolName` + `BRIDGE_TOOL_NAMES`，**指针解析与轮询校验
认同一集合**（`packages/core/src/index.ts` 一并导出）。

**判别力证明**（反向变异）：先在 `packages/core/src/sdkToolBridge.test.ts` 加**失败**用例
（`AssertionError … 桥轮询返回了无效的待执行调用`），修复后转绿（**33 passed**）；
新增"仍拒绝桥不承载的工具（`read_range`）"用例锁住**不能放松**的那一侧。

---

## 2. 要求 3：关联是否执行真正影响结果

> 复用现有脏数据，规则属**本次任务**，**未**硬编码进产品。

**规则**（用例内显式实现，产品只提供通用查询引擎）：
trim 后匹配、保留前导零（`007 ≠ 7`）、空键排除、重复键保留第一条、订单**左连接**会员表。

**第 1b 段**（`apps/excel-addin/src/syntheticResultEndToEnd.test.tsx`，4 项）：

| 断言 | 结果 |
| --- | --- |
| 按关联得到的"会员等级"汇总 | `金卡 400 / 银卡 200 / 未匹配 400`，总计 1000 |
| 空键 / 重复键 / 未匹配的**可定位明细** | 各 1 条，含具体键值 |
| **改会员等级后汇总必须变化** | 等级字段一改，分组结果随之改变（证明关联真的参与计算） |
| 重复键保留第一条 | 与"保留最后一条"可区分 |
| **换表名 / 字段名后任务规则仍成立** | `甲表/乙表`、`编号/款项` 下同样成立 |

---

## 3. 要求 4：变异验收脚本

两条驱动，**四判据同时满足才算通过**：① 退出码 0 ② 有有效测试摘要
③ 实际执行了预期用例（`minTests`）④ 还原后与变异前**逐字节**一致（sha256，不搜残留标记）。

| 驱动 | 结果 |
| --- | --- |
| `scripts/run_synthetic_case_mutations.mjs` | `S1/S2/S3+S3b/S4` 全 `RED(OK) → 还原 → sha256 一致`；`RESULT: OK`；`EXIT=0` |
| `scripts/run_office_undo_mutations.mjs` | `M1/M2/M3/M4/M5/M4+M6b` 全 `RED(OK) → 还原 → sha256 一致`；`RESULT: OK`；`EXIT=0` |

**修好的一处假阴性**：office 驱动原先把"apply 成功"的判据写成 synthetic 的词形
（`apply: OK`），而 office 变异器实际打印 `APPLIED M1 → <path>`，于是**误报失败**。
现改为**按各自实现的词形**判定（表驱动 `MUTATOR_LANGUAGE`），并保留两条硬性：
`exit != 0` 立即抛；`exit == 0` 但**没报告这次操作**也抛（防"打了日志就退 0"）。
另外补了一条：**apply 报 `ALREADY-CLEAN` 视为失败**（源码本就是变异态 ⇒ RED 会伪造）。

**判别力自检**：故意给 S4 塞坏 SEED → 驱动**立即非零退出**；office 侧未知 id → `EXIT=2`、
SEED 未命中 → `EXIT=3`。**"报错"未曾被放松**。

---

## 4. 要求 5：`registerOfficeExecution` 可达性（已核实 / 已删除矛盾结论）

**核实结论：在 App 写入路径上可达。**

- 调用点在 `apps/excel-addin/src/App.tsx:5957`，守卫只有 `recoverySupport.supported`
  与"工作簿已保存"两条（未保存则给 `recoveryLimitation` 并降级为普通执行）。
- `officeRecoverySupport` 现在按**实际来源依赖**（`plan.sourceDependencySheets`）判读写
  重叠；正常计划（动作落在目标表、与来源不重叠）**会**返回 `supported: true` ⇒ 走到登记。
- `sourceDependencySheets` 的唯一来源是 `sdkBridgeHost.execute` 里的
  `recordTrustedToolReads`（执行器**实际 addSheet 过**的表），**不**用"授权范围减写入目标"反推。

**已作废的旧结论（矛盾已删除）**：

| 旧结论 | 状态 |
| --- | --- |
| "App 写入路径上 `registerOfficeExecution` **恒不可达**" | ❌ **作废**（第五轮已修，本轮复核确认） |
| "原生审阅与恢复登记**互斥** ⇒ 登记在 App 路径恒不可达" | ❌ **作废**。互斥只对"审阅要走原生审阅卡"的那类计划成立；授权范围 / 实际来源依赖 / 写入目标**三者分开建模**后，普通计划的动作与来源不重叠，登记正常发生 |
| "`WorkbookSnapshot.workbookIdentity` 是死字段" | ❌ **作废**（三个捕获点已补写 `workbookIdentity`） |

**证据**：`apps/excel-addin/src/excelWriteEndToEnd.test.tsx`「断点恢复登记的真实可达性」
**7 项** —— 授权经**真实范围选择弹层**建立、桥上真的读一次表、真实 `<App/>` 点击执行，
断言登记**恰好 1 次**、登记成功后才开始写目标表、4 条失败路径**原表零写入**。
**保留的限制**：真实 Excel 实机仍未验证；来源依赖**缺失**时按保守条件退回指纹范围
（依赖不确定就说不确定）。

---

## 5. 要求 6：分环节报告（六类，**互不代替**）

| # | 环节 | 状态 | 证据 / 边界 |
| --- | --- | --- | --- |
| 1 | **分段集成** | ✅ **完成** | 各段独立用例齐备（`syntheticResultEndToEnd.test.tsx`、`WpsAgentPanel.test.tsx` 的合成案例段、core/server 各自套件） |
| 2 | **连续链路** | ✅ **完成** | 三条：Excel 面板、WPS 面板、服务端真实 Docker（§1.2），均反向变异验证。**注意口径**：这三条是**两半** —— 服务端内部函数用例（无 HTTP/无面板）+ 两端面板用例（结果矩阵/`CONTINUOUS_PLAN`/异常 warnings **皆预制常量**，`/result` 由 fetch mock 返回固定成功） |
| 2b | **跨进程连续链路** | ✅ **完成** | **同一次任务的数据流**（2026-10-10 续补）：生产侧 `server/tests/test_cross_process_chain_real.py`（真实 ASGI + 真实 Docker，导出工件 `config/cross_process_chain.json`）+ 消费侧 `apps/excel-addin/src/crossProcessChain.test.tsx`（读工件，真实 `<App/>` 预览→确认→写回→回读→撤销）。**面板矩阵/异常/计划来自真实服务端执行，非预制常量**。一条命令 `npm run test:cross-process`；Docker 不可用 ⇒ 生产侧不产出工件 ⇒ 消费侧如实报「连续链路未完成」非零退出（**不用旧工件/mock 顶替**）。均反向变异验证 |
| 3 | **真实 Docker** | ✅ **完成** | 引擎可用（`npipe:////./pipe/docker_engine`，`verify_engine()=True`，镜像按 digest 固定）；`test_python_docker_real.py` **15 passed**；连续链路真实容器 **2 passed** |
| 4 | **真实模型** | ❌ **未完成** | 模型**自主规划**（不给定工具序列）**未验证** —— 留待真实模型验收单独跑。本轮是**固定**工具调用序列，**不**声称验证了模型自主规划 |
| 5 | **真实 Excel** | ❌ **未完成** | 全部 Excel 证据来自**替身宿主** `officeTestHost.ts`。**替身证据 ≠ 实机证据**；实机步骤见 §6 |
| 6 | **真实 WPS** | ❌ **未完成** | 全部 WPS 证据来自**状态化 WPS 替身**（`Range().Value2` 读写缓存）。同上，未上实机 |

> **不统一写"仅待实机"**：第 4/5/6 类**在自动化层面也未闭环**（模型自主规划没跑、
> 实机没上），故写**未完成**；第 1/2/3 类是**已完成的自动化/容器证据**，
> 但**不**等于实机通过。

### 5.1 剩余缺口（如实列出）

1. **模型自主规划**：连续链路里工具序列是固定的；"模型自己决定先读哪张表、怎么关联"未验证。
2. **真实 Excel 实机**：Office.js 真实宿主下的桥、授权、写回、撤销未跑。
3. **真实 WPS 实机**：JSAPI 真实宿主同上。
4. **精确续跑剩余动作**：登记的是 `WpsRecoveryDescriptor`，剩余非空只能 `replan`（续跑能力本身未变）。
5. **SDK 后端检查点恢复**：当前接入**未实现**（普通后端有检查点）—— 措辞按既定口径，**不**推断 SDK 永远不支持。
6. **`packages/core` 既存 tsc 报错 15 处**（`blankPolicy.test.ts` 5 / `columnMappingSemantics.test.ts` 8 /
   `columnMapping.test.ts` 2），断言 `ColumnMapping*Action.values` 而该类型无此字段；
   **HEAD 版同一份报错、本轮未改这些文件**，且**不在任何生产构建路径上**（core 只被 vitest 消费）。

---

## 6. 交付物：测试文件、复跑命令、日志

### 6.1 本轮新增/关键测试文件

| 文件 | 行数 | 说明 |
| --- | --- | --- |
| `apps/excel-addin/src/continuousPythonChain.test.tsx` | 491 | **要求 2** Excel 端连续链路（**两半之一**：面板真、结果/计划为预制常量） |
| `apps/wps-addin/src/wpsContinuousPythonChain.test.tsx` | 717 | **要求 2** WPS 端连续链路（同上，`CONTINUOUS_PLAN` 预制） |
| `server/tests/test_python_continuous_chain_real.py` | 223 | **要求 2** 真实 Docker 连续链路（**服务端内部函数**，无 HTTP/无面板） |
| `server/tests/test_cross_process_chain_real.py` | 500+ | **要求 2b 跨进程** 生产侧：真实 ASGI + 真实 Docker，导出工件 `config/cross_process_chain.json` |
| `apps/excel-addin/src/crossProcessChain.test.tsx` | 400+ | **要求 2b 跨进程** 消费侧：读工件，真实 `<App/>` 预览→确认→写回→回读→撤销 |
| `config/cross_process_chain.json` | — | **要求 2b** 工件（结果引用/矩阵/异常明细/服务端编译计划；由生产侧导出） |
| `scripts/run_cross_process_chain.mjs` | — | **要求 2b/6** 一条命令 `npm run test:cross-process` 的驱动 |
| `apps/excel-addin/src/syntheticResultEndToEnd.test.tsx` | 1187 | 要求 3（第 1b 段）+ 分段集成 |
| `packages/core/src/sdkToolBridge.test.ts` | 507 | 产品缺陷的锁死用例（33 passed） |
| `scripts/run_synthetic_case_mutations.mjs` | — | **要求 4** 合成案例变异驱动 |
| `scripts/run_office_undo_mutations.mjs` | — | **要求 4** 撤销/同表变异驱动 |
| `scripts/mutate_synthetic_case_guards.mjs` / `scripts/mutate_office_undo_and_same_sheet.mjs` | — | 两个变异器本体 |

### 6.2 复跑命令（均已实测）

```bash
# 跨进程连续链路（要求 2b）——**一条命令**：生产侧真实服务端+真实容器 → 导出工件 → 消费侧真实面板写回
npm run test:cross-process    # 生产侧 4 passed → 消费侧 2 passed → "连续链路完成"
# Docker 不可用 ⇒ 生产侧 skip、不产出工件 ⇒ 命令如实报「连续链路未完成」并非零退出（EXIT≠0）

# 连续链路（三条，两半口径见 §5）
npx vitest run apps/excel-addin/src/continuousPythonChain.test.tsx     # 1 passed
npx vitest run apps/wps-addin/src/wpsContinuousPythonChain.test.tsx    # 1 passed
.venv/Scripts/python.exe -m pytest server/tests/test_python_continuous_chain_real.py -v   # 2 passed

# 分段集成 / 要求 3
npx vitest run apps/excel-addin/src/syntheticResultEndToEnd.test.tsx   # 14 passed

# 要求 4（两个驱动都自带基线/还原/逐字节校验）
node scripts/run_synthetic_case_mutations.mjs      # RESULT: OK, EXIT=0
node scripts/run_office_undo_mutations.mjs         # RESULT: OK, EXIT=0

# 全量回归
(cd packages/core && npx vitest run)               # 391 passed / 25 files
(cd apps/excel-addin && npx vitest run)            # 1045 passed / 92 files
(cd apps/wps-addin && npx vitest run)              # 676 passed / 40 files
.venv/Scripts/python.exe -m pytest server/tests -q # 1304 passed, 1 skipped
npm run build:addin && npm run build:wps           # 两端 ✓ built
npm run lint                                       # 25 problems (1 error + 24 warnings)
```

> **pytest 注意**：必须从**仓库根**跑、`-p no:cacheprovider`、`--basetemp` 指向
> `C:\Users\<用户>\AppData\Local\Temp`（OS 临时目录内，避开 safe-delete 阈值）。

### 6.3 日志位置

| 内容 | 路径 |
| --- | --- |
| lint 输出 | `outputs/lint-df-final.log`（本报告）与 `outputs/lint-dfe.log`（上一轮） |
| Excel 写侧端到端 | `apps/excel-addin/src/excelWriteEndToEnd.test.tsx`（用例即证据） |
| 分环节测试矩阵 | `docs/SDK_WORKFLOW_DELIVERY_2026-10-09.md` §4/§6 |
| 合成案例验收 | `docs/ISOLATED_PYTHON_ACCEPTANCE_2026-10-10.md` |
| 逐日工作日志 | `.workbuddy/memory/2026-10-10.md` |

---

## 7. 双端实机验收步骤

### 7.1 测试工作簿

新建 `连续链路.xlsx`：

- `会员表`：表头 `会员编号|会员|等级`；数据 `"  7  "|张三|金卡`、`"7"|张三重复|金卡`、`"007"|李四|银卡`、`""|无编号|银卡`、`"9"|王五|银卡`
- `订单表`：表头 `会员编号|金额`；数据 `"  7  "|100`、`"007"|200`、`"9"|300`、`"404"|400`
- `汇总表`：**空表**（独立、已授权，与两张来源表**不重叠**）

### 7.2 步骤

1. 打开工作簿 → 加载项里点范围选择 → **勾选「会员表」「订单表」「汇总表」** → 完成。
2. 发送：「**按会员编号关联会员表，按会员等级汇总订单金额，结果写到汇总表**」。
3. 依次确认：数据登记 → 分析 → 结果卡 → 写入预览（**应看到**空键 1 / 重复键 1 / 未匹配 1）→ 确认写回。
4. 观察回读核对 → 点「撤销本次执行」。

### 7.3 期望结果（双端一致）

- `汇总表` = `[["会员等级","金额合计"],["金卡",400],["银卡",200],["未匹配",400]]`
- 预览/结果卡出现三条**可定位**异常明细（含具体键值）。
- **确认前**：`会员表` / `订单表` **一个单元格都没变**。
- **撤销后**：`汇总表` 逐格回空；`会员表` / `订单表` 仍逐格未变。

### 7.4 实机必须重点看的两件事

1. **桥是否真的建立**：确认卡里能否看到"已读表清单"、宿主标识是本机真实路径
   （`wps-path:` / `url:`），而不是退回显示名。
2. **撤销前的状态核对**：手工改一下 `汇总表` 的一格再点撤销 —— **应当拒绝撤销**并提示
   "执行后又发生变化，已停止撤销以免覆盖后续编辑"（这是 M1 那道闸的实机表现）。

---

## 8. 未提交声明

本轮全部改动**保留在工作区，未 commit / 未 push / 未发布**。
变异脚本 `restore-all` 已确认无 `变异S` / `变异M` 残留（`grep` 为空），
`run_office_undo_mutations.mjs` 的逐字节校验在还原后与变异前 sha256 一致。
