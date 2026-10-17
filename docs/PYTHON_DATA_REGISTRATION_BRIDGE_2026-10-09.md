# 隔离 Python 数据登记桥（`prepare_python_data`）—— 交付报告（2026-10-09 续）

## 0. 一句话结论

**服务端、共享层与两端宿主已把「数据登记桥工具」接通**：模型在一次 SDK `query()` 内可调用
`prepare_python_data`，服务端按任务绑定**从可信侧反查**来源范围、经本机工具桥让宿主读**已授权**表、
把原始值矩阵登记进易失 `PythonDatasetRegistry` 并建立 `coordinator` 绑定，回执只给模型
`dataset://` 句柄字符串（**绝不含值、绝不含 fingerprint**）；随后 `run_python_analysis` 能用这些句柄
在（环境就绪时的）隔离执行器里跑分析。句柄可信性、源指纹闸、folder 拒绝、形状/类型/字节上界、出站
投影全部在可信侧失败关闭。

> 本报告只覆盖「B/C：可信任务与数据接线」这一段的**数据登记桥**。真实 Docker 隔离执行仍**环境阻断**
> （引擎未运行，`_runner_status()=None` 正确 fail-closed），实机链路一律标**未验证**，不与自动化证据混谈。

## 1. 修改内容

### 1.1 服务端

| 文件 | 改动 |
| --- | --- |
| `server/app/models.py` | 新增 `BridgeDatasetReceipt`（`sourceIndex` + `values` 原始矩阵）；`BridgeToolReceipt` 增 `datasets` 字段，与 `result` / `errorCode` 三者互斥（`validate_one_of`） |
| `server/app/task_bridge.py` | `deliver` 增 `datasets` 形参并按 `pending.tool` 分流；新增 `_deliver_python_datasets`（源指纹闸 → 登记 → 绑定 → 返回 `dataset://` 句柄）、`_build_python_scope`（从 binding 派生 task/host/identityHash/authorizationVersion，只有 `sources` 来自模型参数）、`_bridge_origin(host_id)`（桥的稳定 origin 键，保证 `coordinator.bind` 与 `registry.create` 同源） |
| `server/app/agent_tools.py` | 新增 `PREPARE_PYTHON_DATA_TOOL`、`prepare_python_data_tool_definition()`、`project_prepare_python_data()`（只出站句柄列表 + 固定状态，绝不出值/fingerprint），登记进 `LOCAL_NON_TERMINAL_PROJECTIONS` |
| `server/app/claude_sdk_agent.py` | `_APPROVED_NON_TERMINAL_TOOLS` 增 `prepare_python_data`；`python_execution_enabled` 门控下追加工具定义；`handle()` 分派 `prepare_python_data` → `session.request` 走桥 |
| `server/app/main.py` | `sdk_tool_bridge_receipt` 透传 `payload.receipt.datasets` |

### 1.2 共享层（`packages/core`）

| 文件 | 改动 |
| --- | --- |
| `packages/core/src/sdkToolBridge.ts` | `BridgeToolName` 增 `prepare_python_data`；`bridgeRequestedSheetNames` 对 prepare **从 `sources[].sheet` 枚举**（不是整份授权范围，否则越权来源测不出来）；`bridgeToolPointer` 接受 prepare |

### 1.3 双端宿主

| 文件 | 改动 |
| --- | --- |
| `apps/excel-addin/src/contracts.ts` | `SdkBridgeToolName` 增 prepare；`SdkBridgeDatasetReceipt`；`SdkBridgeToolReceipt` 增 `datasets`；`parseBridgeToolPointer` 接受 prepare |
| `apps/excel-addin/src/dataTools.ts` | 新增 `executePreparePythonData(sources, allowedSheetNames, options)`：按来源 A1 区域读**原始值矩阵**（不做表头识别/分组），范围越权 `SCOPE_VIOLATION` |
| `apps/excel-addin/src/sdkBridgeHost.ts` | `execute` 按 `call.tool` 分流 prepare → `prepareData`（可注入依赖）；`receipt` 对 prepare 回 `datasets` |
| `apps/wps-addin/src/wpsSdkBridgeHost.ts` | `execute` 分流 prepare → 读整表 → `sliceWpsPythonSources` 按 A1 切片；新增 `asWpsPreparePythonSources`、`parseA1Range`、`sliceWpsPythonSources` |

## 2. 关键口径（安全边界）

- **句柄不是授权凭据**：模型只拿 `dataset://` 字符串；`execute_python_analysis` 按 `task_id` 查
  `coordinator` binding，从句柄反查 `PythonDatasetReceipt` 拿 fingerprint，句柄不在登记里就 `rejected`。
- **来源范围只来自可信侧**：`_build_python_scope` 的 `taskId`/`hostId`/`workbookIdentityHash`/
  `authorizationVersion` 全部取自 binding（模型拿不到）；只有 `sources`（sheet+range）来自模型参数，且每张
  sheet 都要落在授权范围内（宿主 `verify` 端口的 `bridgeRequestedSheetNames` 按 `sources[].sheet` 枚举核对）。
- **源指纹闸**：回执来源与绑定时不一致 → `source_changed`，与 `query_table` 同一道闸。
- **folder 拒绝**：folder 宿主没有隔离 Python 读表通道，`_build_python_scope` 直接返回 None → `invalid_arguments`。
  （该闸与 `PythonDataScope.hostId` 的 `Literal["excel-office","wps-et"]` **冗余**，见 §3 变异结论。）
- **出站投影白名单**：`project_prepare_python_data` 只产出 `{tool, status, datasetHandles}`；`values`/`fingerprint`/
  任何自由文本一律不出现。

## 3. 测试与反向变异

### 3.1 新增测试

| 套件 | 项 | 覆盖 |
| --- | --- | --- |
| `server/tests/test_prepare_python_data_bridge.py` | 9 | 闭环（句柄被 `run_python_analysis` 消费 → 本地结果 `[[10]]`）、源变化拒绝、folder 拒绝、空 sources、缺 datasets、形状不符 → `result_too_large`、投影不含值、未知错误映射 internal、绑定建立 |
| `apps/excel-addin/src/sdkBridgeHost.test.ts` | +2 | datasets 回执（`result`/`errorCode` 均 undefined）、越权来源不读宿主只回固定错误码 |
| `apps/wps-addin/src/wpsSdkBridgeHost.test.ts` | +2 | A1 切片原始值回 datasets、越权来源固定错误码 |

### 3.2 反向变异结论

| 变异 | 判别点 | 结果 |
| --- | --- | --- |
| M1 服务端源指纹闸拆除 | `_deliver_python_datasets` 的 `source_fingerprint` 比较 | **RED(OK)** |
| M3 投影值白名单拆除 | `project_prepare_python_data` 的显式白名单重建 | **RED(OK)** |
| M4 core prepare 来源枚举退化为整份授权 | `bridgeRequestedSheetNames` 的 `sources[].sheet` 枚举 | **RED(OK)** |
| M2 服务端 folder 宿主闸拆除 | `_build_python_scope` 的 `host_id not in (...)` | **GREEN(NG)** —— 与 `PythonDataScope.hostId` 的 `Literal` 约束**冗余**，非独立判别点。显式闸保留作早失败+自文档化，但不声称有独立判别力 |

## 4. 五类证据

| 类别 | 结论 |
| --- | --- |
| **通过** | 服务端 9 项数据登记桥测试、Excel/WPS 各 +2 项宿主测试、全量回归（服务端 1241/22 skipped、Excel 89 文件/995、WPS 38 文件/628、core 345）全绿；两端 `tsc --noEmit` EXIT=0、`compileall` EXIT=0；反向变异 M1/M3/M4 三处 RED/还原 GREEN |
| **失败** | 无（M2 folder 闸为冗余层，非缺陷，见上） |
| **环境阻断** | 真实 Docker 隔离执行仍环境阻断：引擎未运行，`_runner_status()=None` 正确 fail-closed，`test_python_docker_real.py` 整体 skip |
| **实机待验** | WPS `prepare_python_data` 用「读整表 + A1 切片」实现、Excel 端走 Office.js 原始区域读，**均未经真实宿主实机验证**；两端「数据登记 → 隔离执行 → 写回 → 回读 → 撤销」的真实实机链路**未取得** |
| **未验证** | 真实模型 / 真实 Excel / 真实 WPS 三类端到端；真实 Docker 沙箱内的隔离/限额/清理证据 |

## 5. 后续待办（未完成，如实保留）

- **真实隔离执行验证**：启动 Docker 引擎后跑 `test_python_docker_real.py`（核验隔离/超时/取消/资源限额/不可用拒绝）。
- **WPS 单表审阅限制**：`stageWpsNativeComparison` 目前限制单张授权表，需支持「两表来源、单表目标」。
- **通用验收案例**：合成数据覆盖清洗关联键、关联、分组汇总、重复键/未匹配明细、预览确认写回、回读、撤销/恢复。
- **两个锁死 `product_not_ready` 的测试**需复核语义（旧 `execute_python_analysis` 保留作失败关闭对照）。
- **文档**：README / ARCHITECTURE / DEVELOPMENT / 计划表同步本里程碑。

> 指导：`docs/SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md` §3（数据与授权分离）；工程纪律见 `AGENTS.md`。
