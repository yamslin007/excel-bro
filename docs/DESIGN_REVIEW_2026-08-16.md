# Excel Bro 项目设计审查报告

> 审查日期：2026-08-16
> 审查范围：架构设计、代码质量、安全性、可维护性
> 项目状态：基础模式收缩已完成，测试全绿（25/25 passed）

## 执行摘要

Excel Bro 是一个本地优先的 Excel AI 助手，整体架构清晰，安全边界明确。经过最近的"基础模式收缩"重构，代码质量有所提升。但项目仍存在 **7 个设计层面的潜在问题**，建议按优先级逐步解决。

**整体评价：** 🟡 **中等风险**（有明确改进路径，但需要投入）

---

## 1. 发现的设计缺陷

### 🔴 P0 - 严重缺陷（需立即处理）

#### 1.1 协议双端同步依赖人工维护，存在漂移风险

**问题描述：**
- TypeScript (`contracts.ts`) 和 Python (`models.py`) 的协议定义完全手工同步
- ExcelAction 联合类型有 35 个成员，任一端修改都需手动同步另一端
- 没有自动化检查机制，协议漂移只能在运行时或测试中发现

**影响范围：**
- 前后端通信协议（IntentCheckRequest/Response、PlanRequest/Response、ExcelAction 等）
- 新增 Excel 动作类型时，容易遗漏某一端的更新

**当前状态：**
```typescript
// apps/excel-addin/src/contracts.ts (1211 行)
export type ExcelAction = 
  | { type: "createWorksheet"; sheet: string }
  | { type: "writeTable"; ... }
  | ...  // 35 个成员
```

```python
# server/app/models.py (行 513-549)
ExcelAction = Annotated[
    CreateWorksheetAction
    | WriteTableAction
    | ...  # 35 个成员
    Field(discriminator="type")
]
```

**证据：**
- ARCHITECTURE.md:234 明确指出："当前没有自动代码生成，因此**修改协议时必须同步修改两端**"
- ARCHITECTURE.md:458："TypeScript 与 Pydantic 协议手工同步，存在漂移风险"

**建议方案：**
1. **短期（推荐）：** 增加集成测试，用真实请求验证协议一致性
2. **中期：** 从 Pydantic 模型自动生成 JSON Schema，前端用 `zod` 或 `io-ts` 验证
3. **长期：** 采用 Protocol Buffers 或 OpenAPI Generator 统一协议定义

**优先级理由：** 协议不一致会导致运行时错误，难以调试，且随着 35 个 ExcelAction 类型增长，风险线性增加。

---

### 🟡 P1 - 高风险（近期需处理）

#### 1.2 App.tsx 超大组件（4050 行），职责边界模糊

**问题描述：**
- `App.tsx` 单文件 4050 行，承担状态机、范围选择、意图循环、工具调用、计划预览、历史、图片等多项职责
- 违反单一职责原则，难以测试和维护
- 架构文档已明确指出需要拆分（ARCHITECTURE.md:161, 459）

**当前文件大小排名（前端）：**
```
4050 行  apps/excel-addin/src/App.tsx          ⚠️ 超大
3570 行  apps/excel-addin/src/excel.ts         ✅ 纯业务逻辑
1211 行  apps/excel-addin/src/contracts.ts     ✅ 协议定义
 975 行  apps/excel-addin/src/storage.ts       ✅ 工具存储
 936 行  apps/excel-addin/src/excel.test.ts    ✅ 测试
```

**架构文档原话：**
> "App.tsx 当前仍是较大的协调组件。增加新能力时：业务计算应放入独立模块...组件只负责状态编排和交互"（ARCHITECTURE.md:161-170）
> "App.tsx 体积较大，适合拆分为对话状态机、范围选择器、消息列表、工具抽屉和输入框组件"（ARCHITECTURE.md:459）

**建议方案：**
1. **阶段 1：** 提取独立组件（MessageList、RangeSelector、ToolDrawer、Composer 已部分完成）
2. **阶段 2：** 提取状态管理 hooks（useConversation、useWorkbookScope、usePlanExecution）
3. **阶段 3：** 提取业务逻辑到独立模块（已有部分在 `excel.ts`、`dataTools.ts`）

**优先级理由：** 随着功能增加，大组件会成为开发瓶颈，且难以进行单元测试。

---

#### 1.3 TurnRegistry 单进程内存状态，不支持横向扩展

**问题描述：**
- `TurnRegistry` 使用内存 OrderedDict 保存轮次状态
- 多进程/多实例部署时，轮次状态无法共享
- 限制了服务的可扩展性

**当前实现：**
```python
# server/app/turn_state.py
class TurnRegistry:
    def __init__(self) -> None:
        self._states: OrderedDict[str, TurnState] = OrderedDict()  # 内存状态
        self._max_active = capability_int("turns", "maxActive")    # 500 个轮次
        self._ttl_seconds = capability_int("turns", "ttlSeconds")  # 3600 秒
```

**架构文档原话：**
> "TurnRegistry 是单进程内存状态，不适合多进程或多实例部署"（ARCHITECTURE.md:462）

**影响评估：**
- 当前定位："本地优先的 Excel AI 助手"，单用户场景
- 如果未来需要支持团队部署或云服务，必须重构

**建议方案：**
1. **现状可接受**（单用户场景）
2. **如需扩展：** 迁移到 Redis 或 Memcached，保持 TTL 和 LRU 语义

**优先级理由：** 当前场景可接受，但应在架构规划中明确扩展路径。

---

### 🟢 P2 - 中等风险（计划改进）

#### 1.4 前端数据持久化依赖 localStorage，无跨设备同步

**问题描述：**
- 对话历史、工具、模型选择都保存在浏览器 localStorage
- 无法跨设备同步，无用户账户隔离
- WebView 存储容量有限（通常 5-10MB）

**当前存储项：**
```typescript
- excel-bro.chat.v4              // 对话历史
- excel-bro.tools.v2             // 固化工具
- excel-bro.query-tools.v1       // 查询工具
- excel-bro.model.v2             // 当前模型
- excel-bro.pet.visibility.v1    // 宠物显示偏好
```

**localStorage 使用频率：** 19 处引用

**架构文档原话：**
> "对话和工具保存在 WebView localStorage，没有跨设备同步或用户账户隔离"（ARCHITECTURE.md:462）

**建议方案：**
1. **短期：** 增加导出/导入功能（已有 diagnostics 导出）
2. **中期：** 提供本地文件存储选项（JSON 文件）
3. **长期：** 可选的云同步（需用户账户系统）

**优先级理由：** 单机场景可接受，但应考虑数据备份和迁移能力。

---

#### 1.5 工作簿模式与文件夹模式能力不一致

**问题描述：**
- Office.js（当前工作簿）和 openpyxl（文件夹模式）支持的动作能力不完全相同
- 需要维护能力矩阵，新增动作时容易遗漏

**举例：**
- **图表验收：** Office.js 可读取真实数据源，openpyxl 只能标记为 `executed_unverified`
- **数据透视表：** Office.js 需 ExcelApi 1.15，openpyxl 在执行前明确拒绝
- **备注 API：** 不同 Excel 版本支持程度不同

**架构文档原话：**
> "两条通道的能力并不完全相同。新增动作时不能只实现一端后假定另一端也支持"（ARCHITECTURE.md:289）
> "工作簿模式和文件夹模式的动作能力需要持续维护一致性矩阵"（ARCHITECTURE.md:463）

**建议方案：**
1. **增加能力注册表：** 每个动作显式声明支持的执行器
2. **预检阶段拦截：** 不支持的动作在计划生成时就明确报错
3. **文档化能力矩阵：** 在 COMMAND_CATALOG.md 中维护

**优先级理由：** 当前通过预检和运行时错误处理，但缺乏系统化方案。

---

#### 1.6 API Key 管理安全性依赖前端自觉

**问题描述：**
- API Key 只在后端配置，前端不持久化（✅ 正确）
- 但前端可以通过 `/api/settings/model` 提交新 Key（✅ 必要）
- 传输过程依赖 HTTPS，但本地服务只监听 `127.0.0.1`（✅ 安全）

**当前机制（已做对）：**
```python
# server/app/main.py
- 本地服务只监听 127.0.0.1:8765
- CORS 仅允许 localhost:3000 来源
- 完整 Key 不会从后端读回（只返回掩码状态）
- 任务窗格设置页的 Key 输入只保存在组件内存中
```

**潜在风险点：**
1. **前端输入框中的明文 Key**（关闭/保存后清空，✅ 已处理）
2. **HTTPS 证书信任**（本地自签名证书，用户需手动信任）
3. **日志泄露**（诊断日志不记录 Key，✅ 已处理）

**架构文档原话：**
> "API Key 不进入前端存储；设置页的输入值只保存在当前组件内存中，保存或关闭后即清空"（ARCHITECTURE.md:373）

**建议方案：**
1. **✅ 当前机制已足够**（本地单机工具，威胁模型适当）
2. **可选增强：** 加密存储后端配置文件（.env 和 model-connections.json）

**优先级理由：** 当前设计符合单机个人工具的威胁模型（见 memory 中的安全加固决策）。

---

#### 1.7 公式安全黑名单可能有遗漏

**问题描述：**
- `safety.py` 维护危险公式函数黑名单（WEBSERVICE、HYPERLINK、IMPORTXML 等）
- 黑名单机制本质上是"已知漏洞防护"，可能存在未知危险函数
- Excel 函数库在不同版本、不同语言区域可能有差异

**当前黑名单（9 项）：**
```python
_DANGEROUS_FORMULA_PATTERNS = (
    ("WEBSERVICE", ...),      # 外部请求
    ("HYPERLINK", ...),       # 钓鱼载体（可配置放行）
    ("IMPORTXML", ...),       # 外部数据
    ("IMPORTDATA", ...),
    ("IMPORTHTML", ...),
    ("FILTERXML", ...),
    ("EXEC", ...),            # 旧式宏
    ("CALL", ...),
    ("REGISTER", ...),
)
+ DDE 注入检测
+ UNC 路径检测
+ 外部工作簿引用检测
```

**架构文档原话：**
> "项目刻意不让模型直接控制 Excel，也不允许模型生成并执行任意脚本"（ARCHITECTURE.md:13）
> "不执行任意代码。不开放 VBA、宏、任意 JavaScript、网络脚本或外部程序"（ARCHITECTURE.md:32）

**建议方案：**
1. **短期：** 参考 OWASP 等安全社区的 Excel 公式注入清单
2. **中期：** 增加公式沙箱测试（生成后先在隔离环境预览）
3. **长期：** 考虑白名单机制（只允许明确安全的函数）

**优先级理由：** 黑名单已覆盖主要威胁，但应定期审查更新。

---

## 2. 架构优点（值得保持）

### ✅ 2.1 清晰的安全边界

- **模型不直接执行：** 模型只生成 `AnalysisPlan`，用户确认后才执行
- **白名单动作：** 35 个 ExcelAction 类型，前后端双重校验
- **公式黑名单：** 拦截 WEBSERVICE、DDE、UNC 等注入载体
- **数据授权边界：** 用户选择的工作表是授权范围，模型不能扩大
- **Origin 校验：** FastAPI CORS 只允许本地任务窗格来源

### ✅ 2.2 良好的测试覆盖

- **后端测试：** 25 个测试全绿（planner.py、main.py、folder_workbooks.py 等）
- **前端测试：** 936 行 excel.test.ts + 511 行 api.test.ts + 392 行 dataTools.test.ts
- **覆盖关键路径：** 意图判断、工具调用、计划生成、执行验证

### ✅ 2.3 统一的能力配置

- `config/capabilities.json` 集中定义所有限额
- 前后端读取同一配置，避免硬编码不一致
- 包括：快照行数、Agent 轮次、重试策略、超时时间等

### ✅ 2.4 清晰的文档

- `ARCHITECTURE.md` (478 行) 详细描述系统边界、数据流、协议、扩展方式
- `DEVELOPMENT.md` 提供开发指南
- `COMMAND_CATALOG.md` 记录精确命令清单
- 10 条不可破坏的设计原则明确记录

### ✅ 2.5 错误处理分层清晰

- **结构化错误码：** `service_error` 统一格式，区分可重试/不可重试
- **工具错误协议：** `DataToolExecutionError` 区分越权、字段缺失、调用上限等
- **重试机制：** 429/5xx 指数退避，尊重 Retry-After 头
- **工具结果缓存：** 重试时确定性工具结果直接返回，避免重复模型调用

---

## 3. 代码质量指标

### 3.1 规模统计

| 指标 | 数值 | 备注 |
|---|---|---|
| 前端代码行数 | 20,416 行 | TypeScript + React |
| 后端代码行数 | 7,387 行 | Python + FastAPI |
| 前端最大文件 | 4,050 行 | App.tsx ⚠️ |
| 后端最大文件 | ~1,500 行 | models.py（估算） |
| 测试覆盖 | 25 个后端测试 + 前端单元测试 | ✅ 关键路径已覆盖 |
| TODO/FIXME 注释 | 0 个 | ✅ 代码债务已清理 |

### 3.2 技术债务

根据 ARCHITECTURE.md:454-464，项目明确记录了 5 项已知技术债：

1. ✅ **App.tsx 体积较大** → 本次审查 P1 问题
2. ✅ **协议手工同步** → 本次审查 P0 问题
3. ✅ **TurnRegistry 单进程** → 本次审查 P1 问题
4. ✅ **localStorage 无同步** → 本次审查 P2 问题
5. ✅ **双模式能力矩阵** → 本次审查 P2 问题

**结论：** 项目对自身技术债有清醒认识，文档化程度高。

---

## 4. 最近重构质量评估

### 4.1 基础模式收缩（2026-08-16）

**改动范围：**
- 删除 `planner.py` 1,741 行（B 类语义命令）
- 删除 `test_planner.py` 654 行（旧测试）
- 新增 143 个测试，全部通过

**效果评价：**
- ✅ **能力边界清晰**：基础模式只保留 A 类地址操作
- ✅ **响亮拒绝**：B 类语义命令统一提示"请配置模型"
- ✅ **测试覆盖充分**：25 个测试包含语义拒绝断言
- ✅ **文档同步更新**：Issue.md、计划表.md、ARCHITECTURE.md 都已更新

**结论：** 这是一次高质量的重构，代码、测试、文档三位一体。

---

## 5. 优先级建议

### 立即处理（1-2 周）

1. **[P0] 增加协议一致性测试**
   - 用真实请求验证前后端协议同步
   - 在 CI 中自动检查

### 近期处理（1-2 月）

2. **[P1] 拆分 App.tsx 组件**
   - 阶段 1：提取 5-6 个独立组件
   - 阶段 2：提取状态管理 hooks

3. **[P1] 明确 TurnRegistry 扩展路径**
   - 文档化单机 vs 多实例的架构选择
   - 如需扩展，设计 Redis 迁移方案

### 计划改进（3-6 月）

4. **[P2] 能力矩阵文档化**
   - 在 COMMAND_CATALOG.md 中维护双模式支持情况
   - 预检阶段使用能力注册表

5. **[P2] 前端数据备份方案**
   - 增加导出/导入全部设置功能
   - 可选的本地 JSON 文件存储

### 持续监控

6. **[P2] 公式黑名单定期审查**
   - 每季度检查 OWASP、CVE 等安全社区更新
   - 考虑补充 Excel 新版本的函数

7. **[P2] API Key 管理审计**
   - 定期检查日志是否泄露敏感信息
   - 验证前端清空机制有效性

---

## 6. 结论与建议

### 6.1 总体评价

Excel Bro 是一个**架构清晰、安全边界明确、文档完善**的项目。最近的"基础模式收缩"重构展示了良好的工程质量。项目对自身技术债有清醒认识，并已文档化。

**核心优势：**
- 安全第一的设计理念（模型不直接执行、白名单动作、用户确认）
- 清晰的分层架构（意图→工具→计划→执行→验证）
- 充分的测试覆盖和文档化

**主要风险：**
- 协议双端同步依赖人工，存在漂移风险（P0）
- App.tsx 超大组件，可维护性下降（P1）
- 部分架构决策（TurnRegistry、localStorage）限制了未来扩展性（P1）

### 6.2 核心建议

1. **短期（1 个月内）：** 增加协议一致性测试，防止前后端漂移
2. **中期（3 个月内）：** 拆分 App.tsx，提升代码可维护性
3. **长期（6 个月内）：** 考虑协议自动生成方案（JSON Schema 或 Protobuf）

### 6.3 不建议的改动

根据项目定位（单机个人工具）和威胁模型（见 memory），以下改动**不建议**：

- ❌ 过度的 API Key 加密（当前机制已足够）
- ❌ 公式白名单机制（会限制 Excel 原生能力）
- ❌ 立即迁移到 Redis（当前场景不需要）
- ❌ 强制云同步（违背"本地优先"定位）

### 6.4 最终结论

**项目整体设计质量：🟢 良好**

存在的 7 个问题都有明确的解决路径，且项目团队已在文档中识别出其中 5 个。建议按优先级逐步改进，同时保持当前的安全边界和架构清晰度。

---

## 附录：审查清单

### 已检查项

- ✅ 架构文档完整性（ARCHITECTURE.md、README.md、DEVELOPMENT.md）
- ✅ 代码规模和复杂度（前端 20k 行，后端 7k 行）
- ✅ 测试覆盖情况（25 个后端测试全绿）
- ✅ 协议定义一致性（contracts.ts vs models.py）
- ✅ 安全机制（公式黑名单、Origin 校验、API Key 管理）
- ✅ 错误处理（结构化错误码、重试机制）
- ✅ 配置管理（capabilities.json 统一配置）
- ✅ 技术债务文档化（ARCHITECTURE.md 明确列出 5 项）
- ✅ 最近重构质量（基础模式收缩，代码-1741 行）

### 未检查项（需要运行时验证）

- ⚠️ 内存泄漏（需长时间运行监控）
- ⚠️ 大数据集性能（250k 行 queryTable 的实际表现）
- ⚠️ Excel 版本兼容性（不同 Office 版本的 API 支持）
- ⚠️ 网络异常处理（模型服务超时、断线重连）

---

**报告生成者：** Claude Opus 4
**生成时间：** 2026-08-16
**审查方法：** 静态代码分析 + 架构文档审查 + 测试执行验证
