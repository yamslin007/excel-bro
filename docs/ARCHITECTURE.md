# Excel Bro 架构文档

## 系统架构

### 整体架构

- 前端：React + TypeScript + Office.js
- 后端：Python FastAPI
- 通信：HTTP REST API
- 存储：localStorage（前端）+ JSON 文件（后端规则）

```text
┌───────────────────────────────────────────────────────┐
│                      Excel 任务窗格                     │
│                                                       │
│  React App.tsx                                        │
│    ├── 13 个 Custom Hooks                              │
│    ├── UI 组件                                         │
│    └── 工具模块                                        │
│                                                       │
│  Office.js ──> 当前 Excel 工作簿                       │
└──────────────────────┬────────────────────────────────┘
                       │ HTTPS / REST
┌──────────────────────▼────────────────────────────────┐
│                 FastAPI 本地服务                       │
│                                                       │
│  /api/turn  /api/intent  /api/plan                    │
│  /api/models  /api/settings/model                     │
│  /api/folder  /health                                 │
│                                                       │
│  Python 服务模块                                       │
│    ├── models.py / contracts.py                       │
│    ├── intent.py / planner.py / excel_agent.py        │
│    ├── folder_workbooks.py / folder_data.py           │
│    └── capabilities.py / safety.py                    │
└──────────────────────┬────────────────────────────────┘
                       │ OpenAI-compatible API
┌──────────────────────▼────────────────────────────────┐
│                  外部模型服务                          │
└───────────────────────────────────────────────────────┘
```

### 前端架构

#### 核心组件层级

```text
App.tsx (5946 行) - 业务协调层
├── 13 个 Custom Hooks - 状态管理层
│   ├── 状态管理类 (6 个)
│   │   ├── useConversation
│   │   ├── useUIState
│   │   ├── useExecutionApproval
│   │   ├── useScopeSelection
│   │   ├── useActivityProgress
│   │   └── useUndoSnapshot
│   ├── 资源管理类 (4 个)
│   │   ├── useImageAttachments
│   │   ├── useModelManagement
│   │   ├── useToolManagement
│   │   └── useWorkbookContext
│   └── 交互辅助类 (3 个)
│       ├── useSlashCommands
│       ├── useCopyFeedback
│       └── useServiceHealth
├── UI 组件（实际接线）
│   ├── PetCompanion / RuleManager / SlashCommandAutocomplete / ThemePanel
│   └── 其余界面（消息流、范围浮层、设置/工具/历史抽屉、输入区等）
│       内联渲染在 App.tsx
│   （曾把上述界面抽成独立展示组件文件，但从未被引入，已作为死代码删除）
└── 工具模块
    ├── conversation.ts
    ├── format.ts
    ├── intent.ts
    ├── range.ts
    ├── utils.ts
    ├── excel.ts
    ├── dataTools.ts
    └── storage.ts
```

#### Hooks 架构

Hook 之间不直接调用彼此的内部状态，主要通过 props 或回调通信：

```text
useServiceHealth ──refreshServiceHealth──> useModelManagement

useConversation ──setMessages──> App.tsx

useScopeSelection ──workbook/snapshot──> App.tsx

useActivityProgress ──onPersistLog──> App.tsx

useUndoSnapshot ──onAfterUndo/onMessage──> App.tsx
```

#### 数据流向

```text
[用户交互]
    ↓
[App.tsx 协调层]
    ↓
[Hooks 状态管理层]
    ↓
[API / Office.js]
    ↓
[FastAPI / 当前工作簿 / localStorage]
```

### 后端架构

`server/app/` 的主要模块：

```text
server/app/
├── main.py                 # FastAPI 应用、路由、CORS
├── models.py               # Pydantic 请求/响应模型
├── intent.py               # 意图判断和澄清
├── planner.py              # 基础模式与最终规划
├── excel_agent.py          # 模型工具循环
├── turn_state.py           # turnId 状态、TTL 和缓存
├── folder_workbooks.py     # 文件夹执行
├── folder_data.py          # 文件夹确定性查询
├── capabilities.py         # 共享能力配置
├── safety.py               # 安全与协议校验
├── rule_generator.py       # /function 公式生成主流程
├── formula_tools/          # /function 工具调用架构（schema/编译器/注册表/Prompt）
├── formula_tools_integration.py  # 工具调用路径与主流程的衔接
└── llm/                    # 模型调用、错误和重试
```

### /function 公式生成链路

`/function` 短链优先走「本地确定性跨表公式」（前端 `crossTableFormula.ts`，
不调用 AI）。覆盖不了时后端按两级架构生成：

1. **工具调用路径**（`formula_tools/` + `formula_tools_integration.py`）：
   模型只返回结构化函数调用 JSON（`FormulaToolCallResponse`），由代码编译器
   递归编译成公式文本，保证语法正确、字面量不丢失；编译结果与主流程一样经过
   `dangerous_formula` 安全检查。
2. **端到端兜底**（`rule_generator.py` 原有逻辑）：工具调用路径任何一步失败
   （模型未配置、JSON 无效、编译失败、安全检查拦截）都自动回退到原有
   「模型直接生成公式文本」方式，用户无感知。

协议不变：请求仍为 `GenerateFormulaRequest`，响应仍为
`GenerateFormulaResponse`（现代版 + 兼容版两条公式）。工具调用路径只生成一版
公式：`compatFormula` 与 `modernFormula` 相同，但 `compatExplanation` 会附加
「⚠️ 此公式使用了 Excel 365+ 函数，可能不兼容旧版本」警示，避免承诺虚假兼容；
真正的兼容版转换（XLOOKUP → VLOOKUP 等）留待后续实现。

### 通信协议

主要 API：

- `POST /api/turn`：统一轮次入口
- `POST /api/turn/stream`：SSE 流式规划
- `POST /api/intent`：兼容意图入口
- `POST /api/plan`：兼容规划入口
- `GET /api/models`：模型目录
- `GET /api/settings/model`：模型配置
- `PUT /api/settings/model`：更新模型配置
- `POST /api/settings/model/connections`：新增连接
- `DELETE /api/settings/model/connections`：删除连接
- `POST /api/formulas/generate`：/function 短链生成原生公式（可携带文件夹模式勾选的外部工作簿工作表上下文 `extraSheets`）
- `POST /api/folders/select`：选择并扫描文件夹
- `POST /api/folders/refresh`：刷新文件夹文件列表（保持 sessionId 不变）
- `GET /api/diagnostics`：诊断信息

前端协议定义在 `apps/excel-addin/src/contracts.ts`，后端协议镜像在 `server/app/models.py`。

## 关键设计决策

### 为什么选择 Custom Hooks？

- 状态和副作用能按业务域拆分，避免 `App.tsx` 继续膨胀。
- Hook 与 React 渲染模型天然集成，无需引入额外运行时。
- 单一 Hook 可独立测试，测试成本低。
- 组件只消费清晰的返回接口，降低耦合。

### 为什么不使用状态管理库？

- 当前状态主要是局部 UI、对话、工具和模型配置，跨组件共享需求有限。
- 引入 Redux/Zustand 等库会增加包体积、样板代码和学习成本。
- Custom Hooks + props 回调已经能覆盖现有协作场景。
- 后续如出现复杂全局状态，再评估轻量状态库。

### 为什么使用 localStorage？

- 本地优先，无需账户和跨设备同步。
- 保存对话、工具、模型选择和 UI 偏好已经足够。
- 对 API Key 只保存到后端本地配置，不进入前端 localStorage。
- 数据量较小，读写频率可接受。

## 性能考虑

### Bundle 大小

- 当前前端为单任务窗格应用，主要逻辑在 `taskpane` bundle。
- 优化策略：
  - 按需拆包
  - 避免重复依赖
  - 工具函数模块化，减少重复内联
  - 未来可评估 `React.lazy`

### 渲染性能

- 关键指标：任务窗格加载、规划流式渲染、本地查询进度。
- 优化点：
  - 保持 `useCallback` 依赖完整，减少无意义重建。
  - 对长消息列表按需渲染。
  - 进度更新使用函数式 state，避免闭包过期。
  - 清理定时器和事件监听，防止后台开销。

## 安全考虑

### 公式黑名单

- 本地 `formulaSafety.ts` 拦截危险公式。
- 服务端仍通过协议和动作白名单限制最终写入。

### API Key 管理

- 完整 Key 只保存在后端本地配置。
- 前端只接收掩码状态和连接 ID。
- 日志和诊断不记录完整 Key、提示词和原始数据行。

### Origin 校验

- FastAPI 只监听 `127.0.0.1`。
- CORS 只允许本地任务窗格来源。
- 当前工作簿模式通过 Office.js，不暴露任意文件访问。

## 扩展性

### 如何添加新功能

1. 判断属于语义决策还是确定性计算。
2. 协议层同步修改 `contracts.ts` 和 `models.py`。
3. 前端执行层实现 Office.js，后端决定是否支持 openpyxl。
4. 补充动作说明、风险判断、验收条件和测试。
5. 更新 `docs/ARCHITECTURE.md` 和 `docs/HOOKS_GUIDE.md`。

### 如何添加新 Hook

1. 在 `apps/excel-addin/src/hooks/` 新建文件。
2. 定义选项和返回类型。
3. 使用 `useCallback` / `useEffect` 并完整声明依赖。
4. 添加副作用 cleanup。
5. 编写单元测试。
6. 运行 `npm run test:addin` 和 `npm run build:addin`。
7. 更新 Hook 文档。

### 如何添加新命令

1. 在 `useSlashCommands` 或相关命令目录中注册命令。
2. 定义命令值、标签和描述。
3. 在 `App.tsx` 或对应组件中处理命令结果。
4. 补充命令过滤和自动补全测试。

## 参考资料

- P1 重构总结：`docs/P1_REFACTOR_SUMMARY.md`
- Hooks 使用指南：`docs/HOOKS_GUIDE.md`
- 设计审查报告：`docs/DESIGN_REVIEW_FINAL_SUMMARY.md`
