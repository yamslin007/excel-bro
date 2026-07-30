# Excel Bro 开发指南

这份文档用于让新开发者快速启动、定位代码并安全地扩展功能。系统设计原因和数据流参见 [ARCHITECTURE.md](./ARCHITECTURE.md)。

## 1. 环境要求

- Windows 10/11
- Microsoft Excel 桌面版
- Node.js 与 npm
- Python 3.9+
- PowerShell

Excel 加载项开发使用本地 HTTPS 证书。首次旁加载时 Windows 会请求信任证书。

## 2. 首次安装

在仓库根目录执行：

```powershell
npm install
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r server\requirements.txt
Copy-Item server\.env.example server\.env
```

如果 PowerShell 阻止虚拟环境脚本，可仅为当前终端放开：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.\.venv\Scripts\Activate.ps1
```

如果同时安装了 Conda，终端出现 `(.venv) (base)` 不影响项目运行。希望新终端不再自动进入 base，可执行：

```powershell
conda config --set auto_activate_base false
```

该设置只对之后新开的终端生效。

## 3. 模型配置

推荐直接在任务窗格顶部点击“添加模型”。如需使用旧环境变量方式，可编辑
`server/.env`：

```env
AI_BASE_URL=https://api.moonshot.cn/v1
AI_API_KEY=你的密钥
AI_MODEL=kimi-k3
AI_MODELS=
AI_VISION_MODELS=kimi-k3
AI_INTENT_TIMEOUT_SECONDS=60
AI_TIMEOUT_SECONDS=60
```

注意：

- `AI_MODEL` 必须是供应商实际接受的模型 ID。
- `AI_MODELS` 用英文逗号分隔可切换模型。
- 只有列入 `AI_VISION_MODELS` 的模型才能接收图片。
- 本地 Ollama、LM Studio 等无鉴权网关可以留空 `AI_API_KEY`。
- 新安装的配置模板不会预置服务商或模型；未配置时只显示基础模式和首次添加引导。
- 任务窗格顶部模型菜单同时提供选择、添加和管理入口。每个独立连接拥有自己的服务地址、模型 ID 和 Key，并写入配置文件同目录下的 `model-connections.json`。
- 设置页修改会立即生效，完整 Key 不会返回前端。`model-connections.json` 包含密钥，已被 Git 忽略，不能截图、记录日志或手动提交。
- “测试连接”会发送一个内容固定、输出极短的模型请求，不会读取或发送工作簿数据，也不会保存尚未确认的连接。
- 新增连接保存后会自动切换到新模型。编辑连接时，Key 留空会保留旧值；需要接入无鉴权本地服务时可显式勾选清除。
- 同一服务地址和模型 ID 不能重复添加；如果 `model-connections.json` 损坏，服务会明确报错并保留原文件，避免静默覆盖密钥。
- 手动修改 `.env` 中的服务地址、模型列表或其他参数后，仍需重启 FastAPI 服务。
- 不要把真实 Key 写入 `.env.example`、前端代码、截图、日志或提交记录。

## 4. 日常启动

建议使用三个终端。

### 终端 1：后端

```powershell
.\.venv\Scripts\Activate.ps1
npm run dev:server
```

健康检查：

```powershell
Invoke-RestMethod http://127.0.0.1:8765/health
Invoke-RestMethod http://127.0.0.1:8765/api/models
Invoke-RestMethod http://127.0.0.1:8765/api/settings/model
```

### 终端 2：任务窗格前端

```powershell
npm run dev:addin
```

前端地址是 `https://localhost:3000`。

### 终端 3：旁加载 Excel

```powershell
npm run start:excel
```

该命令会创建并打开一个临时测试工作簿，加载项只旁加载到该工作簿。普通新建
工作簿不会继承这次开发旁加载。停止旁加载：

```powershell
npm run stop:excel
```

修改 `manifest.xml` 后应停止再重新旁加载。普通 React/CSS 修改通常只需刷新任务窗格。
普通工作簿需要通过 Windows 共享加载项目录进行本地测试安装；正式发布使用
Microsoft 365 集中部署或 AppSource。

### 个人 Windows 安装包

执行 `npm run build:installer` 生成的安装包会自动创建
`\\<电脑名>\ExcelBroAddins` 本机共享并注册到 Office 可信加载项目录。使用实际
计算机名而不是 `localhost`，以兼容只枚举标准 UNC 服务器名的 Office 版本。创建共享
需要一次 UAC 管理员授权。安装后完全退出并重开 Excel，再依次选择
**开始 → 加载项 → 更多加载项 → 高级 → 共享文件夹**，选择 **Excel Bro**
并点击 **添加**。Office 要求用户完成这次首次确认，安装器不能代替点击。

卸载时还会请求一次 UAC 权限以删除本机共享，并清除可信目录、启动项、证书和
程序文件；`%LOCALAPPDATA%\Excel Bro` 中的个人模型配置默认保留。

## 5. 浏览器与 Excel 调试

### 浏览器模式

打开 `https://localhost:3000`：

- 使用 `src/demo.ts` 的演示工作簿；
- 适合调试布局、历史对话、模型选择和普通 API；
- 不会真实写入 Excel；
- 无法完整验证 Office.js 行为。

### Excel 模式

必须在 Excel 中验证：

- 当前工作表和选区读取；
- 多工作表顺序；
- Office.js 动作；
- 图表、透视表、图片等 API 兼容性；
- 写入后的真实验收；
- 任务窗格宽度和触屏交互。

开发者工具可从 Excel 加载项调试入口打开。前端错误优先查看 WebView 控制台，后端错误查看运行 FastAPI 的终端。

## 6. 常用命令

```powershell
# 前端测试
npm run test:addin

# 前端类型检查和生产构建
npm run build:addin

# 后端测试
npm run test:server

# 全量测试
npm test

# Python 语法编译检查
python -m compileall -q server/app

# 构建 Windows 本地安装包
npm run build:installer
```

提交前推荐：

```powershell
npm run test:addin
npm run build:addin
python -m pytest server/tests -q
python -m compileall -q server/app
```

## 7. 从哪里开始改

| 需求 | 首要文件 |
| --- | --- |
| 对话布局、按钮、输入框 | `apps/excel-addin/src/App.tsx`, `styles.css` |
| Excel 读取或写入 | `apps/excel-addin/src/excel.ts` |
| 本地筛选、统计、占比 | `apps/excel-addin/src/dataTools.ts` |
| 单元格类型标准化 | `apps/excel-addin/src/cellNormalization.ts` |
| 拆分工作表并聚合 | `apps/excel-addin/src/splitAggregate.ts` |
| API 请求与错误显示 | `apps/excel-addin/src/api.ts`, `server/app/main.py` |
| 需求确认机制 | `server/app/intent.py` |
| 模型配置、鉴权和 HTTP 适配 | `server/app/llm/` |
| 模型 Agent 工具 | `server/app/excel_agent.py` |
| 基础模式和模型选择 | `server/app/planner.py` |
| Excel 动作协议 | `contracts.ts`, `server/app/models.py` |
| 文件夹 Excel | `server/app/folder_workbooks.py`, `server/app/folder_data.py` |
| 固化工具 | `apps/excel-addin/src/storage.ts`, `deterministicTools.ts` |
| 本地诊断 | `apps/excel-addin/src/diagnostics.ts`, `server/app/main.py` |
| 限额 | `config/capabilities.json` |
| 功能区按钮与地址 | `apps/excel-addin/manifest.xml` |

## 8. 标准开发流程

### UI 小改动

1. 在浏览器模式确认不同宽度。
2. 在 Excel 窄任务窗格确认真实布局。
3. 检查鼠标、键盘和触屏状态。
4. 运行前端测试和构建。

### 协议或 Excel 动作改动

1. 先写清请求/响应或动作结构。
2. 同步修改 `contracts.ts` 和 `models.py`。
3. 实现 Office.js 执行。
4. 决定文件夹模式支持还是明确拒绝。
5. 增加执行预览文字和风险分类。
6. 增加执行后验收，并明确该动作是 `verified` 还是 `executed_unverified`；不能用工作表存在代替复杂效果验证。排序应读取真实值顺序，筛选应读取 AutoFilter 条件，表格应读取真实名称、范围和表头状态；格式、数据验证和冻结窗格也应读取保存后的具体属性；图表和数据透视表应核对真实数据源与位置。无法跨执行通道稳定核对的属性必须保留为未独立验证。
7. 执行器先预检整份计划的工作表依赖、区域与矩阵尺寸、对象名称和 API 版本；中途失败时返回 `succeeded`、`failed` 和 `not_run` 动作明细。
8. 检查固化工具是否需要新增参数绑定。
9. 补前后端测试。
10. 更新架构文档。

图表强验收依赖 ExcelApi 1.12 的系列维度读取；数据透视表强验收依赖
ExcelApi 1.15 的数据源字符串读取。修改最低版本时必须同步预检测试和用户
可读错误，不能在执行后才因缺少 API 失败。

### 模型行为改动

1. 判断这是语义决策还是确定性计算。
2. 语义交给模型；计算尽量写成本地受控工具。
3. 不向意图模型发送原始数据行。
4. 有关键歧义时提供互斥选项，并保留“其他想法”。
5. 写入仍然只返回计划，不直接执行。
6. 为结构化输出失败设置有限修复次数，禁止无限重试。

## 9. 测试组织

前端测试与实现文件同目录：

- `contracts.test.ts`：协议断言
- `api.test.ts`：请求和错误封装
- `dataTools.test.ts`：确定性查询
- `cellNormalization.test.ts`：日期、编码和展示值标准化
- `deterministicTools.test.ts`：固化查询本地执行与模型调用指标
- `diagnostics.test.ts`：无敏感内容的结构化诊断
- `excel.test.ts`：执行辅助逻辑
- `splitAggregate.test.ts`：拆分聚合
- `storage.test.ts`：固化工具与迁移
- `tableSchema.test.ts`：表头识别
- `workbookIdentity.test.ts`：文件名信息

后端测试：

- `test_main.py`：API 和错误协议
- `test_models.py`：Pydantic 模型
- `test_planner.py`：基础模式和规划
- `test_folder_workbooks.py`：文件夹执行
- `test_folder_data.py`：pandas 完整数据、跨文件 ID、权限边界、合并和关联
- `test_no_business_hardcoding.py`：防止行业字段重新硬编码

修复缺陷时优先先增加一个能够复现问题的测试。

## 10. 常见问题

### 前端能打开，但显示“本地服务未连接”

检查：

```powershell
Invoke-RestMethod http://127.0.0.1:8765/health
Get-NetTCPConnection -LocalPort 8765 -State Listen
```

### 端口被旧进程占用

先找进程：

```powershell
Get-NetTCPConnection -LocalPort 3000,8765 -State Listen |
  Select-Object LocalPort,OwningProcess
```

再只停止确认无误的 PID：

```powershell
Stop-Process -Id <PID>
```

不要对未确认的进程或宽泛路径执行强制删除。

### Kimi 返回 401

- Key 与 `.cn`/`.ai` 服务区域是否匹配；
- `.env` 是否保存；
- 后端是否重启；
- Key 是否包含多余引号或空格。

### 模型返回 400

- 模型 ID 是否真实存在；
- 当前模型是否支持 tools/function calling；
- 图片是否只发给 `AI_VISION_MODELS` 中的模型；
- 查看服务端返回正文和 FastAPI 日志。

### 502/504

- 502 通常表示网关、网络、鉴权或上游 HTTP 错误，不等同于“模型繁忙”；
- 504 表示在配置超时内没有完成；
- 先看稳定错误码，再决定重试或修改配置。

### 新建终端自动进入 `.venv`

通常是 VS Code Python 扩展自动激活所选解释器，这是正常行为。项目测试和后端开发应使用 `.venv`。

## 11. 代码评审检查表

- 是否出现了具体行业字段、数据值或单个客户规则的硬编码？
- 是否扩大了用户选择的数据范围？
- 是否把本可本地计算的完整数据发给模型？
- 是否允许写入绕过预览？
- 是否同时更新 TypeScript 与 Python 协议？
- 是否处理 Excel 和文件夹两种模式？
- 是否增加了结构化、可读且可重试性明确的错误？
- 是否在窄任务窗格中验证交互？
- 是否补充了测试？
- 是否更新了相关文档？

## 12. 文档维护

以下变更必须同步更新 `docs/ARCHITECTURE.md`：

- 新增服务、数据流或持久化方式；
- 新增模型供应商适配层；
- 修改授权或安全边界；
- 新增本地工具；
- 修改动作协议或执行通道；
- 改变 `/api/turn` 状态机。

启动命令、环境变量、端口或调试方式变化时，更新本文和根目录 `README.md`。
