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
工作簿不会继承这次开发旁加载。`start:excel` 使用 `--no-debug`：它只注册并
旁加载清单，不开启依赖 Visual Studio 的 Office 直接调试器；前端仍由 Vite
提供热更新，浏览器开发者工具仍可检查任务窗格。停止旁加载：

```powershell
npm run stop:excel
```

修改 `manifest.xml` 后应停止再重新旁加载。普通 React/CSS 修改通常只需刷新任务窗格。
普通工作簿需要通过 Windows 共享加载项目录进行本地测试安装；正式发布使用
Microsoft 365 集中部署或 AppSource。

### 个人 Windows 安装包

构建机需安装 Inno Setup 7，并让 `ISCC.exe` 可从 `PATH` 找到，或设置
`INNO_SETUP_COMPILER` 指向它。执行 `npm run build:installer` 会先用 PyInstaller
冻结本地服务，再由 Inno Setup 生成标准 Windows 安装向导。安装向导提供安装目录、
开始菜单、确认、进度和完成页面，默认目录为
`%LOCALAPPDATA%\Programs\Excel Bro`，但允许用户修改。

Office 加载项清单的 `<Version>` 必须至少为 `1.0.0.0`，并包含顶层 `IconUrl`
和 `HighResolutionIconUrl`；构建脚本会在冻结运行时前检查这些要求。需要做完整
兼容性验证时，运行
`node node_modules\office-addin-manifest/cli.js validate apps/excel-addin/manifest.xml`
调用 Microsoft manifest 验证服务。

安装过程会自动创建
`\\<电脑名>\ExcelBroAddins` 本机共享并注册到 Office 可信加载项目录，同时写入
`HKCU\Software\Microsoft\Office\16.0\Wef\Developer\{清单 Id}` 直连注册（与
`office-addin-debugging` 旁加载相同）。使用实际计算机名而不是 `localhost`，
以兼容只枚举标准 UNC 服务器名的 Office 版本。创建共享需要一次 UAC 管理员授权。
安装后完全退出并重开 Excel，**Excel Bro** 标签会直接出现在功能区，普通工作簿
可用，无需再在“共享文件夹”里手动添加。卸载时 Inno Setup 会移除该注册值。

卸载入口由 Inno Setup 注册到 Windows“已安装的应用”和开始菜单。标准卸载器会先
要求用户确认，再请求一次 UAC 删除本机共享；UAC 被取消或共享删除失败时卸载会
中止，不会先删掉卸载入口或留下半卸载状态。成功后由 Inno Setup 删除程序文件、
可信目录、启动项和证书；`%LOCALAPPDATA%\Excel Bro` 中的个人模型配置默认保留。

### macOS 安装包

mac 包只能在 macOS 上构建，仓库通过 `.github/workflows/build-mac.yml` 在
GitHub Actions 上完成：`workflow_dispatch` 手动触发或推送 `v*` 标签，matrix
同时产出 x86_64（macos-13）和 arm64（macos-14）两个 `.pkg`。本地有 Mac 时也可
直接执行：

```bash
pip install -r server/requirements.txt pyinstaller
bash packaging/build_mac.sh
```

`packaging/build_mac.sh` 与 Windows 的 `build_installer.ps1` 步骤一致（manifest
守卫、前端构建、证书生成、PyInstaller onedir 冻结），最后用 `pkgbuild` +
`productbuild` 生成安装包。系统集成由 `packaging/mac/pkg-scripts/` 中 root 运行
的 preinstall/postinstall 完成，对应 Windows 的 `install_tasks.ps1`：

- 旁加载不依赖共享目录，直接把 `manifest.xml` 拷入 Mac Excel 容器的 `wef` 目录
  （必须以桌面用户身份写入，否则 Excel 沙盒读不到）；
- `localhost` 证书通过 `security add-trusted-cert` 信任到系统钥匙串，替换同名
  旧证书；
- 自启动使用 LaunchAgent `com.excelbro.runtime`，安装末尾轮询 `/health`，失败
  回滚证书与自启动并让安装器报错。

卸载使用安装目录内的 `uninstall.sh`（需要 sudo），保留
`~/Library/Application Support/Excel Bro` 中的模型配置。安装包未做 Apple 签名
公证，分发时告知用户首次右键打开即可。

## 5. 浏览器与 Excel 调试

### 备份目录打开

备份管理中的“打开文件夹”不使用 file://，因为 Excel WebView 会拦截该导航。前端先从备份列表响应取得短期 openDirectoryToken，再调用 POST /api/folders/backups/open。服务端只允许该令牌对应的当前工作目录下 .excel-bro-backups，不接收目标路径；过期令牌、路径越界、同名文件以及符号链接、junction 或 reparse point 都会被拒绝。备份目录尚不存在时列表不返回令牌，打开与复制路径按钮均禁用，打开接口不会创建空目录。此能力是项目“不启动外部程序”边界的窄例外，仅用于调用 Windows 文件资源管理器打开上述受管目录。

逐份删除备份使用 `POST /api/folders/backups/delete`。备份列表在确认统一备份目录已经存在且通过安全检查后签发短期 `backupManagementToken`；当前工作簿入口和文件夹模式使用同一授权方式。请求只能包含令牌、32 位 `undoId`、明确确认、列表取得的 manifest revision 与级联失效 ID，不能提交磁盘路径。后端复核影响后先将完整目录原子移动到 `.trash/<operation_id>/`，再同步同根目录各会话的撤销资格并递归清理；递归清理失败必须返回 `cleanupPending`，且 `bytesFreed` 不能把仍在暂存中的资料算作已释放。自动清理必须复用同一事务，并安全重试合法的 `.trash` 残留。前端使用卡片内确认区，不使用 Excel WebView 可能拦截的 `window.confirm`。

“取出副本”使用 `POST /api/folders/backups/export`，前端为一次用户动作生成稳定 `operationId`；结果不确定时重试必须复用该 ID。服务端固定写入 `exports/<operation_id>/` 并持久保存来源、清单、指纹和完成状态，同一操作重放只返回已核验结果。导出期间将源恢复点加入并发保护集合，删除、撤销和自动清理不能同时改变它。

文件夹执行中的 `changed` 必须由本地确定性逻辑计算，不能交给模型判断。只有能够无歧义比较的动作才可以返回 `false`；新动作或复杂对象尚未实现比较时必须默认为 `true`。当本次计划没有任何实际变化时，执行器不得替换原工作簿，也不得创建 `.excel-bro-backups`。恢复点 manifest 的 `recoveryAction` 是向后兼容的新增字段，读取旧记录时仍需通过 `backupRelativePath` 推导 `restore` 或 `remove`。

### 文件夹任务中断与恢复调试

开发和安装环境使用同一应用数据解析规则。设置 `EXCEL_BRO_CONFIG_DIR` 时以该目录为准；
否则 Windows 使用 `%LOCALAPPDATA%\Excel Bro`，macOS 使用
`~/Library/Application Support/Excel Bro`。文件夹任务账本位于 `tasks/state.sqlite3`，
工作簿副本和恢复事务仍只放在用户已登记工作目录的 `.excel-bro-backups` 中，不复制到
应用数据目录。

- 前端发起文件夹写入前生成一次稳定 `operationId`，刷新、断线重连和状态查询必须复用；
  不得因旧会话失效生成新标识。
- `progress.folderOperationTtlSeconds` 只控制进程内热缓存；持久终态保留期由
  `taskRecovery.operationRetentionDays` 控制。到期后只清除事件/结果载荷并保留幂等墓碑。
- 服务重启后，已完成任务可重放结果；遗留 `running` 文件夹任务先进入核对。只有单个
  `splitGroupAggregate` 且暂存产物、manifest 与真实文件指纹都完整时，才会通过
  `POST /api/folders/execute/resume/stream` 继续剩余发布；其他动作不会按步骤序号跳过。
- `GET /api/folders/recovery/directories` 只返回后端登记目录。独立恢复必须携带目录 ID、
  `undoId`、manifest revision、稳定恢复操作 ID 和明确确认。
- `folder_write_guard` 对规范化工作目录提供进程内和跨进程排他。执行、撤销、恢复均须
  经过它，测试不能用不同大小写或路径别名绕开。

排查刷新后任务时，先查看 `GET /api/folders/execute/status?operationId=...`：`completed`
表示可重放结果，`failed` 表示有确定失败，`needs_reconciliation` 表示不能安全判断是否
完成。HTTP 410 `OPERATION_EXPIRED` 是安全终态，不允许重新发起同一标识。测试应用数据
必须通过 `EXCEL_BRO_CONFIG_DIR` 指向隔离临时目录，不读取或修改用户真实账本。

### 规划预算暂停与恢复调试

流式规划通过 `POST /api/turn/stream` 建立持久 turnId。软时间份额到达后只在完整模型响应
或完整单次工具调用边界暂停，事件流返回 `planning_pause`；前端保存 turnId、revision、
lastSequence、稳定 requestId 和原来源指纹。刷新或网络中断后应先调用
`GET /api/turn/{turnId}/status`，用户明确继续时再调用 `POST /api/turn/continue/stream`。

- 同一继续动作重试必须复用 requestId；旧 revision 应返回冲突，不得重复增加预算。
- 当前默认首次/每次追加均为 60 秒，硬上限 300 秒，最多追加 4 次；配置位于
  `taskRecovery`。继续只增加时间，不提高 `agent.maxTurns`、`maxTotalTokens` 或错误上限。
- checkpoint 保存在 `tasks/state.sqlite3`，包含消息、LoopState、工具缓存、工具模式、修复
  状态和完整待处理模型响应。不要把完整 checkpoint 写入普通诊断日志。
- `/api/plan` 是兼容非流式入口，不具备 B2 的跨重启继续保证。流式路径会逐个保存完整
  工具调用；模型错误暂停后可由用户显式切换到已配置连接，但任意 DAG 仍未实现。
- 意外异常必须把任务标记为 `failed` 并发出安全 SSE 错误；若状态长期为 `running`，先检查
  账本写入失败和服务日志，不要直接新建同一语义的写入任务。
- 若状态查询返回 `interrupted`，只有“账本仍为 running、内存无活动执行者、完整 checkpoint
  存在”时允许继续或显式换模型。顶部模型切换后，恢复按钮应显示具体目标模型，并且请求仍为
  `POST /api/turn/continue/stream`，不得重新调用 `/api/turn/stream` 或意图识别。
- 前端恢复时应复用原过程消息、`turnId` 和 `operationId`。`afterSequence` 取客户端最后实际收到
  的序号，并以状态响应的 `lastSequence` 夹紧；若续跑出现第二张空过程卡或旧步骤从第 1 步重现，
  优先检查 `activityMessageId`、`persistenceKey` 和持久日志中的 `sequence`。

### Office.js 动作断点恢复调试

当前只对已保存工作簿中的 `writeValues`、`writeFormulas`（非 R1C1 铺满）、`setFill`、
`setFont` 和 `setNumberFormat` 开放。来源表原地修改、工作表结构变化、图表、透视表等仍按
普通模式执行，并明确提示不支持跨重启自动续跑。

- 首个写入前调用 `POST /api/office/executions/register`；每一步写入前后分别提交
  `actions/start` 和 `actions/complete`。完成回执必须早于下一动作。
- 刷新后先用本地 operationId 调用状态查询；本地记录只是发现索引，不能替代服务端计划、
  授权、revision 或检查点。
- 在途动作只能通过实际目标范围判定：匹配预期结果时补记完成，匹配写前指纹时允许重试，
  两者都不匹配时保持挂起。不得为了恢复而新建 operationId 或改用文件夹执行器。
- 工作簿身份以 `Office.context.document.url` 的稳定 URL 绑定；未保存工作簿和同名不同路径
  文件必须拒绝自动接续。
- 状态查询不包含 ownerToken。登记、显式领取或成功核对会取得/轮换写入令牌；旧 token 或
  revision 的动作回执必须返回冲突。
- 浏览器演示模式不能验收这些行为。真实 Excel 至少注入：写前登记后退出、`context.sync()`
  后回执前退出、回执后退出、工作簿切换、人工修改在途范围和最终完成回执丢失。

### Custom Functions 资源

清单中的 Custom Functions `<Script>` 必须指向可直接执行的 JavaScript 文件，而不是 `functions.html`。`npm run dev:addin` 的 `predev` 钩子和 `npm run build:addin` 都会生成 `apps/excel-addin/dist/functions.js`；开发服务会通过 `/functions.js` 提供该 IIFE bundle，安装包则直接从前端静态目录提供它。修改清单后，需停止并重新旁加载 Excel 加载项。

### 浏览器模式

打开 `https://localhost:3000`：

- 使用 `src/demo.ts` 的演示工作簿；
- 适合调试布局、历史对话、模型选择和普通 API；
- 不会真实写入 Excel；
- 无法完整验证 Office.js 行为。

也可以直接打开 `https://localhost:3000/focus.html` 检查专注窗口的响应式布局。
直接打开时只有本机缓存的最近展示快照；窗口通信、实时内容同步和关闭行为仍须
从 Excel 任务窗格中的“显示方式 → 打开专注窗口”验证。

### Excel 模式

必须在 Excel 中验证：

- 当前工作表和选区读取；
- 多工作表顺序；
- Office.js 动作；
- 图表、透视表、图片等 API 兼容性；
- 写入后的真实验收；
- 标准宽度（约 350px）和加宽任务窗格的布局、键盘及触屏交互；
- `TaskPaneApi 1.1` 不可用时的加宽降级提示；
- 专注窗口打开、内容同步、返回任务窗格，以及 `DialogApi 1.2` 不可用时的
  本地快照回退；
- 专注窗口中没有 Excel 执行入口，所有写入仍回到任务窗格预览确认。

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
| 任务进度、事件归并与审查轨迹 | `contracts.ts`, `models.py`, `useActivityProgress.ts`, `main.py` |
| 文件夹任务身份、重连与独立恢复 | `task_ledger.py`, `folder_mutex.py`, `folderOperationRecovery.ts`, `useFolderBackups.ts` |
| 需求确认机制 | `server/app/intent.py` |
| 模型配置、鉴权和 HTTP 适配 | `server/app/llm/` |
| 模型 Agent 工具 | `server/app/excel_agent.py` |
| 基础模式和模型选择 | `server/app/planner.py` |
| Excel 动作协议 | `contracts.ts`, `server/app/models.py` |
| 文件夹 Excel | `server/app/folder_workbooks.py`, `server/app/folder_data.py` |
| 固化工具 | `apps/excel-addin/src/storage.ts`, `deterministicTools.ts` |
| 本地诊断 | `apps/excel-addin/src/diagnostics.ts`, `server/app/main.py` |
| 限额与 Excel 批处理参数 | `config/capabilities.json` |
| 功能区按钮与地址 | `apps/excel-addin/manifest.xml` |

## 8. 标准开发流程

### UI 小改动

1. 在浏览器模式确认不同宽度。
2. 在 Excel 的标准与加宽任务窗格确认真实布局。
3. 涉及长内容时同时检查专注窗口。
4. 检查文字是否只使用 `styles.css` 中 12/14/16/20/24px 的语义字号，并确保正文
   默认采用 14/20px。
5. 检查鼠标、键盘和触屏状态。
6. 运行前端测试和构建。

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

为动作增加跨重启续跑能力时，还必须提供动作专用的写前状态、预期写后状态和 `unknown`
判定测试。不能仅因动作已经支持普通执行，就把它加入恢复白名单；结构动作和复杂对象必须
先建立稳定对象身份及完整效果核对。

文件夹模式的物理备份必须写入所选工作目录的 `.excel-bro-backups`，不得恢复为原文件旁散落备份。新增清理路径时必须复用统一事务删除，并同时验证 manifest、活跃撤销 ID、路径边界、符号链接和 Windows reparse point；不能直接递归删除 `executions/<undoId>`。manifest 使用同目录临时文件加 `os.replace()` 原子更新。`undo.maxEntries` 只限制快速撤销资格；历史磁盘副本按最短保留期、软上限、每文件最后完好副本、主动保留和异常状态分别判断。磁盘安全余量使用 `undo.folderBackupMinFreeBytes` 配置，不得在空间不足时降级成无备份写入。

新增耗时操作时必须发送统一 `TaskProgressEvent`，并使用稳定 `operationId`、单调
`sequence` 和可复用 `stepId`。用户可见 `trace` 只能由受控代码根据已授权元数据、
工具安全摘要和验收结果构造；不得传递模型隐藏思维、提示词、API 凭据或原始数据行。
文件夹写入必须优先使用 `/api/folders/execute/stream`，网络中断只能用同一
`operationId` 和最后 `afterSequence` 重连，不得回退为可能重复写入的新执行请求。
白名单工具事件必须填写真实 `toolName`；Excel 动作事件必须保留 `actionType` 和授权
范围内的 `target`。可见参数和结果放入 `facts` 时只使用受控摘要标签，不得把工具原始
返回值拼入事件。前端解析器仍会把 facts 限制为 12 项、单值 500 字符，并对疑似凭据、
未知标签和越界绝对路径执行防御性隐藏。实时与历史过程统一由
`TaskProgressTimeline` 渲染，默认展开且步骤内部不再增加折叠层。

进度问题可通过“设置 → 导出诊断”获取脱敏事件元数据。诊断只包含
`operationId`、`executionId`、`sequence`、`stepId`、阶段、状态、工具、动作和归并结果，
不得记录 summary、detail、facts.value 或工作表数据。服务端日志使用同一组关联 ID。

图表强验收依赖 ExcelApi 1.12 的系列维度读取；数据透视表强验收依赖
ExcelApi 1.15 的数据源字符串读取。修改最低版本时必须同步预检测试和用户
可读错误，不能在执行后才因缺少 API 失败。

拆分聚合的 Office.js 同步批次由
`excelExecution.splitAggregateBatchSheets` 控制。调整该值时需要同时测试
少量输出、超过一个批次、覆盖已有工作表及批次失败清理；不要恢复为每张
结果表一次 `context.sync()`。

### 模型行为改动

1. 判断这是语义决策还是确定性计算。
2. 语义交给模型；计算尽量写成本地受控工具。
3. 不向意图模型发送原始数据行。
4. 有关键歧义时提供互斥选项，并保留“其他想法”。
5. 写入仍然只返回计划，不直接执行。
6. 为结构化输出失败设置有限修复次数，禁止无限重试。
7. 澄清选项若改变拆分维度、输出结构、指标口径或目标，必须返回 replace 模式和
   可独立执行的完整 resolution；只补充缺失参数时才使用 append。
8. 计划格式修正轮不得继续携带完整失败参数，也不得重新开放只读工具；输出额度
   使用 agent.repairMaxOutputTokens。
9. splitGroupAggregate.splitBy 会自动写入输出列，提示词和工具 schema 都必须
   明确禁止在 groupBy 中重复；模型仍重复提交时，服务端只允许确定性移除完全
   同名项，并把修正写入计划 warnings，不得猜测或替换其他业务字段。
10. 多个循环闸门同时命中时，连续结构错误优先于累计 token 展示为停止原因，避免
    把格式错误造成的 token 增长误报成根因。
11. 模型提交的 AnalysisPlan 顶层字段若由 default_factory 提供默认集合，null 可在
    模型提交规范化阶段按省略处理；必填字段、普通可空字段及动作业务参数不得借此
    放宽校验。

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
- `test_task_ledger.py`：持久任务身份、事件重放、异常退出和过期墓碑
- `test_folder_mutex.py`：同进程及跨进程文件夹写入排他
- `test_managed_directories.py`、`test_managed_recovery.py`：受管目录发现和独立恢复
- `test_models.py`：Pydantic 模型
- `test_planner.py`：基础模式和规划
- `test_agent_loop_protection.py`：Agent 预算、循环检测、降级与收尾提示
- `test_agent_integration.py`：Agent 主循环熔断和警告注入
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
