# Excel Bro 开发指南

### AI 40% 比例验收（2026-10-05）

自动化结果：Excel 730、WPS 406、core 213、服务端全量 790；两端 tsc/Vite、compileall 与 diff --check 通过，最终无失败或环境阻断。WPS 既有 act 警告不变。自动化未操作真实 Excel，不能替代下述验收。

重启 Python 后台（SDK 启动命令不变）、刷新 EB，关闭旧临时副本且不保存，在独立测试工作簿重新生成并排副本。右侧 Excel 窗口内部应为原表约 60%、AI 约 40%；外层仍为副本 40%、原表与 AI 60%。右侧客户区 1200px 时目标 480px，1000px 时目标 400px；700px 时下限 320px，500px 时 Office 半窗口上限 250px 优先。改变右侧窗口宽度，等待约 1.5 秒；拖动 AI 分隔线后再改窗口宽度，AI 不应被自动抢回；新审阅重新启用比例，退出审阅不改变普通宽度偏好。同步滚动与保存确认不受此尺寸调整影响。

测量只允许本机 Excel 来源，HTTP 请求只含两种身份摘要及本次 nonce；旧后台 404、窗口关闭/换表/副本另存、TaskPaneApi 不支持或静默忽略时显示手动指引，不猜测宿主宽度、不自动执行。新增回归覆盖真实 fetch 字段投影、严格回执、DPI 96/144/192、唯一原表/副本与测量后身份重查、手动优先、程序 resize 不回环、忽略/抛错、异步卸载和无重叠请求。真实宿主尤其 100%/150%/200% 系统缩放、WebView 缩放与多显示器需单独验收。

### 间隙与同步滚动验收（2026-10-05）

本次后台新增原生比较操作和可见边框补偿，必须停止旧服务后按原启动命令重启；SDK：`powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start_claude_sdk_server.ps1`。刷新 EB，手动关闭旧临时副本且不保存，再生成新副本。prepare/arrange 必须使用版本 3，旧的版本 2 会提示重启，不解释为同步成功。

左副本 40% / 右原表与 AI 60% 不变，默认可见间隙 0 DIP；DWM 边框补偿消除不可见拖拽边框导致的额外空隙。检查上下/左右滚动：同实例、仅两份窗口且原生比较与同步开关实际回读成功时，卡片显示「已启用 Excel 原生并排比较与同步滚动」。跨实例只能排列，自动同步未实现；有其他窗口或开关不可核实则手动「视图 → 并排查看 → 同步滚动」，不得用排列成功当成同步成功。不同实例的手动指引需先将副本保存到本机并从原实例打开，不覆盖原文件；不自动执行另存或关闭用户文档。

自动化：Excel 711 / 68 文件、WPS 406 / 28 文件通过，两端 tsc/Vite 通过；服务端全量 779 通过（窗口模块 67），compileall 与 git diff --check 通过；最终无失败或环境阻断。先新增用例 8 红，再实现；真实 fetch 回执、两条 staging 路径、UI 实际点击、COM 包装/集合独立身份和属性静默忽略均覆盖。`check_excel_layout_guards.py` 核对 7 项身份/凭据守卫、2 项顺序/可见边框补偿变异、4 项原生同步独立守卫。初次 COM Hwnd 变异被集合判据遮住，改用独立包装对象与集合证据后被抓住。全程无真实 Excel 操作；WPS 既有 act 告警保留。必须另验真实滚动、边框间隙、DPI、多显示器与 AI 留存。

### 审阅时原表被 AI 挤窄（2026-10-04）

外层仍是左副本 40%、右原表与 AI 60%；右侧内部 AI 在审阅时尝试临时收窄到 360 CSS px（配置 `excelWindowLayout.reviewTaskPaneWidthCssPx`）。支持 `TaskPaneApi 1.1` 才调用宽度接口，静默忽略/异常/不支持时给手动提示：将 Excel Bro 左侧分隔线向右拖，给原表留出空间。审阅期间不恢复普通加宽偏好、不保存临时宽度，退出审阅恢复原偏好。此次前端修复只需刷新加载项，不要求重启 Python 服务。

新增 6 项测试覆盖进入/退出审阅、焦点恢复、偏好隔离、不支持/静默忽略/异常、手动拖动与真实卡片提示，定向 34 项通过。全量 Excel 698、WPS 406、服务端 765 通过，两端 tsc/Vite 与 compileall 通过，git diff --check 通过；WPS 保留既有 act 告警，无最终失败或环境阻断。实机待验：刷新后生成新副本，检查窗格实际收窄、原表可见列和退出审阅的宽度恢复；未操作真实工作簿。自动调用不保证所有 Excel 版本均可调整，也不保证任意分辨率下三栏固定比例。

### 图 2 布局生效检查（2026-10-04）

目标为左侧审阅副本 40%、右侧原表与 AI 60%，间隔 6 DIP。SDK 后台不热重载：只刷新 Excel 前端不会加载新的 Python 排列代码。停止原后台终端后用原启动命令重启（SDK 使用 `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start_claude_sdk_server.ps1`），刷新加载项，关闭旧临时副本且不保存，重新生成副本。新版 prepare 必须返回 `layoutVersion: 2`；旧后台在 UI 明确提示重启，不报已按新顺序排列。

本轮新增真实 fetch 旧凭据零排列请求、旧通用成功/错误左右回执拒绝、兼容副本状态与重启提示点击测试，以及服务端无几何回执不得报成功。定向前端 61 条与窗口模块 53 条通过；全量 Excel 692、WPS 406、服务端 765 通过，两端构建与 compileall 通过。七项独立守卫及两项几何内存变异均被抓住。WPS 有既有 act 警告；最终无失败/环境阻断。目标左右布局仍需后台重启后的宿主复测，本轮未操作真实窗口/工作簿。

### 顶部图标按钮（2026-10-04）

Excel/WPS 的模型、新对话、工具、备份入口只显示图标，保留 `title` 悬停提示与 `aria-label` 名称。模型使用共享 `ModelPickerIcon`，入口统一 32×32 CSS px；恢复数量以角标显示，不挤占图标空间。品牌下不再展示模型名称与正常连接状态行，当前模型可在悬停提示和菜单选中项查看；加载、未配置与故障提示仍保留。图标明确覆盖旧 `justify-content: space-between`，水平/垂直居中。现有禁用、菜单展开、授权与写入行为不变。

头部按实际空间自然换行，删除旧 ≤700px 强制两行布局与 WPS ≤460px 隐藏品牌的覆盖。品牌最小宽度由文本的 min-content 决定，不人为预留固定宽度；六个图标归入同一工具栏。`check_taskpane_layout.mjs` 在 320/350/400/430/500/620/700/900 CSS px 下检查真实组件的图标尺寸、名称、中心偏移≤0.5 CSS px、品牌无模型行与菜单选中项；当前默认字体下 400px 及以上必须一行，320px 必须自然换行且不裁切。`--force-header-wrap` 在测试内存里强制按钮分行，用于验证断言确实能抓住提前换行。WPS 面板测试检查图标入口实际打开模型、工具和恢复抽屉。浏览器/DOM 通过不等于宿主实机验收。

### 窄窗格与进度交接回归（2026-10-04）

`styles.test.ts` 检查宽度约束与 App 的实际请求准备交接位置；`useActivityProgress.test.ts` 覆盖本地准备 → 服务端理解/生成 → 等待确认，以及缺少结束事件时不得保留运行状态或假报成功。
浏览器尺寸测量单独运行 `node scripts/check_taskpane_layout.mjs [Playwright包路径]`，加 `--wps` 渲染真实 WPS 面板与宿主覆盖样式；需已有 Playwright 和 Microsoft Edge，不自动下载安装。真实组件用隔离上下文与模拟模型/健康响应，源表、模型服务和用户浏览器存储不参与；时间线与本地澄清夹具挂在消息容器中。`--without-containment` 只在测试内存中移除新宽度约束，用于证明布局断言能复现裁切，不修改生产文件。

通过浏览器尺寸检查后，仍须在 Excel/WPS 任务窗格拉窄、加宽并检查模型菜单、表单填写、键盘与真实步骤事件。jsdom/CSS 文本断言不能代替布局证据，Edge 无头渲染不能代替宿主验收。

## 2026-10-03：网页对比移除

不再构建 Excel `review.html`，WPS 不再切换到全幅网页双网格。`ChangeReviewSummaryCard` 仅显示定位、计数与真实核对状态；状态文案及 `ChangeReviewUndoNote` 保留在 `ChangeReviewPresentation.tsx`。原生 WPS 副本审阅、执行器回读证据、撤销与恢复不受此删除影响。

`webComparisonRemoval.test.ts` 验证两端旧入口及专用页面移除，同时检查证据收集与原生副本链路仍被消费。删除的是网页渲染专用测试与网格投影测试，不是执行核对测试。

## 2026-10-03：Excel 原生副本审阅接线

已保存的原工作簿、单张授权表、区域写值/矩阵公式/清空内容计划，在计划卡点击「生成副本并排审阅」。右侧仅包含该表数据与基础格式，不导出整本工作簿；核对后点「保存变更到原表」，放弃则只关闭临时副本。关闭时 Excel 若询问是否保存临时副本，请选择不保存；原文件保存仍按 Excel 正常流程进行。多表、复杂对象、R1C1 铺满和跨表/外部依赖未实现此审阅方式，不能列为“仅待实机验收”。

基础路径需支持 ExcelApi 1.13 以及 createWorkbook。支持 ExcelApiDesktop 1.1 时走自动模式：实际新窗口必须在同一 Application.windows 集合可发现，活动表必须匹配本次唯一标识；若打开到另一实例，停止并提示手动关闭副本，不写原表。不要放宽身份守卫来令假宿主通过。

2026-10-04 增加 Office 2024 等无 Desktop 1.1 宿主的兼容模式。点击「生成副本并排审阅」直接打开包含公式、基础格式与计划变化标记的单表文件，不调用不支持的 Office Window API、不先写原表。既有本地服务会尝试将原表与本次副本左右排列；失败时明确提示手动并排，不假报成功。检查公式栏及实际计算结果后勾选确认，再点「保存变更到原表」。若结果不符则放弃并手动关闭副本且不保存；保存也只执行原计划，不导入副本人工编辑。原表冲突检查与首次写入前守卫保留。不能自动验证/关闭兼容副本，UI 与最终结果不得伪称已回读副本或与副本一致。复杂对象、多表与跨表引用仍未实现。

新增 `manualReviewFile.test.ts` 检查公式类型、常量、格式、颜色、绝对坐标和大小限制；严格的旧宿主替身一触碰 Office 窗口 API 就抛错，测试实际 `createWorkbook` 请求体包含本地敏感哨兵。布局桥接不可用时 fetch/存储无新增调用；实际布局请求的隐私另由 `excelWindowLayout.test.ts` 覆盖：即使错误附加完整本地计划，真实 fetch 请求体也只能包含摘要/nonce 或 token。点击测试覆盖确认门槛、保存、冲突与放弃，自动布局不自动确认审阅。运行 `node scripts/check_manual_review_file.mjs` 可用生产序列化器生成纯合成内存副本，并用本地 openpyxl 独立验证 OOXML 可读性、公式/格式/标记与重算设置；不访问真实工作簿，不运行 Excel，不验证公式计算结果。自动与兼容两模式的实际宿主行为均待实机验收。

### 本机自动窗口排列的启动与验收（2026-10-04）

无需第二个服务或脚本进程。Windows 依赖 `pywin32==312` 已列在 `server/requirements.txt`（仅 win32 安装）；当前虚拟环境已有该版本。改动后重启现有 Python 服务（普通 `npm run dev:server` 或原 SDK 启动命令，不能同时占用 8765），刷新 Excel 加载项。使用单个服务进程，避免 prepare/arrange 落到不同 worker；热重载也会清掉一次性凭据。不要为了访问 Excel 把服务改成系统后台会话或提升管理员权限。

在独立、已保存的本机测试工作簿，激活已授权目标表，关闭上次临时副本且不保存，然后重新点击「生成副本并排审阅」。新版本服务在打开前记录原表与已有窗口，故不能把旧副本补认成本次新副本。正常应显示左侧审阅副本 40%、右侧原表与 AI 窗格 60%，版本 3 中间可见间隔默认为 0 DIP（DPI 缩放后换算），补偿不可见拖拽边框，两窗高度使用显示器工作区；若能力、身份、权限或多窗口证据不满足，则显示「自动排列未完成」并保留手动审阅流程。未成功定位不得改用标题猜测或全体窗口重排。Windows 系统边框样式不修改，后续宽度/间距调整只改 `excelWindowLayout.originalWidthFraction` / `windowGapDip` 并重启服务。用户已反馈上一版自动排列可用；本次顺序、间距和不同 DPI 尺寸仍需重新实测。

`server/tests/test_excel_window_layout.py` 验证原表路径、nonce、仅新窗口、进程身份、令牌来源/期限/单次使用、忙/模糊/换表、几何验证与失败回滚，以及真实 HTTP 请求、无账本持久调用和固定错误响应。尺寸测试覆盖 96/192 DPI、负坐标工作区、左副本/右 AI、6 DIP 间隔与非法尺寸零移动。`python scripts/check_excel_layout_guards.py` 对生产类做六项独立安全守卫内存变异，另独立测试改回旧顺序和取消间隔缩放的两项回归；必须逐项被测试抓住，不编辑生产源文件，不枚举或操作真实 Excel。真实显示器 DPI、Office 2024 可访问性元数据、另一实例新窗口与权限隔离仍需实机验证；自动布局不代表已实现兼容副本回读、关闭或同步滚动。

定向测试：`nativeReviewFile.test.ts`（OOXML 坐标/类型/文本/上限）、`excelNativeComparison.test.ts`（真实目标写入记账、原表零写入、副本公式、各种独立现场守卫、关闭取消、执行器最后时刻守卫）、`ExcelNativeReviewActions.test.tsx`（实际点生成/保存/放弃、消息归属、冲突和重复提交）。这些不证明真实宿主窗口布局。实机在独立测试工作簿中验证双窗口、左右顺序、任务窗格、公式计算、颜色、保存前人工改原表/副本的拒绝、放弃和取消关闭。不得用 jsdom 通过宣称实机完成。

这份文档用于让新开发者快速启动、定位代码并安全地扩展功能。系统设计原因和数据流参见 [ARCHITECTURE.md](./ARCHITECTURE.md)。

## 1. 环境要求

- Windows 10/11
- Microsoft Excel 桌面版
- Node.js 与 npm
- Python 3.9+（可选 Claude Agent SDK 后端需 Python 3.10+）
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

### 规划追问的交互契约

规划 Agent 缺少用户决定时使用 `submit_clarification`，不要用 `submit_answer` 输出追问。
每项待填写信息放入 `clarification.fields`（`id`、`label`、可选 `placeholder`），
没有合理选项时 `options: []`。该协议同时用于 SDK 与 OpenAI-compatible 后端，
Office.js、文件夹模式和 WPS 均复用需求确认卡；不得在前端匹配中文句式生成表单。
表单补齐后仍重新核对授权范围、重新规划并等待预览确认，不能跳过审批写入。
规划后的澄清从新的 turn 继续，避免持久规划指纹与补充后请求冲突；原需求和回答都保留。
回归重点：两后端工具/JSON 出口、HTTP 缓存/续答、SSE 字段不丢失、必填校验、
WPS 面板提交续答与不提前写入。真宿主布局仍需用户重测；旧普通文字回复不会自动转成表单。

推荐直接在任务窗格顶部点击“添加模型”。如需使用旧环境变量方式，可编辑
`server/.env`：

```env
AI_BASE_URL=https://api.moonshot.cn/v1
AI_API_KEY=你的密钥
AI_MODEL=kimi-k3
AI_MODELS=
AI_VISION_MODELS=kimi-k3
AI_INTENT_TIMEOUT_SECONDS=180
AI_TIMEOUT_SECONDS=180
```

注意：

- `AI_MODEL` 必须是供应商实际接受的模型 ID。
- `AI_MODELS` 用英文逗号分隔可切换模型。
- 只有列入 `AI_VISION_MODELS` 的模型才能接收图片。
- 本地 Ollama、LM Studio 等无鉴权网关可以留空 `AI_API_KEY`。
- 新安装的配置模板不会预置服务商或模型；未配置时模型列表为空，顶部入口显示「添加模型」并给出首次添加引导。
- 模型目录区分四种状态：加载中、加载失败（可重试，且不代表你没有配置模型）、加载成功但为空、已配置真实连接。读取失败时不清空已加载的列表，也不覆盖本机已保存的模型选择。
- WPS 任务窗格打开时只读取一次模型目录；失败后不按定时器或窗口焦点持续轮询。恢复后台后可在模型菜单手动点击“重试读取模型配置”，模型设置读取成功或保存连接后也会刷新目录。
- 点击“添加模型连接”立即打开可填写的表单，同时异步检查本机已保存的连接并刷新目录；检查中、已找到连接、确实为空、检查失败分别显示，失败不阻塞填写，也不误称为“未配置”。相同服务地址和模型 ID 允许新增独立连接，测试成功后保存；重复模型 ID 在列表中显示 `(1)`、`(2)` 后缀，各自凭唯一连接 ID 使用自己的 API Key。
- 界面显示的模型与请求携带的 `modelId` 必须来自同一次解析（Excel 与 WPS 共用 `modelSelection.ts`）；没有任何可用连接时不发送模型请求，也不退回规则模式。旧的 `local` 记录会在目录加载成功后迁移到有效真实连接，删除连接后重新计算选择。
- 任务窗格顶部模型菜单同时提供选择、添加和管理入口。每个独立连接拥有自己的服务地址、模型 ID 和 Key，并写入配置文件同目录下的 `model-connections.json`。
- WPS 的“管理模型连接”列表可直接点击可用的非当前连接切换模型，设置窗格保持打开并原位更新“当前使用”；目录暂不可用时显示手动重试入口。新增表单在连接区域展开，编辑表单紧贴被点击的连接卡片，取消不会保存草稿。
- 设置页修改会立即生效，完整 Key 不会返回前端。`model-connections.json` 包含密钥，已被 Git 忽略，不能截图、记录日志或手动提交。
- “测试连接”会发送一个内容固定、输出极短的模型请求，不会读取或发送工作簿数据，也不会保存尚未确认的连接。
- 新增连接保存后会自动切换到新模型。编辑连接时只显示现有 Key 的已配置状态和脱敏尾号，不回显完整 Key；留空保留旧值，输入新 Key 才替换，“显示新 Key”只用于查看本次输入。需要接入无鉴权本地服务时可显式勾选清除。
- 同一服务地址和模型 ID 可以重复添加，不能按模型 ID 覆盖已有连接；如果 `model-connections.json` 损坏，服务会明确报错并保留原文件，避免静默覆盖密钥。
- 手动修改 `.env` 中的服务地址、模型列表或其他参数后，仍需重启 FastAPI 服务。
- 不要把真实 Key 写入 `.env.example`、前端代码、截图、日志或提交记录。

### 可选 Claude Agent SDK 后端

本机 `claude-agent-sdk-python/` 是用户提供的 SDK 源码；它本身不含原生 Claude Code
二进制。Windows 上使用官方平台 wheel，内含原生 `claude.exe`。本项目固定 SDK
`0.2.159` 与兼容 FastAPI 的 MCP/Starlette 组合；本机源码标为 `0.2.163`，但该版本
当时未在包索引提供可安装 wheel，不要直接用 editable 安装覆盖现有 wheel：

```powershell
.\.venv\Scripts\python.exe -m pip install --only-binary=claude-agent-sdk -r server\requirements-claude-sdk.txt
.\.venv\Scripts\python.exe -m pip check
```

若已在 Excel/WPS「添加模型连接」中保存 DeepSeek 连接，无须再次输入模型 ID 或 Key。
运行 `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start_claude_sdk_server.ps1`：
脚本只给本次服务进程设置仓库外会话目录，默认位置为本机
`LocalAppData/ExcelBro/ClaudeAgentSdk`，不读取或复制密钥、不更改既有模型配置。
在模型菜单中选择由原连接衍生的「Agent SDK · …」；原连接仍可按旧通道选择，默认值也不变。
该脚本**替代**下面日常启动中的 `npm run dev:server`，不要同时启动两套后端。
启动前检测 8765 端口；已有服务占用时明确报错，不自动停止用户进程。
Windows 上 SDK 启动采用单进程 `--loop asyncio --workers 1`，不加 `--reload`。
当前 Uvicorn 的 Windows 热重载/多 worker 模式会选用 Selector 事件循环，不能启动
SDK 的原生子进程；普通模型通道仍可使用原有热重载命令。修改 SDK 后端代码后手动重启。
如果误在该循环里选用 SDK，后端会在调用前提示正确启动命令，不再误报网络或 `AI_BASE_URL`。
仅官方 DeepSeek `https://api.deepseek.com` 或 `/v1` 的已保存连接会生成衍生选项；
SDK 请求使用官方 `https://api.deepseek.com/anthropic` 兼容端点，其他服务商不能只靠
修改模型 ID 进入 SDK。模型 ID 必须仍是 DeepSeek 当前实际接受的 ID。

若要使用 Anthropic 原生连接，可另在 `server/.env` 私下配置
`CLAUDE_AGENT_SDK_MODEL`、`ANTHROPIC_API_KEY`、`CLAUDE_AGENT_SDK_STATE_DIR`
（仓库外绝对路径），对应模型菜单中的独立「Claude Agent SDK」选项。
会话目录会保存 SDK 对话资料，应限制本机访问、按隐私策略保留/清理，不能提交 Git。
该 `.ps1` 保持 ASCII 文本，以兼容 Windows PowerShell 5.1 对无 BOM UTF-8 脚本的读取；
若修改脚本提示文案，不要直接加入非 ASCII 字符。
不要把真实密钥写进示例文件。SDK 的 `max_budget_usd` 不保证等于 DeepSeek
实际计费，试用时应查看 DeepSeek 平台账单与额度。

当前 SDK 通道仅用于表结构意图与 `planningReadPolicy=none` 的无数据行规划：
既不上传原始数据行、上一轮紧凑结果、历史对话、图片或本地文件路径，也不开
SDK 的 Read/Bash/Edit 等工具。需要这些输入的请求会明确拒绝并提示用户切回原模型；
依赖上一轮对话指代的请求应在本轮重新说明完整需求，不能期待 SDK 自动读取旧会话。
`/function`、AI 编辑计划等未接入的子功能同样不会把 SDK 凭据误发到 OpenAI 兼容端点。
SDK 通过 EB 白名单终局工具提交回答、列映射规则或预览计划，写入仍需用户确认。
终局工具使用 `mcp__excel_bro__*` 全名；供应商返回合法结构化 JSON 时也经同一
终局 handler 校验，不接受纯文字声称生成计划，不跳过动作能力检查。
自动化使用假 SDK 验证分发、脱敏边界与拒绝行为；即便本机已有原生 CLI，
没有经授权的真实 API 调用仍不能把这组测试当作模型端到端验收。

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

### WPS 表格开发验证

apps/wps-addin 是独立的 WPS 表格加载项工作区，不使用 Office 清单或 Office.js
初始化。W0 核心加载已在目标个人版通过；W1 提供明确授权的分块读取；W2 接入对话 Agent；
W3 已提供受控写入、回读校验和本机恢复点。

依次运行 npm run test:wps、npm run build:wps 和 npm run dev:wps。dev:wps 调用
固定版本的 wpsjs，将类型为 et、名称为 excel-bro-wps 的开发加载项登记到官方
publish 调试配置，并在 http://127.0.0.1:3889 提供资源。WPS 子项目的 `dev`
脚本不单独声明端口，统一服从 `vite.config.ts` 的 3889 严格端口配置，避免任务窗格
意外从 3890 加载后被后台 CORS 预检拒绝。已有 WPS 进程可能不会
热加载新登记；保存打开的文档并完全重启 WPS 后，再验收功能区与任务窗格。不要为了
测试强制结束含未保存文档的进程。
功能区入口应位于独立的“Excel Bro”选项卡，而不是“开始”内的附加分组；
修改 `ribbon.xml` 后仅刷新任务窗格不会更新功能区，须重启 WPS 验收。
点击“打开 Excel Bro”后应显示任务窗格并切回“开始”；若当前宿主不支持
`ActivateTabMso`，仅跳过切换，不影响任务窗格。
WPS 可能恢复先前的任务窗格而不再次创建；新建后和恢复时均在侧栏可见后设置
1080 的默认宽度，仅迁移仍处于旧窄宽度的侧栏，不覆盖已手动拉宽的尺寸。

开发服务器只输出资源路径和布尔能力的脱敏诊断，不记录工作簿名、工作表名、选区值
或文件路径。首次通过必须记录目标 WPS 完整版本、功能区入口、任务窗格连接状态、
重启行为和停用行为。后续授权读取与只读 Agent 接入按照
[WPS 个人版适配指导](./WPS_PERSONAL_EDITION_ADAPTATION_GUIDE.md) 的 W1/W2 阶段实施。
W1 的自动化与真实宿主验收记录见
[WPS W1 授权读取验证](./WPS_W1_VALIDATION_2026-09-22.md)。读取限制统一配置在
config/capabilities.json 的 wpsRead，调整前应记录目标版本的大表耗时。

W2 开发页从 http://127.0.0.1:3889 访问本地服务；服务端只对白名单中的 3889
localhost/127.0.0.1 来源开放 CORS。只读 Agent 的真实验收见
[WPS W2 验证记录](./WPS_W2_VALIDATION_2026-09-22.md)。WPS 代码可以复用
dataTools.ts 的纯计算入口，但不得调用 executeQueryTableTool 或 Excel.run。

W3 不再要求用户点击“验证读取”。工作表勾选同时构成读取和写入的不可扩大授权边界。
模型返回写入计划后，通常先调用 `/api/plans/validate`；WPS 本地计算出的列映射计划
因验收矩阵含完整源值而不上传此接口，直接进入本地受控预检。随后执行 WPS 白名单、文件身份、
只读、保护、合并、尺寸、重叠与限额检查；只有预览卡中的“确认写入”会触发宿主写入。
`config/capabilities.json` 的 `wpsWrite` 控制开关、宿主版本族、单次单元格上限、分块行数和
动作白名单。开发验收通道当前允许 `writeValues`、同形 A1 公式矩阵 `writeFormulas` 与
`clearRange:contents`；R1C1 铺满公式和结构删除仍未开放。公式写入后用 `Formula` 回读校验，
不能用 `Value2` 计算结果代替；正式构建仍须等待独立实机验收。
实机 WPS 表格 `12.1.0.28488` 的 `Application.Version` 实际报告 `12.0`，并非安装目录的
产品版本。开发调试加载项允许 `developmentVersionPrefixes: ["12.0"]` 通过版本闸，
用于独立测试工作簿的写入、回读和撤销验收；先前基于安装路径误设的
`verifiedVersionPrefixes: ["12.1"]` 已清空，生产构建不使用临时名单。
这不是把所有报告 `12.0` 的 WPS 版本视为已验收。正式开放前必须找到可靠产品版本信号，
或完成可区分版本的能力与实机验收；版本缺失、其他版本及任一预检失败仍拒绝写入。
恢复点暂存于 WPS `PluginStorage`，任务窗格刷新后仍可读取，但关闭加载项后不持久化，
不能作为跨关闭恢复保证；撤销会先核对写入后指纹，冲突时不得覆盖用户后续修改。持久恢复的实施关卡见[双宿主能力对齐指导](./WPS_EXCEL_CAPABILITY_PARITY_GUIDE_2026-09-28.md)。目标版本真实宿主测试清单记录在
[W3-A 验证记录](./WPS_W3_A_VALIDATION_2026-09-23.md)。

**撤销入口的归属（2026-09-30 修正）**：按钮必须渲染在**产生恢复点的那次执行结果卡**内，
按 `executionId` 与执行绑定；不得挂在消息流末尾、不得挂"最新一条回复"下，也不得按计划标题
或消息位置推断。「备份与恢复」抽屉只留快捷入口。
判据是纯函数 `apps/wps-addin/src/wpsExecutionRecovery.ts`：只有当前会话里活跃的恢复点
与卡片 `executionId` 一致且允许撤销时才给可点按钮；每张卡如实区分
**可撤销 / 撤销冲突 / 已撤销 / 恢复点已不可用**。
确有恢复点的结果不止"回读完全通过"——写入成功但计算值与副本不符或无法核实、
以及可恢复的部分写入同样画卡，文案与色调都不冒充"验证通过"；**执行器没给证据才不画卡**。
反向变异基线：忽略 `executionId` 归属 → `WpsAgentPanel.test.tsx` 与
`wpsExecutionRecovery.test.ts` 各 1 条红；无条件渲染按钮 → 4 条红；
把全局撤销卡挂回消息流末尾 → 4 条红。

跨宿主能力开发先改共享 `planCore.ts` 与 `AnalysisPlan` 校验，再接 Office.js/WPS
适配器。当前第一批共用 A1 矩形、写值矩阵形状和清空内容矩阵；对应测试位于
`apps/excel-addin/src/planCore.test.ts`。下一批应把计划预览描述、动作能力协商与
回读验收规则逐步移出各自任务窗格，避免重复开发。WPS JSAPI `12.0` 尚未进入
`verifiedVersionPrefixes`；开发模式的临时放行不能用单元测试代替真实宿主的写入、撤销验收。

### 规划前能力交集（已接线）

两端在规划请求里自报 `hostId` / `hostVersion` / `selfReportedActions`，服务端据此把
未验收动作排除在规划之外。接线点：

| 位置 | 作用 |
| --- | --- |
| `config/capabilities.json` → `hostCapabilities` | 可信上界（三宿主 × 版本 × 动作，三态） |
| `packages/core/src/actionCapability.ts` | TS 侧纯函数（L2） |
| `server/app/action_capability.py` | Python 侧镜像 |
| `server/app/excel_agent.py` → `_enforce_action_capabilities` | 计划出口强制（`submit_plan` 与 `_map_column_plan` 两处） |
| `apps/excel-addin/src/officeActions.ts` | Office 自报清单 |
| `apps/excel-addin/src/officeHostAdapter.ts` → `officeHostRequestFields` | Office 请求字段组装 |
| `apps/wps-addin/src/wpsHostAdapter.ts` → `wpsReportedActionTypes` | WPS 自报集合 |

`config/action_capability_cases.json` 是跨语言共享夹具，TS 与 Python 两侧加载同一份文件
断言同一批值。**这不是** `test_safety_mirror.py` 那种抓源码文本的做法：它抓的是行为，
单边修改实现会立刻有一侧变红（可直接验证：把 TS 的 `matchesVersionPrefix` 换成裸
`startsWith`，TS 侧 2 条失败而 Python 侧 15 条仍绿）。

两条维护红线：

- 版本前缀按"相等或后接 `.`"匹配，不用裸 `startsWith`（否则 `"12.0"` 会命中 `"12.01"`）；
  重叠前缀取最长匹配。
- 新增动作先落 `excel.ts` 的 `executeAction` 分支，再登记 `officeActions.ts`；
  `officeActions.test.ts` 会对 `contracts.ts` 联合体与执行分支双向比对，漏报/虚报都会红。

未接线项：两端 `HostAdapter` 仍无生产消费者（阶段 4 前置，尚未替换 `executeAction`）；
WPS `EB()` 公式的异步语义未核实，不得用表面相似的入口冒充对齐。

WPS 任务窗格直接导入 Excel 端 `styles.css`，产品壳和控件类名应优先与 Excel EB 保持一致；
`apps/wps-addin/src/styles.css` 只添加 WPS 宿主内容和窄窗适配，不能再维护第二套主题。页面至少
保持模型、新对话、工具、备份、格仔、更多菜单、历史、范围选择、消息和输入区一致。

**输入区的「上拉」只有一份实现**：`apps/excel-addin/src/hooks/useComposerResize.ts`。
拖动上方把手调高度、内容变多时自动长高，到上限后**改为可滚动**（不是把文字切掉）。
两端都挂同一份共享样式（`.composer-input-region` + `.composer-resize-handle`），
把手必须待在该 region 的上内边距里、位于 `.composer-box` **之外**——挪进输入框内部会
盖住文本区抢走光标。**不许再用 CSS 原生 `resize: vertical`**：WPS 的 WebView 里那个把手
拖不动，而且和继承来的 `overflow-y: hidden` 叠加会把超出部分切掉（这正是"输入框不能上拉、
也看不到后面的字"的根因）。改这块要同时跑：
`npm run test:wps`（`WpsAgentPanel.test.tsx` 的 DOM 断言）、
`apps/excel-addin` 的 `useComposerResize.test.tsx`（纯逻辑，含"超过上限必须 auto"）与
`composerResizeParity.test.ts`（契约：两端同源、把手位置、CSS 不再用原生 resize）。
反向变异基线：把溢出改回一律 `hidden` → 上面两条"可滚动"用例各红一条。
WPS 原生并排审阅已按用户 2026-09-30 修正的交互接线：左原表、右独立副本、
任务窗格保留对话。预览后点「生成右侧副本并排审阅」只写副本并标色；点
「确认写回原表」前重新核对原表与副本，才走原有白名单计划执行器；点
「放弃副本」关闭副本且不保存。当前只支持单张已授权工作表，多表计划明确拒绝。
**侧栏只负责状态、摘要、风险与决策**：长状态说明必须留在
`.wps-native-review-status`，操作区 `.wps-native-review-actions` **只放按钮** ——
混在同一个单行 flex 里会把按钮压到逐字折行（真机竖排的直接原因）。
按钮 `white-space: nowrap` + `min-height: 44px`；≥560 CSS px 水平、更窄整行纵排。
逐格公式默认收进 `<details>` 并用等宽字横向滚动，不重引入窄 HTML 双网格。
区分时**必须以计划里真实的 `action.type` 为判据**：`writeFormulas` 才说"已试算公式并回读
计算值"，否则说"已写入固定值"；"预计（尚未写入）"与"副本回读"是两个不同的入口和标签，
不得互相冒充，也不得在试算失败时给出可写回的主按钮。
**jsdom 不计算布局**：DOM 层的绿灯（按钮在、类名对、`disabled` 生效）与 CSS 声明契约
都不能替代真机视口验收，真机 350/430/700 CSS px 的截图必须另报。
本地列映射的 WPS 副本审阅采用公式优先：仅同表、非空文本且满足
`columnMappingFormula.maxCases/maxCharactersPerCell` 的规则生成候选公式；
副本回读的 `Formula` 与 `Value2` 必须分别等于计划公式和本地计算值。
公式不兼容或值不符时关闭试写副本，改以固定值再生成待审阅副本；用户仍需
明确确认写回。未知新组合公式返回 `#N/A`，不要承诺能自动处理所有未来组合。
Excel 当前工作簿与文件夹模式目前仍使用固定值，不能写成三端公式对齐。

**`valueRank.rank` 的方向必须在 schema 里讲死**。core 约定 rank 越大越优先，
而业务等级名里的数字常与之**反向**（名字里最"0"的最严重）。模型照抄字面数字不会报错，
只会让整条链路一致地取到**较低**的等级。改 `map_column_values` 的 schema 或系统提示时，
"用户说更严重/更优先的等级给更大的 rank"这句和那个**抽象**例子都不能删；
`server/tests/test_map_column_tool.py` 有正向锚点 + "生产提示词不得出现具体业务等级
字面量"的负向断言。极性逻辑本身由 `column_mapping_cases.json` 的
"等级优先级由规则给出：反过来声明则取另一个"跨语言覆盖，两侧各跑一遍。
**自动化通过不等于 WPS 真机布局已通过**：须用独立测试 `.xlsx` 按
[原生并排审阅验收清单](./WPS_NATIVE_SIDE_BY_SIDE_REVIEW_2026-09-30.md)
检查左右顺序、底色、同步滚动、任务窗格、保存/放弃、失败时未写原表与撤销。
旧的 Web 宽窗/任务窗格内全幅探针仅留作历史探索，不是当前 WPS 产品路径。
2026-09-30 真机首次点击原生并排入口时，`ActiveWorkbook === 预览时引用` 在复制前误拒。
并排审阅对**已保存**原表改用工作簿名、完整路径、目录与计划快照复核；副本记录创建时
的名称与未保存状态，后续若被保存成文件须拒绝写回。测试替身应覆盖同一工作簿每次返回
不同 JSAPI 包装对象及同名异路径拒绝。修正后的完整窗口流程仍需按上述清单实机复测。
过程展示也复用共用组件：WPS 本地读写进度按步骤写进消息历史并交给
`TaskProgressTimeline` 展示，刷新后未完成记录必须显示中断；只允许记录安全的动作、目标、
行数和回读状态，不能把原始单元格值写入历史。复制按钮应覆盖用户与助手消息，
并验证 WPS WebView 拒绝 `navigator.clipboard` 时的降级路径。当前尚未接入服务端
规划 SSE 的完整事件时间线，不可把本地步骤误称为 Excel 端全部进度能力已对齐。
逻辑复用不得突破宿主能力：工作目录与图片输入暂未开放；工作流仍须验证并预览；固化查询必须在当前
授权表内由本地确定性执行器运行。自动范围在每次发送时重新读取活动表，手动多表授权不得自动
增加工作表。修改这些区域后至少运行 `npm run test:wps` 与 `npm run build:wps`，并在目标 WPS
版本分别检查窄窗和加宽窗格。

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

### 任务中断后继续调试（第一阶段）

用户任务因断网、服务重启、任务窗格重载而中断时，不应该让他重新输入需求或重新回答
同一个问题。共享口径在 `packages/core/src/taskResume.ts`，先读它再改两端。

- **阶段 / 阻塞原因 / 可续跑性分开表达**：`phase`（走到哪一步）、`blocker`（为什么停）、
  `resumability`（下一步能安全做什么）、`writeState`（原表写入状态）是四个独立维度。
  不要新增"一个字段同时表达停在哪和能不能继续"的快捷方式，那会立刻退回
  "网络中断 / 等待回答 / 需重新授权 / 结果不确定"混为一谈的老问题。
- **只有 `checkpoint` 能说"从上次保存的进度继续"**：`taskResumeDecision()` 是唯一判定
  入口。改它的文案时不要删掉"重新生成预览"这一类说法，那是对"没有检查点"的诚实交代。
- **持久化必须走 `sanitizeTaskResumeRecord()`**：它是显式白名单重建，多余键一律丢弃。
  新增字段要同时加进白名单，否则字段会被静默丢弃（不会报错）；`originalPrompt` 为空时
  返回 `null` 是**故意**的 —— 缺原需求时宁可明确报"无法恢复"。
- **`authorizationRevision` 只是版本，不是许可凭据**：恢复时必须重新读取真实宿主元数据。
  不要把它当"已经授权过"的通行证。
- **执行身份不得翻新**：继续动作只能复用既有 `operationId` / `executionId`，重试、
  状态查询、断线重连都不许新建 —— `officeOperationRecovery.ts` / `folderOperationRecovery.ts`
  里保存的只是**发现索引**，不是权威状态。

断点位置（改任一处都要跑 WPS 与 Excel 两端对应用例）：

| 位置 | 作用 |
| --- | --- |
| `packages/core/src/taskResume.ts` | 阶段/阻塞/可续跑性的共享语义、白名单消毒、可续跑判定 |
| `apps/wps-addin/src/WpsAgentPanel.tsx` → `TaskResumeRecoveryCard` / `resumeInterruptedTask` / `cancelInterruptedTask` | WPS 中断卡与重新授权闭环 |
| `apps/wps-addin/src/WpsProbePanel.tsx` | 范围勾选=草稿，「确认范围」才提交并重读宿主元数据 |
| `apps/excel-addin/src/hooks/useScopeSelection.ts` | `committedSheetNames` / `scopeDraftDirty` / `confirmScopeSelection` / `cancelScopeEdit` |
| `apps/wps-addin/src/wpsTaskResume.test.tsx` | 重载后可继续、答案不丢、取消是终态、写入状态未知、隐私边界 |
| `apps/wps-addin/src/wpsScopeDraft.test.tsx` | 勾选只改草稿、确认范围重读元数据 |
| `apps/excel-addin/src/hooks/useScopeSelection.test.ts` | 草稿/已提交分离 |
| `packages/core/src/taskResume.test.ts` | 白名单消毒、诚实文案、任务身份推进 |

排查"重开后任务不见了"时先看三件事：① 消息上有没有 `taskResume`（没有就说明建卡那条
路径没挂记录）；② 它的 `resumability` 是不是被用户取消成了 `unavailable`
（取消是终态，这是预期行为）；③ `sanitizeTaskResumeRecord` 是不是返回了 `null`
（缺原需求或 schemaVersion 不认识时会整体丢弃，这是 fail-closed 而不是 bug）。
**`localStorage` 里的记录只是恢复线索**，真实执行状态仍以服务端账本为准，
不要用本地记录反推"写没写成功"。

两条维护红线：**恢复卡不能挂在消息流末尾或"最新一条回复"下** —— 它必须渲染在
产生它的那张卡内（与撤销入口同一纪律）；**`writeState` 未知时不得显示"原表未修改"**。

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
| 确定性地址命令和模型选择 | `server/app/planner.py`, `apps/excel-addin/src/modelSelection.ts` |
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

普通规划首轮不再携带全部动作 schema。通用计划先调用 `select_plan_actions` 选择
动作类型，再提交裁剪后的参数；列映射直接提交规则并由客户端计算，不先读数据。
检查点必须保留 selectedActionTypes。完整计划出口校验及预览确认不可裁剪。
回归包含映射工具载荷体积、按类型 schema、恢复后相同工具集合、未知宿主/动作拒绝、
无数据策略下拒绝未公开读取工具、截断响应不得执行、SSE 停止状态落盘及两端失败显示。
`answer.outcome` 缺省兼容 completed；预算熔断必须 failed，并带 stopReason。
SDK 与兼容 API 的预算统计不同，诊断日志只打印 finish_reason 的固定枚举及 usage 计数；
不能打印模型原始响应、提示词、工具参数或数据行。宿主与真实模型重测另报，不用替身绿灯替代。

普通问候和能力咨询应返回 `conversation` 直接答复；只有任务有关键歧义才显示澄清卡。
WPS 意图阶段只读已授权表的候选表头，查询确定后才读取完整数据行。

这条"只读候选表头"的判据要落在**真实取值范围**上：`wpsWorkbookReader.readChunkRange`
按 `purpose` 构造不同的取值面 —— `"schema"` 只取第 1 行 × 前 `structureColumns` 列，
`"values"` 才取完整授权矩形。报告给调用方的行列数仍是整块已用区域。
新增或修改读取路径时必须在 `wpsWorkbookReader.test.ts` 里按替身宿主真实的
`Value2` 访问坐标记账，断言结构扫描**没有取到任何数据行的值**；
不要用"这次读取的目的是 schema"来充当证据（那是调用方的自我声明）。
`chunkRows` 在结构扫描路径上不参与取值面计算（表头扫描始终只有一次整行读取），
它只控制数据读取的分块。

**`purpose` 绝不传给宿主**：真宿主的 `Worksheet.Range` 只接受 `cell1`/`cell2` 两个
参数，多传抛 `Too many parameters.`（2026-09-30 真机证据）。替身是 JS 函数，多传的
实参会被静默忽略 —— 这正是"本地全绿、真机爆炸"的根因。替身的 `Range` 已内置
实参个数检查（多传即抛同一错误，反向变异实测 12 条红）；给任何宿主调用添加
自有标签/旁路实参前，先确认真宿主签名，并在替身里模拟同样的严格性。

澄清卡与本地映射计划的两条补充约束：

WPS 本地列映射若在**读源值后**发现未覆盖值、未声明的分隔形式，或**真的有空白格而空白
怎么处理还没定**，应停在本地澄清卡，而非将带原始来源值的 `UnmappedCategoryError` /
`BlankPolicyNeededError` 写成系统消息。三类缺口按 `reason` 给三组互斥选项，空白那一支
必须发生在真实读取之后：`blankCount` 要来自真读到的格数，**没读到就不许声称有空白**。
未启用拆分时用户可为每个
整格明确指定目标；已有拆分规则但缺少**拆分项**时必须回输入框补全映射与优先级，
不能让新增映射的声明顺序偷偷决定结果。卡片内真实来源值及目标草稿只存 React 内存，
落盘消息只存数量和状态。提交后必须复核授权身份/表头/对象，再读当前源值、重新算
预览；新增缺口继续问，不能拿上次源快照兜底，更不能把补充映射或计划矩阵送到
`/api/turn` / `/api/plans/validate`。检查 `wpsLocalPlanOutbound.test.tsx` 的真实
`fetch` 字节、localStorage 负向断言以及改源后再次询问用例；最后做 WPS 真机验收。

1. 待答澄清绑定**授权身份**（工作簿名 + 已授权表的索引与名称，含顺序）与
   **表头指纹**，不绑定工作表索引。任一项变化都要作废澄清并把原需求退回输入框。
   指纹由 `wpsClarificationScope.ts` 生成，只含表名与表头，不含单元格值。
   **两条判据必须读不同的数据源**：身份读工作表元数据（`listWpsWorksheets`），
   表头必须读第 1 行的真实单元格值（`WpsTurnResult.headerRows` /
   复核时经 `readAuthorizedWpsSheetSchemas` 重读）。若表头判据改回读工作表
   元数据，它就退化成身份判据的副本 —— 只禁用身份判据时它会顶上来把用例撑绿，
   "身份判据没必要"的假象随之出现。   改这两条判据时，必须用**单判据变异**复验
   （只禁身份 → `wpsClarificationIdentityNecessity.test.tsx` 应有 4 条红；
   只禁表头 → 应有 1 条红），**不得同时禁用两条来充当证据**。
   **作答前只能有三种结果**（`wpsClarificationVerdict.ts` 是唯一判定入口）：
   `same` 才放行；`changed`（有正面证据：路径/表名变了、索引越界、工作簿关掉、
   真的读到不同表头）→ `invalidated` + "数据范围已经变化"；`unknown`
   （引用换了但无其它证据、读到失败、面板刚重开）→ `unconfirmed`：
   **不谎称数据已变化**，保留原需求与用户刚填的回答，请用户重新确认授权后
   作为**新一轮**重新读取与预览。
   **"引用不相等"属于 `unknown` 而不是 `changed`** —— 宿主可能对同一张表返回新包装
   对象（改一格又撤回就会触发），说成"数据范围已经变化"是假话。
   **指纹读到 `null` 不算变化证据**（读不到 ≠ 变了），这是此前另一个假话来源。
   反向变异：把 `unknown` 退回旧的 `changed` → 应有 **8 条红**。
   另：指纹只有工作簿**名**、没有路径，所以**单靠指纹**区分不了不同目录的同名
   工作簿，也区分不了"同索引、同名、同表头删除后重建"的工作表。这两类由**第三条
   判据**关闭：建卡时用 `captureWpsObjectIdentity` 在**内存**记下工作簿与每张已授权
   表的**对象引用**，作答前用 `verifyWpsObjectIdentity` 逐项复核（`wpsObjectIdentity.ts`）。
   纪律是**放行需要正面证明，不是"没发现不同"**：引用对不上就判失效，不存在用
   名称/表头/尺寸/稳定 ID 反向补证的路径（它们在"删除后重建""整表复制"里都能保持
   一致）。稳定 ID（`CodeName` / `SId`）与尺寸只作**加固**——引用已相等时再核一遍，
   防对象池复用；读不到就是 `undefined`，**不得凭空造 ID**。
   路径**只在内存比对**，不得写进 localStorage、请求体、日志或错误文案。
   面板重开由**独立**的读取路径判据保障（`loadWpsChatHistory` 把 `pending` 收敛为
   `invalidated`），不依赖对象身份。改动任一条判据都必须做**单判据反向变异**：
   只禁对象身份 → `wpsClarificationCollision.test.tsx` 应有 2 条红，且**必须正是**
   "不同目录同名工作簿"与"同索引同名同表头重建"这两条；只禁表头 → 1 条红；
   只禁重开失效 → 1 条红；复原后全绿。**不得同时禁用多条来充当证据。**
   体验代价：若真宿主每次 `Item()` 返回新包装对象，合法续答也会被拒——这是显式
   接受的 fail-closed 代价，属**实机待验项**，不得为用例好看而放宽。
2. "本地读取"与"向模型/服务端发送"是两个独立的证据面，要分别取证：
   前者看宿主的 `Range`/`Value2` 访问范围，后者在 `fetch` 请求体层面断言
   源列值与目标矩阵都不出站。
3. **重新授权不是澄清答案**。点"重新确认数据范围"只推进授权版本
   （`authorizationRevision`），**不得**往 `clarification.answer` 里写任何套话：
   有真实答案就原样保留并推进到 `resolved`，没有就落到 `continued`
   （`clarificationAfterReauthorization` 是唯一入口），说明文案走
   `REAUTHORIZATION_NOTICE` 进消息流。用户**改过**的答案必须**替换**旧确认
   （`questionId` + `mergeKeyedConfirmation`），追加会让模型同时看到两条冲突确认。
   反向变异：丢掉真实答案 → 2 条红；无答案时伪造"已回答" → 1 条红。
4. **空白策略不许预先拦截**。`blankLabel` 不再是必填；`undecided` 只在真读到
   空白格时才升级成一次澄清（带真实 `blankCount`）。`keep` 的语义是**把该位置
   从写入范围里排除**（`splitTargetActions` 拆连续段），不是读回原值再写回。
   反向变异：`blankCount > 0` 改成 `>= 0` → 3 条红；`keep` 改成写空值 → 3 条红。
   四类意图（保留 / 清空 / 写指定值 / 写空串）在 UI 上是四个互斥选项，选了
   "写指定值"却没填值时提交按钮必须保持不可用——实现不替用户挑默认标签。

1. 判断这是语义决策还是确定性计算。
2. 语义交给模型；计算尽量写成本地受控工具。
3. 不向意图模型发送原始数据行。
4. 有关键歧义时提供互斥选项，并保留“其他想法”。
5. 写入仍然只返回计划，不直接执行。
6. 为结构化输出失败设置有限修复次数，禁止无限重试。
7. 澄清选项若改变拆分维度、输出结构、指标口径或目标，必须返回 replace 模式和
   可独立执行的完整 resolution；只补充缺失参数时才使用 append。
8. 计划格式修正轮不得继续携带完整失败参数，也不得重新开放只读工具；输出额度
   使用 agent.repairMaxOutputTokens（默认 null，省略 max_tokens，使用供应商额度）。
   intentContext.maxOutputTokens、agent.maxOutputTokens 同样默认 null；只有正整数
   才设置显式上限。不要把 length 归为格式错误并改用更小的额度重试。
   agent.maxTotalTokens 默认 null，关闭普通累计 token 预算停止，计数和恢复检查点
   仍保留；正整数可以重新开启（下一模型轮前检查）。LoopConfig 默认 None，同样
   支持显式正整数；0、负数、布尔或非法类型报配置错误，不将坏配置静默视为无限。
   单次 HTTP 默认 180 秒，环境覆盖仍优先。SDK 1 美元预算未动，需单独授权取消。
   这些设置需要后端重启；客户端刷新不能重载缓存的能力配置。
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
- `hooks/useComposerResize.test.tsx`：输入框上拉（自动长高、超上限改 `overflow-y: auto`、
  拖动/键盘 clamp、双击回自动、pointerId 串扰）；需要 `// @vitest-environment jsdom`
- `composerResizeParity.test.ts`：两端输入区上拉的契约（同引共享 hook、把手在 region 内、
  文案来自共享常量、WPS 不再用原生 `resize: vertical`）

后端测试：

- `test_main.py`：API 和错误协议
- `test_task_ledger.py`：持久任务身份、事件重放、异常退出和过期墓碑；另含 WPS 执行
  账本（`wps-path:` 身份命名空间、taskId 绑定、跨前缀拒绝）
- `test_wps_execution_api.py`：CAP-03 WPS 执行路由。register 授予 ownerToken 且
  **只回 descriptor 不回 plan**；跨前缀 / 无前缀身份 422；描述里出现矩阵字段 422；
  `clearRange` 非 contents 422；旧 revision + 旧 ownerToken 409。
  其中一条**直读 SQLite `office_executions.plan_json`** 断言哨兵不在库里。
  另覆盖执行权仍活跃拒绝接管、重复登记不泄露已有 token、Office 路由不得绕过 WPS 守卫、
  在途核对须精确检查点、最终回执丢失、剩余动作未完成不得补记、查询无记录不新建任务、
  各回执矩阵拒收及 SQLite 原始记录取证、旧敏感记录拒绝并且不回显值。
  本轮独立变异：描述 `extra="forbid"` 改 ignore → 1 红；去掉事务内执行权检查 → 1 红；
  去掉完成 SQL 的剩余动作守卫 → 1 红。
- `test_folder_mutex.py`：同进程及跨进程文件夹写入排他
- `test_managed_directories.py`、`test_managed_recovery.py`：受管目录发现和独立恢复
- `test_models.py`：Pydantic 模型
- `test_planner.py`：确定性地址命令与规划；无可用模型时抛稳定错误而非降级
- `test_agent_loop_protection.py`：Agent 预算、循环检测、降级与收尾提示
- `test_agent_integration.py`：Agent 主循环熔断和警告注入
- `test_folder_workbooks.py`：文件夹执行
- `test_folder_data.py`：pandas 完整数据、跨文件 ID、权限边界、合并和关联
- `test_no_business_hardcoding.py`：防止行业字段重新硬编码

WPS 侧与"读取边界 / 授权身份"直接相关的用例：

- `wpsWorkbookReader.test.ts`：分块读取、授权校验，以及**按真实 `Value2` 访问坐标**
  断言结构扫描不取数据行
- `wpsClarificationScope.test.ts`：授权身份指纹与表头指纹（改名 / 重排 / 换表 / 换簿）
- `wpsClarificationIdentityNecessity.test.tsx`：**两条判据各自的必要性**。
  替身的表头/尺寸/数据值全部相同，只有身份可变；配单判据变异使用
  （只禁身份 → 4 条红，只禁表头 → 1 条红，两条都开 → 6 全绿）
- `wpsObjectIdentity.test.ts`：**内存对象身份判据本身**（19 条）。替身每次
  打开/重建都返回**全新对象**（贴合真宿主"重建 = 新对象"语义），`Worksheets.Count`
  必须是 getter 否则"删表后数量不变"、越界判据测不到。覆盖替换工作簿、只改路径、
  重建表、无稳定 ID 时重建、引用没变但 CodeName 被换、只改尺寸、删表、关簿、
  未保存同会话可证明、宿主抛错判 unknown
- `wpsClarificationCollision.test.tsx`：**两类碰撞必须被识破 + 重生路径**（7 条）。
  不再断言"认不出来"，而是断言"必须失效"：① 不同目录的同名工作簿；② 同索引、
  同名、同表头删除后重建。另含指纹兜底（重排 / 改名）、重开后 `pending` 直接
  `invalidated`、未保存工作簿同会话内**不得误杀**、以及落盘历史不含完整路径 /
  稳定 ID / 数据行
- `WpsAgentPanel.test.tsx`：待答澄清在身份变化时失效，身份未变时可正常继续
- `wpsLocalPlanOutbound.test.tsx`：本地映射计划的真实出站面（澄清档零读取、
  算得档必须读，两者都不把源值与目标矩阵发出去）；另覆盖“改为说明拆分规则”
  留在当前卡片、用户确认优先顺序后本地重读并预览，且没有新增模型请求
- `apps/excel-addin/src/localMappingResolution.test.ts` 与
  `components/LocalMappingResolution.test.tsx`：Excel/WPS 共用的结构化补全纯函数及
  DOM 交互；Excel 侧重开历史时，只有内存草稿的待答卡必须失效。
  注意它们**不是** Excel `App.tsx` 整棵挂载或 Office 宿主实测的替代证据
- `wpsExecutionRecovery.test.ts`：撤销按钮归属与恢复点状态的**纯判定**
  （executionId 绑定、四种状态、三种执行结果文案互不冒充）
- `wpsExecutionReconcile.test.ts`：CAP-03 重载核对的**纯判定**（23 条）。身份完整字符串
  比较、canResume、剩余动作为空 / 非空、在途动作 + 本机现场四类、在途缺动作号；
  每条都同时断言 `wpsReconcileActions` 给出的按钮集。当前独立变异：执行权 → 1 红；
  `canResume` → 2 红；身份 → 1 红；现场冲突 → 3 红；未知现场 → 2 红；零剩余分支 → 2 红。
- `wpsRecoveryInspection.test.ts`（13 条）：真实登记坐标读取与当前授权边界，先验证全部
  动作再读任何值，已完成/剩余/在途分别核对，缺采集与空白不混淆，公式文本不能假匹配。
- `wpsRecoveryPrivacy.test.ts`：CAP-03 **隐私边界的真实出站证据**（10 条）。先证明哨兵
  确实在本机计划里（防空过），再 stub 掉 `fetch` 截获**真实请求体**断言不含哨兵、
  不含 `plan`、键集合固定；真实覆盖登记与全部回执/异常通道，额外顶层计划丢弃；
  证据内部矩阵和合法键 `cellCount` 中的矩阵均在 fetch 前拒绝。
  当前独立变异：请求投影改回 spread → 1 红；写前证据检查删除 → 1 红；计数值类型检查删除 → 1 红。
- `wpsTaskResume.test.tsx`：任务恢复卡的重载语义（待答重载、已提交答案、取消终态、
  写入状态未知/待核对**只允许先核对**、已完成任务不再出现继续入口、隐私边界）
- `WpsAgentPanel.test.tsx`：撤销入口在**对应执行结果卡**内、追加普通回复后仍在旧卡、
  连续两次执行只有活跃恢复点可点、撤销冲突 / 成功撤销 / 面板重开 / 新建对话各态；
  并证明消息流里不存在没有卡片容器的裸撤销按钮。
  另含 CAP-03 的恢复入口点击链路：点真实按钮后断言真实 API 调用与
  **宿主写入次数** —— 补记完成走 `resume` + `/complete` 且 `host.writes() === 0`、
  重新预览走 `runTurn` 且写入 0、canResume=false / 身份不一致 / 现场冲突 /
  缺写后指纹仅给清除提示与只读重新核对，不提交任何回执；领取后再次读取现场；
  最后动作回执丢失提交精确指纹后只补记；重新预览须确认当前范围且规划失败保留索引。
  当前独立变异：点击时判定守卫删除 → 3 红；领取后再次现场检查删除 → 1 红。
- `wpsTaskResume.test.tsx`：**任务中断后继续**（6 条）。重载后任务仍在且给出
  「确认此范围并继续」；不谎称"数据范围已经变化"；三件事（中断原因/已保留信息/
  写入状态）分开展示并摆出当前文件与待授权工作表；取消是终态、重开不再兜售；
  写入状态未知时**不出现**"原表未写入"；续答断线后答案跨重载存活、继续时随原需求
  一起重发；恢复记录落盘的键集**恰好**是白名单那一组。
  反向变异基线：只让 `resumableRecord` 恒返回 `null` → 6 条中 4 条红（另 2 条不依赖恢复卡）。
- `wpsScopeDraft.test.tsx`：**范围勾选只改草稿**（5 条）。取消勾选再选回时已提交范围
  一字未动；「确认范围」才提交，且**提交时重新读取真实宿主元数据**（弹层开着时消失的
  表不会被提交成授权）；「放弃修改」与关闭不改已提交授权、重开时草稿回到已提交范围；
  手动模式下没有勾选时不能确认。
  反向变异基线：把 `toggleSheet` 写回已提交范围 → 5 条中 4 条红。

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
