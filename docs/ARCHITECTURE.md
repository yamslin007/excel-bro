# Excel Bro 架构文档

## 2026-10-07：易失结果登记与双端本地预览准备（最新状态，产品仍关闭）

内部 `run_registered_result_internal` 把真实容器的严格矩阵结果交给同一易失授权登记器，发布随机 resultHandle、原始 JSON 字节 SHA256 和矩阵尺寸，不返回原始值。结果绑定原授权及有序数据句柄/指纹；替换源快照删除依赖结果，撤销/重授权/任务 TTL/服务重启失效。数量和编码总字节限额在 `pythonResult`。没有持久结果库或模型工具成功通道。

新增本机 `/api/python/data/result` 与 `/result/discard`，复用回环、Host、Origin/hostId、总开关、严格 JSON 和有界请求守卫，固定错误文案、no-store。协议同步 core/Excel contracts/models.py，不扩展 PlanRequest。core 在本机读取并核对实际响应字节指纹，读取前后重读源值；结果过期不冒充文件变化。仅接受有界、规则、有限安全标量矩阵；共享编译器只生成已有 `writeValues` + `rangeEquals`，目标由可信用户流程给定，限定已登记授权表的精确 A1 范围。拒绝疑似公式/宿主文本转义前缀，不让输出携带坐标/脚本/动作。

Office/WPS 已有只读预览准备器：核对原数据会话绑定的宿主、源快照、计划不可变及目标原值/公式；Office 调用本地 preflight 并按授权 sheet ID 读取目标，WPS 复用 prepareWpsPlan。均不调用执行器、原生副本或持久账本。Office 完整可写性检查/原生副本、WPS 原生副本及两端确认写回仍需产品接线；不得直接把含结果矩阵的计划传给旧远端 plan_validate、聊天历史或持久执行/撤销记录。详细双端差距、证据与下一批门槛见 [实施记录](./PYTHON_RESULT_PREVIEW_2026-10-07.md)。`pythonExecution.enabled=false`；SDK 内部转录限制、文件夹适配及安装包未实现。

以下为前一阶段快照，当前分层状态以上节为准。

## 2026-10-07：内核门禁与授权数据接线（前一阶段，产品仍关闭）

worker 在生成代码之前安装不可撤销的 Linux/amd64 seccomp 附加过滤，拒绝 exec/派生进程/线程/socket，安装失败不执行。真实新测试 5 项通过，原有 15 项保持通过。独立 `/api/python/data/{scope,snapshot,revoke}` 仅接收已允许的本机加载项请求；易失登记绑定任务/文件 hash/授权版本/精确区域，随机句柄、撤销/TTL/大小限制，不进模型、历史或账本。两端使用同一 core 客户端，各自复用宿主强身份与只读原语；新协议未扩展 PlanRequest。

本批仅接通内部登记数据→容器分析。产品工具、结果库/AnalysisPlan 编译、双端 UI/生命周期、文件夹入口及安装包未接线；`pythonExecution.enabled=false`。SDK 内部转录限制仍存在。新镜像 RepoDigest 和当前能力证据见 [授权数据实施记录](./PYTHON_AUTHORIZED_DATA_2026-10-07.md)。下节是此前运行时阶段快照，其中“禁 exec/句柄未实现”不再作为当前状态。

## 2026-10-07：Docker 本机运行时验收（产品入口仍关闭）

用户启动 Docker Desktop 并授权构建后，本机 Linux 引擎、专用镜像构建及 15 项真实合成数据测试已跑通。当前不再处于“本机引擎不可用”的阻断状态；旧记录仅代表此前阶段，不据此推断 WSL 或代理策略的因果。实际镜像身份与证据见 [本机验收记录](./DOCKER_PYTHON_REAL_VALIDATION_2026-10-07.md)。

新增可信构建入口 `scripts/build_python_sandbox.py`：仅把 `.dockerignore`、`docker/Dockerfile`、`app/docker_worker.py` 三个明确文件打包为 stdin tar；不遍历/发送仓库。使用临时无凭据 Docker CLI 配置，显式固定本机端点与 daemon-bound `default`/`docker` builder，拒绝其它 driver、多节点或远端 endpoint；基础镜像先解析真实 RepoDigest 再构建，前后核对引擎身份。构建镜像不启用产品功能，不改用户全局 context，不用本地 image ID 冒充 RepoDigest。

运行器新增固定容器线程环境（OpenBLAS/OpenMP/MKL/NumExpr 各为 1），使 numpy/pandas 在既有单进程限额内可用，不扩大 PID 配额；输出读取用 `read1` 及时交付短输出，不等满 4096 字节或 EOF。已用真实管道先红后绿及容器内运行/取消验证。

`pythonExecution.enabled=false` 保持不变。完全禁止容器内 exec、datasetHandle 授权登记、结果句柄、结果编译到 `AnalysisPlan`、Excel/WPS/文件夹入口及安装包仍未实现；SDK 内部转录投影限制仍存在。15 项通过不代表完整产品 Python 或全部威胁边界已经完成。

## 2026-10-06：隔离 Python 分析的设计授权（历史阶段记录）

以下“当前状态/环境阻断”描述为当时快照；2026-10-07 构建及真实验证的最新状态以本节之前的本机运行时验收为准。

用户批准的例外是模型可编写 Python，在经过验证的隔离执行器中分析授权数据，不是开放 SDK Bash/Read 或宿主任意脚本。模型选用普通后端或 SDK 不改变数据/写入授权；共用受控工具与结果协议，Excel/WPS 通过各自宿主适配读取当前授权范围，文件夹模式通过明确的文件/工作表 ID 读取。不默认复制整表给模型或远程服务。

计划新增 `run_python_analysis`：输入任务内数据句柄与代码，隔离工作进程无网络、宿主目录/工作簿句柄、凭据和外部程序权限；有资源配额、取消及整个执行树终止。隔离缺失则拒绝，不能用 cwd、venv、AST 检查或禁词当沙箱，也不能退回同用户裸执行。输出矩阵只在本地，模型出口由可信投影器控制；代码产生的摘要、stdout、错误均不能默认发送。结果经本地适配编译到既有 `AnalysisPlan`、预览、确认、回读链路，Python 本身不写工作簿。

当前状态（2026-10-07 八批）：失败关闭骨架与**隐私/入参闭环已落地并补修多轮绕过路径**，并已新增 **Docker 运行时适配器（代码落地并修复剩余缺陷，真实隔离实测受环境阻断）**。受控工具 `run_python_analysis` 由 `python_analysis.py` 失败关闭协调器承载，出站走 `NON_TERMINAL_TOOL_PROJECTIONS` 有类型投影。多批形成的隐私边界：① 历史投影是可信侧最小白名单（`code`→固定占位 + `datasetHandles`→空列表，句柄登记未实现故不保留模型自报的任何句柄原值）；② 恢复现场统一投影与归属核对（孤立 ID/名称冲突/关联歧义失败关闭拒绝恢复）；③ 执行路径异常统一安全映射（普通后端与 SDK 同源固定回执，不回显原始异常）；④ loopState 收口（涉及 Python 时清空错误/警告自由文本）。**运行时层**：新增三层分离（协调器 `python_analysis.py` → 可替换运行时适配器 `python_docker_runner.py` → 容器内固定工作程序 `docker_worker.py`）。Docker 是第一种运行时实现，不是项目强制启动依赖；适配器可替换（未来可加 AppContainer 等）。固定启动链由可信代码决定（`--network none`/只读根/非 root/`--cap-drop ALL`/禁提权/CPU·内存·swap·pids 限额/`--log-driver none`），模型只提供 code 经 stdin（绝不进命令行、不用 pickle、不拼 shell、不挂载用户目录/仓库/工作簿/凭据/socket）。**剩余缺陷已修复并各配红灯测试与独立反向变异（第八批 B1–B6）**：① **输出运行期有界**（不是"有界临时文件"）——stdout/stderr 由读取线程并发收集、分通道限字节，任一通道触顶即**中途终止执行**并进入容器终止核实，协议结果/纯 stdout/stderr 预算分离，保留 `--log-driver none`；② **清理可核实 + 真实取消**（不是"清理闭环已完成"）——绑定可信侧运行身份与精确容器身份，`rm -f` 后 `container inspect` 区分"确认已消失"/"无法查询"，清理失败或状态未知**不冒充成功**（不发布结果、不复用沙箱），解决启动/取消竞态，取消后到达的迟到结果被拒收，只清理本次运行拥有的容器（禁止全局 prune/宽前缀删除）；③ **核实并锁定真实本地端点**（不是"剔除变量即绑定本机"，也不是"判 scheme 前缀"）——`docker context inspect` + `docker info` 实测端点，只接受**经结构校验的本机形态**（Windows 命名管道 `//./pipe/<name>` 或扩展长度前缀 `//?/pipe/<name>`；Unix socket 本机绝对路径），拒绝 `tcp://`/`ssh://`/`fd`/未知、远端主机段命名管道（如 `npipe:////remote-host/pipe/...`）、带远端 authority 的 unix socket、空/畸形地址，**不把任意主机别名改写为 local**；整轮锁定到已核实端点，不改用户全局 context，不把 `.OSType=linux` 当本机身份证明；端点解析取结构化字段 `Endpoints.docker.Host`（不能只取第一个 URL 匹配，否则会误取 otel 端点的 `unix://`）；执行路径不按对象名信任 `VerifiedEndpoint`，入口与 CLI 最内层各做一次结构校验；**运行期端点核实 ≠ builder 核实**（buildx 远端 builder 由构建流程单独核实，不因 runner 的 `-H` 锁定而视为安全）；④ **收紧配置/镜像/结果校验**——非 root 校验拒绝 `0:1000`/`root:1000`/`00:0` 等，移除错误的 `maxProcesses + 1` 接线，CPU 速率与累计 CPU 时间分开实现（不可实现则失败关闭），PID 限额不再冒充"禁止 exec"，结果字节/类型/深度/元素数边界 + 拒 NaN/Infinity/畸形 UTF-8/多包协议，区分合法 `result=null` 与缺失 `result`，绑定已核实不可变镜像身份（`imageDigest=null` 时失败关闭，不静默降级为可变 tag）；⑤ **最小化构建上下文**——`server/.dockerignore` 白名单，排除 `.env`/`model-connections.json`/凭据/测试/缓存/账本/工作簿/无关源码，核实真正进入上下文的文件集；⑥ **修真实验收测试的假通过路径**——网络隔离配**可达性正向对照**（不接受"Python 还活着"当对照）、宿主哨兵先真创建再核对、凭据先注入合成哨兵再验证未被继承、资源限额测试须证明探针已启动并给出限额证据、容器查询失败不得返回空列表而冒充"无残留"、真实测试中的实现异常必须失败而非被宽 `except` 吞成 skip。**边界如实记录**：普通后端模型/持久化/异常出口已审查收口；**SDK CLI 托管内部转录无法投影**，不宣称双后端同等隐私闭环；Docker 日志不作为代码/数据/输出的持久化出口（`--log-driver none` 为环境强制手段，非仅文档承诺）。**真实隔离实测受环境阻断**：本机 Docker 客户端存在但无可用 Linux 容器引擎（Docker Desktop 走 WSL2 后端，`wsl.exe` 被代理黑名单拦截）——mock 层已全绿，真实用例整体 skip（明确标注环境阻断，未跑不等于通过）。**代码层面的修复与判别力证据不等于安全闭环完成**：隔离执行器真实隔离证据、句柄登记、结果转 `AnalysisPlan`、双端入口、镜像构建摘要均未完成，`pythonExecution.enabled` 保持 `false`。下面 SDK"受限表结构通道"说明仍是当前实现状态。完整威胁边界、分阶段实施、双端与真实隔离验收见 [隔离 Python 设计](./SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md)，实施交接见 [DeepSeek 提示词](./DEEPSEEK_ISOLATED_PYTHON_PROMPT_2026-10-06.md)。

## 项目原则：新增功能必须双宿主对齐（2026-10-05）

Excel/WPS 同等用户任务能力是新功能的完成条件，不只是未来目标。新增功能从设计起同时列出两端入口、语义、授权/隐私、执行前确认、结果核验、异常与恢复、窄窗交互及测试证据。共享语义/协议/确定性计算优先复用，宿主差异通过适配层处理，不要求调用同一个 API，也不允许为了表面一致扩大授权或降低验证。双端自动化和实机证据分别报告；一端未实现时不能标整体完成。确有宿主限制或用户明确批准分批时，计划表保留具体未对齐项和验收条件，并区分未实现、宿主限制、实机待验。

审阅 AI 40% 自适应已接入 Excel 与 WPS，比例/手动优先共享，尺寸获取与宽度恢复由宿主适配；两端自动化已通过，真实宿主证据仍独立待验，不能称所有版本已验收。

## 2026-10-05：WPS 审阅窗格比例对齐

`packages/core/src/reviewPaneSizing.ts` 提供双端比例目标、实际宽度确认与手动优先状态机。相关限额从 `excelWindowLayout` 移到 `config/capabilities.json.reviewPaneSizing`，默认 40%、320 CSS px 建议下限、1500ms 检查、200ms 防抖、8px 容差。至少给表格留半窗口是共享产品上限，在 Excel 同时满足 Office 原生上限，不宣称 WPS API 有相同硬性限制。Excel 原有接口与普通偏好恢复不变。

`useWpsReviewPaneRatio` 仅在本次 `stagedReview` 存在时启动。`wpsReviewPaneSizing` 获取已存在的 EB 窗格 ID，不创建窗格；要求右停靠、可见、正宽度以及稳定 Application.Hwnd。Workbooks 集合中原表完整身份和未保存副本身份必须各唯一，当前文档须为其中之一；只读 Name/Path/FullName 等身份，不读 Range/Cells，也不调用 Activate。主窗口 `Application.Width` 是磅，按 96/72 换算逻辑像素，不能用两张表任一子窗口或 iframe/screen 当作主窗口。WPS [Application 文档](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Application/obj) 声明宽度单位；[TaskPane.Width 文档](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/TaskPane/member/Width) 未明确单位，因此用进入审阅时 native Width/真实 WebView viewport 对进行初始校准，不在 resize 尚未落定时反复校准而制造尺寸漂移。实际 viewport 到达目标才清除手动指引，不把 setter 返回当成功。DPR 改变即停止自动路径并提示手动/重新生成审阅；初始非标准 WebView 缩放、跨显示器等必须实机核实，不能凭替身声明通用支持。

普通宽度仅在内存中捕获，退出/放弃/确认后或卸载时尽力恢复；窗格替换、隐藏、宿主切换时不得误调别的窗格。Ribbon 与面板是不同 WebView，使用 `wpsReviewPaneLease` 的短期所有权标记防止 Ribbon 把审阅窄宽重新迁移到默认 1080。PluginStorage 仅保存 paneId/随机 owner/过期时间，6 秒 TTL，可见或隐藏期间续期，手动优先后只续期不调宽；退出按 owner 清除，支持没有 removeItem 的宿主；失效标记不能永久阻止普通初始化。没有路径、计划、矩阵、公式、请求模型/后台/执行账本，也不放宽现有保存与原表核对。

缺少 Width/Hwnd/Workbooks/右停靠现有窗格等证据，原表/副本关闭或身份模糊、API 抛错时停止自动并显示手动路径，绝不改用屏幕大小猜测。界面指引嵌入现有副本审阅状态卡，不另起全局操作卡。WPS 主窗口剩余空间继续由原生比较承载两表，未改变左右关系、同步滚动或审阅确认。openpyxl 文件夹模式无原生任务窗格，不引入桌面窗口操作。

## 2026-10-05：审阅 AI 窗格按右侧窗口 40% 自适应

替代下面 2026-10-04 的固定 360px 行为。外层副本左 40% / 原表与 AI 右 60% 不变；右侧内部 AI 目标是原表 XLMAIN 客户区宽度的 40%，不是屏幕或任务窗格 iframe 的 40%。新增只读 `/api/office/review-layout/measure`，复用 prepare 的严格三字段身份摘要/随机副本 nonce 请求模型与回环、Host、Excel Origin、JSON 守卫。必须唯一匹配已保存原表和本次未保存 nonce 副本，测量前后复核双方身份；仅返回 `clientWidthCssPx`，不返回句柄、路径、计划、值矩阵或标题，不落盘或领取执行权。Win32 GetClientRect 在显式 per-monitor 线程 DPI 上读取，再用 GetDpiForWindow 转成 96-DPI 逻辑像素，finally 恢复线程 DPI；最小化/隐藏/身份变化安全失败。TS/Python 同步增加测量回执，不改既有版本 3 排列与同步回执。

`useReviewPaneRatio` 从活跃副本 session 的原表身份、表名和副本 nonce 取测量范围。TaskPaneApi 1.1 可用才请求宽度；最小建议 320px，极窄时半窗口上限优先（例如 500px 原窗口只能请求 250px）。只读尺寸检查默认 1500ms，resize/focus 防抖 200ms，不重叠请求，页面隐藏暂停原生检查，离开审阅/卸载停止；不是滚动轮询，也不驱动任意窗口。只在真实 host 宽度变化时 setWidth，避免自身 resize 回环；窗格单独改变且不匹配程序请求时，判为用户手动调整，本份审阅停止自动调宽。窗口证据丢失/API 抛错停止自动路径；setWidth 返回不等于成功，实际 iframe 宽度才清除指引。审阅中不覆盖普通宽度偏好，退出后沿用普通恢复逻辑。未获得 session 前不先跳到固定 360px。

WPS 原生比较、openpyxl 文件夹模式、数据授权、执行与保存确认不变；不自动执行或回写任何单元格。真实 Excel 的 DPI/多显示器/WebView 缩放、拖动时序仍待实测；自动化只证明替身与代码行为。

## 2026-10-05：可见边框贴齐与原生同步滚动

外层副本左 40% / 原表与 AI 右 60% 保持。`windowGapDip` 默认改为 0；在物理 DPI 上读取 DWM `DWMWA_EXTENDED_FRAME_BOUNDS`，用它与 GetWindowRect 的差补偿不可见拖拽边框，定位后再次核对可见矩形。不修改系统边框样式，也不以原始 Win32 矩形接近当成视觉间隙已消除；异常或不可获取可见尺寸时回退手动。原有 6 DIP 的说明是此前版本行为。

本地桥接现在不仅摆放窗口：对同 PID、唯一原表与本次新建未保存 nonce 副本，使用原生对象模型 `Windows.CompareSideBySideWith`，显式设置 `SyncScrollingSideBySide=true` 并回读，初始行列视口对齐到原表。两条前端副本路径都消费此桥接。不运行 VBA/宏/脚本，不读取 Range/Cells，不创建 Application、不打开/另存文件，不安装滚动钩子或后台轮询。只允许两份已核实窗口；有第三个应用窗口时不替换潜在无关比较。COM 包装对象 Hwnd、原表路径/表名摘要、副本 nonce/未保存状态、进程与截止时限在变更前核对；宽度移动前和移动后再次核对现场。COM 指针及线程初始化按作用域释放，不把对象带出回调。

TS/Python 回执同步升级为 `layoutVersion:3`，在真实几何回执中新增 `synchronousScrolling: enabled | different_instance | unavailable`。旧版仅排列回执不可冒充联动，前端提示重启后台。仅 Compare 返回真且同步属性回读为真、行列视口相同，才显示已启用；跨实例或证据不足仅显示贴齐排列及手动指引。跨实例自动同步滚动未实现，不是待验项；不为它放宽既有隐私、身份或写入边界。窗口/开关回执不是单元格验证，也不自动勾选审阅。WPS 原生比较和文件夹模式不变。

真实宿主的双向滚动、DWM 时序、多显示器/DPI 与任务窗格留存仍需独立测试工作簿验收；自动化不证明已操作真实 Excel。后台修改需按原启动方式重启、刷新加载项并生成新副本。

## 2026-10-04：审阅时临时收窄 AI 窗格

外层副本左 40% / 原表与 AI 右 60% 不变。`useUIState` 接收原生审阅创建中或存在的状态；进入审阅时，若 `TaskPaneApi 1.1` 可用且当前窗格过宽，仅请求一次配置中的 `excelWindowLayout.reviewTaskPaneWidthCssPx`（默认 360 CSS px）。审阅期间不在焦点/可见性恢复时套用普通宽度，也不将临时 resize 覆盖普通聊天宽度偏好；退出审阅后恢复普通偏好。已经更窄的窗格不强制加宽，用户后续手动拖动不被反复重置。

`setWidth` 可能被宿主静默忽略；只有真实窗格宽度测量能清除“向右拖动 AI 左边界”的提示，不能把请求成功当成尺寸验收。无法使用接口时保留手动路径，不使用任意脚本或扩大窗口桥接权限。此变更仅 Excel 前端，WPS/文件夹模式、窗口几何协议、数据授权与保存确认不变。

## 2026-10-04：窗口排列版本与真实回执

兼容模式的 prepare 回执新增 `layoutVersion: 2`；前端拒绝没有此版本的旧凭据，不向旧服务发起排列。arrange 请求仍只投影一次性 token，回执包含 `status`、`layoutVersion`、`copySide: left`、`originalSide: right`。回执由 Win32 方法在实际矩形、左右关系及双方进程身份核对后产生，HTTP 路由不得只凭调用返回就构造成功。版本/顺序不符不显示「已自动排列」；旧服务有专门重启提示，兼容副本仍可手动审阅，保存确认门槛不变。无新增数据行、单元格、路径或句柄出站/落盘，WPS 与文件夹模式不接此版本协议。

本机观察到运行服务 12:19:33 启动，布局源文件 12:25:46 修改，服务命令不含 reload；运行中的 OpenAPI 仍只定义 `status: arranged`。因此刷新前端不能更新该后台，代码更新后须停止旧进程并按原命令重启、刷新加载项和重新生成副本。版本守卫避免旧服务泛化成功回执被新界面解释为图 2 布局。没有自动重启/终止用户进程，也不补认已有副本。

## 2026-10-04：窄窗格与规划进度交接

任务窗格单列显式使用 `minmax(0, 1fr)`，避免 CSS Grid 的自动最小列宽被内部头部和长文本撑开后遭外层 `overflow: hidden` 裁切。共享头部采用自然 flex 换行，品牌与图标在实际放得下时保持同一行，不再在 ≤700 CSS px 时强制分行；极窄时允许换行，不隐藏品牌或操作。澄清与判断轨迹的网格和长字段也约束在可用宽度内。

Excel `analyzeConfirmedIntent` 的本地步骤只表示「准备规划请求」，范围复核及请求整理完成后明确结束，再接服务端规划事件；它不表示尚在运行的整段「生成预览」。活动结束时，未收到结束确认的其他步骤收敛为警告（失败/取消则沿用对应状态），不保留假运行图标，也不因后续步骤完成而冒充这些步骤成功。活动归属、事件去重、授权与写入协议不变。

隔离 Edge 渲染检查使用真实 Excel App 或 WPS 面板的头部与输入区，以及生产 `TaskProgressTimeline` / `LocalMappingResolution` 组件，覆盖 320/350/400/430/500/620/700/900 CSS px、长模型名/范围/源值与整格/拆分两个表单状态。尺寸证据不等于真实 Excel/WPS WebView 验收，宿主截图仍待重测。

## 2026-10-03：网页前后对比退役

按用户要求移除 Excel 的 `review.html` 对话框与 WPS 的窗格全幅网页双网格，并删除专用组件、传输状态、样式、构建入口及 core 网格投影模块。侧栏保留变更摘要（定位、计数、真实回读状态）及已撤销的历史说明；执行器的前后证据采集、授权核对、验证与恢复能力不删除。

WPS 独立副本原生并排、确认写回与放弃副本保持不变。Excel 原生单表副本审阅已接线（范围和限制见下一节），不等于完成两端全部原生能力对齐。专注对话窗口与只读宿主能力探针不是网页前后对比，继续保留。旧双网格指导文件仅为历史记录，不代表当前产品入口。

## 2026-10-03：Excel 原生单表副本并排审阅

**2026-10-04 兼容路径补充**：`excelNativeReviewMode` 先检测基础 `ExcelApi 1.13` / `createWorkbook`，再区分支持 `ExcelApiDesktop 1.1` 的 Office 窗口模式和不支持它的兼容模式（内部仍名为 manual，如本机 Office 2024 Volume）。兼容模式使用 `manualReviewFile.ts` 将授权表快照与白名单计划直接投影为本地 OOXML：包含真实公式、常量、基础格式及绿色计划变化标记，要求 Excel 打开后重算，不伪造公式缓存结果；列宽按点到字符宽度近似转换。副本数据只通过 `createWorkbook` 交给本机 Excel，不调用不支持的 Office 窗口 API，不向原表写入，不新增单元格数据出站或持久存储。范围、动作、复杂对象与公式依赖限制和 Office 窗口模式相同，未支持的计划不能绕过限制。

兼容模式现通过既有 Python 本地服务尝试自动左右排列（见下节）；不能确认窗口时才明确提示手动「视图 → 并排查看」或排列不同实例。用户须勾选「已在 Excel 中核对副本」才能保存，hook 内也验证确认，不仅依赖按钮禁用。保存前及执行器首次写入前仍复核原表身份、授权、计划和完整来源快照；只执行原计划，不导入副本人工修改。此模式**不能自动回读副本单元格或关闭另一实例的副本**：窗口布局成功不是计算结果验证，结果验证只针对原表，不宣称与副本计算值一致；放弃/保存后均明确要求手动关闭临时副本且不保存。打开、重算、真实窗口排列和宿主复核时序待实机验收。以下 Office 窗口集合与副本实际回读描述仅适用于支持 Desktop 1.1 的模式。

### 兼容副本的本机窗口排列桥接（2026-10-04）

复用现有本地 FastAPI 进程，新增 `/api/office/review-layout/prepare`、`/arrange`，两端契约同步。只接受回环客户端、回环 Host、Excel 3000 端口的精确 Origin 和 JSON；前端拒绝向远程服务发送布局请求。prepare 载荷严格限定为 Windows 完整路径规范化后的 SHA-256、授权表名 SHA-256、本次随机 `EB_<20hex>` 表名；arrange 只接收一次性凭据。前端显式投影字段，服务端 extra=forbid，不能传计划、值矩阵、完整路径、窗口句柄或命令。原表身份从既有 Office URL 推导，仅支持已保存本机磁盘/UNC 地址，不修改既有授权及持久身份协议。

Windows `XLMAIN`/`EXCEL7` 可访问性接口只读取窗口关联的活动表名和工作簿 Path/FullName 等身份元数据，验证进程为 Excel.exe。不读取 Range/Cells，不创建 COM Application，不运行外部程序/VBA/脚本，不修改工作簿数据；身份扫描不激活窗口，版本 3 比较阶段仅可激活已核实原窗口、调整本次两窗视口及同步开关。prepare 唯一匹配原表并记录所有既有 Excel HWND/PID；createWorkbook 后 arrange 只接受新窗口、未保存副本和本次 nonce 三者同时成立的唯一候选。已有“工作簿1”、仅同名或同路径的多窗口都不放行；移动前再次核对双方身份及 HWND/PID，过期或模糊证据不移动。

Win32 仅在原表所在显示器工作区排列本次两窗：**左侧审阅副本 40%，右侧原表与 AI 窗格 60%**，先扣除配置间隔再分配宽度（版本 3 默认可见间隔 0 DIP，并补偿不可见拖拽边框），高度填满工作区。不修改 Excel 自带的系统边框样式或单元格格式；任务窗格的审阅临时宽度由前端单独处理；尺寸由 `originalWidthFraction` 与 `windowGapDip` 配置，间隔按窗口 DPI 换算，负坐标显示器同样按工作区计算。跨 Excel 进程也可尝试；DPI 处理后核对实际矩形，不一致或失败则尽力恢复原窗口位置，禁止报成功。限额放在 `config/capabilities.json` 的 `excelWindowLayout`：凭据 120 秒、最多 20 项、等待新窗口 12 秒、轮询 250ms、请求超时 20 秒、矩形误差 4px。凭据绑定 Origin、单次使用，只有内存元数据，不落执行账本。COM 元数据调用可能因 Excel 忙而阻塞，前端请求超时并回退，服务端检查等待截止后不得发起迟到的移动；不是所有原生 COM 调用都可强制取消。服务重启会丢弃凭据，必须重新生成副本。窗口布局成功不能自动勾选审阅、放宽写回；版本 3 的同步滚动独立回执只在原生比较和开关回读成功时为 enabled，跨实例自动同步未实现；兼容副本回读/自动关闭仍未实现。

仅 Windows 交互桌面且与 Excel 同用户、可访问的权限级别适用；非 Windows、云文件、不支持的身份、依赖缺失、后台服务会话、权限隔离均回退手动。文件夹 openpyxl 路径不启动窗口布局，WPS 保持既有原生比较通道不变。真实 Office 2024 COM 身份接口、跨实例排列、DPI/多显示器与聊天留存仍待实机验收。

`excelNativeComparison.ts` 用 `Excel.createWorkbook(base64)` 打开本机生成的单表 OOXML 数据种子，不用整本工作簿导出、不用后端生成文件。`Application.windows` 中必须唯一发现本次新增窗口，且其活动表匹配一次性随机名称，才允许通过 `Window.activeWorksheet` 写副本；不能识别时明确失败，不用当前工作簿代替副本。窗口并排使用 `compareCurrentSideBySideWith` / `resetPositionsSideBySide`，聊天仍在原加载项窗格。原生窗口发现、实际左右顺序、任务窗格留存、关闭提示仍需 Excel 真机验收；`createWorkbook` 若打开到另一 Excel 实例，本路径会安全报错。

仅支持一张已授权表上的互不重叠 `writeValues`、矩阵 `writeFormulas`、`clearRange(contents)`。`config/capabilities.json` 的 `excelNativeComparison` 限制已用/合并显示范围 20,000 格、OOXML 8 MiB。保留常用单元格格式、边框、行高列宽以及宿主确认的真实公式；文字常量即使以 `=` 开头也保持种子内的文本类型，不回放整块 formulas 矩阵。复杂对象（表格、合并、图表、形状、透视表、名称、条件格式、数据验证）明确拒绝；跨表、外部、结构化、INDIRECT/OFFSET 引用不能隔离试算，亦拒绝。不是全保真工作簿复制，打印配置、隐藏状态等未对齐。未支持的复杂计划保留既有显式执行流程，不自动从原生审阅失败降级为写原表。

右侧按白名单计划试写、只重算副本工作表、实际回读并标记变更。公式文本回读不等于语义验证，仍由原执行器验收条件核实计算结果；写回后另以 `compareExcelNativeCommittedResult` 比较原表目标的真实公式及计算值与副本，不同或证据不可用时分别收敛为失败/未核实，变更摘要同步降级，不用计划值冒充回读。`useExcelNativeComparison` 仅在内存中按消息 ID 持有一个副本会话；快照、值和 OOXML 不挂到聊天历史，不新增网络/存储通道。**这不是对旧 Office 执行账本数据边界的重构**，确认写回仍复用既有登记/执行/恢复接口。

确认写回先核对工作簿身份、授权、计划、原表全范围数据/公式/基础格式、副本全已用范围数据/公式/数字格式；再在 `executePlan.beforeFirstWrite` 中于持久登记回调之后、首次原表写入之前复核一次。仅执行已审阅的计划，副本手动编辑不导入。失败或部分写入禁止直接重复保存，沿用原执行证据和恢复点处理；关闭失败执行的副本不宣称“原表未写入”，也不撤销既有写入。原生审阅遇到过期来源不提供旧结果直写旁路。放弃只关闭已识别的副本；`Window.close()` 返回 false 时保持会话，不冒充已关闭。写回完成但副本关闭失败，不重复写入。写回不等于保存原文件到磁盘。

## 系统架构

### 整体架构

- 前端：React + TypeScript；Excel 宿主使用 Office.js，WPS 宿主已通过 W0/W1 接入读取、W2 Agent 核心验收，并已实现 W3 受控写入闭环
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
                       │ 可选后端：OpenAI-compatible API / Claude Agent SDK
┌──────────────────────▼────────────────────────────────┐
│                  外部模型服务                          │
└───────────────────────────────────────────────────────┘
```

### 可选 Claude Agent SDK 后端

`/api/models` 有两种显式 SDK 选项：三项环境配置 `CLAUDE_AGENT_SDK_MODEL`、
`ANTHROPIC_API_KEY`、仓库外绝对路径 `CLAUDE_AGENT_SDK_STATE_DIR` 同时存在时列出
`claude-agent-sdk`；已有托管 DeepSeek 连接使用官方 `https://api.deepseek.com`
或 `/v1` 端点、带 Key，且配置了会话目录时，另列出按连接 ID 绑定的
`claude-agent-sdk:<id>`。后者复用该连接的模型 ID 与 Key，只把 SDK 请求明确指向
DeepSeek 官方 `https://api.deepseek.com/anthropic` 兼容端点；不接受任意地址、
HTTP、其它路径或按相似域名推断。连接改址、删掉或清掉 Key 后衍生选项即失效。
原有 OpenAI-compatible 连接及默认选择保持不变；用户在 Excel/WPS 共用模型菜单中显式选 SDK，
`intent.py` / `planner.py` 才分发给 `claude_sdk_agent.py`，不会在请求失败时自动换供应商。
SDK 依赖本地 Python 包与原生 Claude Code 可执行文件，子进程在专用会话目录启动；
Windows SDK 服务由 `scripts/start_claude_sdk_server.ps1` 单进程启动、不启用热重载。
当前 Uvicorn 的 reload/多 worker 模式会安装不支持子进程的 Selector 事件循环；
适配器在 SDK 调用前检测这一组合并明确报不可用。启动脚本同时拒绝已占用的 8765 端口，
与普通 `npm run dev:server` 二选一。CLI 启动/退出异常只返回固定 SDK 提示，
不将原始 stderr、密钥或输入内容透传给前端。
该目录仍用于 CLI 配置；适配器添加 `--no-session-persistence`，本次 SDK 请求不保存会话转录。
旧版可能已有会话资料，不能放进仓库、项目日志或前端存储。
SDK 子进程显式收到选中连接的 base URL 和 Key，避免继承外部环境中的另一服务商地址；
DeepSeek 的 Claude Code 内部模型档位也固定到该连接的模型 ID。SDK 的美元预算参数
不等同于 DeepSeek 的实际账单上限，费用仍以服务商账单为准。

首批 SDK 接入是受限的表结构通道：意图阶段只发送当前需求和结构摘要，
规划阶段只发送工作簿名、已授权表名、范围、尺寸、表头，不发送
`dataRows`、`displayRows`、`dataResults`、`lastResult`、历史对话或 `sourceFile` 路径。
要求数据行的规划仍显式报不可用，不静默转到旧连接。用户主动添加的 PNG/JPEG/WebP
图片是独立输入：意图阶段仅发送数量，规划阶段通过 SDK 异步 user 消息的原生 image/base64
content block 发送；不创建临时图片文件，不开放 Read/Bash。只有明确声明视觉能力的连接才能接收。
SDK 内置文件、命令、网络和代理工具禁用；仅注册 Excel Bro 的终局白名单工具，
SDK 提示显式使用 MCP 全名提交终局结果。若服务商按共用提示返回结构化 JSON，
适配器先解析，再经同一个 `submit_plan`/`submit_answer`/`submit_clarification` handler 复核；
纯文字不能代替计划，JSON 计划也不能绕过宿主能力校验。
`submit_plan` 仍由 `AnalysisPlan`、宿主能力交集和原有预览确认链路校验。
SDK 没有工作簿句柄，也不能直接执行写入。原有文件夹模式与 Office/WPS 都经过
同一分发器，但文件夹请求若需要数据结果亦明确拒绝，不扩大用户选表授权。
预算上限在 `config/capabilities.json` 的 `agent.claudeSdkMaxBudgetUsd`。

#### 两条后端的检查点差异（CAP-02 第 6 条）

两条后端产出**同一联合类型** `AssistantResponse`（plan / answer / clarification），
并共用**同一份工具契约**：`run_sdk_plan` 从 `excel_agent.AGENT_TOOLS_WITH_DATA`
取定义注册进 MCP，终局也走同一批 `_execute_tool` 校验
（`test_sdk_backend_shares_the_blank_policy_contract_not_a_forked_schema` 断言逐字同一份）。

但**检查点能力不对等**，必须在 `planner.py` 显式说明而不是静默降级：

| 维度 | 普通模型后端 | SDK 后端 |
|---|---|---|
| 检查点现场 | `excel_agent` 循环里维持 `loopState` / `messages` / `toolCache` | **当前接入未实现检查点恢复** —— 本次接入把对话状态交给 CLI 子进程托管，服务端拿不到那份现场 |
| 数据行 | 支持 `planningReadPolicy=ifNeeded` | 仅表结构（`none`） |
| 跨后端续跑 | 可被 SDK 请求方引用（会被拒绝） | **显式 `ModelNotAvailableError`** |

> **口径**：表格里写的是"**当前接入**未实现检查点恢复"，说的是**这次接入方式**
> （CLI 子进程托管对话状态、服务端读不到现场）这一事实，**不是**对 Claude Agent SDK
> 本身能力的推断。SDK 是否提供别的会话/恢复机制属于**未核实**，文档不得写成
> "SDK 永远不支持检查点"；将来换接入方式或补做现场托管时，改这张表而不是改历史结论。

`planner.py` 在分发到 SDK **之前**检查 `resume_checkpoint`：

```python
if config.adapter == "claude_agent_sdk":
    if resume_checkpoint is not None:
        raise ModelNotAvailableError(
            "Claude Agent SDK 不能恢复另一规划后端的检查点，请重新发起预览。"
        )
```

拒绝的理由不是"做不到"，而是**静默重新规划会制造假成功**：用户点的是"从上次进度继续"，
拿到的却是一次从头开始的新规划 —— 可能重复调用模型与工具，界面还显示"继续成功"。
拒绝发生在启动 CLI 之前，一次模型调用都不发出
（`test_sdk_refuses_another_backends_checkpoint_instead_of_replanning` 断言 `fake.calls == []`）。

#### 规划阶段的结构化澄清

两种模型后端共用 `submit_clarification` 终结工具，返回 `kind: clarification`，
复用 `IntentClarificationResponse`，不把追问混入普通 `answer`，也不解析助手文字猜测问题。
`IntentClarification` 新增可选 `fields`（最多 4 个、唯一 id、label、placeholder）；
至少有一个填写字段或至少两个选项。旧选项卡兼容不变。Excel/WPS 共用无 hooks 的
`ClarificationFields`，各字段独立受控输入，全部非空后才可提交；用户答案按标签组合，
连同原需求经原有授权/表头复核再规划，不直接执行。取消后不生成预览，不读额外数据。
本地未覆盖源值仍走仅面板内存的 `LocalMappingResolution`，不借这条新出口上传源值。

非流式 turn 收到规划澄清时进入 `awaiting_clarification`，相同请求可缓存问题，
用户补充后可以继续。Excel 流式规划调用的请求指纹不可修改，因此规划澄清后的回答
启动新的 turn（保留完整原需求和授权范围），不复用旧持久规划任务。续跑记录新增
`originalPrompt`，旧记录缺少原需求时明确要求重发，不用问题摘要冒充原任务。
自动化验证与两端宿主实机验收分开报告。
Excel 的填写卡保留原规划读取策略，SDK 的首轮与自由文本续答均在本端显式选用
`planningReadPolicy: none` 并剥离数据行；不因“不是预设选项”恢复上传快照。

### WPS 宿主接入状态

apps/wps-addin 使用独立的 WPS Ribbon 与任务窗格入口；功能区采用独立的
“Excel Bro”选项卡，不向 WPS 内置“开始”选项卡插入分组。点击入口打开任务窗格后，
尝试切回内置“开始”选项卡；宿主不支持切换时仍保留已打开的任务窗格。W0 已在 WPS 表格个人版
12.1.0.28488 实测显示功能区、任务窗格并读取当前工作簿、活动表和选区。W1 增加
明确工作表授权、来源绑定、Value2 标准化、限额、分块进度、取消和块边界复核。

WPS 入口不会伪造 Office 全局对象，也不会把打开工作簿透明降级为磁盘文件读取。
W2 已把 WPS 读取器接入现有 query_table 确定性计算与模型连接服务。意图阶段只发送
授权表的字段结构；完整值留在本地计算，回答阶段只发送紧凑结果。WPS 响应出口过去拒绝
所有 `AnalysisPlan`，这项旧限制已在 W3 移除；查询仍走确定性本地计算，写入计划则必须通过
服务端规范化校验和 WPS 本地白名单校验，并在用户确认后执行。

W3 开发验收通道当前开放 `writeValues`、同形 A1 公式矩阵 `writeFormulas` 与 `clearRange(applyTo=contents)`。执行器只允许已授权工作表、
明确 A1 矩形区域、本地已保存且非只读的 `.xlsx`，并拒绝保护表、合并区域、公式注入、
重叠目标和超限写入。执行前完整快照暂存于 WPS `PluginStorage`，写入按行分块，
每块更新恢复点，完成后独立回读；公式动作比对 `Formula` 矩阵而非计算后的 `Value2`，
撤销前必须核对执行后指纹，检测到后续修改时拒绝覆盖。`PluginStorage` 关闭加载项后不持久化，
会话内恢复点不能承诺跨关闭恢复；跨关闭的写入耐久恢复见下「WPS 写入耐久恢复（CAP-03）」。
`#EB_RULES` 与 `#EB_RULES_BACKUP` 是内部系统表：WPS 枚举不会展示，授权入口和执行前复核也会拒绝。
整表删除、行列结构删除、格式、R1C1 铺满公式、拆表等动作仍默认拒绝，未知能力不降级执行。
目标版本的真实写入、异常中断和公式恢复仍须在独立测试工作簿完成验收，详见
[W3-A 验证记录](./WPS_W3_A_VALIDATION_2026-09-23.md)。

**WPS 写入耐久恢复（CAP-03，2026-10-02）**：会话内 `PluginStorage` 恢复点之外，首次原表写入
之前把执行身份、授权与**恢复描述**持久登记到服务端 SQLite 账本（`office_executions` 表，复用
`FolderOperationLedger` 的 revision CAS + ownerToken 轮换 + 动作级检查点），走独立的
`/api/wps/executions/*` 路由 —— **不把 WPS 冒充 Office 接入 `/api/office/executions/*`**。
身份用 `wps-path:` 前缀 + 已保存工作簿的小写磁盘完整路径，与 Office 的 `url:` 前缀是两个
命名空间，账本按完整字符串比较，跨前缀重放必然拒绝。

登记载荷**不是完整计划**：只发 `WpsRecoveryDescriptor`（动作类型 / 目标表与区域 / 目标形状 /
写前指纹 `previewFingerprint` / 写后逻辑指纹 `expectedFingerprint` / `applyTo` / `formulaR1C1`）。此前直接上传 `AnalysisPlan` 会把 `writeValues` 的目标
值矩阵带出站并落进 `plan_json` —— 恢复功能不得为"能精确续跑"放宽这条约束。三重闸：类型层
（`registerWpsExecution` 只收 descriptor）、服务端（`WpsRecovery*Descriptor` 用
`extra="forbid"`，矩阵字段直接 422）、客户端入站（`assertWpsExecutionState` 拒收含矩阵的状态）。

动作级（不是分块）写前提交写前指纹、写后提交写后指纹，完成/失败分别提交
`/complete`/`/suspend`。本地只存发现索引（operationId / 身份 / taskId），服务端也只存恢复元数据，
不存完整现场矩阵。写前证据、写后检查点、核对回执、完成结果同样使用严格字段白名单，
完成结果只收计数；异常请求使用固定码与固定说明，不上传宿主错误中的单元格值。
请求的 runtime 投影也丢弃额外顶层 `plan`，不能只依靠 TypeScript 类型约束。

恢复执行权保护由 `taskRecovery.wpsOwnerLeaseSeconds` 配置（当前 120 秒）。仍在 `running`
且近期有账本活动时，`ownerActive=true`、`canResume=false`；领取与核对在同一 SQLite
事务中复核执行权 + revision，不能抢占仍活跃的执行者。状态 GET 和重复登记均不泄露已有
ownerToken；Office 路由先核对 `url:` 命名空间，不能绕过 WPS 守卫。该时间保护不是续跑
许可证，过期后仍必须核对现场，未知/冲突均拒绝操作。允许用户点击只读「重新核对」，不轮询。

重载后读索引 → 查服务端 → 身份核对 → **五态核对**（`completed` 补记完成 /
`reconcile` 核对在途动作 / `replan` 重新读取并生成预览 / `blocked` 人工核对 /
`insufficient` 挂起）。按钮只由 `wpsReconcileActions` 给出，判定不给就不画；`blocked` 与
`insufficient` 一个写入或补记按钮都没有。恢复卡挂在消息流顶部。判定逐条要正面证明：
此前 `running`/`failed` 且无在途动作就返回"可以继续"（不看 `canResume`、身份、剩余动作、现场）
已被收紧。

现场证据来自 `inspectWpsRecoveryState`：先校验**全部**动作是否在当前已确认授权内，
再仅回读登记目标矩形；已完成动作比对写后指纹，剩余动作比对写前指纹，在途动作单独归类。
不依赖 PluginStorage 的整次执行状态，不补齐未采集的值；缺写后指纹的旧描述仅判证据不足。
点击按钮再次查询 + 回读，补记完成在领取后还要再次回读；全部动作完成只提交回执，宿主写入为零。
已有完成记录只显示历史事实，不宣称当前单元格仍未被修改。重规划前须明确勾选确认当前范围，
旧执行先领取处置权再挂起；规划失败保留发现索引，新预览仍要单独确认写回。

校验指纹不是加密，也不是公式计算正确性的证据。公式跨关闭仅核对逻辑公式文本，
**计算值跨关闭分离核对未实现**。旧版含矩阵的持久记录拒绝解释/下发/接管，错误不回显
敏感内容；本轮不自动删除历史账本，需要另行决定存量迁移/清理。

**精确续跑剩余动作未实现**（不是"待实机验收"）：`WpsPreparedPlan` 是内存对象（确认快照、
值矩阵、来源快照、授权索引、宿主版本），重载后重新 `prepareWpsPlan` 会读到"写后"现场，
无法重建首次执行的确认快照；值矩阵不落盘也不出站。PluginStorage 恢复点里虽有目标矩阵，
但跨加载项关闭即失效，且不含授权索引与来源指纹 —— 拿它续跑等于绕过写前的授权与源区核对。
因此剩余动作非空只能走 `replan`，面板明说"无法精确续跑剩余动作 …… 那是新的一轮"。


**恢复点的入口挂在"产生它的那次执行结果卡"内**，按 `executionId` 与执行绑定 ——
不按消息位置、不按计划标题、也不按"最近一次回复"推断。旧实现把一张撤销卡无条件挂在
消息流末尾，于是**上一轮**写入留下的恢复点看起来像本轮唯一能做的下一步。
「备份与恢复」抽屉只保留一个快捷入口。
每张卡如实区分四种状态：**可撤销 / 撤销冲突 / 已撤销 / 恢复点已不可用**；
只有当前会话里活跃的那条恢复点本人可点 —— 新执行取代旧恢复点、面板重开
（加载项会话已结束）或撤销完成后，旧卡不再保留可点按钮，卡片降级为状态说明。
确有恢复点的结果**不止"回读完全通过"**：写入成功但计算值与副本不符或无法核实、
以及可恢复的部分写入也会画卡（`outcome` 分别记为 `readback_passed` / `unverified` /
`partial`），文案与色调都不冒充"验证通过"；**执行器没给证据才不画卡**。
归属判定与文案在 `apps/wps-addin/src/wpsExecutionRecovery.ts`（纯函数），
渲染与 live 恢复点在 `WpsAgentPanel.tsx`。

**任务中断卡的"原表写入状态"是一条独立的生命周期，必须如实推进**（2026-10-02）。
恢复记录（`@excel-bro/core` 的 `taskResume`）在写路径上按四个时点推进：

| 时点 | `phase` | `writeState` | `resumability` |
| --- | --- | --- | --- |
| 预览已生成、等用户确认 | `preview_ready` | `not_written` | `replan` |
| **第一次宿主写入之前** | `executing` | `unknown` | `reconcile` |
| 写入与回读完成 | `executing` | `confirmed_result` | `unavailable`（终态） |
| 部分写入 / 回读不符 | `executing` | `needs_review` | `reconcile` |
| 用户放弃副本 | `preview_ready` | `not_written` | `unavailable`（终态） |

**顺序是关键**：登记必须早于第一次宿主写入。否则写入途中断线 / 重载时记录里仍是
`not_written`，卡片会对用户说"原表未写入"—— 而原表可能已被改过。宁可先说"未知"。
面板重开时 `unknown` 与 `needs_review` 都收敛为"先核对"，中断原因保留为
"写入结果无法确认"，**不降级成笼统的"连接中断"**；已完成/已放弃的任务是终态，
不再出现"继续"入口。规划身份（`turnId`）随记录一起保存，供后续状态查询定位该任务。

### 规划前能力交集（双宿主对齐 P0）

两端加载项在发规划请求时会自报宿主身份：`hostId`（`excel-office` / `wps-et`）、
`hostVersion` 与 `selfReportedActions`（本端能执行的动作清单）。服务端据此计算可规划动作：

```text
客户端自报 A  →  ∩  ← 服务端按 config/capabilities.json 的可信上界 B
                │
           A ∩ B ⊆ A（客户端只能收窄，不能扩大）
```

- 可信上界来自 `config/capabilities.json` 的 `hostCapabilities`，按
  `hosts.<hostId>.versions.<versionPrefix>.actions.<actionType>` 三层寻址；
  版本前缀按"相等或后接 `.`"匹配（`12.0` 命中 `12.0.1`，不命中 `12.01`），
  重叠前缀取**最长匹配**。
- 动作能力是三态：`supported`（已实机验收，可进计划）、`unverified`（可能支持但未验收，
  **不进计划**）、`unsupported`（宿主 API 不具备，不进计划）。后两者行为相同，区别只在说明文案。
- 请求**未声明** `hostId` 时保持原行为（兼容 `/api/plan` 与文件夹通道）；
  声明了但**无法识别**的宿主按上空集处理并拒绝整份计划 —— 不给"我不认识这个 id"留绕过口子。
- `clearRange` 按 `applyTo` 分键（`clearRange:contents` 等），父动作出现不等于全部变体放开。
- 这是第二道防线。执行预检仍是第三道，不因为有了规划前收窄而删除。

同一份判定在 TS 与 Python 两侧各有一份实现，由跨语言共享夹具
`config/action_capability_cases.json` 对齐；单边改动实现会立刻让另一侧变红。
Office 侧的自报清单 `apps/excel-addin/src/officeActions.ts` 由
`officeActions.test.ts` 对着 `contracts.ts` 的 `ExcelAction` 联合体与 `excel.ts`
的 `executeAction` 执行分支双向锁定，防止漏报（功能静默消失）或虚报（确认后才报错）。
具体阶段与验收矩阵见
[WPS 个人版适配指导](./WPS_PERSONAL_EDITION_ADAPTATION_GUIDE.md)。

WPS 与 Excel 任务窗格共用 `apps/excel-addin/src/styles.css` 作为产品视觉基线，并保持相同的
头部、模型菜单、新对话、工具、备份、格仔、更多菜单、消息流、数据范围入口和输入区结构。

**输入区的「上拉」是共享实现**（`apps/excel-addin/src/hooks/useComposerResize.ts`）：
拖动 `.composer-resize-handle` 调高度、内容变多时自动长高，长到上限之后**允许滚动查看**。
把手放在 `.composer-input-region` 的上内边距里、位于 `.composer-box` **之外** ——
挪进输入框内部会盖住文本区、抢走光标。WPS 侧曾用 CSS 原生的 `resize: vertical` 代替，
在 WPS 的 WebView 里既拖不动，又与继承来的 `overflow-y: hidden` 叠加，把超出高度的几行
直接切掉；那段覆盖已删除，`min-height` 与 hook 的 `minHeight` 保持同一个值。
跨端一致性由 `composerResizeParity.test.ts`（契约）与 `useComposerResize.test.tsx`
（行为）把关。
2026-09-30 用户明确修正 WPS 双视图需求：原表与**独立复制的工作簿**使用 WPS 自带
「并排比较」，对话留在任务窗格。WPS 计划预览后先复制已授权的单张工作表，
只在未保存副本上试写、回读和标色；「确认写回原表」重新核对两边状态后才调用
既有 `executeWpsPlan`，放弃则 `Close(false)`（按钮为「放弃副本」）。
**侧栏只负责状态、摘要、风险与决策**：原生表格负责比较，不重新引入窄 HTML 双网格，
也不默认把逐格长公式铺成主画面。成功结果不再展示旧的双网格。
副本数据只活在当前会话内存；多表计划不通过复制整本工作簿变相扩大授权。
此原生路径**已接线且自动化通过，待 WPS 真机验收左右顺序、并排窗口、底色及任务窗格留存**；
首次真机点击曾因把已保存工作簿的 JSAPI 包装对象用 `===` 判身份而在复制前误拒。
并排审阅现用已保存文件的名称/完整路径/目录、授权表及计划快照复核原表；
新建副本另绑定创建时的名称/路径与写后快照。该修复不放宽未保存工作簿的澄清身份判据。
详见[WPS 原生并排审阅](./WPS_NATIVE_SIDE_BY_SIDE_REVIEW_2026-09-30.md)。

**审阅侧栏的信息层级**（2026-09-30）。状态说明独占 `.wps-native-review-status`，
操作区 `.wps-native-review-actions` **只放按钮**（「放弃副本」「确认写回原表」）。
两者混进同一个单行 flex 容器会把按钮压到逐字折行 —— 那是真机截图里按钮竖排的直接
原因。按钮 ≥44px 高且 `white-space: nowrap`（不靠缩小字号或截断文案塞进一行）；
≥560 CSS px 水平排列，更窄时 `flex-basis: 100%` 整行纵排。完整公式与计划默认收进
`<details>`，用等宽字 + 横向滚动，不再用 `overflow-wrap: anywhere` 在窄列里任意折断。
**默认收起不等于删除信息**：授权范围、覆盖风险、未知类别策略与实际写入类型仍须可查。

**"预计"与"副本回读"必须严格区分**。前者只能来自可信的本地确定性预期，必须标明
*尚未写入*；后者必须来自同一次副本试写的真实回读，不得拿计划值冒充。公式试算与
固定值回退用**不同的真实状态文案**，判据是计划里真实的 `action.type`（不是文案猜测）；
试算或回读失败时不显示"已回读通过"，也不给可写回的主按钮。WPS 没有 Excel 式逐项
`checks`，不得借 UI 伪造验证证据。
Excel 独立 Web 审阅窗口及此前的 WPS Web 宽窗探索见
[宽视图指导](./WIDE_CHANGE_REVIEW_GUIDE_2026-09-29.md)，不应再被当成 WPS 当前目标形态。
宿主差异只保留在读取、写入、恢复和功能可用性适配器中；WPS 尚不支持的工作目录
或结构动作必须明确禁用或拒绝，不能用外观相同的空入口伪装已实现。WPS 的“跟随当前”范围在
每次发送或运行工具时重新解析活动工作表；手动多表范围只会收缩到仍存在的已选表，不会自动
扩大。固化查询直接复用本地确定性执行器，并再次检查保存的数据源是否包含在本轮授权中；
对话历史保存在 WPS 加载项自己的本地存储中。模型连接仍由本机 FastAPI 配置统一提供。
WPS 的本地读取、计算及写入回调会整理为 Excel 共用的 `TaskProgressTimeline` 数据形态，
记录步骤、阶段、动作类型、已授权目标、行数和终态，不把原始单元格值写进过程日志。
过程消息保存在对话历史中；刷新时尚无终态的步骤显示为中断，不冒充继续运行或成功。
消息流仅在用户靠近底部时自动跟随更新；上翻阅读时显示“查看最新进度”，不强制拉回底部。
这只对齐了当前 WPS 本地可见过程；服务端规划流的逐事件补发仍须单独接线与验收。

写入能力按共享核心与宿主适配器分层：`apps/excel-addin/src/planCore.ts` 提供两端共用的
A1 矩形边界、二维写入矩阵形状和清空内容值矩阵。`AnalysisPlan` 与服务端
`/api/plans/validate` 仍是同一计划协议；Excel 的 Office.js 执行器与 WPS 的 JSAPI
执行器只负责宿主对象、授权复核、执行前恢复点、实际写入和回读。新增共同写入动作应
先在共享计划语义中实现和测试，再分别接入两个宿主适配器；不能复制一份业务规则。
宿主版本验证与动作可用性独立于共享语义，未经验收的版本不因核心复用而自动开放写入。
当前开发调试加载项对实机报告的 `Application.Version="12.0"` 提供临时验收通道，
让独立测试工作簿能够验证写入、回读和撤销；这不是把 `12.0` 认定为产品版本或
正式已验收版本。先前把安装目录的 `12.1` 误放进正式名单，现已清空
`verifiedVersionPrefixes`；生产构建在真实版本信号和实机验收完成前拒绝 WPS 写入。
预检、计划准备
和确认执行共用同一版本规则，工作簿、授权、动作白名单和恢复点检查不因临时通道放宽。
WPS 写入计划中的空字符串会原样保留；独立回读若得到 `null`，不能再折叠后报告
`verified`。新恢复点使用带 `v2:` 前缀的类型化指纹，区分空字符串与空单元格；
旧 v1 记录仍用旧算法核对，无空值歧义时可以撤销，有歧义时保留恢复资料并拒绝
自动覆盖。预览绑定最终计划、授权表、工作簿身份和目标写前指纹；本地列映射
还绑定源列在计算与预览时的值，执行前复核源和目标，变化时必须重新预览。
WPS 预览显示每个动作前 5 行、前 5 列的待写值；完整计划可按与 Excel 一致的
JSON 方式编辑。应用修改会重新校验计划和实时授权、生成新预览修订，旧确认不能
直接用于修改后的矩阵；本地映射计划仍不把可逆源值送去服务端校验。
本地列映射在同表纯文本且公式可安全限长表示时先生成 `writeFormulas`：共享
确定性映射仍是预期值权威，未知新组合返回 `#N/A`；WPS 只在未保存的右侧副本
试写并核对公式文本及计算值。试算不符会关闭该副本，改用 `writeValues` 重新生成
待审阅副本，不会因公式试错写原表。Excel 当前工作簿与文件夹通道尚无等价的
写入前隔离试算，继续用固定值，不冒充公式能力对齐。
这仍不能代替 WPS 实机写入、保存重读、撤销和异常中断验收。

意图协议增加 `conversation` 结果：问候、能力咨询等非工作簿任务由模型直接回复，
不进入查询工具、规划或澄清卡。只有关键歧义会返回 `clarification`；明确查询继续由
本地 `query_table` 计算，写入仍必须先生成白名单计划并由用户预览确认。WPS 在意图判断
前仅读取已授权表的候选表头和尺寸，完整数据只在模型请求确定性查询后读取；
结构摘要不会包含原始数据行。澄清卡提供推荐选项和自由文本补充。

列映射的匹配顺序是**先整格、后拆分**：来源单元格整体命中映射表就用它；没命中且
规则声明了分隔符（`delimiters`）时，只按已声明的分隔符拆分（多个候选按在格中
**出现位置最靠前**者切分，不按声明序——声明序表达优先级，不是切分方式），拆出的
项逐项精确匹配后按 `valueRank` 取最高级（未声明 rank 的项视为最低）。规则的
`delimiters` 与 `valueRank` 必须**成对**给出：只声明分隔符表达不了"取哪一个"，
服务端会判定规则不完整并返回澄清，而不是按 `mapping` 顺序硬排一个——那个顺序是
模型的声明顺序，不代表用户说的优先级。该校验排在 `planningReadPolicy` 分叉**之前**，
两条通道（服务端算 / 客户端算）用的是同一份 `@excel-bro/core` 实现，规则不全时
两侧都算不出来。

**`rank` 是优先级，不是等级名里的数字**（2026-09-30 真机回归）。core 的约定是
**rank 越大越优先**，而业务等级常见的命名（名字里最小/最"0"的那个最严重）与之**反向**。
模型若照抄字面数字（把最严重的填成 rank 0），**不会报任何错**：服务端 `_rank_entries`
原样保留、core 按 rank 取最大、公式再忠实编码本地预期 —— 整条链路**一致地**给出那个
**较低**的等级，用户只会看到结果莫名变化，排查时又处处"自洽"。所以：
- 极性本身由构造保证，且已被夹具 `column_mapping_cases.json` 的"反过来声明则取另一个"
  用例跨语言钉住（`packages/core` 与 `server/app/column_mapping.py` 各跑一遍）；
- `rank` 与"拆分后取哪一个"的**方向说明**必须留在 `map_column_values` 的 schema 里
  （参数级 `rank` 说明 + 工具级说明 + 系统提示三处），并用**抽象**等级名举例 ——
  生产提示词里不得出现任何具体业务等级字面量，那条负向断言在
  `server/tests/test_map_column_tool.py`。

算不出正确答案时的归因刻意分两类，不得互相冒充：**归因确定**——已按声明的分隔符
拆开、拆出的项里有映射表没覆盖的类别，抛 `UnmappedCategoryError`，补救动作只有
「补充类别映射」；**归因不确定**——整格没命中、又像是用了某种没被声明为可分分隔符
的写法，抛 `NeedsClarificationError`，补救动作有两个可能（补分隔符 / 补映射），
必须让用户选。界面统一用一张澄清卡（复用既有需求澄清卡的状态机），但选项**按归因
收窄**：归因确定时不给「指定其他分隔符」，否则会让用户去改一个本来没问题的设置。
可选项只由 `reason` 推导（`columnMappingClarificationCard`），不直接信服务端下发的
`resolvedBy`。澄清内容**不含来源格原文**，只有归因、缺失类别与"无法处理的格数"——
原文属于本端数据，要展示的阶段界面有自己的源值；错误文案里也不嵌入整格原文。

**两条通道都要在读取源值之前分流澄清**。Office 侧在 `App.tsx` 的 `local_plan_rule`
分支查 `rule.clarification`；WPS 侧在 `runWpsTurn` 的同名分支做同样判断，且两处都必须
排在读取映射源列**之前**——服务端已判定"算不出正确答案"时再读一遍源值，等于白取一次
用户数据只为了说一句"缺映射"。

另有一种**只有本端读到源值后才能发现**的缺口：规则结构完整，但实际来源值未命中映射、
整格含未声明的分隔形式，或**源区域真的存在空白格而空白怎么处理还没定**。Excel 与 WPS
均把 core 的 `UnmappedCategoryError` / `NeedsClarificationError` / `BlankPolicyNeededError`
接到共用的 `LocalMappingResolution`，不再把它们作为终止系统错误。三类缺口按 `reason`
分成三组互斥的选项，**不许互相冒充**：`unmapped` 补类别、`separator_or_mapping` 让用户
在分隔符与映射之间选、`blank` 只给空白的四种处理。空白那一支必须发生在真实读取之后 ——
没读到空白就先问"空白怎么办"是刻意消灭的行为（历史上 `blankLabel` 是必填字段，模型为了
填满它只能在读源值之前提问，用户被迫回答一个可能不存在的问题）。
卡片在当前面板内存展示真实未覆盖值，允许用户选择逐项精确映射，或明确新增分隔符、
补齐拆分项映射、排列**全部目标标签**的优先顺序并勾选确认。候选分隔符只是建议；
不能自动拆分、用 `blankLabel` 冒充未知值，或暗用映射声明顺序决定“最高”。
卡片的持久化消息只含数量、范围和状态；原始来源值与补充规则不放进澄清消息，
目标矩阵不随澄清或规划请求出站。**生成预览后的 Excel `plan` 仍沿用现有计划历史
持久化策略**，不能把“澄清卡不存矩阵”误写成“整个聊天历史不存矩阵”。Excel 重开时
因内存草稿已消失而使待答卡失效。WPS 作答前复核授权身份、
真实表头和宿主对象；Excel 复核工作簿身份及勾选范围，没有可验证身份时不展示
无法安全续答的假卡片，而是要求先保存工作簿并重新发送需求；通过后均**重新读取**授权源列，
在本地生成新的白名单 `AnalysisPlan` 预览，仍须用户确认才写入。重读若又出现缺口，
在原卡继续询问，不沿用旧快照；不调用 `/api/plans/validate`（可逆目标矩阵不能出站）。
服务端预先返回 `rule.clarification` 的语义缺参路径仍使用通用需求澄清卡，
不能与这里“本地读取后发现的来源值缺口”混称同一个闭环。

**空白怎么处理是四条互不相同的路**，不能塞进一个字符串：

| 策略 | 落地方式 | 与相邻策略的差别 |
|---|---|---|
| `keep` | 该位置**从写入范围里排除**（`splitTargetActions` 拆成连续段） | 不是"读回原值再写回"——后者是一次真实写入，会改公式与格式 |
| `clear` | `clearRange` + `applyTo: "contents"` | 得到真空单元格，不是常量空字符串 |
| `value`（含 `""`） | 写一个用户指定的值 | 写空串时 `ISBLANK()` 为假，与 `clear` 不是一回事 |
| `undecided` | 尚未决定 | 只有真的读到空白格才升级成一次澄清；静默折成默认标签是被禁止的 |

协议上用结构化 `BlankPolicy` 表达，`contracts.ts` 与 `models.py` 同步；旧 `blankLabel`
非空串等价于 `{kind:"value"}`，语义无损。跳过位置过多时按
`capabilities.columnMapping.maxSplitActions` 拒绝并说明原因，不降级成写空值。WPS 的公式
通道只在目标是一整块连续写入时才启用——公式会覆盖区域内每一格，表达不了"这一格别动"。

**授权动作与澄清答案是两种状态，不得互相冒充**：点"重新确认数据范围"只证明"你现在授权
这几张表"，不产生答案。实现上把重授权落到 `continued` 终态（无真实答案时原样保留卡片），
说明文案走 `REAUTHORIZATION_NOTICE` 进消息流；只有用户真的填过答案才推进到 `resolved`，
且答案带 `questionId`，改答案**替换**旧确认而不是追加——追加会让模型同时看到两条互相
冲突的确认。

这条判据必须落在**真实取值范围**上，不能靠调用方贴的标签自证。WPS 侧的结构扫描
（`readAuthorizedWpsSheetSchemas`）与数据读取共用 `readChunkRange`，所以该函数
**按目的构造不同的取值面**：结构扫描只取**第 1 行 × 前 `structureColumns` 列**
（表头行），报告的行列数仍是整块已用区域 —— 授权边界与上限判定不受读取范围影响；
数据读取才取完整授权矩形。"澄清档没读源值"因此是**可被証伪的事实**：
`wpsWorkbookReader.test.ts` 按替身宿主真实的 `Value2` 访问坐标记账，断言结构扫描
取过值的格子里**没有一格属于数据行**，而"算得档"必须真的读到源列（对照锚点）。

**宿主调用不得多传自有实参**。`purpose` 只在 `readChunkRange` **内部**决定取值面的
构造与断言，**绝不传给宿主** —— 真宿主对多余实参不宽容，曾因假设"真宿主会忽略
多余实参"而把它塞进 `sheet.Range(first, last, purpose)`，2026-09-30 真机实测
WPS ET 直接抛 `Too many parameters.`，结构扫描第一步就失败。替身是 JS 函数，
多传的实参会被**静默忽略**，所以本地全绿、真机爆炸。防线已内置：替身的
`Range` 按 rest 参数检查实参个数，多传即抛同一错误（反向变异实测 12 条红），
今后任何"往宿主调用里塞自有标签"的改动都会在本地测试当场暴露。

同一个口径也适用于出站面，但两者是**独立的证据面**：读进内存 ≠ 发出去。
本地映射计划在 `fetch` 请求体层面另有一条负向断言——源列值、映射源列地址、
规则本体与算好的目标矩阵都不出现在任何请求体里（`wpsLocalPlanOutbound.test.tsx`）。

**待答澄清绑定授权身份，不绑定索引**。索引只是"第几张表"：重排、插入、删除或
替换之后同一个索引会指向**另一张表**，用户对着旧表回答的澄清会被拿去算新表。
宿主没有跨会话稳定的工作表 ID，因此 WPS 侧保守地把**工作簿名 + 每个已授权表的
索引与名称（含顺序）**算成身份指纹，再加上结构扫描实际读到的**表头指纹**；
任一项变化就作废待答澄清、把原需求退回输入框，要求重新授权、读取与预览
（`wpsClarificationScope.ts` + `WpsAgentPanel` 的 `clarificationScopeStillValid`）。
代价是"表改名"会让澄清卡失效——这是必须付的代价，不得为了"改名不断对话"
牺牲授权边界。指纹只含表头与表名，不含单元格数据值。
这两条是可序列化指纹的判据；**同会话内**还有第三条不可序列化的对象身份判据
（见下），三条并列，任一条判"不同"即失效。

**两条判据必须读不同的数据源，否则后一条只是前一条的副本。** 这是踩过的坑：
"表头指纹"曾一度喂的是 `listWpsWorksheets().sheets.map((s) => [s.index, s.name])`
——**工作表元数据**，与身份指纹同源。后果是只禁用身份判据时，表头判据顶上来把
用例撑绿，造成"身份判据没必要"的假象（实测：仅禁用身份判据，5 条定向用例只红了
1 条）。现在表头指纹只取自 `runWpsTurn` 结构扫描真实读到的第 1 行单元格值
（`WpsTurnResult.headerRows`），复核时经 `readAuthorizedWpsSheetSchemas` 重读。
两侧的独立必要性由 `wpsClarificationIdentityNecessity.test.tsx` 用**单判据变异**
分别钉住：只禁身份 → 4 条红；只禁表头 → 1 条红。不得同时禁用两条来充当证据。

**第三条判据：同会话内存对象身份。** 指纹是**可序列化**的，只负责"跨会话还能证明的
那部分"；它**结构上**就看不见两类碰撞：① 两个**不同目录的同名工作簿**（指纹只带
`workbook.Name`——路径含用户名等信息，不进指纹、**不进 localStorage、不进请求体、
不进日志与错误文案**）；② **同索引、同名、同表头**删除后重建的工作表（复制粘贴一张
表即可，无需任何"看起来危险"的操作）。这两类只能靠**不可序列化的内存对象引用**证明：
它只在当前会话活着，恰好覆盖"指纹看不见、但同会话内还能对得上"的区间。建卡时
`captureWpsObjectIdentity` 记录工作簿对象引用与每张已授权表的对象引用，作答前
`verifyWpsObjectIdentity` 逐项复核（`wpsObjectIdentity.ts`）。**会话内内存 ref，不落盘。**

**放行需要正面证明，不是"没发现不同"。** 三层判据：① 工作簿对象引用；② 路径/完整名
（**只在内存比对**，未保存工作簿两边都为空时**不因此判相同**）；③ 逐张表的对象引用。
**只要引用对不上就判 different，不存在 fall-through 到 same**——名称、表头、尺寸、
甚至稳定 ID 都能在"删除后重建""整表复制"里保持一致，拿它们反向补证等于把
fail-closed 拧回 fail-open（这是实现中实际踩到并纠正过的错：初版判据在引用不等时用
`CodeName` 补证，结果"重建后 CodeName 相同"被判成 same）。稳定 ID 与尺寸只作**加固**：
在引用**已经相等**时再核一遍，防宿主对象池复用；结构上无法证明身份时**一律失效**。

**稳定 ID 只作加固，不作放行依据。** `Worksheet.CodeName` 文档写明"只读、改名不影响"，
但**未说明唯一性**，也**未说明保存重开后是否保留**，且本项目**未在桌面 ET 实机验证**；
`Worksheet.SId` 属 **WebOffice JSAPI**（promise-based），桌面 ET 宿主**未必提供**。
两者都按可选字段探测式读取（`readSheetStableId`，逐候选 `try/catch`），读不到一律回退
`undefined`，**不得凭空造 ID**；捕获身份失败（拿不到工作簿、序号越界、表不存在）时
**抛错**并把 ref 置空，复核随即 fail-closed，不静默返回空身份。

**体验代价（显式接受，不得悄悄优化掉）**：若真宿主每次 `Item()` 返回的是**新包装对象**
而非同一引用，合法续答也会被拒、要求重新授权与重新预览。这是 fail-closed 的必然代价，
属**实机待验项**；不得为了让用例更好看而放宽判据。反向风险——"未保存工作簿的合法续答
被误杀"——已由定向用例保护：同会话内引用可证明时应正常作答。
另：不得为了验身份重读数据行，也不得扩大用户勾选的工作表范围（仍用 `turnSheetIndexes`）。

**作答前只能有三种结果，且不得互相冒充**（2026-09-30 真机回归后修正；
`wpsClarificationVerdict.ts` 是唯一判定入口）：

| 结果 | 条件 | 处理 |
| --- | --- | --- |
| `same` | 对象身份被**正面证明**相同 | 继续回答，复用该轮授权 |
| `changed` | 有**正面变化证据**（路径/表名变了、索引越界、工作簿关掉了、或真的读到了不同的表头） | 旧卡 `invalidated`，提示"数据范围已经变化" |
| `unknown` | **证明不了**（引用换了但无其它证据、读到失败、面板刚重开） | 旧卡 `unconfirmed`：**不谎称数据已变化**，保留原需求与用户刚填的回答，请他在数据范围里重新确认原工作簿与授权工作表；确认后作为**新一轮**重新读取结构、重新规划与预览 |

**"引用不相等"属于 `unknown`，不属于 `changed`。** 宿主可能对**同一张表**返回**新的包装
对象**（用户只改了一格又撤回就会触发）。把它判成"变了"会对着范围根本没变的用户说
"数据范围已经变化" —— 那是假话，且他无从纠正。真正的变化靠**正面证据**判（路径、表名、
索引、表头），不靠"引用对不上"反推。

**读到失败（`null`）不算变化证据。** 指纹比对时任一侧读不到就记"没有可比信息"；
把读失败当成变化是这套判据此前的另一个假话来源。

**面板重开**由独立的读取路径判据保障：`loadWpsChatHistory` 把 `pending` 收敛为
`unconfirmed`（不是 `invalidated`）—— 重开本身没有给出任何变化证据。

三条判据的独立必要性由**单判据反向变异**钉住（`wpsObjectIdentity.test.ts` 19 条 +
`wpsClarificationCollision.test.tsx` 8 条 + `wpsClarificationVerdict.test.ts` 11 条）：
只禁对象身份 → 2 条红，且**正是**"不同目录同名工作簿"与"同索引同名同表头重建"两条；
只禁表头 → 1 条红；只禁重开失效 → 1 条红；把 `unknown` 退回旧的 `changed` → **8 条红**；
复原后全绿。不得同时禁用多条来充当证据。

**`unknown` 不是放行。** 它同样挡住作答，只是不撒谎：用户的原需求与回答被原样放回输入框，
重新确认范围后走**全新**一轮（重新读取、规划、预览），不复用旧卡的授权，
也仍然要经过写入前的预览与确认。

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
├── structural_intent.py    # 可确定识别的高风险工作表结构命令
├── planner.py              # 确定性地址命令与最终规划
├── excel_agent.py          # 模型工具循环
├── agent_loop_protection.py # Agent 预算、循环检测和安全降级
├── turn_state.py           # turnId 进程内热状态、TTL 和缓存
├── task_ledger.py          # SQLite 持久任务、规划现场、事件、预算和幂等状态
├── folder_workbooks.py     # 文件夹执行
├── folder_data.py          # 文件夹确定性查询
├── capabilities.py         # 共享能力配置
├── safety.py               # 安全与协议校验
├── rule_generator.py       # /function 公式生成主流程
├── formula_tools/          # /function 工具调用架构（schema/编译器/注册表/Prompt）
├── formula_tools_integration.py  # 工具调用路径与主流程的衔接
└── llm/                    # 模型调用、错误和重试
```

### Agent 循环保护

2026-10-03 普通规划通道改为按需加载动作 schema：首轮保留列映射、回答、澄清和
`select_plan_actions`（需要数据的兼容入口另保留只读工具），不携带全部动作参数。
模型选择当前宿主允许的动作类型后，下一轮只提供这些类型的完整 `submit_plan` 参数；
完整 AnalysisPlan、宿主能力、授权与安全校验不裁剪。选择结果随规划检查点保存与恢复，
这不是具体行业关键词路由，也不执行任何写入。SDK 当前仍使用原来的终局工具集合。
Excel 当前工作簿的新规划默认 `planningReadPolicy=none`，与 WPS 本地规则通道一致；
无本轮查询结果的结构规划不夹带上轮结果值（客户端与模型载荷分别剥离 lastResult）。
文件夹及显式 ifNeeded 请求仍保留请求快照读取能力，不冒充宿主实时读取。
预算停止返回 `answer.outcome=failed` 与结构化 `stopReason`；两端主路径/Excel续跑消费失败，
服务端热状态与持久规划账本也记 failed，不再因返回 answer 而假报完成。
老回答缺 outcome 仍按正常回答解释。上游 finish_reason=length/content_filter 时拒绝
整份输出（包括看似合法的工具调用）；length 单独停止为 output_length，不缩额重试，
content_filter 沿用有限修复。只记录结束类型及用量计数，不记录原文。
两种后端预算口径仍不同，不把 SDK 成功解释为 token 更少或底层模型更强。

excel_agent.py 的每轮模型调用都由 LoopState 追踪工具路径、缓存命中、连续结构
错误、累计 token、墙上时间和快照截断。循环开始前按固定顺序检查四层预算闸：
最大轮次、连续错误、可选累计 token、墙上时间。累计 token 默认不触发停止，
仅当 agent.maxTotalTokens 配置为正整数时启用；计数与检查点恢复不受影响。
检测层还会识别重复缓存查询以及
2 至 4 步的重复工具路径，并把一次性警告附加到工具结果；接近预算尾部时，
根据发散、收敛失败或循环状态向最后一条用户消息追加收尾提示。
一批只读工具成功返回后，后续轮次只暴露终局工具与按需动作选择工具，要求
Agent 基于已有上下文直接收敛，避免重复读取工作簿耗尽 token 预算。
结构化澄清选项一经用户点选即视为最终语义，不再把确认结果送回意图模型重复
提问。前端通过 `PlanRequest.planningReadPolicy="none"` 标记该路径；规划 Agent
只获得映射规则、澄清、回答与按需计划工具，并依据已授权工作表的准确表头生成
运行时计划。`splitGroupAggregate` 等动作仍在执行阶段读取完整数据，因此不受
200 行模型快照限制。
澄清选项通过 resolutionMode 区分 append 与 replace。append 只补充缺失参数，
并明确最终确认高于原始背景；replace 的 resolution 必须是完整最终需求，前端会
彻底丢弃与之冲突的原始指令。计划提交格式失败后，历史工具参数会压缩为最小占位，
下一轮只保留 submit_plan，并使用 agent.repairMaxOutputTokens 限制输出；生产环境
连续两次结构错误即停止，避免完整规划反复消耗 token。自由文本提交失败时也只在
历史中保留最小占位。splitGroupAggregate 的 splitBy 会自动成为输出列；模型
若又把同名字段放入 groupBy，服务端会在校验前确定性移除并在计划警告中披露，
而不是消耗一轮模型调用修复不改变语义的协议冗余。模型把采用 default_factory
的顶层集合字段提交为 null 时，同样按省略处理并使用模型默认值。

当前输入若自身包含完整动作和目标，必须作为新任务覆盖最近对话和上一轮结构化
意图；只有明显省略式追问才承接旧任务。“删除除某表外的所有工作表”由前后端
共同使用的确定性语法识别，规划器直接生成 `deleteWorksheet` 白名单动作，不再
允许模型复用上一轮拆分或汇总计划。
拆分聚合执行后，两条执行通道都会根据本次实际生成的工作表名称建立动态
`rangeEquals` 验收条件，并重新读取写入范围核对表头、分组字段、指标和占比；
超过 500 行的结果会分块验收，不以“工作表存在”代替具体效果验证。
数值比较采用与 Office.js 相同的浮点容差，只忽略 Excel 保存重读产生的机器精度
末位差异；可见数值变化、计数错误和占比口径错误仍会导致验收失败。
文件夹执行在原子替换成功后、独立验收开始前登记服务端恢复记录，因此验收
`failed` 或 `executed_unverified` 不会丢失撤销入口。恢复记录只保存服务端生成的
目标路径、备份路径与写入前后内容指纹；`/api/folders/undo` 按 LIFO 级联恢复，
已有文件从自动备份原子替换，本次新建文件恢复为不存在。恢复前先验证所有当前
文件仍等于对应写入后指纹、备份仍等于写入前指纹；任何一项变化都在磁盘写入前
整体拒绝，避免撤销覆盖后续人工编辑。文件夹备份集中位于所选工作目录的
`.excel-bro-backups/executions/<undo_id>/`，同一次执行的多个文件共享一个
manifest 并保留原相对路径。manifest 中每个目标使用 `recoveryAction` 区分从副本
恢复的 `restore` 与撤销时移除本次新建文件的 `remove`；旧 manifest 缺少该字段时
根据 `backupRelativePath` 兼容推导。manifest v2 记录 preparing、prepared、applying、
ready、recovering、restored、rolled_back 与 needs_attention 生命周期以及文件级前后
指纹；旧 v1 记录保守降级为只读历史，不冒充可直接撤销。快速撤销栈淘汰和撤销成功只
改变资格或历史状态，不直接删除磁盘副本。

清理先保护当前撤销链、执行/恢复/导出中的记录、主动保留和异常资料，并为每个原文件
保留最后一份可核对副本；7 天是自动副本最短保留期，1 GB 是软上限。自动清理和手动
删除复用同一事务：验证 manifest、登记根目录、整条路径和 Windows reparse point 后，
先将完整执行目录原子移动到 `.trash/<operation_id>/`，再递归清理；占用导致的残留纳入
待清理空间并可在下一次“清理旧备份”安全重试。删除确认同时绑定 manifest revision 与
级联失效 ID，影响变化时拒绝旧确认。撤销自身的回滚救援副本集中保存在
`operations/<operation_id>/rollback/`，仅在恢复无法完整回滚时保留。

备份列表为已登记目录签发短期 `backupManagementToken`；打开、取出、保留、删除、清理
和旧版归拢均不接收任意目标路径，也不扩大工作表数据授权。取出副本写入固定 exports，
使用稳定 operationId 和原子清单实现断线幂等重试。打开接口只解析列表登记的
`.excel-bro-backups`，目录不存在时不签发令牌，也不创建空目录。文件夹保存临时结果前
及备份复制前检查磁盘余量，无法保留配置的安全余量时在替换原文件前停止。

Office.js 执行快照按真实 executionId 绑定，成功、动作失败和验证异常都返回已经捕获的
恢复资料。快照记录写后指纹、工作簿身份、新建表状态与覆盖完整度；撤销前逐项复核，
覆盖不完整、工作簿切换或后续人工修改时停止直接写回。

文件夹动作结果可携带 `changed`。执行器对能够确定比较的写值、写公式、写表格、
排序、清空内容或超链接、清除筛选、数字格式和冻结窗格进行本地前后状态比较；
全部动作均未改变目标时不进入文件保存与恢复点创建。暂时无法可靠比较的动作保守地
视为已变化，避免为了节省空间而漏掉恢复点。

达到 token、轮次、连续错误等硬限制后不再抛出裸异常，而是返回 provider 为 system、
outcome 为 failed 的 AnswerResponse，其中包含停止原因、压缩后的工具调用记录和针对截断、格式错误
或循环的建议。流式规划另受 `taskRecovery` 的时间预算控制：首次 60 秒，在完整模型响应
或完整单次工具调用边界持久保存现场并暂停；用户通过带 revision 和稳定 requestId 的明确操作继续，
每次增加 60 秒且累计不超过 300 秒。继续不会提高 token、轮次或错误上限。
OpenAI-compatible 适配器统一返回 payload 与 usage；缺失 usage 的服务会得到零值
token 统计。

### 任务进度与可审查判断轨迹

规划与两条执行通道统一使用 `TaskProgressEvent`。事件包含 `operationId`、单调递增
`sequence`、阶段、步骤、状态、可选计数、目标、`toolName`、`facts` 和 `trace`；
前后端分别在 `contracts.ts` 与 `models.py` 校验。`toolName` 由本地工具执行器填写，
`facts` 只允许工具级白名单摘要；前端仍会限制数量和长度、隐藏疑似凭据并把越界
绝对路径降级为文件名。`trace` 只记录可核对的理解、证据、假设、决策、冲突和
纠偏，不传递模型隐藏思维、提示词或只读工具返回的原始数据行。

`/api/turn/stream` 继续兼容旧 `step` 事件，但载荷已使用统一协议。Office.js
执行器直接向活动 Hook 推送预检、撤销数据捕获、逐动作、拆分子表和验证事件。
文件夹执行通过 `/api/folders/execute/stream` 推送预检、逐动作、逐子表、集中备份、
逐文件原子保存和验证事件。客户端断线时用同一 `operationId` 和最后收到的
`afterSequence` 重连。`FolderOperationLedger` 把请求指纹、单调事件、终态结果和错误
持久保存到本机应用数据目录的 SQLite；内存记录只是热缓存。完成/失败记录按
`taskRecovery.operationRetentionDays` 保留完整载荷，之后压缩成永久幂等墓碑，因此旧
`operationId` 不会在 TTL、容量淘汰或服务重启后变成一份新的写入。旧的非流式写入
入口统一拒绝执行。状态通过 `GET /api/folders/execute/status` 查询。

前端在本地保存当前待挂接的文件夹操作身份；任务窗格刷新后先查询持久状态，再用同一
操作标识重连。已完成结果可以跨服务重启重放；重启时仍是 `running` 或记录证据不足的
任务转换为 `needs_reconciliation`，不会自动重做 Excel 写入。用户已登记的工作目录由
持久索引定位，启动只扫描这些目录的受管恢复元数据，不扫描整盘，也不自动扩展授权。
执行、撤销和独立恢复共享规范化目录的进程内及跨进程排他锁；独立恢复还必须核对
manifest revision、生命周期、路径边界和写前/写后文件指纹。

这是 [任务断点恢复指导 v0.5](TASK_CHECKPOINT_RECOVERY_GUIDE.md) 的分阶段实现：阶段 A
提供持久身份、事件补收、安全挂起和恢复到执行前；B1 仅对单个
`splitGroupAggregate` 保存完整暂存产物和逐文件发布 checkpoint，重启后核对指纹并跳过
已发布文件；B2 保存流式规划的 messages、LoopState、工具缓存、开放工具模式、修复状态和
完整待处理模型响应，可在安全边界暂停并显式继续。阶段 C 进一步按单个工具调用保存完成
状态，允许用户在完整模型调用边界显式选择已配置备用连接，并为已保存工作簿中的
`writeValues`、`writeFormulas`（非 R1C1 铺满）、`setFill`、`setFont` 和
`setNumberFormat` 提供动作级续跑。任意 DAG、结构动作和复杂对象续跑仍未开放。

### 任务中断后继续（第一阶段）

上面的机制覆盖的是**执行**与**流式规划预算**的中断。第一阶段补的是它们之外、
用户最容易撞上的那一类：**需求已经说清楚、授权也勾好了，任务却因为断网、
服务重启、任务窗格重载而停在半路**，而界面上只剩一句"已失效，请重新发送需求"。
用户既没改需求也没改范围，却被迫从头再来。

阶段、阻塞原因、能否安全继续**是三件不同的事**，分开表达（共享纯语义层
`packages/core/src/taskResume.ts`）：

| 维度 | 类型 | 说明 |
| --- | --- | --- |
| `phase` | `awaiting_clarification` / `answer_submitted` / `reading_source` / `planning` / `preview_ready` / `executing` | 任务走到哪一步 |
| `blocker` | `network` / `user_input` / `authorization` / `source_changed` / `execution_unknown` / `storage` | **为什么停**（与阶段正交） |
| `resumability` | `checkpoint` / `replan` / `reauthorize` / `reconcile` / `unavailable` | 下一步**能安全做什么** |
| `writeState` | `not_written` / `confirmed_result` / `needs_review` / `unknown` | 原表写入状态 |

把"网络中断 / 等待回答 / 需要重新授权 / 执行结果不确定"一律归为"澄清卡已失效"，
会导致两种假话：范围根本没变却被告知"数据已经变化"，或者写入结果未知时被诱导重跑。
`taskResumeDecision()` 是唯一判定入口，**只有 `resumability === "checkpoint"` 才允许说
"从上次保存的进度继续"**；`replan` 一律说明会"重新生成预览"，不冒充精确续跑。

身份分层（复用既有字段，不新造并行账本）：

- `taskId` —— 用户任务的稳定身份，由加载项在一次**新的用户请求**时生成并随恢复记录
  持久化。同一任务内的重新规划、重新授权、重新生成预览都沿用它。
- `turnId` —— 一次规划尝试，仍由服务端 `planning_tasks` 权威持有；重新规划可以换
  `turnId`，但 `taskId` 不变。
- `authorizationRevision` —— 用户**显式确认过**的数据授权版本。它不是许可凭据：
  恢复时必须重新读取真实宿主元数据并重新确认。
- `operationId` / `executionId` —— 真实执行身份，沿用既有文件夹/Office 执行恢复，
  重试、查询、重连**不得**新建。
- `requestId` —— 一次继续动作的幂等身份，沿用既有 `PlanningContinueRequest`。

**持久化走显式白名单。** `sanitizeTaskResumeRecord()` 只重建显式列出的字段，
任何多余键（`plan`、`values`、目标矩阵、`sourceSnapshots`、宿主对象引用、完整路径）
都会被丢弃 —— "原始数据行 / 完整源值 / 目标矩阵不落盘"因此是**结构性**保证，
而不是靠调用方自觉。`originalPrompt` 为空时消毒返回 `null`，调用方据此明确反馈
"无法恢复"，不用问题摘要冒充完整任务。记录里也**不存在**"永久有效授权"这样的字段。

两端接线：

- **WPS**（`WpsAgentPanel.tsx`）：澄清卡与恢复卡是**同一张卡**的不同状态。面板重开时
  待答澄清仍按既有判据收敛为 `unconfirmed`（重开本身没有给出任何变化证据，
  不许说"范围已变"），但卡片不再只剩一句失效说明，而是渲染中断卡：**中断原因 /
  已保留的信息 / 原表写入状态**，加「确认此范围并继续」与「取消任务」。
  「确认此范围并继续」做的是一次**显式重新授权** —— 重新读取当前工作簿、工作表与
  表头，捕获新的对象身份，推进一版 `authorizationRevision`，再用"原需求 + 用户已确认
  的信息"发一轮**全新**请求。旧读取快照、旧计划与旧确认资格就此废弃，因此
  **重新授权之后不会再被同一个旧对象引用检查卡住**。若确认时读到的身份与建卡时
  不同（双方都读到才算证据），先明确告知"任务将应用于新的数据范围"，不悄悄转移授权。
  已提交的回答会写进恢复记录：续答那一轮断线时卡片转为可继续，继续时把原需求与
  确认内容一起重新发出，**不让用户重复回答同一个问题**。「取消任务」是终态
  （`resumability: unavailable`），无论怎样重开都不会再被自动恢复。
- **Excel**：`useScopeSelection` 把**草稿**与**已提交授权**分开，新增
  `committedSheetNames` / `scopeDraftDirty` / `confirmScopeSelection()` /
  `cancelScopeEdit()`。勾选、取消勾选、全选/清空只改草稿；「确认范围」才提交，
  且**提交时按当前真实工作簿清单复核**草稿（弹层开了很久时，已经删掉的表不会被
  提交成授权）。「取消范围编辑」把草稿放回已提交授权，一字不改 —— 用户点开只是
  想看一眼别的表时，不该把正在等待回答的任务打掉。
- **WPS 范围面板**（`WpsProbePanel.tsx`）同样改成草稿/已提交两态：勾选只改草稿，
  底部动作是「确认范围」与「放弃修改」（面板重开时草稿回到已提交范围）。
  授权版本只在**已提交范围**真正变化时递增。

第一阶段**支持**：等待澄清、已提交答案尚未生成预览、读取或本地计算中断、规划中断、
已生成预览但需要重新读取或重新授权这五类中断在两端保留任务与已确认信息，
并给出绑定 `taskId` 与原消息的继续入口。

第一阶段**不支持**（明确记录，不冒充已实现）：WPS 写入中断后的自动续写；
Excel 侧的图形化中断卡（共享语义层已就绪，Excel 本轮只落地了范围确认闭环与
`planningResume` / `officeResume` / `folderResume` 既有入口）；服务端
`planning_tasks` 尚未冗余保存 `taskId`（`taskId` 当前只在加载项侧作为任务稳定身份，
服务端仍以 `turnId` 为规划权威身份）；跨设备/跨浏览器恢复；豁免预览与写回确认的
"一键重做"。**恢复预览不等于确认写回**：继续之后仍然要经过白名单 `AnalysisPlan`、
预览和用户确认才会写入。

规划连接中断后，持久状态仍为 `running`、进程内没有活动执行者且存在完整 checkpoint 时，
对外解释为 `interrupted`。用户可明确选择能力匹配的已配置模型接管同一 `turnId`；状态查询和
继续接口使用同一资格结果，并在账本 CAS 领取与进程内活动登记之间保持互斥。前端续跑复用
原活动消息及 `operationId`，以客户端最后实际归并的事件序号请求补发；服务端最新序号只用于
校验上界，不能替代客户端游标。事件补发只恢复同一条时间线，不代表重新执行已完成工具。

Office 动作恢复由服务端 SQLite 账本作为权威现场：确认执行后、首个写入前登记计划、原始
授权和稳定工作簿 URL 身份；每一步在 `context.sync()` 前提交目标范围的写前快照，写入并
读取实际状态后再提交累计检查点。刷新或异常退出后，前端本地只保存 operationId 等发现
索引；继续前重新查询服务端、核对当前工作簿身份，并对在途范围做三态判断。只有完全匹配
预期结果时补记为已完成，完全匹配写前状态时重试，其他状态保持挂起。写入权通过 revision
和轮换 ownerToken 约束，状态查询不返回令牌。

活动 Hook 从 `started` 起就把步骤加入有序时间线，按 `operationId + sequence` 隔离
并去重事件，再按 `stepId` 原位更新。每段活动在消息流中拥有独立消息和稳定
`persistenceKey`；实时进度与完成日志更新同一个消息节点，禁止再通过“最后一条助手
消息”推断日志归属。用户消息会排在本轮过程消息之前，需求确认、计划和结果会追加在
过程消息之后，因而不会在完成时重排或跳位。旧记录中附着在普通助手消息上的
`activityLog` 继续兼容显示，刷新时会丢弃没有完成日志的临时活动占位。

实时卡和历史消息共用默认展开的单层时间线。完成后保存包含工具、动作、目标、安全
事实和判断轨迹的完整步骤快照；活动级 `outcome` 与最后一步状态分离，需求分析步骤
可以完成而整段活动显示“等待确认”。网络中断只记录连接状态并保留最后进度，不能冒充
失败、完成或已取消；验证失败使用 warning/conflict 表达，不把“文件已经写入”错误
显示成“任务未执行”。消息流在用户接近底部时自动跟随新增步骤；用户主动向上阅读后
不抢夺滚动位置，而是显示返回最新进度的提示。

### 规划暂停在面板重开后的恢复（CAP-02 第 5 条）

与上面「任务中断后继续」是**两条独立通路**，刻意不合并：

| | 任务中断后继续（第一阶段） | 规划暂停恢复（CAP-02 §5） |
|---|---|---|
| 处理对象 | 需求/授权已就绪但停在半路（断网、重启、重载） | 规划**主动暂停**等用户决定（预算/模型故障/被打断） |
| 权威索引 | 恢复记录（`taskId` + 原需求） | 历史消息里的 `planningPause.turnId` |
| 服务端现场 | `planning_tasks` / `office_executions` | `planning_tasks`（状态查询端点） |
| 是否碰过单元格 | 可能已写入（需核对 `writeState`） | **一个都没碰过** |

合并会让"只是规划暂停"也走写入核对，反过来"写入待核对"被当成可续跑的规划 ——
两类事实混在一张卡上，用户分不清到底写没写。

三层分工：

| 层 | 位置 | 职责 |
|---|---|---|
| 共享层（纯函数） | `packages/core/src/planningStream.ts` → `planningPausedEntriesFromHistory` | 从历史挑暂停条目：只挑 `planningPause`、同 `turnId` 取**最后一条**、载荷取不到不挑、空白 `turnId` 跳过 |
| 宿主适配层（查询） | `apps/wps-addin/src/wpsReadonlyAgent.ts` → `reinspectWpsPlanningPause` | 拿 `turnId` 查真实状态；**只查询不继续**；`payload` 优先用服务端状态，历史快照仅供显示 |
| 界面层 | `apps/wps-addin/src/WpsAgentPanel.tsx` 挂载核对 effect + `planningPauseReload` | 核对中不画操作区；核对通过才重建活跃暂停；读不到进度给保守说明 |

四条纪律（都有反向变异用例）：

1. **重开只核对，不自动续跑。** 自动续跑会在用户完全没点过任何东西的情况下触发
   一次模型调用；有副作用的事必须由用户点按钮发起。
2. **按钮只由服务端结构化状态推导。** 唯一允许产出按钮的地方是
   `wpsPlanningResumeView`（→ 共享层 `planningResumeDecision`），它直接吃
   `getWpsPlanningTaskStatus` 的返回。历史快照能推出按钮，但**不许**用它推。
3. **`revision` / `lastSequence` 一律采信服务端。** 历史快照里的值只用于显示。
4. **查不到时不下结论。** 既不说"能继续"，也不说"任务没了"，只给保守说明。

两个判据**分开写**：`canContinue` 决定"服务端是否放行"，`action` 决定"放行时走哪条
通路"。合并成一个条件会把"放行但不是精确续跑通路"误报成"该任务当前不能精确续跑"，
那是替服务端下一个它没下的结论 —— 面板手里并没有这个证据。

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

### 模型配置与选择语义

模型入口只面对用户自行配置的外部 API 模型，模型列表里不再有虚拟的“基础模式”
占位项。`/api/models` 在没有任何真实连接时返回 `models=[]`、`defaultModelId=null`；
前端把目录状态分为四种并分别表达：加载中、加载失败、加载成功但为空、已配置真实连接。
读取失败不清空已加载的列表，也不据此覆盖本机保存的选择。
WPS 新增模型连接时先打开可填写的表单，再异步读取本机共用的已保存连接；
检查失败与确实为空分开显示，且不会清除输入内容。同一模型 ID（即使服务地址也相同）
可以保存多条连接；目录按保存顺序显示 `(1)`、`(2)` 后缀，后缀不写入模型 ID。
每条连接具有独立的 `connection:<id>` 和 API Key；模型调用只能按此 ID 解析凭据，
不能按模型 ID 合并或选取第一条连接。

选择规则集中在前端 `modelSelection.ts`（Excel 与 WPS 共用），按“本次明确指定 →
本机保存且仍有效 → 服务端声明的默认 → 目录首个可用真实连接 → 无可用”的顺序解析。
界面展示的模型与请求携带的 `modelId` 都取自这一次解析结果，因此不存在“显示首项、
实际请求 `local`”这类不一致。旧的 `local` 记录会在目录成功加载后迁移到有效连接；
删除连接后重新计算选择，删除最后一个连接回到未配置状态。

服务端按入口契约决定是否需要模型，缺模型时抛出稳定错误并由 `/api/turn` 系列入口
映射为 `MODEL_NOT_CONFIGURED`（一条连接都没有，引导「添加模型」）或
`MODEL_NOT_AVAILABLE`（指定的连接已删除、失效或不在允许列表）。普通 Agent 请求
不会静默降级为规则模式，也不会擅自改用其他供应商；模型调用失败保持原任务状态并
返回明确错误。健康检查、模型设置、既有结果查看、确定性工具和恢复资料访问不受
模型校验拦截。

`_local_address_plan()`（清空区域、已知地址写值/写公式）与
`_local_bulk_worksheet_deletion_plan()` 只做字面解析，排在模型解析之前，因此不依赖
模型——它们是保留下来的确定性能力，不是基础模式入口的回归。`provider="local"`
始终表示“结果来自本地确定性工具”，与用户选择了哪个模型无关，历史消息照常展示。

### 通信协议

主要 API：

- `POST /api/turn`：统一轮次入口；规划完成可返回 `plan`、`answer` 或不含单元格值的 `local_plan_rule`。后者由客户端在已授权工作表内本地计算计划（Excel 与 WPS 共用 `packages/core` 的同一套确定性实现，仅宿主读取方式不同），仍须预览确认后写入。服务端只在 `planningReadPolicy: "none"`（客户端声明源值不外发）时下发规则；此模式下客户端送出的规划快照会剥掉数据行，服务端据此判越界、产规则，但拿不到任何源值
- `POST /api/turn/stream`：SSE 流式规划
- `GET /api/turn/{turnId}/status`：查询持久规划状态、预算和续跑能力
- `POST /api/turn/continue/stream`：按 revision 与稳定 requestId 显式继续暂停的规划
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
- `POST /api/folders/execute`：兼容的非流式文件夹执行入口
- `POST /api/folders/execute/stream`：带 operationId 幂等保护的 SSE 文件夹执行入口
- `GET /api/folders/execute/status`：查询可恢复的文件夹执行状态和最终结果
- `POST /api/folders/execute/resume/stream`：继续经核验的 B1 拆分任务剩余发布
- `GET /api/folders/backups`：列出当前工作目录的受管理备份
- `POST /api/folders/backups/export`：按 operationId 幂等取出副本到固定 `exports/` 目录
- `POST /api/folders/backups/delete`：按修订号和影响摘要复核后事务删除指定受管备份
- `POST /api/folders/backups/cleanup`：安全清理过期、孤立和超软上限备份
- `POST /api/folders/backups/legacy/collect`：确认后归拢旧版散落备份
- `GET /api/diagnostics`：诊断信息

前端协议定义在 `apps/excel-addin/src/contracts.ts`，后端协议镜像在 `server/app/models.py`。
AnswerResponse.provider 支持 model、local 和用于预算熔断降级的 system；
PlanResponse.provider 仍只允许 model 或 local。

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
- 用户将 Excel Bro 任务窗格拉宽到 380px 以上时，前端保存实际宽度；下次打开或重新显示窗格时，在 `TaskPaneApi 1.1` 可用的 Excel 中通过 `Office.extensionLifeCycle.taskpane.setWidth()` 恢复。旧版 Excel 不支持该 API 时继续使用手动拖动引导，不伪装成已恢复。
- 常用工作目录的扫描目录与工作表勾选会本地保存；首次产品引导的“以后再设置”也使用版本化 UI 偏好记录。重新使用目录时仍由后端扫描和校验，缓存不能代替文件访问检查。
- 设置工作目录只建立候选文件范围，不等于授权其中全部工作表；授权边界仍由用户的工作表勾选确定，首次设置不会自动勾选。
- 对 API Key 只保存到后端本地配置，不进入前端 localStorage。
- Office.js 待恢复任务在 localStorage 中只保存 operationId、对话定位和稳定工作簿身份；
  完整计划、授权、写前快照和动作检查点仅由服务端持久账本提供。
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

### 工作表授权边界

- 计划编辑器调用 `/api/plans/validate` 时，通过 `PlanValidateRequest` 同时发送本次授权工作表和当前现有工作表名称。
- Office.js 执行器在最终预检中以授权集合检查所有现有来源和目标表；新输出表只有在执行前不存在时才可越出授权名称集合。
- 文件夹执行器以会话内已选“文件 + 工作表”引用作为授权集合，并拒绝读取或覆盖输出工作簿中的未授权现有表。
- “仅保留某表、删除其他工作表”不会自动把未选表加入授权集合；前端先切换到多选范围并要求用户勾选全部涉及工作表，授权完整后才生成删除预览。
- 文件夹快照中的动作目标使用“文件 › 工作表”限定名；确定性结构命令可用 `sourceSheet` 解析用户输入的原始表名，但一次只能作用于一个已选文件，避免同名表歧义或跨文件误删。
- 授权检查必须发生在任何写入、备份或文件替换之前。

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

### 普通模型输出额度与截断（2026-10-03）

- 意图、规划、格式修复分别从 `intentContext.maxOutputTokens` / `agent.maxOutputTokens` /
  `agent.repairMaxOutputTokens` 读取 `null | 正整数`。默认 null：兼容客户端省略
  `max_tokens`，不以极小额度提前切断推理及最终结构化输出，也不硬编码供应商/业务类别。
- 单次额度与循环预算独立：agent.maxTotalTokens 现默认 null，关闭普通通道累计
  token 停止（此前 131072 已取消）。LoopConfig 默认同样为 None，避免备用调用路径
  暗中恢复旧 20000 上限；正整数可以重新开启，非法数值不可默默当作无限。
  开启时在下一模型轮前检查，不是正在生成时的硬费用封顶。用量与检查点计数继续保留，
  超过旧额度的已保存模型响应可以续跑，不通过清空计数假装任务获得预算。
  8 轮、连续结构错误、重复工具警告及授权/预览守卫仍保留，不承诺无限任务运行。
  模型 HTTP 默认超时 180 秒，现有环境覆盖优先；规划整体硬时限仍 300 秒。
- `finish_reason=length` 不等于格式错误：意图不采纳哪怕能解析的残片，走既有安全路由
  兜底；规划（包括恢复检查点中的 `_incomplete=length`）终止为
  `answer.outcome=failed, stopReason=output_length`，不执行工具、不缩小额度重试。
  Python 与两宿主共享 TS 契约同步。真正的格式错误仍执行有限结构修复。
- 默认仍受模型/网关输出与上下文容量限制，不宣称无限输出；SDK 独立预算机制未改。
  诊断仅记 finish_reason、输入/输出/总 token、可用时的 reasoning_tokens 和请求额度模式，
  不输出思考正文、用户文本或矩阵。reasoning_tokens 仅诊断，不再叠加进 total_tokens。
- 自动化证据：真实 HTTP 请求体覆盖意图与规划无 max_tokens、非空 messages/tools，
  本地敏感哨兵存在但不出站；35000 输出 token 的响应后仍能进入下一轮提交；
  截断的合法工具/JSON 不得接受。真实模型成功率与真宿主仍需实测。

- P1 重构总结：`docs/P1_REFACTOR_SUMMARY.md`
- Hooks 使用指南：`docs/HOOKS_GUIDE.md`
- 设计审查报告：`docs/DESIGN_REVIEW_FINAL_SUMMARY.md`

### 澄清续答与重新授权（2026-10-05）

目标确认前移：两端在需求分析入口调用共享 `checkForAnalysis`，已有强身份证据且授权一致时直接进入分析；弱包装对象、身份证据缺失或目标变化时，先暂停并显示独立目标选择卡。确认前不读取表头/数据行、不请求意图或规划模型、不创建需求分析活动。确认期间再次采集证据与元数据，变化仍拒绝。业务补全表单只做规则填写和既有授权复核，不授予新目标同意。WPS 同任务的显式前置同意只许可重新读取/规划；有匹配的私有目标键、实际表头与授权指纹时，可继续语义澄清及本地计算，但不把原生对象 unknown 改成强身份相同，执行器不消费此许可。新任务清除弱证据同意，路径/表身份/授权集合冲突仍暂停；写入强身份、快照和独立确认不变。文件夹模式、模型协议不变。详见 `docs/AUTHORIZATION_BEFORE_ANALYSIS_2026-10-05.md`；原内联勾选方案仅保留历史记录。

授权目标前置确认：两宿主共享内存基线、空闲元数据监测和操作入口检查。变化先弹目标/工作表选择卡，确认时重新核对；旧预览不能移植到新授权。此层不替代原有强身份及快照守卫；文件夹模式不变。实现范围、未实现项与宿主验收见 `docs/AUTHORIZED_TARGET_PREFLIGHT_2026-10-05.md`。

回答与授权继续分离，但不能先拦截授权、再丢弃答案：WPS 作答先存真实确认，下一轮携带原 taskId、原需求与全部答案。本地规则仅在内存按原消息定位；unknown 时显式重新授权后直接重读计算并生成新预览，不能重新启动原需求而反复问同一道题。Excel 身份不足同样保留本地规则并给出确认入口，读取新结构后才能继续。正面的路径、表身份或授权集合冲突仍拒绝旧规则，预览不能代替确认写入。详见 `docs/CLARIFICATION_LOOP_FIX_2026-10-05.md`。
