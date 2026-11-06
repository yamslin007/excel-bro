## 2026-11-06：macOS 双宿主安装包

Mac arm64 包同时构建 Excel 与 WPS 前端，共享一个冻结运行时及后端：Excel HTTPS localhost:3000，WPS HTTP 127.0.0.1 使用 capabilities.devServers.wpsPortCandidates 首项（当前 3889），不漂移端口。安装器以桌面用户身份将 Excel manifest 写入 wef，将 WPS ET 插件登记合并到 ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml；保留其他插件、注释与原配置备份。损坏 XML、实体声明及符号链接拒绝改写；卸载只移除本插件的 name/type 条目，保留用户模型配置。重开 Excel/WPS 后加载入口，WPS 版本须具备 JS 加载项能力。

PyInstaller 显式 collect-all claude_agent_sdk，包含平台原生 CLI；Mac 构建时执行冻结程序 --self-test，检查后端导入、双端资源、CLI --version，再在临时 home 验证注册/卸载，不启动 8765 或操作工作簿。产出 .pkg 及 SHA-256；当前未做 Apple 签名/公证。Windows 的打包/注册流程不在本次变更范围，已有 Excel-only 安装器须单独升级，不能据此宣称 Windows 双宿主安装已完成。

自动化已实现；Mac 双端实际入口显示、安装器权限、沙盒和 Excel/WPS API 兼容性仍待实机验收。具体构建结果见本轮交付。

## 2026-11-06：本轮提交前验证

本轮打包双端审阅/恢复与副本生命周期、对话与外观、统一模型连接和 SDK 协议适配的代码、测试、配置与文档；不纳入临时目录、运行数据、真实工作簿及测试产物。

提交前检查：core 全量 1261 passed；WPS 全量 1334 passed、8 skipped；服务端 2113 passed、1 skipped；Excel/WPS build（含 tsc）、服务端 compileall、暂存 diff --check 均通过。WPS 旧多表前态用例改为分别核实目标越界拒绝、缺少整表前态拒绝、按动作索引单表投影与真实副本重建，不放松运行时校验。

Excel 全量两次各有 1 项授权界面用例失败（首次为选择器待提交状态未在时限内解除，限制单工作进程后为文件夹选择器异步状态断言）；两次均为 1853 passed、1 failed、9 skipped。随后独立顺序复查 excelWriteEndToEnd.test.tsx 与 folderScopeSafety.test.tsx，119 passed。此处保留全量失败记录，不将定向复测等同于全量全绿。日志 outputs/git-round-*.log；本轮未做原生宿主实机验收。

## 2026-11-05：撤销后的颜色标记清理

Excel 补齐旧规则 id 失效时的精确归属回查和删除后回读，WPS 补齐删除后回读；两端外部撤销等待审阅收尾。清理失败保留本次规则记录，通过既有入口重试清理，不再次恢复单元格。只认本次随机公式，不删除用户原有颜色或其他规则。

自动化覆盖规则 id 变化、归属歧义、静默删除失效、其他规则保留、撤销反馈等待收尾，以及真实组件放弃后规则数归零。日志 outputs/undo-marks-*.log。宿主接口由替身提供，不代表实机通过；真实双端需在独立测试工作簿检查撤销后颜色消失、业务格式保留。刷新前端窗格加载修复，无需重启后端；已记录的旧残留会走现有重载清理，不扫描或抹除无归属的黄色。

## 2026-11-05：旧撤销记录不注入新对话

WPS 原有兜底在找不到 verification.executionId 匹配消息时，把发现的终态恢复记录追加在当前消息流末尾，导致欢迎页/新对话也常驻旧撤销摘要。已移除该兜底，只在原任务匹配的执行消息内保留 RecoveryTerminalRecord；不清账本、历史、恢复材料，不改变恢复判据或操作。Excel 没有该兜底，无须添加新展示。待核对/阻断/未完成的恢复入口不受影响。

回归覆盖有归属与无归属两类已撤销执行后发送普通聊天、已完成任务连续重开、新任务仍可审阅；保持零新增写入和恢复材料保留。日志 outputs/orphan-terminal-*.log，两端构建通过。真实宿主未操作，刷新 WPS 窗格后验收；无需重启后端。

## 2026-11-05：普通聊天后的空待办计数

两端 pushPending 在构造 TaskExecutionNotice 节点时就登记任务状态，但该组件在空闲、普通发送等状态返回 null，产生“待处理 1 项 · 任务状态”的空壳。共享 taskExecutionNoticeText 同时驱动父级登记与组件渲染，有实际状态文本才计数。暂停、停止中、不可中断及审阅说明仍保留；真实清理/恢复事项继续由各自入口登记。WPS 的已撤销摘要是消息历史，不属于待办，保留其记录和展开入口，没有删除账本或恢复材料。

Excel 提示组件和真实 App 结构测试 13 passed；WPS 面板 254 passed、5 skipped，包含已撤销任务后普通聊天无待办栏且历史摘要保留。两端 build（含 tsc）通过，日志 outputs/empty-pending-*.log。未操作真实工作簿，原生宿主待刷新窗格验收；本轮只需刷新前端，无需重启后端。

## 2026-11-05：放弃恢复与展示副本收尾分离

已证实：Excel closeCompatReviewCopy 为了满足旧的活动表匹配而激活 session.original，可能把用户从最后比对的子表带走；窗口桥 close_copy 又用活动子表匹配多子表副本，原表/副本标签不同会返回 identity_lost。正常、续跑、刷新恢复三条 Excel closeBefore 原先丢弃具体关闭状态，续跑 restore 也漏了 discard 来源标记；useUndoSnapshot 的底层成功消息与 ResultReview 终态消息重复。共享 closeBefore 在保存副本登记时抛错会越过清理分类，被外层误报成放弃失败；已有结构化清理清单缺少独立重试入口。

本轮关闭采用 v2 隐藏 nonce + 未保存身份 + 唯一副本 + 原簿路径哈希，不再依赖活动子表。最终仍由 Windows 适配器复核持有的 Workbook、隐藏标识、无路径、格式及窗口身份，先 SaveCopyAs 归档再 Close(false)；另存、重复/缺失身份及不可核实关闭仍保留。旧 v1 nonce 表签约束不放宽。查看、排列、切表接口继续保留各自的表级校验。前端关闭不再调用 sheet.activate，因此不制造收尾导航、也不更新授权基线。原生关闭前复用身份核实过的 suspend_pair 恢复原窗口位置并解除并排，再重新核实双方身份；失败保留副本，不伪造退出成功。恢复后的 scan 改为被等待，使数据范围栏刷新完成后才结束恢复调用。

共享层按宿主副本身份对清理调用去重，不只是对提示去重；清理、登记存储错误保持 discarded/confirmed 业务终态。新增 retryCopyCleanup，仅终态、非处理中、审阅决定已落账才可调用；retryDecision 与 retryMarkerCleanup 保留。界面为三种收尾提供对应重试，清掉其他 notice 不会藏起仍存在的 copyCleanup。已恢复但登记失败时不再次恢复，未能核实完整恢复时不发布 discarded。

Excel 逆序逐动作恢复后立即回读本动作涉及的公式/值、数字格式、逐格格式属性、行高和列宽；允许 null/空字符串表示同一个空格，其他内容仍严格比较。重叠动作按各步中间前态核对，避免误拿早期快照与最终状态比较。回读失败提示可能已有部分恢复，不声明全部完成；原有工作簿、授权、写后现场、指定恢复栈顶校验不放宽。Reviewed discard 不再广播底层“已撤销上一次执行”，只由审阅终态发布一条主结论；真正撤销入口继续原有消息。WPS 原本已区分 undo/discard，保留其恢复和整本副本内容保护，补齐副本身份去重、独立重试及具体关闭原因。

当前工作区已有 ComposerPendingBar + composer-dock 两层结构，保留：通知在整个输入区上方，工具栏和输入框相邻，没有把清理提示插入其中；保留壁纸、模型和输入框样式。

自动化证据（独立内存工作簿/宿主替身，不是实机）：共享审阅 58 passed；Excel 恢复/格式/原生副本/真实 App 端到端/工具栏布局七文件 218 passed，后续新增独立回读失败用例与重试按钮用例另有定向日志；WPS 原生副本/执行器/真实面板 348 passed、5 skipped；窗口桥 pytest 115 passed（含关闭前恢复原窗口位置、活动子表不变）。覆盖多子表、不同活动标签、退出重进、刷新接回、重复放弃、关闭失败重试不再恢复、另存/已关闭/身份丢失、真实范围变化仍拦截及确认不重写。双端 320/400/500px 浏览器布局检查通过。日志 outputs/discard-fix-*.log；追加界面测试 17 passed、恢复回读测试 14 passed；最终 build:addin/build:wps（含 tsc）均通过，compileall 通过。上述定向用例有重叠，不累加成全量测试数。

Excel/WPS 原生实机均未验证。本轮未操作用户工作簿、未删除归档/账本、未启动额外 8765 服务、未提交 Git。需要重启已有后端以加载新的关闭匹配规则，并刷新两端窗格；不得为验收重放原用户任务。独立测试簿实机待验：三张结果表中切到结果3、原副本异表标签放弃，检查原表值/公式/格式与恢复点一致、副本按身份清理、当前表与范围栏一致；再验关闭失败重试与实际换簿仍受守卫保护。

## 2026-11-05：双视图审阅说明折叠与异常保留

本轮仅调整展示字段与说明归属：ResultReviewState.comparisonHelp.scope/layout 分别承载常规范围说明和布局帮助，notice 继续承载操作异常；不按中文内容判断隐藏。Excel 的静态值、基础格式/合并区域及不含图表/条件格式说明、WPS 的冻结值与基础格式/列宽及复杂对象来自打开时原表的限制，均进入「查看详情 → 比对范围」。自动调窄失败和 Excel 自动并排失败的手动帮助进入「布局说明」，默认折叠。主操作区保留状态、选表、退出、确认/放弃；间距由 10px 收至 6px，不预留说明占位。未增加方向标识或新按钮。

WPS 缺完整前态、标记清理无法核实、身份/切表/恢复失败继续由 notice 或 compareUnavailable 显示；没有统一隐藏 notice。现有执行、授权、确认、恢复和副本清理调用顺序未改，退出继续保留副本。

验证：共享审阅测试 52 passed；Excel 审阅组件及原生适配器 56 passed；WPS 原生适配器与面板 299 passed、5 skipped；build:addin/build:wps 均通过。双端样式浏览器探针覆盖 320/400/500px，三颗审阅按钮均在视口内、八张长表名 3/3/2 分页无横向溢出、长错误独立滚动、详情默认折叠。日志 outputs/review-help-*.log。首轮 WPS 测试依赖旧通用说明失败，补回详情中的通用说明后全文件复跑通过；新增 Excel 范围断言最初放到旧 staged 路径，移至真实 before-result 路径后复跑通过。

两端均已实现；原生实机待验：独立测试簿中查看正常比对、自动宽度不可用和真实错误时的提示归属，展开范围/布局说明，检查窄窗格按钮。未操作用户真实工作簿，未重新执行任务或清理审阅状态；浏览器夹具不能代替原生验收。

## 2026-11-05：多子表副本与切换验证

新比对为一份副本包含全部目标子表。Excel 新增 select-sheet 本机桥；更新后在现有服务终端 Ctrl+C，再运行 `npm run dev:server`，刷新前端即可，不重新执行用户计划。不启动第二个隐藏服务占用 8765。

专项测试覆盖：一份 XLSX 多子表、冻结前值/共享格式、一次创建、切表不排列/不暂停/不重复保存、退出复用、恢复认领、旧记录升级、确认后单次关闭、身份/授权错误拒绝、WPS 非当前子表额外修改保留。测试仅使用合成数据及宿主替身；独立 openpyxl 解析合成 XLSX 也验证通过，不能据此宣称原生切换耗时已达标。

原生验收：独立测试簿执行一次三表任务，打开后仅一份三子表副本；结果1→结果3→结果1 两侧同名，窗口宽度不反复变化；退出/重开复用；确认/放弃成功后关闭，任一恢复失败继续保留。完整记录见 [双视图专项](DUAL_VIEW_RETENTION_2026-11-05.md)。

## 2026-11-05：Excel 对比副本表签修复与本机服务更新

用户截图的 Excel 副本表签为 EB_nonce，而 WPS 保留“结果1”。现已分离显示名与隐藏身份；兼容恢复查看旧 nonce 副本，仅更新该临时副本标签和固定隐藏名称。旧服务 OpenAPI 实测缺失 suspend/resume，导致退出失败；已核验本项目进程后按既有 SDK 启动脚本重启，健康检查通过，现有两个接口与 copyIdentityVersion 1/2。

Excel 专项 270 passed（含逐表名称对应、复用不重写、请求只含元数据）；窗口桥 112 passed（含标识缺失/错误、重名/另存拒绝、旧副本迁移）；WPS 原生比对替身 45 passed。两端生产构建通过。独立 openpyxl 解析生成的合成 XLSX，核对单个“结果1”可见表签、原值与隐藏常量标识均通过。后端编译及差异检查通过。日志 `outputs/copy-label-*.log`。

未实际修改用户工作簿内容，未自动确认/放弃；原生只读窗口元数据检查不替代表签迁移/并排交互的实机验收。完整实机流程仍待验，已有副本可退出再进入以触发核对升级，无需重发写入任务。材料或身份不明时继续保留待审阅并提示原因。

## 2026-11-05：任务状态栏垂直居中

截图中的暂停按钮偏低来自独立按钮遗留的 `margin-top: 8px`；圆点共用动画含向上位移，窄窗格还覆盖成顶对齐。共享状态栏内取消按钮外边距，统一按钮居中、字号及行高，圆点改为固定中心的透明度动画，窄窗格保留垂直居中。仅修改 styles.css 与 appearance.css，不改变任务行为。

Excel/WPS 均加载这两份样式，已实现。真实 Chromium 加载各端样式，验证有效宽度 320/420/600/1000px、100%/150% 缩放共 16 组：按钮中心对齐、圆点不位移、无水平溢出；已有 TaskProgressTimeline 测试 26 项通过，两端构建通过。浏览器证据在 `outputs/probe/status-alignment/`，原生宿主视觉仍待刷新窗格后验收。

## 2026-11-05：Excel 前态上下文对齐与双端副本保留

实现、限额、失败语义及分端验收见 [双视图专项记录](DUAL_VIEW_RETENTION_2026-11-05.md)。须重启本机后端（新增 suspend/resume 路由）并刷新前端。旧任务缺写前整表快照时，不可重发任务来补造历史。自动化使用宿主替身；浏览器验证本机持久存储；本轮未修改用户工作簿，原生窗口行为仍须在独立测试簿验收。

## 2026-11-05：重开审阅后无法放弃修复

复现：Excel 执行完成、未处理审阅，销毁前端并清空缓存，再从账本找回审阅；放弃报“本次执行不是当前恢复栈顶”。审阅回执已恢复，但内存撤销栈为空。WPS 同类回归也复现：恢复包水合成功，但旧 undoRecord 为空而拒绝恢复。两端已接回真实恢复材料，保留指定执行、现场冲突与撤销顺序保护。

| 端 | 实现 | 自动化验收 | 原生验收 |
| --- | --- | --- | --- |
| Excel | 已实现 | E2E、撤销 hook、授权范围测试 140 passed；生产构建通过 | 实机待验 |
| WPS | 已实现 | 面板、执行器、恢复检查测试 357 passed、5 skipped；生产构建通过 | 实机待验 |

覆盖重开后确认不重写、放弃恢复三格、后续人工修改零覆盖；Excel 另验其他任务栈顶与核对期间栈变化拒绝，WPS 另验水合旧记录保持栈顺序。日志 outputs/review-discard-excel.log、outputs/review-discard-wps-final.log、outputs/review-discard-build-*.log。没有修改用户工作簿、账本或恢复包。

原生最短验收：独立测试簿完成多步写入后刷新窗格，找回审阅并放弃，核对所有目标恢复前值且没有新增执行；另测刷新后手改一个目标，应拒绝覆盖并保留审阅。不要重发原任务来恢复。

## 2026-11-05：双视图目标定位与窗格接入修复

用户报告 Excel 可以打开双视图，但内容/排列不正确。代码定位到两处断点：执行后 ResultReview 未接入旧流程已有的窗格比例管理；Excel 打开副本前未激活所选比对表，而原生排列桥按工作簿与活动工作表共同匹配原表。

- Excel：核实工作簿及快照工作表 ID/名称，激活所选表并定位修改区域；副本设置同一初始活动单元格。新审阅副本接入宽度管理，打开期间禁止普通聊天偏好覆盖，退出恢复普通宽度。
- WPS：既有目标表激活保留，将 ResultReview 的当前副本接入现有宽度/租约适配器。
- 共享：仅通过内存 getter 提供当前比对句柄，不持久化；不再比对或进入终态即释放布局接入。宿主拒绝自动调宽时，审阅区显示手动调整提示。

| 端 | 实现状态 | 验证证据 | 尚需原生验收 |
| --- | --- | --- | --- |
| Excel | 已实现 | 新增单元测试：活动标签错位、同名表替换拒绝、执行前值不被新结果替代、无新增写入、退出恢复宽度；原生副本/OOXML/比例/输入区专项通过；生产构建通过 | 本机 Excel 打开、多表切换、窗口排列与实际任务窗格宽度 |
| WPS | 已实现同类宽度接入 | 副本/执行前快照/宽度/真实面板替身测试 333 passed、5 skipped；生产构建通过 | 本机 WPS 并排与退出宽度恢复 |

共享 ResultReview 44 项通过。Excel 全量首次 1810 passed、9 skipped、1 超时；随后端到端文件复跑出现一次选表授权等待超时，再隔离运行 103/103 通过；最新提示用例 8/8 通过。保留全量运行的不稳定性记录，不宣称全量首次通过。日志 `outputs/dual-view-fix-excel-tests.log`、`outputs/dual-view-fix-e2e.log`。

原生桌面查看工具两次启动失败（kernel exited / helper_unknown_error），本次未实机验证，也未修改真实工作簿；自动化使用有状态宿主替身。布局操作不执行 AnalysisPlan，不改变确认/放弃边界。

## 2026-11-04：选表面板防遮挡

Excel 与 WPS 的选表面板均挂载到聊天外壳的独立 `.scope-layer`，边界由窗格高度决定，不再依附输入框向上定位。标题与操作区固定，正文为唯一滚动区；工作表及文件列表不再套第二层纵向滚动。宽窗格并排显示跟随当前/选择多个，窄窗格自动堆叠。勾选即授权和确认关闭的语义保持不变。

| 端 | 实现 | 验证 | 剩余验收 |
| --- | --- | --- | --- |
| Excel | 独立层、共享高度与滚动规则 | 真实 App 层级/选择测试、生产构建、浏览器布局夹具 | 原生任务窗格内长草稿、缩放、重开选表待验 |
| WPS | 移除独立高度上限、共享高度与滚动规则 | WpsAgentPanel、选表/工作目录测试、生产构建、浏览器布局夹具 | 原生任务窗格内长草稿、缩放、重开选表待验 |

`node scripts/check_scope_overlay_layout.mjs` 加载生产样式，以两端行结构构造离线布局夹具，检查 320×480、400×360、740×760、1280×900 与 72/300px 输入区的 16 种组合：标题和确认完整可见、只有正文滚动、末行与确认可达、无横向溢出。截图和测量位于 `outputs/probe/scope-overlay`。这是浏览器布局证据，不等同原生宿主验收；本次未操作真实工作簿。

## 2026-11-04：双端阅读展示

两端共用 ConversationMessage 与 conversation.css，助手正文、过程状态、澄清问题、结果操作以连续对话呈现；壁纸阅读底色按内容占位，空白画布透明。样式顺序为 styles → 宿主补充 → chatScrollbars → appearance → conversation。新增展示形态应进入共享层，不在两端面板另造消息外壳。刷新 Excel/WPS 窗格生效，原生宿主观感待验。[设计与验证](APPEARANCE_PREVIEW_2026-11-04.md)。

## 2026-11-04：壁纸与外观验收

Excel 前端全量 1807 passed / 9 skipped，WPS 1323 passed / 8 skipped；最后异步互斥补充后 useTheme 9 项复跑通过，两端再次构建。node scripts/check_appearance_layout.mjs 覆盖两端五档尺寸、样式一致、取消、IndexedDB 事务中断、保存及重载；截图 outputs/probe/appearance。刷新窗格 → EB Logo → 壁纸即可使用。原生宿主最终观感仍待验，未操作真实工作簿。[完整记录](APPEARANCE_PREVIEW_2026-11-04.md)。

## 2026-11-04：通用 OpenAI 兼容接入验收

后端更新后重启原启动脚本，再刷新 Excel/WPS。连接选择 OpenAI 或 Anthropic 兼容，按供应商文档填写基础地址、模型与 Key，点击统一“测试连接”；不再选择用途，文本正常回复即可使用。两种协议文本及公式调用：`test_unified_model_connection.py`；共享转换与真实本机 SDK 链路：`test_openai_sdk_transport.py`；双端前端覆盖旧 Agent 元数据为 false 时仍选择并发送原连接 ID。开发用 SDK 工具探针仍保留，不能把连接测试当作全功能证明。具体证据和宿主验收步骤见 [统一连接验收](UNIFIED_MODEL_CONNECTION_2026-11-04.md)。

## 2026-11-03：模型工具栏验收

两端模型入口位于 `.composer-toolbar`，顶部不再重复。验收：空目录显示「模型」；配置或切换后显示上游模型 ID 与对应厂商 Logo；选择仍保存连接目录 ID；长 ID 在窄窗格省略；菜单向上展开且不越界。旧版目录缺 `modelId` 时，托管连接按连接设置精确补齐，失败时保留完整连接标签。定向回归：`modelCatalogCompatibility.test.ts`、`components/ModelPickerLabel.test.tsx`；浏览器探针 `check_scope_entry_layout.mjs` 覆盖两端三档宽度和 DeepSeek Logo。原生 Excel/WPS 刷新、切换后视觉效果仍需实机验收。

验证：Excel `styles.test.ts` / `messageFirstRouting.test.tsx` 共 91 passed，模型切换路由断言单独复跑 1 passed；WPS `WpsAgentPanel.test.tsx` 247 passed / 5 skipped；服务端模型连接与 SDK 目录相关 67 passed，compileall 通过。两端构建通过；`node scripts/check_scope_entry_layout.mjs` 检查两端 320/400/500px 布局。真实 Excel/WPS 窗格仍待验，本轮不操作工作簿。

## 2026-11-03：重开窗格后待审阅还在——怎么跑、怎么验

缺陷与设计见 [审阅状态持久化](REVIEW_STATE_PERSISTENCE_2026-11-03.md)。本机跑法：

```
# 需要后端与账本一起升（schema v14 会加 review_state / review_state_at 两列，不回填）
./.venv/Scripts/python.exe -m pytest server/tests -q -p no:cacheprovider --basetemp="$(mktemp -d)"
cd packages/core && npx vitest run
npm run test:addin
npm run test:wps
npm run build:addin && npm run build:wps
python -m compileall -q server/app
```

定向复跑（改这两端任何一处审阅路径时）：

```
npx vitest run src/excelWriteEndToEnd.test.tsx -t "重开窗格后审阅入口回来"
cd apps/wps-addin && npx vitest run src/WpsAgentPanel.test.tsx -t "执行完成但没处理审阅"
```

真机步骤（**本轮未做，待验**）：

1. 重启后端、重新加载两端加载项（账本加列在启动时迁移）。
2. 跑一次会真正写入的小任务，等过程框结束、出现「确认全部变更 / 放弃全部变更」。
3. **不要点任何按钮**：刷新任务窗格（Excel：重新打开加载项；WPS：关掉窗格再打开）。
4. 期望：审阅入口回来了，提示里没有"本次没有需要你处理的后续操作"；
   原表内容与刷新前逐格一致（没有被重写）。
5. 点「确认全部变更」：只应出现"已确认变更"的收尾，原表内容不变；
   再刷新一次，审阅入口**不再**出现。
6. 反向：另跑一次，第 3 步刷新后点「放弃全部变更」，原表应回到执行前值；
   再刷新不出现审阅入口，重复点击不重复恢复。

已知不可重建：双视图（执行前快照只在上一次窗格内存里）；WPS 的颜色标记与"保存为工具"资格
（标记公式只在上一次窗格内存里，账本也不存单元格值）。Excel 会按计划重新加条件格式标记，
那不改单元格内容。

## 对话窗格加宽入口

Excel/WPS 的“更多”菜单提供“加宽对话窗格”，不访问单元格。Excel 使用 TaskPaneApi 1.1，缺少能力时提示手动拖动；WPS 仅使用当前已登记、可见且右侧停靠的任务窗格，并从当前原生宽度和 WebView 宽度换算单位。请求幅度 160 CSS px，上限分别为屏幕可用宽度或宿主窗口宽度的 70%；宿主可能不采纳请求，界面不承诺调整成功。双视图自动宽度管理期间拒绝本入口，避免覆盖布局。

自动化：WPS 宽度及双视图租约定向测试 30 条通过；双端真实宿主是否接受宽度请求仍待验。当前双端构建受既有恢复发现类型错误阻断，不能记录为构建通过。

## 相同公式的结果确认与中断判定

真机确认提示“写前写后的公式文本相同”定位到 WPS 确认误用 inspectWpsUndoRecord。新增 inspectWpsConfirmedResult 供普通/恢复完成后的两个确认入口：要求记录 verified、全部动作 appliedRows 等于 rowCount，再只读核对当前结果。公式矩阵中每格为 written 或 indeterminate 时符合预期；混入 before/unknown 仍拒绝。inspectWpsUndoRecord 与公式续跑闸保持原判据，不能凭相同公式猜测写入历史。Excel 仍用 assertUndoPostState 核对写后状态，无需放宽恢复判据。

自动化覆盖相同公式、相同与不同公式混合、目标被改回写前/第三种公式、未完成记录及缺失记录；确认不写单元格、不改变恢复记录。原生 WPS 当前审阅轮仍待用户点击验证。

## 2026-11-02：确认失败保留核对原因

ResultReview.confirm 按真实失败阶段判断 markerPending，只有 clearMarker 失败提示重试清理。validateResult 失败保留具体原因、不接受结果、不重复执行。WPS 普通与恢复后的确认入口均透传 inspectWpsUndoRecord.detail，不再把区域冲突、公式写前写后相同、读取失败及记录缺失合成“结果已变化或无法核对”。这不改变核对判据；当前真机究竟命中哪一个 detail 仍待获取，不能据笼统文案断言用户改过结果。

## 2026-11-02：审阅工作表分页搜索

Excel/WPS 共用 ReviewSheetPicker：每页最多三张表，中间纵向显示完整名称，两侧仅一组上一页/下一页。三张以内隐藏搜索与分页；超过三张显示搜索，匹配结果仍每页三张。搜索和翻页不派发 comparison，点击名称才调用既有对比接口。中文组合输入结束后筛选，清空搜索回到当前对比表所在页。列表不设第二层滚动，审阅动作行继续固定。

自动化覆盖八表 3/3/2、搜索无匹配、清空定位、IME、忙碌禁用、双端切表零重写。浏览器探针 check_recovery_composer_layout.mjs 使用八张长表名检查三档宽度。真实 Excel/WPS 窗格仍需查看；未改变写入、授权或恢复逻辑。

## 2026-11-01：SDK 主线证据对齐

主线计划新增 §0 当前队列，A1 剩余表修正 Excel 公式 A/B、同会话/重开撤销、保存文件与 undone_at 的过时未验结论；S1 保留 10-26 双端注入原生证据，不能外推全部手动改源场景。WPS 两表值复制终态九步清单与四表公式 A/B 分开，不再互相引用为同一夹具。

发现 copyRange 已有精确续跑用例，复跑 `cd apps/wps-addin; npx vitest run src/wpsExecutionResume.test.ts -t "copyRange 的精确续跑"`：1 passed / 65 未选中；这里只覆盖 values 展开、前两步写后中断、续跑仅 Target!A3，生产实现未改。未重跑全量/build，未重新操作真实工作簿或读取原账本。A1-GATE 仍未通过，A2 未启动。

## 2026-11-01：对话定位条验证

悬停目录调整：单条浮动预览替换为向左展开的可滚动摘要列表；行尾保留刻度，移出/Escape 收回，点击才跳转。目录 overscroll-behavior: contain，不带动消息流。共享组件与滚动 hook 10 passed；两端 build 通过；六十轮浏览器夹具在两端三档宽度验证展开边界、独立滚动、点击跳转和新内容不抢位置。日志 outputs/conversation-nav-menu-*.log，截图仍在下述目录；原生宿主待验。

两端接入共享 ConversationNavigation。Excel 的 ConversationNavigation / useMessageAutoFollow / messageFirstRouting 定向共 71 passed；另 Fast Refresh 回归 2 passed。WPS WpsAgentPanel 232 passed / 5 skipped。两端 build 通过（含类型检查），仅既有大 chunk 提示。日志 outputs/conversation-nav-*.log。

`node scripts/check_conversation_navigation_layout.mjs` 使用真实共享组件、滚动 hook 与双端样式，验证 320×640、400×720、500×760 的右侧留位、预览边界、输入区不遮挡、点击定位及新消息不抢滚动；截图在 outputs/probe/conversation-navigation。浏览器夹具通过不等于原生宿主验收。原生 Excel/WPS 待查看：两轮以上对话出现刻度，悬停展开目录、滚轮浏览、点击旧轮保持位置，回到底部恢复跟随，切换会话不串位置。本轮未操作工作簿，未修改后端或协议，未提交推送。

## 2026-11-01：全部积累改动提交前检查

执行了 test:addin / test:wps / test:core、两端 build、服务端全量 pytest 和 compileall。核心 1101 passed；双端构建通过。首次 Excel 全量 1676 passed / 2 failed / 9 skipped，WPS 1239 passed / 22 failed / 8 skipped：失败为问卷替换后的旧 DOM/文案定位断言，按新位置修正，保留身份、授权、图片隐私与零写入断言。Excel 两个失败文件复跑 31 passed；WPS 四个失败文件先 45 passed / 1 failed，最后图片用例改等待实际问题而非已移除摘要后文件通过。面板凭据用例改用明确的短假凭据，相关三项复跑通过。没有把定向复跑记成整套重新全绿。

服务端首次全量 1949 passed / 28 skipped / 2 failed，失败在 Python 事件投递（Docker 引擎未运行）。用户启动 Docker 后，该文件 6 passed。日志 outputs/git-all-*.log、git-fix-*.log、git-docker-retest.log、git-credentials-final.log 留在本机，不提交。暂存内容排除临时目录、代理本机资料、独立 SDK 仓库、真实工作簿和账本。推送前精简：17 份根目录历史报告、5 个一次性取证脚本及根目录手动 SSE 测试入口仅留本机，不纳入提交。正式测试、可复用回归脚本与设计/验收文档保留。

## 2026-11-01：三个选项与固定问题标题

共用配置 clarification.maxOptions=3；生成提示要求最多三项、短标题和一句说明。TS 与 Python 新增三项接受/四项拒绝且不修改原数组的断言。共享组件增加 clarification-choices-scroll，问题标题留在滚动区域外，其他输入及操作行仍固定。Excel/WPS 共用此组件与协议。定向 TS 32 passed，服务端消息理解 30 passed；双端三档浏览器布局通过（包括滚动前后标题位置不变）。日志 outputs/clarification-three-*.log。已存在的四项问题不被截断或重写；用新问题验收新生成规则。

## 2026-11-01：澄清问卷标题与操作行收紧

“需要确认”与问题同排；去掉未选择时的通用提示和回答模式底部提示，字段缺项校验仍保留。取消任务从滚动区移到提交回答旁，使用按钮样式。共享样式双端应用。浏览器布局脚本验证三档宽度下标题同排、取消/提交同排且可见。Excel 消息路由 61 passed、共享组件 7 passed；WPS 面板 232 passed / 5 skipped；双端 build 通过。组件旧 radiogroup 断言改为框选按钮组的“单选/可多选”标签断言，与实际 aria-pressed 按钮一致。日志 outputs/clarification-actions-*.log；原生宿主仍待用户查看。

## 2026-11-01：澄清问卷替换输入框验证

Excel 定向组件、消息路由与 contracts：99 passed；WPS 面板：232 passed / 5 skipped；服务端消息理解与身份：45 passed。新增两端真实面板测试覆盖框选/其他互斥、点选不提交、消息澄清回答不授权读取和执行。两端 build（含 tsc）通过。日志 outputs/clarification-replace-*.log。

node scripts/check_clarification_inbox_layout.mjs 使用真实共享组件、主按钮与双端样式，在 320×640、400×720、500×760 验证问卷在 composer-box 内、自滚、无圆形单选标记、“其他”输入框、无横向溢出及提交按钮可见。截图 outputs/probe/clarification-inbox.png。这是浏览器夹具，原生 Excel/WPS 窗格尚待验；未操作真实工作簿。后端新增消息理解 clarification 协议，真机须加载更新后的前后端，不只刷新旧服务的前端。

## 2026-10-31：历史对话对齐回归

双端复用 HistoryDeleteDialog 与 core/conversationHistory；WPS 现有 30×120 移入 capabilities.conversation.wps，不能直接改用 Excel 20×60。定向测试覆盖取消/指定删除、其他保护标记、30×120 往返、重开普通续聊及零宿主写入；既有失败提示与历史计划失效测试继续保留。完整命令、真实结果和无敏感聊天真机步骤见 [验收记录](CHAT_HISTORY_ALIGNMENT_2026-10-31.md)。实际加载 URL 必须在退出前后记录，不将测试端口当作真实地址，不清缓存/删键。无需重启后端，A2 未启动。

## 2026-10-31：重开后撤销未完成执行（Excel）怎么跑、怎么验

跑：`cd apps/excel-addin && npx vitest run src/excelWriteEndToEnd.test.tsx -t "重开后"`（五条：
材料恢复入口 / 目标被改过 / 授权收窄 / 已处置 / 材料缺失）；`npx vitest run src/hooks/useUndoSnapshot.test.ts`
（去重、覆盖、`ensure` 之后同 tick 撤销、失败保留重试）；`npm run test:core` 里的
`executionUndoRebuild.test.ts`；服务端 `pytest server/tests/test_execution_disposition.py -q`。

变异：`node scripts/mutate_office_undo_rebuild.mjs`（先打印基线 sha256 与"工作区是否干净"），
一次一个：`U1`（材料挂不进栈）/`U2`（把 needs_reconciliation 当 completed）/`U3`（服务端不投影
disposition）/`U4`（同一执行两份快照都留在栈里）。每个都按四判据判定（退出码、预期用例红了、
失败原文带自家证据、还原后逐字节一致）并**还原后复跑绿**；脚本装信号钩子，结束必须四个目标全部
`clean`。pytest 的 `--basetemp` 固定在 OS 临时目录（仓库内的临时目录会被环境里的安全删除钩子牵连）。

判据边界（改之前先读）：`packages/core/src/executionUndoRebuild.ts` 是"重开后能不能撤销"的**唯一**
出处，入参刻意不含来源证据；撤销的目标必须由 `recoverableActionRange` 解析（与执行同一个函数）。
入口位置纪律不变：可点的那一下只能在输入区上方的 `OfficeRecoveryPanel`，消息里只留只读记录 ——
恢复卡接管同一条执行时，消息里那条同名按钮要让位。留痕只在 `onRestored` 之后写，`undone` 与
`abandoned` 不互冒充。

WPS 侧未改：撤销记录在宿主 `PluginStorage` + 本机恢复包（TTL），判据在 `wpsPlanExecutor` /
`wpsExecutionReconcile` 一线，未接 `rebuildExecutionUndo`；不要为了"看起来一致"去重写已验证链路。

## 2026-10-31：公式恢复登记与探针取消回归

运行 `node`/vitest 对应的 `excelFormulaSourceEvidence.test.ts` 与 `excelWriteEndToEnd.test.tsx`：覆盖模型有/无工具读取的公式三动作中断、本机读取失败、越权、异步范围变化、混合计划保持原证据、写前取消显示“未执行”。运行 `python scripts/mutate_formula_probe_registration.py` 逐个验证 F1（移除本机来源读取）与 F2（丢失未执行分类）；脚本自动还原原始字节并复跑，不能与其他测试或源文件编辑同时运行。日志位于 `outputs/formula-probe-fix-20261030/`。

## 2026-10-27：未完成执行的发现入口 + 测试网络闸门

**发现（产品侧）**：浏览器 `pending-wps-execution.v1` 从"唯一入口"降级为**缓存**，权威发现改为
`GET /api/{wps|office}/executions/pending?workbookIdentity=…`（按当前工作簿稳定身份精确查找）。
面板在三种时机重跑：挂载、工作簿身份变化（间隔 + 焦点 + 可见性）、授权范围刷新；迟到响应由
代数闸作废。多候选只给选择入口，服务不可达显示"暂时读不到"。发现全程只读。

**怎么加新用例**：如果你写的测试会挂载面板/执行器，而 `api.ts` 里还有函数没替身，现在会**直接
失败**——两端 `setupFiles` 装了 `packages/core/src/testNetworkGate.ts`：

- 未替身的 `fetch` / `XMLHttpRequest` / `WebSocket` / `sendBeacon` 立刻抛错，`afterEach` 也会因为
  留下未处理请求让用例失败（吞掉错误同样算失败）；
- 发现入口的替身：`vi.mock("…/api")` 里加 `...pendingDiscoveryApiStubs()`（来自
  `apps/wps-addin/src/wpsPendingDiscovery.testHost.ts`），需要"真的发现出东西"时用
  `pendingDiscoveryBackend().publish({...})`，不要另写假响应；
- 故意要跑真实实现、由用例自己 `vi.stubGlobal("fetch")` 驱动的接缝（桥轮询、暂停、计划校验）用
  `guardNetworkSeam(real)` 包一层：只有本用例装了 fetch 替身才真的调用；
- **不要**为了让测试变绿去关掉用户那台 8765，也不要加"允许某个地址"的白名单。

**反向变异**：`node scripts/mutate_wps_recovery_discovery.mjs D1 check`（D1–D6，每条只打一个
判别点），四判据是"退出码非 0 / 预期用例红了 / 失败原文带它自己那处证据 / 还原后逐字节一致"。
脚本自己打印并比对 `WpsAgentPanel.tsx` 的 sha256。**变异体不许留在工作区**。

**用户处置留痕（放弃 / 撤销）**：改账本这两列或那个端点之后，跑
`python scripts/mutate_execution_disposition.py D1 check`（D2/D3 同）——
D1 拆"首次写赢"、D2 让处置顺手 `revision + 1`、D3 把两个取值映射到同一列。
**改了端上接线（不是账本）就换另一个脚本**：`node scripts/mutate_excel_disposition_wiring.mjs E1 check`
—— E1 把 Excel `onRestored` 的判定写死成 `abandoned`、E2 把 `onRestored` 挪到
`await undoExecution(...)` 之前（恢复失败也留痕）、W3 把 WPS 的 `origin` 判定写死成 `abandoned`；
`check` 模式额外有 ④b「还原后复跑那条用例必须绿」。
服务端用例在 `server/tests/test_execution_disposition.py`（含旧库迁移与"两端同取值"两条），
两端接线分别在 `apps/excel-addin/src/excelWriteEndToEnd.test.tsx` 与
`apps/wps-addin/src/WpsAgentPanel.test.tsx`（各有"上报失败不重试、不二次撤销"与
"恢复失败不留痕"两条）。**加新端点后记得给所有 `vi.mock("…/api")` 的替身补一条**：
那些替身用 `importOriginal` 展开真实模块，漏一条就会让"网络闸门"把用例判红（这不是产品缺陷）。
留痕的判据位置只有一处：`onRestored` / `undoExecution` 的**成功分支**（恢复失败一律不写）。
Excel「确认后撤销」的完整调用链见
[处置留痕文档](EXECUTION_DISPOSITION_2026-10-29.md) §2.4.1。

**删除前的恢复材料留痕（U5）**：动 `DELETE /api/execution-recovery/{id}` 的顺序、或动
`recovery_package_*` 四列之后，跑
`python scripts/mutate_recovery_package_retirement.py R1 check`（R2–R5 同）——
R1 拆"删前取指纹"、R2 拆"落账失败就不删"、R3 把四列从 `_REQUIRED_OFFICE_COLUMNS` 里拿掉、
**R4 让删除失败也写"删除确认"**、**R5 让删除确认不首次写赢**。
用例在 `server/tests/test_recovery_package_retirement.py`（12 条）。**顺序不许反过来**：
先落账、再删、删成功才记确认；这与处置留痕（U4，恢复成功之后才写）正好相反。
**两个事实不许混**：`retired_at` 只证明"记下了删前摘要"，`removed_at` 才证明"文件确已删除"，
读账本时不能只凭前者宣称删除成功。边界见 [材料回收留痕](RECOVERY_PACKAGE_RETIREMENT_2026-10-30.md)。

**最短原生验收步骤**（本轮未执行，需要真实宿主）：加载新版 → 打开那个仍是 111/222/-303 的文件 →
面板应直接出现"待核对在途动作"（不点任何东西、不重发需求）→ 点"核对在途动作"（只补记，不重写）
→ 出现"继续剩余动作" → 点击后只写第三步 → 审阅里确认/放弃。不要求重发任务，不要重建账本或 WAL。

专项：`packages/core/src/pendingExecutionDiscovery.test.ts`、
`server/tests/test_pending_execution_discovery.py`、`WpsAgentPanel.test.tsx` 的 CAP-03 新增十条。
细节见 [架构说明](ARCHITECTURE.md) 与 [恢复包第五轮](SDK_WPS_RECOVERY_PACKAGE_2026-10-27.md) §28。

## 2026-10-28：发现/选定不该被"缓存旧键"改判（WPS 恢复阻断）

真机阻断：用户选了执行 ID 尾号 `…0f5bffdc`（在途、待核对），界面却显示另一条 `failed` 执行的
「需重新规划：还有 1 个动作没写」。机制在 `discoverWpsRecovery` 两处：①缓存优先里 `identity === null`
也采纳索引（同名异路径旧键串台）；②`replan` 被当成「能做事」，挡住后端发现。修法（`WpsAgentPanel.tsx`）：

- 身份 `null` 时**不采纳**浏览器索引（不猜是哪个簿、不据此清卡）；缓存路径要求
  `identity !== null && pending.workbookIdentity === identity`；
- 「能做事」收窄为**只有** `reconcile` / `resume`；`replan` / `completed` 都不算，继续走后端发现；
- 恢复卡新增 `source: "selected" | "auto"`：用户点选过的卡在自动发现里不改判、不退回候选列表；
- `selectWpsRecoveryCandidate` 作废判据由「发现代数」改为「工作簿身份是否变化」—— 点选不该被并发
  的另一次发现判作废；
- 材料缺失措辞：`ensureWpsRecoveryMaterial` 返回 `none`（宿主缓存与持久包**都**没有）时说
  「本机没有这次执行的恢复材料」，不再落到「找不到原执行的本机恢复记录」（那会把"缓存没了但持久
  包在"误读成"材料丢了"）。

回归：`WpsAgentPanel.test.tsx` 新增 4 条（旧索引挡路 / 身份不可读不猜 / 点选不改判 / 材料缺失措辞），
并修正既有「missing」分支断言。反向验证：临时把后端发现调用换成 `no_identity`，旧索引回归立即转红。
**注意**：持久材料与后端发现经隔离实例证明本就正常，本轮只改前端，未改服务端。

## 2026-10-28：授权范围变化后，同一条已选执行必须重新核验（WPS 恢复）

**上一节那条保护写过头了**。`source: "selected"` 的本意是「别让另一次发现把用户点选的这条改判成
别的执行」，实现却写成「这条以后一律不重算」：用户在恢复卡上看到「需重新规划：工作簿、宿主版本或
授权已变化」，回到数据范围面板取消误选的那张表（三张 → 原两张）后，`useEffect` 里的
`authorizationKey` / `scopeMode` 依赖**确实**变了、发现也确实重跑了，但 `discoverWpsRecovery` 在
`current?.source === "selected"` 处直接 `return`，于是一条已经可以续跑的执行继续挂着旧阻断，
用户只能靠重新规划（等于丢掉这次执行）脱身。

**修法（`apps/wps-addin/src/WpsAgentPanel.tsx`）**：抽 `reverifySelectedWpsRecovery(isCurrent)`，
把 selected 分支从「什么都不做」改成「按当前现场重算这一条」：

- 只重算**这一条**：`operationId` / `taskId` 原样沿用，`getWpsExecutionStatus(operationId)` 直接问
  后端，不再走候选列表（`loadPendingWpsExecution` 那一步跳过），所以不会把用户的点选换成别的执行；
- 判据不复制：仍然调 `wpsRecoveryVerdictFor(latest)`（≈`wpsIdentityVerdict` +
  `inspectWpsRecoveryState` + `recoveryMaterialReason` + `wpsResumeUnavailableReason`），
  「能不能续跑」只有这一份判据，不在面板里另写一套宽松规则；
- 只读：这一步只发 `GET /api/wps/executions/{id}`，不认领、不补记、不执行、不写宿主存储；
- 安全边界不变：①`isCurrent()` 代数闸（`discoveryGenerationRef`）在**两次** `await` 之后各查一次，
  迟到响应一律丢弃；②`recoveryActionLock.current` 为真时既不进也不写（核对/续跑动作在飞时让给它）；
  ③`setWpsRecovery` 里再比一次 `operationId` + `taskId`，两次 await 之间被换成别的执行就不覆盖；
  ④身份读不出来（`identity === null`）时保持原状、不猜、不清卡，仍走上一节的缓存规则。

现在的 `discoverWpsRecovery` 分三种情况，别再把它们合成一个 `return`：

| 现场 | 行为 |
| --- | --- |
| selected 且身份一致（或原执行没登记身份） | `reverifySelectedWpsRecovery`：同一条重新算 |
| selected 但身份已变（换了工作簿） | 保持原卡不动，交给换簿逻辑，不在这里重算 |
| 没有 selected（`auto` / 空） | 走后端发现 + 多候选列表（上一节流程） |

**第二处（`apps/wps-addin/src/wpsPlanExecutor.ts` 的 `assertWpsResumeSources`）**：判据本身一直是
fail-closed 的，但三种不匹配共用一句「工作簿、宿主版本或授权已变化，不能续跑」，用户没法知道该去
改哪一样，只能笼统怀疑「材料是不是丢了」。现在按检查顺序各自点名，**顺序与严格性一字不改**：

1. 身份（`wps-path:<fullName 小写>`）→「当前工作簿与原执行登记的身份不一致；请打开原工作簿后再续跑，不猜是不是同一个文件。」
2. 授权集合（排序后逐字比 `authorizedSheetNames`）→「当前授权 3 张表，原执行登记 2 张表，两者不是同一组；授权范围与原执行不同，不能沿用原执行的写入边界续跑。」
3. 宿主版本 →「宿主版本已变化（登记 12.0，当前 12.1），不能续跑。」

三条文案都**只报数量与版本号，不打印路径/表名**（面板可能被投影）。

**回归（先红后绿）**：

- `WpsAgentPanel.test.tsx` 新增 5 条：把 `assertWpsResumeSources` 的合并文案临时还原、或把
  `selected` 分支临时改回 `return`，这 5 条与 `wpsExecutionResume.test.ts` 的 3 条按预期变红，
  还原后全绿；`selectInFlightUnderScope` 用三表夹具 + 一条写前失败候选（`wps-resume-0002`）证明
  「刷新不采纳旧索引/别的候选」；
- 反向变异固化成脚本：`node scripts/mutate_wps_recovery_reverify.mjs R1 check`（selected 分支
  改回 `return`）与 `R2 check`（迟到写回的三道闸一次拆全），四判据依次是退出码、预期用例名、
  **失败原文带它自己那处证据**、还原后逐字节 sha256（基线 `97c0ba01…2dc0cd`）。
  另：`scripts/mutate_wps_recovery_discovery.mjs` 的 `D2` SEED 因上一节删掉"身份为 null 也采纳
  索引"而失效，已重新定点，D1–D6 六条重跑全成立；
- 覆盖 a–e：①三张表阻断 → 收回两张后同一条改判（`当前授权 3 张表，原执行登记 2 张表`）→
  再收窄到 `[2]` 时「继续剩余动作」立刻消失（旧结论不许残留）；②刷新只问 `resumeOperationId`；
  ③连续切授权时手动控制两次响应到达顺序，迟到的旧响应不覆盖新现场，换簿后恢复卡不画回来；
  ④刷新全程零写入（claim/reconcile/start/complete/finish/suspend 均未调用、`f.writes` 为空、
  `PluginStorage` 序列化前后一致、`jobs.snapshot` 不变）。

**Excel 侧差距（如实记录）**：服务端 `GET /api/office/executions/pending` 与
`apps/excel-addin/src/api.ts` 的 `listPendingOfficeExecutions` 都在，但 Excel 窗格**没有接线**——
没有恢复卡、没有 `source: "selected" | "auto"`、也没有 `assertWpsResumeSources` 的对应实现，
所以这个缺陷在 Excel 侧**尚不存在**（而不是「已修」）。接恢复卡时要连同本节三条一起搬过去。

**本轮未验证**：真实 WPS 窗格里点过没有、两张表下恢复卡会不会真的亮出「继续剩余动作」——都只有
自动化证据（真实组件 + 真实服务端 + 替身宿主）。真机最小步骤见下。

## 2026-10-27：多动作字面常量入口

新增 `plan_literal_edits` 工具，edits 每项为 sheet/range/value（标量），按序编译为
独立 writeValues；`agentPlanTools.maxLiteralActions=50`，全部动作合计仍受 maxLiteralCells=1000
限制。字面来源标记共用原单区域协议。WPS 恢复来源证据 version=2 只允许空来源及纯
writeValues，version=1 仍必须有来源。后端与前端必须同时更新后重新发起任务。

专项：服务端 test_sdk_agent_route.py / test_execution_recovery_package.py；共享层
literalPlanOrigin.test.ts / executionRecoveryPackage.test.ts；两端真实组件的三步常量确认/放弃；
WPS wpsExecutionResume.test.ts 的 lost host cache 常量链。实际命令、失败记录及最短原生操作
见 [本轮交付](SDK_ORDERED_LITERAL_WRITES_2026-10-27.md)。测试通过不代替真实 WPS 完整退出。

## 2026-10-27：完整退出宿主后的本机恢复包（WPS）

真实缺陷：第二步写入回读后、完成回执前中断，完整退出 WPS 重开，账本能核对第二步却提示
「找不到原执行的本机恢复记录」。根因是 `prepareWpsResume` 只读宿主 `PluginStorage`；恢复材料
因此改存本机后端应用数据目录，`PluginStorage`/`localStorage` 只作缓存与发现索引。

改动的文件（照这个顺序读）：

1. `packages/core/src/executionRecoveryPackage.ts` —— 协议与判定（白名单、版本、身份、指纹、形状）；
2. `server/app/execution_recovery.py` —— 原子文件存储（缺失/过期/损坏/冲突/配额各有状态）；
3. `server/app/main.py` —— `PUT/GET/DELETE /api/execution-recovery/{operation_id}`（先核账本身份再动手）；
4. `apps/wps-addin/src/wpsRecoveryPackage.ts` —— 构建、写前落盘、跨进程水合、终态回收；
5. `apps/wps-addin/src/wpsPlanExecutor.ts` —— 写前落盘硬闸（失败即零宿主写入）。

跑测试：

```bash
# 共享层（含双端共读样例 config/execution_recovery_package_cases.json）
cd packages/core && npx vitest run src/executionRecoveryPackage.test.ts
# 服务端（basetemp 必须在 OS 临时目录，否则 safe-delete 闸门会让整批测试成片失败）
.venv/Scripts/python.exe -m pytest server/tests/test_execution_recovery_package.py -q \
  -p no:cacheprovider --basetemp="C:/Users/<you>/AppData/Local/Temp/eb-pytest-recovery"
# WPS 端：跨进程恢复 + 写前落盘失败 + 版本不兼容
cd apps/wps-addin && npx vitest run src/wpsExecutionResume.test.ts
# 面板：完整退出后仍能续写第三步（确认 / 放弃两条）
npx vitest run src/WpsAgentPanel.test.tsx -t "完整退出 WPS"
```

测试替身怎么放才不算作弊：`apps/wps-addin/src/wpsRecoveryStore.testHost.ts` 是**本机后端**的替身
（WPS 重启它不该消失），只替换 `save/load/delete` 三个网络函数，`persistWpsRecoveryPackage` /
`ensureWpsRecoveryMaterial` 仍是生产实现，材料以**文本**跨边界。**不要**用宿主临时文件
（`Env.GetTempPath` + `FileSystem`）去凑这条链路 —— 那不是本产品的存储方案，写出来的"证据"是假的。

反向变异（两条机制各一个判别点，四判据：退出码 / 红的那几条 / 失败原文带自家证据 / 还原后 sha256）：

```bash
node scripts/mutate_wps_recovery_material.mjs            # 基线哈希
node scripts/mutate_wps_recovery_material.mjs M1         # 退回"只信宿主缓存" ⇒ 跨进程恢复必须红
node scripts/mutate_wps_recovery_material.mjs M1 restore
node scripts/mutate_wps_recovery_material.mjs M2         # 拆掉写前落盘 ⇒ "中断后确实有材料"必须红
```

设计与双端状态见 [本机恢复包](SDK_WPS_RECOVERY_PACKAGE_2026-10-27.md)。**原生未做**，A1-GATE 未通过。

## 2026-10-26：S1 原生验收实例与任务关联

验收脚本改用专用 `scripts.a1e3_native_server:create_app --factory` 包装真实应用，记录
真实故障启用/绑定/消费结果，不替换判据；普通 main:app 不暴露验收审计端点。实例token
仅在本地manifest保存，CLI按run与精确taskId只读关联账本；不再按最新记录猜测。
配置/账本每轮隔离，SDK运行目录在仓库外；前端通过frontend子命令显式接入。stop核验
完整进程身份，夹具/日志不覆盖。服务端pytest在收集前隔离配置与SDK目录，避免默认账本
进入用户应用数据。授权、写前复核、白名单执行和先执行再审阅边界均不变。
启动与证据边界见 [验收包](A1_E3_NATIVE_ACCEPTANCE_2026-10-25.md)；本轮结果见
[修复记录](SDK_A1_NATIVE_TOOLING_REPAIR_2026-10-26.md)。

## 2026-10-24：A1 缺失恢复快照拒绝续写

Excel 已完成前缀却缺失恢复快照时，执行器现在在读写之前拒绝续写，防止生成只覆盖后缀的撤销点。WPS 新增实际本机材料删除/过期后的连续重开测试，均无自动认领或写入。两端存储机制不同，不统一硬套过期时间。细节与未闭环项见 [材料恢复证据](SDK_A1_MATERIAL_RECOVERY_2026-10-24.md)；A1 尚未整体完成。

## 2026-10-24：A1 恢复入口并发与重复重开

本批故障回归复现两端同帧双击竞态：WPS 在 targetGuard.check 的 await 之后才取得恢复锁，Excel 恢复入口没有同步互斥。两端现在均在首次 await 前取得本地恢复锁，所有退出路径 finally 释放。服务器 revision/owner 校验保持不变；本地锁不替代跨窗格服务端互斥。

证据：WPS 同帧两次点击、暂停认领 Promise 时只允许一个 claim，完成后仅一个 finish、零宿主内容写入。Excel 原三动作中断用例改为同帧双击，修复前两个审阅分支 finish=2 失败，修复后仍仅续写 B4、start 序号 [0,1,2]、complete 序号 [0,2]、finish 一次。WPS 认领和检查点请求等待期间来源变化/授权收窄四种组合均拒绝后续写入。两端完成后连续两次挂载均不调用写入或补记接口；Excel 清除索引，WPS 保留可清除的只读历史提示，两者展示不同但都不重放。

这些是生产组件/执行器 + 替身宿主自动化证据，双击不是原生实机证据；服务端两客户端竞争由 test_execution_recovery_a1.py 真实 HTTP + SQLite 覆盖。A1 尚缺真正杀进程恢复、引用完整过期链、跨窗格竞争等，不能宣称 A1 全部完成或已进入 A2。用户上一轮“歐克了”作为范围恢复反馈，不扩张为这些故障场景均实机通过。

## 2026-10-24：WPS 固定选表重载与只读恢复刷新

WPS 已保存 xlsx 按既有完整路径身份保存固定选择的表名/序号，不保存单元格值。启动时只恢复完整匹配的选择；缺表、改名、重排或损坏记录均转为手动空范围并提示重选，不自动授权当前表。显式切回跟随当前会删除该工作簿的固定记录。未保存工作簿不跨重载恢复固定选择。运行中固定范围解析也不再静默过滤缺失表。

恢复发现随授权集合、范围模式和来源改变重新进行 GET 与本机只读核对；旧请求由 effect 清理失效，不自动认领、补记、规划或写入。正在执行恢复动作时不并发发起此检查。

两端状态：WPS 本批已实现、原生实机待验；Excel 本批未改，其恢复审阅已由用户验收，但不能据此宣称 Excel 固定选择跨重载持久化已实现。共享层提供严格元数据恢复判据，宿主存储接线在 WPS。A1 仍未整体完成。

实机步骤：加载新版后在独立测试簿重新勾选原来的固定工作表一次（旧版本没有保存记录）；重新加载窗格，检查仍为原固定范围。出现待恢复任务时，改变已选范围应自动只读刷新判据，恢复原集合后无需再点只读核对。仅明确点续跑/补记才推进账本或剩余动作。不要为验证范围恢复重复执行已完成任务。未使用 computer-use。
## 2026-10-24：WPS 最终回执与补记审阅

修复续跑提交非白名单 resumed 字段而被真实 API 校验拒绝。删除该字段，测试接入真实结果载荷校验。已写完的值写入任务可只补记最终回执并进入原结果审阅，无需重写。见 [修复和实机步骤](SDK_WPS_FINAL_RECEIPT_2026-10-24.md)。

## 2026-10-24：WPS 重新规划后的恢复卡收尾

新执行持久登记接管发现索引后，不再刷新并复活同任务的旧恢复卡；旧账本保持真实状态，不伪造完成。续跑证据比较不依赖 JSON 对象字段顺序；失败卡新增具体原因。字段顺序敏感是已修复缺陷，但用户此次精确续跑被拒的实机根因仍未独立证实。见 [实现与证据边界](SDK_WPS_REPLAN_RECOVERY_2026-10-24.md)。

## 2026-10-24：WPS 隐藏重开窗格的恢复入口

用户在 WPS 第二步中断后重开仍找不到核对按钮。代码确认 ribbon.showTaskPane 复用已有窗格，仅设置 Visible=true；原恢复发现只在 React mount 时执行，执行失败收尾不设置 wpsRecovery。这构成无需页面卸载即可复现的入口缺口，不要求用户把隐藏当刷新。

现在执行异常挂起账本后，立即使用既有身份、当前授权及实际目标检查建立恢复卡。不自动补记、不自动续写；执行权未释放或证据不足时仍显示受限判据和只读重新核对入口。恢复卡放在消息末尾，避免被长历史埋在顶部。真正重新加载仍使用原持久发现索引。

新增面板回归覆盖回执失败、页面不重挂载立即出现恢复卡、位于最新消息之后、没有自动认领/补记/完成。既有面板用例继续覆盖核对按钮与剩余续写。用户实际现场是否还存在身份/授权/存储问题仍需更新后的恢复卡判据确认，不能用本修复推断所有无按钮情况均已解决。

本批未修改 Excel、服务端协议或授权边界。未使用 computer-use。日志：outputs/wps-recovery-entry-*.log。实机待用户刷新新版页面复测。

## 2026-10-24：恢复完成后接回结果审阅

Excel 续跑完成后通过共享 ResultReviewController.adoptCompleted 接回审阅；该入口不接受执行回调。双视图使用整份原始恢复点，确认不重放，放弃恢复全部动作。续跑测试证明第二步回执前中断后仅写第三步；用户该次实机调用次数尚无独立日志证据。双端回归通过，原生交互待复测。详见 [修复与证据](SDK_RECOVERY_REVIEW_2026-10-24.md)。

## 2026-10-24：精确中断验收入口

双端增加开发专用、显式单次启用的第二步写入回读后/完成回执前中断入口。正常账本、来源核对、恢复与唯一执行权路径不绕过。DEV 默认可见，dist 验收构建设置 VITE_ENABLE_RECOVERY_PROBE=true；普通生产构建默认关闭。新增 Excel 面板重挂载与 WPS 执行器恢复测试；修复 Excel 复制来源指纹误包含目标导致自身写入阻止续跑。实机待验，操作与证据边界见 [精确中断验收](SDK_PRECISE_INTERRUPT_2026-10-24.md)。

## 2026-10-24：Excel 值复制接入自动执行结果审阅

缺陷：新交付的 copyRange 被 resultReviewPlanIssue 复用的旧副本试算白名单排除，enqueueProducedPlan 静默返回，界面回退为“执行到 Excel”。

修复：本次支持来源与目标表分离、无筛选、不跳空白的 copyRange(values)，含明确矩形/单格锚点及转置。几何准入仅作范围检查，不执行其投影。实际执行前在本机读取授权来源，通过两端共用 copyRangePlan 转换为白名单值写入，同时生成独立回读条件。保留规划基线检查，读取前后复核指纹与工作簿/会话/授权身份，执行侧继续复核；值不在意图阶段外发。快照、标记、回读与撤销统一使用转换后的动作，确认不重放。不可支持或越权的复制明确显示原因，不退回手动执行按钮。其他非复制动作的既有路径本批未迁移。

Excel 新增真实 App/执行器+替身宿主的三动作自动执行、确认不重放、放弃整份恢复，以及来源与范围守卫测试。WPS 原有 prepareWpsPlan 会将第二次复制视为前序数据修改而拒绝，本批仅为来源目标表完全分离的纯值复制批次开放同源转换。多动作来源含公式时双端拒绝，避免预读改变逐步复制语义。补三动作执行/撤销回归。两端真实窗格本批未执行，用户仍需重新加载新版 Excel 窗格后复验；不使用 computer-use。

## 2026-10-24：A1 WPS 值写入剩余动作续跑

已接入重开窗格后的「继续剩余动作」：使用原 operationId/taskId，取得唯一 owner 后再次核对授权、来源和目标，只写 nextActionIndex 之后的动作。保持先执行、再审阅；确认只清标记，放弃恢复原整份计划的写前状态。

本批限定：已登记 resumeContext、30 分钟内本机撤销记录仍在、来源表与目标表分离、有完整来源快照、目标不重叠的 writeValues。来源元数据不含单元格矩阵；复用原 PluginStorage 撤销记录，不新增原始数据存储或出站。格式、公式、结构动作、缺失来源快照的 SDK Python 结果计划仍不能跨窗格精确续写，不标为实机待验。真实 WPS 恢复仍待验；Excel 恢复路径保持，双端 A1 尚未整体完成，不进入 A2。详见 [本批实现与验收](SDK_WPS_RESUME_2026-10-24.md)。

## 2026-10-24：A1 第一批执行恢复取证

新增双端真实 HTTP/SQLite 故障注入 10 条，覆盖在途拒绝、三动作回执缺失后按序号续接、unknown 保持挂起、双客户端并发认领及旧 owner 拒绝；四项内存变异检出。WPS 删除 revision 校验的变异仍由 owner lease 拦截，不宣称该层独立验证。详见 [A1 证据矩阵](SDK_A1_EXECUTION_EVIDENCE_2026-10-24.md)。

第一批取证时的差距（后续值写入子集进展见本文顶部）：WPS 重开窗格后精确续跑剩余动作当时**未实现**，目前只可核对/补记最后已完成动作。服务端可推进动作序号不等于端上能续写。A1 未整体完成，双端原生恢复实机仍未验；不进入 A2。

## 2026-10-24：先执行、再审阅（当前唯一交互约定）

在用户明确授权的任务与所选工作表范围内，白名单 AnalysisPlan 通过计划校验、来源复核及恢复准备后执行一次，再回读并展示结果供审阅。确认变更只接受结果并清除审阅标记，不再次执行；放弃变更使用已验证恢复点。不得增加执行前确认按钮。授权缺失、目标变化或现场不明仍须先解决；仅要求查询或预览的任务不得擅自写入。

本文历史记录中的执行前确认约定已废止；目标授权和缺失信息澄清不属于执行前确认按钮。

## 2026-10-24：进度浮标改为透明底（用户裁决）

用户原话：「把这个锚的图标改一下，背景改为透明的」。截图里 42px 的浅色圆底正好压在消息流某行文字上（"…并移除了审阅标记"），把那一行切掉一段。

**裁决 = 这个控件不再有自己的底。** `.new-progress-button` 常态 `background: transparent`、`box-shadow: none`；`:hover` 只保留 `transform`（以前是渐变整体深一档）、`:active` 只做 `scale(0.94)`（以前是渐变 + `--shadow-sm`）。命中区域仍是 42px。投影一并去掉：`--shadow-*` 是按"有一个圆形底"算的，底透明之后只剩一团没有主体的灰影，读成脏点。

**两处随之改口径，不是纯视觉换色：**

1. 对比度基准从"按钮渐变最暗端点"改成 **`.message-stream` 的 `--color-background-subtle`** —— 那才是字形背后真实的颜色。`styles.test.ts` 不写死数值，从 CSS 读令牌名、从 `useTheme.ts` 读色值、按四个预设重算，门槛仍是 WCAG 1.4.11 的 3:1。
2. 守卫从"渐变必须是令牌"改成"**底色必须真的是 transparent** + 不许再出现渐变/图片底 + `box-shadow` 必须是 none + 悬停/按下规则里连 `background` 都不许出现"。四条分开写，各自被一个变异打红（M2/M3/M8），不靠一条粗断言兼任。

`border: none` **不能省**：作者侧不写边框属性时 UA 默认的 `2px outset ButtonBorder` 会接手，透明底上更明显（凭空一圈黑环）。判据是"值会不会真的画出一条线"（`none`/`hidden`/`0` 放行，`solid` + 非零宽度拦下），变异 M4。

反向变异脚本 `scripts/mutate_progress_dot_contrast.mjs` 已**重新定点**：旧 M2（渐变退回深色药丸）/ 旧 M3（把 `brandHover` 改深以验证"最暗端点"）随渐变删除而失去被拆对象（拿旧 SEED 跑会 SKIP，那种"没跑起来"不能当成证据），改为 M2=常态底色改不透明令牌、M3=hover 补底、M8=把投影加回来。M1/M5/M6/M7 与 M4 不变。

**实机待验**：真实 Excel / WPS 窗格里透明底字形在四个主题下的可读性、无底 hover/active 的反馈是否足够、`focus-visible` 环是否仍清楚。自动化只到"真实样式表 + 真实守卫 + 真渲染小样（本机无头浏览器不出图，本轮未能重生成）"。
`preset-sheet*.png` 是**改透明前**的旧图，不能当本轮证据。

## 2026-10-24：用户 Excel 实机验收反馈

用户在 Python验收-Excel.xlsx 实机确认：双视图对比、放弃变更并恢复、确认变更、确认后撤销均正常。提供的四张截图均显示“已固定 3 个工作表”，本轮未再出现范围丢失。这是用户实机反馈，限定于本次 Excel 场景；不扩展为 WPS 或所有异常交错均已验收，也不据此宣称此前模式切回 auto 的根因已完整复现。

## 2026-10-24：后台扫描不得改写固定工作表选择

Excel 执行/撤销收尾会调用 scan。范围 Hook 的扫描回调改为读取当前模式，避免被旧任务保留的 auto 闭包把后来固定的多表草稿覆盖成活动表。固定选择在目录刷新中原样保留；缺失表仍由准入/目标守卫拒绝，不自动收窄、不自动扩大。WPS 同簿固定模式本就保留选择，本轮不改宿主适配。

新增回归先证明旧扫描回调会收窄勾选，再验证修复；真实 App 覆盖固定三表执行和确认后仍为三表。用户截图中“模式本身变成 auto”的完整实机交错尚未复现，不能把这一处已证明的勾选覆盖缺陷当成全部现场根因。未执行真实窗格验收。

## 2026-10-24：恢复 Excel 执行后双视图窗口排列

执行前值快照副本必须保留窗口桥约定的 EB_ 加 20 位十六进制身份；中文前缀会被 prepare/close 协议拒绝，导致只打开副本而无法并排或自动关闭。现恢复该身份格式，不放宽服务端识别与关闭边界。排列失败通过审阅状态展示具体原因，保留副本清理句柄以及确认/放弃入口，不重放主表写入。WPS 仍使用既有原生 CompareSideBySideWith 路径。

回归覆盖真实 Excel 面板生成种子、桥协议名称、调用排列以及 prepare/arrange 失败提示；双端构建与回归另行记录于 outputs/dual-layout-*。真实 Excel/WPS 窗口排列仍需实机验证，生成文件不能替代并排成功证据。

## 2026-10-24：双视图使用本次修改区域的执行前值快照

Excel 不再为执行后双视图调用整表 capture。执行前仅按计划的动作矩形读取 values、numberFormat、直接格式与行列尺寸；对比画布中动作间的空隙留白，不读取或导出空隙、其他区域或依赖表。副本只写固定值，不导出公式，不复制图表、名称、合并及条件格式效果。画布仍受既有限额约束；无法读取可靠快照仍明确失败。主表执行、回读、确认和恢复链不变。

WPS 保留宿主原生单表副本适配，但本次修改区域按执行前固定值重建，不重新写回公式；clearRange 的执行前对比同样恢复旧值，不再次清空。副本其他区域仅是宿主背景，不作为执行前证据。列宽和无法采集格式等既有宿主限制仍明确提示；不宣称复制整簿或全部复杂对象。两端共享审阅条说明对比范围、公式不重算及对象限制。

Excel 审阅期间的输入框提示改为“等待审阅结果”，其他不可暂停的处理中状态不再误称暂停失效。双端自动化覆盖值快照、公式不重算、原表不重复写入与恢复；Excel 检查区域外数据不进入副本，WPS 检查清空前值。真实双端窗格仍待验。

## 2026-10-24：可选双视图复制限制不再阻止主表执行

Excel captureExcelBeforeResult 曾复用完整副本的 capture 检查；工作簿名称、复杂对象和已有跨表公式导致可选素材采集失败，阻断自动执行。现用显式 ComparisonUnavailableError 区分复制能力限制；仅这类限制保存为不可用结果，写入仍经过既有执行器的授权、来源、恢复快照和回读检查。未知宿主错误、保护、授权与计划错误继续抛出，不作宽泛忽略。点击双视图时明确说明不可用，不伪造不完整副本、不重放执行；确认与放弃仍由已验证回执与恢复点处理。

WPS captureWpsBeforeResult 已将可选对比限制与动作恢复分离，本轮未改其行为。新增 Excel 真实 App/执行器加替身宿主用例覆盖名称、跨表公式、确认不重放、对比失败后撤销、未知读取错误及格式恢复快照失败零写入。替身 getSpecialCellsOrNullObject(Formulas) 现在从真实单元格公式状态返回区域，不再恒为空。

真实 Excel/WPS 窗格仍待验证。本次不认定截图中具体是哪一种复杂对象，也不修改用户工作簿。

## 2026-10-23：续跑计划接入执行后审阅

用户确认当前交互为计划就绪后执行并展示结果，再选择确认变更、放弃变更或双视图；不新增“确认执行 / 取消计划”门槛。本节覆盖此前本轮误提的执行前确认方案。确认变更只清除审阅标记，不重放写入；放弃变更使用既有恢复点。

Excel 首次 appendMessage 会登记自动执行候选，续跑 applyPlanningContinuationResult 只更新消息，导致计划落到旧副本入口。两条路径现共用 enqueueProducedPlan，普通计划与本地规则计划都登记到当前回复段的执行后审阅链；按消息身份去重，加载历史不触发。保留工作簿身份、授权、来源复核、白名单执行、回读和恢复边界。WPS 保持原有自动执行路径，本次误加的批准状态已撤回。

新增 excelWriteEndToEnd 回归：暂停后继续交付计划，无需再次点击执行即可写入并回读，默认不建副本，结果有放弃与双视图入口，确认变更不重复写入。真实 Office.js/WPS 窗格仍待验，自动化仅为真实组件/执行器加替身宿主。

## 2026-10-23：Excel 续跑结果落在当前回复段

真实服务只读记录 turn-4152c91387e24ef19fcbaef4cd7cf1ce 显示 submit_result_plan 已成功，恢复描述为汇总表 A1:G9 的白名单 writeValues 计划；用户反馈最新 Python 结果卡后没有预览入口。Excel applyPlanningContinuationResult 仍把计划写入旧暂停消息，和新回复段的进度分离。现捕获同 turnId 活动段的 persistenceKey，计划/回答落到该段；同任务旧暂停入口退役，历史进度不改写。无匹配活动段才使用原消息。WPS 既有路径在末尾 append 计划预览，本次不改执行器。

测试 messageFirstRouting.test.tsx 新增继续后预览与最新过程卡同属一条消息、仅一个计划的断言。临时退回旧目标消息时用例失败，还原后通过。只读服务事实不等于已验收真实窗格新版本；重新加载后的 Office.js 实机仍待验。没有绕过预览确认，没有改用户工作表。

## 2026-10-23：进度展示语义修订

两端共用 TaskProgressTimeline：有后续段的历史卡标明“本段已结束 · 已接续”；接续段保持接续标题，不再因 awaiting_input 改名为需求分析。普通 awaiting_input 只表示等待下一步操作；已收到 preflight-plan 完成证据时才提示预览已生成、待确认写入。closedByFinish 中的步骤以中断符号展示，真实 warning 保持警告；不修改历史事件或终态。planning.model 明确标作包含工具调用的整体阶段，耗时不可与工具耗时相加，原始事件顺序保持。

Python 同次分析的 analysis_completed 与 pythonResult 在同一个 python.analysis:<run_id> 完成事件内投递；无代码工件的兼容分支仍发送独立结果事件。结果引用、代码正文不进账本的边界不变。旧的持久化重复事件不按标题合并，避免误合并真实的多次分析。暂停详情的 resolved 表示本段已接续，不宣称整个任务或写入已完成。

双端自动化与服务端回归日志：outputs/progress-presentation-*.log。真实 Excel/WPS 窗格尚待验收；仅调整展示和事件身份，不改变写入授权或自动执行策略。

## 2026-10-23：发送与任务控制分离

发送按钮固定为 ↑，空输入/空表单提交不恢复任务。Excel/WPS 共用 `ComposerPrimaryButton`，固定在输入框**右下角**，形态由共享判据 `composerPrimaryAction` 一次算出（发送 / 发送回答 / 暂停 / 继续 / 核对在途动作 / 继续剩余动作），输入框上方只留**无按钮**的状态行 `TaskExecutionNotice`。■ 暂停，暂停请求处理中显示禁用的 …，暂停后 ▶ 要由用户另一次点击才恢复；核对与「继续剩余动作」是两件事，核对只补记回执、不自动续跑。按钮始终 type=button，形态变化用不同 React key，不复用为提交按钮；忽略双击的第二次点击。输入文字时主按钮**永远是发送**（不吞字），发送与既有语义续跑仍走原意图路径。固定区域（审阅栏 / 恢复卡）由 `ComposerActionDock` 承载，一次只展开一块，其余压成可切的 chip。
**（已被 2026-11-01 B 批取代）** 那颗按钮现在在输入框**右下角**（`ComposerPrimaryButton`），输入区上方只剩无按钮的状态行
（`TaskExecutionNotice`）；`type=button` / 换 `key` / 忽略 `detail>1` / 空提交不恢复任务四条约定原样保留。

本次覆盖旧约定“空输入点发送即可继续”。暂停后的自动核对仍只查询状态；这次调整不新增写入能力、不改变授权和预览确认边界。组件原生 click 用例验证同步渲染时不会连带 form submit 或 resume；双端面板验证空提交无动作、明确 ▶ 才恢复。真实宿主验收尚未执行。

## 2026-10-23：暂停回执归属与历史段计时修订

两端暂停请求增加独立的回执代数：新一轮运行或新一次暂停会使旧回执失效，迟到成功、失败、超时不能覆盖当前输入框提示。Excel 续跑的活动收尾也检查原 controller 的归属；这是防守性保护，不宣称已复现实机交错。

共享 TaskProgressTimeline 的非活动段固定显示结束时的完整计时快照，不再每秒重绘。这修订此前“暂停后历史卡的等待进度应继续增长”的展示约定；底层等待区间仍保留供恢复时累计。冻结显示不代表后台已停止，服务端结论仍由暂停回执和状态查询决定，pauseRequested 只表示请求被接受。

双端回归覆盖输入框/卡片两入口各三轮暂停继续、迟到回执不污染下一轮、8 秒超时及超时后的迟到成功。Excel 新段位置额外比较对应的最新用户消息。服务端经 ASGI HTTP 路由检查重复暂停、清理后 interrupted、active=false 和未知任务 404。

证据：outputs/pause-followup-*.log。真实 Excel/WPS 窗格尚未验收；不把组件替身测试记为实机通过。未扩展写入/回读中途暂停。

## 2026-10-23（第二批）：暂停收尾归属 + 测试竞态取证 + Excel 三轮覆盖

**测试竞态（已定因，不是产品缺陷）**：`messageFirstRouting.test.tsx` 的暂停循环偶发
`getByRole` 找不到按钮。探针在 `checkIntent` **派发那一刻**同步读 DOM：首轮已是
`■/暂停任务/is-stop/disabled:false` + 活卡 1；**点「继续任务」后的第二轮只有 `▶/继续任务`、
活卡 0**，一个宏任务后才变 ■。即"请求已在飞"不等于"界面已提交"。改法：点击前
`await view.findByRole("button", { name })` 且 `disabled === false`（等的是用户真实可操作的条件），
同文件 27 处运行态派生按钮的点击一并改成等待式；**不用** sleep、扩超时、重试或吞异常。
新增确定性回归 `a pause click after continue waits for the button to be painted…`。
取证件：`outputs/pause-fix-2026-10-23/evidence/race-probe.log`。

**Excel 端新增 6 条**：三轮暂停/继续（每轮验：可暂停、本轮 signal 已 abort、收尾时活卡严格 0、
执行耗时与步骤行 1.1s 冻结、继续后活卡恰好 1 且段数 2/3/4、旧段不被改写、同 turnId）、
迟到进度不复活运行态、暂停回执三态（拒绝/被接受/没有在跑 → `unconfirmed`/`acknowledged`/
`local-stopped`）。测试数 1490 → **1496**，全量连跑 3 次同数。

**Excel 收尾按归属**：`resumePlanningTaskFromMessage` 的 `finally` 里 `setStatus("idle")`
与并发闸（`useRef<AbortController|null>`，记持有者）原来不按归属，现与桥/controller 同一判据。
这条**构造不出可达交错**（续跑有并发闸、发送/早期重启有 `busy` 闸），如实标为防守性对齐，
不声称修好可复现缺陷。WPS 侧不同，P5 能打红。

**两个易误判为缺陷的既有口径**：① 暂停后卡头 `暂停 X · 等待进度 Y` 的"等待进度"会继续走
（`finishTimelineAt`：有服务端任务时本端只确认停止接收，后台可能仍在跑），冻结断言只能盯
执行耗时与步骤行；② `is-live` 归零不等于 `finishActivity` 生效（`markActivityInterrupted`
就够）——P9 下新增用例仍绿，真实判别点是早期暂停那两条。

**本批实测**：core 790；Excel 1496 + 3 skipped（连跑 3 次）；WPS 1020 + 3 skipped；
双端 build 成功、`tsc --noEmit` 双端 0 错误；lint 0 error / 37 既有 warning；`git diff --check` 0；
`pytest server/tests -q -x --basetemp=.tmp-pytest-pause-<stamp>` **1727 passed**；
`compileall -q server/app` 0。变异 P1/P3/P4/P7 各打红对应新增用例，P8 红 35 条，
P5 红 6 条，P2/P6/P10 为冗余层（如实记录），`restore` 后 sha256 与基线逐字一致。
**真实 Excel / WPS 窗格仍未验收**，一切结论止于自动化。

## 2026-10-23：继续与输入框暂停回归

字段阶段补充验收：不得先释放模拟请求，再断言暂停成功。Excel 与 WPS 真实面板分别点击卡片/输入框暂停，请求持续挂起时必须退出等待；继续后再次执行同一检查，再释放旧回包，证明旧段与新段均不被推进。WPS 用例使用真实 `runWpsReadonlyTurn`，仅替换模型接口与宿主。服务端使用真实 ASGI `http.disconnect`，覆盖三个早期入口及取消后仍返回成功的模型替身，确认取消已到达且不登记迟到决定。原测试曾只覆盖回包后的拦截，不能证明立即暂停。证据目录为 outputs/acceptance-2026-10-23/field-pause/。此轮真实 WPS 窗格诊断工具启动失败，不能算实机通过；后端重启和窗格重新加载仍是生效前提。

双端共同验收：已消费的查询事件重放不进入新段；继续保留原 taskId，旧段冻结；接回进度行正常完成。让补发请求不响应（且替身忽略 AbortSignal），点击输入框暂停必须退出等待，随后回包不得更新卡片。WPS 结果校验等待也必须满足同一判据；暂停后的自动核对即使收到 completed，也不能再发布预览。

服务端验证 Python 成功/取消/不可用三个分支，仅成功登记引用；下一轮 before_model 保留已有完成记录。暂停端点需证明取消真实规划子任务、重复请求仅一次取消、清理前互斥闸仍在，完成后请求无副作用。代码和日志见 outputs/acceptance-2026-10-23/continuation/。这些为组件/服务端自动化，真实 Excel/WPS、供应商和 Docker 尚未验收。后端需重启，两端需构建并重新加载。

# Excel Bro 开发指南

## 2026-10-22：独立复核缺陷修复验收

本轮补充三类定向回归：真实协调器/执行适配流程下的启动前取消、Popen 失败、CLI 失败、工作程序确认后取消、旧镜像成功回包；跨任务 A→B→继续 A 与 closed 父段的双向定位；WPS 两条生产写回路径的迟到更新/收尾守卫。工作程序测试使用可信 exec 替身，不在宿主执行模型代码。WPS 守卫移除变异必须导致专属用例失败，之后按原字节恢复并重跑。记录位于 outputs/acceptance-2026-10-22/codex-fixes/。

Excel 与 WPS 均需完整测试、构建并刷新；服务端需重启。实时 Python 执行确认还需要按既有镜像构建/验证流程更新不可变镜像摘要；旧镜像可使用合法成功回包确认完成，但没有运行中确认。不要为获得运行提示而修改隔离开关或放宽摘要校验。

当前验收层级：自动化采用真实组件/共享逻辑/适配器与替身宿主、Docker OS 边界。真实 Excel/WPS 窗格、真实 Docker 与模型链路未执行，不代表实机验收。generating/summarizing/generation_failed 的生产发出点、写入/回读中途暂停仍未实现，本轮不扩展该范围。

## 2026-10-22（第二批）：四处已复现缺陷的修复（自动化已过，实机未验收）

独立审查确认了四处缺陷，本批逐条修复并补反向变异。**判定口径**：自动化通过 ≠ 实机
通过；下面每一条都标了它到哪一层证据为止。

1. **早期（理解消息 / 读表结构 / 字段意图）暂停后的继续仍然更新旧卡** —— 现在与规划
   续跑走**同一个**段机制（`openReplySegmentMessage`），旧卡保留暂停那一刻的内容与
   耗时并给出「查看后续进展」。这三阶段**没有**服务端任务身份，所以段上**不写**
   `taskId`：不伪造后端任务，也不宣称 SDK 内部断点续算；归属如实是"本机按原需求重跑"。
   已完成的模型阶段不再重调（`messageUnderstanding` 复用），未完成的按现场重跑，
   工作簿结构一律重读。三条入口（输入「继续」/ 空输入点发送 / 点卡片上的继续）经
   `restartEarlyAnalysis` 落到同一条路径。
2. **连续恢复的段链漏接** —— WPS 用 `messages.find(...)` 从前往后找父段，第二次继续
   仍然改写第一段（A→C，B 孤立）。现在父段由共享的 `planSegmentContinuation`
   （内部 `latestReplySegmentForTask`）判定，`openReplySegment` 只改**直接父段**的
   `nextSegmentId`。Excel 同一处也改成走共享判据。回归里有一条**三停三续**用例：
   4 段、3 个「查看后续进展」、只有最后一段是活动段、历史段逐字不变。
   **顺带补的一处**：核查发现 `segmentAcceptsProgress` 当时只有单元用例在调（生产零
   调用点）。现在接到"活动日志写回消息"这条唯一写路径——Excel
   `activityMessages.ts::persistActivityMessage`、WPS `updateActivity`/`endActivity`
   ——`superseded`/`closed` 的段结构上不可能再被迟到进度重新激活。Excel 新增 3 条用例，
   变异体 S8 短路该谓词即判红；**WPS 那两处没有专属红用例**（测试台没有伪造"给已冻结段
   补一条迟到进度"的入口），只算"已接线、无专属反向证据"。
3. **Python 没执行就报"正在运行"** —— 执行前检查拆成 `prepare_python_execution`
   （不碰执行器）与 `run_prepared_execution`（真起容器）。`running` /
   `executedVersion` 只由 `on_runner_start`（容器即将启动时的同步回调）触发；没开始
   执行的走新增的 `not_started` 态并带具体原因，校验不过走 `validation_failed`，
   启动前取消说"未开始执行"、运行中取消才说"已暂停"，异常路径留终态。
4. **代码增量按文本去重损坏内容** —— `applyPythonCodeEvent` 改为按事件身份判重
   （`sequence` → `attemptId` → `version`），增量**无条件追加**；缺片记 `gap` 并关掉
   "复制全部"、压掉 `complete`；完整快照整体替换。`f(g(1` + `)` + `)` 现在是
   `f(g(1))`。

**测试口径**：`messageFirstRouting.test.tsx`（42，含早期三阶段 × 两入口的段数/位置/
历史内容/模型调用次数/段链落盘）、`wpsPlanningResumeClosure.test.tsx`（36，含三停三续
与段链跳转数）、`wpsPlanningPauseCard.test.tsx`（28，含早期三阶段 × 两入口）、
`replySegments.test.ts`（18）、`pythonCodeArtifact.test.ts`（31）、
`apps/excel-addin/src/activityMessages.test.ts`（12，新增 3 条活动段守卫）、
`server/tests/test_python_code_stage_evidence.py`（8，新增：正常启动/功能关闭/无运行器/
无绑定/入参非法/启动前取消/运行中取消/执行异常）、`test_python_event_delivery.py`、
`test_python_async_offload.py`。

**反向变异**：`node scripts/run_reply_segment_mutations.mjs`（S1/S7 段链父段、
S2 增量判重、S3 Excel 早期恢复就地复用、S4 WPS 早期恢复不接链、S5 运行证据提前置真、
S6 未开始执行也带执行版本、S8 写回时短路"目标段还是不是活动段"）。四判据：基线七个
套件全绿 → 每个变异只把它自己声明的那条用例判红（失败原文里有它自己那处证据）→
还原逐字节 sha256 一致 → 复跑全绿。

**本批实测**（全部为本轮重新跑出的数字，没有复用上一批）：core 765 passed；Excel
test:addin **1482** passed + 3 skipped（比上一轮多 3 条 = 新增的活动段守卫用例）；
build:addin 退出码 0；WPS test:wps 1005 passed + 3 skipped；build:wps 退出码 0；
`tsc --noEmit` 由两端 build 覆盖、0 错误；`npm run lint` 0 error / 37 既有 warning；
`git diff --check` 退出码 0；服务端 `pytest server/tests -q`（独立 `--basetemp`）
**1708 passed**，`compileall -q server/app` 退出码 0。
服务端这一遍之前曾出现"跑到 42% 卡住 + `test_managed_recovery.py` 3 条红"：真实原因是
当时有**两个残留 pytest 进程**在和本轮并行跑同一个 `--basetemp`（`test_managed_recovery.py`
单独跑 5 passed / 3.68s，前缀 1–33 单独跑 720 passed，全部干净），杀掉残留进程后
全量连跑两次都是 1708 passed。
**实机（Excel / WPS 窗格里真的点一次「继续」）仍然一条都没执行**；真实供应商流式增量
也没有验收（当前供应商只返回完整代码）。需要重启后端、重新构建并刷新两端加载项。

## 2026-10-22：对话 / 执行进度 / Python 代码展示改造（自动化已过，实机未验收）

**约定变更（先读这条）**：「继续后更新原卡」「一任务只有一张过程卡」**已作废**。
任务身份仍然只有一个，但展示上继续会**新开一个回复段**，旧段冻结为历史并给
「查看后续进展」入口。判据在 `packages/core/src/replySegments.ts`，两端口径一致。

重启后端、重新构建并刷新两端加载项；只用独立测试工作簿。

1. 发一条会跑多个工具的表格需求，跑到中途按暂停：过程卡状态是"已暂停"，步骤
   编号保留；此时界面上仍只有这一个回复段。
2. 点「继续分析」（或输入「继续」）：**必须在这条用户消息之后立刻出现新的助手
   回复段**，进度落在新段里；上方旧段**逐字保留**它暂停时的样子，并出现
   「查看后续进展」按钮，点它把视口带到新段。
3. 旧段不得残留可点的"暂停"按钮（历史段已终态）；新段步骤编号**不重排**，只说明
   "继续（此前 N 步已完成）；此前那些步骤没有重做"。
4. 连点「继续」两次：只开**一个**新段、只一次状态查询 + 一次补发轮询。
5. 后端停掉后点「继续」：如实说"读不到进度、未做任何继续操作"，不重发需求。
6. Python 分析：代码块状态按真实事件走（代码已生成 → 正在校验 → 正在隔离运行 →
   分析完成）；正文按 `codeId` 从本机回环端点取回，取不回就说"正文没有保留"；
   代码块**没有任何执行按钮**；长代码内部滚动、不拉走阅读位置。
7. 计时：暂停时间与断线等待时间分开记；后台仍在执行的那段不算暂停。

自动化：`replySegments.test.ts`（11）、`pythonCodeArtifact.test.ts`（17）、
`PythonCodeBlock.test.tsx`（8）、`useActivityProgress.test.ts`（+2 代码工件）、
`wpsActivityLog.test.ts`（+2 代码工件）、`contracts.test.ts`（+1 解析帧）、
`messageFirstRouting.test.tsx`（42，含多段行为判据）、
`wpsPlanningResumeClosure.test.tsx`（36）、`test_python_analysis.py`（74）、
`test_python_event_delivery.py`、`test_python_async_offload.py`、
`test_python_code_artifacts.py`（8，取回端点与登记表）。

本批实测：core 744；Excel 1479 + 3 skipped；WPS 1005 + 3 skipped；服务端
1700 passed；两端 `tsc --noEmit` 0 错误；两端 build 通过；eslint 0 error
（37 条既有 warning）。**实机（Excel / WPS 窗格里点「继续」）一条都没执行，
不得当成已验收。**

## 2026-10-21：暂停 / 继续 / 进度展示（自动化已过，实机未验收）

> **约定已被 2026-10-22 修订**：下面第 1、2 条里"只留一条链 / 接回原来那张卡"
> 已作废，改为"新开回复段"。其余（连点合并、断线不重发、输入框、计时）不变。

重启后端、重新构建并刷新两端加载项；只用独立测试工作簿。

1. 发一条会跑多个工具的需求，跑到中途按暂停：过程卡状态是
   "已暂停 + 详细信息可折叠"，不重置编号。
2. 点「继续分析」：原任务在服务端仍跑时，界面应先把**已落盘**的进度接到**新开的
   回复段**（旧段冻结）；**不得**出现第二次规划。
3. 在输入框里打「继续」「接着刚才暂停的任务往下做」：走意图理解后落到**同一条**
   恢复入口；已经完成的步骤不得被重新激活。
4. 连点「继续」两次：只产生一次恢复（一次状态查询 + 一次补发轮询），不是两条。
5. 把后端停掉后点「继续」：如实说"读不到进度、未做任何继续操作"，
   **不得**退回重新发送需求；后端起来后再核对一次。
6. 输入框：按下发送那一刻消息就进对话、输入框立刻清空；恢复失败时不得覆盖
   用户后来输入的草稿（`restorePromptUnlessEdited` / `keepRequestInComposer`）。
7. 计时：暂停时间与断线等待时间分开记；后台仍在执行的那段不算暂停。

自动化：`activityTimeline.test.ts`（16）、`useActivityProgress.test.tsx`（2）、
`wpsPlanningResumeClosure.test.tsx`（28，含 running 接回与连点）、
`messageFirstRouting.test.tsx`（36，含 Excel 侧 running 接回）、
`test_agent_tool_bridge.py`（桥步骤身份与终态）、`test_main.py`（补发端点与
中断步骤收尾）、`test_folder_workbooks.py`（每个 started 步骤都有终态）。

反向变异：`node scripts/run_activity_resume_mutations.mjs`（T1–T6，四条判据：
基线绿 / 预期**那条**用例红且失败原文带它自己那处证据 / 还原逐字节 sha256 /
复跑全绿）；服务端侧为 `server/tests/_mutate_step_terminals.py`（M1–M6）。

本批最终自动化：core 16（时间线）+ core 全量见下；Excel 1459 + 3 skipped；
WPS 994 + 3 skipped；服务端 1661 + 1 skipped；两端 `tsc --noEmit` 与 build 均 0；
eslint 0 error（37 条既有 exhaustive-deps warning）。
**实机（Excel / WPS 窗格里点「继续」）一条都没执行，不得当成已验收。**

## 2026-10-18：边框专项验收（实机未执行）

重启后端、重新构建并刷新两端；仅用独立测试工作簿。WPS 新增边框执行，Excel 复核标记不遮挡；保留原有用户条件格式。

1. 在 A1:B2 设置不同主题色/自动色/RGB 边框及混合线型，执行外框和内部横纵线，核对外围与内线均正确，值/公式不变。确认只去标记，撤销恢复各边原样。
2. 只设边框时审阅底色为浅黄，业务边框仍可见；同时改填充与边框时区域直接显示真实格式，不叠审阅底色/边框。部分重叠及反序计划都核对。
3. 执行后手改一条被改边界，再撤销必须拒绝覆盖；手改无关字体/边界后撤销则保留这些改动。
4. 可选双视图显示执行前边框；放弃恢复主表。确认/放弃不得写第二次计划或关闭原表。
5. ThemeColor/TintAndShade/ColorIndex 等读不到时必须写前整体拒绝，不得声称已支持。准备后改目标、来源未授权、展开超限也必须拒绝。单行区域仅要求内部横线属于空动作，明确拒绝。
6. 相邻动作分别指定 A1.bottom 与 A2.top（同一物理线）当前拒绝，不能宣称完整支持所有边框组合。重复通过同一侧指定的动作可执行和撤销。

自动化：`wpsBorders.test.ts`、`borderEdges.test.ts`、双端 `resultReviewMarkers.test.ts`、服务端恢复协议测试。反向验证脚本新增 border-restore / border-alias / border-mark-visibility，逐项 RED、SHA256 还原、GREEN 后才记录通过。实机通过与自动化通过分开记录。


边框批次最终自动化：core 597、Excel 1318 + 3 skipped、WPS 947 + 3 skipped、服务端 1511 + 1 skipped；两端 build 与 compileall 均通过。日志与反向验证分项见能力差距表顶部；本批实机未执行。

## 2026-10-18：WPS 新范围能力验收

重启后端、重新构建并刷新 WPS 加载项后，在独立测试工作簿中勾选两张测试表，第三张保持不选。不得使用业务工作簿测试删除或恢复。

1. 混合填充区域设置非对称颜色（如 #123456），确认实际颜色正确；确认只清审阅标记，撤销恢复各格原填充/图案。
2. 百分比、日期、常规混合区域统一数字格式后撤销，逐格核对恢复；执行后手改一个目标格式，撤销应拒绝覆盖该改动。
3. 写入两列表头与两行数据：只表头加粗/浅绿，正文格式不变；撤销恢复原公式与格式。缺失填充读写接口的宿主必须在写任何数据前拒绝。
4. 从已选来源复制值到另一已选表，测试转置、跳过空白、link；跳过的位置原公式必须保留，链接须指向原表正确单元格。修改来源后重用已保存工具，必须重新读取新值，不复用旧矩阵。
5. 复制来源未授权、矩阵不完整、越界、all/formulas/formats 变体必须拒绝。准备后改变来源，原表零写入。复制不可绕过当前工作簿身份。
6. 执行后可选查看之前状态；确认不执行第二次；放弃与撤销均恢复；来源值不出现在服务端恢复描述中。

7. 垂直对齐：A1:A3 分别顶端/居中/底端，统一后撤销应还原混合值；手改其中一格后撤销须拒绝覆盖。读取不到旧垂直属性时零写入。
8. 条件复制：有表头、命中/未命中行混合，分别测试 values/link + transpose + skipBlanks；链接引用必须为原始行号。条件清空只清命中数据行，表头及未命中行保持。字段缺失或重名应澄清，不得默认为全表或零行。
9. computeValues：在已授权来源表按字段过滤、分组并合计，写入另一已授权现有表；表头样式与 Excel 一致，确认不重复写，撤销恢复目标公式和格式。重用工具前改来源应得到新结果。来源含日期/前导零编号时核对归一结果。未授权来源或目标、来源矩阵不完整、来源在登记等待期间改变必须零写入。

10. 组合动作等待期间手改已执行区域，后续写入和撤销应拒绝覆盖人工改动；宿主写操作中途失败，已改格必须保留恢复记录。回读目标格式缺失时不得显示验证通过，也不得用缺失字段覆盖恢复证据。自动化的注入故障只证明这些分支，不代替宿主实际故障验收。

自动化日志为 outputs/wps-capabilities-*；判别力脚本 node scripts/check_wps_capability_guards.mjs 现含 14 项种子，必须独占运行并核对 SHA256 恢复。自动化通过不等于上述实机步骤已完成。删除工作表仍未接通，不应期待本批刷新后可以删表。


本批最终自动化：core 593；Excel 1317 + 3 skipped；WPS 938 + 3 skipped；服务端 1510 + 1 skipped；两端 build 与 compileall exit 0；14 项反向验证均判红并逐字节还原、复跑绿。完整日志映射见 [能力差距表](HOST_ACTION_CAPABILITY_GAP_2026-10-14.md)。尚未完成的 26 类动作继续明确标为未实现，本批实机未执行。

前一批自动化（保留历史证据）：WPS 911 passed / 3 skipped；Excel 1314 passed / 3 skipped；core 587 passed；服务端 1507 passed / 1 skipped；两端 build、compileall 均通过；7 项反向验证通过并还原。具体日志与未完成能力见能力差距表顶部。

## 2026-10-18：连接协议与用途测试

更新后重新构建两端、用 `npm run dev:server` 重启 SDK 后端并刷新加载项。不要并行启动旧后端；本轮没有停止用户服务或修改保存的模型连接。

1. 打开 Excel 或 WPS 的模型连接设置。按供应商文档选择接口协议；自定义 SDK 端点选择“Anthropic 兼容”，地址原样填写。已有 DeepSeek 官方 OpenAI 连接无需重新录入密钥。
2. 测试用途选择“表格 Agent”，成功必须明确显示 SDK 与受控工具已完成往返。仅“独立公式”测试成功不能替代；编辑协议/用途后要重新测试。
3. 普通 OpenAI 网关若没有兼容 SDK 端点，仍可保存为独立公式连接；Agent 菜单应禁用并说明原因。不得改成猜测的 URL 或静默换供应商。公式选择器拒绝仅支持 Anthropic 的连接；全局为 Anthropic 时应单独设置 OpenAI 兼容公式连接。
4. 老连接 ID、密钥和默认配置保留，SDK 别名迁移到同一连接且只显示一条。原选中连接不兼容时，应由用户明确选择兼容连接。
5. 两端宿主实机核对上述表单、窄窗格可读性、菜单禁用和重新测试流程。本轮真实 SDK 随机令牌探测不代表宿主视觉验收，也不保证所有网关都兼容。

自动化：`server/tests/test_model_connection_purposes.py`、两端模型管理用例；反向验证 `node scripts/check_sdk_connection_guards.mjs`。日志在 `outputs/sdk-connection-*`。真实探测记录 `outputs/sdk-connection-real-probe.json`（只发送随机令牌，未读取工作簿，连接存储哈希一致）。


## 2026-10-18：SDK 唯一路由启动与核验

基础依赖现包含 SDK；原 `requirements-claude-sdk.txt` 保留为兼容安装入口。安装后只启动一个后端：

```powershell
.\.venv\Scripts\python.exe -m pip install --only-binary=claude-agent-sdk -r server\requirements.txt
npm run dev:server
```

`dev:server` 等同 `scripts/start_claude_sdk_server.ps1`，不再启动 `--reload`。占用 8765 时先确认当前进程归属再重启；本轮没有停用户服务。SDK 状态目录必须在仓库外，脚本会设置；隔离 Python 开关仍沿现有验收配置，未自动开启。

已知 DeepSeek 官方连接可直接沿用原连接 ID。未知自定义网关不能自动变为 SDK 连接，须选已支持的连接或按现有显式 SDK 配置方式配置；不修改用户凭据或默认模型。独立公式连通性成功不证明 SDK 兼容性；当前目录和用途测试已分开，见本页第二批说明。

本轮验收要点：

1. 独立测试簿授权三张表，需求“保留 sheet1、sheet2，删除其他表”由 SDK 生成结构化列表；保留名称不得合并。WPS 不支持删表时如实拒绝，不为对齐开放白名单。
2. “清空 A1 后把汇总写入 B1”等复合需求进入完整 Agent 流程，不被地址正则提前截断。关联键、重复项口径未指定时主动澄清；回答后保留原任务。未授权目标通过数据范围调整，答案不能代替授权。
3. 查询需触发工具桥；请求不夹带预计算 `dataResults`，模型只收引用与结构。旧普通连接 ID 也同样要求 SSE；断流不得切回非流式等候工具。
4. 结果仍按当前流程在主表执行一次、标记、可选查看执行前副本；确认不再执行，放弃恢复。上述交互本轮未改。

针对性用例：`test_sdk_agent_route.py`、`test_claude_sdk_agent.py`、两端模型选择/流式请求测试。`node scripts/check_sdk_route_guards.mjs` 独占修改四个种子，验证业务断言变红并 SHA256 还原，勿与其他测试或编辑并发。在线运行器只替换 SDK 模型传输，实际 MCP 工具、HTTP/SSE、Docker 与宿主适配器继续执行；模型替身不算真实模型证据。

本轮真实模型合成工件：`outputs/sdk-route-real-model/report.json` / `synthetic-result.json`，使用普通 DeepSeek 连接 ID 经新路由进入 SDK。通道通过与事实核对分别记录在主线状态；没有操作真实业务工作簿。SDK 原生会话恢复、安装包内 SDK CLI 收集、模型目录的 Agent 兼容性细分仍需后续收口，不能从开发态测试宣称已发布可用。

## 2026-10-18：受控 JS 验收

入口：Excel 的「编辑计划」，WPS 的「我的工具 → 编辑计划」中展开「受控 JS · 多表操作」。这是一种生成白名单计划的受限语法，当前不支持原生 VBA/宏、任意 JS、单元格读取或公式生成。

在独立测试工作簿建三张表，勾选其中两张。点击「填入多表示例」生成计划，应只包含两张已勾选表的 `A1:B2` 加粗操作。生成时主表必须无改动；保存后沿现有计划/工具运行流程执行，两张表应加粗，第三张表不变。确认后撤销应恢复两张表。单张目标可用双视图；多张写入目标应明确提示当前对比限制，不能阻止执行和撤销。

将脚本某个 `sheet` 改为未勾选表名后重新生成，整份编译应失败且草稿不变；生成后撤销某张表授权再执行，执行器必须拒绝、零写入。尝试 `Application.Run(...)`、`fetch(...)`、`eval(...)` 或构造器访问，必须拒绝。现有编辑器中的原计划 JSON 功能不受影响。

定点自动化：`npm run test --workspace packages/core -- src/controlledScript.test.ts`；Excel 的 `ControlledScriptEditor.test.tsx` / `controlledScriptExecution.test.ts`；WPS 的 `WpsAgentPanel.test.tsx` / `wpsPlanExecutor.test.ts`。不以这些替身测试声称真实宿主或原生宏已验收。

本轮全量结果：core 587 passed；Excel 1308 passed / 3 skipped；WPS 887 passed / 3 skipped；服务端 1463 passed / 1 skipped，compileall 通过；两端 build（含 tsc）通过。日志位于 `outputs/controlled-script-*.log`。`node scripts/check_controlled_script_guards.mjs` 验证授权门、多表快照限制两项变异，分别记录 baseline / red / restored 与逐字节 SHA256。共享层单独 `npm run build --workspace packages/core` 仍有 15 处类型错误，位于本轮未修改的 `blankPolicy.test.ts` / `columnMapping.test.ts` / `columnMappingSemantics.test.ts`（联合类型未收窄），不能写成所有项目类型检查全绿。

## 2026-10-17：真实 SDK 主线验证

`scripts/diagnose_sdk_http.py --python --model <真实 SDK 连接 ID> --prompt-file <合成验收需求文本> --out <新的证据目录>` 使用独立回环后端、真实模型和 Docker；加 `--causal-change` 只更改合成来源，模型不接收期望结果。该脚本的 `passed/transportPassed` 仅表示通道成功，`businessAcceptance=not_checked` 明确要求另行核对事实，不是实机通过。

本轮工具说明与结果尺寸投影改动需重启验收后端才生效；不要停掉未知运行进程。本次未重启用户的 8765 服务，默认 Python 开关未改。桌面自动化仍报 sandbox 初始化失败且无连接 Excel 会话，不能用脚本替身代替双端实机点击。当前步骤见 [实机清单顶部更新](MANUAL_PYTHON_ACCEPTANCE_2026-10-10.md)。

2026-10-16（夜间）Excel 侧**不再识别文件名里的日期** —— 两个调用点一起下线，识别机制整条删除。
这条口径在两轮里走完，**别把第二轮读成第一轮的重复**：

* 第一轮（本轮稍早）：用户附截图说「文件名中未识别到报表日期，这个文字不应该出现呀」。
  那张表叫 `评分表.xlsx`，本来就不含日期。当时的修法是**只删否定式说明**，
  文件名里有日期时**照旧显示**「报表日期 2026-03-15」。
* 第二轮（现在）：用户回一句「**不需要 识别日期**」。于是
  `extractWorkbookDataPeriod` 与它派生出的 `workbookDataPeriod` **整条删除**，
  不是留着一个没人读的派生值 —— 留着它，任何一处 JSX 都能把它请回来。

**根因是"把内部状态当成了要告诉用户的事实"，但只删那句提示是不够的。**
`workbookDataPeriod` 由 `extractWorkbookDataPeriod(workbook.name)` 从**文件名**正则派生
（`workbookIdentity.ts`，识别 `2026-03-15` / `2026年3月15日` 这类；单个匹配或两个相同 ⇒ 该日期，
两个不同 ⇒ `A 至 B`，一个都没有 ⇒ `null`）。`null` 只说明**我们的正则没匹配上** ——
它既不是用户此刻要做的决定，也不代表文件有问题。但只要这条派生还在，
界面就始终在"猜用户的日期"，用户看到的仍然是"这软件在猜我的日期"。

**删掉的四处（一处不留）**：

1. `App.tsx` 里本地那份 `const workbookDataPeriod = useMemo(…)` 整段删除，
   `extractWorkbookDataPeriod` 的 import 一并删掉。
2. **当前文件卡**（`div.workbook-identity` 的第二行）：三元表达式换成单行
   `<em>当前工作表 {workbook.activeWorksheet}</em>`。
3. **数据范围触发器**（`.scope-trigger` 的 `strong`）：`auto` 支与「已固定」支上的
   `${workbookDataPeriod ? … : ""}` 前缀一并删掉 —— 只会漏在这儿：触发器收起来时
   它显示的是「2026-03-15 · 当前表 Sheet2」，比卡片上那句话更早被看到。
4. `hooks/useWorkbookContext.ts` 里同名的派生状态与导出、
   `workbookIdentity.ts` 的 `extractWorkbookDataPeriod` + `normalizedDate`。

**连带清掉的**：`workbookIdentity.test.ts` 里两条钉它的用例（文件从 8 条降到 6 条）、
`hooks/useWorkbookContext.test.ts` 里 `vi.mock("../workbookIdentity")` + 那条
"should derive workbook data period"（10 条 → 9 条）、`docs/HOOKS_GUIDE.md` 接口清单里的
`workbookDataPeriod`。**这些断言留着就是在说谎**：它们宣称产品会认日期，而产品已经不认了。

**判据（`apps/excel-addin/src/workbookIdentityCard.test.tsx`，2 条，挂真实的 `App.tsx`）：
两个调用点 × 两个文件名状态，四角全在场。**
每条都同时断言 ① 卡片第二行、② `.scope-trigger strong` 的文案，外加
`not.toContain("报表日期")` / `not.toContain("未识别到")` / 那一行 `not.toContain("·")` /
触发器 `not.toMatch(/\d{4}/)`，并在旁边放正向锚点（文件名的 `strong` 仍在、`em` 仍在、
所述工作表名仍在），免得负向断言靠"整行删掉"假绿。

* 文件名 `评分表.xlsx` ⇒ 卡片「当前工作表 Sheet2」、触发器「当前表 Sheet2」；
* 文件名 `评分表-2026-03-15.xlsx` ⇒ **一模一样，不出现日期**。

第二条是关键：一个"有则显示"的版本能通过第一条。只测"没有日期的那支"等于没测。

**反向变异**：`node scripts/run_workbook_identity_card_mutations.mjs`（4 个变异，先跑定点闸）——
H1（把日期识别请回**卡片**）、H2（退回「文件名中未识别到报表日期」否定式）、
H3（整行 `<em>` 删掉）、H4（把日期加回**触发器**）。
实测：基线 GREEN（2/2）→ H1 只红「有日期」那条、H2/H3 两条都红、H4 只红「有日期」那条
（失败清单 1 / 2 / 2 / 1）→ restore 后 `App.tsx` **逐字节一致**（sha256 `9a613481…`）→ 复跑全绿。
两处新写法值得记住：

* **驱动要连"失败原文"一起卡**。H1 与 H4 红的都是同一条用例（"文件名里有日期也不显示"），
  只靠用例名根本分不出是**卡片**被塞了日期还是**触发器**被塞了日期 —— 那就等于两个调用点里
  永远有一个没被测到。所以 `MUTATIONS[].errorContains` 要求失败原文里出现**它自己那一处**的证据：
  H1 是 `报表日期 2026-03-15 · 当前工作表 Sheet2`（卡片那一行），H4 是 `2026-03-15 · 当前表 Sheet2`
  （触发器那一行）。
* **变异不许靠"引用一个已删除的标识符"来变红**。H1/H4 都是把日期识别**就地内联**，
  不依赖 `extractWorkbookDataPeriod`；否则测试非 0 退出是因为标识符不存在，
  那是运行期错误冒充判红，与"判据有判别力"无关。

证据分档（本轮）：**自动化已通过** —— excel 全量 `107 passed | 2 skipped`（文件）、
`1281 passed | 3 skipped（1284）`、退出码 0
（= 上一轮的 1284 减掉随能力一起删除的 3 条：`workbookIdentity.test.ts` 2 + `useWorkbookContext.test.ts` 1）；`npm run build:addin` 退出码 0（内含 `tsc --noEmit`）
并重刷 `dist/`，重刷后 `apps/excel-addin/dist` 里已搜不到「未识别到报表日期」，说明之前那份是
它的**陈旧构建产物**、不是第二处源码；两个既有像素脚本复跑仍退出码 0（动作条 12 条断言 / 选项行 8 条断言）；
三套变异脚本的种子定点闸（`mutate_scope_action_bar_layer` / `mutate_scope_option_rows` /
`mutate_workbook_identity_card`）均 `OK` —— 这一轮又动了 `App.tsx`，按铁律必须回头确认旧种子没被打散。
肉眼对照：`node scripts/render_workbook_identity_card_demo.mjs <playwright-core 路径>` 用**真实样式表** +
逐字抄自 `App.tsx` 的卡片与触发器结构画**四格**（最初·否定式 / 上一版·卡片与触发器都带日期 /
现在·无日期 / 现在·有日期），真浏览器读回两处文案复核 + 截图 →
`outputs/workbook-identity-card-2026-10-16/`。**前两格是反向对照**（第二格专门画出那个
"卡片删了、触发器还留着"的半截状态），证据等级 = 那两条 Excel 夹具泳道，**是夹具页、不是实机 Excel 截图**。
**没有对端判据**：WPS 侧本来就没有日期（`.scope-trigger` 与当前文件卡都只有表名），
所以这一轮只有 Excel 一侧 —— 不把"两端一致"写成新增证据。改动是纯文案语义、**不涉像素**，
故没有实机截图这一档。`docs/DEVELOPMENT.md` 上前两轮记录的 `App.tsx` sha256（`6c762763…` / `e79ca7eb…`）
已相继被取代，现值 **`9a613481…`**。


2026-10-16（傍晚）Excel 与 WPS 的「完成 / 确认」按钮**不在同一层** —— 先只做诊断，再对齐。
用户先要求「不写代码」地找原因，确认后追加一句「把 excel 的按钮和层次修改和 wps 一样」。

**根因不是 padding，是结构。** 两个弹层是**两段各自写的 JSX**（`App.tsx` 与 `WpsProbePanel.tsx`），
所以动作条落在**两个不同的类名、两层不同的容器**里：

* WPS：`section.scope-popover` → `div.scope-popover-scroll`（收口）→ `div.wps-scope-actions`
  = 滚动容器的**兄弟** ⇒ 固定页脚。样式里有 `flex: 0 0 auto` + `border-top` + `padding`。
* Excel（改前）：`section.scope-popover` → `div.scope-popover-scroll` → `fieldset.scope-selection-fields`
  → `div.scope-controls` → `div.scope-action-buttons` = **在滚动容器里面**，样式只有
  `display: flex / gap / justify-content`，**没有** `flex`、`border-top`、`background`
  ⇒ 它只是"内容末尾的一行 flex"，内容一滚就跟着滚走。

`.scope-popover` 是 `column flex + overflow: hidden`，`.scope-popover-scroll` 是
`flex: 1 1 auto; min-height: 0; overflow-y: auto` —— 这两条合起来才是"钉住"的机制，
所以**只有当动作条是滚动容器的兄弟时才被钉住**；放在容器里，`flex: 0 0 auto` 挡不住容器整体滚动。
Excel 其实**已经有**这个正确形态，就在文件夹模式的 `div.folder-selection-footer`
（`flex: 0 0 auto` + `border-top` + `background` + `z-index: 2` + `min-height: 68px`，且在滚动容器之外）——
分叉只出在工作簿模式没复用它。

**改法**（三条一起，缺一条就会重新分叉）：

1. **搬层**：工作簿模式的动作条（与它下面的错误条）从 `scroll → fieldset → scope-controls`
   里移出来，改成 `section.scope-popover` 的直接子节点、`div.scope-popover-scroll` 的兄弟。
   Excel 现有的 `scopeDraftDirty` / `committedSheetNames` / `cancelScopeEdit` 本来就是现成的
   （`hooks/useScopeSelection.ts:196` 的语义与 WPS 的 `discardScopeEdit` 是同一件事 ——
   草稿回到已提交授权，已提交授权一字不改），此前一直没人用。
2. **一份声明**：`apps/excel-addin/src/styles.css` 里那条页脚规则改成
   `.scope-action-buttons, .wps-scope-actions { display: flex; flex: 0 0 auto;
   justify-content: flex-end; gap: 7px; border-top: 1px solid var(--color-border);
   padding: 10px var(--spacing-md); }`，同时**删掉** `apps/wps-addin/src/styles.css:1963`
   那条重复的同名规则。两端从此读同一份；哪一端再把容器搬回去，一端的分叉会立刻显形。
3. **按钮集合**：Excel 工作簿模式与 WPS 一致 = `刷新 / 确认 / 取消`，「取消」只在
   `scopeDraftDirty` 时出现，点击走 `cancelScopeEdit`（弹层不关，这是"我点开只是想看看"的出口）。
   「完成」→「确认」，**文件夹模式那条也一并改名**：两处调的是同一个 `confirmSheetSelection()`，
   同一个动作有两种叫法正是本轮要消灭的那种漂移。

**判据分两层，缺一层都不算过**（jsdom 没有布局引擎，量不出像素）：

① 结构：`apps/excel-addin/src/scopeActionBarLayer.test.tsx`（5 条）—— 挂**真实的 `App.tsx`**，
   断言 `bar.parentElement === section.scope-popover` 且 `bar.previousElementSibling === 滚动容器`，
   **反向**断言 `scroll.contains(bar) === false`；按钮集合恰为 `["↻ 刷新","确认"]`，造出草稿差异后
   变成 `[...,"取消"]` 且「完成」「确认范围」「放弃修改」三个旧文案一个都不许出现；点「取消」后
   勾选态与触发器文案都退回已提交授权、弹层不关；错误条是弹层的直接子节点且紧跟在动作条之后
   （不写 `>` 选择器 —— 写了之后"被塞进页脚里"这个回归只会让判据以**超时**变红，看不出是哪条口径坏了）。
   WPS 侧同一判据落在 `wpsScopeDraft.test.tsx` 新增的第 17 条里（`actions.parentElement === popover`
   + `actions.previousElementSibling === 滚动容器` + 反向 `scroll.contains(actions) === false`）。
   **工作目录模式的调用点另有两条**（`folderScopeSafety.test.tsx`「Excel 工作目录：固定页脚的层与格式（对齐 WPS）」）：
   页脚住在弹层上、确认在页脚里、承认汇总紧跟在页脚下方、工作目录模式**没有** `.scope-action-buttons`；
   以及「勾选只改草稿 → 页脚随即出现取消 → 点它把草稿退回已提交授权、弹层不关」。
   `App.tsx` 里「完成」这个文案还被 **14 处**测试调用点按 `getByRole("button", { name })` 找着
   （`folderScopeSafety.test.tsx` 10 处、`messageFirstRouting.test.tsx` 4 处），全部随之改为「确认」；
   `syntheticResultEndToEnd` / `crossProcessChain` / `continuousPythonChain` / `excelWriteEndToEnd` /
   `onlineLink` 里「范围『完成』按钮尚未出现」的报错文案同步改掉。

② **真实像素**：`node scripts/check_scope_action_bar_layer.mjs <playwright-core 路径>` ——
   三条泳道（Excel 夹具 / WPS 夹具 / Excel 夹具带错误条）+ 一页**真实的 `WpsProbePanel`**
   （只把 `.scope-popover` 夹到 300px 制造滚动）。判据只有两根数：
   **内容滚动时动作条自己的位移 = 0px**（固定页脚的可证伪定义）、
   **面板一打开动作条就贴住面板下沿**（贴底距离 = 面板自己的 1px 下描边，且落在可视区内）。
   实测：Excel 夹具 0px / WPS 夹具 0px / 带错误条 0px / 真实 WPS 组件 0px（溢出 398–512px）。
   另配一根**防假绿锚点**：同一容器里标题的位移必须 > 100px（实测 398–512px），
   否则"动作条位移 0"可能只是因为容器压根没滚动。
   `--legacy` 把两端的结构换回修复前（动作条放进滚动容器）：**12 条断言全部判红**，
   量到的就是修复前的真实样子 —— 动作条在滚动容器里、随内容位移 **−446px / −512px**、
   且打开时它在面板**下方** 433px（视口里根本看不见，正是用户截图里"按钮不在那一层"的样子）。
   这一段另有一处诚实的边界：真实 `WpsProbePanel` **不受 `--legacy` 影响**（它就是仓库里那段代码），
   两种模式下都保持修复后的形状 —— 这恰好反证夹具 legacy 的差别确实只来自结构。

**反向变异**：`node scripts/run_scope_action_bar_layer_mutations.mjs`（**9 个**变异，先跑定点闸）——
E1（Excel 动作条搬回滚动容器里）、E2（WPS 同）、B1（工作簿动作条的「取消」出现条件写反成 `!scopeDraftDirty`）、
B2（确认按钮文案退回「完成」）、**B3（工作目录固定页脚的「取消」出现条件写反成 `!folderDraftDirty`）**、
C1（WPS 侧又自己写一遍页脚）、C2（共享声明被拆成只管 Excel 一条）、
F1（错误条被塞进动作条里面）、**L1（宽屏两栏覆盖插回共享样式表）**。
**B1 与 B3 是两个调用点、两份草稿，互不代替** —— 所以 excel 套件
从「只跑 `scopeActionBarLayer.test.tsx`」扩成「+ `folderScopeSafety.test.tsx`（11）= **17**」，
否则 B3 无论怎么写都不会红。实测：基线两端 GREEN（excel 17/17、wps 17/17）→ 九处**各自把它自己
那一端判红**（E1 2 条、E2 1 条、B1 2 条、B3 2 条、B2 3 条、L1/C1/C2/F1 各 1 条）→ restore 后四个目标文件**逐字节一致**
（`App.tsx` sha256 `6c762763…`（**当轮的值**，此后被文案改动取代为 `e79ca7eb…`、
再被「不再识别日期」取代为 `9a613481…`、补上工作目录页脚的取消后为 `d39ce591…`）、
`WpsProbePanel.tsx` `763ebf69…`（改措辞后为 `c1b022f0…`）、
`excel-addin/src/styles.css` `bdc669ae…`（**删掉宽屏两栏覆盖后为 `f15129e0…`**）、`wps-addin/src/styles.css` `c3465801…`）→ 复跑两端全绿。

**文件夹主从区是「进入式单栏」，两端共用、任何宽度都一样（2026-10-16 深夜，按用户口径）**：
原来有一段 `@media (min-width: 520px)` 把 `.folder-master-detail` 在宽面板切成两栏（文件列表一直挂着、
`.folder-back-button` 被藏），于是 WPS 任务窗格里是"左边勾文件、右边勾子表"，而用户熟悉的 Excel 交互是
「点文件 → 进入该文件 → 勾它的子表（← 文件 返回）」。**整段删掉**（styles.css 里留了说明注释），让
窄屏形态的基础规则在所有宽度生效：files 视图只显示文件列表（`.folder-sheet-detail { display: none }`）、
点文件进 sheets 视图只显示该文件的子表（`.folder-picker-view-sheets .folder-file-list { display: none }`）、
返回按钮常在。判据是 `scopeActionBarLayer.test.tsx` 的 CSS 契约用例（读**剥掉注释后**的生效文本，
防换写法复活），反向变异 L1 把那段覆盖原样插回。**顺带修掉了宽面板下三层滚动条的问题**（原来左栏、
右栏、外层弹层各一条）。
E1/E2 是**两步定点**（先把滚动容器的 `</div>` 删掉、再在错误条之前补回去），因为一次结构搬迁
在 JSX 里必然是两步；每步都要求**命中恰好 1 处**，不存在"删对了插错了还能静默通过"。
`check` 的判定因此写成 `muted` 优先：**插入型变异的 `from` 串天然是 `to` 串的子串**，
先判 `unmuted` 会把"已变异"误报成"未变异"。

证据分档（本轮）：**自动化已通过** —— excel `106 passed | 2 skipped`、`1282 passed | 3 skipped
（1285）`、退出码 0（比上一轮记录少 13 条失败：那 13 条在本轮跑时已经全绿，期间有**并发写入者**
在同一工作区改过文件）；两端 `tsc --noEmit` 退出码 0；`build:addin` / `build:wps` 退出码 0。
**wps `72 failed | 779 passed | 3 skipped（854）`：与本轮改动无关，属并发写入者的在途改动。**
判定依据（两条，缺一不算）：① 本轮对 WPS 源码的**唯一**改动是删掉 `styles.css` 里那条重复的
`.wps-scope-actions`；把这条规则**临时加回去再跑**，失败依旧（`应展示写入计划`），
加回后已按 sha256 `c3465801…` 逐字节还原。② 失败集中在 `WpsAgentPanel.test.tsx`（70）/
`wpsContinuousPythonChain.test.tsx`（1）/ `WpsAgentPanelFormatReview.test.tsx`（1），
全是"计划预览 / 回写 / 原生审阅"这一族，而 `WpsAgentPanel.tsx` 的最后修改时间是 **12:59**
（本轮改动都在 13:23 之后）。**上一轮记录的 3 条 `getServerSnapshot` 失败已不在列表里，换成了这一批。**
**真实 Excel 2024 / 真实 WPS 的像素验收仍未执行** —— 本轮的像素证据来自无头浏览器（本机 Edge）
+ 逐字照抄的夹具，以及一页真实 `WpsProbePanel`；**替身/夹具证据 ≠ 实机证据**。

2026-10-16（下午）数据范围面板「选项行」四条口径 —— **这一轮推翻了同日「勾选不改变行高」那条的做法**。
用户给了四条编号改动（②③④附两端截图）：① 行内不要再出现「可分析」「修改仍需预览确认」这些字；
② WPS 的行内选择控件比 Excel 好，Excel 对齐它；③ 列表行高「太细」（38px 不够），参照「工作目录」
那些选择框统一；④ Excel 里块与块的间隔太窄，参照 WPS 处理。

**① 行内授权文案整条下线（推翻上一轮）。** 上一轮的结论是「把徽标收进表名所在的同一个行盒」，
行高因此不再跳动；但用户这次给的两张截图正是那两个行盒 —— 他不要这句话出现在行里。
于是 `ScopeAdmissionBadge` 组件、两处行内调用点（工作簿列表 / 工作目录列表）、WPS 行内那句
`<small>` 与 import 一起删除，`.scope-admission-badge` 的四条样式随之删除。授权状态**没有消失**：
它本来就是"整次选择"的属性，唯一去处是列表下方的 `ScopeAdmissionSummary`（按 reason 展开明细，
`SCOPE_ADMISSION_MESSAGES[reason]` 只在那里渲染）。上一轮的判据文件与两个脚本
（`wpsSheetPickerRowHeight.test.tsx` / `check_wps_sheet_picker_row_height.mjs` /
`mutate_wps_sheet_row_height.mjs` / `run_wps_sheet_row_height_mutations.mjs`）**已删除** ——
它们的核心断言（"徽标必须在行盒里"）与这一轮正好相反，留着只会两头说谎。
上一轮那条判据也没有丢：`wpsScopeDraft.test.tsx` 里"行内看得见写入状态"那条**搬到汇总行**
（`写入状态待核实` 现在断言在 `.scope-admission-summary` 上，并逐条扫 `SCOPE_ADMISSION_MESSAGES`
确认行内一条都不在）。

**② Excel 行内勾选框对齐 WPS。根因是一条后代选择器泄漏。** 实测差异：WPS 行内勾选框
20×20、1px 描边、**不填色**；Excel 那颗是 **25×25 的灰方块**（实测填充 `rgb(241, 240, 238)`）。
根因不是勾选框自己写错，而是 `.scope-mode-list i` 是**后代**选择器，而整棵工作表选择器都挂在
`.scope-mode-list` 的子树里 —— 模式卡那颗 25px + `--color-surface-hover` 底色的图标规则
把行内勾选框一起吃掉，特异度 0-1-1 压过 `.sheet-checkbox` 的 0-1-0。修法两步：图标规则收窄成
**`.scope-mode-list > button > i`**；勾选框统一成**一份**共享声明 `.sheet-checkbox`
（20px / `1px solid var(--color-border-strong)` / `background: transparent`，勾上由
`button.selected > .sheet-checkbox` 换成品牌绿 + 白勾）。WPS 侧那条高特异度的
`.wps-sheet-picker-options > button i`（0-1-2）同时删除 —— 要"一份"就不能留两条。

**③ 行高统一成一个 token。** 原先 WPS 列表行自定 `38px`、Excel 工作表行没有 `min-height`
（由两行内容撑出 57px），而模式卡（跟随当前 / 选择多个 / 工作目录）是 59px。现在三处模式/列表行
都读 `--scope-option-height: 59px`（定义在 `apps/excel-addin/src/styles.css` 的 `:root`，
两端共用同一份样式表），"统一"由**一处声明**保证，而不是靠两个恰好相等的数字。
59 的来历 = 模式卡自己的盒高：两个 12px/20px 的行盒 + 1px 行距 + `--spacing-sm` × 2 + 上下 1px 边框。

**④ Excel 块间距被内层 fieldset 吃掉。** Excel 的正文多包了一层
`fieldset.scope-selection-fields`（提交期间整体 `disabled`，WPS 侧没有），外层
`.scope-popover-scroll { gap: 10px }` 因而不起作用 —— 实测卡片→模式卡间距 **Excel 0px / WPS 10px**。
修法：间距提成 `.scope-popover` 上的 `--scope-popover-block-gap: 10px`，外层滚动容器与内层
fieldset 都读它，并且 fieldset 改成 `display: grid`、`border: 0; padding: 0`。

**判据分两层，缺一层都不算过**（jsdom 没有布局引擎，量不出像素）。
① 结构与契约：`apps/excel-addin/src/scopeOptionRow.test.tsx`（6 条，按 ①②③④ 分组）钉 App.tsx
**代码**（剥掉注释后）里不再有徽标与那句文案、两处列表与两处行名还在、`<ScopeAdmissionSummary`
还在、模式图标规则已收窄且仍保留 `width: 25px` + `--color-surface-hover`、`.sheet-checkbox`
是 20px/1px/不填色、`--scope-option-height` 是 59px 且三处选项行都读它、
`--scope-popover-block-gap` 是 10px 且外层与内层都读它。
`apps/wps-addin/src/wpsScopeOptionRow.test.tsx`（3 条）挂**真实的 `WpsProbePanel`**，
断言勾选前后行内都没有徽标元素、也没有那句文案的任何一个字，而表名、「✓」勾选反馈、
下方汇总行都还在。
② **真实像素**：`node scripts/check_scope_option_rows.mjs <playwright-core 路径>` 在真实浏览器
（本机 Edge）里挂**真实的 `WpsProbePanel`**（聊天面板换替身，只留范围面板槽位）与一段从
`App.tsx` 誊下来的 Excel 面板夹具，用**真实样式表**量 8 个值：**两端工作表行 59px
（未勾选 → 已勾选不变）、两端块间距 10px、两端勾选框 20×20 不填色**，8/8 通过。

**这条测量有判别力吗 —— `--legacy`：** 同一脚本加 `--legacy` 把四类写法整体换回修复前
（内层 fieldset `display: block; gap: 0`、Excel 工作表行退回 `min-height: auto`、
WPS 工作表行退回 `38px`、`.scope-mode-list i` 退回**后代**选择器、WPS 外层 `gap: 0`），
**8 条断言全部判红**，量到的就是修复前的真实样子：WPS 行 38px、Excel 行 57px、
Excel 勾选框 **25×25 填充 `rgb(241, 240, 238)`**、三处间距 0px。
WPS 那条 `gap: 0` 是**专门补进去的** —— WPS 本来就没坏在这一项上，不加就没法证明那条断言
真挂在间距上。Excel 那条退回写的是 `min-height: auto` 而**不是** 38px：HEAD 里
`.workbook-sheet-list > button` 根本没有 `min-height`，38px 是 **WPS** 的值 —— 反向对照里
编一个从未存在的"修复前"，可信度就是这么丢的（数值碰巧一样也不行）。

三个必须保留的坑：`round is not defined`（页面函数看不见 Node 侧的辅助函数，量里各写一份）；
`.scope-picker-reveal` 展开有 360ms `0fr → 1fr` 过渡，脚本里必须关过渡并过一道前置闸
（列表高度 > 60px），否则量到中间帧（典型的假绿）；对照页截图必须**等布局稳定**再截，
且高度不能读 `documentElement.scrollHeight` —— 页面带着**应用自己的样式表**，
`html, body { height: 100% }` 会把它钉在视口高度，照它截就会把工作表列表那一段整段切掉。

**反向变异**：`node scripts/run_scope_option_row_mutations.mjs` —— M1（WPS 行内插回文案）、
M2（Excel 行内插回文案）、M3（Excel 图标规则退回后代选择器）、M4（WPS 侧另写一套行内 `i`）、
M5（WPS 行高退回 38px）、M6（Excel 行高退回 57px）、M7（内层 fieldset 间距退回 0）。实测：
基线两端 GREEN（excel 6/6、wps 3/3）→ 七处**各自把它自己那一端判红**（每处 1 条失败；
M6 因为两端读同一份共享样式表，WPS 侧跟着红）→ restore 后四个目标文件**逐字节一致**
（`WpsProbePanel.tsx` sha256 `763ebf69…`、`wps-addin/src/styles.css` `6de28296…`、
`excel-addin/src/styles.css` `db6bd1a1…`、`App.tsx` `5abcf609…`）→ 复跑两端全绿。
M1 另单独验过：它同时把 `wpsScopeOptionRow.test.tsx` 与 `wpsScopeDraft.test.tsx` 判红
（`expected '✓明细可分析 · 修改仍需预览确认' to be '✓明细'`），所以"搬到汇总行"那条判据
没有因为搬家而丢掉判别力。

**变异脚本这次真抓到过一次缺陷，护栏因此加了一道。** M5 原先的种子只写
`  min-height: 38px;` / `  min-height: var(--scope-option-height);`，不含选择器。而
`.connection-shell > button`（WPS 连接页那颗按钮）本身就是 `min-height: 38px`，于是**还原**时
`split(mutant).join(seed)` 把**两处一起**替换，把连接页按钮顺手撑成了 59px（`git diff` 才看出来）。
三处修：种子/变异串改成**带选择器上下文**的定点锚（`.wps-sheet-picker-options > button` 里
让 `grid-template-columns` 紧挨 `min-height`，这一对在本文件里唯一）；`restore` 改成**只替第一处**、
且**命中 ≠ 1 就拒绝**（不再是"命中 0 算 ALREADY-CLEAN、其余一律盲替"）；新增 `check` 前置闸
（每个变异同时验 SEED==1 且 MUTANT==0）并挂进驱动脚本开头。这条闸自己也有判别力：拿旧那个
歧义串去判，它报 `MISS 旧歧义串 SEED=1 MUTANT=1`。**改完源码（换行、重排声明）必跑 `check`。**

证据分档：**自动化已通过**（excel `test:addin` 13 failed / 1264 passed / 3 skipped（1280），
失败与改动前一致、全是既存问题：`excelWriteEndToEnd.test.tsx` 11 条「审阅按钮尚未出现」、
`sdkIntentRouting.test.tsx` 1 条、`styles.test.ts` 1 条 `expected ['12px'] to deeply equal []`
（`.result-review-bar` 的既存字面圆角）；wps `test:wps` 3 failed / 846 passed / 3 skipped（852），
失败仍是既存 `WpsAgentPanel.test.tsx` 的 `Missing getServerSnapshot` ×3；两端 `tsc --noEmit`
与 `build:addin` / `build:wps` 见下）；**真实 Excel 2024 / 真实 WPS 任务窗格的像素验收未执行** ——
上面量的是"真实浏览器 + 真实面板 + 真实样式表"，**不是真实宿主布局**。

**人工对照页**：`… --demo` 写出 `outputs/scope-option-rows-2026-10-16/scope-option-rows-demo.html`
（四栏：WPS 列表 ×2、Excel 面板 ×2；同一份样式表，「修复前」的旧规则只在 `.legacy` 作用域里），
页面自己把量到的值填进 `[data-measure]`，四个舞台都钉在 **396px** 内容宽 —— 与判据脚本
420px 视口下的面板同宽，两处报出的行高因此是同一个数（换宽度会因折行而不同：700px 宽时
legacy 那行量到 76.5px，396px 宽才是 57px）。Excel 那两栏的行高一律量**未勾选**的那一行：
「修复前」一侧把被投诉的文案也画了出来（那是**结构**差异），勾选行因此更高 ——
拿带文案的行去比就是把两种机制混在一根数上。`--shots` 另存裁剪截图；WPS 那两栏**故意不截图**
—— 变体面板是 `position: absolute; bottom: calc(100% + 7px)`，裁出来基本是白的（PIL 逐列扫
非白像素确认过），要看就用 `--demo`。`--shots` 同时用同一个 `buildDemoPage()` 重画并整页截
`demo-preview.png`，图与页同源，不会出现"页面改了图还是旧的"。

2026-10-16 WPS 任务窗格顶部那条开发诊断导航条**整条移除**：原来是三颗按钮
（「主面板」/「宽幅审阅 · 能力探针」/「WPS 格式能力 · 受控探测」），仅 DEV 渲染
（`isDev()` = `import.meta.env.DEV`），在真实宿主里独占一整行。按产品要求删掉整个
`DevNav` 组件、`DEV_ROUTES` 清单与 `MAIN_ROUTE` 常量，三处渲染点的 `<DevNav />` 一并去掉，
`styles.css` 里只剩给它的 `.dev-nav` 规则也删了。**诊断页本身保留**：DEV 下改地址栏 hash
（`#/wide-view-poc` / `#/dev-format-probe`）仍可进入，非 DEV 依旧打不开——`useWpsHashRoute`
的共享 store 必须留着，外部 `hashchange` 就是唯一入口了。
**判据**（`main.test.tsx`）：DEV 下断言 `.dev-nav` 为 null、`nav` 为 null、**页面上一个按钮都没有**、
且三颗按钮各自的文案（"能力探针"/"受控探测"）一个字都不在；只断言空类名选择器不够——换个类名
挂回来就会漏过。正向锚点是「主面板确实渲染了」，免得整页空白把负向断言蒙对。
**反向变异**：`node scripts/run_wps_dev_nav_removed_mutations.mjs` —— M1 把导航条挂回主面板
（红色是断言 `expected <nav class="dev-nav" …> to be null`，不是编译不过）、M2 去掉诊断路由的
DEV 门控；各自变红、`restore` 逐字节一致（`AppRoutes.tsx` sha256 `93ac3eda…`）、复跑全绿 6/6。
证据分档：**自动化已通过**（`test:wps` 843 passed，仅剩 3 条既存失败：`WpsAgentPanel.test.tsx` 的
`useSyncExternalStore` 缺 `getServerSnapshot`；`tsc --noEmit` 0 错、`build:wps` 退出码 0，
新产物里已搜不到 `dev-nav`）；**真实宿主（刷新加载项后的 WPS 任务窗格）未验收**。

2026-10-16 发送即送达（输入框不再等「理解消息」）：**症状**是用户点发送后，那句话一直留在
输入框里，要等「理解消息」跑完才像"发出去"了；而且运行中的发送键停在「打断并补充」（↑）
而不是「暂停」（■）——再按一次 Enter 就会拿着**同一句**去转向重跑。**根因**是 `sendMessage`
把 `setPrompt("")` 和用户气泡都放在 `await understandMessage` / 操作目标核对 / 授权表结构读取
**之后**（原 `App.tsx` 5794/5800 一带），而这三步都是可能需要几秒的 await。WPS 那边在
`await understandMessage` **之前**就把用户气泡 append 进对话（`WpsAgentPanel.tsx:2575`），
Excel 没有 —— 两端在这里本来就分叉。

**修法（受理即送达）**：在 `/function` 短路之后、本函数**第一个 await 之前**就 append 用户气泡
并清空输入框。两处例外不能提前送：① **回答澄清卡**（由 `resolveClarification` 承载，不另起
气泡）；② **目标确认后的回调重入**——`retrySendRef` 会带着同一句再进 `sendMessage` 一次，
新增显式标记 `targetConfirmedRetry`（与 WPS 同名同义）把它挡掉，否则同一句会进对话两遍。
同时删掉后面 4 处重复的用户气泡（会话式回复、结构化删除三条、主路径），并把
`pendingLocalMapping` / `pendingClarification` 两个 `find` 提到前面（只依赖 `messages` 与
`enteredText`，取值不变；「取消本地映射澄清 / 按补充回答继续」两段逻辑留在原位）。

**配套（清空早了就必须补还原）**：新增 `keepRequestInComposer()`，把「停下来让用户先调整、
这句话一个字都没被消费」的出口——无工作簿 / 手工范围草稿脏 / 目标核对未通过 / 本地读表结构失败 /
未选工作表 / 文件夹范围缺——的请求放回输入框。只填**空的**输入框：停在这里期间用户可能已经打了
新草稿，不能覆盖（同款判据见暂停出口的 `current.trim() === text ? "" : current`）。
`folderScopeSafety.test.tsx:222` 与 `messageFirstRouting.test.tsx:415` 这两条既有断言本来就是钉
这个口径的——旧实现靠"从不清空"连带满足，现在必须显式还原（否则两条立刻变红，本轮实测过）。

**铁律**：凡"受理即送达"这类改动，判据必须在**第一个 await 之前的窗口**里取证（等到 promise
resolve 之后再断言，就什么也测不出）；清空一旦提前，**每一个"停下来先让用户调整"的出口都必须
把它还回输入框**，否则用户白打一遍。

**判据（三条各自独立的现场）**：`messageFirstRouting.test.tsx` 新增 2 条 ——
① `delivers the message the moment it is sent…`：在理解 promise 挂起期间断言**输入框值 === ""**、
**`.message-row.user` 已含该文本**、**发送键可访问名是「暂停任务」且没有「打断并补充」**，
收尾后再断言用户气泡**只有一条**；② `the retry after target confirmation does not deliver the
same message twice`：走真实确认链路（点「检查数据范围」→ 完成 → 回调重入）后用户气泡仍只有一条、
`understandMessage` 只调一次，且停住时请求确实回到了输入框。
**反向变异**：`node scripts/run_message_delivery_on_send_mutations.mjs` —— M1 受理时不清空输入框、
M2 受理时不落气泡、M3 去掉 `targetConfirmedRetry` 判据、M4 目标核对出口不还原请求；四处**各自变红**
（1/5/1/2 条失败），`restore` 逐字节一致（sha256 比对），复跑全绿 21/21。
证据分档：**自动化已通过**（`test:addin` 1258 passed，仅剩 13 条既存失败：`excelWriteEndToEnd` 11 条
`Error: 审阅按钮尚未出现` + `sdkIntentRouting` 1 + `styles.test.ts` 1，均与本改动无关，已用消融法比对过
失败数一字不差）；**真实宿主（刷新加载项后的 Excel 任务窗格）未验收**。

2026-10-16 复制按钮「已复制」在 Excel 端整条失效（对齐 WPS）：**根因是挂载期开关被
StrictMode 重放打坏，不是按钮本身**。`apps/excel-addin/src/hooks/useCopyFeedback.ts` 里
`mountedRef` 只在 cleanup 里置 `false`、**不在 setup 里置回 `true`**；而 `main.tsx` 用
`React.StrictMode` 包 `<App/>`，开发构建下 React **有意**重放 effect（setup → cleanup → setup），
重放之后这个开关**永久**停在 `false`，于是每次点复制都在"写『已复制』状态"之前被
`if (!mountedRef.current) return;` 拦掉 —— 界面表现就是"点复制毫无反应"（`copied` 永远是
false，所以不出现「已复制」）。WPS 面板不走这个 hook（自有 `copiedMessageId` +
内联 `copyMessage`），所以只有 Excel 端出问题。
修法即**在 setup 里重新武装** `mountedRef.current = true;`。
**铁律**：凡"挂载期开关"（`useRef` 的 mounted/alive 类布尔），**setup 必须显式置回真值**；
只在 cleanup 里写 `false` 一律视为缺陷。全仓库同类点已核查：`useAuthorizedTarget.ts` /
`useScopeAdmission.ts` / `LocalMappingResolution.tsx` 都已在 setup 里重新武装，`useCopyFeedback`
是唯一漏网的一个。

**测试口径（踩过的坑）**：`@testing-library/react` 的 `render` / `renderHook` **不会**触发
StrictMode 重放（实测日志只有 `setup`，真实 `createRoot` 才有 `setup, cleanup, setup`）。
所以"只在重放时才暴露"的缺陷，用 RTL 挂载**测不到**，会假绿。这类判据必须用
**真实 `createRoot` + `StrictMode`**（照 `main.tsx` 的挂法）取证，并在用例里加一条
**前置闸**断言重放真的发生过（本仓库做法：记录组件实例真实经历的 `setup/cleanup` 序列并断言
`["setup","cleanup","setup"]`），否则用例可以在机制被拆掉之后依然全绿。

**反向变异**：`node scripts/run_copy_feedback_mount_guard_mutations.mjs` —— M1 去掉 setup 里的
重新武装（退回缺陷写法）、M2 把消息复制的自动清除窗口从 1600ms 拉到 3600ms、M3 让按钮不再渲染
「已复制」这三个字；三处**各自变红**、`restore` 逐字节一致（sha256 比对）、复跑全绿。
判据读两条不同数据源：hook 状态（`useCopyFeedback.test.ts` 的 StrictMode 用例）与**界面上真的
渲染出来的文本**（`MessageCopyButton.test.tsx` 里把真实 hook 接在真实按钮上的集成用例）。
证据分档：**自动化已通过**（上述 10 条用例 + 两端 `tsc --noEmit` 0 错）；**真实宿主（刷新加载项
后的 Excel 任务窗格点复制）未验收** —— 本轮改动按开发构建（StrictMode 生效）的路径验证，
生产构建下 StrictMode 不重放，该缺陷本不显现。仍未对齐的一处**有意差异**：复制**失败**时
Excel 追加一条系统消息「复制失败，请选中文字后手动复制。」，WPS 弹 `role="status"` 轻提示
「复制未成功；可以选中文字后使用 Ctrl+C」；本轮只对齐「已复制」这一项，未改失败路径。

2026-10-15 双执行方式系统完成：用户在计划预览卡明确选择「直接执行到原表」或「先生成副本比对」。
直接执行只跳过副本比对，不跳过任何别的守卫（授权、校验、复核、快照、回读、撤销）；副本比对确认后
写回原表并只关闭指定副本，原表始终保持打开。切换执行方式时旧审阅失效（写回身份包含 `mode`），
在途请求禁止切换与重复提交。两端共用语义层（`executionMode.ts`、`reviewCopyLifecycle.ts`），
各自实现宿主适配。完整设计与验收见 `docs/EXECUTION_MODE_SELECTION_2026-10-15.md`。

2026-10-14 撤销后结果卡状态收口（身份匹配缺口收尾）：`executionUndone.ts` 的归属规则改为
**精确标识优先 + planId 回退只在候选唯一时生效** —— `executionId` / `undoId` 命中只标那一条；
快照缺精确标识时走 `planId` 回退，但**只有"匹配同一 planId 的候选消息恰好一条"才自动标**，
多条缺 ID 旧卡共用 planId ⇒ **一条都不标**，并返回 `ambiguous: true`，由 `App.tsx` 追加
「本次执行已恢复，但无法确认它对应哪一条历史记录，未自动标记「已撤销」。」。不用"恢复记录消失"
推断撤销成功。工作簿、文件夹两条路径同口径；现代快照与缺 ID 历史卡混合时不误标旧卡。

2026-10-14 撤销后结果卡状态收口：新增仅界面的 `ChatMessage.executionUndone`（`types/chat.ts`），
由 `useUndoSnapshot` / `useFolderUndo` 的 `onRestored` 回调在**每次撤销真正成功**时触发；
`conversation.ts` 恢复历史时 `=== true` 归一化。撤销成功后原结果卡标题改为「已撤销并恢复」、
摘要按模式区分（文件夹=文件、工作簿=内容与格式），**不再把旧验证结论/旧变更摘要当当前状态**；
撤销冲突、恢复失败、清空栈、容量淘汰、重载历史都不标「已撤销」
（无恢复记录时显示「当前没有可用的撤销记录；这不表示执行已撤销。」）。失败/未验证/部分完成/已撤销/
恢复记录不可用各自有明确文字。不改变授权、预览确认、执行与恢复边界，不改 `excel.ts` 的恢复算法。

**证据分档（本项）**：① 自动化 **已通过** —— `executionUndone.test.ts` 13 条归属规则
（含负向：同 planId 多次执行按 executionId/undoId 精确归属、**多张缺 ID 旧卡同 planId 一条不标**、
现代/旧卡混合不误标、跨模式不误标、精确 ID 匹配不到时 `ambiguous:true`）；`useUndoSnapshot` +4 /
`useFolderUndo` +3 / `chatHistory` +1、`excelWriteEndToEnd` 端到端（撤销后卡标题「已撤销并恢复」、
不含「验证通过」、按钮消失；冲突卡**不含**「已撤销」）；`test:addin` 1191 ✓ / 3 skipped，
`test:wps` 796 ✓ / 3 skipped，两端 `tsc --noEmit` 0 错、`build:addin` / `build:wps` 退出码 0。
② **浏览器视觉检查未完成** —— agent-browser 未安装/未初始化、无 playwright，沙箱无真实布局引擎；
`executionResultLayout.test.ts` 只是**静态 CSS 守卫**（锁防溢出规则不被删），**不是视觉证据**。
③ **真实宿主（刷新加载项后的 Excel/WPS 视觉与撤销交互）未验收**。

2026-10-14 消息结果展示整合：Excel/WPS 共用 `ExecutionResultCard`，结合分区变更摘要与中性底色状态标记，撤销入口保留在对应执行卡内。WPS 成功明细默认收起，异常明细默认展开；格式动作显示实际动作名称。这里只统一展示，保留两端各自的撤销可用条件、恢复范围和副本审阅语义。双端组件自动化与构建通过，宿主内视觉效果仍需刷新加载项核对。

2026-10-14 Excel 格式副本审阅：同区域格式组合允许依次试算，原生核对保存整份计划最终格式及行列尺寸；兼容模式需人工核对生成文件。测试分档与复跑步骤见 [格式副本审阅记录](EXCEL_FORMAT_COPY_REVIEW_2026-10-14.md)。

实机 Python 验收准备（2026-10-10）：使用
`powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start_manual_python_acceptance.ps1`。
该入口沿用正常模型连接与 SDK 单进程启动器，只给本次服务使用专用能力表开启 Python，
默认配置保持关闭；先加 `-PreflightOnly` 可验证 Docker/固定镜像而不启动服务或调用模型。
测试工作簿、双后端选择、操作和期望见 [实机验收清单](MANUAL_PYTHON_ACCEPTANCE_2026-10-10.md)。

2026-10-12 启动配置复核：普通 `start_claude_sdk_server.ps1` 不会自动打开 `pythonExecution.enabled`。开关为 false 时，普通后端与 SDK 都不会注册 `prepare_python_data` / `run_python_analysis`；提示词不能启用工具。SDK 启动器现打印实际开关；Python 验收请始终使用上面的专用入口，勿同时启动两份服务。本次曾观察到更换后端进程后能力回落为 false；恢复正确入口后须重新发起任务，旧分析引用不会跨进程保留。

### 双端在线连续集成（2026-10-10，最新）

一条命令：`npm run test:online`（`scripts/run_online_link.mjs`）。它起**测试专用服务端**
`server/tests/online_entry.py`（只绑回环 + 专用端口 + 临时能力表打开 `pythonExecution`
+ `EXCEL_BRO_CONFIG_DIR` 隔离模型目录），就绪检查（`/health`）通过后依次跑
`apps/excel-addin/src/onlineLink.test.tsx` 与 `apps/wps-addin/src/wpsOnlineLink.test.tsx`，
结束后只关本次进程。Docker 不可用 ⇒ 打印环境阻断并**非零退出**，不用 mock 或旧工件顶替。

**与工件回放集成的分工不许混**：`npm run test:cross-process`、`continuousPythonChain` /
`wpsContinuousPythonChain` 是**工件回放**（面板消费服务端算好的结果，面板自己的
`prepare_python_data` 回执**没有参与**那次计算，`/result` 由桩顶替）—— 它们是有效证据，
如实命名即可，**不要删**。在线用例才把两半接在同一条链上。

**写这类用例的硬规矩**：

- 两端共用 `config/online_link_scenario.json`；期望矩阵**只用于断言**，绝不用来构造
  产品返回值。
- 只替身两类东西：固定模型工具调用序列（只在 `online_entry.py` 注入）与有状态宿主
  替身（Excel 用 `officeTestHost`；WPS 只 mock `./wpsHost` 的 `getWpsApplication()`）。
  `./wpsReadonlyAgent`、`./wpsNativeComparison`、`./wpsPlanExecutor`、`./wpsWorkbookReader`
  与 `../../excel-addin/src/api` **必须跑真实实现** —— mock 掉它们等于绕过产品环节。
- 判据要落在"真的写进来了"上：撤销后目标区留着 6×3 的空矩阵，**只看行数会被空矩阵
  蒙过去**，要等表头真的出现（本轮踩过）。
- 同一视图里历史消息会各自渲染自己的操作卡，取按钮要取**最后一个**匹配；因果验证
  那一轮要先 `cleanup()` + `localStorage.clear()` 开新对话，否则会点到上一轮已撤销的卡。

**两后端终局结果工具共用**：语义只在 `server/app/result_submission.py`（归属校验、
结果卡、按引用编译计划、固定错误码）；`excel_agent` 与 `claude_sdk_agent` 只注册、
分派、返回。改一处必须同时改两端回归（`server/tests/test_normal_backend_result_tools.py`
里有"两端实现同一性"断言），文档里"两后端共用"的说法必须与实际实现一致。

**WPS 并排审阅副本只承载写入目标表**：判据是"写入只落在一张已授权表上"，不是"授权
表只有一张"（两表关联汇总会授权 3 张）。副本仍然只有一张表；跨表写入仍拒绝；来源核对
在原表上做。反向变异：`node scripts/run_wps_write_sheet_mutations.mjs`（M1 复制授权表
第一张 / M2 去掉跨表写入拒绝 / M3 退回授权张数判据，各 1 红，还原逐字节一致）。

**本机环境坑**：流水线里 node **无法用管道接子进程 stdout**（`spawnSync`/`spawn` 一律
`EBUSY`）—— 脚本要用 `stdio: "inherit"` 或把**文件描述符**传给 `stdio`，不要用 `"pipe"`。
另外驱动脚本会清空 `outputs/online-link`，不要把日志重定向到该目录内（会被删掉）。

### 2026-10-09 本地查询总数与能力提示

修改 `dataTools.ts` 的汇总说明时，总记录数必须使用筛选后原始记录计数，不能累加结果卡指标或截断后的分组。总数仅随本机结果传递，禁止加进 SDK 模型投影。工具 schema 的 union 限制只针对当前工具桥，不应修改共享 `DataCombine` 而破坏文件夹关联能力。SDK handler 仍须拒绝绕过 schema 的 join/deduplicate 调用，并保留模型按合法参数继续查询的路径。

部署这次补修需要重启 `scripts/start_claude_sdk_server.ps1` 启动的服务，重新构建并刷新两端加载项。验收：两表追加统计显示总数；限制只展示一组时总数仍正确；缺字段时不声称完整；关联/去重不读取宿主，失败有明确限制提示。实机与自动化分别记录。

### SDK 任务工具桥（2026-10-09）

SDK 在 Python 后端运行，`query_table` 要读的是任务窗格里打开且可能未保存的工作簿 —— MCP handler 读不到它。因此本地查询只能走**易失任务工具桥**：SDK handler 提交请求 → 桥按 `(taskId, hostId, workbookName, authorizedKey)` 登记并绑定 → 宿主适配器经回环 `POST /api/sdk/tool-bridge/poll` 拉取 → 本机执行**现有确定性 `query_table`** → `POST /api/sdk/tool-bridge/receipt` 回执 → handler 投影后交回 SDK。四个回环端点（`poll`/`receipt`/`register`/`result`）必须保持回环客户端 + 回环 Host + 显式 JSON + Origin 白名单守卫，不能靠 CORS 或可猜 taskId 放行。

**通道分离不许打破**：进度事件只带不透明指针 `bridgeTool{operationId,tool}`（可落账本）；请求参数只存易失桥，不进账本；值矩阵只在回环 `receipt`/`result` 出现。**模型可见回执**只允许 `status` + `resultRef` + 允许结构 + 固定 `errorCode`，`rows`/`title`/`calculation`/`warnings`/`scannedRows`/原始异常一律不出站；投影登记在 `agent_tools.LOCAL_NON_TERMINAL_PROJECTIONS`（与 `excel_agent.NON_TERMINAL_TOOL_PROJECTIONS` 合并），新增工具必须同时登记投影，否则失败关闭。

**幂等/取消/迟到**：同 `operationId` 同 `requestFingerprint` 幂等（返回 `duplicate`），同 ID 不同参 `conflict`，取消/换簿/重授权后 `late`/`cancelled`；绑定来源与回执来源不一致 → `source_changed`，结果不接、不登记。重授权或换簿由 `TaskToolBridge.open` 用 `authorizedKey` 差异检测并 `supersede` 旧会话。

**结果→计划**：`compile_result_plan` 只生成单动作 `writeValues`（验收条件由 `AnalysisPlan` 自动补 `rangeEquals`），并校验矩阵类型/尺寸、目标表授权、WPS 格数上限、疑似公式文本与来源版本；`submit_result_plan` 只产出计划进入预览，**SDK 工具绝不直接写工作簿**。`AnswerResponse.localResult` 只带引用与结构，`title` 用固定标签，值由回环 `result` 端点按引用取回。

新增限额集中在 `config/capabilities.json` 的 `agentToolBridge`（每任务工具次数、等待时间、待执行上限、结果 TTL/保留上限、结果行/列/格/字符与任务累计读取格数）。协议同步 `server/app/models.py` 与 `apps/excel-addin/src/contracts.ts`。

**三段生命周期必须分开建模**（混在一起已经产生过三类真实缺陷）：

- **结果生命周期**（`local_results`）：`close()` **只结束会话，不删结果**。此前 `finally` 里直接销毁结果，导致终局 `resultRef` 取不到值 —— 结果卡的取值发生在会话结束之后，靠的是**授权归属 claim** 而不是活跃会话。
- **会话生命周期**（`ToolBridgeSession`）：`register` 要求**活跃且归属一致**；`result` 只按归属核对（会话可已关闭）。两者判据不同，不能互相顶替。
- **授权生命周期**（`OwnerClaim` + `authorizationRevision`）：重授权只推进版本并作废旧结果，不编造澄清答案。

**有值计划不进持久账本**：`ledger_privacy.py` 按**字段名**剥离矩阵（`values`/`rows`/`headers`/`formulas`/`expected`…），同时作用于 `actions` **和** `acceptanceCriteria`（`rangeEquals.expected` 同样嵌值）。回环 SSE 仍把完整值交给本机客户端。新增内嵌矩阵的字段时，必须同步加进 `MATRIX_FIELD_NAMES`，否则会静默落盘。

**工作簿身份**：桥按 `identity`（`url:` / `wps-path:` / `name:` 命名空间）而不是显示名判归属 —— 只看名字无法隔离两个同名工作簿。客户端的归一必须与服务端 `BridgeBinding.from_request` **逐字一致**，否则轮询会被判成"不属于本次任务"，症状是"分析卡住、一个回执都没有"而不是报错。归一函数两端共用 `bridgeWorkbookIdentity`。

**授权收窄必须拒读**：`verifyBridgeAuthorization` 是两端**同一份**判据（`packages/core`）。判据必须拿**绑定时的授权范围快照**跟当前授权比；只看"当前授权非空"会放行一个少读一张表、静默算出错答案的查询。快照读取器与执行前复核读取器必须是**同一个**（Excel 由 hook 的 `begin` 快照，WPS 由 `createWpsBridgeLiveContext` 统一读）—— 两份分开写迟早对不上。

**两端接线（照此实现，不要另起一套）**：Excel `hooks/useSdkToolBridge.ts` 是**共用** hook（`hostId` + `createPorts` 是仅有的宿主差异）；WPS 传 `createWpsSdkBridgeHostPorts`。纪律：一轮内**串行**服务（重入会把同一次调用执行两遍）；取消/失败即停且不接受迟到回执；`end()` **保留绑定**（终局结果卡要用）；计算两端**同源**（`executeQueryTableData`），只有读宿主各走本宿主。WPS 的 `turnId` 由服务端 `checkIntent` 生成，绑定要**惰性**建立并按 `turnId` 判重。

**断言纪律**：新守卫要有**反向变异**证据（拆掉被测机制用例必须变红）。本批三处已证：授权收窄判据（core 2 红 / Excel 2 红）、`bridgeTool` 透传（WPS 1 红）、`wps-path:` 身份归一（WPS 3 红，含同名不同路径隔离）。测试不得用预置结果替代需要验收的计算路径 —— 前端用例把 Office.js/JSAPI 读表换成内存矩阵，但**计算仍走生产函数**。

**两端测试运行口径**：`apps/excel-addin` 有若干用例用 `readFileSync("src/App.tsx")` 按 **cwd** 解析跨 workspace 文件，必须 `cd apps/excel-addin` 再跑（`npm run test:addin`），否则 4 个文件报 `ENOENT`。core 的 `tsc` 有 15 处**既有**报错（`blankPolicy` / `columnMapping` / `columnMappingSemantics` 三个测试文件的联合类型），与 `sdkToolBridge.ts` 无关。

**仍未完成（不得据此宣称端到端已通）**：真实 Excel / 真实 WPS / 真实模型端到端**均未验证**（2026-10-10 的在线链路用的是固定模型序列 + 有状态宿主替身，只证明"接通"，不证明真实模型自主规划与真实客户端）；WPS 侧无源内容指纹（`sourceFingerprint` 为 `null` = 不比对）；未保存的同名 WPS 工作簿无法区分。



### Agent 整合与澄清接线（2026-10-08）

两端和两种模型后端共用 `taskContext`，仅投影原始用户需求和明确提交的业务决定；禁止从本地源值、映射草稿、计划矩阵或模型转录构造。问题按稳定 questionId 合并，改答替换旧答案；新任务/显式替换不得继承旧确认。答案先存再发模型请求，失败重试不能丢失或重复追加。taskContext 不是授权或写入许可，真实源值覆盖仍由本地工具决定。

本地补全复用完整的显式优先级，不从业务标签推断；新目标或排序不明仍须确认。空项/公式等错误在字段旁显示并阻断提交。协议变更须重启后端、刷新双端，SDK 不使用热重载；真实点击链路、完整请求与恢复记录均须覆盖。实施/限制及实机步骤见 [整合记录](./AGENT_TASK_INTEGRATION_2026-10-08.md)。

### 手动规划暂停纪律（2026-10-08）

**交互补修**：暂停后自动按原 `turnId` 核对一次，但不自动继续、不重发。已提交需求不能回填普通发送框；需求保留在原消息/任务中，未发送的新草稿不被覆盖。过程卡和输入区都称“暂停”。回归覆盖实际按钮点击、自动 GET 且零 POST、空输入 Enter/发送不创建新任务、连续暂停、运行中/查询失败/无检查点/已完成、新草稿保留及授权复核；无任务身份的消息理解暂停不查询、不编造检查点。没有增加持续轮询或服务端协议；刷新双端即可。

客户端暂停只能证明接收已中止，不能声称服务端检查点已经保存。先按原 `turnId` 查状态，运行中/查询失败只提供核对；确认可续检查点后才显式继续，不能自动重发原需求。普通继续不得传入全局 `modelId` 冒充已确认换模型；事件游标来自客户端实际接收，不从服务端最新序号推断，旧历史缺失从零补收。测试必须通过实际面板操作覆盖连续暂停、刷新不自动继续、无检查点、真实换簿/授权变化及零写入，不能用未监听的自定义事件证明守卫生效。双端实现边界和实机步骤见 [实施记录](./MANUAL_PLANNING_PAUSE_2026-10-08.md)。

### 范围提醒只保留检查入口（2026-10-08 补修）

Excel/WPS 正式面板的共享 `TargetScopeNotice` 只保留“检查数据范围”，不能通过清除提示的“暂不操作”或独立“重新检查”改变任务状态。忽略提醒保持阻断及输入，点击仅打开已有选择器；只有选择器核验并确认后才续接一次新分析，旧写入不重放。真实变化展示原/当前文件与表名，不展示私有路径；未知身份不声称已变化。独立 WPS 无范围入口时的兼容回退不变。本补修不改变授权判据和守卫触发条件，无需重启后端。

回归须覆盖四种提醒状态都只有一个按钮、核验期间禁用、需求保留、双端真实组件在确认前零字段读取/规划，以及确认后只继续一次；不能仅靠静态接线测试。

### 选择器状态与刷新（2026-10-08 补修）

`ScopeAdmissionSummary` 不再提供独立重试动作和固定授权说明；正常为一行状态，失败/只读/保护/写入能力未知保留原因，重试统一由原“刷新”触发。刷新即使目录不变也重新校验，刷新失败不能保持旧成功结果可提交。两端刷新只更新自动草稿，手动勾选与已确认授权不变；Excel 新增手动草稿必须确认后才能进入字段读取。WPS 当前表两处显示使用同一元数据，不冒充持续实时观察。回归包含真实组件点击、零字段读取与未提前授权，不仅检查静态源码。前端刷新即可，本补修未修改服务端协议。

### 数据范围统一入口、选中即校验（2026-10-08 最新）

不要在既有范围选择器之上再弹同名文件二选一窗口。实际双宿主使用 `TargetScopeNotice` 非模态提示引导到原入口，选择器展开时隐藏重复提示；有效授权内的追问静默复核。`useScopeAdmission` 的结果只存在内存，必须按选中 ID 精确返回，异步结果与本轮选择不符即丢弃；检查中/不可用不得提交，分析可用不代表允许写入。确认时再次检查并绑定所见目录的私有文件/文档身份；别仅比较显示名，也不能自动去掉失效项后确认剩余范围。

Office 不取单元格/usedRange；WPS 的 `readWpsHostProbe(...,{includeSelectionPreview:false})` 不访问 Selection，表目录/尺寸不是数据行。文件夹 `/api/folders/admission` 只能检查已登记文件，不接受任意路径，不提交 selected_targets；XML 元数据检查与文件哈希不会向模型发送行。限额在 `config/capabilities.json`，已知只读/保护/锁定限制前置显示，执行前还需复核。openpyxl 对受保护的既有写入目标必须显式拒绝，不能因为库允许写就绕过宿主保护。

唯一确认入口续接的是新分析，不是旧写入；`continuePendingAnalysis` 必须只消费分析入口登记的回调，取消保留用户输入。测试覆盖共享 hook、双端实际点击、同名文档替换、缺失表、未知/伪造状态、检查超时/迟到、范围确认后仅一次分析、旧写入不重放、文件夹真实合成文件与接口载荷白名单。需重启后端并刷新 Excel/WPS；真实宿主未验收，WPS 文件夹入口仍未实现。见 [记录](./SINGLE_SCOPE_SELECTOR_2026-10-08.md)。本节覆盖下方历史“独立目标卡”的正式入口要求。

### 当前表授权三态（2026-10-08）

一致弱证据的首次分析同意使用 `confirmation_required` 单份范围确认，不是身份失败。`unavailable` 仅提供重新检查，不能调用确认覆盖失败；暂缺按 `authorizedTarget.verificationAttempts`（默认 2）有限重试，已知变化或任意身份键冲突立即暂停，第三份缺失不抹除前两份冲突。真实变更保留原目标/当前目标选择。重试期间保留已有同意，耗尽或冲突后撤销；普通追问复用同会话已确认范围，关闭重载仍失效。新状态不序列化出站，不能用 UI 简化扩大读取授权。

双端共用组件与 hook；用例须覆盖实际点击、有限重试、直接调用确认被拒、变化不被重试掩盖、旧写入不重放与零字段/数据读取。只需刷新前端，真实宿主尚待验收，见 [实施记录](./TARGET_CONFIRMATION_STATES_2026-10-08.md)。

2026-11-01 补充（证据分级）：分析前置闸现在接受**文档级**证据 `documentBound`（路径身份 +
宿主声明的工作表稳定 ID）。它是**有条件的**：宿主真的报出稳定 ID 时才成立，跨"重载任务窗格"
仍有效；宿主给不出稳定 ID 或工作簿未保存时它不成立，仍确认一次（文案不得暗示用户改过勾选）。
`conclusive`（对象级）含义不变，写入/恢复/隔离 Python 数据读取仍只认它。三键冲突、证据缺失、
元数据范围变化一律照旧拦下；`requireExplicitConsent` **只**禁止用文档级证据替调用方跳过显式
确认（恢复挂起任务卡），对象级证明或本会话同键同意仍然直接放行。细则见
`docs/ARCHITECTURE.md` 同日章节（含 **2026-11-01 真机复测：这一支在真机上没有生效** 的修正）。
**真实宿主仍未取得运行时布尔值**。

### 文件夹刷新不等于重新授权（2026-10-08）

`applyFolderCatalog` 的 `preserveAuthorization` 仅用于候选刷新：同绝对根、同会话、已选目标均存在才保留原确认，不能将草稿/localStorage 勾选转为授权。重选目录与会话变化默认失效。澄清范围必须用本地目录根身份，不能仅凭目录显示名与相对文件名；根路径不得加进模型意图请求。目录转当前工作簿时用 `captureWorkbookCatalog` 检查名称，再读明确选择的表。后端快照必须读取前后核对全部选中文件指纹，并在通过后提交授权现场；执行前指纹保护不得省略。

回归包含真实合成工作簿中途变更、真实 Excel 组件的打开/关闭范围、同名异目录续答及非勾选活动表读取哨兵；实机尚未运行。WPS 文件夹入口仍未实现，不以本轮修复宣称对齐。需重启 Python 后端并刷新 Excel，见 [修复与实机步骤](./FOLDER_SCOPE_SAFETY_FIX_2026-10-08.md)。

### 产品身份与内部任务分开（2026-10-08）

三个模型阶段必须复用 `assistant_identity.py`：用户面对完整 Excel Bro 助手，内部理解/路由/规划分工不作为自我介绍。只在用户询问底层模型时使用服务端当前连接的配置标识；SDK 名称不等于模型，配置标识不证明网关背后的真实厂商。配置投影仅包含 `configuredModel`，禁止复制连接对象、密钥、地址或路径；未知标识如实说明，不能根据旧对话猜测。

`ActivityLog.defaultCollapsed` 仅用于成功普通聊天的本地展示，手动展开后不因重渲染折回；正在运行、失败、等待确认、警告和中断不可借此隐藏。双端共用过程卡，不影响授权或写入审批。回归见 `test_assistant_identity.py`、真实双端发送测试及 `TaskProgressTimeline.test.tsx`；实际模型输出与宿主体验仍需实测。更新后重启后端、刷新双端，旧历史回复不改写。详见 [实施记录](./ASSISTANT_IDENTITY_2026-10-08.md)。

### 先理解消息，再按需读取字段（2026-10-08）

新消息先走 `/api/intent/message`，仅当前消息、限长纯文本历史、图片数量和模型选择；不附带 scope/workbook/priorResult/priorIntent/图片二进制。模型 `conversation` 可直接结束；只有 `needs_workbook_context` 才复核目标、扫描字段和调用字段意图。无效/截断响应不能默认读表，不能用本地关键词分类或调整标题替代语义路由。既有业务补全、专用命令与恢复仍走原受控流程，不绕过写前确认。

Office 初始化/刷新目录必须用 `captureWorkbookCatalog`，禁止启动自动读取单元格/usedRange；`catalogOnly` 表示尺寸/字段未取，范围 UI 不把占位零显示成零行空表。任务按已确认范围重新读取真实结构。WPS 复用同源消息 API，直接回复不能调用工作簿 agent。文件夹发送也先消息理解，用户主动扫描不变。目标检查发生在消息模型等待之后，期间换文件仍暂停。授权前消息理解不等于数据许可，下方旧批次“模型调用为零”仅指字段意图/规划，不再指无上下文消息路由。

回归：`messageUnderstanding.test.ts`（实际 JSON 白名单/目录取值）、`messageFirstRouting.test.tsx`（真实 Excel 组件）、WPS 面板/只读代理/真实 fetch 链路、`test_message_understanding.py`（双模型后端/协议/截断/失败）。需重启后端并刷新双宿主；实机与真实模型另报。详细设计见 [实施记录](./MESSAGE_FIRST_ROUTING_2026-10-08.md)。

### 范围确认与连续追问回归（2026-10-08）

范围面板的确认必须调用共享 `confirmScope`，不能只更新勾选 UI；确认前只检查文件/表元数据与宿主证据。分析同意绑定当前会话内已确认的文件/表范围，而不是每条聊天消息；每次追问仍重新检查。检测到变更即撤销旧同意，回到原文件也不能悄悄恢复；关闭/重载不保留弱同意。显式重新授权仍按既有规则推进 WPS 授权版本，普通追问不推进。写入仍有独立身份/快照/预览确认守卫。

定向用例：`useAuthorizedTarget.test.tsx`、`officeScopeAuthorization.test.tsx`、`WpsAgentPanel.test.tsx`、`wpsScopeDraft.test.tsx`。覆盖已选范围接线、连续追问、弱新包装对象、同名异路径、同 URL 文档上下文替换、确认期间变更/删表、空/重复选择、旧请求不重放，以及真正变更后旧同意失效。不要把静态接线断言当成真实 Office.js 验收。只需刷新前端；详细结果与实机清单见 [修复记录](./SCOPE_AUTHORIZATION_FOLLOWUP_FIX_2026-10-08.md)。

### 易失 Python 结果与本机预览准备（2026-10-07 最新）

结果只进任务内存，不落盘/日志/聊天/检查点；持有句柄不等于获得写入授权。只用服务器签发、当前 grant 下输入顺序/指纹完全匹配的结果，关闭/撤销/源变化/TTL 后重读重算。不要用 JS 重新序列化结果计算服务端指纹：Python 的 `1.0` 与 JS 的 `1` 不同，核对实际响应 UTF-8 字节。矩阵与回执独立严格校验，疑似公式文本失败关闭，不能静默转义；明确值矩阵只编译 writeValues。

两端 `prepareOfficePythonResult` / `prepareWpsPythonResult` 共用 stagePythonResultPreview，返回只读准备数据和 assertCurrent。它们尚未接入面板或写入按钮；后续必须在实际首次写入之前重新调用活性守卫、执行各宿主完整可写性检查并明确确认，原生副本不是本批功能。不能复用会保存矩阵的旧执行/撤销账本而声称“不落盘”；需要单独实现非持久执行与会话内撤销，或者先获得变更隐私边界的授权。

内部验收程序可调用 run_registered_result_internal；**没有 HTTP/模型执行路由**，总开关保持 false。模型出口仍拒绝内部 completed/receipt，不因已有结果登记就擅自扩大 SDK 权限。结果读取/丢弃接口仍默认关闭。测试、限制和下一批接线门槛见 [结果预览记录](./PYTHON_RESULT_PREVIEW_2026-10-07.md)。

### 内核门禁与授权数据通道（2026-10-07 前一阶段）

当前镜像在生成代码之前安装附加 seccomp；不要用 pids-limit 冒充禁 exec，也不要在宿主运行 worker 代码以凑分析绿灯。只支持验证过的 Linux/amd64，安装失败关闭。独立本机数据接口与 core 只读客户端不属于 PlanRequest：数据、秘密 token、路径不得进模型/存储/检查点，URL 必须本机且禁重定向；两端使用前后复核真实身份/源值，任务终止或授权变化调用 dispose。当前面板调用及文件夹适配未实现，不可直接开启开关。数据登记是受控客户端绑定检查，不是服务器可独立证明当前 Excel 文件。新镜像、API、限额、测试及后续门槛详见 [实施记录](./PYTHON_AUTHORIZED_DATA_2026-10-07.md)。

### 隔离 Python 分析的实施门槛（2026-10-07，Docker 本机运行时已验收，产品入口仍关闭）

本机最新证据见 [Docker 真实验收记录](./DOCKER_PYTHON_REAL_VALIDATION_2026-10-07.md)。用户启动 Docker Desktop 后，已构建专用镜像并回填真实 RepoDigest；15 项真实 Docker 合成数据测试通过。不要继续沿用历史“无引擎、整体 skip”作为当前结论，也不要把本机通过推广为朋友电脑或产品入口已可用。

在仓库根运行 `.venv\Scripts\python.exe -m scripts.build_python_sandbox` 是当前可信构建入口，需要事先批准下载公开依赖/镜像与使用磁盘。只提交三个明确文件的 tar 上下文；独立临时 CLI 配置不读取用户 Docker 凭据，固定本机 `-H` 和 `--builder default`，核实唯一运行中 docker driver 节点与引擎前后身份。若插件目录布局不受支持、builder 不可核实或 RepoDigest 缺失，拒绝构建/配置，不能换远端 builder 或把 image ID 当 repository digest。脚本不自动改功能开关；镜像记录为机器本地前置，不是发布安装包。

容器内固定 `OPENBLAS_NUM_THREADS=1`、`OMP_NUM_THREADS=1`、`MKL_NUM_THREADS=1`、NumExpr 两项线程变量为 1；不得继承宿主覆盖值或为解决库导入而扩大进程限额。流读取采用 `read1`（无此接口的原始流用 `read`），避免短输出直到 EOF 才可观察；取消测试必须证明分析代码已开始，再取消并核实确切容器消失。

允许模型生成 Python 分析代码，不允许宿主裸执行、任意 Bash/Read、直接修改工作簿或默认数据出站。先读 [设计](./SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md) 和 [DeepSeek 实施提示词](./DEEPSEEK_ISOLATED_PYTHON_PROMPT_2026-10-06.md)，不复制已停用的 SDK 代码草稿。默认关闭；隔离不可用时明确拒绝并保留现有受控工具，不能静默切到非隔离路径。运行环境、依赖安装、镜像构建或付费模型验证须另获明确授权。

**运行时三层分离**：协调器 `python_analysis.py` → 可替换运行时适配器 `python_docker_runner.py` → 容器内固定工作程序 `docker_worker.py`。Docker 是第一种实现、非强制依赖；适配器可替换。固定启动链由可信代码决定，模型只提供 code 经 stdin 传入（绝不进命令行、不用 pickle、不拼 shell）。容器安全配置必须由环境强制并**逐项校验**（`validate_isolation_config`，缺任一项拒绝执行）：`--network none`、`--read-only` + 限大小 tmpfs scratch、非 root（**拒绝一切 root UID 形式**）、`--cap-drop ALL`、`--security-opt no-new-privileges`、内存/swap/CPU 速率/累计 CPU 时间（`--ulimit cpu`）/pids 限额、`--log-driver none`；不挂载用户目录/仓库/工作簿/凭据/Docker socket。

**输出运行期有界（不是"有界临时文件"）**：`run_python_isolated` 用 `Popen` + 两条 reader 线程**并发**读取 stdout/stderr，各通道独立字节上限（`_OutputCollector`）；任一通道触顶立即 kill 容器并进入终止核对——不把内存风险转移到宿主磁盘，也不等进程自然结束。原始输出**不落**宿主临时文件/日志/历史/账本（只在有界内存中暂存用于协议解析）。`maxStderrBytes` 与 `maxOutputBytes` 是两个独立预算，stderr 不得无界。

**清理可核实（不是"清理闭环已完成"）**：每次运行绑定可信侧运行身份（唯一容器名 `--name excel-bro-py-<uuid>`）；`rm -f` 之后必须 `container inspect` 核实，返回 `gone`（确认消失）/`unknown`（无法确认）。清理未确认时**不发布结果、不复用沙箱**。取消经 `cancel_check` 回调实现，清理后按原契约传播，取消后的迟到结果一律拒绝。禁止全局 prune / 宽泛前缀删除，只清理本容器。

**端点核实并锁定（不是"剔除变量即绑定本机"）**：仅剔除 `DOCKER_HOST`/`DOCKER_CONTEXT` 环境变量不等于绑定本机——Docker 仍会读取磁盘上持久化的 current context。`verify_local_endpoint()` 用 `docker context inspect` + `docker info` 查询**实际端点**，只接受本机 `npipe://`/`unix://`，拒绝 `tcp://`/`ssh://`/未知；`VerifiedEndpoint` 一旦核实即把整个生命周期（health/image/run/stop/rm/verify）用 `-H <host>` 锁到该端点，运行期不再重新解析 context；不修改用户全局 context。端点解析**必须取结构化字段 `Endpoints.docker.Host`**（context inspect 的权威来源），不能只取第一个 URL 匹配——真实输出里 `Metadata.otel.OTEL_EXPORTER_OTLP_ENDPOINT`（`unix://...`）会出现在 docker `Host` 之前，取第一个匹配会把 otel 端点误锁进生命周期；结构化字段缺失时才退回正则（`docker info` 自由文本形态）。

**端点必须是经结构校验的本机形态（不是判 scheme 前缀）**：`startswith("npipe://"/"unix://")` 不足——`npipe:////remote-host/pipe/docker_engine` 前缀合法、实为远端主机，会被前缀判定误接受。`_classify_local_endpoint()` 按各形态的**结构**逐段校验：Windows 命名管道只接受本机命名空间 `//./pipe/<name>` 与扩展长度前缀 `//?/pipe/<name>`（远端主机段、空管名、缺失 `pipe/` 段一律拒绝）；Unix socket 只接受本机**绝对路径**（`//host/...` 远端 authority、相对路径、根目录一律拒绝）；`tcp`/`ssh`/`fd`/`http`/空/畸形一律拒绝；**不把任意主机别名改写为 local**。执行路径**不按对象名信任** `VerifiedEndpoint`——`endpoint_is_locally_verified()` 对传入实例重新做结构校验，`run_python_isolated` 入口与 `_run_docker_cli` 最内层各有一道校验，非法端点在任何 docker 调用前失败。**运行期端点核实 ≠ builder 核实**：本 runner 的 `-H` 锁定只证明运行生命周期锁到本机；镜像构建阶段的远端 buildx builder 由构建流程单独核实，不因 runner 的 `-H` 锁定而被视为安全。

**回包严格校验**：`completed` 必须带 `result` 键（显式 null 合法，缺键不合法），拒绝 NaN/Infinity/过深/元素爆炸/畸形 UTF-8/多包协议；`status=error`（即便退出码 0）一律映射为 `error`。镜像身份必须**不可变**（`name@sha256:...`）；缺 digest 时 `image_ref()` 返回不可解析哨兵（失败关闭），不静默降级为可变 tag。

构建归档内部路径以 `server/` 为根，但不把整个目录交给 Docker：`build_python_sandbox.py` 明确只打包 `.dockerignore`、`docker/Dockerfile`、`app/docker_worker.py`。`server/.dockerignore` 同时采用默认全排除，不再重纳入 `.env.example`。Docker 日志不作为代码/数据/输出的持久化出口（`--log-driver none` 强制）。

**真实验证与未实现分列**：`server/tests/test_python_docker_real.py` 本机 15 项通过；其它机器的引擎/镜像缺失仍显式 skip，不冒充通过。完全禁 exec、授权句柄登记、结果库/计划编译及双端入口仍未实现，产品入口保持关闭。mock 层缺陷复现与独立反向变异在 `server/tests/test_python_docker_defects2.py` + `scripts/mutate_python_docker_defects2.py`（M1–M21）、`server/tests/test_python_docker_defects.py` + `scripts/mutate_python_docker_defects.py`（M1/M5/M6）、`scripts/mutate_task_ledger_wal.py`（M13/M14，SQLite 并发首建只读）。**代码与部分真实证据不等于安全闭环完成**；不得只凭这 15 项通过打开 `pythonExecution.enabled`。

**资源参数严格校验（不是只查 `isinstance(int)`/`< 0`）**：Python 里 `bool` 是 `int` 子类，`capability_int` 的 `isinstance(value, int)` 会把 JSON 的 `true` 当合法配额（`True` 冒充 1，限额被静默放大）；`capability_float` 不拒绝 `NaN`/`Infinity`（`NaN` 与任何上限比较都为 `False`，可绕过全部限额）。修复：`capability_int` 显式剔除 `bool`，`capability_float` 增加 `math.isfinite` 校验（两者都是 Python 侧通用 bug，`config/capabilities.json` 中的 `bool` 只出现在 `enabled`/`allowHyperlink` 字段，均不由这两个 helper 消费，故不改变其它配置规则）。runner 侧另有 **Python 专用**严格校验：`_require_positive_int`/`_require_positive_finite_float`/`_validate_resource_limits` 在 `build_run_arguments`（适配器入口，构造 CLI 前）与 `run_python_isolated`（执行入口，起 Docker 前）各校验一次，拒绝 `bool`/非整数/零/负数/`NaN`/`Infinity`/非数值，**不得静默钳位或降级为默认/无限**。注意语义区分：`cpu_quota`（`--cpus`）是 CPU **速率**上限，累计 CPU 时间另由 `cpu_seconds`（`--ulimit cpu=<s>:<s>`，RLIMIT_CPU）表达；`pids_limit`（`--pids-limit`）限制进程/线程数，**不表示禁止 exec**。

**冗余守卫的变异纪律（WPS 原生比较）**：原表侧漂移由两道同性质守卫共同保证——`assertOriginalUnchanged`（重新读原表、比对快照 → 文案「原表或授权范围在预览后发生变化」）与后道逐格快照复核（文案「原表在并排审阅后发生变化」）。拆其中任一道，另一道仍可能拒绝，因此**不能**用"拆一道后用例变红"冒充该道守卫的独立变异；测试断言**前一道的确切文案**才能把判别力落到 `assertOriginalUnchanged`。副本侧漂移先于原表侧被检查，其用例独立判别副本侧值比较。常量单元格改动不联动公式（`Range.Value2` setter 才联动），按实际合法拦截路径验证即可，不冒充公式重算。

已落地的隐私/入参防线（修改受控工具时不得拆掉）：`run_python_analysis` 的参数在 assistant 消息进入消息历史、检查点与恢复现场**之前**做可信侧最小白名单投影——只保留 `code` 固定占位，`datasetHandles` 投影为空列表（句柄登记未实现，不保留模型自报的任何句柄原值），未知字段、嵌套对象、任意自由文本、畸形 JSON 一律丢弃（不做哨兵扫描、不做全局替换），入参校验与历史投影是两道独立边界，被拒绝的参数不会因"没执行"就原样进历史；同轮执行参数走受控内存，恢复现场（旧检查点的 `messages` 与 `pendingAssistantMessage`）统一投影后判 `code_not_preserved`。恢复现场还必须做**回执归属核对**（assistant 调用 ID→工具名 的关联纳入 pending，孤立 ID/名称冲突/关联歧义失败关闭拒绝恢复，合法非 Python 回执保持内容顺序、一致镜像不误拒）与 **loopState 收口**（Python 相关状态判定同时考虑 messages 与 pending，涉及 Python 时清空 `lastErrorMessage`/`lastWarning`，不得进入任何新持久记录）。回执投影按封闭枚举组合再生 `message`；执行路径（解析/校验/协调器/分发/投影/序列化）异常统一映射为固定 `execution_error` 回执，普通后端与 SDK 同源，`asyncio.CancelledError` 继续传播。继续开发时保持纪律：先写能复现缺陷的合成哨兵测试再修生产代码；模型出口、持久化出口（SQLite 账本全列、每次保存边界）、异常出口、恢复现场分别取证；普通后端与 SDK 的工具契约逐字同源，SDK 内部转录无法投影这一点要如实记录而非假装覆盖。既有反向变异（历史投影/回执校验/未知字段/执行异常/恢复投影/归属核对/loopState 收口）作为回归护栏。

开工核对源码，不把 `planningReadPolicy=none` 的拒绝改为全局放行，也不把任意工具 dict 序列化给模型。仅为批准工具添加有类型的输入与模型出口投影，协议同步 TS/Python。客户端读值只能投递给经验证的本地执行器；本地执行授权和模型接收数据授权分开。Python 的 stdout/异常/摘要/值矩阵与嵌入数据的代码按敏感输出处理，检查日志、SSE、SDK 转录、缓存及恢复账本所有出口。

必须有真实隔离集成测试，不只 mock SDK 或检查配置文本：越权读宿主文件/凭据、联网、派生进程、资源耗尽、超时/取消整树终止、运行器缺失/配置损坏的拒绝。使用纯合成敏感哨兵，先证明哨兵存在于本地输入/输出，再断言真实 fetch、模型消息、日志和持久记录不包含它；摘要与错误出口单独测试。某平台测试跳过就是该平台未验证，不等于通过。做独立反向变异，但变异红灯不能代替真实隔离证据。

Excel/WPS 同步覆盖实际授权、运行、结果审阅、确认写入与取消点击链路，确认前新增宿主写入为 0；授权变化/现场冲突停止旧结果，不丢补充答案、不循环追问。按 AGENTS.md 跑双端 test/build 与服务端 pytest/compileall，修改 core 另跑 test:core；分别报告失败、环境阻断与真实宿主待验，未实现不能记作待验。本轮文档更新仅检查 diff/链接，未运行模型或宿主、未启用 Python。

### 输入失败双守卫的独立变异（2026-10-07）

`test_python_docker_defects2.py` 用 `Event` 控制真实 `_feed_stdin` 记录错误的时机，不用固定 sleep 排序。运行中输入失败须返回 `error / execution_error`，不接受 timeout 兜底；自然结束场景只在循环内检查之后释放输入失败，即便 stdout 是合法完成回包也不得发布结果。

`python scripts/mutate_python_docker_defects2.py M23a` 只拆循环内守卫，`M23b` 只拆自然结束后的守卫。兼容命令 `M23` 顺序运行两项，任一 RED-MISS 即失败；不得同时拆两道守卫充当独立证据。这两项在独立 pytest 进程内存中加载可信仓库源码的单守卫变异，不写回生产源码，不运行模型提供的分析代码。脚本注册、另一道守卫保留、兼容入口顺序和失败停止由 `test_python_docker_mutations.py` 回归。此证据不替代真实 Docker 验收，产品入口保持关闭。

### 新功能双端完成门槛（2026-10-05）

新增功能必须同时覆盖 Excel/WPS，按相同用户任务、确认与安全语义验收，不只检查外观。开工前在计划表列出双端现状与缺口，实施后分别报告实现、接线、自动化及真实宿主证据；单端上线不得标整体完成。宿主限制必须有证据与真实降级，分批实施须用户明确同意；未实现不可写成待实机验收。共享协议/逻辑优先复用，宿主特定 API 留在适配层。

双端代码改动至少运行 `npm run test:addin`、`npm run build:addin`、`npm run test:wps`、`npm run build:wps`，以及服务端 pytest/compileall；修改共享 core 时另跑 `npm run test:core`。只修改说明/计划等文档时检查 diff 与链接，不以文档更新冒充代码验收。审阅布局现为「按整个显示器可用工作区 35/35/30」（见 2026-10-14 一节），两端自动化已接线，真实宿主截图验收仍独立待验。

### 审阅窗格 35/35/30 验收（2026-10-14）

**口径先对齐**：AI 比例与外侧比例的唯一参照都是**整个显示器可用工作区宽度**（扣除任务栏），不是原表窗口客户区、不是右侧窗口、不是任务窗格 iframe。默认 **副本 35% / 原表+AI 65% / AI 30%**；AI = `clamp(0.3W, 320, ⌊原表/2⌋)`，原表可见表格 = 原表 − AI，`degraded` 在原表可见宽 < 280 或 AI < 320 时为真。旧口径（AI = 右侧窗口客户区的 40%）正是把原表挤到约 24% 可视宽度的算法，见到的任何"右侧窗口约 40%"文案都属于回归。

**Excel 最短验收步骤**（需重启 Python 服务 + 刷新加载项）：
1. 在独立、已保存的测试工作簿里生成并排副本；确认左副本约占整屏可用宽度 35%、右原表+AI 占 65%。
2. 看 AI 窗格：应尝试收到整屏可用宽度的约 30%。若卡片出现"窗口已排列…AI 窗格自动比例尚未确认：请拖动「Excel Bro」左侧分隔线，把 AI 调到整个屏幕可用宽度约 30%"——这是真实回读结论（宿主不支持 `TaskPaneApi 1.1` 或静默忽略），**不代表窗口没排好**，也不要当成"已按 35/35/30 自动完成"。
3. 手动把 AI 拖到别的宽度，再改窗口大小或挪到另一台显示器：本次审阅**不应**再自动抢回宽度。
4. 点一次「重新排列」：才允许重套默认比例（并清掉"已手动拖过"标记）；不点则任何消息更新、流式输出、回读、保存工具都不触发重排。
5. 放弃/确认审阅后，只恢复**本次流程自动改过**的宽度；用户中途拖过的宽度不被强行覆盖。

**WPS 最短验收步骤**（只需刷新加载项）：
1. 独立、已保存的测试工作簿里进副本审阅；确认出现「重新排列」入口，且它在生成副本时**不会**自己触发。
2. 点「重新排列」：应如实报告"WPS 窗口已重新排列"；若环境缺并排比较能力则报未确认。WPS 的 `Window` 没有可写宽度成员 ⇒ **不能**由代码设置两张表的精确宽度，只能靠原生并排比较；AI 窗格宽度走 `TaskPane.Width`，失败时显示"WPS 无法由代码设置 AI 窗格宽度…拖到整个可用宽度约 30%"的手动指引。
3. 手动拖过 AI 边界后，改主窗口宽度不应被自动纠正；Ribbon 重新打开也不应套用默认 1080。
4. 全程必须原表零写入；重排只改窗口几何，不产生撤销记录，也不等于核对通过。

**自动化与实机分开报告**：本轮自动化覆盖 35/35/30 由整个工作区计算（且断言 AI ≠ 右侧窗口的 30%）、不同工作区尺寸/DPI（96/144/192）/负坐标、真实降级指引、消息更新不重排、手动不被覆盖、`rearm` 才重套默认、「重新排列」调用对应宿主适配器、排列失败不影响业务状态。**本轮未操作真实 Excel/WPS 窗口**，上述截图项属未执行的实机验收，不得用替身或静态 CSS 冒充实机通过。

### 审阅副本生命周期验收（2026-10-14）

**先对齐四件事**：① 副本一律是宿主**内存里的未保存工作簿**，全产品**没有**磁盘临时文件；② 「副本关掉了」与「原表写回成功」是两件事，任何一边失败都不许改写另一边；③ 身份判据是本次随机的 `EB_<20hex>` 表名 + 宿主身份元数据 + 唯一性，"工作簿N"、`EB_` 前缀、活动窗口、文件名相似都不是判据；④ **无提示丢弃副本前必须过三项判断**：身份可信、没被另存、**副本里没有额外人工修改** —— 前两项成立**推不出**第三项。

**Excel 最短验收步骤**（需重启 Python 服务 + 刷新加载项；旧服务没有 `/close` 路由，会按 `unsupported` 降级为手动关闭）：
1. 已保存的测试工作簿生成并排副本 → 首屏应为**标题「请核对审阅副本」＋「左侧是审阅副本，右侧是原表。原表尚未修改。」＋范围行＋兼容模式必要说明＋确认框「我已核对副本，确认按此计划修改原表。」**；按钮为「重新排列」「放弃变更」「确认写回原表」，**未勾选时写回入口是禁用的**。勾选后点「确认写回原表」→ 应出现业务结论「已写回原表」＋「原表已按本卡审阅的计划写入，并通过执行侧逐格回读核对。」，副本关闭状态**另起一处独立**显示「审阅副本已关闭。」。重启提示、随机副本表名、关闭诊断、磁盘清理结果**只在「诊断详情」折叠里**，不得出现在首屏。
2. **关不掉时的措辞**（原生模式）：把副本窗口留着不手动关，副本状态应显示「关闭尚未确认，请到副本窗口前确认。」并**明确列举**"可能是保存提示，也可能是别的对话框或关闭较慢"——**不得**出现「正在询问是否保存」这类推断。到窗口前处理后，再点一次「关闭审阅副本」应报「此前已由你关闭」，**不是**报错。
3. 放弃路径（**终态**）：点「放弃变更」→ 确认框与三个审阅按钮**立刻消失**，主卡只显示「已放弃变更」＋「本次计划未写入原表。」；原表必须一格未写。副本没关掉时另给一句手动顺序「审阅副本仍打开，请核对窗口名称后手动关闭。若副本中有需要保留的修改，请先另存。」——**不得**无条件让用户选「不保存」。
4. **副本被人工修改**（第三项判断）：生成副本后，**在副本里手动改一格**（做一下明显标记），再点「放弃变更」→ 副本状态必须报「副本里还有未经确认的人工修改，未自动关闭。」且**副本窗口仍在、你改的内容还在**（加载项**不关**副本、不丢弃任何改动）；原表必须一格未写。若改成"读数拿不到"的情形，报的是「无法确认副本是否被额外修改，未自动关闭。」——两档不许混用。
5. **副本被人工修改后写回**：同样先手动改副本一格，再点「确认写回原表」→ 必须被阻断并提示副本已变化；**原表一格未写**，副本里的人工修改**不导入原表**。
6. 写回失败/结果不确定：面板应显示「已保留审阅副本供你核对。」，**不得**自动关掉副本（排查材料必须留在现场），按钮为「保留副本并结束审阅」。两档文案不许互相冒充（守卫在写入前拦下 ⇒「原表可能只改了一部分，也可能一格未改」；执行器跑到一半没确认 ⇒「部分动作已经执行，但整体结果无法确认」）。
7. 复制/另存干扰：先在副本上「另存为」一个文件再结束审阅 → 应报「审阅副本已被另存为文件，未自动关闭」，**不得**自动丢弃你自己保存的文件。
8. 同名工作簿 + 切换活动窗口：另开一个同名（甚至同名工作表）的工作簿并切到它，结束审阅仍只能关掉本次那一份副本。
9. 撤销：关闭副本后原表仍能撤销本次写入（清理副本 ≠ 撤销）。

**WPS 最短验收步骤**（只需刷新加载项，不重启服务）：
1. 生成副本 → 首屏与 Excel 端**同一句话**，只把宿主名换成 WPS（「请在 WPS 中核对副本。…」）；勾选确认后点「确认写回原表」→ 应出现业务结论「已写回原表」＋「审阅副本已关闭。」（副本状态独立成句，不再重复业务结论）。
2. **WPS 的第三项判断（重点）**：生成副本后，**在副本里手动改一格**，再点「放弃变更」→ 必须报「副本里还有未经确认的人工修改，未自动关闭。」；副本**没有被关掉**，你的改动还在。**绝不能**出现"点了放弃、副本连同改动一起消失"。同一步：改为点「确认写回原表」→ 必须被阻断，原表一格未写，副本改动不导入原表。
3. 关闭被拒或返回未关闭：文案应变「关闭尚未确认，请到副本窗口前确认。」并**只作次要状态**跟在「本次计划未写入原表。」后面；**不得**说成"执行失败"，也不得说成"未写回原表"（关闭被拒是 `Close(false)` 返回 false 的歧义信号，加载项没有对话框信息，不许推断成"宿主在问是否保存"）。
4. 手动把副本「另存为」后再放弃 → 应报「已被另存为文件，未自动关闭」，不动该文件。

**自动化与实机分开报告**：替身自动化覆盖身份唯一性/未保存标记、只关本次副本、缺少原表身份时不动手、多候选与用户另存不自动丢弃、**副本被人工修改时放弃与写回都不得静默丢失或导入原表**、**副本读数拿不到与确认被改过必须分成两档**、`close_unconfirmed`/`failed`/`unsupported`/`already_closed` 的幂等与"不推断保存提示"、自相矛盾回执按失败处理、关闭延迟与**非保存对话框阻塞**都不得当成"已关闭"、临时文件 `none` 与越界拒绝、关闭失败不改写业务结论、清理后仍可撤销。`scripts/mutate_review_copy_lifecycle.mjs` 13 个 SEED 逐个 apply/restore，全部让目标用例变红且源码逐字节还原。

**审阅卡四阶段与"旧回调不可写回"另有三条硬证据**（2026-10-14 文案/层级重构时补）：
- **放弃后旧按钮与回调不可写回**：拿放弃**之前**那个闭包（`__reactProps$` 上的 `onClick`）再调一次，原表一格都不许动。这条用例在修复前是**红**的（`expected 1 to be 0`）：闸门当时读的是渲染期闭包变量，闭包里的"计划还在"是过期事实。修复后闸门只读实时事实（`wpsWriteBackBlockReason` + 授权检查后二次判定 + 在途闸），两端同组用例转绿。
- **清理失败不恢复审阅状态**：`close_unconfirmed` / `failed` / `identity_lost` / `extra_edits` / `edits_unverified` 八种没收尾结果逐个断言"确认框与三个审阅按钮都不许回来"，副本状态只作次要行。
- **不同结果不混用文案**：写回成功 / 写回未完成 / 写回结果不确定 / 已放弃四档互相 `not.toContain` 对方的关键句。两端各有一组显式命名的**四阶段界面**用例（`审阅卡四阶段界面`），逐阶段验收首屏与折叠层的分工。

**本轮未操作真实 Excel/WPS 窗口**，「副本被自动关掉、宿主保存提示出现、用户选不保存后窗口消失」以及上面 Excel 第 4/5 条与 WPS 第 2 条（人工修改不被静默丢弃）必须真机跑一遍才算闭环。

**未消费的配置项不算已实现能力**：`reviewCopyLifecycle` 只保留服务端真正消费的 `closeWaitSeconds` / `closePollIntervalMs`；`maxCloseAttempts` / `tempCleanupWaitSeconds` / `tempRootLabel` 已删除。**没有自动重试循环** —— 重试由用户再点一次驱动。

### WPS AI 40% 比例验收（2026-10-05）

刷新/重新打开 WPS EB，使用独立、已保存的测试工作簿生成本次原生副本。AI 应约占包含两张表的 WPS 主窗口 40%，不是任一表格子窗口的 40%；主窗口逻辑宽 1200px 时目标为 480 CSS px，极窄时共用半窗口产品上限。调主窗口宽度等待约 1.5 秒应跟随；手动拖 AI 边界后再改主窗口宽度，本次审阅不抢宽，Ribbon 重新打开也不应套用 1080；新副本重新启用自动比例。放弃/确认/卸载尝试恢复进入审阅前宽度，未生成副本不临时调宽。不要把尺寸调整当作自动批准写入。

缺少窗口/窗格接口、另存副本、切换到其他文档、关闭副本、身份重复等显示手动指引，不创建新窗格、不读写单元格；API 成功但实际宽度未达到目标时同样不得报成功。100%/150%/200% DPI、WebView 缩放与跨显示器实测仍待验：DPR 发生变化时本轮停止自动调宽，按指引手动调整或重新生成审阅重新校准；不是已实现无缝跨缩放随动。此 WPS 前端变更无需重启 Python，也不接 Excel measure 端点。

新增 `wpsReviewPaneSizing.test.tsx` 覆盖比例/单位校准、实际宽度确认、主窗口随动/防抖/隐藏/卸载、手动与 Ribbon 所有权、退出恢复、身份变化、能力缺失、隐私哨兵和 StrictMode；面板用实际按钮验证进入/放弃审阅，原表零写入。共享目标与判定在 core 测试，Excel 原 12 条比例回归保留。反向变异拆掉手动优先、现场身份、普通宽度恢复分别 1/3/5 条红，还原后复跑。自动化与真实宿主必须分开报告。

本轮最终：WPS 433、Excel 730、core 216、服务端 790；两端 tsc/Vite、compileall、diff --check 通过。无最终失败或环境阻断；既有 React act 告警、pytest cache 写权限警告保留，不阻断测试。尚未操作真实 Excel/WPS，不能以此把所有版本标为验收通过。

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
恢复入口的**位置**（"要用户点的那一下"在不在视口里）单独验：
`node scripts/check_recovery_entry_viewport.mjs` —— 用**真实 `<App/>` 构建 + 真实 `styles.css`**
在 320x640 / 400x720 / 500x760 三档量入口按钮矩形，并额外跑一档"320x640 + 很长的阻断原因"
验证面板限高与正文自滚；脚本末尾把同一张卡塞回第一条消息再滚到底，必须量出"看不见"
（灵敏度对照）。走**原始 CDP**（`--remote-debugging-port` + WebSocket +
`Emulation.setDeviceMetricsOverride`）：本机不一定有 Playwright，且 Chrome 是 GUI 子系统
程序，Windows 上 `--dump-dom` 的 stdout 拿不到。`--shots` 出图到 `outputs/probe/`，
`PROBE_BROWSER=<路径>` 换浏览器。**这是 Chrome 的布局引擎，不是宿主窗格**。

反向变异：`node scripts/mutate_office_recovery_placement.mjs`（P1 入口搬回消息流 /
P2 记录里补按钮 / P3 选择器取第一条），`check P1` 判是否已干净变异，还原后要比 sha256。

浏览器尺寸测量单独运行 `node scripts/check_taskpane_layout.mjs [Playwright包路径]`，加 `--wps` 渲染真实 WPS 面板与宿主覆盖样式；需已有 Playwright 和 Microsoft Edge，不自动下载安装。真实组件用隔离上下文与模拟模型/健康响应，源表、模型服务和用户浏览器存储不参与；时间线与本地澄清夹具挂在消息容器中。`--without-containment` 只在测试内存中移除新宽度约束，用于证明布局断言能复现裁切，不修改生产文件。

通过浏览器尺寸检查后，仍须在 Excel/WPS 任务窗格拉窄、加宽并检查模型菜单、表单填写、键盘与真实步骤事件。jsdom/CSS 文本断言不能代替布局证据，Edge 无头渲染不能代替宿主验收。

## 2026-10-03：网页对比移除

不再构建 Excel `review.html`，WPS 不再切换到全幅网页双网格。`ChangeReviewSummaryCard` 仅显示定位、计数与真实核对状态；状态文案及 `ChangeReviewUndoNote` 保留在 `ChangeReviewPresentation.tsx`。原生 WPS 副本审阅、执行器回读证据、撤销与恢复不受此删除影响。

`webComparisonRemoval.test.ts` 验证两端旧入口及专用页面移除，同时检查证据收集与原生副本链路仍被消费。删除的是网页渲染专用测试与网格投影测试，不是执行核对测试。

## 2026-10-03：Excel 原生副本审阅接线

已保存的原工作簿、单张授权表、区域写值/矩阵公式/清空内容计划，在计划卡点击「生成副本并排审阅」。右侧仅包含该表数据与基础格式，不导出整本工作簿；核对后点「保存变更到原表」，放弃则只关闭临时副本。关闭时 Excel 若询问是否保存临时副本，请选择不保存；原文件保存仍按 Excel 正常流程进行。多表、复杂对象、R1C1 铺满和跨表/外部依赖未实现此审阅方式，不能列为“仅待实机验收”。

基础路径需支持 ExcelApi 1.13 以及 createWorkbook。支持 ExcelApiDesktop 1.1 时走自动模式：实际新窗口必须在同一 Application.windows 集合可发现，活动表必须匹配本次唯一标识；若打开到另一实例，停止并提示手动关闭副本，不写原表。不要放宽身份守卫来令假宿主通过。

2026-10-04 增加 Office 2024 等无 Desktop 1.1 宿主的兼容模式。点击「生成副本并排审阅」直接打开包含公式、基础格式与计划变化标记的单表文件，不调用不支持的 Office Window API、不先写原表。既有本地服务会尝试将原表与本次副本左右排列；失败时明确提示手动并排，不假报成功。检查公式栏及实际计算结果后勾选确认，再点「保存变更到原表」。若结果不符则放弃并手动关闭副本且不保存；保存也只执行原计划，不导入副本人工编辑。原表冲突检查与首次写入前守卫保留。仍**不能自动回读**兼容副本（UI 与最终结果不得伪称已回读副本或与副本一致）；**关闭**已实现：结束审阅时加载项请求本机桥按本次 nonce 身份关闭该副本，宿主若弹保存提示就请用户选「不保存」，关不掉则如实降级为手动关闭，详见 2026-10-14 生命周期验收一节。复杂对象、多表与跨表引用仍未实现。

新增 `manualReviewFile.test.ts` 检查公式类型、常量、格式、颜色、绝对坐标和大小限制；严格的旧宿主替身一触碰 Office 窗口 API 就抛错，测试实际 `createWorkbook` 请求体包含本地敏感哨兵。布局桥接不可用时 fetch/存储无新增调用；实际布局请求的隐私另由 `excelWindowLayout.test.ts` 覆盖：即使错误附加完整本地计划，真实 fetch 请求体也只能包含摘要/nonce 或 token。点击测试覆盖确认门槛、保存、冲突与放弃，自动布局不自动确认审阅。运行 `node scripts/check_manual_review_file.mjs` 可用生产序列化器生成纯合成内存副本，并用本地 openpyxl 独立验证 OOXML 可读性、公式/格式/标记与重算设置；不访问真实工作簿，不运行 Excel，不验证公式计算结果。自动与兼容两模式的实际宿主行为均待实机验收。

### 本机自动窗口排列的启动与验收（2026-10-04）

无需第二个服务或脚本进程。Windows 依赖 `pywin32==312` 已列在 `server/requirements.txt`（仅 win32 安装）；当前虚拟环境已有该版本。改动后重启现有 Python 服务（普通 `npm run dev:server` 或原 SDK 启动命令，不能同时占用 8765），刷新 Excel 加载项。使用单个服务进程，避免 prepare/arrange 落到不同 worker；热重载也会清掉一次性凭据。不要为了访问 Excel 把服务改成系统后台会话或提升管理员权限。

在独立、已保存的本机测试工作簿，激活已授权目标表，关闭上次临时副本且不保存，然后重新点击「生成副本并排审阅」。新版本服务在打开前记录原表与已有窗口，故不能把旧副本补认成本次新副本。正常应显示左侧审阅副本 40%、右侧原表与 AI 窗格 60%，版本 3 中间可见间隔默认为 0 DIP（DPI 缩放后换算），补偿不可见拖拽边框，两窗高度使用显示器工作区；若能力、身份、权限或多窗口证据不满足，则显示「自动排列未完成」并保留手动审阅流程。未成功定位不得改用标题猜测或全体窗口重排。Windows 系统边框样式不修改，后续宽度/间距调整只改 `excelWindowLayout.originalWidthFraction` / `windowGapDip` 并重启服务。用户已反馈上一版自动排列可用；本次顺序、间距和不同 DPI 尺寸仍需重新实测。

`server/tests/test_excel_window_layout.py` 验证原表路径、nonce、仅新窗口、进程身份、令牌来源/期限/单次使用、忙/模糊/换表、几何验证与失败回滚，以及真实 HTTP 请求、无账本持久调用和固定错误响应。尺寸测试覆盖 96/192 DPI、负坐标工作区、左副本/右 AI、6 DIP 间隔与非法尺寸零移动。`python scripts/check_excel_layout_guards.py` 对生产类做六项独立安全守卫内存变异，另独立测试改回旧顺序和取消间隔缩放的两项回归；必须逐项被测试抓住，不编辑生产源文件，不枚举或操作真实 Excel。真实显示器 DPI、Office 2024 可访问性元数据、另一实例新窗口与权限隔离仍需实机验证；自动布局不代表已实现兼容副本**单元格回读**；副本**关闭**已由 `/api/office/review-layout/close` 实现（2026-10-14 一节），跨实例同步滚动仍未实现。

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
会话目录用于 CLI 配置，本次请求使用 `--no-session-persistence`；旧版可能留下的资料
仍应限制本机访问、按隐私策略保留/清理，不能提交 Git。
该 `.ps1` 保持 ASCII 文本，以兼容 Windows PowerShell 5.1 对无 BOM UTF-8 脚本的读取；
若修改脚本提示文案，不要直接加入非 ASCII 字符。
不要把真实密钥写进示例文件。SDK 的 `max_budget_usd` 不保证等于 DeepSeek
实际计费，试用时应查看 DeepSeek 平台账单与额度。

当前 SDK 通道仅用于表结构意图与 `planningReadPolicy=none` 的无数据行规划：
既不上传原始数据行、上一轮紧凑结果、历史对话或本地文件路径，也不开
SDK 的 Read/Bash/Edit 等工具。用户显式添加的图片现已通过原生 image content block 接入，
不转成文件路径、不创建临时文件。数据行请求仍明确拒绝并提示用户切回原模型；
同一任务显式追问现通过 `taskContext` 保留原需求与用户确认，两端在字段意图/规划中连续传递；
它不是全量历史或 CLI 会话，泛指上一轮结果的请求仍应明确说明，不能期待 SDK 自动读取旧会话。
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

依次运行 npm run test:wps、npm run build:wps 和 npm run dev:wps。**WPS 开发前端只有这一个入口**：
`npm run dev:wps` → `apps/wps-addin` 的 `debug` 脚本 → `scripts/start_wps_dev.mjs`。它不再调用
`wpsjs debug`，原因是 wpsjs 2.2.3 的端口链路不可控：`util.js` 的 `portfinder` 从 basePort
往上找"能 bind 的端口"，把 Windows **TCP 排除/保留端口**的 `EACCES` 也算作"不可用"而静默跳过
（本机 3887–3986 被保留，实测 3889 → 3987）；`debug_publish.js` 再把这个结果写进 WPS 的
`publish.xml`，并把 `package.json` 的 `scripts.dev` 改写成 `vite --port <扫出来的端口>`
（CLI `--port` 会覆盖 `vite.config.ts`，`strictPort` 只保证"不换端口"，不保证"是哪个端口"）。
三处端口虽然当时一致，但端口是"扫"出来的：一旦落在后端 `ALLOWED_ORIGINS` 之外，任务窗格的
`POST` 被 403 `FORBIDDEN_ORIGIN` 打回、`GET` 也读不到响应，界面只剩 `Failed to fetch`。

启动器按顺序做四件事，每件都有可见输出：

1. 从 `config/capabilities.json` 的 `devServers.wpsPortCandidates`（默认 `3889、3890、3987、3988`）
   **按声明顺序**实测 `127.0.0.1` 能否 bind，取第一个可用端口；全部不可用时打印排查命令
   （`netsh int ipv4 show excludedportrange protocol=tcp`、`netstat -ano | findstr :<端口>`）并以
   退出码 2 失败。**不会**换到列表外的端口，也不删保留端口、不改用管理员运行。
2. 用该端口更新 `%APPDATA%\kingsoft\wps\jsaddons\publish.xml` 里本加载项那一条（其余插件条目
   逐字保留），写后读回校验；地址不一致就中止，不落盘。
3. 以 `EB_WPS_DEV_PORT=<端口>` 起 Vite（直接 spawn vite 入口，避免 `npm run` 的 cmd 外壳留下
   占端口的孤儿进程）。`vite.config.ts` 只认这个变量并保持 `strictPort: true`：端口在检查后被抢占
   就报错退出，不漂移；`EB_WPS_DEV_PORT` 不在候选内会直接拒绝（后端不认识它）。
4. 打印三项自检：登记地址 vs 监听地址 vs 后端来源。后端两项实测是 `GET /health` 是否回
   `access-control-allow-origin`、以及来源守卫探测（`POST` 到不存在的路由：被允许 → 404，
   被拒 → 403），二者都不写账本。后端不认识该来源时提示**手动**重启 `npm run dev:server`
   ——启动器不会停掉正在运行的后端。

`devServers.excelPort`（3000）是 Excel 开发页端口，WPS 候选里出现它会在配置校验阶段直接报错：
修 WPS 不允许占掉 Excel 的启动能力。改端口只改 `config/capabilities.json` 一处——后端
`ALLOWED_ORIGINS` 由同一份配置派生（Excel 的四个来源在前，各候选的 `127.0.0.1`/`localhost`
精确来源在后），`_review_layout_origin` 仍只认 Excel 那四个；**改完必须重启后端**，否则白名单还是旧的。
登记变更后需保存文档并完全重启 WPS（功能区/登记只在重启后生效）；只刷新任务窗格不会重读登记。

`npm run test:wps` 会先执行 `pretest` 的 `scripts/check_wps_dev_ports.mjs`：覆盖端口被系统保留
（EACCES）与被真实 socket 占用、候选内不漂移（挑不到就报错）、登记地址等于监听地址、登记改写
保留其他插件且在结构不确定时 fail-closed、`package.json` 不再被当成运行时配置、后端不再自带
第二份端口清单。来源一侧由 `server/tests/test_dev_origins.py` 钉住：候选内每个来源都能读
`/health` 与只读发现端点，候选外来源的 `POST` 仍 403，Excel 布局守卫仍拒绝 WPS 来源。

已有 WPS 进程可能不会
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

W2 开发页从 `devServers.wpsPortCandidates` 里由 `npm run dev:wps` 当时选中的回环端口访问本地
服务（启动器自检会把它打出来）；服务端只对白名单中的这些精确来源开放 CORS，白名单与启动器读
同一份 `config/capabilities.json`。只读 Agent 的真实验收见
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

**WPS 开发导航与格式探测入口（2026-10-13）**：任务窗格的路由由 `apps/wps-addin/src/useWpsHashRoute.ts`
统一管理——它用 `useSyncExternalStore` 订阅模块级共享 store，以 `window.location.hash` 为唯一真源，
`DevNav` 导航条与 `AppRoutes` 页面选择共用同一份 hash 状态（此前 `DevNav` 只写
`window.location.hash`、`Root` 只读一次 hash 却不订阅 `hashchange`，点按钮不触发重渲染，
界面停在聊天面板）。hash 归一：`""`/`/` 都归一到主面板 `/taskpane`，未知 hash 兜底到主面板；
支持首次直入诊断地址、点击切换、浏览器前进/后退与返回主面板入口。诊断路由
（`#/wide-view-poc`、`#/dev-format-probe`）由独立模块 `apps/wps-addin/src/isDev.ts` 的
`isDev()`（`import.meta.env.DEV`）门控，非 DEV 手改 hash 也兜底到主面板，**不能打开诊断页**。
导航本身不触发格式探测或写入：路由只决定渲染哪个面板组件，运行前校验
（`assertWpsFormatProbeTarget`）与用户确认仍在各面板内部。诊断入口位于任务窗格顶部的
`DevNav` 导航条（仅 DEV 渲染）：主面板点「WPS 格式能力 · 受控探测」进入 `#/dev-format-probe`。
若宿主缓存导致导航不更新，需完全重启 WPS / 重新旁加载（沿上文「完全重启 WPS」的说明），
不要用整页 reload 掩盖路由缺陷。相关测试：`useWpsHashRoute.test.tsx`、`main.test.tsx`
（断言**渲染页面变化**而非 hash 字符串）、`main.production.test.tsx`（非 DEV 门控）。

**格式探测能力组选择与实机证据（2026-10-13）**：`runWpsFormatProbe` 支持按能力组
选择探测（`options.groups`，组枚举 `font`/`alignment`/`fill`/`borders`/`size`/`autofit`），
面板 `WpsFormatProbePanel` 顶部提供多选；使填充/边框这类此前因前置失败被 abort 而未测到的
组能在**新的独立测试基线**上单独重跑。**同一轮内任一成员恢复失败仍立即停止后续写入**，
不允许在已受污染的状态上继续。报告新增 `groups` 字段留痕本轮探测范围。
真实 WPS（`Application.Version` 报告 `12.0`）实机探测结论见
[证据文档](./WPS_FORMAT_PROBE_EVIDENCE_2026-10-13.md)：8 项 `restore_verified`
（Font.Bold/Italic/Name/Size/Color、HorizontalAlignment、WrapText、RowHeight）；
`Font.Underline` 因哨兵误用布尔（应为 `XlUnderlineStyle` 枚举）而 `failed`，
`ColumnWidth` 因量化漂移 `failed`（写 12 回读 12.0088…、恢复 8.4533…≠8.4444…），
两者均**不得用宽松比较掩盖**；Interior/Borders 未测、AutoFit 未验证。原始结论
`baselineRestored=false`、`aborted=true` 保留不篡改。

**格式动作闭环（2026-10-13，Task #53）**：已把实机 `restore_verified` 的属性子集接入完整产品链路
（能力声明 → 白名单计划 → 校验 → 执行 → 逐格/逐行回读 → 用户审阅 → 撤销与恢复核对），**按属性开放，
不因某类通过就整类放开**。已接通的动作与属性子集：
- `setFont`：`bold` / `italic` / `name` / `size` / `color`；
- `setAlignment`：`horizontal` / `wrapText`；
- `resizeRange`：**仅** `rowHeight`。

属性级守卫在 `apps/wps-addin/src/wpsFormatCapability.ts`：`assertWpsVerifiedSetAlignment`
明确拒绝 `vertical`（哨兵与原值同值，尚未实机验证）、`assertWpsVerifiedResizeRange` 明确拒绝
`columnWidth`（量化漂移，恢复未通过）；`setFill`/`setBorders`/`autofit`/`setNumberFormat` 未接通，
均**显式拒绝、不静默忽略**。枚举/颜色映射同模块（`wpsHorizontalAlignment`、`wpsColorFromHex`）。

逐格/逐行快照在 `apps/wps-addin/src/wpsFormatSnapshot.ts`（`captureWpsFormatSnapshot` /
`restoreWpsFormatSnapshot` / `sameWpsFormatSnapshot`）：因为 WPS 区域级 `range.Font.Bold` 对混合格式
返回不确定，单格探测不证明混合格式恢复，必须 `range.Item(r+1, c+1)` 逐格读取字体/对齐/换行 + 逐行行高。
**读不到（undefined/null）的属性存 `null`，恢复时跳过、不写回默认值**；任一目标属性读不到即
`captured:false`，`prepareWpsPlan` 拒绝整份计划（fail-closed）。**多格/多行缺 `Item` 时 `captured:false`
拒绝、不回退整块区域**（单格区域 range 本身就是单元格，不需要 `Item`）。执行器接入
`apps/wps-addin/src/wpsPlanExecutor.ts`（`prepareFormatAction` / `applyFormatAction` / `asFormatFields`），
与值动作（writeValues/writeFormulas/clearRange）共用 `prepareWpsPlan`/`executeWpsPlan`/`undoWpsExecution`
外壳。`config/capabilities.json` 的 `wpsWrite.allowedActions` 与
`hostCapabilities.wps-et.12.0` 的 supported 集合已同步（6 个动作）。

**格式动作闭环第二轮（2026-10-13，Task #53 三阻断修复）**：
1. **格式动作接入副本审阅**：`prepareFormatAction` 读**真实值矩阵**作 `values`/`beforeValues`/
   `previewFingerprint`（不用空矩阵冒充原值快照）；`stageWpsNativeComparison` 循环内对格式动作分派
   （采集 before 格式 → `applyFormatAction` → 回读 → 值/公式不变核对 → `changedCells` 标整块区域）；
   `verifyWpsNativeComparison`/`compareCommittedWpsResultToCopy` 补格式快照核对（`copyFormatAfter`/
   `originalFormatBefore` 字段）。
2. **属性级恢复**：`formatActionMatches` 拆成三函数 `formatActionApplied`（全字段）/`formatActionStillApplied`
   （声明字段）/`formatActionRestored`（声明字段）；撤销用 `restoreWpsFormatSnapshot(..., declaredFormatFields(...))`
   只恢复动作声明过的属性，**保留用户对其他属性的后续修改**；`overlaps` 放行同区域两个格式动作；执行循环内
   现场采集 `beforeFormatSnapshot`（同区域多格式动作的写前状态是前一个执行后的状态）。
3. **多格采集粒度**：`captureWpsFormatSnapshot` 多格/多行缺 `Item` ⇒ `captured:false` 拒绝（不回退整块伪造
   逐格快照）。

预览对格式动作显示格式描述（`describeWpsFormatAction`，含"影响整行/整列"提示），确认卡首屏不再显示
"写入值"值矩阵对照（`wpsPlanChangePreview.ts` 的 `planActionLabel` + `buildPlanChangePreview` 产
`formatDescription`）。测试：`wpsFormatCapability.test.ts`（11）、`wpsFormatSnapshot.test.ts`（5）、
`wpsFormatExecutor.test.ts`（10）、`wpsFormatFlow.test.ts`（5，走真实 `stageWpsNativeComparison`/
`executeWpsPlan`/`undoWpsExecution`，只替身宿主，覆盖副本审阅值/公式不变、加粗后改字色/行高撤销不覆盖、
多格缺 Item 拒绝零写入、同区域多格式动作、「放弃变更」不碰原表）；反向变异
`scripts/run_wps_format_closed_loop_mutations.mjs`（M1 拆副本格式应用 2 红、M2 拆属性级恢复 1 红、
M3 伪造逐格快照 1 红、M4 拆重叠放行 1 红）还原后逐字节一致。

**格式动作闭环第三轮：组合状态收口（2026-10-13，Task #53）**。第二轮把格式动作接进副本审阅后，
同区域**组合动作**（如"加粗→居中"、同属性被后续动作覆盖）暴露出两个状态口径错误，本轮只修口径、不扩新能力：
1. **动作级执行证据 vs 整份计划最终状态**。`stageWpsNativeComparison` 循环内逐动作采集的原 `copyFormatAfter`
   是**中间态**（动作 0 执行后只有加粗），而确认（`verifyWpsNativeComparison`）与提交后对照
   （`compareCommittedWpsResultToCopy`）拿的是**最终现场**，两者一比就对"加粗→居中"误报
   "副本的格式在生成对照后发生变化"。改为字段 `copyFormatFinal`，循环内先按 `index` 写**动作级证据**，
   循环结束后新增 finalize 循环用 `captureWpsFormatSnapshot` **整体覆盖**成整份计划执行后的最终状态；
   确认与提交后对照一律用最终状态。真实外部改动（审阅后改副本/改原表）仍因"现场 ≠ 最终状态"被拦截。
2. **撤销冲突检测改为"最终恢复影响范围"**。原 `formatActionStillApplied` 逐动作要求"本动作声明的目标值仍成立"，
   同属性被后续动作覆盖时（字号 12→14→16，第一步声明 14 已被第二步覆盖成 16）会误判冲突。
   `StoredAction` 新增 `afterFormatSnapshot`（整份计划执行后最终格式）；新增
   `formatActionUnchangedSinceFinal`（现场 vs 最终状态逐格比**本动作声明过的属性**）与统一入口
   `formatActionSafeToUndo`（有最终快照按最终状态比，历史/部分写入记录缺快照时**退回**旧口径
   `formatActionStillApplied`，不静默放行）。`inspectWpsUndoRecord` 与 `undoWpsExecution` 的冲突检测都走它；
   恢复循环保持 `[...record.actions].reverse()` **逆序属性级恢复**，同属性覆盖叠加后正确回到原始混合值
   （14→13/12，12→13/12），部分重叠区域与整行行高的实际影响范围由逐格/逐行快照覆盖。

**证据分档（不得合并宣称）**：`wpsFormatFlow.test.ts` 走真实 `stageWpsNativeComparison`/`executeWpsPlan`/
`undoWpsExecution`，是**模块集成证据**（非面板点击证据）；`WpsAgentPanelFormatReview.test.tsx` 是**面板点击
证据**：真实 React 面板 `<WpsAgentPanel/>` ＋真实审阅/执行模块（`./wpsNativeComparison` / `./wpsPlanExecutor`
**不 mock**），走真实入口按钮"生成右侧副本并排审阅"→"确认写回原表"→撤销按钮
`aria-label="撤销本次执行"`。**模拟范围（不得笼统写成"仅替身宿主"）**：该文件同时替身 ① 宿主
（`./wpsHost` 的 `getWpsApplication`）、② 模型规划（`runWpsReadonlyTurn` + `api.understandMessage`）、
③ 计划校验与服务端执行账本接口（`api.validatePlan` + `registerWpsExecution`/`getWpsExecutionStatus`/
`startWpsAction`/`completeWpsAction`/`suspendWpsExecution`/`claimWpsExecutionResume`/
`reconcileWpsActionReceipt`/`finishWpsExecution`）。所以它**不是**真实模型、真实服务端或真实客户端证据。
反向变异
`scripts/run_wps_format_closed_loop_mutations.mjs` 增 M5（拆 finalize 循环 ⇒ FLOW 2 红 + PANEL 1 红）、
M6（`formatActionSafeToUndo` 丢掉最终快照退回旧口径 ⇒ 1 红），M1–M6 全红、三文件逐字节还原、复跑全绿。
**WPS 格式动作真实宿主闭环（副本审阅→确认→写回→回读→撤销）仍未验收，上述自动化均为替身证据。**

**格式动作闭环第四轮：组合边界证据（2026-10-13，Task #53 第 18 轮）**。新增
`apps/wps-addin/src/wpsFormatCombinationBoundary.test.ts`（11 条），保留真实执行器/审阅/回读/撤销，
只替身**有状态宿主**，与 `wpsFormatFlow.test.ts` 的关键差别是**行高按「整行共享状态」实现**：
旧替身给每个单元格各存一个 `rowHeight`，掩盖了"行高是行级属性"。覆盖：
- **部分重叠区域**：`A1:B3` 字号 14 → `B2:C4` 字号 16，重叠区 `B2:B3` 最终 16，其余按动作顺序，
  撤销逐格恢复原始字号（10..21 各自原值）。
- **行高的真实影响范围**：`A2:A3` 行高 20 → `C3:C4` 行高 25，最终第 2 行 20、第 3/4 行 25；用 A 列 /
  C 列读同一行读到**同一份**行高（整行语义）；预览经 `buildPlanChangePreview` 明确提示"影响整行"；
  撤销恢复各行原始行高；外部改**同一行的无关列**（B3）也判撤销冲突（整行语义）；外部改**无关行**
  则撤销成功且保留该行改动、只恢复声明行。
- **外部修改**：审阅后改副本 / 改原表的声明属性 → 拒绝写回；执行后改本次声明的属性 → 撤销冲突；
  执行后改无关属性 → 撤销保留。
- **值与公式保护**：格式目标区域 `A2:C2` **内部**含文本（A2）、数字（B2）、真公式（C2 `=B2&"元"`）；
  确认前原表零写入，执行与撤销后值/公式一字不变。

反向变异扩到 **8 发**：M7（`formatActionUnchangedSinceFinal` 关掉行高比对 ⇒ 组合边界 1 红）、
M8（`restoreWpsFormatSnapshot` 关掉行高恢复 ⇒ 组合边界 2 红）；M1–M8 全红、三文件逐字节还原、复跑全绿。
**实机验收材料**：`outputs/format-acceptance-2026-10-13/`（专用测试工作簿 `格式验收.xlsx` +
`验收步骤与核对表.md`，由 `scripts/make_format_acceptance_fixture.py` 生成/复位）。**实机仍未验收。**

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
「确认写回原表」前重新核对原表与副本（并先勾选「我已核对副本，确认按此计划修改原表。」），
才走原有白名单计划执行器；点「放弃变更」关闭副本且不保存。当前只支持单张已授权工作表，多表计划明确拒绝。
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
逻辑复用不得突破宿主能力：工作目录暂未开放；图片输入已与 Excel 共用附件处理 Hook；
工作流仍须验证并预览；固化查询必须在当前
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
当前产出 arm64（macos-15）`.pkg`，未产出 Intel 包。macos-14 runner 已退役，首次构建在调度阶段被取消，已更新 runner。本地有 Mac 时也可
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
- **WPS 规划暂停在面板重开后只能"核对"，不能"自动续跑"**（CAP-02 §5）：重开效果走
  `planningPausedEntriesFromHistory`（共享层）→ `reinspectWpsPlanningPause`（查询）→
  面板挂载 effect。三条不可越过：① 核对中**不画操作区**；② 按钮只由
  `wpsPlanningResumeView` 从**服务端刚返回的**状态推导，历史快照只许用于显示；
  ③ `revision` / `lastSequence` 一律采信服务端。改这里的条件时**不要把
  `canContinue` 和 `action` 合并成一个 if** —— 那会把"放行但非精确续跑通路"误报成
  "不能精确续跑"，等于替服务端下结论。
- **两条后端差异必须显式说明，且只描述"当前接入"**：普通模型后端有检查点
  （`loopState`/`messages`/`toolCache`）；Claude Agent SDK 后端**当前接入未实现
  检查点恢复** —— 本次接入把对话状态交给 CLI 子进程托管，服务端读不到那份现场。
  写文档时**不要**推断成"SDK 永远不支持检查点"：那是没核实过的结论。将来换接入
  方式或补做现场托管，改这句话而不是改历史结论。`planner.py` 在分发前用
  `ModelNotAvailableError` 挡住跨后端续跑。**不要**改成静默重新规划 —— 用户点的是
  "从上次进度继续"，静默从头规划会重复调用模型与工具，界面还显示"继续成功"。
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

### 规划时间策略、暂停与恢复调试

流式规划通过 `POST /api/turn/stream` 建立持久 turnId。**2026-10-10 起提醒与暂停分离**：
`planningReminderSeconds`（60 秒）只发一条**非阻塞**进度事件（`planning.reminder.<n>`），
任务继续保持运行、界面不显示"继续"按钮；真正的停止阈值是 `planningHardWallTimeSeconds`
（1800 秒），按**累计活跃执行时间**计算，到达后由 `note_progress(...,
enforce_total_wall_time=True)` 在启动下一次模型/工具调用前收口（它复用同一个边界**唯一一次**
`save_boundary`，所以 `before_model` 不会重复落检查点），事件流返回 `planning_pause`
（`pauseType=budget_exhausted`）；前端保存
turnId、revision、lastSequence、稳定 requestId 和原来源指纹。刷新或网络中断后应先调用
`GET /api/turn/{turnId}/status`，用户明确继续时再调用 `POST /api/turn/continue/stream`。

- `pauseType` 四种含义必须分开呈报：`budget`（可接续现场）、`budget_exhausted`（总时限
  已用完，不得显示"必然可以继续"）、`model_error`（换模型接续）、`interrupted`（重启找回）。
- 同一继续动作重试必须复用 requestId；旧 revision 应返回冲突。配置位于 `taskRecovery`。
  继续只接续现场，不提高 `agent.maxTurns`、`maxTotalTokens` 或错误上限。
- **检查点保存不得因取消短时暂停而被删掉**：`save_boundary` 是检查点的唯一写入点，
  每个可恢复边界都要落盘；"取消暂停"删的只是"到期即停"这一件事。
- 累计活跃耗时只算真实运行区间（模型等待、本地计算、保存核对），**等待用户不计入**；
  跨重启沿用（写在检查点 `loopState.activeElapsedSeconds`），重连/继续/重试都不归零。
- checkpoint 保存在 `tasks/state.sqlite3`，包含消息、LoopState、工具缓存、工具模式、修复
  状态和完整待处理模型响应。不要把完整 checkpoint 写入普通诊断日志。
- **SDK 后端没有检查点恢复能力**：跨后端续跑由 `ModelNotAvailableError` 挡住；这是当前
  接入的限制，不代表 SDK 永远不支持。
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
- 反向变异证据：`server/tests/_mutate_planning_time_policy.py`（M1 拆提醒 / M2 拆总时限
  收口 / M3 把墙上时间熔断加回 `should_stop` / M4 拆 `save_boundary`），4 个 SEED 必须全红，
  还原后 `server/tests/test_planning_time_policy.py` 7 条全绿。

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

WPS 的专注窗口入口在任务窗格输入框左下角 EB 菜单的「专注窗口」（`WpsAgentPanel.tsx`
的 `openFocusWindow`），生产构建含 `focus.html`（`apps/wps-addin/vite.config.ts` 多页
`rollupOptions.input`）。WPS 没有 Office `dialog.messageChild`，跨窗口通信走
`apps/wps-addin/src/wpsFocusChannel.ts`：按 sessionId 优先 `BroadcastChannel`、回退同源
`localStorage` 邮箱，仅传已校验展示快照（不传密钥/凭据/原始数据行），子窗口失联/父窗格
刷新/切换工作簿/关闭都会显示失联并停止会话。**跨 WebView 的 `BroadcastChannel` 是否真的
可用、`ShowDialog` 是否真的打开独立窗口与可关闭，只能实机验证**，见
[WPS 专注窗口交付记录](./WPS_FOCUS_WINDOW_2026-11-03.md)。

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

先确认目标再分析回归（2026-10-05，授权生命周期于 2026-10-08 修正）：共享 `useAuthorizedTarget.test.tsx` 验证弱证据先确认、强身份直接分析、取消不分析、撤销/检测到变更后不继承弱同意，以及确认/分析检查期间身份证据改变。普通追问复用当前已确认范围，但重载不恢复内存同意。`LocalMappingResolution.test.tsx` 验证业务卡没有目标授权勾选，提交仍复核现场并保留规则；两端 `*AuthorizedTarget.test.ts` 验证宿主证据不读取单元格值。WPS `wpsLocalPlanOutbound.test.tsx` 用真实 fetch 与宿主读取计数证明确认前表头/数据行/模型请求/写入均为 0，确认后补全直接到新预览；换文件和源值改变仍拒绝陈旧计划。`wpsTaskResume.test.tsx` 验证弱证据恢复先走独立目标卡，保留真实答案。2026-10-05 独立反向变异：绕过 WPS 分析前检查 1 条红、移除弱证据同意条件 4 条红、移除确认时证据键一致检查 1 条红，均已还原。完整证据与实机项目见 `docs/AUTHORIZATION_BEFORE_ANALYSIS_2026-10-05.md`。

授权目标前置确认（2026-10-05）：重点回归共享 `useAuthorizedTarget.test.tsx`、`authorizedTarget.test.ts`、`officeAuthorizedTarget.test.ts`、WPS 面板和只读代理，以及 Excel 原生审阅入口。必须检查实际点击未确认前模型/读取/写入为 0、确认期间再次切换拒绝、后台监测不重置勾选、旧写入不自动重放。实机项目见 `docs/AUTHORIZED_TARGET_PREFLIGHT_2026-10-05.md`。

以下变更必须同步更新 `docs/ARCHITECTURE.md`：

- 新增服务、数据流或持久化方式；
- 新增模型供应商适配层；
- 修改授权或安全边界；
- 新增本地工具；
- 修改动作协议或执行通道；
- 改变 `/api/turn` 状态机。

澄清恢复回归（2026-10-05）：不仅断言“调用了新一轮”，还须完整点击答案提交、范围确认、预览入口，验证确认不丢失、不重复询问；WPS 替身的 ActiveWorkbook 与 Worksheets.Item 每次都返回新对象。检查真实 fetch 与历史的隐私哨兵，以及预览路径原表写入为 0。同名异路径必须拒绝私有规则续接。实现与证据见 `docs/CLARIFICATION_LOOP_FIX_2026-10-05.md`；Excel 共享表单测试不代替真实 Office.js 验收。

启动命令、环境变量、端口或调试方式变化时，更新本文和根目录 `README.md`。

消息自动跟随回归（2026-10-08）：`useMessageAutoFollow.test.ts` 覆盖上翻后不抢滚动、新建/切换对话清除未读提示、短对话无需滚动事件也不误报；Excel `messageFirstRouting.test.tsx` 与 WPS `WpsAgentPanel.test.tsx` 从实际面板点击新对话和历史对话，验证状态不跨对话残留且保留原三点按钮结构。自动化测试不代替宿主实机滚动验收。

开发态白屏补修（2026-10-08）：上述 Hook 的 `.ts` 实现未获得 Vite React 刷新签名，新增 `useRef` 后可能复用旧 Hook 顺序。实现迁入 `.tsx`，旧 `.ts` 地址只做转发。`useMessageAutoFollow.refresh.test.ts` 使用双端默认 React 转换器及真实 Fast Refresh runtime 验证签名、旧入口与不兼容状态重置；测试使用独立依赖缓存，不得覆盖运行中开发服务的 `.vite` 缓存。已崩溃的任务窗格需要重新加载页面；收起/展开复用窗格不等于重新加载。不要清除聊天历史或重新授权来解决加载问题。WPS 实际恢复仍需宿主验证。

Excel 写侧恢复登记与来源依赖（2026-10-09 第五轮）：`sourceDependencySheets` 由 `apps/excel-addin/src/planSourceDependency.ts` 单独建模，记账**唯一入口**是 `sdkBridgeHost.execute` 里确定性执行器返回之后（取 `DataToolResult.sourceSheets`，不是模型给的参数）；计划自带的值一律被 `attestSourceDependency` 覆盖或剥离，**没有记录就删除字段而不是留空数组**。`officeRecoverySupport` 的读写重叠只看这份依赖：依赖含目标 ⇒ 拒绝；依赖不确定 ⇒ 保留既有保守拒绝（退回 `sourceFingerprintSheets` 口径）。**不要**用"授权范围减去写入目标"反推来源（那是伪造，不是记账）。回归必须包含 `planSourceDependency.test.ts`（8 项）与 `excelWriteEndToEnd.test.tsx`（22 项，"断点恢复登记的真实可达性"7 项、续跑 2 项、审阅期收窄 1 项、快照稳定身份 1 项），以及"规划请求 ↔ 桥回执同一基线"的逐字断言。Excel 侧每轮规划前先读一次现场作为本轮基线（身份 + 指纹 + 范围），并同时喂给规划请求、桥绑定与计划指纹 —— 三个捕获点（`captureWorkbook` / `captureWorkbookCatalog` / `captureWorkbookStructure`）也都要写 `workbookIdentity`，任一处漏写都会让桥在已保存工作簿上恒判 `source_changed`。反向变异 14 项（M1–M14）由 `outputs/_mut17.mjs` + `_mut17.sh` 驱动；**解析 vitest 摘要必须先清 ANSI 颜色码**，否则 14 条全被误判成"未抓住"。自动化证据是宿主替身的，真实 Excel 实机仍未验证。


### 2026-10-11：SDK 完成目标与结果事件调试

- 意图 proceed 回传 planningGoal（read_only / write_preview），两端传入规划；服务端同 turnId 意图记录优先，终局工具不能降级已绑定目标。旧直接调用缺目标仍走兼容语义。
- Python 进度结果必须通过 main.py SSE 字段投影、contracts.ts 共用解析、两端消费三层。TaskProgressEvent.pythonResult 使用严格 LocalResultCard，禁止值矩阵。普通后端与 SDK 均传事件回调。
- 真实模型并发诊断：`.venv\Scripts\python.exe scripts/diagnose_sdk_http.py --model <SDK目录ID> --python --out outputs/sdk-python-audit-new`。使用真实模型额度和 Docker，仅合成宿主，不写实际工作簿；诊断服务仅监听回环，使用独立任务账本，不重启现有服务。请用新输出目录区分每次证据。
- 新 pytest 测试使用仓库已有 asyncio.run 方式；勿假定安装 pytest-asyncio。账本测试必须 monkeypatch 独立临时 FolderOperationLedger，不能读写用户默认数据库。

### 2026-10-12：任务账本损坏排查

WPS 数字文本写入：副本、正式写回、撤销须共用 `wpsWriteMatrix`，临时文本格式后恢复原格式；不能通过放宽严格回读来掩盖数字文本转换。测试与实机边界见 [WPS 文本保真修复](WPS_TEXT_WRITE_FIX_2026-10-12.md)。

SDK 流接口提示不应覆盖服务端 422/500/503。`TASK_STATE_CORRUPT` 表示本地 SQLite 账本损坏，须停止后端、完整备份主库及 WAL/SHM 后离线处理，不能自动删库重试。`scripts/recover_task_ledger.py --source <state.sqlite3> --backup <同目录下的新备份目录>` 会隔离历史规划并完整保留可读执行记录；非规划表不可读即停止。操作影响及本次证据见 [WPS 账本恢复记录](WPS_TASK_LEDGER_RECOVERY_2026-10-12.md)。诊断脚本新增 `--base-url http://127.0.0.1:8765` 可验证实际运行后端及实际账本；会产生合成诊断任务，使用真实模型额度，不写真实工作簿。

## 审阅结束与副本处理（2026-10-15，本轮范围）

- Excel、WPS 审阅卡移除“重新排列”；生成副本时的首次自动排列保留。
- 放弃计划后立即进入终态，旧写回入口失效；确认写回成功后才收尾副本。写回失败或结果不确定时保留副本。
- Excel 兼容路径启用 v4 本机关闭协议。先核对唯一副本与原表身份，再获取 Workbook 对象，直接检查该对象的未保存状态、唯一工作表 nonce 与 XLSX 格式；不在核对后重新获取关闭对象，不使用 WM_CLOSE、ActiveWorkbook、全局 DisplayAlerts 或关闭整个 Application。
- 同一 Workbook 先 SaveCopyAs 到应用数据目录 `review-copies/<随机目录>/review-copy.xlsx`，检查文件为可读 XLSX 后，再复核对象身份并 Close(False)。归档失败、身份不符时不关；关闭失败仍保留已生成的恢复文件。恢复文件包含副本中的额外人工修改，**不会写回原表**。
- 默认恢复目录为 `%LOCALAPPDATA%\Excel Bro\review-copies`；配置 `EXCEL_BRO_CONFIG_DIR` 时位于该目录下。结果卡显示实际路径；文件仅在本机保留，不上传模型、不自动删除，用户确认不需要后自行删除。原表文件和用户自行另存的副本均不删除、不自动关闭。
- 请求新增必填 `closeVersion: 4`，旧服务在校验时拒绝；返回 `status / layoutVersion:4 / copyClosed / recoveryCopyPath`。排列接口仍是 v3。升级后需重启 Python 服务并刷新加载项。
- Excel 原生与 WPS 保留既有内容核对后关闭逻辑：额外修改或读数未知时保留副本；没有生成磁盘文件时不宣称已删除文件。关闭副本不清除原表撤销记录。
- 实现与自动化证据已补齐；本轮未操作真实 Excel/WPS 窗口，自动化通过不表示实机关闭已验收。

最短实机核对：在专用测试工作簿生成副本，分别点“放弃变更”和确认写回；确认只关闭副本、原表保持打开且无原表保存提示，写回后仍可撤销。兼容副本中另改一格后放弃，打开卡片列出的恢复文件核对该修改仍在。另存副本后再放弃应保留已另存的工作簿。归档目录不可写时应保留副本并报告未关闭。

## WPS 工作目录（第三种数据来源）接线（2026-10-16，本轮范围）

- WPS 数据范围面板的「工作目录」卡**不再是禁用的死按钮**，副标题与 Excel 逐字一致（`批量处理多个本地工作簿`）。副标题上那句「WPS 当前仅支持打开的工作簿」是**能力说明**，本轮整条拆掉，**不是**换个说法留着。
- 口径按 Excel 同一套走：选目录（由**本机服务端**弹系统目录框，`POST /api/folders/select`）→ 文件/表**两级勾选** → 承认汇总 → **确认才提交**授权 → 草稿与取消 → 刷新增表不作废、换目录/目标消失即作废。
- 纯判据全在共享层 `packages/core/src/folderScope.ts`（`folderSheetKey` / 可勾选名单 / 存在性名单 / `canPreserveFolderAuthorization` / `folderScopeDraftDirty` / `workDirectorySendRefusal`）。Excel 侧（`utils.ts` / `excel.ts` / `hooks/useScopeSelection.ts`）改成读**同一份**，只留转发 —— 两端不再是"行为相似的副本"。**判据不许再往任一端写第二份。**
- **可勾选名单与存在性名单是两个判据**：前者过滤 EB 系统遗留表（`#EB_RULES` 等），后者**不过滤**。合并成一个会让一个**真存在**的键被判成"目标已消失"，从而误作废授权。
- **工作目录范围下发送侧明确拒绝**：零出站、零工作簿读取；提示说清是哪个目录、几张表，输入框里的字原样保留。**绝不**拿当前打开的 WPS 工作簿顶上。判据 `workDirectorySendRefusal`（只看来源，文案里不许出现工作簿名）。
- 触发器在工作目录模式下**只**说工作目录（`工作目录 <目录名> · 已选 N 张表`），**不显示**当前工作簿名；切回工作簿来源后照旧显示工作簿与当前表。
- **工作目录模式的页脚与动作条按 Excel 的形状改（2026-10-16 深夜）**：WPS 的固定页脚 `.folder-selection-footer` 与 Excel **逐字同构** —— 内部顺序「计数 + ⓘ → 确认」（靠共享样式表里既有的 `.folder-selection-footer .scope-confirm { margin-left: auto }` 把确认贴到右边），承认汇总 `.scope-admission-summary` 落在页脚**下方**。**`.wps-scope-actions`（刷新/确认/取消）只在工作簿来源渲染** —— 工作目录来源下刷新在目录工具条上、确认在页脚里，不再有独立动作条（两端各一份固定页脚 = 一个面板两条固定页脚，那是缺陷）。
- **固定页脚的「层」判据（两端同名同义）**：页脚必须是 `.scope-popover-scroll` 的**兄弟**、直接住在 `.scope-popover` 上 —— 放进滚动容器它就只是"内容末尾的一行"，内容一滚就跟着滚走。工作簿来源下动作条与页脚都满足这条；`wpsScopeDraft.test.tsx` 钉"动作条是滚动容器的兄弟"，`wpsWorkDirectoryScope.test.tsx` 的「固定页脚的层与格式（对齐 Excel）」钉**工作目录那半条**（页脚 + 承认汇总）。**两条判据两个调用点**，不许用一条冒充另一条。
- **两端文件夹页脚到此逐字同形，没有偏离项**：内部顺序「计数 + ⓘ → 确认 →（可选）取消」。Excel 侧原本**只有工作簿动作条**有取消，工作目录页脚没有出口（误勾一张就只能关掉弹层 —— 那是用户**猜**出来的语义）。现已补上：`folderDraftDirty`（草稿 `folderSheetKeys` vs 已提交 `confirmedFolderScopeRef.current.keys`，按共享层 `folderScopeDraftDirty` 比、**根不一致一律按没有已提交范围**）成立时出现，点击 = 草稿键回到已提交授权、已提交一字不改、弹层不关（`cancelFolderScopeEdit`）。判据与回退读**同一个** `confirmedFolderKeys`，免得"按钮因差异出现、点下去却回退到另一份键"。
- **未完成，不得合并宣称**：服务端的查询/执行链路本轮**没有**接到 WPS 面板（发送被明确拦下）；真实 WPS / 真实 Excel 的实机验收未执行。自动化证据是**替身级**的。

判据与反向变异：
- `apps/wps-addin/src/wpsWorkDirectoryScope.test.tsx`（11，授权语义 + 固定页脚的层与格式）、`wpsWorkDirectorySendGate.test.tsx`（3，挂**真实** `WpsAgentPanel`，用计数宿主量读写、记录每个出站请求）、`packages/core/src/folderScope.test.ts`（14）；`wpsScopeDraft.test.tsx` 补齐工作目录来源下的层次判据。
- `node scripts/run_wps_work_directory_mutations.mjs`（W1–W8 + C1 共 9 个变异）：基线 31 + 14 全绿，每个变异**只**红它声明的那几条且失败原文含本处证据，还原逐字节一致（sha256），复跑全绿。**W8 钉的正是"工作目录那半条层次判据"**：它只动工作目录侧（页脚 + 承认汇总），动作条层面保持正确，于是红的恰好是两条 —— `wpsScopeDraft.test.tsx` 的「动作条是滚动容器的兄弟」与「固定页脚的层与格式」。**这条不能用 E2 代替**：E2（`run_scope_action_bar_layer_mutations.mjs`）把动作条本身卷进滚动容器，`actions.parentElement === popover` 先红，工作目录那半句永远轮不到。
- 对照图 `node scripts/render_wps_work_directory_demo.mjs "…/playwright-core"` → `outputs/wps-work-directory-2026-10-16/`。是**夹具页 + 复印的结构**，不是实机截图。

最短实机核对（尚未执行）：
1. 重启 Python 服务、刷新 WPS 加载项，打开一个带宏观配置的工作簿。
2. 点数据范围触发器 → 「工作目录」应为**可点**；点「选择工作目录」应弹出**系统目录框**（服务端弹的）；点取消，范围一个字段都不许变。
3. 选一个含 `.xlsx` 的目录 → 一级列文件、二级列工作表；`#EB_RULES` 这类 EB 遗留表**不应**出现在可勾选列表里。
4. 勾两张表 → 页脚与承认汇总应是「已选 2 张表」，**此时还没授权**；点「确认」后触发器应变成 `工作目录 <目录名> · 已选 2 张表` 且**不带工作簿名**。
5. 在该范围下发一条消息 → 应**发不出去**，出现诚实提示，输入框里的字原样保留，当前工作簿一格都没被读、也没被写。
6. 在别处新建一个工作簿放进该目录后点「刷新」→ 原有 2 张表的授权**不该被动**，新表**不该**被自动勾上。



## 2026-10-18：给 WPS 增加一个"值重排型"动作（以 sortRange 为样板）

新动作按**两道闸 + 一条守卫函数 + 一份反向变异**落地，缺一条都不算完成。

1. **共享层先写判据**（`packages/core/src/tableTransforms.ts` 仿照）。凡是"用户能看出顺序/集合
   差异"的语义（排序顺序、空白位次、重复行保留哪一条），一律放共享层纯函数，两端引用同一份；
   不要在 `excel.ts` / `wpsPlanExecutor.ts` 各写一份比较器。
2. **两道闸一起开**：`WPS_SUPPORTED_ACTION_KEYS`（执行器真的实现了）与
   `config/capabilities.json` 的 `wpsWrite.allowedActions`（对外放开）。两者由
   `wpsPlanExecutor.test.ts` 的 lockstep 用例钉住，只改一处会红。
3. **能力表逐项登记**：`config/capabilities.json` → `hostCapabilities["wps-et"].versions["12.0"]`
   加 `state` 与 `reason`，并在 `docs/HOST_ACTION_CAPABILITY_GAP_2026-10-14.md` 写明
   **属性子集与拒绝条件**。`state:"supported"` 只表示"可进计划"，不等于实机通过。
4. **守卫函数写在适配层**（`assertSortableRange` 的形状）：把"本实现做不到、但宿主能做"的
   场景**在写入前**拒绝，并在错误文案里说清楚"为什么"和"用户可以怎么做"。
   典型：公式的相对引用、格式随行移动、表格对象、宿主缺关键接口（缺接口 ⇒ 按"无法确认"拒绝，
   不当作"没有"）。
5. **执行复用既有链路**：把它折算成"目标矩阵"，而不是新增一条宿主写入原语 ——
   这样写前复核、分块回读、恢复点、撤销冲突判定全部自动成立。
6. **测试要能看见真实状态变化**：替身宿主必须**带网格状态**（`Range(address)` 返回真实子区域、
   `Item(r,c)` 定位真实单元格、写入改变再读到的值），断言"回读到的顺序变了"，
   不能只断言"方法被调用过"。
7. **反向变异证明判别力**：`scripts/check_wps_sort_guards.mjs` 的形状 —— 一次只拆一处守卫，
   红必须是那一条业务断言（含它自己的证据串），还原后逐字节 SHA256 一致再复跑全绿。
   语法错误 / `Failed Suites` / `no tests ran` 一律不算通过。

**结构变更类动作（去重、插入/删除区域）不适用上面这套**：它们会移动下方数据，
恢复必须先有经过实机取证的"行/表插入 + 整块还原"，在此之前**只拒绝、不开白名单**，
也不要用"压实值 + 清空尾行"之类的降级冒充原语义。


## 2026-10-19：删表完整检查点验收

> 2026-10-19（三续）口径变更：结果卡上的「↶ 撤销本次执行」已改为**在原工作簿原地撤销**（不打开副本），"打开执行前副本"降级为「备份与恢复」里的备用入口。下面的实机步骤因此需要按新行为改写，见文末「原地撤销实机验收（待执行）」。

删除工作表采用独立完整归档方案，最初经用户选择为"打开恢复副本、保留当前原表"。与范围动作的原地撤销不同。重启 Python 服务并刷新两端加载项后，只在独立验收文件运行下列步骤：

1. Excel 与 WPS 各建本机 XLSX，包含三张表、公式、业务填充、批注及图表；保存后再做一处未保存编辑。勾选全部三张表，要求“保留 Sheet1，删除其他工作表”。
2. 确认只删除指定两张表，Sheet1 仍在。结果必须有“打开执行前恢复副本”；旧的重复执行按钮不出现。检查点生成失败时一张都不能删。
3. 在当前原表继续编辑一格，点恢复按钮。应打开另一份 XLSX，三张原表及删除前未保存编辑均在，公式/格式/对象人工对照；当前原表保留新编辑、仍只剩 Sheet1。不得关闭、覆盖或自动保存原表。
4. 在恢复副本做修改，再从原窗格打开恢复，第二份仍来自未修改归档；两次恢复文件与原表路径均不同。
5. 仅授权部分表时，完整备份需要额外许可：明确显示未勾选的表名，确认前零备份/零删除；额外许可只用于本地备份。最后可见表也列入删除、未保存/受保护工作簿、混合计划仍在首删前拒绝。
6. 删除途中宿主报错或服务连接断开：显示未完成/待确认，刷新只查询，不自动重删；完整检查点仍可打开。文件夹通道保持原有流程。

目前以上**实机步骤未执行**。归档保存在应用数据目录 `worksheet-checkpoints` 中，不做自动过期清理；恢复窗口中的文件也是独立保留文件。清理策略没有在本轮加入，不能声称已回收磁盘空间。

自动化入口：`test_worksheet_checkpoint.py`、两端面板的 `native worksheet checkpoint routing`、`useWorksheetCheckpoint.test.tsx`、`worksheetCheckpointReceipt.test.ts`；反向验证见 `scripts/check_worksheet_checkpoint_guards.py`。服务端状态持久化与完整备份均使用临时测试目录，测试不触碰活动业务工作簿。


本轮验证记录：core 615 通过；Excel 全量 1332 通过 / 3 跳过；WPS 全量 964 通过 / 3 跳过；服务端全量 1530 通过 / 1 跳过，compileall 通过。最后“响应丢失后查询成功必须向调用方返回成功”的共享 hook 修正，两端参数化用例与 Excel 面板链路 33 条再次通过。两端生产构建含 TypeScript 检查通过。C1–C4 反向变异均命中业务断言并逐字节还原。日志在 `outputs/checkpoint-*.log`；没有执行本节的真实宿主步骤，没有操作业务工作簿、提交或推送。


### 隐藏表备份许可专项验收（2026-10-19）

重启 Python 服务并刷新两端加载项，在独立本机 XLSX 中准备可见 Sheet1–Sheet3 和非常隐藏的测试表（可使用 `#EB_RULES`、`#EB_RULES_BACKUP` 复现范围面板过滤）。

1. 勾选三张可见表，要求保留 Sheet1、删除 Sheet2/Sheet3。应显示“请确认本地备份范围”与两张隐藏表的准确名称，尚无归档或删除。
2. 点“暂不删除”：表保持不变；重新发起后点“允许仅本地备份并继续”，只删除两张目标表，隐藏表保留且未加入分析授权。
3. 打开恢复副本，检查可见表、隐藏状态、单元格内容/对象与删除前一致；原表和随后编辑保留。
4. 等待许可时切换文件、更改选表或改动工作表列表，再确认必须阻断。刷新后旧卡只读，不能继续旧删除。
5. Excel 同样验证；窄窗格中按钮换行、长表名不溢出，旧无归档失败记录折叠。尚未做真实宿主和浏览器视觉验收。

针对性自动化：服务端隐藏表需许可/精确名单/名单变化/额外许可不可授权删除；共享 hook 的实时授权、重载、暂缓；两端真实面板的确认入口（宿主与网络替身）。检查点反向守卫 C1 为“不经同意便备份”，新增 C5 为“等待时表列表变化仍执行”。


本次隐藏表许可修复验证：core 615 通过；Excel 1340 通过 / 3 跳过；WPS 965 通过 / 3 跳过；两端 build（含 tsc）退出码 0；服务端 1538 通过 / 1 跳过，compileall 通过。C1–C5 共 5 项反向守卫全部命中业务断言，还原 SHA256 一致。日志 `outputs/checkpoint-consent-*.log`。服务端真实 ZIP 备份测试包含 veryHidden 状态；两端面板用宿主/网络替身，不能据此宣称 COM 实机删除已通过。业务工作簿仅读取过表名/可见性用于诊断，没有备份、删除、保存、关闭；未提交或推送。


### 原地撤销实机验收（2026-10-19，**未执行**）

重启 Python 服务并刷新两端加载项，**只在专用测试工作簿**（例：`444.xlsx`）里做，不碰业务文件。

1. 建本机 XLSX：Sheet1…Sheet9，Sheet2/3/4 有公式与格式、一张很隐藏的表、一条引用被删表的工作簿级名称、一张表里有图表与批注。勾选全部表，要求"保留 Sheet1、Sheet 5–9，删除 Sheet2/3/4"。
2. 删除完成后，**保持原窗口不动**，在 Sheet1 的空白格写一句话（模拟"删除后的编辑"）。
3. 点结果卡右下角「↶ 撤销本次执行」。预期：Sheet2/3/4 回到**原位置**（Sheet5 之前）；当前工作簿仍是 `444.xlsx`、没有新窗口、没有弹保存提示；第 2 步写的那句话还在。
4. 核对完整性：公式原文（含跨表引用）回来、`#REF!` 已被修回、名称定义指向原区域、隐藏状态与行高列宽/图表/批注一致。结果卡写"已撤销本次删除 · 已恢复 3 张工作表"，明细里有"已放回"，按钮收起。
5. 撤销后在表里新建一张同名表再撤销另一条更早的记录：必须判冲突、一张都不动、不说"已撤销"，并指路「备份与恢复 → 打开执行前副本」。
6. 撤销刚点下去就再点一次（或用两个窗口），不允许发出第二次恢复请求、不允许出现重复表。
7. 在撤销期间切换工作簿 / 改授权：不得误写、不得只写一半还报成功；失败时保留检查点与重试入口。
8. 「备份与恢复 → 打开执行前副本」仍可用，且不会被说成"撤销成功"；原有单元格撤销、结果审阅链路不受影响。
9. Excel 与 WPS 各做一遍；两端结果卡文案、按钮位置、明细顺序一致。

自动化入口：`server/tests/test_worksheet_restore.py`（11 条）、共享判据 `packages/core/src/worksheetRestore.test.ts`、`apps/excel-addin/src/hooks/useWorksheetCheckpoint.test.tsx`、`apps/excel-addin/src/hooks/useAuthorizedTarget.test.tsx`、`apps/excel-addin/src/worksheetCheckpointReceipt.test.ts`、两端 `native worksheet checkpoint routing`。反向变异：`node scripts/mutate_worksheet_restore.mjs R1..R8`，驱动 `node scripts/run_worksheet_restore_mutations.mjs`（四判据 + 还原 sha256 一致，日志落 `.tmp-mutations/`）。

**替身证据 ≠ 实机证据**：以上 9 步尚未在真实 Excel / WPS 上执行，COM 侧"整表复制是否保留全部对象""活动表与滚动位置是否被影响"都还没有实机结论。


### 2026-10-19：Excel 身份读取与重复范围确认回归

- 后端检查点原生读取：最多 3 次、间隔 150ms，配置见 `config/capabilities.json` 的 `worksheetCheckpoint.nativeReadAttempts/nativeReadRetryDelayMs`。只重试元数据读取，不重复删除或恢复。路径/窗口身份明确变化仍直接拒绝。
- 范围确认：选择当前已打开的文件后提交范围应一次完成，不能再弹“数据范围变化”。真实替换工作簿、外部删表、未确认草稿仍须拦截。
- 本轮串行全量：Excel 1376 passed/3 skipped；WPS 965 passed/3 skipped；两端 build（含 tsc）成功；服务端 1561 passed/1 skipped，compileall 成功。日志在 `outputs/excel-identity-fix-*-20261019.log`。
- 红→绿：新增范围测试先复现空授权名单和 scope_changed；原生读取测试先复现将 None 误报身份变化。新增测试涵盖同一保留对象、暂时失败、持续失败、真实身份变化、窗口复用、重复候选、不得重试写入。
- 实机只读：旧代码在同一 Excel 工作簿上复现 NativeOM E_FAIL；仅重试无间隔仍出现读不到对象，未当作完成。最终同对象读取+有间隔重试后，连续三次读取 `ss.xlsx` 身份成功。此证据不代表真实删除/原地撤销已通过。
- 生效：用户自行重启 Python 后端并刷新 Excel 加载项。在专用测试簿里选范围→删除→撤销，核对没有重复范围确认、表恢复到原簿、后续编辑保留；不要拿旧失败操作的“查询执行状态”当成重新执行（查询只返回旧回执）。本轮未重启用户服务、未写业务簿、未提交/推送/发布。

### 数据范围授权同步实机验收（2026-10-19，**未执行**）

自动化只证明了替身宿主上的链路。下面 9 步必须在真实 Excel 与真实 WPS 上各做一遍，
并把"是否还弹范围提示"逐条记下来（弹了就是没通过，不要用"可能是别的提示"含糊过去）。

1. 在**专用测试簿**（三张表：A、B、C）里打开加载项，默认「跟随当前」，点「数据范围」→「确认」→
   发一句需要读表的话。应正常执行，不弹「工作表范围发生变化」。
2. 「数据范围」→「选择多个」→ 勾 A、B →「确认」→ 发一句需求。执行后**等 10 秒**（越过 2s 轮询）再发
   第二句：两句都不应要求再确认。
3. 接上一步，在 Excel 里点 C 的标签（切换活动表）→ 等 10 秒 → 再发一句：**不应**弹提示
   （固定范围与活动表无关）；若弹了，记录当时「原范围 / 当前」两行文案。
4. 接上一步，打开「数据范围」→ 点「跟随当前」→ 点「取消」：触发器应退回「已固定 2 个工作表」，
   再发一句不弹提示。（旧行为的症状：触发器变成「当前表 C」并弹提示。）
5. 打开「数据范围」→ 点「跟随当前」→ **直接点触发器关掉弹层**（不点取消也不点确认）→ 等 10 秒：
   后台不应弹提示；此时发一句，应停在「请先确认当前对话的数据范围」，输入框内容保留、未读字段。
6. 删一张表（B）→ 结果卡正常 → 等 10 秒 → 再发一句：不弹提示；「数据范围」里已提交范围应已收窄到 A、C。
7. 点「↶ 撤销本次执行」→ B 回到原位置 → 等 10 秒 → 再发一句：不弹提示，授权仍是 A、C（不自动新增 B）。
8. 制造真实变化：在 Excel 里手动删掉 A，或把 A 改名，或**换一个工作簿**（同名文件也算）→ 发一句：
   **必须**被拦下（「工作表范围发生变化」/「当前文件与已授权文件不同」），且不读字段、不写入。
9. 慢机器上重复第 2、6 步各三次：观察是否出现"刚确认完又要求确认"。出现即记录当时的
   ①是否点过取消 ②是否点过「跟随当前」 ③提示里"原范围 / 当前"分别是什么。

自动化入口：`apps/excel-addin/src/excelWriteEndToEnd.test.tsx`（`数据范围授权同步` 一组：
①③④⑥⑦⑧⑨）、`apps/excel-addin/src/folderScopeSafety.test.tsx`（②）、
`apps/excel-addin/src/hooks/useAuthorizedTarget.test.tsx`（`读数在飞期间基线被推进` 两条）、
`apps/excel-addin/src/hooks/useScopeSelection.test.ts`（`重读宿主只刷新草稿`）。
反向变异：`node scripts/mutate_scope_authorization.mjs A1..A6` + 驱动
`node scripts/run_scope_authorization_mutations.mjs`（四判据 + 还原 sha256 一致，日志落 `.tmp-mutations/`）。

**替身证据 ≠ 实机证据**：以上 9 步尚未在真实 Excel / WPS 上执行；真实宿主的读数延迟与 2s 轮询
叠加后，"切表是否提示"是否仍显得频繁没有实机结论。

### 已确认多表范围实机验收（2026-10-19，**未执行**）

本轮修的是「已确认多表范围，发送任务仍显示『暂时无法核实当前目标』」。自动化只证明了
替身宿主上的链路；下面 6 步必须在**真实 Excel 与真实 WPS** 上各做一遍。
不操作业务工作簿：准备一个专用测试簿，含三张表，**目录顺序与你要点的顺序不同**
（例如目录 A、B、C，而活动表是 C，点选顺序 C → B → A）。

1. 打开加载项 →「数据范围」→「选择多个」→ 按 **C → B → A** 勾三张 →「确认」→
   立刻发一句需要读表的话。应**正常执行**，不出现「暂时无法核实当前目标」。
2. 接上一步，等 10 秒（越过 2s 轮询）再发第二句：仍不提示。
3. 换一组顺序（B → A → C）重新确认一次，再发一句：仍不提示。
4. 制造真实变化：外部删掉一张已授权表 / 把一张改名 / 换一个工作簿（同名文件也算）→
   发一句：**必须**被拦下（「工作表范围发生变化」或「当前文件与已授权文件不同」，
   或证据缺失时的「暂时无法核实当前目标」），且不读字段、不写入。
   这一步是"规范形没有把真实变化也吞掉"的实机判据。
5. WPS 侧重复 1–4（两端口径必须一致）。
6. 若仍复现，先记录**本地脱敏元数据**再报问题：文件身份是否存在、`document` 引用是否
   仍是同一个对象、证据是缺失（`null`）还是冲突（两个不同的键）、授权版本是否变化。
   **不得**记录单元格内容，也不得把"读不出来"统一吞成 `null` 后猜原因。

自动化入口：`apps/excel-addin/src/excelWriteEndToEnd.test.tsx`
（`A.按非目录顺序确认三张表后…`）、`apps/excel-addin/src/officeAuthorizedTarget.test.ts`
（证据键与顺序无关一条、换文件/换 `document` 一条）、
`apps/excel-addin/src/hooks/useAuthorizedTarget.test.tsx`（`已确认范围的证据冲突判据` 两条）、
`apps/wps-addin/src/wpsClarificationTargetConfirmation.test.ts`（顺序无关一条）、
`packages/core/src/authorizedTarget.test.ts`（规范形六条）。
反向变异：`node scripts/mutate_scope_identity.mjs I1..I4` +
`node scripts/run_scope_identity_mutations.mjs`（四判据 + 还原 sha256 一致）。

**替身证据 ≠ 实机证据**：以上 6 步尚未在真实 Excel / WPS 上执行。

### 刷新后不重复确认范围实机验收（2026-11-01，**未执行**）

修的是「已固定勾选、未改勾选，重新加载 WPS 插件后发送任务仍要求确认数据范围」。
自动化用**替身宿主**证明了两条分支（文档级证据成立 ⇒ 不再询问；证据不足 ⇒ 仍确认一次），
但真实 WPS 是否真的每次返回新包装对象、`CodeName`/`SId` 是否可读，只能实机看。
用专用测试簿（含来源 + 结果1/2/3 四张表），不操作业务工作簿。

1. 「数据范围」→ 勾四张 →「确认」→ 发一句读表任务：正常执行，不弹确认。
2. **只重载加载项**（不改勾选）→ 发同一类任务：**应不再弹「请确认本次使用的数据范围」**，
   直接进入分析。这一步是本次修复的实机判据。
3. 若仍弹确认：**按标题精确对号**（四种标题各自唯一对应一个 `notice.reason`，
   `TargetScopeNotice` 是唯一出处；「请确认本次使用的数据范围」只可能是
   `confirmation_required`）。
   - 「请确认本次使用的数据范围」⇒ `confirmation_required`：元数据一致、三份证据键也一致，
     但**都是弱证据** —— 即 `conclusive=false` 且 `documentBound=false`。**已经实测到过一次**
     （见下），意味着这台宿主没给稳定 ID；
   - 「暂时无法核实当前目标」⇒ `unavailable`：证据缺失/三键冲突/读不出来，不是"变了"；
   - 「工作表范围发生变化」⇒ `scope_changed`：元数据比对失败；
   - 「当前文件与已授权文件不同」⇒ `file_changed`：真的换了文件（同名不同路径也算）。
   **不得**为了让它通过而关守卫或强制 `conclusive`。
3b. 要拿到布尔值而不是标题猜测：DEV 构建下每条判定都会落一条诊断
   （`wpsTargetDiagnostics.ts`）。回收方式二选一：
   - 窗格 console（若可见）：`[wps-target] {...}`；
   - 磁盘：`%APPDATA%/kingsoft/wps/addons/data/win-x64/cef/2/jsapi/cache/Local Storage/leveldb`
     里的 `excel-bro.wps.target-diag.v1`（只有布尔值/原因码，不含路径、表名、证据键）。
   要看到 `branch / conclusive / documentBound / hasConsent / explicitRequired /
   stableId / pathIdentity / round / seq`，这几项就够判"新请求又被拦"还是"旧提示没清"：
   同一 `round` 且 `seq` 不再增长 = 旧提示；`seq` 增长而 `round` 变大 = 新的一轮。
4. 收窄/扩一张表 → 重载 → 发任务：**必须**要求确认（范围真的变了），且说的是范围变化，
   不是"暂时无法核实"。
5. 换一个工作簿（含同名不同目录）→ 重载 → 发任务：**必须**拦下。
6. 真做一次「同名工作表替换」（删掉一张已授权表、新建同名表）→ 发任务：允许它要求确认，
   但**不得**宣称"数据已变化"；记录实际结论。
7. Excel 侧重复 1–2（Excel 应当一直不弹；若弹了就是另一个问题）。

自动化入口：`apps/wps-addin/src/wpsAuthorizedTarget.test.ts`（文档级/对象级/缺稳定 ID/
未保存/顺序无关）、`apps/excel-addin/src/hooks/useAuthorizedTarget.test.tsx`
（`文档级证据成立时刷新后直接分析`、`只有对象级弱证据时刷新后仍确认一次`、
`文档级证据也不能掩盖身份键冲突`、`requireExplicitConsent 只挡文档级捷径` 三条、
`范围恢复原样后旧提示会被清掉`、`诊断只有布尔值与原因码`）、
`apps/wps-addin/src/WpsAgentPanel.test.tsx`（`刷新后范围未变、宿主换了包装对象`、
`宿主给不出工作表稳定 ID`）、`apps/wps-addin/src/wpsTargetDiagnostics.test.ts`。
反向核对：摘掉共享闸里的 `documentBound` 放行条件，面板用例立刻变红（卡在"应按原授权
直接进入分析"），还原后 sha256 一致。

**已执行到的程度（2026-11-01，只到第 1–3 步的一半）**：用户重载窗格后发任务，第 2 步的
预期**没有**出现 —— 弹的是「请确认本次使用的数据范围」（= `confirmation_required`）。
按 3b 回收到了那条记录：`branch=weak_evidence`、`conclusive=false`、`documentBound=false`、
**`stableId=none`**、`pathIdentity=true`、`hasConsent=false`、`explicitRequired=false`。
即**这台 ET 宿主没有报出工作表稳定 ID**，路径身份正常，对象级证明也不成立。
因此第 4–7 步（范围变化/换簿/同名表替换/Excel 对照）**未执行**。
修复前不得再宣称"刷新后与 Excel 一致"；要让刷新后不再询问，只能放宽那条证据、
由用户拍板（见 `docs/ARCHITECTURE.md` 同日章节末）。

**最短真机步骤（2026-11-01，交互闭环后；只跑这一遍）**

1. 在 WPS 里打开那份四表工作簿，重载加载项（任务窗格重开）。
2. 发一句读表任务（例如"按来源表算三步公式"）。
3. 预期：提示卡写「窗格已重新加载，请确认继续使用已选的 4 张工作表。」+「工作表勾选没有变动」，
   并有主按钮「确认范围并继续」；**不要**去点「检查数据范围」。
4. 点「确认范围并继续」→ 只继续这一次，原需求与澄清答案保留；**不要**取消勾选再勾回。
5. 再发一句同类任务 → 不再询问（同会话范围未变）。
6. 负向：把其中一张表改名或删一张 → 提示变成「工作表范围发生变化」，**且没有**
   「确认范围并继续」（原样确认不能替"范围变了"背书）。

只要第 3–6 步里有一步不符，就把当时那条 `excel-bro.wps.target-diag.v1` 留着（我按
`branch`/`round`/`seq` 判读）；不需要反复刷新碰运气。

**替身证据 ≠ 实机证据**：以上 7 步尚未在真实 WPS / Excel 上全部执行。

**最短真机步骤（2026-11-01 另一起：只读查询被误判。与上面那条范围确认问题分开跑）**

准备：工作簿里**四张表都加入授权范围**（例如 `来源`、`结果1`、`结果2`、`结果3`），
**四张表的 B 列都放数字**（各 100 之类），`来源` 的 `A1:B5` 是没有表头的五对值（B 列 1~5），
再在 `来源!B6` 写 100 作为区域外诱饵。

1. 发送：「只读取工作簿里的来源表 A1:B5，按无表头处理，计算 B 列五个数的合计，
   并告诉我参与计算的记录数。不修改任何工作表。」
2. 预期：合计 **15**、记录数 **5**；**不是** 315（15+100×3，说明范围没收窄）。
3. 预期：结果卡的「来源：」后面**只有「来源」**，没有「结果1/2/3」；
   **不出现**任何「某某没有字段」的提示；**不出现**「结果不完整」「扫描已达上限」。
4. 预期：**不出现**任何写入计划/写入预览/目标单元格，表格一格未变。
5. 负向（有表头那一侧不许被改坏）：另起一张表首行写「数量」，下面三行写 1、2、3，
   问合计 → 必须回 **6**（表头行不算数据）。
6. 负向（显式声明压过按列名推断）：另起一张表首行真的就写「列1」「列2」，下面两行数据，
   问笔数 → 必须回 **2**。
7. 负向（不许把别的表混进来）：问「四张表里一共有多少条记录」→ 这次是**真跨表**，
   模型必须把四张表都列进 `sheets`，结果卡的来源才该出现四张表。
8. WPS 与 Excel 各跑 1–7。任一步不符，把当时的本地结果卡（含 `scannedRows`）与
   工具调用参数留下；参数现在会落在进度事件的 `bridgeTool.bindings` 里
   （模式/scope/表名/区域/表头口径），不需要再靠 `sourceSheets` 反推。

只需注意：文件夹模式没有按区域裁剪这一步，`range` 在那里会被**明确拒绝**（不是静默忽略）；
正确定位是"改用工作簿模式"，而不是改成写入去凑范围。

### 第一批 SDK 业务闭环实机验收（2026-10-19，**未执行**）

自动化已完成：真实模型 + SDK + Docker 的结果核对，替身宿主上的执行/回读/确认/撤销，
以及跨进程工件回放。下面这组必须在**真实 Excel 与真实 WPS** 各做一遍；使用专用合成
测试簿，不操作业务工作簿。

准备：一个工作簿，三张表——`会员表`（含空编号、重复编号、前导零）、`订单表`（含找不到
会员的订单）、`汇总表`（空白）；三张都加入授权范围。期望值只用于你核对，不填进任何
提示词。

1. 发送任务：按会员编号左连接，去首尾空格、保留前导零，空编号不参与匹配，重复编号
   保留首次出现；按等级汇总订单金额，未匹配归入"未匹配"；输出汇总、总计与可定位的
   空编号/重复编号/未匹配明细到汇总表。
2. 观察：**不点任何"执行"按钮**也应在汇总表写一次并回读；确认前两张来源表一格未变。
3. 核对：金卡/银卡/未匹配金额与总计；三类异常明细都带来源表与行号；空编号明细不能
   写成"共 0 条"。
4. 审阅：出现变更标记；**可选**打开双视图——不打开也必须能完成任务；打开时显示的是
   执行前状态，且不重新执行。
5. 点「确认变更」：只清除本次标记，**写入次数不增加**；出现「保存为工具」入口。
6. 点「↶ 撤销本次执行」：汇总表逐格回到执行前；来源表不变；不跳转到恢复副本。
7. 因果变化：只改会员表里键 7 第一次出现的等级（金卡→银卡），重新发送；金卡与银卡
   金额应随之变化，异常定位不变。
8. 反向场景：授权收窄（外部删一张已授权表）、换工作簿（同名文件也算）、等待期间改来源
   值、取消本轮、取消后迟到的结果——都不得产生额外写入。
9. WPS 侧重复 1–8；涉及删表时沿用既有「原地撤销」验收步骤，不另设流程。

若失败，记录脱敏元数据（任务 id、计划 id、结果引用、表名、行号、错误码），**不要**记录
单元格内容或密钥。自动化对应入口：`apps/excel-addin/src/onlineLink.test.tsx`、
`apps/wps-addin/src/wpsOnlineLink.test.tsx`、`apps/excel-addin/src/crossProcessChain.test.tsx`。

### 2026-10-20：桥准备及检查点故障验证

自动化入口：

- `wpsReadonlyAgent.test.ts`、`wpsPlanningPauseCard.test.tsx`：意图已有任务身份后，准备桥未结束时
  不发规划；准备失败、抛错或取消均无规划/备用请求；真实面板接线使用同一判据。
- `useSdkToolBridge.test.tsx`、`sdkIntentRouting.test.tsx`：初始化授权不可读时不发请求；
  运行期间现场不可读仍拒绝读表，不泄露宿主异常详情。
- `server/tests/test_checkpoint_persistence.py`：短暂/持续替换失败、复制成功后异常、
  回执丢失后重载、重复请求、故障前零写入、API 错误及状态回读。工作簿全部为合成替身。

实机仍待验：仅使用专用合成 XLSX，分别在 Excel/WPS 验证首次 SDK 读表、在途断开后核对、
继续分析与旧轮次迟到事件；核查实际请求时序，不能仅凭按钮文字判断。
持久化故障须在独立测试环境注入，不对业务工作簿制造文件锁。
持续失败后应先查询状态；未知的已出现工作表不得靠反复点撤销来“补齐”。
已有执行前副本可用于人工核对，主撤销流程不自动切换到副本。

本轮自动化结果（2026-10-20）：core 688；Excel 1411（3 skipped）；WPS 977（3 skipped）；
服务端 1642（1 skipped）；两端构建（含 tsc）、compileall 均通过。改动文件 eslint
0 error / 10 warning，diff --check 通过。日志为 `outputs/codex-bridge-persistence-20261020-*`。
`scripts/verify_bridge_persistence_faults.py` 的 P1–P4 均先基线通过、再因目标业务断言失败、
还原 SHA256 一致、复跑通过；日志在 `outputs/bridge-persistence-mutations/`。
真实 SDK、真实 Excel/WPS、cross-process/online 本轮未运行。

### 2026-10-20：暂停后输入“继续”的回归验收

更新后自行重启后端并刷新 Excel/WPS 加载项，前后端须同时使用新消息理解协议。
仅在专用测试工作簿上执行：

1. 发起分析，在规划期间暂停接收。保留原任务卡，不重新发送原需求。
2. 输入普通聊天，再输入“继续”。应查询原任务 ID，不创建第二个规划任务；
   原任务仍运行时，原卡保留“核对任务状态”入口。
3. 服务端核实有可续现场后出现“继续分析”。只有点击后才发继续请求，且重新核实
   工作簿、授权与来源。状态不明时不重试写入、不自动重新规划。
4. 在暂停后发送另一个问题，旧卡仍能核对旧任务；切换会话不能继承其他会话任务。
5. 重开 WPS 窗格后重复第 2–3 步；缺少本机上下文时只查询，不猜测原目标。
6. 换簿、改授权、改来源的反向场景仍须拒绝。恢复完成后回读、审阅确认及撤销沿用
   原流程；不要以出现“继续分析”按钮作为完整通过证据。

自动化入口：`messageFirstRouting.test.tsx`、`messageUnderstanding.test.ts`、
`wpsPlanningPauseCard.test.tsx`、`wpsPlanningResumeClosure.test.tsx`、
`planningRecoveryRouting.test.ts` 与 `server/tests/test_message_understanding.py`。
日志前缀 `outputs/recovery-routing-*`。本批没有操作真实工作簿，没有重启用户后端；
真实模型语义路由及双端中断续跑仍待按以上步骤验收。

本批验证：core 690；Excel 全量 1415（3 skipped），最后的归属收尾调整后定向
35 passed；WPS 最终全量 982（3 skipped）；服务端 1647（1 skipped）。
两端 build（含 tsc）、compileall、diff --check 通过；改动 TS 文件 eslint
0 error / 17 个 hook 依赖警告。没有运行真实 SDK、在线集成或双端实机故障注入。

### 2026-10-20：早期暂停缺口（用户实机反馈后补修）

此前日志不能证明所有阶段都可继续。新增针对“理解消息 / 读取字段 / 结合字段理解任务”
暂停的验收：显示“需求分析已暂停”和“继续分析”，说明重新分析且尚无规划检查点；
输入“继续”或点击按钮应把原始完整需求送回分析，不能把字面“继续”作为新的业务需求，
不能调用不存在的规划任务状态接口，也不能重复追加原需求气泡。

双端反向场景：核对时换簿、改授权或口径、切会话、双击、旧意图响应在取消后迟到。
已有规划身份的任务仍必须走原状态核对/检查点链，不能因没拿到结果就重分析。
本机保留只在当前窗格会话有效，附件和授权对象不新增落盘；文件夹模式不启用该入口的
自动重分析。只在专用测试工作簿复测，用户自行重启后端并刷新加载项。

自动化入口新增 `hooks/usePrePlanningPause.test.tsx`，面板用例仍在
`messageFirstRouting.test.tsx` 和 `wpsPlanningPauseCard.test.tsx`；迟到响应反例还覆盖
`wpsReadonlyAgent.test.ts`。日志前缀 `outputs/early-pause-*`。
`early-pause-real-sdk.log` 是使用已配置 SDK 兼容连接、当前代码和合成文字的独立验证：
返回 `task_recovery / preplanning-synthetic-early-pause`。最初用默认非 SDK 兼容连接
验证被 `ModelNotAvailableError` 拒绝，保留在 `early-pause-real-model.log`；没有修改
默认模型、重启服务或操作工作簿。真实 Excel/WPS 的截图所示操作仍需用户复测。

本批结果：Excel 全量 1425 passed / 3 skipped，最后增加关闭窗格的保护后共享 Hook
定向 7 passed；WPS 全量 985 passed / 3 skipped；服务端 1647 passed / 1 skipped。
两端最终 build（含 tsc）、compileall、diff --check 通过；eslint 0 error / 18 warning。
首次 WPS 全量唯一失败是旧用例仍等待“尚未形成可续跑检查点”的旧提示，更新为
“原需求入口存在但没有服务端检查点卡，且没有自动调用规划”后通过，旧失败日志保留。

### 2026-10-20：恢复后仅两次“理解消息”的修复验收

日志前缀 `outputs/restart-handoff-*`。在修复前，双端测试改为“凡提供暂停候选就返回
该候选”，Excel 三条、WPS 两条均在“原需求必须进入后续分析”断言失败。修复后覆盖
按钮/自然语言两种入口及旧渲染闭包读取实时候选：定向 Excel 41、WPS 25 通过；
全量 Excel 1430 passed / 3 skipped，WPS 987 passed / 3 skipped，两端 build 含
类型检查通过。本次没有修改服务端、协议或能力开关，未重启后端、未操作工作簿。

`restart-handoff-real-sdk.log` 使用已配置 SDK 兼容连接与合成文字连续调用真实
消息理解：输入“继续”选择暂停 ID，然后原完整提示不带候选，返回需要工作簿上下文。
这比单次路由验证多覆盖一次交接，但仍不是宿主业务链验收。

实机复测：刷新加载项（刷新前的本机早期暂停引用不会保留），在专用测试簿重新发起
完整需求，于消息理解或字段判断阶段暂停；分别测试输入“继续”和点击“继续分析”。
应继续读取授权字段、产生分析结果，不应停在两条已完成的消息理解时间线。再核对
换簿、授权变化时仍拒绝接回。两端实机复测待执行，本次前端更新无需重启后端。

### 2026-10-21：输入框暂停/继续验收

本批涉及 `task_recovery.action` 协议：重启后端，刷新 Excel/WPS 加载项后在专用
测试工作簿验证。暂停后默认只留一行“已暂停”，可展开详情；空输入时右下角按钮
应启用并标注“继续任务”。分别点击按钮、按空输入 Enter、发送“继续”，确认原任务
能进入后续分析；自动状态核对不得自行续跑。新文字应正常发送，不默认恢复旧任务。
仅询问“只查看状态，不要继续”不得发继续请求。附件存在时按钮走新消息发送。
换簿、改授权、来源变化或无法核实仍需停止，不能为了简化提示放行。

日志前缀 `outputs/pause-composer-*`。新增/更新双端面板测试覆盖早期暂停接回、
服务端任务核对后继续、WPS 续跑桥实际读表与回执、旧空发送契约变更以及协议
拒绝未知 action。`pause-composer-real-sdk.log` 是真实 SDK 合成文字证据：
“继续”返回 continue，“只查看状态，不要继续”返回 inspect；不是实机验收。

本批自动化结果：Excel 1433 passed / 3 skipped；WPS 最终 988 passed / 3 skipped；
服务端 1651 passed / 1 skipped；双端 build（含 tsc）、compileall、改动文件 eslint
及 diff --check 通过（eslint 保留原 hooks 警告）。首次 WPS 全量的一个失败来自旧
用例“空输入提交不得继续”；按本次用户要求更新为“只继续原任务一次、不重发需求、
不写入工作簿”后全量通过，首次日志保留。真实宿主未自动操作，实机仍待复测。

### 2026-10-21：续跑保留步骤编号验收

WPS 面板用例先复现续跑仍显示独立“任务过程”而未接回原时间线，日志
`outputs/resume-timeline-before.log`；修复后第 7 行原地更新、下一事件为第 8 行，
连续暂停与输入继续均不增加过程卡。Excel Hook 验证七行接续、前六行不变及排除
暂停时间，真实 App 组件验证继续前后卡片和步骤数量不增加。模型输出在组件测试
中为替身，包含“接着刚才暂停的任务往下做”的继续意图，不代表真实模型命中率。

双端全量日志 `outputs/resume-timeline-{excel,wps}-all.log`；后续定向复跑与构建
使用同前缀。仅前端变更，未改服务端协议及能力配置，未操作业务工作簿。
实机待验：刷新加载项，在专用测试簿的第 7 步暂停，通过空输入按钮或表达继续的
自然语言恢复，核对原第 7 行接续、后续追加第 8 行，再暂停继续仍在同一卡内。
遇到真实授权或来源变化仍应拦截；无检查点重分析不得宣称精确断点续跑。

### 2026-10-21：任意分析阶段接续与精简暂停 UI

早期三阶段（消息理解、读表结构、字段分析）分别验证文字继续和按钮继续：恢复原需求、
保留一个过程卡、完成步骤不重复、后续步骤顺序追加。共享计时用例覆盖第 1/2/4/7/10 行，
第 7 行仅是示例，不是恢复规则。组件测试核对暂停动作在原标题栏，不再存在灰色说明卡。
日志：`outputs/early-timeline-*`；Excel 全量 1439 passed / 3 skipped，WPS 最终全量
994 passed / 3 skipped，新增界面与定向复跑 69 passed。首次 WPS 全量有一个旧文案断言失败，
更新为实际入口和不伪造规划任务的判据后通过；不恢复用户已要求删除的文案。

实机待验：刷新两端加载项，在专用测试簿分别于上述阶段暂停，再用空发送或自然语言继续；
核对原卡、原编号、当前阶段完成后追加下一阶段，连续暂停仍然如此；身份/授权变化仍拦截。
本轮未操作真实宿主。写入与回读阶段目前不能暂停，安全动作边界暂停仍属未实现项。

### 2026-10-21：WPS 暂停后授权误拦回归

`outputs/planning-consent-before.log`：同一范围已有确认后被包装对象误拦、未知身份
被误报变化，两条业务断言修复前失败。新增面板用例覆盖已确认继续、未确认先阻止再
确认后继续、同名异路径拒绝、稳定表 ID 变化拒绝；记录模型规划只发一次及零宿主写入。
最终替身每次 ActiveWorkbook / Worksheets.Item 读取都返回新包装对象，验证不依赖
偶然稳定的 JS 引用。原等待状态查询期间换簿反向用例仍通过。

WPS 全量 998 passed / 3 skipped；Excel 共享授权与消息入口回归 82 passed。
日志前缀 `outputs/planning-consent-*`。首次构建检出新增测试漏传 scopeDock，已补全并
重新构建；生产实现只改 WPS 面板。未改后端及能力配置，未操作真实工作簿。
实机待验：刷新 WPS 加载项，确认测试簿三表，发新需求并暂停，通过空发送或文字继续；
同一已确认范围应继续原任务，真实改选范围/换簿仍拦截。未知身份应保留任务，不报失败。

### 2026-10-21：共享授权在临时读取失败后保持有效

修复前 `scope-consent-retention-before.log` 的两条用例分别复现元数据读失败、分析证据
持续缺失耗尽重试后原确认键被清除。正向验证原范围恢复可读后重新检查沿用原授权、
不再次 adopt；反向验证换文件、增选表、证据冲突均撤销确认，返回旧范围也不能自动
复用已撤销的确认。全量回归：Excel 1446 passed / 3 skipped；WPS 998 passed / 3 skipped。

新增重新检查入口时补查写入边界：`scope-consent-write-replay-before.log` 证明原 recheck
会将 check 的任意待定回调登记成分析重试；已改为先校验既有分析回调归属，再核验。
后续定向回归、两端构建和 lint 日志为 `outputs/scope-consent-final-*`。无服务端改动，
未操作真实宿主；实机需核对暂时不可读后恢复时无需重新勾选，真正范围变化仍暂停。


### 2026-10-21：勾选即授权验收

本轮前端行为覆盖旧的“确认才授权、取消回退”测试约定。自动化覆盖：勾选自动提交、收起重开保留、
清空即撤销、连续勾选旧核验失效、未完成选择阻止后台误报、同范围请求不重复确认、真实换簿/删表阻止、
目录会话重建复用本窗格授权。日志前缀 outputs/immediate-scope-*。没有修改后端、能力配置或重启服务。

实机待验（未执行）：刷新 Excel / WPS 加载项，在专用测试工作簿中依次进行以下操作。

1. 选择多个并勾选三张表，不点任何确认，收起再展开；勾选与固定张数保持一致。
2. 不改变文件和勾选，静置及切换活动表，再发需求、暂停并继续；不重复要求确认范围。
3. 连续勾选/取消不同表，最终授权以最后一次勾选为准；清空后不能读取活动表作为兜底。
4. 临时读取失败时保留选择，刷新后能继续；真正换文件、删除/改名已选表仍停止。
5. 工作目录勾选自动生效、清空撤销，刷新同根同表保留；更换根目录不沿用旧授权。

勾选仅改变授权，不自动恢复旧任务或执行旧计划。现有写回、回读、审阅确认与撤销流程保持原约定。

本轮验证：Excel 全量 1454 passed / 3 skipped，WPS 全量 988 passed / 3 skipped；
授权守卫与选择调度定向 60 passed，最后的选择页脚与样式复核 52 passed。
两端最终构建（含 tsc --noEmit）均 exit 0；改动文件 eslint 为 0 error / 25 warning，
git diff --check 通过。具体日志：`outputs/immediate-scope-{addin,wps}-final.log`、
`outputs/immediate-scope-build-{addin,wps}-final.log`、`outputs/immediate-scope-eslint-final.log`。
本轮只修改前端，未重跑服务端测试；以上自动化证据不代替真实 Excel / WPS 验收。

界面补充验收：顶部固定显示“勾选即授权”，底部“确认”只关闭面板，勾选仍即时生效；工作簿和工作目录
共用此语义。相关既有用例按新布局更新后 Excel 49 passed、WPS 24 passed；两端构建通过。
日志前缀 `outputs/scope-picker-layout-*`。真实宿主未复验，需刷新加载项核对窄窗格、长列表与目录页脚。

### 2026-10-22：双端暂停/恢复补充验收

本轮新增的双端面板用例让服务端任务必须等到真实本地适配器生成成功读表回执后才能完成：
覆盖继续后新到达的工具指针、暂停前已收到但未完成回执的指针，断言 result 存在且无 errorCode，
并确认工具不写工作表、不重新发送规划需求。网络与 Office.js/WPS 宿主为替身，面板、工具桥、
本地适配器和共享状态逻辑为真实实现。该证据不等于真实宿主验收。

另补充：WPS 消息理解阶段暂停/继续及完成即收尾；WPS 继续详情与按钮同源；查询期间换文件
两端均拒绝接回；步骤终态补发不重新计时；暂停/等待时间分开显示；任务进行中不会因最近一步
完成而提前显示完成。WPS“点击继续时后台已完成”新增用例先失败于“应取回已完成结果”，修复后
走同一结果链。证据日志 `outputs/codex-resume-completed-before.log`。

实机待验（本轮未操作原生窗格）：在 Excel 与 WPS 各用独立测试簿和新会话，勾选相同范围后，
分别于消息理解、字段读取、规划、等待本地查询时暂停。用空发送或自然语言继续，核对原卡、
原编号、已完成步骤耗时冻结、本地查询最终有结果；后台已完成应直接取回结果。连点继续只发起
一次恢复；再次暂停后不能自动重新接回。同文件同勾选正常恢复，查询期间换文件/撤销授权则停止。
当前写入与回读期间禁止中断，不能将上述分析阶段证据当作写入阶段可中断的验收。

本轮最终检查：

| 检查 | 结果 | 证据 |
| --- | --- | --- |
| Excel 全量 | 1465 passed，3 skipped | `outputs/codex-resume-addin-final.log` |
| WPS 全量 | 1000 passed，3 skipped | `outputs/codex-resume-wps-final.log` |
| 共享层全量 | 708 passed | `outputs/codex-resume-core-full.log` |
| 服务端全量 | 1661 passed，1 skipped；compileall exit 0 | `outputs/codex-resume-server.log` |
| 两端最终构建（含 tsc --noEmit） | exit 0 | `outputs/codex-resume-build-{addin,wps}-final.log` |
| ESLint / 补丁格式 | 0 error，37 项现存依赖警告；diff --check exit 0 | `outputs/codex-resume-lint-final.log` |

最终全量后，构建检出新 Excel 用例响应未补齐协议字段，以及 WPS 提前处理 completed 后留下的
不可达比较。补全字段、删除不可达比较后，两端重新构建通过；受影响的面板恢复用例复跑
Excel 39 passed、WPS 33 passed（`outputs/codex-resume-{addin,wps}-edge.log`）。没有用类型断言
绕过这两处错误。本轮没有修改服务端实现、协议或能力配置，没有重启服务或操作业务工作簿。

| 本轮能力 | Excel | WPS |
| --- | --- | --- |
| 运行中恢复本机工具回执、复用原任务和时间线 | 已实现，组件及适配器回归通过，实机待验 | 已实现，组件及适配器回归通过，实机待验 |
| 任务完成取回结果、状态/耗时展示 | 已实现，自动化通过，实机待验 | 已实现，自动化通过，实机待验 |
| 写入/回读阶段中断、跨窗格早期恢复、SDK 会话持久化 | 未实现，本轮未扩展 | 未实现，本轮未扩展 |


### 2026-10-22：额度不足提醒验收

| 能力 | Excel | WPS |
| --- | --- | --- |
| 已知额度/上下文/输出/预算/轮数限制的原因与建议 | 已实现，共享 API 与面板组件回归通过，实机待验 | 已实现，共享 API 与面板组件回归通过，实机待验 |
| 普通 429 不误报余额不足；明确额度不足不自动重试 | 共用服务端与前端分类，自动化通过 | 共用服务端与前端分类，自动化通过 |
| 故障后的写入状态 | 本地结果卡为准，不从模型失败推断 | 本地结果卡为准，不从模型失败推断 |

后端验证覆盖：HTTP 402、429 insufficient_quota / billing_hard_limit_reached、普通限流、
上下文超限、SDK billing_error / ResultError / 输出截断 / 预算及轮数限制、错误后成功重试、
意图 JSON 与规划 SSE 传递、明确额度不足只调用供应商一次。原有工具能力提示只在普通响应
失败或轮数耗尽时优先；不能覆盖额度、上下文、输出或预算错误。

两端真实组件验证：错误原因与处理建议显示于对话，原输入保留，消息理解失败时不进入规划或
宿主写入；提示不泄露原始供应商错误。没有调用真实付费 API 制造欠费，也未进行原生窗格验收。
加载新版本需要重启本地后端，并刷新 Excel/WPS 加载项。实机复验应使用独立测试工作簿，分别
验证额度不足、限流和有效连接下的正常执行；已有执行记录不能因后续模型报错而被改写为失败。

本轮最终回归结果（2026-10-22）：

| 检查 | 结果 | 日志 |
| --- | --- | --- |
| Excel 全量 | 1468 passed，3 skipped | `outputs/codex-usage-excel-full.log` |
| WPS 全量 | 1003 passed，3 skipped | `outputs/codex-usage-wps-full.log` |
| 共享层全量 | 716 passed | `outputs/codex-usage-core.log` |
| 服务端最终全量 | 1691 passed，1 skipped | `outputs/codex-usage-server-final.log` |
| 两端构建及 TypeScript | exit 0 | `outputs/codex-usage-build-{excel,wps}.log` |
| ESLint | 0 error，37 条既有 warning | `outputs/codex-usage-lint.log` |
| compileall / diff --check | exit 0 | 本地命令验证 |

补充覆盖：SDK 已报告 billing_error 后发生进程退出或缺少最终帧，仍显示额度不足；未识别到
供应商限制证据的本机进程故障继续保留原故障提示。本轮未重启用户运行中的服务或宿主。

### 2026-10-25：来源锚点回归与门控宿主

SDK 固定值回归需模拟真实前端桥绑定/受控读取，不能直接返回计划后塞一个自称有效的
指纹。无桥、错误 taskId、来源变化测试必须关闭普通成功用例的桥便捷适配，独立控制
事件时序。WPS 来源读取替身要支持 UsedRange.Item 与 Range(firstCell,lastCell)。
默认套件不执行 A1-E2 工件消费用例；更改门控用例、桥或来源复核后必须单独运行
`npm run test:a1-e2`。该命令消费已留存工件，不等于本轮重新执行真实模型或原生宿主。
本轮普通/门控结果与日志见 `SDK_A1_E2_RECOVERY_CHAIN_2026-10-25.md` §9；原生步骤见 §7。


2026-10-26 S2 收尾：stop 的创建时间按 UTC ticks 比较，兼容 JSON 小数尾零正规化与 DateTime 解析；仍校验路径、完整命令和专用启动模块。脚本回归见 test_a1e3_native_stop.py；原生记录见 SDK_A1_S2_NATIVE_RECOVERY_2026-10-26.md。

### 2026-10-27：账本 schema 升级与"旧库缺列"恢复

**症状**：恢复包 `PUT /api/execution-recovery/{id}` 返回 500，浏览器显示 `Failed to fetch`，
服务端原文 `sqlite3.OperationalError: no such column: recovery_package_expires_at`；随后 suspend
200、`GET` 404。客户端表现是"这次执行的恢复材料存不下来"，重开只能看到"需要重新规划"。

**根因**：`_schema_is_current` 只比 `PRAGMA user_version`。版本号被写到 10 而列没建时（旧库/
迁移中途失败），检查此后永久跳过补列。异常又穿出路由、响应缺 CORS 头，被浏览器折成网络错误。

**升级（用户侧最短操作）**：只重启服务进程，不重建账本、不删 WAL、不重发任务。

```
# 停掉当前 dev:server，然后在同一工作区：
npm run dev:server
```

新进程第一次实例化账本时补列并把 `user_version` 写到 10；已存在的任务、`next_action_index`、
`checkpoint`、`before_snapshot` 与已登记的首次期限全部保留。**不要**用"删库重建"或"清空恢复
索引"来绕过 —— 那会把"这次执行做到哪一步"一起丢掉。

**排查顺序**（下次再遇到同一族问题时）：先看服务端日志有没有 `OperationalError`，再看
`/api/diagnostics` 的 schema；浏览器只报 `Failed to fetch` 时不要从网络层查起。

**相关纪律**：加列必须同时改 `_schema_is_current` 的判据（读 `PRAGMA table_info`），
迁移用守卫式 `ALTER TABLE`；schema 已是最新时不许重跑 bootstrap（会抢写锁，见
`ARCHITECTURE.md`「版本号是意图，列是事实」）。回归见
`test_task_ledger.py::test_migration_keeps_action_progress_and_checkpoint`、
`test_migration_keeps_an_existing_deadline_when_only_the_version_lags` 与既有
`test_read_is_not_serialized_behind_concurrent_write_lock`；客户端反向对照见
`WpsAgentPanel.test.tsx` 的四条（变异脚本 `scripts/mutate_wps_recovery_clear_gate.mjs` M6–M10）。

本轮没有重启用户运行中的服务、没有调用其 API、没有改用户的库文件，也没有打开真实 Excel/WPS。
**A1-GATE 仍未通过，A2 未启动。** 证据见 `docs/SDK_WPS_RECOVERY_PACKAGE_2026-10-27.md` §21–§27
与 `outputs/recovery-package-20261027/r4-*.log`。


## 2026-11-02：审阅输入框最短检查

Excel/WPS 已执行并回读后，输入框内显示「已完成，待你审阅」，上方无重复审阅卡、无空 textarea 或发送按钮；范围条保留。「查看详情」默认折叠；处理失败原因不折叠。确认只接受结果，放弃走原恢复，两者成功后普通输入恢复；标记清理失败保留独立重试。澄清或中文组合输入不被卸载。

本轮自动化：Excel E2E/共享组件 120 passed；WPS 面板 244 passed、5 skipped；双端 build（含 tsc）通过。布局探针 scripts/check_recovery_composer_layout.mjs 覆盖 320/400/500px 审阅、长错误正文及既有恢复模式共 18 档，按钮完整可见。日志 outputs/review-composer-*.log。未操作真实工作簿、账本、材料，未提交推送。原生 Excel/WPS 各检查一次审阅布局和确认/放弃即可；不要重复整轮 B 恢复验收。


## 2026-11-02：逐表双视图对比验证

双端多表执行完成后，审阅输入框提供「对比工作表」。选择结果1、结果2、结果3，分别看该表执行前快照和当前结果；切换不新增原表写入。确认/放弃仍针对完整任务。副本有人工修改、关闭不明或身份变化时，保持现有保护并给原因。

自动化：核心全量 1162 passed；Excel E2E/判据/组件 111 passed；WPS 面板+执行器 291 passed、5 skipped，原生副本/前态重建 55 passed。用不同表相同地址不同旧值检测串表；测试确认不重写、关闭失败保留旧句柄、排列失败后仍可收尾。两端构建含 tsc 通过，320/400/500px 浏览器布局通过。真实宿主窗口排列未验，不宣称实机通过。日志 outputs/multi-review-*.log。
## 2026-11-02：审阅摘要与长表名切换

Excel/WPS 共用 ResultReviewBar：输入区摘要为“待审阅”，有明确对比表清单时显示表数，不从任务标题推算单元格数。任务描述进入折叠详情；错误和无法对比原因保持可见。确认/放弃按钮明确作用于本次全部修改。

ReviewSheetPicker 提供前后切表、当前位置、可搜索列表；长名称在常驻行省略，列表完整换行。列表在正文内展开，独立限高滚动，不覆盖底部操作。切表仍调用既有 comparison 请求，不确认或重复执行计划。原生 Excel/WPS 长表名、多表及键盘操作仍待实机检查。

## 2026-11-05：底部结构 / 反馈归属的检查点

结构契约（读源码，双端同一条）：`apps/excel-addin/src/composerDockParity.test.ts` —— `.composer-dock` 只许有 `scope-entry-row` 与 `composer-input-region` 两个子节点且相邻，`TargetScopeNotice` / `WorkDirectoryGuide` / `HistoryStorageNotice` / `ComposerActionDock` / `RecoveryPendingEntry` / `TaskExecutionNotice` / `composer-service-status` / `composer-pending` / 恢复探针一律不许回到两者之间；同时断言 `.message-stream` 仍 `min-height: 0; overflow-y: auto`、`.chat-shell` 仍是 `minmax(0, 1fr)`，以及短系统反馈的局部底色选择器。

真实 DOM（默认态）：`WpsAgentPanel.test.tsx` 的「底部固定容器只放工具栏与输入框，通知不夹在两者之间」。变异自检：在 `<div className="composer-dock">` 后插入任意带 `composer-service-status` 的元素，该用例必须变红（本轮已实测）。

判据单测：`packages/core/src/operationFeedback.test.ts`（按 executionId 归属、不同执行不合并、找不到执行保留原行、反馈行不能当 anchor、副本按身份去重、相同错误文本不同副本算两份、摘要不冒充恢复失败）、`packages/core/src/composerPending.test.ts`（决定类默认展开、纯状态默认收起、超两条只报数量）、`packages/core/src/resultReview.test.ts` 的「展示副本清理失败」四条（多入口重复触发只更新同一份清单、真关掉后清空、业务结论仍为 confirmed）。

界面组件：`components/ComposerPendingBar.test.tsx`、`components/ExecutionFeedbackSummary.test.tsx`。`<details>` 的展开由 React 状态驱动并对 `summary` 的 click `preventDefault()` —— jsdom 不实现原生展开，只靠 `onToggle` 既测不到也控不住。

运行：
`npm run test:core`、`npm run test:addin`、`npm run test:wps`、`npm run build:addin`、`npm run build:wps`。
本轮结果见下面「本轮最终数字」（这一节写完后又在真实 Chromium 上补了布局证据并修掉一处让位方向错误，数字以最终那次为准）。
已知非本轮失败（记录时与最终状态的差别）：`excelWriteEndToEnd.test.tsx` 那两条断言旧文案「公式采用当时的值」的用例，在最终那次全量里已经通过（并行改动把实现与断言对齐了）；`wpsBeforeComparison.test.ts` 一条仍期望「授权快照」而实现抛「对比目标不属于本次执行。」，仍是并行改动尚未收口，未在本轮修 —— 修它需要动那条判据的措辞，属于另一处判据的所有权。
另：`wpsContinuousPythonChain.test.tsx` 的 `./wpsNativeComparison` 替身缺 `saveWpsComparisonCopies` 等新出口，会在 `ResultReview.closeBefore()` 抛未处理拒绝；本轮补齐替身存根（不影响任何断言）。

### 布局证据（真实 Chromium，不是 jsdom）

`node scripts/check_recovery_composer_layout.mjs`（Excel 与 `PROBE_HOST=wps` 各一次）：320x640 / 400x720 / 500x760 三档下，审阅动作行与主操作完整落在视口内、详情折叠、长错误正文自己滚；`reconcile` / `resume` / `long` / `entry` 四态同样通过。

`node scripts/check_scope_entry_layout.mjs` 本轮改了两件事，否则它量到的不是线上结构：

1. **样式清单补齐**：原先只拼 `styles.css + chatScrollbars + composerPrimaryButton.css`，**没有 `conversation.css`** —— 底部两层结构全在那个文件里，等于没测。现按真实导入顺序拆成 `base`（styles + 组件 css）与 `late`（chatScrollbars + appearance + conversation），WPS 的 `styles.css` 仍插在 `base` 之后、`late` 之前；夹具页顶部加 `*{scroll-behavior:auto !important}`（否则量滚动位置量的是动画中间态）。
2. **夹具换成新结构**：`<footer class="composer"><ComposerPendingBar>…</ComposerPendingBar><div class="composer-dock">工具栏 + 输入区</div></footer>`，消息区塞 24 条消息 + `#last-row`。

新增断言（两端 × 三档宽度）：`.composer-dock` 恰好两个子节点且顺序为 `scope-entry-row` → `composer-input-region`；两者**紧贴**（`rowBottom` 与 `regionTop` 差 ≤ 1px，即中间没夹东西）；`.composer-pending` 整个在输入容器上方；发送按钮完整在视口内、页面无横向溢出；**展开长待办详情后**详情自己滚、摘要行仍在视口内、发送按钮仍完整可见；消息区可滚且滚到底后最后一条不被底部盖住。

变异自检（本轮实跑）：把 `.chat-shell .composer-dock > .scope-entry-row { margin-bottom: 0 }` 改成 `12px`，探针报「工具栏与输入框之间不许夹别的东西: … rowBottom:514, regionTop:526」；还原后 sha256 与变异前一致。

另一条反向证据来自真实缺陷：给 `.chat-shell .composer` 加 `min-height: 0`，`check_recovery_composer_layout.mjs` 在 320x640 直接报「审阅动作行 bottom=728.9 > 640」。已写进 `composerDockParity.test.ts` 的「让位的是常驻摘要，不是底部容器」（断言 `.composer` 规则不含 `min-height: 0`、`.composer-pending` 是 `flex: 0 1 auto`、`.composer-dock` 是 `flex: 0 0 auto`）。

### 本轮最终数字

`npm run test:core` 1261 passed / 81 files；`npx vitest run`（excel-addin）**1844 passed / 9 skipped / 0 failed**；`npx vitest run`（wps-addin）1331 passed / 8 skipped / **1 failed**；`npm run build:addin`、`npm run build:wps` 均 EXIT=0；`npx tsc --noEmit` WPS 0 错，Excel 1 错在 `src/excelUndoScope.test.ts(435)`（该文件 22:18 仍在被并行改动，不在本轮范围）。

WPS 那 1 条是既有文案漂移：`wpsBeforeComparison.test.ts:298` 期望抛错含「授权快照」，实现抛「对比目标不属于本次执行。」。上一轮记录的两条 Excel 旧文案（「公式采用当时的值」）已随并行改动一并消失，本轮全量 0 失败。

另：跑测试期间 `apps/excel-addin/src/excel.ts` 与 `excelWriteEndToEnd.test.tsx` 在 22:13:16 被并行改写，一度让「公式续跑完成后进入审阅收尾（discard）」报「恢复后回读不一致」而失败；文件稳定后单跑与全量均通过。**并行会话在写同一棵树时，中途失败要按 mtime 复核，别当成自己的回归。**

实机未验：两端真实窗格里的底部间距、窄/宽窗格与缩放、多行草稿、刷新后恢复状态、正式待审阅的确认/放弃/双视图入口位置，以及本轮全部反馈归属（已撤销 / 已放弃并恢复 / 副本清理失败）与待办摘要展开，均只到「真实组件 + 真实服务端 + 替身宿主」、源码/DOM 契约与真实 Chromium 布局，未在原生 Excel/WPS 窗格点过。
