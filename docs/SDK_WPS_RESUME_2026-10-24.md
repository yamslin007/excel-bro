# A1：WPS 受限值写入续跑（2026-10-24）

> **2026-10-27 更新（材料来源已换）**：本文件里"沿用 PluginStorage 撤销记录作为续跑材料"的做法
> 在**完整退出宿主**这条路径上不成立 —— 真机重开后那里是空的，界面提示「找不到原执行的本机恢复记录」。
> 现在不可变材料存本机后端应用数据目录，`PluginStorage` 只作缓存/索引；窗口期一分为二
> （仅缓存 1800s / 持久材料 7 天）。续跑判定本身未改（仍是 `prepareWpsResume`），所以下文关于
> 来源证据、动作范围、在途核对、审阅接回的描述继续有效。见
> [本机恢复包](SDK_WPS_RECOVERY_PACKAGE_2026-10-27.md)。

## 当前范围

| 项目 | Excel | WPS |
| --- | --- | --- |
| 已登记动作恢复与原身份认领 | 既有实现，本批回归 | 已实现 |
| 完整来源证据、来源目标分离、非重叠值写入的剩余动作 | 既有实现 | 本批已实现 |
| 来源快照缺失的 SDK Python 结果计划 | 按既有来源守卫判断 | 本批未实现，不宣称能续跑 |
| 格式/公式/结构动作跨窗格精确续跑 | 受既有能力限制 | 未实现 |
| 原生宿主恢复验收 | 待验 | 待验 |

A1 尚未整体完成。用户此前对 Excel 双视图、确认、放弃、撤销的实机反馈不能替代恢复验收。

## 实现

- `wpsResumeContext` 从经过可信来源装配的 prepared 提取来源定位、形状、逻辑指纹和独立值指纹。缺少快照或 presentationFingerprint 来源不登记续跑材料；后者涉及显示格式依赖，不能只核对值冒充完整证据。
- 同步 contracts/models 严格协议：`resumeContext` 可缺省，兼容历史记录；未知字段与矩阵拒绝。服务端校验来源授权、尺寸、限额和来源目标分离。
- 既有 PluginStorage 撤销记录实际已有目标值和原始写前快照（不是仅存在内存）。本批复用它，新增 taskId 与不含单元格值的来源描述，不新建矩阵存储。原创建时间不刷新；`taskRecovery.wpsResumeMaxAgeSeconds=1800`，来源总格数限额 250000。到期只失去续跑资格，不擅自删除撤销点；既有 undo.maxEntries 淘汰仍生效。
- `prepareWpsResume` 在重建 whitelist plan 前校验归属、原始快照与登记指纹、来源、授权、宿主版本及游标。重新从可信坐标构造几何信息，拒绝直接执行本地记录的任意字段。指纹沿用既有 FNV 逻辑核对用途，不是密码学认证。
- `resumeWpsExecution` 只 claim 原身份，不调用模型或重新注册；领取后复核，动作写前/写后提交账本，每次等待后及值分块前重验来源与授权。沿用现有执行器的目标、已写区域、只读、保护、合并、回读守卫。
- 在途动作先单独核对 applied/not_applied，unknown/conflict 保持阻断。重复认领由服务端 revision/owner 控制，响应不明不自动重试。
- 原始整份计划的撤销快照跨多次中断保留。续跑后接入 ResultReview，确认不重写；放弃恢复整份计划。双视图按原快照重建，沿用现有宿主限制。

## 自动化证据

- 新执行器故障注入：原三动作第一/二步后丢回执，模拟重开，仅续写后缀；整份撤销、重复中断、原始时间不刷新。
- 反例：来源、已完成目标、剩余目标、授权变化、过期、taskId 错位、本机目标矩阵被改、历史记录无材料、在途未核对、领取冲突；均不续写。
- 领取与写前 checkpoint 等待期间改变来源，未产生新的宿主写入。
- 真实面板两条：点「继续剩余动作」→仅第三步写入→确认不重写 / 放弃恢复三步；模型和新登记调用数为零。
- 服务端真实 HTTP/SQLite：来源描述往返不含值；拒绝原始矩阵、越权、源目标重叠、错误形状、不支持动作、非指纹字符串与超限。
- 日志：`outputs/a1-wps-resume-*.log`。本轮最终结果如下。

| 检查 | 结果 |
| --- | --- |
| core | 790 passed |
| Excel 全量 | 1518 passed / 3 skipped |
| WPS 最终全量 | 1042 passed / 3 skipped（较基线新增 19 条） |
| server 全量 | 1744 passed / 1 skipped，130.28s |
| 双端 build（含 tsc） | 退出码 0 |
| lint | 0 error / 37 既有 warning |
| compileall / diff-check | 退出码 0 |

最终日志：`a1-wps-resume-addin-full.log`、`a1-wps-resume-wps-verified.log`、`a1-wps-resume-server-full.log`、`a1-wps-resume-build-addin.log`、`a1-wps-resume-build-verified.log`、`a1-wps-resume-core.log`、`a1-wps-resume-lint.log`。均在 outputs/。

中间失败未隐去：新用例最初把单格实际写入地址 A3 误写为 A3:A3；两条旧面板用例仍断言旧文案；后补 copyRange 用例误用了不存在的 targetStartCell（运行测试和 tsc 均检出）。已分别修正，最终全量及构建通过。早期服务端定向测试有 pytest cache 写权限警告；最终全量用 no:cacheprovider，独立 basetemp。新增逻辑本批未做反向变异，不把旧 A1 服务端变异结果算成本批新守卫证据。

## 实机步骤（未执行）

在独立 WPS 测试簿中选来源及目标两表，执行带来源证据的三动作值写入；在第二步写后回执前中断。重开后先核对 applied，再点继续，核实第三步只写一次、前两步不重写。分别验确认、放弃、撤销和可选双视图；改变来源/授权、等待过期及双窗格并发认领均不得新增错误写入。Excel 复跑对应恢复链，不把旧审阅验收当本轮恢复验收。

后端与 WPS 面板需加载本批版本才识别 resumeContext。没有修改默认 Python 开关，没有开启原生 SDK 会话持久化。
