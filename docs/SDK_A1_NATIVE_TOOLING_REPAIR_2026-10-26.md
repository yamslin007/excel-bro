# S1 验收工具修复与真实窗格记录（2026-10-26）

范围：修复第十一轮验收工具和测试隔离，随后在独立测试簿做真实S1。用户已授权computer-use。
未增加执行前确认；没有修改生产models/contracts或宿主执行器；未操作8765；未提交Git。

## 修复

- 裸arm：交互终端明确选择；非交互缺筛选非零退出，不假装已启用。
- cancel发JSON；消费后用专用审计保留实际绑定任务，接受生产taskBound=false的真实语义。
- evidence按run/attempt/task/宿主/簿身份/时间关联，拒绝缺材料、取消、未消费与错归属；
  只查该任务执行记录，403/500/网络错误不伪装成另一宿主404。每次导出独立文件，不覆盖。
- 只有完整+明确zero+已知单位才称零调用。没有执行记录还必须有同任务SOURCE_CHANGED失败收尾。
- 每轮新目录、绝对脚本路径、Base/Port一致；stop比较PID/创建时间/可执行文件/命令，拒停无关进程。
- 独立配置与账本；SDK目录在仓库外。服务端pytest收集前设置独立环境，已知污染测试逐条换临时账本。

## 自动化

最初三条回归全部红：取消无JSON、partial+zero称零、缺账本被当未进入执行。修复后通过。
现专项30条通过，覆盖真实fault返回值与真实SQLite、消费后的绑定证据、多候选、取消、过期、
错任务/错身份/错宿主/旧任务、接口错误、错误进程PID、目录与夹具覆盖防护。

全量：server第一轮1838/1 skip；Excel首轮并行有3条5000ms超时，单独全量复跑1560/9 skip；
WPS1130/8 skip，双端build通过。完整日志保留 outputs/codex-native-*-20261026*.log。
最终新增4条后专项30通过；最终服务端全量1842 passed/1 skipped（120.03s），compileall通过。
core845、lint 0 error/36既有warning；A1-E2门控Excel6/WPS188均通过，消费已有真实工件，
本轮未重新声称产生Docker过期链证据。双端build、PowerShell AST解析、git diff --check及
三份新脚本的no-index --check均退出0。

| 检查 | 日志 |
| --- | --- |
| 服务端最终1842/1 skip | outputs/codex-native-server-final-20261026.log |
| Excel独立全量1560/9 skip | outputs/codex-native-addin-20261026-solo.log |
| Excel首轮3超时（保留） | outputs/codex-native-addin-20261026.log |
| WPS1130/8 skip | outputs/codex-native-wps-20261026.log |
| 双端构建 | outputs/codex-native-addin-build-20261026.log / codex-native-wps-build-20261026.log |
| core / lint / A1-E2 | outputs/codex-native-core-20261026.log / codex-native-lint-20261026.log / codex-native-a1e2-20261026.log |


## 真实 Excel S1

- 测试实例：run=20261026T101759-747dff27，8877，使用已有DS/deepseek-v4-flash SDK连接。
- 工作簿：先前独立run 20261026T082204-9c000735/a1e3-验收-副本.xlsx（修正SDK目录后继续使用；
  不是业务工作簿）。固定授权仅来源、目标。
- taskId=turn-7f578003becd4dd7a12811c1486ec15b；绑定和触发完全相同，账本failed/SOURCE_CHANGED，
  executions=[]，记录为「未进入执行」，没有编造host-writes零观测。
- 原生界面观察：执行卡失败；出现“来源已变化，当前任务已停止：没有继续读取来源，没有继续计算，
  也没有写入。请按最新数据重新发起需求。”；未出现继续按钮。来源仍显示111/250/333，目标B2
  公式栏为-101。随后通过真实UI保存副本，独立openpyxl只读核对完整来源111/250/333与目标-101/-202/-303均不变；不以仅看B2代替整组核对。
- JSON：outputs/a1e3-native/runs/20261026T101759-747dff27/evidence/20261026T023435-b28d0c66/evidence.json。
  CLI把nativeUiAndWorkbook留为pending；本节记录独立UI观察，避免自动脚本替人判过。
- 首次SDK目录误放仓库内导致模型不可用，已修且补回归。第一次export被沙箱拒绝，重获文件写权限后成功。

## 真实 WPS S1

- 同一8877实例，WPS打开当前run内独立a1e3-验收-副本.xlsx，仅勾选来源/目标。
- taskId=turn-b8c9effb11d74b2a8589faf054a1b49e；绑定=消费，账本failed/SOURCE_CHANGED，
  executions=[]，未进入执行。JSON：该run/evidence/20261026T025656-2317b90e/evidence.json。
- 真实窗格出现同一原文停止提示、执行步骤失败，没有继续入口；固定授权仍两张表。
- 通过UI保存测试副本，独立只读核对来源111/250/333与目标-101/-202/-303全部不变。
- 首次打开新簿时加载项保留旧未保存簿的范围保护，通过“检查数据范围”明确重新选择新测试簿
  后解除；未绕过守卫。后端重启后的旧模型目录通过“管理模型连接”只读刷新，未改密钥。

两端原生截图、保存后簿与独立观察在该run的native/excel-office、native/wps-et及
native/observations.json。evidence.json仍保留nativeUiAndWorkbook=pending，独立观察记录在本节
和observations.json，不改写自动导出结果。截图为Windows真实窗格，通过computer-use采集。
S2/S3本轮未重跑，真实手动改源检测未替代。
A1-GATE未通过，A2未开始。


## 原生证据文件

- [Excel停止截图](../outputs/a1e3-native/runs/20261026T101759-747dff27/native/excel-office/stopped.png)
- [WPS停止截图](../outputs/a1e3-native/runs/20261026T101759-747dff27/native/wps-et/stopped.png)
- [双端保存后数值与UI观察](../outputs/a1e3-native/runs/20261026T101759-747dff27/native/observations.json)

本轮未收集S1宿主写调用上报：原因是规划阶段已经终止，没有executionId；记录“未进入执行”，
不把空host-writes当作complete+zero。本轮没有验证Docker分析（故障在分析前触发）。


## 环境收尾

8877专用服务按manifest核实进程身份后已停止。两套验收前端只在祖先进程确为本轮
`a1e3_native.ps1 frontend`时停止，随后按原默认配置恢复3000/3889 Vite，均已就绪。
未启动或停止8765；测试簿、截图、日志、隔离账本与仓库外SDK临时目录保留。
已有工作区其他改动未回退，未提交/推送Git。

本轮完成项是“第十一轮工具修复 + 双端S1注入停止原生验收”。下一步仍是S2原生精确中断/
双击只续第三步、S3确认/放弃后重开不重执行；真正来源变化检测另验。A1-GATE不据本轮单项改为通过。
