# 有序常量写入与无来源依赖恢复（2026-10-27）

用户的 WPS恢复验收.xlsx 截图：目标 B2/B3/B4 保持 -101/-202/-303，计划明确写入
111/222/333，仍报“没有记录它是按哪一版数据算出来的”。这是写前拒绝，尚未进入第二步中断。

## 根因与实现

原可信 plan_range_edit 只能提交单区域、统一标量；共享内存封印也限制一个动作。
三步不同常量无法使用该可信入口。新增 plan_literal_edits，保留 edits 顺序且不合并动作，
逐项复用单区域编译器的表名/A1/公式注入/能力校验，整批受总动作和总格数限制。
SDK 指令明确引导多步用户常量走该工具，计算结果仍走 submit_result_plan；原始 JSON
不能伪造 literal_range_edit 来源。共享封印绑定全部动作、目标、数值及顺序。

新增恢复链回归进一步发现：没有来源表的可信常量计划拿不到 wpsResumeContext，
因而没有持久恢复包。WPS 增加来源证据 v2（纯常量、无来源依赖，sources 必须为空）；
v1 继续要求非空来源。只允许带活体内存封印的纯 writeValues 生成 v2。
服务端 models、客户端 contracts、共享恢复包校验同步；TS/Python 共读同一份
sourceEvidenceCases，拒绝 v1 空来源、v2 非空来源及未知版本。

恢复前仍比对记录与账本来源证据、计划/动作/值指纹、授权、宿主版本、目标现场、
在途动作及执行权。通过完整校验后才为重建的常量计划重新建立封印。不以序列化 JSON
上的字段直接取得来源豁免；计算计划缺锚点仍拒绝。确认只清标记，放弃恢复原始快照。

## 证据与局限

先红：共享层有序常量封印用例失败；服务端新增工具相关 10 failed / 43 passed，
失败包含 unknown_plan_compiler 与真实 SDK handler 中缺少工具的 KeyError。
新增 WPS 恢复链初次失败在 kind=none（没有持久材料）；补来源证据 v2 后通过。

测试修正如实记录：Excel 组件测试最初漏掉服务端本来会自动添加的 rangeEquals，
补成真实响应形状并在服务端断言这些条件；WPS 最初选用只接受矩阵的旧替身，
改用现有支持单格标量的 bridgeRoundTripHost。服务端第一轮默认 basetemp 触发环境
清理限制，改独立 OS 临时目录。环境错误未当成产品失败证据。

专项及全量结果记录在 outputs/literal-batch-*.log；最终结果见文末验证表。
本轮证据为生产编译器/执行器/真实面板组件及宿主或 SDK 替身。未运行真实模型、Docker、
真实 WPS/Excel；不能把新增恢复链写成原生已验收。未提交或推送 Git。

## 用户操作

1. 重启当前使用的后端，重新加载 WPS 的 Excel Bro 窗格。新版 dist 已构建；旧页面需重载。
2. 打开原测试簿，确认目标表 B2/B3/B4 仍为 -101/-202/-303，固定授权目标表及来源表。
3. 展开“恢复验收（开发专用）”，启用下一次第二步中断。
4. 重新发需求：“按顺序执行三个独立的常量写入：目标表 B2 写入 111，B3 写入 222，
   B4 写入 333。不要合并动作，其他单元格保持不变。直接执行后展示结果供我审阅。”
5. 预期停在 111/222/-303。先保存工作簿并保留截图，再按原完整退出验收流程继续；
   重开后须先核对第二步，再续写第三步，不能自动跳过核对。

旧失败计划不会被自动升级；没有恢复材料的旧执行也不能凭空补出快照。
A1-GATE 保持未通过，A2 未启动。

## 最终验证

| 检查 | 实跑结果 | 日志 |
|---|---|---|
| Excel 全量 | 1566 passed / 9 skipped | literal-batch-addin-full-final.log |
| WPS 全量（含常量恢复及原始三步撤销） | 1148 passed / 8 skipped | literal-batch-wps-full-final.log |
| core 全量 | 866 passed | literal-batch-core-full-final.log |
| 服务端全量 | 1901 passed / 1 skipped | literal-batch-server-full-final.log |
| 服务端最后新增 v2 HTTP 往返后，三个相关文件复跑 | 161 passed | literal-batch-final-api.log |
| Excel / WPS build（均含 tsc） | exit 0 / 0 | literal-batch-build-addin.log / literal-batch-build-wps.log |
| lint | 0 error / 36 warning | literal-batch-lint.log |
| A1-E2 门控 | Excel 6 / WPS 194 passed，exit 0 | literal-batch-a1-e2.log |
| compileall server/app / git diff --check | exit 0 / 0 | 本轮命令输出 |

日志均在 outputs/。A1-E2 消费既有六份真实工件，没有重新运行真实模型或 Docker。
最后仅为已有 HTTP 往返用例增加 version=2 参数（生产代码未再修改），相关三文件复跑161条；
未为这一条测试参数再重复全量服务端。现存 React act、WPS chunk 与 LF/CRLF 提示未隐藏。
本轮没有新增反向变异实验，不复用前轮变异冒充本轮证据。
