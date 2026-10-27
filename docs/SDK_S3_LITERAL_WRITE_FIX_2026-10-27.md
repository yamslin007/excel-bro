# S3 字面值写入修复（2026-10-27）

## 原生反馈与根因

WPS-S3-确认.xlsx：用户要求只在目标!A1写入文本「完成」，其他单元格不变。可信 plan_range_edit 生成的 writeValues 被动作类型统一归为计算结果；没有数据读取桥/来源锚点时被拒。WPS 拒绝位置又处在执行异常处理内，导致未开始写入却显示结果不确定、恢复记录缺失。

## 修改与边界

- 服务端 compile_plan_tool 仅给 plan_range_edit/writeValues 返回 planOrigin=literal_range_edit；raw模型JSON解析清除此字段，双方响应协议同步。
- core literalPlanOrigin 对单一、统一字面值矩阵加绑定id/actions的内存封印；双端接收路线接入。JSON往返、修改动作、伪造普通属性均不具备封印。
- core 来源闸识别封印；其他 writeValues/writeTable 计算结果的来源变化、缺失锚点、来源不可读仍拒绝。字面值路线不替代授权、目标检查、恢复准备与回读。
- WPS 审阅准备先调用来源检查，执行器仍再次检查；共享审阅控制器区分 execute 调用前拒绝(not_started)与执行中失败(failed)。不以该状态推断所有失败都未写入。
- 确认仍只接受结果并清标记，放弃走恢复点，不增加执行前确认。

## 自动化证据

新增覆盖：可信编译器来源与原始JSON伪造、封印失效、审阅准备拒绝/执行中失败分开、两端无桥字面值执行且确认不重写、WPS缺锚点计算结果零写入且重开无错误恢复卡。

| 检查 | 结果 | 日志 |
|---|---|---|
| core | 851 passed | outputs/literal-s3-core-full.log |
| Excel | 1564 passed / 9 skipped | outputs/literal-s3-addin-full.log |
| WPS | 1134 passed / 8 skipped | outputs/literal-s3-wps-full.log |
| 服务端 | 1844 passed / 1 skipped | outputs/literal-s3-server-full.log |
| 两端build（含类型检查） | exit 0 | outputs/literal-s3-build-addin.log、literal-s3-build-wps.log |
| lint | 0 error / 36 warning | outputs/literal-s3-lint.log |
| A1-E2门控 | exit 0；复用既有工件，非重新运行真实模型实验 | outputs/literal-s3-a1-e2.log |

没有本轮变异证据，也没有修复后原生宿主证据。WPS已有React act告警、构建chunk提示未隐藏。另尝试的独立core workspace build仍受既有ColumnMapping测试类型错误阻断（blankPolicy/columnMapping/columnMappingSemantics的values类型）；两端生产build及上述测试通过，不将独立core build写成通过。

## 用户复验

1. 重启正在使用的后端，重新加载WPS的Excel Bro窗格（协议和前端均有变更，本轮未代为重启）。
2. 在独立测试簿WPS-S3-确认.xlsx，选中并授权「目标」。旧失败计划不自动升级，不点击旧恢复卡去重放；重新发送：仅在「目标」工作表A1写入文本「完成」，其他单元格保持不变，直接执行后展示结果供我审阅。
3. 预期A1自动写入完成，B2:B4仍为-101/-202/-303，出现双视图/放弃变更/确认变更。点击确认仅清标记，不再写一次。
4. 另用独立放弃测试簿重复，点击放弃后A1恢复为空；关闭重开窗格不应出现本轮未执行的错误恢复提示。

旧版本已落盘的恢复提示不在本次批量清除，避免删除其他真实中断任务。如果仍显示旧提示，需按任务身份核对，不能据新一轮成功就任意清空所有恢复记录。

A1-GATE仍未通过，A2未启动；未操作用户工作簿、未使用computer-use、未提交/推送Git。


### 2026-10-27：用户完成 S3 确认 / 放弃后的窗格重开复验

用户提交六张真实 WPS 截图，并明确回复「我已重開」。确认路径：目标!A1 保留「完成」，黄色审阅标记移除，重开后显示「恢复点已不可用」；这是会话内撤销的既定边界。放弃路径：A1 恢复为空，B2:B4 保持 -101/-202/-303，重开后显示已撤销并恢复，没有本轮待恢复入口。S3 的这两条 WPS 可见行为已有原生证据；用户先前 Excel 截图也显示确认后保留「完成」、放弃后恢复为空。

证据归档：outputs/a1e3-native/s3-reopen-20261027/（六张 WPS 截图与 evidence.json）。本次重开由用户操作；不据此认定 WPS 进程或 WebView 已销毁重建。

重复写入审计未闭环：默认本机账本 C:/Users/jia/AppData/Local/Excel Bro/tasks/state.sqlite3 使用 mode=ro 查询 office_executions 时返回 sqlite3.DatabaseError: database disk image is malformed；四个隔离验收账本均未匹配 S3 执行。未修改、重建或删除任何账本，也未操作正在运行的后端。单元格值相同不能证明未重复写入，调用次数与 observation 均记为未采集。后续须先定位本次服务实际账本，并对数据库损坏单独诊断，保存原文件及 WAL 后再决定修复。

A1-GATE 保持未通过，A2 未启动。下一步优先闭环账本可读性与 S3 不重复写入证据，再补 WPS 完整重载和原生故障矩阵。此轮只归档证据及更新文档，未改生产代码、未提交或推送 Git。


### 2026-10-27：账本损坏结论撤回，真实账本及 S3 WPS 审计通过

Codex MSIX 将默认账本读取重定向到自己的 LocalCache 旧主文件，前轮把旧主文件和真实 WAL 混读，误判生产损坏。通过真实后端文件句柄取得的主文件为 2289664 字节，副本合并 WAL 后 integrity_check=ok、foreign_key_check 空；130 条执行、129 条规划、68 条宿主写观测。原数据库无需修复；未改写、重建或删 WAL。后端短暂停住用于保持句柄位置，已恢复，真实审计接口 GET 随后成功；撤回暂停新写入建议。

WPS S3 确认 c4244afe… 与放弃 d1545de6… 两轮均 completed，动作推进 1/1；每轮 hostWriteCalls=3 wps_api_calls、mutatingCalls=1、readback=verified。服务端均返回 complete/nonzero、无缺失/越界观测；记录内未见确认/重开导致重复计划内容写入。放弃/审阅辅助写不属于该计数。Excel S3 此轮未匹配到记录，不借用 WPS 计数。

最终证据：outputs/ledger-diagnosis-20261027/REPORT.md、handle-snapshot-111950/verification.json、s3-live-host-writes.json。旧505页/损坏结论及相应阻塞项已被本条取代；旧错误样本保留且标记 superseded。A1-GATE 仍未通过，A2 未启动；原生矩阵与完整重载等缺口不因本条被标完。
