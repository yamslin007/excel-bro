# A1 独立进程退出与竞争验收（2026-10-24）

本批只新增测试和证据文档，没有修改生产执行协议、授权、写入或审阅逻辑。先执行、再审阅约定不变。未使用 computer-use，未关闭用户 Excel/WPS 或运行中的后端。

## 实际完成

`server/tests/test_execution_recovery_a1.py` 新增两个参数化场景，Office/WPS 各两条，共 4 条。测试经真实 FastAPI ASGI 路由访问临时 SQLite；每个 worker 使用独立解释器、独立 ledger 实例和测试配置目录。不是网络套接字、真实宿主或真实窗格测试。

1. 第一动作完成，独立 worker 提交第二动作 start 并确认 SQLite 已提交；父进程仅 kill 自己创建的 worker，不调用 suspend，也不走 graceful shutdown。新 worker 读取相同临时数据库，nextActionIndex=1、inFlightActionIndex=1、revision 保持。提交合成 applied 核对证据后游标到 2，仅提交第三动作 start/complete，最终 completed。
2. 两个独立 worker 用同一 revision 竞争恢复，断言一个 200、一个 409、revision 只增加 1。这比单进程内两个 TestClient 多覆盖跨解释器/跨 ledger 实例共享 SQLite，但不冒充两个原生窗格。

WPS 进程退出后持久租约仍有效，立即补记应返回 409。测试显式推进测试进程时钟到租约过期后才核对，没有更改数据库时间、降低生产租约或把退出当成 suspend。这一段是模拟过期，不是真实等待租约实测。

## 不能据此宣称

- 未杀死真实 Excel/WPS 或 WebView；未验证宿主断电、崩溃、未保存文件恢复。
- 第二步仅在协议账本中登记在途；测试没有实际写单元格。applied 的指纹/检查点是合成证据。宿主只写第三步的证据仍来自前一批真实组件/执行器加替身宿主计数，不能把两个测试拼成单条原生端到端证据。
- 未实现或验收两原生窗格同时操作同一工作簿；本批并发仅为独立 API 进程。
- A1 仍未整体完成，A2 未开启。

## 复跑

在仓库根目录运行（每次使用新的临时目录名）：

```powershell
.\.venv\Scripts\python.exe -m pytest server/tests/test_execution_recovery_a1.py -q -p no:cacheprovider --basetemp=.tmp-a1-process-manual-001
```

定向结果：14 passed，8.90s；日志 `outputs/a1-process-focused.log`。新增 4 条无生产缺陷暴露，不为“修复”而修改通过验证的生产逻辑。

## 下一项

继续 A1 的恢复材料过期/缺失闭环：两端分别核对过期后是否只读说明原因、拒绝旧计划续写、显式重新规划使用新身份、完成后旧提示不会复活；既有原生恢复审阅反馈继续有效。真实宿主进程终止、跨窗格竞争维持待验，需独立测试文件及实际调用日志，不能仅靠截图值正确判定未重写。

全量验证：服务端 1748 passed / 1 skipped（86.07s），日志 outputs/a1-process-server-full.log；compileall 与 git diff --check 通过。本批仅测试/文档变化，未重复构建前端，未将上一轮双端测试数冒充本轮结果。
