# 本地阶段性检查点 — 2026-10-17

本提交保存截至本次操作的累计源码、测试与项目文档，供后续修复接续；不表示整体交付完成。

## 本次实际复跑

- Excel：1286 passed / 3 skipped，退出码 0。
- WPS：28 failed / 839 passed / 3 skipped，退出码 1。
- core：542 passed，退出码 0。
- build:addin、build:wps：退出码 0（包含类型检查）。
- 服务端：2 failed / 1426 passed / 28 skipped。两条失败位于 test_python_event_delivery.py，预期 completed、实际 unavailable；本次未进一步确认执行器不可用的具体环境原因。
- compileall server/app：退出码 0。

本机完整复跑日志位于 outputs/git-check-*.log（按现有规则不入库）。git diff --check 仍报告三个既有 EOF 空行：dataTools.ts、ARCHITECTURE.md、DEVELOPMENT.md；本次 Git 归档未修改其代码或内容。

## 待完成

- WPS 三个测试文件中的 28 条失败需逐项核对新结果审阅流程和真实行为，不能一概视为过时测试。
- WPS 工作目录目前完成选择与授权交互，查询/执行链路尚未接通。
- 主表结果、临时标记、可选执行前双视图、确认后保存工具与撤销的新流程，仍需完整收口与真实宿主验收。

## 归档范围

纳入已跟踪改动及 apps、packages、scripts、server/app、server/tests、docs 下新增项目文件，和非生成的 config 文件。排除测试临时目录、outputs、本机助手配置、外部 SDK checkout、根目录草稿报告与生成的 config/cross_process_chain.json。

.workbuddy/memory 按既有 .gitignore 保留在本机，不强制加入 Git；本次未删除任何排除项。未推送或发布。