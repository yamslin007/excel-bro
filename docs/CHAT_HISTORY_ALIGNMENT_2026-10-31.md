# Excel / WPS 历史对话对齐（2026-10-31）

## 代码核实与差异

核实的是当前工作区代码；本轮未读取宿主 localStorage 或真实聊天正文，未操作工作簿。

| 能力 | Excel 原实现 | WPS 原实现 | 本轮结果 / 证据边界 |
| --- | --- | --- | --- |
| 列表、新建保留、切换消息、继续普通聊天 | 已实现，useConversation + App | 已实现，WpsAgentPanel；并非缺少历史功能 | 双端组件及既有回归覆盖；WPS 新增创建甲乙、重挂载、切回甲续聊；原生实机待验 |
| 删除 | 有取消/确认 | 原先直接删除 | 共用 HistoryDeleteDialog；取消不改历史，确认仅删除指定 ID；双端真实面板测试 |
| 安全标记 | 删除保留其他记录标记 | 原删除重建对象时遗漏 legacyCredentialConversationIds | 共用 core/conversationHistory.removeHistoryConversation，保留对象其他元数据，仅过滤不再存在的 ID；切换保留标记 |
| 保存失败 | 四态 HistoryStorageNotice（含旧键迁移状态） | 已有 storage-write-warning；成功后清提示 | 保留既有机制，自动化覆盖失败不报成功；不把错误原文写入提示 |
| 重开加载 | excel-bro.chat.v4；有既有 v3 兼容路径 | excel-bro.wps.chat.v1；已有加载/脱敏/保护标记 | 本轮不新建存储、不改键、不添加迁移清理；组件重挂载已验，完整退出宿主持久性待验 |
| 容量 | conversation：20 个、持久每个 60 条；内存每个 100 条 | 硬编码 30 个、持久每个 120 条 | WPS 改读 conversation.wps 显式 30 / 120；Excel 数值不变；30×120 合成数据加载→切换→保存→重挂载保留首尾消息 |
| 历史计划 | 展示不等于执行资格 | 已有“历史不恢复可执行计划”回归 | 本轮继续保持；加载/切换/删除不调用执行、认领、补记；替身宿主零写入 |
| 菜单说明 | 查看并恢复以前的对话 | 查看当前 WPS 会话记录 | 共用“查看并打开已保存的对话” |

双端功能代码均已实现；自动化使用真实面板、纯逻辑与合成聊天，宿主和网络为替身。原生宿主、存储配额与完整进程退出行为不由 jsdom 证明。本轮未发现足以引入新后端历史存储的证据。

## 兼容与边界

WPS 不借用 Excel 较小的默认值，原键和数据结构继续使用。30×120 既有记录不会因本轮配置接入变成 20×60。容量策略仍是有限历史：超过各端原限额时沿用既有保留/淘汰规则；这不是无限归档保证。未来统一限额需先明确非破坏迁移、导出或用户可见的容量处置策略，不能直接调小配置并静默裁剪。

共享的是菜单说明、确认组件、删除纯函数；WPS 的历史序列化/反序列化、脱敏收口和宿主适配保留，未大幅重构面板。Excel 既有旧键迁移规则未改动，本轮没有添加任何删旧键或清空存储的操作。

打开历史只恢复展示；继续普通聊天仍通过现有理解入口，新业务请求仍复核当前工作簿、授权和来源/目标现场。删除聊天不会撤销执行或删除账本、恢复包、未完成任务。SDK 转录持久化开关、U3/U2、后端协议均未改动，A2 未启动。已有工作区其他变更保留，未提交、未推送。

## 最短真机验收（待用户执行）

代码与自动化完成后，用户自行加载新版窗格；本轮无需重启后端。仅进行无敏感普通聊天，不选择表数据、不发起表格任务。

1. 先记录 WPS 窗格**实际完整 URL**（开发者工具 Console 只读查看 `location.href`，不要查看 localStorage）；同时记协议、主机、端口和路径。实际地址本轮未采集，不能用 vite 默认端口或测试日志冒充。
2. 发“普通聊天：请回复测试甲”，新建后发“普通聊天：请回复测试乙”。打开历史，分别切回查看，确认甲乙内容对应。
3. 用户自行完整退出 WPS，再重开窗格；再次记录实际 URL，核对同一来源和两条记录。收起窗格或组件重挂载不能替代完整退出。
4. 删除甲，先取消：甲乙都在；再确认删除甲：只有甲消失，乙可打开。无需改动任何单元格。
5. 记录两次 URL、甲乙是否保留、删除取消/确认结果、是否出现存储失败提示。若来源（协议/主机/端口）不同，先归类为不同存储分区；不要清缓存、删键或重启服务来“修复”。同一来源仍丢失才是待定位的宿主持久性问题。

Excel 同样待验菜单、确认取消、切换和重开行为；WPS 完整退出持久性为本轮重点待验项。可使用空白无敏感工作簿，但验收聊天不得读取或修改工作表，不使用已验收工作簿。

## 自动化结果

日志统一在 `outputs/history-align-*`。所有聊天夹具均为合成文字。

| 实际命令 | 真实结果 | 日志 |
| --- | --- | --- |
| `npm run test --workspace @excel-bro/addin -- --maxWorkers=1` | 1638 passed / 9 skipped | history-align-addin-full-final.log |
| `npm run test:wps` | 1258 passed / 8 skipped | history-align-wps-full.log |
| `npm run test --workspace @excel-bro/addin -- src/chatHistory.test.ts src/messageFirstRouting.test.tsx src/continuousPythonChain.test.tsx src/excelWriteEndToEnd.test.tsx --maxWorkers=1` | 194 passed | history-align-excel-regression.log |
| `npm run test --workspace excel-bro-wps -- src/WpsAgentPanel.test.tsx -t 'history alignment\|does not restore an executable plan\|warns without leaking' --maxWorkers=1` | 5 passed / 229 未选中；使用修正后的宿主夹具 | history-align-wps-final-focused.log |
| `npm run test --workspace @excel-bro/addin -- src/messageFirstRouting.test.tsx -t 'history alignment' --maxWorkers=1` | 1 passed / 56 未选中；最终增加执行/认领/补记/处置/恢复包删除接口零调用断言 | history-align-excel-final-focused.log |
| `npm run build:addin` / `npm run build:wps` | 均 exit 0，均包含 tsc --noEmit；保留既有大包提示 | history-align-build-addin-final3.log / history-align-build-wps-final2.log |
| `.\.venv\Scripts\python.exe -m pytest server/tests/test_capability_matrix_coverage.py server/tests/test_action_capability_shared_cases.py -q -p no:cacheprovider --basetemp <独立系统临时目录>` | 22 passed | history-align-server-config-final.log |
| `.\.venv\Scripts\python.exe -m compileall -q server/app` | exit 0 | 本地命令 |
| `npx eslint apps/excel-addin/src/components/HistoryDeleteDialog.tsx packages/core/src/conversationHistory.ts apps/excel-addin/src/conversation.ts apps/excel-addin/src/messageFirstRouting.test.tsx apps/wps-addin/src/WpsAgentPanel.test.tsx` | 0 error / 1 条原有 Hook warning | history-align-lint.log |
| `git diff --check -- <本轮涉及文件>` | exit 0 | 本地命令 |

首次 Excel 定向测试暴露新增测试缺少宿主初始化参数，已修正；首次最终构建又发现测试夹具返回值误用及 Testing Library 不支持的 exact 查询选项，均已修正并重新通过。首次并行全量 Excel 有三条既有执行链用例 5 秒超时（1635 passed / 3 failed / 9 skipped），相关四文件单工作进程 194 passed，随后 Excel 全量单工作进程通过；不修改业务代码或扩大超时来掩盖失败。首次服务端定向测试因默认 pytest 临时目录 WinError 5 全部初始化失败，改为全新系统临时目录后 22 passed；未删除旧目录。

本轮未修改后端实现，服务端按共享配置影响跑上述 22 条回归，未跑全量 server/tests。最早标为 focused 的日志实际因根 npm 脚本转发丢失 -t 跑了整文件；以上最终定向命令直接指定 workspace，日志明确给出选中与未选中数。首次失败日志保留，不以最终成功覆盖。
