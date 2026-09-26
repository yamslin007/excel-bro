# WPS W3-A 写入能力预检记录（2026-09-23）

> 2026-09-25 更新：W3 代码闭环已经实现并由自动化测试覆盖，包括计划确认门禁、
> `writeValues`、`clearRange:contents`、执行前持久化快照、分块写入、回读校验、撤销和
> 冲突拒绝。`wpsWrite.enabled` 现已开启。下文保留 2026-09-23 当时的历史预检结论；
> 真实 WPS 12.1.0.28488 上的写入、公式恢复、刷新后撤销和异常中断仍需按清单复验。

## 当前结论

W3-A 的代码侧预检已加入；目标 WPS 中的实际写入、读回、公式恢复和中断恢复尚未验收。
`config/capabilities.json` 的 `wpsWrite.enabled` 默认为 `false`，WPS 对话仍保持只读。
本记录不能作为开放写入的依据。

## 已完成

- 本机运行的 WPS 表格进程路径包含 `12.1.0.28488`，与 W0～W2 的目标版本一致。
- 为 WPS 宿主类型补充 `Value2`、`Formula`、`MergeCells`、`ProtectContents`、`ReadOnly` 和 `Version` 的读取声明。
- 增加只读预检，检查本地已保存 `.xlsx`、明确的非只读状态、已选工作表、A1 区域读取、公式读取、合并状态和工作表保护状态。
- 预检只展示能力结果与宿主版本，不输出单元格原值、公式内容或文件完整路径。
- 开发版任务窗格的“授权读取范围”下增加“W3 写入能力检查（只读）”。
- 能力配置增加 `wpsWrite.enabled=false`、`maxCells=1000`、`chunkRows=100`。
- WPS 自动化测试 21 项及构建通过；探测不会调用任何写入接口。
- Excel 前端测试 442 项及构建通过；服务端测试使用隔离 `--basetemp` 后 437 项通过，`compileall` 通过。
- 当前开发服务的 3889 端口已能返回更新后的诊断组件；尚未取得任务窗格内的实际检查结果。

## API 依据

WPS 官方文档说明 [`Range.Value2`](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Range/member/Value2)
可读写单元格值，[`Range.Formula`](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Range/obj)
可读写 A1 公式，[`Range.MergeCells`](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Range/member/MergeCells)
可判断合并状态，[`Worksheet.ProtectContents`](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Worksheet/obj)
可判断内容保护，[`Workbook.ReadOnly`](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Workbook/member/ReadOnly)
可判断只读打开状态。这些资料说明接口设计，不能代替在 `12.1.0.28488` 中实测。

## 待目标宿主实测

先打开专用测试工作簿，运行 `npm run dev:wps`，在 WPS “开始”选项卡打开 Excel Bro，展开底部
工作簿范围，再展开“W3 写入能力检查（只读）”并点击检查。记录每项结果和宿主版本。

- [ ] 本地已保存 `.xlsx` 上的预检结果与 WPS 实际状态一致。
- [ ] 只读文件、受保护表、合并单元格与未保存文件均被正确识别。
- [ ] 单格、单行、单列和二维范围的 `Value2` 读回形状及赋值效果。
- [ ] 原单元格为空白、常量、公式时，`Formula` 与 `Value2` 可区分并完整恢复。
- [ ] 写入中异常后能确定实际改动区域；无法确定时保持挂起。
- [ ] 工作簿切换、同名文件、另存为和任务窗格刷新后不会误认原文件。

后四项需要在独立测试文件中完成受控写入与恢复试验。当前开发诊断只做只读预检，
不会在用户工作簿中试写。W3-B 写入执行器须等这些接口行为获得实测证据后继续。

## 现有恢复协议差异

当前 `/api/office/executions/register` 只接受 `url:` 开头的 Office.js 工作簿身份，
并禁止来源工作表与写入目标重合。WPS 的本地 `.xlsx` 身份和首版“修改已授权现有表”
无法直接套用该接口。后续实现应在同一任务账本基础上增加宿主区分和独立的 WPS
工作簿身份校验，不能把本地路径伪装成 `url:`，也不能取消现有 Office.js 的来源保护。
协议字段变更须同时更新前端 `contracts.ts` 和后端 `models.py`。
