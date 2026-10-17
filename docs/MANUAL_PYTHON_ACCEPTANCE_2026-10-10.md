# 隔离 Python 双端实机验收（2026-10-10）

## 交付与启动

独立测试工作簿：`outputs/manual-acceptance-20261010/Python验收-WPS.xlsx` 和
`Python验收-Excel.xlsx`。各含空白「汇总表」和两张来源表，数据来自
`config/online_link_scenario.json`；编号按文本保存，保留空格、前导零和空编号。
没有预填分析结果或公式。

项目根目录执行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start_manual_python_acceptance.ps1 -PreflightOnly
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start_manual_python_acceptance.ps1
```

启动器给本次后端设置 `outputs/manual-acceptance-20261010/server-capabilities.json`，
仅把 `pythonExecution.enabled` 打开；默认能力表保持关闭。不注入模型替身，不设置测试
模型目录，沿用已保存连接。预检核实本机 Docker Linux 引擎及固定 digest 的镜像，缺失即退出。
正式启动复用 SDK 启动脚本：正常产品应用、回环 8765、单 worker、无热重载，两种后端共用。
端口已占用时脚本拒绝启动，不自动停止用户进程。

只读就绪检查：

```powershell
Invoke-RestMethod http://127.0.0.1:8765/health
(Invoke-RestMethod http://127.0.0.1:8765/api/capabilities).pythonExecution.enabled
(Invoke-RestMethod http://127.0.0.1:8765/api/models).models | Select-Object id,label,available
```

WPS 前端：`npm run dev:wps`（3889）；已有前端时复用并刷新任务窗格。
Excel 前端：`npm run dev:addin`（HTTPS 3000）。已核实本机旁加载工具支持 `--document`，
可直接加载准备好的 Excel 文件：

```powershell
.\node_modules\.bin\office-addin-debugging.cmd start apps/excel-addin/manifest.xml desktop --no-debug --document "E:\codex\excel-bro\outputs\manual-acceptance-20261010\Python验收-Excel.xlsx"
```

普通 `npm run start:excel` 则给它自动创建的临时文件旁加载。不要把另一工作簿缺少加载项
当作 Python 缺陷。

退出验收后端，再运行原 `start_claude_sdk_server.ps1` 可恢复默认关闭 Python 的配置。
验收配置不写入生产默认文件。实机不得使用 `server/tests/online_entry.py`。

## 后端选择

本机菜单的只读检查结果：普通连接 `DS · deepseek-v4-flash`、SDK 连接
`Agent SDK · DS · deepseek-v4-flash`，另有普通连接 `vec · claude-opus-4-6`。
标签可用只证明配置可选，不代表真实端点或工具调用已经通过。
先 WPS 普通连接，再 WPS SDK；每轮恢复基线、新建对话。之后在 Excel 重复。
三张表都要授权，授权与写入确认分别进行。

## 发送指令

> 用隔离 Python 分析这两张已授权来源表：以订单表为左表，按会员编号左连接会员表。
> 编号只去除首尾空格，保留前导零；空编号不参与匹配；会员表重复编号保留首次出现，
> 后续重复记录单独说明。按关联得到的等级汇总订单金额，未匹配订单归入“未匹配”。
> 展示汇总与可定位的空编号、重复编号、未匹配明细，并给出金额总计。
> 将汇总和异常明细安排到汇总表，使用定值矩阵，先展示写入预览，等我确认后再执行。
> 不修改会员表、订单表，也不把原始数据行发送给模型。

允许结果行序、文字和布局不同，事实必须一致：

| 检查项 | 基线 | 修改来源后 |
|---|---:|---:|
| 金卡金额 | 400 | 300 |
| 银卡金额 | 200 | 300 |
| 未匹配金额 | 400 | 400 |
| 金额总计 | 1000 | 1000 |
| 空编号 | 会员表第 5 行，1 条 | 同基线 |
| 重复编号 | 会员表第 3 行，编号 7，1 条 | 同基线 |
| 未匹配订单 | 订单表第 5 行，编号 404，1 条 | 同基线 |

## 操作步骤

1. 打开已保存的对应测试文件，确认汇总表为空，授权三张表并完成选择。
2. 选当前要验收的模型连接，新建对话，人工发送上述指令。
3. 核对预览结果与基线；确认前原工作簿零写入，两张来源表不变。
4. WPS 生成副本并排审阅，核对后确认写回原表；Excel 按当前可用入口明确确认执行。
5. 核对写入结果和产品回读成功提示，点击该次执行的撤销，检查目标完整恢复为空。
6. 将 `会员表!C2` 从金卡改为银卡，新建对话，重跑同一指令。
7. 应得到 300 / 300 / 400，总计仍 1000；完成确认、回读与撤销，然后把 C2 还原为金卡。
8. 换后端或宿主前恢复基线，只确认当前轮计划。

只有异常数量而没有定位信息，异常明细验收未完成。若模型拒绝、工具失败、没有计划、
写入被拒或回读不符，记录原文与 taskId，不以预制结果补齐。

## 实机证据（准备完成不等于通过）

| 宿主 | 后端 | 自主规划/计算 | 预览/确认/回读/撤销 | 因果验证 |
|---|---|---|---|---|
| WPS | 普通 | 未完成 | 未完成 | 未完成 |
| WPS | SDK | 未完成 | 未完成 | 未完成 |
| Excel | 普通 | 未完成 | 未完成 | 未完成 |
| Excel | SDK | 未完成 | 未完成 | 未完成 |

每条记录应有客户端版本、模型标签、时间、任务标识、预览与写后/撤销截图。
自动化测试不能代填此表。本轮准备不发送真实模型请求，不代替用户点击写入确认。
