# Excel Bro 开发约定

开始修改前先阅读：

- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`

必须保持的边界：

- 模型负责语义判断，本地工具负责确定性读取与计算。
- 用户选择的工作表是不可扩大的授权边界。
- 意图阶段不向模型发送原始数据行。
- 所有 Excel 写入必须使用白名单 `AnalysisPlan`，先预览、再由用户确认执行。
- 不加入宿主中的任意脚本、VBA、宏或任意系统命令执行。唯一获用户批准的代码执行例外是“模型生成的 Python 只在已验证的隔离执行器内做分析”；不得直接修改工作簿、访问宿主文件/网络/密钥或回退到宿主裸执行。具体边界见 `docs/SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md`；本次是设计授权，运行时尚未实现，默认不开放。
- 不在运行时代码中硬编码行业字段或客户数据。
- 协议变更必须同步 `apps/excel-addin/src/contracts.ts` 与 `server/app/models.py`。
- 项目限额优先放在 `config/capabilities.json`。
- 新功能应同时考虑 Office.js 当前工作簿模式和 openpyxl 文件夹模式。
- 新增功能必须同时设计、实现和验收 Excel 与 WPS 两端，按用户任务能力、交互语义、授权/隐私/写入/恢复边界对齐；不能只实现一端就宣称整体完成。
- 共用语义、协议与确定性逻辑优先放入共享层，宿主 API 差异放入各端适配层；不以同款 UI 或单端测试通过代替双端能力证据。
- 受宿主限制或获用户明确同意分批实施时，在计划表分别记录两端“已实现 / 未实现 / 宿主限制 / 实机待验”、具体差距及验收标准；未实现不能标为实机待验，不为对齐放宽安全边界。

提交前至少运行：

```powershell
npm run test:addin
npm run build:addin
npm run test:wps
npm run build:wps
python -m pytest server/tests -q
python -m compileall -q server/app
```

架构、协议、启动方式或环境变量变化时，同步更新 `docs/` 和根目录 `README.md`。
