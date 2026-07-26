# Excel Bro 开发约定

开始修改前先阅读：

- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`

必须保持的边界：

- 模型负责语义判断，本地工具负责确定性读取与计算。
- 用户选择的工作表是不可扩大的授权边界。
- 意图阶段不向模型发送原始数据行。
- 所有 Excel 写入必须使用白名单 `AnalysisPlan`，先预览、再由用户确认执行。
- 不加入任意脚本、VBA、宏或外部程序执行。
- 不在运行时代码中硬编码行业字段或客户数据。
- 协议变更必须同步 `apps/excel-addin/src/contracts.ts` 与 `server/app/models.py`。
- 项目限额优先放在 `config/capabilities.json`。
- 新功能应同时考虑 Office.js 当前工作簿模式和 openpyxl 文件夹模式。

提交前至少运行：

```powershell
npm run test:addin
npm run build:addin
python -m pytest server/tests -q
python -m compileall -q server/app
```

架构、协议、启动方式或环境变量变化时，同步更新 `docs/` 和根目录 `README.md`。
