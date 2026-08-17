# 性能基线（2026-08-17）

## Bundle 大小

构建命令：`npm run build:addin`

主要产物位于 `apps/excel-addin/dist/`：

- taskpane JS：`assets/taskpane-DW6V5o-C.js`，222.12 KB（217.0 KiB）
- 主样式 CSS：`assets/styles-SD97rhBM.css`，68.58 KB（67.0 KiB）
- 主样式 JS：`assets/styles-CKZ1k6c1.js`，192.52 KB（188.0 KiB）
- 专注窗口 CSS：`assets/focus-9Mpkc-nF.css`，5.69 KB（5.6 KiB）
- 专注窗口 JS：`assets/focus-BgKBczoe.js`，5.87 KB（5.7 KiB）
- 总大小：491.83 KiB（503,631 字节）

> 说明：当前 Vite 产物没有名为 `taskpane.css` 的固定文件，任务窗格样式被拆到
> 哈希命名的 `styles-*.css` / `styles-*.js` 中；上表按实际产物记录。

## 测试性能

测试命令：`npm run test:addin`

- 测试数量：258
- Vitest 运行时间：4.32s
- npm 端到端耗时：5.78s
- 平均每个测试：16.7ms（按 Vitest 4.32s 计算）

## 运行环境

- Node 版本：v24.12.0
- npm 版本：11.8.0
- 操作系统：Windows 11（Microsoft Windows NT 10.0.26100.0）
- Vitest 版本：3.2.7

## 备注

建立基线后的优化目标：

- Bundle 大小不应显著增长（允许 ±10%）
- 测试时间不应变慢（允许 ±20%）
