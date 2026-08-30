# Bundle 大小分析

分析日期：2026-08-30

## 结论

当前没有未使用的大型第三方依赖；Office.js 仍通过 Microsoft CDN 外部加载，不在 bundle
中。主要体积来自 React DOM、`App.tsx` 和 Excel 执行层。计划编辑器已经是独立入口，
继续为小面板做懒加载的稳定收益不足 10KB。

本轮决策为“不立即优化”。唯一可能超过门槛的方向是把 `excel.ts` 拆成“首屏结构读取”
与“执行/验收”模块，但它跨越 Office.js 与文件夹模式，回归风险高；没有真实加载慢反馈
时不值得在本轮实施。

可视化报告见 [bundle-visualizer.html](performance/bundle-profile/bundle-visualizer.html)，
原始数据见 [bundle-stats.json](performance/bundle-profile/bundle-stats.json)，生成脚本见
[`bundle-profile.mjs`](../apps/excel-addin/scripts/bundle-profile.mjs)。

## 构建结果

命令：`npm run build:addin`

- Vite 7.3.6，87 个模块；构建 1.14s。
- `dist` 总计 616,104 字节（601.66KiB）。
- 2026-08-17 基线为 503,631 字节（491.83KiB）；当前增加 112,473 字节，约 22.3%。
- 单看 taskpane 入口从 222.12KB 降到 157.16KB，但现在有 113.76KB 的共享
  `useTheme-*` chunk、13.72KB 的计划编辑器入口，主 CSS 也从 68.58KB 增到
  115.29KB，所以不能用 taskpane 单文件下降来宣称整体变小。

当前主要产物：

| 产物 | 原始大小 | gzip |
| --- | ---: | ---: |
| `styles-*.js`（React/ReactDOM 共享运行时） | 192.49KB | 60.35KB |
| `taskpane-*.js` | 157.16KB | 49.69KB |
| `styles-*.css` | 115.29KB | 19.31KB |
| `useTheme-*.js`（Excel/存储/数据工具共享代码） | 113.76KB | 35.25KB |
| `planEditor-*.js` | 13.72KB | 4.89KB |
| `focus-*.js` | 5.91KB | 1.96KB |

`index.html` 会预加载 taskpane、React 共享运行时、`useTheme-*`、functions 和主 CSS；
按产物原始字节合计约 581.7KB，按 Vite 输出的逐文件 gzip 合计约 166.3KB（不含外部
Office.js）。

## 前五大模块

下表为 Rollup visualizer 的 rendered 模块估算；它用于比较模块占比，不等于磁盘上的
minified chunk 大小。逐模块 gzip 也不能简单相加，因为整 chunk 压缩可以跨模块复用。

| 排名 | 模块 | rendered | gzip 估算 | 占全部 rendered |
| ---: | --- | ---: | ---: | ---: |
| 1 | `react-dom-client.production.js` | 552,879B | 95,428B | 46.3% |
| 2 | `src/App.tsx` | 242,719B | 44,374B | 20.3% |
| 3 | `src/excel.ts` | 118,552B | 23,839B | 9.9% |
| 4 | `src/storage.ts` | 22,651B | 5,513B | 1.9% |
| 5 | `src/components/PlanEditorDialog.tsx` | 20,179B | 4,408B | 1.7% |

React 本身和 scheduler 分别约 18.3KB、11.1KB rendered；项目没有 UI 框架、图表库、
日期库等可直接删除的大型依赖。

## 依赖与代码分割检查

### 已生效

- Office.js 使用外部 CDN，不计入 bundle。
- `planEditor`、`focus`、`functions` 和 taskpane 是独立入口。
- React/ReactDOM 被抽成共享 chunk，避免多个入口重复打包。
- 计划编辑器业务入口只有 13.72KB；它不是 taskpane 首屏内联的大型对话框。

### 可选但暂不实施

1. **拆分 `excel.ts`（潜在 15–25KB gzip，风险高）**

   Visualizer 估算 `excel.ts` 为 23.84KB gzip，并进入首屏预加载的共享 chunk。理论上可把
   计划执行、验收、撤销等非首屏函数动态加载，只保留工作簿结构扫描。但文件内部共享
   大量类型、安全校验和 Office.js helper，同时要兼顾 openpyxl 文件夹模式；需要专项
   设计和真实冷启动测量，不作为本轮顺手优化。

2. **拆分计划编辑器 CSS（约 10.29KB minified raw，gzip 预计明显低于 10KB）**

   `styles.css` 第 6415 行后的计划编辑器区段约 12.34KB 源码、10.29KB minified。它刚到
   原始字节门槛，但压缩收益低于 10KB，还会增加共享样式边界维护，暂不做。

3. **继续拆 `App.tsx`（bundle 收益接近 0）**

   抽成同步子组件只改善可维护性，不会减少首屏下载；改成动态组件才可能减少体积。
   阶段 2 又没有稳定渲染瓶颈，因此不能把结构重构包装成性能收益。

### 不值得做

- `petTips.ts` gzip 估算约 3.1KB，低于 10KB。
- ThemePanel、Slash 自动补全等模块单体均低于门槛。
- 用 Preact 或自定义运行时替换 React 可能减少体积，但会显著扩大兼容与维护风险。
- 对已经独立的计划编辑器再套 `React.lazy`，不会减少 taskpane 当前预加载的共享 Excel
  层，收益有限。

## 优先级与决策

| 优先级 | 建议 | 预期收益 | 维护成本 | 决策 |
| --- | --- | ---: | --- | --- |
| P1 按需 | 用户明确反馈冷启动慢时，专项拆分 `excel.ts` | 15–25KB gzip（需实验确认） | 高 | 暂缓 |
| P2 | 分离计划编辑器专属 CSS | <10KB gzip | 中 | 不做 |
| P3 | 懒加载小面板、提示文案或增加手工 chunk | <10KB | 中 | 不做 |

虽然总 `dist` 已超过旧基线 22.3%，新增体积主要对应计划编辑器、共享 Excel 能力和样式，
不是一个可删除的偶然依赖。应在这些功能定稿后更新基线；如果产品出现“首次打开慢”的
真实反馈，再以网络传输、解析/执行和 Excel WebView 首屏三段数据决定是否拆 `excel.ts`。
