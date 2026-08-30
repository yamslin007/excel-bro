# 渲染性能分析

分析日期：2026-08-30

## 结论

当前没有发现稳定、可复现且超过 16ms 的常规交互 commit，因此本轮不进入阶段 3，
不为消息流、范围选择器或计划编辑器增加 `React.memo`、虚拟列表或额外状态层。

最值得继续观察的是“我的工具”上限场景：50 个工具首次展开的最大 commit 中位数为
12.61ms，低于 16ms，但 12 次中有 4 次超过阈值。受当前自动化环境的 GC 和调度
抖动影响，这些长尾还不能直接归因为 React 组件；若低配设备或 Excel WebView 中有
实际卡顿反馈，应优先复测这一场景。

## 测量方法与限制

- 生产构建：`npm run build:addin`，Vite 7.3.6 构建成功，用时 1.14s。
- 组件采样：React 19 Profiler + jsdom，直接挂载当前 `App` 和
  `PlanEditorDialog`，每个场景预热 1 次、记录 12 次。
- 上限数据：100 条会话消息、50 个已保存工具；这两个规模分别来自
  `config/capabilities.json` 的现行上限。
- 网络请求保持 pending，数据只覆盖本地 React render/commit，不混入模型、FastAPI、
  Excel COM 或 Office.js 耗时。
- `actualDuration` 由 Profiler 采集；墙钟时间包含测试 DOM 更新和调度。
- 当前 Codex Windows 沙箱持续报 `helper_unknown_error: setup refresh had errors`，
  in-app browser / React DevTools 无法启动。因此本报告提供可重复的 commit 时序图，
  不是浏览器 DevTools 火焰图；也不包含真实 WebView 的布局、绘制和 GPU 时间。

原始结果见 [results.json](performance/render-profile/results.json)，采样脚本见
[`render-profile.tsx`](../apps/excel-addin/scripts/render-profile.tsx)。

## 关键场景

阈值按单次 commit 的 16ms 判断；“总 actual”可能包含多个 commit，不能直接作为掉帧
结论。中位数反映稳定路径，`>16ms` 同时保留长尾信息。

| 场景 | 数据规模 | commit 数中位 | 最大 commit 中位 | total actual 中位 | 墙钟中位 | >16ms 采样 |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| 打开插件 | 100 消息，工具箱关闭 | 3 | 12.15ms | 18.41ms | 38.18ms | 3/12 |
| 发送消息 | 100 条后新增并裁剪 | 2 | 3.20ms | 3.20ms | 10.15ms | 3/12 |
| 切换工作表 | 100 消息、2 张演示表 | 1 | 2.41ms | 2.41ms | 4.58ms | 1/12 |
| 打开“我的工具” | 50 个工具 | 1 | 12.61ms | 12.61ms | 15.78ms | 4/12 |
| 打开计划编辑器 | 50 个白名单动作 | 2 | 5.65ms | 5.82ms | 14.07ms | 0/12 |

采样时序图：

- [打开插件](performance/render-profile/initial-render.svg)
- [发送消息](performance/render-profile/send-message.svg)
- [切换工作表](performance/render-profile/switch-worksheet.svg)
- [打开“我的工具”](performance/render-profile/open-tools.svg)
- [打开计划编辑器](performance/render-profile/open-plan-editor.svg)

### 长尾解释

计划编辑器 12 次全部低于 16ms，而其他场景偶发 29–61ms 的孤立峰值；峰值会在不同
场景、不同轮次迁移，并非随列表规模单调增加。这与同一 Node 进程反复挂载/卸载大型
DOM 时的 GC/调度暂停一致。它说明存在分配压力，但在真实 Excel WebView 复测前，不足以
证明某个组件需要改造。报告不隐藏这些峰值，也不据此声称已经满足真实设备 P95。

## 频繁重渲染

连续输入“分析当前工作表”7 个字符会产生 7 次 `App` commit，符合当前受控输入的
状态结构。在 100 条历史消息下：

- 7 次 commit 的最大 commit 中位数：2.75ms；
- 7 次 total actual 中位数：16.02ms；
- 整个输入突发墙钟中位数：30.26ms；
- 12 次中有 1 次单 commit 超过 16ms。

因此快速打字时确实可能超过 5 次重渲染/秒，但单次稳定成本远低于一帧，不构成当前
瓶颈。对应时序图见 [typing-burst.svg](performance/render-profile/typing-burst.svg)。

## 组件与计算审计

### 消息流

`App.tsx` 直接执行 `messages.map`，没有虚拟化，也没有独立的 memo 消息组件。当前上限
为 100 条，发送场景的稳定最大 commit 仅 3.20ms；现在引入虚拟化会增加滚动定位、动态
卡片高度和无障碍维护成本，收益不足。

### 工具列表

工具箱一次渲染最多 50 个卡片，是最接近阈值的路径。首次展开需要挂载全部卡片，
`React.memo` 对首次挂载无效；分页或虚拟化会改变用户体验。现阶段只列为复测优先级 1，
不直接实施。

### 范围选择器

`filteredWorksheets` 已用 `useMemo`，切换选中状态的稳定最大 commit 为 2.41ms。选择判断
目前是数组 `includes`，即使工作表数增加也应先在真实大工作簿复测，再考虑改为 `Set`。

### 计划编辑器

JSON 校验、高亮和行号均在 `JsonCodeEditor` 中用 `useMemo` 缓存；50 个动作场景的最大
commit 中位数为 5.65ms，且没有超过 16ms 的采样。当前无需更换编辑器或延迟高亮。

### memo、回调与上下文

- `App` 有 5 处 `useMemo`；各业务 Hook 和计划编辑器对可复用操作使用了
  `useCallback`。
- `npm run lint -- --quiet` 通过，没有发现 Hooks 依赖错误。
- 项目没有 React Context，不存在 Context 更新引发全树渲染的问题。
- 当前没有 `React.memo`。这不是缺陷：主要列表的稳定 commit 都低于阈值，盲目 memo
  会增加 props 比较和维护成本。

## 瓶颈与优先级

| 优先级 | 路径 | 证据 | 决策 |
| --- | --- | --- | --- |
| P1 复测 | 50 工具首次展开 | 中位 12.61ms；4/12 长尾超 16ms | 在 Excel WebView/低配机复测；确认稳定超阈值后再分页或虚拟化 |
| P2 观察 | 100 消息初始加载 | 最大 commit 中位 12.15ms；3/12 长尾 | 暂不改；真实首屏卡顿时再拆消息流并采集火焰图 |
| P3 不优化 | 快速输入、发送、切表、计划编辑器 | 稳定最大 commit 2.41–5.65ms | 收益预计 <10ms，不实施 |

## 阶段 3 决策

本轮不修改运行时代码。稳定数据没有越过 16ms，任何预期收益都低于任务设定的 10ms
门槛；对 50 项工具做虚拟化或拆状态属于证据不足的过度工程化。

后续触发条件：用户在 Excel 中明确反馈工具箱或长对话卡顿，且 React DevTools 在同一
交互连续 3 次记录到 >16ms commit。届时优先验证工具卡片分页/虚拟化，而不是先给整个
`App` 添加 memo。
## 验证结果

- `npm run lint -- --quiet`：通过。
- `npm run build:addin`：通过，87 个模块，最终构建 1.04s。
- `python -m pytest server/tests -q`：把 pytest 临时目录改到工作区后，289 项通过。
- `python -m compileall -q server/app`：通过。
- `npm run test:addin`：335 项中 332 项通过、3 项失败。失败位于既有
  `crossTableFormula.test.ts`（2 项外部工作簿引用引号不一致）和
  `useScopeSelection.test.ts`（1 项新目录应用后的选区重置断言），与本轮新增的性能
  采样脚本和文档无关。本轮遵守任务边界，没有顺手修改这些业务行为。