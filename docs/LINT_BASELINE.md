# Lint 基线（2026-08-17）

## 配置现状

- 配置格式：ESLint 8 legacy config（`.eslintrc.cjs`）
- 解析器：`@typescript-eslint/parser`
- Hooks 规则：
  - `react-hooks/rules-of-hooks`: error
  - `react-hooks/exhaustive-deps`: warn
- 检查范围：`apps/excel-addin/src`（仓库实际前端路径；原任务描述中的
  `client/src` 不存在）

## 依赖版本

- eslint：8.57.1
- eslint-plugin-react-hooks：4.6.2
- @typescript-eslint/parser：8.67.0

## 检查命令

```powershell
npm run lint
```

根目录 `package.json` 已新增对应脚本。

## 首次全量结果

```text
2 problems (0 errors, 2 warnings)
```

### warning 清单

1. `apps/excel-addin/src/App.tsx:1302`
   - 规则：`react-hooks/exhaustive-deps`
   - 缺失依赖：`clearUndoSnapshot`, `scan`, `setSelectedSheetNames`,
     `setSelectionConfirmed`
   - 判定：需要人工确认的真实提示。该 effect 是 `Office.onReady` 的
     mount-only 初始化逻辑，当前刻意使用空依赖；若直接补齐依赖，会在相关
     callback 变化时重复注册 watcher，可能引入新的竞态。
   - 建议：接入更稳定的初始化方案后，再决定是否用 ref 或重构 effect。

2. `apps/excel-addin/src/hooks/useActivityProgress.ts:122`
   - 规则：`react-hooks/exhaustive-deps`
   - 缺失依赖：`activity`
   - 判定：偏误报/有意窄化依赖。计时 effect 只读取
     `activity.startedAt`，依赖数组使用 `activity?.startedAt` 可避免
     `advanceActivity` 更新步骤时反复重建 interval。
   - 建议：如需消除告警，可先把 `startedAt` 提取为局部常量，再让 effect
     仅依赖该常量；当前不做业务修改。

## 结论

- 没有 `react-hooks/rules-of-hooks` 的 error 级违规。
- 两条 exhaustive-deps 均为 warn，暂不批量修复，符合“先建立基线”的阶段目标。
