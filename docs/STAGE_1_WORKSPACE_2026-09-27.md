# 阶段 1 执行记录：workspace 统一与 React 对齐(2026-09-27)

## 1. 目标与退出闸

把 `apps/wps-addin` 从独立安装的子项目收进根 workspaces,消除双份 React,
为阶段 2 起 `packages/core` 携带 React hooks 做前置准备。
退出闸:`npm ls react` 只剩一个版本,且两端构建与测试全绿、阶段 G 的预期失败数量与原因不变。

## 2. 实际改动

| 文件 | 改动 |
|---|---|
| `package.json` | workspaces 加入 `apps/wps-addin` 与 `packages/*`;`build:wps`/`test:wps`/`dev:wps`/`publish:wps` 由 `npm --prefix` 改 `--workspace excel-bro-wps`;新增 `test:core`;`test:server` 带 `--basetemp=.tmp-pytest`;`test` 串起三端;`lint` 覆盖 `packages/core/src` |
| `apps/wps-addin/package-lock.json` | 删除(git 中一并删除),依赖解析统一由根 lock 负责 |
| `apps/wps-addin/node_modules` | 删除(237 个包),重装后由根目录 hoist |
| `apps/wps-addin/package.json` | `react`/`react-dom` 由精确 `19.1.1` 改 `^19.1.1`,与 excel 端一致 |
| 两端 `package.json` | `@types/office-js` 统一 `^1.0.600`(取高) |
| 两端 `tsconfig.json` | 新增 `paths` 指向 `packages/core/src` |
| 三个 vite 配置 | 新增 `resolve.alias` 指向 `packages/core/src` 与 `resolve.dedupe: ["react","react-dom"]` |
| `packages/core/` | 新建骨架:`package.json`(React 仅作 peerDependency)、`tsconfig.json`、`vitest.config.ts`、空 `src/index.ts` |
| `.gitignore` | 忽略 `.tmp-npm-cache/` |

## 3. 两处需要记录的偏差

### 3.1 `@types/office-js` 不能从 WPS 端移除(阶段 6 才能移除)

原判断:`apps/wps-addin/src` 内 `Office.` 与 `Excel.` 引用数为 0,故 WPS 不需要 Office 类型。
这个 grep 结论本身没错,但**作用域只到 WPS 自己的源码**。移除后 `tsc --noEmit` 报 9 条错误:

```
../excel-addin/src/contracts.ts(1120,14): error TS2503: Cannot find namespace 'Excel'.
../excel-addin/src/dataTools.ts(295,14): error TS2503: Cannot find namespace 'Excel'.
../excel-addin/src/dataTools.ts(324,12): error TS2503: Cannot find namespace 'Excel'.
../excel-addin/src/dataTools.ts(910,10): error TS2304: Cannot find name 'Excel'.
（另有 5 条由上述缺失类型派生的 TS7006 隐式 any）
```

WPS 反向导入的 `contracts.ts` 与 `dataTools.ts` 自身引用 `Excel` 命名空间,
类型依赖是**传递性**的。这正是本次重构要消除的耦合:WPS 不用 Office API,
却必须装 Office 类型才能编译。

处置:恢复 `@types/office-js`(统一到 `^1.0.600`),在 `apps/wps-addin/tsconfig.json`
的 `types` 上留注释说明它是反向导入的产物,阶段 6 清掉 30 处 `../../excel-addin/src/...`
之后再移除。**不在阶段 1 提前拆**——阶段 1 的契约是"不移动任何代码"。

### 3.2 npm 全局缓存目录不可写

`npm install` 首次失败:

```
npm error code EPERM
npm error syscall open
npm error path D:\AdvancedSoftware\Node.js\install\node_cache\_cacache\tmp\bcc63048
npm error   requiredBy: 'apps/wps-addin'
```

与本项目 pytest 需要 `--basetemp` 指向项目内目录是同一类 Windows 权限问题。
改用 `npm install --cache .tmp-npm-cache` 通过,该目录已加入 `.gitignore`。
后续在本机重装依赖需沿用此参数。

## 4. 退出闸与回归

```
$ npm ls react react-dom
excel-bro@0.1.0 E:\codex\excel-bro
+-- @excel-bro/addin@0.1.0 -> .\apps\excel-addin
| +-- @testing-library/react@16.3.2
| | +-- react-dom@19.2.8 deduped
| | `-- react@19.2.8 deduped
| +-- react-dom@19.2.8
| | `-- react@19.2.8 deduped
| `-- react@19.2.8
+-- @excel-bro/core@0.1.0 -> .\packages\core
| +-- react-dom@19.2.8 deduped
| `-- react@19.2.8 deduped
`-- excel-bro-wps@0.1.0 -> .\apps\wps-addin
  +-- react-dom@19.2.8 deduped
  `-- react@19.2.8 deduped
```

单一 `19.2.8`,全部 deduped;`apps/wps-addin/node_modules` 与
`apps/excel-addin/node_modules` 均已不存在。`@types/office-js` 实际安装 `1.0.600`。

| 命令 | 结果 |
|---|---|
| `npm run lint` | 0 errors,11 warnings(全部为既有 `react-hooks/exhaustive-deps`,数量未变) |
| `npm run build:addin` | 通过(97 modules + functions bundle 3 modules) |
| `npm run build:wps` | 通过(62 modules) |
| `npm run test:addin` | 47 files,474 passed |
| `npm run test:wps` | 8 files,57 passed |
| `npm run test:server` | 457 passed,3 xfailed |

阶段 G 的预期失败**数量与原因均未变**:

- `wpsEmptyStringBaseline.test.ts` 5 tests 全过,即 2 条 `it.fails` 仍按原因失败
  (vitest 中 `it.fails` 若意外通过会报 XPASS 失败)
- `test_column_mapping_baseline.py` 6 passed,3 xfailed,3 条 XFAIL 原因行与阶段 G 记录一致

## 5. 本阶段未做的事

- 未移动任何源码文件,`packages/core/src/index.ts` 仍是空导出
- 未新增或修改任何生产逻辑
- 未触碰服务端代码(457 passed 用于佐证)
- `docs/STAGE_G_BASELINE_2026-09-26.md` §7 的核实结论在本次执行中得到确认:
  WPS 确有被 git 跟踪的独立 lock 与 237 个包的 `node_modules`,删除是必要步骤
