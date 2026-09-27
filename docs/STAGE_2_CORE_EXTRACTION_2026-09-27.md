# 阶段 2 执行记录：抽 L1 纯核心 + 兼容 shim(2026-09-27)

## 1. 目标

按方案 §5.4 抽取 L1 层,用它验证 workspace、alias、测试与构建链路真的通。
原路径留再导出 shim,使 App.tsx 与 WPS 的导入点零改动。

## 2. 方案 §5.4 清单与实测依赖面的差距

§5.4 写"`planCore` + `format` + `range` + `utils` + `tableSchema`,共约 400 行,
**零环境依赖,有现成测试**"。逐个核实后,这句话有三处与实际不符:

| 模块 | 行数 | 实测依赖 | 自带测试 | WPS 是否导入 | 本阶段 |
|---|---|---|---|---|---|
| `planCore.ts` | 61 | 仅 `import type CellValue` | 有(28 行) | **是** | 已迁 |
| `range.ts` | 54 | **零 import** | **无** | 否 | 已迁,并补测试 |
| `tableSchema.ts` | 51 | `type CellValue` + 直接 import `capabilities.json` | 有(26 行) | **是** | 已迁 |
| `format.ts` | 142 | `type CellValue` + **`type SavedTool` from `storage.ts`** | 无 | 否 | **未迁** |
| `utils.ts` | 104 | **运行时 import `DataToolExecutionError` from `dataTools.ts`** + `type ModelConnectionDraft` from `hooks/` | 无 | 否 | **未迁** |

三处差距:

1. **只有 2/5 自带测试**,不是"有现成测试"。`format`/`range`/`utils` 没有测试文件,
   且 grep 确认 `suggestWriteTarget`/`columnLetterFromIndex`/`firstCellOfRange`/
   `startColumnOfRange` 在整个 addin 测试套件里零引用——迁移前无任何回归网。
2. **`utils.ts` 不是零环境依赖**,详见 §3.1。
3. **只有 `planCore` 与 `tableSchema` 是 WPS 真正共用的**(各 1 处导入)。
   `format`/`range`/`utils` WPS 一处都没用,`utils.ts` 仅 App.tsx 单一消费者。

## 3. 两个模块本阶段不迁的原因

阶段 2 的契约是"机械移动 + 留 shim",不改生产逻辑。下面两个模块要迁必须先改代码,
因此各自归到对应阶段,而不是在本阶段夹带。

### 3.1 `utils.ts` —— 运行时边指向 Office 耦合模块

```ts
import { DataToolExecutionError } from "./dataTools";   // 运行时导入，非 type-only
```

`isAbortError`(`utils.ts:78-89`)用 `reason instanceof DataToolExecutionError` 判断,
这是**值**依赖。而 `dataTools.ts` 自身直接用 `Excel` 命名空间
(`:295 Excel.Worksheet`、`:324 Excel.RequestContext`、`:910 Excel.run`)。
把 `utils.ts` 原样搬进 core,会顺着这条边把 Office 依赖拖进宿主无关层——
正是本次重构要消除的东西。

另有两处次要问题:`import type ModelConnectionDraft from "./hooks/useModelManagement"`
是 L1→L4 反向边;`copyTextToClipboard` 用 `navigator.clipboard`/`document`,
`CACHE_SESSION_ID`(`:30`)在模块加载时读 `globalThis.crypto` 并产生一次性副作用。
后两者是 DOM 而非宿主依赖,core 可接受,但说明这个文件本就不是纯函数集合。

**前置条件**:把 `DataToolExecutionError` 与 `DataToolErrorCode` 提到 core
(错误类型本身不碰 Office,只是当前和 Office 代码同住一个文件)。
这是拆 `dataTools.ts` 的一部分,归**阶段 5**。

### 3.2 `format.ts` —— 类型边指向 L2

```ts
import type { SavedTool } from "./storage";   // storage.ts 按 §5.3 是 L2
```

`workflowToolPresentation`(`:17`)的入参是 `SavedTool`。虽是 type-only、无运行时代价,
但 L1→L2 是反向边,迁入 core 会在第一刀就破坏 §5.2 的"依赖只能从上往下"。

`SavedTool` 是纯 DTO,按 §5.2 本应落在 L0。但 `storage.ts` 有 1000 行,
其中混着迁移与校验逻辑,拆它属于 §5.3 的 `storage` 处理范围,归**阶段 3/4**。
届时 `format.ts` 可整体迁入。

另注:`actionLabel`(`:51-142`)是 38 个动作的中文展示串,严格说属展示层而非 L1;
迁移时应考虑放 L5 而不是 L1。

## 4. 实际改动

| 变更 | 文件 |
|---|---|
| 新建 L0 类型种子 | `packages/core/src/contracts.ts`,目前仅 `CellValue` |
| `git mv` 迁入 core | `planCore.ts` + `planCore.test.ts`、`range.ts`、`tableSchema.ts` + `tableSchema.test.ts` |
| 原路径留 shim | `apps/excel-addin/src/{planCore,range,tableSchema}.ts` 改为从 `@excel-bro/core/*` 再导出 |
| `CellValue` 单一来源 | `apps/excel-addin/src/contracts.ts:1` 改为 `import type ... from "@excel-bro/core/contracts"` 后再导出 |
| core 出口 | `packages/core/src/index.ts` 导出 L0 + L1 全部符号 |
| 补测试 | 新建 `packages/core/src/range.test.ts`(迁移前该模块零测试) |
| shim 身份测试 | 新建 `apps/excel-addin/src/coreShims.test.ts` |

### 4.1 相对路径深度恰好一致,无需改写

`apps/excel-addin/src/X` 与 `packages/core/src/X` 距仓库根同为三层,
故 `tableSchema.ts` 里的 `../../../config/capabilities.json` 迁移后**指向同一文件**,
一个字都不用改。`./contracts` 也自然指向 core 自己的 L0 文件。

`tableSchema.ts` 仍直接 import `capabilities.json`。方案要求改为经 core 的读取器,
那属于 capabilities 重组,本阶段不动,只在此登记。

### 4.2 shim 用引用相等而非行为相似来锁

`coreShims.test.ts` 断言的是 `shim.fn === core.fn`(函数引用相等),
不是"两边行为一致"。若有人把 shim 改回本地实现,引用不再相同,测试立刻报红;
只比行为的话,一份复制粘贴的实现能悄悄通过。另有一条断言锁住导出符号集合,
防止 shim 漏导出某个符号却无人发觉。阶段 6 删 shim 时本文件一并删除。

### 4.3 `export type { X } from` 的一处踩坑

`contracts.ts` 首行先写成:

```ts
export type { CellValue } from "@excel-bro/core/contracts";
```

`tsc` 报 40 余条 `TS2304: Cannot find name 'CellValue'`。
再导出**不把名字引入本地作用域**,而 `contracts.ts` 内部有大量 `CellValue` 引用。
改为 `import type` + 单独 `export type { CellValue }` 后通过。
`vite build` 不报此错(只做转译,不做类型检查),是 `tsc --noEmit` 抓到的——
说明 `build` 脚本里 `tsc --noEmit &&` 这一段不能省。

## 5. 回归

| 套件 | 迁移前 | 迁移后 |
|---|---|---|
| `npm run test:core` | 不存在 | **3 files,14 passed** |
| `npm run test:addin` | 47 files,474 passed | 46 files,**474 passed** |
| `npm run test:wps` | 8 files,57 passed | 8 files,**57 passed** |
| `npm run test:server` | 457 passed,3 xfailed | **457 passed,3 xfailed** |
| `npm run lint` | 0 errors,11 warnings | **0 errors,11 warnings** |
| `npm run build:addin` | 通过(97 modules) | 通过(**99 modules**) |
| `npm run build:wps` | 通过(62 modules) | 通过(**64 modules**) |

addin 测试文件数 47→46:迁出 2 个(`planCore.test.ts`、`tableSchema.test.ts`),
新增 1 个(`coreShims.test.ts`)。用例数 474 不变:迁出 4 条、新增 4 条。
core 的 14 条 = 迁入的 4 条 + `range.test.ts` 新增的 10 条。

模块数 +2 是 core 的 `index.ts` 与 `contracts.ts` 进入依赖图。

阶段 G 预期失败**数量、位置、原因均未变**:

- `wpsEmptyStringBaseline.test.ts` 5 tests 全过(2 条 `it.fails` 仍按原因失败)
- `test_column_mapping_baseline.py` 6 passed + 3 XFAIL,三条 XFAIL 节点 id 与阶段 G 记录逐字一致

## 6. 本阶段未做的事

- 未改动任何生产逻辑,迁入 core 的三个文件内容逐字未变(仅位置变化)
- 未触碰 `contracts.ts` 除首行外的任何内容
- 未触碰服务端
- WPS 的 30 处 `../../excel-addin/src/...` 一处未改,其中 `planCore` 与
  `tableSchema` 两处现经 shim 转发到 core,阶段 6 再改为直接导入

## 7. 交给后续阶段的事项

| 事项 | 归属 |
|---|---|
| `SavedTool` 提到 L0,`format.ts` 随后整体迁入(`actionLabel` 考虑归 L5) | 阶段 3/4(拆 `storage.ts`) |
| `DataToolExecutionError` 提到 core,`utils.ts` 随后可迁 | 阶段 5(拆 `dataTools.ts`) |
| `tableSchema.ts` 改用 core 的 capabilities 读取器,不再直接 import JSON | capabilities 重组 |
| 删 shim 与 `coreShims.test.ts`,WPS 改 `@excel-bro/core/...` | 阶段 6 |
