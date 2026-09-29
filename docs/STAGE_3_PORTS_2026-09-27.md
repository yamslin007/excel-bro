# 阶段 3 执行记录：L3 端口、WorkbookLocator 与 SavedTool 下沉(2026-09-27)

## 1. 目标

按方案 §2.3 第 1 条与 §5.2:

1. 引入共享层内部判别式 `WorkbookLocator`,身份折叠收口到唯一一处
2. 定义 L3 端口接口(`HostAdapter` / `KeyValueStore` / `ApiClient` / `Clock`),只定接口不写实现
3. 把 `SavedTool` 纯 DTO 下沉到 L0,解除 `format.ts`(L1)对 `storage.ts`(L2)的反向类型边

**不改现有 Office 恢复协议。** 出网字段 `workbookIdentity` 仍是 `url:` 字符串。

## 2. 实际改动

| 变更 | 文件 | 行数 |
|---|---|---|
| L0 追加 `WorkbookLocator`、`AuthorizedSheetRef` | `packages/core/src/contracts.ts` | +22 |
| L0 新建工具 DTO | `packages/core/src/tools.ts` | 122 |
| L2 新建身份折叠 | `packages/core/src/locator.ts` | 78 |
| L3 新建端口接口 | `packages/core/src/ports.ts` | 119 |
| 折叠规则测试 | `packages/core/src/locator.test.ts` | 124(14 条) |
| `storage.ts` DTO 改为再导出 | `apps/excel-addin/src/storage.ts` | −94 |
| `format.ts` 改从 core 取 `SavedTool` | `apps/excel-addin/src/format.ts` | |
| Excel 身份构造改走 core | `apps/excel-addin/src/excel.ts` | |
| WPS 身份判别式 + 测试 | `apps/wps-addin/src/wpsPlanExecutor.{ts,test.ts}` | +3 条测试 |

## 3. 四个判断与理由

### 3.1 `SavedTool.planTemplate` 泛型化,而不是让 L0 依赖 `AnalysisPlan`

`SavedTool` 其余字段都是纯标量,唯一的外部依赖是 `planTemplate: AnalysisPlan`。
`AnalysisPlan` 挂在 `ExcelAction` 这个 38 成员联合上,连带 1000 余行,
方案 §5.4 明确不从 `contracts.ts` 开刀。

所以 core 里是 `SavedTool<TPlanTemplate = unknown>`,app 侧:

```ts
export type SavedTool = CoreSavedTool<AnalysisPlan>;
```

结构完全一致,13 处使用点零改动。`AnalysisPlan` 将来迁入 L0 时可收紧默认值。

### 3.2 迁移时发现两处与原定义不符,已按原文修正

写 `tools.ts` 时有两处偏离了 `storage.ts` 的原声明,typecheck 前就核对出来:

| 成员 | 我先写成 | 原文 |
|---|---|---|
| `OutputWorksheetToolParameter.bindings[].property` | `"sheet"` | `"sheet" \| "sourceSheet"` |
| `FieldToolParameter.sourceParameterId` | 可选 | **必填** |

两者都会放宽类型。已用 `git show HEAD:apps/excel-addin/src/storage.ts` 取出原
17-111 行与迁入后的块做 `diff`,除一处尾随空行外**逐字一致**。

### 3.3 非 Office 通道折叠时抛错,不返回伪身份

```ts
if (locator.kind !== "office-url") {
  throw new Error(`仅 Office 通道可折叠为 workbookIdentity，收到 kind=${locator.kind}`);
}
```

WPS 恢复按方案 §2.3 第 2 条保持客户端权威,服务端不参与。
若某处把 `wps-path` 塞进 `/api/office/...`,那是调用方分流错了。
拼一个能通过服务端 `url:` 校验的假串,会让两条通道的现场混进同一个账本 ——
这正是方案反复强调的"不在两套语义之间架桥"。反向解析同理:
缺前缀返回 `null` 而不是猜成 `office-url`,否则后续身份比对会静默通过。

### 3.4 `HostAdapter` 不含撤销成员

三条通道的撤销机制与粒度全不同(方案 §2.2):服务端账本 / `PluginStorage` /
文件级副本;区域 / 区域 / 整文件原子替换。塞进一个接口就是架桥。
撤销按通道各自定义,阶段 4 再决定共享的是哪一层。

同理 `runBatch` 的注释里写明**不承诺整体原子性**:Office 有 `Excel.run`,
WPS 没有等价物,顺序执行到一半失败就是半写状态。跨宿主一致的是"半写可恢复",
靠检查点与写前快照兜底,不是假装有事务。

## 4. 身份归一:两端各自的既有形态都没动

折叠规则逐字复制 `excel.ts:3766`,不是重写:

```ts
`url:${trimmed.replace(/\\/g, "/").toLocaleLowerCase()}`
```

`excel.ts` 的 `currentOfficeWorkbookIdentity` 改为调用 core,不再自行拼前缀。
原先是同一条规则的第二份实现,漂移的后果是已落盘账本里的旧记录身份对不上,
在途任务失去续跑与撤销入口。

WPS 侧**不统一**归一规则:`workbookIdentity`(`wpsPlanExecutor.ts:161`)只小写、
保留反斜杠,已落盘的 `WpsUndoRecord.workbookIdentity` 就是这个形态。
`wpsWorkbookLocator` 沿用它,新增测试把 locator 的 `value` 与执行路径写入撤销记录
的身份锁在一起。core 的 `sameLocator` 比对 `wps-path` 时对两种分隔符都归一,
所以比对侧不受影响。

`locator.test.ts` 另锁了一条容易踩的:值里再含 `url:` 时用 `slice(prefix.length)`
而非 `replace`,防止有人改成 `replaceAll` 后静默截断路径。

## 5. 回归

| 套件 | 阶段 2 | 阶段 3 |
|---|---|---|
| `npm run test:core` | 3 files,14 passed | **4 files,28 passed** |
| `npm run test:addin` | 46 files,474 passed | 46 files,**474 passed** |
| `npm run test:wps` | 8 files,57 passed | 8 files,**60 passed** |
| `npm run test:server` | 457 passed,3 xfailed | **457 passed,3 xfailed** |
| `npm run lint` | 0 errors,11 warnings | **0 errors,11 warnings** |
| `npm run build:addin` | 通过(99 modules) | 通过(**100 modules**) |
| `npm run build:wps` | 通过(64 modules) | 通过(**64 modules**) |

core +14 条 = `locator.test.ts`。WPS +3 条 = `wpsWorkbookLocator`。
addin 用例数不变:DTO 下沉是纯类型改动,无行为变化。

模块数 +1 是 `locator.ts` 进入 addin 依赖图。`ports.ts` 与 `tools.ts` 是纯类型,
不进产物;WPS 只 import 了 `WorkbookLocator` 类型,故 64 不变。

阶段 G 预期失败**数量、节点 id、原因均未变**:三条 XFAIL 仍是
`test_a3_...[same-sheet-D-to-E]`、`[cross-sheet-B-to-H]` 与
`test_unmatched_source_value_does_not_get_guessed`。

### 5.1 一次未能复现的服务端偶发失败

首轮 `npm run test:server` 出现 `1 failed, 456 passed, 3 xfailed`,失败节点是
`test_a7_a9_write_and_undo_round_trip[baseline-D-to-E]`。随后**单文件 1 次、
定向 5 次、全量 4 次**均未复现(全部 457 passed / 3 xfailed)。

**失败原因未查明。** 首轮只抓了汇总行,没留失败堆栈,复现尝试又全部通过,
所以现在说不出它因何失败。

已核实的只有两点:本阶段未触碰服务端任何文件;该用例走 openpyxl 的
文件级原子替换。同轮另有 `.pytest_cache` 的 `WinError 5`,但那是 pytest 缓存目录
的权限问题,**不构成同因证据** —— 两者都发生在 Windows 文件系统上,
仅此而已,不足以推断是同一个句柄/权限争用。

**未修复,如实登记。** 若重现,第一件事是取失败堆栈与 `-l`(局部变量),
按文件夹通道撤销路径的独立问题排查,不要归到阶段 3,也不要沿用上面那条
未经证实的猜测。

## 6. 本阶段未做的事

- `ports.ts` 只定义阶段 4 需要的原语(`setValues`/`readValues`/`readFormulas`/
  `clearContents`)。没有预先铺 38 个动作:无使用点的接口无法被测试约束,
  只会在阶段 5 逐动作推进时被改写
- 未写任何适配器实现。两端仍走各自现有函数,只有身份构造收口到 core
- `format.ts` **未迁入 core**。反向边已解除,但 `actionLabel`(38 个动作的中文
  展示串)的 L1/L5 归属仍未定,方案 §5.4 要求先判断再迁
- 未触碰 `utils.ts`(等阶段 5 拆 `dataTools` 的错误类型)
- 未改服务端、未改 `capabilities.json`

## 7. 交给后续阶段的事项

| 事项 | 归属 |
|---|---|
| 写两端 `HostAdapter` 实现,`readSheets`/`runBatch` 接原语 | 阶段 4 |
| `format.ts` 迁入 core 前定 `actionLabel` 的层归属 | 阶段 4/5 |
| 能力协商若新增请求字段,同步 `contracts.ts` 与 `models.py` 并加拒绝伪造集合测试 | 阶段 4 |
| `DataToolExecutionError` 提到 core,`utils.ts` 随后可迁 | 阶段 5 |
| `AnalysisPlan` 迁入 L0 后收紧 `SavedTool` 的泛型默认值 | 阶段 5/6 |
| 删 shim 与 `coreShims.test.ts`,WPS 改 `@excel-bro/core/...` | 阶段 6 |
