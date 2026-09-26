# 阶段 G 基线记录

> 日期:2026-09-26
> 依据:[双宿主改造方案 v2.1](./DUAL_HOST_REFACTOR_PROPOSAL_2026-09-26.md) 第 3 节与第 7 节
> 性质:**只新增测试与记录,未改动任何生产代码**

## 1. 本阶段的判定标准

阶段 G 不要求全绿。要求的是**可信基线**:

- 已具备的底层能力 → 正常通过
- 尚未实现的能力 → 预期失败,且**失败原因被锁定**
- 已知缺陷 → 预期失败,且**当前实际行为被正向断言钉住**

"锁定原因"是关键。仅仅"失败了"不构成基线 —— 必须保证它是因为文档说的那个原因
失败。因此每条预期失败都配了一条断言当前实际行为的正向测试;缺陷一旦变化,
正向测试会先报错,而不是让预期失败悄悄换个理由继续红着。

## 2. 自动化基线:文件夹通道(openpyxl)

`server/tests/test_column_mapping_baseline.py`,6 条。

| 测试 | 覆盖断言 | 结果 |
|---|---|---|
| `test_a7_a9_write_and_undo_round_trip[baseline-D-to-E]` | A7 持久化重读一致、A9 撤销回到执行前且源未被触碰 | PASSED |
| `test_a7_a9_write_and_undo_round_trip[variant-B-to-H]` | 同上,换源列与目标列 | PASSED |
| `test_a4_authorization_does_not_expand` | A4 授权集合不含未选表 | PASSED |
| `test_a12_target_changed_between_preview_and_execute` | A12 预览与执行之间源被改动 → 拒绝且未写入 | PASSED |
| `test_a3_column_mapping_plan_is_generated_from_parameters` | A3 参数化映射生成单个 `writeValues` | **XFAIL** |
| `test_b1_unmatched_category_requires_clarification` | B1-B3 未匹配类别必须澄清不猜测 | **XFAIL** |

两条 XFAIL 用 `strict=True, raises=ImportError`,锁定在
`server.app.column_mapping` 尚不存在。阶段 4 实现后会 XPASS,
`strict=True` 使 XPASS 报错,迫使移除标记 —— 不会静默转绿。

### 2.1 参数化的验证方式

`D→E` 与 `B→H` 是同一实现的两组参数,期望值由 `_expected_for(mapping, blank_label)`
从映射表推导,不硬编码。映射表、源/目标列、工作表名全部是夹具参数。
这条约束是为了防止"测试通过,产品却只会 D→E"。

### 2.2 两条断言的修正记录

初版有两条断言写错了,核实后是我的预期不对,不是产品缺陷:

1. **A4 授权名格式** —— 文件夹通道用「文件 › 工作表」限定名
   (`orders.xlsx › 订单`),而不是裸工作表名。这是既有设计,
   目的是避免同名表跨文件歧义,见 `docs/ARCHITECTURE.md` 工作表授权边界一节。
2. **A12 拒绝文案** —— 实际拒绝原因是"预览中所选文件已发生变化,
   请重新生成预览后再执行",不含"指纹"字样。拒绝行为本身是对的,
   断言改为匹配实际语义,并追加"目标列仍为空"确认确实没有写入。

## 3. 已知缺陷基线:空串写入语义(WPS 通道)

`apps/wps-addin/src/wpsEmptyStringBaseline.test.ts`,5 条,全部按预期。

| 测试 | 性质 |
|---|---|
| 缺陷 1:`prepareWpsPlan` 应保留计划中的空串 | `it.fails` |
| 缺陷 2:无法保真空串的宿主不应报告 `verified` | `it.fails` |
| 锁定缺陷 1 的确切表现:空串在准备阶段变成 `null` | 正向断言当前行为 |
| 锁定缺陷 2 的确切表现:宿主折叠空串后仍报告 `verified` | 正向断言当前行为 |
| 参照组:不含空串的写入在同一宿主正常验证通过 | 正向 |

`it.fails` 在断言变为通过时会报错,与 pytest `strict=True` 语义一致。

### 3.1 缺陷比方案 §1.1 描述的更早一层

方案 §1.1 定位在 `matricesEqual`(`wpsPlanExecutor.ts:147-149`)双侧折叠导致
回读核对失明。核实后发现这是**第二层**,前面还有一层:

`valuesForAction`(`wpsPlanExecutor.ts:195`)在**准备阶段**就对计划值调用
`canonicalCell`,`""` 在任何宿主交互之前已变成 `null`。也就是说写入意图在
预检阶段就丢了,根本没走到回读这一步。

这不改变修复方向,但**修复点是两处而非一处**:

| 层 | 位置 | 问题 |
|---|---|---|
| 准备 | `wpsPlanExecutor.ts:195` `valuesForAction` | 计划里的 `""` 被折成 `null`,写入意图丢失 |
| 验收 | `wpsPlanExecutor.ts:147-149` `matricesEqual` | 双侧折叠,`""` 与 `null` 差异不可见 |

只修第二层不够 —— 准备阶段已经把 `""` 变成 `null`,回读时 expected 本身就是
`null`,比对自然通过。阶段 4 必须两层一起修。

### 3.2 空串不作跨宿主持久化断言

已核实 openpyxl 把 `""` 写入 `.xlsx` 后保存重载会读回 `None`,
因此不能要求三条通道都把持久化后的 `""` 与空单元格区分开。
文件夹用例第 5 行改为**空白单元格 + 「空白→未分类」映射策略**,
空串误判问题由本节的 WPS 负向测试单独固定,不藏在正向用例里。

## 4. 两宿主实机清单(待执行)

文件夹通道可全自动,两个宿主的真实写入需实机。以下清单在阶段 4 执行,
按 `docs/WPS_W*_VALIDATION_*.md` 既有格式记录实际读回值。

### 4.1 共用前置

- 临时工作簿,不使用真实业务数据
- 记录宿主完整版本号
- 每条记录实际读回值,不只记"通过"

### 4.2 清单

| # | 步骤 | Excel 期望 | WPS 期望 |
|---|---|---|---|
| M1 | 打开夹具工作簿,授权单表 | 授权集合仅该表 | 同 |
| M2 | 给出含映射表的指令 | 进入 `plan`,单个 `writeValues` | 同 |
| M3 | 不确认,检查工作簿 | 无任何写入 | 同 |
| M4 | 编辑预览中任一行 | 确认绑定编辑后计划 | 同 |
| M5 | 确认执行 | 目标列 5 行与预览一致 | 同 |
| M6 | 回读核对 | 通过,记录实际值 | 同 |
| M7 | 撤销 | 目标回执行前,源未动 | 同 |
| M8 | 执行后外部改动再撤销 | 拒绝覆盖 | 同 |
| M9 | 换目标列重跑 M2-M7 | 行为一致 | 同 |
| M10 | 不给映射表 | 返回澄清,不生成写入 | 同 |

WPS 侧另需:M11 未验收版本(如 12.0)应明确拒绝并说明原因,不出现外观相同的空入口。

## 5. 阶段 G 退出状态

| 退出条件 | 状态 |
|---|---|
| 现有通过项仍绿 | 达成,见第 6 节 |
| 参数化用例预期输出与授权边界已冻结 | 达成 |
| 新能力缺口有可重复失败证据 | 达成,2 条 strict XFAIL |
| `""`/`null` 误判有可重复失败证据 | 达成,2 条 `it.fails` + 2 条行为锁定 |
| 两宿主实机基线记录归档 | **清单已定,实机待执行** |
| 未改动生产代码 | 达成 |

最后一项尚未完成,因为实机验证需要目标版本的 WPS 与 Excel 环境。
自动化基线已可支撑阶段 1 与阶段 2(它们都不触碰写入语义)。
**阶段 4 开始前必须补齐实机记录** —— 那一阶段才真正改动写入与验收语义。

## 6. 全量回归

| 套件 | 结果 |
|---|---|
| `npm run test:addin` | 474 passed |
| `npm run test:wps` | 57 passed(含新增 5 条) |
| `python -m pytest server/tests --basetemp=.tmp-pytest` | 455 passed,2 xfailed |

## 7. 阶段 1 的一处方案更正

核实依赖现状时发现方案 v2.1 §阶段 1 的注记有误,已在该文档内更正:

| 位置 | React | 独立 lock | 独立 `node_modules` |
|---|---|---|---|
| 根 | 19.2.8 | 有 | 有 |
| `apps/wps-addin` | 19.1.1 | 有,且已被 git 跟踪 | 有,237 个包 |

双份 React 是**当前就存在的事实**,不是"someone runs npm install 才触发"的潜在风险。
`package.json` 的 `test:wps`/`build:wps` 用 `npm --prefix` 而非 `--workspace`,
印证 WPS 从一开始就是独立安装的子项目。

因此 v1 写的"删 `apps/wps-addin/package-lock.json`"这一步**是必要的**,
v2 删掉它是基于错误核实。阶段 1 必须删该 lock(git 中一并删除)、
删 `apps/wps-addin/node_modules`、纳入根 workspaces、脚本改 `--workspace`、根目录重装,
退出闸 `npm ls react` 只剩一个版本。
