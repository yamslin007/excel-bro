# EB 系统收缩执行计划

> 承接 `WORK_LOG_2026-08-08.md` 第 314 行待办。基于 2026-08-08 两轮讨论 + 全链路代码核查得出。
> 决策已定，实施尚未开始。

## 1. 核心决策

**否定"让普通用户自己造规则"这个前提。** 不是优化 EB 规则的实现，而是把"造规则"这件事从用户手里拿走。

- 用户原话要点：EB 规则对普通用户较难、泛化性不够、"很容易成为负担"。
- 违背既定原则"预制能力必须泛化"——现有设计把业务复杂度（规则/别名/参数类型/依赖）转嫁给了用户。

**新主路径**：对话内 `/function <描述>` **短链**生成原生 Excel 公式，直接写单元格、立即重算。
- 关键：绕开 planner 的规划/思维链（不走 `intent → create_plan` 重链），命中式单发生成。类比 `/commit`。
- `=EB(...)` 自定义函数**仅**兜正则清理等原生公式表达不了的场景。预制 CLEAN/EMPTY 保留。

## 2. 为什么慢 —— 根因确诊

| 根因 | 证据 | 性质 |
|------|------|------|
| 别名函数走 functions.json，必须重启 Excel | RuleCreator.tsx:188/279/395 明写"重启后可用" | 慢反馈头号元凶 |
| 表单式创建 = 前置仪式 | RuleCreator 三步向导，看结果前先声明一切 | 心智负担 |
| 纯函数运行时读不到其他表 | 向导测试(RuleCreator.tsx:99)在沙箱跑必过，真单元格里字典表读不到/区域被 scalar 截断 | **假反馈** |
| 依赖引擎是重造轮子 | ebDependencies 297 行，`refreshAllEBFunctions` 最终只调 `application.calculate(full)` | 死重 |

**一句话**：80% 的代码在应用层重建 Excel 公式引擎已免费提供的东西（依赖排序、重算、具名函数注册），这层重建正是所有慢反馈的来源。

## 3. 文件级去留清单

### 🟢 保留（真正的泛化资产）

| 文件/部分 | 理由 |
|------|------|
| `ebRules.ts` 的 `cleanInvisibleChars` / `isBlankCell` | 正则清理，原生公式表达不了。EB 唯一硬价值。 |
| `functions.ts` 的 `EB` 分发器 + 预制别名注册 | `=EB.CLEAN` 运行时。**删第 105-116 行自定义规则加载**。 |
| ⚠️ `functions.html` | manifest.xml:96-97 绑定 CF 运行时宿主 + vite.config.ts:46 构建入口。**删了 EB 全废，务必保留。** |
| `functions.json` 的 EB / EBCLEAN / EBEMPTY 三条 | manifest.xml:48 必需元数据。 |
| `ebStorage.ts` 的预制规则 load + 初始化 | 给分发器读预制规则。 |

### 🔴 删除（错误前提的产物 + 慢反馈元凶）

| 文件/部分 | 行数 | 关联清理 |
|------|------|------|
| `RuleCreator.tsx` | 447 | 表单造规则，含第 99 行假反馈测试 |
| `ebDependencies.ts` + `ebDependencies.test.ts` | 297+ | 无自定义规则后依赖图无意义 |
| `rule_generator.py` 的 `regenerate_functions_metadata` + `functions_json_path`(176-285) | ~110 | 连带删 `main.py:64-70` import、`:591` 端点、`api.ts:452` `regenerateFunctionsMetadata`、`contracts.ts:26-40` 三个类型 |
| `ebStorage.ts` 的 saveRule/deleteRule/import/export/backup | ~200 | 公式落单元格自解释，无需持久化自定义规则 |
| `RuleManager.integration.example.tsx` | - | 未提交示例 |
| `request-log.txt` | 718 | 调试日志残留，不该进仓库 |

### 🟡 改造 / 降级

| 文件 | 处理 |
|------|------|
| `rule_generator.py` `generate_rule`(1-172) | 出口 `{logic,compiled}` → `{formula}`；prompt 改产原生公式或判断直接填值；**保持单发不走 planner** |
| `contracts.ts` `GenerateRuleResponse` | 删 `compiled` / `dependencies`，加 `formula` |
| `RuleManager.tsx`(296) | **砍成只读说明书**（决策 A 已定）：只列预制函数 + 用法说明，删创建/导入/导出/删除按钮和 RuleCreator 引用 |
| `App.tsx:3830` "EB规则"按钮 | 保留，改开只读说明窗口（文案可改"函数说明"）|
| `docs/` 6 个 EB 文档 | 标 deprecated 或归档 |

**净账**：删约 1400+ 行；保留真资产 = 两个正则函数 + EB 分发器 + 一条改造后的短链生成能力。

## 4. 分阶段实施

### 阶段 1：建新主路径（`/function` 短链）
1. 改造 `rule_generator.generate_rule`：出口换成 `{formula, explanation}`，prompt 改为产原生公式；保持单发调用，不接 planner。
2. `contracts.ts` / `api.ts` 同步改契约。
3. 前端接 `/function` 斜杠命令：抓选区列头+样本、发现"数据字典"表则读内容、绕开 intent/create_plan → 调改造后的端点。
4. **预览试算（决策 B）**：拿到公式后，先在对话里出预览卡——展示公式 + 用选区真实数据试算一行结果，让用户判断。用户确认后才写入。
5. **写入 + 撤销兜底（决策 B）**：确认后经现有 `writeFormulas` 操作(contracts.ts:127)写进选中单元格、立即重算；写入前记录 `ExecutionUndoSnapshot`，复用现有 `undoExecution` + "↶ 撤销"按钮(App.tsx:57/2760/5879)作兜底。
6. 验证：`/function 根据数据字典算最高错误级` → 出预览卡试算 → 确认写入 → "↶ 撤销"能还原。

**复用红利**：撤销机制(`ExecutionUndoSnapshot`/`undoExecution`)、公式写入(`writeFormulas`)均已存在，阶段 1 主要是接线，非从零造。

### 阶段 2：拆慢反馈元凶
1. 删 `regenerate_functions_metadata` 链（rule_generator + main.py import/端点 + api.ts + contracts.ts 三类型）。
2. `functions.ts` 删自定义规则别名加载（105-116）。
3. 验证：`npm run build` + `tsc --noEmit` + 后端 import 通过；`=EB.CLEAN`/`=EB.EMPTY` 仍可用。

### 阶段 3：删死重 + 简化持久化
1. 删 `ebDependencies.ts` + 测试。
2. `ebStorage.ts` 砍到只留预制规则 load + 初始化（删 save/delete/import/export/backup）。
3. 删 `RuleCreator.tsx`、`RuleManager.integration.example.tsx`、`request-log.txt`。
4. `RuleManager.tsx` 砍成只读说明书：保留搜索 + 函数卡片列表（名称/描述/用法示例），删创建/导入/导出/删除按钮和 RuleCreator 引用、handleDelete/handleExport/handleImport/handleRefresh。App.tsx 按钮保留、文案改"函数说明"。
5. 验证：全量测试通过，`=EB()` 分发器 + 预制函数运行时无回归。

### 阶段 4：文档收尾
1. 6 个旧 EB 文档标 deprecated 或移入 `docs/archive/`。
2. 更新 WORK_LOG:314 待办为已实施。

## 5. 待决策项

- **A（已定 ✅）**：保留一个**只读函数说明窗口**——用户只"触发"和"了解"预制函数（怎么用），**不再创造** EB 函数。后续新需求由维护者写进预制库，不经用户界面。→ RuleManager 砍成只读说明书，RuleCreator 删除。
- **B（已定 ✅）**：先出预览卡（用选区真实数据试算一行）供用户判断 → 确认后写入 → 写入后可撤销兜底。预览降低写错概率，撤销兜住漏网的。底层 `writeFormulas` + `undoExecution` 均可复用。
- **C**：`request-log.txt` 疑似运行时日志误入仓库，确认后应加进 `.gitignore`。

## 6. 风险 / 约束

- `functions.html` + `functions.json` 分发器条目**绝不能删**（manifest 依赖，删则 EB 运行时全废）。
- 阶段间保持每步可 build、可测，`=EB.CLEAN` 全程可用，避免大爆炸式改动。
- 原生公式表达不了的场景（正则清理）仍须能落到 `=EB(...)`，收缩不等于废弃兜底。
