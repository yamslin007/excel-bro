# `/help` 命令实现方案

> **任务目标**：为基础模式添加 `/help` 斜杠命令，展示基础模式的完整功能说明，帮助用户了解可用命令。

---

## 📋 需求概述

在基础模式下，用户输入 `/help` 立即显示基础模式的完整能力说明，让用户知道能做什么。

### 核心特性

- `/help` 是立即执行的命令（不需要用户继续输入参数）
- 显示完整的基础模式功能文档
- 引导用户配置 AI 模型以解锁完整能力
- 在基础模式和 AI 模式下都可用

---

## 🎯 实现方案

### 1. 前端：注册 `/help` 斜杠命令

**文件**：`apps/excel-addin/src/App.tsx`  
**位置**：约第 960 行，`slashCommands` 数组定义处

**改动**：在 `slashCommands` 数组开头添加 `help` 命令

```typescript
const slashCommands: SlashCommand[] = useMemo(
  () => [
    {
      value: "help",
      description: "查看基础模式支持的命令（离线功能说明）"
    },
    {
      value: "function",
      description: "AI 生成原生 Excel 公式（快捷输入，智能补全）"
    },
    {
      value: "model",
      description: "切换模型（使用你已配置的 API 连接）"
    }
  ],
  []
);
```

---

### 2. 前端：处理 `/help` 命令选择

**文件**：`apps/excel-addin/src/App.tsx`  
**位置**：约第 3447 行，`handleSlashCommandSelect` 函数

**改动**：在 `handleSlashCommandSelect` 函数中，添加 `/help` 的特殊处理逻辑

在 `} else {` 分支的开头（处理一级命令的地方），添加以下代码：

```typescript
function handleSlashCommandSelect(value: string) {
  if (slashMode === "model") {
    handleSelectModel(value);
    setPrompt("");
    setShowSlashAutocomplete(false);
  } else {
    // 🆕 特殊命令：/help 立即执行并显示帮助内容
    if (value === "help") {
      appendMessage({
        role: "assistant",
        text: `📘 基础模式功能说明

基础模式是离线、零延迟的确定性操作模式，适合快速的地址操作。

✅ 支持的命令

1️⃣ 清空区域
   · 清空 A1:B10
   · 清空 Sheet1!C5:D20
   · 清除 A1

2️⃣ 写入值
   · 把 A1 填入 100
   · 在 B2 写入 "标题"
   · 把 Sheet1!C5 设为 3.14

3️⃣ 写入公式
   · 把 D1 设为 =SUM(A1:A10)
   · 在 E5 填入 =AVERAGE(B2:B100)

📝 语法说明
• 支持带工作表前缀（Sheet1!A1）或省略（默认当前表）
• 支持单元格（A1）或区域（A1:B10）
• 写入值支持：数字、文本（加引号）、布尔值、空值
• 写入公式：以 = 开头

❌ 不支持的功能
基础模式不支持需要理解字段含义的操作，例如：
• 查找、定位、统计（"销售额最高的是谁"）
• 去重、排序、筛选（"去掉重复行"）
• 汇总、分组（"按部门统计平均分"）

💡 解锁完整能力
输入 /model 即可配置 AI 模型，使用自然语言进行复杂分析
（支持免费本地模型：Ollama、LM Studio 等）`
      });
      setPrompt("");
      setShowSlashAutocomplete(false);
      // 光标聚焦回输入框
      const input = composerInputRef.current;
      if (input) {
        input.focus();
      }
      return;
    }
    
    // 一级命令：/function 填入等继续输入描述；/model 进二级菜单。
    if (value === "model") {
      setPrompt("/model ");
      setSlashMode("model");
      setSlashFilter("");
      // 保持补全开启，接下来会切到模型列表。
    } else {
      // ... 原有代码保持不变
    }
  }
}
```

**关键实现点**：
- `/help` 命令立即执行，不等待用户继续输入
- 使用 `appendMessage` 追加 `assistant` 角色消息
- 清空输入框 `setPrompt("")`
- 关闭自动补全 `setShowSlashAutocomplete(false)`
- 光标聚焦回输入框，方便用户继续输入

---

### 3. 后端：优化拒绝消息

**文件**：`server/app/planner.py`  
**位置**：约第 238-252 行，`_local_analysis` 函数

**改动**：将详细的拒绝消息改为简短引导

**当前代码**：
```python
def _local_analysis(request: PlanRequest) -> AssistantResponse:
    address_plan = _local_address_plan(request)
    if address_plan is not None:
        return address_plan
    return AnswerResponse(
        provider="local",
        message=(
            "基础模式只支持明确的地址/范围操作，例如：\n"
            "· 清空 A1:B10\n"
            "· 把 A1 填入 100\n"
            "· 把 A1 设为 =SUM(B2:B10)\n\n"
            "其他需要理解表格字段或数据含义的需求（查找、汇总、去重、比较等），"
            "请配置 AI 模型后使用。"
        ),
    )
```

**修改为**：
```python
def _local_analysis(request: PlanRequest) -> AssistantResponse:
    address_plan = _local_address_plan(request)
    if address_plan is not None:
        return address_plan
    return AnswerResponse(
        provider="local",
        message=(
            "基础模式不支持此操作。\n\n"
            "💡 输入 /help 查看支持的命令\n"
            "或配置 AI 模型解锁完整能力。"
        ),
    )
```

**改动原因**：
- 原有消息过长，与 `/help` 内容重复
- 改为简短引导，指向 `/help` 获取完整文档
- 形成分工：拒绝消息负责快速引导，`/help` 负责完整文档

---

## 📂 涉及文件总结

| 文件 | 改动内容 | 行数位置 |
|------|---------|---------|
| `apps/excel-addin/src/App.tsx` | 1. 在 `slashCommands` 添加 help 命令<br>2. 在 `handleSlashCommandSelect` 添加 `/help` 处理逻辑（约 30 行代码） | 约 960 行<br>约 3447 行 |
| `server/app/planner.py` | 修改 `_local_analysis` 的拒绝消息（约 5 行） | 约 238-252 行 |

---

## ✅ 验收标准

### 功能验收

1. **命令注册**
   - ✅ 输入 `/` 时，自动补全列表中出现 `/help`
   - ✅ 描述为"查看基础模式支持的命令（离线功能说明）"
   - ✅ `/help` 在列表顶部（第一个）

2. **命令执行**
   - ✅ 选择 `/help` 后立即显示帮助内容（不等待继续输入）
   - ✅ 输入框自动清空
   - ✅ 光标自动聚焦回输入框

3. **帮助内容**
   - ✅ 包含：支持的命令（清空、写入值、写入公式）
   - ✅ 包含：语法说明（工作表前缀、单元格/区域、值类型）
   - ✅ 包含：不支持的功能（查找、统计、去重等）
   - ✅ 包含：解锁完整能力的引导（配置 AI 模型）
   - ✅ 格式清晰：emoji + 编号列表 + 示例代码
   - ✅ 消息角色为 `assistant`

4. **拒绝消息优化**
   - ✅ 基础模式下输入不支持的操作（如"计算平均值"）
   - ✅ 显示简短拒绝消息
   - ✅ 消息中引导用户使用 `/help` 或配置模型

5. **跨模式可用**
   - ✅ 基础模式下可以使用 `/help`
   - ✅ AI 模式下也可以使用 `/help`（查看基础模式能力）

### 交互验收

1. **自动补全**
   - ✅ 输入 `/h` → 过滤出 `/help`
   - ✅ 输入 `/he` → 只显示 `/help`
   - ✅ 按 Tab 或 Enter → 执行 `/help`

2. **键盘操作**
   - ✅ 方向键可以在补全列表中导航
   - ✅ 选中 `/help` 后按 Enter 立即执行
   - ✅ 按 Esc 关闭补全列表

3. **鼠标操作**
   - ✅ 鼠标悬停 `/help` 时高亮
   - ✅ 点击 `/help` 立即执行

---

## 🧪 测试场景

### 场景 1：基础模式下主动查看帮助

**操作步骤**：
1. 切换到基础模式（选择 "基础模式（离线）"）
2. 输入框输入 `/`
3. 从补全列表选择 `/help`

**预期结果**：
- 立即显示完整的基础模式功能说明
- 输入框清空
- 可以继续输入下一条命令

---

### 场景 2：基础模式下触发拒绝消息

**操作步骤**：
1. 切换到基础模式
2. 输入不支持的命令，如"计算平均值"
3. 发送

**预期结果**：
- 显示拒绝消息："基础模式不支持此操作。💡 输入 /help 查看支持的命令或配置 AI 模型解锁完整能力。"
- 用户看到提示后可以输入 `/help` 查看详细文档

---

### 场景 3：基础模式下执行支持的命令

**操作步骤**：
1. 先输入 `/help` 查看帮助
2. 根据帮助文档，输入支持的命令，如"清空 A1:B10"
3. 发送

**预期结果**：
- 成功生成清空操作计划
- 显示预览和确认按钮

---

### 场景 4：AI 模式下查看帮助

**操作步骤**：
1. 配置并切换到 AI 模型
2. 输入 `/help`

**预期结果**：
- 也能显示基础模式的功能说明
- 用户可以了解"如果切换到基础模式，能做什么"

---

### 场景 5：Tab 自动补全

**操作步骤**：
1. 输入框输入 `/he`
2. 按 Tab 键

**预期结果**：
- 自动补全为 `/help` 并立即执行
- 显示帮助内容

---

## 📝 帮助内容（最终版本）

以下是 `/help` 命令显示的完整内容，确保实现时使用此版本：

```
📘 基础模式功能说明

基础模式是离线、零延迟的确定性操作模式，适合快速的地址操作。

✅ 支持的命令

1️⃣ 清空区域
   · 清空 A1:B10
   · 清空 Sheet1!C5:D20
   · 清除 A1

2️⃣ 写入值
   · 把 A1 填入 100
   · 在 B2 写入 "标题"
   · 把 Sheet1!C5 设为 3.14

3️⃣ 写入公式
   · 把 D1 设为 =SUM(A1:A10)
   · 在 E5 填入 =AVERAGE(B2:B100)

📝 语法说明
• 支持带工作表前缀（Sheet1!A1）或省略（默认当前表）
• 支持单元格（A1）或区域（A1:B10）
• 写入值支持：数字、文本（加引号）、布尔值、空值
• 写入公式：以 = 开头

❌ 不支持的功能
基础模式不支持需要理解字段含义的操作，例如：
• 查找、定位、统计（"哪一项的总和最高"）
• 去重、排序、筛选（"去掉重复行"）
• 汇总、分组（"按部门统计平均分"）

💡 解锁完整能力
输入 /model 即可配置 AI 模型，使用自然语言进行复杂分析
（支持免费本地模型：Ollama、LM Studio 等）
```

**注意**：
- 使用 emoji 增强可读性
- 使用编号列表清晰分类
- 示例命令以 `·` 开头
- 最后引导用户配置模型以解锁完整能力

---

## 🎨 可选增强（后续迭代）

以下功能不在本次实现范围内，但可作为未来优化方向：

1. **首次引导**
   - 用户首次选择"基础模式"时，自动显示一次 `/help`
   - 需要在本地存储记录"是否已显示过首次引导"

2. **快捷图标**
   - 基础模式顶部状态栏添加 `[?]` 图标
   - 点击图标触发 `/help`

3. **示例可点击**
   - 帮助内容中的命令示例可点击直接填入输入框
   - 需要扩展消息渲染逻辑，支持交互式内容

4. **上下文敏感帮助**
   - AI 模式下输入 `/help` 显示不同的内容
   - 说明"当前已配置 AI 模型，可使用自然语言操作"

---

## 🚀 实现检查清单

实现前请确认：

- [ ] 已阅读完整文档
- [ ] 理解 `/help` 是立即执行的命令（不需要参数）
- [ ] 理解前后端分工：前端负责显示，后端拒绝消息引导到前端

实现后请检查：

- [ ] 前端：`slashCommands` 数组已添加 help 命令
- [ ] 前端：`handleSlashCommandSelect` 已添加 `/help` 处理逻辑
- [ ] 后端：`_local_analysis` 的拒绝消息已优化
- [ ] 帮助内容与文档中的最终版本完全一致
- [ ] 所有验收标准已通过测试
- [ ] 至少完成 5 个测试场景

---

## 📞 实现疑问

如有实现疑问，请参考：

1. **现有斜杠命令实现**
   - 参考 `/function` 和 `/model` 的实现方式
   - 位置：`App.tsx` 约 3447-3473 行

2. **消息追加方式**
   - 使用 `appendMessage` 函数
   - 示例：`appendMessage({ role: "assistant", text: "..." })`

3. **输入框引用**
   - 使用 `composerInputRef.current` 访问输入框
   - 聚焦：`composerInputRef.current?.focus()`

4. **斜杠命令自动补全**
   - 组件：`SlashCommandAutocomplete`
   - 过滤逻辑在 `useSlashCommands` Hook 中

---

**文档版本**：v1.0  
**创建日期**：2026-08-18  
**最后更新**：2026-08-18  
**状态**：已实现 ✅
