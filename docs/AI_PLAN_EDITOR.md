# AI 辅助计划编辑功能 - 2026-08-23

## 功能概述

为计划编辑器添加 AI 辅助功能，用户可以用自然语言描述修改需求，AI 返回修改后的完整计划 JSON。

## 用户界面

### 两栏布局

**左栏（默认 400px，可拖动宽度、可折叠）：** AI 对话界面
- 显示用户与 AI 的对话历史（带头像和气泡样式）
- 底部输入框，支持多行输入，Enter 发送，Shift+Enter 换行
- AI 思考时显示三点动画，回复下方有「应用到编辑器 →」按钮
- 可折叠成左侧窄条，为 JSON 腾出空间

**右栏（自适应）：** JSON 编辑器（`JsonCodeEditor`，与任务窗格抽屉共用）
- 语法高亮（key/字符串/数字/布尔/null 分色），浅色跟随主题、可切深色
- 固定显示行号（错误定位直接指向行号），提供复制、格式化按钮
- 编辑时本地实时校验：格式错误立即定位到第几行第几列
- 底部状态栏显示光标行列和是否已修改
- 保存流程：本地 JSON 语法 + 白名单结构校验，快速保存（不请求服务端）

### 交互流程

1. **用户在左侧输入**：例如 "添加一个筛选步骤，只保留销售额大于 1000 的记录"
2. **AI 思考并回复**：
   - 文本说明：简洁描述做了什么修改（2-3 句话）
   - 修改后的完整计划 JSON（后台返回）
3. **用户点击「应用到编辑器」**：右侧 JSON 编辑器更新为 AI 生成的内容
4. **用户可继续修改**：
   - 在右侧手动编辑 JSON
   - 或在左侧继续与 AI 对话
5. **保存**：点击底部「保存」按钮，本地检查格式与结构后保存（Ctrl+S 保存并留在窗口）

## 技术实现

### 服务端（Python）

**新增文件：** `server/app/plan_chat.py`

```python
async def chat_modify_plan(
    current_plan: AnalysisPlan,
    user_message: str,
    connection: ModelConnection,
    timeout: float,
) -> tuple[str, AnalysisPlan]:
    """AI 辅助修改计划。
    
    返回：(AI 回复文本, 修改后的完整计划)
    """
```

**新增 API 端点：** `POST /api/plans/chat-modify`

请求体：
```json
{
  "plan": { /* 当前的 AnalysisPlan */ },
  "message": "用户的修改需求",
  "modelId": "optional-model-id"
}
```

响应：
```json
{
  "explanation": "AI 对修改的简短说明",
  "modifiedPlan": { /* 修改后的完整 AnalysisPlan */ }
}
```

**工作流程：**
1. 使用 OpenAI Function Calling 格式调用模型
2. 系统提示词包含计划结构说明和常见操作类型
3. AI 返回 explanation + modifiedPlan
4. 后端自动删除 AI 生成的 acceptanceCriteria（由服务端重算）
5. 保留原有的 sourceFingerprint
6. 通过 Pydantic 校验确保返回的计划合法

### 前端（TypeScript + React）

**修改文件：** `apps/excel-addin/src/components/PlanEditorDialog.tsx`

**新增状态：**
```typescript
type DialogPhase = "loading" | "editing" | "validating" | "error" | "chatting";

interface ChatMessage {
  role: "user" | "assistant";
  content: string;
  modifiedPlan?: AnalysisPlan;
}

const [chatMessages, setChatMessages] = useState<ChatMessage[]>([]);
const [chatInput, setChatInput] = useState("");
```

**新增函数：**
- `handleChatSubmit()`: 发送用户消息给 AI
- `handleApplyPlan(plan)`: 应用 AI 生成的计划到编辑器

**新增 API 客户端：** `apps/excel-addin/src/api.ts`
```typescript
export async function chatModifyPlan(
  request: PlanChatRequest,
  signal?: AbortSignal
): Promise<PlanChatResponse>
```

### 样式（CSS）

**新增样式：** `apps/excel-addin/src/styles.css`

主要样式类：
- `.plan-editor-dialog-split`: 两栏布局容器
- `.plan-editor-chat-panel`: 左侧 AI 对话面板
- `.plan-editor-chat-messages`: 对话消息列表
- `.plan-editor-chat-message`: 单条消息样式
  - `.user`: 用户消息（绿色背景）
  - `.assistant`: AI 消息（白色背景）
- `.plan-editor-chat-input`: 底部输入框
- `.apply-plan-button`: 应用按钮
- `.plan-editor-json-panel`: 右侧 JSON 编辑器面板

## 设计亮点

1. **保持手动编辑能力**：用户可以在 AI 辅助和手动编辑之间自由切换
2. **完整计划输出**：AI 返回完整的计划 JSON，而非增量 diff，避免合并冲突
3. **应用前可预览**：用户可以看到 AI 的修改说明，决定是否应用
4. **多轮对话**：支持连续与 AI 对话，每次基于当前编辑器的内容
5. **快速保存**：保存只做本地 JSON 语法 + 白名单结构检查，不等待服务端往返

## 错误处理

- **当前 JSON 格式错误**：无法使用 AI 修改，提示用户先修复 JSON
- **AI 调用失败**：显示错误消息，可重试
- **结构校验失败**：保存时触发，显示具体错误信息

## 构建结果

```
✓ built in 994ms
styles.css     +2.40 kB (103.12 kB)
planEditor.js  +1.92 kB (5.25 kB)
```

## 待测试功能

1. **基本 AI 对话**
   - [ ] 发送修改需求，收到 AI 回复
   - [ ] 点击「应用到编辑器」，JSON 更新
   - [ ] 多轮对话，每次基于当前编辑器内容

2. **手动编辑兼容**
   - [ ] 应用 AI 计划后，手动修改 JSON
   - [ ] 手动修改后，继续与 AI 对话

3. **错误场景**
   - [ ] JSON 格式错误时使用 AI
   - [ ] AI 调用超时/失败
   - [ ] 应用计划后保存失败（校验不通过）

4. **UI 响应**
   - [ ] 发送消息时显示「AI 思考中...」
   - [ ] 对话历史滚动到底部
   - [ ] Enter 发送，Shift+Enter 换行

## 后续优化

1. **对话历史持久化**：保存到 localStorage，下次打开恢复
2. **撤销/重做**：支持撤销应用 AI 计划的操作
3. **消息内代码片段高亮**：左侧消息中显示 JSON diff 时高亮
4. **快捷命令**：例如 "/explain" 解释当前计划，"/simplify" 简化步骤

已实现（2026-08-25 UI 重构）：JSON 编辑器语法高亮与行号、编辑时本地实时校验、
复制/格式化、Ctrl+S 快捷保存（提示「已保存」，窗口保持打开可继续编辑）、
保存按钮保存后关闭窗口、未保存修改离开确认、AI 思考中动画、聊天栏折叠与
宽度拖动、深/浅代码皮肤。

---

**状态：** 代码已完成，构建通过  
**完成时间：** 2026-08-23
