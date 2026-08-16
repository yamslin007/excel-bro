# P1 问题分析：App.tsx 超大组件拆分方案

> 当前状态：4050 行，84 个 hooks，职责过多
> 目标：拆分为多个专注的组件和 hooks，提升可维护性

## 问题分析

### 当前状态统计

```
文件：apps/excel-addin/src/App.tsx
行数：4050 行
useState 数量：50+ 个
useEffect 数量：84 个 hooks 总数
导入依赖：150+ 行
```

### 职责识别

通过代码分析，App.tsx 承担了以下职责：

#### 1. 工作簿上下文管理
- `workbook`: WorkbookSnapshot
- `folderCatalog`: FolderCatalog
- `folderSelection`: FolderSelection
- `contextOpen`, `sheetSearch`
- 工作簿结构监听和刷新

#### 2. 对话状态管理
- `chatHistory`: ChatHistoryState
- `prompt`, `replyTarget`
- `clarificationDrafts`
- `status`, `activity`, `activitySeconds`
- 对话历史加载/保存/删除

#### 3. 消息处理流程
- Intent 判断
- 工具调用（query_table）
- 计划生成和执行
- 验证报告处理

#### 4. 模型管理
- `modelOptions`, `selectedModelId`
- `modelSettings`, `modelMenuOpen`
- `connectionDraft`
- 模型连接的 CRUD 操作

#### 5. 工具管理（"我的工具"）
- `tools`, `queryTools`
- `toolsOpen`, `selectedToolId`
- `toolDrawerView`, `toolDetailMode`
- `saveCandidate`
- 工具的创建、编辑、删除、执行

#### 6. 图片附件
- `pendingImages`, `imageError`
- `draggingImage`
- 图片选择、拖拽、验证

#### 7. UI 状态
- `historyOpen`, `settingsOpen`
- `petVisible`, `focusOpening`
- `moreMenuOpen`
- 各种抽屉和弹窗的开关

#### 8. 斜杠命令
- `showSlashAutocomplete`
- `slashFilter`, `slashMode`
- 命令补全和执行

#### 9. 首次引导
- `isNarrowPane`, `widenStepDone`
- `modelCatalogLoaded`, `modelGuideDismissed`

#### 10. 撤销/恢复
- `lastUndoSnapshot`
- 撤销快照管理

#### 11. 服务健康检查
- `serverOnline`, `serviceHealth`
- 健康检查轮询

#### 12. 公式生成（/function 命令）
- 公式生成、预览、写入
- 目标单元格选择

## 拆分方案

### 阶段 1：提取自定义 Hooks（优先级：高）

#### 1.1 useWorkbookContext
```typescript
// hooks/useWorkbookContext.ts
export function useWorkbookContext() {
  const [workbook, setWorkbook] = useState<WorkbookSnapshot | null>(null);
  const [folderCatalog, setFolderCatalog] = useState<FolderCatalog | null>(null);
  const [folderSelection, setFolderSelection] = useState<FolderSelection | null>(null);
  const [contextOpen, setContextOpen] = useState(false);
  const [sheetSearch, setSheetSearch] = useState("");
  
  // 工作簿刷新逻辑
  const refreshWorkbook = useCallback(async () => {
    // ...
  }, []);
  
  // 文件夹选择逻辑
  const handleFolderSelect = useCallback(async () => {
    // ...
  }, []);
  
  return {
    workbook,
    folderCatalog,
    folderSelection,
    contextOpen,
    setContextOpen,
    sheetSearch,
    setSheetSearch,
    refreshWorkbook,
    handleFolderSelect,
  };
}
```

**影响范围：** ~500 行 → 单独文件  
**收益：** 工作簿相关逻辑独立，可单独测试

#### 1.2 useConversation
```typescript
// hooks/useConversation.ts
export function useConversation() {
  const [chatHistory, setChatHistory] = useState<ChatHistoryState>(loadChatHistory);
  const [status, setStatus] = useState<Status>("idle");
  const [activity, setActivity] = useState<ActivityProgress | null>(null);
  
  // 创建新对话
  const createNewConversation = useCallback(() => {
    // ...
  }, []);
  
  // 切换对话
  const switchConversation = useCallback((conversationId: string) => {
    // ...
  }, []);
  
  // 删除对话
  const deleteConversation = useCallback((conversationId: string) => {
    // ...
  }, []);
  
  return {
    chatHistory,
    status,
    activity,
    currentConversation: chatHistory.conversations[chatHistory.activeConversationId],
    createNewConversation,
    switchConversation,
    deleteConversation,
  };
}
```

**影响范围：** ~600 行 → 单独文件  
**收益：** 对话状态管理独立

#### 1.3 useModelManagement
```typescript
// hooks/useModelManagement.ts
export function useModelManagement() {
  const [modelOptions, setModelOptions] = useState<ModelOption[]>([...]);
  const [selectedModelId, setSelectedModelId] = useState(...);
  const [modelSettings, setModelSettings] = useState<ModelSettings | null>(null);
  const [connectionDraft, setConnectionDraft] = useState<ModelConnectionDraft | null>(null);
  
  // 加载模型目录
  const loadModelCatalog = useCallback(async () => {
    // ...
  }, []);
  
  // 保存模型连接
  const saveConnection = useCallback(async (data: UpsertModelConnectionRequest) => {
    // ...
  }, []);
  
  // 删除模型连接
  const deleteConnection = useCallback(async (id: string) => {
    // ...
  }, []);
  
  return {
    modelOptions,
    selectedModelId,
    setSelectedModelId,
    modelSettings,
    connectionDraft,
    setConnectionDraft,
    loadModelCatalog,
    saveConnection,
    deleteConnection,
  };
}
```

**影响范围：** ~400 行 → 单独文件  
**收益：** 模型管理逻辑隔离

#### 1.4 useToolManagement
```typescript
// hooks/useToolManagement.ts
export function useToolManagement() {
  const [tools, setTools] = useState<SavedTool[]>(loadTools);
  const [queryTools, setQueryTools] = useState<SavedQueryTool[]>(loadQueryTools);
  const [selectedToolId, setSelectedToolId] = useState<string | null>(null);
  const [toolDrawerView, setToolDrawerView] = useState<ToolDrawerView>("library");
  
  // 保存工具
  const saveTool = useCallback((plan: AnalysisPlan, name: string) => {
    // ...
  }, []);
  
  // 删除工具
  const deleteTool = useCallback((id: string) => {
    // ...
  }, []);
  
  // 执行工具
  const executeTool = useCallback(async (id: string, params: Record<string, string>) => {
    // ...
  }, []);
  
  return {
    tools,
    queryTools,
    selectedToolId,
    setSelectedToolId,
    toolDrawerView,
    setToolDrawerView,
    saveTool,
    deleteTool,
    executeTool,
  };
}
```

**影响范围：** ~500 行 → 单独文件  
**收益：** 工具管理逻辑隔离

#### 1.5 useImageAttachments
```typescript
// hooks/useImageAttachments.ts
export function useImageAttachments() {
  const [pendingImages, setPendingImages] = useState<PendingImage[]>([]);
  const [imageError, setImageError] = useState("");
  const [draggingImage, setDraggingImage] = useState(false);
  
  // 添加图片
  const addImage = useCallback(async (file: File) => {
    // 验证、压缩、转 base64
    // ...
  }, []);
  
  // 移除图片
  const removeImage = useCallback((index: number) => {
    // ...
  }, []);
  
  // 处理拖拽
  const handleDrop = useCallback((e: DragEvent) => {
    // ...
  }, []);
  
  return {
    pendingImages,
    imageError,
    draggingImage,
    addImage,
    removeImage,
    handleDrop,
    setDraggingImage,
  };
}
```

**影响范围：** ~200 行 → 单独文件  
**收益：** 图片处理逻辑隔离

#### 1.6 useSlashCommands
```typescript
// hooks/useSlashCommands.ts
export function useSlashCommands(options: {
  onModelSelect: (modelId: string) => void;
  onFunctionGenerate: (description: string) => void;
}) {
  const [showAutocomplete, setShowAutocomplete] = useState(false);
  const [filter, setFilter] = useState("");
  const [mode, setMode] = useState<"command" | "model">("command");
  
  // 检测斜杠命令
  const detectSlashCommand = useCallback((text: string) => {
    // ...
  }, []);
  
  // 执行命令
  const executeCommand = useCallback((command: SlashCommand) => {
    // ...
  }, []);
  
  return {
    showAutocomplete,
    filter,
    mode,
    detectSlashCommand,
    executeCommand,
    setShowAutocomplete,
    setFilter,
    setMode,
  };
}
```

**影响范围：** ~150 行 → 单独文件  
**收益：** 斜杠命令逻辑隔离

### 阶段 2：提取业务逻辑模块（优先级：中）

#### 2.1 messageProcessor.ts
```typescript
// services/messageProcessor.ts
export class MessageProcessor {
  async processUserMessage(params: {
    prompt: string;
    workbook: WorkbookSnapshot;
    images: PendingImage[];
    modelId: string;
    lastResult?: ResultContext;
  }): Promise<{
    type: "clarification" | "tool_request" | "plan" | "answer";
    data: IntentCheckResponse;
  }> {
    // Intent 判断逻辑
    // ...
  }
  
  async handleToolRequest(params: {
    request: DataToolRequest;
    workbook: WorkbookSnapshot;
  }): Promise<DataToolResult> {
    // query_table 执行逻辑
    // ...
  }
  
  async generatePlan(params: {
    prompt: string;
    workbook: WorkbookSnapshot;
    toolResults?: DataToolResult[];
  }): Promise<AnalysisPlan> {
    // 计划生成逻辑
    // ...
  }
}
```

**影响范围：** ~800 行 → 单独文件  
**收益：** 核心业务逻辑可独立测试

#### 2.2 planExecutor.ts
```typescript
// services/planExecutor.ts
export class PlanExecutor {
  async execute(plan: AnalysisPlan, workbook: WorkbookSnapshot): Promise<{
    report: VerificationReport;
    undoSnapshot?: ExecutionUndoSnapshot;
  }> {
    // 计划执行逻辑
    // ...
  }
  
  async undo(snapshot: ExecutionUndoSnapshot): Promise<void> {
    // 撤销逻辑
    // ...
  }
}
```

**影响范围：** ~300 行 → 单独文件  
**收益：** 执行逻辑隔离，可测试

### 阶段 3：优化 JSX 结构（优先级：低）

App.tsx 最终只保留：
- 核心状态编排
- Hooks 调用
- 高层事件处理
- JSX 布局

```typescript
export default function App() {
  // 使用自定义 hooks
  const workbookContext = useWorkbookContext();
  const conversation = useConversation();
  const modelManagement = useModelManagement();
  const toolManagement = useToolManagement();
  const imageAttachments = useImageAttachments();
  const slashCommands = useSlashCommands({...});
  
  // 核心业务处理
  const messageProcessor = useMemo(() => new MessageProcessor(), []);
  const planExecutor = useMemo(() => new PlanExecutor(), []);
  
  // 高层事件处理
  const handleSendMessage = async () => {
    const result = await messageProcessor.processUserMessage({
      prompt,
      workbook: workbookContext.workbook,
      images: imageAttachments.pendingImages,
      modelId: modelManagement.selectedModelId,
    });
    
    // 根据结果更新状态
    // ...
  };
  
  return (
    <div className="app-container">
      <ChatHeader {...} />
      <MessageList messages={conversation.currentConversation.messages} />
      <Composer onSend={handleSendMessage} {...} />
      
      {/* 各种抽屉和弹窗 */}
      <SettingsDrawer {...} />
      <ToolDrawer {...} />
      <HistoryDrawer {...} />
    </div>
  );
}
```

**目标行数：** 4050 行 → **~800 行**

## 实施计划

### 第一周：提取核心 hooks
- [x] 创建 `hooks/` 目录结构
- [ ] 实现 `useWorkbookContext`（影响 ~500 行）
- [ ] 实现 `useConversation`（影响 ~600 行）
- [ ] 在 App.tsx 中使用新 hooks
- [ ] 运行测试确保无破坏

### 第二周：提取业务逻辑
- [ ] 实现 `useModelManagement`（影响 ~400 行）
- [ ] 实现 `useToolManagement`（影响 ~500 行）
- [ ] 实现 `useImageAttachments`（影响 ~200 行）
- [ ] 实现 `useSlashCommands`（影响 ~150 行）

### 第三周：提取业务模块
- [ ] 创建 `services/` 目录
- [ ] 实现 `MessageProcessor`（影响 ~800 行）
- [ ] 实现 `PlanExecutor`（影响 ~300 行）
- [ ] 为业务模块编写单元测试

### 第四周：优化和测试
- [ ] 优化 App.tsx JSX 结构
- [ ] 补充集成测试
- [ ] 性能优化（useCallback, useMemo）
- [ ] 文档更新

## 预期收益

### 代码质量
- **行数：** 4050 → ~800（App.tsx）+ 6 个 hooks + 2 个服务类
- **可测试性：** 从"难以测试"到"每个模块可独立测试"
- **可维护性：** 职责清晰，修改影响范围明确

### 开发效率
- **新增功能：** 不再需要在 4050 行中找位置
- **Bug 修复：** 快速定位到具体 hook 或服务
- **代码审查：** 每个 PR 只涉及特定模块

### 团队协作
- **并行开发：** 不同开发者可同时修改不同 hooks
- **知识分享：** 新成员可从单个 hook 开始理解
- **减少冲突：** Git 合并冲突大幅减少

## 风险与缓解

### 风险 1：重构破坏现有功能
**缓解：**
- 每个阶段独立测试
- 保持现有测试覆盖
- 增量式重构，随时可回滚

### 风险 2：性能回归
**缓解：**
- 使用 React DevTools Profiler 监控
- 合理使用 useMemo/useCallback
- 必要时使用 React.memo

### 风险 3：状态同步问题
**缓解：**
- 明确状态所有权
- 使用 Context 共享跨模块状态（如需要）
- 文档化状态流转

## 成功标准

- [ ] App.tsx 行数降至 1000 行以下
- [ ] 提取至少 6 个自定义 hooks
- [ ] 提取至少 2 个业务服务类
- [ ] 所有现有测试通过
- [ ] 新增至少 10 个单元测试
- [ ] 无性能回归（Lighthouse 评分持平或提升）

---

**文档作者：** Claude Opus 4  
**创建日期：** 2026-08-16  
**相关问题：** P1 - App.tsx 超大组件（4050 行）  
**预计工时：** 4 周（每周 8-10 小时）
