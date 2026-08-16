# Hooks 使用指南

本文档覆盖 `apps/excel-addin/src/hooks/` 下 13 个自定义 Hook，说明每个 Hook 的职责、导入方式、返回接口、典型用法和注意事项。

## 概述

P1 重构将原先集中在 `App.tsx` 中的状态、副作用和业务逻辑拆分为 13 个单一职责 Hook。Hook 之间通过 props、回调或共享服务模块协作，避免跨 Hook 直接读写彼此的内部状态。

整体分类：

- 状态管理类：负责领域状态和派生状态
- 资源管理类：负责工具、模型、工作簿、图片等资源生命周期
- 交互辅助类：负责命令输入、复制反馈、服务健康等 UI 辅助能力

## Hook 分类

### 状态管理类

- `useConversation`
- `useUIState`
- `useExecutionApproval`
- `useScopeSelection`
- `useActivityProgress`
- `useUndoSnapshot`

### 资源管理类

- `useImageAttachments`
- `useModelManagement`
- `useToolManagement`
- `useWorkbookContext`

### 交互辅助类

- `useSlashCommands`
- `useCopyFeedback`
- `useServiceHealth`

---

## 详细 API 文档

### useConversation

**职责：** 对话历史和消息管理。

**导入：**

```typescript
import { useConversation } from "./hooks/useConversation";
```

**接口：**

```typescript
const {
  chatHistory,
  activeConversation,
  messages,
  pendingDeleteConversationId,
  setMessages,
  newChat,
  openConversation,
  deleteConversation,
  confirmDeleteConversation,
  cancelDeleteConversation,
  setChatHistory,
  setPendingDeleteConversationId
} = useConversation();
```

**使用示例：**

```typescript
const handleNewChat = () => {
  newChat();
};

const handleSendMessage = (text: string) => {
  setMessages((current) => [
    ...current,
    {
      id: crypto.randomUUID(),
      role: "user",
      text,
      createdAt: new Date().toISOString()
    }
  ]);
};
```

**注意事项：**

- `messages` 会自动持久化到 `localStorage`
- 最多保存 `MAX_STORED_CONVERSATIONS` 个对话
- 每个对话最多 `MAX_MESSAGES_PER_CONVERSATION` 条消息
- `setMessages` 会同步更新当前对话标题和更新时间

---

### useUIState

**职责：** UI 抽屉、菜单、宠物可见性和窗格宽度管理。

**导入：**

```typescript
import { useUIState } from "./hooks/useUIState";
```

**接口：**

```typescript
const {
  toolsOpen,
  historyOpen,
  settingsOpen,
  modelMenuOpen,
  moreMenuOpen,
  isRuleManagerOpen,
  petVisible,
  focusOpening,
  isNarrowPane,
  widenStepDone,
  composerHeight,
  setToolsOpen,
  setHistoryOpen,
  setSettingsOpen,
  setModelMenuOpen,
  setMoreMenuOpen,
  setIsRuleManagerOpen,
  setPetVisible,
  setFocusOpening,
  setIsNarrowPane,
  setWidenStepDone,
  setComposerHeight,
  closeAllDrawers,
  closeAllMenus,
  togglePetVisibility
} = useUIState();
```

**使用示例：**

```typescript
const handleOpenSettings = () => {
  closeAllDrawers();
  setSettingsOpen(true);
};
```

**注意事项：**

- `petVisible` 会持久化到 `localStorage`
- `window.resize` 事件会在 Hook 内自动监听
- 抽屉和菜单关闭操作应优先使用 `closeAllDrawers` / `closeAllMenus`

---

### useExecutionApproval

**职责：** 已验证计划跟踪和工具固化风险批准。

**导入：**

```typescript
import { useExecutionApproval } from "./hooks/useExecutionApproval";
```

**接口：**

```typescript
const {
  saveCandidate,
  setSaveCandidate,
  approveFixedContent,
  setApproveFixedContent,
  approveDestructive,
  setApproveDestructive,
  verifiedPlanIds,
  markPlanVerified,
  saveEligibility,
  beginSaveCandidate,
  closeSaveCandidate
} = useExecutionApproval();
```

**使用示例：**

```typescript
const handleSaveAsTool = (plan: AnalysisPlan) => {
  beginSaveCandidate(plan);
};

const handlePlanVerified = (planId: string) => {
  markPlanVerified(planId);
};
```

**注意事项：**

- `saveEligibility` 根据 `saveCandidate` 自动派生
- 只有经过 `markPlanVerified` 的计划才允许进入工具保存流程
- `beginSaveCandidate` 会重置两类风险批准状态

---

### useScopeSelection

**职责：** 数据源模式、工作表选择和工具保存字段管理。

**导入：**

```typescript
import { useScopeSelection } from "./hooks/useScopeSelection";
```

**接口：**

```typescript
const {
  toolName,
  setToolName,
  toolDescription,
  setToolDescription,
  selectedSheetNames,
  setSelectedSheetNames,
  selectionConfirmed,
  setSelectionConfirmed,
  sourceMode,
  setSourceMode,
  workbookScopeMode,
  setWorkbookScopeMode,
  folderCatalog,
  setFolderCatalog,
  folderSheetKeys,
  setFolderSheetKeys,
  applyWorkbookSnapshotSelection,
  toggleSheet,
  toggleFolderSheet,
  applyFolderCatalog,
  chooseAutomaticScope,
  chooseManualScope,
  chooseFolderScope,
  folderSelections,
  selectedNamesFor,
  selectAllSheets,
  clearSelectedSheets
} = useScopeSelection();
```

**使用示例：**

```typescript
const handleAutoScope = () => {
  chooseAutomaticScope(workbook);
};

const handleToggleSheet = (sheetName: string) => {
  toggleSheet(sheetName);
};
```

**注意事项：**

- `sourceMode` 与 `workbookScopeMode` 的切换会影响 `selectionConfirmed`
- 文件夹模式下应使用 `folderSheetKeys` 和 `folderSelections`
- 手动模式无选择时，确认按钮应保持禁用

---

### useActivityProgress

**职责：** 活动进度、步骤记录和计时器管理。

**导入：**

```typescript
import { useActivityProgress } from "./hooks/useActivityProgress";
```

**接口：**

```typescript
const {
  activity,
  activitySeconds,
  startActivity,
  advanceActivity,
  updateActivityDetail,
  completeActivity
} = useActivityProgress({
  onPersistLog: (log) => {
    // 将 activityLog 写入最近一条 assistant 消息
  }
});
```

**使用示例：**

```typescript
startActivity("正在规划", "正在分析需求");
advanceActivity("正在确认", "正在检查字段", "需求已确认");
updateActivityDetail("正在读取 Sheet1");
completeActivity();
```

**注意事项：**

- `startActivity` 和 `completeActivity` 会重置秒数
- `advanceActivity` 可追加已完成步骤
- `useEffect` 按秒更新 `activitySeconds`

---

### useUndoSnapshot

**职责：** 撤销快照管理和撤销执行。

**导入：**

```typescript
import { useUndoSnapshot } from "./hooks/useUndoSnapshot";
```

**接口：**

```typescript
const {
  lastUndoSnapshot,
  setLastUndoSnapshot,
  clearUndoSnapshot,
  undoLastExecution
} = useUndoSnapshot({
  isBusy,
  onStatusChange: setStatus,
  onMessage: (text) => appendMessage({ role: "system", text }),
  onAfterUndo: () => scan()
});
```

**使用示例：**

```typescript
const handleUndo = async () => {
  await undoLastExecution();
};
```

**注意事项：**

- 当前工作簿结构变化时应调用 `clearUndoSnapshot`
- `undoLastExecution` 会自行管理 `executing` / `idle` 状态

---

### useImageAttachments

**职责：** 图片附件、拖拽事件和数量限制管理。

**导入：**

```typescript
import { useImageAttachments } from "./hooks/useImageAttachments";
```

**接口：**

```typescript
const {
  pendingImages,
  imageError,
  draggingImage,
  addImage,
  removeImage,
  clearImages,
  handleDrop,
  handleDragEnter,
  handleDragLeave,
  handleDragOver,
  setImageError,
  setDraggingImage
} = useImageAttachments();
```

**使用示例：**

```typescript
const handleFileInput = async (file: File) => {
  await addImage(file);
};
```

**注意事项：**

- 拖拽文件会按当前状态截断，防止超过 `MAX_IMAGE_ATTACHMENTS`
- `prepareImageFile` 负责格式、大小和尺寸验证

---

### useModelManagement

**职责：** 模型选择、连接 CRUD 和设置管理。

**导入：**

```typescript
import { useModelManagement } from "./hooks/useModelManagement";
```

**接口：**

```typescript
const {
  selectedModelId,
  modelSettings,
  apiKeyDraft,
  showApiKey,
  connectionDraft,
  pendingDeleteConnectionId,
  settingsSaving,
  settingsTesting,
  settingsLoading,
  settingsFeedback,
  modelGuideDismissed,
  selectModel,
  dismissModelGuide,
  openSettings,
  openConnectionCreator,
  saveApiKey,
  editModelConnection,
  verifyConnection,
  saveConnection,
  removeConnection,
  saveFormulaModel,
  setSelectedModelId,
  setModelSettings,
  setApiKeyDraft,
  setShowApiKey,
  setConnectionDraft,
  setPendingDeleteConnectionId,
  setSettingsSaving,
  setSettingsTesting,
  setSettingsLoading,
  setSettingsFeedback,
  setModelGuideDismissed
} = useModelManagement({
  refreshServiceHealth: refreshServiceState
});
```

**使用示例：**

```typescript
const handleAddConnection = async () => {
  await openConnectionCreator();
};
```

**注意事项：**

- `refreshServiceHealth` 由 `useServiceHealth` 提供
- API Key 完整值不会从服务端读回前端
- 保存连接后会自动刷新模型目录并选择可用模型

---

### useToolManagement

**职责：** 工作流工具、查询工具和工具抽屉管理。

**导入：**

```typescript
import { useToolManagement } from "./hooks/useToolManagement";
```

**接口：**

```typescript
const {
  tools,
  queryTools,
  selectedToolId,
  selectedQueryToolId,
  toolDrawerView,
  toolDetailMode,
  pendingToolDeletion,
  copiedToolDslId,
  toolParameterValues,
  saveTool,
  saveQueryTool,
  deleteTool,
  deleteQueryTool,
  requestToolDeletion,
  confirmToolDeletion,
  resetToolDrawer,
  openWorkflowToolDetail,
  openQueryToolDetail,
  selectTool,
  fieldOptions,
  updateToolParameter,
  copyToolDsl,
  setTools,
  setQueryTools,
  setSelectedToolId,
  setSelectedQueryToolId,
  setToolDrawerView,
  setToolDetailMode,
  setPendingToolDeletion,
  setCopiedToolDslId,
  setToolParameterValues
} = useToolManagement({
  workbook,
  onToolDslCopyError
});
```

**使用示例：**

```typescript
const handleSelectTool = (tool: SavedTool) => {
  selectTool(tool, workbook);
  setToolDrawerView("run");
};
```

**注意事项：**

- `selectTool` 会根据工作表快照准备参数值
- `updateToolParameter` 会联动更新 range 和 field 参数
- DSL 复制反馈定时器会在卸载时清理

---

### useWorkbookContext

**职责：** 工作簿快照、结构监听和文件夹上下文管理。

**导入：**

```typescript
import { useWorkbookContext } from "./hooks/useWorkbookContext";
```

**接口：**

```typescript
const {
  workbook,
  selectedSheetNames,
  selectionConfirmed,
  sourceMode,
  workbookScopeMode,
  folderCatalog,
  folderSheetKeys,
  workbookDataPeriod,
  scan,
  toggleSheet,
  toggleFolderSheet,
  browseFolder,
  chooseAutomaticScope,
  chooseManualScope,
  chooseFolderScope,
  getFolderSelections,
  setWorkbook,
  setSelectedSheetNames,
  setSelectionConfirmed,
  setSourceMode,
  setWorkbookScopeMode,
  setFolderCatalog,
  setFolderSheetKeys
} = useWorkbookContext();
```

**使用示例：**

```typescript
const handleRefreshWorkbook = async () => {
  await scan({ announce: true });
};
```

**注意事项：**

- 当前工作簿模式下，`scan` 会通过 `captureWorkbookStructure` 更新快照
- 文件夹模式下，`browseFolder` 会重新选择文件夹并清空已选表
- 该 Hook 目前与 `useScopeSelection` 存在职责重叠，正式接入 `App.tsx` 前应选择其一并统一

---

### useSlashCommands

**职责：** 斜杠命令检测、自动补全和模式切换。

**导入：**

```typescript
import { useSlashCommands, filterSlashCommands } from "./hooks/useSlashCommands";
```

**接口：**

```typescript
const {
  showAutocomplete,
  filter,
  mode,
  detectSlashCommand,
  closeAutocomplete,
  enterModelMode,
  exitModelMode,
  reset,
  setShowAutocomplete,
  setFilter,
  setMode
} = useSlashCommands();
```

**使用示例：**

```typescript
const visible = detectSlashCommand(text, cursorPosition);
const matches = filterSlashCommands(commands, filter);
```

**注意事项：**

- `/` 后出现空格时，命令模式会关闭补全
- `/model` 模式允许后续空格并进入模型选择

---

### useCopyFeedback

**职责：** 消息和公式复制的反馈状态与定时器管理。

**导入：**

```typescript
import { useCopyFeedback } from "./hooks/useCopyFeedback";
```

**接口：**

```typescript
const {
  copiedMessageId,
  copiedFunctionPreviewId,
  copyMessageText,
  copyFunctionFormula
} = useCopyFeedback({
  onMessageCopyError: () => {
    // 处理消息复制失败
  },
  onFormulaCopyError: (messageId) => {
    // 处理公式复制失败
  }
});
```

**使用示例：**

```typescript
await copyMessageText(messageId, text);
await copyFunctionFormula(messageId, formula);
```

**注意事项：**

- Hook 会在卸载时清理两个反馈定时器
- 空文本不会触发剪贴板操作

---

### useServiceHealth

**职责：** 本地服务健康轮询和模型目录加载。

**导入：**

```typescript
import { useServiceHealth } from "./hooks/useServiceHealth";
```

**接口：**

```typescript
const {
  serverOnline,
  serviceHealth,
  modelOptions,
  modelCatalogLoaded,
  refreshServiceState,
  markServerOnline,
  markServerOffline,
  setServerOnline,
  setServiceHealth,
  setModelOptions,
  setModelCatalogLoaded
} = useServiceHealth();
```

**使用示例：**

```typescript
const handleRefresh = async () => {
  const catalog = await refreshServiceState();
  if (catalog) {
    console.log(catalog.modelOptions);
  }
};
```

**注意事项：**

- 每 5 秒轮询一次 `/health`
- 窗口获得焦点时会立即刷新
- `listModels` 失败时仍保留健康状态

---

## 常见问题

### Q: 如何添加新的 Hook？

A: 遵循以下流程：

1. 在 `apps/excel-addin/src/hooks/` 新建文件。
2. 定义清晰的选项和返回类型。
3. 使用 `useCallback` 稳定回调，完整声明依赖。
4. 对副作用添加 cleanup。
5. 编写对应 `*.test.ts`。
6. 运行 `npm run test:addin` 和 `npm run build:addin`。
7. 更新本文档。

### Q: Hook 之间如何通信？

A: 优先通过 props 或回调传递。例如 `useModelManagement` 接收 `refreshServiceHealth`，而不是直接调用 `useServiceHealth`。跨 Hook 共享的确定性逻辑应提取到 `utils.ts` 或独立服务模块。

### Q: 如何测试使用了 Hooks 的组件？

A: 使用 `@testing-library/react` 的 `renderHook` 和 `act`。对外部模块使用 `vi.mock`。涉及异步副作用时，先 `await act(async () => { ... })` 再断言。

## 最佳实践

1. 单一职责：每个 Hook 只管理一类状态。
2. 依赖完整：`useCallback` / `useEffect` 依赖数组必须完整。
3. 清理副作用：定时器、事件监听器、订阅必须清理。
4. 类型安全：避免使用 `any`，明确类型定义。
5. 测试覆盖：新增 Hook 必须编写单元测试。
6. 文档同步：接口或职责变化时更新本指南。

## 参考资料

- React 官方文档：https://react.dev/reference/react
- Testing Library：https://testing-library.com/docs/react-testing-library/intro/
- 项目代码审查报告：`docs/CODE_REVIEW_REPORT.md`
