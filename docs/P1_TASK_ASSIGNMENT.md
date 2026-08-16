# P1 重构任务分配

**目标：** 将 App.tsx 的状态管理和业务逻辑提取到 Hooks 和服务类

**协作模式：** Kiro (我) + Codex 并行工作

---

## ✅ 已完成 (Kiro)

1. **useImageAttachments** - 图片附件管理 ✅ 已集成
2. **useSlashCommands** - 斜杠命令检测 ✅ 已集成
3. **useServiceHealth** - 服务健康检查 ✅ 已集成
4. **useWorkbookContext** - 工作簿上下文管理 ✅ 已集成
5. **useModelManagement** - 模型管理 Hook ✅ 已创建（未集成）

**成果：**
- 删除 185 行 App.tsx 代码
- useState 从 69 → 58 个
- 提取 726 行 Hook 代码

---

## 🔨 Codex 任务清单

### Task 1: useModelManagement 集成到 App.tsx ✅ 已完成 ⏱️ 2h

**状态：** 已完成

**完成情况：**
1. ✅ 删除了 12 个重复函数（~237 行代码）
2. ✅ 添加了 3 个 UI 协调包装函数（handleOpenSettings, handleOpenConnectionCreator, handleSelectModel）
3. ✅ 更新了函数调用点（selectModel → handleSelectModel）
4. ✅ 清理了未使用的 imports
5. ✅ 编译通过验证

**代码行数：**
- App.tsx: 4048 行 → 3811 行（-237 行，-5.9%）

**成果：**
- 删除了 openSettings、selectModel、dismissModelGuide、saveApiKey、editModelConnection、verifyConnection、saveConnection、removeConnection、saveFormulaModel、refreshModelsAfterSettings、connectionRequest 等重复函数
- 清理了 deleteModelConnection、getModelSettings、saveModelConnection、setFormulaModel、testModelConnection、updateModelSettings、ModelSettings、UpsertModelConnectionRequest、emptyModelConnectionDraft、chooseAvailableModel 等未使用的 imports
- UI 状态管理和业务逻辑清晰分离

---

### Task 2: useToolManagement Hook 创建 + 集成 ✅ 已完成 ⏱️ 3h

**状态：** 已完成 (Codex)

**完成情况：**
1. ✅ 创建了 `apps/excel-addin/src/hooks/useToolManagement.ts`（342 行）
2. ✅ 提取了 9 个工具管理状态（tools, queryTools, selectedToolId, selectedQueryToolId, toolDrawerView, toolDetailMode, pendingToolDeletion, copiedToolDslId, toolParameterValues）
3. ✅ 提取了工具操作函数（saveTool, deleteTool, deleteQueryTool, confirmToolDeletion, prepareToolParameters, updateToolParameter, copyToolDsl 等）
4. ✅ 类型定义导出（ToolDrawerView, ToolDetailMode, PendingToolDeletion）
5. ✅ 更新了 App.tsx 和相关组件的调用点
6. ✅ 编译和测试全部通过

**验证结果：**
- `npm run build:addin` ✓ 编译通过
- `npm run test:addin` ✓ 18 个测试文件，150 项测试全部通过

**成果：**
- 工作流工具和查询工具状态集中管理
- 工具抽屉交互逻辑完全封装在 Hook 内
- 工具保存/删除/参数管理/DSL 复制统一提供
- App.tsx 只保留业务协调逻辑
- 组件可直接引用 Hook 的类型定义，无需依赖 App.tsx

---

### Task 3: 工具保存逻辑简化（可选优化） ⏱️ 1h
   - `toolParameterValues` / `setToolParameterValues`

3. 提取相关函数：
   - `saveTool()`
   - `deleteTool()`
   - `deleteQueryTool()`
   - 其他工具 CRUD 操作

4. 集成到 App.tsx

**参考模式：**
```typescript
export function useToolManagement() {
  const [tools, setTools] = useState<SavedTool[]>(loadTools);
  // ... 其他状态
  
  const saveTool = useCallback(async (tool: SavedTool) => {
    // ...
  }, []);
  
  return {
    // 状态
    tools,
    queryTools,
    selectedToolId,
    // ...
    
    // 操作
    saveTool,
    deleteTool,
    // ...
    
    // 内部状态设置
    setTools,
    setQueryTools,
    // ...
  };
}
```

**验证：**
```bash
npm run build:addin
```

---

### Task 3: 工具保存逻辑简化（可选优化） ⏱️ 1h

**状态：** 已完成

**任务：**
- 当前工具保存涉及 `saveCandidate`、`approveFixedContent`、`approveDestructive` 等状态
- 可以考虑将这些状态也纳入 useToolManagement
- 或者创建单独的 useToolSaving Hook

**完成情况：**
1. ✅ 创建 `apps/excel-addin/src/hooks/useExecutionApproval.ts`
2. ✅ 迁移 `saveCandidate`、`approveFixedContent`、`approveDestructive`、`verifiedPlanIds` 四个状态
3. ✅ 派生 `saveEligibility` 固化资格检查并集成到 App.tsx
4. ✅ 提供 `markPlanVerified`、`beginSaveCandidate`、`closeSaveCandidate` 操作
5. ✅ `npm run build:addin` 与 `npm run test:addin` 通过

---

### Task 4: useExecutionApproval Hook ✅ 已完成 ⏱️ 2h

**状态：** 已完成

**完成情况：**
1. ✅ 创建 `apps/excel-addin/src/hooks/useExecutionApproval.ts`
2. ✅ 提取 4 个执行审批状态：
   - `saveCandidate`
   - `approveFixedContent`
   - `approveDestructive`
   - `verifiedPlanIds`
3. ✅ 提取审批相关操作：
   - `markPlanVerified`
   - `beginSaveCandidate`
   - `closeSaveCandidate`
4. ✅ 派生 `saveEligibility` 固化资格检查并集成到 App.tsx
5. ✅ 删除 App.tsx 中对应的 `useState` 声明，改为 Hook 解构使用
6. ✅ `npm run build:addin` 与 `npm run test:addin` 通过

**代码行数：**
- 新增 `useExecutionApproval.ts`：57 行
- App.tsx 中移除 4 个状态声明和 `saveEligibility` 派生逻辑

---

### Task 5: useScopeSelection Hook ✅ 已完成 ⏱️ 3h

**状态：** 已完成

**完成情况：**
1. ✅ 创建 `apps/excel-addin/src/hooks/useScopeSelection.ts`
2. ✅ 提取 8 个范围选择与工具保存字段状态：
   - `toolName`
   - `toolDescription`
   - `selectedSheetNames`
   - `selectionConfirmed`
   - `sourceMode`
   - `workbookScopeMode`
   - `folderCatalog`
   - `folderSheetKeys`
3. ✅ 提取范围操作函数：
   - `applyWorkbookSnapshotSelection`
   - `toggleSheet`
   - `toggleFolderSheet`
   - `applyFolderCatalog`
   - `chooseAutomaticScope`
   - `chooseManualScope`
   - `chooseFolderScope`
   - `folderSelections`
   - `selectedNamesFor`
   - `selectAllSheets`
   - `clearSelectedSheets`
4. ✅ 保持 sourceMode/workbookScopeMode 切换时的选择状态依赖关系
5. ✅ 删除 App.tsx 中对应的 `useState` 声明，改为 Hook 解构使用
6. ✅ `npm run build:addin` 与 `npm run test:addin` 通过

**代码行数：**
- 新增 `useScopeSelection.ts`：162 行
- App.tsx：6149 行 → 6092 行（-57 行）

---

### Task 7: useActivityProgress Hook ✅ 已完成 ⏱️ 2h

**状态：** 已完成

**完成情况：**
1. ✅ 创建 `apps/excel-addin/src/hooks/useActivityProgress.ts`
2. ✅ 提取 2 个活动进度状态：
   - `activity`
   - `activitySeconds`
3. ✅ 提取进度操作函数：
   - `startActivity`
   - `advanceActivity`
   - `updateActivityDetail`
   - `completeActivity`
4. ✅ 提取计时器 `useEffect`，按秒刷新 `activitySeconds`
5. ✅ 提取 activity 日志固化逻辑，通过 `onPersistLog` 回调更新最近一条 assistant 消息
6. ✅ 导出 `ActivityStep`、`ActivityProgress`、`ActivityLog` 类型，并从 App.tsx 保持兼容导出
7. ✅ 删除 App.tsx 中对应的 `useState`、函数和计时器逻辑，改为 Hook 解构使用
8. ✅ `npm run build:addin` 与 `npm run test:addin` 通过

**代码行数：**
- 新增 `useActivityProgress.ts`：121 行
- App.tsx：6092 行 → 6010 行（-82 行）

---

### Task 8: useCopyFeedback Hook ✅ 已完成 ⏱️ 1h

**状态：** 已完成

**完成情况：**
1. ✅ 创建 `apps/excel-addin/src/hooks/useCopyFeedback.ts`
2. ✅ 提取 2 个复制反馈状态和 2 个 timer ref：
   - `copiedMessageId`
   - `copiedFunctionPreviewId`
   - `copyFeedbackTimerRef`
   - `functionCopyTimerRef`
3. ✅ 提取复制操作：
   - `copyMessageText`
   - `copyFunctionFormula`
4. ✅ 添加 `useEffect` cleanup，自动清理两个定时器
5. ✅ 删除 App.tsx 中本地重复的 `copyTextToClipboard`，改用 `utils.ts` 中的版本
6. ✅ 删除 App.tsx 中对应的 `useState`、`useRef`、函数和 cleanup 逻辑
7. ✅ `npm run build:addin` 与 `npm run test:addin` 通过

**代码行数：**
- 新增 `useCopyFeedback.ts`：83 行
- App.tsx：6010 行 → 5952 行（-58 行）

---

### Task 9: useUndoSnapshot Hook ✅ 已完成 ⏱️ 1h

**状态：** 已完成

**完成情况：**
1. ✅ 创建 `apps/excel-addin/src/hooks/useUndoSnapshot.ts`
2. ✅ 提取 `lastUndoSnapshot` 状态
3. ✅ 提取撤销相关操作：
   - `setLastUndoSnapshot`
   - `clearUndoSnapshot`
   - `undoLastExecution`
4. ✅ 将撤销执行、状态切换、提示和刷新逻辑封装到 Hook，通过回调与 App.tsx 协作
5. ✅ 删除 App.tsx 中对应的 `useState`、`undoExecution` 导入和 `undoLastExecution` 函数
6. ✅ `npm run build:addin` 与 `npm run test:addin` 通过

**代码行数：**
- 新增 `useUndoSnapshot.ts`：57 行
- App.tsx：5952 行 → 5946 行（-6 行）

**最终统计：**
- App.tsx 最终行数：5946 行
- App.tsx 剩余 useState 数量：12 个
- 相对当前 HEAD：6691 行 → 5946 行，累计减少 745 行

---

## 🚀 Kiro 任务清单

### Task A: useConversation Hook 创建 ✅ 已完成 ⏱️ 4h

**状态：** 已完成

**完成情况：**
1. ✅ 创建了 `apps/excel-addin/src/hooks/useConversation.ts`（166 行）
2. ✅ 提取了对话管理相关状态和函数
3. ✅ 添加了 UI 协调包装函数（handleNewChat, handleOpenConversation, handleDeleteConversation）
4. ✅ 更新了所有调用点
5. ✅ 修复了类型导出问题（ChatMessage, FunctionPreview, ActivityProgress 等）
6. ✅ 编译通过验证

**删除的重复代码：**
- `setMessages()` - 30 行
- `newChat()` - 36 行  
- `openConversation()` - 14 行
- `deleteConversation()` - 11 行
- `confirmDeleteConversation()` - 7 行
- `pendingDeleteConversationId` 状态声明 - 2 行
- localStorage 持久化 useEffect - 19 行
- **总计：约 119 行**

**成果：**
- 对话状态（chatHistory）和派生状态（activeConversation, messages）集中管理
- localStorage 持久化逻辑封装在 Hook 内
- 对话操作（新建、切换、删除）通过 Hook 统一提供
- UI 状态协调通过包装函数处理
- 类型定义正确导出，其他组件可复用

---

### Task B: 创建进度报告 ⏱️ 0.5h

**任务：**
更新 `docs/P1_IMPLEMENTATION_PROGRESS.md`，记录：
- 已完成的 Hook
- 代码行数变化
- 遇到的问题和解决方案
- 下一步计划

---

## 📋 协作流程

1. **任务认领：** Codex 从上面选择任务开始
2. **进度同步：** 每个任务完成后更新本文档状态
3. **代码审查：** 完成后运行 `npm run build:addin` 验证
4. **提交规范：** 
   - Codex: `feat: integrate useModelManagement hook`
   - Codex: `feat: create and integrate useToolManagement hook`
   - Kiro: `feat: create and integrate useConversation hook`

---

## 📊 总体目标

- **App.tsx**: 从 4000+ 行缩减到 ~3000 行
- **useState**: 从 69 个缩减到 ~40 个 (-40%)
- **关键收益**: 代码组织、可测试性、可维护性

---

## 🎯 优先级

**高优先级（本轮完成）：**
1. Task 1: useModelManagement 集成 ← **Codex 先做这个**
2. Task 2: useToolManagement Hook ← **Codex 接着做这个**
3. Task A: useConversation Hook ← **Kiro 并行做这个**

**中优先级（视情况）：**
- Task 3: 工具保存逻辑优化
- 其他小的 Hook 提取

**低优先级（下一轮）：**
- 业务服务类提取（MessageProcessor, PlanExecutor）
- 深度业务逻辑重构

---

## 📝 注意事项

**给 Codex：**
1. 参考已完成的 Hook 结构（useWorkbookContext 是最好的例子）
2. 每个 Hook 完成后立即编译验证
3. 遇到问题及时反馈，不要卡住
4. 可以先完成简单的，再做复杂的
5. 编译通过 ≠ 功能正确，需要确保逻辑没有变化

**给 Kiro：**
1. 监控 Codex 的进度
2. useConversation 最复杂，需要仔细处理
3. 完成后审查 Codex 的代码
4. 整理最终的重构报告

---

**更新时间：** 2026-08-13
**状态：** 协作进行中
