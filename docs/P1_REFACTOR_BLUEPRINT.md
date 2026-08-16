# P1 重构蓝图 - 整体架构设计

**设计者：** Kiro  
**执行者：** Codex  
**当前状态：** App.tsx 6345 行，30 个 useState

---

## 📊 当前进度

### ✅ 已完成 Hooks (5/10)
1. **useImageAttachments** - 图片附件管理 (Kiro)
2. **useSlashCommands** - 斜杠命令检测 (Kiro)
3. **useServiceHealth** - 服务健康检查 (Kiro)
4. **useWorkbookContext** - 工作簿上下文管理 (Kiro)
5. **useModelManagement** - 模型管理 (Codex)
6. **useToolManagement** - 工具管理 (Codex)
7. **useConversation** - 对话管理 (Kiro)
8. **useUIState** - UI 状态管理 (Kiro)

**成果：** 已删除 ~600 行代码，useState 从 69 → 30

---

## 🎯 剩余状态分类与提取策略

### 第一优先级：执行审批状态 (4 个 useState)

**Hook 名称：** `useExecutionApproval`

**状态清单：**
```typescript
const [saveCandidate, setSaveCandidate] = useState<AnalysisPlan | null>(null);
const [approveFixedContent, setApproveFixedContent] = useState(false);
const [approveDestructive, setApproveDestructive] = useState(false);
const [verifiedPlanIds, setVerifiedPlanIds] = useState<Set<string>>(() => new Set());
```

**职责：**
- 管理方案保存候选（saveCandidate）
- 管理用户审批状态（approveFixedContent, approveDestructive）
- 跟踪已验证的方案 ID（verifiedPlanIds）

**提取难度：** ⭐⭐ 中等（状态独立，逻辑清晰）

---

### 第二优先级：范围选择状态 (5 个 useState)

**Hook 名称：** `useScopeSelection`

**状态清单：**
```typescript
const [toolName, setToolName] = useState("");
const [toolDescription, setToolDescription] = useState("");
const [selectedSheetNames, setSelectedSheetNames] = useState<string[]>([]);
const [selectionConfirmed, setSelectionConfirmed] = useState(false);
const [sourceMode, setSourceMode] = useState<SourceMode>("workbook");
const [workbookScopeMode, setWorkbookScopeMode] = useState<WorkbookScopeMode>("auto");
const [folderCatalog, setFolderCatalog] = useState<FolderCatalog | null>(null);
const [folderSheetKeys, setFolderSheetKeys] = useState<string[]>([]);
```

**职责：**
- 管理数据源模式（workbook/folder）
- 管理工作簿范围模式（auto/selected）
- 管理工作表选择状态
- 管理文件夹目录和表键
- 管理工具保存时的名称/描述

**提取难度：** ⭐⭐⭐ 较高（状态间有依赖关系）

---

### 第三优先级：活动进度状态 (1 个 useState + derived)

**Hook 名称：** `useActivityProgress`

**状态清单：**
```typescript
const [activityProgress, setActivityProgress] = useState<ActivityProgress | null>(null);
const [activitySeconds, setActivitySeconds] = useState(0);
```

**职责：**
- 管理执行进度状态
- 管理进度计时器
- 提供进度更新函数（startActivity, advanceActivity, completeActivity）

**提取难度：** ⭐⭐ 中等（需要提取相关函数）

---

### 第四优先级：复制反馈状态 (2 个 useState + timers)

**Hook 名称：** `useCopyFeedback`

**状态清单：**
```typescript
const [copiedMessageId, setCopiedMessageId] = useState<string | null>(null);
const [copiedFunctionPreviewId, setCopiedFunctionPreviewId] = useState<string | null>(null);
const copyFeedbackTimerRef = useRef<number | null>(null);
const functionCopyTimerRef = useRef<number | null>(null);
```

**职责：**
- 管理消息复制反馈状态
- 管理函数预览复制反馈状态
- 自动清除反馈定时器

**提取难度：** ⭐ 简单（状态独立，逻辑简单）

---

### 第五优先级：撤销快照 (1 个 useState)

**Hook 名称：** `useUndoSnapshot`

**状态清单：**
```typescript
const [lastUndoSnapshot, setLastUndoSnapshot] = useState<ExecutionUndoSnapshot | null>(null);
```

**职责：**
- 管理最后一次执行的撤销快照
- 提供撤销操作函数

**提取难度：** ⭐ 简单（状态独立）

---

### 暂不提取：核心业务状态 (保留在 App.tsx)

**保留原因：** 这些状态是 App.tsx 核心业务逻辑，提取后反而会增加复杂度

```typescript
// 对话输入状态 - 与 UI 交互紧密耦合
const [prompt, setPrompt] = useState("");
const [status, setStatus] = useState<Status>("idle");

// 斜杠命令状态 - 与输入框交互紧密耦合
const [showSlashAutocomplete, setShowSlashAutocomplete] = useState(false);
const [slashFilter, setSlashFilter] = useState("");
const [slashMode, setSlashMode] = useState<"command" | "model">("command");

// 意图澄清状态 - 对话流程核心
const [clarification, setClarification] = useState<IntentClarification | null>(null);
const [clarificationImages, setClarificationImages] = useState<PendingImage[]>([]);

// Ref 状态 - 不占用 useState 计数
// messageEndRef, composerInputRef, imageInputRef, etc.
```

---

## 📋 执行计划

### Codex 任务队列

**Task 4: useExecutionApproval Hook** ⏱️ 2h
- 创建 `hooks/useExecutionApproval.ts`
- 提取 4 个审批相关状态
- 提供审批操作函数
- 集成到 App.tsx
- 验证编译和测试

**Task 5: useScopeSelection Hook** ⏱️ 3h  
- 创建 `hooks/useScopeSelection.ts`
- 提取 8 个范围选择状态
- 处理状态间依赖关系
- 集成到 App.tsx
- 验证编译和测试

**Task 7: useActivityProgress Hook** ⏱️ 2h
- 创建 `hooks/useActivityProgress.ts`
- 提取进度状态和相关函数（startActivity, advanceActivity, completeActivity）
- 处理定时器逻辑
- 集成到 App.tsx
- 验证编译和测试

**Task 8: useCopyFeedback Hook** ⏱️ 1h
- 创建 `hooks/useCopyFeedback.ts`
- 提取 2 个复制反馈状态和定时器
- 自动清理定时器
- 集成到 App.tsx
- 验证编译和测试

**Task 9: useUndoSnapshot Hook** ⏱️ 1h
- 创建 `hooks/useUndoSnapshot.ts`
- 提取撤销快照状态
- 提供撤销操作函数
- 集成到 App.tsx
- 验证编译和测试

---

## 🎯 目标检查点

**当前：**
- App.tsx: 6345 行
- useState: 30 个

**目标：**
- App.tsx: ~5800 行 (-500 行，-8%)
- useState: ~15 个 (-15 个，-50%)

**关键收益：**
- 状态管理清晰分离
- 业务逻辑更易理解
- 单元测试覆盖面更广
- 后续维护成本降低

---

## 📝 执行规范

**给 Codex 的指令：**

1. **每个 Hook 独立完成**
   - 创建 Hook 文件
   - 集成到 App.tsx
   - 运行 `npm run build:addin` 验证
   - 运行 `npm run test:addin` 验证
   - 统计代码行数变化
   - 更新 P1_TASK_ASSIGNMENT.md

2. **保持行为一致性**
   - 不改变任何业务逻辑
   - 保持原有的状态初始化逻辑
   - 保持原有的副作用（useEffect）逻辑
   - 只做代码搬运和封装

3. **代码质量标准**
   - 每个 Hook 添加 JSDoc 注释说明职责
   - 导出必要的类型定义
   - 使用 useCallback 包装回调函数
   - 清理 App.tsx 中未使用的 imports

4. **遇到问题立即反馈**
   - 状态依赖关系不清晰
   - 函数调用链过于复杂
   - 类型定义冲突
   - 测试失败

---

## 🚀 开始执行

**Codex，请按以下顺序执行：**

1. Task 4: useExecutionApproval Hook
2. Task 5: useScopeSelection Hook  
3. Task 7: useActivityProgress Hook
4. Task 8: useCopyFeedback Hook
5. Task 9: useUndoSnapshot Hook

每完成一个任务，报告：
- Hook 文件路径和行数
- App.tsx 行数变化
- 删除的重复代码统计
- 编译和测试结果

**开始执行 Task 4。**
