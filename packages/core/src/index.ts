// 宿主无关共享层。分层顺序：contracts → pure → semantics → ports → orchestration → ui，
// 依赖只能从上往下，不允许反向边。
//
// L0 contracts
export type { CellValue, WorkbookLocator, AuthorizedSheetRef } from "./contracts";
export type {
  WorksheetToolParameter,
  OutputWorksheetToolParameter,
  FieldToolParameter,
  RangeToolParameter,
  ToolParameter,
  ToolRiskApproval,
  ToolEligibilityIssue,
  ToolEligibility,
  SavedTool
} from "./tools";

// L1 pure
export type { A1Rect } from "./planCore";
export {
  parseExplicitA1Rect,
  assertRectangularValues,
  writeMatrixShapeIssue,
  blankCellMatrix
} from "./planCore";
export {
  firstCellOfRange,
  columnLetterFromIndex,
  columnIndexFromLetter,
  startColumnOfRange,
  suggestWriteTarget
} from "./range";
export {
  normalizeField,
  detectHeaderIndex,
  detectSheetFields,
  looksLikeHeaderRow,
  isSyntheticField,
  syntheticFields
} from "./tableSchema";
// `query_table` 的显式只读区域：写法判定、交集与裁剪三件事只在这里定。
export type { AbsoluteRect, QueryRect } from "./queryRange";
export {
  parseQueryRange,
  queryRangeIssue,
  intersectQueryRect,
  sliceForQueryRange,
  queryRectAddress
} from "./queryRange";
// 「这次查哪几张表」的唯一判据：授权集合是上限，不是本次查询范围。
export type {
  QueryTargetRequest,
  QueryTargetResolution
} from "./queryTarget";
export { resolveQueryTarget } from "./queryTarget";
// 结果"为什么不完整"的唯一出处：原因由生产它的代码写下，文案只负责复述。
export type {
  QueryCompleteness,
  QueryIncompleteReason,
  QueryTruncation
} from "./queryCompleteness";
export {
  completenessOf,
  describeIncompleteReason,
  describeIncompleteReasons,
  describeTruncation
} from "./queryCompleteness";
// 表格区域变换（排序 / 去重）的确定性判据：两端读同一份，避免顺序语义漂移。
export type { SortKeySpec } from "./tableTransforms";
export {
  compareSortValues,
  sortRowOrder,
  reorderRows,
  sortedMatrix,
  duplicateRowPartition,
  deduplicatedMatrix
} from "./tableTransforms";
// 写进任何用户可见 / 可持久化文本之前，先过一遍凭据脱敏。
export {
  REDACTED_PLACEHOLDER,
  redactCredentials,
  containsLikelyCredential
} from "./redact";

// 双栏变更对照的模型与纯比较。三端共用同一套空值/公式语义，
// 出口给活动日志或聊天历史时一律走 changeReviewSummary，不让单元格内容外泄。
export type {
  ChangeValue,
  ChangeCell,
  ChangeSnapshot,
  ChangePhase,
  ChangeReviewStatus,
  ChangeKind,
  ChangeReviewRow,
  ChangeReviewLimit,
  RangeChangeReview,
  ChangeReviewSummary,
  CellDisplayText,
  BuildRangeChangeReviewInput
} from "./changeReview";
export {
  DEFAULT_CHANGE_REVIEW_LIMIT,
  VALUE_NOT_CAPTURED_TEXT,
  buildRangeChangeReview,
  canClaimWriteSuccess,
  cellDisplayText,
  changeReviewSummary,
  onlyChangedRows
} from "./changeReview";
// 「实际写入了几次」的唯一口径：请求次数、认领结果、宿主写调用数、结果与回读
// 各自能证明什么。两端执行器与验收说明读这一份，不各写一套。
export type {
  HostWriteAudit,
  HostWriteObservation,
  HostWriteObservationVerdict,
  HostWriteSummary
} from "./hostWriteAudit";
export {
  describeHostWriteObservation,
  normalizeHostWriteAudit,
  summarizeHostWrites
} from "./hostWriteAudit";

// 常用格式操作的宿主无关语义：要改哪些属性、真实影响范围（行高=整行 / 列宽=整列）、
// 以及「属性 × 修改前/修改后」对照。三端共用，宿主 API 差异留在各端适配层。
export type {
  FormatActionType,
  FormatActionLike,
  FormatPropertyKey,
  FormatScalar,
  FormatEdit,
  FormatImpact,
  FormatSnapshot,
  FormatChangeRow,
  FormatChangeReview,
  FormatChangeSummary,
  BuildFormatChangeReviewInput,
  FormatPlanPropertyGroup,
  FormatPlanPreviewEntry
} from "./formatActions";
export {
  FORMAT_PROPERTY_ORDER,
  FORMAT_PROPERTY_LABELS,
  FORMAT_VALUE_NOT_CAPTURED_TEXT,
  FORMAT_IMPACT_UNRESOLVED_TEXT,
  isFormatActionType,
  asFormatAction,
  selectFormatActions,
  borderDescription,
  formatActionEdits,
  formatActionImpact,
  formatImpactNotice,
  formatValueText,
  buildFormatChangeReview,
  formatChangeSummary,
  formatPlanPropertyGroups,
  formatPlanPreview
} from "./formatActions";

// 审阅标记的呈现方式判据（浅黄底色 vs 只边框），两端标记实现读同一份。
export type { ReviewMarkStyle } from "./reviewMarkStyle";
export {
  actionTouchesFill,
  planTouchesFill,
  reviewMarkStyleForAction,
  REVIEW_MARK_COLORS,
  REVIEW_MARK_LEGEND
} from "./reviewMarkStyle";

// L2 semantics：身份折叠是唯一允许拼接/解析 `url:` 前缀的地方。
export {
  officeIdentityFromUrl,
  officeIdentityFromLocator,
  locatorFromOfficeIdentity,
  sameLocator
} from "./locator";
// 与 server/app/column_mapping.py 镜像，一致性由 config/column_mapping_cases.json 锁住。
export type {
  ColumnMappingInput,
  ColumnMappingAction,
  ColumnMappingClearAction,
  ColumnMappingPlanAction,
  ColumnMappingRankEntry,
  ColumnMappingClarification,
  ColumnMappingClarificationReason,
  ColumnMappingResolution,
  MappingCellGrid
} from "./columnMapping";
export {
  ColumnMappingError,
  UnmappedCategoryError,
  NeedsClarificationError,
  ShapeMismatchError,
  OverlappingRangeError,
  BlankPolicyNeededError,
  BlankSplitLimitError,
  DEFAULT_MAX_MAPPING_ACTIONS,
  mapColumnValues,
  mapColumnValuesWithOutcomes,
  splitTargetActions,
  buildColumnMappingActions
} from "./columnMapping";
// 空白来源格的四种处理意图必须分开表达：一个 `blankLabel` 字符串表达不了
// 「保留目标原值」，也分不清「清空」与「写空串」。
export type {
  BlankPolicy,
  BlankCellOutcome,
  BlankPolicyResolution
} from "./blankPolicy";
export {
  BlankPolicyError,
  UNDECIDED_BLANK_POLICY,
  BLANK_POLICY_LABEL,
  BLANK_POLICY_CHOICE_TEXT,
  blankPolicyFromLegacy,
  parseBlankPolicy,
  isBlankValue,
  blankCellOutcome,
  blankCellPositions
} from "./blankPolicy";
export type { ColumnMappingFormulaCandidate } from "./columnMappingFormula";
export { columnMappingFormulaCandidate } from "./columnMappingFormula";
// 把「待澄清的归因」翻成一张卡片的内容。归因不同、能做的动作不同 ——
// 分隔符已经对上时给出"补分隔符"只会让用户去改一个没问题的设置。
export type { ColumnMappingClarificationCard } from "./columnMappingClarification";
export { columnMappingClarificationCard } from "./columnMappingClarification";

// L3 ports：只有接口，没有实现。
export type {
  HostId,
  KeyValueStore,
  ActionCapability,
  CapabilityResolver,
  HostPrimitiveOps,
  HostAdapter,
  ApiClient,
  Clock
} from "./ports";

// L2 semantics：宿主 × 版本 × 动作 的可信上界与交集（方案 §4.2/§4.3）。
// 与 server/app/action_capability.py 镜像，一致性由 config/action_capability_cases.json 锁住。
export type {
  HostActionTable,
  HostCapabilityMatrix,
  CapabilityMatrix
} from "./actionCapability";
export {
  matchesVersionPrefix,
  parseHostId,
  resolveHostActionBound,
  intersectActionCapabilities,
  resolveAllowedActionTypes,
  actionCapabilityKeys,
  isActionAllowedByKeys
} from "./actionCapability";

// L5 ui：任务窗格的口径（计划卡/结果卡的标题与标签）。
// Excel/WPS 两端共用，避免同一句话在两处各自漂移；证据内容不在此列。
export {
  PLAN_NOTE_SECTIONS,
  PLAN_FORMAT_SECTION,
  PLAN_FORMAT_IMPACT_LABEL,
  SESSION_RECOVERY_LABEL,
  SESSION_RECOVERY_NOTICE_PREFIX,
  VERIFICATION_PASSED_LABEL
} from "./planInformationHierarchy";

// L2 semantics：澄清的**答案**与**授权动作**分开。重新确认数据范围是授权，
// 不是回答那道题 —— 写进 `clarification.answer` 会伪造证据。
export type {
  ClarificationAnswerCard,
  KeyedConfirmation
} from "./clarificationAnswers";
export type { AgentTaskContext } from "./agentTaskContext";
export { createAgentTaskContext, confirmAgentTaskDecision } from "./agentTaskContext";
// SDK 任务工具桥的共享状态机：把不透明进度指针翻译成轮询、跟踪待回执调用，
// 并在发网络请求之前判定 accepted / duplicate / conflict / late。两端面板共用。
export type {
  BridgeHostId,
  BridgeToolName,
  SdkBridgeBinding,
  BridgeToolPointer,
  PendingBridgeCall,
  BridgePollResponse,
  BridgeReceiptResponse,
  BridgeReceiptDecision,
  SdkBridgeState,
  BridgeCallOutcome,
  BridgeAuthorizationVerdict,
  BridgeLiveAuthorization,
  BridgeHostPorts,
  BridgeOutcomeEntry,
  BridgeServiceResult
} from "./sdkToolBridge";
export {
  createSdkBridgeState,
  isBridgeToolName,
  bridgeToolPointer,
  applyBridgeProgressEvent,
  assertBridgePollResponse,
  applyBridgePollResponse,
  pendingBridgeCall,
  decideBridgeReceipt,
  markBridgeDelivered,
  cancelBridge,
  shouldPollBridge,
  bridgeWorkbookIdentity,
  bridgeRequestedSheetNames,
  BRIDGE_UNVERIFIABLE_ERROR_CODE,
  verifyBridgeAuthorization,
  verifyBridgeAuthorizationOrRefuse,
  serviceBridgeOnce,
  runBridgeUntilIdle
} from "./sdkToolBridge";
export {
  REAUTHORIZATION_NOTICE,
  clarificationAfterReauthorization,
  mergeKeyedConfirmation,
  renderConfirmedPrompt
} from "./clarificationAnswers";

// L2 semantics：「任务中断后继续」的共享口径。
// 阶段 / 阻塞原因 / 可续跑性三件事分开表达；持久化走显式白名单消毒，
// 单元格内容与宿主对象引用结构上无法进入该记录。
export type {
  TaskPhase,
  TaskBlocker,
  TaskResumability,
  TaskWriteState,
  TaskResumeConfirmation,
  TaskResumeRecord,
  TaskResumeAction,
  TaskResumeDecision
} from "./taskResume";
export {
  TASK_RESUME_SCHEMA_VERSION,
  TASK_RESUME_MAX_CONFIRMATIONS,
  TASK_RESUME_MAX_TEXT,
  TASK_PHASE_LABEL,
  TASK_BLOCKER_LABEL,
  TASK_RESUMABILITY_LABEL,
  TASK_WRITE_STATE_LABEL,
  taskResumeDecision,
  sanitizeTaskResumeRecord,
  isTaskResumeRecord,
  newTaskId,
  withTaskResumeState,
  withTaskResumeConfirmation,
  planSettlesExecution
} from "./taskResume";

// L2 semantics：规划流式消费与检查点续跑的身份/游标/终态口径。
// 两端共用，避免各自再写一套会漂移的规划状态机。权威现场仍在服务端
// `planning_tasks`；本模块只做归并、游标判定与"下一步能做什么"的推导。
export type {
  PlanningStreamPhase,
  PlanningTaskStatus,
  PlanningPauseType,
  PlanningStepEventView,
  PlanningCursor,
  PlanningEventApplication,
  PlanningResumeDecision,
  PlanningStatusView,
  PlanningContinuationView,
  PlanningStepLifecycle,
  PlanningPausedEntry
} from "./planningStream";
export {
  MODEL_ERROR_CODE_PREFIX,
  PLANNING_RESUME_HOST_WRITES,
  newPlanningCursor,
  planningAfterSequence,
  applyPlanningEvent,
  reconcilePlanningCursor,
  planningResumeDecision,
  isSamePlanningResume,
  resolveTurnIdForNewRequest,
  newPlanningRequestId,
  closeDanglingSteps,
  planningPausedEntriesFromHistory
} from "./planningStream";

export type {
  PlanningContinuationEvidence,
  PlanningEvidenceFailure,
  PlanningEvidenceInput,
  PlanningEvidenceResult,
  PlanningSourceVerification
} from "./planningContinuationEvidence";
export {
  buildPlanningContinuationEvidence,
  PLANNING_EVIDENCE_FAILURE_MESSAGES
} from "./planningContinuationEvidence";

export type {
  SourceDependencyOrigin,
  SourceDependencyRecord
} from "./planSourceDependency";
export {
  recordTrustedToolReads,
  trustedToolReads,
  resetTrustedSourceReads,
  attestSourceDependency
} from "./planSourceDependency";
export { attestResultSourceDependency } from "./planSourceDependency";

export type {
  SourceAnchorPlan,
  PlanSourceAnchor,
  PlanSourceVerdict,
  PlanSourceAnchorProblem,
  PlanSourceGateVerdict
} from "./planSourceAnchor";
export {
  PLAN_SOURCE_STALE_MESSAGE,
  PLAN_SOURCE_UNVERIFIABLE_MESSAGE,
  PLAN_SOURCE_MISSING_MESSAGE,
  PLAN_SOURCE_SCOPE_MISMATCH_MESSAGE,
  BAKED_WRITE_ACTION_TYPES,
  planHasBakedValues,
  planSourceComparisonNeeded,
  judgePlanSourceGate,
  anchorSheetsWithin,
  planSourceGateMessage,
  planSourceAnchor,
  comparePlanSourceAnchor,
  sourceAnchorFromBinding,
  stampPlanSourceAnchor,
  sameSheetSet
} from "./planSourceAnchor";

export type {
  PendingExecutionNamespace,
  PendingExecutionStatus,
  PendingExecutionCandidate,
  PendingExecutionDiscoveryResponse,
  PendingExecutionSelection
} from "./pendingExecutionDiscovery";
export {
  PENDING_DISCOVERY_SCHEMA_VERSION,
  PENDING_DISCOVERY_LIMIT,
  PENDING_DISCOVERY_MAX_IDENTITY_CHARS,
  PENDING_DISCOVERY_CHOOSE_DETAIL,
  PENDING_DISCOVERY_TRUNCATED_DETAIL,
  PENDING_DISCOVERY_UNAVAILABLE_DETAIL,
  assertPendingExecutionDiscovery,
  orderPendingExecutionCandidates,
  pendingExecutionCandidateLabel,
  pendingExecutionSelection
} from "./pendingExecutionDiscovery";

export type {
  PendingResumeBlockReason,
  PendingResumeReadiness,
  PendingResumeInput
} from "./pendingExecutionResume";
export {
  pendingResumeReadiness,
  pendingResumeTargetSheets
} from "./pendingExecutionResume";

export type {
  ExecutionRecoveryAction,
  ExecutionRecoveryKind,
  ExecutionReconcileOutcome,
  ExecutionRecoverySourceVerdict,
  ExecutionRecoveryTerminalReason,
  ExecutionRecoveryVerdict
} from "./executionRecoveryView";
export {
  EXECUTION_RECOVERY_ACTION_LABEL,
  EXECUTION_RECOVERY_ENTRY_HINT,
  EXECUTION_RECOVERY_NOTE,
  EXECUTION_RECOVERY_STATE_LABEL,
  EXECUTION_RECOVERY_TERMINAL_SUMMARY,
  EXECUTION_RECOVERY_TITLE,
  EXECUTION_RECOVERY_UNDONE_DETAIL,
  EXECUTION_RECOVERY_UNDONE_NOTE,
  executionRecoveryActions,
  executionRecoveryBadge,
  executionRecoveryDetail,
  executionRecoveryIsTerminal,
  executionRecoveryNote,
  executionRecoveryProgress,
  executionRecoveryResumeVerdict,
  executionRecoveryTerminalReason,
  executionRecoveryTerminalSummary,
  selectExecutionRecoveryEntry
} from "./executionRecoveryView";

export type {
  ResumeSceneInput,
  ResumeSceneObservation,
  ResumeSceneObservationState,
  ResumeSceneOutcome,
  ResumeSceneReason
} from "./executionResumeScene";
export { verifyResumeScene } from "./executionResumeScene";
export type {
  ExecutionDisposition,
  ExecutionDispositionOutcome
} from "./executionDisposition";
export {
  EXECUTION_DISPOSITION_LABEL,
  executionDispositionTarget,
  normalizeExecutionDisposition,
  recordExecutionDisposition
} from "./executionDisposition";

export type {
  UndoRebuildBlockReason,
  UndoRebuildCheckpoint,
  UndoRebuildDisposition,
  UndoRebuildInput,
  UndoRebuildOutcome,
  UndoRebuildRange,
  UndoRebuildTarget
} from "./executionUndoRebuild";
export {
  UNDO_RECOVERY_LABEL,
  rebuildExecutionUndo,
  undoRecoveryDetail
} from "./executionUndoRebuild";

// 专注窗口跨窗口协议（Excel DialogApi / WPS WebDialog 共用）
export type {
  FocusCellValue,
  FocusPlanSummary,
  FocusResultTable,
  FocusMessage,
  FocusConversation,
  FocusTool,
  FocusWorkbookRef,
  FocusSnapshot,
  FocusMessageKind,
  FocusEnvelope,
  FocusRejectReason,
  FocusExpectation,
  FocusReadResult,
  FocusChildStatus,
  FocusChildState,
  FocusChildEvent
} from "./focusProtocol";
export {
  FOCUS_PROTOCOL_CHANNEL,
  FOCUS_PROTOCOL_VERSION,
  FOCUS_HEARTBEAT_INTERVAL_MS,
  FOCUS_SYNC_TTL_MS,
  FOCUS_HELLO_RETRY_MS,
  FOCUS_HANDSHAKE_TIMEOUT_MS,
  FOCUS_CACHE_KEY_PREFIX,
  FOCUS_CACHE_TTL_MS,
  FOCUS_SNAPSHOT_LIMITS,
  FOCUS_REJECT_TEXT,
  createFocusSessionId,
  focusCacheKey,
  parseFocusEnvelope,
  sanitizeFocusSnapshot,
  readFocusMessage,
  initialFocusChildState,
  reduceFocusChild,
  nextFocusSeq,
  createFocusEnvelope
} from "./focusProtocol";
