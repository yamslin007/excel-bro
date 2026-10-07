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
export { normalizeField, detectHeaderIndex, detectSheetFields } from "./tableSchema";
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
  withTaskResumeConfirmation
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
