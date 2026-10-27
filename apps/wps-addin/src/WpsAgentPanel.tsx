import { stampLiteralPlanOrigin } from "@excel-bro/core/literalPlanOrigin";
import { assertWpsPlanSourceCurrent } from "./wpsPlanExecutor";
import { RecoveryProbeControl } from "@excel-bro/core/RecoveryProbeControl";
import { bindRecoveryProbe, interruptRecoveryProbe } from "@excel-bro/core/recoveryProbe";
import { ResultRefFaultControl } from "@excel-bro/core/ResultRefFaultControl";
import { resultRefFaultVisible } from "@excel-bro/core/resultRefFault";
import { wpsResultRefFaultTarget } from "./wpsResultRefFaultTarget";
import { canResumeWpsExecution, resumeWpsExecution, wpsResumeUnavailableReason } from "./wpsExecutionResume";
import { prepareWpsResume, prepareWpsCompletedReview } from "./wpsPlanExecutor";
import { TaskExecutionControls } from "../../excel-addin/src/components/TaskExecutionControls";
import { acceptsPlanningReplay, planningReceivedSequence, waitForPlanning } from "@excel-bro/core/planningReplay";
import {
  PAUSE_UNAVAILABLE_REASON,
  PauseTimeoutError,
  controllerOwnerOf,
  executionPauseState,
  newExecutionId,
  pauseButtonState,
  planExecutionTeardown,
  recordPauseDiagnostic,
  tagControllerOwner,
  tailOf,
  withPauseTimeout
} from "@excel-bro/core/pauseControl";
import {
  armResultRefFault,
  cancelResultRefFault,
  getResultRefFaultStatus,
  pausePlanningTask
} from "../../excel-addin/src/api";
import checkpointCapabilities from "../../../config/capabilities.json";
import { planningRecoveryCandidates, isOfferedPlanningRecovery } from "@excel-bro/core/planningRecoveryRouting";
import { usePrePlanningPause } from "../../excel-addin/src/hooks/usePrePlanningPause";
import { PrePlanningPauseCard } from "../../excel-addin/src/components/PrePlanningPauseCard";
import { PauseDetails } from "../../excel-addin/src/components/PauseDetails";
import { wpsWriteVersionStatus } from "./wpsWriteVersionGate";
import {
  isWorksheetDeletion,
  reconcileWorksheetDeletionScope,
  worksheetDeletionIssue
} from "@excel-bro/core/worksheetDeletion";
import { reconcileWorksheetRestoreScope } from "@excel-bro/core/worksheetRestore";
import { isEBSystemSheet } from "@excel-bro/core/folderScope";
import { useWorksheetCheckpoint } from "../../excel-addin/src/hooks/useWorksheetCheckpoint";
import {
  WorksheetCheckpointCard,
  WorksheetCheckpointRecoverySection
} from "../../excel-addin/src/components/WorksheetCheckpointCard";
import { isFormatActionType } from "./wpsPlanExecutor";
import { agentModelAvailable, agentModelUnavailableReason } from "../../excel-addin/src/modelSelection";
import { Fragment, type ReactNode, useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { ResultReview } from '@excel-bro/core/resultReview';
import { ResultReviewBar } from '../../excel-addin/src/components/ResultReviewBar';
import { ControlledScriptEditor } from '../../excel-addin/src/components/ControlledScriptEditor';
import { markWpsResult, clearWpsResultMarks, recoverWpsResultMarks, type WpsResultMarks } from './wpsResultReviewMarkers';
import {
  bindWpsReviewTarget,
  identityOf,
  isWpsReviewCopyWorkbook,
  releaseWpsReviewTarget,
  wpsReadWorkbook
} from './wpsResultReviewTarget';
import { rearrangeWpsComparison } from './wpsNativeComparison';
import { flushSync } from "react-dom";
import { createAgentTaskContext, confirmAgentTaskDecision, type AgentTaskContext } from "@excel-bro/core/agentTaskContext";
// 工作目录范围的发送判据：共享层唯一一份声明（Excel 端读同一份）。
import { workDirectorySendRefusal } from "@excel-bro/core/folderScope";
import {
  newTaskId,
  PLAN_NOTE_SECTIONS,
  redactCredentials,
  sanitizeTaskResumeRecord,
  SESSION_RECOVERY_LABEL,
  SESSION_RECOVERY_NOTICE_PREFIX,
  TASK_BLOCKER_LABEL,
  TASK_RESUME_SCHEMA_VERSION,
  TASK_WRITE_STATE_LABEL,
  taskResumeDecision,
  withTaskResumeState,
  type TaskBlocker,
  type TaskPhase,
  type TaskResumability,
  type TaskResumeConfirmation,
  type TaskResumeRecord,
  type TaskWriteState
} from "@excel-bro/core";
import {
  containsCredentialText,
  isLegacyCredentialConversation,
  redactHistoryCredentials,
  remainingLegacyCredentialConversationIds,
  replaceRestoredSecrets
} from "../../excel-addin/src/conversation";
import { brandLogoUrl } from "../../excel-addin/src/brandLogo";
import {
  ApiRequestError,
  apiErrorCategory,
  listModels,
  reportWpsHostWrites,
  understandMessage,
  validatePlan,
  type ModelCatalogStatus,
  type ModelOption
} from "../../excel-addin/src/api";
import type { HostWriteAudit } from "@excel-bro/core/hostWriteAudit";
import { assertEditablePlan, type AnalysisPlan, type ClarificationField, type IntentOption, type ResultContext } from "../../excel-addin/src/contracts";
import { ClarificationFields } from "../../excel-addin/src/components/ClarificationFields";
import { useTheme } from "../../excel-addin/src/hooks/useTheme";
import { useLongPress } from "../../excel-addin/src/hooks/useLongPress";
import { useModelManagement } from "../../excel-addin/src/hooks/useModelManagement";
import {
  chooseAvailableModel,
  isRealModelId
} from "../../excel-addin/src/modelSelection";
import ThemePanel from "../../excel-addin/src/components/ThemePanel";
import PetCompanion from "../../excel-addin/src/PetCompanion";
import { RuleManager } from "../../excel-addin/src/RuleManager";
import {
  analyzeToolEligibility,
  createTool,
  deleteTool,
  instantiateTool,
  instantiateQueryTool,
  loadQueryTools,
  loadTools,
  saveTool,
  updateToolPlanTemplate,
  type SavedQueryTool,
  type SavedTool,
  type ToolRiskApproval
} from "../../excel-addin/src/storage";
import { executeQueryTableData } from "../../excel-addin/src/dataTools";
import TaskProgressTimeline from "../../excel-addin/src/components/TaskProgressTimeline";
import { SaveToolDialog } from "../../excel-addin/src/components/SaveToolDialog";
import MessageCopyButton from "../../excel-addin/src/components/MessageCopyButton";
import { ExecutionResultCard, ExecutionResultDetails } from "../../excel-addin/src/components/ExecutionResultCard";
import ScrollToLatestButton from "../../excel-addin/src/components/ScrollToLatestButton";
import { ChangeReviewUndoNote } from "../../excel-addin/src/components/ChangeReviewPresentation";
import ChangeReviewSummaryCard from "../../excel-addin/src/components/ChangeReviewSummaryCard";
import { LocalMappingResolution } from "../../excel-addin/src/components/LocalMappingResolution";
import { wpsConfirmationTargetKey } from "./wpsClarificationTargetConfirmation";
import {
  changeReviewSummary,
  type ChangeReviewSummary,
  type RangeChangeReview
} from "@excel-bro/core/changeReview";
import {
  REAUTHORIZATION_NOTICE,
  clarificationAfterReauthorization,
  newPlanningRequestId,
  renderConfirmedPrompt
} from "@excel-bro/core";
import { resumeActivityLog, type ActivityLog, type ActivityOutcome } from "../../excel-addin/src/hooks/useActivityProgress";
import type { ReplySegment } from "@excel-bro/core/replySegments";
import {
  freezeSegmentLog,
  openReplySegment,
  planSegmentContinuation
} from "@excel-bro/core/replySegments";
import { useMessageAutoFollow } from "../../excel-addin/src/hooks/useMessageAutoFollow";
import { useComposerResize } from "../../excel-addin/src/hooks/useComposerResize";
import { useAuthorizedTarget } from "../../excel-addin/src/hooks/useAuthorizedTarget";
import { useExecutionApproval } from "../../excel-addin/src/hooks/useExecutionApproval";
import { AuthorizedTargetConfirmation } from "../../excel-addin/src/components/AuthorizedTargetConfirmation";
import { TargetScopeNotice } from "../../excel-addin/src/components/TargetScopeNotice";
import { inspectWpsAuthorizedTarget, isWpsReviewCopyTarget, wpsAnalysisTargetEvidence } from "./wpsAuthorizedTarget";
import { checkAuthorizedTarget, TargetConfirmationRequiredError, type AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { buildPlanningContinuationEvidence, PLANNING_EVIDENCE_FAILURE_MESSAGES } from "@excel-bro/core/planningContinuationEvidence";
import { copyTextToClipboard } from "../../excel-addin/src/utils";
import WpsModelSettingsPanel from "./WpsModelSettingsPanel";
import { ModelPickerIcon } from "../../excel-addin/src/components/ModelPickerIcon";
import {
  PET_VISIBILITY_STORAGE_KEY,
  normalizePetVisibility
} from "../../excel-addin/src/conversation";
import { getWpsApplication } from "./wpsHost";
import { useWpsReviewPaneRatio } from './useWpsReviewPaneRatio';
import { REVIEW_PANE_LIMITS } from '@excel-bro/core/reviewPaneSizing';
// 审阅副本生命周期：与 Excel 端同一套状态/文案，只把宿主适配留在本端。
import { reviewCopyCleanupState, type ReviewCopyBusinessOutcome, type ReviewCopyCleanupState } from "@excel-bro/core/reviewCopyLifecycle";
// 审阅卡的信息层级与文案口径：与 Excel 端**共用**，只把宿主差异当入参喂进去。
import {
  REVIEW_CARD_COMPAT_DETAIL,
  REVIEW_CARD_COMPAT_DETAIL_TITLE,
  REVIEW_CARD_DIAGNOSTICS_TITLE,
  planTargetCells,
  reviewCardBusiness,
  reviewCardCompatNotice,
  reviewCardHasDiagnostics,
  reviewCardView,
  reviewChangeSummary,
  reviewScopeLocations,
  type ReviewCardDiagnostics,
  type ReviewCardScope,
  type ReviewCardView,
  type ReviewOutcome
} from "@excel-bro/core/reviewCardPresentation";
// 写回请求的现场身份（同一次审阅 / 同一计划版本 / 同一原工作簿）：两端共用同一套判据。
import { reviewWriteBackTargetChanged, type ReviewWriteBackTarget } from "@excel-bro/core/reviewWriteBackGate";
// 写回身份里仍带着"这一次是哪条执行路径"（与 Excel 端共用同一套判据）。
import {
  executionModeTargetChanged,
  type ExecutionModeWriteBackTarget as DirectExecutionTarget,
  type PlanExecutionMode
} from "@excel-bro/core/executionMode";
import { probeWideViewDialog, type WideViewDialogProbe } from "./wpsWideViewPoc";
import {
  createProbeNonce,
  startParentChannelProbe,
  type ChannelProbeStatus
} from "./wpsWideViewChannelPoc";
import {
  authorizeWpsSheets,
  listWpsWorksheets,
  readAuthorizedWpsSheetSchemas,
  readAuthorizedWpsSheets
} from "./wpsWorkbookReader";
import { probeWpsWriteCapabilities } from "./wpsWriteCapabilities";
import { columnMappingClarificationMessage } from "../../excel-addin/src/columnMappingClarification";
import {
  wpsHeaderFingerprint,
  wpsScopeChanged,
  wpsSheetIdentity
} from "./wpsClarificationScope";
import {
  wpsClarificationScopeNotice,
  wpsClarificationVerdict,
  type WpsClarificationVerdict
} from "./wpsClarificationVerdict";
import {
  captureWpsObjectIdentity,
  verifyWpsObjectIdentity,
  type WpsIdentityCheck,
  type WpsObjectIdentity
} from "./wpsObjectIdentity";
import {
  executeWpsPlan,
  inspectWpsUndoRecord,
  latestWpsUndo,
  prepareWpsPlan,
  undoWpsExecution,
  wpsUndoRecordFor,
  type WpsPreparedPlan,
  type WpsUndoRecord
} from "./wpsPlanExecutor";
import { buildPlanChangePreview, planActionLabel } from "./wpsPlanChangePreview";
import {
  classifyWpsWriteEvidence,
  normalizeWpsExecutionOutcome,
  taskResumePatchForEvidence,
  WPS_EXECUTION_OUTCOME_DETAIL_SUFFIX,
  WPS_EXECUTION_OUTCOME_LABEL,
  WPS_RECOVERY_CARD_LABEL,
  wpsExecutionCardTone,
  wpsRecoveryCardState,
  type WpsExecutionOutcome,
  type WpsWriteEvidence
} from "./wpsExecutionRecovery";
import {
  compareCommittedWpsResultToCopy,
  discardWpsNativeComparison,
  stageWpsNativeComparison,
  stageWpsBeforeComparison,
  captureWpsBeforeResult,
  type WpsBeforeResult,
  verifyWpsNativeComparison,
  WpsFormulaTrialError,
  type WpsNativeComparison
} from "./wpsNativeComparison";
import { fixedValueFallbackFromFormula } from "./wpsColumnMappingPlan";
import { describeWpsFormatAction } from "./wpsFormatCapability";
import {
  advanceWpsActivity,
  executionActivityUpdate,
  finishWpsActivity,
  readonlyActivityUpdate,
  startWpsActivity,
  updateWpsActivityMessage,
  type WpsActivityUpdate
} from "./wpsActivityLog";
import {
  resumeWpsLocalMapping,
  runWpsReadonlyTurn,
  wpsProgressFromTaskEvent,
  type WpsReadonlyTurnProgress,
  type WpsTurnClarification,
  type WpsTurnResult
} from "./wpsReadonlyAgent";
import {
  claimWpsExecutionResume,
  clearPendingWpsExecution,
  completeWpsAction,
  finishWpsExecution,
  getWpsExecutionStatus,
  loadPendingWpsExecution,
  reconcileWpsActionReceipt,
  registerWpsWriteIntent,
  savePendingWpsExecution,
  startWpsAction,
  suspendWpsExecution,
  wpsWorkbookIdentity,
  type PendingWpsExecution
} from "./wpsExecutionLedger";
import {
  wpsIdentityVerdict,
  wpsReconcileVerdict,
  wpsReconcileActions,
  WPS_RECONCILE_LABEL,
  WPS_RECOVERY_ACTION_LABEL,
  WPS_RECOVERY_DETAIL,
  type WpsIdentityVerdict,
  type WpsReconcileAction,
  type WpsReconcileVerdict
} from "./wpsExecutionReconcile";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import { inspectWpsRecoveryState } from "./wpsRecoveryInspection";
import { useImageAttachments } from "../../excel-addin/src/hooks/useImageAttachments";
import type { PendingImage } from "../../excel-addin/src/imageAttachments";
import {
  inspectWpsPlanningTask,
  planningPauseFromStatus,
  reinspectWpsPlanningPause,
  resumeWpsPlanning,
  wpsPlanningResumeView,
  type WpsPlanningResumeView
} from "./wpsReadonlyAgent";
import type { WpsPlanningTaskEvents, WpsPlanningTaskStatus } from "./wpsPlanningStream";
import { getWpsPlanningTaskEvents } from "./wpsPlanningStream";
import type { PlanningPauseResult } from "../../excel-addin/src/contracts";
import {
  attestResultSourceDependency,
  attestSourceDependency,
  planningPausedEntriesFromHistory,
  sourceAnchorFromBinding,
  stampPlanSourceAnchor,
  trustedToolReads
} from "@excel-bro/core";
import { LocalResultCard as LocalResultCardView } from "../../excel-addin/src/components/LocalResultCard";
import { useSdkToolBridge } from "../../excel-addin/src/hooks/useSdkToolBridge";
import type { BridgeResultLookup } from "../../excel-addin/src/sdkBridgeHost";
import type { LocalResultCard } from "../../excel-addin/src/contracts";
import {
  createWpsBridgeLiveContext,
  createWpsSdkBridgeHostPorts,
  loadWpsBridgeResult,
  type WpsSdkBridgeLiveContext
} from "./wpsSdkBridgeHost";
import { wpsWorkbookBridgeIdentity } from "./wpsWorkbookReader";

type ChatMessage = {
  id: string;
  role: "user" | "assistant" | "system";
  text: string;
  attachmentNames?: string[];
  hasImageAttachments?: boolean;
  resultContext?: ResultContext | null;
  /**
   * SDK 本地查询结果卡：**只有引用与结构，没有任何值**。
   * 真实值由本机回环端点按 `resultRef` 取回后渲染，不进模型/历史/持久账本。
   */
  localResult?: LocalResultCard;
  /**
   * 产生该结果那一轮的桥绑定。会话关闭后服务端按**授权归属**校验，
   * 所以要回带这一轮的绑定，不能用"当前"绑定顶替。
   */
  localResultBinding?: BridgeResultLookup;
  activityLog?: ActivityLog;
  /**
   * 这条助手回复属于哪一段展示（与执行身份分开）。
   *
   * 一个任务可以有多段回复（暂停 → 继续 → 再暂停 → 再继续），同一时刻只有一段
   * 是 `live`。段身份不替代 `operationId` / `requestId`：步骤身份、事件序号、
   * 结果引用与写入回执仍然只认执行身份。
   */
  replySegment?: ReplySegment;
  planPreview?: {
    title: string;
    status: "pending" | "executed" | "cancelled";
    /**
     * 这份预览是"编辑计划后重新预览"的结果。跟着消息走，所以新计划、取消、
     * 执行后都会自动回到未修订态，不需要额外的组件状态。
     */
    revised?: boolean;
  };
  /**
   * 这次**写入执行**的结果卡。它同时是恢复点的宿主：撤销按钮就渲染在这张卡里，
   * 靠 `executionId` 与执行绑定 —— 不按消息位置、不按计划标题、不按"最近一次回复"。
   *
   * 只要执行真的留下了恢复点就会画卡，不止"回读完全通过"那一种：
   * `outcome` 如实区分完全回读通过 / 计算值未获证实 / 部分写入。
   */
  verification?: {
    executionId: string;
    actionCount: number;
    cellCount: number;
    /**
     * 卡片生命周期：`recorded` = 这次写入执行留下了卡；`undone` = 已撤销（终端态）。
     *
     * 它**不**表达"写得好不好" —— 那由 `outcome` 说，别拿它冒充"验证通过"。
     * 旧历史里这个字段曾写 `"verified"`，读取时应视为 `recorded`
     * （当时只有完全回读通过才会画卡）。
     */
    status: "recorded" | "undone";

    /**
     * 这次执行到底发生了什么。旧历史没有这个字段，读取时按"完全回读通过"
     * 归一化（那时只有这种才会画卡）。
     */
    outcome?: WpsExecutionOutcome;
    targets: { type: string; sheetName: string; range: string }[];
  };
  /**
   * 写入失败时附上的双栏明细索引。只存 id —— 矩阵留在 `changeReviewsRef` 里，
   * 不随这条消息持久化（方案 §7）。
   */
  changeReviewExecutionId?: string;
  /**
   * 失败/冲突的**可展开诊断详情**（撤销未完成、恢复冲突等）。
   *
   * `text` 只说真实状态的结论，具体是哪个区域、哪一道校验、执行身份是谁放在这里，
   * 默认折叠、用户点开才看。这样首屏不被诊断文字淹没，但也不存在"静默隐藏"——
   * 结论与详情都在同一条消息里。
   *
   * 只放**定位与状态**，不放单元格内容；写入前过凭据脱敏。
   */
  errorDetail?: string;
  /**
   * 这张卡所属任务的**可持久化恢复记录**（见 `@excel-bro/core` 的 `taskResume`）。
   *
   * 它只携带"定位 + 阶段 + 原需求 + 用户已确认的信息"，没有任何单元格内容、
   * 宿主对象引用或完整路径 —— 白名单消毒是结构性的，不靠调用方自觉。
   * 有了它，任务窗格重载后这张卡才**仍然是一件可继续的任务**，
   * 而不是一句"已失效，请重新发送需求"。
   */
  taskResume?: TaskResumeRecord;
  /**
   * 规划暂停卡（CAP-02）。
   *
   * 只带**定位与归因**：`turnId`（服务端账本主键）、暂停类型、原因，以及由
   * 服务端结构化状态推导出的按钮动作。不存暂停载荷全文（`PlanningPauseResult`
   * 里含 `availableModelConnections` 等连接元数据，没必要进历史）。
   *
   * 关键纪律：卡片上"能不能继续"**不来自这里**，而是每次点击时
   * **重新查询服务端真实状态**再判断 —— 落盘的这份只是当时那一刻快照。
   */
  planningPause?: {
    turnId: string;
    pauseType: "budget" | "budget_exhausted" | "model_error" | "interrupted";
    reason: string;
    canContinue: boolean;
    action: "continue_checkpoint" | "regenerate" | "none";
    actionLabel: string;
    detail: string;
    awaitingStatus?: boolean;
    clientLastSequence?: number;
    resolved?: boolean;
    terminalReason?: string;
  };
  clarification?: {
    summary: string;
    question: string;
    reason: string;
    scopeLabel: string;
    options: { label: string; description: string }[];
    fields?: ClarificationField[];
    /**
     * `"unconfirmed"` 与 `"invalidated"` **刻意分开**：
     * 前者是"证明不了还是原授权对象"（不许谎称数据已变，原需求与回答要保留），
     * 后者是"有证据表明授权对象或表头已变"。
     *
     * `"continued"` 是**重新授权后的终态**：用户重新确认了数据范围、任务
     * 已作为新一轮继续 —— 它**不等于**回答了那道题，所以与 `"resolved"`
     * 分开（后者代表用户真的给了答案）。
     */
    status:
      | "pending"
      | "resolved"
      | "invalidated"
      | "unconfirmed"
      | "continued"
      | "cancelled";
    answer?: string;
    /**
     * 建卡时的**授权身份指纹**（工作表索引 + 名称 + 顺序 + 工作簿名）。
     *
     * 只存索引是不够的：重排、删除或替换之后同一个索引会指向**另一张表**，
     * 用户对着旧表回答的问题就会被拿去算新表。宿主没有跨会话稳定的工作表 ID，
     * 所以这里保守地把名称与顺序也算进身份 —— 任何一项变化都让卡片失效。
     */
    sheetIdentityFingerprint?: string;
    /**
     * 建卡时结构扫描读到的**表头**指纹。
     *
     * 表头改名同样会改变"这一列是什么"的语义，所以它也参与失效判定。
     * 只含表头行：数据行不进指纹，否则用户改数据就会把卡片打掉。
     */
    headerFingerprint?: string;
  };
  createdAt: string;
};

type WpsConversation = {
  id: string;
  title: string;
  messages: ChatMessage[];
  createdAt: string;
  updatedAt: string;
};

type WpsChatHistory = {
  activeConversationId: string;
  conversations: WpsConversation[];
  /**
   * 旧记录里曾带真凭据、已被读取路径脱敏的会话 id。独立且可持久化，
   * 用于在脱敏之后继续拦截这些会话的续聊（判据不能是"还能不能匹配密钥"）。
   */
  legacyCredentialConversationIds?: string[];
};

type PendingClarification = {
  messageId: string;
  originalPrompt: string;
  round: number;
  options: IntentOption[];
  /**
   * 这道题的**稳定身份**（不是消息 id —— 重新生成预览会换消息）。
   *
   * 用户改主意时，新的确认靠它**替换**旧的；没有它只能追加，而追加会把
   * 改过的旧答案留在提示词里，形成两条互相冲突的"补充确认"。
   * 模型路径的澄清没有稳定语义身份，留空（保守追加）。
   */
  questionId?: string;
  /** 建卡时的授权身份指纹（见 `ChatMessage.clarification`），作答前复核。 */
  sheetIdentityFingerprint?: string;
  /** 建卡时结构扫描读到的表头指纹，作答前复核。 */
  headerFingerprint?: string;
  /** 只活在当前面板内存；源值和用户补充映射不得进入 ChatMessage/localStorage。 */
  localMapping?: NonNullable<WpsTurnClarification["localMapping"]>;
  /** 仅用于新一轮本地缺口时重置表单，不把来源值放进 React key。 */
  localMappingRevision?: number;
  selectedSheetIndexes?: readonly number[];
};

type WpsAgentPanelProps = {
  selectedSheetIndexes: readonly number[];
  resolveSelectedSheetIndexes?: () => readonly number[];
  scopeMode?: "auto" | "manual";
  /**
   * 这次对话的数据来源。
   *
   * `"folder"`（工作目录）时，数据由**本机服务端**读取，跟当前打开的 WPS 工作簿
   * 毫无关系。下游因此**不得**把 `selectedSheetIndexes` 当成这次的范围 ——
   * 那等于用一个用户没授权的工作簿顶替他要的目录。
   */
  scopeSource?: "workbook" | "folder";
  /** 工作目录范围摘要；`null` = 这次不是工作目录范围（或还没确认）。 */
  folderScope?: { folderName: string; sheetCount: number } | null;
  onAuthorizeTarget?: (indexes: number[]) => void;
  onScopeConfirmationReady?: (handler: ((target: AuthorizedTarget, options?: { isCurrent?: () => boolean; allowEmpty?: boolean }) => Promise<boolean>) | null) => void;
  scopeSelectionPending?: boolean;
  onOpenScope?: () => void;
  scopePickerOpen?: boolean;
  scopeDock: ReactNode;
  /**
   * 数据范围**面板本体**的独立呈现层。
   *
   * 为什么和 `scopeDock`（输入框上方那一行触发器）分开：面板原来是
   * `.composer` 的后代 + `position:absolute; bottom: calc(100% + 7px)`，于是它被
   * 输入框那一层牵着走 —— 审阅条一出现、窗口一矮，面板就被挤扁/截掉，还和输入区叠。
   * 放到外壳的直接子层里，它才不再受输入框的定位、`backdrop-filter` 与裁剪影响。
   */
  scopeOverlay?: ReactNode;
};

const WPS_CHAT_STORAGE_KEY = "excel-bro.wps.chat.v1";

/**
 * 接回"仍在运行的原任务"时的轮询节奏。
 *
 * 上限而不是死等：服务端跑一次规划有硬时限，超过这个窗口还没结束就不是"等一等
 * 就能好"的事，如实告诉用户仍在运行、让他稍后核对，比无限转圈更诚实。
 */
const RUNNING_TAIL_INTERVAL_MS = 1200;
const RUNNING_TAIL_MAX_MS = 90_000;

/**
 * 暂停请求的确认上限。
 *
 * 本端 abort 是同步的，但"服务端收到没有"只能靠 HTTP 回执。没有上限时，一次
 * 卡住的请求会让"正在请求暂停"永远挂着 —— 用户以为按住了，其实什么都没发生。
 * 超时不是失败结论：如实说"未确认"。
 */
const PAUSE_REQUEST_TIMEOUT_MS = 8_000;

const LEGACY_MODEL_CATALOG_NOTICE =
  "暂时无法读取模型配置，请点击上方模型菜单里的「重试」。已输入的内容仍保留在输入框中。";
/** AI 窗格占**整个可用工作区**的目标百分比。与 Excel 共用同一份 `reviewAiWidthFraction`。 */

/**
 * 规划暂停的成因文案（CAP-02）。
 *
 * 只按服务端给的 `pauseType` 取值 —— 这是结构化判据，不是从助手文本里读出来的。
 * 各成因的**处置方式不同**，所以文案必须分开：
 *  - `budget`：本轮计算量达到上限，可继续或换模型；
 *  - `budget_exhausted`：累计活跃执行时间已达总时限，**不再承诺"必然可以继续"**；
 *  - `model_error`：模型连接故障，通常需要换一个可用连接；
 *  - `interrupted`：被打断（如刷新/断线），现场仍在。
 */
const WPS_PLANNING_PAUSE_LABEL: Record<
  "budget" | "budget_exhausted" | "model_error" | "interrupted",
  string
> = {
  budget: "本轮处理时间已达上限，进度已保存",
  budget_exhausted: "本轮有效运行总时长已用完，现场已保存",
  model_error: "模型连接中断，进度已保存",
  interrupted: "规划被中断，进度已保存"
};

function messageId(): string {
  return `wps-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

/**
 * 当前**授权身份指纹**：对已勾选的工作表，逐个取出宿主上的真实名称。
 *
 * 为什么不能只用索引：索引只是"第几张表"。重排、插入、删除或替换之后，
 * 同一个索引会指向另一张表 —— 用户对着"订单"回答的澄清会被拿去算"客户"，
 * 授权边界被绕过而界面上看不出来。宿主没有跨会话稳定的工作表 ID，
 * 所以保守地把名称与顺序一起算进身份：**任何一项变化都让待答澄清失效**。
 *
 * 代价是"表改名"会让上一张卡失效 —— 这是必须付的代价，
 * 不得为了"改名不断对话"牺牲授权边界。
 */
function currentWpsSheetIdentity(): string | null {
  try {
    const metadata = listWpsWorksheets(getWpsApplication());
    return wpsSheetIdentity({
      workbookName: metadata.name,
      sheets: metadata.sheets.map((sheet) => ({ index: sheet.index, name: sheet.name }))
    });
  } catch {
    // 读不到宿主身份时返回 null：调用方据此**不判为变化**（避免误伤），
    // 但这不代表授权被放宽 —— 真正的读取仍由 `authorizeWpsSheets` 卡住。
    return null;
  }
}

/**
 * 当前**表头指纹**：读宿主上每张已授权表的**第 1 行真实单元格值**。
 *
 * 与 `currentWpsSheetIdentity` 的区别是判据对象不同 —— 那个看"是不是同一张表"，
 * 这个看"这份结构的列名有没有被改写"。两者必须读**不同的数据源**，否则后者
 * 只是前者的副本：禁用前者时后者会顶上来把用例撑绿，让人误以为身份判据没必要。
 *
 * 只读第 1 行、且只经由 `readAuthorizedWpsSheetSchemas`（它已把取值面收窄到表头），
 * 所以这里**不会**把数据行读进内存。读不到宿主时返回 `null`，不参与比对（不误伤）。
 */
async function currentWpsHeaderFingerprint(sheetIndexes: readonly number[]): Promise<string | null> {
  try {
    const application = getWpsApplication();
    const authorization = authorizeWpsSheets(application, sheetIndexes);
    const schema = await readAuthorizedWpsSheetSchemas(application, authorization);
    return wpsHeaderFingerprint(
      schema.map((sheet) =>
        (sheet.values[0] ?? []).map((cell) => (cell == null ? "" : String(cell)))
      )
    );
  } catch {
    return null;
  }
}

function turnErrorMessage(error: unknown): string {
  if (apiErrorCategory(error) === "local_service") {
    return "无法连接 Excel Bro 后端。请确认本地服务已启动，然后重试；本次未修改工作表。";
  }
  // 非连接类错误的 message 会被写进对话历史并随之持久化。服务端错误文本源于
  // Python 侧 `str(error)`，而请求里带着 apiKey —— 原样透出就等于把凭据存进
  // 宿主 localStorage。先过共享脱敏，再截断，避免长文把历史撑爆。
  const raw = error instanceof Error ? error.message : "任务处理失败；本次未修改工作表。";
  const redacted = redactCredentials(raw).trim();
  if (!redacted) return "任务处理失败；本次未修改工作表。";
  return redacted.length > 500 ? `${redacted.slice(0, 500)}…` : redacted;
}

/**
 * 把「业务结果 × 关闭结果」映射成共享文案（与 Excel 端逐字同源）。
 *
 * 分开陈述的理由：写回成功与副本关没关是**两件事**。
 * "副本未能自动关闭"绝不能读成"执行失败"，"原表已写回"也绝不能冒充"副本已清理"。
 */
export function wpsCopyCleanupState(business: ReviewCopyBusinessOutcome, copy: Parameters<typeof reviewCopyCleanupState>[0]["copy"]): ReviewCopyCleanupState {
  return reviewCopyCleanupState({ business, copy, tempFilePath: null });
}
export function wpsCopyCleanupText(state: ReviewCopyCleanupState): string {
  return [state.copy.headline, state.copy.detail, state.file.headline].filter(Boolean).join(" ");
}
/** 副本是否已确认收尾（可以安全释放审阅会话）。其余结果都保留现场让用户决定。 */
export function wpsCopyRetired(state: ReviewCopyCleanupState): boolean {
  return state.copy.outcome === "closed" || state.copy.outcome === "already_closed";
}

/**
 * 副本状态**独立**陈述：不带业务前缀。
 *
 * 写回成功时业务结论已经由执行结果卡说了，这里再念一遍"已写回并验证"就是把
 * 同一件事说两遍，还会让人以为副本状态是业务结论的一部分。
 */
export function wpsCopyStatusLine(state: ReviewCopyCleanupState): string {
  return [state.copy.statusText, state.copy.detail].filter(Boolean).join(" ");
}

/**
 * 终态的对话陈述：业务结论在前，副本处理指引与副本状态在**后**（次要）。
 *
 * 顺序就是主次：先回答"我的计划怎么了"，再说"副本还要你做什么"。
 */
export function wpsReviewOutcomeText(view: ReviewCardView): string {
  return [
    view.business.headline,
    view.business.detail,
    view.actualResult,
    view.copyOpenNotice,
    view.copyStatus ? `${view.copyStatus.headline} ${view.copyStatus.detail}` : null
  ].filter(Boolean).join(" ");
}

/**
 * WPS 端审阅卡：主卡口径与 Excel 端**共用**，只把宿主差异当入参。
 *
 * 真实宿主差异（不是文案推断）：WPS 的 `Window` 没有可写宽度成员，
 * 所以 AI 窗格宽度**只能手动拖** —— `paneWidthSettable: false`。
 */
export function wpsReviewCard(input: {
  outcome: ReviewOutcome;
  scope: ReviewCardScope | null;
  compatNotice: string | null;
  paneNeedsResize: boolean;
  copy: ReviewCopyCleanupState["copy"] | null;
  executionDetail?: string | null;
  syncScrollingEnabled?: boolean;
  diagnostics?: Partial<ReviewCardDiagnostics>;
}): ReviewCardView {
  return reviewCardView({
    ...input,
    paneWidthSettable: false,
    compatDetail: REVIEW_CARD_COMPAT_DETAIL["wps-native"]
  });
}

/**
 * 终态审阅卡：放弃 / 写回成功 / 写回失败。
 *
 * 它只用来生成**对话里那段陈述**（WPS 的审阅卡在终态会被撤下），所以
 * 不再讲兼容模式怎么核对，也不带任何审阅入口。
 */
export function wpsTerminalCard(input: {
  outcome: Exclude<ReviewOutcome, "reviewing">;
  scope: ReviewCardScope | null;
  copy: ReviewCopyCleanupState["copy"] | null;
  staged: WpsNativeComparison;
  executionDetail?: string | null;
  developmentWrite: boolean;
}): ReviewCardView {
  return wpsReviewCard({
    outcome: input.outcome,
    scope: input.scope,
    compatNotice: null,
    paneNeedsResize: false,
    copy: input.copy,
    executionDetail: input.executionDetail ?? null,
    diagnostics: wpsReviewDiagnostics(input.staged, false, input.developmentWrite)
  });
}

/**
 * 诊断层信息：开发环境提示、排列指引、副本标识、关闭诊断、磁盘清理。
 *
 * **不进首屏** —— 它们是排查用的，不是"要不要按确认"需要的。
 */
/**
 * 写回入口的**唯一闸门**判据。
 *
 * 返回值是"为什么现在不能写回"（`null` = 可以写）。抽成纯函数是为了让它**可被直测**：
 * 「放弃之后拿着旧回调来点」与「没勾选就来点」都必须在**同一处**被拦下 ——
 * 只要判据只有一份，就不存在"界面拦了、回调没拦"的缝隙。
 *
 * 前两项是**终态**：计划或副本已经不在场（已放弃 / 已写回），写回入口整体失效。
 * 它们必须排在授权检查之前：已经决定不写的时候，连"要不要确认授权范围"都不该发起。
 */
export type WpsWriteBackBlock = "no_review" | "busy" | "editing_plan" | "not_acknowledged";

export function wpsWriteBackBlockReason(state: {
  acknowledged: boolean;
  hasPlan: boolean;
  hasReview: boolean;
  editingPlan: boolean;
  busy: boolean;
}): WpsWriteBackBlock | null {
  if (!state.hasPlan || !state.hasReview) return "no_review";
  if (state.busy) return "busy";
  if (state.editingPlan) return "editing_plan";
  if (!state.acknowledged) return "not_acknowledged";
  return null;
}

/**
 * **直接执行**路径的写回闸门判据（与 `wpsWriteBackBlockReason` 分工明确）。
 *
 * 与副本路径的差别只有一条：**不要求审阅副本在场**。其余每一条都相同 ——
 * 没计划、忙、正在编辑计划，一样挡。
 *
 * 为什么不直接把 `hasReview` 改成可选：副本路径"必须先有副本"是那条路径的
 * **全部意义**（用户核对的是副本）。放松它会让"没核对就写回"重新变成可达状态。
 * 两条判据分开，直接执行才敢声明自己只跳过副本比对。
 */
export type WpsDirectWriteBackBlock = "no_plan" | "busy" | "editing_plan";

export function wpsDirectWriteBackBlockReason(state: {
  hasPlan: boolean;
  editingPlan: boolean;
  busy: boolean;
}): WpsDirectWriteBackBlock | null {
  if (!state.hasPlan) return "no_plan";
  if (state.busy) return "busy";
  if (state.editingPlan) return "editing_plan";
  return null;
}

/**
 * **异步等待回来之后**的复核：这一次写回请求还能不能继续？
 *
 * 与 `wpsWriteBackBlockReason` 分工不同 —— 那是**发起点**的闸门（能不能发起），
 * 这里是**回来后**的复核（还能不能继续）。两者不可互相顶替：
 *
 *  - 发起时"有审阅"是真的，等回来时那份审阅可能已经被放弃了；
 *  - 更险的是"有审阅"仍然为真，但已经**换成了另一份** —— 用户放弃之后又生成了
 *    新副本并勾了新的确认框。此时"计划与副本都在场 + 已勾选"全部成立，
 *    旧请求若据此继续写回，就是把**新副本的确认状态借过来**写进原表：
 *    用户核对过的那一份从来没被写，没核对过的那一份反而被写了。
 *
 * 所以除了"现场还在"，还必须逐项比对身份三元组：同一次审阅、同一计划版本、
 * 同一原工作簿。任一侧拿不到身份（`null`）一律按"已变化"处理（fail-closed）。
 *
 * 导出是为了让这条判据能被**直接**判定 —— 面板里"旧请求在途期间换副本"这条复合
 * 路径目前被 `busy` 闸门整个挡在界面之外（`busy` 含 `targetGuard.checking`：写回
 * 等待期间放弃 / 生成副本 / 发新任务的入口全是禁用的），端到端造不出来。所以它是
 * **纵深防御**：等哪天 `busy` 不再覆盖这个窗口（或出现已排队的点击 / 快捷键 /
 * 程序直调），判据已经在位，并被单独判定。
 *
 * 「现场还在」那一半由调用点判（顺带把 `plan` / `review` 的 null 收窄），这里只判
 * **还是不是同一份现场** —— 两半都在调用点各占一行，缺一不可。
 */
export function wpsWriteBackStillCurrent(
  issued: ReviewWriteBackTarget | null,
  now: { target: ReviewWriteBackTarget | null }
): boolean {
  return !reviewWriteBackTargetChanged(issued, now.target);
}

/**
 * 直接执行请求回来之后的复核：**计划版本 + 原工作簿 + 执行方式**三项全同才继续。
 *
 * 与副本路径的差别是"同一次审阅"换成了"同一种执行方式"：直接执行没有副本，
 * 但它仍必须确认"我发起时看到的计划与授权，就是现在要写的那份"。
 * 方式也必须进身份：用户在途期间切到副本比对再回来，旧请求不该按旧前提继续。
 */
export function wpsDirectWriteBackStillCurrent(
  issued: DirectExecutionTarget | null,
  now: { target: DirectExecutionTarget | null }
): boolean {
  return !executionModeTargetChanged(issued, now.target);
}

export function wpsReviewDiagnostics(
  staged: WpsNativeComparison,
  paneNeedsResize: boolean,
  developmentWrite: boolean
): Partial<ReviewCardDiagnostics> {
  const restart: string[] = [];
  if (developmentWrite) restart.push("当前 WPS JSAPI 版本仅在开发调试模式开放写入；请先在独立测试工作簿验收写入与撤销。");
  const arrangement = [
    "WPS 不能由代码设置 AI 窗格宽度。若两张表的可见宽度差得多，请拖动「Excel Bro」左侧分隔线手动调整。",
    "未确认窗口重新排列时，请在 WPS「视图 → 并排比较」中手动选择本次副本。"
  ];
  if (paneNeedsResize) arrangement.push("AI 窗格宽度尚未确认落到目标比例。");
  return {
    restart,
    arrangement,
    identity: [
      `原表：${redactCredentials(staged.original.Name)}`,
      `副本：${redactCredentials(staged.copy.Name)}`,
      `副本差异格数：${staged.changedCells.length}`
    ],
    close: [],
    cleanup: ["本次没有需要清理的磁盘文件：审阅副本是宿主内存中的未保存工作簿。"]
  };
}

function conversationTitle(messages: readonly ChatMessage[]): string {
  const firstQuestion = messages.find((message) => message.role === "user")?.text.trim();
  if (!firstQuestion) return "新对话";
  const compact = firstQuestion.replace(/\s+/g, " ");
  return compact.length > 34 ? `${compact.slice(0, 34)}…` : compact;
}

function createWpsConversation(messages: ChatMessage[] = []): WpsConversation {
  const now = new Date().toISOString();
  return {
    id: messageId(),
    title: conversationTitle(messages),
    messages,
    createdAt: now,
    updatedAt: now
  };
}

function restoreActivityLog(log: ActivityLog | undefined): ActivityLog | undefined {
  if (!log || !Array.isArray(log.steps)) return undefined;
  if (log.outcome) return log;
  return {
    ...log,
    outcome: "interrupted",
    connectionState: "interrupted",
    connectionMessage: "任务窗格曾被关闭或刷新；此过程未取得终态，请核对工作表和恢复点。",
    steps: log.steps.map((step) => step.status === "started" || step.status === "progress"
      ? { ...step, status: "warning" as const }
      : step)
  };
}

/**
 * 已提交范围指纹。与面板里的 `authorizationKey` 同一算法 —— 只对**已提交**
 * 的范围取值，勾选草稿的变化不进这里。
 */
function authorizationKeyOf(indexes: readonly number[]): string {
  return [...indexes].sort((left, right) => left - right).join(",");
}

/**
 * 从恢复记录重建这次继续要发的请求文本。
 *
 * 原始需求在前，用户已确认的信息逐条附在后面 —— 顺序固定、内容原样，
 * 因此"重新规划"不会把用户确认过的参数丢掉，也不会把它们当成新的语义。
 */
function resumePrompt(record: TaskResumeRecord): string {
  // 同一道题只出现**最后一次**的答案（合并由共享 agentTaskContext 完成）。
  // 早期实现直接把历史确认全部追加，于是改过答案的用户会送进去两条
  // 互相冲突的"补充确认"，模型只能替他挑一个。
  return renderConfirmedPrompt(record.originalPrompt, record.confirmations);
}

/**
 * 构造一条恢复记录。走一遍共享白名单消毒：只有合法字段留得下来，
 * 顺带保证本端不会写出连自己都读不回来的记录（消毒失败即不写）。
 */
function buildTaskResumeRecord(input: {
  taskId: string;
  phase: TaskPhase;
  blocker: TaskBlocker;
  resumability: TaskResumability;
  writeState: TaskWriteState;
  originalPrompt: string;
  confirmations: TaskResumeConfirmation[];
  authorizationRevision: number;
  /**
   * 真实执行/规划身份。**重试与重连必须复用**，不得新建 ——
   * 这是"同一件事"的唯一凭据，也是服务端幂等与并发领取的键。
   */
  turnId?: string;
  operationId?: string;
  requestId?: string;
}): TaskResumeRecord | undefined {
  const now = new Date().toISOString();
  return (
    sanitizeTaskResumeRecord({
      schemaVersion: TASK_RESUME_SCHEMA_VERSION,
      createdAt: now,
      updatedAt: now,
      ...input
    }) ?? undefined
  );
}

/**
 * 预览已生成、**等待用户确认**时的恢复记录。
 *
 * 写前状态仍是 `not_written`（一个单元格都还没碰），可续跑性退回 `replan`：
 * 预览本身带着完整的目标矩阵，**不落盘**，所以重开后只能"重新读取、重新生成预览"，
 * 不能说成"从上次进度精确续跑"。
 */
function previewTaskResume(
  taskId: string,
  originalPrompt: string,
  authorizationRevision: number,
  turnId?: string | null,
  /**
   * 从历史任务继续时把**原记录里已确认的信息**带过来。丢掉它，用户就要重答
   * 同一个问题 —— 而那正是"不必重复描述需求和回答问题"要避免的。
   */
  confirmations: TaskResumeConfirmation[] = []
): TaskResumeRecord | undefined {
  return buildTaskResumeRecord({
    taskId,
    phase: "preview_ready",
    blocker: "none",
    resumability: "replan",
    writeState: "not_written",
    originalPrompt,
    confirmations,
    authorizationRevision,
    // 服务端这次规划的身份：断线后先用它查真实状态，有检查点才续跑。
    turnId: turnId ?? undefined
  });
}

/**
 * 写入失败提示的后半句：只说**有证据**支持的那部分。
 *
 * `not_started` 不加任何后缀（原表确实没动，不需要恐吓用户）；
 * `needs_review` 说"可能已部分写入"；`unknown` 直说"无法确认"并明确
 * **不要直接重跑** —— 不拿"未写入"糊过去。
 */
function evidenceSuffix(evidence: WpsWriteEvidence): string {
  switch (evidence.kind) {
    case "needs_review":
      return " 本次可能已部分写入，恢复资料已保留；请检查执行过程与撤销入口。";
    case "unknown":
      return " 这次写入的结果无法确认（没有可用的执行记录，或读取执行记录失败）；请先核对原表，不要直接重跑。";
    default:
      return "";
  }
}

/**
 * 面板重开后把持久化的恢复记录收敛成"需要用户重新确认范围"的形态。
 *
 * 收敛的是**可续跑性**，不是事实：重开本身没有给出任何"数据变了"的证据
 * （与 `wpsClarificationVerdict` 的 `unknown` 同一条纪律）。所以这里不说
 * "范围已变化"，只承认"没法再证明它没变"，把决定权交回用户。
 */
function reloadTaskResume(record: TaskResumeRecord): TaskResumeRecord {
  if (record.resumability === "unavailable") return record;
  /*
   * 写入状态**未知或待核对**时，重开只允许走核对 —— 不能因为"重开了一次"
   * 就退化成"重新生成预览并直接写"。后者会在原表可能已被改过的情况下再写一遍。
   */
  if (record.writeState === "unknown" || record.writeState === "needs_review") {
    return withTaskResumeState(record, {
      blocker: "execution_unknown",
      resumability: "reconcile"
    });
  }
  if (record.blocker === "user_input") {
    // 用户**还没回答**：先让他把数据范围重新确认一遍，再继续那道问题。
    return withTaskResumeState(record, {
      blocker: "authorization",
      resumability: "reauthorize"
    });
  }
  /*
   * 其它中断（断网 / 服务重启 / 任务窗格重载）保持"重新生成预览"的可续跑性。
   * 继续这个动作**本身**就是一次重新授权（重新读真实宿主元数据、推进授权版本），
   * 所以不必再额外要用户点一次"确认范围" —— 那只会让他多按一下同样的意图。
   */
  return withTaskResumeState(record, { blocker: "network" });
}

/** resolved / cancelled 是终态，优先于恢复卡渲染（否则已完成的任务会被再次兜售）。 */
function resolvedOrCancelledOutcome(message: ChatMessage): string | null {
  const status = message.clarification?.status;
  if (status === "resolved") {
    return `已按「${message.clarification?.answer ?? "补充说明"}」继续`;
  }
  /*
   * `"continued"`：重新确认了数据范围、任务作为新一轮继续 ——
   * **不是**回答了那道题。话必须这么说，否则等于用一次授权动作
   * 冒充一条答案，写入资格看起来也跟着"齐了"。
   */
  if (status === "continued") {
    return "已重新确认数据范围并按原需求作为新一轮继续（这不是对那道题的回答）";
  }
  if (status === "cancelled") return "已取消本次澄清，未修改工作表";
  return null;
}

/**
 * 这张卡是否还是一件事**可继续**的任务。
 *
 * 判据落在"有没有合法的恢复记录 + 该记录判定的动作是否可继续"上，
 * 而不是"消息是不是最新一条"。用户显式取消后记录被置为 `unavailable`，
 * 这里自然返回 `null` —— 取消是终态，不会被自动恢复。
 */
function resumableRecord(message: ChatMessage): TaskResumeRecord | null {
  if (!message.taskResume) return null;
  const record = sanitizeTaskResumeRecord(message.taskResume);
  if (!record) return null;
  return taskResumeDecision(record).canContinue ? record : null;
}

/** "已保留的信息"只说**保留了什么**，不复述内容（卡片上本来就看得到）。 */
function preservedSummary(record: TaskResumeRecord, localAnswerPending = false): string {
  const localNotice = localAnswerPending ? "；本地补充规则已提交（仅当前窗格内保存）" : "";
  if (record.confirmations.length === 0) {
    return localAnswerPending ? `原需求已保留${localNotice}` : "原需求已保留（尚未提交补充回答）";
  }
  const labels = record.confirmations.map((entry) => entry.label).join("、");
  return `原需求与你已提交的 ${record.confirmations.length} 项回答（${labels}）${localNotice}`;
}

/**
 * 中断卡本体。展示三件必须分开说的事：**中断原因 / 已保留的信息 / 原表写入状态**，
 * 再给出"继续任务"与"取消任务"。写入状态未知时如实写"未知"，绝不显示"原表未修改"。
 */
function TaskResumeRecoveryCard({
  record,
  scopeLabel,
  busy,
  resuming,
  canPickScope,
  localAnswerPending = false,
  onResume,
  onCancel
}: {
  record: TaskResumeRecord;
  scopeLabel: string;
  busy: boolean;
  resuming: boolean;
  canPickScope: boolean;
  localAnswerPending?: boolean;
  onResume: () => void;
  onCancel: () => void;
}) {
  const decision = taskResumeDecision(record);
  return (
    <div className="clarification-recovery">
      <dl className="clarification-recovery-facts">
        <div>
          <dt>中断原因</dt>
          <dd>{TASK_BLOCKER_LABEL[record.blocker]}</dd>
        </div>
        <div>
          <dt>已保留的信息</dt>
          <dd>{preservedSummary(record, localAnswerPending)}</dd>
        </div>
        <div>
          <dt>原表写入状态</dt>
          <dd>{TASK_WRITE_STATE_LABEL[record.writeState]}</dd>
        </div>
      </dl>
      <p className="clarification-recovery-detail">{decision.detail}</p>
      <p className="clarification-recovery-scope">
        当前文件与待授权工作表：{scopeLabel}
      </p>
      <div className="clarification-recovery-actions">
        <button
          type="button"
          disabled={busy || resuming || !canPickScope}
          onClick={onResume}
        >
          {decision.actionLabel}
        </button>
        <button type="button" disabled={busy || resuming} onClick={onCancel}>
          取消任务
        </button>
      </div>
      {!canPickScope && (
        <p className="clarification-recovery-hint">
          还没有选中任何工作表，先在数据范围里确认要使用的工作表。
        </p>
      )}
    </div>
  );
}

function invalidatePendingClarifications(messages: ChatMessage[]): ChatMessage[] {
  return messages.map((message) => message.clarification?.status === "pending"
    ? { ...message, clarification: { ...message.clarification, status: "invalidated" as const } }
    : message);
}

/**
 * 落盘前的收窄 + 脱敏，**唯一一份实现**。
 *
 * 现在有两条落盘路径：常规的 `useEffect`，以及"写入前必须已经在存储里"的
 * **同步**登记。两条必须用同一份收窄与脱敏规则 —— 各写一份迟早会漂移成
 * "同步那条忘了脱敏"，而那是隐私问题，不是风格问题。
 */
function serializeHistoryForStorage(history: WpsChatHistory): string {
  const compact: WpsChatHistory = {
    activeConversationId: history.activeConversationId,
    legacyCredentialConversationIds: remainingLegacyCredentialConversationIds(
      history as never
    ) as string[],
    conversations: [...history.conversations]
      .sort((left, right) => right.updatedAt.localeCompare(left.updatedAt))
      .slice(0, 30)
      .map((conversation) => ({
        ...conversation,
        messages: conversation.messages.slice(-120)
      }))
  };
  return JSON.stringify(redactHistoryCredentials(compact as never));
}

function loadWpsChatHistory(): WpsChatHistory {
  try {
    const stored = JSON.parse(localStorage.getItem(WPS_CHAT_STORAGE_KEY) ?? "null") as
      | Partial<WpsChatHistory>
      | null;
    const conversations = Array.isArray(stored?.conversations)
      ? stored.conversations.filter((conversation): conversation is WpsConversation =>
          Boolean(
            conversation &&
              typeof conversation.id === "string" &&
              Array.isArray(conversation.messages)
          )
        ).map((conversation) => ({
          ...conversation,
          messages: conversation.messages
            .filter((message) =>
              !(message.role === "system" && message.text === LEGACY_MODEL_CATALOG_NOTICE)
            )
            .map((message) => ({
              ...message,
              activityLog: restoreActivityLog(message.activityLog),
              /*
               * 恢复记录必须**结构化地**过一遍白名单，而不是原样信任本地存储：
               * 旧版本、被手改过、或混进越界字段的记录一律收敛（或整个丢弃）。
               * 宁可让用户重新发一次需求，也不把来源值之类的字段带回内存。
               */
              taskResume: (() => {
                const record = message.taskResume
                  ? sanitizeTaskResumeRecord(message.taskResume)
                  : null;
                return record ? reloadTaskResume(record) : undefined;
              })(),
              /*
               * 面板重开 = 内存对象身份必然丢失，属于"**证明不了**仍是原授权对象"，
               * 因此收敛为 `"unconfirmed"`（而不是 `"invalidated"`）。
               * 两者的区别不是措辞：`invalidated` 断言"范围已变"，重开后我们
               * 并没有那个证据，不能替用户下这个结论。
               */
              clarification: message.clarification?.status === "pending"
                ? { ...message.clarification, status: "unconfirmed" as const }
                : message.clarification
            }))
        }))
      : [];
    if (conversations.length > 0) {
      // 读取路径与 Excel 侧保持一致：先在原始存储上检测并脱敏，再进入可渲染
      // 状态（理由详见 conversation.ts 的 `loadChatHistory`）。凡是原始文本里
      // 检出凭据的会话，都登记到一个**独立且可持久化**的标记里 —— 脱敏之后
      // 就再也匹配不到密钥了，靠"还能不能匹配"做发送闸会在重开后自动松开。
      const legacyIds = new Set<string>(
        Array.isArray(stored?.legacyCredentialConversationIds)
          ? stored.legacyCredentialConversationIds.filter(
              (id): id is string => typeof id === "string"
            )
          : []
      );
      for (const conversation of conversations) {
        const carriedCredential =
          containsCredentialText(conversation.title) ||
          conversation.messages.some((message) =>
            containsCredentialText(message.text)
          );
        if (carriedCredential) legacyIds.add(conversation.id);
      }
      const sanitized = redactHistoryCredentials({
        activeConversationId: String(stored?.activeConversationId ?? ""),
        conversations,
        legacyCredentialConversationIds: [...legacyIds]
      } as never) as unknown as WpsChatHistory;

      // 这里**不做**措辞还原（`replaceRestoredSecrets`）：返回对象会被落盘
      // effect 再 `redactHistoryCredentials` 一次，若现在把 `【已隐藏】` 换成
      // 人话，落盘时匹配不到占位符，下一轮读取又会把真凭据渲染出来。措辞还原
      // 由 `renderWpsRestoredSecrets` 在挂载时做一次。
      const activeConversationId = sanitized.conversations.some(
        (conversation) => conversation.id === stored?.activeConversationId
      )
        ? String(stored?.activeConversationId)
        : sanitized.conversations[0].id;
      return { ...sanitized, activeConversationId };
    }
  } catch {
    // 损坏或不可用的宿主存储不应阻断任务窗格启动。
  }
  const conversation = createWpsConversation();
  return {
    activeConversationId: conversation.id,
    conversations: [conversation],
    legacyCredentialConversationIds: []
  };
}

/**
 * 把「落盘占位符」还原成界面措辞 —— 只做一次，且只在挂载时。
 *
 * 与 Excel 端同一考量：落盘 effect 每次都会把 state 重新脱敏写回存储，所以在
 * 它之前保留原样占位符是安全的；把还原放到读取路径里则会让占位符在落盘时消失，
 * 脱敏再也匹配不到，真凭据会在下一轮循环里重新渲染。
 */
function renderWpsRestoredSecrets(history: WpsChatHistory): WpsChatHistory {
  return {
    ...history,
    conversations: history.conversations.map((conversation) => ({
      ...conversation,
      title: replaceRestoredSecrets(conversation.title ?? ""),
      messages: conversation.messages.map((message) => ({
        ...message,
        text: replaceRestoredSecrets(message.text ?? "")
      }))
    }))
  };
}

function formatConversationTime(value: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleString("zh-CN", {
    month: "numeric",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

/** 是否为格式动作（走格式快照路径，预览显示格式变化而非"写入值"）。 */


/** 动作类型的短标签（预览标题用）。格式动作给明确描述，不冒充"写入值"。 */
function actionTypeLabel(type: string): string {
  if (type === "clearRange") return "清空内容";
  if (type === "writeFormulas") return "写入公式";
  if (type === "writeValues") return "写入值";
  if (type === "setFont") return "设置字体";
  if (type === "setAlignment") return "设置对齐";
  if (type === "resizeRange") return "调整尺寸";
  return "写入值";
}

export default function WpsAgentPanel({
  selectedSheetIndexes,
  resolveSelectedSheetIndexes,
  scopeMode = "manual",
  scopeSource = "workbook",
  folderScope = null,
  onAuthorizeTarget,
  onScopeConfirmationReady,
  scopeSelectionPending = false,
  onOpenScope,
  scopePickerOpen = false,
  scopeDock,
  scopeOverlay
}: WpsAgentPanelProps) {
  // 模型目录只包含用户真实配置的连接。初始列表为空、状态单独表达，
  // 不再注入虚拟的 “基础模式” 占位项——目录尚未加载或读取失败时，
  // 那会让人误以为“已经选好模型，可以开始对话”。
  const [models, setModels] = useState<ModelOption[]>([]);
  const [defaultModelId, setDefaultModelId] = useState<string | null>(null);
  const [catalogStatus, setCatalogStatus] = useState<ModelCatalogStatus>("idle");
  const [prompt, setPrompt] = useState("");
  const {
    pendingImages, setPendingImages, preparingImages, imageError, setImageError,
    draggingImage, addImageFiles, clearImages, removeImage, handleDrop,
    handleImagePaste, handleDragEnter, handleDragLeave, handleDragOver
  } = useImageAttachments();
  const imageInputRef = useRef<HTMLInputElement>(null);
  // Only names/required flags enter history; image bytes live in task-scoped memory.
  const taskImagesRef = useRef(new Map<string, PendingImage[]>());
  /**
   * 输入框的固定高度（`null` = 跟着内容自动长高）。与 Excel 端同义：由共享的
   * `useComposerResize` 读写，拖动上方把手或按 ↑/↓ 时变化，双击把手回到自动。
   */
  const [composerHeight, setComposerHeight] = useState<number | null>(null);
  const [chatHistory, setChatHistory] = useState<WpsChatHistory>(() =>
    renderWpsRestoredSecrets(loadWpsChatHistory())
  );
  const messages = useMemo(
    () =>
      chatHistory.conversations.find(
        (conversation) => conversation.id === chatHistory.activeConversationId
      )?.messages ?? [],
    [chatHistory]
  );
  /**
   * 当前对话是否被标为"旧记录里曾带真凭据"。
   *
   * 判据是**持久化标记**（`legacyCredentialConversationIds`），不是"文本里还能
   * 不能匹配到密钥"。读取路径已经把这些会话脱敏，脱敏后当然匹配不到；拿"匹配
   * 不到"当"已安全"会让重开之后发送闸自动打开。
   */
  const staleCredentials = useMemo(
    () =>
      isLegacyCredentialConversation(
        chatHistory as never,
        chatHistory.activeConversationId
      ),
    [chatHistory]
  );
  const [progress, setProgress] = useState<WpsReadonlyTurnProgress | null>(null);
  const [activeActivityMessageId, setActiveActivityMessageId] = useState<string | null>(null);
  const [writeProgress, setWriteProgress] = useState<string | null>(null);
  const [pendingPlan, setPendingPlan] = useState<WpsPreparedPlan | null>(null);
  type ResultReceipt = { executionId: string; verified: boolean; undoable: boolean; record: WpsUndoRecord };
  const [resultReview] = useState(() => new ResultReview<WpsBeforeResult, ResultReceipt, WpsResultMarks, WpsNativeComparison>());
  /*
   * 刷新/重载后认领上一次遗留的审阅标记：只按持久记录里的随机公式匹配，
   * 同区域里其它规则（别的任务、用户自建）一律不碰。
   */
  const recoveredMarks = useRef(false);
  useEffect(() => {
    if (recoveredMarks.current) return;
    recoveredMarks.current = true;
    try {
      const note = recoverWpsResultMarks(getWpsApplication());
      if (note) append({ role: "system", text: note });
    } catch { /* 认领失败不能影响面板启动 */ }
  }, []);
  const resultReviewState = useSyncExternalStore(resultReview.subscribe, resultReview.getSnapshot, resultReview.getSnapshot);
  const [resultTitle, setResultTitle] = useState('');
  /*
   * 终态**写回对话**：固定操作条在确认/放弃成功后收起，业务状态不能跟着消失。
   * 这里只补一条状态行（去重），保存工具与撤销入口仍由结果卡承载。
   */
  const settledLoggedRef = useRef<string | null>(null);
  useEffect(() => {
    const { phase, key, executionId } = resultReviewState;
    if (!key || (phase !== "confirmed" && phase !== "discarded")) return;
    /*
     * 这次 discard 是**撤销入口**造成的（`settleExternally` 只同步状态），
     * 不是用户点了「放弃变更」—— 那条路的结果文案已经由 `undoExecution` 给了
     * （"已撤销「…」，并回读确认恢复成功"），这里再补一句"已放弃并恢复"就是
     * 同一件事的重复反馈。跳过它，并清掉标记等下一轮。
     */
    if (phase === "discarded" && executionId && executionId === undoSettledExecutionIdRef.current) {
      undoSettledExecutionIdRef.current = null;
      return;
    }
    const stamp = `${key}:${phase}`;
    if (settledLoggedRef.current === stamp) return;
    settledLoggedRef.current = stamp;
    append({
      role: "assistant",
      text: phase === "confirmed"
        ? "已确认变更：已接受当前结果并移除本次审阅标记；保存为工具与「撤销本次执行」可用。"
        : "已放弃并恢复：主表已按本次恢复点还原。"
    });
  }, [resultReviewState.phase, resultReviewState.key, resultReviewState.executionId]);
  /**
   * 已执行计划的**展示副本**：只用于事后复核"按什么口径改的"。
   *
   * 执行后 `pendingPlan` 会被清空（它是"待执行"的现场），但这张卡要留在对话里，
   * 否则用户事后只能看到"已执行"，看不到采用了哪些口径、格式动作改了什么。
   */
  const [executedPlan, setExecutedPlan] = useState<{ messageId: string; plan: WpsPreparedPlan; revised: boolean } | null>(null);
  const resultUndoRef = useRef<(id: string, origin?: "undo" | "discard") => Promise<boolean>>(async () => false);
  // 原生左右子窗口只活在本次面板会话；副本矩阵绝不进入聊天历史或持久存储。
  const [stagedReview, setStagedReview] = useState<WpsNativeComparison | null>(null);
  /**
   * 审阅确认。**换一份副本就要重新确认** —— 用副本名做键，名字变了就归零。
   * 与 Excel 端同一套语义：没勾选就不是"已核对"。
   */
  const [reviewAcknowledged, setReviewAcknowledged] = useState(false);
  /** 执行方式：默认副本比对（多一道预览），用户可改为直接执行。 */
  /*
   * 执行方式只剩一种：**先在授权范围内执行，再审阅结果**。
   *
   * 用户已明确取消"执行方式选择器"与"先生成副本再写回"的旧流程。保留这个常量是
   * 为了让写回身份里仍带着"这一次是哪条路径"——它不是可选项，也不需要界面。
   */
  const executionMode: PlanExecutionMode = "direct";
  const stagedCopyKey = stagedReview ? stagedReview.copy.Name ?? null : null;
  useEffect(() => { setReviewAcknowledged(false); }, [stagedCopyKey]);
  // 首次生成副本时调整窗格；消息更新不会主动重新排列。
  const reviewPaneNeedsResize = useWpsReviewPaneRatio(stagedReview);
  const [pendingPlanMessageId, setPendingPlanMessageId] = useState<string | null>(null);
  const [planEditDraft, setPlanEditDraft] = useState<string | null>(null);
  const [planEditError, setPlanEditError] = useState<string | null>(null);
  useEffect(() => {
    setPlanEditDraft(null);
    setPlanEditError(null);
  }, [pendingPlan?.previewRevision]);
  const [undoRecord, setUndoRecord] = useState<WpsUndoRecord | null>(null);
  /**
   * 正在撤销的那一次执行（`executionId`）。
   *
   * 撤销**不新增执行过程卡片**：处理中只把原入口的文字换成「正在撤销…」并禁用
   * 重复点击，逐步进度留在内部（`writeProgress` 单行指示）不进对话。
   * 所以这里只需要知道"哪一次正在撤"，好让对应的按钮改字、变灰。
   */
  const [undoingExecutionId, setUndoingExecutionId] = useState<string | null>(null);
  /**
   * 因「撤销本次执行」而收尾的那次审阅（`executionId`）。
   *
   * `settleExternally` 只同步审阅状态，**不应**再生成一条放弃文案 ——
   * 那条消息是用户点「放弃变更」时才有的结果。这里记下是哪一次，
   * 收尾消息效应据此跳过，避免"已撤销…"与"已放弃并恢复…"重复反馈。
   */
  const undoSettledExecutionIdRef = useRef<string | null>(null);
  /**
   * CAP-03 重载恢复：面板重开后从本地发现索引查服务端账本，核出这次未完成写入
   * 的现场。`null` = 没有需要核对的写入；否则是服务端状态 + 身份/现场核对结论。
   */
  const [wpsRecovery, setWpsRecovery] = useState<{
    operationId: string;
    /** 这次执行所属任务。重新预览要靠它找回原始需求，不能按"最近一次任务"猜。 */
    taskId: string;
    state: WpsExecutionState | null;
    identity: WpsIdentityVerdict;
    verdict: WpsReconcileVerdict;
    checking: boolean;
    /** 正在提交补记/核对/重新预览。防止连点造成重复回执。 */
    acting: boolean;
  } | null>(null);
  // 双栏的完整前后矩阵只活在这里（按 executionId 索引的会话内内存）。
  const recoveryActionLock = useRef(false);
  const [recoveryScopeConfirmed, setRecoveryScopeConfirmed] = useState(false);
  useEffect(() => { setRecoveryScopeConfirmed(false); }, [wpsRecovery?.operationId, selectedSheetIndexes]);
  // 刻意不挂到 ChatMessage 上：消息会持久化到 WPS_CHAT_STORAGE_KEY，
  // 矩阵不许进历史、也不许出站（方案 §7）。消息里只留一个 executionId 当索引。
  const changeReviewsRef = useRef(new Map<string, RangeChangeReview[]>());
  // 撤销成功后完整矩阵要释放，但"这次确实改过、后来撤销了"是历史事实，不能一并抹掉。
  // 留下的是 changeReviewSummary 形态：只有定位与计数，没有任何单元格内容。
  const changeReviewSummariesRef = useRef(new Map<string, ChangeReviewSummary[]>());
  const [clarification, setClarification] = useState<PendingClarification | null>(null);
  const [clarificationDraft, setClarificationDraft] = useState("");
  const [clarificationFieldDrafts, setClarificationFieldDrafts] = useState<Record<string, string>>({});
  /**
   * 当前活跃的规划暂停（CAP-02）。
   *
   * `payload` 是服务端给的暂停现场快照，只活在内存里（含模型连接元数据，不进历史）。
   * `view` 是当时那一刻推出的续跑方案；**点击续跑时会重新查询服务端状态**，
   * 因为 `revision` 可能已被后台推进 —— 用旧 revision 继续会直接被 CAS 拒绝。
   */
  const [planningPause, setPlanningPause] = useState<{
    turnId: string;
    payload: PlanningPauseResult;
    view: WpsPlanningResumeView;
    originalPrompt: string;
    sourceMessageId: string;
  } | null>(null);
  /**
   * 面板重开后正在核对的暂停（CAP-02 第 5 条）。
   *
   * 历史里那些暂停卡是**当时那一刻**的快照。重开后必须先拿 `turnId` 去服务端
   * 查真实状态，再决定还画不画"继续任务"按钮 —— 中间它可能已经跑完了，
   * 或者被别的窗口接手了。核完之前一律**不画可点按钮**，避免用户点到一份
   * 基于过期 `revision` 的动作（服务端 CAS 必然拒绝，界面上却像"按钮坏了"）。
   *
   * 只放 `turnId`，不放任何载荷 —— 载荷以服务端返回为准。
   */
  const [planningPauseReload, setPlanningPauseReload] = useState<{
    /** 正在核对 / 已核对的那一条。`null` = 历史里没有暂停需要核对。 */
    turnId: string | null;
    checking: boolean;
    /** 核对结论：服务端说不能继续时的真实原因（展示用，不参与判定）。 */
    blockedReason: string | null;
    canRecheck?: boolean;
  } | null>(null);
  /**
   * 续跑请求的幂等键（CAP-02 第 2 条）。
   *
   * **同一次继续操作固定一个 requestId**：用户点"继续"时生成一次，此后无论
   * 超时重试多少次、网络抖动重发多少次，都沿用同一个值。服务端按
   * `(turn_id, request_id)` 判重，同一个值只会追加一次预算、只跑一轮规划。
   *
   * 只有新的一次继续轮次（上一次已经结束、用户重新点击）才换新身份。
   */
  const resumeRequestRef = useRef<{ turnId: string; requestId: string } | null>(null);
  /**
   * 续跑并发闸。
   *
   * 用 ref 而非 `busy`：`busy` 是异步渲染出来的，连点两下时两帧都读到 `false`，
   * 于是两次请求都发出去 —— 正好是幂等键要防的那种重复。ref 的读写是同步的，
   * 第二次点击必然看到第一次已经置位。
   */
  const invalidPlanningTurnsRef = useRef(new Set<string>());
  const resumingPauseRef = useRef(false);
  const [continuingTurnId, setContinuingTurnId] = useState<string | null>(null);
  /**
   * 续跑要带回的处理上下文：原需求、原授权表序号、原澄清轮次。
   *
   * 续跑结果进入 `applyTurnResult` 时需要"这是哪一件事"的上下文，否则计划
   * 预览会挂到一个空的 taskId 上、澄清卡的轮次也会归零。这些字段在发起
   * 规划时就记下，暂停后仍留在内存里（不落盘：它们含原需求文本）。
   */
  const pauseCarryRef = useRef<{
    turnId: string;
    originalPrompt: string;
    selectedSheetIndexes: number[];
    clarificationRound: number;
  } | null>(null);
  /**
   * 本轮规划在服务端的任务身份，由 `runWpsTurn` 在意图阶段上报。
   *
   * 断线恢复（第 4 条）靠它去查真实状态。没拿到身份就**不能**声称"服务端没有
   * 这条任务" —— 那会把"我不知道"谎报成"它不存在"，进而放行一次重复规划。
   */
  const turnIdRefForStream = useRef<string | null>(null);
  const streamCursorRef = useRef<{ turnId: string; sequence: number } | null>(null);
  const pauseRequestedRef = useRef<AbortController | null>(null);
  /**
   * 暂停请求到了哪一步（三件事分开，不许互相冒充）：
   *
   *  - `requesting`：刚发出请求，还没有任何答案；
   *  - `local-stopped`：**本端**已经停止接收（这是按下暂停那一刻就成立的事实）；
   *  - `acknowledged`：服务端接受了暂停请求（`pauseRequested: true`）—— 仍然
   *    **不等于**它已经停下，停下与否只由任务状态决定；
   *  - `unconfirmed`：请求失败/超时，服务端状态未知（如实说不知道）。
   *
   * 旧实现把"断开 SSE"直接当成"已暂停"，失败分支只吞一句"未能确认"。
   */
  const pauseRequestGenerationRef = useRef(0);
  const [pauseRequestState, setPauseRequestState] = useState<{
    phase: "requesting" | "local-stopped" | "acknowledged" | "unconfirmed";
    detail: string;
  } | null>(null);
  const prePlanningPause = usePrePlanningPause<{
    prompt: string; images: PendingImage[]; taskContext?: AgentTaskContext;
    activityMessageId: string; messageUnderstanding?: Awaited<ReturnType<typeof understandMessage>>;
  }>(JSON.stringify([chatHistory.activeConversationId, scopeSource, scopeMode, selectedSheetIndexes]));
  const planningInspectionRef = useRef<AbortController | null>(null);
  const pendingInspectionMessageRef = useRef<string | null>(null);
  // 原任务的宿主身份只驻内存，不落盘、不传给模型。
  const planningTargetRef = useRef<WpsObjectIdentity | null>(null);
  /**
   * 当前用户任务的**稳定身份**。
   *
   * 每次新的用户请求（`showUser !== false`）换一个 —— 一个 taskId 对应一件事；
   * 同一任务内的重新规划、重新授权、重新生成预览都沿用它，因此"换了 turnId
   * 仍是同一件任务"在界面上是可判定的，恢复卡也据此绑定到原消息。
   */
  const taskIdRef = useRef<string>(newTaskId());
  // Keep recovery ownership separate from whichever message is currently being understood.
  const planningContextsRef = useRef(new Map<string, {
    conversationId: string;
    carry: NonNullable<typeof pauseCarryRef.current>;
    target: WpsObjectIdentity | null;
    taskId: string;
    cursor: typeof streamCursorRef.current;
    resumeIdentity: typeof resumeRequestRef.current;
  }>());
  const retainPlanningContext = () => {
    const carry = pauseCarryRef.current;
    if (!carry) return;
    planningContextsRef.current.set(carry.turnId, {
      conversationId: chatHistory.activeConversationId, carry,
      target: planningTargetRef.current, taskId: taskIdRef.current,
      cursor: streamCursorRef.current, resumeIdentity: resumeRequestRef.current
    });
  };
  /**
   * 从历史任务继续时要**带过去**的东西：原 `taskId` 与用户已确认过的回答。
   *
   * 为什么必须显式携带：`taskIdRef` 的初值是面板挂载时随机生成的 —— 重开之后
   * 它就是**另一个身份**。若继续时不用记录里的 `taskId`，新的记录会挂到一个
   * 与历史无关的 id 上，"同一件事"当场断掉；确认信息也会被丢掉，用户被迫重答。
   *
   * 只在"继续历史任务"这一轮生效；该轮结束后清空，避免影响后续新任务。
   */
  const resumeCarryRef = useRef<{
    taskId: string;
    confirmations: TaskResumeConfirmation[];
    originalPrompt?: string;
  } | null>(null);
  // 本地补全结果只在内存中等待显式重新授权，绝不进入历史或模型请求。
  const localMappingResumeRef = useRef<{
    pending: PendingClarification;
    rule: NonNullable<WpsTurnClarification["localMapping"]>["rule"];
    identity: ReturnType<typeof captureWpsObjectIdentity> | null;
  } | null>(null);
  const retryTurnRef = useRef<(question: string, options?: {
    intentConfirmed?: boolean; clarificationRound?: number; showUser?: boolean; targetConfirmedRetry?: boolean; images?: PendingImage[]; messageUnderstanding?: Awaited<ReturnType<typeof understandMessage>>; taskContext?: AgentTaskContext;
  }) => void>(() => {});
  const resumeTaskRef = useRef<(message: ChatMessage) => void>(() => {});
  const targetGuard = useAuthorizedTarget({
    selectionPending: () => scopeSelectionPending,
    enabled: true,
    reviewActive: stagedReview !== null,
    monitorPaused: progress !== null || writeProgress !== null,
    /*
     * 副本识别只认**本次审阅创建时登记的身份**（`wpsResultReviewTarget`）。
     * 旧实现挂在 `stagedReview` 上，而新流程的副本由结果审阅通道创建、
     * `stagedReview` 恒为 null —— 于是副本一获得焦点就被判成"换了工作簿"。
     */
    isReviewCopy: (target) => isWpsReviewCopyTarget(target) || isWpsReviewCopyWorkbook(target.fileRef as never),
    inspect: () => inspectWpsAuthorizedTarget(scopeMode === "auto" ? null : selectedSheetIndexes),
    analysisEvidence: wpsAnalysisTargetEvidence,
    adopt: (target) => {
      onAuthorizeTarget?.(target.selected.map(Number));
      /*
       * **只有授权真的变了才递增版本。**
       *
       * 原来在比较之前就 `+1`：用户只是"核对一下范围"（尤其副本获得焦点后
       * 再核对）也会把 `authorizationRevision` 推进一格，于是正在进行的审阅
       * 被判成"计划/授权已变化"，确认与放弃反复失败。版本号是"已提交范围的
       * 身份"，不是"用户点过确认"的计数。
       */
      const nextKey = authorizationKeyOf(target.selected.map(Number));
      if (nextKey === authorizationRevisionBaseRef.current) return;
      authorizationRevisionBaseRef.current = nextKey;
      setAuthorizationRevision((value) => value + 1);
      // Old data-derived artifacts cannot be adopted by a newly authorized target.
      localMappingResumeRef.current = null;
      clarificationObjectIdentityRef.current = null;
      setPendingPlan(null);
      setPendingPlanMessageId(null);
      setPlanEditDraft(null);
      setClarification(null);
      setChatHistory((current) => ({ ...current, conversations: current.conversations.map((conversation) =>
        ({ ...conversation, messages: invalidatePendingClarifications(conversation.messages).map((message) =>
          message.planPreview?.status === "pending"
            ? { ...message, planPreview: { ...message.planPreview, status: "cancelled" as const } } : message) })) }));
    }
  });
  const scopeConfirmationRef = useRef<typeof targetGuard.confirmScope>(targetGuard.confirmScope);
  scopeConfirmationRef.current = (target, options) => targetGuard.confirmScope(target, { continuePendingAnalysis: !options, ...options });
  useEffect(() => {
    onScopeConfirmationReady?.((target, options) => scopeConfirmationRef.current(target, options));
    return () => onScopeConfirmationReady?.(null);
  }, [onScopeConfirmationReady]);
  /**
   * 这一轮该用的任务身份与要携带的确认信息：从历史继续时取自原记录，
   * 否则用当前任务的。**集中在这里**而不是散在各调用点 —— 漏掉任何一处，
   * 新的记录就会悄悄挂到另一个身份上。
   */
  const carried = () => ({
    taskId: resumeCarryRef.current?.taskId ?? taskIdRef.current,
    confirmations: resumeCarryRef.current?.confirmations ?? []
  });
  /**
   * 显式确认的数据授权版本。只在**已提交范围**真正变化时递增 ——
   * 勾选草稿的增删不改它，否则用户每取消一次勾选就会把待恢复任务打掉。
   */
  const [authorizationRevision, setAuthorizationRevision] = useState(1);
  const authorizationRevisionBaseRef = useRef(authorizationKeyOf(selectedSheetIndexes));
  /** 正在继续的任务所属消息：用于禁用重复点击。 */
  const [resumingTaskMessageId, setResumingTaskMessageId] = useState<string | null>(null);
  /**
   * 上一轮 `runTurn` 是否以失败告终。
   *
   * `runTurn` 自己吞掉异常并写系统消息，调用方拿不到结果；续答路径需要据此
   * 决定"这张卡是已完成，还是变成了可恢复的中断卡"。
   */
  const lastTurnFailedRef = useRef(false);
  const [themeOpen, setThemeOpen] = useState(false);
  const [petVisible, setPetVisible] = useState(() => {
    try {
      return normalizePetVisibility(localStorage.getItem(PET_VISIBILITY_STORAGE_KEY));
    } catch {
      return true;
    }
  });
  const [modelMenuOpen, setModelMenuOpen] = useState(false);
  const [modelSettingsOpen, setModelSettingsOpen] = useState(false);
  const [toolsOpen, setToolsOpen] = useState(false);
  const [backupOpen, setBackupOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [moreMenuOpen, setMoreMenuOpen] = useState(false);
  // 开发期能力探针只活在当前面板内存，不进入聊天历史或本地存储。
  const [wideViewProbe, setWideViewProbe] = useState<WideViewDialogProbe | null>(null);
  const [wideViewProbeError, setWideViewProbeError] = useState<string | null>(null);
  const [wideViewChannelStatus, setWideViewChannelStatus] = useState<ChannelProbeStatus | null>(null);
  const wideViewChannelStopRef = useRef<(() => void) | null>(null);
  useEffect(() => () => wideViewChannelStopRef.current?.(), []);
  const [ruleManagerOpen, setRuleManagerOpen] = useState(false);
  const [tools, setTools] = useState(loadTools);
  const [queryTools] = useState(loadQueryTools);
  // 已保存工具的「编辑计划模板」抽屉：非空表示正在编辑某个固化工具的计划 JSON。
  const [toolPlanEdit, setToolPlanEdit] = useState<{ toolId: string; draft: string; error: string | null } | null>(null);
  // 验证通过后、可供「保存为工具」的最新计划（绑定实际已验证的执行身份）。
  const verifiedPlanForSaveRef = useRef<{ plan: AnalysisPlan; sourceSheetNames: string[]; executionId: string } | null>(null);
  // 保存为工具的候选上下文（执行身份 + 授权表名），与共用的 saveCandidate 配套。
  const saveToolContextRef = useRef<{ executionId: string; sourceSheetNames: string[] } | null>(null);
  const [toolName, setToolName] = useState("");
  const [toolDescription, setToolDescription] = useState("");
  /**
   * 保存为工具的资格与批准状态，复用 Excel 同一套状态机（`useExecutionApproval`）：
   * `beginSaveCandidate` 重置确认项、`saveEligibility` 由 `analyzeToolEligibility` 派生。
   */
  const {
    saveCandidate: saveToolCandidate,
    approveFixedContent,
    setApproveFixedContent,
    approveDestructive,
    setApproveDestructive,
    saveEligibility,
    beginSaveCandidate,
    closeSaveCandidate
  } = useExecutionApproval();
  const [copiedMessageId, setCopiedMessageId] = useState<string | null>(null);
  const [copyError, setCopyError] = useState(false);
  const [newConversationNotice, setNewConversationNotice] = useState<{
    id: number;
    previousSaved: boolean;
  } | null>(null);
  const [newConversationConfirmOpen, setNewConversationConfirmOpen] = useState(false);
  const controllerRef = useRef<AbortController | null>(null);
  /**
   * 当前读表桥的主人（执行身份）。`null` = 桥已收或从未建立。
   *
   * 与 `controllerRef` 分开记：核对接管 controller 之后，桥仍然属于立它的那次
   * 执行。绑在同一个布尔上，就会出现"controller 被接走 ⇒ 桥没人收"。
   */
  const bridgeOwnerRef = useRef<string | null>(null);
  /** 本轮"建桥失败已经说过一次"的闸；每轮规划开始时复位。 */
  /** 每次执行前重读的宿主现场；由 runTurn 在发起本轮规划前填好。 */
  const wpsBridgeLiveContextRef = useRef<() => WpsSdkBridgeLiveContext>(() => ({
    workbookName: "",
    workbookIdentity: "",
    sourceFingerprint: null,
    authorizedSheetNames: [],
    selectedSheetIndexes: [],
    activeWorksheet: ""
  }));
  /**
   * SDK 工具桥（WPS 端）。
   *
   * 与 Excel 端共用同一个 hook 与同一份状态机，只把"怎么读宿主"换成 JSAPI。
   * 桥只做**读**：轮询 → 复核当前授权 → 真实读取 + 确定性计算 → 回执。
   */
  const sdkBridge = useSdkToolBridge({
    hostId: "wps-et",
    createPorts: (options, binding) =>
      createWpsSdkBridgeHostPorts(
        {
          ...options,
          liveContext: () => wpsBridgeLiveContextRef.current()
        },
        binding
      )
  });
  /**
   * 建立/复用本次轮次的活桥。
   *
   * 首次规划与"继续分析"走**同一个入口** —— 两份实现迟早分叉，而分叉的表现是
   * "首次能用、续跑读不到表"这种只在特定路径出现的哑失败。
   *
   * `sheetIndexes` 由调用方给：首次是 `turnSheetIndexes`，续跑是建卡时记下的
   * 授权表序号。桥读取的范围必须与"刚刚核实过并交给服务端的证据"是同一批表。
   *
   * `owner` 是**这一次执行**的身份：桥立起来时记下主人，收尾时只由主人关。
   * 不记主人的后果是实测过的 —— 旧执行（或其核对）把新执行的桥关掉，新一轮
   * 一条读表指针都发不出去。
   */
  const ensureWpsBridge = async (
    turnId: string,
    sheetIndexes: readonly number[],
    signal?: AbortSignal,
    deferSource = false,
    owner: string | null = null
  ): Promise<boolean> => {
    let includeSource = !deferSource;
    wpsBridgeLiveContextRef.current = () =>
      createWpsBridgeLiveContext(getWpsApplication(), sheetIndexes, includeSource);
    const live = wpsBridgeLiveContextRef.current();
    const ready = await sdkBridge.ensureLive({
      taskId: turnId,
      workbookName: live.workbookName,
      workbookIdentity: live.workbookIdentity,
      sourceFingerprint: live.sourceFingerprint,
      signal,
      liveContext: () => wpsBridgeLiveContextRef.current(),
      onProgress: ({ scannedRows, totalRows, sheet }) => {
        /*
         * **迟到进度不许复活运行态。**
         *
         * 暂停之后再到达的读取进度会重新 `setProgress`，界面于是又出现一条
         * "正在读取「甲」"，暂停按钮也重新变成 ■ —— 而本端已经没有可中止对象。
         * 判据用**立桥那次执行自己的 signal**：它一 abort，这条桥的任何进度都不
         * 再属于界面。
         */
        if (signal?.aborted) {
          recordPauseDiagnostic({ kind: "late-progress-ignored", source: "system", executionId: owner ?? undefined });
          return;
        }
        setProgress({
          stage: "reading",
          message: `正在读取「${sheet}」：${scannedRows} / ${totalRows} 行`
        });
      }
    });
    includeSource = true;
    if (ready && owner) bridgeOwnerRef.current = owner;
    return ready;
  };
  /**
   * 暂停这一次执行。
   *
   * 三件事必须分开说，因为它们成立的时刻不同：
   *  ① **请求取消**：`abort()` 一调用，本端的连接立刻断 —— 这是本端唯一能自己
   *     确定的事实；
   *  ② **服务端暂停**：只有 `/pause` 返回 `pauseRequested: true` 才算**请求被接受**，
   *     是否真的停下仍以任务状态为准；
   *  ③ **界面停止**：活动卡、计时、按钮状态各自由收尾逻辑处理，与 ①② 都无关。
   *
   * 旧实现在没有可中止对象时**静默返回**（点了没反应），并且把失败分支并入一句
   * 事后提示 —— 用户无从分辨"没停"和"停了但界面没跟上"。
   */
  const pauseCurrentTurn = (source: "composer" | "card" = "composer") => {
    const controller = controllerRef.current;
    const owner = controllerOwnerOf(controller);
    const taskId = streamCursorRef.current?.turnId ?? turnIdRefForStream.current;
    recordPauseDiagnostic({
      kind: "pause-click",
      source,
      executionId: owner ?? undefined,
      turnIdTail: tailOf(taskId),
      controller: !controller ? "none" : controller.signal.aborted ? "aborted" : "owned",
      bridge: bridgeOwnerRef.current ? (bridgeOwnerRef.current === owner ? "owned" : "other") : "none"
    });
    if (!controller || controller.signal.aborted) {
      /*
       * 没有任何可中止对象：**不许静默**。
       *
       * 把原因说清楚，同时把界面状态更新到位 —— 否则用户看到的是一个"点了没反应"
       * 的按钮，而它其实根本不该显示成可暂停。
       */
      setProgress(null);
      setPauseRequestState({ phase: "unconfirmed", detail: PAUSE_UNAVAILABLE_REASON });
      recordPauseDiagnostic({
        kind: "pause-unavailable", source, executionId: owner ?? undefined, reasonCode: "no-controller"
      });
      recordPauseDiagnostic({ kind: "pause-requested", source, executionId: owner ?? undefined, outcome: "skipped" });
      return;
    }
    /*
     * 用户真的按了暂停：这一次继续动作到此结束。
     *
     * 只有在这里清空"本次继续已开的段"，才能同时满足两件互相拉扯的事：
     * ① 同一次继续内部的多条分支（接回 running → 取回结果）只落**一段**；
     * ② 用户再次暂停后重新继续，应当**另开一段**，而不是继续往旧段里灌。
     * 若改在收尾（`endActivity`）清空，① 就会断掉 —— 中间那次收尾会把段放走，
     * 后续分支再开一段，一次点击串出两段。
     */
    resumeSegmentRef.current = null;
    pauseRequestedRef.current = controller;
    const pauseGeneration = ++pauseRequestGenerationRef.current;
    setPauseRequestState({ phase: "requesting", detail: "正在请求暂停…" });
    controller.abort();
    if (!taskId) {
      /*
       * 还没有服务端任务身份（意图阶段就按了暂停）：没有可请求的对象，就**不请求**。
       * 只说本端已经停止接收 —— 编一个请求发出去只会得到 404，再把 404 说成"已暂停"。
       */
      setPauseRequestState({
        phase: "local-stopped",
        detail: "已在本端停止接收进度；这次执行还没有提交到后台任务，不需要后台暂停。"
      });
      recordPauseDiagnostic({ kind: "pause-request", source, executionId: owner ?? undefined, outcome: "skipped" });
      return;
    }
    void (async () => {
      try {
        const ack = await withPauseTimeout(
          pausePlanningTask(taskId),
          PAUSE_REQUEST_TIMEOUT_MS,
          "暂停请求"
        );
        if (pauseGeneration !== pauseRequestGenerationRef.current) return;
        if (ack.pauseRequested) {
          recordPauseDiagnostic({ kind: "pause-ack", source, executionId: owner ?? undefined, turnIdTail: tailOf(taskId), outcome: "confirmed" });
          setPauseRequestState({
            phase: "acknowledged",
            detail: "本端已停止接收；服务端已接受暂停请求（是否已停下以任务状态为准）。"
          });
          return;
        }
        // `pauseRequested: false` = 服务端此刻没有在跑这条任务（多半已经结束）。
        recordPauseDiagnostic({ kind: "pause-ack", source, executionId: owner ?? undefined, turnIdTail: tailOf(taskId), outcome: "skipped" });
        setPauseRequestState({
          phase: "local-stopped",
          detail: "本端已停止接收；服务端这条任务此刻不在运行（可能已结束），稍后核对即可看到结果。"
        });
      } catch (error) {
        if (pauseGeneration !== pauseRequestGenerationRef.current) return;
        const timedOut = error instanceof PauseTimeoutError;
        recordPauseDiagnostic({
          kind: timedOut ? "pause-timeout" : "pause-failed",
          source,
          executionId: owner ?? undefined,
          turnIdTail: tailOf(taskId),
          outcome: "unconfirmed"
        });
        setPauseRequestState({
          phase: "unconfirmed",
          detail: timedOut
            ? `本端已停止接收，但暂停请求在 ${Math.round(PAUSE_REQUEST_TIMEOUT_MS / 1000)} 秒内没有确认；后台可能仍在运行。`
            : "本端已停止接收，但未能确认后台暂停；请稍后核对原任务状态。"
        });
      }
    })();
  };
  const recordPlanningResult = (update: WpsReadonlyTurnProgress) => {
          // Keep the result attached to its producing task, even if planning
          // later fails. Receiving a result never confirms or executes a plan.
          if (update.pythonResult && update.turnId === sdkBridge.binding()?.taskId) {
            const card = update.pythonResult;
            const binding = localResultBindingOf(card);
            const id = `python-result-${update.turnId}-${card.resultRef}`;
            updateMessages((messages) => messages.some((message) => message.id === id)
              ? messages : [...messages, { id, role: "assistant", text: "隔离 Python 分析结果已生成。", createdAt: new Date().toISOString(),
                localResult: card, localResultBinding: binding }]);
          }
  };
  const recordPlanningProgress = (update: WpsReadonlyTurnProgress) => {
    if (!update.turnId) return;
    turnIdRefForStream.current = update.turnId;
    const previous = streamCursorRef.current;
    const sequence = Number.isSafeInteger(update.sequence) && update.sequence! >= 0 ? update.sequence! : 0;
    streamCursorRef.current = { turnId: update.turnId,
      sequence: Math.max(previous?.turnId === update.turnId ? previous.sequence : 0, sequence) };
  };
  const receivedSequence = (turnId: string) => planningReceivedSequence(turnId, [
    { taskId: streamCursorRef.current?.turnId, sequence: streamCursorRef.current?.sequence },
    ...messages.flatMap((message) => [
      { taskId: message.activityLog?.operationId, sequence: message.activityLog?.sequence },
      { taskId: message.planningPause?.turnId, sequence: message.planningPause?.clientLastSequence }
    ])
  ]);
  /**
   * 建卡时用的**授权表序号**。
   *
   * 复核表头指纹时要对着同一批表重新读第 1 行；若不复用建卡时的序号，
   * 用户中途改了勾选，就会拿另一批表的表头去比，判据失去意义。
   */
  const clarificationScopeIndexesRef = useRef<readonly number[]>([]);
  /**
   * 建卡时捕获的**宿主对象身份**（工作簿/工作表对象引用 + 可选稳定 ID + 结构）。
   *
   * 只有它活得过"面板重开"以外的场景：它**不可序列化**，所以既不进
   * localStorage，也不随消息持久化 —— 面板重开后自然为空，待答澄清随之失效。
   * 指纹负责跨会话能证明的部分，它负责"同会话内还是不是那个对象"。
   *
   * 完整路径也在里面，所以这个 ref **只用于内存比对**：不得写进请求、
   * 日志、错误文案或落盘历史。
   */
  const clarificationObjectIdentityRef = useRef<WpsObjectIdentity | null>(null);
  const composerInputRef = useRef<HTMLTextAreaElement | null>(null);
  /**
   * 输入框的「上拉」行为与 Excel 端共用同一份实现（拖动把手调高度、内容变多时
   * 自动长高，长到上限之后**可滚动**）。之前这里靠 CSS 的原生 `resize: vertical`，
   * 在 WPS 的 WebView 里拖不动，超出部分还被继承来的 `overflow-y: hidden` 切掉。
   */
  const { handleProps: composerResizeHandleProps } = useComposerResize({
    inputRef: composerInputRef,
    height: composerHeight,
    setHeight: setComposerHeight,
    value: prompt,
    // 与 styles.css 里 `.wps-chat-shell .composer-box textarea` 的 min-height 对齐。
    minHeight: 56
  });
  const authorizationKey = useMemo(
    () => authorizationKeyOf(selectedSheetIndexes),
    [selectedSheetIndexes]
  );
  /**
   * 「当前文件 + 待授权工作表」的展示文本。
   *
   * 恢复卡必须让用户看清自己**在哪份文件、授权了哪几张表**上继续 ——
   * 光说"重新确认范围"而不把当前范围摆出来，等于让他闭着眼睛签字。
   * 读不到宿主时如实说读不到，不拿旧名字冒充当前状态。
   */
  const currentScopeLabel = useMemo(() => {
    try {
      const metadata = listWpsWorksheets(getWpsApplication());
      const names = metadata.sheets
        .filter((sheet) => selectedSheetIndexes.includes(sheet.index))
        .map((sheet) => sheet.name);
      if (names.length === 0) return "尚未选择工作表";
      return `${metadata.name}：${names.join("、")}`;
    } catch {
      return "无法读取当前工作簿信息";
    }
  }, [authorizationKey, selectedSheetIndexes]);
  const authorizationKeyRef = useRef(authorizationKey);
  const writeCapability = useMemo(() => {
    try {
      return probeWpsWriteCapabilities(getWpsApplication(), selectedSheetIndexes);
    } catch {
      return null;
    }
  }, [authorizationKey, selectedSheetIndexes]);
  const canWrite = writeCapability?.writeEnabled === true;
  const developmentWrite = writeCapability?.versionStatus === "development";
  const writeStatusReason = writeCapability?.checks.find((check) => !check.passed)?.detail
    ?? "确认任务时会检查当前工作簿是否满足写入条件";
  const latestActivity = [...messages].reverse().find((message) => message.activityLog)?.activityLog;
  const scrollSignal = [
    chatHistory.activeConversationId,
    messages.length,
    latestActivity?.totalMs ?? 0,
    latestActivity?.steps.length ?? 0,
    latestActivity?.steps.at(-1)?.current ?? "",
    latestActivity?.outcome ?? "",
    progress?.message ?? "",
    writeProgress ?? ""
  ].join(":");
  const {
    containerRef: messageStreamRef,
    awayFromBottom,
    handleScroll: handleMessageStreamScroll,
    scrollToLatest
  } = useMessageAutoFollow(scrollSignal, chatHistory.activeConversationId);
  const theme = useTheme();
  const logoLongPress = useLongPress(() => setThemeOpen(true));
  const worksheetCheckpoint = useWorksheetCheckpoint("wps-et");
  /*
   * 检查点归属用的工作簿身份：与写入账本同一个规范化函数（`wps-path:` 前缀），
   * 这样"刷新后回挂到原消息"和"写入恢复"认的是同一本工作簿。
   * 路径仍保留原始写法，供「备份与恢复」按工作簿查询恢复点。
   */
  const checkpointWorkbookPath = useMemo(() => {
    try { return listWpsWorksheets(getWpsApplication()).fullName || null; } catch { return null; }
  }, [authorizationKey, selectedSheetIndexes]);
  const checkpointWorkbookIdentity = useMemo(
    () => (checkpointWorkbookPath ? wpsWorkbookIdentity(checkpointWorkbookPath) ?? checkpointWorkbookPath : null),
    [checkpointWorkbookPath]
  );
  /*
   * A1-E1 验收入口的绑定目标：交给 `wpsResultRefFaultTarget` 一处构造，
   * 它读的正是桥建绑定时发出去的那一组同源取值（名字 = `Workbook.Name`，
   * 身份 = `wpsWorkbookBridgeIdentity`）。服务端按这三者全等匹配任务，
   * 任何一项自己拼装都会让"点了按钮却没生效"。
   *
   * 依赖**只**是"当前这本工作簿"（`checkpointWorkbookPath` 是它的稳定身份）：
   * 绑定目标与选中表、与授权版本无关 —— 换表不换簿，目标还是同一个；
   * 换簿（含另存为、未保存→保存）才会变。
   *
   * 这个依赖在函数体里**真的参与判断**，不是"缓存失效开关"：读现场的同时核对
   * "读到的仍是面板刚观察过的那本簿吗"。两者不一致（切簿/另存为的中间态）就不给
   * 目标 —— 宁可这一次入口不可用（按钮置灰），也不把另一本簿的名字与身份交给
   * 服务端，让这次启用悄悄绑到别的任务上去。
   */
  const resultRefFaultTargetOfPanel = useMemo(() => {
    try {
      const application = getWpsApplication();
      // 未保存的簿没有路径（`checkpointWorkbookPath` 为 null）：桥那时发的是会话
      // 内令牌，目标照样要能给，所以这里只做"有路径就核对"，不是拿路径当闸门。
      if (checkpointWorkbookPath && (listWpsWorksheets(application).fullName || null) !== checkpointWorkbookPath) {
        return null;
      }
      return wpsResultRefFaultTarget(application);
    } catch {
      // 读不到现场 ⇒ 没有可绑定的目标；`ResultRefFaultControl` 据此禁用启用按钮。
      return null;
    }
  }, [checkpointWorkbookPath]);
  /*
   * 刷新后只查询**当前会话 + 当前工作簿**里已有操作的状态：不重新执行删除，
   * 也不把别的会话/工作簿的删除卡带进聊天流。
   */
  useEffect(() => {
    if (!checkpointWorkbookIdentity) return;
    void worksheetCheckpoint.hydrate({
      conversationId: chatHistory.activeConversationId,
      workbookIdentity: checkpointWorkbookIdentity
    });
  }, [chatHistory.activeConversationId, checkpointWorkbookIdentity]);
  const deletionLive = useRef({ indexes: selectedSheetIndexes, resolve: resolveSelectedSheetIndexes, scopeSource, authorizationRevision, conversationId: chatHistory.activeConversationId });
  deletionLive.current = { indexes: selectedSheetIndexes, resolve: resolveSelectedSheetIndexes, scopeSource, authorizationRevision, conversationId: chatHistory.activeConversationId };
  /**
   * `requestMessageId` = 这次删除所属的那条请求消息。
   *
   * 结果卡只画在它下面；用户中途切换会话时，回执仍然只更新这一条，
   * 不会变成聊天底部的常驻卡。
   */
  async function executeStructuralDeletion(
    plan: AnalysisPlan,
    indexes: readonly number[],
    requestMessageId: string | null
  ) {
    const issued = deletionLive.current;
    if (!checkpointCapabilities.wpsWrite.enabled || wpsWriteVersionStatus(String(getWpsApplication().Version ?? "")) === "blocked") throw new Error("当前 WPS 版本或写入开关未开放，未删除工作表。");
    if (issued.scopeSource === "folder") throw new Error("工作目录不能使用当前宿主的工作表删除通道。");
    const app = getWpsApplication();
    const before = listWpsWorksheets(app);
    const authorized = before.sheets.filter(sheet => indexes.includes(sheet.index)).map(sheet => sheet.name);
    const issue = worksheetDeletionIssue(plan, authorized, before.sheets.map(sheet => sheet.name));
    if (issue) throw new Error(issue);
    const binding = JSON.stringify({ path: before.fullName, authorized });
    if (!(await targetGuard.check())) throw new TargetConfirmationRequiredError();
    const live = deletionLive.current;
    if (live.scopeSource !== issued.scopeSource || live.authorizationRevision !== issued.authorizationRevision || live.conversationId !== issued.conversationId) throw new Error("授权或会话变化，未删除。");
    const now = listWpsWorksheets(getWpsApplication());
    const currentIndexes = live.resolve?.() ?? live.indexes;
    const current = now.sheets.filter(sheet => currentIndexes.includes(sheet.index)).map(sheet => sheet.name);
    if (binding !== JSON.stringify({ path: now.fullName, authorized: current })) throw new Error("授权或工作簿变化，未删除。");
    if (!now.fullName || !app.Hwnd) throw new Error("当前宿主未提供可验证的原工作簿身份，未删除。");
    const beforeResume = async () => {
      if (!(await targetGuard.check())) throw new TargetConfirmationRequiredError();
      const fresh = deletionLive.current;
      if (fresh.scopeSource !== issued.scopeSource || fresh.authorizationRevision !== issued.authorizationRevision || fresh.conversationId !== issued.conversationId) throw new Error("授权或会话变化，请重新发起；未继续删除。");
      const freshApp = getWpsApplication();
      const freshBook = listWpsWorksheets(freshApp);
      const freshIndexes = fresh.resolve?.() ?? fresh.indexes;
      const freshNames = freshBook.sheets.filter(sheet => freshIndexes.includes(sheet.index)).map(sheet => sheet.name);
      if (freshApp.Hwnd !== app.Hwnd || binding !== JSON.stringify({ path: freshBook.fullName, authorized: freshNames })) throw new Error("授权或工作簿变化，请重新发起；未继续删除。");
    };
    const key = `${issued.conversationId}:${plan.id}`;
    const workbookIdentity = wpsWorkbookIdentity(now.fullName) ?? now.fullName;
    /*
     * 删表成功后按回执核对过的**预期收窄**：把剩余授权范围与目标基线同步过去。
     *
     * 不做这件事，下一轮 `targetGuard.check()` 会把"表少了"读成"用户改了授权"，
     * 弹出「工作表范围发生变化 / 尚未写入」，与刚完成的结果直接冲突。
     * 但"预期"必须被证明：共享层核对同一工作簿身份、回执剩余表与真实目录逐名相同、
     * 删除清单恰好等于授权表差集；随后 `bindVerifiedShrink` 再独立卡一遍"只减不增"。
     * 任何一条不成立就不动基线 —— 守卫仍会照常拦截外部改动与换簿。
     *
     * 为什么没有 `onAuthorizeTarget` 就直接放弃：面板的 `inspect` 读的是受控属性
     * `selectedSheetIndexes`。只推进守卫基线、却不把新范围写回读取方，等于让"基线"
     * 与"读数"从此各说一套 —— 下一轮照样判范围变化。宁可原样保留（用户照常被问一次），
     * 也不做一半的同步。
     */
    const adoptVerifiedShrink = async (): Promise<void> => {
      if (!onAuthorizeTarget) return;
      const receipt = worksheetCheckpoint.receiptForMessage(requestMessageId, issued.conversationId);
      const remaining = listWpsWorksheets(getWpsApplication());
      const shrink = reconcileWorksheetDeletionScope({
        beforeIdentity: workbookIdentity,
        afterIdentity: wpsWorkbookIdentity(remaining.fullName) ?? remaining.fullName ?? null,
        authorizedSheets: current,
        receipt,
        actualSheets: remaining.sheets.map(sheet => sheet.name),
        isSystemSheet: isEBSystemSheet
      });
      if (!shrink.ok) return;
      const nextIndexes = remaining.sheets
        .filter(sheet => shrink.scope.includes(sheet.name)).map(sheet => sheet.index);
      if (nextIndexes.length !== shrink.scope.length) return;
      let rebound = false;
      try {
        rebound = targetGuard.bindVerifiedShrink(await inspectWpsAuthorizedTarget(nextIndexes));
      } catch {
        rebound = false;
      }
      // 基线没推进就不同步宿主范围：两边必须指向同一份"剩下的表"。
      if (rebound) onAuthorizeTarget(nextIndexes);
    };
    const onCompleted = () => adoptVerifiedShrink();
    const succeeded = await worksheetCheckpoint.execute(key, {
      workbookPath: now.fullName, hostWindowId: app.Hwnd, authorizedSheets: current, plan
    }, {
      conversationId: issued.conversationId,
      workbookIdentity,
      requestMessageId,
      beforeResume,
      onCompleted
    });
    if (succeeded) await adoptVerifiedShrink();
    return succeeded ? "completed" : worksheetCheckpoint.waitingForConsent(key) ? "awaiting_input" : "failed";
  }

  /**
   * 原地撤销成功后的**预期目录增加**：把授权与目标基线就地认下来。
   *
   * 撤销把表放回工作簿，目录因此比授权时多几张。不做这件事，下一轮
   * `targetGuard.check()` 会把它读成"范围变了"，弹「尚未写入」——与刚撤销的事实冲突。
   * 放过它同样要证据（共享层 `reconcileWorksheetRestoreScope`）：同簿身份、授权过的表
   * 全在、回执里放回的表真的出现、目录新增只来自这次撤销。授权集合本身不变。
   */
  const adoptVerifiedRestore = async (identity: string | null, restored: readonly string[]): Promise<boolean> => {
    if (!identity || !onAuthorizeTarget) return false;
    const live = deletionLive.current;
    const remaining = listWpsWorksheets(getWpsApplication());
    const authorized = remaining.sheets
      .filter(sheet => (live.resolve?.() ?? live.indexes).includes(sheet.index))
      .map(sheet => sheet.name);
    const check = reconcileWorksheetRestoreScope({
      beforeIdentity: identity,
      afterIdentity: wpsWorkbookIdentity(remaining.fullName) ?? remaining.fullName ?? null,
      authorizedSheets: authorized,
      restoredSheets: restored,
      actualSheets: remaining.sheets.map(sheet => sheet.name),
      isSystemSheet: isEBSystemSheet
    });
    if (!check.ok) return false;
    const nextIndexes = remaining.sheets
      .filter(sheet => check.authorized.includes(sheet.name)).map(sheet => sheet.index);
    if (nextIndexes.length !== check.authorized.length) return false;
    let rebound = false;
    try {
      rebound = targetGuard.bindVerifiedRestore(await inspectWpsAuthorizedTarget(nextIndexes), restored);
    } catch {
      rebound = false;
    }
    // 基线没推进就不同步宿主范围：两边必须指向同一份"授权集合"。
    if (rebound) onAuthorizeTarget(nextIndexes);
    return rebound;
  };

  const baseBusy = worksheetCheckpoint.pending || progress !== null || writeProgress !== null || targetGuard.checking;
  // `busy` = 交互锁定：它多含 `phase === 'reviewing'`（结果已就绪、等用户确认/放弃），
  // 用来防重复执行与操作互斥。**不能**拿它当"EB 正在跑" —— 那会让悬浮标在只等用户
  // 查看结果时也显示三点动画。
  const busy = scopeSelectionPending || baseBusy || resultReviewState.pending || resultReviewState.phase === 'reviewing';
  /*
   * 「实际运行状态」与「交互锁定状态」分开（用户口径）：
   *  - 三点动画：模型请求、执行、回读、生成对比、确认收尾、恢复等**真的在途**；
   *  - 锚图标：只是结果已生成、在等用户查看/确认/放弃。
   *
   * `resultReviewState.pending` 已经覆盖审阅通道自己的每个操作（start / confirm /
   * discard / compare / retryMarkerCleanup / settleExternally，见 `resultReview.ts`
   * 的 `finally { pending:false }`），所以这里**不减** `reviewing`、也不加新条件：
   * 减掉 `reviewing` 就是答案 —— `reviewing` 是"结果在那儿、等人"，不是"还在跑"。
   */
  const running = baseBusy || resultReviewState.pending;
  /*
   * 新一轮一开始就收掉上一次暂停请求的结论：那条小字说的是「上一轮停到哪一步」，
   * 留着它会让人以为**这一轮**也已经停了。
   */
  useEffect(() => {
    if (running) {
      pauseRequestGenerationRef.current += 1;
      setPauseRequestState(null);
    }
  }, [running]);

  /**
   * 读取模型目录。
   *
   * 返回 null 表示本次读取失败：调用方必须保留原有选择，不能把失败当成
   * “用户没有配置模型”——那会把一次后台不可用误报成用户从没配过。
   */
  const refreshModelCatalog = useCallback(async () => {
    setCatalogStatus((current) => (current === "ready" ? current : "loading"));
    try {
      const catalog = await listModels();
      setModels(catalog.models);
      setDefaultModelId(catalog.defaultModelId);
      setCatalogStatus("ready");
      return {
        modelOptions: catalog.models,
        defaultModelId: catalog.defaultModelId
      };
    } catch {
      setCatalogStatus("failed");
      return null;
    }
  }, []);

  const modelManagement = useModelManagement({
    refreshServiceHealth: refreshModelCatalog
  });

  /**
   * 界面展示的模型必须与请求实际发送的 modelId 完全一致：
   * 这里和 Excel 端共用同一份选择规则，解析不出有效连接时就是「未配置」，
   * 绝不回退到目录首项，也不再有任何虚拟的基础模式占位项。
   */
  const effectiveModelId = useMemo(
    () => chooseAvailableModel(models, modelManagement.selectedModelId, defaultModelId),
    [models, modelManagement.selectedModelId, defaultModelId]
  );

  const selectedModel = useMemo(
    () => models.find((model) => model.id === effectiveModelId) ?? null,
    [effectiveModelId, models]
  );

  const modelCatalogFailed = catalogStatus === "failed";
  const modelCatalogLoading = catalogStatus === "loading" || catalogStatus === "idle";
  const modelNotConfigured = catalogStatus === "ready" && models.length === 0;
  const brandStatusText = modelCatalogFailed
    ? isRealModelId(modelManagement.selectedModelId)
      ? "等待 Excel Bro 后台"
      : "点击添加模型开始"
    : modelCatalogLoading
      ? "正在读取模型配置"
      : modelNotConfigured
        ? "尚未配置模型"
        : "";

  useEffect(() => {
    let active = true;
    void refreshModelCatalog().then(() => {
      if (!active) return;
    });
    return () => {
      active = false;
      controllerRef.current?.abort();
      planningInspectionRef.current?.abort();
    };
  }, [refreshModelCatalog]);

  useEffect(() => {
    try {
      setUndoRecord(latestWpsUndo(getWpsApplication()));
    } catch {
      setUndoRecord(null);
    }
  }, []);

  useEffect(() => {
    /*
     * CAP-03 重载恢复：从本地发现索引找回上次未完成的 WPS 写入，查服务端账本
     * 真实状态，并对身份与现场做核对。索引里只有 operationId / 身份 / taskId，
     * 完整计划、授权与写前/写后指纹都在服务端账本 —— 本地记录不是权威。
     */
    if (scopeSource !== "workbook" || recoveryActionLock.current) return;
    let active = true;
    const pending = loadPendingWpsExecution();
    if (!pending) return;
    setWpsRecovery({
      operationId: pending.operationId,
      taskId: pending.taskId,
      state: null,
      identity: { kind: "match" },
      verdict: { kind: "insufficient", detail: "正在查询服务端写入现场…" },
      checking: true,
      acting: false
    });
    void (async () => {
      let state: WpsExecutionState | null = null;
      try {
        state = await getWpsExecutionStatus(pending.operationId);
      } catch {
        state = null;
      }
      if (!active || loadPendingWpsExecution()?.operationId !== pending.operationId) return;
      const currentIdentity = (() => {
        try {
          return wpsWorkbookIdentity(wpsReadWorkbook(getWpsApplication())?.FullName);
        } catch {
          return null;
        }
      })();
      const identity = wpsIdentityVerdict({
        currentIdentity,
        recordedIdentity: state?.workbookIdentity ?? ""
      });
      const localInspection = (() => {
        try {
          return state ? inspectWpsRecoveryState(getWpsApplication(), state,
            resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes) : "unknown" as const;
        } catch {
          return "unknown" as const;
        }
      })();
      const verdict = wpsReconcileVerdict({
        state,
        identityMatch: identity.kind === "match",
        localInspection,
        resumeUnavailableReason: wpsResumeUnavailableReason(getWpsApplication(), state, resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes),
        resumeAvailable: canResumeWpsExecution(getWpsApplication(), state, resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)
      });
      setWpsRecovery({
        operationId: pending.operationId,
        taskId: pending.taskId,
        state,
        identity,
        verdict,
        checking: false,
        acting: false
      });
    })();
    return () => {
      active = false;
    };
    // Scope changes only repeat GET + local inspection, never claim/reconcile/write.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [authorizationKey, scopeMode, scopeSource]);

  /**
   * CAP-02 第 5 条：面板刷新后恢复可恢复的规划任务。
   *
   * 与上面 CAP-03 那段是**两条独立通路**，刻意不合并：
   *
   *  - CAP-03 处理的是"**写入**留下了未完成现场"（`pendingWpsExecution` 索引
   *    → 服务端执行账本），它关心单元格有没有被写过；
   *  - 这一段处理的是"**规划**停在等用户决定"（历史里的 `planningPause`
   *    → 服务端规划账本 `turnId`），它**一个单元格都没碰过**。
   *
   * 合并会让"只是规划暂停"也走写入核对，反过来"写入待核对"被当成可续跑的
   * 规划 —— 两类事实混在一张卡上，用户分不清到底写没写。
   *
   * 关键纪律：**重开只核对，不自动续跑**。自动续跑会在用户完全没点过任何
   * 东西的情况下触发一次模型调用；有副作用的事必须由用户点按钮发起。
   */
  useEffect(() => {
    /*
     * 筛选在共享层（`planningPausedEntriesFromHistory`），那是一条确定性规则：
     * 只挑暂停、同 `turnId` 取最后一条。本函数只负责"查回来挂到卡上"。
     */
    const paused = planningPausedEntriesFromHistory(
      messages.filter((message) => !message.planningPause?.resolved),
      (turnId) => {
        const message = [...messages].reverse().find(
          (item) => item.planningPause?.turnId === turnId
        );
        return message?.planningPause ?? null;
      }
    );
    if (paused.length === 0) return;
    // 只核对**最近一条**：同一次重开里把历史所有暂停都查一遍既慢又没意义，
    // 而更早的那些任务的按钮本来就该由"是不是当前活跃暂停"来决定。
    const latest = paused[paused.length - 1];
    let active = true;
    const controller = new AbortController();
    setPlanningPauseReload({
      turnId: latest.turnId,
      checking: true,
      blockedReason: null
    });
    void (async () => {
      try {
        const inspected = await reinspectWpsPlanningPause(
          { turnId: latest.turnId, pausedAt: latest.payload as never },
          controller.signal
        );
        if (!active) return;
        /*
         * 只有**服务端确认还能精确续跑**时才把它接成当前活跃暂停。
         *
         * 两个判据必须分开看，不能合并成一个条件：
         *  - `canContinue`：服务端是否放行（允许继续）；
         *  - `action`：放行时走哪条通路。
         *
         * 合并成 `canContinue && action === "continue_checkpoint"` 的话，"
         * 放行但不是精确续跑通路"会掉进 else，被报成"该任务当前不能精确续跑" ——
         * 那是替服务端下一个它没下的结论，面板手里并没有这个证据。
         * 而"不放行"时也不能重建活跃暂停：重建会让卡片重新长出按钮，
         * 按钮点下去服务端会直接拒绝。
         */
        const canResumeExactly = inspected.view.canContinue === true &&
          inspected.view.action === "continue_checkpoint";
        if (canResumeExactly) {
          setPlanningPause({
            turnId: latest.turnId,
            payload: inspected.payload,
            view: inspected.view,
            // 原需求在历史里没存（暂停卡只留定位与归因），续跑本身也不需要它 ——
            // 继续只带身份与游标。这里如实留空，不编一句"原需求"出来。
            originalPrompt: "",
            sourceMessageId: latest.messageId
          });
          streamCursorRef.current = { turnId: latest.turnId,
            sequence: Number.isSafeInteger(latest.payload.clientLastSequence) && latest.payload.clientLastSequence! >= 0
              ? latest.payload.clientLastSequence! : 0 };
          setPlanningPauseReload({
            turnId: latest.turnId,
            checking: false,
            blockedReason: null
          });
          return;
        }
        setPlanningPause(null);
        pendingInspectionMessageRef.current = latest.messageId;
        setPlanningPauseReload({
          turnId: latest.turnId,
          checking: false,
          /*
           * 不放行时如实转述服务端的拒绝理由；放行但只能重新规划时**不编理由** ——
           * 卡片会落到"已被新的任务取代"之外的中性说明，不冒充服务端的结论。
           */
          blockedReason: inspected.view.canContinue ? null : inspected.view.reason,
          canRecheck: inspected.status.status === "running" ||
            (inspected.status.status === "completed" && latest.payload.awaitingStatus === true)
        });
      } catch {
        // 查不到（服务重启 / 这条任务的服务端记录已不在）时保持"核对中"的
        // 保守结论：既不说"能继续"，也不说"任务没了" —— 没有证据的事不下结论。
        if (!active) return;
        setPlanningPause(null);
        pendingInspectionMessageRef.current = latest.messageId;
        setPlanningPauseReload({
          turnId: latest.turnId,
          checking: false,
          blockedReason: "暂时读不到这条任务的真实进度，请确认本地 AI 服务仍在运行后重试。",
          canRecheck: true
        });
      }
    })();
    return () => {
      active = false;
      controller.abort();
    };
    // 只在挂载时核对一次：`messages` 之后的变化都由常规轮次处理，重跑会
    // 把用户刚刚点出来的活跃暂停又打回"核对中"。
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /**
   * 重开后重算恢复结论：拿到了新的服务端状态就重新核对一次。
   *
   * 与初始化那次走**同一条** `wpsReconcileVerdict`，不另写一套宽松规则 ——
   * 补记/核对之后再判断"接下来还能做什么"，必须和第一次一样严。
   */
  const refreshWpsRecoveryVerdict = (
    operationId: string,
    taskId: string,
    state: WpsExecutionState | null
  ) => {
    const currentIdentity = (() => {
      try {
        return wpsWorkbookIdentity(wpsReadWorkbook(getWpsApplication())?.FullName);
      } catch {
        return null;
      }
    })();
    const identity = wpsIdentityVerdict({
      currentIdentity,
      recordedIdentity: state?.workbookIdentity ?? ""
    });
    const localInspection = (() => {
      try {
        return state ? inspectWpsRecoveryState(getWpsApplication(), state,
          resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes) : "unknown" as const;
      } catch {
        return "unknown" as const;
      }
    })();
    setWpsRecovery({
      operationId,
      taskId,
      state,
      identity,
      verdict: wpsReconcileVerdict({
        state,
        identityMatch: identity.kind === "match",
        localInspection,
        resumeUnavailableReason: wpsResumeUnavailableReason(getWpsApplication(), state, resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes),
        resumeAvailable: canResumeWpsExecution(getWpsApplication(), state, resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)
      }),
      checking: false,
      acting: false
    });
  };

  /**
   * CAP-03 恢复动作。
   *
   * 按钮只可能来自 `wpsReconcileActions`（判定给什么才画什么），这里再复核一次
   * 才执行 —— 不画"点了没反应"的按钮，也不拿「重新生成预览」冒充断点续跑：
   * `replan` 走重新读取和规划，使用当前授权，通过校验后执行，再审阅；生成新计划、
   * 新执行身份，与这次未完成的执行不是同一件事。
   */
  const refreshWpsRecovery = async () => {
    const current = wpsRecovery;
    if (!current || busy || current.acting || current.checking || recoveryActionLock.current) return;
    recoveryActionLock.current = true;
    setWpsRecovery({ ...current, acting: true });
    try {
      refreshWpsRecoveryVerdict(current.operationId, current.taskId,
        await getWpsExecutionStatus(current.operationId));
    } catch (error) {
      append({ role: "system", text: `恢复状态查询失败；未进行写入。${turnErrorMessage(error)}` });
      setWpsRecovery((previous) => previous ? { ...previous, acting: false } : previous);
    } finally { recoveryActionLock.current = false; }
  };

  const runWpsRecoveryAction = async (action: WpsReconcileAction) => {
    let current = wpsRecovery;
    if (!current || current.checking || current.acting || busy || !current.state || recoveryActionLock.current) return;
    if (action === "replan" && !recoveryScopeConfirmed) return;
    if (!wpsReconcileActions(current.verdict).includes(action)) return;
    const operationId = current.operationId;
    let revision = current.state.revision;
    recoveryActionLock.current = true;
    try {
      // Lock before the first await: two clicks can otherwise pass the same render's guard.
      if (action !== "dismiss" && !(await targetGuard.check())) return;
      setWpsRecovery({ ...current, acting: true });
      if (action === "dismiss") {
        clearPendingWpsExecution(operationId);
        setWpsRecovery(null);
        return;
      }
      // A displayed badge is not authority. Requery AND reread on each click.
      const fresh = await getWpsExecutionStatus(operationId);
      const identity = wpsIdentityVerdict({
        currentIdentity: wpsWorkbookIdentity(wpsReadWorkbook(getWpsApplication())?.FullName),
        recordedIdentity: fresh.workbookIdentity
      });
      const verdict = wpsReconcileVerdict({ state: fresh,
        resumeAvailable: canResumeWpsExecution(getWpsApplication(), fresh, resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes), identityMatch: identity.kind === "match",
        localInspection: inspectWpsRecoveryState(getWpsApplication(), fresh,
          resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes) });
      if (!wpsReconcileActions(verdict).includes(action)) {
        refreshWpsRecoveryVerdict(operationId, current.taskId, fresh);
        return;
      }
      current = { ...current, state: fresh, identity, verdict };
      revision = fresh.revision;
      /*
       * 补记完成：服务端已把全部动作记为"回读一致后完成"，缺的只是最终回执。
       * 这里**一次宿主写入都不做** —— 先领取执行权（旧 token 同时作废），
       * 再用新 revision + 新 token 提交完成回执。领取失败就不碰账本。
       */
      const completedRecord = action === "backfill" ? wpsUndoRecordFor(getWpsApplication(), operationId) : null;
      if (action === "resume" || (action === "backfill" && completedRecord?.state === "verified" &&
          completedRecord.actions.every(item => item.type === "writeValues"))) {
        const application = getWpsApplication();
        const indexes = () => resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes;
        const {prepared, record: prior} = action === "resume"
          ? prepareWpsResume(application, fresh, indexes()) : prepareWpsCompletedReview(application, fresh, indexes());
        const revisionAtStart = authorizationRevision;
        const isCurrent = () => resultFacts.current.authorizationRevision === revisionAtStart &&
          String(wpsReadWorkbook(getWpsApplication())?.FullName ?? "") === prepared.authorization.workbookFullName;
        let formula: string | null = null;
        releaseWpsReviewTarget();
        setResultTitle(prepared.plan.title);
        await resultReview.start(`resume:${operationId}:${fresh.revision}`, {
          isCurrent,
          prepare: async () => ({prepared, workbookFullName: prepared.authorization.workbookFullName!,
            values: prior.actions.map(a => ({values: a.beforeValues, formulas: a.beforeFormulas, valuesCaptured: true})),
            formats: prior.actions.map(() => undefined)}),
          execute: async () => {
            let record = prior;
            if (action === "resume") {
              record = await resumeWpsExecution(application, fresh, indexes, update => setWriteProgress(update.message));
            } else {
              const claimed = await claimWpsExecutionResume(operationId, fresh.revision);
              if (!claimed.ownerToken) throw new Error("没有取得补记执行权。");
              try {
                if (!isCurrent() || claimed.operationId !== operationId || claimed.taskId !== fresh.taskId) throw new Error("补记时任务或授权身份已变化。");
                prepareWpsCompletedReview(application, claimed, indexes());
                await finishWpsExecution(operationId, {expectedRevision: claimed.revision,
                  ownerToken: claimed.ownerToken, result: {backfilled: true, actionCount: claimed.totalActions}});
              } catch (error) {
                try { await suspendWpsExecution(operationId, {expectedRevision: claimed.revision,
                  ownerToken: claimed.ownerToken, code: "WPS_BACKFILL_FAILED", message: "补记未完成，请重新核对。"}); } catch { /* retain discovery */ }
                throw error;
              }
            }
            setWriteProgress(null);
            setUndoRecord(record);
            clearPendingWpsExecution(operationId);
            setWpsRecovery(null);
            append({role: "system", text: action === "resume" ? "已按原执行身份续跑剩余动作并回读验证，已完成动作未重复写入。请审阅主表结果。" : "已补记完成回执，未重写原表。请审阅主表结果。"});
            return {record, executionId: record.executionId, verified: true, undoable: true};
          },
          // 公式在第一次 Add 之前就到手：标记部分失败时双视图仍能按它清本次残留。
          mark: async () => { const marks = markWpsResult(application, prepared, value => { formula = value; }); return marks; },
          clearMarker: async marks => { clearWpsResultMarks(application, marks); },
          showBefore: async before => {
            const staged = stageWpsBeforeComparison(application, before, formula ?? undefined);
            bindWpsReviewTarget({original: staged.original,
              originalIdentity: identityOf(staged.original) ?? {name: staged.plan.authorization.workbookName,
                fullName: String(staged.plan.authorization.workbookFullName ?? ""), path: staged.plan.authorization.workbookPath ?? null},
              copyIdentity: {name: staged.copyIdentity.name, fullName: String(staged.copyIdentity.fullName ?? ""), path: staged.copyIdentity.path ?? null},
              writeSheet: {index: staged.writeSheet.index, name: staged.writeSheet.name}});
            if (rearrangeWpsComparison(application, staged) !== "arranged") throw new Error("双视图未能排列，请手动查看执行前副本。");
            return staged;
          },
          comparisonNotice: copy => {
            const marks = copy.reviewMarks;
            if (!marks || marks.unconfirmed === 0) return null;
            return `执行前副本的审阅标记未能核实清理（${marks.unconfirmed} 处区域，可能读不到或删不掉）；两张表可能显示相同底色，请以原表与回读结果为准。`;
          },
          closeBefore: async copy => {
            const outcome = discardWpsNativeComparison(copy);
            if (outcome === "closed" || outcome === "already_closed") { releaseWpsReviewTarget(); return null; }
            return "执行前副本仍保留，请手动核对。";
          },
          restore: async receipt => {
            if (!await resultUndoRef.current(receipt.executionId, "discard")) throw new Error("恢复未完成，未声明已放弃。");
          },
          validateResult: async receipt => {
            if (inspectWpsUndoRecord(application, receipt.executionId).kind !== "completed") throw new Error("结果已变化，未确认变更。");
          },
          accepted: receipt => { verifiedPlanForSaveRef.current = {plan: prepared.plan,
            sourceSheetNames: prepared.authorization.sheets.map(sheet => sheet.name), executionId: receipt.executionId}; }
        });
        setWriteProgress(null);
        // ResultReview captures errors in its notice; retain recovery discovery and allow a fresh inspection.
        setWpsRecovery(previous => previous ? {...previous, acting: false} : previous);
        if (resultReview.getSnapshot().phase !== "reviewing") {
          refreshWpsRecoveryVerdict(operationId, fresh.taskId, await getWpsExecutionStatus(operationId));
        }
        return;
      }
      if (action === "backfill") {
        const claimed = await claimWpsExecutionResume(operationId, revision);
        if (!claimed.ownerToken) {
          append({
            role: "system",
            text: "服务端没有授予补记这次执行的权限，账本未改动；请重开面板后重新核对。"
          });
          refreshWpsRecoveryVerdict(operationId, current.taskId, claimed);
          return;
        }
        if (inspectWpsRecoveryState(getWpsApplication(), fresh,
          resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes) !== "completed") {
          await suspendWpsExecution(operationId, { expectedRevision: claimed.revision,
            ownerToken: claimed.ownerToken, code: "WPS_RECOVERY_CONFLICT", message: "领取后现场变化。" });
          throw new Error("领取执行权后现场发生变化，未补记完成。");
        }
        const finished = await finishWpsExecution(operationId, {
          expectedRevision: claimed.revision,
          ownerToken: claimed.ownerToken,
          result: { backfilled: true, actionCount: claimed.totalActions }
        });
        clearPendingWpsExecution(operationId);
        setWpsRecovery(null);
        append({
          role: "system",
          text:
            `已补记这次写入的完成回执（${finished.totalActions} 个动作）。` +
            "补记只改账本状态，原表未再做任何写入。"
        });
        return;
      }
      /*
       * 核对在途动作：结论来自授权目标的实际回读（`inspectWpsRecoveryState`），
       * 只有当现场与写后一致 / 仍是写前这两类**正面证据**时才提交。
       * 冲突与未知在判定层就已经拦下，这里不再补送任何回执。
       */
      if (action === "reconcile") {
        if (current.verdict.kind !== "reconcile") return;
        const reconciled = await reconcileWpsActionReceipt(operationId, {
          expectedRevision: revision,
          actionIndex: current.verdict.actionIndex,
          outcome: current.verdict.outcome,
          ...(current.verdict.outcome === "applied" ? { checkpoint: {
            sheetName: fresh.descriptor.actions[current.verdict.actionIndex].sheet,
            range: fresh.descriptor.actions[current.verdict.actionIndex].range,
            type: fresh.descriptor.actions[current.verdict.actionIndex].type,
            beforeFingerprint: fresh.descriptor.actions[current.verdict.actionIndex].previewFingerprint,
            afterFingerprint: fresh.descriptor.actions[current.verdict.actionIndex].expectedFingerprint
          }} : {})
        });
        // A final action receipt was lost: finish the ledger with the newly
        // granted token, never replay any host action.
        if (current.verdict.outcome === "applied" && reconciled.nextActionIndex === reconciled.totalActions &&
            reconciled.ownerToken && inspectWpsRecoveryState(getWpsApplication(), reconciled,
              resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes) === "completed") {
          await finishWpsExecution(operationId, { expectedRevision: reconciled.revision,
            ownerToken: reconciled.ownerToken, result: { backfilled: true, actionCount: reconciled.totalActions } });
          clearPendingWpsExecution(operationId);
          setWpsRecovery(null);
          append({ role: "system", text: "已核对最后一个动作并补记完成；原表未重复写入。" });
          return;
        }
        if (reconciled.ownerToken) await suspendWpsExecution(operationId, {
          expectedRevision: reconciled.revision, ownerToken: reconciled.ownerToken,
          code: "WPS_RECOVERY_RECHECK_REQUIRED", message: "核对完成，剩余动作等待来源与本机材料复核。"
        });
        const suspended = await getWpsExecutionStatus(operationId);
        refreshWpsRecoveryVerdict(operationId, current.taskId, suspended);
        append({
          role: "system",
          text:
            current.verdict.outcome === "applied"
              ? `已把第 ${current.verdict.actionIndex + 1} 个动作补记为已写入（现场与写后一致），未重复写入。`
              : `已把第 ${current.verdict.actionIndex + 1} 个动作补记为未写入（现场仍是写前）。`
        });
        return;
      }
      /*
       * 本机材料不满足续跑条件时，重新读取并规划，生成新的执行身份。
       * 新计划仍先执行后审阅，不能把它称为原动作的精确续跑。
       */
      const taskId = current.taskId;
      const record = messages
        .map((message) => (message.taskResume ? sanitizeTaskResumeRecord(message.taskResume) : null))
        .find((item) => item?.taskId === taskId);
      append({
        role: "system",
        text:
          "这次写入还有未完成的动作，本机材料不满足精确续跑条件。" +
          "下面按原需求重新读取和规划，在当前授权范围内通过校验后执行，再审阅结果；" +
          "不是从上次进度接着写。"
      });
      if (!record) {
        append({
          role: "system",
          text: "没有找到这次任务的原始需求记录，请重新输入需求，通过校验后执行，再审阅结果。"
        });
        return;
      }
      // Fence the previous owner before generating a NEW preview. This does not
      // resume its plan, and no cells are written by this recovery handler.
      const claimed = await claimWpsExecutionResume(operationId, revision);
      if (!claimed.ownerToken) throw new Error("未取得旧执行的处置权，不能重新规划。");
      await suspendWpsExecution(operationId, {
        expectedRevision: claimed.revision, ownerToken: claimed.ownerToken,
        code: "WPS_REPLAN_REQUESTED", message: "用户选择重新读取并规划。"
      });
      // 沿用原 taskId 与已确认信息：重新规划仍是同一件事，用户不必重答。
      resumeCarryRef.current = {
        taskId: record.taskId,
        confirmations: record.confirmations
      };
      await runTurn(resumePrompt(record), { showUser: false });
      // Keep the durable discovery index until a new preview is confirmed or the
      // user explicitly dismisses it; a failed planner must not erase evidence.
      // A newly registered execution supersedes discovery of the old attempt.
      // Do not revive its card after the new plan has already executed/reached review.
      if (loadPendingWpsExecution()?.operationId === operationId) {
        const state = await getWpsExecutionStatus(operationId);
        if (loadPendingWpsExecution()?.operationId === operationId)
          refreshWpsRecoveryVerdict(operationId, current.taskId, state);
      } else {
        setWpsRecovery(previous => previous?.operationId === operationId ? null : previous);
      }
    } catch (error) {
      append({
        role: "system",
        text: `恢复动作没有完成；请核对执行记录与原表现场，不会自动重试。${turnErrorMessage(error)}`
      });
      setWpsRecovery((previous) =>
        previous && previous.operationId === operationId
          ? { ...previous, acting: false }
          : previous
      );
    } finally {
      recoveryActionLock.current = false;
    }
  };

  useEffect(() => {
    /*
     * 授权版本只认**已提交**的范围。WpsProbePanel 里勾选只改草稿，
     * 直到用户点「确认范围」才把新的已提交范围传下来 —— 所以"取消勾选又选回来"
     * 不会疯狂递增版本，也不会顺手作废用户正要回答的那张澄清卡。
     */
    const nextKey = authorizationKeyOf(selectedSheetIndexes);
    if (authorizationRevisionBaseRef.current !== nextKey) {
      authorizationRevisionBaseRef.current = nextKey;
      setAuthorizationRevision((current) => current + 1);
    }
  }, [selectedSheetIndexes]);

  useEffect(() => {
    if (authorizationKeyRef.current === authorizationKey) return;
    if (stagedReview) {
      // 授权范围变了 ⇒ 旧的审阅作废。副本已确认收尾才继续；否则保留现场让用户决定。
      if (!wpsCopyRetired(wpsCopyCleanupState("discarded", discardWpsNativeComparison(stagedReview)))) return;
      setStagedReview(null);
    }
    authorizationKeyRef.current = authorizationKey;
    setPendingPlan(null);
    setPendingPlanMessageId(null);
    setChatHistory((current) => ({
      ...current,
      conversations: current.conversations.map((conversation) =>
        conversation.id === current.activeConversationId
          ? { ...conversation, messages: invalidatePendingClarifications(conversation.messages) }
          : conversation
      )
    }));
    setClarification(null);
  }, [authorizationKey, stagedReview]);

  // 目录成功加载后才迁移旧记录：本机存过已被移除的 local、或指向已删除连接的
  // 选择，会在这里被换成当前有效模型；目录读取失败时不迁移，以免依据空列表
  // 删掉仍然有效的选择。
  useEffect(() => {
    if (catalogStatus !== "ready") return;
    modelManagement.migrateModelSelection({
      modelOptions: models,
      defaultModelId
    });
  }, [
    catalogStatus,
    models,
    defaultModelId,
    modelManagement.migrateModelSelection
  ]);

  useEffect(() => {
    if (!newConversationNotice) return;
    const timer = window.setTimeout(() => setNewConversationNotice(null), 1500);
    return () => window.clearTimeout(timer);
  }, [newConversationNotice]);

  useEffect(() => {
    if (!copiedMessageId && !copyError) return;
    const timer = window.setTimeout(() => {
      setCopiedMessageId(null);
      setCopyError(false);
    }, 1600);
    return () => window.clearTimeout(timer);
  }, [copiedMessageId, copyError]);

  /**
   * 落盘失败提示的展示状态。与 Excel 端一致：失败时明确告知用户，不静默降级。
   *
   * 用 state 而不是直接 append 系统消息，理由有两条：
   * 1. 落盘失败本身也是"这条历史不会被保留"，把它写进历史会在下一次落盘时
   *    又被写丢，自相矛盾；提示应当独立于可持久化的内容存在。
   * 2. 避免"提示写进历史 → 触发落盘 → 又失败 → 再提示"的循环。
   */
  const [historyNotice, setHistoryNotice] = useState<string | null>(null);
  const storageFailedRef = useRef(false);

  useEffect(() => {
    try {
      // 落盘前统一脱敏：对话文本可能来自服务端错误、模型回答或用户粘贴，
      // 任何一条都不该把凭据带进 localStorage。与 Excel 端同一实现。
      // `legacyCredentialConversationIds` 一并写下去：脱敏之后就只能靠这个标记
      // 判断哪些会话是旧记录，靠它重开后发送闸才仍然有效。
      localStorage.setItem(WPS_CHAT_STORAGE_KEY, serializeHistoryForStorage(chatHistory));
    } catch {
      // 宿主写入失败：内存里的历史在读取路径已经脱敏（原文此时已不存在），
      // 所以这里要做的不是"再脱一遍原文"，而是明确告知用户本次会话不会被保留。
      // 关键：**不把错误对象原文写进提示或日志** —— 异常里可能带着写入失败
      // 的上下文（键名、配额乃至被拒绝的值），一律折叠成固定文案。
      //
      // 提示只提示一次；发送闸不受影响（判据是内存里的 `legacyCredentialConversationIds`，
      // 与是否写进存储无关），所以"落盘失败"不会让旧记录被误当成正常历史续聊。
      if (!storageFailedRef.current) {
        storageFailedRef.current = true;
        setHistoryNotice(
          "无法把对话历史写入本机存储（宿主可能禁止了本地存储或空间已满）。" +
            "本次会话仍可继续，但关掉任务窗格后不会被保留。"
        );
      }
      return;
    }
    storageFailedRef.current = false;
    setHistoryNotice(null);
  }, [chatHistory]);

  useEffect(() => {
    try {
      localStorage.setItem(
        PET_VISIBILITY_STORAGE_KEY,
        petVisible ? "visible" : "hidden"
      );
    } catch {
      // 存储不可用时，本次会话仍保留当前选择。
    }
  }, [petVisible]);

  /**
   * 更新当前对话的消息。
   *
   * `clearLegacyMark` 决定这次写入是否算"用户明确放弃旧记录拦截"。
   *
   * **默认必须为 false。** 起初这里无条件清标记，结果是：用户对一段旧记录
   * 提问 → 发送闸拦下并追加一条系统提示 → 那条提示本身把标记清掉了 →
   * 第二次提问就畅通无阻。等于"拦一次就失效"，发送闸形同虚设。
   *
   * 只有真正把用户的新提问（`runTurn` 的 `append({role:"user"})`）写进历史时，
   * 才显式传 true —— 那一刻才是"这段历史已经翻篇，后续是新内容"。
   */
  const updateMessages = useCallback(
    (
      updater: (current: ChatMessage[]) => ChatMessage[],
      clearLegacyMark = false
    ) => {
      setChatHistory((current) => {
        const now = new Date().toISOString();
        const legacyIds = (current as WpsChatHistory).legacyCredentialConversationIds;
        return {
          ...current,
          legacyCredentialConversationIds: clearLegacyMark
            ? (legacyIds ?? []).filter(
                (id) => id !== current.activeConversationId
              )
            : legacyIds,
          conversations: current.conversations.map((conversation) => {
            if (conversation.id !== current.activeConversationId) return conversation;
            const nextMessages = updater(conversation.messages);
            return {
              ...conversation,
              title: conversationTitle(nextMessages),
              messages: nextMessages,
              updatedAt: now
            };
          })
        };
      });
    },
    []
  );

  const append = (message: Omit<ChatMessage, "id" | "createdAt">): string => {
    const id = messageId();
    updateMessages((current) => [
      ...current,
      { id, createdAt: new Date().toISOString(), ...message,
        hasImageAttachments: message.hasImageAttachments ||
          Boolean(message.taskResume && taskImagesRef.current.get(message.taskResume.taskId)?.length) }
    ]);
    return id;
  };

  const beginActivity = (step: WpsActivityUpdate) => {
    const startedAt = Date.now();
    const id = append({ role: "assistant", text: "", activityLog: startWpsActivity(step, startedAt) });
    setActiveActivityMessageId(id);
    return { id, startedAt };
  };

  const updateActivity = (activity: { id: string; startedAt: number }, step: WpsActivityUpdate) => {
    updateMessages((current) => updateWpsActivityMessage(current, activity.id,
      (log) => advanceWpsActivity(log, step, activity.startedAt), step.operationId));
  };

  /**
   * 「继续」永远开一个**新的回复段**（与 Excel 端同一口径）。
   *
   * 以前是把进度接回原来那张卡：用户按了继续之后，最新消息下面没有任何反馈，
   * 进度出现在上方一条更早的回复里。现在换成"同一任务、多段回复"——
   * 旧段保留它停止时的状态并给出"查看后续进展"入口，新段承接后续进度。
   *
   * 执行身份一个都不换：`turnId` 仍是原任务，事件序号仍从原游标往后接。
   * 段链判据（父段是谁、谁降级成 `superseded`、`nextSegmentId` 指向谁）**只**
   * 走 `packages/core/replySegments.ts` —— 两端各写一遍同构但会漂移的链式改写，
   * 正是"第二次继续覆盖第一段后继"那类缺陷的来源。
   */
  const openReplySegmentActivity = (params: {
    /** 同一次继续里重复调用必须用**同一个**键，否则会多开一段。 */
    key: string;
    /** 这一段在展示哪个执行任务；早期暂停没有服务端任务身份，这里给 null。 */
    taskId?: string | null;
    /** 上一段承载的消息。 */
    previous?: ChatMessage;
    step: WpsActivityUpdate;
  }): { id: string; startedAt: number } => {
    const { key, taskId = null, previous, step } = params;
    // 同一次继续里的第二次调用：复用已经开出的那一段，不另开新段。
    // 判据全部来自 ref 自己（`messages` 在这条链上还是渲染期的旧值）。
    const already = resumeSegmentRef.current;
    if (already && already.key === key) {
      setActiveActivityMessageId(already.messageId);
      return { id: already.messageId, startedAt: already.startedAt };
    }
    const startedAt = Date.now();
    const id = append({
      role: "assistant",
      text: "",
      activityLog: startWpsActivity(step, startedAt)
    });
    const parentId = previous?.id ?? null;
    updateMessages((items) => {
      /*
       * 段身份 = 承载它的消息 id：段链与「查看后续进展」都按同一个键对齐。
       * 早期暂停的旧消息可能还没有段身份（比段机制更早落盘），这里只给"这一轮
       * 要被接管的上一段"补一条 `live` —— 给所有带日志的老消息补 `live` 会把它们
       * 集体降级成 `superseded`，历史卡上的「继续」入口就全没了。
       */
      const seeded = items.map((item) =>
        parentId && item.id === parentId && !item.replySegment
          ? { ...item, replySegment: { id: item.id, status: "live" as const } }
          : item
      );
      const planned = openReplySegment(
        seeded.flatMap((item) => (item.replySegment ? [item.replySegment] : [])),
        {
          id,
          taskId,
          continuesTask: true,
          previousSegmentId: parentId
        }
      );
      const byId = new Map(planned.map((segment) => [segment.id, segment]));
      return seeded.map((item) => {
        const segment = byId.get(item.replySegment?.id ?? item.id);
        if (!segment) return item;
        const supersededNow =
          segment.status === "superseded" &&
          item.replySegment?.status !== "superseded";
        return {
          ...item,
          replySegment: segment,
          ...(supersededNow &&
          item.activityLog &&
          item.activityLog.outcome === undefined
            ? {
                // 只冻结还在进行中的那一段：已有结论的段不能在这里重算耗时，
                // 否则暂停时长会被并进执行时间。
                activityLog: freezeSegmentLog(
                  item.activityLog,
                  startedAt
                ) as unknown as ActivityLog
              }
            : {})
        };
      });
    });
    setActiveActivityMessageId(id);
    resumeSegmentRef.current = { key, messageId: id, startedAt };
    return { id, startedAt };
  };

  const resumeTaskActivity = (turnId: string, sourceMessageId?: string) => {
    // 同一次继续里的第二次调用：复用已经开出的那一段。
    const already = resumeSegmentRef.current;
    if (already && already.key === `task:${turnId}`) {
      setActiveActivityMessageId(already.messageId);
      return { id: already.messageId, startedAt: already.startedAt };
    }
    // Older histories did not bind the timeline to a task. Only use the log
    // immediately preceding this pause within the same user turn.
    const sourceIndex = messages.findIndex((item) => item.id === sourceMessageId);
    const preceding = sourceIndex < 0 ? [] : messages.slice(0, sourceIndex).reverse();
    const boundary = preceding.findIndex((item) => item.role === "user");
    const legacy = (boundary < 0 ? [] : preceding.slice(0, boundary))
      .find((item) => item.activityLog);
    /*
     * 父段 = 该任务**最新**那一段（共享判据），不是"列表里第一条匹配这个
     * turnId 的消息"。用后者的话第二次继续仍然改写第一段：A 的
     * `nextSegmentId` 被反复覆盖，B 永远没有后继，A→B→C→D 的链断在中间。
     */
    const continuation = planSegmentContinuation(
      messages.flatMap((item) => (item.replySegment ? [item.replySegment] : [])),
      turnId
    );
    const original =
      (continuation.previousSegmentId
        ? messages.find((item) => item.id === continuation.previousSegmentId)
        : undefined) ??
      [...messages]
        .reverse()
        .find((item) => item.activityLog?.operationId === turnId) ??
      (legacy?.activityLog?.operationId ? undefined : legacy);
    if (!original?.activityLog) {
      const begun = beginActivity({
        id: "planning-continue", label: "继续规划（此前未记录步骤）", phase: "planning"
      });
      resumeSegmentRef.current = { key: `task:${turnId}`, messageId: begun.id, startedAt: begun.startedAt };
      return begun;
    }
    return openReplySegmentActivity({
      key: `task:${turnId}`,
      taskId: turnId,
      previous: original,
      step: { id: "planning.attach", label: "接回原任务进度", phase: "planning" }
    });
  };

  /**
   * 「查看后续进展」：只把视口带到接续这一段的那条回复。
   *
   * 不是"把状态搬回去"——目标任务、事件游标、结果引用都不动，历史段保留它
   * 停止时的内容。找不到那条消息时静默不动。
   */
  /**
   * 这一次「继续」已经开出的展示段。
   *
   * 一次继续动作可能分几段代码走：先接回 running 任务、任务结束后再取结果，两条
   * 分支都会要求"把进度落到当前回复"。**它们属于同一次继续**，只能有一段新的
   * 回复 —— 记在这里，第二次直接复用，不再另开一段（否则一次点击会串出两段）。
   * 这一段收尾时清空，下一次继续自然开新段。
   */
  const resumeSegmentRef = useRef<{
    /**
     * 这一次继续的键。规划续跑用 `task:<turnId>`，早期分析（还没有服务端任务）
     * 用 `early:<原消息 id>`。**不伪造 turnId**：没有服务端任务就没有这个键值。
     */
    key: string;
    messageId: string;
    startedAt: number;
  } | null>(null);

  const jumpToMessage = (id: string) => {
    const row = document.querySelector<HTMLElement>(`[data-message-id="${id}"]`);
    row?.scrollIntoView?.({ block: "start", behavior: "smooth" });
  };

  const discardRecoveryRoutingActivity = (activity: { id: string }) => {
    updateMessages((items) => items.filter((item) => item.id !== activity.id));
    setActiveActivityMessageId((current) => current === activity.id ? null : current);
  };

  const endActivity = (activity: { id: string; startedAt: number }, outcome: ActivityOutcome, presentation: { defaultCollapsed?: boolean } = {}) => {
    updateMessages((current) => updateWpsActivityMessage(current, activity.id, (log) => ({
      ...finishWpsActivity(log, outcome, activity.startedAt),
      ...(outcome === "completed" && presentation.defaultCollapsed === true ? { defaultCollapsed: true } : {})
    })));
    setActiveActivityMessageId((current) => current === activity.id ? null : current);
  };

  /**
   * 只关**自己立的**桥。
   *
   * 以前四处 `sdkBridge.end()` 各自判断"这一轮还归我吗"，而 `controllerRef` 会被
   * 状态核对接管：核对一接走，`ownsTurn` 就假，桥永远不收 —— 迟到的读表进度于是
   * 又能把按钮点亮。桥的主人是**立它的那次执行**，与 controller 无关。
   */
  const endBridgeOwnedBy = (owner: string): boolean => {
    if (bridgeOwnerRef.current !== owner) {
      recordPauseDiagnostic({
        kind: "teardown",
        source: "system",
        executionId: owner,
        bridge: bridgeOwnerRef.current ? "other" : "none",
        outcome: "skipped"
      });
      return false;
    }
    sdkBridge.end();
    bridgeOwnerRef.current = null;
    recordPauseDiagnostic({
      kind: "teardown", source: "system", executionId: owner, bridge: "owned", outcome: "ended"
    });
    return true;
  };
  /**
   * 一次执行的收尾：活动、controller、桥**分别**按归属判断。
   *
   * 判据在共享层（`planExecutionTeardown`），两端读同一份；这里只负责把三件事
   * 落到各自的对象上。重复调用是安全的：活动已收、controller 已不是自己、桥已收，
   * 都会得到 `false`，于是什么都不做。
   */
  const settleExecution = (params: {
    owner: string;
    activity: { id: string; startedAt: number } | null;
    activitySettled: boolean;
    /** 活动收尾用哪种结局。 */
    outcome: ActivityOutcome;
  }) => {
    const plan = planExecutionTeardown({
      owner: params.owner,
      controllerOwner: controllerOwnerOf(controllerRef.current),
      bridgeOwner: bridgeOwnerRef.current,
      hasActivity: params.activity !== null,
      activitySettled: params.activitySettled
    });
    if (plan.settleActivity && params.activity) {
      endActivity(params.activity, params.outcome);
      recordPauseDiagnostic({
        kind: "teardown", source: "system", executionId: params.owner, outcome: "settled"
      });
    }
    if (plan.clearController) {
      controllerRef.current = null;
      setProgress(null);
      recordPauseDiagnostic({
        kind: "teardown", source: "system", executionId: params.owner, controller: "owned", outcome: "cleared"
      });
    } else {
      /*
       * 不是我的 controller：**可能**别人正拿它执行（不许动它的进度），也**可能**
       * 谁都没拿着（`controllerRef` 已空）。后一种情况下本端并没有任何可中止的
       * 执行，进度提示必须收掉 —— 否则界面会一直显示"处理中"，按钮却点不动。
       */
      const holder = controllerOwnerOf(controllerRef.current);
      if (holder === null && controllerRef.current === null) setProgress(null);
      recordPauseDiagnostic({
        kind: "teardown",
        source: "system",
        executionId: params.owner,
        controller: holder ? "stale" : "none",
        outcome: "skipped"
      });
    }
    endBridgeOwnedBy(params.owner);
    return plan;
  };

  const markPlanPreview = (status: "executed" | "cancelled") => {
    if (!pendingPlanMessageId) return;
    updateMessages((current) => current.map((message) =>
      message.id === pendingPlanMessageId && message.planPreview
        ? { ...message, planPreview: { ...message.planPreview, status } }
        : message
    ));
  };

  const completeNewConversation = () => {
    if (stagedReview) {
      const cleanup = wpsCopyCleanupState("discarded", discardWpsNativeComparison(stagedReview));
      if (!wpsCopyRetired(cleanup)) {
        append({ role: "system", text: `请先处理右侧审阅副本，再新建对话；原表尚未写入。${wpsCopyCleanupText(cleanup)}` });
        return;
      }
      setStagedReview(null);
    }
    planningInspectionRef.current?.abort();
    planningInspectionRef.current = null;
    pendingInspectionMessageRef.current = null;
    setPlanningPause(null);
    setPlanningPauseReload(null);
    pauseCarryRef.current = null;
    streamCursorRef.current = null;
    planningTargetRef.current = null;
    const previousSaved = messages.some((message) => message.role === "user");
    if (previousSaved) {
      const conversation = createWpsConversation();
      setChatHistory((current) => ({
        activeConversationId: conversation.id,
        conversations: [
          conversation,
          ...current.conversations.map((item) => item.id === current.activeConversationId
            ? { ...item, messages: invalidatePendingClarifications(item.messages) }
            : item)
        ].slice(0, 30),
        // 标记跟着会话走：只保留切到最后 30 条里仍然存在的那些。
        legacyCredentialConversationIds: (
          current.legacyCredentialConversationIds ?? []
        ).filter((id) =>
          current.conversations.slice(0, 29).some((item) => item.id === id)
        )
      }));
    }
    setPrompt("");
    setProgress(null);
    setActiveActivityMessageId(null);
    setClarification(null);
    setPendingPlan(null);
    setPendingPlanMessageId(null);
    setCopiedMessageId(null);
    setHistoryOpen(false);
    setToolsOpen(false);
    setBackupOpen(false);
    setMoreMenuOpen(false);
    setModelMenuOpen(false);
    setNewConversationConfirmOpen(false);
    clearImages();
    taskImagesRef.current.clear();
    setNewConversationNotice({ id: Date.now(), previousSaved });
    window.requestAnimationFrame(() => composerInputRef.current?.focus());
  };

  const startNewConversation = () => {
    if (busy) return;
    if (pendingPlan) {
      setHistoryOpen(false);
      setToolsOpen(false);
      setBackupOpen(false);
      setMoreMenuOpen(false);
      setModelMenuOpen(false);
      setNewConversationConfirmOpen(true);
      return;
    }
    completeNewConversation();
  };

  const openConversation = (conversationId: string) => {
    if (stagedReview) {
      append({ role: "system", text: "请先保存或放弃右侧审阅副本，再切换对话。" });
      return;
    }
    planningInspectionRef.current?.abort();
    planningInspectionRef.current = null;
    pendingInspectionMessageRef.current = null;
    setPlanningPause(null);
    setPlanningPauseReload(null);
    pauseCarryRef.current = null;
    streamCursorRef.current = null;
    planningTargetRef.current = null;
    clearImages();
    taskImagesRef.current.clear();
    setChatHistory((current) =>
      current.conversations.some((conversation) => conversation.id === conversationId)
        ? {
            ...current,
            activeConversationId: conversationId,
            conversations: current.conversations.map((conversation) =>
              conversation.id === current.activeConversationId
                ? { ...conversation, messages: invalidatePendingClarifications(conversation.messages) }
                : conversation
            )
          }
        : current
    );
    setPrompt("");
    setClarification(null);
    setClarificationDraft("");
    setPendingPlan(null);
    setPendingPlanMessageId(null);
    setCopiedMessageId(null);
    setHistoryOpen(false);
  };

  const deleteConversation = (conversationId: string) => {
    if (stagedReview) {
      append({ role: "system", text: "请先保存或放弃右侧审阅副本，再删除对话。" });
      return;
    }
    setChatHistory((current) => {
      const remaining = current.conversations.filter(
        (conversation) => conversation.id !== conversationId
      );
      if (remaining.length === current.conversations.length) return current;
      if (remaining.length === 0) {
        const replacement = createWpsConversation();
        return { activeConversationId: replacement.id, conversations: [replacement] };
      }
      return {
        activeConversationId:
          current.activeConversationId === conversationId
            ? remaining[0].id
            : current.activeConversationId,
        conversations: remaining
      };
    });
    setClarification(null);
    setPendingPlan(null);
    setPendingPlanMessageId(null);
  };

  const closeHeaderSurfaces = () => {
    setModelMenuOpen(false);
    setToolsOpen(false);
    setBackupOpen(false);
    setHistoryOpen(false);
    setMoreMenuOpen(false);
  };

  const copyMessage = async (message: ChatMessage) => {
    const copied = await copyTextToClipboard(message.text.trim());
    if (copied) {
      setCopiedMessageId(message.id);
      setCopyError(false);
    } else {
      setCopyError(true);
    }
  };

  const runTurn = async (
    question: string,
    options: { intentConfirmed?: boolean; clarificationRound?: number; showUser?: boolean; targetConfirmedRetry?: boolean; images?: PendingImage[]; messageUnderstanding?: Awaited<ReturnType<typeof understandMessage>>; taskContext?: AgentTaskContext; restartingAnalysis?: boolean; activityMessageId?: string } = {}
  ) => {
    const sentImages = options.images ?? (options.showUser === false && !pendingImages.length
      ? taskImagesRef.current.get(taskIdRef.current) ?? [] : pendingImages);
    const text = question.trim() || (sentImages.length ? "请结合附件图片分析当前工作簿，并说明发现的问题。" : "");
    const originalTarget = targetGuard.binding;
    const originalMode = `${scopeSource}:${scopeMode}`;
    let completedUnderstanding = options.messageUnderstanding;
    /**
     * 失败时把原需求放回输入框，但**不覆盖用户后来打的字**。
     *
     * 一次请求失败可能发生在几秒之后；这段时间里用户完全可能已经开始写下一句。
     * 无条件 `setPrompt(text)` 会把他刚写的东西冲掉 —— 那不是"恢复"，是丢草稿。
     */
    const restorePromptUnlessEdited = (value: string) =>
      setPrompt((current) => (current.trim() === "" ? value : current));
    const rememberEarlyPause = () => {
      const messageId = activity.id;
      prePlanningPause.remember({ messageId, conversationId: chatHistory.activeConversationId,
        target: originalTarget, mode: originalMode,
        data: { prompt: text, images: [...sentImages], taskContext: options.taskContext,
          activityMessageId: activity.id, messageUnderstanding: completedUnderstanding } });
    };
    if (!text || busy || preparingImages) return;
    if (sentImages.length && !selectedModel?.supportsVision) {
      setImageError("当前模型不支持图片，请切换到支持视觉输入的模型后发送。附件已保留。");
      setModelMenuOpen(true);
      return;
    }
    if (pendingPlan) {
      append({ role: "system", text: "请先确认或取消当前写入计划，再开始新的任务。" });
      return;
    }
    // 续聊前的旧记录闸：与 Excel 端同一判据 —— 查的是读取时登记的持久标记
    // （`legacyCredentialConversationIds`），不是"文本里还能不能匹配到凭据"。
    // 这段历史在读取时已被脱敏，真凭据不会进入界面、内存或请求；标记只是为了
    // 让用户明确知道这是一段旧记录、并据此新开一段对话。
    if (staleCredentials) {
      append({
        role: "system",
        text:
          "这段历史里的凭据是在旧版本写进去的，已被拦下，不会随本次提问发送。" +
          "请点上方「新对话」开始一段干净的对话，或先删除这段历史。"
      });
      setPrompt(text);
      return;
    }
    // 没有可用模型时必须明确反馈，并保留用户已经输入的任务文本：
    // 不清空输入框、不伪造一次成功回答，也不擅自改用别的连接。
    if (!effectiveModelId) {
      setModelMenuOpen(true);
      return;
    }
    /**
     * 工作目录范围下的发送闸。
     *
     * 这一轮**没有**把服务端的查询/执行链路接到 WPS 上，所以这里必须明确拒绝，
     * 而不是拿当前打开的工作簿顶上 —— 那会让用户以为"问的是目录里的文件"，
     * 实际读的却是他屏幕上那本，属于用未经授权的对象冒充他要的数据。
     * 判据与文案都在共享层 `workDirectorySendRefusal`（两端同一份声明）。
     * 输入框里的字**原样保留**，用户切个范围就能直接发。
     */
    const folderRefusal = workDirectorySendRefusal({
      sourceMode: scopeSource,
      folderName: folderScope?.folderName,
      sheetCount: folderScope?.sheetCount
    });
    if (folderRefusal) {
      append({ role: "system", text: folderRefusal });
      setPrompt(text);
      return;
    }
    const conversation = messages
      .filter((message) => message.role !== "system" && message.text.trim().length > 0)
      .slice(-6)
      .map((message) => ({ role: message.role as "user" | "assistant", text: message.text.slice(0, 1000) }));
    retainPlanningContext();
    const earlyTasks = prePlanningPause.pendingEntries().filter((entry) => entry.conversationId === chatHistory.activeConversationId);
    // A selected original request is a fresh analysis, not another recovery
    // selection (including other still-paused tasks in this conversation).
    const recoveryTasks = options.restartingAnalysis ? [] : planningRecoveryCandidates(messages.map((message) =>
      message.planningPause && !message.planningPause.resolved &&
      (message.planningPause.awaitingStatus !== false || message.planningPause.canContinue)
        ? message.planningPause.turnId : earlyTasks.find((entry) => entry.messageId === message.id)?.id));
    if (options.showUser !== false && !options.targetConfirmedRetry && !options.restartingAnalysis) {
      updateMessages((current) => [...current, {
        id: messageId(), role: "user", text,
        attachmentNames: sentImages.map((image) => image.name),
        hasImageAttachments: sentImages.length > 0, createdAt: new Date().toISOString()
      }], true);
      /*
       * **本地接受即清空。**
       *
       * 以前这里等异步的"理解消息"跑完才清输入框，于是用户按了发送、消息已经
       * 进了对话流，输入框里却还留着同一句话（看起来像没发出去，有人会再按一次，
       * 于是同一句话发两遍）。清空只需要"这句话已经进了对话流"这一个条件。
       *
       * 原需求不会丢：`text` 一直在本轮的闭包里，失败时按下面的
       * `restorePromptUnlessEdited` 放回去。
       */
      setPrompt("");
      clearImages();
    }
    const retainedActivity = options.activityMessageId && messages.find((item) => item.id === options.activityMessageId);
    /*
     * 早期分析（理解消息 / 读表结构 / 字段意图）的继续：**另开一个新回复段**。
     *
     * 这一步还没有服务端任务身份（没形成规划检查点），所以段上不写 `taskId`、
     * 不伪造 turnId、也不声称 SDK 内部断点续算 —— 它是本机按原需求重跑，归属
     * 就是这条本机消息。把进度接回旧卡正是"用户又发了一句话，下方什么都没有"
     * 的那个缺陷。旧卡保留它暂停那一刻的内容与耗时，并给出「查看后续进展」。
     */
    const activity = retainedActivity && retainedActivity.activityLog
      ? openReplySegmentActivity({
          key: `early:${retainedActivity.id}`,
          previous: retainedActivity,
          step: {
            id: "message-intent",
            label: "继续分析",
            phase: "intent",
            detail: "正在核验当前现场并继续原需求"
          }
        })
      : beginActivity({
      id: "message-intent", label: "理解消息", phase: "intent",
      detail: "先判断是否需要工作簿上下文；尚未读取字段或数据。"
    });
    const messageController = new AbortController();
    /*
     * 这一次发送的执行身份：从"理解消息"一直活到本轮收尾。活动、controller、
     * 桥三样资源的归属都按它算（`planExecutionTeardown`），不按"ref 还指着我吗"。
     */
    const executionOwner = newExecutionId("send");
    tagControllerOwner(messageController, executionOwner);
    controllerRef.current = messageController;
    setProgress({ stage: "intent", message: "正在理解消息…" });
    let messageUnderstanding: Awaited<ReturnType<typeof understandMessage>>;
    let needsContext = false;
    try {
      messageUnderstanding = options.messageUnderstanding ?? await waitForPlanning(understandMessage({
        prompt: text, conversation, imageCount: sentImages.length, modelId: effectiveModelId,
        ...(recoveryTasks.length ? { recoveryTasks } : {})
      }, messageController.signal), messageController.signal);
      if (messageController.signal.aborted) throw new DOMException("Aborted", "AbortError");
      updateActivity(activity, { id: "message-intent", label: "理解消息", phase: "intent", status: "completed" });
      if (messageUnderstanding.kind === "conversation") {
        append({ role: "assistant", text: messageUnderstanding.message });
        setPrompt("");
        clearImages();
        endActivity(activity, "completed", { defaultCollapsed: true });
        return;
      }
      if (messageUnderstanding.kind === "task_recovery") {
        const turnId = messageUnderstanding.turnId;
        if (!sentImages.length && isOfferedPlanningRecovery(turnId, recoveryTasks) && earlyTasks.some((entry) => entry.id === turnId)) {
          discardRecoveryRoutingActivity(activity);
          setPrompt("");
          if (messageUnderstanding.action === "continue") await restartEarlyAnalysis(turnId);
          else append({ role: "assistant", text: "原需求已暂停，尚未写入。可以继续分析。" });
          return;
        }
        const original = messages.find((message) => message.planningPause?.turnId === turnId && !message.planningPause.resolved);
        if (sentImages.length || !isOfferedPlanningRecovery(turnId, recoveryTasks) || !original) {
          throw new Error("无法确定要恢复的原任务；请在原任务卡片上核对状态。");
        }
        setPrompt("");
        discardRecoveryRoutingActivity(activity);
        if (messageUnderstanding.action === "continue") await continuePlanningMessage(original);
        else await inspectPlanningMessage(original);
        return;
      }
      needsContext = true;
      completedUnderstanding = messageUnderstanding;
    } catch (error) {
      endActivity(activity, messageController.signal.aborted && pauseRequestedRef.current === messageController
        ? "awaiting_input" : error instanceof DOMException && error.name === "AbortError" ? "cancelled" : "failed");
      if (messageController.signal.aborted && pauseRequestedRef.current === messageController) {
        setPrompt((current) => current.trim() === text ? "" : current);
        clearImages();
        rememberEarlyPause();
        pauseRequestedRef.current = null;
      } else restorePromptUnlessEdited(text);
      if (!(error instanceof DOMException && error.name === "AbortError")) append({ role: "system", text: turnErrorMessage(error) });
      return;
    } finally {
      if (!needsContext) {
        if (controllerRef.current === messageController) {
          controllerRef.current = null;
          setProgress(null);
        }
      }
    }
    const pauseBeforeContext = () => {
      if (controllerRef.current === messageController) controllerRef.current = null;
      setProgress(null);
    };
    // Only a new workbook task changes the live execution context. Ordinary replies
    // and recovery requests must not discard the original task before classification.
    if (options.showUser !== false && !options.targetConfirmedRetry) {
      pendingInspectionMessageRef.current = null;
      setPlanningPause(null);
      setPlanningPauseReload(null);
      streamCursorRef.current = null;
      planningTargetRef.current = null;
      pauseRequestedRef.current = null;
      localMappingResumeRef.current = null;
      taskIdRef.current = newTaskId();
      turnIdRefForStream.current = null;
      pauseCarryRef.current = null;
      resumeRequestRef.current = null;
      taskImagesRef.current.set(taskIdRef.current, sentImages);
    }
    // The message-only model call grants no data access. Recheck the live scope now,
    // before reading fields, including changes that happened while understanding.
    if (!(await targetGuard.checkForAnalysis(() => retryTurnRef.current(text, {
      ...options, images: sentImages, targetConfirmedRetry: true, messageUnderstanding
    })))) {
      /*
       * 目标需要重新确认：这条需求**还没有被接受**，输入框里那句话说到底还是
       * 用户的草稿。清空过的草稿要放回去（不覆盖用户在此之后又写的内容），
       * 否则界面一边提示"检查数据范围后继续"，一边把待继续的需求藏了起来。
       */
      restorePromptUnlessEdited(text);
      endActivity(activity, "awaiting_input");
      pauseBeforeContext();
      return;
    }
    if (messageController.signal.aborted) {
      if (pauseRequestedRef.current === messageController) {
        pauseRequestedRef.current = null;
        rememberEarlyPause();
        endActivity(activity, "awaiting_input");
      } else endActivity(activity, "cancelled");
      pauseBeforeContext();
      return;
    }
    const turnSheetIndexes = [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)];
    const sheetIdentity = currentWpsSheetIdentity();
    let headerFingerprint: string | undefined;
    if (!turnSheetIndexes.length) {
      endActivity(activity, "awaiting_input");
      append({ role: "system", text: "请先选择需要操作的工作表。" });
      pauseBeforeContext();
      return;
    }
    taskImagesRef.current.set(taskIdRef.current, sentImages);
    setPrompt("");
    clearImages();
    if (clarification && options.showUser !== false) {
      updateMessages((current) => current.map((message) =>
        message.id === clarification.messageId && message.clarification
          ? { ...message, clarification: { ...message.clarification, status: "invalidated" } }
          : message
      ));
    }
    setClarification(null);
    const controller = messageController;
    setProgress({ stage: "reading", message: "消息需要工作簿上下文，正在读取授权表结构…" });
    lastTurnFailedRef.current = false;
    let preparedBridgeTurnId: string | null = null;
    try {
      planningTargetRef.current = captureWpsObjectIdentity(getWpsApplication(), turnSheetIndexes);
      const result = await waitForPlanning(runWpsReadonlyTurn({
        application: getWpsApplication(),
        selectedSheetIndexes: turnSheetIndexes,
        prompt: text,
        images: sentImages,
        taskContext: options.taskContext ?? createAgentTaskContext(
          options.showUser === false ? resumeCarryRef.current?.originalPrompt ?? text : text,
          options.showUser === false ? resumeCarryRef.current?.confirmations ?? [] : []
        ),
        modelId: effectiveModelId || null,
        intentConfirmed: options.intentConfirmed,
        clarificationRound: options.clarificationRound,
        messageUnderstanding,
        conversation,
        signal: controller.signal,
        beforeRead: async () => {
          if (!(await targetGuard.checkForAnalysis())) throw new TargetConfirmationRequiredError();
        },
        preparePlanningBridge: async (turnId) => {
          if (controller.signal.aborted || controllerRef.current !== controller) return false;
          const ready = await ensureWpsBridge(turnId, turnSheetIndexes, controller.signal, true, executionOwner);
          if (!ready || controller.signal.aborted || controllerRef.current !== controller) return false;
          preparedBridgeTurnId = turnId;
          return true;
        },
        onProgress: (update) => {
          if (controller.signal.aborted || controllerRef.current !== controller) return;
          // 意图阶段一旦拿到服务端 taskId 就记下：断线恢复要靠它查真实状态。
          recordPlanningProgress(update);
          setProgress(update);
          updateActivity(activity, readonlyActivityUpdate(update));
          recordPlanningResult(update);
          // The intent task id was bound before sending the planning request.
          // Progress must never resurrect an ended bridge.
          if (!update.bridgeTool) return;
          const turnId = update.turnId ?? turnIdRefForStream.current;
          if (!turnId || turnId !== preparedBridgeTurnId) return;
          if (sdkBridge.binding()?.sourceFingerprint === null) {
            let pinned = false;
            try {
              pinned = sdkBridge.pinSourceFingerprint(turnId, wpsBridgeLiveContextRef.current());
            } catch { /* Do not expose host exception details. */ }
            if (!pinned) throw new Error("无法核对本轮来源，已停止规划；未执行工作簿写入。");
          }
          sdkBridge.handleProgressEvent(update);
        }
      }), controller.signal);
      controller.signal.throwIfAborted();
      headerFingerprint = await applyTurnResult(result, {
        activity,
        text,
        turnSheetIndexes,
        clarificationRound: options.clarificationRound ?? 0,
        images: sentImages
      });
    } catch (error) {
      if (controllerRef.current !== controller) return;
      const manualPause = error instanceof DOMException && error.name === "AbortError" && pauseRequestedRef.current === controller;
      if (sentImages.length && !manualPause) { setPendingImages(sentImages); restorePromptUnlessEdited(text); }
      if (error instanceof TargetConfirmationRequiredError) {
        endActivity(activity, "awaiting_input");
        restorePromptUnlessEdited(text);
      } else if (error instanceof DOMException && error.name === "AbortError") {
        endActivity(activity, manualPause ? "awaiting_input" : "cancelled");
        if (manualPause) {
          // 已提交需求保留在原任务，不重新塞进普通发送入口。
          setPrompt((current) => current.trim() === text ? "" : current);
          clearImages();
          pauseRequestedRef.current = null;
          if (turnIdRefForStream.current) {
            const turnId = turnIdRefForStream.current;
            const cardId = rememberPlanningInterruption(turnId, text, turnSheetIndexes, options.clarificationRound ?? 0);
            // 自动核对一次，只查询/取回原结果；续跑仍由用户单独确认。
            // `attachRunning = false`：暂停不接回进度轮询，只留一条真实状态。
            await reconcileInterruptedTurn(turnId, text, turnSheetIndexes, cardId, false);
          } else {
            rememberEarlyPause();
          }
        } else append({ role: "system", text: "任务已停止；未确认的计划不会修改工作表。" });
      } else {
        endActivity(activity, "failed");
        lastTurnFailedRef.current = true;
        if (apiErrorCategory(error) === "local_service" && options.showUser !== false) restorePromptUnlessEdited(text);
        append({
          role: "system",
          text: turnErrorMessage(error)
        });
        /*
         * **断线恢复（CAP-02 第 4 条）。**
         *
         * 网络错误只说明"这条连接断了"，**不说明任务没跑**。此前这里直接提示
         * 用户重发需求 —— 而重发就是一次**全新的规划**：重复调用模型、重复读取，
         * 而服务端那条任务可能已经跑完、正在跑、或停在可续跑的检查点上。
         *
         * 正确顺序：拿着本轮已经用过的 `turnId` 去查真实状态，按状态给动作：
         *  - 已接收 / 运行中 / 已完成 / 有可续跑现场 → **明确不许重发**，
         *    把现状如实说给用户（已完成就把结果取回来，可续跑就给暂停卡）；
         *  - 只有服务端**明确说没有这条任务**（404）→ 才允许"重发需求"。
         */
        const interruptedTurnId = turnIdRefForStream.current;
        if (
          interruptedTurnId &&
          options.showUser !== false &&
          apiErrorCategory(error) === "local_service"
        ) {
          await reconcileInterruptedTurn(interruptedTurnId, text, turnSheetIndexes);
        }
      }
    } finally {
      /*
       * **只收自己的**。桥按"立桥那次执行"归属，controller 按对象归属。
       *
       * 旧写法把两者塞进同一个 `ownsTurn`：核对接管 controller 之后它就全假，
       * 于是本轮的桥永远不收 —— 迟到的读表进度又能把"可暂停"点亮，而这时本端
       * 已经没有可中止对象。活动在这里不参与（该函数每条出口都已经自己收过）。
       */
      settleExecution({
        owner: executionOwner,
        activity: null,
        activitySettled: true,
        outcome: "awaiting_input"
      });
    }
  };

  /**
   * 网络中断后核对原任务：**先查状态，再决定能不能重发**。
   *
   * 五态处理（与 `planningResumeDecision` 同一份判据，不另立一套）：
   *
   *  1. `completed` 且有结果 → 把结果接回处理链（`applyTurnResult`），
   *     **不重发**。用户要的东西其实已经算好了，只是回程断在半路。
   *  2. `running` → 明说仍在处理中，**不重发**（重发会与在跑的那一轮打架）。
   *  3. 可继续（`waiting_user` / 模型故障暂停）→ 落暂停卡，**不重发**。
   *  4. 有任务但不可继续 → 如实说明，**不重发**（重发也无济于事，且可能重复计费）。
   *  5. 服务端明确 404「没有这条任务」→ **这才允许**提示重新发送。
   *
   * 查询本身失败（后端仍未起）→ 保守处理：既不重发也不谎称任务不存在，
   * 只提示"稍后重试核对"。
   */
  // 中断只是客户端事实，不伪造服务端检查点。先留身份，再核对真实状态。
  const rememberPlanningInterruption = (
    turnId: string, text: string, indexes: readonly number[], round: number,
    sourceMessageId?: string
  ) => {
    pauseCarryRef.current = { turnId, originalPrompt: text, selectedSheetIndexes: [...indexes], clarificationRound: round };
    retainPlanningContext();
    const metadata: NonNullable<ChatMessage["planningPause"]> = {
      turnId, pauseType: "interrupted", reason: "已暂停接收进度，正在核对原任务；未写入工作表。",
      canContinue: false, action: "none", actionLabel: "核对任务状态", detail: "只查询原任务，不重发需求。",
      awaitingStatus: true, clientLastSequence: receivedSequence(turnId)
    };
    const id = sourceMessageId ?? append({ role: "assistant", text: "", planningPause: metadata });
    if (sourceMessageId) updateMessages((items) => items.map((item) => item.id === id ? { ...item, planningPause: metadata } : item));
    pendingInspectionMessageRef.current = id;
    setPlanningPause(null);
    setPlanningPauseReload({ turnId, checking: false, blockedReason: null, canRecheck: true });
    return id;
  };

  const checkOriginalPlanningTarget = async (indexes: readonly number[], analysisOnly = false) => {
    if (!(await targetGuard.check())) return false;
    const selected = [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)].sort((a, b) => a - b);
    const original = [...indexes].sort((a, b) => a - b);
    let identityCheck: WpsIdentityCheck = "same";
    const recorded = planningTargetRef.current;
    if (recorded) {
      try { identityCheck = verifyWpsObjectIdentity(getWpsApplication(), recorded); }
      catch { identityCheck = "unknown"; }
    }
    if (JSON.stringify(selected) !== JSON.stringify(original) || identityCheck === "different") {
      append({ role: "system", text: "原任务的工作簿或授权范围已变化；不会继续旧任务或写入原表。请按当前范围重新生成预览。" });
      return false;
    }
    if (identityCheck === "unknown") {
      // Native proxies may be recreated on every host read. Explicit scope consent
      // permits analysis against the same evidence, never adoption of old writes.
      const recordedKey = recorded && wpsConfirmationTargetKey(recorded);
      if (analysisOnly && recordedKey && targetGuard.hasAnalysisConsent(recordedKey)) {
        try {
          if (await targetGuard.checkForAnalysis()) {
            const currentKey = wpsConfirmationTargetKey(captureWpsObjectIdentity(getWpsApplication(), indexes));
            if (currentKey === recordedKey && targetGuard.hasAnalysisConsent(recordedKey)) return true;
          }
        } catch { /* Unreadable evidence remains unverified. */ }
      }
      append({ role: "system", text: "暂时无法核实原任务的目标身份，不代表授权范围已变化。原任务已保留；请在数据范围中确认原工作簿与工作表后继续。" });
      return false;
    }
    return true;
  };

  const composerPausedMessage = [...messages].reverse().find((message) =>
    (message.planningPause && !message.planningPause.resolved &&
      (message.planningPause.awaitingStatus || message.planningPause.canContinue)) ||
    prePlanningPause.entries.some((entry) => entry.messageId === message.id &&
      entry.conversationId === chatHistory.activeConversationId));
  /*
   * 三种"继续"进**同一个恢复入口**，但入口不同层：
   *  - 独立 ▶ 按钮 → 直接走暂停卡（空提交不恢复任务）；
   *  - 在输入框里打「继续」「接着做」→ 仍交给意图理解（模型判语义）后落到
   *    同一条 `resumePlanningFromPause`，不在本地再写一套关键词分类。
   */
  const composerCanContinue = Boolean(composerPausedMessage);
  /*
   * 按钮该长什么样，由**共享判据**决定（`pauseButtonState`，Excel 端读同一份）。
   *
   * 三件事在这里被显式分开，不再靠 `progress` 一个变量推断：
   *  - `taskRunning`：界面上确实有一条在执行（进度/写入）；
   *  - `controllerAlive`：真的存在一个没被 abort 的对象 —— **只有**它能让 ■ 成立；
   *  - `uninterruptible`：写入/回读期间不允许中断。
   *
   * 旧写法 `progress ? ■ : ↑` 的实际后果：`progress` 非空但没有可中止对象时，
   * 界面画出一个"能暂停"的按钮，点下去却静默返回。
   */
  const composerPauseButton = pauseButtonState({
    hasInput: false,
    paused: composerCanContinue,
    taskRunning: progress !== null || writeProgress !== null,
    controllerAlive: Boolean(controllerRef.current && !controllerRef.current.signal.aborted),
    uninterruptible: writeProgress !== null
  });
  /** 过程卡/进度条上的「停止」：同一条判据，但不受输入影响。 */
  const cardPauseButton = executionPauseState({
    taskRunning: progress !== null || writeProgress !== null,
    controllerAlive: Boolean(controllerRef.current && !controllerRef.current.signal.aborted),
    uninterruptible: writeProgress !== null
  });
  /** 写入/回读期间**禁止**中断：按钮据此置灰（其余"不可暂停"仍要给反馈，不置灰）。 */

  const submitComposer = async () => {
    if (busy || preparingImages || (!prompt.trim() && pendingImages.length === 0)) return;
    return runTurn(prompt);
  };
  const resumeComposerTask = async () => {
    if (busy || preparingImages || !composerPausedMessage || pauseRequestState?.phase === "requesting") return;
    const early = prePlanningPause.pendingEntries().find((entry) =>
      entry.messageId === composerPausedMessage.id && entry.conversationId === chatHistory.activeConversationId);
    if (early) await restartEarlyAnalysis(early.id);
    else await continuePlanningMessage(composerPausedMessage);
  };

  const restartEarlyAnalysis = async (id: string) => {
    const issued = controllerRef.current;
    const entry = prePlanningPause.pendingEntries().find((item) => item.id === id);
    if (busy || preparingImages) return;
    if (pendingPlan) {
      append({ role: "system", text: "请先确认或取消当前写入计划，再继续原需求。" });
      return;
    }
    if (entry?.data.images.length && !selectedModel?.supportsVision) {
      setImageError("原需求包含图片，请切换到支持视觉输入的模型后继续。");
      setModelMenuOpen(true);
      return;
    }
    await prePlanningPause.restart(id, chatHistory.activeConversationId, `${scopeSource}:${scopeMode}`,
      async () => {
        if (!(await targetGuard.check())) return null;
        return inspectWpsAuthorizedTarget(scopeMode === "auto" ? null : selectedSheetIndexes);
      },
      async (saved) => {
        await runTurn(saved.prompt, { images: saved.images, taskContext: saved.taskContext,
          messageUnderstanding: saved.messageUnderstanding, activityMessageId: saved.activityMessageId, restartingAnalysis: true });
      }, () => controllerRef.current === issued);
  };

  const continuePlanningMessage = async (message: ChatMessage) => {
    const pause = await inspectPlanningMessage(message);
    if (pause) await resumePlanningFromPause(pause);
  };

  const inspectPlanningMessage = async (message: ChatMessage): Promise<NonNullable<typeof planningPause> | undefined> => {
    const metadata = message.planningPause;
    if (!metadata || metadata.resolved || invalidPlanningTurnsRef.current.has(metadata.turnId) || planningInspectionRef.current || resumingPauseRef.current) return;
    retainPlanningContext();
    const saved = planningContextsRef.current.get(metadata.turnId);
    if (saved && saved.conversationId === chatHistory.activeConversationId) {
      pauseCarryRef.current = saved.carry;
      planningTargetRef.current = saved.target;
      taskIdRef.current = saved.taskId;
      streamCursorRef.current = saved.cursor;
      resumeRequestRef.current = saved.resumeIdentity;
    }
    const carry = pauseCarryRef.current?.turnId === metadata.turnId ? pauseCarryRef.current : null;
    if (!carry) {
      // History without live ownership can still be queried, but never adopted
      // using the current workbook/sheet selection as if it belonged to the old task.
      const inspection = new AbortController();
      planningInspectionRef.current = inspection;
      controllerRef.current = inspection;
      setPlanningPause(null);
      setPlanningPauseReload({ turnId: metadata.turnId, checking: true, blockedReason: null });
      setProgress({ stage: "answering", message: "正在查询原任务状态…" });
      let continuation: NonNullable<typeof planningPause> | undefined;
      try {
        const status = await inspectWpsPlanningTask(metadata.turnId, inspection.signal);
        if (inspection.signal.aborted || planningInspectionRef.current !== inspection) return;
        const view = wpsPlanningResumeView(status);
        if (view.canContinue && view.action === "continue_checkpoint") {
          // Same admission as reopening the pane: the server has a checkpoint,
          // but the eventual POST must still verify freshly read identity/scope/source.
          pauseCarryRef.current = null;
          planningTargetRef.current = null;
          streamCursorRef.current = { turnId: metadata.turnId, sequence: metadata.clientLastSequence ?? 0 };
          continuation = { turnId: metadata.turnId, payload: planningPauseFromStatus(status, null),
            view, originalPrompt: "", sourceMessageId: message.id };
          setPlanningPause(continuation);
        }
        const reason = status.status === "running" ? "原任务仍在服务端处理中。"
          : status.status === "completed" ? "原任务已完成。"
          : "已查询原任务状态。";
        updateMessages((items) => items.map((item) => item.id === message.id && item.planningPause
          ? { ...item, planningPause: { ...item.planningPause,
              reason: `${reason}本机缺少原任务的目标上下文，未继续、未写入。` } } : item));
      } catch {
        if (!inspection.signal.aborted && planningInspectionRef.current === inspection) {
          updateMessages((items) => items.map((item) => item.id === message.id && item.planningPause
            ? { ...item, planningPause: { ...item.planningPause,
                reason: "暂时无法查询原任务状态；请稍后核对，不要重复发送原需求。" } } : item));
        }
      } finally {
        if (planningInspectionRef.current === inspection) {
          planningInspectionRef.current = null;
          if (controllerRef.current === inspection) controllerRef.current = null;
          setPlanningPauseReload({ turnId: metadata.turnId, checking: false, blockedReason: null, canRecheck: true });
          setProgress(null);
        }
      }
      return !inspection.signal.aborted ? continuation : undefined;
    }
    pendingInspectionMessageRef.current = message.id;
    return await reconcileInterruptedTurn(metadata.turnId, carry.originalPrompt, carry.selectedSheetIndexes, message.id);
  };

  // Check the original target before reconnecting either a running or checkpointed task.
  const preparePlanningBridge = async (
    turnId: string,
    status: WpsPlanningTaskStatus,
    sheetIndexes: readonly number[],
    signal: AbortSignal,
    /**
     * 立桥的那次执行。桥立起来之后**只由它关闭** —— 核对把 controller 接走时，
     * 这个身份不会跟着变。
     */
    owner: string
  ) => {
    if (!(await checkOriginalPlanningTarget(sheetIndexes, true))) {
      return null;
    }
    if (signal.aborted) throw new DOMException("Aborted", "AbortError");
    /*
     * **现场证据必须此刻现读**，而且要与交给桥的是同一份。
     *
     * 之前这里写的是"读不到就不声称核实过（字段省略）"，而服务端当时也确实
     * 允许省略 —— 两边合起来就是"读不到宿主也能继续"。现在证据是协议必填，
     * 读不到就**停下来**：
     *
     * - 工作簿身份用 `wpsWorkbookIdentity(fullName)`，不拿文件名兜底；
     * - 授权成员取"建卡时记下的授权表序号"过滤出的名字（与桥绑定的范围一致）；
     * - 来源指纹三态由共享层判定，"读取失败"不会退化成"尚未采集"。
     *
     * `createWpsBridgeLiveContext` 一次读齐这三项，正是桥在绑定/执行前复核时
     * 用的**同一个读取器** —— 证据与绑定不同源，迟早对不上。
     */
    let live: ReturnType<typeof createWpsBridgeLiveContext> | null = null;
    try {
      live = createWpsBridgeLiveContext(getWpsApplication(), sheetIndexes);
    } catch {
      live = null;
    }
    if (!live) {
      setPlanningPauseReload({ turnId, checking: false,
        blockedReason: "读取当前工作簿现场失败。", canRecheck: true });
      append({
        role: "system",
        text: "读取当前工作簿现场失败，无法确认是否仍是原任务那一份，已停止本次继续；原表未写入。"
      });
      return null;
    }
    const evidenceResult = await buildPlanningContinuationEvidence({
      workbookIdentity: live.workbookIdentity,
      authorizedSheetNames: live.authorizedSheetNames,
      // 服务端从检查点判定；缺省按"需要核实"处理。
      sourceRequired: status.continuation?.sourceVerificationRequired !== false,
      readSourceFingerprint: () => live!.sourceFingerprint
    });
    if (!evidenceResult.ok) {
      setPlanningPauseReload({ turnId, checking: false,
        blockedReason: PLANNING_EVIDENCE_FAILURE_MESSAGES[evidenceResult.failure], canRecheck: true });
      append({ role: "system", text: PLANNING_EVIDENCE_FAILURE_MESSAGES[evidenceResult.failure] });
      return null;
    }
    const evidence = evidenceResult.evidence;
    if (status.sourceFingerprintSheets?.length && JSON.stringify([...status.sourceFingerprintSheets].sort()) !==
        JSON.stringify([...evidence.authorizedSheetNames].sort())) {
      append({ role: "system", text: "原任务的授权工作表范围已变化，未接回读表工具。" });
      return null;
    }
    const expectedFingerprint = status.sourceFingerprint ??
      (sdkBridge.binding()?.taskId === turnId ? sdkBridge.binding()?.sourceFingerprint : null);
    if (
      evidence.sourceVerification === "verified" &&
      expectedFingerprint &&
      evidence.sourceFingerprint !== expectedFingerprint
    ) {
      // 读到了但与原任务记录的不同：数据变过，不能沿用旧现场。
      const reason = "来源数据已变化，请按当前数据重新发送需求。";
      invalidPlanningTurnsRef.current.add(turnId);
      updateMessages((items) => items.map((item) => item.planningPause?.turnId === turnId
        ? { ...item, planningPause: { ...item.planningPause, resolved: true,
            terminalReason: reason, canContinue: false, awaitingStatus: false, detail: reason } }
        : item));
      setPlanningPause((current) => current?.turnId === turnId ? null : current);
      setPlanningPauseReload((current) => current?.turnId === turnId ? null : current);

      append({
        role: "system",
        text: "原任务使用的数据已经变化，不能继续旧的规划现场；请按当前数据重新发送需求。"
      });
      return null;
    }
    /*
     * **继续之前先把桥立起来。** 上一轮收尾时桥已 `end()`，绑定虽然还在（结果
     * 按引用取值要用），但已不能服务。少了这一步，续跑里模型发出的读表指针会被
     * 交给一条已收尾的桥，谁也不再轮询，最后只能等服务端超时。
     */
    const bridgeReady = await ensureWpsBridge(turnId, sheetIndexes, signal, false, owner);
    if (!bridgeReady) {
      /*
       * **桥没立起来就不发续跑请求。**
       *
       * 以前这里只把 `bridgeReady` 用在"进度事件要不要交给桥"上，请求照发：
       * 服务端一发出 `bridgeTool` 指针就没人接，模型干等到超时，而界面上看起来
       * 续跑是"成功"的 —— 因为 `resumeWpsPlanning` 真的返回了。
       *
       * `ensureLive` 返回 false 只有一种含义：这一次建立没有成立（初始化期间被
       * 取消、被收尾，或被另一轮取代）。不发送才是如实。
       */
      if (signal.aborted) throw new DOMException("Aborted", "AbortError");
      const notice = "本机读表桥没有为本次继续建立起来，没有发送继续请求；任务现场已保留，可以稍后再继续。";
      setPlanningPauseReload({
        turnId,
        checking: false,
        blockedReason: notice,
        canRecheck: true
      });
      append({ role: "system", text: notice });
      return null;
    }
    return evidence;
  };

  const completePlanningAttachment = (activity: ReturnType<typeof beginActivity>) => {
    updateActivity(activity, { id: "planning.attach", label: "已接回原任务进度", phase: "planning", status: "completed" });
  };

  // Replay only unseen committed events into the current reply segment.
  const tailRunningTurn = async (
    turnId: string,
    sourceMessageId: string | undefined,
    current: () => boolean,
    signal: AbortSignal
  ): Promise<"settled" | "timeout"> => {
    const activity = resumeTaskActivity(turnId, sourceMessageId);
    let cursor = receivedSequence(turnId);
    const deadline = Date.now() + RUNNING_TAIL_MAX_MS;
    let handedOff = false;
    try {
      await waitForPlanning(sdkBridge.pollPending(), signal);
      for (;;) {
        if (!current()) return "timeout";
        let page: WpsPlanningTaskEvents;
        try {
          page = await waitForPlanning(getWpsPlanningTaskEvents(turnId, cursor, signal), signal);
        } catch {
          // 读不到就停下：断线不等于可以重发需求，也不等于任务已经结束。
          return "timeout";
        }
        if (!current()) return "timeout";
        completePlanningAttachment(activity);
        for (const event of page.events) {
          if (!acceptsPlanningReplay(turnId, cursor, event)) continue;
          // 事件序号只前进不后退；重复/乱序帧由共享时间线层自己挡。
          cursor = Math.max(cursor, event.sequence);
          const progressEvent = wpsProgressFromTaskEvent(turnId, event);
          recordPlanningProgress(progressEvent);
          recordPlanningResult(progressEvent);
          updateActivity(activity, readonlyActivityUpdate(progressEvent));
          sdkBridge.handleProgressEvent(progressEvent);
        }
        if (!page.active) {
          // 任务已经不在跑了：成功就让下游的"取回结果"路径收尾（它会把结果接进
          // 同一条时间线）；失败/暂停则在这里如实收尾，不留一条"进行中"。
          if (page.status !== "completed") {
            endActivity(activity, page.status === "waiting_user" ? "awaiting_input" : "warning");
          }
          handedOff = page.status === "completed";
          return "settled";
        }
        if (Date.now() >= deadline) return "timeout";
        await waitForPlanning(new Promise((resolve) => setTimeout(resolve, RUNNING_TAIL_INTERVAL_MS)), signal);
      }
    } finally {
      if (!handedOff) endActivity(activity, "awaiting_input");
    }
  };

  const reconcileInterruptedTurn = async (
    turnId: string,
    text: string,
    turnSheetIndexes: readonly number[],
    sourceMessageId?: string,
    /**
     * 发现原任务**仍在服务端跑**时，是否把已落盘的进度接回那张过程卡。
     *
     * 默认接回（用户点了"继续分析 / 核对任务状态"，他就是要看原任务现在到哪了）。
     * 但**暂停之后那一次自动核对要传 false**：用户刚按下暂停，服务端任务还在跑，
     * 这时去轮询补发等于把"已暂停"悄悄变成"其实还在进行中"，还会把这次核对卡住
     * 最多一个补发上限。暂停只该留下一条简短的真实状态。
     */
    attachRunning = true
  ): Promise<NonNullable<typeof planningPause> | undefined> => {
    if (planningInspectionRef.current) return;
    const inspection = new AbortController();
    /*
     * 核对是**独立的一次执行**：它接管 `controllerRef`，但不接管别人开的活动。
     * 身份贴着 controller 走（`tagControllerOwner`），于是发起核对的那次续跑回来
     * 时能看清"controller 已经换人了"，而不是把所有收尾一起放弃。
     */
    const inspectionOwner = newExecutionId("inspect");
    tagControllerOwner(inspection, inspectionOwner);
    let attachedBridge = false;
    planningInspectionRef.current = inspection;
    controllerRef.current = inspection;
    if (!sourceMessageId) {
      rememberPlanningInterruption(turnId, text, turnSheetIndexes, 0);
      sourceMessageId = pendingInspectionMessageRef.current ?? undefined;
    }
    const current = () => !inspection.signal.aborted && planningInspectionRef.current === inspection;
    const settleInspection = (reason: string | null, canRecheck: boolean) => {
      setPlanningPauseReload({ turnId, checking: false, blockedReason: reason, canRecheck });
    };
    const replaceCard = (metadata: NonNullable<ChatMessage["planningPause"]>) => {
      updateMessages((items) => items.map((item) => item.id === sourceMessageId ? { ...item, planningPause: metadata } : item));
    };
    setPlanningPause(null);
    setPlanningPauseReload({ turnId, checking: true, blockedReason: null });
    setProgress({ stage: "answering", message: "正在核对原任务的真实进度…" });
    try {
      /*
       * 最多两轮：第一轮如果发现原任务**还在服务端跑**，就把已落盘的事件接回原来
       * 那张过程卡（`tailRunningTurn`），等它结束后再判一次状态，走"取回结果"或
       * "从检查点续跑"这两条既有分支 —— 不再只是回一句"稍后核对"。
       */
      let status: WpsPlanningTaskStatus | null = null;
      for (let attempt = 0; ; attempt += 1) {
        try {
          status = await waitForPlanning(inspectWpsPlanningTask(turnId, inspection.signal), inspection.signal);
          if (!current()) return;
        } catch (error) {
          if (!current()) return;
          /*
           * 区分两件事：**服务端说没有这条任务**（404，可以重发）与
           * **读不到服务端**（后端没起，禁止重发 —— 那只是把同样的失败再跑一遍）。
           * 只有前者才放行"重新发送"。
           */
          const notFound =
            error instanceof ApiRequestError &&
            (error.status === 404 || error.code === "PLANNING_TASK_NOT_FOUND");
          append({
            role: "system",
            text: notFound
              ? "服务端没有这条任务的记录，可以重新发送需求。本次未修改工作表。"
              : "连接中断，且暂时读不到该任务的真实进度。请先确认本地 AI 服务已启动，稍后重试核对；在核对清楚之前请不要重复发送同一需求，以免与原任务重复规划。"
          });
          if (notFound) replaceCard({ turnId, pauseType: "interrupted", reason: "服务端没有该任务记录，不能从检查点继续。",
            canContinue: false, action: "none", actionLabel: "", detail: "原需求保留在对话中；如需处理请重新开始任务。",
            awaitingStatus: false, clientLastSequence: receivedSequence(turnId) });
          settleInspection(notFound ? "服务端没有该任务记录，不能从检查点继续。" : "暂时读不到原任务状态，请稍后核对。", !notFound);
          return;
        }
        if (status.status !== "running" || attempt > 0 || !attachRunning) break;
        // 原任务仍在跑：先把进度接回来（不重发、不重新规划），再判一次。
        setProgress({ stage: "answering", message: "原任务仍在服务端处理中，正在接回进度…" });
        if (!(await preparePlanningBridge(turnId, status, turnSheetIndexes, inspection.signal, inspectionOwner)) || !current()) return;
        attachedBridge = true;
        const attached = await tailRunningTurn(turnId, sourceMessageId, current, inspection.signal);
        if (!current()) return;
        if (attached === "settled") continue;
        settleInspection("原任务仍在服务端处理中；稍后核对，不重复规划。", true);
        replaceCard({ turnId, pauseType: "interrupted", reason: "原任务仍在服务端处理中；进度已接回，稍后核对，不重复规划。",
          canContinue: false, action: "none", actionLabel: "核对任务状态", detail: "只查询原任务，不重发需求。",
          awaitingStatus: true, clientLastSequence: receivedSequence(turnId) });
        return;
      }
      if (!status) return;
      const view = wpsPlanningResumeView(status);
      if (status.status === "completed" && !attachRunning) {
        const reason = "原任务已完成，结果已保留；继续时取回结果。";
        replaceCard({ turnId, pauseType: "interrupted", reason, canContinue: false,
          action: "none", actionLabel: "核对任务状态", detail: reason, awaitingStatus: true,
          clientLastSequence: receivedSequence(turnId) });
        settleInspection(reason, true);
        return;
      }
      if (status.status === "completed" && status.result) {
        // Structural deletion normally enters the checkpoint writer immediately.
        // A status query is not permission to invoke that writer.
        if (status.result.kind === "plan" && isWorksheetDeletion(status.result.plan)) {
          const reason = "原任务已生成删除方案；本次仅查询状态，未执行删除。";
          replaceCard({ turnId, pauseType: "interrupted", reason, canContinue: false,
            action: "none", actionLabel: "", awaitingStatus: false,
            detail: "删除方案需重新进入授权与执行检查，不能通过状态查询直接执行。",
            clientLastSequence: receivedSequence(turnId) });
          settleInspection(reason, false);
          return;
        }
        if (!(await checkOriginalPlanningTarget(turnSheetIndexes)) || !current()) {
          if (current()) settleInspection("请先核对原任务的数据范围，再取回结果。", true);
          return;
        }
        // 任务其实跑完了，只是回程断线：把结果接回同一条处理链，不重发。
        const activity = resumeTaskActivity(turnId, sourceMessageId);
        completePlanningAttachment(activity);
        try {
          await applyTurnResult(status.result as WpsTurnResult, {
            activity,
            text,
            turnSheetIndexes: [...turnSheetIndexes],
            clarificationRound: pauseCarryRef.current?.clarificationRound ?? 0,
            images: []
          });
          updateMessages((items) => items.map((item) => item.planningPause?.turnId === turnId
            ? { ...item, planningPause: { ...item.planningPause, resolved: true, canContinue: false, awaitingStatus: false,
                clientLastSequence: receivedSequence(turnId) } } : item));
          settleInspection("任务已完成，结果已取回。", false);
          replaceCard({ turnId, pauseType: "interrupted", reason: "任务已完成，结果已取回。", canContinue: false,
            action: "none", actionLabel: "", detail: "未自动写入。", resolved: true, clientLastSequence: receivedSequence(turnId) });
          pendingInspectionMessageRef.current = null;
          pauseCarryRef.current = null;
        } catch (applyError) {
          if (inspection.signal.aborted) {
            endActivity(activity, "awaiting_input");
            return;
          }
          endActivity(activity, applyError instanceof TargetConfirmationRequiredError ? "awaiting_input" : "failed");
          append({ role: "system", text: turnErrorMessage(applyError) });
          settleInspection("未能接回结果；请核对原任务的数据范围后重试。", true);
        }
        return;
      }
      if (status.status === "running") {
        settleInspection("原任务仍在服务端处理中；稍后核对，不重复规划。", true);
        replaceCard({ turnId, pauseType: "interrupted", reason: "原任务仍在服务端处理中；稍后核对，不重复规划。",
          canContinue: false, action: "none", actionLabel: "核对任务状态", detail: "只查询原任务，不重发需求。",
          awaitingStatus: true, clientLastSequence: receivedSequence(turnId) });
        return;
      }
      if (view.canContinue && view.action === "continue_checkpoint") {
        if (!(await checkOriginalPlanningTarget(turnSheetIndexes, true)) || !current()) {
          if (current()) settleInspection("请先核对原任务的数据范围。", true);
          return;
        }
        /*
         * 可续跑：落一张暂停卡，而不是让用户重发。
         *
         * 用**服务端刚返回的真实状态**建卡，并记下续跑上下文 —— 与正常暂停
         * 路径共用同一张卡与同一个处理链。
         */
        pauseCarryRef.current = {
          turnId,
          originalPrompt: text,
          selectedSheetIndexes: [...turnSheetIndexes],
          clarificationRound: pauseCarryRef.current?.clarificationRound ?? 0
        };
        const metadata: NonNullable<ChatMessage["planningPause"]> = {
          turnId, pauseType: status.pauseType ?? "interrupted", reason: status.reason ?? "已核实可续跑检查点。",
          canContinue: true, action: view.action, actionLabel: view.actionLabel, detail: view.detail,
          clientLastSequence: receivedSequence(turnId)
        };
        replaceCard(metadata);
        const continuation = {
          turnId,
          payload: planningPauseFromStatus(status, null),
          view,
          originalPrompt: text,
          sourceMessageId: sourceMessageId!
        };
        setPlanningPause(continuation);
        settleInspection(null, false);
        pendingInspectionMessageRef.current = null;
        return continuation;
      }
      append({
        role: "system",
        text: `该任务当前不能继续分析：${view.reason} 原需求仍保留；不会自动重新规划。`
      });
      replaceCard({ turnId, pauseType: status.pauseType ?? "interrupted", reason: view.reason,
        canContinue: false, action: "none", actionLabel: "", detail: view.detail, awaitingStatus: false,
        clientLastSequence: receivedSequence(turnId) });
      settleInspection(view.reason, false);
    } finally {
      if (planningInspectionRef.current === inspection) {
        /*
         * 桥按**立桥那次执行**收，不按"这一轮还归我吗"收：核对可能把
         * `controllerRef` 又交给了别人（用户在其间按了继续），那不代表它立的桥
         * 该留给下一轮。
         */
        if (attachedBridge) endBridgeOwnedBy(inspectionOwner);
        if (inspection.signal.aborted) settleInspection("已暂停状态核对；尚未继续原任务。", true);
        if (pauseRequestedRef.current === inspection) pauseRequestedRef.current = null;
        planningInspectionRef.current = null;
        if (controllerRef.current === inspection) controllerRef.current = null;
        setProgress(null);
      }
    }
  };

  /**
   * 把一次**结构化结果**接进正常的处理链。
   *
   * 抽出来的唯一理由：`continue_checkpoint` 续跑回来的结果与一次性轮次
   * 是**同一个联合类型**（`WpsTurnResult`）。此前续跑只处理了 `answer` 和
   * `planning_pause`，`plan` / `clarification` / 本地映射一律落到
   * "请重新发送需求" —— 那等于**让用户重做一遍已经做过的事**，而且把
   * "服务端其实已经算好了"这个事实说成"算不了"。
   *
   * 这条纪律必须只有一份实现：两条通道各写一份分支，迟早会漂移成
   * "一次性通道能出预览、续跑通道只能重发"。
   *
   * 返回值为**澄清这一轮读到的表头指纹**（无澄清时为 undefined）。
   * 调用方据此建卡；本函数不替它决定要不要落卡。
   */
  const applyTurnResult = async (
    result: WpsTurnResult,
    context: {
      activity: ReturnType<typeof beginActivity>;
      text: string;
      turnSheetIndexes: readonly number[];
      /** 本轮是第几轮澄清（续跑通道从原暂停载荷里带过来）。 */
      clarificationRound: number;
      /** 出错时要回填的待发送图片（续跑通道为空）。 */
      images: readonly PendingImage[];
    }
  ): Promise<string | undefined> => {
    const { activity, text, turnSheetIndexes } = context;
    const signal = controllerRef.current?.signal;
    const checkActive = () => { if (signal?.aborted) throw new DOMException("Aborted", "AbortError"); };
    checkActive();
    const clarificationRound = context.clarificationRound;
    const sentImages = context.images;
    let headerFingerprint: string | undefined;
      /**
       * 结构扫描已经在这一轮里发生（`runWpsTurn` 的第一步），它**真的读到了
       * 表头单元格**。直接把那份表头拿来做指纹 —— 不要再回头去读工作表名
       * （那是身份判据的数据，混用会让两条判据互相遮蔽）。
       *
       * `headerRows` 缺失时留 `undefined`，该判据不参与比对（不误伤）。
       */
      if (result.kind === "clarification" && result.headerRows) {
        headerFingerprint = wpsHeaderFingerprint(result.headerRows);
      }
      if (result.kind === "clarification") {
        const sheetIdentity = currentWpsSheetIdentity();
        // 记下这次建卡针对的授权表序号，复核表头时要用同一批表。
        clarificationScopeIndexesRef.current = [...turnSheetIndexes];
        /**
         * 捕获**宿主对象身份**并留在内存里。
         *
         * 指纹证明不了"同索引、同名、同表头的另一张表"，也证明不了"不同目录的
         * 同名工作簿"—— 那两类只有对象引用（+ 路径 + 稳定 ID）能分辨。
         * 读不到宿主身份时置空：复核会据此 **fail-closed**，而不是放行。
         */
        try {
          clarificationObjectIdentityRef.current = captureWpsObjectIdentity(
            getWpsApplication(),
            turnSheetIndexes
          );
        } catch {
          clarificationObjectIdentityRef.current = null;
        }
        endActivity(activity, "awaiting_input");
        if (result.clarification.localMapping) {
          const localMapping = result.clarification.localMapping;
          const clarificationMessageId = append({
            role: "assistant",
            text: "",
            taskResume: buildTaskResumeRecord({
              taskId: carried().taskId,
              phase: "reading_source",
              blocker: "user_input",
              resumability: "reauthorize",
              writeState: "not_written",
              originalPrompt: resumeCarryRef.current?.originalPrompt ?? text,
              confirmations: carried().confirmations,
              authorizationRevision
            }),
            clarification: {
              summary: "本地映射需要你决定",
              question: localMapping.allowExactMapping
                ? `源区域有 ${localMapping.values.length} 个不同来源值未能按现有规则归类。请为每项明确填写整格的目标值。`
                : `源区域有 ${localMapping.values.length} 个拆分项未被规则覆盖。请说明新增项的映射及多项优先级。`,
              reason: !localMapping.allowExactMapping
                ? "现有规则会拆分来源格；请在卡片内补齐映射并明确优先顺序，再重新读取生成预览。"
                : localMapping.reason === "separator_or_mapping"
                ? "其中可能含有未声明的分隔形式；这里不自动拆分，也不猜测优先级。可在本卡片内确认拆分规则。"
                : "这些值未被规则覆盖；不会用空白标签代替。若一格代表多项，也可改为说明拆分与优先级规则。",
              scopeLabel: `${localMapping.rule.sourceSheet}!${localMapping.rule.sourceRange}`,
              options: [],
              status: "pending",
              sheetIdentityFingerprint: sheetIdentity ?? undefined,
              headerFingerprint
            }
          });
          setClarification({
            messageId: clarificationMessageId,
            originalPrompt: text,
            round: clarificationRound,
            options: [],
            // 本地映射的题面由"停在哪种缺口"决定，与消息 id 无关。
            questionId: `local-mapping:${localMapping.reason}`,
            sheetIdentityFingerprint: sheetIdentity ?? undefined,
            headerFingerprint,
            localMapping,
            localMappingRevision: 1,
            selectedSheetIndexes: [...turnSheetIndexes]
          });
          return;
        }
        if (result.clarification.columns) {
          // 列映射澄清：**不读源值、不算计划**，只把归因接进同一张澄清卡。
          // 卡片文案与选项都来自 `@excel-bro/core`（Excel/WPS 共用），本端不重复实现。
          const { columns } = result.clarification;
          const message = columnMappingClarificationMessage({
            clarification: columns.clarification,
            scopeLabel: columns.scopeLabel,
            originalPrompt: text,
            scopeFingerprint: sheetIdentity ?? "",
            round: (clarificationRound) + 1
          });
          const clarificationMessageId = append({
            role: "assistant",
            text: columns.summary,
            taskResume: buildTaskResumeRecord({
              taskId: carried().taskId,
              phase: "awaiting_clarification",
              blocker: "user_input",
              resumability: "reauthorize",
              writeState: "not_written",
              originalPrompt: resumeCarryRef.current?.originalPrompt ?? text,
              confirmations: carried().confirmations,
              authorizationRevision
            }),
            clarification: {
              summary: redactCredentials(message.summary),
              question: redactCredentials(message.question),
              reason: redactCredentials(message.reason),
              scopeLabel: redactCredentials(message.scopeLabel),
              options: message.options.map((option) => ({
                label: redactCredentials(option.label),
                description: redactCredentials(option.description)
              })),
              status: "pending",
              sheetIdentityFingerprint: sheetIdentity ?? undefined,
              headerFingerprint
            }
          });
          setClarification({
            messageId: clarificationMessageId,
            originalPrompt: resumeCarryRef.current?.originalPrompt ?? text,
            round: message.round,
            options: message.options,
            questionId: `column-mapping:${columns.clarification.reason}`,
            sheetIdentityFingerprint: sheetIdentity ?? undefined,
            headerFingerprint
          });
          return;
        }
        const response = result.clarification.response;
        if (!response) return;
        setClarificationDraft("");
        setClarificationFieldDrafts({});
        const clarificationMessageId = append({
          role: "assistant",
          text: "",
          taskResume: buildTaskResumeRecord({
            taskId: carried().taskId,
            phase: "awaiting_clarification",
            blocker: "user_input",
            resumability: "reauthorize",
            writeState: "not_written",
            originalPrompt: resumeCarryRef.current?.originalPrompt ?? text,
            confirmations: carried().confirmations,
            authorizationRevision
          }),
          clarification: {
            summary: redactCredentials(response.clarification.summary),
            question: redactCredentials(response.clarification.question),
            reason: redactCredentials(response.clarification.reason),
            scopeLabel: redactCredentials(response.clarification.scopeLabel),
            options: response.clarification.options.map(
              (option: { label: string; description: string }) => ({
                label: redactCredentials(option.label),
                description: redactCredentials(option.description)
              })
            ),
            fields: response.clarification.fields?.map(
              (field: { id: string; label: string; placeholder?: string }) => ({
                id: field.id, label: redactCredentials(field.label),
                placeholder: redactCredentials(field.placeholder ?? "")
              })
            ),
            status: "pending",
            sheetIdentityFingerprint: sheetIdentity ?? undefined,
            headerFingerprint
          }
        });
        setClarification({
          messageId: clarificationMessageId,
          originalPrompt: text,
          round: clarificationRound,
          options: response.clarification.options,
          questionId: response.clarification.id,
          sheetIdentityFingerprint: sheetIdentity ?? undefined,
          headerFingerprint
        });
      } else if (result.kind === "plan") {
        if (sdkBridge.binding()?.taskId === result.turnId && sdkBridge.failure()) {
          throw new Error('本轮本地读取未通过现场核对，请重新读取后生成结果；原表未写入。');
        }
        setProgress({ stage: "answering", message: "正在校验写入计划和授权边界…" });
        updateActivity(activity, {
          id: "preflight-plan",
          label: "校验写入计划与授权",
          phase: "preflight",
          detail: "准备可供用户检查的计划预览；此时不写入工作簿。"
        });
        const application = getWpsApplication();
        const metadata = listWpsWorksheets(application);
        const selectedNames = metadata.sheets
          .filter((sheet) => turnSheetIndexes.includes(sheet.index))
          .map((sheet) => sheet.name);
        // 本端算出的计划不送服务端校验：验收条件里带着完整值矩阵，上传等于
        // 把刚才刻意留在本地的源值送出去。理由详见 WpsTurnResult.computedLocally。
        const normalizedPlan = result.computedLocally
          ? result.plan
          : await waitForPlanning(validatePlan(
              result.plan,
              selectedNames,
              metadata.sheets.map((sheet) => sheet.name), signal
            ), signal);
        checkActive();
        /*
         * 计划的**实际来源依赖**（与 Excel 端逐字同源）：
         * - 由本地结果编译出的计划（带 `localResultRef`）：来源由服务端按"实际
         *   交给执行器的输入句柄"认定，取它与本端记账的**交集**；
         * - 其余计划：只认本端这一轮桥上的确定性读取记录（`recordTrustedToolReads`
         *   在读取那一刻记账）。没有记录 = 依赖不确定 = 删掉该字段。
         * 缺失这条，WPS 端的恢复登记会恒判"依赖不确定"，与 Excel 端分叉。
         */
        const trustedReads = trustedToolReads(
          result.turnId ?? turnIdRefForStream.current ?? ""
        );
        const planWithDependency = normalizedPlan.localResultRef
          ? attestResultSourceDependency(normalizedPlan, trustedReads ? [trustedReads] : [])
          : attestSourceDependency(normalizedPlan, trustedReads);
        /*
         * **来源锚点 = 模型读到的那一版。**
         *
         * 服务端编译出来的计划里是**固定值**（隔离容器算完的矩阵）。它到达面板时
         * 再读一次"现在的指纹"是错的：只要用户在"模型读完数据"到"计划到达"之间
         * 改了源表，这次现读记的就是改后的现场，写前复核再读一次当然相等，
         * 一份按旧来源算出来的值就被静默写进单元格（2026-10-25 实测：新增写入 1 次）。
         *
         * 所以锚点取**本轮桥绑定**的指纹（`pinSourceFingerprint` 在第一个受控读取
         * 时固定下来的那一版），与 Excel 端同一个来源、同一份判据
         * （`@excel-bro/core/planSourceAnchor`）。
         *
         * 本地映射计划（`computedLocally`）不套这个锚点：它的值是**刚刚**从现场读出来
         * 算的，且执行器已经按 `prepared.sources` 的 `previewFingerprint` 逐区域复核。
         */
        const anchoredPlan = stampLiteralPlanOrigin(result.computedLocally
          ? planWithDependency
          : stampPlanSourceAnchor(
              planWithDependency,
              sourceAnchorFromBinding(
                sdkBridge.binding(),
                result.turnId ?? turnIdRefForStream.current ?? ""
              )
            ), result.planOrigin);
        if (result.computedLocally && !result.sourceSnapshots?.length) {
          throw new Error("本地映射计划缺少源区域快照，不能安全预览写入。");
        }
        if (!(await waitForPlanning(targetGuard.check(), signal))) throw new TargetConfirmationRequiredError();
        checkActive();
        if (isWorksheetDeletion(anchoredPlan)) {
          const outcome = await executeStructuralDeletion(anchoredPlan, turnSheetIndexes, activity.id);
          endActivity(activity, outcome);
          return;
        }
        const prepared = prepareWpsPlan(application, turnSheetIndexes, anchoredPlan, result.sourceSnapshots, { requireSourceAnchor: !result.computedLocally });
        updateActivity(activity, { id: "preflight-plan", label: "写入计划与授权已校验", phase: "preflight", status: "completed" });
        endActivity(activity, "awaiting_input");
        setPendingPlan(prepared);
        setPendingPlanMessageId(append({
          role: "assistant",
          text: "",
          // 预览已生成、等用户确认：写前状态仍是"未写入"，但预览本体（含目标矩阵）
          // 不落盘，所以重开后只能重新生成预览，不能说成"从进度精确续跑"。
          taskResume: previewTaskResume(
            carried().taskId,
            resumeCarryRef.current?.originalPrompt ?? text,
            authorizationRevision,
            result.turnId,
            carried().confirmations
          ),
          planPreview: { title: prepared.plan.title, status: "pending" }
        }));
      } else if (result.kind === "planning_pause") {
        /*
         * 规划暂停（CAP-02）。
         *
         * 这不是失败：服务端的规划现场已经持久化，用户可以选择继续或换模型。
         * 关键纪律：**不在这里写任何单元格**，也不把暂停当"已完成"收尾 ——
         * 它是"等用户决定"的中间态。按钮与文案全部由 `pause` 的结构化字段决定
         * （`wpsPlanningResumeView`），不看助手文本、不猜业务关键词。
         */
        updateMessages((items) => items.map((item) => item.id === activity.id && item.activityLog
          ? { ...item, activityLog: { ...item.activityLog, operationId: result.pause.turnId } } : item));
        endActivity(activity, "awaiting_input");
        setPrompt((current) => current.trim() === text ? "" : current);
        if (sentImages.length) clearImages();
        const pauseView = wpsPlanningResumeView(result.pause);
        /*
         * 记下"这是哪一件事"的上下文，供续跑结果回接处理链时使用。
         * 与暂停卡生命同周期；续跑成功/失败后由 `resumePlanningFromPause` 清理。
         */
        pauseCarryRef.current = {
          turnId: result.pause.turnId,
          originalPrompt: resumeCarryRef.current?.originalPrompt ?? text,
          selectedSheetIndexes: [...turnSheetIndexes],
          clarificationRound: clarificationRound
        };
        retainPlanningContext();
        setPlanningPause({
          turnId: result.pause.turnId,
          payload: result.pause,
          view: pauseView,
          originalPrompt: text,
          sourceMessageId: append({
            role: "assistant",
            text: "",
            planningPause: {
              turnId: result.pause.turnId,
              pauseType: result.pause.pauseType,
              reason: result.pause.reason,
              canContinue: result.pause.canContinue,
              action: pauseView.action,
              actionLabel: pauseView.actionLabel,
              detail: pauseView.detail,
              clientLastSequence: receivedSequence(result.pause.turnId)
            }
          })
        });
      } else {
        endActivity(activity, result.outcome === "failed" ? "failed" : "completed");
        if (result.outcome === "failed") {
          lastTurnFailedRef.current = true;
          if (sentImages.length) {
            setPendingImages([...sentImages]);
            setPrompt(text);
          }
        }
        append({
          role: result.outcome === "failed" ? "system" : "assistant",
          text: result.message,
          resultContext: result.resultContext,
          // SDK 本地结果卡：只带引用与结构，值由本机回环按引用取回后渲染。
          localResult: result.localResult ?? undefined,
          // 会话关闭后服务端按授权归属校验，所以要带上**这一轮**的绑定。
          localResultBinding: localResultBindingOf(result.localResult)
        });
      }
    return headerFingerprint;
  };

  /**
   * 把"这一轮"的桥绑定固化到消息上。
   *
   * 结果卡的取值发生在**会话结束之后**，此时不能再问控制器"当前绑定是谁"——
   * 下一轮可能已经换了绑定。所以按消息快照一份：任务身份 + 工作簿身份 + 源指纹。
   */
  const localResultBindingOf = (card: unknown): BridgeResultLookup | undefined => {
    if (!card) return undefined;
    const binding = sdkBridge.binding();
    if (!binding) return undefined;
    return {
      taskId: binding.taskId,
      hostId: binding.hostId,
      workbookName: binding.workbookName,
      workbookIdentity: binding.workbookIdentity,
      sourceFingerprint: binding.sourceFingerprint
    };
  };

  retryTurnRef.current = (question, options) => { void runTurn(question, options); };

  /**
   * 处理"规划暂停卡"上的动作（CAP-02）。
   *
   * 与普通"重试"最本质的区别：**这不会重新发送需求**。它先查询服务端真实状态，
   * 再按检查点继续。因此不会重复调用模型、不会重复读取、不会重复规划。
   *
   * 顺序不可调换：
   *  1. 查状态 —— 界面上的 revision 是"当时那一刻"的，可能已经过期；
   *  2. 由状态推按钮 —— 不能继续就不硬续（服务端已完成/已停止时硬续会重复规划）；
   *  3. 按检查点继续 —— 只带身份与游标。
   *
   * 全程**不写任何单元格**：续跑只推进"规划"，写入仍要经过既有的预览与确认。
   */
  const resumePlanningFromPause = async (pause: NonNullable<typeof planningPause>) => {
    if (busy || invalidPlanningTurnsRef.current.has(pause.turnId) ||
        messages.some((message) => message.planningPause?.turnId === pause.turnId && message.planningPause.resolved)) return;
    /*
     * 幂等键：**一次点击 = 一个 requestId**。
     *
     * 在函数入口生成并放进 `resumeRequestRef`，整个"查状态 → 继续 → 超时重试"
     * 期间固定不变；只有用户在这张卡上**重新点击**（上一次已经结束）才换新身份。
     * 服务端的判重是 `(turn_id, request_id)`：沿用同一个值 → 重复领取只追加一次
     * 预算；每次重试换新值 → 每重试一次就多扣一份预算、多跑一轮规划。
     *
     * 同一张暂停卡的重复点击由 `busy`/`progress` 与下面的自增代次一起挡住，
     * 不依赖 UI 的禁用状态做正确性保证。
     */
    const resumeIdentity = resumeRequestRef.current;
    const requestId =
      resumeIdentity && resumeIdentity.turnId === pause.turnId
        ? resumeIdentity.requestId
        : newPlanningRequestId();
    resumeRequestRef.current = { turnId: pause.turnId, requestId };
    /*
     * 并发闸：同一时刻只允许一个续跑在飞。用 ref 而不是 `busy`（后者是异步
     * 渲染出来的状态，连点两下时两帧都读到 false）。
     */
    if (resumingPauseRef.current) return;
    resumingPauseRef.current = true;
    setContinuingTurnId(pause.turnId);
    /*
     * **接管权：点"继续"意味着这一次核对以本次为准。**
     *
     * 按暂停时自动触发的那次状态核对可能还在飞（它自己也持着
     * `planningInspectionRef`）。下面 running 分支收尾会转进
     * `reconcileInterruptedTurn`，而它开头有"已有核对在飞就不再开一次"的守卫 ——
     * 不让位的话，用户在"原任务仍在运行"这张卡上点"继续"就会撞上守卫**直接返回**，
     * 界面上等于点了没反应。这里先中止并摘掉它：它的 `finally` 会看到 ref 已换人，
     * 因此不会反过来清掉本次的进度提示、也不会把暂停卡重置成"尚未继续"。
     *
     * 摘掉的只是**状态查询**，不是授权校验：真正会写表的那一道（
     * `checkOriginalPlanningTarget` / 续跑前证据）在下面照旧对当下宿主重核。
     */
    planningInspectionRef.current?.abort();
    planningInspectionRef.current = null;

    const controller = new AbortController();
    const executionOwner = newExecutionId("resume");
    tagControllerOwner(controller, executionOwner);
    controllerRef.current = controller;
    setPauseRequestState(null);
    let activity: ReturnType<typeof beginActivity> | null = null;
    let activitySettled = false;
    try {
      setProgress({ stage: "answering", message: "正在核对服务端的任务进度…" });
      // 第 1 步：查真实状态。界面上的快照只用来定位任务，不用来当续跑凭据。
      let status: WpsPlanningTaskStatus;
      try {
        status = await waitForPlanning(inspectWpsPlanningTask(pause.turnId, controller.signal), controller.signal);
      } catch (error) {
        if (controller.signal.aborted) throw new DOMException("Aborted", "AbortError");
        /*
         * **断线不等于可以重发需求。** 查不到状态只说明"现在读不到"，不说明
         * 服务端没有这条任务 —— 恰恰相反，我们手上有一个 turnId，说明它很可能
         * 还在（或已跑完）。这里保持暂停卡不动，让用户稍后重试；
         * **绝不**退回一次重新规划（那会重复调用模型与工具，且把"接着跑"
         * 谎报成"重新跑了一遍"）。
         */
        setPlanningPause((current) => (current ? { ...current, view: { ...current.view, detail: "暂时读不到服务端进度，尚未做任何继续操作。" } } : current));
        pendingInspectionMessageRef.current = pause.sourceMessageId;
        setPlanningPauseReload({ turnId: pause.turnId, checking: false, blockedReason: "暂时读不到原任务状态。", canRecheck: true });
        setProgress(null);
        append({
          role: "system",
          text: "无法读取该任务的真实进度，未做任何继续操作。请确认本地 AI 服务仍在运行后重试。"
        });
        return;
      }
      if (controller.signal.aborted || controllerRef.current !== controller) {
        throw new DOMException("Aborted", "AbortError");
      }
      if (status.status === "completed") {
        await reconcileInterruptedTurn(
          pause.turnId, pause.originalPrompt,
          pauseCarryRef.current?.selectedSheetIndexes ?? selectedSheetIndexes,
          pause.sourceMessageId
        );
        return;
      }
      // 第 2 步：按钮只由结构化状态推导，不看助手文本、不凭界面上的旧快照。
      const view = wpsPlanningResumeView(status);
      if (status.status === "running") {
        /*
         * 原任务仍在服务端跑：把**已落盘**的进度接回原来那张过程卡，等它结束后
         * 交给同一张卡的结果处理链（`reconcileInterruptedTurn` 的 completed 分支）。
         * 绝不再发一次需求：那会重复规划、重复读表，而原任务还在跑。
         */
        setProgress({ stage: "answering", message: "原任务仍在服务端处理中，正在接回进度…" });
        const indexes = pauseCarryRef.current?.selectedSheetIndexes ?? selectedSheetIndexes;
        if (!(await preparePlanningBridge(pause.turnId, status, indexes, controller.signal, executionOwner))) return;
        const attached = await tailRunningTurn(
          pause.turnId,
          pause.sourceMessageId,
          () => !controller.signal.aborted && controllerRef.current === controller,
          controller.signal
        );
        if (controller.signal.aborted) throw new DOMException("Aborted", "AbortError");
        if (attached === "settled") {
          endBridgeOwnedBy(executionOwner);
          await reconcileInterruptedTurn(
            pause.turnId,
            pause.originalPrompt,
            indexes,
            pause.sourceMessageId
          );
          return;
        }
        setPlanningPause((current) => (current
          ? { ...current, view: { ...current.view, detail: "原任务仍在运行；进度已接回，稍后再核对一次即可。" } }
          : current));
        setPlanningPauseReload({ turnId: pause.turnId, checking: false,
          blockedReason: "原任务仍在服务端处理中。", canRecheck: true });
        setProgress(null);
        append({ role: "system", text: "原任务仍在服务端处理中；进度已接回，稍后再核对一次即可，不要重复发送需求。" });
        return;
      }
      if (!view.canContinue || view.action !== "continue_checkpoint") {
        setPlanningPause((current) => (current ? { ...current, view } : current));
        pendingInspectionMessageRef.current = pause.sourceMessageId;
        setPlanningPauseReload({ turnId: pause.turnId, checking: false, blockedReason: view.reason,
          // Running and completed tasks have already entered their recovery paths.
          canRecheck: false });
        append({
          role: "system",
          text: `该任务当前不能精确续跑：${view.reason}`
        });
        return;
      }
      /*
       * 第 2.5 步：**授权复核**（续跑前）。
       *
       * 续跑会把"原任务的结果"落到**当前**工作簿上。等待网络响应的这段时间里，
       * 用户完全可能切了工作簿 / 改了勾选范围 —— 那种情况下把旧结果套到新文件
       * 上，等于用 A 表的意图去写 B 表。这里对着**当下**的宿主重新核对，
       * 与"重新预览"走同一个判据（`checkAuthorizedTarget`），不另立一套。
       */
      const evidence = await preparePlanningBridge(pause.turnId, status,
        pauseCarryRef.current?.selectedSheetIndexes ?? selectedSheetIndexes, controller.signal, executionOwner);
      if (!evidence) return;
      // 第 3 步：按检查点继续。`resumeWpsPlanning` 内部再查一次状态并取
      // revision/游标；这里不再自己拼请求体，避免"界面拿一份旧 revision 去 CAS"。
      activity = resumeTaskActivity(pause.turnId, pause.sourceMessageId);
      const continued = await waitForPlanning(resumeWpsPlanning({
        turnId: pause.turnId,
        requestId,
        evidence,
        // 不传 modelId = 保留原模型。换模型必须另行明确确认，不能跟随全局选择。
        afterSequence: Math.min(receivedSequence(pause.turnId), status.lastSequence),
        signal: controller.signal,
        onProgress: (event) => {
          if (controller.signal.aborted || controllerRef.current !== controller) return;
          completePlanningAttachment(activity!);
          recordPlanningProgress(event);
          recordPlanningResult(event);
          updateActivity(activity!, readonlyActivityUpdate(event));
          setProgress({
            stage: "answering",
            stepId: event.stepId || "planning-continue",
            title: event.title,
            message: event.title || "正在继续规划"
          });
          // 续跑的进度里同样会带 `bridgeTool`：交给同一座桥去轮询/复核/读取/回执。
          sdkBridge.handleProgressEvent(event);
        }
      }), controller.signal);
      /*
       * 续跑可能被服务端**当场拒绝**（`outcome` 为 null）：已完成 / 已停止 /
       * 没有可验证现场。这时**不能**假装续跑过，必须把真实状态落到卡上。
       */
      if (!continued.outcome) {
        const refusedView = wpsPlanningResumeView(continued.status);
        setPlanningPause((current) => (current ? { ...current, view: refusedView } : current));
        pendingInspectionMessageRef.current = pause.sourceMessageId;
        setPlanningPauseReload({ turnId: pause.turnId, checking: false, blockedReason: refusedView.reason,
          canRecheck: continued.status.status === "running" || continued.status.status === "completed" });
        append({ role: "system", text: `未能继续该任务：${refusedView.reason}` });
        return;
      }
      if (controller.signal.aborted) throw new DOMException("Aborted", "AbortError");
      completePlanningAttachment(activity);
      // 继续成功后这条暂停卡退役：任务已经重新动起来了。
      setPlanningPause(null);
      updateMessages((items) => items.map((item) => item.planningPause?.turnId === pause.turnId
        ? { ...item, planningPause: { ...item.planningPause, resolved: true, clientLastSequence: receivedSequence(pause.turnId) } } : item));
      resumeRequestRef.current = null;
      const outcome = continued.outcome;      if (outcome.kind === "planning_pause") {
        // 又暂停了（例如再次触发预算闸）：落一张新卡，而不是复用旧卡。
        const nextView = wpsPlanningResumeView({
          turnId: outcome.turnId,
          status: outcome.status,
          revision: outcome.revision,
          lastSequence: outcome.lastSequence,
          canContinue: outcome.canContinue,
          pauseType: outcome.pauseType,
          reason: outcome.reason,
          // 再次暂停时同样要带上服务端刚给的继续资格：丢了它，新卡就会用
          // "暂停过 = 有现场"的旧判据画出一个点不动的「继续分析」。
          continuation: outcome.continuation
        });
        setPlanningPause({
          turnId: outcome.turnId,
          payload: outcome,
          view: nextView,
          originalPrompt: pause.originalPrompt,
          sourceMessageId: append({
            role: "assistant",
            text: "",
            planningPause: {
              turnId: outcome.turnId,
              pauseType: outcome.pauseType,
              reason: outcome.reason,
              canContinue: outcome.canContinue,
              action: nextView.action,
              actionLabel: nextView.actionLabel,
              detail: nextView.detail,
              clientLastSequence: receivedSequence(outcome.turnId)
            }
          })
        });
        // 再次暂停仍沿用原任务的授权与需求；不能在 finally 清掉。
        return;
      }
      /*
       * **续跑结果闭环**（CAP-02 第 1 条）。
       *
       * `plan` / `clarification` / `local_plan_rule` 与一次性轮次是同一个联合，
       * 必须走**同一条处理链**（`applyTurnResult`）：计划进预览、澄清进可填写
       * 卡片、本地规则按当前授权读取计算。
       *
       * 此前这三支一律报"请重新发送需求" —— 那不但让用户白做一遍，还把
       * "服务端已经算好了"说成"算不了"。更危险的是：如果哪天有人把"重发"
       * 实现成自动重跑，就会变成一次**未确认的重新规划**。
       *
       * 全程不写单元格：`applyTurnResult` 里没有任何宿主写动作，计划只到
       * "预览待确认"为止（与一次性通道完全一致）。
       */
      const carry = pauseCarryRef.current;
      const carryIndexes = carry?.selectedSheetIndexes ?? [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)];
      try {
        await applyTurnResult(outcome as WpsTurnResult, {
          activity,
          text: carry?.originalPrompt ?? pause.originalPrompt,
          turnSheetIndexes: carryIndexes,
          clarificationRound: carry?.clarificationRound ?? 0,
          images: []
        });
        activitySettled = true;
        pauseCarryRef.current = null;
      } catch (error) {
        endActivity(activity, error instanceof TargetConfirmationRequiredError ? "awaiting_input" : "failed");
        activitySettled = true;
        throw error;
      }
    } catch (error) {
      if (error instanceof TargetConfirmationRequiredError) {
        append({
          role: "system",
          text: "数据范围在继续过程中发生了变化，已暂停本次继续；原表未写入。请重新确认数据范围。"
        });
      } else if (error instanceof DOMException && error.name === "AbortError") {
        if (pauseRequestedRef.current === controller) {
          /*
           * **先把自己的三样资源收干净，再做后台状态核对 —— 而且不等它回来。**
           *
           * 顺序反过来就会漏：核对第一件事是把 `controllerRef` 换成自己的
           * controller，等它返回时"这一轮还归我吗"已经不成立 —— 活动于是永远停在
           * 运行中（卡一直转、耗时继续涨），桥也没人收（迟到的读表进度还能把按钮
           * 点亮）。
           *
           * 更不能"等核对回来再收"：核对是一次真实的状态读取，它慢、它失败、它
           * 卡住都只说明**服务端状态未确认**，与"本端是否已经停下"无关。用户按下
           * 暂停那一刻，本端就该立刻停。所以这里按归属收尾（`settleExecution`），
           * 后面 `finally` 里那次是同一条判据的**幂等**兜底。
           */
          settleExecution({
            owner: executionOwner,
            activity,
            activitySettled,
            outcome: "awaiting_input"
          });
          activitySettled = true;
          const carry = pauseCarryRef.current;
          const text = carry?.originalPrompt ?? pause.originalPrompt;
          const indexes = carry?.selectedSheetIndexes ?? selectedSheetIndexes;
          const cardId = rememberPlanningInterruption(pause.turnId, text,
            carry?.selectedSheetIndexes ?? selectedSheetIndexes, carry?.clarificationRound ?? 0, pause.sourceMessageId);
          pauseRequestedRef.current = null;
          await reconcileInterruptedTurn(pause.turnId, text, indexes, cardId, false);
        } else append({ role: "system", text: "已停止继续；未确认的计划不会修改工作表。" });
      } else {
        append({ role: "system", text: turnErrorMessage(error) });
      }
    } finally {
      /*
       * 三样资源**分别**按归属清理（判据在 `planExecutionTeardown`，两端同一份）：
       *  - 活动：本次执行开的就由本次收，哪怕 controller 已被核对接管；
       *  - controller：只有仍指向自己那个对象时才清；
       *  - 桥：只有自己仍是桥主人才收（保留绑定给结果卡按引用取值）。
       *
       * 重复收尾是安全的：已收的活动、已换人的 controller、已收的桥都判 false。
       * **旧执行动不到新执行**：新一轮有自己的身份，三者都不匹配。
       */
      settleExecution({
        owner: executionOwner,
        activity,
        activitySettled,
        outcome: "awaiting_input"
      });
      resumingPauseRef.current = false;
      setContinuingTurnId((current) => current === pause.turnId ? null : current);
      updateMessages((items) => items.map((item) => item.id === pause.sourceMessageId && item.planningPause
        ? { ...item, planningPause: { ...item.planningPause, clientLastSequence: receivedSequence(pause.turnId) } } : item));
    }
  };

  const stagePendingPlan = async () => {
    if (!(await targetGuard.check())) return;
    if (!pendingPlan || busy || planEditDraft !== null || stagedReview) return;
    setWriteProgress("正在创建独立副本并核对结果…");
    try {
      const currentSheetIndexes = [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)].sort((a, b) => a - b);
      const previewSheetIndexes = pendingPlan.authorization.sheets.map((sheet) => sheet.index).sort((a, b) => a - b);
      if (JSON.stringify(currentSheetIndexes) !== JSON.stringify(previewSheetIndexes)) {
        throw new Error("授权工作表范围在预览后发生变化，请重新预览。");
      }
      const application = getWpsApplication();
      let staged: WpsNativeComparison;
      let usedFixedFallback = false;
      try {
        staged = stageWpsNativeComparison(application, pendingPlan);
      } catch (error) {
        // An edited plan may carry a user-edited expected matrix. It is no longer proof of
        // the deterministic mapping result, so only the unedited local plan may auto-fallback.
        const fallback = error instanceof WpsFormulaTrialError && !planRevised
          ? fixedValueFallbackFromFormula(pendingPlan.plan)
          : null;
        if (!fallback) throw error;
        const fixedPrepared = prepareWpsPlan(application, currentSheetIndexes, fallback, pendingPlan.sourceSnapshots, { requireSourceAnchor: pendingPlan.sourceAnchorRequired });
        staged = stageWpsNativeComparison(application, fixedPrepared);
        setPendingPlan(fixedPrepared);
        usedFixedFallback = true;
        if (pendingPlanMessageId) {
          updateMessages((current) => current.map((message) =>
            message.id === pendingPlanMessageId && message.planPreview
              ? { ...message, planPreview: { ...message.planPreview, title: fixedPrepared.plan.title, revised: true } }
              : message
          ));
        }
      }
      setStagedReview(staged);
      append({
        role: "assistant",
        text: `${usedFixedFallback ? "公式试算未能通过计算值核对，已关闭试写副本并改用固定值重新生成。" : ""}已在右侧独立副本生成「${staged.plan.plan.title}」的结果，并标记 ${staged.changedCells.length} 个变化单元格。左侧原表尚未写入；请并排核对后选择保存或放弃。`
      });
    } catch (error) {
      append({ role: "system", text: turnErrorMessage(error) });
    } finally {
      setWriteProgress(null);
    }
  };

  const discardStagedPlan = () => {
    if (!stagedReview || busy) return;
    /*
     * 「放弃变更」是**终态**：确认框与写回入口立刻撤下，同一份副本此后不可能再写回。
     *
     * 副本清理失败**不恢复审阅状态** —— 它只作为**次要**状态跟在"计划已取消"后面，
     * 绝不冲淡主结论。
     */
    const cleanup = wpsCopyCleanupState("discarded", discardWpsNativeComparison(stagedReview));
    const outcomeText = wpsReviewOutcomeText(wpsTerminalCard({
      outcome: "discarded",
      scope: pendingPlan ? {
        locations: reviewScopeLocations(pendingPlan.plan.actions),
        changes: reviewChangeSummary(pendingPlan.plan.actions),
        totalCells: pendingPlan.totalCells
      } : null,
      copy: cleanup.copy,
      staged: stagedReview,
      developmentWrite
    }));
    setStagedReview(null);
    markPlanPreview("cancelled");
    /*
     * 用户显式放弃：这是**终态**，而且是在任何宿主写入之前。
     * 置成 `unavailable` 之后，重开面板也不会再把它摆成"可以继续"。
     */
    advanceTaskResume(pendingPlanMessageId, {
      phase: "preview_ready",
      blocker: "none",
      resumability: "unavailable",
      writeState: "not_written"
    });
    setPendingPlan(null);
    setPendingPlanMessageId(null);
    append({ role: "assistant", text: outcomeText });
  };

  /**
   * 写回入口的**实时**事实。
   *
   * `confirmPendingPlan` 是渲染期闭包：旧闭包会一直看到它诞生那一刻的
   * `pendingPlan` / `stagedReview` / `reviewAcknowledged`。放弃之后若有人拿着
   * 放弃之前那个闭包再调用一次（已排队的点击、快捷键、程序直调），闭包里的
   * "计划还在"就是**过期事实** —— 拿它当判据会把已经放弃的计划写进原表。
   *
   * 所以写回闸门**只读这里的实时值**，闭包变量一个字都不参与判断。
   * 每次渲染都会刷新，包括放弃（计划与副本被清空）那一次。
   *
   * `target` 是这份现场的**身份三元组**（同一次审阅 / 同一计划版本 / 同一原工作簿），
   * 供异步授权检查回来后复核用：光看"当前有审阅"会把**新副本**的确认状态借给旧请求。
   */
  const liveWriteBack = useRef({
    acknowledged: reviewAcknowledged,
    plan: pendingPlan,
    review: stagedReview,
    editingPlan: planEditDraft !== null,
    busy: baseBusy,
    executionMode,
    target: null as ReviewWriteBackTarget | null,
    directTarget: null as DirectExecutionTarget | null
  });
  liveWriteBack.current = {
    acknowledged: reviewAcknowledged,
    plan: pendingPlan,
    review: stagedReview,
    editingPlan: planEditDraft !== null,
    busy: baseBusy,
    executionMode,
    // 计划与副本都在场才谈得上"同一份现场"；任一缺席 ⇒ 身份为 null ⇒ 一律按已变化处理。
    target: stagedReview && pendingPlan
      ? {
          review: stagedReview,
          plan: pendingPlan,
          workbookIdentity: String(pendingPlan.authorization.workbookFullName ?? "")
        }
      : null,
    // 直接执行的身份里**没有审阅**（没有副本），但有执行方式：切到副本比对再回来时前提已变。
    directTarget: pendingPlan
      ? {
          review: null,
          plan: pendingPlan,
          workbookIdentity: String(pendingPlan.authorization.workbookFullName ?? ""),
          mode: executionMode
        }
      : null
  };

  /**
   * 一次写回正在跑。
   *
   * `busy` 是渲染期派生值：同一 tick 里的第二次调用读到的还是 false
   * （`setWriteProgress` 要等下一次渲染才生效）。这个 ref 同步置位，
   * 重复派发（双击、旧闭包连点）在这里就断掉，写不了两次。
   */
  const writeBackInFlight = useRef(false);

  /**
   * 异步等待回来之后，这一次写回请求还能不能继续？
   *
   * 判两件事，缺一不可：
   *  1. 现场还在（勾选未撤、计划与副本都在场）；
   *  2. **还是同一份现场** —— 同一次审阅、同一计划版本、同一原工作簿
   *     （`reviewWriteBackTargetChanged`）。
   *
   * 只判第 1 条是不够的：用户放弃之后又生成了**新副本**并勾了新的确认框时，现场
   * 同样"在"，但那是**别人的**确认 —— 旧请求借它继续写回，结果是用户核对过的那一份
   * 从来没被写，没核对过的那一份反而被写进原表。
   *
   * 导出是为了让这条判据能被**直接**判定：面板里"旧请求在途期间换副本"这条复合路径
   * 目前被 `busy` 闸门整个挡在界面之外（`busy` 含 `targetGuard.checking`：写回等待期间
   * 放弃 / 生成副本 / 发新任务的入口全是禁用的，`runTurn` 自己也挡 `busy`），端到端
   * 造不出来。所以这段防护是**纵深防御**：等哪天 `busy` 不再覆盖这个窗口（或出现
   * 已排队的点击 / 快捷键 / 程序直调），判据已经在位并被单独判定。
   */
  /**
 * 一次 WPS 写入。
 *
 * `direct = true` 表示「直接执行到原表」：**只**跳过副本比对这一步 ——
 * 不做 `verifyWpsNativeComparison`、不做与副本的最终对照、不关任何副本。
 * 授权闸、执行前计划复核、授权范围复核、恢复标记、服务端账本首写前登记、
 * 动作回读、撤销入口、失败证据分级**全部共用同一段代码**，这是"直接执行没跳过
 * 任何守卫"唯一可靠的证据形式：守卫只有一份实现，副本路径也没有多跑一步。
 *
 * 参数化而不是复制一份函数，就是为这个：复制出来的"直接执行版"迟早会少一两道闸，
 * 而且没人能一眼看出少在哪。
 */
const confirmPendingPlan = async (direct = false, afterExecutionReview = false): Promise<WpsUndoRecord | undefined> => {
    /*
     * 写回入口的**唯一闸门**（见 `wpsWriteBackBlockReason`）。
     *
     * 本地判据先跑、授权检查后跑：
     *  - 「没勾选就是没核对」——确认框是真闸，不是装饰；按钮已禁用，这里再挡一次，
     *    旧闭包、快捷键、测试直调都绕不过去。
     *  - 「放弃 / 写回已成功」是终态：计划与副本都不在场，写回入口整体失效 ——
     *    这条路径连授权确认都不该发起。
     */
    /*
     * **"忙"的口径按调用方区分。**
     *
     * 结果审阅通道（`afterExecutionReview`）本身就是"pending"的制造者：它置
     * `resultReviewState.pending = true` 之后才发起这次执行。若这里把那个 pending
     * 也算成"忙"，闸门就会拦下自己发起的写入 —— 而它拦不拦得住取决于这之间有没有
     * 恰好发生一次重渲染，于是表现为**时灵时不灵**：同一份代码，有的用例能写、
     * 有的用例报"执行未完成"。真实用户那边就是"有时执行不了、重试一次又好了"。
     * 所以这条通道只认自己的在途锁 `writeBackInFlight`；界面忙态由调用方（结果是
     * 否已有审阅在跑）在**发起点**判定，不在这里和它互相否决。
     */
    const gateBusy = afterExecutionReview
      ? writeBackInFlight.current
      : (liveWriteBack.current.busy || writeBackInFlight.current);
    const block = direct
      ? wpsDirectWriteBackBlockReason({
        hasPlan: liveWriteBack.current.plan !== null,
        editingPlan: liveWriteBack.current.editingPlan,
        busy: gateBusy
      })
      : wpsWriteBackBlockReason({
        acknowledged: liveWriteBack.current.acknowledged,
        hasPlan: liveWriteBack.current.plan !== null,
        hasReview: liveWriteBack.current.review !== null,
        editingPlan: liveWriteBack.current.editingPlan,
        busy: gateBusy
      });
    if (block === "not_acknowledged") {
      append({ role: "system", text: "请先勾选确认已核对副本，再写回原表；原表尚未写入。" });
      return;
    }
    if (block) return;
    /*
     * **发起点**的现场身份：这一次请求只认这一份审阅 / 这个计划版本 / 这一本原工作簿。
     * 必须在 `await` **之前**取，`await` 之后再取就等于"跟着新副本走"了。
     */
    // 发起点的现场身份：**按本次走的那条路径**取，不能两条都取再挑一条用 ——
    // 直接执行没有副本，取副本身份会得到 null，回来后一律判"已变化"，直接执行永远发不出去。
    const issuedTarget = direct
      ? liveWriteBack.current.directTarget
      : liveWriteBack.current.target;
    if (!(await targetGuard.check())) return;
    /*
     * 授权检查是异步的：等它回来时用户可能已经点了「放弃变更」，甚至放弃之后又生成了
     * 一份**新副本**并勾了新的确认框。
     *
     * 所以这里做两件事，缺一不可：
     *  1. 再判一次实时事实（放弃之后计划与副本都不在场）；
     *  2. **核对身份三元组**：仍是同一次审阅、同一计划版本、同一原工作簿。
     *     只判"当前有审阅"是不够的 —— 新副本的确认状态会被旧请求借走：用户核对过的
     *     那一份从来没被写，没核对过的那一份反而被写进了原表。
     */
    const now = liveWriteBack.current;
    if (!now.plan) return;
    // 副本路径：勾选未撤、计划与副本都在场（顺带把类型收窄）。
    // 直接执行：没有副本，只需计划在场 —— 但**执行方式必须仍是 direct**。
    if (direct) {
      // ② 直接执行的身份：同一计划版本、同一原工作簿、**同一种执行方式**。
      //    方式进身份是因为在途期间切到副本比对再回来时，前提已经变了。
      if (!afterExecutionReview && now.executionMode !== "direct") return;
      if (!wpsDirectWriteBackStillCurrent(issuedTarget as DirectExecutionTarget | null,
        { target: now.directTarget })) return;
    } else {
      if (!now.acknowledged || !now.review) return;
      // ② 副本路径的身份三元组：同一次审阅、同一计划版本、同一原工作簿。
      if (!wpsWriteBackStillCurrent(issuedTarget, now)) return;
    }
    if (writeBackInFlight.current) return;
    writeBackInFlight.current = true;
    const plan = now.plan;
    // 直接执行没有副本会话：`review` 为 null，后续所有副本专属步骤都据此跳过。
    const review = direct ? null : now.review!;
    const activity = beginActivity({
      id: "preflight",
      label: "复核计划与授权范围",
      phase: "preflight",
      detail: `${plan.actions.length} 个动作，目标 ${plan.totalCells} 个单元格；执行前重新核对工作簿与授权。`
    });
    setWriteProgress("正在保存执行前恢复点…");
    // 失败路径也要拿到 executionId 才能把明细挂到提示消息上，而那时 record 拿不到。
    let reviewExecutionId: string | null = null;
    // 服务端持久账本的执行身份与写入权。声明在 try 外，失败路径要据此挂起账本。
    let operationId: string | null = null;
    /*
     * **本次执行留下的恢复记录**，要在函数末尾返回给调用方（结果审阅通道）。
     *
     * 少了这个返回值，审阅通道拿到 `undefined` 就判"执行未完成"—— 而表已经写了、
     * 恢复点也在。用户看到的是"执行未完成"，实际状态却是"已写入且可撤销"，
     * 这是最糟的一种报告：它让人要么重复执行，要么不敢撤销。
     */
    let completedRecord: WpsUndoRecord | null = null;
    let wpsLedgerState: Awaited<ReturnType<typeof registerWpsWriteIntent>> | null = null;
    /*
     * 宿主观测（A1-E3）：执行器在每个动作窗口里数**真的进入宿主写 API 的调用**
     * （`wpsHostWriteCounter`），这里先收着，等账本 finish/suspend **之后**再补报。
     * 顺序是有意的：观测不能插在"写入 → 账本提交"之间，否则上报失败会变成新的
     * 中断窗口。失败也没关系 —— 少一条证据，不改执行结论。
     */
    const hostWriteAudits: HostWriteAudit[] = [];
    const flushHostWriteAudits = () => {
      if (!operationId || hostWriteAudits.length === 0) return;
      void reportWpsHostWrites(operationId, {
        ownerToken: wpsLedgerState?.ownerToken ?? null,
        entries: [...hostWriteAudits]
      }).catch(() => {
        // 观测失败只少一条证据：不重试、不打断、不改写任何执行结论。
      });
    };
    try {
      // 副本未被改动这一项只属于副本路径；直接执行没有副本可核对。
      if (review) verifyWpsNativeComparison(getWpsApplication(), review);
      const currentSheetIndexes = [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)].sort((a, b) => a - b);
      const previewSheetIndexes = plan.authorization.sheets.map((sheet) => sheet.index).sort((a, b) => a - b);
      if (JSON.stringify(currentSheetIndexes) !== JSON.stringify(previewSheetIndexes)) {
        throw new Error("授权工作表范围在预览后发生变化，请重新预览并确认。");
      }
      /*
       * **先同步落盘**"写入结果未知"，确认写进去了才动原表。
       *
       * 用同步路径而不是 `advanceTaskResume`（后者经 `useEffect` 异步落盘）：
       * 这里要的是顺序保证。落盘失败就**不写**——没有持久证据的写入一旦中断，
       * 用户看到的信息仍然是"原表未写入"。
       */
      const markerPersisted = persistTaskResumeNow(pendingPlanMessageId, {
        phase: "executing",
        blocker: "none",
        resumability: "reconcile",
        writeState: "unknown"
      });
      /*
       * 落盘失败就**当场放弃这次写入**。
       *
       * 这里不能"先写着再说"：一旦写到一半断线，本机既没有恢复标记也没有恢复点，
       * 卡片却还在说"原表未写入"。没有证据的写入比不写更糟 —— 用户会照着这句
       * 假话去手动改原表，把真正的现场彻底搞乱。
       */
      if (!markerPersisted) {
        endActivity(activity, "failed");
        append({
          role: "system",
          text: "执行前恢复标记无法写入本机存储，已取消本次写入；原表未被修改。请清理加载项存储空间后重新预览并确认。"
        });
        return;
      }
      /*
       * CAP-03 首写前持久登记：把执行身份、授权与**恢复描述**持久存到服务端账本，
       * 拿到有效写入权（ownerToken）才允许动原表。登记失败、身份缺失或服务端不可用
       * 一律中止 —— 没有持久证据的写入一旦中断，重开后无法判断"写没写"。
       *
       * 描述**不含目标值矩阵**（`writeValues.values` 等）：那些内容只停在本机内存，
       * 不随登记请求出站，也不落进服务端账本。
       *
       * 身份用 `wps-path:` 前缀（已保存工作簿的小写磁盘完整路径），不与 Office 的
       * `url:` 共享命名空间。同名异路径靠完整路径区分，表替换靠授权表名 + 索引。
       */
      operationId = `wps-execution-${newTaskId().slice("task-".length)}`;
      const workbookIdentity = wpsWorkbookIdentity(plan.authorization.workbookFullName);
      wpsLedgerState = null;
      if (workbookIdentity) {
        try {
          wpsLedgerState = await registerWpsWriteIntent({
            operationId,
            // 传**已确认的 WpsPreparedPlan**：描述投影在 `wpsExecutionLedger` 里
            // 一次性完成，登记请求不带任何单元格矩阵。
            prepared: plan,
            authorizedSheetNames: plan.authorization.sheets.map((sheet) => sheet.name),
            workbookIdentity,
            taskId: carried().taskId
          });
        } catch (error) {
          // 登记失败（服务端不可用 / 身份或计划不被接受）→ 中止写入，不静默继续。
          endActivity(activity, "failed");
          append({
            role: "system",
            text: `无法持久登记本次写入现场，已取消写入；原表未被修改。${turnErrorMessage(error)}`
          });
          return;
        }
        if (!wpsLedgerState.ownerToken) {
          endActivity(activity, "failed");
          append({
            role: "system",
            text: "服务端没有授予本次 WPS 写入权，已停止执行。"
          });
          return;
        }
        const pendingSaved = savePendingWpsExecution({
          schemaVersion: 1,
          operationId,
          conversationId: chatHistory.activeConversationId,
          activityMessageId: pendingPlanMessageId ?? `activity-${operationId}`,
          workbookIdentity,
          taskId: carried().taskId,
          createdAt: new Date().toISOString()
        });
        if (!pendingSaved) {
          // 本地发现索引写不进去：重开后无法自动定位这次执行。宁可中止，不留
          // "已经登记到服务端、却在本机找不到入口"的半截状态。
          endActivity(activity, "failed");
          append({
            role: "system",
            text: "无法保存本次写入的本地恢复索引，已停止执行以避免丢失恢复入口。"
          });
          return;
        }
      } else {
        /*
         * 拿不到稳定身份（未保存 / 非 `.xlsx`）就不写。
         *
         * 正常流程到不了这里 —— `prepareWpsPlan` 已经要求已保存的 `.xlsx` ——
         * 但守卫不能只靠上游：一旦上游放宽，这里若"没身份也照写"，就会变成
         * 一次**没有任何持久证据**的原表写入，中断后无法判断写没写。
         */
        endActivity(activity, "failed");
        append({
          role: "system",
          text: "当前工作簿没有可用于跨关闭核对的稳定身份，已取消写入；原表未被修改。请先保存为 .xlsx 再执行。"
        });
        return;
      }
      // The persisted discovery index now belongs to this replacement execution.
      setWpsRecovery(previous => previous && previous.operationId !== operationId &&
        previous.taskId === carried().taskId ? null : previous);
      bindRecoveryProbe((import.meta.env.DEV || import.meta.env.VITE_ENABLE_RECOVERY_PROBE === "true"), workbookIdentity, operationId, plan.plan.actions, Boolean(wpsLedgerState?.ownerToken && wpsLedgerState.descriptor.resumeContext));
      const registeredOperationId = operationId;
      const record = await executeWpsPlan(getWpsApplication(), plan, (update) => {
        setWriteProgress(update.message);
        updateActivity(activity, executionActivityUpdate(update));
      }, (reviews, executionId) => {
        // `executionId` 一定要记下来：撤销按钮与变更明细都按它归属，
        // 而**变化为空**的那次执行同样可能留下恢复点（例如整块写回原值），
        // 只按"有没有明细"记账会让那张卡拿不到自己的执行 id。
        reviewExecutionId = executionId;
        if (reviews.length === 0) return;
        changeReviewsRef.current.set(executionId, reviews);
      }, wpsLedgerState
        ? async (phase, checkpoint) => {
            if (!wpsLedgerState?.ownerToken) {
              throw new Error("WPS 写入权已经失效，已停止写入。");
            }
            if (phase === "before") {
              wpsLedgerState = await startWpsAction(registeredOperationId, {
                expectedRevision: wpsLedgerState.revision,
                ownerToken: wpsLedgerState.ownerToken,
                actionIndex: checkpoint.actionIndex,
                beforeSnapshot: {
                  sheetName: checkpoint.sheetName,
                  range: checkpoint.range,
                  type: checkpoint.type,
                  beforeFingerprint: checkpoint.beforeFingerprint
                }
              });
            } else {
              interruptRecoveryProbe((import.meta.env.DEV || import.meta.env.VITE_ENABLE_RECOVERY_PROBE === "true"), registeredOperationId, checkpoint.actionIndex);
              wpsLedgerState = await completeWpsAction(registeredOperationId, {
                expectedRevision: wpsLedgerState.revision,
                ownerToken: wpsLedgerState.ownerToken,
                actionIndex: checkpoint.actionIndex,
                checkpoint: {
                  sheetName: checkpoint.sheetName,
                  range: checkpoint.range,
                  type: checkpoint.type,
                  beforeFingerprint: checkpoint.beforeFingerprint,
                  afterFingerprint: checkpoint.afterFingerprint
                }
              });
            }
          }
        : undefined,
        operationId, {taskId: wpsLedgerState?.taskId},
        (entry) => { hostWriteAudits.push(entry); });
      let comparison: "matched" | "different" | "unavailable" = "unavailable";
      let comparisonFailed = false;
      if (review) {
        try {
          comparison = compareCommittedWpsResultToCopy(review);
        } catch {
          comparisonFailed = true;
        }
      }
      updateActivity(activity, {
        id: "verified-result",
        label: !review
          ? "写入与回读核对完成"
          : comparison === "matched" ? "写入与回读核对完成" : "写入完成，计算结果仍需核对",
        phase: "verifying",
        note: !review
          ? "本次直接执行未生成副本，因此没有与副本比对；动作回读与恢复点已核对。"
          : comparison === "matched"
            ? "目标与审阅副本的公式和值均匹配；恢复点仅保证当前加载项会话内可用。"
            : "动作回读已完成，但无法证明原表计算值与审阅副本一致；请检查并按需撤销。",
        facts: [
          { label: "操作", value: `${record.actions.length} 步` },
          { label: "已核对", value: `${plan.totalCells} 个单元格` },
          { label: "恢复", value: SESSION_RECOVERY_LABEL }
        ]
      });
      endActivity(activity, "completed");
      setUndoRecord(record);
      completedRecord = record;
      // 写入与回读都完成：提交服务端完成回执，释放恢复索引。回执失败只提示，不推翻
      // 已经发生的写入 —— 服务端账本仍保留这次执行，重开后可核对或补记。
      if (operationId && wpsLedgerState?.ownerToken) {
        try {
          wpsLedgerState = await finishWpsExecution(operationId, {
            expectedRevision: wpsLedgerState.revision,
            ownerToken: wpsLedgerState.ownerToken,
            result: { actionCount: record.actions.length, cellCount: plan.totalCells }
          });
          clearPendingWpsExecution(operationId);
        } catch (error) {
          append({ role: "system", text: `写入已完成，但完成回执提交失败；重开面板后可核对本次执行。${turnErrorMessage(error)}` });
        }
        flushHostWriteAudits();
      }
      if (!review) {
        // 直接执行：如实说"没有副本可比"，而不是让"与副本不同"这类话缺席 ——
        // 缺席会被读成"核对过了"。
        append({ role: "system", text: "原表已按计划写入并完成动作回读；本次直接执行未生成副本，未与副本比对。" });
      } else if (comparisonFailed) {
        append({ role: "system", text: "原表已按计划写入，但无法完成与审阅副本的最终逐格对照；请检查原表。" });
      } else if (comparison === "different") {
        append({ role: "system", text: "原表已按计划写入并通过动作回读，但计算结果与右侧副本不同。请检查原表；如不符合预期，可使用本次恢复点撤销。" });
      } else if (comparison === "unavailable") {
        append({ role: "system", text: "原表已按计划写入，但宿主未提供完整显示值，无法证明计算结果与审阅副本完全相同。" });
      }
      /*
       * 业务结论已由上面的执行结果卡说了；这里**只**陈述副本状态，独立成一条。
       * 副本关不掉绝不能把上面那条"原表已按计划写入"改写掉 —— 两者是两件事。
       */
      if (review) {
        append({ role: "assistant", text: wpsCopyStatusLine(wpsCopyCleanupState("verified", discardWpsNativeComparison(review))) });
      }
      /*
       * 写入与回读都完成了：这条任务到此**结束**，不再是"可以继续的任务"。
       * 置成 `unavailable` 而不是删记录 —— 保留"这件事做过、结果已确认"的事实，
       * 同时让它不再出现在中断卡里（已完成的任务不该再被兜售"继续"按钮）。
       */
      advanceTaskResume(pendingPlanMessageId, {
        phase: "executing",
        blocker: "none",
        resumability: "unavailable",
        writeState: "confirmed_result"
      });
      setStagedReview(null);
      /*
       * 执行完成之后**保留这份计划的摘要**。
       *
       * 为什么不能直接丢掉：新流程里计划卡不再是"确认入口"，但它承载着用户事后
       * 复核所需的**事实**——采用了哪些口径、有哪些补充说明、格式动作具体改什么、
       * 行高列宽是"影响整行/整列"。执行一结束就把这些抹掉，用户看到的是
       * "已经改了"，却看不到"按什么口径改的、改的是格式还是值"。
       */
      if (pendingPlanMessageId) setExecutedPlan({ messageId: pendingPlanMessageId, plan, revised: planRevised });
      markPlanPreview("executed");
      setPendingPlan(null);
      setPendingPlanMessageId(null);
      /*
       * 「完全通过」的判据按**执行方式**取，两条都要求动作回读一致：
       *  - 副本路径：还要与副本计算值一致（多一道比对才敢说得更满）；
       *  - 直接执行：没有副本可比，只以动作回读为准 —— 少了那一层就说"未获证实"，
       *    会让用户以为直接执行跳过了验证，而它并没有。
       * 两种情况下失败 / 部分完成 / 回读未过都记成 `unverified`。
       */
      const outcome: WpsExecutionOutcome =
        !comparisonFailed && (review ? comparison === "matched" : true) ? "readback_passed" : "unverified";
      // 只有「完全通过」（回读一致且与副本计算值一致）才具备保存为工具的资格：
      // 把已实际验证的计划版本连同其执行身份一并记下，供结果卡里的「保存为工具」使用。
      // 失败 / 部分完成 / 未验证的写入绝不写入这里，也就拿不到保存入口。
      if (outcome === "readback_passed" && !afterExecutionReview) {
        verifiedPlanForSaveRef.current = {
          plan: plan.originalPlan ?? plan.plan,
          sourceSheetNames: plan.authorization.sheets.map((sheet) => sheet.name),
          executionId: record.executionId
        };
      }
      append({
        role: "assistant",
        text: outcome === "readback_passed"
          ? `已执行「${plan.plan.title}」，并完成目标区域回读核对。`
          : `已执行「${plan.plan.title}」，但最终计算值未获证实；请检查原表，必要时使用恢复点撤销。`,
        verification: {
          executionId: record.executionId,
          actionCount: plan.actions.length,
          cellCount: plan.totalCells,
          // `status` 只表达"这张卡记录的是哪次已执行写入 / 后来被撤销了"，
          // 写入本身好不好由 `outcome` 说，别拿它冒充"验证通过"。
          status: "recorded",
          outcome,
          targets: plan.actions.map((action) => ({
            type: action.type,
            sheetName: redactCredentials(action.sheetName),
            range: action.range
          }))
        }
      });
      return record;
    } catch (error) {
      // CAP-03：写失败时把服务端账本挂起（有在途动作则 needs_reconciliation）。
      // 挂起失败只提示，不掩盖已经发生的写入 —— 账本里 in_flight 状态仍可核对。
      if (operationId && wpsLedgerState?.ownerToken && wpsLedgerState.status === "running") {
        try {
          wpsLedgerState = await suspendWpsExecution(operationId, {
            expectedRevision: wpsLedgerState.revision,
            ownerToken: wpsLedgerState.ownerToken,
            code: "WPS_EXECUTION_INTERRUPTED",
            message: "WPS 写入中断，需要核对现场。"
          });
        } catch {
          // 挂起失败：服务端账本保持 running + 在途动作，重开后核对路径仍能接手。
        }
        flushHostWriteAudits();
      }
      // WPS reopens an existing pane by setting Visible=true; React need not remount.
      // Publish the recovery state now, after suspension, using the same read-only
      // identity/authorization/target verdict as reload. Never automatically reconcile.
      if (operationId && wpsLedgerState) {
        refreshWpsRecoveryVerdict(operationId, wpsLedgerState.taskId, wpsLedgerState);
      }
      let latest: WpsUndoRecord | null = null;
      /** 读取宿主存储本身失败 —— 与"读到了但没有记录"是两件事，不能合并。 */
      let evidenceReadFailed = false;
      try {
        latest = latestWpsUndo(getWpsApplication());
        setUndoRecord(latest);
      } catch {
        setUndoRecord(null);
        evidenceReadFailed = true;
      }
      /*
       * 归属只认 `executionId`。曾经的 `planId` + 时间戳只是"像是这一次"的旁证：
       * 同一份计划执行第二次时，它也成立，会把恢复点算到上一次头上。
       */
      const reviewId = reviewExecutionId;
      /*
       * 证据读的是**按 executionId 查到的原始记录**，而不是 `latestWpsUndo`：
       * 后者只返回"当前还能撤销"的那一条，会把 `restored`（执行器已把区域放回写前）
       * 与 `ready`/`applying` 中间态一起过滤掉 —— 而 `restored` 恰恰是"未开始写入"
       * 唯一的正面证据。
       */
      let rawRecord: WpsUndoRecord | null = null;
      if (reviewId && !evidenceReadFailed) {
        try {
          rawRecord = wpsUndoRecordFor(getWpsApplication(), reviewId);
        } catch {
          evidenceReadFailed = true;
        }
      }
      /*
       * 失败之后**只有拿到正面证据**才敢断言原表的状态：
       * - `restored` → 未开始写入；
       * - `verified` → 写完并回读过；
       * - 部分写入 / 现场冲突 → 需要核对；
       * - 读失败、回执丢失、找不到记录、中间态 → **一律未知**。
       * 把任何一项折成"未写入"，就是对着可能已被改过的原表说假话。
       */
      const evidence: WpsWriteEvidence = classifyWpsWriteEvidence({
        executionId: reviewId,
        record: rawRecord,
        readFailed: evidenceReadFailed
      });
      const partiallyWritten = evidence.kind === "needs_review";
      if (pendingPlanMessageId) {
        // 这条不需要抢在某个动作之前，走常规的异步落盘即可。
        // 一并记下执行身份：会话内还能靠它做只读核对（CAP-03 的跨关闭账本还没实现）。
        advanceTaskResume(pendingPlanMessageId, {
          ...taskResumePatchForEvidence(evidence),
          ...(reviewId ? { operationId: reviewId } : {})
        });
      }
      endActivity(activity, evidence.kind === "not_started" ? "failed" : "warning");
      // §6.5：先展示实际变更（哪些步写入了、哪些没写），由用户决定撤销还是重新生成计划。
      // 不自动重试，也不拿这份 after 当新计划的 before。
      append({
        role: "system",
        text: `${error instanceof Error ? error.message : "WPS 写入失败。"}${evidenceSuffix(evidence)}`,
        ...(reviewId && changeReviewsRef.current.has(reviewId)
          ? { changeReviewExecutionId: reviewId }
          : {}),
        /*
         * 写失败 ≠ 没有恢复点。只要这一次执行真的留下了恢复资料（`needs_review`），
         * 就必须给出结果卡（`outcome: "partial"`），并把撤销入口挂在**它**上面 ——
         * 否则用户只能从抽屉里看到一个不知道属于哪次执行的恢复点。
         * 拿不到执行 id、或宿主里查不到对应记录时**不画卡**：没有证据就不画。
         */
        ...(partiallyWritten && latest
          ? {
              verification: {
                executionId: latest.executionId,
                actionCount: plan.actions.length,
                cellCount: plan.totalCells,
                status: "recorded" as const,
                outcome: "partial" as WpsExecutionOutcome,
                targets: plan.actions.map((action) => ({
                  type: action.type,
                  sheetName: redactCredentials(action.sheetName),
                  range: action.range
                }))
              }
            }
          : {})
      });
    } finally {
      writeBackInFlight.current = false;
      setWriteProgress(null);
    }
    return completedRecord ?? undefined;
  };

  const applyEditedPlan = async () => {
    if (!pendingPlan || planEditDraft === null || busy) return;
    if (!(await targetGuard.check())) return;
    setWriteProgress("正在重新校验修改后的计划…");
    try {
      const parsed: unknown = JSON.parse(planEditDraft);
      assertEditablePlan(parsed);
      const application = getWpsApplication();
      const currentSheetIndexes = [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)];
      const metadata = listWpsWorksheets(application);
      const selectedNames = metadata.sheets
        .filter((sheet) => currentSheetIndexes.includes(sheet.index))
        .map((sheet) => sheet.name);
      // Locally computed mapping values must never be sent to plan validation.
      const validated = pendingPlan.sourceSnapshots.length
        ? parsed
        : await validatePlan(parsed, selectedNames, metadata.sheets.map((sheet) => sheet.name));
      const updated = prepareWpsPlan(application, currentSheetIndexes, validated, pendingPlan.sourceSnapshots, { requireSourceAnchor: pendingPlan.sourceAnchorRequired });
      setPendingPlan(updated);
      if (pendingPlanMessageId) {
        updateMessages((current) => current.map((message) =>
          message.id === pendingPlanMessageId && message.planPreview
            ? { ...message, planPreview: { ...message.planPreview, title: updated.plan.title, revised: true } }
            : message
        ));
      }
      setPlanEditDraft(null);
      setPlanEditError(null);
    } catch (error) {
      setPlanEditError(error instanceof Error ? error.message : "计划编辑内容无效。");
    } finally {
      setWriteProgress(null);
    }
  };

  /**
   * 撤销**某一次具体执行**。入口有两处（执行结果卡内 / 备份抽屉的快捷入口），
   * 但归属键只有一个：`executionId`。绝不按"最近一次"或消息位置猜。
   *
   * 只有当前会话里活跃的那个恢复点本人才可点，所以这里再核一次 id：
   * 卡片上的按钮万一因历史/重开而残留，也不会撤销到别的执行上去。
   */
  const undoExecution = async (
    executionId: string,
    fromResultReview = false,
    origin: "undo" | "discard" = "undo"
  ): Promise<boolean> => {
    if (fromResultReview ? baseBusy : busy) return false;
    const record = undoRecord && undoRecord.executionId === executionId ? undoRecord : null;
    if (!record) return false;
    if (!(await targetGuard.check())) return false;
    /*
     * 撤销**不新增执行过程卡片**：逐步进度只进内部单行指示（`writeProgress`），
     * 不进对话。原入口靠 `undoingExecutionId` 改字、变灰，防重复点击。
     * 内部的身份校验、逐动作恢复与回读核对一步都没少（见 `undoWpsExecution`）。
     */
    setUndoingExecutionId(executionId);
    setWriteProgress("正在核对撤销条件…");
    try {
      await undoWpsExecution(getWpsApplication(), record.executionId, (update) => {
        setWriteProgress(update.message);
      });
      // 重新问宿主：撤销成功后那条记录已变成 `restored`，不再是活跃恢复点；
      // 若更早的执行还留着可撤销记录，`latestWpsUndo` 会把它交出来（真正的 LIFO 栈）。
      try {
        setUndoRecord(latestWpsUndo(getWpsApplication()));
      } catch {
        setUndoRecord(null);
      }
      // 区域已回到执行前，这份对照描述的现场不再存在：按 §7 生命周期一并释放，
      // 而不是留着一张写着"已写入并回读"的过期明细。
      // 但"改过又撤销"本身是操作记录，只摘掉单元格内容后留一句摘要。
      const released = changeReviewsRef.current.get(record.executionId);
      if (released && released.length > 0) {
        changeReviewSummariesRef.current.set(
          record.executionId,
          released.map((review) => ({ ...changeReviewSummary(review), sheet: redactCredentials(review.sheet) }))
        );
      }
      changeReviewsRef.current.delete(record.executionId);
      // 结果卡回标「已撤销并恢复」—— 这是撤销成功的唯一权威展示位。
      updateMessages((current) => current.map((message) =>
        message.verification?.executionId === record.executionId
          ? { ...message, verification: { ...message.verification, status: "undone" } }
          : message
      ));
      /*
       * 该执行已被撤销 ⇒ 它的「保存为工具」资格一并作废。
       * 不能拿"已经被撤销的写入"去固化工具：清掉已验证身份，让任何仍开着的保存
       * 对话框进入"资格已失效"（确认按钮禁用；`saveToolInvalidationReason` 在确认
       * 时再拦一道），并复位已勾选的风险确认项，不留任何已批准的状态。
       */
      /*
       * 从**结果卡**撤销之后，固定操作条与审阅状态必须同步收尾：
       * 否则它会继续显示"请查看主表中的执行结果"并提供「确认变更」，
       * 而那份结果已经被还原了。`settleExternally` 只同步状态、不再次恢复。
       *
       * 记下这一次，收尾消息效应据此**跳过**：用户点「撤销」时不该再听到一句
       * 「已放弃并恢复」—— 那是「放弃变更」的结果文案，两件事不能重复反馈。
       */
      if (origin === "undo") {
        undoSettledExecutionIdRef.current = record.executionId;
        void resultReview.settleExternally(record.executionId);
      }
      if (verifiedPlanForSaveRef.current?.executionId === record.executionId) {
        verifiedPlanForSaveRef.current = null;
      }
      if (saveToolContextRef.current?.executionId === record.executionId) {
        setApproveFixedContent(false);
        setApproveDestructive(false);
      }
      /*
       * 最多一条简短成功消息，且只属于「撤销」这条路径。
       * 「放弃变更」走的是 `resultReview.discard()` 自己的恢复，它的结果文案由
       * 收尾消息效应给（"已放弃并恢复：…"），这里再补一句就成了两条。
       */
      if (origin === "undo") {
        append({ role: "assistant", text: `已撤销「${record.planTitle}」，并回读确认恢复成功。` });
      }
      return true;
    } catch (error) {
      try {
        setUndoRecord(latestWpsUndo(getWpsApplication()));
      } catch {
        setUndoRecord(null);
      }
      /*
       * 失败/冲突**不静默、不误报成功**：结果卡保持原状（不会变成"已撤销"），
       * 这里只说真实状态。撤销是逐动作反序写的，中途失败时"已处理的区域保持
       * 恢复后、未处理的区域保持执行后"是可能的中途现场 —— 所以用只读核对
       * (`inspectWpsUndoRecord`) 看清现场再下结论，不猜。
       */
      const reason = error instanceof Error ? error.message : "WPS 撤销失败。";
      const inspection = inspectWpsUndoRecord(getWpsApplication(), record.executionId);
      const state = inspection.kind === "not_started"
        ? "目标区域已全部回到执行前状态，但恢复记录未能收尾，请重试一次。"
        : inspection.kind === "conflict"
          ? "目标区域处于部分恢复状态：一部分已回到执行前，一部分仍是执行后。"
          : inspection.kind === "unknown"
            ? `无法确认目标区域当前状态：${inspection.detail}`
            : "目标区域未做任何修改，仍是执行后的状态。";
      append({
        role: "system",
        text: `撤销未完成：${reason} ${state}`,
        errorDetail: [
          `执行：${redactCredentials(record.planTitle)}`,
          `执行身份：${record.executionId}`,
          `现场核对：${inspection.kind === "conflict" ? inspection.detail : inspection.kind}`,
          `原因：${redactCredentials(reason)}`
        ].join("\n")
      });
      return false;
    } finally {
      setUndoingExecutionId(null);
      setWriteProgress(null);
    }
  };
  /*
   * `origin` 让恢复知道自己是**被谁触发**的：
   *  - `"discard"`：用户点了「放弃变更」，结果文案走放弃那一套（收尾消息效应），
   *    这里不再补一句"已撤销…"，避免同一件事出现两条成功消息；
   *  - `"undo"`：用户点了撤销入口，成功消息由 `undoExecution` 给。
   * 两条路径都复用同一个恢复实现（算法不变），只是**反馈分工**不同。
   */
  resultUndoRef.current = (id, origin) => undoExecution(id, true, origin ?? "undo");
  const resultFacts = useRef({ pendingPlan, authorizationRevision });
  resultFacts.current = { pendingPlan, authorizationRevision };
  /*
   * 结果审阅通道：计划就绪即在**既有授权范围内**执行到主表，随后由固定操作条
   * 承担「双视图对比 / 放弃变更 / 确认变更」。
   *
   * 为什么不再有"写入前确认"这一步：用户已明确授权"先执行、后审阅、可选双视图"。
   * 执行仍受全部守卫约束（授权边界、计划校验、执行前复核、恢复点、回读），
   * 变化只在**确认时机**——确认的对象从"这份计划"变成"主表中的实际结果"。
   * 因此 `confirmPendingPlan(true, true)` 的第二参数为真：不在执行期改写保存工具资格，
   * 那件事由 `accepted()` 在用户确认之后做。
   */
  useEffect(() => {
    if (!pendingPlan || !pendingPlanMessageId || baseBusy || resultReviewState.pending || resultReviewState.phase === 'reviewing') return;
    const prepared = pendingPlan, key = pendingPlanMessageId, revision = authorizationRevision;
    const planKey = JSON.stringify(prepared.plan);
    // 本次审阅标记的**完整随机公式**：mark 阶段生成、showBefore 建副本时用来按完整公式
    // 清掉被 Copy() 带进副本的本次规则（不用前缀匹配，避免误删另一轮审阅的标记）。
    let currentMarkFormula: string | null = null;
    /*
     * 绑定属于"**本次**审阅"。开跑就先解掉上一轮可能留下的绑定：
     * 它若不清除，一次没有正常收尾的旧审阅会一直把某本工作簿认成"本次副本"，
     * 新的一次审阅反而被它带偏。
     */
    releaseWpsReviewTarget();
    const isCurrent = () => {
      const now = resultFacts.current;
      return now.authorizationRevision === revision && (!now.pendingPlan || now.pendingPlan === prepared) &&
        JSON.stringify(prepared.plan) === planKey &&
        /*
         * 原表身份：双视图打开后 WPS 会把**副本**激活成 `ActiveWorkbook`，直接读活动簿
         * 就会把副本判成"换了工作簿"，于是用户什么都没做却被拦下（真机症状）。
         * `wpsReadWorkbook` 只在"活动簿就是本次审阅那份副本、且原表仍在场"时解析回原表；
         * 切到无关工作簿、原表被关掉，它仍然返回活动簿 ⇒ 该拦的还是拦得住。
         */
        String(wpsReadWorkbook(getWpsApplication())?.FullName ?? '') === prepared.authorization.workbookFullName;
    };
    setResultTitle(prepared.plan.title);
    void resultReview.start(key, {
      isCurrent,
      /*
       * 执行前只留**内存快照**，不建副本。
       *
       * 旧写法在这里就 `stageWpsNativeComparison(..., true)`：于是"宿主支持工作表复制"
       * 变成了写入的前置条件，用户根本没点双视图也会被挡下（真机与替身都会）。
       * 副本改在 `showBefore` 里建 —— 那是用户真的要看执行前状态的时候。
       */
      prepare: async () => {
        if (!isCurrent()) throw new Error('当前计划或授权已变化，未执行。');
        assertWpsPlanSourceCurrent(getWpsApplication(), prepared);
        return captureWpsBeforeResult(getWpsApplication(), prepared);
      },
      execute: async () => {
        const record = await confirmPendingPlan(true, true);
        if (!record) throw new Error('执行未完成，请检查执行记录与恢复入口；不会自动重试。');
        return { record, executionId: record.executionId, verified: true, undoable: true };
      },
      mark: () => {
        // 公式在**第一次 Add 之前**就交给这里：标记建到一半失败时双视图仍能按它
        // 精确清掉本次残留（`markWpsResult` 会先回调公式、再抛错）。
        const marks = markWpsResult(getWpsApplication(), prepared, formula => { currentMarkFormula = formula; });
        return Promise.resolve(marks);
      },
      clearMarker: marks => Promise.resolve(clearWpsResultMarks(getWpsApplication(), marks)),
      showBefore: async before => {
        /*
         * 只有用户**真的**点「双视图对比」时才建立副本，而且必须按**执行前快照**重建。
         *
         * 不能拿"复制当前表"当执行前副本：执行已经发生，复制到的是执行后状态，
         * 对比就失去意义；`stageWpsNativeComparison(..., true)` 还会拿预览快照去比对
         * 已经变过的原表而直接抛错。所以走 `stageWpsBeforeComparison`。
         */
        const staged = stageWpsBeforeComparison(getWpsApplication(), before, currentMarkFormula ?? undefined);
        /*
         * 副本一建立就登记"本次审阅的目标绑定"：之后 WPS 把副本激活成活动簿时，
         * 确认/放弃与范围核对仍按**原表**进行，而不是把副本当成新授权目标。
         * 只认本次创建时记下的身份，无关工作簿不受影响（见 `wpsResultReviewTarget`）。
         */
        bindWpsReviewTarget({
          original: staged.original,
          originalIdentity: identityOf(staged.original) ?? {
            name: staged.plan.authorization.workbookName,
            fullName: String(staged.plan.authorization.workbookFullName ?? ""),
            path: staged.plan.authorization.workbookPath ?? null
          },
          copyIdentity: {
            name: staged.copyIdentity.name,
            fullName: String(staged.copyIdentity.fullName ?? ""),
            path: staged.copyIdentity.path ?? null
          },
          writeSheet: { index: staged.writeSheet.index, name: staged.writeSheet.name }
        });
        if (rearrangeWpsComparison(getWpsApplication(), staged) !== 'arranged') {
          throw new Error('未能打开双视图，请手动查看本次执行前副本；执行结果与撤销不受影响。');
        }
        return staged;
      },
      /*
       * 副本侧"本次标记没清干净"**必须说出来**：那种情况下两张表可能显示同样的
       * 高亮，用户会把执行前状态看成执行后状态。清掉了就返回 null，保留原有提示。
       */
      comparisonNotice: copy => {
        const marks = copy.reviewMarks;
        if (!marks || marks.unconfirmed === 0) return null;
        return `执行前副本的审阅标记未能核实清理（${marks.unconfirmed} 处区域，可能读不到或删不掉）；两张表可能显示相同底色，请以原表与回读结果为准。`;
      },
      closeBefore: async before => {
        const outcome = discardWpsNativeComparison(before);
        const settled = outcome === 'closed' || outcome === 'already_closed';
        // 副本确实收尾了才解除绑定；没关掉就继续认它，否则它会被当成"另一本工作簿"。
        if (settled) releaseWpsReviewTarget();
        return settled ? null : '执行前副本仍保留，请手动核对；原表不会被关闭。';
      },
      restore: async receipt => {
        if (!(await resultUndoRef.current(receipt.executionId, "discard"))) throw new Error('本次结果未恢复，未声明已放弃。');
      },
      validateResult: async receipt => {
        const inspection = inspectWpsUndoRecord(getWpsApplication(), receipt.executionId);
        if (inspection.kind !== 'completed') throw new Error('本次结果已变化或无法核对，未确认变更。');
      },
      accepted: receipt => {
        verifiedPlanForSaveRef.current = { plan: prepared.originalPlan ?? prepared.plan,
          sourceSheetNames: prepared.authorization.sheets.map(sheet => sheet.name), executionId: receipt.executionId };
      }
    });
  }, [pendingPlan, pendingPlanMessageId, baseBusy, resultReviewState.pending, resultReviewState.phase]);

  /**
   * 某张执行结果卡此刻该显示的恢复点状态。
   *
   * 活跃恢复点就是 `undoRecord`（`latestWpsUndo` 的产物 = 当前会话里最新的那条），
   * 所以"新执行取代旧恢复点"天然成立：旧卡拿到的 `live.executionId` 对不上，
   * 只会显示"恢复点已不可用"，而不是一个点了就报错的假按钮。
   */
  const wpsRecoveryCardStateFor = (verification: NonNullable<ChatMessage["verification"]>) =>
    wpsRecoveryCardState({
      executionId: verification.executionId,
      undone: verification.status === "undone",
      live: undoRecord ? { executionId: undoRecord.executionId, state: undoRecord.state } : null
    });

  const prepareSavedWorkflow = async (tool: SavedTool) => {
    if (!(await targetGuard.check())) return;
    if (busy) return;
    try {
      const toolSheetIndexes = [
        ...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)
      ];
      if (toolSheetIndexes.length === 0) {
        throw new Error("请先选择这个工具可以使用的工作表");
      }
      const application = getWpsApplication();
      const metadata = listWpsWorksheets(application);
      const selectedNames = metadata.sheets
        .filter((sheet) => toolSheetIndexes.includes(sheet.index))
        .map((sheet) => sheet.name);
      const values = Object.fromEntries(
        tool.parameters.map((parameter) => [parameter.id, parameter.defaultValue])
      );
      const plan = instantiateTool(tool, values);
      const normalizedPlan = await validatePlan(
        plan,
        selectedNames,
        metadata.sheets.map((sheet) => sheet.name)
      );
      if (isWorksheetDeletion(normalizedPlan)) {
        setToolsOpen(false);
        const requestMessageId = append({ role: "assistant", text: "" });
        const outcome = await executeStructuralDeletion(normalizedPlan, toolSheetIndexes, requestMessageId);
        if (outcome === "failed") {
          updateMessages(current => current.map(message => message.id === requestMessageId && !message.text
            ? { ...message, text: "工作表删除未完成；可在下方结果卡查询状态或打开执行前副本。" }
            : message));
        }
        return;
      }
      const prepared = prepareWpsPlan(application, toolSheetIndexes, normalizedPlan);
      setPendingPlan(prepared);
      setPendingPlanMessageId(append({
        role: "assistant",
        text: "",
        // 固化工具没有自由文本需求，原需求就是"运行哪个工具"本身。
        taskResume: previewTaskResume(
          carried().taskId,
          `运行固化工具：${tool.name}`,
          authorizationRevision
        ),
        planPreview: { title: prepared.plan.title, status: "pending" }
      }));
      setToolsOpen(false);
    } catch (error) {
      append({
        role: "system",
        text:
          error instanceof Error
            ? `这个工具暂时不能在 WPS 中运行：${error.message}`
            : "这个工具暂时不能在 WPS 中运行。"
      });
      setToolsOpen(false);
    }
  };

  // —— 个人工具管理（保存 / 编辑模板 / 删除），对齐 Excel 的既有任务 ——

  /**
   * 打开「保存为工具」确认对话框。
   *
   * 保存资格绑定**实际已验证的计划版本与执行身份**（`verifiedPlanForSaveRef`）：
   * 只有完全通过（`readback_passed`）的写入才会被记进来；失败、部分完成、
   * 未验证的计划根本拿不到入口。打开时把当前计划设为候选并**重置确认项**
   * （`beginSaveCandidate` 会把固定内容/破坏性勾选复位为未勾选）。
   */
  const openSaveToolDialog = (executionId: string) => {
    const verified = verifiedPlanForSaveRef.current;
    if (!verified || verified.executionId !== executionId) {
      append({ role: "system", text: "这份计划当前没有可保存为工具的已验证版本。" });
      return;
    }
    saveToolContextRef.current = {
      executionId,
      sourceSheetNames: verified.sourceSheetNames
    };
    setToolName(verified.plan.title);
    setToolDescription(verified.plan.summary);
    beginSaveCandidate(verified.plan);
  };

  /** 取消保存：清空候选与上下文，不留任何批准状态。 */
  const cancelSaveTool = () => {
    saveToolContextRef.current = null;
    closeSaveCandidate();
  };

  /**
   * 候选在打开对话框后是否已失效。确认为**必检项**，不只是界面禁用：
   * 计划被替换（新执行）、或该执行已被撤销，都会让旧候选失去保存资格。
   */
  const saveToolInvalidationReason = (): string | null => {
    const candidate = saveToolCandidate;
    const context = saveToolContextRef.current;
    const verified = verifiedPlanForSaveRef.current;
    if (!candidate || !context) return "没有可保存的候选计划。";
    if (!verified || verified.executionId !== context.executionId || verified.plan !== candidate) {
      return "这份计划的可保存资格已失效（执行已被新的执行替换，或本次执行已撤销）。请重新执行并通过验证后再保存。";
    }
    return null;
  };

  /** 确认保存：确认时**重新检查**资格，再调用 createTool（其固化的审批门不放宽）。 */
  const confirmSaveToolForWps = () => {
    const candidate = saveToolCandidate;
    const context = saveToolContextRef.current;
    if (!candidate || !context) {
      cancelSaveTool();
      return;
    }
    const invalidReason = saveToolInvalidationReason();
    if (invalidReason) {
      append({ role: "system", text: invalidReason });
      return;
    }
    try {
      // createTool 内部会重新跑 analyzeToolEligibility 并校验审批是否齐全；
      // 这里绝不把确认项预先置真、也不放宽它的守卫。
      const tool = createTool(candidate, toolName, toolDescription, context.sourceSheetNames, {
        fixedContent: approveFixedContent,
        destructive: approveDestructive
      });
      setTools(saveTool(tool));
      cancelSaveTool();
      append({ role: "system", text: `已把「${tool.name}」保存为个人工具。` });
    } catch (reason) {
      append({
        role: "system",
        text: reason instanceof Error ? reason.message : "无法保存工具"
      });
    }
  };

  /** 打开固化工具的「编辑计划模板」JSON 抽屉。 */
  const openToolPlanEditor = (tool: SavedTool) => {
    setToolPlanEdit({
      toolId: tool.id,
      draft: JSON.stringify(tool.planTemplate, null, 2),
      error: null
    });
  };

  const scriptAuthorizedSheets = (): string[] => {
    if (folderScope) throw new Error("WPS 受控 JS 当前仅支持已勾选的工作簿工作表");
    const indexes = resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes;
    return listWpsWorksheets(getWpsApplication()).sheets
      .filter(sheet => indexes.includes(sheet.index)).map(sheet => sheet.name);
  };

  /**
   * 应用编辑后的工具模板：本地 JSON + 白名单校验，写回后按当前能力/授权重新校验才可运行。
   *
   * 编辑会重算**保存资格**（`updateToolPlanTemplate` 内部重新跑 `analyzeToolEligibility`），
   * 不会沿用旧版本不再适用的批准；若所需审批项发生变化，如实告知用户。
   */
  const applyEditedToolTemplate = () => {
    if (!toolPlanEdit) return;
    let parsed: unknown;
    try {
      parsed = JSON.parse(toolPlanEdit.draft);
    } catch {
      setToolPlanEdit({ ...toolPlanEdit, error: "JSON 格式有误，请检查括号和引号是否匹配" });
      return;
    }
    try {
      assertEditablePlan(parsed);
    } catch (reason) {
      setToolPlanEdit({
        ...toolPlanEdit,
        error: reason instanceof Error ? reason.message : "计划结构不合法"
      });
      return;
    }
    const previous = tools.find((tool) => tool.id === toolPlanEdit.toolId);
    const before = new Set<ToolRiskApproval>(previous?.approvals ?? []);
    try {
      const updated = updateToolPlanTemplate(toolPlanEdit.toolId, parsed as AnalysisPlan);
      const after = new Set<ToolRiskApproval>(
        analyzeToolEligibility(parsed as AnalysisPlan).requiredApprovals
      );
      const changed =
        before.size !== after.size || [...after].some((approval) => !before.has(approval));
      setTools(updated);
      setToolPlanEdit(null);
      append({
        role: "system",
        text: changed
          ? "已更新工具计划，并重新分析了保存资格：所需风险确认项有变化，运行前仍会按当前工作簿与授权重新校验、预览并让你确认。"
          : "已更新工具计划。运行前仍会按当前授权重新校验并预览。"
      });
    } catch (reason) {
      setToolPlanEdit({
        ...toolPlanEdit,
        error: reason instanceof Error ? reason.message : "写回工具失败"
      });
    }
  };

  /** 删除个人工具（对齐 Excel 的工具管理）。 */
  const deleteSavedTool = (tool: SavedTool) => {
    setTools(deleteTool(tool.id));
    append({ role: "system", text: `已删除工具「${tool.name}」。` });
  };

  const runSavedQuery = async (tool: SavedQueryTool) => {
    if (!(await targetGuard.check())) return;
    if (busy) return;
    const querySheetIndexes = [
      ...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)
    ];
    if (querySheetIndexes.length === 0) {
      append({ role: "system", text: "请先选择这个固化查询可以使用的工作表。" });
      return;
    }
    const controller = new AbortController();
    controllerRef.current?.abort();
    controllerRef.current = controller;
    const activity = beginActivity({
      id: "saved-query-preflight",
      label: "检查固化查询的数据范围",
      phase: "preflight",
      detail: "只使用本次明确授权的工作表。"
    });
    try {
      if (tool.sourceMode !== "workbook") {
        throw new Error("这个固化查询来自工作目录；WPS 当前只支持已打开的工作簿");
      }
      const application = getWpsApplication();
      const metadata = listWpsWorksheets(application);
      const authorizedSheets = metadata.sheets.filter((sheet) =>
        querySheetIndexes.includes(sheet.index)
      );
      const requiredNames = new Set(tool.sourceSheetNames);
      const missingNames = tool.sourceSheetNames.filter(
        (name) => !authorizedSheets.some((sheet) => sheet.name === name)
      );
      if (missingNames.length > 0) {
        throw new Error(
          `当前授权范围缺少：${missingNames.join("、")}。请在数据范围中明确选择后再运行`
        );
      }
      const targetIndexes = (requiredNames.size > 0
        ? authorizedSheets.filter((sheet) => requiredNames.has(sheet.name))
        : authorizedSheets
      ).map((sheet) => sheet.index);
      setProgress({ stage: "reading", message: "正在读取固化查询已授权的数据源…" });
      updateActivity(activity, {
        id: "saved-query-read",
        label: "读取授权数据",
        phase: "preflight",
        detail: "读取固化查询实际需要的工作表。"
      });
      const authorization = authorizeWpsSheets(application, targetIndexes);
      const results = await readAuthorizedWpsSheets(application, authorization, {
        signal: controller.signal,
        onProgress: ({ sheet, rowsRead, totalRows }) => {
          setProgress({
            stage: "reading",
            message: `正在读取「${sheet.name}」：${rowsRead} / ${totalRows} 行`
          });
          updateActivity(activity, {
            id: `saved-query-read:${sheet.name}`,
            label: `读取「${sheet.name}」`,
            phase: "preflight",
            target: { sheet: sheet.name },
            current: rowsRead,
            total: totalRows,
            unit: "rows"
          });
        }
      });
      setProgress({ stage: "calculating", message: "正在本地确定性执行固化查询…" });
      updateActivity(activity, {
        id: "saved-query-calculate",
        label: "执行固化查询",
        phase: "executing",
        toolName: "query_table",
        detail: "在本机对完整授权数据执行确定性计算。"
      });
      const result = executeQueryTableData(
        instantiateQueryTool(tool),
        results.map((item) => ({ name: item.sheet.name, values: item.values }))
      );
      endActivity(activity, "completed");
      append({
        role: "assistant",
        text: `固化查询「${tool.name}」已完成：扫描 ${result.scannedRows.toLocaleString()} 行，返回 ${result.rows.length} 行；模型调用 0 次。`,
        resultContext: {
          kind: "table",
          title: result.title,
          headers: result.headers.map(String),
          rows: result.rows,
          sourceSheets: result.sourceSheets,
          warnings: result.warnings
        }
      });
      setToolsOpen(false);
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        endActivity(activity, "cancelled");
        append({ role: "system", text: "固化查询已停止，没有修改任何工作表。" });
      } else {
        endActivity(activity, "failed");
        append({
          role: "system",
          text: error instanceof Error ? error.message : "固化查询运行失败。"
        });
      }
    } finally {
      if (controllerRef.current === controller) controllerRef.current = null;
      setProgress(null);
    }
  };

  /**
   * 待答澄清的**授权范围复核**：三条**独立**判据。
   *
   * 返回 `true` 表示可以继续作答；`false` 表示范围/身份已经变化，卡片被作废、
   * 原需求退回输入框，要求用户重新授权、重新读取、重新预览。
   *
   * 1. **指纹判据**（可序列化，进对话历史）：工作簿名 + 每张授权表的
   *    `[index, name]`（含顺序）。它负责跨会话能证明的那部分。
   * 2. **表头判据**（真实第 1 行单元格值）：与指纹读**不同数据源**，
   *    否则它会退化成指纹的副本、让指纹看起来"没必要"。
   * 3. **宿主对象身份**（只在内存）：指纹**结构性看不见**的两类替换 ——
   *    不同目录的同名工作簿、同索引同名同表头的重建工作表 —— 只能靠对象引用
   *    （+ 完整路径 + 稳定 ID + 结构信号）分辨。
   *
   * **第 3 条采用 fail-closed**：身份未被捕获（面板重开、宿主读不到、未保存
   * 工作簿且引用已换）时一律作废，而不是"读不到就放行"。仅凭相同名称与表头
   * **不得**放行 —— 那正是本模块要堵的缺口。
   *
   * 三条判据读不同数据源、各自独立必要，不得互相替代，也不得只留其一。
   */
  const clarificationScopeVerdict = async (
    clarification: { sheetIdentityFingerprint?: string; headerFingerprint?: string }
  ): Promise<WpsClarificationVerdict> => {
    const scopeIndexes = clarificationScopeIndexesRef.current;
    const currentIdentity = currentWpsSheetIdentity();
    const currentHeader = await currentWpsHeaderFingerprint(scopeIndexes);
    const objectIdentity = clarificationObjectIdentityRef.current;
    let check: WpsIdentityCheck | null = null;
    if (objectIdentity) {
      try {
        check = verifyWpsObjectIdentity(getWpsApplication(), objectIdentity);
      } catch {
        check = "unknown";
      }
    }
    const verdict = wpsClarificationVerdict({
      identity: { recorded: clarification.sheetIdentityFingerprint, current: currentIdentity },
      header: { recorded: clarification.headerFingerprint, current: currentHeader },
      objectIdentity: check
    });
    if (verdict === "unknown" && check === "unknown" && objectIdentity &&
        clarification.sheetIdentityFingerprint && clarification.sheetIdentityFingerprint === currentIdentity &&
        clarification.headerFingerprint && clarification.headerFingerprint === currentHeader) {
      const recordedKey = wpsConfirmationTargetKey(objectIdentity);
      const currentKey = wpsConfirmationTargetKey(captureWpsObjectIdentity(getWpsApplication(), scopeIndexes));
      // Explicit PRE-ANALYSIS scope consent permits fresh reads/replanning within this task.
      // It does not prove native object identity and is never consumed by the write executor.
      if (recordedKey && recordedKey === currentKey && targetGuard.hasAnalysisConsent(recordedKey)) return "same";
    }
    return verdict;
  };

  /**
   * 作答前的范围复核。只有 `"same"` 才放行。
   *
   * `"changed"` 与 `"unknown"` **都不放行**，但说法与善后**必须不同**：
   * 前者有证据证明范围变了；后者只是证明不了（宿主换了包装对象、读表失败、
   * 面板刚重开）。把后者也说成"数据范围已经变化"是对用户的假话 ——
   * 他只改了一格又撤回，范围根本没变。所以 `"unknown"` 时保留他的原需求与
   * 刚填的回答，请他重新确认授权，再作为**新一轮**重新读取、规划与预览。
   */
  const reviewClarificationScope = async (
    clarification: {
      messageId: string;
      originalPrompt: string;
      sheetIdentityFingerprint?: string;
      headerFingerprint?: string;
    },
    preservedAnswer: string | null
  ): Promise<WpsClarificationVerdict> => {
    const verdict = await clarificationScopeVerdict(clarification);
    if (verdict === "same") return verdict;

    updateMessages((current) => current.map((item) =>
      item.id === clarification.messageId && item.clarification
        ? {
            ...item,
            clarification: {
              ...item.clarification,
              status: verdict === "changed" ? "invalidated" : "unconfirmed",
              answer: preservedAnswer ?? item.clarification.answer
            },
            taskResume: item.taskResume ? withTaskResumeState(item.taskResume, {
              blocker: "authorization", resumability: "reauthorize"
            }) : undefined
          }
        : item
    ));
    append({ role: "system", text: wpsClarificationScopeNotice(verdict) });
    if (verdict === "changed") {
      // 范围确实变了：原需求还回去，让用户按新范围重发。
      setPrompt(clarification.originalPrompt);
      setClarificationDraft("");
    } else {
      // 证明不了：把原需求连同他刚写的回答一起放回输入框，
      // 他确认范围后直接发送即为一轮**全新**请求（重新读取、规划、预览），
      // 不复用旧卡的授权，也不会绕过写入前的确认。
      setPrompt(
        preservedAnswer
          ? `${clarification.originalPrompt}\n补充确认：${preservedAnswer}`
          : clarification.originalPrompt
      );
      setClarificationDraft(preservedAnswer ?? "");
    }
    return verdict;
  };

  // Business forms only verify the pre-analysis authorization; they never grant new consent.
  const checkLocalMappingScope = async (pending: PendingClarification) => {
    if (!(await targetGuard.check())) throw new TargetConfirmationRequiredError();
    const indexes = [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)];
    if (JSON.stringify(indexes) !== JSON.stringify(pending.selectedSheetIndexes ?? [])) {
      await targetGuard.requestConfirmation();
      throw new TargetConfirmationRequiredError();
    }
    const verdict = await clarificationScopeVerdict(pending);
    if (verdict === "changed") {
      await targetGuard.requestConfirmation();
      throw new TargetConfirmationRequiredError();
    }
    if (verdict === "same") return;
    const original = clarificationObjectIdentityRef.current;
    const current = captureWpsObjectIdentity(getWpsApplication(), indexes);
    const originalKey = original && wpsConfirmationTargetKey(original);
    const targetKey = wpsConfirmationTargetKey(current);
    const identity = currentWpsSheetIdentity();
    const header = await currentWpsHeaderFingerprint(indexes);
    // Read failures and positive changes cannot be fixed by prior scope consent.
    if (!originalKey || originalKey !== targetKey || !identity || !header ||
        identity !== pending.sheetIdentityFingerprint || header !== pending.headerFingerprint ||
        wpsConfirmationTargetKey(captureWpsObjectIdentity(getWpsApplication(), indexes)) !== targetKey) {
      await targetGuard.requestConfirmation();
      throw new TargetConfirmationRequiredError();
    }
    if (!targetKey || !targetGuard.hasAnalysisConsent(targetKey)) {
      await targetGuard.requestConfirmation();
      throw new TargetConfirmationRequiredError();
    }
  };

  const resolveLocalMapping = async (
    resolvedRule: NonNullable<WpsTurnClarification["localMapping"]>["rule"],
    reauthorized?: { pending: PendingClarification; revision: number }
  ) => {
    const pending = reauthorized?.pending ?? clarification;
    if (!pending?.localMapping || busy || pendingPlan) return;
    if (!(await targetGuard.check())) return;
    const localMapping = pending.localMapping;
    const currentIndexes = [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)];
    const recordedIndexes = [...(pending.selectedSheetIndexes ?? [])];
    const sameAuthorization = JSON.stringify(currentIndexes) === JSON.stringify(recordedIndexes);
    if (!reauthorized) await checkLocalMappingScope(pending);
    const verdict = sameAuthorization
      ? reauthorized ? "same" : await clarificationScopeVerdict(pending)
      : "changed";
    if (verdict !== "same") {
      localMappingResumeRef.current = verdict === "unknown"
        ? { pending, rule: resolvedRule, identity: clarificationObjectIdentityRef.current }
        : null;
      updateMessages((current) => current.map((message) =>
        message.id === pending.messageId && message.clarification
          ? { ...message,
              clarification: { ...message.clarification, status: verdict === "changed" ? "invalidated" : "unconfirmed",
                answer: verdict === "unknown" ? "本地规则已提交，等待重新确认范围" : undefined },
              taskResume: message.taskResume ? withTaskResumeState(message.taskResume, {
                phase: "answer_submitted", blocker: "authorization", resumability: "reauthorize"
              }) : undefined }
          : message
      ));
      append({ role: "system", text: wpsClarificationScopeNotice(verdict) });
      setPrompt(pending.originalPrompt);
      setClarification(null);
      return;
    }
    localMappingResumeRef.current = null;
    const localTask = messages.find((item) => item.id === pending.messageId)?.taskResume;
    if (reauthorized) {
      setClarification(pending);
      updateMessages((current) => current.map((item) => item.id === pending.messageId && item.clarification
        ? { ...item, clarification: { ...item.clarification, status: "pending", answer: undefined } } : item));
    }
    const activity = beginActivity({
      id: "read-mapping-source",
      label: "重新读取授权源列",
      phase: "preflight",
      detail: "按用户补充的映射、分隔符和优先顺序在本机重新计算；源值不出站。"
    });
    const controller = new AbortController();
    controllerRef.current?.abort();
    controllerRef.current = controller;
    setProgress({ stage: "reading", message: "正在重新读取授权源列并计算预览…" });
    try {
      const previewTarget = captureWpsObjectIdentity(getWpsApplication(), recordedIndexes);
      const result = await resumeWpsLocalMapping({
        application: getWpsApplication(),
        selectedSheetIndexes: recordedIndexes,
        rule: resolvedRule,
        context: localMapping.context,
        additions: [],
        signal: controller.signal,
        beforeRead: async () => {
          if (!(await targetGuard.check())) throw new TargetConfirmationRequiredError();
          const header = await currentWpsHeaderFingerprint(recordedIndexes);
          const current = captureWpsObjectIdentity(getWpsApplication(), recordedIndexes);
          const key = wpsConfirmationTargetKey(previewTarget);
          if (!header || header !== pending.headerFingerprint ||
              (key ? key !== wpsConfirmationTargetKey(current)
                : verifyWpsObjectIdentity(getWpsApplication(), previewTarget) !== "same")) {
            await targetGuard.requestConfirmation();
            throw new TargetConfirmationRequiredError();
          }
        },
        onProgress: (update) => {
          setProgress(update);
          updateActivity(activity, readonlyActivityUpdate(update));
        }
      });
      if (result.kind === "clarification" && result.clarification.localMapping) {
        // 用户补充期间数据可能改变。新缺口仍停在本地问，不复用旧源快照。
        setClarification({
          ...pending,
          localMapping: result.clarification.localMapping,
          localMappingRevision: (pending.localMappingRevision ?? 0) + 1
        });
        updateMessages((current) => current.map((message) =>
          message.id === pending.messageId && message.clarification
            ? { ...message, clarification: {
                ...message.clarification,
                status: "pending", answer: undefined,
                question: result.clarification.localMapping!.reason === "blank"
                  ? `源区域重新读取后有 ${result.clarification.localMapping!.blankCount ?? 0} 个空白格，请确认这些位置如何处理。`
                  : `源区域重新读取后仍有 ${result.clarification.localMapping!.values.length} 个来源值未覆盖。请在卡片内补全规则。`
              } }
            : message
        ));
        endActivity(activity, "awaiting_input");
        return;
      }
      if (result.kind !== "plan" || !result.sourceSnapshots?.length) {
        throw new Error("本地映射没有生成可复核的写入计划。请重新发送需求。");
      }
      const prepared = prepareWpsPlan(getWpsApplication(), recordedIndexes, result.plan, result.sourceSnapshots);
      updateMessages((current) => current.map((message) =>
        message.id === pending.messageId && message.clarification
          ? { ...message, clarification: { ...message.clarification, status: "resolved", answer: "已补全本地映射并重新预览" } }
          : message
      ));
      setClarification(null);
      setPendingPlan(prepared);
      setPendingPlanMessageId(append({
        role: "assistant",
        text: "",
        // 同一件任务推进到"预览已生成"：沿用原需求与 taskId，不新开一件任务。
        taskResume: previewTaskResume(
          localTask?.taskId ?? carried().taskId,
          localTask?.originalPrompt ?? resumeCarryRef.current?.originalPrompt ?? pending.originalPrompt,
          reauthorized?.revision ?? authorizationRevision,
          result.turnId,
          localTask?.confirmations ?? carried().confirmations
        ),
        planPreview: { title: prepared.plan.title, status: "pending" }
      }));
      endActivity(activity, "awaiting_input");
    } catch (error) {
      if (error instanceof TargetConfirmationRequiredError) {
        // Keep the editable form and private answer; scope confirmation is not a failed task.
        endActivity(activity, "awaiting_input");
        throw error;
      }
      endActivity(activity, "failed");
      localMappingResumeRef.current = { pending, rule: resolvedRule, identity: clarificationObjectIdentityRef.current };
      setClarification(null);
      updateMessages((current) => current.map((item) => item.id === pending.messageId && item.clarification
        ? { ...item, clarification: { ...item.clarification, status: "unconfirmed",
            answer: "本地规则已提交，等待重新确认范围" },
          taskResume: item.taskResume ? withTaskResumeState(item.taskResume, {
            phase: "answer_submitted", blocker: "network", resumability: "reauthorize"
          }) : undefined } : item));
      // 宿主错误可能把正在处理的单元格原文串进 message；本地澄清路径
      // 不把这种异常原文写入会持久化的消息历史。卡片仍留着，用户可重试。
      append({ role: "system", text: "重新读取或生成预览失败，原表未写入。请核对授权范围后重试；若仍失败，请重新发送需求。" });
    } finally {
      if (controllerRef.current === controller) controllerRef.current = null;
      setProgress(null);
    }
  };

  const resolveClarification = async (optionIndex: number) => {
    if (!clarification || busy) return;
    if (!(await targetGuard.check())) return;
    const option = clarification.options[optionIndex];
    if (!option) return;
    let answeredRecord: TaskResumeRecord | null;
    try {
      answeredRecord = settleTaskAnswer(clarification.messageId, {
        label: option.label, value: option.resolution, questionId: clarification.questionId
      }, option.resolutionMode);
    } catch {
      append({ role: "system", text: "任务确认内容为空或过长，请缩短说明后继续；原表未写入。" });
      return;
    }
    if ((await reviewClarificationScope(clarification, option.label)) !== "same") {
      setClarification(null);
      return;
    }
    updateMessages((current) => current.map((message) =>
      message.id === clarification.messageId && message.clarification
        ? { ...message, clarification: { ...message.clarification, status: "resolved", answer: redactCredentials(option.label) } }
        : message
    ));
    append({ role: "user", text: option.label });
    const resolvedPrompt =
      option.resolutionMode === "replace"
        ? option.resolution
        : `${clarification.originalPrompt}\n补充确认：${option.resolution}`;
    // 先把"用户已确认什么"写进任务记录：即使这一轮立刻断线，重开后
    // 也不必让他重新回答同一个问题。
    const targetMessageId = clarification.messageId;
    resumeCarryRef.current = answeredRecord;
    try {
      await runTurn(option.resolutionMode === "replace" || !answeredRecord
        ? resolvedPrompt : resumePrompt(answeredRecord), {
        intentConfirmed: true, clarificationRound: clarification.round + 1, showUser: false,
        ...(option.resolutionMode === "replace" ? { taskContext: createAgentTaskContext(resolvedPrompt) } : {})
      });
    } finally { resumeCarryRef.current = null; }
    if (lastTurnFailedRef.current) markTaskInterrupted(targetMessageId);
  };

  const resolveClarificationWithText = async (submittedAnswer = clarificationDraft) => {
    if (!clarification || !submittedAnswer.trim() || busy) return;
    if (!(await targetGuard.check())) return;
    const answer = submittedAnswer.trim();
    let answeredRecord: TaskResumeRecord | null;
    try {
      answeredRecord = settleTaskAnswer(clarification.messageId, {
        label: "补充说明", value: answer, questionId: clarification.questionId
      });
    } catch {
      append({ role: "system", text: "任务确认内容为空或过长，请缩短说明后继续；原表未写入。" });
      return;
    }
    if ((await reviewClarificationScope(clarification, answer)) !== "same") {
      setClarification(null);
      return;
    }
    updateMessages((current) => current.map((message) =>
      message.id === clarification.messageId && message.clarification
        ? { ...message, clarification: { ...message.clarification, status: "resolved", answer: redactCredentials(answer) } }
        : message
    ));
    append({ role: "user", text: answer });
    const targetMessageId = clarification.messageId;
    resumeCarryRef.current = answeredRecord;
    try {
      await runTurn(answeredRecord ? resumePrompt(answeredRecord)
        : `${clarification.originalPrompt}\n补充确认：${answer}`, {
        intentConfirmed: true, clarificationRound: clarification.round + 1, showUser: false
      });
    } finally { resumeCarryRef.current = null; }
    if (lastTurnFailedRef.current) markTaskInterrupted(targetMessageId);
  };

  /**
   * 从**原任务卡**继续一件被中断的任务。
   *
   * 这不是"直接重做"，也不等于授权写入：它做的是**显式重新授权** ——
   * 重新读取当前真实宿主元数据（工作簿、工作表、表头），捕获新的对象身份，
   * 把授权版本推进一版，然后用"原需求 + 用户已确认的信息"作为一轮**全新**
   * 请求重新读取与规划。旧的读取快照、旧计划与旧确认资格就此废弃。
   *
   * 重新授权后**不会再**被同一个旧对象引用检查卡住：检查的基准已经是这次
   * 新建的那一版，而不是面板重开前那一版。
   */
  const resumeInterruptedTask = async (message: ChatMessage) => {
    if (!(await targetGuard.check())) return;
    const record = message.taskResume
      ? sanitizeTaskResumeRecord(message.taskResume)
      : null;
    if (!record || busy || pendingPlan || resumingTaskMessageId !== null) return;
    const decision = taskResumeDecision(record);
    if (message.hasImageAttachments && decision.action !== "reconcile" &&
      !taskImagesRef.current.get(record.taskId)?.length && !pendingImages.length) {
      setImageError("原任务包含图片，图片内容不保存在聊天历史中。请重新添加图片后再继续。");
      composerInputRef.current?.focus();
      return;
    }
    const resumeImages = pendingImages.length ? pendingImages : taskImagesRef.current.get(record.taskId) ?? [];
    if (resumeImages.length && decision.action !== "reconcile" && !selectedModel?.supportsVision) {
      setImageError("当前模型不支持图片，请切换到支持视觉输入的模型后继续。");
      setModelMenuOpen(true);
      return;
    }

    /*
     * `continue_checkpoint` 需要服务端检查点续跑通路（CAP-02），本项目尚未接线。
     * 明确告知，**不**降级成一次重新规划来冒充"从进度续跑成功"。
     */
    if (decision.action === "continue_checkpoint") {
      append({
        role: "system",
        text:
          "这个任务本来可以从服务端保存的进度继续，但当前版本还没有接通检查点续跑通路，" +
          "所以无法从上次进度接着跑。原需求与你已确认的信息都还在 —— " +
          "请重新发送需求或重新生成预览；那是重新读取、重新规划，不是精确续跑。"
      });
      return;
    }

    /*
     * `reconcile` **绝不重新规划**。上一次写入的结果无法确认时，跑一遍普通轮次
     * 会读取数据、生成一份新计划，用户很可能顺手确认 —— 那等于在原表**可能已被
     * 改过**的情况下再写一遍。这里只做只读核对；拿不到证据就保持挂起。
     */
    if (decision.action === "reconcile") {
      reconcileSuspendedTask(message, record);
      return;
    }

    // A restored card is not strong host evidence. Confirm before fresh reads/planning.
    if (!(await targetGuard.checkForAnalysis(() => resumeTaskRef.current(message)))) return;

    const indexes = [...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)];
    if (indexes.length === 0) {
      append({
        role: "system",
        text: "尚未选择工作表，不会读取或修改任何数据。请先在数据范围里确认要使用的工作表，再继续这个任务。"
      });
      return;
    }
    setResumingTaskMessageId(message.id);
    /*
     * 从历史继续：**沿用原 taskId 与已确认信息**。
     * `taskIdRef` 的初值是面板挂载时随机生成的，重开之后就是另一个身份 ——
     * 不带过去，"同一件事"当场断掉，用户还要重答一遍。
     */
    resumeCarryRef.current = {
      taskId: record.taskId,
      confirmations: record.confirmations,
      originalPrompt: record.originalPrompt
    };
    try {
      const currentIdentity = currentWpsSheetIdentity();
      /*
       * 范围变了就**明说**，不悄悄转移授权。只在双方都读到了身份、且确实不同时
       * 才说 —— 任一侧读不到就没有证据，不能替用户下"你换了工作簿"这个结论。
       */
      const recordedIdentity = message.clarification?.sheetIdentityFingerprint;
      if (
        recordedIdentity &&
        currentIdentity &&
        recordedIdentity !== currentIdentity
      ) {
        append({
          role: "system",
          text: `这个任务将应用于新的数据范围：${currentScopeLabel}。原范围不再作为授权依据。`
        });
      }
      const localResume = localMappingResumeRef.current;
      // 只复用这道题已提交的规则；不同文件/表/表头/授权集合不得挪用本地答案。
      const canResumeLocal = localResume?.pending.messageId === message.id &&
        JSON.stringify(indexes) === JSON.stringify(localResume.pending.selectedSheetIndexes) &&
        currentIdentity === localResume.pending.sheetIdentityFingerprint &&
        await currentWpsHeaderFingerprint(indexes) === localResume.pending.headerFingerprint &&
        localResume.identity !== null &&
        verifyWpsObjectIdentity(getWpsApplication(), localResume.identity) !== "different";
      // 建立这一版授权：新的对象身份 + 推进一版授权版本。
      clarificationScopeIndexesRef.current = [...indexes];
      try {
        clarificationObjectIdentityRef.current = captureWpsObjectIdentity(
          getWpsApplication(),
          indexes
        );
      } catch {
        clarificationObjectIdentityRef.current = null;
      }
      const nextRevision = authorizationRevision + 1;
      authorizationRevisionBaseRef.current = authorizationKeyOf(indexes);
      setAuthorizationRevision(nextRevision);
      updateMessages((current) => current.map((item) =>
        item.id === message.id
          ? {
              ...item,
              taskResume: withTaskResumeState(record, {
                phase: "answer_submitted",
                blocker: "none",
                resumability: "replan",
                authorizationRevision: nextRevision
              }),
              /*
               * 重新确认范围是**授权动作**，不是对那道题的回答。
               *
               * 早期实现往 `answer` 里写一句"已重新确认数据范围并继续"：
               * 一是把用户真正的回答覆盖掉，二是让"写入资格"看起来已经具备，
               * 而重新授权其实只证明"你现在授权这几张表"。
               *
               * 现在分两态：有真实答案 → 保留答案并记为已作答（`resolved`）；
               * 没有答案 → 记为"已重新授权并作为新一轮继续"（`continued`），
               * **绝不**编一条答案出来。
               */
              clarification: item.clarification
                ? item.clarification.answer?.trim()
                  ? clarificationAfterReauthorization(item.clarification)
                  : { ...item.clarification, status: "continued" as const }
                : item.clarification
            }
          : item
      ));
      setClarification(null);
      setClarificationDraft("");
      // 系统说明，不是答案：只进消息流，不写进任何卡片的 `answer`。
      append({ role: "system", text: REAUTHORIZATION_NOTICE });
      if (canResumeLocal && localResume) {
        await resolveLocalMapping(localResume.rule, { pending: localResume.pending, revision: nextRevision });
      } else {
        localMappingResumeRef.current = null;
        await runTurn(resumePrompt(record), {
          intentConfirmed: true,
          clarificationRound: Math.min(record.confirmations.length, 1),
          showUser: false,
          images: resumeImages
        });
      }
      if (lastTurnFailedRef.current) markTaskInterrupted(message.id);
    } finally {
      // 携带只在"继续历史任务"这一轮生效，结束即清空，不影响后续新任务。
      resumeCarryRef.current = null;
      setResumingTaskMessageId(null);
    }
  };

  /**
   * 核对一次"写入结果未知"的任务。**只读：不重新规划、不写原表。**
   *
   * 能核对的前提是拿得到**可靠证据** —— 记录里带着 `operationId`，且同一个加载项
   * 会话内宿主仍保存着那条恢复记录。这时读目标区域、和记录里的写前/写后指纹比，
   * 结论如实分四种：已完成 / 确实没开始 / 现场冲突 / 读不到。
   *
   * CAP-03 的跨关闭持久执行账本尚未实现，所以"重开面板且插件存储已清"这一档
   * 必然落到"读不到"。**宁可明说核对不了，也不假装核对过**，更不给一个点了就
   * 重新写入的按钮。
   */
  resumeTaskRef.current = (message) => { void resumeInterruptedTask(message); };

  const reconcileSuspendedTask = (message: ChatMessage, record: TaskResumeRecord) => {
    const executionId = record.operationId;
    if (!executionId) {
      append({
        role: "system",
        text:
          "无法自动核对这次写入：本次执行没有留下可跨会话读取的身份记录" +
          "（持久执行账本尚未实现）。请先在原表里人工确认目标区域：\n" +
          "· 已按预期改好 → 按完成处理，不必重跑；\n" +
          "· 完全没变 → 可以重新发送需求，重新生成预览后确认写入；\n" +
          "· 只改了一部分 → 请先修好原表，再重新提出需求。"
      });
      return;
    }
    let inspection: ReturnType<typeof inspectWpsUndoRecord>;
    try {
      inspection = inspectWpsUndoRecord(getWpsApplication(), executionId);
    } catch {
      inspection = { kind: "unknown", detail: "读取原表现场失败。" };
    }
    if (inspection.kind === "completed") {
      advanceTaskResume(message.id, taskResumePatchForEvidence({ kind: "completed", executionId }));
      append({
        role: "system",
        text: "核对结果：原表目标区域与这次执行的**写后**状态一致 —— 写入其实已经完成，不会再写一遍。"
      });
      return;
    }
    if (inspection.kind === "not_started") {
      advanceTaskResume(message.id, taskResumePatchForEvidence({ kind: "not_started", executionId }));
      append({
        role: "system",
        text: "核对结果：原表目标区域仍是写入**之前**的状态 —— 这次写入确实没有开始。可以重新生成预览后再确认写入。"
      });
      return;
    }
    if (inspection.kind === "conflict") {
      advanceTaskResume(message.id, taskResumePatchForEvidence({ kind: "needs_review", executionId }));
      append({
        role: "system",
        text: `核对结果：${inspection.detail} 任务保持挂起；请先核对原表，不要直接重跑。`
      });
      return;
    }
    append({
      role: "system",
      text: `无法自动核对这次写入：${inspection.detail} 任务保持挂起；请人工核对原表后再决定。`
    });
  };

  /**
   * 用户显式取消任务。取消是**终态**：恢复记录被改成 `unavailable`，
   * 之后无论怎样重开面板都不会再把它摆成"可以继续"。
   */
  const cancelInterruptedTask = (message: ChatMessage) => {
    const record = message.taskResume
      ? sanitizeTaskResumeRecord(message.taskResume)
      : null;
    if (!record || busy) return;
    updateMessages((current) => current.map((item) =>
      item.id === message.id
        ? {
            ...item,
            taskResume: withTaskResumeState(record, {
              phase: "answer_submitted",
              blocker: "none",
              resumability: "unavailable"
            }),
            clarification: item.clarification
              ? { ...item.clarification, status: "cancelled" as const }
              : item.clarification
          }
        : item
    ));
    if (clarification?.messageId === message.id) {
      setClarification(null);
      setClarificationDraft("");
    }
  };

  /**
   * 把**某条消息**上的恢复记录推进到新状态（消息没有记录时什么都不做）。
   *
   * 按 `messageId` 定位而不是按"最新一条"：同一屏里可能有好几件任务的卡，
   * 按位置推断就会把状态写到别人的卡上。
   */
  const advanceTaskResume = (
    messageIdValue: string | null,
    patch: Parameters<typeof withTaskResumeState>[1]
  ) => {
    if (!messageIdValue) return;
    updateMessages((current) => current.map((item) => {
      if (item.id !== messageIdValue || !item.taskResume) return item;
      const record = sanitizeTaskResumeRecord(item.taskResume);
      if (!record) return item;
      return { ...item, taskResume: withTaskResumeState(record, patch) };
    }));
  };

  /**
   * **同步**把某条消息上的恢复记录写进存储，并返回是否成功。
   *
   * 为什么不复用 `useEffect` 那条落盘路径：它在 commit 之后才跑，是**异步**的。
   * 而"第一次宿主写入之前，恢复标记必须已经在存储里"要的是**顺序保证** ——
   * 依赖 effect 的调度顺序等于没有保证：真在写入途中被关掉，存储里可能还是
   * 写前的 `not_written`，界面于是会对用户说"原表未写入"。
   *
   * 为什么用 `flushSync` 而不是读渲染闭包里的 `chatHistory`：同一个点击里
   * 前面刚 `append` 过活动消息，闭包里的值**还没包含它** —— 拿它整份覆盖会把
   * 那条消息抹掉（本轮就踩到过：执行过程时间线整段消失）。`flushSync` 里的
   * `current` 由 React 保证是最新值。
   *
   * 返回 `false` 表示宿主拒绝写入；此时内存也会回滚，保持"界面 == 存储"。
   * 调用方**必须**据此中止这次写入 —— 没有落盘的"恢复标记"在关闭之后并不存在。
   */
  const persistTaskResumeNow = (
    messageIdValue: string | null,
    patch: Parameters<typeof withTaskResumeState>[1]
  ): boolean => {
    if (!messageIdValue) return true;
    let before: WpsChatHistory | null = null;
    let next: WpsChatHistory | null = null;
    flushSync(() => {
      setChatHistory((current) => {
        before = current;
        const now = new Date().toISOString();
        next = {
          ...current,
          conversations: current.conversations.map((conversation) => {
            if (conversation.id !== current.activeConversationId) return conversation;
            return {
              ...conversation,
              messages: conversation.messages.map((item) => {
                if (item.id !== messageIdValue || !item.taskResume) return item;
                const record = sanitizeTaskResumeRecord(item.taskResume);
                if (!record) return item;
                return { ...item, taskResume: withTaskResumeState(record, patch) };
              }),
              updatedAt: now
            };
          })
        };
        return next;
      });
    });
    if (!next) return true;
    try {
      localStorage.setItem(WPS_CHAT_STORAGE_KEY, serializeHistoryForStorage(next));
      return true;
    } catch {
      // 存不下就把内存回滚：否则界面说"已登记"、存储里却什么都没有。
      const rollback = before;
      if (rollback) setChatHistory(rollback);
      return false;
    }
  };

  /**
   * 把用户刚提交的回答写进恢复记录。
   *
   * 这一步把"已提交答案"变成**可跨重载**的事实：即使紧接着断线，重开后卡片
   * 也带着这条回答，不必让用户重新回答同一个问题。
   */
  const settleTaskAnswer = (
    messageIdValue: string,
    confirmation: TaskResumeConfirmation,
    mode: "append" | "replace" = "append"
  ) => {
    const previous = messages.find((item) => item.id === messageIdValue)?.taskResume;
    const record = previous ? sanitizeTaskResumeRecord(previous) : null;
    const context = record ? confirmAgentTaskDecision(
      createAgentTaskContext(record.originalPrompt, record.confirmations), confirmation, mode
    ) : null;
    // Replacement must reset the recovery record, not just one outgoing
    // request. Otherwise the next question revives the old task and rules.
    const sanitized = record && context ? sanitizeTaskResumeRecord({
      ...record, originalPrompt: context.originalPrompt, confirmations: context.confirmations
    }) : null;
    const answered = sanitized ? withTaskResumeState(sanitized, { phase: "answer_submitted" }) : null;
    updateMessages((current) => current.map((item) => {
      if (item.id !== messageIdValue || !item.taskResume) return item;
      if (!answered) return item;
      return {
        ...item,
        taskResume: answered
      };
    }));
    return answered;
  };

  /** 续答那一轮失败时，把卡还原成"可恢复的中断卡"，并如实标明是连接中断。 */
  const markTaskInterrupted = (messageIdValue: string) => {
    updateMessages((current) => current.map((item) => {
      if (item.id !== messageIdValue) return item;
      const record = item.taskResume
        ? sanitizeTaskResumeRecord(item.taskResume)
        : null;
      return {
        ...item,
        taskResume: record
          ? withTaskResumeState(record, {
              blocker: "network",
              resumability: "replan"
            })
          : item.taskResume,
        clarification: item.clarification
          ? { ...item.clarification, status: "unconfirmed" as const }
          : item.clarification
      };
    }));
  };

  const cancelClarification = () => {
    if (!clarification || busy) return;
    updateMessages((current) => current.map((message) =>
      message.id === clarification.messageId && message.clarification
        ? { ...message, clarification: { ...message.clarification, status: "cancelled" } }
        : message
    ));
    setClarification(null);
    setClarificationDraft("");
  };

  const closeModelSettings = () => {
    setModelSettingsOpen(false);
    modelManagement.setApiKeyDraft("");
    modelManagement.setShowApiKey(false);
    modelManagement.setConnectionDraft(null);
    modelManagement.setPendingDeleteConnectionId(null);
    modelManagement.setSettingsFeedback("");
  };

  const openModelSettings = async (createConnection: boolean) => {
    setModelMenuOpen(false);
    setModelSettingsOpen(true);
    if (createConnection) {
      // 先展示可填写的表单，再异步发现 Excel/WPS 共用的已有连接。
      modelManagement.openConnectionCreator();
      void modelManagement.checkExistingConnections();
    } else {
      await modelManagement.openSettings();
    }
  };

  // 「编辑计划」后重新预览时给出可见标识：屏幕上这份已不再是模型最初给的那版。
  const pendingPlanMessage = pendingPlanMessageId
    ? messages.find((message) => message.id === pendingPlanMessageId)
    : undefined;
  const planRevised = pendingPlanMessage?.planPreview?.revised === true;

  /**
   * 目标行与"试算了什么"的真实状态。
   *
   * `writeFormulas` 是计划里真实记录的动作类型 —— 侧栏据此区分"公式已试算"
   * 与"固定值"，而不是靠文案猜测。跨多动作时只要有一个不是公式就按固定值说，
   * 不为好看而夸大。
   */
  const planTargetLine = pendingPlan
    ? pendingPlan.actions.length === 1
      ? `${pendingPlan.actions[0].sheetName}!${pendingPlan.actions[0].range}`
      : pendingPlan.actions.map((action) => `${action.sheetName}!${action.range}`).join("、")
    : "";
  const stagedWriteKind = stagedReview
    ? stagedReview.plan.actions.every((action) => action.type === "writeFormulas")
      ? "formula"
      : "fixed"
    : null;
  /**
   * 首屏「位置 / 修改前 / 修改后」预览。
   *
   * 纯计算、零宿主读取：只用已经在内存里的计划动作与来源快照。
   * 拿不到原值的格子如实标"原值未取得"，不编造。
   */
  const planChangePreview = pendingPlan
    ? buildPlanChangePreview(pendingPlan)
    : { actions: [], totalCells: 0 };

  /**
   * 审阅卡视图。**只在副本已生成后**才有 —— 没生成副本时展示的是"写入前确认"，
   * 那不是审阅卡，不能套用审阅中的标题与确认框。
   *
   * `reviewPaneNeedsResize` 是真实回读结论：为假时首屏一句调宽提示都不出现。
   */
  const reviewCard: ReviewCardView | null = pendingPlan && stagedReview
    ? wpsReviewCard({
      outcome: "reviewing",
      scope: {
        locations: reviewScopeLocations(pendingPlan.plan.actions),
        changes: reviewChangeSummary(pendingPlan.plan.actions),
        // 已备好的计划里 `totalCells` 是权威计数；测试另有一条守卫断言它与
        // 按范围重算的面积一致，防止两个口径悄悄漂移。
        totalCells: pendingPlan.totalCells
      },
      compatNotice: reviewCardCompatNotice("WPS"),
      paneNeedsResize: reviewPaneNeedsResize,
      copy: null,
      diagnostics: wpsReviewDiagnostics(stagedReview, reviewPaneNeedsResize, developmentWrite)
    })
    : null;

  /** 「修改前 → 修改后」预览体。审阅中会被折叠，所以单独抽出来复用。 */
  const planChangePreviewBodyFor = (preview: ReturnType<typeof buildPlanChangePreview>) => (
    <>
      <strong className="wps-plan-change-title">修改前 → 修改后</strong>
      {preview.actions.map((action) => (
        <div className="wps-plan-change-block" key={action.key}>
          <span className="wps-plan-change-head">
            {action.label} · {action.sheetName}!{action.range}
            <small>共 {action.totalCells} 格</small>
          </span>
          {action.formatDescription != null ? (
            <p className="wps-plan-change-format">{action.formatDescription}</p>
          ) : (
            <table className="wps-plan-change-table">
              <thead>
                <tr>
                  <th scope="col">位置</th>
                  <th scope="col">修改前</th>
                  <th scope="col">修改后</th>
                </tr>
              </thead>
              <tbody>
                {action.rows.map((row) => (
                  <tr key={row.key} className={row.changed ? "changed" : undefined}>
                    <th scope="row">{row.address}</th>
                    <td>
                      {row.before.kind === "value"
                        ? row.before.text
                        : <span className="wps-plan-change-unknown">原值未取得</span>}
                    </td>
                    <td>{row.after.kind === "value" ? row.after.text : "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
          {action.truncated && (
            <small className="wps-plan-change-more">
              此处仅展示代表性样例（前几行、前几列）；完整内容可展开下方「查看预计结果」。
            </small>
          )}
        </div>
      ))}
    </>
  );

  const planCard = pendingPlan && (
    <section className="wps-plan-card" aria-label="WPS 写入计划预览">
      {/* 执行已由结果审阅通道自动完成：这张卡是**本次执行的记录**，不是确认入口。 */}
      <span className="clarification-eyebrow">本次执行计划</span>
      <strong>{pendingPlan.plan.title}</strong>
      <p>{pendingPlan.plan.summary}</p>
      {planRevised && (
        <p className="wps-plan-revised">已按你的修改更新；写入前请重新核对下面的目标与内容。</p>
      )}
      {/*
        开发环境提示**只在诊断入口**展示：它是排查用的，不是"要不要按确认"需要的。
        见下面的「诊断详情」。
      */}
      {/*
        目标与规模是决策必需信息，留在主卡（默认可见）。
        长说明不再与按钮同排 —— 那正是真机上按钮被压到逐字折行的原因。
      */}
      <div className="wps-plan-target" title={planTargetLine}>
        <span>{planTargetLine}</span>
        <small>{pendingPlan.totalCells} 格 · {pendingPlan.actions.length} 个操作</small>
      </div>
      <div className="wps-plan-meta">
        <span>{SESSION_RECOVERY_LABEL}</span>
      </div>
      {/*
        决策必需的事实：**到底改了什么**。
        未生成副本时留在首屏 —— 用户要在按下「生成右侧副本并排审阅」之前做决定。
        **审阅中**则收进折叠：那时首屏留给审阅卡（标题/说明/范围/确认框/按钮），
        否则"改了什么"会把真正的决策项挤下去。
        数据全部来自已在内存里的计划与授权读取快照 —— 这里**不做**额外宿主读取。
      */}
      {stagedReview ? (
        <details className="wps-plan-disclosure">
          <summary>查看本次计划的改动（修改前 → 修改后）</summary>
          <div className="wps-plan-change-preview" aria-label="本次修改的修改前与修改后">
            {planChangePreviewBodyFor(planChangePreview)}
          </div>
        </details>
      ) : (
        <div className="wps-plan-change-preview" aria-label="本次修改的修改前与修改后">
          {planChangePreviewBodyFor(planChangePreview)}
        </div>
      )}
      {/*
        审阅中：**首屏只留主卡**。
        标题 / 说明 / 范围 / 兼容模式必要说明 / 可收起提示 / 确认框 / 按钮 ——
        其余（重启提示、副本标识、关闭诊断、磁盘清理）全部下沉到「诊断详情」。
      */}
      {/*
        结果状态由固定操作条（`ResultReviewBar`）承担：它常驻在滚动区之外，
        所以"确认变更"这个入口不会随对话滚走。这里不再重复一遍状态，
        也不再放任何确认按钮 —— 两个入口会让人以为要确认两次。
      */}
      {pendingPlan.plan.assumptions.length > 0 && (
        <div className="wps-plan-notes">
          <strong>{PLAN_NOTE_SECTIONS.assumptions}</strong>
          <ul className="wps-plan-assumptions">
            {pendingPlan.plan.assumptions.map((item) => <li key={item}>{item}</li>)}
          </ul>
        </div>
      )}
      {/*
        正文默认收起，但**默认收起不等于删除信息**：授权范围、覆盖风险、
        未知类别策略与实际写入类型都仍可从这两处展开查到。
        公式用等宽字 + 横向滚动，绝不在窄列里任意折断（`overflow-wrap:anywhere`
        曾让一条长公式占掉多屏高度）。
      */}
      <details className="wps-plan-disclosure">
        <summary>查看预计结果（尚未写入）</summary>
        <p className="wps-plan-disclosure-note">
          以下是本地确定性计算给出的<strong>预计</strong>值，尚未写入任何工作表。
        </p>
        <div className="wps-plan-preview">
          {pendingPlan.actions.map((action) => (
            <div key={`expected:${action.actionIndex}:${action.sheetName}:${action.range}`}>
              <strong>{actionTypeLabel(action.type)} · {action.sheetName}!{action.range}</strong>
              {isFormatActionType(action.type) ? (
                <div className="wps-plan-values" aria-label={`${action.sheetName}!${action.range} 待应用格式`}>
                  <p>{describeWpsFormatAction(action.formatAction ?? (action as unknown as Parameters<typeof describeWpsFormatAction>[0]))}</p>
                </div>
              ) : (
                <div className="wps-plan-values">
                  <table><tbody>
                    {action.values.map((row, rowIndex) => (
                      <tr key={rowIndex}>
                        <th scope="row">{rowIndex + 1}</th>
                        {row.map((value, columnIndex) => (
                          <td key={columnIndex}>{value === null ? "空单元格" : value === "" ? "空字符串" : String(value)}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody></table>
                </div>
              )}
            </div>
          ))}
        </div>
      </details>
      <details className="wps-plan-disclosure">
        <summary>查看公式与完整计划</summary>
        <div className="wps-plan-actions-list">
          {pendingPlan.actions.map((action) => (
            <div key={`${action.actionIndex}:${action.sheetName}:${action.range}`}>
              <strong>{actionTypeLabel(action.type)}</strong>
              <span>{action.sheetName}!{action.range}</span>
              <small>{action.rowCount} 行 × {action.columnCount} 列</small>
              {isFormatActionType(action.type) ? (
                <div className="wps-plan-values" aria-label={`${action.sheetName}!${action.range} 待应用格式`}>
                  <p>{describeWpsFormatAction(action.formatAction ?? (action as unknown as Parameters<typeof describeWpsFormatAction>[0]))}</p>
                </div>
              ) : (
                <div className="wps-plan-values" aria-label={`${action.sheetName}!${action.range} 待写内容预览`}>
                  <table><tbody>
                    {action.preview.map((row, rowIndex) => (
                      <tr key={rowIndex}>
                        <th scope="row">{rowIndex + 1}</th>
                        {row.map((value, columnIndex) => (
                          <td key={columnIndex}>{value === null ? "空单元格" : value === "" ? "空字符串" : String(value)}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody></table>
                  {(action.rowCount > 5 || action.columnCount > 5) && <small>此处仅显示前 5 行、前 5 列；编辑计划可查看完整矩阵。</small>}
                </div>
              )}
            </div>
          ))}
        </div>
      </details>
      {stagedReview && (
        <details className="wps-plan-disclosure">
          <summary>查看副本回读结果（已试算）</summary>
          <p className="wps-plan-disclosure-note">
            以下是<strong>副本</strong>试算后的真实回读值 —— 与上面的"预计"是两回事，原表尚未写入。
            {stagedWriteKind === "formula" ? "本次副本已试算公式并回读计算值。" : "本次副本已写入固定值并回读。"}
            副本共标记 {stagedReview.changedCells.length} 处差异；此处是副本回读结果，不是预计值。
          </p>
          <div className="wps-plan-preview">
            {stagedReview.copyAfter.map((after, index) => (
              <div key={`readback:${index}`}>
                <strong>回读值 · {stagedReview.plan.actions[index]?.sheetName}!{stagedReview.plan.actions[index]?.range}</strong>
                <div className="wps-plan-values">
                  <table><tbody>
                    {after.values.map((row, rowIndex) => (
                      <tr key={rowIndex}>
                        <th scope="row">{rowIndex + 1}</th>
                        {row.map((value, columnIndex) => (
                          <td key={columnIndex}>{value === null ? "空单元格" : value === "" ? "空字符串" : String(value)}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody></table>
                  {!after.valuesCaptured && <small>宿主本次未交出显示值。</small>}
                </div>
              </div>
            ))}
          </div>
        </details>
      )}
      {pendingPlan.plan.warnings.length > 0 && (
        <div className="wps-plan-notes warning">
          <strong>{PLAN_NOTE_SECTIONS.warnings}</strong>
          <ul className="wps-plan-warnings">
            {pendingPlan.plan.warnings.map((warning) => <li key={warning}>{warning}</li>)}
          </ul>
        </div>
      )}
      {planEditDraft !== null && (
        <div className="wps-plan-editor">
          <ControlledScriptEditor getAuthorizedSheets={scriptAuthorizedSheets}
            draft={planEditDraft} onApply={value => { setPlanEditDraft(value); setPlanEditError(null); }} disabled={busy} />
          <label htmlFor="wps-plan-edit-json">编辑写入计划 JSON</label>
          <textarea
            id="wps-plan-edit-json"
            aria-label="编辑写入计划 JSON"
            rows={12}
            value={planEditDraft}
            disabled={busy}
            onChange={(event) => { setPlanEditDraft(event.target.value); setPlanEditError(null); }}
          />
          {planEditError && <p role="alert">{planEditError}</p>}
        </div>
      )}
      {/*
        兼容模式说明：格式恢复与"为什么加载项读不回副本"的技术解释。
        默认收起 —— 它是"为什么"，不是"要不要按确认"。
      */}
      {reviewCard && reviewCard.compatDetail.length > 0 && (
        <details className="wps-plan-disclosure">
          <summary>{REVIEW_CARD_COMPAT_DETAIL_TITLE}</summary>
          <ul className="wps-plan-diagnostics-list">
            {reviewCard.compatDetail.map((entry) => <li key={entry}>{entry}</li>)}
          </ul>
        </details>
      )}
      {/*
        诊断详情：重启/开发环境提示、排列指引、副本标识、关闭诊断、磁盘清理。
        **只在**这里出现 —— 首屏不摆排查材料。
      */}
      {reviewCard && reviewCardHasDiagnostics(reviewCard.diagnostics) && (
        <details className="wps-plan-disclosure">
          <summary>{REVIEW_CARD_DIAGNOSTICS_TITLE}</summary>
          {(["restart", "arrangement", "identity", "close", "cleanup"] as const).map((group) =>
            reviewCard.diagnostics[group].length > 0 && (
              <ul key={group} className={`wps-plan-diagnostics-list ${group}`}>
                {reviewCard.diagnostics[group].map((entry) => <li key={entry}>{entry}</li>)}
              </ul>
            ))}
        </details>
      )}
      {/*
        操作区**只放按钮**：长说明留在上面的 `.wps-native-review-status`。
        宽窗格水平排列、窄窗格整行纵排（见 styles.css），按钮文案不逐字断行。
      */}
    </section>
  );

  /**
   * 已执行计划的摘要卡。
   *
   * 与执行前的计划卡**同一套事实**（目标、规模、采用了哪些口径、格式动作具体改什么、
   * 行高列宽影响整行），只是没有任何按钮：执行已经发生，确认入口属于结果操作条。
   * 没有它，用户事后只能看到"已执行"，看不到"按什么口径改的"。
   */
  const executedPlanCard = executedPlan && (() => {
    const planned = executedPlan.plan;
    const preview = buildPlanChangePreview(planned);
    const targets = planned.actions.length === 1
      ? `${planned.actions[0].sheetName}!${planned.actions[0].range}`
      : planned.actions.map(action => `${action.sheetName}!${action.range}`).join("、");
    return (
      <section className="wps-plan-card" aria-label="本次执行的计划摘要">
        <span className="clarification-eyebrow">本次执行计划</span>
        <strong>{planned.plan.title}</strong>
        <p>{planned.plan.summary}</p>
        {executedPlan.revised && <p className="wps-plan-revised">已按你的修改更新。</p>}
        <div className="wps-plan-target" title={targets}>
          <span>{targets}</span>
          <small>{planned.totalCells} 格 · {planned.actions.length} 个操作</small>
        </div>
        <div className="wps-plan-meta"><span>{SESSION_RECOVERY_LABEL}</span></div>
        <div className="wps-plan-change-preview" aria-label="本次修改的修改前与修改后">
          {planChangePreviewBodyFor(preview)}
        </div>
        {planned.plan.assumptions.length > 0 && (
          <div className="wps-plan-notes">
            <strong>{PLAN_NOTE_SECTIONS.assumptions}</strong>
            <ul className="wps-plan-assumptions">
              {planned.plan.assumptions.map(item => <li key={item}>{item}</li>)}
            </ul>
          </div>
        )}
        {planned.plan.warnings.length > 0 && (
          <div className="wps-plan-notes warning">
            <strong>{PLAN_NOTE_SECTIONS.warnings}</strong>
            <ul className="wps-plan-warnings">
              {planned.plan.warnings.map(warning => <li key={warning}>{warning}</li>)}
            </ul>
          </div>
        )}
        {/*
          完整计划的**只读**入口。执行已经发生，所以这里不再提供编辑 ——
          但"到底按哪份动作写的"必须可查：这是事后复核与排查的唯一依据。
        */}
        <details className="wps-plan-disclosure">
          <summary>查看本次执行的完整计划</summary>
          <pre className="wps-plan-json" aria-label="本次执行计划 JSON">
            {JSON.stringify(planned.plan, null, 2)}
          </pre>
        </details>
      </section>
    );
  })();

  const progressCard = (progress || writeProgress) && (
    <section className="agent-progress" aria-label="任务执行过程">
      <div className="agent-progress-heading">
        <strong>{writeProgress ? "写入过程" : "任务过程"}</strong>
        <span>{writeProgress ? SESSION_RECOVERY_LABEL : "处理中"}</span>
      </div>
      <div className="agent-progress-current">
        <span className="progress-dot" aria-hidden="true" />
        <strong>{writeProgress ? "执行与校验" : "处理中"}</strong>
        <span>{writeProgress ?? progress?.message}</span>
        {/*
          「暂停」只在**真的能停**时出现（共享判据）。以前只要 `progress` 非空就画，
          于是出现过"按钮在那儿、点下去没反应"。不能停时如实说明原因，不留死按钮。
        */}
        {cardPauseButton.canPause
          ? <button type="button" onClick={() => pauseCurrentTurn("card")}>暂停</button>
          : cardPauseButton.kind === "unavailable" && !writeProgress
            ? <span className="agent-progress-note" role="status">{cardPauseButton.reason}</span>
            : null}
      </div>
    </section>
  );

  return (
    <section className="wps-chat-shell chat-shell">
      {targetGuard.notice && !onOpenScope && <AuthorizedTargetConfirmation notice={targetGuard.notice}
        error={targetGuard.error} busy={targetGuard.confirming}
        onConfirm={(choice, selected) => void targetGuard.confirm(choice, selected)} onCancel={targetGuard.cancel}
        onRecheck={() => void targetGuard.recheck()} />}
      <header
        className={`chat-header${modelMenuOpen || moreMenuOpen ? " menu-open" : ""}`}
      >
        <img className="brand-mark" src={brandLogoUrl} alt="" />
        <div className="brand-copy">
            <strong className="app-logo" {...logoLongPress}>Excel Bro</strong>
            {brandStatusText && <span className={modelCatalogFailed ? "offline" : "model-online online"}>
              <i aria-hidden="true" />
              {brandStatusText}
            </span>}
        </div>
        <div className="header-actions">
          <div
            className="model-picker"
            onBlur={(event) => {
              if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
                setModelMenuOpen(false);
              }
            }}
          >
            <button
              className="model-picker-trigger"
              type="button"
              disabled={busy}
              aria-haspopup="menu"
              aria-expanded={modelMenuOpen}
              aria-label={modelNotConfigured || modelCatalogFailed ? "添加模型" : "选择或添加模型"}
              title={selectedModel ? `选择或添加模型 · 当前模型：${selectedModel.label}` : "选择或添加模型"}
              onClick={() => {
                const nextOpen = !modelMenuOpen;
                closeHeaderSurfaces();
                setModelMenuOpen(nextOpen);
              }}
            >
              <ModelPickerIcon />
            </button>
            {modelMenuOpen && (
              <div className="model-menu" role="menu" aria-label="模型">
                <span className="model-menu-title">选择模型</span>
                {modelCatalogLoading && (
                  <p className="eb-model-menu-note">正在读取模型配置…</p>
                )}
                {modelCatalogFailed && (
                  <p className="eb-model-menu-note">
                    暂时无法取得已保存的模型连接。可先填写模型信息；测试和保存前需启动 Excel Bro 后台。已有配置会在连接恢复后显示。
                  </p>
                )}
                {modelNotConfigured && (
                  <p className="eb-model-menu-note">
                    尚未配置模型。添加一个模型连接后，即可通过对话查询和分析表格。
                  </p>
                )}
                {models.map((model) => (
                  <button
                    key={model.id}
                    type="button"
                    role="menuitemradio"
                    aria-checked={model.id === effectiveModelId}
                    disabled={!agentModelAvailable(model)}
                    title={!agentModelAvailable(model) ? agentModelUnavailableReason(model) : undefined}
                    onClick={() => {
                      if (!isRealModelId(model.id)) return;
                      modelManagement.selectModel(model.id);
                      setModelMenuOpen(false);
                    }}
                  >
                    <i aria-hidden="true" />
                    <span>{model.label}{agentModelAvailable(model) ? "" : `（${agentModelUnavailableReason(model)}）`}</span>
                    {model.id === effectiveModelId && <b aria-hidden="true">✓</b>}
                  </button>
                ))}
                <div className="model-menu-actions">
                  <button type="button" role="menuitem" onClick={() => void openModelSettings(true)}>
                    ＋ 添加模型连接
                  </button>
                  <button type="button" role="menuitem" onClick={() => void openModelSettings(false)}>
                    管理模型连接
                  </button>
                  {modelCatalogFailed && (
                    <button
                      type="button"
                      role="menuitem"
                      onClick={() => void refreshModelCatalog()}
                    >
                      重试读取模型配置
                    </button>
                  )}
                </div>
              </div>
            )}
          </div>
          <button
            className="header-button new-chat-entry"
            type="button"
            disabled={busy}
            title="新对话"
            aria-label="新对话"
            onClick={startNewConversation}
          >
            <span aria-hidden="true">＋</span>
          </button>
          <button
            className="header-button tools-entry"
            type="button"
            disabled={busy}
            title="我的工具"
            aria-label="我的工具"
            onClick={() => {
              const nextOpen = !toolsOpen;
              closeHeaderSurfaces();
              setToolsOpen(nextOpen);
            }}
          >
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <path d="M8.5 7V5.8c0-1 .8-1.8 1.8-1.8h3.4c1 0 1.8.8 1.8 1.8V7" />
              <rect x="3" y="7" width="18" height="12.5" rx="2.5" />
              <path d="M3.5 11.5h17M10 11.5v2h4v-2" />
            </svg>
          </button>
          <button
            className="header-button folder-backup-entry"
            type="button"
            disabled={busy}
            title="备份与恢复"
            aria-label="备份与恢复"
            onClick={() => {
              const nextOpen = !backupOpen;
              closeHeaderSurfaces();
              setBackupOpen(nextOpen);
            }}
          >
            <span aria-hidden="true">↶</span>
            {undoRecord && <b>1</b>}
          </button>
          <button
            className={`header-button pet-toggle${petVisible ? " active" : ""}`}
            type="button"
            aria-label={petVisible ? "隐藏格仔" : "显示格仔"}
            aria-pressed={petVisible}
            title={petVisible ? "隐藏格仔" : "显示格仔"}
            onClick={() => setPetVisible((visible) => !visible)}
          >
            <svg aria-hidden="true" viewBox="0 0 24 24">
              <ellipse cx="6.3" cy="10" rx="1.9" ry="2.4" />
              <ellipse cx="10.4" cy="7.1" rx="1.9" ry="2.5" />
              <ellipse cx="14.6" cy="7.1" rx="1.9" ry="2.5" />
              <ellipse cx="18.7" cy="10" rx="1.9" ry="2.4" />
              <path d="M7.4 16c0-2.7 2-4.5 4.6-4.5s4.6 1.8 4.6 4.5c0 2.4-1.9 4-4.6 4s-4.6-1.6-4.6-4Z" />
            </svg>
          </button>
          <div className="view-menu-wrap">
            <button
              className="header-button"
              type="button"
              aria-label="更多"
              title="更多"
              aria-haspopup="menu"
              aria-expanded={moreMenuOpen}
              onClick={() => {
                const nextOpen = !moreMenuOpen;
                closeHeaderSurfaces();
                setMoreMenuOpen(nextOpen);
              }}
            >
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <circle cx="5" cy="12" r="1.6" />
                <circle cx="12" cy="12" r="1.6" />
                <circle cx="19" cy="12" r="1.6" />
              </svg>
            </button>
            {moreMenuOpen && (
              <>
                <div className="more-menu-backdrop" aria-hidden="true" onClick={() => setMoreMenuOpen(false)} />
                <div className="view-menu" role="menu" aria-label="更多">
                  <button type="button" role="menuitem" aria-label="历史对话" onClick={() => { setMoreMenuOpen(false); setHistoryOpen(true); }}>
                    <strong>历史对话</strong><span>查看当前 WPS 会话记录</span>
                  </button>
                  <button type="button" role="menuitem" onClick={() => { setMoreMenuOpen(false); setRuleManagerOpen(true); }}>
                    <strong>EB 函数说明</strong><span>查看内置 =EB() 函数用法</span>
                  </button>
                  <button type="button" role="menuitem" onClick={() => { setMoreMenuOpen(false); setThemeOpen(true); }}>
                    <strong>外观与壁纸</strong><span>设置主题或自定义背景</span>
                  </button>
                  {import.meta.env.DEV && (
                    <button
                      type="button"
                      role="menuitem"
                      onClick={() => {
                        setMoreMenuOpen(false);
                        setWideViewProbeError(null);
                        wideViewChannelStopRef.current?.();
                        const nonce = createProbeNonce();
                        if (nonce) {
                          wideViewChannelStopRef.current = startParentChannelProbe(nonce, setWideViewChannelStatus);
                        } else {
                          setWideViewChannelStatus("unsupported");
                        }
                        try {
                          const probe = probeWideViewDialog(getWpsApplication(), nonce ?? undefined);
                          setWideViewProbe(probe);
                          if (probe.status !== "requested") {
                            wideViewChannelStopRef.current?.();
                            setWideViewChannelStatus(null);
                          }
                        } catch (error) {
                          wideViewChannelStopRef.current?.();
                          setWideViewChannelStatus(null);
                          setWideViewProbe(null);
                          setWideViewProbeError(error instanceof Error ? error.message : String(error));
                        }
                      }}
                    >
                      <strong>探测独立宽窗口</strong><span>开发验证：保留当前对话，不读取工作表</span>
                    </button>
                  )}
                </div>
              </>
            )}
          </div>
        </div>
      </header>

      {import.meta.env.DEV && (wideViewProbe || wideViewProbeError) && (
        <aside className={`poc-panel-result ${wideViewProbe?.status ?? "error"}`} role="status" aria-label="宽窗口探测结果">
          <div className="poc-probe-result-header">
            <strong>{wideViewProbe?.status === "requested" ? "宽窗口请求已发出，请检查实际布局" : "宽窗口探测未成功"}</strong>
            <button type="button" aria-label="关闭探测结果" onClick={() => { wideViewChannelStopRef.current?.(); setWideViewProbe(null); setWideViewProbeError(null); setWideViewChannelStatus(null); }}>×</button>
          </div>
          <p>宿主版本：{wideViewProbe?.hostVersion ?? "未报告"}；API：{wideViewProbe?.presentApis.join("、") || "无"}；打开方式：{wideViewProbe?.openedWith ?? "未成功"}</p>
          {wideViewChannelStatus && (
            <p>跨窗口内存通信：{{ waiting: "正在握手", connected: "已双向连通", unsupported: "不可用", timeout: "超时" }[wideViewChannelStatus]}</p>
          )}
          {(wideViewProbe?.error || wideViewProbeError) && <p className="poc-panel-error-detail">{wideViewProbe?.error ?? wideViewProbeError}</p>}
          <p className="poc-panel-note">此结果只在当前页面显示；目标是在宽窗口内同时呈现双表格与对话。请核对弹窗内的通信结果，并确认能手动关闭返回。</p>
        </aside>
      )}

      {toolsOpen && (
        <aside className="tool-drawer" aria-label="我的工具">
          <div className="tool-drawer-header">
            <div><strong>我的工具</strong><span>选择后仍会先校验并预览，不会立即修改工作簿</span></div>
            <button type="button" aria-label="关闭" onClick={() => setToolsOpen(false)}>×</button>
          </div>
          <div className="tool-library">
            {tools.length === 0 && queryTools.length === 0 ? (
              <div className="tool-empty">
                <i>▦</i><strong>工具箱还是空的</strong><span>执行并验证一个计划后，可以把它保存到这里。</span>
              </div>
            ) : (
              <>
                <div className="tool-library-intro">
                  <span>工具箱</span><strong>今天想用哪一个？</strong>
                  <p>这里显示当前 WPS 加载项保存的工具；运行前仍会按 WPS 白名单和本次授权重新校验。</p>
                </div>
                <div className="tool-card-grid">
                  {tools.map((tool) => (
                    <article className="tool-card-wrap" key={tool.id}>
                      <button className="tool-card" type="button" onClick={() => void prepareSavedWorkflow(tool)}>
                        <span className="tool-card-icon tone-green">▦</span>
                        <span className="tool-card-copy">
                          <small>工作流</small><strong>{tool.name}</strong><span>{tool.description}</span>
                        </span>
                      </button>
                      <div className="tool-card-footer">
                        <span className="tool-card-meta">{tool.planTemplate.actions.length} 个步骤 · 点击后预览</span>
                        <span className="tool-card-actions">
                          <button type="button" className="tool-card-edit" onClick={() => openToolPlanEditor(tool)}>编辑计划</button>
                          <button type="button" className="tool-card-delete" onClick={() => deleteSavedTool(tool)}>删除</button>
                        </span>
                      </div>
                    </article>
                  ))}
                  {queryTools.map((tool) => (
                    <article className="tool-card-wrap" key={tool.id}>
                      <button
                        className="tool-card"
                        type="button"
                        onClick={() => void runSavedQuery(tool)}
                      >
                        <span className="tool-card-icon tone-blue">查</span>
                        <span className="tool-card-copy">
                          <small>固化查询</small><strong>{tool.name}</strong><span>{tool.description}</span>
                        </span>
                      </button>
                      <div className="tool-card-footer"><span className="tool-card-meta">本地确定性执行 · 模型调用 0 次</span></div>
                    </article>
                  ))}
                </div>
              </>
            )}
          </div>
        </aside>
      )}

      {toolPlanEdit && (
        <aside className="tool-drawer plan-editor-drawer" aria-label="编辑工具计划模板">
          <div className="tool-drawer-header">
            <div>
              <strong>编辑工具计划模板</strong>
              <span>保存会检查 JSON 格式与白名单结构；运行前仍会按当前授权重新校验并预览</span>
            </div>
            <button type="button" aria-label="关闭" onClick={() => setToolPlanEdit(null)}>×</button>
          </div>
          <div className="plan-editor-body">
            <ControlledScriptEditor getAuthorizedSheets={scriptAuthorizedSheets}
              draft={toolPlanEdit.draft} onApply={value => setToolPlanEdit({ ...toolPlanEdit, draft: value, error: null })} disabled={busy} />
            <textarea
              className="plan-editor-textarea"
              aria-label="编辑工具计划 JSON"
              rows={14}
              value={toolPlanEdit.draft}
              disabled={busy}
              onChange={(event) => setToolPlanEdit({ ...toolPlanEdit, draft: event.target.value, error: null })}
            />
            {toolPlanEdit.error && <p role="alert" className="plan-editor-error">{toolPlanEdit.error}</p>}
            <div className="inline-plan-buttons">
              <button type="button" className="secondary-button" onClick={() => setToolPlanEdit(null)}>放弃修改</button>
              <button type="button" className="primary-button" onClick={applyEditedToolTemplate}>保存</button>
            </div>
          </div>
        </aside>
      )}

      {backupOpen && (
        <aside className="tool-drawer" aria-label="备份与恢复">
          <div className="tool-drawer-header">
            <div><strong>备份与恢复</strong><span>WPS 写入前保存完整快照，并在撤销前核对后续修改</span></div>
            <button type="button" aria-label="关闭" onClick={() => setBackupOpen(false)}>×</button>
          </div>
          <div className="wps-backup-content">
            {undoRecord ? (
              <section className="wps-backup-card">
                <span>最近恢复点</span><strong>{undoRecord.planTitle}</strong>
                {/*
                  标明它属于哪一次执行：恢复点是**上一轮**写入留下的，
                  不是当前这轮请求的下一步。短 id 足够区分，长 id 不必要地占地方。
                */}
                <small>所属执行 {undoRecord.executionId.slice(0, 8)}</small>
                <small>{undoRecord.state === "undo_conflict" ? "检测到后续修改，需要先处理冲突" : "当前可安全撤销"}</small>
                {/*
                  处理中原入口改字 + 禁用：不新增过程卡片，也不追加逐步进度消息。
                  重复点击在这一步就被挡掉（`undoingExecutionId` 命中即禁用）。
                */}
                <button type="button"
                  disabled={busy || undoRecord.state === "undo_conflict" || undoingExecutionId === undoRecord.executionId}
                  onClick={() => void undoExecution(undoRecord.executionId)}>
                  {undoingExecutionId === undoRecord.executionId ? "↶ 正在撤销…" : "↶ 撤销本次执行"}
                </button>
              </section>
            ) : (
              <div className="tool-empty"><i>↶</i><strong>还没有恢复点</strong><span>只有确认并完成写入后才会创建，不会生成空备份。</span></div>
            )}
            <WorksheetCheckpointRecoverySection
              controller={worksheetCheckpoint}
              workbookPath={checkpointWorkbookPath}
            />
          </div>
        </aside>
      )}

      {historyOpen && (
        <aside className="tool-drawer history-drawer" aria-label="历史对话">
          <div className="tool-drawer-header">
            <div><strong>历史对话</strong><span>新建对话不会覆盖之前的记录</span></div>
            <button type="button" aria-label="关闭" onClick={() => setHistoryOpen(false)}>×</button>
          </div>
          <div className="history-list">
            {[...chatHistory.conversations]
              .sort((left, right) => right.updatedAt.localeCompare(left.updatedAt))
              .map((conversation) => {
                const active = conversation.id === chatHistory.activeConversationId;
                const questionCount = conversation.messages.filter(
                  (message) => message.role === "user"
                ).length;
                return (
                  <article className={`history-item${active ? " active" : ""}`} key={conversation.id}>
                    <button
                      className="history-open-button"
                      type="button"
                      disabled={busy}
                      onClick={() => openConversation(conversation.id)}
                    >
                      <span className="history-title-row">
                        <strong>{conversation.title}</strong>
                        {active && <em>当前</em>}
                      </span>
                      <span>
                        {formatConversationTime(conversation.updatedAt)}
                        {questionCount > 0 ? ` · ${questionCount} 条提问` : " · 尚未提问"}
                      </span>
                    </button>
                    <button
                      className="history-delete-button"
                      type="button"
                      disabled={busy}
                      title={`删除「${conversation.title}」`}
                      aria-label={`删除「${conversation.title}」`}
                      onClick={() => deleteConversation(conversation.id)}
                    >
                      删除
                    </button>
                  </article>
                );
              })}
          </div>
          <button className="history-new-button" type="button" disabled={busy} onClick={startNewConversation}>＋ 新建对话</button>
        </aside>
      )}

      <div className={`message-stream-shell${newConversationNotice ? " conversation-starting" : ""}`}>
      {newConversationNotice && (
        <div className="new-conversation-toast" role="status">
          {newConversationNotice.previousSaved ? "已新建对话" : "当前已是新对话"}
        </div>
      )}
      {copyError && (
        <div className="new-conversation-toast" role="status">
          复制未成功；可以选中文字后使用 Ctrl+C
        </div>
      )}
      <main ref={messageStreamRef} className="message-stream" aria-live="polite" onScroll={handleMessageStreamScroll}>
        {messages.length === 0 && (
          <article className="message-row assistant welcome">
            <div className="message-content">
              <div className="message-author">Excel Bro</div>
              <p className="message-text">
              {modelCatalogFailed
                ? "你好。点击上方模型图标开始连接；如果你已配置过模型，Excel Bro 后台连接后会自动显示。"
                : modelNotConfigured
                ? "你好。添加一个模型连接后，即可查询、分析或修改已授权的工作表。"
                : canWrite
                  ? "你好。直接告诉我你想查询、分析或修改什么；涉及写入时，我会在授权范围内完成校验和执行，再展示结果供你审阅。"
                  : `你好。你可以查询和分析工作表；写入前还需通过当前文件检查：${writeStatusReason}`}
              </p>
            </div>
          </article>
        )}

        {messages.map((message) => (
          <Fragment key={message.id}>
            <article className={`message-row ${message.role}`} data-message-id={message.id}>
              <div className="message-content">
              {message.role === "assistant" && (
                <div className="message-author">
                  Excel Bro
                  {message.resultContext ? <em>本地工具</em>
                    : message.clarification ? <em>需求确认</em>
                    : message.activityLog || message.planPreview || message.verification ? null
                    : <em>Agent</em>}
                </div>
              )}
              {message.resultContext ? (
                <section className="query-result-card">
                  <p className="message-text">{message.text}</p>
                  <div className="wps-query-result-meta">
                    <strong>{message.resultContext.title}</strong>
                    <span>显示 {Math.min(20, message.resultContext.rows.length)} / {message.resultContext.rows.length} 行</span>
                    {message.resultContext.sourceSheets.length > 0 && (
                      <small>数据来源：{message.resultContext.sourceSheets.join("、")}</small>
                    )}
                  </div>
                  <div className="preview-table-wrap">
                    <table>
                      <thead>
                        <tr>
                          {message.resultContext.headers.map((header) => <th key={header}>{header}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {message.resultContext.rows.slice(0, 20).map((row, rowIndex) => (
                          <tr key={rowIndex}>
                            {row.map((cell, columnIndex) => (
                              <td key={columnIndex}>{cell == null ? "" : String(cell)}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {message.resultContext.warnings.length > 0 && (
                    <ul className="wps-query-result-warnings">
                      {message.resultContext.warnings.map((warning, index) => <li key={`${index}:${warning}`}>{warning}</li>)}
                    </ul>
                  )}
                </section>
              ) : message.text && !message.planningPause &&
                  !prePlanningPause.entries.some((entry) => entry.messageId === message.id) ? (
                <>
                  <p className="message-text">{message.text}</p>
                  {/*
                    失败/冲突的诊断详情：默认折叠，用户点开才看。
                    首屏只放真实状态的结论，不被诊断文字淹没；但结论与详情同在一条
                    消息里，不存在"静默隐藏"。
                  */}
                  {message.errorDetail && (
                    <details className="message-error-detail">
                      <summary>查看错误详情</summary>
                      <pre>{message.errorDetail}</pre>
                    </details>
                  )}
                </>
              ) : null}
              {message.localResult && (
                <LocalResultCardView
                  card={message.localResult}
                  load={(resultRef) =>
                    message.localResultBinding
                      ? loadWpsBridgeResult(message.localResultBinding, resultRef)
                      : Promise.resolve(null)
                  }
                />
              )}
              {message.clarification && (
                <section className={`clarification-card ${message.clarification.status}`} aria-label="需求澄清">
                  <div className="clarification-heading">
                    <div>
                      <span>需要确认</span>
                      <strong>{message.clarification.summary}</strong>
                    </div>
                    <em>{message.clarification.scopeLabel}</em>
                  </div>
                  <p className="clarification-question">{message.clarification.question}</p>
                  <p className="clarification-reason">{message.clarification.reason}</p>
                  {message.clarification.status === "pending" && clarification?.messageId === message.id ? (
                    <>
                      {clarification.localMapping ? (
                        <LocalMappingResolution
                          key={`${message.id}:${clarification.localMappingRevision ?? 0}`}
                          rule={clarification.localMapping.rule}
                          values={clarification.localMapping.values}
                          allowExactMapping={clarification.localMapping.allowExactMapping}
                          reason={clarification.localMapping.reason}
                          blankCount={clarification.localMapping.blankCount}
                          busy={busy || Boolean(pendingPlan)}
                          checkScope={() => checkLocalMappingScope(clarification)}
                          onResolve={(rule) => resolveLocalMapping(rule)}
                          onCancel={cancelClarification}
                        />
                      ) : (<>
                      <div className="clarification-options">
                        {message.clarification.options.map((option, index) => (
                          <button type="button" key={index} disabled={busy} onClick={() => void resolveClarification(index)}>
                            <strong>{option.label}</strong>
                            <span>{option.description}</span>
                          </button>
                        ))}
                      </div>
                      {message.clarification.fields?.length ? (
                        <ClarificationFields fields={message.clarification.fields}
                          values={clarificationFieldDrafts} busy={busy}
                          onChange={setClarificationFieldDrafts}
                          onSubmit={(answer) => void resolveClarificationWithText(answer)} />
                      ) : <>
                      <span className="clarification-custom-hint">其他想法也可以直接说明</span>
                      <div className="clarification-custom-input">
                        <textarea
                          aria-label="其他想法"
                          value={clarificationDraft}
                          placeholder="补充你希望我怎么处理…"
                          onChange={(event) => setClarificationDraft(event.target.value)}
                          onKeyDown={(event) => {
                            if (event.key === "Enter" && !event.shiftKey) {
                              event.preventDefault();
                              void resolveClarificationWithText();
                            }
                          }}
                        />
                        <button type="button" disabled={!clarificationDraft.trim() || busy} onClick={() => void resolveClarificationWithText()}>发送</button>
                      </div>
                      </>}
                      <div className="clarification-actions">
                        <button type="button" disabled={busy} onClick={cancelClarification}>取消本次澄清</button>
                      </div>
                      </>)}
                    </>
                  ) : resolvedOrCancelledOutcome(message) ? (
                    <div className="clarification-outcome">
                      {resolvedOrCancelledOutcome(message)}
                    </div>
                  ) : resumableRecord(message) ? (
                    <TaskResumeRecoveryCard
                      record={resumableRecord(message) as TaskResumeRecord}
                      scopeLabel={currentScopeLabel}
                      busy={busy}
                      resuming={resumingTaskMessageId !== null}
                      canPickScope={selectedSheetIndexes.length > 0}
                      localAnswerPending={localMappingResumeRef.current?.pending.messageId === message.id}
                      onResume={() => void resumeInterruptedTask(message)}
                      onCancel={() => cancelInterruptedTask(message)}
                    />
                  ) : (
                    <div className="clarification-outcome">
                      这张澄清卡已失效，且没有可用的恢复记录；如需继续，请重新发送需求。
                    </div>
                  )}
                </section>
              )}
              {message.planPreview && (
                pendingPlan && pendingPlanMessageId === message.id && !stagedReview
                  ? planCard
                  : executedPlan?.messageId === message.id
                    // 执行完成后留住计划摘要：用户仍要能核对"按什么口径改的"。
                    ? executedPlanCard
                    : <div className="wps-plan-outcome" role="status">
                        <strong>{message.planPreview.title}</strong>
                        <span>{message.planPreview.status === "executed"
                          ? "计划已执行并回读确认"
                          : message.planPreview.status === "cancelled"
                            ? "计划已取消，未写入工作表"
                            : "计划预览已失效；如需写入，请重新提出任务"}</span>
                      </div>
              )}
              {!message.activityLog && prePlanningPause.entries.filter((entry) => entry.messageId === message.id &&
                entry.conversationId === chatHistory.activeConversationId).map((entry) =>
                <PrePlanningPauseCard key={entry.id} busy={busy} error={entry.error}
                  onContinue={() => void restartEarlyAnalysis(entry.id)} />)}
              {message.planningPause && (() => {
                /*
                 * 规划暂停卡（CAP-02）。
                 *
                 * 三条纪律落在这张卡上：
                 *  1. 操作区只放 `<button>`，状态文字单独放在一个 role="status" 的
                 *     节点里 —— 与 WPS 结果卡同一条既有约定（状态与操作分离）。
                 *  2. 按钮只在**当前活跃暂停就是这条**时出现；已经过去的暂停卡只留
                 *     状态说明，避免用户点一条早就被取代的任务。
                 *  3. 文案全部来自结构化字段（`pauseType` / `reason` / `actionLabel`），
                 *     不看助手文本、不猜关键词。
                 */
                const active = planningPause?.sourceMessageId === message.id;
                const view = active ? planningPause!.view : null;
                /*
                 * 面板刚重开、还在这条任务的核对中。
                 *
                 * 这段刻意**排在 `active` 之前**：核对没完之前谁也不知道按钮
                 * 该不该画，画出来就是基于过期快照的假按钮。状态文字照旧给
                 * （用户知道这条任务还在），但操作区先不出现。
                 */
                const reloading =
                  planningPauseReload?.checking === true &&
                  planningPauseReload.turnId === message.planningPause!.turnId;
                const continuing = continuingTurnId === message.planningPause!.turnId;
                const reloadBlocked =
                  planningPauseReload?.checking === false &&
                  planningPauseReload.turnId === message.planningPause!.turnId
                    ? planningPauseReload.blockedReason
                    : null;
                return (
                  <PauseDetails active={reloading || continuing} resolved={message.planningPause!.resolved}
                    blocked={message.planningPause!.awaitingStatus === false && !message.planningPause!.canContinue && !message.planningPause!.resolved}>
                  <section
                    className="wps-planning-pause-card"
                    aria-label="规划已暂停"
                  >
                    <div className="wps-planning-pause-status" role="status">
                      {message.text && <p>{message.text}</p>}
                      <strong>
                        {message.planningPause!.resolved ? (message.planningPause!.terminalReason ?? "本段已接续，查看后续结果")
                          : continuing ? "正在继续"
                          : reloading ? "正在核对原任务"
                          : message.planningPause!.awaitingStatus ? "已暂停接收，待核对原任务"
                          : message.planningPause!.awaitingStatus === false && !message.planningPause!.canContinue ? "任务不能继续"
                          : WPS_PLANNING_PAUSE_LABEL[message.planningPause!.pauseType]}
                      </strong>
                      <span>{message.planningPause!.reason}</span>
                    </div>
                    {reloading || continuing ? (
                      <small className="wps-planning-pause-note">
                        {continuing ? "正在接续未完成的步骤…" : "正在核对该任务的真实进度…"}
                      </small>
                    ) : !message.planningPause!.resolved &&
                        (message.planningPause!.awaitingStatus === true || message.planningPause!.canContinue) &&
                        !(planningPauseReload?.turnId === message.planningPause!.turnId &&
                          planningPauseReload.checking === false && planningPauseReload.canRecheck === false) &&
                        !(active && view?.canContinue && view.action === "continue_checkpoint") ? (
                      <div className="wps-planning-pause-actions">
                        <button type="button" disabled={busy} onClick={() => {
                          void inspectPlanningMessage(message);
                        }}>核对任务状态</button>
                        <small>{reloadBlocked ?? "先核对原任务；只查询，不自动继续、不写入。"}</small>
                      </div>
                    ) : active && view && view.canContinue && view.action === "continue_checkpoint" ? (
                      <div className="wps-planning-pause-actions">
                        <button
                          type="button"
                          disabled={busy}
                          onClick={() => void resumePlanningFromPause(planningPause!)}
                        >
                          {view.actionLabel || "继续分析"}
                        </button>
                        <small>{view.detail}</small>
                      </div>
                    ) : (
                      <small className="wps-planning-pause-note">
                        {message.planningPause!.resolved ? message.planningPause!.detail : reloadBlocked
                          ? `该任务当前不能继续分析：${reloadBlocked}`
                          : message.planningPause!.awaitingStatus === false && !message.planningPause!.canContinue
                            ? `${message.planningPause!.detail} 如需处理请重新开始任务。`
                          : active && view && !view.canContinue
                            ? `该任务当前不能继续分析：${view.reason}`
                            : "原任务记录仍保留；请先核对状态，不要重复发送原需求。"}
                      </small>
                    )}
                  </section></PauseDetails>
                );
              })()}
              {message.verification && (
                <ExecutionResultCard
                  className="wps-verification-card"
                  tone={wpsExecutionCardTone({
                    outcome: normalizeWpsExecutionOutcome(message.verification.outcome),
                    undone: message.verification.status === "undone"
                  })}
                  title={message.verification.status === "undone"
                      ? "已撤销并恢复"
                      : WPS_EXECUTION_OUTCOME_LABEL[normalizeWpsExecutionOutcome(message.verification.outcome)]}
                  summary={<>
                    {message.verification.actionCount} 个动作 · {message.verification.cellCount}{" "}
                    {WPS_EXECUTION_OUTCOME_DETAIL_SUFFIX[normalizeWpsExecutionOutcome(message.verification.outcome)]}
                  </>}
                >
                  {message.changeReviewExecutionId && (
                    <ChangeReviewSummaryCard
                      reviews={changeReviewsRef.current.get(message.changeReviewExecutionId) ?? []}
                    />
                  )}
                  <ExecutionResultDetails open={!["verified", "undone"].includes(wpsExecutionCardTone({
                    outcome: normalizeWpsExecutionOutcome(message.verification.outcome),
                    undone: message.verification.status === "undone"
                  }))}>
                  <ul>
                    {message.verification.targets.slice(0, 5).map((target, index) => (
                      <li key={`${message.id}-target-${index}`}>
                        {planActionLabel(target.type)} · {target.sheetName}!{target.range}
                      </li>
                    ))}
                  </ul>
                  {message.verification.targets.length > 5 && (
                    <small>另有 {message.verification.targets.length - 5} 个目标，详见执行过程。</small>
                  )}
                  </ExecutionResultDetails>
                  <small>{SESSION_RECOVERY_NOTICE_PREFIX}；关闭后不能从此卡片撤销。</small>
                  {/*
                    撤销入口就挂在这张卡里，归属靠 `executionId`。
                    只有"当前会话里的活跃恢复点正是这一次、且允许撤销"才给可点按钮；
                    被新执行取代、面板重开、或已撤销之后，这里只会留一句状态说明。
                  */}
                  <div className={`wps-undo-inline execution-recovery-footer ${wpsRecoveryCardStateFor(message.verification)}`}>
                    <span>{WPS_RECOVERY_CARD_LABEL[wpsRecoveryCardStateFor(message.verification)]}</span>
                    {wpsRecoveryCardStateFor(message.verification) === "undoable" && (
                      <button
                        type="button"
                        aria-label="撤销本次执行"
                        className="secondary-button"
                        disabled={busy || undoingExecutionId === message.verification!.executionId}
                        onClick={() => void undoExecution(message.verification!.executionId)}
                      >
                        {undoingExecutionId === message.verification!.executionId ? "↶ 正在撤销…" : "↶ 撤销本次执行"}
                      </button>
                    )}
                  </div>
                  {/*
                    只有「完全通过」且尚未撤销的写入才提供保存为工具入口。
                    资格绑定实际已验证的计划版本与执行身份（verifiedPlanForSaveRef）；
                    同 planId 的旧版本验证通过也不放行已修改的版本。
                  */}
                  {message.verification.status !== "undone" &&
                    normalizeWpsExecutionOutcome(message.verification.outcome) === "readback_passed" &&
                    verifiedPlanForSaveRef.current?.executionId === message.verification.executionId && (
                    <div className="wps-save-tool-inline">
                      <button
                        type="button"
                        className="secondary-button"
                        disabled={busy}
                        onClick={() => openSaveToolDialog(message.verification!.executionId)}
                      >
                        保存为工具
                      </button>
                    </div>
                  )}
                </ExecutionResultCard>
              )}
              {!message.verification && message.changeReviewExecutionId && (
                <ChangeReviewSummaryCard
                  reviews={changeReviewsRef.current.get(message.changeReviewExecutionId) ?? []}
                />
              )}
              {message.verification && message.verification.status === "undone" && (
                <ChangeReviewUndoNote
                  summaries={changeReviewSummariesRef.current.get(message.verification.executionId) ?? []}
                />
              )}
              {message.attachmentNames?.length ? (
                <p className="message-text">附件：{message.attachmentNames.join("、")}</p>
              ) : null}
              {message.activityLog && message.activityLog.steps.length > 0 && (
                <TaskProgressTimeline
                  log={message.activityLog}
                  live={activeActivityMessageId === message.id}
                  segment={{
                    live: activeActivityMessageId === message.id,
                    continuesTask: message.replySegment?.continuesTask,
                    ...(message.replySegment?.nextSegmentId
                      ? {
                          onShowLater: () => jumpToMessage(message.replySegment!.nextSegmentId!),
                          laterLabel: "查看后续进展"
                        }
                      : {})
                  }}
                  onContinue={
                    // 已被后续段接管的段是历史：不再给"继续"入口（否则用户会对一个
                    // 已经继续过的任务再点一次）。
                    message.replySegment?.status === "superseded"
                      ? undefined
                      : prePlanningPause.entries.some((entry) => entry.messageId === message.id && entry.conversationId === chatHistory.activeConversationId)
                        ? () => void restartEarlyAnalysis(`preplanning-${message.id}`) : undefined
                  }
                  continueBusy={busy}
                  continuationError={prePlanningPause.entries.find((entry) => entry.messageId === message.id)?.error}
                  /*
                   * 卡片上的「停止」与输入框按钮读**同一份判据**：只有真的能停时才给入口。
                   * 以前写成 `activeActivityMessageId === message.id && progress`，只要还有
                   * 进度就画按钮 —— 本端已经没有可中止对象时点它等于什么都没发生。
                   */
                  onStop={activeActivityMessageId === message.id && cardPauseButton.canPause
                    ? () => pauseCurrentTurn("card")
                    : undefined}
                />
              )}
              {/*
                删表结果卡只画在**发起这次删除的那条消息**下面。
                它不是"会话里最新的一张卡"：换会话、换工作簿、刷新都不会改写它的归属。

                位置固定为「执行过程之后」：删表这一轮的执行过程与结果卡同属一条消息，
                卡在时间线之前会让人看到"结果先出、过程后补"。撤销按钮属于这张卡，
                所以它自然也排在过程之后 —— 卡下面不再重复画一遍执行过程。
              */}
              {(() => {
                const checkpoint = worksheetCheckpoint.entryForMessage(
                  message.id,
                  chatHistory.activeConversationId
                );
                if (!checkpoint) return null;
                return (
                  <WorksheetCheckpointCard
                    entry={checkpoint}
                    controller={worksheetCheckpoint}
                    onRestored={async (result, entry) => {
                      await adoptVerifiedRestore(entry.workbookIdentity, result.restoredSheets);
                    }}
                  />
                );
              })()}
              {message.text.trim() && (
                <div className="message-copy-row">
                  <MessageCopyButton
                    copied={copiedMessageId === message.id}
                    onClick={() => void copyMessage(message)}
                  />
                </div>
              )}
              </div>
            </article>
          </Fragment>
        ))}

        {/*
          这里**不再**放撤销卡。它曾经无条件挂在消息流末尾，于是永远出现在
          最新一条回复下面，看起来像"本轮唯一能做的操作" —— 而它其实属于
          上一轮已经完成的写入。恢复点只从「备份与恢复」进入（见上方 tool-drawer），
          那里会标明它属于哪一次执行。
        */}

        {wpsRecovery && (
          <article className="wps-write-recovery" role="status">
            <div className="wps-write-recovery-head">
              <strong>上次写入的恢复核对</strong>
              <span className={`wps-write-recovery-badge wps-write-recovery-${wpsRecovery.verdict.kind}`}>
                {wpsRecovery.verdict.kind === "completed" && wpsRecovery.verdict.recorded
                  ? "已记录完成" : WPS_RECONCILE_LABEL[wpsRecovery.verdict.kind]}
              </span>
            </div>
            {wpsRecovery.identity.kind !== "match" && (
              <p className="wps-write-recovery-detail">{wpsRecovery.identity.detail}</p>
            )}
            <p className="wps-write-recovery-detail">
              {wpsRecovery.checking
                ? "正在查询服务端写入现场…"
                : WPS_RECOVERY_DETAIL(wpsRecovery.verdict)}
            </p>
            <p className="wps-write-recovery-detail">
              恢复只使用当前已授权工作表；续跑通过复核后执行剩余动作，再审阅结果。
            </p>
            {/*
             * 按钮**只**来自判定给的 `wpsReconcileActions`：判定说证据不足/现场冲突
             * 就一个写入或补记按钮都不画，"继续"只在真的能安全续跑时出现。
             */}
            {wpsRecovery.verdict.kind === "replan" && <label>
              <input type="checkbox" checked={recoveryScopeConfirmed}
                onChange={(event) => setRecoveryScopeConfirmed(event.target.checked)} />
              使用当前已选工作表重新读取和规划；校验通过后执行，再审阅结果。
            </label>}
            <div className="wps-write-recovery-actions">
              {wpsReconcileActions(wpsRecovery.verdict).map((action) => (
                <button
                  key={action}
                  className="wps-recovery-action"
                  type="button"
                  disabled={busy || wpsRecovery.checking || wpsRecovery.acting || (action === "replan" && !recoveryScopeConfirmed)}
                  onClick={() => void runWpsRecoveryAction(action)}
                >
                  {WPS_RECOVERY_ACTION_LABEL[action]}
                </button>
              ))}
              {(wpsRecovery.verdict.kind === "blocked" || wpsRecovery.verdict.kind === "insufficient") &&
                <button type="button" disabled={busy || wpsRecovery.checking || wpsRecovery.acting}
                  onClick={() => void refreshWpsRecovery()}>重新核对（只读）</button>}
            </div>
          </article>
        )}

        {(progress || writeProgress) && !activeActivityMessageId && progressCard}
      </main>
      {/*
        用户往上滚离底部时出现，判据是 `useMessageAutoFollow` 的 `awayFromBottom`
        —— **只看位置**，不要求"期间又来了一次新内容"。旧口径（`hasUnseenProgress`）
        只在"离底 **且** 有新内容"时才亮，于是往上翻历史时什么都不出现。
        字形跟着**本轮是否在跑**走：跑着是三点动效，跑完换成锚。判据用 `busy`
        （进度/写入/目标核对任一在途），与 Excel 端同一个组件、同一套语义。
      */}
      {awayFromBottom && (
        <ScrollToLatestButton running={running} onClick={scrollToLatest} />
      )}
      </div>

      {historyNotice && (
        <div className="storage-write-warning" role="alert">
          <strong>对话历史没能保存到本机</strong>
          <span>{historyNotice}</span>
        </div>
      )}

      {staleCredentials && (
        <div className="stale-credentials-warning" role="alert">
          <strong>这段历史来自旧版本，曾包含凭据</strong>
          <span>
            凭据已在读取时抹掉，界面里不会再显示原文。为避免把它当作正常历史
            续聊，发送仍会被拦下；请点上方「新对话」开始一段干净的对话，或删掉
            这段历史。
          </span>
        </div>
      )}

      {/*
        数据范围面板的**独立呈现层**：它是外壳的直接子层，不是 `.composer` 的后代。
        这样它既不会被审阅条/输入框挤掉，也不受那一层的 `backdrop-filter`、
        `z-index` 与外壳裁剪牵制；打开时整格铺开，关闭时彻底不留。
      */}
      {scopeOverlay ? <div className="scope-layer">{scopeOverlay}</div> : null}

      <footer className="composer">
        <ResultReviewBar state={resultReviewState} title={resultTitle}
          onCompare={() => void resultReview.compare(resultReviewState.key ?? '')}
          onConfirm={() => void resultReview.confirm(resultReviewState.key ?? '')}
          onDiscard={() => void resultReview.discard(resultReviewState.key ?? '')}
          onRetryMarker={() => void resultReview.retryMarkerCleanup(resultReviewState.key ?? '')} />
        {targetGuard.notice && onOpenScope && !scopePickerOpen && <TargetScopeNotice notice={targetGuard.notice}
          error={targetGuard.error} busy={targetGuard.confirming} onOpenScope={onOpenScope}
          onRecheck={() => void targetGuard.recheck()}
        />}
        {scopeDock}
        {/*
          与 Excel 端同构：把手放在 `.composer-input-region` 的上内边距里（在输入框
          之外），这样它既不会盖住文本区抢走光标，又能整宽拖动。
        */}
        <div className={`composer-input-region ${draggingImage ? "dragging-image" : ""}`}
          onDragEnter={handleDragEnter} onDragOver={handleDragOver}
          onDragLeave={handleDragLeave} onDrop={(event) => { if (!busy) void handleDrop(event); else event.preventDefault(); }}>
          <div className="composer-resize-handle" {...composerResizeHandleProps}>
            <i />
          </div>
          <TaskExecutionControls reviewing={resultReviewState.phase === "reviewing"} state={composerPauseButton}
            pausing={pauseRequestState?.phase === "requesting"}
            disabled={busy && composerPauseButton.kind === "continue"}
            onPause={() => pauseCurrentTurn("composer")}
            onResume={() => void resumeComposerTask()} />
          <form
            className="composer-box"
            onSubmit={(event) => {
              event.preventDefault();
              void submitComposer();
            }}
          >
            <textarea
              ref={composerInputRef}
              aria-label="给 Excel Bro 发消息"
              value={prompt}
              disabled={busy}
              placeholder="描述你想查询、分析或修改的内容…"
              onChange={(event) => setPrompt(event.target.value)}
              onPaste={(event) => { if (!busy) handleImagePaste(event); }}
              onKeyDown={(event) => {
                if (event.key === "Enter" && !event.shiftKey) {
                  event.preventDefault();
                  void submitComposer();
                }
              }}
            />
            {pendingImages.length > 0 && <div className="composer-attachments">
              {pendingImages.map((image, index) => <figure key={image.id}>
                <img src={image.previewUrl} alt={image.name} />
                <figcaption title={image.name}>{image.name}</figcaption>
                <button type="button" disabled={busy || preparingImages} aria-label={`移除图片 ${image.name}`} onClick={() => removeImage(index)}>×</button>
              </figure>)}
            </div>}
            {preparingImages && <div className="composer-image-error" role="status">正在处理图片…</div>}
            {pendingImages.length > 0 && !selectedModel?.supportsVision &&
              <div className="composer-image-error" role="status">当前模型不支持图片，请切换到支持视觉输入的模型后发送。附件不会自动发送。</div>}
            {imageError && <div className="composer-image-error" role="alert">{imageError}</div>}
            <div className="composer-toolbar">
              <input ref={imageInputRef} className="image-file-input" type="file"
                aria-label="选择图片附件" accept="image/png,image/jpeg,image/webp" multiple
                onChange={(event) => {
                  const files = Array.from(event.target.files ?? []);
                  event.target.value = "";
                  if (!busy) void addImageFiles(files);
                }} />
              <div className="composer-tools-left">
                <button type="button" className="attach-image-button" aria-label="添加图片"
                  title="添加图片，也可以粘贴或拖入截图" disabled={busy || preparingImages}
                  onClick={() => imageInputRef.current?.click()}>
                  <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <rect x="3.5" y="4.5" width="17" height="15" rx="2.5" />
                    <circle cx="9" cy="9.5" r="1.5" />
                    <path d="m5.5 17 4.2-4.4 3.1 3 2.2-2.2 3.5 3.6" />
                  </svg>
                </button>
              </div>
              <span className="composer-shortcut-hint">Enter 发送 · Shift + Enter 换行 · 可粘贴截图</span>
              <button className="send-button composer-send-button" type="submit"
                aria-label="发送" title="发送这条需求"
                disabled={busy || preparingImages ||
                  (!pendingImages.length && !prompt.trim()) ||
                  (pendingImages.length > 0 && !selectedModel?.supportsVision) || selectedSheetIndexes.length === 0}>↑</button>
            </div>
            {/*
              暂停请求的阶段**如实**显示，挂在工具条**外面**（工具条是三列网格，
              多一个子元素会被挤到第二行）：本端已停止接收 ≠ 服务端已暂停，
              `acknowledged` 只说明请求被接受，仍要看任务状态。
            */}
            {pauseRequestState && (
              <span className="composer-pause-note" role="status" data-phase={pauseRequestState.phase}>
                {pauseRequestState.detail}
              </span>
            )}
          </form>
        </div>
        {/*
          「开发验收模式」是**诊断**文案：它只对开发者有意义，普通用户看到只会
          困惑"我是不是不该用"。所以它只在开发构建里出现（`import.meta.env.DEV`，
          与探测面板共用同一个开关）。
          「写入为何还没开放」则是用户该知道的真实状态，任何构建下都要说。
        */}
        {(developmentWrite || !canWrite) && (
          <p className="wps-capability-note">
            {developmentWrite && import.meta.env.DEV
              ? "开发验收模式：仅在独立测试工作簿中确认写入与撤销。"
              : `当前仍可查询；写入将在满足条件后开放：${writeStatusReason}`}
          </p>
        )}
        {(import.meta.env.DEV || import.meta.env.VITE_ENABLE_RECOVERY_PROBE === "true") && <RecoveryProbeControl identity={(checkpointWorkbookPath ? wpsWorkbookIdentity(checkpointWorkbookPath) : null)} disabled={busy} />}
        {/*
          A1-E1 验收入口：让本工作簿的下一次规划任务的第一条 Python 结果引用在
          "已登记、尚未首次写入"处失效。可见性与文案都来自共享层（Excel 端读同一份），
          本端只负责把 WPS 的宿主名与工作簿身份填进去；失效本身发生在服务端解析引用
          的那一刻，界面无法伪造"已过期"。
        */}
        <ResultRefFaultControl
          visible={resultRefFaultVisible(
            import.meta.env.DEV || import.meta.env.VITE_ENABLE_RECOVERY_PROBE === "true"
          )}
          disabled={busy}
          target={resultRefFaultTargetOfPanel}
          load={getResultRefFaultStatus}
          arm={armResultRefFault}
          cancel={cancelResultRefFault}
        />
      </footer>

      {newConversationConfirmOpen && (
        <div className="tool-dialog-backdrop" onMouseDown={() => setNewConversationConfirmOpen(false)}>
          <section
            className="tool-dialog history-delete-dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="new-conversation-confirm-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="tool-dialog-title-row">
              <strong id="new-conversation-confirm-title">放弃待确认的写入计划？</strong>
              <button type="button" aria-label="关闭" onClick={() => setNewConversationConfirmOpen(false)}>×</button>
            </div>
            <p>当前对话还有一份尚未执行的写入计划。新建对话会放弃这份计划，但不会修改工作簿。</p>
            <div className="tool-dialog-actions">
              <button type="button" onClick={() => setNewConversationConfirmOpen(false)}>继续当前对话</button>
              <button className="danger-button" type="button" onClick={completeNewConversation}>放弃计划并新建</button>
            </div>
          </section>
        </div>
      )}

      {saveToolCandidate && (
        <SaveToolDialog
          toolName={toolName}
          toolDescription={toolDescription}
          eligibility={saveEligibility}
          approveFixedContent={approveFixedContent}
          approveDestructive={approveDestructive}
          onToolNameChange={setToolName}
          onToolDescriptionChange={setToolDescription}
          onApproveFixedContentChange={setApproveFixedContent}
          onApproveDestructiveChange={setApproveDestructive}
          onCancel={cancelSaveTool}
          onConfirm={confirmSaveToolForWps}
          blockedReason={saveToolInvalidationReason()}
          subtitle="验证后可从「我的工具」重复使用"
          note="只保存 WPS 白名单操作，不会保存 API Key，也不会执行系统命令。"
        />
      )}

      {petVisible && <PetCompanion busy={busy} />}
      <WpsModelSettingsPanel
        open={modelSettingsOpen}
        models={models}
        currentModelId={effectiveModelId}
        management={modelManagement}
        onClose={closeModelSettings}
        onAddConnection={() => void openModelSettings(true)}
      />
      <ThemePanel open={themeOpen} onClose={() => setThemeOpen(false)} {...theme} />
      <RuleManager isOpen={ruleManagerOpen} onClose={() => setRuleManagerOpen(false)} />
    </section>
  );
}
