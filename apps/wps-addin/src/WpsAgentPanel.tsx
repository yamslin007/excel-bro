import { Fragment, type ReactNode, useCallback, useEffect, useMemo, useRef, useState } from "react";
import { flushSync } from "react-dom";
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
  withTaskResumeConfirmation,
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
import {
  ApiRequestError,
  apiErrorCategory,
  listModels,
  validatePlan,
  type ModelCatalogStatus,
  type ModelOption
} from "../../excel-addin/src/api";
import { assertEditablePlan, type ClarificationField, type IntentOption, type ResultContext } from "../../excel-addin/src/contracts";
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
  instantiateTool,
  instantiateQueryTool,
  loadQueryTools,
  loadTools,
  type SavedQueryTool,
  type SavedTool
} from "../../excel-addin/src/storage";
import { executeQueryTableData } from "../../excel-addin/src/dataTools";
import TaskProgressTimeline from "../../excel-addin/src/components/TaskProgressTimeline";
import MessageCopyButton from "../../excel-addin/src/components/MessageCopyButton";
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
import type { ActivityLog, ActivityOutcome } from "../../excel-addin/src/hooks/useActivityProgress";
import { useMessageAutoFollow } from "../../excel-addin/src/hooks/useMessageAutoFollow";
import { useComposerResize } from "../../excel-addin/src/hooks/useComposerResize";
import { useAuthorizedTarget } from "../../excel-addin/src/hooks/useAuthorizedTarget";
import { AuthorizedTargetConfirmation } from "../../excel-addin/src/components/AuthorizedTargetConfirmation";
import { inspectWpsAuthorizedTarget, isWpsReviewCopyTarget, wpsAnalysisTargetEvidence } from "./wpsAuthorizedTarget";
import { checkAuthorizedTarget, TargetConfirmationRequiredError } from "@excel-bro/core/authorizedTarget";
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
import { buildPlanChangePreview } from "./wpsPlanChangePreview";
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
  verifyWpsNativeComparison,
  WpsFormulaTrialError,
  type WpsNativeComparison
} from "./wpsNativeComparison";
import { fixedValueFallbackFromFormula } from "./wpsColumnMappingPlan";
import {
  advanceWpsActivity,
  executionActivityUpdate,
  finishWpsActivity,
  readonlyActivityUpdate,
  startWpsActivity,
  type WpsActivityUpdate
} from "./wpsActivityLog";
import {
  resumeWpsLocalMapping,
  runWpsReadonlyTurn,
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
import type { WpsPlanningTaskStatus } from "./wpsPlanningStream";
import type { PlanningPauseResult } from "../../excel-addin/src/contracts";
import { planningPausedEntriesFromHistory } from "@excel-bro/core";

type ChatMessage = {
  id: string;
  role: "user" | "assistant" | "system";
  text: string;
  attachmentNames?: string[];
  hasImageAttachments?: boolean;
  resultContext?: ResultContext | null;
  activityLog?: ActivityLog;
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
    pauseType: "budget" | "model_error" | "interrupted";
    reason: string;
    canContinue: boolean;
    action: "continue_checkpoint" | "regenerate" | "none";
    actionLabel: string;
    detail: string;
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
  onAuthorizeTarget?: (indexes: number[]) => void;
  scopeDock: ReactNode;
};

const WPS_CHAT_STORAGE_KEY = "excel-bro.wps.chat.v1";
const LEGACY_MODEL_CATALOG_NOTICE =
  "暂时无法读取模型配置，请点击上方模型菜单里的「重试」。已输入的内容仍保留在输入框中。";

/**
 * 规划暂停的成因文案（CAP-02）。
 *
 * 只按服务端给的 `pauseType` 取值 —— 这是结构化判据，不是从助手文本里读出来的。
 * 三种成因的**处置方式不同**，所以文案必须分开：
 *  - `budget`：本轮计算量达到上限，可继续或换模型；
 *  - `model_error`：模型连接故障，通常需要换一个可用连接；
 *  - `interrupted`：被打断（如刷新/断线），现场仍在。
 */
const WPS_PLANNING_PAUSE_LABEL: Record<
  "budget" | "model_error" | "interrupted",
  string
> = {
  budget: "本轮处理时间已达上限，进度已保存",
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
  // 同一道题只出现**最后一次**的答案（替换发生在 `withTaskResumeConfirmation`）。
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

export default function WpsAgentPanel({
  selectedSheetIndexes,
  resolveSelectedSheetIndexes,
  scopeMode = "manual",
  onAuthorizeTarget,
  scopeDock
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
  // 原生左右子窗口只活在本次面板会话；副本矩阵绝不进入聊天历史或持久存储。
  const [stagedReview, setStagedReview] = useState<WpsNativeComparison | null>(null);
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
  const resumingPauseRef = useRef(false);
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
  /**
   * 当前用户任务的**稳定身份**。
   *
   * 每次新的用户请求（`showUser !== false`）换一个 —— 一个 taskId 对应一件事；
   * 同一任务内的重新规划、重新授权、重新生成预览都沿用它，因此"换了 turnId
   * 仍是同一件任务"在界面上是可判定的，恢复卡也据此绑定到原消息。
   */
  const taskIdRef = useRef<string>(newTaskId());
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
    intentConfirmed?: boolean; clarificationRound?: number; showUser?: boolean; targetConfirmedRetry?: boolean; images?: PendingImage[];
  }) => void>(() => {});
  const resumeTaskRef = useRef<(message: ChatMessage) => void>(() => {});
  const targetGuard = useAuthorizedTarget({
    enabled: true,
    reviewActive: stagedReview !== null,
    monitorPaused: progress !== null || writeProgress !== null,
    isReviewCopy: (target) => isWpsReviewCopyTarget(target, stagedReview),
    inspect: () => inspectWpsAuthorizedTarget(scopeMode === "auto" ? null : selectedSheetIndexes),
    analysisEvidence: wpsAnalysisTargetEvidence,
    adopt: (target) => {
      onAuthorizeTarget?.(target.selected.map(Number));
      setAuthorizationRevision((value) => value + 1);
      if (targetGuard.binding && checkAuthorizedTarget(targetGuard.binding, target) === "same") return;
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
  const [tools] = useState(loadTools);
  const [queryTools] = useState(loadQueryTools);
  const [copiedMessageId, setCopiedMessageId] = useState<string | null>(null);
  const [copyError, setCopyError] = useState(false);
  const [newConversationNotice, setNewConversationNotice] = useState<{
    id: number;
    previousSaved: boolean;
  } | null>(null);
  const [newConversationConfirmOpen, setNewConversationConfirmOpen] = useState(false);
  const controllerRef = useRef<AbortController | null>(null);
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
    hasUnseenProgress,
    handleScroll: handleMessageStreamScroll,
    scrollToLatest
  } = useMessageAutoFollow(scrollSignal);
  const theme = useTheme();
  const logoLongPress = useLongPress(() => setThemeOpen(true));
  const busy = progress !== null || writeProgress !== null || targetGuard.checking;

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
      if (!active) return;
      const currentIdentity = (() => {
        try {
          return wpsWorkbookIdentity(getWpsApplication().ActiveWorkbook?.FullName);
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
        localInspection
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
  }, []);

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
      messages,
      (turnId) => {
        const message = messages.find(
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
          setPlanningPauseReload({
            turnId: latest.turnId,
            checking: false,
            blockedReason: null
          });
          return;
        }
        setPlanningPause(null);
        setPlanningPauseReload({
          turnId: latest.turnId,
          checking: false,
          /*
           * 不放行时如实转述服务端的拒绝理由；放行但只能重新规划时**不编理由** ——
           * 卡片会落到"已被新的任务取代"之外的中性说明，不冒充服务端的结论。
           */
          blockedReason: inspected.view.canContinue ? null : inspected.view.reason
        });
      } catch {
        // 查不到（服务重启 / 这条任务的服务端记录已不在）时保持"核对中"的
        // 保守结论：既不说"能继续"，也不说"任务没了" —— 没有证据的事不下结论。
        if (!active) return;
        setPlanningPause(null);
        setPlanningPauseReload({
          turnId: latest.turnId,
          checking: false,
          blockedReason: "暂时读不到这条任务的真实进度，请确认本地 AI 服务仍在运行后重试。"
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
        return wpsWorkbookIdentity(getWpsApplication().ActiveWorkbook?.FullName);
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
        localInspection
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
   * `replan` 走的是**重新读取 + 重新授权 + 重新确认**的完整一轮，生成的是新预览、
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
    if (action !== "dismiss" && !(await targetGuard.check())) return;
    const operationId = current.operationId;
    let revision = current.state.revision;
    recoveryActionLock.current = true;
    setWpsRecovery({ ...current, acting: true });
    try {
      if (action === "dismiss") {
        clearPendingWpsExecution(operationId);
        setWpsRecovery(null);
        return;
      }
      // A displayed badge is not authority. Requery AND reread on each click.
      const fresh = await getWpsExecutionStatus(operationId);
      const identity = wpsIdentityVerdict({
        currentIdentity: wpsWorkbookIdentity(getWpsApplication().ActiveWorkbook?.FullName),
        recordedIdentity: fresh.workbookIdentity
      });
      const verdict = wpsReconcileVerdict({ state: fresh, identityMatch: identity.kind === "match",
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
          code: "WPS_RECOVERY_PREVIEW_REQUIRED", message: "核对完成，剩余动作需重新预览。"
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
       * 重新预览：这次执行还有动作没写完，但目标矩阵与确认快照只活在会话内内存，
       * 重载后拼不回用户当初确认的那份计划 —— **精确续跑剩余动作未实现**。
       * 所以如实走"重新读取并生成预览"，让用户重新授权、重新确认，生成新的执行身份。
       */
      const taskId = current.taskId;
      const record = messages
        .map((message) => (message.taskResume ? sanitizeTaskResumeRecord(message.taskResume) : null))
        .find((item) => item?.taskId === taskId);
      append({
        role: "system",
        text:
          "这次写入还有未完成的动作，而目标矩阵与确认快照只存在会话内，重开后无法精确续跑剩余动作。" +
          "下面按原需求重新读取并生成预览 —— 那是新的一轮，需要你重新授权并确认后再写入，" +
          "不是从上次进度接着写。"
      });
      if (!record) {
        append({
          role: "system",
          text: "没有找到这次任务的原始需求记录，请重新输入需求，确认预览后再写入。"
        });
        return;
      }
      // Fence the previous owner before generating a NEW preview. This does not
      // resume its plan, and no cells are written by this recovery handler.
      const claimed = await claimWpsExecutionResume(operationId, revision);
      if (!claimed.ownerToken) throw new Error("未取得旧执行的处置权，不能重新规划。");
      await suspendWpsExecution(operationId, {
        expectedRevision: claimed.revision, ownerToken: claimed.ownerToken,
        code: "WPS_REPLAN_REQUESTED", message: "用户选择重新读取并生成预览。"
      });
      // 沿用原 taskId 与已确认信息：重新规划仍是同一件事，用户不必重答。
      resumeCarryRef.current = {
        taskId: record.taskId,
        confirmations: record.confirmations
      };
      await runTurn(resumePrompt(record), { showUser: false });
      // Keep the durable discovery index until a new preview is confirmed or the
      // user explicitly dismisses it; a failed planner must not erase evidence.
      refreshWpsRecoveryVerdict(operationId, current.taskId, await getWpsExecutionStatus(operationId));
    } catch (error) {
      append({
        role: "system",
        text: `恢复动作没有成功，本次未对原表做新的写入。${turnErrorMessage(error)}`
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
      try {
        discardWpsNativeComparison(stagedReview);
        setStagedReview(null);
      } catch {
        // Keep the pending plan and the owned copy visible so the user can resolve it.
        return;
      }
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
    updateMessages((current) => current.map((message) => message.id === activity.id && message.activityLog
      ? { ...message, activityLog: advanceWpsActivity(message.activityLog, step, activity.startedAt) }
      : message));
  };

  const endActivity = (activity: { id: string; startedAt: number }, outcome: ActivityOutcome) => {
    updateMessages((current) => current.map((message) => message.id === activity.id && message.activityLog
      ? { ...message, activityLog: finishWpsActivity(message.activityLog, outcome, activity.startedAt) }
      : message));
    setActiveActivityMessageId((current) => current === activity.id ? null : current);
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
      try {
        discardWpsNativeComparison(stagedReview);
        setStagedReview(null);
      } catch {
        append({ role: "system", text: "审阅副本未能自动关闭。请先处理副本，再新建对话；原表尚未写入。" });
        return;
      }
    }
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
    options: { intentConfirmed?: boolean; clarificationRound?: number; showUser?: boolean; targetConfirmedRetry?: boolean; images?: PendingImage[] } = {}
  ) => {
    const sentImages = options.images ?? (options.showUser === false && !pendingImages.length
      ? taskImagesRef.current.get(taskIdRef.current) ?? [] : pendingImages);
    const text = question.trim() || (sentImages.length ? "请结合附件图片分析当前工作簿，并说明发现的问题。" : "");
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
    // A new task cannot inherit a weak-wrapper consent from a different task.
    // The queued retry is the SAME request after explicit target consent, not an old write.
    if (options.showUser !== false && !options.targetConfirmedRetry) targetGuard.clearAnalysisConsent();
    if (!(await targetGuard.checkForAnalysis(() => retryTurnRef.current(text, {
      ...options, images: sentImages, targetConfirmedRetry: true
    })))) return;
    const turnSheetIndexes = [
      ...(resolveSelectedSheetIndexes?.() ?? selectedSheetIndexes)
    ];
    /**
     * 这一轮的授权身份与表头指纹。
     *
     * 在**发起请求之前**取：`sheetIdentity` 是建卡时的身份，`headerFingerprint`
     * 是这一轮结构扫描实际读到的表头（响应回来后再取，所以下面才赋值）。
     * 两者都写进澄清卡，用户作答时逐项比对。
     */
    const sheetIdentity = currentWpsSheetIdentity();
    let headerFingerprint: string | undefined;
    if (!text || busy || turnSheetIndexes.length === 0) return;
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
    if (options.showUser !== false) {
      localMappingResumeRef.current = null;
      // 新的一次用户请求 = 新的一件任务。同一任务内的重新规划不会走到这里
      // （续答、重新授权都带 `showUser: false`），所以 taskId 在任务内保持稳定。
      taskIdRef.current = newTaskId();
      /*
       * 上一轮的服务端任务身份作废。
       *
       * 留着它会让本轮断线后去查**上一件任务**的状态 —— 查到了、还"可以继续"，
       * 于是把旧任务的结果接回本轮，用户看到的是一份不属于这次需求的东西。
       * 身份必须与任务同生命周期：新任务 = 身份清空，等本轮的意图阶段重新上报。
       */
      turnIdRefForStream.current = null;
      pauseCarryRef.current = null;
      resumeRequestRef.current = null;
      taskImagesRef.current.set(taskIdRef.current, sentImages);
      // 明确放弃"旧记录"拦截：用户真的开始输入新内容了，这段历史翻篇，
      // 后续是新写的、不再携带原始凭据。**只有这里**会清掉陈旧标记 ——
      // 系统提示之类的追加写入不能清，否则发送闸被拦一次就失效。
      updateMessages(
        (current) => [
          ...current,
          { id: messageId(), role: "user", text, attachmentNames: sentImages.map((image) => image.name),
            hasImageAttachments: sentImages.length > 0, createdAt: new Date().toISOString() }
        ],
        true
      );
    }
    const activity = beginActivity({
      id: "read-schemas",
      label: "读取已授权表结构",
      phase: "preflight",
      detail: "只读取字段结构，不向意图模型发送原始数据行。"
    });
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
    const controller = new AbortController();
    controllerRef.current?.abort();
    controllerRef.current = controller;
    setProgress({ stage: "reading", message: "正在读取表结构并理解消息…" });
    lastTurnFailedRef.current = false;
    try {
      const result = await runWpsReadonlyTurn({
        application: getWpsApplication(),
        selectedSheetIndexes: turnSheetIndexes,
        prompt: text,
        images: sentImages,
        modelId: effectiveModelId || null,
        intentConfirmed: options.intentConfirmed,
        clarificationRound: options.clarificationRound,
        conversation: messages
          .filter((message) => message.role !== "system" && message.text.trim().length > 0)
          .slice(-6)
          .map((message) => ({
            role: message.role as "user" | "assistant",
            text: message.text.slice(0, 1000)
          })),
        signal: controller.signal,
        beforeRead: async () => {
          if (!(await targetGuard.checkForAnalysis())) throw new TargetConfirmationRequiredError();
        },
        onProgress: (update) => {
          // 意图阶段一旦拿到服务端 taskId 就记下：断线恢复要靠它查真实状态。
          if (update.turnId) turnIdRefForStream.current = update.turnId;
          setProgress(update);
          updateActivity(activity, readonlyActivityUpdate(update));
        }
      });
      headerFingerprint = await applyTurnResult(result, {
        activity,
        text,
        turnSheetIndexes,
        clarificationRound: options.clarificationRound ?? 0,
        images: sentImages
      });
    } catch (error) {
      if (sentImages.length) { setPendingImages(sentImages); setPrompt(text); }
      if (error instanceof TargetConfirmationRequiredError) {
        endActivity(activity, "awaiting_input");
        setPrompt(text);
      } else if (error instanceof DOMException && error.name === "AbortError") {
        endActivity(activity, "cancelled");
        append({ role: "system", text: "任务已停止；未确认的计划不会修改工作表。" });
      } else {
        endActivity(activity, "failed");
        lastTurnFailedRef.current = true;
        if (apiErrorCategory(error) === "local_service" && options.showUser !== false) setPrompt(text);
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
      if (controllerRef.current === controller) controllerRef.current = null;
      setProgress(null);
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
  const reconcileInterruptedTurn = async (
    turnId: string,
    text: string,
    turnSheetIndexes: readonly number[]
  ) => {
    setProgress({ stage: "answering", message: "连接中断，正在核对该任务的真实进度…" });
    let status: WpsPlanningTaskStatus;
    try {
      status = await inspectWpsPlanningTask(turnId);
    } catch (error) {
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
      if (!notFound) setPrompt(text);
      return;
    }
    const view = wpsPlanningResumeView(status);
    if (status.status === "completed" && status.result) {
      // 任务其实跑完了，只是回程断线：把结果接回同一条处理链，不重发。
      const activity = beginActivity({
        id: "planning-recover",
        label: "取回已完成的规划结果",
        phase: "preflight",
        detail: "任务在服务端已完成；直接取回结果，不重新规划、不写单元格。"
      });
      try {
        await applyTurnResult(status.result as WpsTurnResult, {
          activity,
          text,
          turnSheetIndexes: [...turnSheetIndexes],
          clarificationRound: 0,
          images: []
        });
      } catch (applyError) {
        endActivity(activity, applyError instanceof TargetConfirmationRequiredError ? "awaiting_input" : "failed");
        append({ role: "system", text: turnErrorMessage(applyError) });
      }
      return;
    }
    if (status.status === "running") {
      append({
        role: "system",
        text: "连接中断，但该任务仍在服务端处理中。已保留它的身份，**请不要重复发送同一需求**；稍后可在本卡片上继续或查看结果。"
      });
      return;
    }
    if (view.canContinue && view.action === "continue_checkpoint") {
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
        clarificationRound: 0
      };
      setPlanningPause({
        turnId,
        payload: planningPauseFromStatus(status, null),
        view,
        originalPrompt: text,
        sourceMessageId: append({
          role: "assistant",
          text: "",
          planningPause: {
            turnId,
            /*
             * 落盘的暂停卡字段必须**非空**：`ChatMessage.planningPause` 是
             * 结构化记录，界面据此渲染状态文案。服务端没给 `pauseType` 时
             * 归为 `interrupted`（执行器中断，现场完整）——这是"有任务但
             * 成因未知"里最保守的一档，不等于"预算用完"。
             */
            pauseType: status.pauseType ?? "interrupted",
            reason: status.reason ?? "本轮处理暂停，当前进度已保存。",
            canContinue: status.canContinue === true,
            action: view.action,
            actionLabel: view.actionLabel,
            detail: view.detail
          }
        })
      });
      append({
        role: "system",
        text: "连接中断，但该任务的规划现场已安全保存在服务端。已为你保留它，**不必重复发送需求**。"
      });
      return;
    }
    append({
      role: "system",
      text: `连接中断。该任务当前不能精确续跑：${view.reason} 请不要重复发送同一需求，以免重复规划。`
    });
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
          : await validatePlan(
              result.plan,
              selectedNames,
              metadata.sheets.map((sheet) => sheet.name)
            );
        if (result.computedLocally && !result.sourceSnapshots?.length) {
          throw new Error("本地映射计划缺少源区域快照，不能安全预览写入。");
        }
        if (!(await targetGuard.check())) throw new TargetConfirmationRequiredError();
        const prepared = prepareWpsPlan(application, turnSheetIndexes, normalizedPlan, result.sourceSnapshots);
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
        endActivity(activity, "awaiting_input");
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
              detail: pauseView.detail
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
          resultContext: result.resultContext
        });
      }
    return headerFingerprint;
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
    if (busy) return;
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

    const controller = new AbortController();
    controllerRef.current = controller;
    try {
      setProgress({ stage: "answering", message: "正在核对服务端的任务进度…" });
      // 第 1 步：查真实状态。界面上的快照只用来定位任务，不用来当续跑凭据。
      let status: WpsPlanningTaskStatus;
      try {
        status = await inspectWpsPlanningTask(pause.turnId, controller.signal);
      } catch (error) {
        /*
         * **断线不等于可以重发需求。** 查不到状态只说明"现在读不到"，不说明
         * 服务端没有这条任务 —— 恰恰相反，我们手上有一个 turnId，说明它很可能
         * 还在（或已跑完）。这里保持暂停卡不动，让用户稍后重试；
         * **绝不**退回一次重新规划（那会重复调用模型与工具，且把"接着跑"
         * 谎报成"重新跑了一遍"）。
         */
        setPlanningPause((current) => (current ? { ...current, view: { ...current.view, detail: "暂时读不到服务端进度，尚未做任何继续操作。" } } : current));
        setProgress(null);
        append({
          role: "system",
          text: "无法读取该任务的真实进度，未做任何继续操作。请确认本地 AI 服务仍在运行后重试。"
        });
        return;
      }
      // 第 2 步：按钮只由结构化状态推导，不看助手文本、不凭界面上的旧快照。
      const view = wpsPlanningResumeView(status);
      if (!view.canContinue || view.action !== "continue_checkpoint") {
        setPlanningPause((current) => (current ? { ...current, view } : current));
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
      if (!(await targetGuard.check())) {
        append({
          role: "system",
          text: "数据范围在继续前发生了变化，已停止本次继续；原表未写入。请重新确认数据范围后再继续。"
        });
        return;
      }
      setProgress({ stage: "answering", message: "正在从上次保存的进度继续…" });
      // 第 3 步：按检查点继续。`resumeWpsPlanning` 内部再查一次状态并取
      // revision/游标；这里不再自己拼请求体，避免"界面拿一份旧 revision 去 CAS"。
      const continued = await resumeWpsPlanning({
        turnId: pause.turnId,
        requestId,
        modelId: modelManagement.selectedModelId ?? undefined,
        signal: controller.signal,
        onProgress: (event) => {
          setProgress({
            stage: "answering",
            stepId: event.stepId || "planning-continue",
            title: event.title,
            message: event.title || "正在继续规划"
          });
        }
      });
      /*
       * 续跑可能被服务端**当场拒绝**（`outcome` 为 null）：已完成 / 已停止 /
       * 没有可验证现场。这时**不能**假装续跑过，必须把真实状态落到卡上。
       */
      if (!continued.outcome) {
        const refusedView = wpsPlanningResumeView(continued.status);
        setPlanningPause((current) => (current ? { ...current, view: refusedView } : current));
        append({ role: "system", text: `未能继续该任务：${refusedView.reason}` });
        return;
      }
      // 继续成功后这条暂停卡退役：任务已经重新动起来了。
      setPlanningPause(null);
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
          reason: outcome.reason
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
              detail: nextView.detail
            }
          })
        });
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
      const activity = beginActivity({
        id: "planning-continue",
        label: "继续已暂停的规划",
        phase: "preflight",
        detail: "只从服务端保存的进度继续规划；不写单元格，写入仍须预览与确认。"
      });
      try {
        await applyTurnResult(outcome as WpsTurnResult, {
          activity,
          text: carry?.originalPrompt ?? pause.originalPrompt,
          turnSheetIndexes: carryIndexes,
          clarificationRound: carry?.clarificationRound ?? 0,
          images: []
        });
      } catch (error) {
        endActivity(activity, error instanceof TargetConfirmationRequiredError ? "awaiting_input" : "failed");
        throw error;
      }
    } catch (error) {
      if (error instanceof TargetConfirmationRequiredError) {
        append({
          role: "system",
          text: "数据范围在继续过程中发生了变化，已暂停本次继续；原表未写入。请重新确认数据范围。"
        });
      } else if (error instanceof DOMException && error.name === "AbortError") {
        append({ role: "system", text: "已停止继续；未确认的计划不会修改工作表。" });
      } else {
        append({ role: "system", text: turnErrorMessage(error) });
      }
    } finally {
      if (controllerRef.current === controller) controllerRef.current = null;
      resumingPauseRef.current = false;
      // 续跑上下文只在"这一次继续"期间有意义：无论成败都清掉，
      // 免得下一张卡带着上一件的原需求/授权范围去回接结果。
      pauseCarryRef.current = null;
      setProgress(null);
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
        const fixedPrepared = prepareWpsPlan(application, currentSheetIndexes, fallback, pendingPlan.sourceSnapshots);
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
    try {
      discardWpsNativeComparison(stagedReview);
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
      append({ role: "assistant", text: "已放弃右侧审阅副本；左侧原表未写入。" });
    } catch (error) {
      append({ role: "system", text: turnErrorMessage(error) });
    }
  };

  const confirmPendingPlan = async () => {
    if (!(await targetGuard.check())) return;
    if (!pendingPlan || !stagedReview || busy || planEditDraft !== null) return;
    const plan = pendingPlan;
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
    let wpsLedgerState: Awaited<ReturnType<typeof registerWpsWriteIntent>> | null = null;
    try {
      verifyWpsNativeComparison(getWpsApplication(), stagedReview);
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
        operationId);
      let comparison: "matched" | "different" | "unavailable" = "unavailable";
      let comparisonFailed = false;
      try {
        comparison = compareCommittedWpsResultToCopy(stagedReview);
      } catch {
        comparisonFailed = true;
      }
      updateActivity(activity, {
        id: "verified-result",
        label: comparison === "matched" ? "写入与回读核对完成" : "写入完成，计算结果仍需核对",
        phase: "verifying",
        note: comparison === "matched"
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
      }
      if (comparisonFailed) {
        append({ role: "system", text: "原表已按计划写入，但无法完成与审阅副本的最终逐格对照；请检查原表。" });
      } else if (comparison === "different") {
        append({ role: "system", text: "原表已按计划写入并通过动作回读，但计算结果与右侧副本不同。请检查原表；如不符合预期，可使用本次恢复点撤销。" });
      } else if (comparison === "unavailable") {
        append({ role: "system", text: "原表已按计划写入，但宿主未提供完整显示值，无法证明计算结果与审阅副本完全相同。" });
      }
      try {
        discardWpsNativeComparison(stagedReview);
      } catch {
        append({ role: "system", text: "原表已写入并回读核对，但审阅副本未能自动关闭；请手动关闭副本且不要保存。" });
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
      markPlanPreview("executed");
      setPendingPlan(null);
      setPendingPlanMessageId(null);
      /*
       * 只有**动作回读一致 且 与副本计算值一致**才算完全通过；其余如实记成
       * "未获证实"。但不受影响的是：这次执行留下了恢复点，所以**卡一定在**，
       * 撤销入口就挂在这张卡里（归属靠 `executionId`）。
       */
      const outcome: WpsExecutionOutcome =
        comparison === "matched" && !comparisonFailed ? "readback_passed" : "unverified";
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
      setWriteProgress(null);
    }
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
      const updated = prepareWpsPlan(application, currentSheetIndexes, validated, pendingPlan.sourceSnapshots);
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
  const undoExecution = async (executionId: string) => {
    if (busy) return;
    const record = undoRecord && undoRecord.executionId === executionId ? undoRecord : null;
    if (!record) return;
    if (!(await targetGuard.check())) return;
    const activity = beginActivity({
      id: "undo-preflight",
      label: "核对撤销条件",
      phase: "preflight",
      detail: "先确认目标仍是上次写入后的状态，避免覆盖后续人工修改。"
    });
    setWriteProgress("正在核对撤销条件…");
    try {
      await undoWpsExecution(getWpsApplication(), record.executionId, (update) => {
        setWriteProgress(update.message);
        updateActivity(activity, executionActivityUpdate(update));
      });
      updateActivity(activity, {
        id: "undo-verified",
        label: "恢复结果已回读核对",
        phase: "verifying",
        detail: "目标区域与执行前快照一致。"
      });
      endActivity(activity, "completed");
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
      updateMessages((current) => current.map((message) =>
        message.verification?.executionId === record.executionId
          ? { ...message, verification: { ...message.verification, status: "undone" } }
          : message
      ));
      append({ role: "assistant", text: `已撤销「${record.planTitle}」，并回读确认恢复成功。` });
    } catch (error) {
      endActivity(activity, "failed");
      try {
        setUndoRecord(latestWpsUndo(getWpsApplication()));
      } catch {
        setUndoRecord(null);
      }
      append({ role: "system", text: error instanceof Error ? error.message : "WPS 撤销失败。" });
    } finally {
      setWriteProgress(null);
    }
  };

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
    const answeredRecord = settleTaskAnswer(clarification.messageId, {
      label: option.label, value: option.resolution, questionId: clarification.questionId
    });
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
        intentConfirmed: true, clarificationRound: clarification.round + 1, showUser: false
      });
    } finally { resumeCarryRef.current = null; }
    if (lastTurnFailedRef.current) markTaskInterrupted(targetMessageId);
  };

  const resolveClarificationWithText = async (submittedAnswer = clarificationDraft) => {
    if (!clarification || !submittedAnswer.trim() || busy) return;
    if (!(await targetGuard.check())) return;
    const answer = submittedAnswer.trim();
    const answeredRecord = settleTaskAnswer(clarification.messageId, {
      label: "补充说明", value: answer, questionId: clarification.questionId
    });
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
    confirmation: TaskResumeConfirmation
  ) => {
    const previous = messages.find((item) => item.id === messageIdValue)?.taskResume;
    const record = previous ? sanitizeTaskResumeRecord(previous) : null;
    const answered = record ? withTaskResumeState(
      withTaskResumeConfirmation(record, confirmation), { phase: "answer_submitted" }
    ) : null;
    updateMessages((current) => current.map((item) => {
      if (item.id !== messageIdValue || !item.taskResume) return item;
      const record = sanitizeTaskResumeRecord(item.taskResume);
      if (!record) return item;
      return {
        ...item,
        taskResume: withTaskResumeState(
          withTaskResumeConfirmation(record, confirmation),
          { phase: "answer_submitted" }
        )
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

  const planCard = pendingPlan && (
    <section className="wps-plan-card" aria-label="WPS 写入计划预览">
      <span className="clarification-eyebrow">写入前确认</span>
      <strong>{pendingPlan.plan.title}</strong>
      <p>{pendingPlan.plan.summary}</p>
      {planRevised && (
        <p className="wps-plan-revised">已按你的修改更新；写入前请重新核对下面的目标与内容。</p>
      )}
      {developmentWrite && <p>当前 WPS JSAPI 版本仅在开发调试模式开放写入；请先在独立测试工作簿验收写入与撤销。</p>}
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
        决策必需的事实提到首屏：**到底改了什么**。
        此前只有"预计写入内容"整段收在 `<details>` 里，用户要在按下
        「生成右侧副本并排审阅」之前做决定，却看不到实际改动。
        数据全部来自已在内存里的计划与授权读取快照 —— 这里**不做**额外宿主读取。
      */}
      <div className="wps-plan-change-preview" aria-label="本次修改的修改前与修改后">
        <strong className="wps-plan-change-title">修改前 → 修改后</strong>
        {planChangePreview.actions.map((action) => (
          <div className="wps-plan-change-block" key={action.key}>
            <span className="wps-plan-change-head">
              {action.label} · {action.sheetName}!{action.range}
              <small>共 {action.totalCells} 格</small>
            </span>
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
            {action.truncated && (
              <small className="wps-plan-change-more">
                此处仅展示代表性样例（前几行、前几列）；完整内容可展开下方「查看预计结果」。
              </small>
            )}
          </div>
        ))}
      </div>
      <div className="wps-native-review-status" role="status">
        {stagedReview ? (
          <>
            {/*
              先说用户能懂的结论，再说技术依据。
              结论与依据都要**准确**：副本核对通过 ≠ 原表已写入，所以结论是
              "副本已核对，原表还没有动"，不是"写入成功"。
            */}
            <strong className="wps-review-outcome">
              副本已核对通过，<em>原表还没有动</em>
            </strong>
            <p className="wps-review-line">
              请在 WPS 并排窗口比一眼实际值；你确认后才写回原表。
            </p>
            <p className="wps-review-line wps-review-detail">
              {stagedWriteKind === "formula" ? "副本已试算公式并回读计算值" : "副本已写入固定值并回读"}
              ；
              <span className="wps-review-name" title={redactCredentials(stagedReview.original.Name)}>「{redactCredentials(stagedReview.original.Name)}」</span>
              未写入，副本
              <span className="wps-review-name" title={redactCredentials(stagedReview.copy.Name)}>「{redactCredentials(stagedReview.copy.Name)}」</span>
              标记了 {stagedReview.changedCells.length} 处差异。此处是<strong>副本回读</strong>结果，不是预计值。
            </p>
            {reviewPaneNeedsResize && <p className="wps-review-line">AI 窗格自动比例尚未确认。请拖动「Excel Bro」左侧分隔线，将 AI 调到 WPS 主窗口约 {Math.round(REVIEW_PANE_LIMITS.reviewTaskPaneWidthFraction * 100)}%，为两张表留出空间。</p>}
          </>
        ) : (
          <>
            <strong className="wps-review-outcome">待你确认：尚未写回原表</strong>
            <p className="wps-review-line">生成右侧独立副本后会在副本上试算并回读；只有核对通过、你确认之后才会写回原表。</p>
          </>
        )}
      </div>
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
              <strong>{action.type === "clearRange" ? "清空内容" : action.type === "writeFormulas" ? "写入公式" : "写入值"} · {action.sheetName}!{action.range}</strong>
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
            </div>
          ))}
        </div>
      </details>
      <details className="wps-plan-disclosure">
        <summary>查看公式与完整计划</summary>
        <div className="wps-plan-actions-list">
          {pendingPlan.actions.map((action) => (
            <div key={`${action.actionIndex}:${action.sheetName}:${action.range}`}>
              <strong>{action.type === "clearRange" ? "清空内容" : action.type === "writeFormulas" ? "写入公式" : "写入值"}</strong>
              <span>{action.sheetName}!{action.range}</span>
              <small>{action.rowCount} 行 × {action.columnCount} 列</small>
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
            </div>
          ))}
        </div>
      </details>
      {stagedReview && (
        <details className="wps-plan-disclosure">
          <summary>查看副本回读结果（已试算）</summary>
          <p className="wps-plan-disclosure-note">
            以下是<strong>副本</strong>试算后的真实回读值 —— 与上面的"预计"是两回事，原表尚未写入。
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
        操作区**只放按钮**：长说明留在上面的 `.wps-native-review-status`。
        宽窗格水平排列、窄窗格整行纵排（见 styles.css），按钮文案不逐字断行。
      */}
      <div className="wps-native-review-actions">
        {planEditDraft === null ? (
          stagedReview ? (
            <>
              <button type="button" disabled={busy} onClick={discardStagedPlan}>放弃副本</button>
              <button className="primary" type="button" disabled={busy} onClick={() => void confirmPendingPlan()}>确认写回原表</button>
            </>
          ) : (
            <>
              <button type="button" disabled={busy} onClick={() => { setPlanEditDraft(JSON.stringify(pendingPlan.plan, null, 2)); setPlanEditError(null); }}>编辑计划</button>
              <button type="button" disabled={busy} onClick={() => { markPlanPreview("cancelled"); setPendingPlan(null); setPendingPlanMessageId(null); }}>取消</button>
              <button className="primary" type="button" disabled={busy} onClick={stagePendingPlan}>生成右侧副本并排审阅</button>
            </>
          )
        ) : (
          <>
            <button type="button" disabled={busy} onClick={() => { setPlanEditDraft(null); setPlanEditError(null); }}>取消编辑</button>
            <button className="primary" type="button" disabled={busy} onClick={() => void applyEditedPlan()}>应用修改并重新预览</button>
          </>
        )}
      </div>
    </section>
  );

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
        {!writeProgress && <button type="button" onClick={() => controllerRef.current?.abort()}>停止</button>}
      </div>
    </section>
  );

  return (
    <section className="wps-chat-shell chat-shell">
      {targetGuard.notice && <AuthorizedTargetConfirmation notice={targetGuard.notice}
        error={targetGuard.error} busy={targetGuard.confirming}
        onConfirm={(choice, selected) => void targetGuard.confirm(choice, selected)} onCancel={targetGuard.cancel} />}
      <header
        className={`chat-header${modelMenuOpen || moreMenuOpen ? " menu-open" : ""}`}
      >
        <div className="brand-mark">EB</div>
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
                    disabled={!model.available}
                    onClick={() => {
                      if (!isRealModelId(model.id)) return;
                      modelManagement.selectModel(model.id);
                      setModelMenuOpen(false);
                    }}
                  >
                    <i aria-hidden="true" />
                    <span>{model.label}{model.available ? "" : "（未配置）"}</span>
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
                      <div className="tool-card-footer"><span className="tool-card-meta">{tool.planTemplate.actions.length} 个步骤 · 点击后预览</span></div>
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
                <button type="button" disabled={busy || undoRecord.state === "undo_conflict"} onClick={() => void undoExecution(undoRecord.executionId)}>↶ 撤销本次执行</button>
              </section>
            ) : (
              <div className="tool-empty"><i>↶</i><strong>还没有恢复点</strong><span>只有确认并完成写入后才会创建，不会生成空备份。</span></div>
            )}
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
                  ? "你好。直接告诉我你想查询、分析或修改什么；涉及写入时，我会先展示计划并等你确认。"
                  : `你好。你可以查询和分析工作表；写入前还需通过当前文件检查：${writeStatusReason}`}
              </p>
            </div>
          </article>
        )}

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
              恢复仅核对当前已确认范围内的登记区域；重新预览仍须你再次确认写回原表。
            </p>
            {/*
             * 按钮**只**来自判定给的 `wpsReconcileActions`：判定说证据不足/现场冲突
             * 就一个写入或补记按钮都不画，"继续"只在真的能安全续跑时出现。
             */}
            {wpsRecovery.verdict.kind === "replan" && <label>
              <input type="checkbox" checked={recoveryScopeConfirmed}
                onChange={(event) => setRecoveryScopeConfirmed(event.target.checked)} />
              确认使用数据范围中当前已选工作表重新读取；新预览需另行确认写入。
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

        {messages.map((message) => (
          <Fragment key={message.id}>
            <article className={`message-row ${message.role}`}>
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
              ) : message.text ? (
                <p className="message-text">{message.text}</p>
              ) : null}
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
                pendingPlan && pendingPlanMessageId === message.id
                  ? planCard
                  : <div className="wps-plan-outcome" role="status">
                      <strong>{message.planPreview.title}</strong>
                      <span>{message.planPreview.status === "executed"
                        ? "计划已执行并回读确认"
                        : message.planPreview.status === "cancelled"
                          ? "计划已取消，未写入工作表"
                          : "计划预览已失效；如需写入，请重新提出任务"}</span>
                    </div>
              )}
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
                const active = planningPause?.turnId === message.planningPause!.turnId;
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
                const reloadBlocked =
                  planningPauseReload?.checking === false &&
                  planningPauseReload.turnId === message.planningPause!.turnId
                    ? planningPauseReload.blockedReason
                    : null;
                return (
                  <section
                    className="wps-planning-pause-card"
                    aria-label="规划已暂停"
                  >
                    <div className="wps-planning-pause-status" role="status">
                      <strong>
                        {WPS_PLANNING_PAUSE_LABEL[message.planningPause!.pauseType]}
                      </strong>
                      <span>{message.planningPause!.reason}</span>
                    </div>
                    {reloading ? (
                      <small className="wps-planning-pause-note">
                        正在核对该任务的真实进度…
                      </small>
                    ) : active && view && view.canContinue && view.action === "continue_checkpoint" ? (
                      <div className="wps-planning-pause-actions">
                        <button
                          type="button"
                          disabled={busy}
                          onClick={() => void resumePlanningFromPause(planningPause!)}
                        >
                          {view.actionLabel || "继续任务"}
                        </button>
                        <small>{view.detail}</small>
                      </div>
                    ) : (
                      <small className="wps-planning-pause-note">
                        {reloadBlocked
                          ? `该任务当前不能精确续跑：${reloadBlocked}`
                          : active && view && !view.canContinue
                            ? `该任务当前不能精确续跑：${view.reason}`
                            : "这条任务的暂停现场已被新的任务取代；如需继续，请重新发送需求。"}
                      </small>
                    )}
                  </section>
                );
              })()}
              {message.verification && (
                <section
                  className={`verification-card wps-verification-card ${wpsExecutionCardTone({
                    outcome: normalizeWpsExecutionOutcome(message.verification.outcome),
                    undone: message.verification.status === "undone"
                  })}`}
                  aria-label="写入执行结果"
                >
                  <strong>
                    {message.verification.status === "undone"
                      ? "已撤销并核对恢复"
                      : WPS_EXECUTION_OUTCOME_LABEL[normalizeWpsExecutionOutcome(message.verification.outcome)]}
                  </strong>
                  <span>
                    {message.verification.actionCount} 个动作 · {message.verification.cellCount}{" "}
                    {WPS_EXECUTION_OUTCOME_DETAIL_SUFFIX[normalizeWpsExecutionOutcome(message.verification.outcome)]}
                  </span>
                  <ul>
                    {message.verification.targets.slice(0, 5).map((target, index) => (
                      <li key={`${message.id}-target-${index}`}>
                        {target.type === "clearRange" ? "清空内容" : target.type === "writeFormulas" ? "写入公式" : "写入值"} · {target.sheetName}!{target.range}
                      </li>
                    ))}
                  </ul>
                  {message.verification.targets.length > 5 && (
                    <small>另有 {message.verification.targets.length - 5} 个目标，详见执行过程。</small>
                  )}
                  <small>{SESSION_RECOVERY_NOTICE_PREFIX}；关闭后不能从此卡片撤销。</small>
                  {/*
                    撤销入口就挂在这张卡里，归属靠 `executionId`。
                    只有"当前会话里的活跃恢复点正是这一次、且允许撤销"才给可点按钮；
                    被新执行取代、面板重开、或已撤销之后，这里只会留一句状态说明。
                  */}
                  <div className={`wps-undo-inline ${wpsRecoveryCardStateFor(message.verification)}`}>
                    <span>{WPS_RECOVERY_CARD_LABEL[wpsRecoveryCardStateFor(message.verification)]}</span>
                    {wpsRecoveryCardStateFor(message.verification) === "undoable" && (
                      <button
                        type="button"
                        aria-label="撤销本次执行"
                        disabled={busy}
                        onClick={() => void undoExecution(message.verification!.executionId)}
                      >
                        ↶ 撤销本次执行
                      </button>
                    )}
                  </div>
                </section>
              )}
              {message.changeReviewExecutionId && (
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
                  onStop={activeActivityMessageId === message.id && progress
                    ? () => controllerRef.current?.abort()
                    : undefined}
                />
              )}
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

        {(progress || writeProgress) && !activeActivityMessageId && progressCard}
      </main>
      {hasUnseenProgress && (
        <button type="button" className="new-progress-button" onClick={scrollToLatest} aria-label="查看最新执行进度" title="查看最新进度">
          <span className="progress-dots" aria-hidden="true">
            <span />
            <span />
            <span />
          </span>
        </button>
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

      <footer className="composer">
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
          <form
            className="composer-box"
            onSubmit={(event) => {
              event.preventDefault();
              void runTurn(prompt);
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
                  void runTurn(prompt);
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
              {progress ? (
                <button
                  className="send-button composer-send-button wps-stop-button"
                  type="button"
                  aria-label="停止任务"
                  title="停止当前读取或规划任务"
                  onClick={() => controllerRef.current?.abort()}
                >
                  ■
                </button>
              ) : (
                <button
                  className="send-button composer-send-button"
                  type="submit"
                  aria-label={writeProgress ? "正在执行写入与回读" : "发送"}
                  title={writeProgress ? "写入与回读期间不能中断，以免结果状态不明" : undefined}
                  disabled={Boolean(writeProgress) || preparingImages ||
                    (!pendingImages.length && prompt.trim().length < 2) ||
                    (pendingImages.length > 0 && !selectedModel?.supportsVision) || selectedSheetIndexes.length === 0}
                >
                  {writeProgress ? "…" : "↑"}
                </button>
              )}
            </div>
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
