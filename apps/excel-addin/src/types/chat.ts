import type {
  AnalysisPlan,
  IntentClarification,
  IntentMemory,
  LocalResultCard,
  PlanningPauseResult,
  ResultContext,
  SdkBridgeBindingRequest,
  VerificationReport
} from "../contracts";
import type { CrossTableMatchProposal } from "../crossTableFormula";
import type { ActivityLog } from "../hooks/useActivityProgress";
import type { ReplySegment } from "@excel-bro/core/replySegments";
import type {
  ExecutionRecoveryAction,
  ExecutionRecoveryVerdict
} from "@excel-bro/core/executionRecoveryView";
import type { SourceMode } from "./workbook";

/**
 * Excel 端的恢复结论：`reconcile` 的现场结论在**发现阶段**还没读过
 * （发现只 GET 状态 + 本机判定，不读单元格），所以是 `pending`；
 * 点「核对在途动作」之后服务端回执才把结论换成 applied / not_applied。
 */
export type OfficeRecoveryVerdict = ExecutionRecoveryVerdict<
  "applied" | "not_applied" | "pending"
>;

export type Status = "idle" | "scanning" | "planning" | "tooling" | "executing";
export type MessageRole = "assistant" | "user" | "system";

export interface MessageClarification extends IntentClarification {
  /** 不读取宿主的需求讨论；回答只能回到消息理解，不授予执行资格。 */
  messageOnly?: boolean;
  /** 本地源值澄清的标记。原始值与规则草稿只在任务窗格内存中。 */
  localMapping?: true;
  turnId?: string;
  planningReadPolicy?: "ifNeeded" | "none";
  originalPrompt: string;
  taskContext?: import("@excel-bro/core/agentTaskContext").AgentTaskContext;
  scopeFingerprint: string;
  hadImages?: boolean;
  round: number;
  status: "pending" | "resolving" | "resolved" | "cancelled" | "invalidated";
  resolvedLabel?: string;
}

export interface FunctionPreview {
  phase: "target" | "preview";
  description: string;
  sheet: string;
  writeTarget: string;
  pickingTarget?: boolean;
  targetError?: string;
  version: "modern" | "compat";
  modernFormula: string;
  modernExplanation: string;
  modernResult: string;
  compatFormula: string;
  compatExplanation: string;
  compatResult: string;
  appliedTarget?: string;
  applied?: boolean;
  cancelled?: boolean;
  generateMs?: number;
  /** 公式来源：deterministic=本地确定性拼公式；model=模型生成（默认）。 */
  mode?: "deterministic" | "model";
  /** 跨表匹配提案（目标阶段预填，可改匹配键/取值列；换 AI 生成时清空）。 */
  match?: CrossTableMatchProposal;
  /** 本次公式生成时勾选的外部工作簿文件名（白名单，用于试算与写入放行）。 */
  externalFiles?: string[];
}

/** 归属到某次执行的操作反馈。类型来自共享层，两端不另写一份。 */
export type ChatMessageFeedback = import("@excel-bro/core/operationFeedback").OperationFeedbackRecord;

export interface ChatMessage {
  id: string;
  role: MessageRole;
  text?: string;
  /**
   * 非空 = 这是一条短操作反馈，渲染时按 `executionId` 挂到对应执行记录下面，
   * 不再单独占一行。不同执行身份绝不合并（`groupOperationFeedback` 是唯一判据）。
   */
  feedback?: ChatMessageFeedback | null;
  attachmentNames?: string[];
  plan?: AnalysisPlan;
  resultContext?: ResultContext;
  /**
   * SDK 本地结果卡指针：只有引用与结构，**没有值**。
   * 因此可以安全地随消息持久化；真实值由本机回环端点按引用取回渲染。
   */
  localResult?: LocalResultCard;
  /**
   * 取回结果值时要回带的绑定（产生该结果那一轮的任务 + 工作簿身份）。
   * 会话关闭后服务端按**授权归属**校验，因此不能用"当前"绑定顶替。
   */
  localResultBinding?: SdkBridgeBindingRequest;
  intentMemory?: IntentMemory;
  verification?: VerificationReport;
  clarification?: MessageClarification;
  activityLog?: ActivityLog;
  /**
   * 这条助手回复属于哪一段展示（与执行身份分开）。
   *
   * 一个任务可以有多段回复（暂停 → 继续 → 再暂停 → 再继续），同一时刻只有一段
   * 是 `live`。段身份**不**替代 `operationId` / `requestId`：步骤身份、事件序号、
   * 结果引用与写入回执仍然只认执行身份。
   */
  replySegment?: ReplySegment;
  /** 批量工作表操作超出授权范围时，用结构化引导指向输入框上方的范围入口。 */
  scopeSelectionGuidance?: {
    keeperName: string;
    missingSheetNames: string[];
  };
  /** 仅在本轮活动进行期间存在；刷新后未完成的空占位会被丢弃。 */
  activityPending?: boolean;
  reused?: boolean;
  provider?: "model" | "local" | "system";
  querySourceMode?: SourceMode;
  querySourceSheetNames?: string[];
  querySourceSheetIds?: string[];
  functionPreview?: FunctionPreview;
  executedPlanId?: string;
  /** Set only after the corresponding recovery operation succeeds. */
  executionUndone?: boolean;
  /** 当前工作簿模式下实际执行的唯一 ID，避免同一 plan 重复执行时选错撤销记录。 */
  executedExecutionId?: string;
  /** 文件夹模式下对应的恢复点 ID。 */
  executedUndoId?: string;
  /** 文件夹执行的持久操作 ID，用于刷新恢复结果去重。 */
  folderOperationId?: string;
  /** 服务端已确认存在安全断点；只有用户明确点击后才继续未完成发布。 */
  folderResume?: {
    operationId: string;
    expectedRevision: number;
    status: "pending" | "resuming" | "completed" | "failed";
  };
  /** 已保存工作簿中、受控区域写入的服务端断点。 */
  officeExecutionId?: string;
  officeResume?: {
    operationId: string;
    expectedRevision: number;
    /**
     * 恢复**结论**。与 WPS 读同一份呈现契约
     * （`@excel-bro/core/executionRecoveryView`）：六种 kind、标题、正文、按钮集合
     * 都从共享层来，这里不再用 `blocked | pending` 两态和一句 `reason` 自己拼。
     *
     * `reconcile`（第二步可能已写入、只等补记回执）与 `resume`（第三步还没执行）
     * 是**两个互斥结论**，各自对应一个按钮；不许再合成"核对并继续"。
     */
    verdict: OfficeRecoveryVerdict;
    /**
     * 正在跑的恢复动作（不是判定）。`null` = 空闲。
     *
     * 与 `verdict` 分开，是因为"正在补记"不能把结论改写成另一种结论：
     * 卡片的说法仍按结论给，只是按钮暂时不可点。
     */
    acting?: ExecutionRecoveryAction | null;
    /** 结论正文之外还需要补一句时的附加说明（例如身份不符的具体原因）。 */
    detail?: string;
    /** 服务端已推进到的下一步序号与动作总数：执行记录（详情）里要显示，不放进正文。 */
    nextActionIndex: number;
    totalActions: number;
    /**
     * 下一条要执行的动作目标 —— 就是 `plan.actions[nextActionIndex]` 的 `sheet`/`range`。
     *
     * 只给"写在哪"用的一句话（输入框里那句「继续后将写入「结果3!B2」」）。**不含
     * `values` / `formulas` 矩阵**，也不是新的判定输入：拿不到就传 `null`，
     * 界面改成说"还要执行剩余 N 步"，绝不拼一个假区域。
     */
    nextAction?: { type: string; sheet: string; range?: string | null } | null;
    /**
     * 这次计划的**全部动作目标**（顺序即执行顺序）—— 输入框里那两句
     * 「已写入：结果1!B2、结果2!B2」「尚未写入：结果3!B2」的唯一来源。
     *
     * 仍然是"写在哪"的信息：**不含** `values` / `formulas` 矩阵，也不参与
     * "能不能续跑"的判定。哪些算写入由账本游标与结论里的 `outcome` 决定
     * （`recoveryComposerTargets` 读同一份判据），这里只提供位置。
     *
     * 早于本字段的历史消息没有它 ⇒ `undefined`，正文退回只说步数，绝不编位置。
     */
    targets?: { type: string; sheet: string; range?: string | null }[];
  };
  /** 规划达到本轮安全时长边界后保存的现场；刷新后仍可由用户明确继续。 */
  planningResume?: {
    turnId: string;
    /** Retain the full request when a resumed planner asks for user input. */
    originalPrompt?: string;
    taskContext?: import("@excel-bro/core/agentTaskContext").AgentTaskContext;
    /** 原规划任务的过程消息，用于续跑时复用同一张时间线卡片。 */
    activityMessageId?: string;
    /** 前端最后实际收到并归并的服务端事件序号。 */
    clientLastSequence?: number;
    /** 客户端主动暂停/断线，尚不能声称服务端已有可续跑现场。 */
    awaitingStatus?: boolean;
    expectedRevision: number;
    lastSequence: number;
    status: "pending" | "resuming" | "completed" | "failed";
    reason: string;
    extensionSeconds: number;
    activeElapsedSeconds: number;
    pauseType?: PlanningPauseResult["pauseType"];
    currentModelId?: string | null;
    modelSwitchCount?: number;
    canSwitchModel?: boolean;
    availableModelConnections?: Array<{
      catalogModelId: string;
      label: string;
      modelId: string;
      supportsVision: boolean;
    }>;
    selectedFallbackModelId?: string;
    requestId?: string;
    sourceFingerprint?: string | null;
    sourceFingerprintSheets?: string[];
  };
  /**
   * 计划里含预览时算出的固定值、且源数据已变。带上原计划让用户能在确认后
   * 坚持执行（executePlan 的 acknowledgeStaleSource）。
   */
  staleSourceConfirm?: {
    plan: AnalysisPlan;
    bakedActionCount: number;
    status?: "pending" | "completed";
  };
  createdAt: string;
}

export interface ChatConversation {
  id: string;
  title: string;
  messages: ChatMessage[];
  createdAt: string;
  updatedAt: string;
}

export interface ChatHistoryState {
  activeConversationId: string;
  conversations: ChatConversation[];
  /**
   * 曾经在原始存储里带过真凭据、已被读取路径脱敏的会话 id。
   *
   * 这是一个**独立且可持久化**的标记，用来在脱敏之后继续拦截这些会话的续聊。
   * 不能用"文本里还能不能匹配到凭据"当判据 —— 脱敏之后永远匹配不到，重开后
   * 发送闸会自动打开。只有用户明确清除或该会话被替换时才移除。
   */
  legacyCredentialConversationIds?: string[];
  /**
   * v4 安全记录已经落盘、但 v3 旧键还没能删掉。
   *
   * 只表示"存储里可能仍留有含原文的旧键"，是**迁移清理的进度**，不是历史内容。
   * 它随 v4 一起落盘，所以重开之后界面仍能如实提示"旧键待清理"，而不会因为
   * 一次成功的内存态就谎报存储已干净。清掉旧键后该字段自然消失。
   *
   * 与 `legacyCredentialConversationIds` 刻意分开：后者是"这段会话要拦续聊"，
   * 前者是"本机存储还有残留待收拾"。两件事可以同时成立，也可以只成立一件。
   */
  legacyMigrationCleanupPending?: boolean;
}
