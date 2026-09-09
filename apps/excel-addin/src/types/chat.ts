import type {
  AnalysisPlan,
  IntentClarification,
  IntentMemory,
  ResultContext,
  VerificationReport
} from "../contracts";
import type { CrossTableMatchProposal } from "../crossTableFormula";
import type { ActivityLog } from "../hooks/useActivityProgress";
import type { SourceMode } from "./workbook";

export type Status = "idle" | "scanning" | "planning" | "tooling" | "executing";
export type MessageRole = "assistant" | "user" | "system";

export interface MessageClarification extends IntentClarification {
  turnId?: string;
  originalPrompt: string;
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

export interface ChatMessage {
  id: string;
  role: MessageRole;
  text?: string;
  attachmentNames?: string[];
  plan?: AnalysisPlan;
  resultContext?: ResultContext;
  intentMemory?: IntentMemory;
  verification?: VerificationReport;
  clarification?: MessageClarification;
  activityLog?: ActivityLog;
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
    status: "pending" | "resuming" | "completed" | "failed";
    requiresReconciliation: boolean;
    nextActionIndex: number;
    totalActions: number;
    reason: string;
  };
  /** 规划达到本轮安全时长边界后保存的现场；刷新后仍可由用户明确继续。 */
  planningResume?: {
    turnId: string;
    expectedRevision: number;
    lastSequence: number;
    status: "pending" | "resuming" | "completed" | "failed";
    reason: string;
    extensionSeconds: number;
    activeElapsedSeconds: number;
    pauseType?: "budget" | "model_error" | "interrupted";
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
}
