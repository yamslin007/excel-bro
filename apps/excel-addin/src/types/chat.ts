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
  reused?: boolean;
  provider?: "model" | "local";
  querySourceMode?: SourceMode;
  querySourceSheetNames?: string[];
  querySourceSheetIds?: string[];
  functionPreview?: FunctionPreview;
  executedPlanId?: string;
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
