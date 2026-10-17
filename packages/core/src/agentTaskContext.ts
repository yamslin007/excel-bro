import capabilities from "../../../config/capabilities.json";
import { mergeKeyedConfirmation, type KeyedConfirmation } from "./clarificationAnswers";

/** Semantic decisions only. This is neither authorization nor a model transcript.
 * Never construct it from source values, local mapping drafts or assistant output.
 */
export interface AgentTaskContext {
  schemaVersion: 1;
  originalPrompt: string;
  confirmations: KeyedConfirmation[];
}

function boundedText(value: unknown, max: number): string {
  if (typeof value !== "string" || !value.trim() || value.trim().length > max) {
    throw new Error("任务确认内容为空或过长，请缩短说明后继续；原表未写入。");
  }
  return value.trim();
}

export function createAgentTaskContext(
  originalPrompt: string, confirmations: readonly KeyedConfirmation[] = []
): AgentTaskContext {
  const limits = capabilities.taskContext;
  let projected: KeyedConfirmation[] = [];
  for (const entry of confirmations) {
    const decision: KeyedConfirmation = {
      ...(entry.questionId ? { questionId: boundedText(entry.questionId, limits.maxQuestionIdChars) } : {}),
      label: boundedText(entry.label, limits.maxLabelChars),
      value: boundedText(entry.value, limits.maxAnswerChars)
    };
    projected = mergeKeyedConfirmation(projected, decision);
  }
  if (projected.length > limits.maxConfirmations) throw new Error("本次任务的确认项过多，请合并需求后重新开始；原表未写入。");
  return { schemaVersion: 1, originalPrompt: boundedText(originalPrompt, limits.maxPromptChars), confirmations: projected };
}

export function confirmAgentTaskDecision(
  context: AgentTaskContext, decision: KeyedConfirmation, mode: "append" | "replace" = "append"
): AgentTaskContext {
  return mode === "replace" ? createAgentTaskContext(decision.value)
    : createAgentTaskContext(context.originalPrompt, [...context.confirmations, decision]);
}
