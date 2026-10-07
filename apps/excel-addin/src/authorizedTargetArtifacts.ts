import type { ChatMessage } from "./types/chat";

/** Keep history and confirmed answers, but never transplant an old executable preview. */
export function retireTargetMessage(message: ChatMessage): ChatMessage {
  return { ...message,
    text: message.plan ? `${message.text ?? ""}\n原目标计划：${message.plan.title}（目标已重新授权，需重新生成预览）`.trim() : message.text,
    plan: undefined,
    staleSourceConfirm: undefined,
    functionPreview: message.functionPreview ? { ...message.functionPreview, cancelled: true } : undefined,
    clarification: message.clarification?.status === "pending"
      ? { ...message.clarification, status: "invalidated" } : message.clarification
  };
}
