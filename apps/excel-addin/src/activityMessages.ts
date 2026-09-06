import type { ActivityLog } from "./hooks/useActivityProgress";
import type { ChatMessage } from "./types/chat";

export function insertChatMessage(
  messages: ChatMessage[],
  message: ChatMessage,
  activeActivityMessageId: string | null
): ChatMessage[] {
  if (messages.some((item) => item.id === message.id)) {
    return messages;
  }
  if (message.role !== "user" || !activeActivityMessageId) {
    return [...messages, message];
  }
  const activeIndex = messages.findIndex(
    (item) => item.id === activeActivityMessageId
  );
  if (activeIndex < 0) return [...messages, message];
  return [
    ...messages.slice(0, activeIndex),
    message,
    ...messages.slice(activeIndex)
  ];
}

export function persistActivityMessage(
  messages: ChatMessage[],
  log: ActivityLog,
  targetId: string | null,
  fallbackMessage: ChatMessage
): ChatMessage[] {
  const resolvedTargetId = targetId ?? fallbackMessage.id;
  const targetIndex = messages.findIndex(
    (message) => message.id === resolvedTargetId
  );
  if (targetIndex < 0) return [...messages, fallbackMessage];
  return messages.flatMap((message, index) => {
    if (message.id !== resolvedTargetId) return [message];
    if (index !== targetIndex) return [];
    return [
      {
        ...message,
        activityPending: false,
        activityLog: log
      }
    ];
  });
}
