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
  fallbackMessage: ChatMessage,
  localResultBinding?: ChatMessage["localResultBinding"]
): ChatMessage[] {
  const resolvedTargetId = targetId ?? fallbackMessage.id;
  const targetIndex = messages.findIndex(
    (message) => message.id === resolvedTargetId
  );
  // 多次分析只展示最近一次；没有新结果时保留原卡片。
  const pythonResultStep = [...log.steps].reverse().find((step) => step.pythonResult);
  const pythonResult = pythonResultStep?.pythonResult ?? undefined;
  const patch = {
    activityPending: false,
    activityLog: log,
    ...(pythonResult && localResultBinding ? { localResult: pythonResult, localResultBinding } : {})
  };
  if (targetIndex < 0) return [...messages, { ...fallbackMessage, ...patch }];

  return messages.flatMap((message, index): ChatMessage[] => {
    if (message.id !== resolvedTargetId) return [message];
    if (index !== targetIndex) return [];
    return [
      {
        ...message,
        ...patch
      }
    ];
  });
}
