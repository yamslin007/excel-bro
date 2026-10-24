import { segmentAcceptsProgress } from "@excel-bro/core/replySegments";
import type { ReplySegment } from "@excel-bro/core/replySegments";
import type { ActivityLog } from "./hooks/useActivityProgress";
import type { ChatMessage } from "./types/chat";

/** 消息上现有的全部段身份（老消息没有段 ⇒ 空数组，按"可以接收"处理）。 */
export function replySegmentsOf(messages: readonly ChatMessage[]): ReplySegment[] {
  return messages.flatMap((message) => (message.replySegment ? [message.replySegment] : []));
}

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

  /*
   * 只有**当前活动段**能接收进度：`superseded`/`closed` 的段已经冻结在它自己的
   * 那一刻，迟到的进度事件不许把它重新激活（这正是"继续后仍然更新上方旧卡"的
   * 那类症状）。判据来自共享层，不在这儿另写一份"行为相似"的副本。
   */
  if (
    !segmentAcceptsProgress(
      replySegmentsOf(messages),
      resolvedTargetId,
      log.operationId ?? null
    )
  ) {
    return messages;
  }

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
