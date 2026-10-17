import type { ChatMessage } from "./types/chat";

/**
 * 把「哪条消息该标成已撤销」的归属规则抽成**纯函数**，便于直接断言。
 *
 * 纪律（与界面同口径）：
 *  - 只有**恢复操作真的成功**才调用这两个函数；失败 / 冲突 / 清空栈 / 容量淘汰 /
 *    页面重载都**不**标已撤销（重载历史归一为 false，见 `normalizeStoredMessage`）。
 *  - 归属优先用**精确标识**：当前工作簿模式用 `executionId`，文件夹模式用 `undoId`。
 *    有精确标识时**只标精确命中的那一条**，同计划多次执行不会串卡。
 *  - `planId` 只在**它能唯一证明归属**时才用作回退：即"匹配同一 planId 的候选消息
 *    **恰好只有一条**"才自动标它。若有多条缺 ID 的历史卡共用同一 planId，**一条都不标**，
 *    并在返回值里说明"无法确认恢复归属" —— 宁可少标，也不把没撤销的卡标成已撤销。
 *  - `planId` 回退还要排除"另一模式"的消息（文件夹恢复不误标工作簿执行，反之亦然）。
 */

export interface UndoneMarking {
  messages: ChatMessage[];
  /** 有精确标识却匹配不到、或 planId 回退无法唯一归属时，为 true。 */
  ambiguous: boolean;
}

function messageBelongsToWorkbook(message: ChatMessage): boolean {
  // 有 executedExecutionId 的是当前工作簿执行；有 executedUndoId 的是文件夹执行。
  // 都没有时视为旧工作簿卡（planId 回退候选），但文件夹消息的 planId 回退由另一函数处理。
  return !message.executedUndoId;
}

/**
 * 当前工作簿撤销：某一条快照恢复成功后，标回产生它的那条消息。
 * `snapshot.executionId` 缺失时（旧快照）走 planId 回退，且**只在候选唯一时**才标。
 */
export function markWorkbookExecutionUndone(
  messages: readonly ChatMessage[],
  snapshot: { executionId?: string | undefined; planId: string }
): UndoneMarking {
  // 精确匹配优先：有 executionId 的快照只标 executionId 命中的那一条。
  if (snapshot.executionId) {
    let matched = false;
    const result = messages.map((message) => {
      const hit = message.executedExecutionId === snapshot.executionId;
      if (hit) matched = true;
      return hit ? { ...message, executionUndone: true } : message;
    });
    return { messages: result, ambiguous: !matched };
  }

  // 无 executionId 的快照：planId 回退。候选 = 缺 executionId 的工作簿卡且 planId 相同。
  const candidates = messages.filter(
    (message) =>
      !message.executedExecutionId &&
      messageBelongsToWorkbook(message) &&
      message.executedPlanId === snapshot.planId
  );
  // 只有候选**恰好唯一**才能证明归属；多条缺 ID 旧卡同 planId ⇒ 一条都不标。
  if (candidates.length !== 1) {
    return { messages: [...messages], ambiguous: true };
  }
  const target = candidates[0];
  return {
    messages: messages.map((message) =>
      message === target ? { ...message, executionUndone: true } : message
    ),
    ambiguous: false
  };
}

/**
 * 文件夹撤销：一次请求可能连带恢复多条，按条目逐一标回各自的消息。
 * 任一 `entry` 缺 `undoId` 时走 planId 回退，同样**只在候选唯一时**才标该 entry。
 */
export function markFolderExecutionUndone(
  messages: readonly ChatMessage[],
  entries: readonly { undoId?: string | undefined; planId: string }[]
): UndoneMarking {
  let ambiguous = false;
  const flagged = new Set<ChatMessage>();

  for (const entry of entries) {
    if (entry.undoId) {
      let matched = false;
      for (const message of messages) {
        if (message.executedUndoId === entry.undoId) {
          flagged.add(message);
          matched = true;
        }
      }
      if (!matched) ambiguous = true;
      continue;
    }
    // planId 回退：候选 = 缺 undoId 的文件夹卡且 planId 相同。
    const candidates = messages.filter(
      (message) =>
        !message.executedUndoId &&
        !message.executedExecutionId &&
        message.executedPlanId === entry.planId
    );
    if (candidates.length !== 1) {
      ambiguous = true;
      continue;
    }
    flagged.add(candidates[0]);
  }

  return {
    messages: messages.map((message) =>
      flagged.has(message) ? { ...message, executionUndone: true } : message
    ),
    ambiguous
  };
}
