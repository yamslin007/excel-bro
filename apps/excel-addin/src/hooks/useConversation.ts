import { useState, useCallback, useEffect } from "react";
import type { ChatMessage, ChatHistoryState } from "../types/chat";
import type { HistoryStorageStatus } from "../components/HistoryStorageNotice";
import { redactCredentials } from "@excel-bro/core";
import {
  loadChatHistory,
  createConversation,
  conversationTitle,
  deleteConversationFromHistory,
  isLegacyCredentialConversation,
  isLegacyV3KeyPendingCleanup,
  redactHistoryCredentials,
  remainingLegacyCredentialConversationIds,
  removeLegacyV3Key,
  replaceRestoredSecrets,
  CHAT_STORAGE_KEY,
  MAX_MESSAGES_PER_CONVERSATION,
  MAX_STORED_CONVERSATIONS,
  PERSISTED_MESSAGES_PER_CONVERSATION,
} from "../conversation";

/**
 * 把「落盘占位符」还原成界面措辞。
 *
 * 这是**唯一的措辞还原点**，因为它紧挨着持久化 effect：state 被写回存储的路径
 * 一定会再经过 `redactHistoryCredentials`，占位符因此始终能被识别。若在更早的
 * 地方（例如 `loadChatHistory` 里）就把占位符还原成人话，落盘时那个 `【已隐藏】`
 * 已经不存在了，脱敏匹配不到 —— 真凭据会在下一轮重新渲染出来。
 */
export function renderRestoredSecrets(history: ChatHistoryState): ChatHistoryState {
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

/**
 * 持久化 effect 里对 v3 旧键的清理时机。
 *
 * **写入是清理的前置条件，不是并列步骤。** v3 键是用户在 v4 尚未落盘时**仅剩的**
 * 历史载体，一旦先删掉、而紧接着的 v4 写入又失败，存储里就什么都不剩 —— 历史被
 * 永久清空，比"多留一份原文"严重得多。所以：
 *
 * - `v4Written === false` → **一次都不许碰** v3 键；
 * - `v4Written === true`  → 才允许调用清理（单键 removeItem，见 `removeLegacyV3Key`）。
 *
 * 抽成独立函数是为了让"先写后删"这件事可被单测直接钉住，而不是埋在 effect 里。
 */
export function maybeCleanupLegacyKeyAfterWrite(v4Written: boolean): boolean {
  if (!v4Written) return false;
  return removeLegacyV3Key();
}

/**
 * 从"上一次落盘尝试的事实"推导出存储状态。
 *
 * 关键在于**四个事实必须来自同一次尝试**，不能混用"内存里的标记"和"当前探测到的存储"：
 *
 * - `v4Written`：这一次 `setItem(v4)` 是否真的成功。写失败时**不能**说"已保存"。
 * - `legacyPresent`：这一次是否探测到 v3 旧键还在。没有旧键就不能说"待清理"。
 *
 * 早先的实现在写失败时**无条件**置 `true`，于是"写失败 + 无旧键"这条也会渲染出
 * "旧版本记录待清理 / 历史已经以脱敏形式保存" —— 两句话都不成立。四种组合各自
 * 对应一句真话，所以这里返回状态而不是布尔值。
 *
 * 文案不在这里，而在 `components/HistoryStorageNotice` —— 把文案与判定分开，
 * 是为了让"用户实际看到什么"能被单独渲染断言，而不必去读 hook 的布尔值。
 */
export function resolveHistoryStorageStatus(
  v4Written: boolean,
  legacyPresent: boolean
): HistoryStorageStatus {
  if (v4Written) {
    return legacyPresent ? "saved-legacy-pending" : "saved-clean";
  }
  return legacyPresent ? "save-failed-legacy-present" : "save-failed-no-legacy";
}

/**
 * 对话管理 Hook
 *
 * 职责：
 * - 管理对话历史（chatHistory 状态）
 * - 提供派生状态（activeConversation, messages）
 * - 提供对话操作（新建、切换、删除、更新消息）
 * - 自动持久化到 localStorage
 */
export function useConversation() {
  /**
   * 读取存储**只有一处调用点**，结果同时供 state 初值与存储状态初值使用。
   *
   * `loadChatHistory` 不是纯读：v4 缺失/损坏时它会触发 v3→v4 迁移（写 v4、删 v3）。
   * 早先的实现在这里调了两次 —— `useState` 初值一次、提示初值一次 —— 同一次挂载里
   * 迁移副作用可能被触发两遍。现在合并成一次 `useState`，两处共用同一个快照。
   *
   * **注意口径**：这里能保证的是"源码只有一处调用点"，**不等于**"运行时只调用一次"。
   * `main.tsx` 用 `React.StrictMode` 包裹 `<App />`，开发构建下 React 会有意重放渲染与
   * effect（实测 StrictMode 下这段初始化执行 **2 次**）。所以正确性标准不是调用次数，
   * 而是**结果的不变量**：不丢历史、不泄漏凭据、写入确认前不删 v3、不松开发送闸。
   * 这些由 `chatHistory.test.ts` 的 StrictMode 用例钉住。
   */
  const [initialSnapshot] = useState(() => loadChatHistory());

  // state 里存的是**可渲染**的历史：读取路径留下的落盘占位符在这里统一还原成
  // 界面措辞。持久化 effect 会再脱敏一次，所以占位符的往返是收敛的。
  const [chatHistory, setChatHistory] = useState<ChatHistoryState>(() =>
    renderRestoredSecrets(initialSnapshot)
  );
  const [pendingDeleteConversationId, setPendingDeleteConversationId] = useState<string | null>(null);
  /**
   * 对话历史的存储状态（四种形态之一，见 `resolveHistoryStorageStatus`）。
   *
   * 初值来自**上面那一次**读取：`loadChatHistory` 已经把"清理没完成"写进了历史的
   * `legacyMigrationCleanupPending`；这里直接复用它，不再重复读存储。所以重开任务窗格
   * 也能如实提示，而不是因为本次内存态干净就谎报存储无残留。
   *
   * 注意初值只在"迁移已成功写出 v4、只是旧键没删掉"时才是 `saved-legacy-pending`：
   * 迁移写失败时 `loadChatHistory` 返回的内存态上没有 pending 标记，此时再看
   * `isLegacyV3KeyPendingCleanup()` 才知道有没有旧键要留着。
   */
  const [historyStorageStatus, setHistoryStorageStatus] =
    useState<HistoryStorageStatus>(() => {
      if (initialSnapshot.legacyMigrationCleanupPending) {
        return "saved-legacy-pending";
      }
      // 迁移没写出 v4 时，旧键是否还在决定了提示该说"保存失败（旧记录保留）"
      // 还是"未保存（无旧记录可留）"。
      return resolveHistoryStorageStatus(false, isLegacyV3KeyPendingCleanup());
    });

  // 派生状态：当前活动对话
  const activeConversation =
    chatHistory.conversations.find(
      (conversation) => conversation.id === chatHistory.activeConversationId
    ) ?? chatHistory.conversations[0];

  // 派生状态：当前对话的消息列表
  const messages = activeConversation?.messages ?? [];

  /**
   * 派生状态：当前对话是否被标为"旧记录里曾带真凭据"。
   *
   * 判据是**持久化的标记**（`legacyCredentialConversationIds`），不是"文本里
   * 还能不能匹配到密钥"。读取路径已经把这些会话脱敏过了，脱敏后当然匹配不到；
   * 如果拿"匹配不到"当"已安全"，重开之后发送闸会自动打开，旧会话依然会被当成
   * 正常历史续聊 —— 这正是上一版实现自相矛盾的地方。
   */
  const staleCredentials = isLegacyCredentialConversation(
    chatHistory,
    activeConversation?.id
  );

  /**
   * 更新当前对话的消息列表
   *
   * 用户在当前对话里主动发消息，视为明确放弃"旧记录"的拦截：这条历史已经被
   * 读取路径脱敏，继续聊是新内容，不再携带任何原始凭据。同时清掉标记，否则
   * 发送闸永不解除。标记与消息在同一次 state 更新中一起写，避免中间态把
   * 界面卡在"提示还在、实际已经能发"或反之。
   */
  const setMessages = useCallback((
    update: ChatMessage[] | ((current: ChatMessage[]) => ChatMessage[])
  ) => {
    setChatHistory((current) => {
      const now = new Date().toISOString();
      return {
        ...current,
        legacyCredentialConversationIds: (
          current.legacyCredentialConversationIds ?? []
        ).filter((id) => id !== current.activeConversationId),
        conversations: current.conversations.map((conversation) => {
          if (conversation.id !== current.activeConversationId) {
            return conversation;
          }
          const nextMessages =
            typeof update === "function"
              ? update(conversation.messages)
              : update;
          const trimmedMessages = nextMessages.slice(
            -MAX_MESSAGES_PER_CONVERSATION
          );
          // state 里的文本必须**始终是已脱敏的**。读取路径已经这么做了，但
          // 新写入的消息（用户粘贴、服务端错误原文、模型回答）同样可能带凭据，
          // 而它们会留在内存里被后续请求带走 —— 只在落盘时脱敏不够。
          //
          // 这一步也让"落盘前再脱一次"保持幂等：`【已隐藏】` 不是自然语言，
          // 重复脱敏不会继续改动它。
          const safeMessages = trimmedMessages.map((message) => {
            const safeText = redactCredentials(message.text ?? "");
            return safeText === message.text
              ? message
              : { ...message, text: safeText };
          });
          const safeTitle = redactCredentials(
            conversationTitle(safeMessages)
          );
          return {
            ...conversation,
            title: safeTitle,
            messages: safeMessages,
            updatedAt: now
          };
        })
      };
    });
  }, []);

  /**
   * 创建新对话
   */
  const newChat = useCallback(() => {
    const conversation = createConversation();
    setChatHistory((current) => ({
      activeConversationId: conversation.id,
      conversations: [
        conversation,
        ...current.conversations
      ].slice(0, MAX_STORED_CONVERSATIONS),
      // 标记跟着会话走，只保留仍然存在的那些（新会话本身不带标记）。
      legacyCredentialConversationIds: (
        current.legacyCredentialConversationIds ?? []
      ).filter((id) =>
        current.conversations
          .slice(0, MAX_STORED_CONVERSATIONS - 1)
          .some((item) => item.id === id)
      )
    }));
  }, []);

  /**
   * 切换到指定对话
   */
  const openConversation = useCallback((conversationId: string) => {
    setChatHistory((current) => ({
      ...current,
      activeConversationId: conversationId
    }));
  }, []);

  /**
   * 标记对话待删除（用于确认对话框）
   */
  const deleteConversation = useCallback((conversationId: string) => {
    if (
      !chatHistory.conversations.some(
        (conversation) => conversation.id === conversationId
      )
    ) {
      return;
    }
    setPendingDeleteConversationId(conversationId);
  }, [chatHistory.conversations]);

  /**
   * 确认删除对话
   */
  const confirmDeleteConversation = useCallback(() => {
    if (!pendingDeleteConversationId) return;
    setChatHistory((current) =>
      deleteConversationFromHistory(current, pendingDeleteConversationId)
    );
    setPendingDeleteConversationId(null);
  }, [pendingDeleteConversationId]);

  /**
   * 取消删除对话
   */
  const cancelDeleteConversation = useCallback(() => {
    setPendingDeleteConversationId(null);
  }, []);

  /**
   * 持久化到 localStorage
   * 只保存最近的 PERSISTED_MESSAGES_PER_CONVERSATION 条消息
   *
   * 落盘必须过 `redactHistoryCredentials`：历史进宿主 localStorage，用户可读
   * 可复制、诊断包也可能带走。消息正文、**标题**、以及服务端错误原文都可能
   * 带着凭据，这里是一次性的收口点 —— 漏掉标题就会让"密钥连同问题一起粘进
   * 输入框"成为绕过通道（标题取的就是首条用户消息）。
   *
   * `legacyCredentialConversationIds` 随历史一起落盘：它必须在脱敏之后仍然
   * 存在，否则重开时无从判断哪些会话是旧记录，发送闸会在刷新后失效。
   */
  useEffect(() => {
    const compact: ChatHistoryState =
      chatHistory.conversations.length === 0
        ? { ...chatHistory, legacyMigrationCleanupPending: undefined }
        : {
            activeConversationId: chatHistory.activeConversationId,
            legacyCredentialConversationIds:
              remainingLegacyCredentialConversationIds(chatHistory),
            conversations: chatHistory.conversations
              .slice(0, MAX_STORED_CONVERSATIONS)
              .map((conversation) => ({
                ...conversation,
                messages: conversation.messages.slice(
                  -PERSISTED_MESSAGES_PER_CONVERSATION
                )
              }))
          };
    // 次序是**先写 v4、成功后才清理 v3**。反过来（先前那样：先删再写）在
    // "v4 持续写不进 + v3 删得掉"的组合下会先抹掉用户仅剩的历史，再落盘失败 ——
    // 存储被清空，且无法恢复。写入失败时**一个字节都不许动** v3 键。
    //
    // 提示状态必须由**这一次尝试的两个事实**共同决定，不能只看"是否有 pending 标记"：
    // - 写失败 + 旧键还在 → 说"本次保存失败、旧记录仍保留且未清理"（不能说已保存）；
    // - 写失败 + 没有旧键 → 说"对话历史未保存"（不能说"旧版本记录待清理"，无键可清）；
    // - 写成功 + 删不掉   → 说"脱敏后的 v4 已保存，但旧键仍待清理"；
    // - 写成功 + 删掉了   → 两个警告都消失。
    //
    // 早先的实现写失败时**无条件** `setLegacyCleanupNotice(true)`，把"写失败 + 无旧键"
    // 也渲染成"旧版本记录待清理 / 历史已经以脱敏形式保存" —— 两句都不成立。
    let status: HistoryStorageStatus;
    try {
      localStorage.setItem(
        CHAT_STORAGE_KEY,
        JSON.stringify(redactHistoryCredentials(compact))
      );
    } catch {
      // 存储被浏览器阻止 / 配额已满：本次会话只保留内存态。内存里的历史已在读取
      // 路径脱敏过，所以"没写成功"不会让真凭据泄漏到界面或请求。**旧键原样保留**，
      // 它是用户仅剩的数据；若本次迁移确实带着未清理的旧键，重开时会被重新探测到。
      setHistoryStorageStatus(
        resolveHistoryStorageStatus(false, isLegacyV3KeyPendingCleanup())
      );
      return;
    }
    // 走到这里才算"安全的 v4 已经真的落盘"，此时清理旧键才是安全的。
    let cleanupPending = false;
    if (isLegacyV3KeyPendingCleanup()) {
      cleanupPending = !maybeCleanupLegacyKeyAfterWrite(true);
      if (cleanupPending) {
        // 删不掉：把"待清理"补写进刚刚写成功的 v4 记录，避免重开后丢失。
        // 这次写入失败也不致命（下次启动因旧键还在会重新探测到），但尽力而为。
        try {
          localStorage.setItem(
            CHAT_STORAGE_KEY,
            JSON.stringify(
              redactHistoryCredentials({ ...compact, legacyMigrationCleanupPending: true })
            )
          );
        } catch {
          // 忽略：已经有一个可用的安全 v4，pending 只是进度信息，不影响数据安全。
        }
      }
    }
    // v4 已确认落盘：此时不论旧键清没清掉，都**不是**"保存失败"。
    status = resolveHistoryStorageStatus(true, cleanupPending);
    setHistoryStorageStatus(status);
  }, [chatHistory]);

  return {
    // 状态
    chatHistory,
    activeConversation,
    messages,
    staleCredentials,
    pendingDeleteConversationId,
    /**
     * 对话历史的存储状态：`saved-clean` / `saved-legacy-pending` /
     * `save-failed-legacy-present` / `save-failed-no-legacy`。
     * 界面据此渲染 `HistoryStorageNotice`，不谎报"已保存"或"待清理"。
     */
    historyStorageStatus,
    /** 是否要给出存储提示（`saved-clean` 之外都为 true）。 */
    legacyCleanupNotice: historyStorageStatus !== "saved-clean",

    // 操作
    setMessages,
    newChat,
    openConversation,
    deleteConversation,
    confirmDeleteConversation,
    cancelDeleteConversation,

    // 状态设置（给特殊场景使用）
    setChatHistory,
    setPendingDeleteConversationId,
  };
}
