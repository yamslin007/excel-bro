import { useState, useCallback, useEffect } from "react";
import type { ChatMessage } from "../App";
import {
  loadChatHistory,
  createConversation,
  conversationTitle,
  deleteConversationFromHistory,
  CHAT_STORAGE_KEY,
  MAX_MESSAGES_PER_CONVERSATION,
  MAX_STORED_CONVERSATIONS,
  PERSISTED_MESSAGES_PER_CONVERSATION,
  type ChatHistoryState
} from "../conversation";

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
  const [chatHistory, setChatHistory] = useState<ChatHistoryState>(loadChatHistory);
  const [pendingDeleteConversationId, setPendingDeleteConversationId] = useState<string | null>(null);

  // 派生状态：当前活动对话
  const activeConversation =
    chatHistory.conversations.find(
      (conversation) => conversation.id === chatHistory.activeConversationId
    ) ?? chatHistory.conversations[0];

  // 派生状态：当前对话的消息列表
  const messages = activeConversation?.messages ?? [];

  /**
   * 更新当前对话的消息列表
   */
  const setMessages = useCallback((
    update: ChatMessage[] | ((current: ChatMessage[]) => ChatMessage[])
  ) => {
    setChatHistory((current) => {
      const now = new Date().toISOString();
      return {
        ...current,
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
          return {
            ...conversation,
            title: conversationTitle(trimmedMessages),
            messages: trimmedMessages,
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
      ].slice(0, MAX_STORED_CONVERSATIONS)
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
   */
  useEffect(() => {
    if (chatHistory.conversations.length === 0) {
      localStorage.setItem(CHAT_STORAGE_KEY, JSON.stringify(chatHistory));
    } else {
      const compact: ChatHistoryState = {
        activeConversationId: chatHistory.activeConversationId,
        conversations: chatHistory.conversations
          .slice(0, MAX_STORED_CONVERSATIONS)
          .map((conversation) => ({
            ...conversation,
            messages: conversation.messages.slice(
              -PERSISTED_MESSAGES_PER_CONVERSATION
            )
          }))
      };
      localStorage.setItem(CHAT_STORAGE_KEY, JSON.stringify(compact));
    }
  }, [chatHistory]);

  return {
    // 状态
    chatHistory,
    activeConversation,
    messages,
    pendingDeleteConversationId,

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
