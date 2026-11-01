/** Display history only: no execution, recovery or storage dependencies. */
export function removeHistoryConversation<C extends { id: string }, H extends {
  activeConversationId: string;
  conversations: C[];
  legacyCredentialConversationIds?: string[];
}>(current: H, conversationId: string, create: () => C): H {
  const remaining = current.conversations.filter((item) => item.id !== conversationId);
  if (remaining.length === current.conversations.length) return current;
  if (!remaining.length) remaining.push(create());
  return {
    ...current,
    activeConversationId: current.activeConversationId === conversationId
      ? remaining[0].id : current.activeConversationId,
    conversations: remaining,
    legacyCredentialConversationIds: (current.legacyCredentialConversationIds ?? [])
      .filter((id) => remaining.some((item) => item.id === id))
  };
}
