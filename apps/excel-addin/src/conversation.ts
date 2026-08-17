// 对话 / 历史相关的纯函数与常量（从 App.tsx 抽出，避免单文件过大的维护负担）。
import type {
  ChatConversation,
  ChatHistoryState,
  ChatMessage
} from "./types/chat";
import type {
  QueryTableArguments,
  IntentScopeContext,
  VerificationReport
} from "./contracts";
import capabilities from "../../../config/capabilities.json";

export const CHAT_STORAGE_KEY = "excel-bro.chat.v4";
export const LEGACY_CHAT_STORAGE_KEY = "excel-bro.chat.v3";
export const MODEL_STORAGE_KEY = "excel-bro.model.v2";
export const PET_VISIBILITY_STORAGE_KEY = "excel-bro.pet.visibility.v1";
export const MAX_STORED_CONVERSATIONS =
  capabilities.conversation.maxStoredConversations;
export const MAX_MESSAGES_PER_CONVERSATION =
  capabilities.conversation.maxMessagesPerConversation;
export const PERSISTED_MESSAGES_PER_CONVERSATION =
  capabilities.conversation.persistedMessagesPerConversation;
export const INTENT_HISTORY_MESSAGES =
  capabilities.conversation.intentHistoryMessages;
export const INTENT_MESSAGE_CHARACTERS =
  capabilities.conversation.intentMessageCharacters;
export const MAX_CLARIFICATION_ROUNDS =
  capabilities.conversation.maxClarificationRounds;
export const INTENT_MAX_FIELDS = capabilities.intentContext.maxFieldsPerSheet;
export const INTENT_MAX_PRIOR_RESULT_ROWS =
  capabilities.intentContext.maxPriorResultRows;

export function normalizePetVisibility(value: string | null): boolean {
  return value !== "hidden";
}

export function normalizeStoredVerification(
  report: VerificationReport | undefined
): VerificationReport | undefined {
  if (!report) return undefined;
  return {
    ...report,
    status: report.status ?? (report.passed ? "verified" : "failed"),
    unverifiedActions: Array.isArray(report.unverifiedActions)
      ? report.unverifiedActions
      : []
  };
}

// 原样文本命中的一级 key：去首尾空白、内部空白折叠、小写。
export function normalizePrompt(text: string): string {
  return text.trim().replace(/\s+/g, " ").toLowerCase();
}

// 结构化意图 key：归一化 QueryTableArguments + scope 后排序序列化。
// filters 是 AND 关系、顺序无关 => 排序；groupBy/metrics/fields 顺序影响输出 => 保序。
// 归一化从严：误 miss（多算一次）可接受，误命中（拿错结论）危险。
export function normalizeIntentKey(
  args: QueryTableArguments,
  scope: IntentScopeContext
): string {
  const field = (value?: string | null): string =>
    (value ?? "").trim().toLowerCase();
  const filters = (args.filters ?? [])
    .map((filter) => ({
      f: field(filter.field),
      o: filter.operator,
      v: filter.value ?? null
    }))
    .sort((a, b) =>
      `${a.f}|${a.o}|${String(a.v)}`.localeCompare(
        `${b.f}|${b.o}|${String(b.v)}`
      )
    );
  const metrics = (args.metrics ?? []).map((metric) => ({
    op: metric.operation,
    f: field(metric.field),
    out: field(metric.outputName),
    ratio: field(metric.ratioOutputName)
  }));
  return JSON.stringify({
    mode: args.mode,
    scope: args.scope ?? null,
    fields: (args.fields ?? []).map(field),
    filters,
    groupBy: (args.groupBy ?? []).map(field),
    metrics,
    combine: args.combine ?? null,
    profileField: field(args.profileField),
    sortBy: field(args.sortBy),
    sortDirection: args.sortDirection ?? null,
    limit: typeof args.limit === "number" ? args.limit : null,
    workbookName: scope.workbookName,
    sourceMode: scope.sourceMode,
    sheets: [...scope.sheets.map((sheet) => sheet.name)].sort()
  });
}

export function messageId(): string {
  return globalThis.crypto?.randomUUID?.() ??
    `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

export function welcomeMessage(): ChatMessage {
  return {
    id: messageId(),
    role: "assistant",
    text:
      "你好。直接告诉我你想查询、分析或修改什么；涉及写入时，我会先给你预览。",
    createdAt: new Date().toISOString()
  };
}

export function conversationTitle(messages: ChatMessage[]): string {
  const firstUserMessage = messages.find(
    (message) => message.role === "user" && message.text?.trim()
  )?.text;
  if (!firstUserMessage) return "新对话";
  const compact = firstUserMessage.replace(/\s+/g, " ").trim();
  return compact.length > 34 ? `${compact.slice(0, 34)}…` : compact;
}

export function createConversation(
  messages = [welcomeMessage()]
): ChatConversation {
  const now = new Date().toISOString();
  return {
    id: messageId(),
    title: conversationTitle(messages),
    messages,
    createdAt: now,
    updatedAt: now
  };
}

export function deleteConversationFromHistory(
  current: ChatHistoryState,
  conversationId: string
): ChatHistoryState {
  const remaining = current.conversations.filter(
    (conversation) => conversation.id !== conversationId
  );
  if (remaining.length === current.conversations.length) {
    return current;
  }
  if (remaining.length === 0) {
    const replacement = createConversation();
    return {
      activeConversationId: replacement.id,
      conversations: [replacement]
    };
  }
  return {
    activeConversationId:
      current.activeConversationId === conversationId
        ? remaining[0].id
        : current.activeConversationId,
    conversations: remaining
  };
}

export function loadChatHistory(): ChatHistoryState {
  try {
    const stored = JSON.parse(
      localStorage.getItem(CHAT_STORAGE_KEY) ?? "null"
    ) as Partial<ChatHistoryState> | null;
    if (
      stored &&
      typeof stored.activeConversationId === "string" &&
      Array.isArray(stored.conversations) &&
      stored.conversations.length > 0
    ) {
      const conversations = stored.conversations
        .filter(
          (conversation): conversation is ChatConversation =>
            Boolean(
              conversation &&
                typeof conversation.id === "string" &&
                Array.isArray(conversation.messages)
            )
        )
        .map((conversation) => {
          const now = new Date().toISOString();
          return {
            ...conversation,
            messages: conversation.messages.map((message) => {
              const normalizedMessage = {
                ...message,
                verification: normalizeStoredVerification(
                  message.verification
                )
              };
              return message.clarification
                ? {
                    ...normalizedMessage,
                    clarification: {
                      ...message.clarification,
                      round:
                        typeof message.clarification.round === "number"
                          ? message.clarification.round
                          : 0,
                      status:
                        message.clarification.status === "resolving"
                          ? "pending"
                          : message.clarification.status
                    }
                  }
                : normalizedMessage;
            }),
            title:
              typeof conversation.title === "string"
                ? conversation.title
                : conversationTitle(conversation.messages),
            createdAt:
              typeof conversation.createdAt === "string"
                ? conversation.createdAt
                : now,
            updatedAt:
              typeof conversation.updatedAt === "string"
                ? conversation.updatedAt
                : now
          };
        });
      if (conversations.length > 0) {
        return {
          activeConversationId: conversations.some(
            (conversation) =>
              conversation.id === stored.activeConversationId
          )
            ? stored.activeConversationId
            : conversations[0].id,
          conversations
        };
      }
    }
  } catch {
    // Fall through to legacy migration.
  }

  try {
    const legacy = JSON.parse(
      localStorage.getItem(LEGACY_CHAT_STORAGE_KEY) ?? "[]"
    );
    const conversation = createConversation(
      Array.isArray(legacy) && legacy.length > 0
        ? (legacy as ChatMessage[])
        : [welcomeMessage()]
    );
    return {
      activeConversationId: conversation.id,
      conversations: [conversation]
    };
  } catch {
    const conversation = createConversation();
    return {
      activeConversationId: conversation.id,
      conversations: [conversation]
    };
  }
}
