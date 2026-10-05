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
import {
  containsLikelyCredential,
  redactCredentials,
  REDACTED_PLACEHOLDER
} from "@excel-bro/core";
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

/**
 * 界面上表示「这里原文是凭据，已被抹掉」的占位符。
 *
 * 与 `REDACTED_PLACEHOLDER`（core 里的"已隐藏"，用于**落盘**）刻意分开：
 * 落盘只留一个不含密钥的标记；渲染时把它还原成下面这个更明确的措辞，
 * 既不泄漏原文，也不会让用户以为模型确实收到了"已隐藏"四个字。
 */
export const RESTORED_SECRET_PLACEHOLDER = "（此处原为凭据，已隐藏）";

/**
 * 读取历史时渲染用的措辞还原：把落盘占位符换成更明确的"这里原文是凭据"。
 *
 * 已脱敏的文本里**不会**再含真凭据，所以这一步是纯粹的显示层处理；它把
 * `【已隐藏】` 换成一句人话，让用户知道"这里原本有东西，是凭据，被我抹了"。
 *
 * 它**不是**旧的"读取时脱敏"的替代品。旧实现（只还原、不脱敏）会让旧记录里的
 * 真密钥原样渲染；现在读取路径先脱敏（见 `loadChatHistory`），这里的还原只作用
 * 在已经干净的文本上。
 *
 * `REDACTED_PLACEHOLDER` 用带方括号的结构化写法，正是为了让这一步能安全地区分
 * "脱敏留下的标记"和"用户自己写的'已隐藏'"：普通文字不含 `【】`，不会被误还原。
 */
export function replaceRestoredSecrets(text: string): string {
  return text.split(REDACTED_PLACEHOLDER).join(RESTORED_SECRET_PLACEHOLDER);
}

/**
 * 文本里是否含有会被脱敏改写的疑似凭据原文。
 *
 * 只用于**读取原始存储**时的检测（见 `loadChatHistory`），不作为"是否陈旧"的
 * 持久判据 —— 那件事由 `ChatHistoryState.legacyCredentialConversationIds` 记录，
 * 因为它必须在脱敏、落盘、重开之后仍然成立，而"密钥匹配"在脱敏后就消失了。
 */
export function containsCredentialText(text: string | undefined): boolean {
  return containsLikelyCredential(text ?? "");
}

/**
 * 落盘前的历史脱敏：把**会随历史持久化**的文本里的凭据抹掉。
 *
 * 对话文本可能来自服务端错误（`str(error)` 会把请求里的 apiKey 带出来）、
 * 模型回答，或用户自己粘贴的内容。历史存进宿主 localStorage，
 * 用户可读可复制、诊断包也可能带走，所以不能原样落盘。
 *
 * 覆盖的字段与理由（**这不是"只处理 message.text"**）：
 * - `conversation.title`：标题取首条用户消息的前 34 字，用户完全可能把
 *   密钥连同问题一起粘进输入框；只脱敏 text 会让标题成为绕过通道。
 * - `message.text`：错误的原样文本、模型回答、用户输入。
 *
 * 刻意**不**处理的字段：`plan` / `resultContext` / `activityLog` 等结构化
 * 内容。它们由各自的受控构造器产出，且 `resultContext.rows` 是用户显式要求
 * 查看的查询结果 —— 改变它会破坏功能，其保存策略另行定义，不在这里顺手改。
 */
export function redactHistoryCredentials(
  history: ChatHistoryState
): ChatHistoryState {
  return {
    ...history,
    legacyCredentialConversationIds: remainingLegacyCredentialConversationIds(
      history
    ),
    conversations: history.conversations.map((conversation) => {
      const safeTitle = redactCredentials(conversation.title ?? "");
      const messages = conversation.messages.map((message) => {
        const safe = redactCredentials(message.text ?? "");
        return safe === message.text ? message : { ...message, text: safe };
      });
      if (safeTitle === conversation.title && messages.every((message, index) => message === conversation.messages[index])) {
        // 整条对话无需改动时原样返回，避免每次落盘都重建整棵历史树。
        return conversation;
      }
      return {
        ...conversation,
        title: safeTitle,
        messages
      };
    })
  };
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

export function normalizeStoredMessage(message: ChatMessage): ChatMessage {
  const staleSourceConfirm = message.staleSourceConfirm
    ? {
        ...message.staleSourceConfirm,
        status:
          message.staleSourceConfirm.status === "completed"
            ? ("completed" as const)
            : ("pending" as const)
      }
    : undefined;
  const planningResume = message.planningResume
    ? {
        ...message.planningResume,
        status:
          message.planningResume.status === "resuming"
            ? ("pending" as const)
            : message.planningResume.status
      }
    : undefined;
  const normalizedMessage: ChatMessage = {
    ...message,
    plan: message.plan ?? staleSourceConfirm?.plan,
    verification: normalizeStoredVerification(message.verification),
    staleSourceConfirm,
    planningResume
  };
  if (!message.clarification) return normalizedMessage;
  return {
    ...normalizedMessage,
    clarification: {
      ...message.clarification,
      round:
        typeof message.clarification.round === "number"
          ? message.clarification.round
          : 0,
      status:
        message.clarification.localMapping &&
        (message.clarification.status === "pending" || message.clarification.status === "resolving")
          ? "invalidated"
          : message.clarification.status === "resolving"
            ? "pending"
            : message.clarification.status
    }
  };
}

export function canAcknowledgeStaleSource(message: ChatMessage): boolean {
  const confirmation = message.staleSourceConfirm;
  return Boolean(
    confirmation &&
      confirmation.status !== "completed" &&
      message.plan?.id === confirmation.plan.id
  );
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
  const remainingLegacyIds = remainingLegacyCredentialConversationIds(
    current
  ).filter((id) => id !== conversationId);
  if (remaining.length === 0) {
    const replacement = createConversation();
    return {
      activeConversationId: replacement.id,
      conversations: [replacement],
      legacyCredentialConversationIds: remainingLegacyIds
    };
  }
  return {
    activeConversationId:
      current.activeConversationId === conversationId
        ? remaining[0].id
        : current.activeConversationId,
    conversations: remaining,
    legacyCredentialConversationIds: remainingLegacyIds
  };
}

/**
 * 会话是否被标为"旧记录里曾带真凭据"。
 *
 * 这是**独立且可持久化**的判据：一旦某条记录在读取时被脱敏过，就登记在这里，
 * 之后无论该会话怎样落盘、重开，标记都还在。它与"脱敏后的文本里还能不能匹配
 * 到密钥"无关 —— 后者在脱敏那一刻就永久消失了，用它做判据会导致重开后发送闸
 * 自动放宽，与"始终拦截旧会话"矛盾。
 *
 * 只有用户主动忽略、或该会话在内存里被整体替换（删除 / 新建）时才清除。
 */
export function isLegacyCredentialConversation(
  history: ChatHistoryState,
  conversationId: string | undefined
): boolean {
  if (!conversationId) return false;
  return (history.legacyCredentialConversationIds ?? []).includes(
    conversationId
  );
}

/** 把标记去重、去掉已不存在的会话，并兼容旧对象里没有该字段的情况。 */
export function remainingLegacyCredentialConversationIds(
  history: ChatHistoryState
): string[] {
  const ids = Array.isArray(history.legacyCredentialConversationIds)
    ? history.legacyCredentialConversationIds
    : [];
  const existing = new Set(history.conversations.map((c) => c.id));
  return [...new Set(ids)].filter((id) => existing.has(id));
}

/** 清除某会话的"旧记录"标记（用户明确要求继续这段历史时调用）。 */
export function clearLegacyCredentialMark(
  history: ChatHistoryState,
  conversationId: string
): ChatHistoryState {
  const next = remainingLegacyCredentialConversationIds(history).filter(
    (id) => id !== conversationId
  );
  return { ...history, legacyCredentialConversationIds: next };
}

/** v3 是裸消息数组（没有会话外壳）。 */
function readLegacyV3Messages(): ChatMessage[] | null {
  try {
    const legacy = JSON.parse(
      localStorage.getItem(LEGACY_CHAT_STORAGE_KEY) ?? "null"
    );
    if (!Array.isArray(legacy)) return null;
    return legacy.filter(
      (message): message is ChatMessage =>
        Boolean(message && typeof message === "object" && "text" in message)
    );
  } catch {
    return null;
  }
}

/**
 * v3 旧键是否还躺在存储里等着被清理。
 *
 * 与"迁移是否成功"是**两件事**：安全的 v4 记录可能已经写好，但
 * `removeItem(LEGACY_CHAT_STORAGE_KEY)` 失败（宿主对某些键只读、或整块存储只允许
 * 新增）。这种情况下存储里仍然留着含原文的旧键，所以**不能**对外宣称
 * "迁移完成、本机已无凭据"。
 *
 * 只探测**确切的那一个键**：读别的键既无必要，也可能把用户的其他数据卷进来。
 */
export function isLegacyV3KeyPendingCleanup(): boolean {
  try {
    return localStorage.getItem(LEGACY_CHAT_STORAGE_KEY) !== null;
  } catch {
    // 存储整体不可读时无法判定；宁可保守地报告"可能还有残留"，也不要谎报干净。
    return true;
  }
}

/**
 * 只清理**确切的那一个 v3 旧键**。
 *
 * 刻意不做 `localStorage.clear()`、也不遍历删除 —— 那会连 v4 历史、模型选择、
 * 宠物可见性等用户数据一起抹掉。删不掉（抛错）时返回 false，由调用方决定
 * 是否继续提示，绝不静默当作已完成。
 */
export function removeLegacyV3Key(): boolean {
  try {
    localStorage.removeItem(LEGACY_CHAT_STORAGE_KEY);
  } catch {
    return false;
  }
  try {
    return localStorage.getItem(LEGACY_CHAT_STORAGE_KEY) === null;
  } catch {
    return false;
  }
}

/**
 * v3 → v4 迁移：**先让安全的 v4 记录落盘成功，再清理旧键**。
 *
 * 旧实现把 v3 消息直接交给 `createConversation` 返回，既没有检测/脱敏/打标
 * （凭据旁路），也从不处理 `LEGACY_CHAT_STORAGE_KEY` —— 那个键里的原文会一直
 * 留在存储里，"迁移完成"名不副实。
 *
 * 清理顺序是可恢复的，两步都不能省：
 *
 * 1. **先写安全的 v4**。新记录整体过一遍 `createConversation`：标题取首条用户
 *    消息前 34 字，是绕过 `message.text` 脱敏的第二条通道，所以创建后立刻
 *    `redactHistoryCredentials`（覆盖 title 与 text）。检出凭据的会话登记进
 *    `legacyCredentialConversationIds` —— 独立且可持久化，发送闸查它。
 * 2. **写成功之后才动旧键**。只有 `setItem(v4)` 真的成功，才删掉 `LEGACY_CHAT_STORAGE_KEY`。
 *    写失败时**不删旧键**（否则用户的数据就此丢失），返回内存态时也保持脱敏 +
 *    打标 —— 原文不渲染、不发送，发送闸不因落盘失败而松开。
 *
 * 删旧键失败**不是静默可忽略的**：旧键里仍然是原文，存储整体并不干净。此时只
 * 做两件事 —— 保留 v4 的干净记录、把"待清理"如实报告出去（`recoverable: false`）。
 * 清理会在下一次启动或存储恢复后由 `useConversation` 的持久化 effect 重试；能做到
 * 是因为 `removeLegacyV3Key` 只碰确切的那一个键，重试是安全且幂等的。
 *
 * 与落盘 effect 一样，这里没有 `replaceRestoredSecrets`：返回的对象会再次被
 * `useConversation` 的持久化 effect `JSON.stringify` 覆盖写回 v4。若这里把
 * `【已隐藏】` 还原成人话，`redactHistoryCredentials` 就不再匹配，下一轮循环
 * 真凭据会重新渲染出来（`chatHistory.test.ts` 的
 * "is stable when a rendered history is saved and reopened again" 正是钉这个）。
 * 措辞还原只属于 `useConversation` → `renderRestoredSecrets`。
 */
function migrateLegacyV3Conversation(): ChatHistoryState {
  const legacyMessages = readLegacyV3Messages();
  const conversation = createConversation(
    legacyMessages && legacyMessages.length > 0
      ? legacyMessages
      : [welcomeMessage()]
  );
  const carriedCredential =
    legacyMessages !== null &&
    (containsCredentialText(conversation.title) ||
      conversation.messages.some((message) =>
        containsCredentialText(message.text)
      ));
  const migrated: ChatHistoryState = {
    activeConversationId: conversation.id,
    conversations: [conversation],
    legacyCredentialConversationIds: carriedCredential ? [conversation.id] : []
  };
  // 关键次序：安全的 v4 记录必须**先落盘成功**，旧键才可以清理。
  const safe = redactHistoryCredentials(migrated);
  try {
    localStorage.setItem(CHAT_STORAGE_KEY, JSON.stringify(safe));
  } catch {
    // 写不进就保留旧键，让本次（以及下次）仍有机会重新迁移；上面的 `safe`
    // 已经脱敏，返回出去的内存态不含原文。
    return safe;
  }
  // 旧键删得掉就删；删不掉**不改变已完成的事实**，但要在返回值上如实标出
  // "存储里可能仍有原文"。调用方（`useConversation`）据此决定是否提示用户，
  // 并在下一次存储可用时重试。
  const cleaned = removeLegacyV3Key();
  if (!cleaned) {
    // 删不掉时把进度**写回 v4**：只在内存里标一下不够 —— 重开任务窗格会走 v4
    // 分支，若存储里没有这个标记，界面就拿不到它，等于谎报"存储已干净"。
    // 这里再写一次是安全的：前一次 setItem 刚刚成功，内容也仍不含真凭据。
    const pending: ChatHistoryState = {
      ...safe,
      legacyMigrationCleanupPending: true
    };
    try {
      localStorage.setItem(CHAT_STORAGE_KEY, JSON.stringify(pending));
    } catch {
      // 第二次写入失败就只保留内存标记：本次会话仍会如实提示，下次启动
      // 因为旧键还在会重新探测到"待清理"。不丢历史、不泄漏原文。
    }
    return pending;
  }
  return safe;
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
            messages: conversation.messages.map(normalizeStoredMessage),
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
        // 读取路径必须**先脱敏再进入可渲染状态**。这是本次修正的核心：
        //
        // 早先的实现"只还原落盘占位符、不再脱敏"，理由是"读时脱敏会连用户
        // 正常提问一起改掉"。那个顾虑是真的，但方向反了 —— 它换来的代价是
        // 旧记录里的真密钥会**原样**渲染到界面、复制到剪贴板、留在内存里被
        // 后续请求带走；而且加载后的落盘 effect 会把脱敏版本写回存储，于是
        // "陈旧判据"（脱敏后还能不能匹配到密钥）在重开后自动消失。
        //
        // 现在的分工：
        // 1. 读取原始存储 → 立刻 `redactHistoryCredentials`（覆盖 title 与
        //    text，因为标题取首条用户消息，也是泄漏通道）。返回出去的
        //    `conversations` 已经是干净的，界面 / 复制 / 内存发送历史都不含
        //    真密钥。
        // 2. 凡是在原始文本里检出凭据的会话，登记到 `legacyCredentialConversationIds`
        //    —— 一个**独立且可持久化**的标记。发送闸查这个标记，不查"还能不能
        //    匹配密钥"，所以重开后依然拦截。
        // 返回出去的对象整体是**干净的**（title / text 都不含真密钥，只有
        // `【已隐藏】` 占位符）：界面 / 复制 / 内存发送历史都不会拿到原文。
        // 措辞还原由 `useConversation` 的 `renderRestoredSecrets` 负责。
        const legacyIds: string[] = Array.isArray(
          stored.legacyCredentialConversationIds
        )
          ? stored.legacyCredentialConversationIds.filter(
              (id): id is string => typeof id === "string"
            )
          : [];
        const flagged = new Set(legacyIds);
        for (const conversation of conversations) {
          const carriedCredential =
            containsCredentialText(conversation.title) ||
            conversation.messages.some((message) =>
              containsCredentialText(message.text)
            );
          if (carriedCredential) flagged.add(conversation.id);
        }

        const sanitized = redactHistoryCredentials({
          activeConversationId: stored.activeConversationId,
          conversations,
          legacyCredentialConversationIds: [...flagged]
        });

        // 这里**不做**措辞还原（`replaceRestoredSecrets`）。返回的对象是持久化
        // effect 的输入，而那个 effect 会再次 `redactHistoryCredentials`；若现在
        // 就把 `【已隐藏】` 还原成人话，落盘时脱敏匹配不到占位符，下一轮读取又会
        // 把真凭据渲染出来。措辞还原只在界面这一侧做一次，见
        // `useConversation` → `renderRestoredSecrets`。
        const activeConversationId = sanitized.conversations.some(
          (conversation) => conversation.id === stored.activeConversationId
        )
          ? sanitized.activeConversationId
          : sanitized.conversations[0].id;
        // "旧键还没清掉"这个进度必须**原样带出**：它是随历史落盘的持久标记，
        // 若在这里丢掉，重开之后界面就会拿不到它，等价于谎报"存储已干净"。
        // 与 `legacyCredentialConversationIds` 是两件事，各自独立保留。
        return {
          ...sanitized,
          activeConversationId,
          ...(stored.legacyMigrationCleanupPending
            ? { legacyMigrationCleanupPending: true }
            : {})
        };
      }
    }
  } catch {
    // Fall through to legacy migration.
  }

  // v4 缺失或损坏 → 从 v3 迁移。迁移本身负责检测 / 脱敏 / 打标 / 清理旧键，
  // 不再有"原文直接进可渲染状态"的旁路。
  return migrateLegacyV3Conversation();
}
