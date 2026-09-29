// @vitest-environment jsdom
// 本文件用 localStorage 验证「读取旧历史时的职责边界」以及「落盘收口点」。
// vitest 默认走 node 环境（没有 DOM/存储），这里只给本文件开 jsdom，
// 避免影响其余 40+ 个纯逻辑用例。
import { StrictMode, createElement } from "react";
import { act, renderHook, type RenderHookResult } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { REDACTED_PLACEHOLDER } from "@excel-bro/core";
import {
  CHAT_STORAGE_KEY,
  LEGACY_CHAT_STORAGE_KEY,
  MODEL_STORAGE_KEY,
  RESTORED_SECRET_PLACEHOLDER,
  canAcknowledgeStaleSource,
  containsCredentialText,
  deleteConversationFromHistory,
  isLegacyCredentialConversation,
  isLegacyV3KeyPendingCleanup,
  loadChatHistory,
  normalizePetVisibility,
  normalizeStoredMessage,
  normalizeStoredVerification,
  redactHistoryCredentials,
  remainingLegacyCredentialConversationIds,
  removeLegacyV3Key,
  replaceRestoredSecrets
} from "./conversation";
import {
  maybeCleanupLegacyKeyAfterWrite,
  renderRestoredSecrets,
  useConversation
} from "./hooks/useConversation";
import type { AnalysisPlan } from "./contracts";
import type { ChatHistoryState } from "./types/chat";

/** 模拟 App / 面板真实挂载：读取 → 措辞还原 → 落入 state。 */
function mountLikeApp() {
  return renderHook(() => useConversation());
}

// jsdom 下 hook 内部会读 localStorage；每个用例之间清干净，避免相互污染。
afterEach(() => {
  localStorage.clear();
});

function conversation(id: string, updatedAt: string) {
  return {
    id,
    title: `对话 ${id}`,
    messages: [],
    createdAt: updatedAt,
    updatedAt
  };
}

describe("history deletion", () => {
  it("removes an inactive conversation without changing the active one", () => {
    const next = deleteConversationFromHistory(
      {
        activeConversationId: "active",
        conversations: [
          conversation("active", "2026-07-28T00:00:00.000Z"),
          conversation("old", "2026-07-27T00:00:00.000Z")
        ]
      },
      "old"
    );

    expect(next.activeConversationId).toBe("active");
    expect(next.conversations.map((item) => item.id)).toEqual(["active"]);
  });

  it("selects a remaining conversation after deleting the active one", () => {
    const next = deleteConversationFromHistory(
      {
        activeConversationId: "active",
        conversations: [
          conversation("active", "2026-07-28T00:00:00.000Z"),
          conversation("remaining", "2026-07-27T00:00:00.000Z")
        ]
      },
      "active"
    );

    expect(next.activeConversationId).toBe("remaining");
    expect(next.conversations.map((item) => item.id)).toEqual(["remaining"]);
  });

  it("creates a fresh conversation after deleting the last one", () => {
    const next = deleteConversationFromHistory(
      {
        activeConversationId: "only",
        conversations: [
          conversation("only", "2026-07-28T00:00:00.000Z")
        ]
      },
      "only"
    );

    expect(next.conversations).toHaveLength(1);
    expect(next.activeConversationId).toBe(next.conversations[0].id);
    expect(next.activeConversationId).not.toBe("only");
    expect(next.conversations[0].messages).toHaveLength(1);
  });
});

describe("history verification migration", () => {
  it("normalizes verification reports saved before status was introduced", () => {
    const migrated = normalizeStoredVerification({
      passed: true,
      checks: []
    } as never);

    expect(migrated).toMatchObject({
      status: "verified",
      passed: true,
      unverifiedActions: []
    });
  });
});

describe("stale source plan preview", () => {
  const plan = {
    id: "stale-plan",
    title: "写入汇总结果",
    summary: "把预览时计算的汇总值写入结果表",
    assumptions: [],
    warnings: [],
    actions: [
      {
        type: "writeValues",
        sheet: "结果",
        range: "A1",
        values: [[42]]
      }
    ]
  } as AnalysisPlan;

  it("restores the stored plan as a visible preview before stale execution", () => {
    const normalized = normalizeStoredMessage({
      id: "message-1",
      role: "system",
      text: "数据来源已变化",
      staleSourceConfirm: { plan, bakedActionCount: 1 },
      createdAt: "2026-09-01T00:00:00.000Z"
    });

    expect(normalized.plan).toBe(plan);
    expect(normalized.staleSourceConfirm?.status).toBe("pending");
    expect(canAcknowledgeStaleSource(normalized)).toBe(true);
  });

  it("blocks stale execution without the matching visible preview or after completion", () => {
    const message = normalizeStoredMessage({
      id: "message-2",
      role: "system",
      staleSourceConfirm: { plan, bakedActionCount: 1 },
      createdAt: "2026-09-01T00:00:00.000Z"
    });

    expect(canAcknowledgeStaleSource({ ...message, plan: undefined })).toBe(
      false
    );
    expect(
      canAcknowledgeStaleSource({
        ...message,
        staleSourceConfirm: {
          ...message.staleSourceConfirm!,
          status: "completed"
        }
      })
    ).toBe(false);
  });
});

describe("pet visibility preference", () => {
  it("shows the pet by default", () => {
    expect(normalizePetVisibility(null)).toBe(true);
  });

  it("restores an explicitly hidden pet", () => {
    expect(normalizePetVisibility("hidden")).toBe(false);
    expect(normalizePetVisibility("visible")).toBe(true);
  });
});

describe("persisted history credential redaction", () => {
  // WPS 面板直接复用这个函数（`../../excel-addin/src/conversation`），
  // 所以它是两端共同的落盘边界，行为要在本仓锁住。
  function historyWith(
    text: string,
    title = "对话"
  ): ChatHistoryState {
    return {
      activeConversationId: "c1",
      conversations: [
        {
          id: "c1",
          title,
          createdAt: "2026-09-28T00:00:00.000Z",
          updatedAt: "2026-09-28T00:00:00.000Z",
          messages: [
            {
              id: "m1",
              role: "system",
              text,
              createdAt: "2026-09-28T00:00:00.000Z"
            }
          ]
        }
      ]
    };
  }

  it("replaces a leaked API key while keeping the text diagnosable", () => {
    const secret = "sk-live-leaked-from-upstream-9f3a";
    const next = redactHistoryCredentials(
      historyWith(`模型连接校验失败：上游拒绝 Authorization: Bearer ${secret}`)
    );
    const text = next.conversations[0].messages[0].text!;

    expect(text).not.toContain(secret);
    expect(text).not.toContain("sk-");
    // 正向锚点：错误原因还在，凭据位置换成明确占位符。
    expect(text).toContain("模型连接校验失败");
    expect(text).toContain(`Bearer ${REDACTED_PLACEHOLDER}`);
  });

  it("hides a URL-embedded password", () => {
    const next = redactHistoryCredentials(
      historyWith("连接 https://admin:hunter2@example.com/api 超时")
    );
    const text = next.conversations[0].messages[0].text!;

    expect(text).not.toContain("hunter2");
    expect(text).toContain(`https://${REDACTED_PLACEHOLDER}@example.com/api`);
  });

  it("redacts the conversation title, not just message text", () => {
    // 标题取首条用户消息前 34 字。用户把密钥连同问题一起粘进输入框时，
    // 标题会成为绕过 message.text 脱敏的第二条通道。
    const secret = "sk-live-in-title-7b21";
    const next = redactHistoryCredentials(
      historyWith("普通正文", `用 ${secret} 跑一下统计`)
    );

    expect(next.conversations[0].title).not.toContain(secret);
    expect(next.conversations[0].title).not.toContain("sk-");
    // 正向锚点：标题仍然可读，只是密钥位置被换成占位符。
    expect(next.conversations[0].title).toContain(REDACTED_PLACEHOLDER);
    expect(next.conversations[0].title).toContain("跑一下统计");
  });

  it("redacts a JSON-shaped title the same way as bare keys", () => {
    const next = redactHistoryCredentials(
      historyWith("正文", '{"apiKey":"testsecret123456"}')
    );

    expect(next.conversations[0].title).not.toContain("testsecret123456");
    expect(next.conversations[0].title).toContain(REDACTED_PLACEHOLDER);
    expect(next.conversations[0].title).toContain("apiKey");
  });

  it("leaves clean history structurally intact (no needless copying)", () => {
    const history = historyWith("已按 A1:B10 汇总，共 12 行。");
    const next = redactHistoryCredentials(history);

    // 内容相同，值也相等；不要求引用相等 —— 干净时返回同一对象是**实现细节**，
    // 用来省掉每次落盘整棵树的复制。断言值语义才能同时防住
    // 「没脱敏就返回」和「脱敏把干净文本改坏」两种回归。
    expect(next.conversations[0].messages[0].text).toBe("已按 A1:B10 汇总，共 12 行。");
    expect(next.conversations[0].title).toBe("对话");
    // 除标记字段外，整棵树值语义不变。
    expect(next.conversations).toEqual(history.conversations);
    expect(remainingLegacyCredentialConversationIds(next)).toEqual([]);
  });

  it("round-trips through the persisted placeholder", () => {
    const secret = "sk-live-roundtrip-51c0";
    // 落盘：真凭据 → 持久占位符。
    const saved = redactHistoryCredentials(
      historyWith(`用 ${secret} 跑统计`, "对话")
    );
    const savedText = saved.conversations[0].messages[0].text!;
    expect(savedText).not.toContain(secret);
    expect(savedText).toContain(REDACTED_PLACEHOLDER);

    // 渲染：持久占位符 → 可读措辞，且原文不可能从占位符里还原出来。
    const rendered = replaceRestoredSecrets(savedText);
    expect(rendered).toContain(RESTORED_SECRET_PLACEHOLDER);
    expect(rendered).not.toContain(secret);
    expect(rendered).not.toContain("sk-");
    expect(rendered).toContain("跑统计");

    // 还原是幂等的：再跑一次不会把措辞叠成两段。
    expect(replaceRestoredSecrets(rendered)).toBe(rendered);
    // 落盘之后的存储内容里不该再有任何能被识别成凭据的东西。
    expect(savedText).not.toContain("sk-");
  });
});

/**
 * 读取路径：**原始存储 → 立刻脱敏 → 进入可渲染状态**。
 *
 * 这是本轮修正的核心。旧实现"只还原占位符、不再脱敏"，代价是：
 * 1. 旧记录里的真密钥会**原样**渲染到界面、被复制、留在内存里被后续请求带走；
 * 2. 加载后的落盘 effect 把脱敏版本写回存储，于是"还能不能匹配到密钥"这个
 *    陈旧判据在重开后自动消失 —— 与"始终拦截旧会话"直接矛盾。
 *
 * 现在的判据是一个**独立且可持久化**的标记（`legacyCredentialConversationIds`），
 * 它不随脱敏消失，重开后依然成立。
 */
describe("loading legacy history sanitizes before it can be rendered", () => {
  function seed(
    conversations: ChatHistoryState["conversations"],
    activeId: string,
    extra: Record<string, unknown> = {}
  ) {
    localStorage.setItem(
      CHAT_STORAGE_KEY,
      JSON.stringify({
        activeConversationId: activeId,
        conversations,
        ...extra
      })
    );
  }

  function message(id: string, role: "user" | "system", text: string) {
    return { id, role, text, createdAt: "2026-09-27T00:00:00.000Z" };
  }

  const SECRET = "sk-live-in-legacy-title-4a90";

  it("is stable when a rendered history is saved and reopened again", () => {
    // 完整循环：脱敏 → 还原措辞 → 再次落盘 → 再次还原。
    // 还原后的措辞含"已隐藏"三个字，若它又会被当成占位符，二次循环就会不断
    // 加壳。这里钉住"二次循环结果与一次相同"。
    const secret = "sk-live-double-cycle-88ab";
    const history: ChatHistoryState = {
      activeConversationId: "c1",
      conversations: [
        {
          id: "c1",
          title: "对话",
          createdAt: "2026-09-28T00:00:00.000Z",
          updatedAt: "2026-09-28T00:00:00.000Z",
          messages: [
            {
              id: "m1",
              role: "system",
              text: `用 ${secret} 跑统计`,
              createdAt: "2026-09-28T00:00:00.000Z"
            }
          ]
        }
      ]
    };
    const render = (target: ChatHistoryState): ChatHistoryState => ({
      ...target,
      conversations: target.conversations.map((conversation) => ({
        ...conversation,
        title: replaceRestoredSecrets(conversation.title ?? ""),
        messages: conversation.messages.map((message) => ({
          ...message,
          text: replaceRestoredSecrets(message.text ?? "")
        }))
      }))
    });

    const rendered = render(redactHistoryCredentials(history));
    const renderedAgain = render(redactHistoryCredentials(rendered));

    // 二次循环结果与一次相同 —— 不会把"（此处原为凭据，已隐藏）"再加一层。
    expect(renderedAgain.conversations[0].messages[0].text).toBe(
      rendered.conversations[0].messages[0].text
    );
    expect(rendered.conversations[0].messages[0].text).toContain(
      RESTORED_SECRET_PLACEHOLDER
    );
    expect(rendered.conversations[0].messages[0].text).not.toContain(
      RESTORED_SECRET_PLACEHOLDER.repeat(2)
    );
    // 无论循环多少次都不含真密钥。
    expect(renderedAgain.conversations[0].messages[0].text).not.toContain(
      secret
    );
  });

  it("redacts a real secret in both title and message before the first render", () => {
    seed(
      [
        {
          id: "legacy",
          title: `用 ${SECRET} 统计订单`,
          createdAt: "2026-09-27T00:00:00.000Z",
          updatedAt: "2026-09-27T00:00:00.000Z",
          messages: [
            message("m1", "user", `帮我用 ${SECRET} 统计订单`),
            message("m2", "system", `模型连接失败：Authorization: Bearer ${SECRET}`)
          ]
        }
      ],
      "legacy"
    );

    const loaded = loadChatHistory();
    const conversation = loaded.conversations[0];

    // 正向锚点：非敏感内容完整保留，没有静默删除整段历史。
    expect(loaded.conversations).toHaveLength(1);
    expect(conversation.messages).toHaveLength(2);
    expect(conversation.title).toContain("统计订单");
    expect(conversation.messages[0].text).toContain("统计订单");
    expect(conversation.messages[1].text).toContain("模型连接失败");

    // 关键负向断言：读取路径产出的历史（也就是界面、复制、后续请求看到的同一份）
    // 不得含真密钥 —— 无论藏在标题还是消息正文里。
    const everything = [
      conversation.title,
      ...conversation.messages.map((item) => item.text ?? "")
    ].join("\n");
    expect(everything).not.toContain(SECRET);
    expect(everything).not.toContain("sk-");
    // 凭据位置是**结构化占位符**（不是人话措辞）：措辞还原刻意延后到
    // `renderRestoredSecrets`，否则落盘 effect 会因为匹配不到 `【已隐藏】`
    // 而让真凭据在下一轮循环里重新渲染出来。
    expect(conversation.messages[1].text).toContain("Bearer ");
    expect(conversation.messages[1].text).toContain(REDACTED_PLACEHOLDER);
    expect(conversation.title).toContain(REDACTED_PLACEHOLDER);

    // 界面上看到的措辞是还原后的那一句，且依然不含真密钥。
    const rendered = renderRestoredSecrets(loaded);
    expect(rendered.conversations[0].messages[1].text).toContain(
      RESTORED_SECRET_PLACEHOLDER
    );
    expect(rendered.conversations[0].title).toContain(
      RESTORED_SECRET_PLACEHOLDER
    );
    expect(JSON.stringify(rendered)).not.toContain(SECRET);
  });

  it("keeps the legacy flag after the sanitized history is persisted, so the gate survives a reopen", async () => {
    seed(
      [
        {
          id: "legacy",
          title: `用 ${SECRET} 统计订单`,
          createdAt: "2026-09-27T00:00:00.000Z",
          updatedAt: "2026-09-27T00:00:00.000Z",
          messages: [message("m1", "user", `帮我用 ${SECRET} 统计订单`)]
        }
      ],
      "legacy"
    );

    // 第一次加载：立刻脱敏，并登记旧记录标记。
    const first = loadChatHistory();
    expect(first.legacyCredentialConversationIds).toEqual(["legacy"]);

    // 真实落盘：把脱敏后的历史写回存储（模拟 useConversation 的持久化 effect）。
    localStorage.setItem(
      CHAT_STORAGE_KEY,
      JSON.stringify(redactHistoryCredentials(first))
    );
    const rawAfterSave = localStorage.getItem(CHAT_STORAGE_KEY) ?? "";
    expect(rawAfterSave).not.toContain(SECRET);
    expect(rawAfterSave).toContain("legacyCredentialConversationIds");

    // 重开：脱敏之后文本里当然再也匹配不到密钥，但标记仍在，发送闸不放宽。
    const reopened = loadChatHistory();
    expect(isLegacyCredentialConversation(reopened, "legacy")).toBe(true);
    expect(containsCredentialText(reopened.conversations[0].title)).toBe(false);
    expect(
      reopened.conversations[0].messages.some((item) =>
        containsCredentialText(item.text)
      )
    ).toBe(false);
  });

  it("does not rewrite an ordinary question that merely names a credential scheme", () => {
    // 读取时脱敏最容易误伤的功能面：用户就是在问方案名本身。
    const question = "Bearer 是什么？api key 怎么填？";
    seed(
      [
        {
          id: "ask",
          title: question,
          createdAt: "2026-09-27T00:00:00.000Z",
          updatedAt: "2026-09-27T00:00:00.000Z",
          messages: [message("m1", "user", question)]
        }
      ],
      "ask"
    );

    const loaded = loadChatHistory();
    // 一字不改 —— 既不是脱敏，也不是占位符还原。
    expect(loaded.conversations[0].title).toBe(question);
    expect(loaded.conversations[0].messages[0].text).toBe(question);
    // 也不该被误标成旧记录，否则正常提问会被拦。
    expect(isLegacyCredentialConversation(loaded, "ask")).toBe(false);
  });

  it("leaves the plain word 已隐藏 untouched and does not misread it as a credential", () => {
    // 普通文字里出现"已隐藏"三个字，不该被当成脱敏留下的占位符再还原一层，
    // 也不该被误判成含凭据。
    const plain = "这一列我暂时已隐藏，先不用管。";
    seed(
      [
        {
          id: "plain",
          title: plain,
          createdAt: "2026-09-27T00:00:00.000Z",
          updatedAt: "2026-09-27T00:00:00.000Z",
          messages: [message("m1", "user", plain)]
        }
      ],
      "plain"
    );

    const loaded = loadChatHistory();
    expect(loaded.conversations[0].messages[0].text).toBe(plain);
    expect(loaded.conversations[0].title).toBe(plain);
    // 不得被还原成"（此处原为凭据，已隐藏）"。
    expect(loaded.conversations[0].messages[0].text).not.toContain(
      RESTORED_SECRET_PLACEHOLDER
    );
    expect(isLegacyCredentialConversation(loaded, "plain")).toBe(false);
  });

  it("keeps clean persisted history resumable and preserves the legacy flag it already carries", () => {
    seed(
      [
        {
          id: "clean",
          title: `用 ${REDACTED_PLACEHOLDER} 统计订单`,
          createdAt: "2026-09-28T00:00:00.000Z",
          updatedAt: "2026-09-28T00:00:00.000Z",
          messages: [message("m1", "user", `用 ${REDACTED_PLACEHOLDER} 统计订单`)]
        }
      ],
      "clean",
      // 落盘时写下的标记续到下一次读取。
      { legacyCredentialConversationIds: ["clean"] }
    );

    const loaded = loadChatHistory();
    // 文本本来就没问题：内容保留，占位符留待界面层还原（读取路径不再改措辞）。
    expect(loaded.conversations[0].title).toContain("统计订单");
    expect(loaded.conversations[0].title).toContain(REDACTED_PLACEHOLDER);
    // 渲染时才变成人话。
    const rendered = renderRestoredSecrets(loaded);
    expect(rendered.conversations[0].title).toContain(RESTORED_SECRET_PLACEHOLDER);
    // 标记是持久判据：即便文本已经干净，重开后仍标记为旧记录。
    expect(isLegacyCredentialConversation(loaded, "clean")).toBe(true);
  });

  it("drops a legacy flag that no longer has a matching conversation", () => {
    seed(
      [
        {
          id: "clean",
          title: "正常记录",
          createdAt: "2026-09-28T00:00:00.000Z",
          updatedAt: "2026-09-28T00:00:00.000Z",
          messages: [message("m1", "user", "正常记录")]
        }
      ],
      "clean",
      // "gone" 对应的会话已经不在存储里了。
      { legacyCredentialConversationIds: ["gone", "clean"] }
    );

    const loaded = loadChatHistory();
    // 已消失的会话不该继续占着标记位；仍存在的 "clean" 继续保留。
    expect(remainingLegacyCredentialConversationIds(loaded)).toEqual(["clean"]);
    // 换一种更直接的确认：只剩 "gone" 时会被整个丢掉。
    expect(
      remainingLegacyCredentialConversationIds({
        ...loaded,
        legacyCredentialConversationIds: ["gone"]
      })
    ).toEqual([]);
  });
});

describe("persisted round-trip via CHAT_STORAGE_KEY", () => {
  it("scrubs a secret through the real persist effect and keeps the send gate on reopen", async () => {
    // 锁住"落盘收口点在 hook 里、且真的过脱敏"。以前 useConversation 的持久化
    // effect 直接 `JSON.stringify(compact)`，标题根本没过脱敏。
    //
    // 关键：必须跑**真实的**持久化 effect，而不是在测试里再调一遍
    // `redactHistoryCredentials`（那样只证明这个函数本身没问题，会把 hook 里
    // 的遗漏放过 —— 这正是这条用例的由来）。
    const secret = "sk-live-persist-gap-8e13";
    const view = mountLikeApp();
    try {
      // 通过 hook 的公开操作写入一条带凭据的消息：setMessages 会按首条用户
      // 消息生成 title，正是泄漏通道本身。
      await act(async () => {
        view.result.current.setMessages([
          {
            id: "m1",
            role: "user",
            text: `用 ${secret} 统计订单`,
            createdAt: "2026-09-28T00:00:00.000Z"
          }
        ]);
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });

      // **界面内存**（hook 的 state，也就是消息流与后续请求看到的同一份）
      // 不含真凭据。注意它此刻是**已脱敏**的原文（占位符还是 `【已隐藏】`），
      // 措辞还原只在加载时发生 —— 新输入的消息没有经过那一步。
      const inMemory = JSON.stringify(view.result.current.messages);
      expect(inMemory).toContain(REDACTED_PLACEHOLDER);
      expect(inMemory).not.toContain(secret);
      expect(inMemory).not.toContain("sk-");
      // 标题取首条用户消息，是绕过 message.text 脱敏的第二条通道。
      expect(JSON.stringify(view.result.current.activeConversation?.title)).not.toContain(
        secret
      );
    } finally {
      view.unmount();
    }

    const raw = localStorage.getItem(CHAT_STORAGE_KEY) ?? "";
    // 正向锚点：存储里**确实**写了内容，而且消息还在、标题里还有"统计订单"。
    expect(raw.length).toBeGreaterThan(0);
    expect(raw).toContain("统计订单");
    expect(raw).toContain(REDACTED_PLACEHOLDER);
    // 关键负向断言：任何持久化文本里都不该有真凭据。
    expect(raw).not.toContain(secret);
    expect(raw).not.toContain("sk-");

    // 读回来不该被判成旧记录 —— 它从落盘起就是干净的。
    const loaded = loadChatHistory();
    expect(loaded.conversations.every((item) => !item.messages.length)).toBe(
      false
    );
    expect(isLegacyCredentialConversation(loaded, loaded.activeConversationId)).toBe(
      false
    );
    expect(loaded.conversations[0].title).toContain(REDACTED_PLACEHOLDER);
    expect(loaded.conversations[0].title).not.toContain(secret);
  });

  it("keeps a flagged legacy record flagged across a full save/reload cycle through the real hook", async () => {
    // 端到端：旧存储里的真密钥 → 首次加载脱敏 + 打标 → hook 落盘（标记写下去）
    // → 重开 → 仍然是"需新开会话"，且任何环节都不含真密钥。
    const secret = "sk-live-reopen-2d77";
    localStorage.setItem(
      CHAT_STORAGE_KEY,
      JSON.stringify({
        activeConversationId: "legacy",
        conversations: [
          {
            id: "legacy",
            title: `用 ${secret} 统计订单`,
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: `用 ${secret} 统计订单`,
                createdAt: "2026-09-27T00:00:00.000Z"
              }
            ]
          }
        ]
      })
    );

    // 第一次挂载：hook 读取 → 脱敏 → 落盘（带标记）。
    const first = mountLikeApp();
    expect(first.result.current.staleCredentials).toBe(true);
    await act(async () => {
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
    });
    const rawAfterFirstMount = localStorage.getItem(CHAT_STORAGE_KEY) ?? "";
    expect(rawAfterFirstMount).not.toContain(secret);
    first.unmount();

    // 第二次挂载（等价于重开任务窗格）：文本已干净，但标记让它仍被拦。
    const reopened = mountLikeApp();
    try {
      expect(reopened.result.current.staleCredentials).toBe(true);
      const active = reopened.result.current.activeConversation;
      expect(active.messages[0].text).not.toContain(secret);
      expect(active.title).not.toContain(secret);
      // 界面看到的是还原后的措辞。
      expect(active.title).toContain(RESTORED_SECRET_PLACEHOLDER);
    } finally {
      reopened.unmount();
    }
  });
});

/**
 * v3 → v4 迁移的凭据旁路。
 *
 * 旧实现把 `LEGACY_CHAT_STORAGE_KEY` 里的裸消息数组直接交给 `createConversation`
 * 返回：既不检测、不脱敏、不打标，也从不处理那个旧键。结果 v3 记录里的真密钥
 * 会原样渲染，且"迁移完成"名不副实 —— 原文一直留在存储里。
 */
describe("v3 legacy migration sanitizes, flags and cleans up", () => {
  function seedV3(messages: ChatHistoryState["conversations"][number]["messages"]) {
    localStorage.setItem(LEGACY_CHAT_STORAGE_KEY, JSON.stringify(messages));
  }

  const SECRET = "sk-live-in-v3-array-31ce";

  it("sanitizes v3 messages, flags the migrated conversation and removes the v3 key", () => {
    seedV3([
      {
        id: "old-1",
        role: "user",
        text: `帮我用 ${SECRET} 统计订单`,
        createdAt: "2026-09-26T00:00:00.000Z"
      },
      {
        id: "old-2",
        role: "system",
        text: `模型连接失败：Authorization: Bearer ${SECRET}`,
        createdAt: "2026-09-26T00:00:01.000Z"
      }
    ]);

    const migrated = loadChatHistory();
    const conversation = migrated.conversations[0];

    // 正向锚点：历史没有丢，非敏感内容还在，标题仍可读。
    expect(migrated.conversations).toHaveLength(1);
    expect(conversation.messages).toHaveLength(2);
    expect(conversation.messages[0].text).toContain("统计订单");
    expect(conversation.messages[1].text).toContain("模型连接失败");
    expect(conversation.title).toContain("统计订单");
    // 标题是绕过 message.text 脱敏的第二条通道 —— 也必须干净。
    expect(conversation.title).not.toContain(SECRET);
    expect(JSON.stringify(migrated)).not.toContain(SECRET);
    expect(JSON.stringify(migrated)).not.toContain("sk-");

    // 旧记录标记登记在**新会话 id** 上，发送闸据此拦截续聊。
    expect(
      isLegacyCredentialConversation(migrated, migrated.activeConversationId)
    ).toBe(true);

    // 安全的 v4 记录已经落盘，因此旧键被清理 —— 否则原文会一直留在存储里。
    const v4 = localStorage.getItem(CHAT_STORAGE_KEY) ?? "";
    expect(v4).toContain(conversation.id);
    expect(v4).not.toContain(SECRET);
    expect(v4).not.toContain("sk-");
    expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).toBeNull();
  });

  it("does not flag a clean v3 record and still clears the old key", () => {
    seedV3([
      {
        id: "old-1",
        role: "user",
        text: "按类别汇总一下销售额",
        createdAt: "2026-09-26T00:00:00.000Z"
      }
    ]);

    const migrated = loadChatHistory();
    expect(migrated.conversations[0].messages[0].text).toBe("按类别汇总一下销售额");
    expect(
      isLegacyCredentialConversation(migrated, migrated.activeConversationId)
    ).toBe(false);
    expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).toBeNull();
  });

  it("creates an empty conversation when the v3 key is absent or unusable", () => {
    const migrated = loadChatHistory();
    expect(migrated.conversations).toHaveLength(1);
    expect(migrated.conversations[0].messages).toHaveLength(1);
    expect(migrated.legacyCredentialConversationIds).toEqual([]);
  });

  it("keeps the render, the memory state and the persisted record free of the secret through the real hook", async () => {
    seedV3([
      {
        id: "old-1",
        role: "user",
        text: `用 ${SECRET} 统计订单`,
        createdAt: "2026-09-26T00:00:00.000Z"
      }
    ]);

    const first = mountLikeApp();
    try {
      // 首次渲染就是干净的，且发送闸已经关上。
      expect(first.result.current.staleCredentials).toBe(true);
      const active = first.result.current.activeConversation;
      expect(active.title).not.toContain(SECRET);
      expect(active.messages[0].text).not.toContain(SECRET);
      expect(active.messages[0].text).toContain(RESTORED_SECRET_PLACEHOLDER);
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).toBeNull();
    } finally {
      first.unmount();
    }

    const v4 = localStorage.getItem(CHAT_STORAGE_KEY) ?? "";
    expect(v4).not.toContain(SECRET);
    expect(v4).not.toContain("sk-");
    expect(v4).toContain("legacyCredentialConversationIds");

    // 重开：文本已经干净，但标记让发送闸仍然关着。
    const reopened = mountLikeApp();
    try {
      expect(reopened.result.current.staleCredentials).toBe(true);
    } finally {
      reopened.unmount();
    }
  });

  it("does not delete the v3 key when the safe v4 write fails, and still returns sanitized history", () => {
    seedV3([
      {
        id: "old-1",
        role: "user",
        text: `用 ${SECRET} 统计订单`,
        createdAt: "2026-09-26T00:00:00.000Z"
      }
    ]);

    // 只让 v4 的写入失败：旧键的读取/删除照常工作。
    const originalSetItem = Storage.prototype.setItem;
    const spy = vi
      .spyOn(Storage.prototype, "setItem")
      .mockImplementation(function (this: Storage, key: string, value: string) {
        if (key === CHAT_STORAGE_KEY) throw new Error("QuotaExceededError");
        originalSetItem.call(this, key, value);
      });
    try {
      const migrated = loadChatHistory();
      // 写入失败不能丢历史，也不能渲染原文。
      expect(migrated.conversations[0].messages[0].text).toContain("统计订单");
      expect(JSON.stringify(migrated)).not.toContain(SECRET);
      expect(
        isLegacyCredentialConversation(migrated, migrated.activeConversationId)
      ).toBe(true);
    } finally {
      spy.mockRestore();
    }

    // 旧键**保留**（那是用户仅剩的数据），但它的原文没有进入返回的内存态。
    expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).not.toBeNull();
  });
});

/**
 * v3 旧键删除失败。
 *
 * 迁移的次序（"先安全写 v4，再删 v3"）是对的，但 `removeItem` 失败此前被静默
 * 忽略 —— 旧键里仍是**原文**，而代码却当作迁移已经收尾。下面这组用例专门钉：
 * 1. 删不掉不影响任何安全属性（v4 干净、标记在、首渲染干净、发送闸不松）；
 * 2. 清理**只碰那一个键**，不牵连 v4 与用户的其他存储；
 * 3. 删不掉时要**如实报告"待清理"**，不能宣称存储已干净；
 * 4. 存储恢复后能真的清掉，v4 历史不丢。
 */
describe("v3 legacy key that refuses to be deleted", () => {
  const SECRET = "sk-live-v3-undeletable-6be2";

  function seedV3(messages: ChatHistoryState["conversations"][number]["messages"]) {
    localStorage.setItem(LEGACY_CHAT_STORAGE_KEY, JSON.stringify(messages));
  }

  function seedV3WithSecret() {
    seedV3([
      {
        id: "old-1",
        role: "user",
        text: `用 ${SECRET} 统计订单`,
        createdAt: "2026-09-26T00:00:00.000Z"
      }
    ]);
  }

  /**
   * 只让**确切的 v3 键**删除失败：其余键照常删。
   *
   * 这样既能模拟"宿主对这个键只读"，又能同时验证"实现没有顺手 clear 别的键"。
   */
  function blockLegacyKeyRemoval() {
    const originalRemoveItem = Storage.prototype.removeItem;
    const removed: string[] = [];
    const spy = vi
      .spyOn(Storage.prototype, "removeItem")
      .mockImplementation(function (this: Storage, key: string) {
        removed.push(key);
        if (key === LEGACY_CHAT_STORAGE_KEY) {
          throw new Error("SecurityError: 该存储项不可删除");
        }
        originalRemoveItem.call(this, key);
      });
    return { spy, removed };
  }

  it("keeps the safe record, the flag and the gate while reporting the old key as pending", () => {
    seedV3WithSecret();
    // 用户的其他数据：清理**绝不能**连带删掉这些。
    localStorage.setItem(MODEL_STORAGE_KEY, "connection:first");

    const { spy } = blockLegacyKeyRemoval();
    let migrated: ChatHistoryState;
    try {
      migrated = loadChatHistory();
    } finally {
      spy.mockRestore();
    }

    // 1) 安全属性一条没丢：v4 干净、历史完整、标记在。
    const v4Raw = localStorage.getItem(CHAT_STORAGE_KEY) ?? "";
    expect(v4Raw.length).toBeGreaterThan(0);
    expect(v4Raw).not.toContain(SECRET);
    expect(v4Raw).not.toContain("sk-");
    expect(v4Raw).toContain("统计订单");
    expect(migrated.conversations[0].messages[0].text).toContain("统计订单");
    expect(JSON.stringify(migrated)).not.toContain(SECRET);
    expect(
      isLegacyCredentialConversation(migrated, migrated.activeConversationId)
    ).toBe(true);

    // 2) 如实报告"待清理" —— 不能因为 v4 写成功就宣称存储已无残留。
    expect(migrated.legacyMigrationCleanupPending).toBe(true);
    expect(isLegacyV3KeyPendingCleanup()).toBe(true);

    // 3) 用户的其他存储没被动过。
    expect(localStorage.getItem(MODEL_STORAGE_KEY)).toBe("connection:first");
    // 4) 第一次渲染仍然干净。
    const rendered = renderRestoredSecrets(migrated);
    expect(JSON.stringify(rendered)).not.toContain(SECRET);
    expect(rendered.conversations[0].messages[0].text).toContain(
      RESTORED_SECRET_PLACEHOLDER
    );

    // 5) 旧键仍在，且**原文确实还躺在存储里** —— 这正是不能声称"已清理"的原因。
    expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).not.toBeNull();
    expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY) ?? "").toContain(SECRET);
  });

  it("only touches the exact v3 key when cleaning up", () => {
    seedV3WithSecret();
    localStorage.setItem(MODEL_STORAGE_KEY, "connection:first");
    localStorage.setItem("excel-bro.chat.v4.other", "别动我");
    const v4Before = localStorage.getItem(CHAT_STORAGE_KEY);

    const { removed } = blockLegacyKeyRemoval();
    try {
      // 删除失败 → 返回 false，且**只尝试过这一个键**。
      expect(removeLegacyV3Key()).toBe(false);
    } finally {
      // 恢复真实实现后再看"没被误删"的证据。
      vi.restoreAllMocks();
    }

    expect(removed).toContain(LEGACY_CHAT_STORAGE_KEY);
    // 关键：清理只针对那一个键，没有遍历、没有 clear。
    expect([...new Set(removed)]).toEqual([LEGACY_CHAT_STORAGE_KEY]);
    // v4 与用户的其他存储原样保留。
    expect(localStorage.getItem(MODEL_STORAGE_KEY)).toBe("connection:first");
    expect(localStorage.getItem("excel-bro.chat.v4.other")).toBe("别动我");
    expect(localStorage.getItem(CHAT_STORAGE_KEY)).toBe(v4Before);
    // 旧键也还在（失败就是失败，不假装删掉了）。
    expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).not.toBeNull();
  });

  it("keeps the gate on and the notice up through the real hook, then clears the key once storage recovers", async () => {
    seedV3WithSecret();

    // 第一次挂载：删除失败。界面要如实提示"待清理"，发送闸不松，v4 历史不丢。
    const first = blockLegacyKeyRemoval();
    let firstView: RenderHookResult<ReturnType<typeof useConversation>, unknown> | null = null;
    try {
      firstView = mountLikeApp();
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });
      expect(firstView.result.current.staleCredentials).toBe(true);
      expect(firstView.result.current.legacyCleanupNotice).toBe(true);
      // 界面上的文本不含真密钥。
      const active = firstView.result.current.activeConversation;
      expect(active.messages[0].text).not.toContain(SECRET);
      expect(active.messages[0].text).toContain(RESTORED_SECRET_PLACEHOLDER);
      expect(active.title).not.toContain(SECRET);
    } finally {
      first.spy.mockRestore();
    }
    expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).not.toBeNull();
    const v4Raw = localStorage.getItem(CHAT_STORAGE_KEY) ?? "";
    expect(v4Raw).not.toContain(SECRET);
    // "待清理"这件事本身也写进了存储，所以重开后仍会提示，而不是谎报干净。
    expect(v4Raw).toContain("legacyMigrationCleanupPending");
    firstView?.unmount();

    // 第二次挂载：存储已恢复。旧键这次真的被清掉，v4 历史不丢。
    const reopened = mountLikeApp();
    try {
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });
      expect(reopened.result.current.legacyCleanupNotice).toBe(false);
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).toBeNull();
      // v4 仍在，且标记（拦续聊）没有被"清理成功"顺手抹掉。
      const afterRaw = localStorage.getItem(CHAT_STORAGE_KEY) ?? "";
      expect(afterRaw).toContain("统计订单");
      expect(afterRaw).not.toContain(SECRET);
      expect(reopened.result.current.staleCredentials).toBe(true);
      expect(
        reopened.result.current.activeConversation.messages[0].text
      ).toContain(RESTORED_SECRET_PLACEHOLDER);
    } finally {
      reopened.unmount();
    }

    // 重开一次确认收敛：提示不会回来，旧键不会被无中生有。
    const settled = mountLikeApp();
    try {
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });
      expect(settled.result.current.legacyCleanupNotice).toBe(false);
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).toBeNull();
    } finally {
      settled.unmount();
    }
  });

  it("reports a clear pending state when the removal is blocked but never claims the machine storage is clean", () => {
    seedV3WithSecret();

    // 第一次读取（v3 → v4 迁移）：删除被拒 → 落盘的 v4 必须带上"待清理"。
    const first = blockLegacyKeyRemoval();
    try {
      const migrated = loadChatHistory();
      // 报告口径：**"旧键待清理"**，而不是"已无凭据"。
      expect(migrated.legacyMigrationCleanupPending).toBe(true);
      // 同时仍然完整履行脱敏与拦截职责 —— 两件事互不冲突。
      expect(JSON.stringify(migrated)).not.toContain(SECRET);
      expect(
        isLegacyCredentialConversation(migrated, migrated.activeConversationId)
      ).toBe(true);
    } finally {
      first.spy.mockRestore();
    }

    // 重开（这次走 v4 分支，且旧键**仍然删不掉**）：标记读回来仍在，
    // 否则重开后界面拿不到它，等价于谎报"存储已干净"。
    const second = blockLegacyKeyRemoval();
    try {
      const reopenedAgain = loadChatHistory();
      expect(reopenedAgain.legacyMigrationCleanupPending).toBe(true);
      expect(JSON.stringify(reopenedAgain)).not.toContain(SECRET);
    } finally {
      second.spy.mockRestore();
    }

    expect(isLegacyV3KeyPendingCleanup()).toBe(true);
  });

  it("keeps reporting pending when the storage is entirely unreadable", () => {
    // 存储整体读不到时无法判定，宁可保守报"可能还有残留"。
    const originalGetItem = Storage.prototype.getItem;
    const spy = vi
      .spyOn(Storage.prototype, "getItem")
      .mockImplementation(function (this: Storage, key: string) {
        if (key === LEGACY_CHAT_STORAGE_KEY) {
          throw new Error("SecurityError: 拒绝访问");
        }
        return originalGetItem.call(this, key);
      });
    try {
      expect(isLegacyV3KeyPendingCleanup()).toBe(true);
      expect(removeLegacyV3Key()).toBe(false);
    } finally {
      spy.mockRestore();
    }
  });
});

/**
 * v4 写入**持续失败**时的挂载安全性（真实 renderHook 故障注入）。
 *
 * 这一组补的是上一组留下的缺口：上面所有用例要么只走 `loadChatHistory()`，
 * 要么只让 `removeItem` 失败。**没有一条**覆盖"v4 写不进、v3 删得掉"这个组合
 * 走完真实 hook 的情形 —— 而那正是唯一会**丢数据**的组合。
 *
 * 缺陷形状（修复前）：`useConversation` 的持久化 effect 把清理排在落盘**之前**
 * （`removeStaleLegacyKey()` → `setItem(v4)`）。v3 键是用户仅剩的历史载体，
 * 一旦先删掉，紧接着的 v4 写入又持续失败，存储里就**什么都不剩** ——
 * 历史被永久清空，比"残留原文"严重得多。
 *
 * 下面的故障注入同时表达两件事：
 * - `setItem(CHAT_STORAGE_KEY)` **持续**抛错（迁移内那次、以及挂载后 effect 里那次）；
 * - `removeItem(LEGACY_CHAT_STORAGE_KEY)` **允许成功**。
 *
 * 修复后的正确行为：写入是清理的**前置条件**，不是并列步骤。
 */
describe("v4 write that keeps failing must not cost the v3 record", () => {
  const SECRET = "sk-live-v4-write-fails-41d7";

  function seedV3WithSecret() {
    localStorage.setItem(
      LEGACY_CHAT_STORAGE_KEY,
      JSON.stringify([
        {
          id: "old-v4w",
          role: "user",
          text: `帮我用 ${SECRET} 汇总这张表`,
          createdAt: "2026-09-26T00:00:00.000Z"
        }
      ])
    );
  }

  /**
   * v4 写入持续失败，但 v3 键**允许**删除。
   *
   * 与上一组的 `blockLegacyKeyRemoval()` 恰好互补：上一组是"删不掉、写得进"，
   * 这一组是"写得进不了、删得掉"。两组合起来才覆盖双向的失败组合。
   */
  function blockV4WritesOnly() {
    const originalSetItem = Storage.prototype.setItem;
    const v4Writes: number[] = [];
    const removed: string[] = [];
    const originalRemoveItem = Storage.prototype.removeItem;
    const setSpy = vi
      .spyOn(Storage.prototype, "setItem")
      .mockImplementation(function (this: Storage, key: string, value: string) {
        if (key === CHAT_STORAGE_KEY) {
          v4Writes.push(value.length);
          throw new Error("QuotaExceededError: 空间已满");
        }
        originalSetItem.call(this, key, value);
      });
    const removeSpy = vi
      .spyOn(Storage.prototype, "removeItem")
      .mockImplementation(function (this: Storage, key: string) {
        removed.push(key);
        originalRemoveItem.call(this, key);
      });
    return {
      v4Writes,
      removed,
      restore: () => {
        setSpy.mockRestore();
        removeSpy.mockRestore();
      }
    };
  }

  /**
   * 记录 `removeItem` 被以哪些键调用过（不抛错）。
   *
   * 用于直接盯"写入未确认前一次都不许删"这条不变量：`removed` 为空数组就是证据。
   */
  function blockRemovalAttempt() {
    const originalRemoveItem = Storage.prototype.removeItem;
    const removed: string[] = [];
    const spy = vi
      .spyOn(Storage.prototype, "removeItem")
      .mockImplementation(function (this: Storage, key: string) {
        removed.push(key);
        originalRemoveItem.call(this, key);
      });
    return { removed, restore: () => spy.mockRestore() };
  }

  it("deletes nothing, keeps the gate shut, and never claims the v4 record is safe", async () => {
    seedV3WithSecret();
    localStorage.setItem(MODEL_STORAGE_KEY, "connection:first");

    const fault = blockV4WritesOnly();
    let view: RenderHookResult<ReturnType<typeof useConversation>, unknown> | null = null;
    try {
      view = mountLikeApp();
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });

      // 正向锚点 1：确实尝试过写 v4（迁移内 + effect 里），否则下面的断言是空过。
      expect(fault.v4Writes.length).toBeGreaterThan(0);
      // 正向锚点 2：迁移确实读到了那条 v3 消息（内容进入内存态）。
      const active = view.result.current.activeConversation;
      const hasSummaryText = (messages: { text?: string }[]): boolean =>
        messages.some((message) => (message.text ?? "").includes("汇总这张表"));
      expect(hasSummaryText(active.messages)).toBe(true);

      // 1) **核心断言**：v4 没写成功，v3 键就绝不能被删 —— 它是用户仅剩的数据。
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).not.toBeNull();
      expect(
        localStorage.getItem(LEGACY_CHAT_STORAGE_KEY) ?? ""
      ).toContain("汇总这张表");

      // 2) 界面与内存态都不含真密钥。
      const json = JSON.stringify(view.result.current.activeConversation);
      expect(json).not.toContain(SECRET);
      expect(active.messages.every((message) => !(message.text ?? "").includes(SECRET))).toBe(true);
      expect(active.title).not.toContain(SECRET);
      expect(JSON.stringify(view.result.current.chatHistory)).not.toContain(SECRET);
      // 脱敏后的措辞（占位符已还原成人话）应当出现。
      expect(JSON.stringify(view.result.current.chatHistory)).toContain(
        RESTORED_SECRET_PLACEHOLDER
      );

      // 3) 发送闸仍然关闭：旧会话不许续聊。
      expect(view.result.current.staleCredentials).toBe(true);

      // 4) 提示不得谎称"v4 已安全保存"。旧键还在，就必须如实说待清理。
      expect(view.result.current.legacyCleanupNotice).toBe(true);
    } finally {
      view?.unmount();
      fault.restore();
    }

    // 5) 其他用户数据在尝试过程中没有被牵连删除。
    expect(localStorage.getItem(MODEL_STORAGE_KEY)).toBe("connection:first");
    expect(fault.removed).not.toContain(MODEL_STORAGE_KEY);
  });

  it("never touches the v3 key before the v4 write is confirmed", () => {
    seedV3WithSecret();

    // v4 **未**确认落盘 → 清理函数必须一次都不调用 removeItem。
    const beforeWrite = blockRemovalAttempt();
    try {
      expect(maybeCleanupLegacyKeyAfterWrite(false)).toBe(false);
      expect(beforeWrite.removed).toEqual([]);
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).not.toBeNull();
    } finally {
      beforeWrite.restore();
    }

    // v4 确认落盘 → 才允许清理，且只碰那一个键。
    const afterWrite = blockRemovalAttempt();
    try {
      expect(maybeCleanupLegacyKeyAfterWrite(true)).toBe(true);
      expect(afterWrite.removed).toEqual([LEGACY_CHAT_STORAGE_KEY]);
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).toBeNull();
    } finally {
      afterWrite.restore();
    }
  });

  it("recovers: the v3 key is only cleared after a v4 write finally succeeds", async () => {
    seedV3WithSecret();

    // 阶段一：v4 持续写失败。旧键必须原样留着。
    const fault = blockV4WritesOnly();
    let first: RenderHookResult<ReturnType<typeof useConversation>, unknown> | null = null;
    try {
      first = mountLikeApp();
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });
      expect(first.result.current.legacyCleanupNotice).toBe(true);
      expect(first.result.current.staleCredentials).toBe(true);
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).not.toBeNull();
      // v4 一个字节都没落盘。
      expect(localStorage.getItem(CHAT_STORAGE_KEY)).toBeNull();
    } finally {
      first?.unmount();
      fault.restore();
    }

    // 阶段二：存储恢复（不再注入故障）。重开后应当：写出安全的 v4 → 才清掉 v3。
    const reopened = mountLikeApp();
    try {
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });
      const v4Raw = localStorage.getItem(CHAT_STORAGE_KEY) ?? "";
      // v4 先存在，而且是干净的。
      expect(v4Raw.length).toBeGreaterThan(0);
      expect(v4Raw).not.toContain(SECRET);
      expect(v4Raw).toContain("汇总这张表");
      // 然后 v3 才被清掉 —— 这是"写成功"授权的，不是无条件删的。
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).toBeNull();
      // 历史没丢、闸没松，提示收敛。
      expect(reopened.result.current.legacyCleanupNotice).toBe(false);
      expect(reopened.result.current.staleCredentials).toBe(true);
      const reopenedActive = reopened.result.current.activeConversation;
      const reopenedHasSummary = reopenedActive.messages.some((message) =>
        (message.text ?? "").includes("汇总这张表")
      );
      expect(reopenedHasSummary).toBe(true);
    } finally {
      reopened.unmount();
    }
  });
});

/**
 * `main.tsx` 用 `React.StrictMode` 包裹 `<App />`，开发构建下 React **有意**
 * 重放渲染与 effect（挂载 → 卸载 → 再挂载）。
 *
 * 因此"源码里只有一处 `loadChatHistory()` 调用"**不等于**"运行时只调用一次" ——
 * StrictMode 下初始化逻辑会跑两遍。把"调用次数 == 1"写进断言是**错的正确性标准**：
 * 它要么在 StrictMode 下假红，要么诱使实现去抵消 React 的刻意行为。
 *
 * 这组用例因此只盯**结果不变量**：无论重放几次，都不能丢历史、不能泄漏密钥、
 * 不能在写入确认前删掉 v3、不能放开发送闸。幂等性是**结果**属性，不是次数属性。
 */
describe("React.StrictMode replay must not break migration idempotence", () => {
  const SECRET = "sk-live-strictmode-5ae2";

  // 这个文件是 `.ts`（套件里其余几百条都是纯逻辑，不引入 JSX 转换），
  // 所以用 `createElement` 而不是 JSX 语法来包 StrictMode。
  function mountStrict() {
    return renderHook(() => useConversation(), {
      wrapper: ({ children }) => createElement(StrictMode, null, children)
    });
  }

  it("keeps the history, the gate and the v3 key intact across the double-mount", async () => {
    localStorage.setItem(
      LEGACY_CHAT_STORAGE_KEY,
      JSON.stringify([
        {
          id: "legacy-strict",
          role: "user",
          text: `帮我用 ${SECRET} 汇总这张表`,
          createdAt: "2026-09-26T00:00:00.000Z"
        }
      ])
    );

    const view = mountStrict();
    try {
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });

      // 历史没丢：v3 那条消息进入了内存态（重放两次也不会被抹掉）。
      const active = view.result.current.activeConversation;
      expect(
        active.messages.some((message) => (message.text ?? "").includes("汇总这张表"))
      ).toBe(true);

      // 不泄漏密钥：界面 / 内存态都不含原文，只有还原后的措辞。
      const historyJson = JSON.stringify(view.result.current.chatHistory);
      expect(historyJson).not.toContain(SECRET);
      expect(historyJson).toContain(RESTORED_SECRET_PLACEHOLDER);
      expect(active.title).not.toContain(SECRET);

      // 落盘内容同样不含密钥。
      expect(localStorage.getItem(CHAT_STORAGE_KEY) ?? "").not.toContain(SECRET);

      // 发送闸没松。
      expect(view.result.current.staleCredentials).toBe(true);

      // 正常路径下 v4 写成功后才清 v3 —— 重放不会把顺序颠倒成"先删后写"。
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).toBeNull();
      expect(localStorage.getItem(CHAT_STORAGE_KEY)).not.toBeNull();
    } finally {
      view.unmount();
    }
  });

  it("never deletes the v3 key during the replay when the v4 write is failing", async () => {
    localStorage.setItem(
      LEGACY_CHAT_STORAGE_KEY,
      JSON.stringify([
        {
          id: "legacy-strict-fail",
          role: "user",
          text: `用 ${SECRET} 汇总这张表`,
          createdAt: "2026-09-26T00:00:00.000Z"
        }
      ])
    );

    const originalSetItem = Storage.prototype.setItem;
    const setSpy = vi
      .spyOn(Storage.prototype, "setItem")
      .mockImplementation(function (this: Storage, key: string, value: string) {
        if (key === CHAT_STORAGE_KEY) {
          throw new Error(`写入被拒绝 ${SECRET}`);
        }
        originalSetItem.call(this, key, value);
      });

    const view = mountStrict();
    try {
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });

      // StrictMode 下挂了两次，但 v3 键在这两次里都不能被删 —— 它是仅剩的数据。
      expect(localStorage.getItem(LEGACY_CHAT_STORAGE_KEY)).not.toBeNull();
      expect(localStorage.getItem(CHAT_STORAGE_KEY)).toBeNull();

      // 不丢历史、不泄漏、闸不松。
      const active = view.result.current.activeConversation;
      expect(
        active.messages.some((message) => (message.text ?? "").includes("汇总这张表"))
      ).toBe(true);
      expect(JSON.stringify(view.result.current.chatHistory)).not.toContain(SECRET);
      expect(view.result.current.staleCredentials).toBe(true);
      // 如实报告"保存失败"，不谎称"v4 已保存、仅待清理"。
      expect(view.result.current.historyStorageStatus).toBe(
        "save-failed-legacy-present"
      );
    } finally {
      view.unmount();
      setSpy.mockRestore();
    }
  });
});
