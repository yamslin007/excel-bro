import { describe, expect, it } from "vitest";
import {
  FOCUS_HEARTBEAT_INTERVAL_MS,
  FOCUS_PROTOCOL_CHANNEL,
  FOCUS_PROTOCOL_VERSION,
  FOCUS_SNAPSHOT_LIMITS,
  FOCUS_SYNC_TTL_MS,
  createFocusEnvelope,
  createFocusSessionId,
  focusCacheKey,
  initialFocusChildState,
  nextFocusSeq,
  parseFocusEnvelope,
  readFocusMessage,
  reduceFocusChild,
  sanitizeFocusSnapshot,
  type FocusSnapshot
} from "./focusProtocol";

const SESSION = "0123456789abcdef0123456789abcdef";
const WORKBOOK = { name: "预算.xlsx", identity: "wps-path:c:\\data\\预算.xlsx" };

function snapshot(overrides: Partial<FocusSnapshot> = {}): FocusSnapshot {
  return {
    initialView: "conversation",
    activeConversationId: "c1",
    conversations: [
      {
        id: "c1",
        title: "整理销售表",
        updatedAt: "2026-11-03T08:00:00.000Z",
        messages: [
          { id: "m1", role: "user", text: "把空行删掉", createdAt: "2026-11-03T08:00:00.000Z" },
          {
            id: "m2",
            role: "assistant",
            text: "已完成",
            createdAt: "2026-11-03T08:00:10.000Z",
            plan: { title: "清理", summary: "删除空行", steps: ["读取", "写入"] },
            result: { title: "结果", headers: ["A"], rows: [[1], [2]] }
          }
        ]
      }
    ],
    tools: [
      {
        id: "t1",
        kind: "workflow",
        name: "月度汇总",
        description: "按列汇总",
        category: "工作流程",
        steps: 2,
        stepLabels: ["读取", "写入"]
      }
    ],
    workbook: WORKBOOK,
    ...overrides
  };
}

function envelope(kind: Parameters<typeof createFocusEnvelope>[0]["kind"], extra: Partial<Parameters<typeof createFocusEnvelope>[0]> = {}) {
  return createFocusEnvelope({
    kind,
    sessionId: SESSION,
    seq: 1,
    sentAt: 1_000,
    workbook: WORKBOOK,
    ...(kind === "snapshot" ? { snapshot: snapshot() } : {}),
    ...extra
  });
}

describe("focus protocol envelope", () => {
  it("builds one shape for both hosts and round-trips through the parser", () => {
    const message = envelope("snapshot");
    expect(message.channel).toBe(FOCUS_PROTOCOL_CHANNEL);
    expect(message.version).toBe(FOCUS_PROTOCOL_VERSION);
    const parsed = parseFocusEnvelope(JSON.parse(JSON.stringify(message)));
    expect(parsed).toEqual(message);
  });

  it("refuses messages from another channel, version or kind", () => {
    expect(parseFocusEnvelope({ ...envelope("hello"), channel: "excel-bro.other" })).toBeNull();
    expect(parseFocusEnvelope({ ...envelope("hello"), version: 99 })).toBeNull();
    expect(parseFocusEnvelope({ ...envelope("hello"), kind: "exec-plan" })).toBeNull();
    expect(parseFocusEnvelope({ ...envelope("hello"), sessionId: "" })).toBeNull();
    expect(parseFocusEnvelope(null)).toBeNull();
    expect(parseFocusEnvelope("focus")).toBeNull();
  });

  it("allocates monotonic sequence numbers", () => {
    expect(nextFocusSeq(0)).toBe(1);
    expect(nextFocusSeq(41)).toBe(42);
    expect(nextFocusSeq(Number.NaN)).toBe(1);
  });

  it("creates unpredictable session ids and refuses when the host has no crypto", () => {
    const first = createFocusSessionId();
    const second = createFocusSessionId();
    expect(first).toMatch(/^[0-9a-f]{32}$/);
    expect(first).not.toBe(second);
    expect(createFocusSessionId(null)).toBeNull();
    expect(createFocusSessionId({} as never)).toBeNull();
    expect(focusCacheKey(SESSION)).toContain(SESSION);
  });
});

describe("focus protocol validation", () => {
  const base = { sessionId: SESSION, workbookIdentity: WORKBOOK.identity, now: 1_000, maxAgeMs: 10_000 };

  it("accepts a matching snapshot", () => {
    const result = readFocusMessage(JSON.stringify(envelope("snapshot")), base);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.message.snapshot?.conversations[0].id).toBe("c1");
  });

  it("rejects malformed json and foreign channels", () => {
    expect(readFocusMessage("{", base)).toMatchObject({ ok: false, reason: "not-json" });
    expect(readFocusMessage(JSON.stringify({ ...envelope("hello"), channel: "x" }), base))
      .toMatchObject({ ok: false, reason: "wrong-channel" });
  });

  it("rejects an old session, a stale sequence and an expired message", () => {
    expect(readFocusMessage(envelope("snapshot"), { ...base, sessionId: "another" }))
      .toMatchObject({ ok: false, reason: "session-mismatch" });
    expect(readFocusMessage(envelope("snapshot", { seq: 3 }), { ...base, minSeq: 3 }))
      .toMatchObject({ ok: false, reason: "stale-seq" });
    expect(readFocusMessage(envelope("snapshot", { sentAt: 1 }), { ...base, now: 20_000 }))
      .toMatchObject({ ok: false, reason: "expired" });
  });

  it("rejects a message that belongs to another workbook", () => {
    const foreign = createFocusEnvelope({
      kind: "snapshot",
      sessionId: SESSION,
      seq: 2,
      sentAt: 1_000,
      workbook: { name: "别的.xlsx", identity: "wps-path:c:\\data\\别的.xlsx" },
      snapshot: snapshot({ workbook: { name: "别的.xlsx", identity: "wps-path:c:\\data\\别的.xlsx" } })
    });
    expect(readFocusMessage(foreign, base)).toMatchObject({ ok: false, reason: "workbook-mismatch" });
  });

  it("does not pretend to know a workbook identity it cannot read", () => {
    const unsaved = createFocusEnvelope({
      kind: "snapshot",
      sessionId: SESSION,
      seq: 1,
      sentAt: 1_000,
      workbook: { name: "工作簿1", identity: null },
      snapshot: snapshot({ workbook: { name: "工作簿1", identity: null } })
    });
    expect(readFocusMessage(unsaved, { ...base, workbookIdentity: null }).ok).toBe(true);
  });

  it("requires a snapshot body for snapshot messages", () => {
    const { snapshot: _ignored, ...withoutSnapshot } = envelope("snapshot");
    expect(readFocusMessage(withoutSnapshot, base))
      .toMatchObject({ ok: false, reason: "missing-snapshot" });
  });

  it("gives every rejection a human readable reason", () => {
    const result = readFocusMessage(envelope("hello"), { ...base, sessionId: "other" });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.detail).toContain("另一个专注窗口会话");
  });
});

describe("focus snapshot sanitising", () => {
  it("keeps only display fields and clamps oversized history", () => {
    const oversized = snapshot({
      conversations: [
        {
          id: "c1",
          title: "长对话",
          updatedAt: "2026-11-03T08:00:00.000Z",
          messages: Array.from({ length: FOCUS_SNAPSHOT_LIMITS.messagesPerConversation + 20 }, (_, index) => ({
            id: `m${index}`,
            role: "assistant" as const,
            text: `第${index}条`,
            createdAt: "2026-11-03T08:00:00.000Z",
            result: {
              title: "结果",
              headers: ["A"],
              rows: Array.from({ length: FOCUS_SNAPSHOT_LIMITS.resultRows + 5 }, () => [1])
            }
          }))
        }
      ]
    });
    const clean = sanitizeFocusSnapshot(JSON.parse(JSON.stringify(oversized)))!;
    expect(clean.conversations[0].messages).toHaveLength(FOCUS_SNAPSHOT_LIMITS.messagesPerConversation);
    expect(clean.conversations[0].messages[0].result?.rows).toHaveLength(FOCUS_SNAPSHOT_LIMITS.resultRows);
    // 只留展示字段：没有任何单元格原始数据之外的附加键被带进来。
    expect(Object.keys(clean.conversations[0].messages[0]).sort()).toEqual(
      ["createdAt", "id", "result", "role", "text"]
    );
  });

  it("drops entries without an id instead of inventing one", () => {
    const clean = sanitizeFocusSnapshot({
      initialView: "tools",
      activeConversationId: "c1",
      conversations: [{ title: "没有 id" }, { id: "c1", title: "有 id", messages: [] }],
      tools: [{ name: "没有 id" }],
      workbook: { name: "预算.xlsx", identity: null }
    })!;
    expect(clean.conversations.map((item) => item.id)).toEqual(["c1"]);
    expect(clean.tools).toEqual([]);
  });

  it("rejects shapes that are not a snapshot at all", () => {
    expect(sanitizeFocusSnapshot(null)).toBeNull();
    expect(sanitizeFocusSnapshot({ conversations: [], tools: {} })).toBeNull();
  });
});

describe("focus child state machine", () => {
  it("only becomes ready after a validated snapshot arrives", () => {
    let state = initialFocusChildState();
    expect(state.status).toBe("connecting");
    state = reduceFocusChild(state, { type: "transport", channel: "BroadcastChannel" });
    expect(state.status).toBe("connecting");
    state = reduceFocusChild(state, { type: "snapshot", snapshot: snapshot(), seq: 1, at: 1_000 });
    expect(state.status).toBe("ready");
    expect(state.lastSyncAt).toBe(1_000);
    expect(state.channel).toBe("BroadcastChannel");
  });

  it("marks the content stale when heartbeats stop, and recovers on the next sync", () => {
    let state = reduceFocusChild(initialFocusChildState(), {
      type: "snapshot",
      snapshot: snapshot(),
      seq: 1,
      at: 1_000
    });
    state = reduceFocusChild(state, { type: "tick", at: 1_000 + FOCUS_SYNC_TTL_MS });
    expect(state.status).toBe("ready");
    state = reduceFocusChild(state, { type: "tick", at: 1_000 + FOCUS_SYNC_TTL_MS + 1 });
    expect(state.status).toBe("stale");
    expect(state.snapshot).not.toBeNull();
    expect(state.reason).toContain("已过期");
    state = reduceFocusChild(state, { type: "heartbeat", seq: 2, at: 1_000 + FOCUS_SYNC_TTL_MS + 500 });
    expect(state.status).toBe("ready");
    expect(state.reason).toBeNull();
  });

  it("keeps heartbeats moving the clock but never fakes content", () => {
    const state = reduceFocusChild(initialFocusChildState(), {
      type: "heartbeat",
      seq: 4,
      at: 2_000
    });
    expect(state.status).toBe("connecting");
    expect(state.snapshot).toBeNull();
    expect(state.lastSyncAt).toBe(2_000);
  });

  it("ignores repeated sequence numbers", () => {
    let state = reduceFocusChild(initialFocusChildState(), {
      type: "snapshot",
      snapshot: snapshot(),
      seq: 5,
      at: 1_000
    });
    const next = reduceFocusChild(state, {
      type: "snapshot",
      snapshot: snapshot({ activeConversationId: "c2" }),
      seq: 5,
      at: 1_500
    });
    expect(next).toBe(state);
  });

  it("treats a parent close as terminal", () => {
    let state = reduceFocusChild(initialFocusChildState(), {
      type: "snapshot",
      snapshot: snapshot(),
      seq: 1,
      at: 1_000
    });
    state = reduceFocusChild(state, { type: "closing", reason: "任务窗格已关闭" });
    expect(state.status).toBe("closed");
    expect(state.reason).toBe("任务窗格已关闭");
    const after = reduceFocusChild(state, { type: "snapshot", snapshot: snapshot(), seq: 9, at: 2_000 });
    expect(after.status).toBe("closed");
  });

  it("reports a failed handshake instead of showing an empty but plausible window", () => {
    const failed = reduceFocusChild(initialFocusChildState(), { type: "handshake-timeout" });
    expect(failed.status).toBe("failed");
    expect(failed.reason).toContain("握手");
    const unavailable = reduceFocusChild(initialFocusChildState(), {
      type: "transport-unavailable",
      reason: "宿主不支持跨窗口通道"
    });
    expect(unavailable.status).toBe("failed");
  });

  it("marks cache-only content as a snapshot, not as the live scene", () => {
    const state = reduceFocusChild(initialFocusChildState(), {
      type: "cache-only",
      snapshot: snapshot(),
      reason: "双向通道不可用"
    });
    expect(state.status).toBe("snapshot");
    expect(state.reason).toContain("双向通道不可用");
    // 只靠缓存不能升级成 ready；必须有真实快照消息。
    expect(reduceFocusChild(state, { type: "tick", at: 1_000_000 }).status).toBe("snapshot");
  });

  it("keeps heartbeat cadence well inside the staleness window", () => {
    expect(FOCUS_SYNC_TTL_MS).toBeGreaterThanOrEqual(FOCUS_HEARTBEAT_INTERVAL_MS * 2);
  });
});
