import { describe, expect, it } from "vitest";
import {
  applyBridgePollResponse,
  applyBridgeProgressEvent,
  assertBridgePollResponse,
  bridgeToolPointer,
  bridgeWorkbookIdentity,
  bridgeRequestedSheetNames,
  verifyBridgeAuthorization,
  cancelBridge,
  createSdkBridgeState,
  decideBridgeReceipt,
  markBridgeDelivered,
  pendingBridgeCall,
  runBridgeUntilIdle,
  serviceBridgeOnce,
  shouldPollBridge,
  type BridgeHostPorts,
  type SdkBridgeState,
  type SdkBridgeBinding
} from "./sdkToolBridge";

const BINDING: SdkBridgeBinding = {
  taskId: "turn-bridge-0001",
  hostId: "excel-office",
  workbookName: "测试.xlsx",
  workbookIdentity: "url:https://contoso.example/测试.xlsx",
  sourceFingerprint: "fp-source-1",
  authorizedSheetNames: ["Sheet1", "Sheet2"]
};

const OP = "op_" + "1".repeat(32);
const FP = "a".repeat(64);

function call(overrides: Partial<Record<string, unknown>> = {}) {
  return {
    operationId: OP,
    tool: "query_table" as const,
    arguments: { mode: "aggregate" },
    requestFingerprint: FP,
    expiresAt: 1000,
    ...overrides
  };
}

describe("bridgeToolPointer", () => {
  it("accepts a well-formed opaque pointer", () => {
    expect(bridgeToolPointer({ operationId: OP, tool: "query_table" })).toEqual({
      operationId: OP,
      tool: "query_table"
    });
  });

  it("rejects malformed pointers instead of guessing", () => {
    expect(bridgeToolPointer(null)).toBeNull();
    expect(bridgeToolPointer({ operationId: "op_short", tool: "query_table" })).toBeNull();
    expect(bridgeToolPointer({ operationId: OP, tool: "read_range" })).toBeNull();
  });
});

describe("progress events", () => {
  it("requests a poll when a bridge pointer arrives", () => {
    const state = applyBridgeProgressEvent(createSdkBridgeState(BINDING), {
      stepId: "bridge.query_table",
      status: "started",
      bridgeTool: { operationId: OP, tool: "query_table" }
    });
    expect(state.pollRequested).toBe(true);
    expect(shouldPollBridge(state)).toBe(true);
  });

  it("ignores unrelated events", () => {
    const base = createSdkBridgeState(BINDING);
    expect(applyBridgeProgressEvent(base, { stepId: "planning.model" })).toBe(base);
  });

  it("marks the session cancelled without a pointer", () => {
    const state = applyBridgeProgressEvent(createSdkBridgeState(BINDING), {
      stepId: "bridge.query_table",
      status: "cancelled"
    });
    expect(state.cancelled).toBe(true);
    expect(state.pending).toEqual([]);
  });
});

describe("poll responses", () => {
  it("merges pending calls and clears the poll request", () => {
    const state = applyBridgePollResponse(
      applyBridgeProgressEvent(createSdkBridgeState(BINDING), {
        stepId: "bridge.query_table",
        bridgeTool: { operationId: OP, tool: "query_table" }
      }),
      { calls: [call()], generation: 3 }
    );
    expect(state.generation).toBe(3);
    expect(state.pending).toHaveLength(1);
    expect(state.pollRequested).toBe(false);
    expect(pendingBridgeCall(state, OP)?.requestFingerprint).toBe(FP);
  });

  it("never re-issues a call it has already delivered", () => {
    let state = applyBridgePollResponse(createSdkBridgeState(BINDING), {
      calls: [call()],
      generation: 1
    });
    state = markBridgeDelivered(state, OP, FP);
    state = applyBridgePollResponse(state, { calls: [call()], generation: 2 });
    expect(state.pending).toEqual([]);
  });

  it("rejects a malformed poll response", () => {
    expect(() => assertBridgePollResponse({ calls: [{ operationId: "bad" }] })).toThrow();
  });

  /*
   * 桥支持**两个**工具（`BridgeToolName` = query_table | prepare_python_data）。
   * 校验器必须与指针解析（`bridgeToolPointer`）认同一集合：认指针不认轮询，
   * 会让 `prepare_python_data` 指针拉回来的调用被判无效，异常又被 `drain`
   * 吞掉 → 服务端一个回执都收不到，只能干等超时。这是"数据登记永远拿不到
   * 句柄"的根因，不是可以放宽的加固。
   */
  it("accepts a prepare_python_data call like any other bridge tool", () => {
    expect(() =>
      assertBridgePollResponse({
        calls: [call({ tool: "prepare_python_data", arguments: { sources: [] } })],
        generation: 1
      })
    ).not.toThrow();
    const state = applyBridgePollResponse(createSdkBridgeState(BINDING), {
      calls: [call({ tool: "prepare_python_data" })],
      generation: 1
    });
    expect(pendingBridgeCall(state, OP)?.tool).toBe("prepare_python_data");
  });

  it("still rejects a tool the bridge does not carry", () => {
    expect(() =>
      assertBridgePollResponse({
        calls: [call({ tool: "read_range" })],
        generation: 1
      })
    ).toThrow();
  });
});

describe("receipt decisions mirror the server", () => {
  it("accepts the first delivery of a live call", () => {
    const state = applyBridgePollResponse(createSdkBridgeState(BINDING), {
      calls: [call()],
      generation: 1
    });
    expect(decideBridgeReceipt(state, OP, FP)).toBe("accepted");
  });

  it("treats the same id and fingerprint as a duplicate", () => {
    let state = applyBridgePollResponse(createSdkBridgeState(BINDING), {
      calls: [call()],
      generation: 1
    });
    state = markBridgeDelivered(state, OP, FP);
    expect(decideBridgeReceipt(state, OP, FP)).toBe("duplicate");
  });

  it("rejects the same id with a different fingerprint as a conflict", () => {
    const state = applyBridgePollResponse(createSdkBridgeState(BINDING), {
      calls: [call()],
      generation: 1
    });
    expect(decideBridgeReceipt(state, OP, "b".repeat(64))).toBe("conflict");
  });

  it("rejects an unknown or cancelled operation as late", () => {
    const state = createSdkBridgeState(BINDING);
    expect(decideBridgeReceipt(state, OP, FP)).toBe("late");
    expect(decideBridgeReceipt(cancelBridge(state), OP, FP)).toBe("late");
  });

  it("reports pending calls for the transport to execute", () => {
    const state = applyBridgePollResponse(createSdkBridgeState(BINDING), {
      calls: [call()],
      generation: 1
    });
    expect(pendingBridgeCall(state, OP)?.arguments).toEqual({ mode: "aggregate" });
    expect(pendingBridgeCall(state, "op_" + "2".repeat(32))).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// 宿主适配器编排：轮询 → 复核授权 → 真实读取 → 回执
// ---------------------------------------------------------------------------

interface Recorded {
  polls: number;
  receipts: Array<{ operationId: string; outcome: unknown }>;
  verified: string[];
  executed: string[];
}

function ports(
  overrides: Partial<BridgeHostPorts> & {
    available?: Array<{ operationId: string; requestFingerprint: string }>;
    receiptStatus?: "accepted" | "duplicate" | "rejected";
    recorded?: Recorded;
  } = {}
): { ports: BridgeHostPorts; recorded: Recorded } {
  const recorded: Recorded =
    overrides.recorded ?? { polls: 0, receipts: [], verified: [], executed: [] };
  const available = overrides.available ?? [];
  const status = overrides.receiptStatus ?? "accepted";
  return {
    recorded,
    ports: {
      poll: async () => {
        recorded.polls += 1;
        return {
          calls: available.map((item) => ({
            operationId: item.operationId,
            tool: "query_table" as const,
            arguments: { mode: "aggregate" },
            requestFingerprint: item.requestFingerprint,
            expiresAt: 0
          })),
          generation: recorded.polls
        };
      },
      verify: (binding, call) => {
        recorded.verified.push(call.operationId);
        return { ok: true };
      },
      execute: async (_binding, call) => {
        recorded.executed.push(call.operationId);
        return { kind: "result", result: { rows: [["甲", 2]] } };
      },
      receipt: async (_binding, call, outcome) => {
        recorded.receipts.push({ operationId: call.operationId, outcome });
        return { status };
      },
      ...overrides
    }
  };
}

describe("serviceBridgeOnce", () => {
  it("polls on a pointer, verifies authorization, executes and receipts", async () => {
    const state = applyBridgeProgressEvent(createSdkBridgeState(BINDING), {
      status: "started",
      stepId: "bridge.query_table",
      bridgeTool: { operationId: OP, tool: "query_table" }
    });
    const { ports: host, recorded } = ports({
      available: [{ operationId: OP, requestFingerprint: FP }]
    });

    const result = await serviceBridgeOnce(state, host);

    expect(recorded.polls).toBe(1);
    expect(recorded.verified).toEqual([OP]);
    expect(recorded.executed).toEqual([OP]);
    expect(recorded.receipts[0].operationId).toBe(OP);
    expect(result.executed).toBe(1);
    expect(result.outcomes[0]).toMatchObject({ decision: "accepted", status: "accepted" });
    expect(result.state.pending).toEqual([]);
    expect(result.state.delivered[OP]).toBe(FP);
  });

  it("refuses to read the host when the authorization check fails", async () => {
    const state = applyBridgePollResponse(createSdkBridgeState(BINDING), {
      calls: [call()],
      generation: 1
    });
    const { ports: host, recorded } = ports({
      verify: () => ({ ok: false, errorCode: "unauthorized_scope" })
    });

    const result = await serviceBridgeOnce(state, host);

    expect(recorded.executed).toEqual([]);
    expect(recorded.receipts[0].outcome).toEqual({
      kind: "error",
      errorCode: "unauthorized_scope"
    });
    expect(result.executed).toBe(0);
  });

  it("never re-executes an already delivered operation", async () => {
    // 残留状态：指针重复到达，调用仍在待执行列表里但已经交付过。
    const state: SdkBridgeState = {
      ...createSdkBridgeState(BINDING),
      pending: [call() as never],
      delivered: { [OP]: FP }
    };

    const { ports: host, recorded } = ports({});
    const run = await serviceBridgeOnce(state, host);

    expect(recorded.executed).toEqual([]);
    expect(recorded.receipts).toEqual([]);
    expect(run.executed).toBe(0);
    expect(run.outcomes[0]).toMatchObject({ decision: "duplicate", status: null });
    // 短路后必须把它移出待执行列表，否则调用方会空转。
    expect(run.state.pending).toEqual([]);
  });

  it("drops calls the server already filtered out when re-polling", async () => {
    const state = markBridgeDelivered(
      applyBridgePollResponse(createSdkBridgeState(BINDING), {
        calls: [call()],
        generation: 1
      }),
      OP,
      FP
    );
    const { ports: host, recorded } = ports({
      available: [{ operationId: OP, requestFingerprint: FP }]
    });

    const run = await serviceBridgeOnce({ ...state, pollRequested: true }, host);

    expect(recorded.polls).toBe(1);
    expect(recorded.executed).toEqual([]);
    expect(run.outcomes).toEqual([]);
    expect(run.state.pending).toEqual([]);
  });

  it("stops serving once the session is cancelled", async () => {
    const state = cancelBridge(createSdkBridgeState(BINDING));
    const { ports: host, recorded } = ports({ available: [{ operationId: OP, requestFingerprint: FP }] });

    const result = await serviceBridgeOnce(state, host);

    expect(recorded.polls).toBe(0);
    expect(recorded.executed).toEqual([]);
    expect(result.polled).toBe(false);
  });
});

describe("runBridgeUntilIdle", () => {
  it("serves every call handed out by one poll", async () => {
    const seen: string[] = [];
    const host: BridgeHostPorts = {
      poll: async () => ({
        calls: [
          {
            operationId: "op_" + "1".repeat(32),
            tool: "query_table" as const,
            arguments: {},
            requestFingerprint: FP,
            expiresAt: 0
          },
          {
            operationId: "op_" + "2".repeat(32),
            tool: "query_table" as const,
            arguments: {},
            requestFingerprint: FP,
            expiresAt: 0
          }
        ],
        generation: 1
      }),
      verify: () => ({ ok: true }),
      execute: async (_binding, incoming) => {
        seen.push(incoming.operationId);
        return { kind: "result", result: { rows: [] } };
      },
      receipt: async () => ({ status: "accepted" })
    };

    const result = await runBridgeUntilIdle(
      { ...createSdkBridgeState(BINDING), pollRequested: true },
      host,
      { maxRounds: 4 }
    );

    expect(seen).toEqual(["op_" + "1".repeat(32), "op_" + "2".repeat(32)]);
    expect(result.executed).toBe(2);
    expect(result.state.pending).toEqual([]);
  });

  it("stops polling once nothing is pending and no pointer arrived", async () => {
    let polls = 0;
    const host: BridgeHostPorts = {
      poll: async () => {
        polls += 1;
        return { calls: [], generation: polls };
      },
      verify: () => ({ ok: true }),
      execute: async () => ({ kind: "result", result: {} }),
      receipt: async () => ({ status: "accepted" })
    };

    const result = await runBridgeUntilIdle(createSdkBridgeState(BINDING), host, {
      maxRounds: 4
    });

    expect(polls).toBe(0);
    expect(result.polled).toBe(false);
  });

  it("returns immediately without a binding", async () => {
    const { ports: host, recorded } = ports({});
    const result = await runBridgeUntilIdle(createSdkBridgeState(null), host, {
      maxRounds: 3
    });
    expect(recorded.polls).toBe(0);
    expect(result.executed).toBe(0);
  });
});

describe("bridgeWorkbookIdentity", () => {
  it("falls back to the name namespace when no stable identity is known", () => {
    expect(bridgeWorkbookIdentity("报表.xlsx")).toBe("name:报表.xlsx");
    expect(bridgeWorkbookIdentity("报表.xlsx", "   ")).toBe("name:报表.xlsx");
    expect(bridgeWorkbookIdentity("报表.xlsx", "url:https://a/报表.xlsx")).toBe(
      "url:https://a/报表.xlsx"
    );
  });
});

describe("verifyBridgeAuthorization", () => {
  const live = {
    workbookIdentity: BINDING.workbookIdentity,
    sourceFingerprint: BINDING.sourceFingerprint,
    authorizedSheetNames: ["Sheet1", "Sheet2"],
    activeWorksheet: "Sheet1"
  };

  it("授权未变时放行", () => {
    expect(verifyBridgeAuthorization(BINDING, call(), live)).toEqual({ ok: true });
  });

  it("授权**收窄**必须拒绝 —— 静默少读一张表会算出错答案", () => {
    // 这是判别点：如果判据只看"当前授权是否非空"，这里会错误放行。
    expect(
      verifyBridgeAuthorization(BINDING, call(), {
        ...live,
        authorizedSheetNames: ["Sheet1"]
      })
    ).toEqual({ ok: false, errorCode: "unauthorized_scope" });
  });

  it("授权放宽仍然放行（当前授权覆盖绑定范围即可）", () => {
    expect(
      verifyBridgeAuthorization(BINDING, call(), {
        ...live,
        authorizedSheetNames: ["Sheet1", "Sheet2", "Sheet3"]
      })
    ).toEqual({ ok: true });
  });

  it("授权被全部收回必须拒绝", () => {
    expect(
      verifyBridgeAuthorization(BINDING, call(), { ...live, authorizedSheetNames: [] })
    ).toEqual({ ok: false, errorCode: "unauthorized_scope" });
  });

  it("同名不同工作簿必须拒绝（换簿不得沿用旧授权）", () => {
    expect(
      verifyBridgeAuthorization(BINDING, call(), {
        ...live,
        workbookIdentity: "url:https://evil.example/测试.xlsx"
      })
    ).toEqual({ ok: false, errorCode: "source_changed" });
  });

  it("来源指纹变化必须拒绝", () => {
    expect(
      verifyBridgeAuthorization(BINDING, call(), { ...live, sourceFingerprint: "fp-2" })
    ).toEqual({ ok: false, errorCode: "source_changed" });
  });

  it("scope=active 只看当前活动表，与整簿授权范围解耦", () => {
    const active = call({ arguments: { mode: "aggregate", scope: "active" } });
    expect(
      verifyBridgeAuthorization(BINDING, active, {
        ...live,
        authorizedSheetNames: ["Sheet1"]
      })
    ).toEqual({ ok: true });
    expect(
      verifyBridgeAuthorization(BINDING, active, {
        ...live,
        authorizedSheetNames: ["Sheet2"],
        activeWorksheet: "Sheet1"
      })
    ).toEqual({ ok: false, errorCode: "unauthorized_scope" });
  });
});

describe("bridgeRequestedSheetNames", () => {
  it("默认按绑定时的授权范围读整片", () => {
    expect(bridgeRequestedSheetNames(BINDING, call(), { activeWorksheet: "Sheet9" })).toEqual([
      "Sheet1",
      "Sheet2"
    ]);
  });

  it("scope=active 只读当前活动表", () => {
    expect(
      bridgeRequestedSheetNames(
        BINDING,
        call({ arguments: { mode: "rows", scope: "active" } }),
        { activeWorksheet: "Sheet9" }
      )
    ).toEqual(["Sheet9"]);
  });
});
