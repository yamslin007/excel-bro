// @vitest-environment jsdom
/**
 * Excel 端「重新发现未完成执行」的判据测试。
 *
 * 只测发现这一层：身份怎么读、请求发不发、响应怎么判。**不**测恢复资格
 * （那是 `pendingResumeReadiness` 与面板的事），也**不**碰宿主。
 */
import { describe, expect, it, vi } from "vitest";
import type { PendingExecutionDiscoveryResponse } from "@excel-bro/core/pendingExecutionDiscovery";
import { discoverOfficePendingExecutions } from "./officeRecoveryDiscovery";

const IDENTITY = "url:https://contoso.sharepoint.com/book.xlsx";

function candidate(overrides: Partial<PendingExecutionDiscoveryResponse["candidates"][number]> = {}) {
  return {
    operationId: "ea6a0f0d-1111-4a4a-8a8a-111111111111",
    taskId: "task-1111",
    status: "needs_reconciliation" as const,
    createdAt: 1_790_000_000_000,
    updatedAt: 1_790_000_600_000,
    nextActionIndex: 1,
    totalActions: 3,
    inFlightActionIndex: 1,
    ownerActive: false,
    requiresReconciliation: true,
    canResume: true,
    ...overrides
  };
}

function response(
  overrides: Partial<PendingExecutionDiscoveryResponse> = {}
): PendingExecutionDiscoveryResponse {
  return {
    schemaVersion: 1,
    namespace: "office",
    workbookIdentity: IDENTITY,
    limit: 5,
    truncated: false,
    candidates: [candidate()],
    ...overrides
  };
}

function discover(
  value: PendingExecutionDiscoveryResponse | (() => Promise<PendingExecutionDiscoveryResponse>),
  identity: string | null = IDENTITY
) {
  const load = typeof value === "function" ? vi.fn(value) : vi.fn(async () => value);
  return {
    load,
    run: discoverOfficePendingExecutions({ identity: () => identity, load })
  };
}

describe("Excel 未完成执行发现", () => {
  it("身份读不出来时一个请求都不发", async () => {
    const { load, run } = discover(response(), null);
    await expect(run).resolves.toEqual({ kind: "no_identity" });
    expect(load).not.toHaveBeenCalled();
  });

  it("恰好一条且未截断时才直接给出这一条", async () => {
    const { run } = discover(response());
    await expect(run).resolves.toEqual({
      kind: "resolved",
      selection: { kind: "single", candidate: candidate() }
    });
  });

  it("多条候选交给用户选，不自动挑最新的一条", async () => {
    const newest = candidate({
      operationId: "op-newest",
      status: "failed",
      nextActionIndex: 0,
      totalActions: 3,
      inFlightActionIndex: null,
      updatedAt: 1_790_000_900_000
    });
    const inFlight = candidate({ operationId: "op-in-flight", updatedAt: 1_790_000_100_000 });
    const { run } = discover(response({ candidates: [newest, inFlight] }));
    const outcome = await run;
    expect(outcome.kind).toBe("resolved");
    if (outcome.kind !== "resolved") throw new Error("unreachable");
    expect(outcome.selection.kind).toBe("choose");
    if (outcome.selection.kind !== "choose") throw new Error("unreachable");
    // 有在途动作的排在前面：写前失败的那条"一个动作都没推进"，最不该挡住别人。
    expect(outcome.selection.candidates.map((item) => item.operationId))
      .toEqual(["op-in-flight", "op-newest"]);
    expect(outcome.selection.truncated).toBe(false);
  });

  it("服务端给的不是这个工作簿的记录时不说成“没有”", async () => {
    const { run } = discover(response({ workbookIdentity: "url:https://contoso.sharepoint.com/other.xlsx" }));
    const outcome = await run;
    expect(outcome.kind).toBe("resolved");
    if (outcome.kind !== "resolved") throw new Error("unreachable");
    expect(outcome.selection.kind).toBe("mismatch");
  });

  it("服务不可达时给“暂时读不到”，不折成“没有未完成任务”", async () => {
    const { run } = discover(async () => {
      throw new Error("Failed to fetch");
    });
    const outcome = await run;
    expect(outcome.kind).toBe("unavailable");
    if (outcome.kind !== "unavailable") throw new Error("unreachable");
    expect(outcome.detail).toContain("暂时读不到");
    expect(outcome.reason).toContain("Failed to fetch");
  });

  it("响应夹带执行权令牌时整份拒收", async () => {
    const leaky = { ...response(), ownerToken: "secret" };
    const { run } = discover(leaky as unknown as PendingExecutionDiscoveryResponse);
    const outcome = await run;
    expect(outcome.kind).toBe("unavailable");
    if (outcome.kind !== "unavailable") throw new Error("unreachable");
    expect(outcome.reason).toContain("未知字段");
  });

  it("WPS 命名空间的候选不会出现在 Excel 面板里", async () => {
    // 形状合法的 WPS 响应（身份前缀也对）：共享准入闸放行，靠命名空间这一层拦下。
    const { run } = discover({
      ...response(),
      namespace: "wps",
      workbookIdentity: "wps-path:e:\\codex\\book.xlsx"
    } as unknown as PendingExecutionDiscoveryResponse);
    const outcome = await run;
    expect(outcome.kind).toBe("unavailable");
    if (outcome.kind !== "unavailable") throw new Error("unreachable");
    expect(outcome.detail).toContain("不是 Office 写入的记录");
  });

  it("命名空间与身份前缀对不上时由共享准入闸直接拒收", async () => {
    const { run } = discover({
      ...response(),
      namespace: "wps",
      workbookIdentity: IDENTITY
    } as unknown as PendingExecutionDiscoveryResponse);
    const outcome = await run;
    expect(outcome.kind).toBe("unavailable");
    if (outcome.kind !== "unavailable") throw new Error("unreachable");
    expect(outcome.reason).toContain("工作簿身份无效");
  });
});
