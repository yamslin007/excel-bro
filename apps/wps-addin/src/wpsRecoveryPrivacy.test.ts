// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import { registerWpsExecution, startWpsAction, completeWpsAction, reconcileWpsActionReceipt,
  suspendWpsExecution, finishWpsExecution } from "../../excel-addin/src/api";
import { wpsRecoveryDescriptor } from "./wpsExecutionLedger";
import type { WpsPreparedPlan } from "./wpsPlanExecutor";

/**
 * CAP-03 隐私边界的**真实出站**证据。
 *
 * 这里断言的是**真的要发给服务端的那个字符串**（stub 掉 `fetch` 后截获 body），
 * 不是源码里有没有某行字，也不是某个对象"看起来"少了个字段。
 *
 * 每条"不该有"的断言都配一条"确实有"的锚点：先证明哨兵真的存在于本机计划里，
 * 否则"body 里没有哨兵"可能只是因为哨兵压根没进计划，断言会空过。
 */

/** 只存在于单元格值里的哨兵。刻意不用短数字串 —— 那会被时间戳之类命中。 */
const SENTINEL = "SENTINEL-CONFIDENTIAL-7f3a91c0";

function planWithSentinel(): AnalysisPlan {
  return {
    id: "plan-privacy",
    title: "写入含机密值的区域",
    summary: `把目标区域写成 ${SENTINEL}`,
    assumptions: [],
    warnings: [],
    actions: [
      {
        type: "writeValues",
        sheet: "Sheet1",
        range: "A1:B2",
        values: [
          [SENTINEL, "普通值"],
          ["第三行", SENTINEL]
        ]
      }
    ],
    acceptanceCriteria: []
  };
}

function preparedWithSentinel(): WpsPreparedPlan {
  const values = (
    planWithSentinel().actions[0] as Extract<AnalysisPlan["actions"][number], { type: "writeValues" }>
  ).values as CellValue[][];
  return {
    plan: planWithSentinel(),
    sourceSnapshots: [],
    authorization: {
      workbookFullName: "C:\\data\\测试.xlsx",
      sheets: [{ index: 1, name: "Sheet1" }]
    } as WpsPreparedPlan["authorization"],
    workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
    hostVersion: "12.0",
    actions: [
      {
        actionIndex: 0,
        type: "writeValues",
        sheetIndex: 1,
        sheetName: "Sheet1",
        range: "A1:B2",
        rowCount: 2,
        columnCount: 2,
        cellCount: 4,
        values,
        preview: values,
        beforeValues: [[null, null], [null, null]],
        beforeValuesCaptured: true,
        previewFingerprint: "4:1a2b3c4d",
        rect: { startRow: 1, startColumn: 1, endRow: 2, endColumn: 2 }
      }
    ],
    sources: [],
    totalCells: 4,
    // 这份计划写的是固定值、且没有本机来源复核 ⇒ 本来就需要来源锚点。
    sourceAnchorRequired: true,
    previewRevision: "rev-1",
    confirmationSnapshot: "snap-1"
  };
}

type StubbedFetch = ReturnType<typeof vi.fn>;

function stubFetch(handler: (body: string) => Response): StubbedFetch {
  const fetchMock = vi.fn(async (_url: unknown, init?: { body?: unknown }) => {
    return handler(String(init?.body ?? ""));
  });
  vi.stubGlobal("fetch", fetchMock);
  return fetchMock;
}

function okResponse(state?: Record<string, unknown>): Response {
  return new Response(
    JSON.stringify({
      operationId: "wps-execution-privacy",
      status: "running",
      revision: 1,
      workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
      taskId: "task-privacy",
      descriptor: wpsRecoveryDescriptor(preparedWithSentinel()),
      authorizedSheetNames: ["Sheet1"],
      nextActionIndex: 0,
      totalActions: 1,
      inFlightActionIndex: null,
      canResume: true,
      ownerActive: false,
      requiresReconciliation: false,
      ownerToken: "owner-token-0123456789abcdef0123456789abcdef",
      ...state
    }),
    { status: 200, headers: { "Content-Type": "application/json" } }
  );
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("WPS 恢复登记的隐私边界", () => {
  it("锚点：哨兵确实存在于本机已确认计划里（防止后面的'不该有'空过）", () => {
    const prepared = preparedWithSentinel();
    expect(JSON.stringify(prepared.plan.actions)).toContain(SENTINEL);
    expect(JSON.stringify(prepared.actions)).toContain(SENTINEL);
  });

  it("恢复描述不含目标值矩阵", () => {
    const descriptor = wpsRecoveryDescriptor(preparedWithSentinel());
    const encoded = JSON.stringify(descriptor);
    expect(encoded).not.toContain(SENTINEL);
    expect(encoded).not.toContain("values");
    expect(descriptor.actions).toHaveLength(1);
    // 但恢复必需的定位信息必须在：类型、工作表、区域、形状、指纹。
    expect(descriptor.actions[0]).toMatchObject({
      actionIndex: 0,
      type: "writeValues",
      sheet: "Sheet1",
      range: "A1:B2",
      rowCount: 2,
      columnCount: 2,
      cellCount: 4,
      previewFingerprint: "4:1a2b3c4d"
    });
  });

  it("真实 fetch 请求体不含哨兵，也不携带 plan 字段", async () => {
    const fetchMock = stubFetch(() => okResponse());
    await registerWpsExecution({
      operationId: "wps-execution-privacy",
      descriptor: wpsRecoveryDescriptor(preparedWithSentinel()),
      authorizedSheetNames: ["Sheet1"],
      workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
      taskId: "task-privacy"
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    const body = String((fetchMock.mock.calls[0]?.[1] as { body: string }).body);
    expect(body).not.toContain(SENTINEL);
    const parsed = JSON.parse(body) as Record<string, unknown>;
    // 整份计划（含 writeValues.values）根本不在请求里，而不是"传了但对方没存"。
    expect(parsed.plan).toBeUndefined();
    expect(Object.keys(parsed).sort()).toEqual(
      ["authorizedSheetNames", "confirmed", "descriptor", "operationId", "taskId", "workbookIdentity"].sort()
    );
    const descriptor = parsed.descriptor as { actions: Record<string, unknown>[] };
    expect(descriptor.actions[0].values).toBeUndefined();
  });

  it("真实所有回执与异常请求只含白名单，额外顶层计划也不会出站", async () => {
    const prepared = preparedWithSentinel();
    expect(JSON.stringify(prepared.plan.actions)).toContain(SENTINEL);
    const fetchMock = stubFetch(() => okResponse());
    const evidence = { sheetName: "Sheet1", range: "A1:B2", type: "writeValues", beforeFingerprint: "4:1a2b3c4d" };
    const extra = { plan: prepared.plan };
    await registerWpsExecution({ operationId: "wps-execution-privacy", taskId: "task-privacy",
      workbookIdentity: "wps-path:c:\\data\\测试.xlsx", authorizedSheetNames: ["Sheet1"],
      descriptor: wpsRecoveryDescriptor(prepared), ...extra });
    await startWpsAction("wps-execution-privacy", { expectedRevision: 1, ownerToken: "owner", actionIndex: 0,
      beforeSnapshot: evidence, ...extra });
    await completeWpsAction("wps-execution-privacy", { expectedRevision: 2, ownerToken: "owner", actionIndex: 0,
      checkpoint: { ...evidence, afterFingerprint: "4:11111111" }, ...extra });
    await reconcileWpsActionReceipt("wps-execution-privacy", { expectedRevision: 2, actionIndex: 0,
      outcome: "applied", checkpoint: { ...evidence, afterFingerprint: "4:11111111" }, ...extra });
    await suspendWpsExecution("wps-execution-privacy", { expectedRevision: 3, ownerToken: "owner",
      code: SENTINEL, message: SENTINEL, ...extra });
    await finishWpsExecution("wps-execution-privacy", { expectedRevision: 4, ownerToken: "owner",
      result: { actionCount: 1, cellCount: 4 }, ...extra });
    expect(fetchMock).toHaveBeenCalledTimes(6);
    for (const call of fetchMock.mock.calls) {
      const body = String((call[1] as { body: string }).body);
      expect(body).not.toContain(SENTINEL);
      expect(JSON.parse(body)).not.toHaveProperty("plan");
      expect(body).not.toContain('"values"');
    }
  });

  it.each(["before", "after", "reconcile", "result"])("%s 内嵌值矩阵在 fetch 前拒绝", async (phase) => {
    const prepared = preparedWithSentinel();
    expect(JSON.stringify(prepared.actions)).toContain(SENTINEL);
    const fetchMock = stubFetch(() => okResponse());
    const evidence = { sheetName: "Sheet1", range: "A1:B2", type: "writeValues", beforeFingerprint: "fingerprint",
      values: [[SENTINEL]] };
    const request = phase === "before" ? startWpsAction("operation", { expectedRevision: 1, ownerToken: "owner", actionIndex: 0, beforeSnapshot: evidence })
      : phase === "after" ? completeWpsAction("operation", { expectedRevision: 1, ownerToken: "owner", actionIndex: 0, checkpoint: evidence })
      : phase === "reconcile" ? reconcileWpsActionReceipt("operation", { expectedRevision: 1, actionIndex: 0, outcome: "applied", checkpoint: evidence })
      : finishWpsExecution("operation", { expectedRevision: 1, ownerToken: "owner", result: { actionCount: 1, values: [[SENTINEL]] } });
    await expect(request).rejects.toThrow(/矩阵|单元格内容/);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("合法键 cellCount 也不能塞矩阵，不能只检测键名", async () => {
    expect(JSON.stringify(preparedWithSentinel().plan.actions)).toContain(SENTINEL);
    const fetchMock = stubFetch(() => okResponse());
    await expect(finishWpsExecution("operation", { expectedRevision: 1, ownerToken: "owner",
      result: { actionCount: 1, cellCount: [[SENTINEL]] } })).rejects.toThrow(/单元格内容/);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("服务端若把矩阵塞回状态，客户端直接拒绝接收", async () => {
    stubFetch(() =>
      okResponse({
        descriptor: {
          actions: [
            {
              actionIndex: 0,
              type: "writeValues",
              sheet: "Sheet1",
              range: "A1:B2",
              rowCount: 2,
              columnCount: 2,
              cellCount: 4,
              previewFingerprint: "4:1a2b3c4d",
              values: [[SENTINEL, "普通值"]]
            }
          ]
        }
      })
    );
    await expect(
      registerWpsExecution({
        operationId: "wps-execution-privacy",
        descriptor: wpsRecoveryDescriptor(preparedWithSentinel()),
        authorizedSheetNames: ["Sheet1"],
        workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
        taskId: "task-privacy"
      })
    ).rejects.toThrow(/单元格矩阵/);
  });
});
