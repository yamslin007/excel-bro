// @vitest-environment jsdom
/**
 * Excel 侧桥宿主适配器。
 *
 * 断言的是**实际请求体**：拦截真实的 `fetch`，检查送往 `/api/sdk/tool-bridge/*`
 * 的 JSON 原样包含哪一轮的绑定、哪一个操作 ID 与什么回执。
 * 本地查询执行用注入替身（Office.js 在 jsdom 里不存在），但适配器本身、
 * 授权复核与错误码映射都是生产代码。
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  BRIDGE_UNVERIFIABLE_ERROR_CODE,
  verifyBridgeAuthorization,
  verifyBridgeAuthorizationOrRefuse
} from "@excel-bro/core";

import { DataToolExecutionError } from "./dataTools";
import { resetTrustedSourceReads, trustedToolReads } from "./planSourceDependency";
import {
  bridgeErrorCode,
  createSdkBridgeBinding,
  createSdkBridgeHostPorts,
  loadBridgeResult,
  type SdkBridgeLiveContext
} from "./sdkBridgeHost";

const OP = "op_" + "1".repeat(32);
const FP = "a".repeat(64);

const LIVE: SdkBridgeLiveContext = {
  workbookIdentity: "url:https://contoso.example/报表.xlsx",
  sourceFingerprint: "fp-1",
  authorizedSheetNames: ["Sheet1", "Sheet2"],
  activeWorksheet: "Sheet1"
};

const BINDING = createSdkBridgeBinding({
  taskId: "turn-bridge-0001",
  hostId: "excel-office",
  workbookName: "报表.xlsx",
  workbookIdentity: "url:https://contoso.example/报表.xlsx",
  sourceFingerprint: "fp-1",
  authorizedSheetNames: ["Sheet1", "Sheet2"]
});

function call(overrides: Record<string, unknown> = {}) {
  return {
    operationId: OP,
    tool: "query_table" as const,
    arguments: { mode: "aggregate", groupBy: ["类别"] },
    requestFingerprint: FP,
    expiresAt: 0,
    ...overrides
  };
}

interface Captured {
  url: string;
  body: Record<string, unknown>;
}

let captured: Captured[] = [];

function stubFetch(payloads: unknown[]) {
  let index = 0;
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      captured.push({
        url: String(url),
        body: JSON.parse(String(init?.body ?? "{}"))
      });
      const payload = payloads[Math.min(index, payloads.length - 1)];
      index += 1;
      return new Response(JSON.stringify(payload ?? {}), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    })
  );
}

beforeEach(() => {
  captured = [];
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("createSdkBridgeBinding", () => {
  it("归一化工作簿身份，缺省时退回 name: 命名空间", () => {
    expect(BINDING.workbookIdentity).toBe("url:https://contoso.example/报表.xlsx");
    const fallback = createSdkBridgeBinding({
      taskId: "turn-bridge-0002",
      hostId: "excel-office",
      workbookName: "报表.xlsx"
    });
    expect(fallback.workbookIdentity).toBe("name:报表.xlsx");
  });

  it("快照本轮授权范围，且与传入数组解耦（后续被改动不影响绑定）", () => {
    const scope = ["Sheet1", "Sheet2"];
    const binding = createSdkBridgeBinding({
      taskId: "turn-bridge-0003",
      hostId: "excel-office",
      workbookName: "报表.xlsx",
      authorizedSheetNames: scope
    });
    scope.push("Sheet3");
    expect(binding.authorizedSheetNames).toEqual(["Sheet1", "Sheet2"]);
  });
});

describe("共享复核判据（Excel 现场）", () => {
  it("授权不变时放行", () => {
    expect(verifyBridgeAuthorization(BINDING, call(), LIVE)).toEqual({ ok: true });
  });

  it("同名但不同工作簿必须拒绝（换簿不得沿用旧授权）", () => {
    const verdict = verifyBridgeAuthorization(BINDING, call(), {
      ...LIVE,
      workbookIdentity: "url:https://evil.example/报表.xlsx"
    });
    expect(verdict).toEqual({ ok: false, errorCode: "source_changed" });
  });

  it("来源指纹变化必须拒绝", () => {
    const verdict = verifyBridgeAuthorization(BINDING, call(), {
      ...LIVE,
      sourceFingerprint: "fp-2"
    });
    expect(verdict).toEqual({ ok: false, errorCode: "source_changed" });
  });

  it("请求范围超出当前授权必须拒绝（收窄也拒，哪怕指纹没变）", () => {
    const verdict = verifyBridgeAuthorization(BINDING, call(), {
      ...LIVE,
      authorizedSheetNames: ["Sheet1"],
      sourceFingerprint: "fp-2"
    });
    // 收窄是"授权"问题，要重新勾选范围；指成 source_changed 会把用户引向错误动作。
    expect(verdict).toEqual({ ok: false, errorCode: "unauthorized_scope" });
  });

  it("scope=active 时只核对当前工作表", () => {
    const scoped = verifyBridgeAuthorization(
      BINDING,
      call({ arguments: { scope: "active" } }),
      { ...LIVE, authorizedSheetNames: ["Sheet1"] }
    );
    expect(scoped).toEqual({ ok: true });
  });

  it("读不到现场时 fail-closed 拒绝：不能把异常当成「没变化」放过", async () => {
    await expect(
      verifyBridgeAuthorizationOrRefuse(BINDING, call(), () => {
        throw new Error("宿主读不到：敏感细节");
      })
    ).resolves.toEqual({
      ok: false,
      errorCode: BRIDGE_UNVERIFIABLE_ERROR_CODE
    });
  });

  it("读得到现场时与直接复核同一结论（不许另起一套判据）", async () => {
    await expect(
      verifyBridgeAuthorizationOrRefuse(BINDING, call(), () => LIVE)
    ).resolves.toEqual({ ok: true });
    await expect(
      verifyBridgeAuthorizationOrRefuse(BINDING, call(), async () => ({
        ...LIVE,
        sourceFingerprint: "fp-2"
      }))
    ).resolves.toEqual({ ok: false, errorCode: "source_changed" });
  });
});

describe("bridgeErrorCode", () => {
  it("把本地确定性查询的错误映射成服务端固定错误码", () => {
    expect(bridgeErrorCode(new DataToolExecutionError("SCOPE_VIOLATION", "", false))).toBe(
      "unauthorized_scope"
    );
    expect(
      bridgeErrorCode(new DataToolExecutionError("TOOL_LIMIT_EXCEEDED", "", false))
    ).toBe("result_too_large");
    expect(new DataToolExecutionError("FIELD_NOT_FOUND", "", false));
    expect(bridgeErrorCode(new DataToolExecutionError("FIELD_NOT_FOUND", "", false))).toBe(
      "invalid_arguments"
    );
    expect(bridgeErrorCode(new Error("boom"))).toBe("internal");
  });
});

describe("createSdkBridgeHostPorts", () => {
  function ports(
    overrides: Partial<Parameters<typeof createSdkBridgeHostPorts>[2]> = {},
    live: SdkBridgeLiveContext = LIVE
  ) {
    return createSdkBridgeHostPorts(
      { liveContext: () => live },
      BINDING,
      {
        execute: vi.fn(async () => ({
          requestId: OP,
          tool: "query_table" as const,
          title: "分组结果",
          headers: ["类别", "数量"],
          rows: [["甲", 2]],
          sourceSheets: ["Sheet1"],
          scannedRows: 12,
          complete: true,
          calculation: "按类别计数",
          warnings: []
        })),
        ...overrides
      }
    );
  }

  it("轮询请求体带完整的归属（含工作簿身份）", async () => {
    stubFetch([{ calls: [], generation: 3 }]);
    await ports().poll(BINDING);

    expect(captured).toHaveLength(1);
    expect(captured[0].url).toContain("/api/sdk/tool-bridge/poll");
    expect(captured[0].body).toEqual({
      taskId: "turn-bridge-0001",
      hostId: "excel-office",
      workbookName: "报表.xlsx",
      workbookIdentity: "url:https://contoso.example/报表.xlsx",
      sourceFingerprint: "fp-1"
    });
  });

  it("回执请求体带操作 ID、指纹与真实结果", async () => {
    stubFetch([{ status: "accepted", code: null }]);
    const outcome = await ports().execute(BINDING, call());
    await ports().receipt(BINDING, call(), outcome);

    expect(captured[0].url).toContain("/api/sdk/tool-bridge/receipt");
    expect(captured[0].body).toMatchObject({
      taskId: "turn-bridge-0001",
      workbookIdentity: "url:https://contoso.example/报表.xlsx",
      operationId: OP,
      requestFingerprint: FP
    });
    const receipt = captured[0].body.receipt as Record<string, unknown>;
    expect(receipt.result).toMatchObject({ rows: [["甲", 2]], headers: ["类别", "数量"] });
    expect(receipt.errorCode).toBeUndefined();
  });

  it("执行失败时回执只带固定错误码，不带异常文本", async () => {
    stubFetch([{ status: "accepted", code: null }]);
    const failing = ports({
      execute: vi.fn(async () => {
        throw new DataToolExecutionError("SCOPE_VIOLATION", "敏感细节", false);
      })
    });
    const outcome = await failing.execute(BINDING, call());
    expect(outcome).toEqual({ kind: "error", errorCode: "unauthorized_scope" });

    await failing.receipt(BINDING, call(), outcome);
    expect(captured[0].body.receipt).toEqual({ errorCode: "unauthorized_scope" });
    expect(JSON.stringify(captured[0].body)).not.toContain("敏感细节");
  });

  it("prepare_python_data 走 datasets 回执，值只进回执字段", async () => {
    stubFetch([{ status: "accepted", code: null }]);
    const prepared = createSdkBridgeHostPorts(
      { liveContext: () => LIVE },
      BINDING,
      {
        execute: vi.fn(),
        prepareData: vi.fn(async () => [
          { sourceIndex: 0, values: [[1, 2], [3, 4]] }
        ])
      }
    );
    const outcome = await prepared.execute(
      BINDING,
      call({
        tool: "prepare_python_data" as const,
        arguments: { sources: [{ sheet: "Sheet1", range: "A1:B2" }] }
      })
    );
    expect(outcome.kind).toBe("result");
    const payload = (outcome as { kind: "result"; result: unknown }).result as {
      datasets: unknown[];
    };
    expect(payload.datasets).toHaveLength(1);

    await prepared.receipt(
      BINDING,
      call({
        tool: "prepare_python_data" as const,
        arguments: { sources: [{ sheet: "Sheet1", range: "A1:B2" }] }
      }),
      outcome
    );
    const receipt = captured[0].body.receipt as Record<string, unknown>;
    expect(receipt.datasets).toHaveLength(1);
    expect(receipt.result).toBeUndefined();
    expect(receipt.errorCode).toBeUndefined();
  });

  it("prepare_python_data 范围越权时不读宿主，只回固定错误码", async () => {
    stubFetch([{ status: "accepted", code: null }]);
    const execute = vi.fn();
    const prepared = createSdkBridgeHostPorts(
      { liveContext: () => LIVE },  // live 授权 = binding 授权 = ["Sheet1","Sheet2"]
      BINDING,
      { execute }
    );
    await expect(
      prepared.verify(
        BINDING,
        call({
          tool: "prepare_python_data" as const,
          arguments: { sources: [{ sheet: "Sheet3", range: "A1:B2" }] }
        })
      )
    ).resolves.toEqual({ ok: false, errorCode: "unauthorized_scope" });
    expect(execute).not.toHaveBeenCalled();
  });

  it("授权复核失败时**不读宿主**，只回一个固定错误码", async () => {
    stubFetch([{ status: "accepted", code: null }]);
    const execute = vi.fn();
    const verifying = ports({ execute }, { ...LIVE, authorizedSheetNames: ["Sheet1"] });
    await expect(verifying.verify(BINDING, call())).resolves.toEqual({
      ok: false,
      errorCode: "unauthorized_scope"
    });
    expect(execute).not.toHaveBeenCalled();
  });

  /*
   * 读不到现场 ≠ 没变化。这里必须产出**一张拒绝回执**，而不是让 verify 抛异常：
   * 抛出去会被 hook 的 drain 吞掉，服务端一个回执都收不到，只能干等到超时 ——
   * 那正是"无人回执的等待"。
   */
  it("现场读不到时 verify 也**不抛**，而是给出可回执的固定错误码", async () => {
    const execute = vi.fn();
    const unreadable = createSdkBridgeHostPorts(
      {
        liveContext: () => {
          throw new Error("未保存工作簿身份无法核实：D:\\敏感\\x.xlsx");
        }
      },
      BINDING,
      { execute }
    );

    const verdict = await unreadable.verify(BINDING, call());
    expect(verdict).toEqual({ ok: false, errorCode: "unauthorized_scope" });
    expect(execute).not.toHaveBeenCalled();

    // 真的把回执发出去 —— 只断言"抛没抛"会漏掉"没人发回执"这一半。
    stubFetch([{ status: "accepted", code: null }]);
    const outcome = verdict.ok
      ? await unreadable.execute(BINDING, call())
      : ({ kind: "error", errorCode: verdict.errorCode } as const);
    await unreadable.receipt(BINDING, call(), outcome);
    expect(captured[0].url).toContain("/api/sdk/tool-bridge/receipt");
    expect(captured[0].body.receipt).toEqual({ errorCode: "unauthorized_scope" });
    // 宿主路径是隐私：异常文本绝不出站。
    expect(JSON.stringify(captured[0].body)).not.toContain("敏感");
  });

  /*
   * 「宿主成功读过的范围」——数据登记也必须记账。
   *
   * 与 ``query_table`` 同一条铁律：只有**确定性执行器真的读完**那一刻才写记账，
   * 谁都没读过就不写。漏掉这一条，Python 轮次的"实际来源依赖"恒为空 ——
   * 恢复登记会以为"不知道依赖谁"，于是把有来源的计划当作不可恢复。
   */
  it("prepare_python_data 成功读取后按实际来源记账", async () => {
    stubFetch([{ status: "accepted", code: null }]);
    resetTrustedSourceReads();
    const prepared = createSdkBridgeHostPorts({ liveContext: () => LIVE }, BINDING, {
      execute: vi.fn(),
      prepareData: vi.fn(async () => [
        { sourceIndex: 0, values: [[1, 2]] },
        { sourceIndex: 1, values: [[3, 4]] }
      ])
    });

    const outcome = await prepared.execute(
      BINDING,
      call({
        tool: "prepare_python_data" as const,
        arguments: {
          sources: [
            { sheet: "Sheet1", range: "A1:B2" },
            { sheet: "Sheet2", range: "A1:B2" }
          ]
        }
      })
    );
    expect(outcome.kind).toBe("result");
    expect(trustedToolReads(BINDING.taskId)).toEqual({
      origin: "tool-execution-record",
      sheets: ["Sheet1", "Sheet2"]
    });
  });

  /* 只交输入表的一部分时，记账只有**实际读过**的那张，不虚报另一张。 */
  it("prepare_python_data 只记账实际读取的来源", async () => {
    stubFetch([{ status: "accepted", code: null }]);
    resetTrustedSourceReads();
    const prepared = createSdkBridgeHostPorts({ liveContext: () => LIVE }, BINDING, {
      execute: vi.fn(),
      prepareData: vi.fn(async () => [{ sourceIndex: 0, values: [[1, 2]] }])
    });

    await prepared.execute(
      BINDING,
      call({
        tool: "prepare_python_data" as const,
        arguments: { sources: [{ sheet: "Sheet2", range: "A1:B2" }] }
      })
    );
    expect(trustedToolReads(BINDING.taskId)?.sheets).toEqual(["Sheet2"]);
  });

  /* 读取失败时**不留**记账：没读到就不算"读过"。 */
  it("prepare_python_data 读取失败时不留任何记账", async () => {
    stubFetch([{ status: "accepted", code: null }]);
    resetTrustedSourceReads();
    const prepared = createSdkBridgeHostPorts({ liveContext: () => LIVE }, BINDING, {
      execute: vi.fn(),
      prepareData: vi.fn(async () => {
        throw new DataToolExecutionError("SCOPE_VIOLATION", "表外区域", false);
      })
    });

    const outcome = await prepared.execute(
      BINDING,
      call({
        tool: "prepare_python_data" as const,
        arguments: { sources: [{ sheet: "Sheet1", range: "A1:B2" }] }
      })
    );
    expect(outcome).toEqual({ kind: "error", errorCode: "unauthorized_scope" });
    expect(trustedToolReads(BINDING.taskId)).toBeNull();
  });

  /* 执行阶段才读不到现场（直接调 execute 的路径）也不能让异常冒出去。 */
  it("execute 阶段现场读不到时也只回固定错误码", async () => {
    let readable = true;
    const flaky = createSdkBridgeHostPorts(
      {
        liveContext: () => {
          if (!readable) throw new Error("宿主读不到：D:\\敏感\\x.xlsx");
          return LIVE;
        }
      },
      BINDING,
      { execute: vi.fn() }
    );

    readable = false;
    await expect(flaky.execute(BINDING, call())).resolves.toEqual({
      kind: "error",
      errorCode: "unauthorized_scope"
    });
    expect(captured).toHaveLength(0);
  });
});

describe("loadBridgeResult", () => {
  it("按引用取值时回带该轮的绑定，而不是当前绑定", async () => {
    stubFetch([
      {
        resultRef: "res_" + "0".repeat(32),
        title: "本地查询结果",
        headers: ["类别"],
        rows: [["甲"]],
        sourceSheets: ["Sheet1"],
        warnings: [],
        complete: true,
        calculation: "计数",
        scannedRows: 3
      }
    ]);

    const payload = await loadBridgeResult(
      {
        taskId: "turn-bridge-0001",
        hostId: "excel-office",
        workbookName: "报表.xlsx",
        workbookIdentity: "url:https://contoso.example/报表.xlsx",
        sourceFingerprint: "fp-1"
      },
      "res_" + "0".repeat(32)
    );

    expect(captured[0].url).toContain("/api/sdk/tool-bridge/result");
    expect(captured[0].body).toMatchObject({
      taskId: "turn-bridge-0001",
      workbookIdentity: "url:https://contoso.example/报表.xlsx",
      resultRef: "res_" + "0".repeat(32)
    });
    expect(payload?.rows).toEqual([["甲"]]);
  });

  it("取值失败（过期/授权不符/断线）返回 null，不抛给界面", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => new Response("{}", { status: 404 }))
    );
    expect(
      await loadBridgeResult(
        {
          taskId: "turn-bridge-0001",
          hostId: "excel-office",
          workbookName: "报表.xlsx",
          workbookIdentity: "url:https://contoso.example/报表.xlsx"
        },
        "res_" + "0".repeat(32)
      )
    ).toBeNull();
  });
});
