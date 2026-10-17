// @vitest-environment jsdom
/**
 * 规划流 ↔ 本机桥的接线（Excel 端）。
 *
 * 这里跑的是**完整链路**，只把两处边界换成替身：
 *
 * 1. **网络边界**：`fetch` 是替身，但断言的是真实发出的**请求体**
 *    （poll / receipt 各自带了哪一轮绑定、哪个 operationId、什么回执）。
 * 2. **Office.js 边界**：jsdom 里没有 Excel，所以"读宿主"用固定的内存矩阵
 *    顶上。但**计算本身是真的** —— 矩阵交给既有的 `executeQueryTableData`
 *    算，不喂预置结果。表头识别、跨表合并、分组求和都是生产代码算出来的。
 *
 * 状态机（`runBridgeUntilIdle`）、端口组装（`createSdkBridgeHostPorts`）、
 * 授权复核、回执投影全部是生产实现。
 */

import { act, cleanup, fireEvent, render, renderHook, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DataToolExecutionError, executeQueryTableData, type SheetValues } from "../dataTools";
import type { SdkBridgeLiveContext } from "../sdkBridgeHost";
import type { DataToolRequest, DataToolResult } from "../contracts";
import { useSdkToolBridge } from "./useSdkToolBridge";

const OP = "op_" + "2".repeat(32);
const FP = "b".repeat(64);

/** 两张授权表：同一批列，需要跨表合并后再分组汇总。 */
const WORKBOOK: Record<string, SheetValues> = {
  Sheet1: {
    name: "Sheet1",
    values: [
      ["影院名称", "影片名称", "计数项:影厅"],
      ["影院A", "功夫女足", 11],
      ["影院B", "恐怖游轮", 2]
    ]
  },
  Sheet2: {
    name: "Sheet2",
    values: [
      ["影院名称", "影片名称", "计数项:影厅"],
      ["影院C", "功夫女足", 13],
      ["影院D", "恐怖游轮", 1]
    ]
  }
};

const QUERY_ARGUMENTS = {
  mode: "aggregate" as const,
  scope: "selected" as const,
  groupBy: ["影片名称"],
  metrics: [{ operation: "sum" as const, field: "计数项:影厅", outputName: "影厅数" }],
  sortBy: "影厅数",
  sortDirection: "desc" as const
};

const LIVE: SdkBridgeLiveContext = {
  workbookIdentity: "url:https://contoso.example/影院.xlsx",
  sourceFingerprint: "fp-1",
  authorizedSheetNames: ["Sheet1", "Sheet2"],
  activeWorksheet: "Sheet1"
};

interface Captured {
  url: string;
  body: Record<string, unknown>;
}

let captured: Captured[] = [];

/** 网络替身：按 URL 分流，同时记下每个请求的真实 body。 */
function stubNetwork(options: { poll?: unknown; receipt?: unknown } = {}) {
  captured = [];
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      const target = String(url);
      captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
      const payload = target.includes("/receipt")
        ? options.receipt ?? { status: "accepted", code: null }
        : options.poll ?? {
            calls: [
              {
                operationId: OP,
                tool: "query_table",
                arguments: QUERY_ARGUMENTS,
                requestFingerprint: FP,
                expiresAt: 0
              }
            ],
            generation: 1
          };
      return new Response(JSON.stringify(payload), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    })
  );
}

/**
 * 唯一被替换的生产环节：把 Office.js 读表换成读内存矩阵。
 * 计算走**既有**的 `executeQueryTableData`，不是预置结果。
 */
function hostExecute() {
  const computed: DataToolResult[] = [];
  const execute = vi.fn(
    async (
      request: DataToolRequest,
      allowedSheetNames: string[],
      activeWorksheet: string
    ): Promise<DataToolResult> => {
      const names =
        request.arguments.scope === "active" ? [activeWorksheet] : allowedSheetNames;
      for (const name of names) {
        if (WORKBOOK[name] === undefined) {
          throw new DataToolExecutionError("SCOPE_VIOLATION", "越界", false);
        }
      }
      const result = executeQueryTableData(
        request,
        names.map((name) => WORKBOOK[name])
      );
      computed.push(result);
      return result;
    }
  );
  return { execute, computed };
}

function beginOptions(live: () => SdkBridgeLiveContext = () => LIVE) {
  return {
    taskId: "turn-bridge-0001",
    workbookName: "影院.xlsx",
    workbookIdentity: "url:https://contoso.example/影院.xlsx",
    sourceFingerprint: "fp-1",
    liveContext: live
  };
}

/**
 * 可变的宿主现场：绑定发生在**前**，复核发生在**后**。
 * "授权中途收窄 / 中途换簿"必须这样建模 —— 拿同一个现场调两次是测不出东西的。
 */
function liveHolder(initial: SdkBridgeLiveContext = LIVE) {
  const holder = { current: initial };
  return {
    holder,
    read: () => holder.current
  };
}

const POINTER_EVENT = {
  stepId: "bridge.query_table",
  bridgeTool: { operationId: OP, tool: "query_table" }
};

function pollBody() {
  return captured.find((item) => item.url.includes("/poll"))?.body;
}

function receiptBody() {
  return captured.find((item) => item.url.includes("/receipt"))?.body;
}

beforeEach(() => {
  captured = [];
});

afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("useSdkToolBridge：指针 → 轮询 → 复核 → 真读真算 → 回执", () => {
  it("完整链路跑通，且回执带的是真算出来的分组结果", async () => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));

    await act(async () => { await result.current.begin(beginOptions()); });
    act(() => result.current.handleProgressEvent(POINTER_EVENT));

    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true));

    expect(pollBody()).toEqual({
      taskId: "turn-bridge-0001",
      hostId: "excel-office",
      workbookName: "影院.xlsx",
      workbookIdentity: "url:https://contoso.example/影院.xlsx",
      sourceFingerprint: "fp-1"
    });

    const receipt = receiptBody();
    expect(receipt).toMatchObject({
      taskId: "turn-bridge-0001",
      operationId: OP,
      requestFingerprint: FP,
      workbookIdentity: "url:https://contoso.example/影院.xlsx"
    });
    const payload = (receipt?.receipt as { result?: DataToolResult }).result;
    // 跨两张表的合并 + 分组求和，值是生产计算函数算出来的。
    expect(payload?.headers).toEqual(["影片名称", "影厅数"]);
    expect(payload?.rows).toEqual([
      ["功夫女足", 24],
      ["恐怖游轮", 3]
    ]);
    expect(payload?.sourceSheets).toEqual(["Sheet1", "Sheet2"]);
    expect(host.computed).toHaveLength(1);
  });

  it("绑定时的授权范围取自 liveContext（与执行前复核同源）", async () => {
    stubNetwork();
    const { result } = renderHook(() => useSdkToolBridge({ execute: hostExecute().execute }));
    await act(async () => { await result.current.begin(beginOptions()); });
    expect(result.current.binding()?.authorizedSheetNames).toEqual(["Sheet1", "Sheet2"]);
  });

  it("执行前授权收窄：**一次宿主都不读**，回执只带固定错误码", async () => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    const live = liveHolder();
    await act(async () => { await result.current.begin(beginOptions(live.read)); });
    // 用户在绑定之后撤销了一张表 —— 复核必须拿"当前授权"跟绑定范围比。
    live.holder.current = { ...LIVE, authorizedSheetNames: ["Sheet1"] };

    act(() => result.current.handleProgressEvent(POINTER_EVENT));

    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true));
    expect(host.execute).not.toHaveBeenCalled();
    expect(receiptBody()?.receipt as Record<string, unknown>).toEqual({
      errorCode: "unauthorized_scope"
    });
    // 错误码之外的任何宿主细节都不外发。
    expect(JSON.stringify(receiptBody())).not.toContain("越界");
  });

  it("换簿（同名不同身份）同样拒读", async () => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    const live = liveHolder();
    await act(async () => { await result.current.begin(beginOptions(live.read)); });
    live.holder.current = {
      ...LIVE,
      workbookIdentity: "url:https://evil.example/影院.xlsx"
    };

    act(() => result.current.handleProgressEvent(POINTER_EVENT));

    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true));
    expect(host.execute).not.toHaveBeenCalled();
    expect(receiptBody()?.receipt as Record<string, unknown>).toEqual({
      errorCode: "source_changed"
    });
  });

  it("源数据中途变化同样拒读（不会拿新快照的值顶旧授权）", async () => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    const live = liveHolder();
    await act(async () => { await result.current.begin(beginOptions(live.read)); });
    live.holder.current = { ...LIVE, sourceFingerprint: "fp-2" };

    act(() => result.current.handleProgressEvent(POINTER_EVENT));

    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true));
    expect(host.execute).not.toHaveBeenCalled();
    expect(receiptBody()?.receipt as Record<string, unknown>).toEqual({
      errorCode: "source_changed"
    });
  });

  it("同一个 operationId 再来一次指针不会重复执行", async () => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    await act(async () => { await result.current.begin(beginOptions()); });

    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    await waitFor(() => expect(host.execute).toHaveBeenCalledTimes(1));
    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    await act(async () => {
      await Promise.resolve();
      await Promise.resolve();
    });
    // 服务端已收下该 operationId，本地状态机短路，不再读一次宿主。
    expect(host.execute).toHaveBeenCalledTimes(1);
  });

  it("取消后不再轮询，后续指针也不会读宿主", async () => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    await act(async () => { await result.current.begin(beginOptions()); });

    act(() => {
      result.current.cancel();
      result.current.handleProgressEvent(POINTER_EVENT);
    });
    await act(async () => {
      await Promise.resolve();
    });

    expect(host.execute).not.toHaveBeenCalled();
    expect(captured.filter((c) => c.url.includes("/poll"))).toHaveLength(0);
  });

  it("轮内取消（进度事件 status=cancelled）就地停手", async () => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    await act(async () => { await result.current.begin(beginOptions()); });

    act(() => {
      result.current.handleProgressEvent({ stepId: "bridge.query_table", status: "cancelled" });
      result.current.handleProgressEvent(POINTER_EVENT);
    });
    await act(async () => {
      await Promise.resolve();
    });
    expect(host.execute).not.toHaveBeenCalled();
  });

  it("end() 停轮询但保留绑定：终局结果卡还要靠它按引用取值", async () => {
    stubNetwork();
    const { result } = renderHook(() => useSdkToolBridge({ execute: hostExecute().execute }));
    await act(async () => { await result.current.begin(beginOptions()); });
    act(() => result.current.end());
    expect(result.current.binding()).toMatchObject({
      taskId: "turn-bridge-0001",
      workbookIdentity: "url:https://contoso.example/影院.xlsx"
    });
  });

  it("执行失败时回执只带固定错误码，不带异常文本", async () => {
    stubNetwork();
    const failing = vi.fn(async () => {
      throw new DataToolExecutionError("SCOPE_VIOLATION", "敏感细节", false);
    });
    const { result } = renderHook(() => useSdkToolBridge({ execute: failing }));
    await act(async () => { await result.current.begin(beginOptions()); });

    act(() => result.current.handleProgressEvent(POINTER_EVENT));

    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true));
    expect((receiptBody()?.receipt as Record<string, unknown>)).toEqual({
      errorCode: "unauthorized_scope"
    });
    expect(JSON.stringify(receiptBody())).not.toContain("敏感细节");
  });
});

describe("真实组件点击触发链路", () => {
  it("点击按钮 → 发出真实请求体 → 真算分组结果 → 回执原样带上", async () => {
    stubNetwork();
    const host = hostExecute();

    function Harness() {
      const bridge = useSdkToolBridge({ execute: host.execute });
      return (
        <button
          onClick={() => {
            void (async () => {
              await bridge.begin(beginOptions());
              bridge.handleProgressEvent(POINTER_EVENT);
            })();
          }}
        >
          开始分析
        </button>
      );
    }

    const ui = render(<Harness />);
    fireEvent.click(ui.getByText("开始分析"));
    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true));

    // 真实发出的轮询请求体带齐了归属，回执带的是真算出来的行。
    expect(pollBody()).toMatchObject({ taskId: "turn-bridge-0001", hostId: "excel-office" });
    const payload = (receiptBody()?.receipt as { result?: DataToolResult }).result;
    expect(payload?.rows).toEqual([
      ["功夫女足", 24],
      ["恐怖游轮", 3]
    ]);
    expect(host.execute).toHaveBeenCalledTimes(1);
  });
});
