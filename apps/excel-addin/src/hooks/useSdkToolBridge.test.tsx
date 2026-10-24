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

function beginOptions(
  live: () => SdkBridgeLiveContext | Promise<SdkBridgeLiveContext> = () => LIVE
) {
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

/** 最后一条回执：失败重试/接续之后，只有它代表"这一轮真的交付了什么"。 */
function lastReceiptBody() {
  return captured.filter((item) => item.url.includes("/receipt")).at(-1)?.body;
}

beforeEach(() => {
  captured = [];
});

afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("useSdkToolBridge：指针 → 轮询 → 复核 → 真读真算 → 回执", () => {
  it("pins a deferred source once without reviving or rebinding another task", async () => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    await act(async () => {
      await result.current.begin({ ...beginOptions(), sourceFingerprint: null });
    });
    expect(result.current.pinSourceFingerprint("other-task", LIVE)).toBe(false);
    expect(result.current.pinSourceFingerprint("turn-bridge-0001", { ...LIVE, workbookIdentity: "other" })).toBe(false);
    expect(result.current.pinSourceFingerprint("turn-bridge-0001", { ...LIVE, authorizedSheetNames: ["Sheet1"] })).toBe(false);
    expect(result.current.pinSourceFingerprint("turn-bridge-0001", LIVE)).toBe(true);
    expect(result.current.pinSourceFingerprint("turn-bridge-0001", { ...LIVE, sourceFingerprint: "changed" })).toBe(false);
    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true));
    expect(pollBody()?.sourceFingerprint).toBe("fp-1");
    expect(host.execute).toHaveBeenCalledTimes(1);
    act(() => result.current.end());
    expect(result.current.pinSourceFingerprint("turn-bridge-0001", LIVE)).toBe(false);
    expect(result.current.isLive()).toBe(false);
  });

  it.each(["unreadable", "empty"] as const)("does not report bridge readiness with %s authorization", async (mode) => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    await act(async () => { await result.current.begin(beginOptions()); });
    let ready = true;
    await act(async () => {
      ready = await result.current.begin(beginOptions(() => {
        if (mode === "unreadable") throw new Error("private host detail");
        return { ...LIVE, authorizedSheetNames: [] };
      }));
    });
    expect(ready, "unreadable authorization must stop the planning request").toBe(false);
    expect(result.current.isLive()).toBe(false);
    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    expect(captured).toHaveLength(0);
    expect(host.execute).not.toHaveBeenCalled();
  });

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

/**
 * 桥的**生命周期**（2026-10-20）。
 *
 * 这一组盯的是两个哑失败：
 *
 * 1. `end()` 之后绑定还在（结果卡要按引用取值），但桥已经不能服务。只看
 *    "binding 是否存在"就把新指针交给它 —— 表现是"继续分析之后一直不动"。
 * 2. 旧一轮的 drain 跑得慢，新一轮 `begin` 已经建好绑定；旧 drain 回来把
 *    `stateRef.current` 整个覆盖掉 —— 新一轮从此收不到任何指针。
 */
describe("useSdkToolBridge：收尾之后必须重新立桥，旧事件不得串轮", () => {
  it("end() 之后绑定还在但桥不再服务；ensureLive 用同一任务重建后才继续读表", async () => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));

    await act(async () => { await result.current.begin(beginOptions()); });
    // 一轮结束：桥收尾，但绑定**刻意保留**（结果卡按引用取值要用）。
    act(() => result.current.end());
    expect(result.current.binding()?.taskId).toBe("turn-bridge-0001");
    expect(result.current.isLive()).toBe(false);

    // 把指针交给一条已收尾的桥：不许轮询、不许读表。
    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    await act(async () => { await Promise.resolve(); });
    expect(captured.some((c) => c.url.includes("/poll"))).toBe(false);
    expect(host.execute).not.toHaveBeenCalled();

    // 重新立桥（同一个 taskId：续跑不生成新身份绕过原账本）。
    await act(async () => {
      const ok = await result.current.ensureLive(beginOptions());
      expect(ok).toBe(true);
    });
    expect(result.current.isLive()).toBe(true);

    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true));
    expect(host.execute).toHaveBeenCalledTimes(1);
  });

  it("同一任务连续 ensureLive 只初始化一次（重建会清空 delivered，同一次调用被再读一遍）", async () => {
    stubNetwork();
    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));

    await act(async () => {
      await result.current.ensureLive(beginOptions());
      await result.current.ensureLive(beginOptions());
      await result.current.ensureLive(beginOptions());
    });
    const binding = result.current.binding();
    expect(binding?.taskId).toBe("turn-bridge-0001");

    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    await waitFor(() => expect(captured.some((c) => c.url.includes("/receipt"))).toBe(true));
    // 同一个 operationId 只执行一次。
    expect(host.execute).toHaveBeenCalledTimes(1);
    expect(result.current.binding()).toBe(binding);
  });

  it("旧一轮的 drain 迟到返回时不得覆盖新一轮的绑定（变异点：提交前不校验代数）", async () => {
    // 第一次轮询卡住，等我们把新一轮立好再放行。
    let releasePoll: (() => void) | null = null;
    const gate = new Promise<void>((resolve) => { releasePoll = resolve; });
    captured = [];
    vi.stubGlobal(
      "fetch",
      vi.fn(async (url: string, init?: RequestInit) => {
        const target = String(url);
        captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
        if (target.includes("/poll")) {
          await gate;
          return new Response(
            JSON.stringify({
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
            }),
            { status: 200, headers: { "Content-Type": "application/json" } }
          );
        }
        return new Response(JSON.stringify({ status: "accepted", code: null }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      })
    );

    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));

    await act(async () => { await result.current.begin(beginOptions()); });
    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    // 旧一轮已经进入 poll，并卡在网络上。
    await act(async () => { await Promise.resolve(); });

    // 这一轮收尾，紧接着为**另一个任务**立新桥。
    act(() => result.current.end());
    await act(async () => {
      await result.current.ensureLive({ ...beginOptions(), taskId: "turn-bridge-0002" });
    });
    expect(result.current.binding()?.taskId).toBe("turn-bridge-0002");

    // 放行旧轮询：它的结果**不得**回写，也不得执行读表。
    await act(async () => {
      releasePoll?.();
      await Promise.resolve();
      await Promise.resolve();
    });

    expect(result.current.binding()?.taskId).toBe("turn-bridge-0002");
    // 新一轮仍然是活的：旧一轮的收尾**没有**把它的 live 状态带走。
    expect(result.current.isLive()).toBe(true);
    expect(host.execute).not.toHaveBeenCalled();
    expect(captured.some((c) => c.url.includes("/receipt"))).toBe(false);
  });
});

/**
 * 生命周期**竞态**（2026-10-20 第三批）。
 *
 * 前一组钉的是"旧结果不许覆盖新轮次"，这一组钉的是三件更隐蔽的事：
 *
 * 1. 旧轮次**退出时**，新一轮在它运行期间排过队的指针没人接 —— 桥是活的，但
 *    一个请求都不发；表现是"继续分析之后要再点一次/再等一个事件才动"。
 * 2. 初始化（`await liveContext()`）期间被取消/收尾，迟到的初始化把桥又点亮了。
 * 3. 两次初始化交错完成，先发起的那个后回来，把新轮次覆盖掉。
 *
 * 三条都只用可控 Promise 定序，不靠 sleep 碰概率。
 */
describe("useSdkToolBridge：轮次接续 / 迟到初始化作废", () => {
  it("旧轮次在途时到达的指针，旧轮次退出后必须被接上（不依赖第二个进度事件）", async () => {
    // 旧轮次的 poll 卡在网络里；新轮次的 poll 正常返回同一个待执行调用。
    let releaseOldPoll: (() => void) | null = null;
    const oldPollGate = new Promise<void>((resolve) => { releaseOldPoll = resolve; });
    let polls = 0;
    captured = [];
    vi.stubGlobal(
      "fetch",
      vi.fn(async (url: string, init?: RequestInit) => {
        const target = String(url);
        captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
        if (target.includes("/poll")) {
          polls += 1;
          const generation = polls;
          if (generation === 1) await oldPollGate;
          return new Response(
            JSON.stringify({
              calls: [
                {
                  operationId: OP,
                  tool: "query_table",
                  arguments: QUERY_ARGUMENTS,
                  requestFingerprint: FP,
                  expiresAt: 0
                }
              ],
              generation
            }),
            { status: 200, headers: { "Content-Type": "application/json" } }
          );
        }
        return new Response(JSON.stringify({ status: "accepted", code: null }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      })
    );

    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));

    await act(async () => { await result.current.begin(beginOptions()); });
    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    // 旧轮次已经进入 poll 并被闸住。
    await act(async () => { await Promise.resolve(); });

    // 这一轮收尾，紧接着为**同一个任务**重新立桥（续跑正是这条路径）。
    act(() => result.current.end());
    await act(async () => { await result.current.ensureLive(beginOptions()); });
    // 新一轮**唯一**一个指针：此刻旧 drain 还占着 runningRef，这次只会被记账。
    act(() => result.current.handleProgressEvent(POINTER_EVENT));

    // 放行旧 poll：旧轮次退出，出口必须把新一轮排队的活儿接过去。
    await act(async () => {
      releaseOldPoll?.();
      await Promise.resolve();
      await Promise.resolve();
      await Promise.resolve();
    });

    await waitFor(() => expect(host.execute).toHaveBeenCalledTimes(1));
    // 两次轮询：旧轮次一次（被判失效、不执行），新轮次一次（真的读了表并回执）。
    expect(polls).toBe(2);
    expect(captured.filter((item) => item.url.includes("/receipt"))).toHaveLength(1);
  });

  it("初始化期间 cancel：迟到的初始化不得把桥重新激活", async () => {
    stubNetwork();
    const host = hostExecute();
    let releaseLate: (() => void) | null = null;
    const gate = new Promise<void>((resolve) => { releaseLate = resolve; });
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));

    let pending: Promise<boolean> | null = null;
    await act(async () => {
      pending = result.current.begin(
        beginOptions(async () => {
          await gate;
          return LIVE;
        })
      );
      await Promise.resolve();
    });
    // 初始化还没回来：桥当然还不该是活的。
    expect(result.current.isLive()).toBe(false);

    act(() => result.current.cancel());
    await act(async () => {
      releaseLate?.();
      await pending;
    });

    // 取消之后再回来的初始化**不得**写绑定、更不得把 live 点亮。
    expect(result.current.isLive()).toBe(false);
    expect(result.current.binding()).toBeNull();
    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    await act(async () => { await Promise.resolve(); });
    expect(host.execute).not.toHaveBeenCalled();
  });

  it("初始化期间 end()：迟到的初始化同样作废，绑定不得被写进来", async () => {
    stubNetwork();
    const { result } = renderHook(() => useSdkToolBridge({ execute: hostExecute().execute }));
    let releaseLate: (() => void) | null = null;
    const gate = new Promise<void>((resolve) => { releaseLate = resolve; });

    let pending: Promise<boolean> | null = null;
    await act(async () => {
      pending = result.current.begin(
        beginOptions(async () => {
          await gate;
          return LIVE;
        })
      );
      await Promise.resolve();
    });

    act(() => result.current.end());
    await act(async () => {
      releaseLate?.();
      await pending;
    });

    expect(result.current.isLive()).toBe(false);
    // end() 之后本该"没有绑定"：迟到的初始化不许顺手塞一个进去。
    expect(result.current.binding()).toBeNull();
  });

  it("两次初始化交错：B 先完成、A 后完成，最终绑定必须属于 B", async () => {
    stubNetwork();
    const { result } = renderHook(() => useSdkToolBridge({ execute: hostExecute().execute }));
    let releaseA: (() => void) | null = null;
    const gateA = new Promise<void>((resolve) => { releaseA = resolve; });

    let first: Promise<boolean> | null = null;
    let second: Promise<boolean> | null = null;
    await act(async () => {
      first = result.current.begin({
        ...beginOptions(async () => {
          await gateA;
          return LIVE;
        }),
        taskId: "turn-bridge-A"
      });
      await Promise.resolve();
      second = result.current.begin({ ...beginOptions(), taskId: "turn-bridge-B" });
    });
    await act(async () => { await second; });
    expect(result.current.binding()?.taskId).toBe("turn-bridge-B");

    // A 后回来：它已经被 B 取代，不许覆盖，也不许自称成功。
    await act(async () => {
      releaseA?.();
      await first;
    });
    expect(await first).toBe(false);
    expect(result.current.binding()?.taskId).toBe("turn-bridge-B");
    expect(result.current.isLive()).toBe(true);
  });

  it("ensureLive 在初始化期间被取消时返回 false（不拿别的轮次的 live 冒充成功）", async () => {
    stubNetwork();
    const { result } = renderHook(() => useSdkToolBridge({ execute: hostExecute().execute }));
    let releaseLate: (() => void) | null = null;
    const gate = new Promise<void>((resolve) => { releaseLate = resolve; });

    let ensured: Promise<boolean> | null = null;
    await act(async () => {
      ensured = result.current.ensureLive(
        beginOptions(async () => {
          await gate;
          return LIVE;
        })
      );
      await Promise.resolve();
    });

    act(() => result.current.cancel());
    await act(async () => {
      releaseLate?.();
      await ensured;
    });

    expect(await ensured).toBe(false);
    expect(result.current.isLive()).toBe(false);
  });
});

/**
 * 生命周期**异常退出**（2026-10-20 第四批）。
 *
 * 上一组钉的是"旧轮次**正常返回**、发现自己失效之后仍要接续排队的新指针"。
 * 这一组钉的是同一件事的**另一半**：旧轮次不是正常返回，而是**抛出来**的
 * （取消 → AbortError、断网 → fetch reject、读表/回执自己失败）。控制流会从
 * `runBridgeUntilIdle` 一路抛到 drain 的外层 `catch`；那里只记账不兑现的话，
 * 新一轮的指针就一直躺在那儿 —— 桥是活的、不报错、也一个请求都不发。
 *
 * 后半段（"本轮自身失败"）是防**改过头**的闸：不能为了让异常退出也接续，
 * 就把"同一代里失败"也当成接续 —— 那等于把接续变成自动重试，而回执可能
 * 已经在路上，状态未知的请求不许擅自重放。
 *
 * 全部用可控 Promise 定序，不靠 sleep 碰概率。
 */
describe("useSdkToolBridge：旧轮次异常退出仍须接续，本轮自身失败不得重试", () => {
  /** 每一轮都返回同一个待执行调用的轮询响应。 */
  function callResponse(generation: number) {
    return new Response(
      JSON.stringify({
        calls: [
          {
            operationId: OP,
            tool: "query_table",
            arguments: QUERY_ARGUMENTS,
            requestFingerprint: FP,
            expiresAt: 0
          }
        ],
        generation
      }),
      { status: 200, headers: { "Content-Type": "application/json" } }
    );
  }

  function acceptedResponse() {
    return new Response(JSON.stringify({ status: "accepted", code: null }), {
      status: 200,
      headers: { "Content-Type": "application/json" }
    });
  }

  it("旧轮次轮询抛 AbortError：新一轮排队的指针仍要被接上（不依赖第二个事件）", async () => {
    let releaseOldPoll: (() => void) | null = null;
    const oldPollGate = new Promise<void>((resolve) => { releaseOldPoll = resolve; });
    let polls = 0;
    let receipts = 0;
    captured = [];
    vi.stubGlobal(
      "fetch",
      vi.fn(async (url: string, init?: RequestInit) => {
        const target = String(url);
        captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
        if (target.includes("/poll")) {
          polls += 1;
          if (polls === 1) {
            await oldPollGate;
            // 取消/断网：控制流从端口一路抛到 drain 的外层 catch。
            throw new DOMException("aborted", "AbortError");
          }
          return callResponse(polls);
        }
        receipts += 1;
        return acceptedResponse();
      })
    );

    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));

    await act(async () => { await result.current.begin(beginOptions()); });
    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    // 旧轮次已经进入 poll，并被闸住。
    await act(async () => { await Promise.resolve(); });

    // 这一轮收尾，紧接着为同一个任务重新立桥（续跑正是这条路径）。
    act(() => result.current.end());
    await act(async () => { await result.current.ensureLive(beginOptions()); });
    // 新一轮**唯一**一个指针：此刻旧 drain 还占着 runningRef，这次只会被记账。
    act(() => result.current.handleProgressEvent(POINTER_EVENT));

    // 放行旧轮询：它抛异常退出，出口必须把新一轮排队的活儿接过去。
    await act(async () => {
      releaseOldPoll?.();
      await Promise.resolve();
      await Promise.resolve();
      await Promise.resolve();
    });

    await waitFor(() => expect(host.execute).toHaveBeenCalledTimes(1));
    expect(polls).toBe(2);
    expect(receipts).toBe(1);
    // 旧轮次的异常**不是**新一轮的失败：不许记错误码，也不许把桥判死。
    expect(result.current.failure()).toBeNull();
    expect(result.current.isLive()).toBe(true);
    expect(result.current.binding()?.taskId).toBe("turn-bridge-0001");
  });

  /*
   * 为什么只测轮询与回执两处？
   *
   * `verify` 与 `execute` 两个端口**自己接住了异常**（读不到现场 → 拒绝回执；
   * 宿主异常 → 固定错误码回执，见 `verifyBridgeAuthorizationOrRefuse` 与
   * `createSdkBridgeHostPorts`），它们抛不到 drain。真正能把控制流抛进外层
   * `catch` 的只有两处网络调用：poll 与 receipt。两种都测，证明收尾是同一套。
   */
  it("旧轮次轮询以普通网络错误退出：排队的新指针同样要被接上", async () => {
    let releaseOldPoll: (() => void) | null = null;
    const oldPollGate = new Promise<void>((resolve) => { releaseOldPoll = resolve; });
    let polls = 0;
    let receipts = 0;
    captured = [];
    vi.stubGlobal(
      "fetch",
      vi.fn(async (url: string, init?: RequestInit) => {
        const target = String(url);
        captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
        if (target.includes("/poll")) {
          polls += 1;
          if (polls === 1) {
            await oldPollGate;
            // fetch 自己失败（断网）：TypeError，不是 AbortError。
            throw new TypeError("fetch failed");
          }
          return callResponse(polls);
        }
        receipts += 1;
        return acceptedResponse();
      })
    );

    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    await act(async () => { await result.current.begin(beginOptions()); });
    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    await act(async () => { await Promise.resolve(); });

    act(() => result.current.end());
    await act(async () => { await result.current.ensureLive(beginOptions()); });
    act(() => result.current.handleProgressEvent(POINTER_EVENT));

    await act(async () => {
      releaseOldPoll?.();
      await Promise.resolve();
      await Promise.resolve();
      await Promise.resolve();
    });

    await waitFor(() => expect(receipts).toBe(1));
    expect(polls).toBe(2);
    expect(host.execute).toHaveBeenCalledTimes(1);
    // 失败的那一轮没有留下"已交付"记录：新轮次的回执身份是这一轮的调用。
    expect((receiptBody()?.operationId as string)).toBe(OP);
    expect(result.current.isLive()).toBe(true);
  });

  it("旧轮次在回执处抛错：状态未知也不能把新轮次卡死", async () => {
    let releaseOldReceipt: (() => void) | null = null;
    const oldReceiptGate = new Promise<void>((resolve) => { releaseOldReceipt = resolve; });
    let polls = 0;
    let receiptAttempts = 0;
    let receipts = 0;
    captured = [];
    vi.stubGlobal(
      "fetch",
      vi.fn(async (url: string, init?: RequestInit) => {
        const target = String(url);
        captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
        if (target.includes("/poll")) {
          polls += 1;
          return callResponse(polls);
        }
        receiptAttempts += 1;
        if (receiptAttempts === 1) {
          await oldReceiptGate;
          throw new TypeError("network down");
        }
        receipts += 1;
        return acceptedResponse();
      })
    );

    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    await act(async () => { await result.current.begin(beginOptions()); });
    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    await waitFor(() => expect(receiptAttempts).toBe(1));

    act(() => result.current.end());
    await act(async () => { await result.current.ensureLive(beginOptions()); });
    act(() => result.current.handleProgressEvent(POINTER_EVENT));

    await act(async () => {
      releaseOldReceipt?.();
      await Promise.resolve();
      await Promise.resolve();
      await Promise.resolve();
    });

    await waitFor(() => expect(receipts).toBe(1));
    expect(polls).toBe(2);
    // 新一轮的回执**带的是真算出来的结果**，不是"未知"。
    expect((lastReceiptBody()?.receipt as { result?: DataToolResult }).result?.rows).toEqual([
      ["功夫女足", 24],
      ["恐怖游轮", 3]
    ]);
    expect(result.current.isLive()).toBe(true);
  });

  it("同一代里回执失败：不自动重试、不擅自重放，下一条真指针才重新轮询", async () => {
    const OP2 = "op_" + "3".repeat(32);
    let polls = 0;
    let receiptAttempts = 0;
    let receipts = 0;
    captured = [];
    vi.stubGlobal(
      "fetch",
      vi.fn(async (url: string, init?: RequestInit) => {
        const target = String(url);
        const body = JSON.parse(String(init?.body ?? "{}")) as Record<string, unknown>;
        captured.push({ url: target, body });
        if (target.includes("/poll")) {
          polls += 1;
          // 第二次轮询之前，服务端已经把这条调用撤掉了（本地就不能再执行它）。
          return new Response(
            JSON.stringify({
              calls:
                polls === 1
                  ? [
                      {
                        operationId: OP,
                        tool: "query_table",
                        arguments: QUERY_ARGUMENTS,
                        requestFingerprint: FP,
                        expiresAt: 0
                      }
                    ]
                  : [
                      {
                        operationId: OP2,
                        tool: "query_table",
                        arguments: QUERY_ARGUMENTS,
                        requestFingerprint: FP,
                        expiresAt: 0
                      }
                    ],
              generation: polls
            }),
            { status: 200, headers: { "Content-Type": "application/json" } }
          );
        }
        receiptAttempts += 1;
        if (receiptAttempts === 1) throw new TypeError("network down");
        receipts += 1;
        return acceptedResponse();
      })
    );

    const host = hostExecute();
    const { result } = renderHook(() => useSdkToolBridge({ execute: host.execute }));
    await act(async () => { await result.current.begin(beginOptions()); });

    act(() => result.current.handleProgressEvent(POINTER_EVENT));
    await waitFor(() => expect(receiptAttempts).toBe(1));
    // 失败发生在同一代里：等几拍，确认**没有**自动再轮询一次。
    await act(async () => {
      await Promise.resolve();
      await Promise.resolve();
      await Promise.resolve();
    });
    expect(polls).toBe(1);
    expect(receiptAttempts).toBe(1);

    // 真的来了下一条指针（另一个 operationId）才重新开始：先问服务端，再执行。
    act(() =>
      result.current.handleProgressEvent({
        stepId: "bridge.query_table",
        bridgeTool: { operationId: OP2, tool: "query_table" }
      })
    );
    await waitFor(() => expect(receipts).toBe(1));
    expect(polls).toBe(2);
    // 旧的那条（回执状态未知）没有被本地重放：只回执了 OP2。
    expect(lastReceiptBody()?.operationId as string).toBe(OP2);
    expect(host.execute).toHaveBeenCalledTimes(2);
  });
});
