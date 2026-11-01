// @vitest-environment jsdom
/**
 * CAP-02 暂停卡的面板契约测试。
 *
 * 为什么单独开一个文件：`WpsAgentPanel.test.tsx` 体量已经很大，而这里要钉的是
 * **暂停卡独有的两条纪律**：
 *
 *  1. **状态与操作分离**：状态文字在 role="status" 节点里，操作区只放 `<button>`。
 *  2. **只有真正的"可继续"才画按钮**：`canContinue=false` 或没有检查点时，
 *     卡片必须只剩状态说明 —— 一个"看起来能点、点了其实会重复规划"的按钮
 *     比没有按钮危险得多。
 *
 * 这两条都不是样式问题，而是"用户会不会误以为能从上次进度精确续跑"。
 */
import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { listModels, getModelSettings, understandMessage } from "../../excel-addin/src/api";
import { getWpsApplication } from "./wpsHost";
import { createWpsBridgeLiveContext } from "./wpsSdkBridgeHost";
import { readAuthorizedWpsSheets, readAuthorizedWpsSheetSchemas } from "./wpsWorkbookReader";
import { executeQueryTableData } from "../../excel-addin/src/dataTools";
import type { WpsTurnResult } from "./wpsReadonlyAgent";

/**
 * 只桩掉真正发起网络的那个函数：`runWpsReadonlyTurn`。其余（`wpsPlanningResumeView`
 * 等纯函数）保持真实实现 —— 本文件要断言的正是它们的**输出如何被渲染**，
 * 桩掉它们等于把被测行为一起换掉，断言就恒真了。
 */
vi.mock("./wpsReadonlyAgent", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsReadonlyAgent")>();
  return {
    ...actual,
    runWpsReadonlyTurn: vi.fn()
  };
});

/**
 * 只把**网络缝**桩掉：`getWpsPlanningTaskStatus` 是 `reinspectWpsPlanningPause`
 * 唯一的外部调用。
 *
 * 为什么不直接桩 `reinspectWpsPlanningPause`：那会连带换掉
 * `wpsPlanningResumeView`（由 status 推导按钮）—— 于是"核对结论怎么决定
 * 按钮画不画"这段逻辑完全没被走到，桩住返回值等于同时改掉被测行为。
 */
vi.mock("./wpsPlanningStream", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsPlanningStream")>();
  return {
    ...actual,
    getWpsPlanningTaskStatus: vi.fn(),
    // 续跑的网络缝。单独桩出来是为了让"续跑期间的读表指针"能由用例决定，
    // 断言的仍然是面板有没有把指针交给桥。
    continueWpsPlanning: vi.fn(),
    /*
     * 只读补发历史进度的**另一个**网络缝（`GET /api/turn/{id}/events`）。
     * 以前没桩，于是"重开窗格核对暂停任务"这类用例会真的去读用户那台 8765 的轮次事件：
     * 读不读得到取决于对方在不在跑。这里固定成"读不到"（生产里走的就是
     * `catch ⇒ timeout`：读不到不等于任务结束），需要事件的用例自己给实现。
     */
    getWpsPlanningTaskEvents: vi.fn(async () => { throw new TypeError("Failed to fetch"); })
  };
});

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  const discovery = await import("./wpsPendingDiscovery.testHost");
  /*
   * 「继续分析的读表桥」用例自己 `vi.stubGlobal("fetch", …)` 驱动真实轮询，所以桥与
   * 暂停/校验这几条接缝用 `guardNetworkSeam`：装了 fetch 替身才真的调用，否则按
   * "连不上"失败。漏出去的请求另有 `testSetup.ts` 的闸门当场判失败。
   */
  const gate = await import("@excel-bro/core/testNetworkGate");
  return {
    ...actual, ...discovery.pendingDiscoveryApiStubs(),
    pollSdkBridgeTool: vi.fn(gate.guardNetworkSeam(actual.pollSdkBridgeTool)),
    pausePlanningTask: vi.fn(gate.guardNetworkSeam(actual.pausePlanningTask)),
    validatePlan: vi.fn(gate.guardNetworkSeam(actual.validatePlan)),
    understandMessage: vi.fn(async () => ({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" })),
    listModels: vi.fn(),
    getModelSettings: vi.fn()
  };
});

vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

/**
 * 读表桥外面套一层**可控返回值**（2026-10-20）。
 *
 * 桥自己的行为由 `hooks/useSdkToolBridge.test.tsx` 的 23 条用例 + 四条变异证明，
 * 这里要证的是 WPS 面板的**调用层**：拿到 `ensureLive` 的 `false` 之后到底还发
 * 不发续跑请求。替身只在开关打开时替换返回值，其余一律转发给真实 hook。
 */
const bridgeStub = vi.hoisted(() => ({ begin: false, ensureLive: false }));

vi.mock("../../excel-addin/src/hooks/useSdkToolBridge", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/hooks/useSdkToolBridge")>();
  const React = await import("react");
  type Controller = ReturnType<typeof actual.useSdkToolBridge>;
  return {
    ...actual,
    useSdkToolBridge: (dependencies?: Parameters<typeof actual.useSdkToolBridge>[0]) => {
      const real = actual.useSdkToolBridge(dependencies);
      const latest = React.useRef<Controller>(real);
      latest.current = real;
      return React.useMemo(
        () =>
          new Proxy({} as Controller, {
            get(_target, key: string) {
              if (key === "begin" && bridgeStub.begin) return async () => false;
              if (key === "ensureLive" && bridgeStub.ensureLive) return async () => false;
              const value = (latest.current as unknown as Record<string, unknown>)[key];
              return typeof value === "function" ? value.bind(latest.current) : value;
            }
          }),
        []
      );
    }
  };
});

/**
 * 面板是默认导出，必须在 mock 之后动态引入 —— 静态 import 会被提升到
 * `vi.mock` 之前，于是面板拿到的是真实实现。
 */
const { default: WpsAgentPanel } = await import("./WpsAgentPanel");
const { runWpsReadonlyTurn } = await import("./wpsReadonlyAgent");
const { getWpsPlanningTaskStatus, continueWpsPlanning } = await import("./wpsPlanningStream");

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const listModelsMock = vi.mocked(listModels);
const getModelSettingsMock = vi.mocked(getModelSettings);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);
const statusMock = vi.mocked(getWpsPlanningTaskStatus);
const continueMock = vi.mocked(continueWpsPlanning);
const getWpsApplicationMock = vi.mocked(getWpsApplication);

const connection = {
  id: "connection:first",
  label: "第一个连接",
  provider: "model" as const,
  available: true,
  supportsVision: false
};

let mounted: Array<{ container: HTMLElement; unmount: () => void }> = [];

async function mount(element: ReactElement) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(element);
  });
  const entry = {
    container,
    unmount: () => {
      act(() => root.unmount());
      container.remove();
    }
  };
  mounted.push(entry);
  return entry;
}

async function settle(): Promise<void> {
  await act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  });
}

async function until(check: () => boolean, message: string): Promise<void> {
  for (let index = 0; index < 80; index += 1) {
    if (check()) return;
    await settle();
  }
  throw new Error(message);
}

async function type(element: HTMLTextAreaElement | null, value: string) {
  if (!element) throw new Error("没有输入框");
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")?.set;
  await act(async () => {
    setter?.call(element, value);
    element.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

async function submit(element: HTMLTextAreaElement | null): Promise<void> {
  const form = element?.closest("form");
  if (!form) throw new Error("没有输入区表单");
  await act(async () => {
    form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
  });
  await settle();
}

function textOf(container: HTMLElement): string {
  return container.textContent ?? "";
}

function findButton(container: HTMLElement, label: string): HTMLButtonElement | null {
  /*
   * 按**可访问名**找：主操作按钮（输入框右下角那一颗）只画字形，动作名在
   * `aria-label` 上。按钮位置变了，查找方式跟着变；断言本身一个字没动。
   */
  return Array.from(container.querySelectorAll("button")).find(
    (button) =>
      (button.textContent ?? "").trim() === label || button.getAttribute("aria-label") === label
  ) ?? null;
}

/** 能让面板通过"目标确认"的最小宿主：必须是一个可读取的真实工作簿对象。 */
function pauseHost() {
  const usedRange = { Rows: { Count: 3 }, Columns: { Count: 2 }, Row: 1, Column: 1 } as Record<string, unknown>;
  Object.defineProperties(usedRange, {
    Value2: { get: () => [["类别", "归属"], ["甲", "华东"], ["乙", "华南"]] },
    Item: { value: () => usedRange },
    MergeCells: { value: false }
  });
  const range = { Rows: { Count: 1 }, Columns: { Count: 2 }, MergeCells: false } as Record<string, unknown>;
  Object.defineProperties(range, { Value2: { get: () => [["类别", "归属"]] } });
  const sheet = {
    Name: "Sheet1",
    Index: 1,
    CodeName: "SheetCode1",
    ProtectContents: false,
    UsedRange: usedRange,
    Range: () => range
  };
  const storage = new Map<string, unknown>();
  return {
    Version: "12.0",
    ActiveWorkbook: {
      Name: "测试.xlsx",
      FullName: "C:\\Data\\测试.xlsx",
      Path: "C:\\Data",
      ReadOnly: false,
      Worksheets: { Count: 1, Item: () => sheet },
      ActiveSheet: sheet
    },
    ActiveSheet: sheet,
    Selection: { Address: () => "$A$1" },
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, item: unknown) => { storage.set(key, item); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never;
}

function pauseResult(overrides: Partial<Record<string, unknown>> = {}) {
  return {
    kind: "planning_pause" as const,
    turnId: "turn-pause-1",
    status: "waiting_user" as const,
    revision: 4,
    lastSequence: 9,
    reason: "本轮处理时间已达上限。",
    activeElapsedSeconds: 120,
    approvedWallTimeSeconds: 120,
    extensionSeconds: 60,
    extensionCount: 1,
    maximumExtensions: 3,
    canContinue: true,
    requiresExtension: true,
    pauseType: "budget" as const,
    currentModelId: null,
    modelSwitchCount: 0,
    canSwitchModel: false,
    availableModelConnections: [],
    // 服务端给的继续资格：默认"可以继续，停在工具回执后"。
    continuation: { available: true, code: null, message: "", stage: "after_tool" as const },
    ...overrides
  };
}

function panel() {
  return (
    <WpsAgentPanel
      selectedSheetIndexes={[1]}
      scopeDock={<div className="eb-scope-dock">当前表 Sheet1</div>}
    />
  );
}

/** 让面板加载到"可以发送"的程度，然后发一条需求并等暂停卡出现。 */
async function reachPauseCard(pause: ReturnType<typeof pauseResult>) {
  listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
  getModelSettingsMock.mockResolvedValue({} as never);
  getWpsApplicationMock.mockReturnValue(pauseHost());
  runTurnMock.mockResolvedValue({ kind: "planning_pause", pause } as WpsTurnResult);
  const { container } = await mount(panel());
  await until(
    () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false,
    "应加载模型目录"
  );
  const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
  await type(textarea, "生成统计预览");
  await submit(textarea);
  await until(() => Boolean(container.querySelector(".wps-planning-pause-card")), "应出现规划暂停卡");
  return { container };
}

beforeEach(() => {
  mounted = [];
  localStorage.clear();
  vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" });
  bridgeStub.begin = false;
  bridgeStub.ensureLive = false;
  // 默认"服务端说可以继续"：单条用例按需覆盖成"已完成 / 读不到"。
  statusMock.mockResolvedValue({
    turnId: "turn-pause-1",
    status: "waiting_user",
    revision: 12,
    lastSequence: 30,
    canContinue: true,
    pauseType: "budget",
    reason: null
  });
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * 模拟"关掉任务窗格再打开"。
 *
 * 全程走真实交互拿到落盘历史，再卸载面板 —— 重开后一切都只能从
 * localStorage 重建，内存里的活跃暂停、`planningPauseReload` 都不复存在。
 * 这正是"刷新后恢复"要处理的场景。
 */
async function seedHistoryThenReload() {
  await reachPauseCard(pauseResult());
  const stored = localStorage.getItem("excel-bro.wps.chat.v1");
  expect(stored, "落盘历史应已写入").not.toBeNull();
  // 卸载面板 = 关掉并重开任务窗格。
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  localStorage.setItem("excel-bro.wps.chat.v1", stored!);
  runTurnMock.mockClear();
  statusMock.mockClear();
  statusMock.mockResolvedValue({
    turnId: "turn-pause-1",
    status: "waiting_user",
    revision: 12,
    lastSequence: 30,
    canContinue: true,
    pauseType: "budget",
    reason: null
  });
  getWpsApplicationMock.mockReturnValue(pauseHost());
  return stored!;
}

describe("CAP-02 规划暂停卡", () => {
  it.each((["message", "reading", "fields"] as const).flatMap((stage) =>
    (["text", "button"] as const).map((via) => ({ stage, via }))
  ))("早期暂停 $stage 后通过 $via 接回原需求，不再选择已消费的恢复任务", async ({ stage, via }) => {
    listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
    getModelSettingsMock.mockResolvedValue({} as never);
    getWpsApplicationMock.mockReturnValue(pauseHost());
    const blocked = (_request: unknown, signal?: AbortSignal) => new Promise<never>((_resolve, reject) => {
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    });
    if (stage === "message") vi.mocked(understandMessage).mockImplementationOnce(blocked);
    else runTurnMock.mockImplementationOnce((request) => {
      request.onProgress?.({ stage: "reading", stepId: "read-schemas", title: "读取已授权表结构", message: "读取字段" });
      if (stage === "fields") request.onProgress?.({ stage: "intent", stepId: "intent", title: "结合字段理解任务", message: "分析字段" });
      return blocked(request, request.signal);
    });
    const { container } = await mount(panel());
    await until(() => Boolean(container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型")), "模型已加载");
    await type(container.querySelector(".composer textarea"), "按会员等级汇总订单并列出异常");
    await submit(container.querySelector(".composer textarea"));
    await until(() => stage === "message" ? vi.mocked(understandMessage).mock.calls.length > 0 : runTurnMock.mock.calls.length > 0, "进入可控的暂停阶段");
    const pause = [...container.querySelectorAll<HTMLButtonElement>("button")]
      .find((button) => /^(暂停|暂停任务)$/.test(button.getAttribute("aria-label") ?? button.textContent ?? ""));
    expect(pause).toBeTruthy();
    await act(async () => pause!.click());
    await until(() => Boolean(findButton(container, "继续分析")), "早期暂停也必须保留原需求入口");
    // 暂停那一刻就取快照：后面要逐字比较"历史段有没有被这一轮改写"。
    expect(container.querySelectorAll(".task-progress-timeline")).toHaveLength(1);
    const pausedSteps = [...container.querySelectorAll<HTMLElement>(".activity-steps > li")]
      .map((row) => row.textContent);
    // The restarted request must not offer its already consumed recovery ID.
    vi.mocked(understandMessage).mockImplementation(async (request) => request.recoveryTasks?.length
      ? { kind: "task_recovery", turnId: request.recoveryTasks![0], action: "continue", provider: "model" }
      : { kind: "needs_workbook_context", summary: request.prompt, provider: "model" });
    runTurnMock.mockImplementation(async (request) => {
      request.onProgress?.({ stage: "reading", stepId: "read-schemas", title: "读取已授权表结构", message: "复核字段" });
      request.onProgress?.({ stage: "intent", stepId: "intent", title: "结合字段理解任务", message: "继续分析" });
      return { kind: "answer", message: "原需求重新分析已接通", provider: "model" } as WpsTurnResult;
    });
    const semanticCallsBeforeRestart = vi.mocked(understandMessage).mock.calls.length;
    if (via === "text") {
      await type(container.querySelector(".composer textarea"), "继续");
      await submit(container.querySelector(".composer textarea"));
    } else {
      expect(container.querySelector(".analysis-pause-inline")).not.toBeNull();
      expect(container.textContent).not.toContain("尚未生成规划检查点");
      await act(async () => container.querySelector<HTMLButtonElement>('button[aria-label="继续任务"], [aria-label="继续分析"]')!.click());
    }
    await until(() => textOf(container).includes("原需求重新分析已接通"), "继续必须实际重入原需求");
    expect(understandMessage).toHaveBeenCalledTimes(semanticCallsBeforeRestart + (via === "text" ? 1 : 0) + (stage === "message" ? 1 : 0));
    if (stage === "message") {
      expect(vi.mocked(understandMessage).mock.calls.at(-1)?.[0].prompt).toBe("按会员等级汇总订单并列出异常");
      expect(vi.mocked(understandMessage).mock.calls.at(-1)?.[0].recoveryTasks).toBeUndefined();
    }
    /*
     * 判别点（缺陷 1）：继续开的是**新的回复段**。
     * 旧段逐字停在暂停那一刻（结论 awaiting_input + 「查看后续进展」），新段才是
     * 这一轮的实际进展 —— 不是把旧卡就地改写成"又跑起来了"。
     */
    const cards = [...container.querySelectorAll<HTMLElement>(".task-progress-timeline")];
    expect(cards).toHaveLength(2);
    expect(
      [...cards[0].querySelectorAll(".activity-steps > li")].map((row) => row.textContent)
    ).toEqual(pausedSteps);
    expect(cards[0].getAttribute("data-status")).toBe("awaiting_input");
    expect(cards[0].textContent).toContain("查看后续进展");
    expect(cards[0].textContent).not.toContain("原需求重新分析已接通");
    /*
     * 新段正常结束 ⇒ 过程框收起成一行。展开之后才能数到它带来的步骤 ——
     * 收起只是呈现，内容没有被删。
     */
    expect(cards[1].getAttribute("data-expanded")).toBe("false");
    await act(async () => cards[1].querySelector<HTMLButtonElement>(".task-progress-toggle")!.click());
    expect(cards[1].querySelectorAll(".activity-steps > li").length).toBeGreaterThan(0);
    expect(container.querySelectorAll(".task-progress-jump")).toHaveLength(1);
    expect(container.querySelectorAll(".activity-card.is-live").length).toBeLessThanOrEqual(1);
    expect(container.textContent).not.toContain("已接回原需求，重新分析");
    expect(runTurnMock.mock.calls.at(-1)?.[0].prompt).toBe("按会员等级汇总订单并列出异常");
    expect(runTurnMock).toHaveBeenCalledTimes(stage === "message" ? 1 : 2);
    expect(statusMock).not.toHaveBeenCalled();
    expect(continueMock).not.toHaveBeenCalled();
  });

  it("natural-language continuation queries the original task, and ordinary chat does not erase its recovery context", async () => {
    const { container } = await reachPauseCard(pauseResult());
    vi.mocked(understandMessage).mockResolvedValueOnce({ kind: "conversation", message: "这是暂停状态。", provider: "model" });
    await type(container.querySelector(".composer textarea"), "这个状态是什么意思");
    await submit(container.querySelector(".composer textarea"));
    expect(findButton(container, "继续分析")).not.toBeNull();
    statusMock.mockResolvedValue({ turnId: "turn-pause-1", status: "running", revision: 5, lastSequence: 10, canContinue: false } as never);
    vi.mocked(understandMessage).mockResolvedValueOnce({ kind: "task_recovery", turnId: "turn-pause-1", provider: "model" });
    await type(container.querySelector(".composer textarea"), "继续");
    await submit(container.querySelector(".composer textarea"));
    await until(() => Boolean(findButton(container, "核对任务状态")), "原任务必须保留核对入口");
    expect(vi.mocked(understandMessage).mock.calls.at(-1)?.[0].recoveryTasks).toEqual(["turn-pause-1"]);
    expect(statusMock.mock.calls.at(-1)?.[0]).toBe("turn-pause-1");
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    expect(continueMock).not.toHaveBeenCalled();
    expect(textOf(container)).not.toContain("被新的任务取代");
    expect(textOf(container)).not.toContain("**请不要");
  });

  it("a different workbook task leaves the old card queryable using its original task identity", async () => {
    const { container } = await reachPauseCard(pauseResult());
    runTurnMock.mockResolvedValueOnce({ kind: "answer", message: "另一个问题的回答", provider: "model" } as WpsTurnResult);
    await type(container.querySelector(".composer textarea"), "查询另一项统计");
    await submit(container.querySelector(".composer textarea"));
    const query = findButton(container, "核对任务状态");
    expect(query, "新任务不能销毁旧任务查询入口").not.toBeNull();
    statusMock.mockResolvedValue({ turnId: "turn-pause-1", status: "running", revision: 5, lastSequence: 10, canContinue: false } as never);
    await act(async () => query!.click());
    expect(statusMock.mock.calls.at(-1)?.[0]).toBe("turn-pause-1");
    expect(runTurnMock).toHaveBeenCalledTimes(2);
    expect(continueMock).not.toHaveBeenCalled();
  });

  it("a model-invented recovery ID neither queries nor starts another plan", async () => {
    const { container } = await reachPauseCard(pauseResult());
    const queries = statusMock.mock.calls.length;
    vi.mocked(understandMessage).mockResolvedValueOnce({ kind: "task_recovery", turnId: "another-conversation", provider: "model" });
    await type(container.querySelector(".composer textarea"), "继续");
    await submit(container.querySelector(".composer textarea"));
    expect(statusMock).toHaveBeenCalledTimes(queries);
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    expect(continueMock).not.toHaveBeenCalled();
    /*
     * 有字时主按钮是「发送」（这是这条交互的硬规则：绝不吞掉用户正在写的话）。
     * 核对/续跑的入口仍在，但要求输入框为空 —— 清空之后它必须回来。
     */
    expect(findButton(container, "发送")).not.toBeNull();
    await act(async () => {
      const input = container.querySelector<HTMLTextAreaElement>(".composer textarea")!;
      const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")!.set!;
      setter.call(input, "");
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });
    expect(findButton(container, "继续分析")).not.toBeNull();
  });

  it("查询取回删除方案时不进入删表检查点写入通道", async () => {
    const { container } = await reachPauseCard(pauseResult());
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);
    statusMock.mockResolvedValue({
      turnId: "turn-pause-1", status: "completed", revision: 6, lastSequence: 11, canContinue: false,
      result: { kind: "plan", provider: "model", plan: {
        id: "delete-plan", title: "删除工作表", summary: "", assumptions: [], warnings: [],
        actions: [{ type: "deleteWorksheet", sheet: "Sheet1" }], acceptanceCriteria: []
      } }
    } as never);
    vi.mocked(understandMessage).mockResolvedValueOnce({
      kind: "task_recovery", turnId: "turn-pause-1", provider: "model"
    });
    await type(container.querySelector(".composer textarea"), "核对原任务");
    await submit(container.querySelector(".composer textarea"));
    expect(textOf(container)).toContain("本次仅查询状态，未执行删除");
    expect(fetchMock).not.toHaveBeenCalled();
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    expect(continueMock).not.toHaveBeenCalled();
  });

  it("可继续时画按钮，且状态与操作分属不同节点（变异点：状态与按钮混在一处）", async () => {
    const { container } = await reachPauseCard(pauseResult());
    const card = container.querySelector(".wps-planning-pause-card");
    expect(card).not.toBeNull();

    // 状态独占 role=status 节点。
    const status = card!.querySelector('[role="status"]');
    expect(status).not.toBeNull();
    expect(status!.textContent).toContain("本轮处理时间已达上限");

    /*
     * 操作**已搬到输入框右下角**（同一个动作不再有两个入口）：
     * 卡面动作区只剩一句指路，一颗按钮都没有；状态节点里也不出现动作名。
     * 原来这条断言验的是"状态与操作分属不同节点"，这一点没变 —— 只是操作节点
     * 现在在 `.composer` 里，且它是**唯一**的那颗。
     */
    const actions = card!.querySelector(".wps-planning-pause-actions");
    expect(actions).not.toBeNull();
    expect(actions!.querySelectorAll("button")).toHaveLength(0);
    expect(actions!.textContent).toContain("输入框右下角");
    const primary = container.querySelector<HTMLButtonElement>(".composer .composer-primary");
    expect(primary).not.toBeNull();
    expect(primary!.getAttribute("aria-label")).toBe("继续分析");
    expect(primary!.getAttribute("type")).toBe("button");
    expect(container.querySelectorAll('.composer .composer-primary')).toHaveLength(1);
    // 判据：状态节点里不含动作名，动作名只出现在那颗按钮上。
    expect(status!.textContent).not.toContain("继续分析");
  });

  it("暂停成因来自结构化 pauseType，不同成因文案不同（变异点：统一一句通用文案）", async () => {
    const { container: budget } = await reachPauseCard(pauseResult({ pauseType: "budget" }));
    expect(textOf(budget)).toContain("本轮处理时间已达上限");
    mounted.forEach((entry) => entry.unmount());
    mounted = [];
    localStorage.clear();

    const { container: modelError } = await reachPauseCard(
      pauseResult({ pauseType: "model_error", reason: "模型连接中断。" })
    );
    expect(textOf(modelError)).toContain("模型连接中断");
  });

  it("canContinue=false 时不画按钮，只留状态说明（变异点：无条件放行按钮）", async () => {
    const { container } = await reachPauseCard(pauseResult({ canContinue: false }));
    const card = container.querySelector(".wps-planning-pause-card");
    expect(card).not.toBeNull();
    // 关键：没有可点的"继续分析"。
    expect(findButton(container, "继续分析")).toBeNull();
    expect(card!.querySelector(".wps-planning-pause-actions")).toBeNull();
    // 但状态说明还在 —— 不能把卡片变成一片空白。
    expect(card!.querySelector('[role="status"]')).not.toBeNull();
    expect(card!.querySelector(".wps-planning-pause-note")).not.toBeNull();
  });

  it("暂停卡落盘时不带暂停载荷全文（模型连接元数据不进历史）", async () => {
    const pause = pauseResult({
      availableModelConnections: [{ catalogModelId: "m1", label: "备用连接", modelId: "gpt-x", supportsVision: false }],
      sourceFingerprint: "fingerprint-should-not-persist",
      currentModelId: "gpt-current"
    });
    // 先证明这些字段确实在**服务端响应**里 —— 否则下面的"历史里没有"毫无意义。
    expect(JSON.stringify(pause)).toContain("备用连接");
    expect(JSON.stringify(pause)).toContain("fingerprint-should-not-persist");

    await reachPauseCard(pause);
    const stored = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    const serialized = JSON.stringify(stored);
    // 反向变异点：只要把整个 `pause` 载荷塞进 ChatMessage，这条立即变红。
    expect(serialized).not.toContain("备用连接");
    expect(serialized).not.toContain("availableModelConnections");
    expect(serialized).not.toContain("fingerprint-should-not-persist");
    expect(serialized).not.toContain("sourceFingerprint");
    expect(serialized).not.toContain("gpt-current");
    // 留下来的只有定位与归因这五个字段。
    const message = stored.conversations[0].messages.find(
      (item: { planningPause?: unknown }) => item.planningPause
    );
    expect(Object.keys(message.planningPause).sort()).toEqual([
      "action",
      "actionLabel",
      "canContinue",
      "clientLastSequence",
      "detail",
      "pauseType",
      "reason",
      "turnId"
    ]);
  });
});

/**
 * CAP-02 第 5 条：**面板刷新后恢复可恢复的规划任务**。
 *
 * 重开这一件事本身**不是**任何结论的证据。它只能证明"我不记得了"，所以：
 *
 *  - 必须拿 `turnId` 去服务端重新查真实状态，不能拿落盘快照画按钮；
 *  - 查回来之前不许出现可点按钮（否则用户点到基于过期 revision 的动作）；
 *  - **绝不自动续跑** —— 那会在用户没点过任何东西时触发一次模型调用。
 */
describe("CAP-02 面板刷新后恢复规划暂停", () => {
  it("重开后输入继续仍核对原任务，重新显示可续按钮而不自动续跑", async () => {
    await seedHistoryThenReload();
    const { container } = await mount(panel());
    await until(() => Boolean(findButton(container, "继续分析")), "服务端核对后应有继续入口");
    vi.mocked(understandMessage).mockResolvedValueOnce({
      kind: "task_recovery", turnId: "turn-pause-1", provider: "model"
    });
    await type(container.querySelector(".composer textarea"), "继续");
    await submit(container.querySelector(".composer textarea"));
    await until(() => statusMock.mock.calls.length === 2, "自然语言应再次查询原任务");
    expect(statusMock.mock.calls[1][0]).toBe("turn-pause-1");
    expect(findButton(container, "继续分析")).not.toBeNull();
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(continueMock).not.toHaveBeenCalled();
    expect(textOf(container)).not.toContain("被新的任务取代");
  });

  it("重开后重新查询服务端真实状态（变异点：不核对，直接拿落盘快照）", async () => {
    await seedHistoryThenReload();
    /*
     * 让核对**悬在半空**：不 resolve、不 reject。
     *
     * 判别点不是"网上有没有发出请求"，而是"请求落地之前卡片敢不敢画按钮"。
     * 如果断言发生在状态已落定之后（`waiting_user + canContinue` 会画出按钮），
     * 那这条用例对"不核对"的变异体就失去了判别力 —— 所以必须把响应悬住，
     * 在"正在核对"这一瞬时态取卡片的证据。
     */
    let releaseStatus: ((status: unknown) => void) | null = null;
    statusMock.mockReturnValue(
      new Promise((resolve) => {
        releaseStatus = resolve as (status: unknown) => void;
      })
    );
    const { container } = await mount(panel());
    await until(() => statusMock.mock.calls.length > 0, "重开后应拿 turnId 去查真实状态");
    expect(statusMock.mock.calls[0][0]).toBe("turn-pause-1");
    // 关键：只查询，不继续、不重发需求。
    expect(runTurnMock).not.toHaveBeenCalled();
    // 核对中：状态文字照给（用户知道这条任务还在），但操作区不许出现。
    expect(container.querySelector(".wps-planning-pause-note")?.textContent ?? "").toContain(
      "正在核对该任务的真实进度"
    );
    expect(findButton(container, "继续分析"), "核对完成前不该有可点按钮").toBeNull();
    expect(container.querySelector(".wps-planning-pause-actions")).toBeNull();
    // 放行，避免悬挂的 promise 泄漏到后续用例。
    await act(async () => {
      releaseStatus?.({ turnId: "turn-pause-1", status: "waiting_user", revision: 12, lastSequence: 30, canContinue: true, pauseType: "budget", reason: null });
    });
  });

  it("核对通过后按钮恢复可用（用的是服务端 revision 而不是落盘的那份）", async () => {
    await seedHistoryThenReload();
    const { container } = await mount(panel());
    await until(() => Boolean(findButton(container, "继续分析")), "核对通过后应恢复继续按钮");
    expect(statusMock).toHaveBeenCalledTimes(1);
    // 按钮存在 = 卡片用的是"服务端说还能续"这一结论，而不是落盘快照（快照本身
    // 也能推出按钮，所以下面用 completed 那条做反向判别）。
    expect(container.querySelector(".wps-planning-pause-actions")).not.toBeNull();
  });

  it("服务端说已完成时重开后**不**画继续按钮（变异点：忽略终态，照复活按钮）", async () => {
    await seedHistoryThenReload();
    /*
     * 判别点：服务端把这条任务的终态改口成 `completed`。
     *
     * 落盘快照里 `canContinue: true`、`action: continue_checkpoint` 都还在
     * ——"不核对、直接拿快照重建"的实现会照着快照把按钮画出来。
     * 正确实现拿服务端状态重新推导：终态一律 `action: "none"`，
     * 于是只剩状态说明、没有操作区。
     *
     * 顺带钉住"不许拿旧结论当新结论"：面板唯一允许产按钮的地方是
     * `wpsPlanningResumeView`（由**服务端刚返回的**状态推导）。
     */
    statusMock.mockResolvedValue({
      turnId: "turn-pause-1",
      status: "completed",
      revision: 40,
      lastSequence: 90,
      canContinue: false,
      pauseType: null,
      reason: null,
      result: { kind: "answer", provider: "model", message: "已经完成了" }
    });
    const { container } = await mount(panel());
    await until(() => statusMock.mock.calls.length > 0, "应核对暂停");
    await until(
      () => container.querySelector(".wps-planning-pause-note") !== null,
      "应落到状态说明"
    );
    expect(findButton(container, "继续分析")).toBeNull();
    expect(container.querySelector(".wps-planning-pause-actions")).toBeNull();
  });

  /**
   * 终态词表由共享层 `planningResumeDecision` 定义，面板不重复维护一份。
   * 这条用例在**面板契约层**钉住那个共享结论确实被用上了：服务端即便
   * 语无伦次地同时给出 `completed` 和 `canContinue: true`，终态也必须是终态 ——
   * 不能靠"服务端说能继续"把一条已结束的任务复活（那会重复规划）。
   */
  it("终态优先于放行位：completed + canContinue 也不复活按钮", async () => {
    await seedHistoryThenReload();
    statusMock.mockResolvedValue({
      turnId: "turn-pause-1",
      status: "completed",
      revision: 40,
      lastSequence: 90,
      canContinue: true,
      pauseType: "budget",
      reason: null,
      result: { kind: "answer", provider: "model", message: "已经完成了" }
    });
    const { container } = await mount(panel());
    await until(() => statusMock.mock.calls.length > 0, "应核对暂停");
    await until(
      () => container.querySelector(".wps-planning-pause-note") !== null,
      "应落到状态说明"
    );
    expect(findButton(container, "继续分析")).toBeNull();
    expect(container.querySelector(".wps-planning-pause-actions")).toBeNull();
  });

  it("服务端不放行时重开后只剩状态说明（变异点：核对结论不生效，照画按钮）", async () => {
    await seedHistoryThenReload();
    /*
     * 反向判别：这次让服务端**明确说不能续**。
     *
     * 与上一条互为锚点 —— 上一条证明"放行 → 有按钮"，这一条证明"不放行 → 没按钮"。
     * 两条读同一个 `canContinue`，所以任何"忽略服务端结论"的实现必然至少
     * 在其中一条上变红；只留一条的话，"永远画按钮"和"永远不画按钮"
     * 都能骗过其中一半。
     */
    statusMock.mockResolvedValue({
      turnId: "turn-pause-1",
      status: "waiting_user",
      revision: 41,
      lastSequence: 91,
      canContinue: false,
      pauseType: null,
      reason: "缺少可验证的检查点，只能重新描述需求。"
    });
    const { container } = await mount(panel());
    await until(() => statusMock.mock.calls.length > 0, "应核对暂停");
    await until(
      () => container.querySelector(".wps-planning-pause-note") !== null,
      "应落到状态说明"
    );
    expect(findButton(container, "继续分析")).toBeNull();
    expect(container.querySelector(".wps-planning-pause-actions")).toBeNull();
  });

  it("核对读不到进度时保守处理：不画按钮，也不谎称任务已消失", async () => {
    await seedHistoryThenReload();
    statusMock.mockRejectedValue(new Error("local service down"));
    const { container } = await mount(panel());
    await until(() => statusMock.mock.calls.length > 0, "应尝试核对");
    await until(
      () => (container.textContent ?? "").includes("暂时读不到这条任务的真实进度"),
      "应给出保守说明"
    );
    expect(findButton(container, "继续分析")).toBeNull();
  });

  /**
   * 「能不能继续」只认服务端给的 `continuation`，不认 `pauseType != null`。
   *
   * 这是一条**判别力**用例：服务端同时给了 `canContinue: true`（许可）和
   * `continuation.available: false`（事实不允许，因为这个任务已经登记过写入）。
   * 旧判据（暂停过 = 有现场）会照画一个「继续分析」按钮，点下去就是拿重新分析
   * 绕开执行核对；新判据必须不画按钮，并把服务端的原话摆出来。
   */
  it("服务端说这个任务已经写过原表：不画按钮，原因照抄（变异点：拿 pauseType 当有现场）", async () => {
    await seedHistoryThenReload();
    statusMock.mockResolvedValue({
      turnId: "turn-pause-1",
      status: "waiting_user",
      revision: 41,
      lastSequence: 91,
      canContinue: true,
      pauseType: "budget",
      reason: null,
      continuation: {
        available: false,
        code: "PLANNING_CONTEXT_HAS_EXECUTION",
        message: "这个任务已经登记过写入。先核对原表与执行状态，不能用重新分析绕过。",
        stage: "after_tool"
      }
    });
    const { container } = await mount(panel());
    await until(() => statusMock.mock.calls.length > 0, "应核对暂停");
    await until(
      () => (container.textContent ?? "").includes("先核对原表与执行状态"),
      "应把服务端的原因原样摆出来"
    );
    expect(findButton(container, "继续分析")).toBeNull();
    expect(container.querySelector(".wps-planning-pause-actions")).toBeNull();
    // 负向：不许把"已经写过原表"降级成"没有可验证的现场"。
    expect(container.textContent ?? "").not.toContain("没有可验证的完整规划现场");
  });
});

/**
 * 「继续分析」的**读表桥**（2026-10-20）。
 *
 * 缺口很具体：续跑路径只把进度交给文案，没有把 `bridgeTool` 指针交给本机桥；
 * `resumeWpsPlanning` 里那层进度重组还把这个字段整个丢掉了。两处任何一处成立，
 * 表现都是"点继续之后一直转圈" —— 模型在等服务端/本机读表，而本机在等模型出结果。
 *
 * 这里断言的是一条**真实链路**：面板点继续 → 服务端发指针 → 面板立桥 →
 * 宿主适配器按当前授权真读真算 → 回执发出去。只桩掉网络与"上一轮怎么暂停的"，
 * 桥状态机、授权复核、读取与计算都是生产实现。
 */
describe("继续分析的读表桥", () => {
  const BRIDGE_OP = "op_" + "7".repeat(32);
  const BRIDGE_FP = "c".repeat(64);

  it.each(["button", "text"] as const)("通过 %s 继续后，服务端的读表指针真的被本机读掉并回执", async (via) => {
    await seedHistoryThenReload();
    const bridgeCalls: Array<{ url: string; body: Record<string, any> }> = [];
    let bridgeExecutions = 0;
    let markReceipted: (() => void) | null = null;
    const receipted = new Promise<void>((resolve) => { markReceipted = resolve; });
    vi.stubGlobal(
      "fetch",
      vi.fn(async (url: string, init?: RequestInit) => {
        const target = String(url);
        const body = JSON.parse(String(init?.body ?? "{}"));
        bridgeCalls.push({ url: target, body });
        if (target.includes("/poll")) {
          return new Response(
            JSON.stringify({
              calls: [
                {
                  operationId: BRIDGE_OP,
                  tool: "query_table",
                  arguments: { mode: "values", scope: "selected" },
                  requestFingerprint: BRIDGE_FP,
                  expiresAt: 0
                }
              ],
              generation: 1
            }),
            { status: 200, headers: { "Content-Type": "application/json" } }
          );
        }
        if (target.includes("/receipt")) {
          bridgeExecutions += 1;
          markReceipted?.();
          return new Response(JSON.stringify({ status: "accepted", code: null }), {
            status: 200,
            headers: { "Content-Type": "application/json" }
          });
        }
        throw new Error(`未预期的请求：${target}`);
      })
    );

    statusMock.mockResolvedValue({
      turnId: "turn-pause-1",
      status: "waiting_user",
      revision: 12,
      lastSequence: 30,
      canContinue: true,
      pauseType: "budget",
      reason: null,
      // 这份现场没读过来源 ⇒ 继续前不需要重读数据（纯结构场景）。
      continuation: {
        available: true,
        code: null,
        message: "",
        stage: "after_tool",
        sourceVerificationRequired: false
      }
    });
    continueMock.mockImplementation(async (_body, hooks) => {
      // 服务端在续跑流里发出的**读表指针**。必须是完整契约事件（tsc 会卡）。
      hooks?.onProgress?.({
        schemaVersion: 1,
        operationId: "turn-pause-1",
        sequence: 31,
        phase: "planning",
        stepId: "bridge.query_table",
        status: "progress",
        title: "正在读取授权表",
        timestamp: "2026-10-20T10:00:00.000Z",
        bridgeTool: { operationId: BRIDGE_OP, tool: "query_table" }
      });
      /*
       * 与生产一致的时序：**流在收到回执之后才结束**。
       *
       * 服务端发出指针后是在等本机回执的，这段等待期里流一直开着。如果这里
       * 立刻返回，面板就会走到收尾（`sdkBridge.end()`），在途的读表被正确判成
       * "这一轮已经结束" —— 那是桥的生命周期语义，不是本用例要证的东西。
       */
      await receipted;
      return { kind: "answer", message: "已按当前授权重新分析" } as never;
    });

    const { container } = await mount(panel());
    await until(() => Boolean(findButton(container, "继续分析")), "应出现继续按钮");
    if (via === "button") {
      await act(async () => container.querySelector<HTMLButtonElement>('button[aria-label="继续任务"], [aria-label="继续分析"]')!.click());
    } else {
      vi.mocked(understandMessage).mockResolvedValueOnce({ kind: "task_recovery", turnId: "turn-pause-1", action: "continue", provider: "model" });
      await type(container.querySelector(".composer textarea"), "继续");
      await submit(container.querySelector(".composer textarea"));
    }

    await until(
      () => bridgeCalls.some((call) => call.url.includes("/receipt")),
      "续跑期间的读表请求必须被本机执行并回执，而不是只剩一句进度文案"
    );

    // 归属身份：用的是**原任务**的 turnId，不是新造的任务。
    const poll = bridgeCalls.find((call) => call.url.includes("/poll"));
    expect(poll?.body.taskId).toBe("turn-pause-1");
    const receipt = bridgeCalls.find((call) => call.url.includes("/receipt"));
    expect(receipt?.body).toMatchObject({
      taskId: "turn-pause-1",
      operationId: BRIDGE_OP
    });
    // 同一个 operationId 只执行一次。
    expect(bridgeExecutions).toBe(1);
    expect(continueMock).toHaveBeenCalledTimes(1);
  });

  /**
   * 服务端说这份现场**依赖来源**时，继续请求里必须带着本轮**重读**到的指纹。
   *
   * 这是 A0 补充修复里最容易被"改回省事写法"的一处：把
   * `sourceRequired: status.continuation?.sourceVerificationRequired !== false`
   * 写成 `sourceRequired: false`，共享层就不会去读工作簿，交上去的是
   * `not_collected`，服务端只能拒绝（SOURCE_UNVERIFIED）。界面上看不出区别 ——
   * 用户只看到"点继续没反应"。所以判据必须落在**交上去的请求体**上。
   */
  it("来源变化拒绝后撤下旧任务继续入口，重开也不复活", async () => {
    await seedHistoryThenReload();
    statusMock.mockResolvedValue({ turnId: "turn-pause-1", status: "waiting_user", revision: 12,
      lastSequence: 30, canContinue: true, pauseType: "budget", reason: null,
      sourceFingerprint: "old-source-fingerprint", sourceFingerprintSheets: ["Sheet1"],
      continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: true } });
    const { container } = await mount(panel());
    await until(() => Boolean(findButton(container, "继续分析")), "应出现继续入口");
    await act(async () => container.querySelector<HTMLButtonElement>('button[aria-label="继续任务"], [aria-label="继续分析"]')!.click());
    await until(() => (container.textContent ?? "").includes("原任务使用的数据已经变化"), "应拒绝旧现场");
    expect(continueMock).not.toHaveBeenCalled();
    expect(container.querySelector('button[aria-label="继续任务"], [aria-label="继续分析"]')).toBeNull();
    expect(findButton(container, "继续分析")).toBeNull();
    mounted.forEach((entry) => entry.unmount()); mounted = [];
    const reopened = await mount(panel());
    expect(reopened.container.querySelector('button[aria-label="继续任务"], [aria-label="继续分析"]')).toBeNull();
    expect(continueMock).not.toHaveBeenCalled();
  });

  /**
   * 作废一条任务**只作废那一条**。
   *
   * 判据不是"界面上还剩几个按钮"，而是**哪一条**：暂停状态、落盘历史、输入框 ▶
   * 全都按 `turnId` 认人。把 `current?.turnId === turnId ? null : current` 写成
   * `setPlanningPause(null)`、或把落盘更新写成"所有暂停消息一起 resolved"，这条就红。
   */
  it("来源失效只作废该任务：另一条独立暂停任务仍保有继续入口", async () => {
    const stored = await seedHistoryThenReload();
    // 历史里再放一条**更早**的独立暂停任务（另一个 turnId），它必须活下来。
    const history = JSON.parse(stored);
    const messages = history.conversations[0].messages as Array<{ id: string; planningPause?: { turnId: string } }>;
    const latest = messages.find((item) => item.planningPause)!;
    const independent = JSON.parse(JSON.stringify(latest)) as { id: string; planningPause: { turnId: string } };
    independent.id = "independent-pause-message";
    independent.planningPause.turnId = "turn-pause-2";
    messages.splice(messages.indexOf(latest), 0, independent);
    localStorage.setItem("excel-bro.wps.chat.v1", JSON.stringify(history));

    statusMock.mockImplementation(async (turnId: string) => (turnId === "turn-pause-1"
      ? { turnId, status: "waiting_user", revision: 12, lastSequence: 30, canContinue: true,
          pauseType: "budget", reason: null, sourceFingerprint: "old-source-fingerprint",
          sourceFingerprintSheets: ["Sheet1"],
          continuation: { available: true, code: null, message: "", stage: "after_tool", sourceVerificationRequired: true } }
      : { turnId, status: "waiting_user", revision: 5, lastSequence: 8, canContinue: true,
          pauseType: "budget", reason: null,
          continuation: { available: true, code: null, message: "", stage: "after_tool" } }));

    const { container } = await mount(panel());
    await until(() => Boolean(container.querySelector('button[aria-label="继续任务"], [aria-label="继续分析"]')), "最新暂停任务应有继续入口");
    await act(async () => container.querySelector<HTMLButtonElement>('button[aria-label="继续任务"], [aria-label="继续分析"]')!.click());
    await until(() => textOf(container).includes("原任务使用的数据已经变化"), "应拒绝旧现场");

    const persisted = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    const after = persisted.conversations[0].messages as Array<{
      planningPause?: { turnId: string; resolved?: boolean; canContinue?: boolean; terminalReason?: string } }>;
    const invalidated = after.find((item) => item.planningPause?.turnId === "turn-pause-1")!.planningPause!;
    const survivor = after.find((item) => item.planningPause?.turnId === "turn-pause-2")!.planningPause!;
    expect(invalidated.resolved).toBe(true);
    expect(invalidated.terminalReason).toContain("来源数据已变化");
    // 独立任务没有被一起作废：它仍是"未解决、可继续"，入口也仍然指向它。
    expect(survivor.resolved).toBeUndefined();
    expect(survivor.canContinue).toBe(true);
    /*
     * 它的入口此刻是「核对任务状态」而不是「继续分析」：另一条刚被拒，这条还没核对过，
     * 主按钮按**未核对**处理（核对完才会变成继续）。入口仍然唯一、仍然指向它 ——
     * 这条断言要证的是"另一条的失败没有把它一起作废"，不是按钮叫什么。
     */
    expect(container.querySelector('button[aria-label="核对任务状态"]')).not.toBeNull();

    const before = statusMock.mock.calls.length;
    await act(async () => container.querySelector<HTMLButtonElement>('button[aria-label="核对任务状态"]')!.click());
    await until(() => statusMock.mock.calls.slice(before).some(([turnId]) => turnId === "turn-pause-2"),
      "▶ 必须指向那条独立任务");
    expect(continueMock.mock.calls.some(([request]) => (request as { turnId?: string }).turnId === "turn-pause-1")).toBe(false);
  });

  it("服务端要求核实来源时，继续请求必须带上此刻重读的指纹（变异点：不采集必要来源指纹）", async () => {
    await seedHistoryThenReload();
    const live = createWpsBridgeLiveContext(
      pauseHost() as never,
      [1]
    );
    statusMock.mockResolvedValue({
      turnId: "turn-pause-1",
      status: "waiting_user",
      revision: 12,
      lastSequence: 30,
      canContinue: true,
      pauseType: "budget",
      reason: null,
      // 原任务读过来源：记录里既有指纹也有取指纹的表。
      sourceFingerprint: live.sourceFingerprint,
      sourceFingerprintSheets: ["Sheet1"],
      continuation: {
        available: true,
        code: null,
        message: "",
        stage: "after_tool",
        sourceVerificationRequired: true
      }
    });
    continueMock.mockResolvedValue({ kind: "answer", message: "已按当前数据重新分析" } as never);
    vi.stubGlobal(
      "fetch",
      vi.fn(async (url: string) => {
        const target = String(url);
        if (target.includes("/poll")) {
          return new Response(JSON.stringify({ calls: [], generation: 1 }), {
            status: 200,
            headers: { "Content-Type": "application/json" }
          });
        }
        return new Response(JSON.stringify({ status: "accepted", code: null }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      })
    );

    const { container } = await mount(panel());
    await until(() => Boolean(findButton(container, "继续分析")), "应出现继续按钮");
    await act(async () => {
      findButton(container, "继续分析")!.dispatchEvent(
        new MouseEvent("click", { bubbles: true })
      );
    });
    await until(() => continueMock.mock.calls.length > 0, "应真的发起继续请求");

    const body = continueMock.mock.calls[0][0] as unknown as Record<string, unknown>;
    // 三态必须是"已核实"，而不是"尚未采集"。
    expect(body.sourceVerification).toBe("verified");
    // 指纹必须与此刻宿主读到的那一份**完全相同**（不是旧快照、不是文件名）。
    expect(body.sourceFingerprint).toBe(live.sourceFingerprint);
    expect(body.workbookIdentity).toBe(live.workbookIdentity);
    expect(body.authorizedSheetNames).toEqual(["Sheet1"]);
  });
});

/**
 * 读表桥没立起来时，**续跑请求不许发出去**（2026-10-20）。
 *
 * 与「继续分析的读表桥」那一组相反：那组证明的是"桥活着时指针真的被读掉"，
 * 这组证明的是"桥没立起来时请求根本不该发"。缺口很具体 —— 面板 `await` 了
 * `ensureLive`，却把返回的 `bridgeReady` 只用在"进度事件要不要交给桥"上，
 * `resumeWpsPlanning` 照样调用。桥初始化被取消/被取代时返回 false，请求照样
 * 发出去，服务端一发出 `bridgeTool` 指针就没人接，模型只能等超时；而界面上
 * 续跑看起来"成功"了，因为 `resumeWpsPlanning` 真的返回了。
 *
 * 这里桩的是 `useSdkToolBridge` 的返回值（可控地让 `ensureLive` 返回 false），
 * 桥自己的行为由 `useSdkToolBridge.test.tsx` 的 23 条用例 + 变异证明。断言的
 * 是调用层：`continueWpsPlanning` 真的没被调用，且卡片如实说明停在哪一步。
 */
describe("读表桥没立起来时不发续跑请求", () => {
  it.each([false, true])("首次规划由真实面板在请求之前准备桥（不可用=%s）", async (unavailable) => {
    const actual = await vi.importActual<typeof import("./wpsReadonlyAgent")>("./wpsReadonlyAgent");
    const planning = vi.fn(async () => ({
      kind: "answer" as const, provider: "model" as const, message: "首次规划已返回"
    }));
    listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
    getModelSettingsMock.mockResolvedValue({} as never);
    getWpsApplicationMock.mockReturnValue(pauseHost());
    bridgeStub.ensureLive = unavailable;
    runTurnMock.mockImplementation((input) => actual.runWpsReadonlyTurn(input, {
      checkIntent: async () => ({
        kind: "proceed", provider: "model", turnId: "turn-first-preflight",
        summary: "统计", confirmedPrompt: "统计授权表"
      }),
      streamPlanning: planning,
      readSchemas: readAuthorizedWpsSheetSchemas,
      readSheets: readAuthorizedWpsSheets,
      executeQuery: executeQueryTableData,
      createAssistantResponse: vi.fn(),
      getPlanningStatus: statusMock,
      continuePlanning: continueMock
    }));
    const { container } = await mount(panel());
    await until(
      () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false,
      "应加载模型目录"
    );
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计授权表");
    await submit(textarea);
    await until(() => textOf(container).includes(unavailable ? "未发送规划请求" : "首次规划已返回"),
      `应根据桥准备结果停止或规划：${textOf(container)}`);
    expect(planning).toHaveBeenCalledTimes(unavailable ? 0 : 1);
    expect(textOf(container)).not.toContain("服务端会按超时收尾");
  });

  it("ensureLive 返回 false ⇒ 不发继续请求，卡片留在可重试状态", async () => {
    await seedHistoryThenReload();
    // 让桥"建不起来"。
    bridgeStub.ensureLive = true;
    statusMock.mockResolvedValue({
      turnId: "turn-pause-1",
      status: "waiting_user",
      revision: 12,
      lastSequence: 30,
      canContinue: true,
      pauseType: "budget",
      reason: null,
      continuation: {
        available: true,
        code: null,
        message: "",
        stage: "after_tool",
        sourceVerificationRequired: false
      }
    });

    const { container } = await mount(panel());
    await until(() => Boolean(findButton(container, "继续分析")), "应出现继续按钮");
    await act(async () => {
      findButton(container, "继续分析")!.dispatchEvent(
        new MouseEvent("click", { bubbles: true })
      );
    });

    await until(
      () => textOf(container).includes("没有发送继续请求"),
      "应如实说明没有发送继续请求"
    );
    // 关键断言：续跑请求一次都没有发出去。
    expect(continueMock).not.toHaveBeenCalled();
    // 卡片仍在：用户可以再点一次，而不是被吞掉一次尝试。
    expect(findButton(container, "继续分析")).not.toBeNull();
  });
});
