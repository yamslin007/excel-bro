// @vitest-environment jsdom
/**
 * CAP-02 第 1 条：**续跑结果闭环**的面板契约测试。
 *
 * 这个文件要钉住的事只有一件：点"继续"之后，服务端返回的 `plan` /
 * `clarification` / `local_plan_rule` 必须**接回与一次性轮次相同的处理链** ——
 * 计划进预览、澄清进可填写卡、本地规则按当前授权读取计算。
 *
 * 此前这三支一律落进"请重新发送需求"，等于让用户白做一遍；更严重的是，
 * 一旦有人把"重发"实现成自动重跑，就会变成一次**未确认的重新规划**。
 *
 * 因此这里的断言分两类：
 *
 *  - **正向锚点**：继续后确实进入了对应界面（预览 / 澄清卡）且宿主写入为 0；
 *  - **反向变异**：拆掉"接回处理链"这一环时必须变红。
 *
 * 具体变异体在每条 `it` 的名称里标注。
 */
import { act, type ComponentProps, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ApiRequestError, listModels, getModelSettings, understandMessage, checkIntent, pausePlanningTask } from "../../excel-addin/src/api";
import { getWpsApplication } from "./wpsHost";
import type { WpsTurnResult } from "./wpsReadonlyAgent";

vi.mock("./wpsReadonlyAgent", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsReadonlyAgent")>();
  return {
    ...actual,
    runWpsReadonlyTurn: vi.fn(),
    // 续跑通道单独桩：本文件只关心"继续回来之后怎么处理"，不关心网络细节。
    resumeWpsPlanning: vi.fn()
  };
});

vi.mock("./wpsPlanningStream", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsPlanningStream")>();
  return {
    ...actual,
    getWpsPlanningTaskStatus: vi.fn(),
    // 「原任务仍在运行」时接回进度的唯一入口（只读补发）。
    getWpsPlanningTaskEvents: vi.fn()
  };
});

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  const discovery = await import("./wpsPendingDiscovery.testHost");
  /*
   * 发现入口的替身见上。几条"读表回执 / 暂停 / 校验"用例故意要跑**真实实现**（它们自己
   * `vi.stubGlobal("fetch", …)` 驱动真实字节），所以这里不能把它们换成空的 `vi.fn()`。
   * 用 `guardNetworkSeam`：本用例装了 fetch 替身才真的调用，否则按"连不上"失败 ——
   * 之前那种"看用户那台 8765 有没有在跑"的偶然前提就此消失。
   */
  const gate = await import("@excel-bro/core/testNetworkGate");
  return { ...actual, ...discovery.pendingDiscoveryApiStubs(),
    pollSdkBridgeTool: vi.fn(gate.guardNetworkSeam(actual.pollSdkBridgeTool)),
    pausePlanningTask: vi.fn(gate.guardNetworkSeam(actual.pausePlanningTask)),
    validatePlan: vi.fn(gate.guardNetworkSeam(actual.validatePlan)),
    checkIntent: vi.fn(), understandMessage: vi.fn(async () => ({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" })), listModels: vi.fn(), getModelSettings: vi.fn() };
});

vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

/**
 * 读表桥的探针（2026-10-23）。
 *
 * 只做两件事：把 `ensureLive` 收到的 `onProgress` 记下来（用它注入"暂停之后才
 * 到达的迟到进度"），以及数 `end()` 被调了几次（暂停后桥必须收尾）。
 * 其余一律转发真实实现 —— 桥自己的行为由 `hooks/useSdkToolBridge.test.tsx` 证。
 */
const bridgeProbe = vi.hoisted(() => ({
  progress: null as null | ((progress: { scannedRows: number; totalRows: number; sheet: string }) => void),
  ensureCalls: 0,
  endCalls: 0
}));

vi.mock("../../excel-addin/src/hooks/useSdkToolBridge", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/hooks/useSdkToolBridge")>();
  const React = await import("react");
  type Controller = ReturnType<typeof actual.useSdkToolBridge>;
  type EnsureOptions = Parameters<Controller["ensureLive"]>[0];
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
              const current = latest.current as unknown as Record<string, unknown>;
              if (key === "ensureLive") {
                return async (options: EnsureOptions) => {
                  bridgeProbe.ensureCalls += 1;
                  bridgeProbe.progress = (options?.onProgress as typeof bridgeProbe.progress) ?? null;
                  return latest.current.ensureLive(options);
                };
              }
              if (key === "end") {
                return () => {
                  bridgeProbe.endCalls += 1;
                  latest.current.end();
                };
              }
              const value = current[key];
              return typeof value === "function" ? value.bind(latest.current) : value;
            }
          }),
        []
      );
    }
  };
});

const { default: WpsAgentPanel } = await import("./WpsAgentPanel");
const { runWpsReadonlyTurn, resumeWpsPlanning } = await import("./wpsReadonlyAgent");
const { getWpsPlanningTaskEvents, getWpsPlanningTaskStatus } = await import("./wpsPlanningStream");

(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const listModelsMock = vi.mocked(listModels);
const getModelSettingsMock = vi.mocked(getModelSettings);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);
const resumeMock = vi.mocked(resumeWpsPlanning);
const statusMock = vi.mocked(getWpsPlanningTaskStatus);
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

/**
 * 等**真实的墙钟时间**，而不是"settle 若干次"。
 *
 * 只有一条路径需要它：「原任务仍在运行」的补发轮询按 `RUNNING_TAIL_INTERVAL_MS`
 * （1200ms）退避，`settle()` 的 0ms 循环永远跑不到下一轮 —— 用假定时器又会把
 * 组件里所有真实计时口径一起改掉。所以这里如实等时间，只是给足上限。
 */
async function untilWithin(check: () => boolean, message: string, timeoutMs = 8000): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    if (check()) return;
    if (Date.now() > deadline) throw new Error(message);
    await act(async () => {
      await new Promise<void>((resolve) => setTimeout(resolve, 25));
    });
  }
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

/**
 * 「继续」入口：规划暂停卡上叫「继续分析」，普通暂停走输入框那颗叫「继续任务」。
 * 这些用例验的是续跑生命周期，不是按钮叫什么，所以两种都接受；入口仍然只有一个。
 */
function findResume(container: HTMLElement): HTMLButtonElement | null {
  return findButton(container, "继续分析") ?? findButton(container, "继续任务");
}

/**
 * 回复段快照：每段过程卡的步骤行文本。
 *
 * 「继续之后旧段不被改写」只能靠逐字比较证明 —— 只数卡片数量，区分不了
 * "历史段被冻结"和"历史段被就地改写后又复制了一份"。
 */
function segmentSnapshots(container: HTMLElement) {
  return [...container.querySelectorAll<HTMLElement>(".task-progress-timeline")].map(
    (card) => ({
      status: card.getAttribute("data-status"),
      steps: [...card.querySelectorAll<HTMLElement>(".activity-steps > li")].map(
        (row) => row.textContent ?? ""
      )
    })
  );
}

/** 当前活动段数量：同一任务同一时刻最多一个。 */
function liveSegmentCount(container: HTMLElement): number {
  return container.querySelectorAll(".activity-card.is-live").length;
}

/**
 * 最后一张过程卡的标题行文本（含"· N 步 · 1:23"）。
 *
 * "耗时冻结"只能靠它证明：卡还在活动时，标题里的时间每秒重绘；
 * 暂停之后必须逐字不变。
 */
function lastSegmentHeader(container: HTMLElement): string {
  const headers = container.querySelectorAll<HTMLElement>(".task-progress-heading strong");
  return headers[headers.length - 1]?.textContent ?? "";
}

/** 最后一段过程卡元素。 */
function lastSegment(container: HTMLElement): HTMLElement | null {
  const cards = container.querySelectorAll<HTMLElement>(".task-progress-timeline");
  return cards[cards.length - 1] ?? null;
}

/**
 * 记录宿主写入次数的探针。
 *
 * 用 `Range.Value2` / `Range.Formula` 的 **setter** 计数：写入单元格必然
 * 经过其中之一。读取（getter）不算 —— 续跑要读源值，那不算"写"。
 *
 * 形状刻意对齐 `wpsNativeComparison.test.ts` 的可 staging 宿主：
 * `prepareWpsPlan` 需要 `Formula`（生成公式快照）与 `Interior`（着色标记），
 * 缺了会在预览阶段抛错，于是"计划分支"根本没被走到。
 */
function writeCountingHost() {
  const writes: string[] = [];
  const cell = { value: [["类别", "归属"], ["甲", "华东"], ["乙", "华南"]] as unknown, formula: [["类别", "归属"], ["甲", "华东"], ["乙", "华南"]] as unknown };
  let pendingValue: unknown = cell.value;
  const makeRange = (rows: number, columns: number) => {
    const range = { Rows: { Count: rows }, Columns: { Count: columns }, Row: 1, Column: 1, MergeCells: false } as Record<string, unknown>;
    Object.defineProperties(range, {
      Value2: {
        get: () => pendingValue,
        set: (value: unknown) => { writes.push(JSON.stringify(value)); pendingValue = value; }
      },
      Formula: {
        get: () => (rows === 3 ? cell.formula : [["类别", "归属"]]),
        set: (value: unknown) => { writes.push(JSON.stringify(value)); }
      },
      Interior: { value: { Color: undefined } },
      Item: { value: () => range }
    });
    return range;
  };
  const usedRange = makeRange(3, 2);
  const headerRange = makeRange(1, 2);
  const sheet = {
    Name: "Sheet1",
    Index: 1,
    CodeName: "SheetCode1",
    ProtectContents: false,
    UsedRange: usedRange,
    Range: () => headerRange
  };
  const storage = new Map<string, unknown>();
  const host = {
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
  return { host, writes, storage };
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

/**
 * 走到"暂停卡出现"这一步。
 *
 * 全程走真实交互（真实 `runTurn` 分支、真实暂停卡渲染）—— 只有网络缝
 * （`runWpsReadonlyTurn` / `resumeWpsPlanning` / `getWpsPlanningTaskStatus`）
 * 被桩掉，这样"继续之后怎么处理结果"才是**真的**在跑。
 */
async function reachPauseCard(pause: ReturnType<typeof pauseResult>, numberedSteps = false) {
  const { host, writes } = writeCountingHost();
  listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
  getModelSettingsMock.mockResolvedValue({} as never);
  getWpsApplicationMock.mockReturnValue(host);
  runTurnMock.mockImplementation(async ({ onProgress }) => {
    if (numberedSteps) for (let step = 2; step <= 7; step++) {
      onProgress?.({ stage: "answering", turnId: pause.turnId, sequence: step,
        stepId: `step-${step}`, title: `步骤${step}`, message: `步骤${step}` });
    }
    return { kind: "planning_pause", pause } as WpsTurnResult;
  });
  statusMock.mockResolvedValue({
    turnId: pause.turnId,
    status: "waiting_user",
    revision: 12,
    lastSequence: 30,
    canContinue: true,
    pauseType: "budget",
    reason: null
  });
  const { container } = await mount(panel());
  await until(
    () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false,
    "应加载模型目录"
  );
  const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
  await type(textarea, "生成统计预览");
  await submit(textarea);
  await until(() => Boolean(container.querySelector(".wps-planning-pause-card")), "应出现规划暂停卡");
  return { container, writes };
}

async function clickContinue(container: HTMLElement): Promise<void> {
  const button = findResume(container);
  if (!button) throw new Error("没有继续按钮");
  await act(async () => {
    button.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
}

beforeEach(() => {
  vi.clearAllMocks();
  mounted = [];
  localStorage.clear();
  // 默认"读不到补发事件"：绝大多数用例走到"仍在运行"就该停下，不偷偷发真请求。
  vi.mocked(getWpsPlanningTaskEvents).mockRejectedValue(new TypeError("Failed to fetch"));
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("CAP-02 续跑结果闭环：三分支接回正常处理链", () => {
  it.each(["composer", "card"])("real field intent pause via %s settles before the response, including after continue", async (via) => {
    const actual = await vi.importActual<typeof import("./wpsReadonlyAgent")>("./wpsReadonlyAgent");
    runTurnMock.mockImplementation((input) => actual.runWpsReadonlyTurn(input));
    const { host, writes } = writeCountingHost();
    getWpsApplicationMock.mockReturnValue(host);
    listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
    getModelSettingsMock.mockResolvedValue({} as never);
    vi.mocked(understandMessage).mockResolvedValue({ kind: "needs_workbook_context", summary: "分析", provider: "model" });
    const releases: Array<(value: Awaited<ReturnType<typeof checkIntent>>) => void> = [];
    vi.mocked(checkIntent).mockImplementation(() => new Promise((resolve) => releases.push(resolve)));
    const { container } = await mount(panel());
    await until(() => Boolean(container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型")), "模型就绪");
    const input = container.querySelector<HTMLTextAreaElement>("textarea");
    await type(input, "分析当前表"); await submit(input);
    for (let attempt = 1; attempt <= 2; attempt++) {
      await until(() => releases.length === attempt, "真实流程进入字段理解");
      const pause = via === "composer" ? container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]') : findButton(container, "暂停");
      expect(pause).toBeTruthy();
      await act(async () => pause!.click());
      await until(() => !container.querySelector('[aria-label="暂停任务"]'), "请求未返回也应停止等待");
      expect(container.querySelector('[aria-label="继续任务"], [aria-label="继续分析"]')).toBeTruthy();
      const frozen = segmentSnapshots(container);
      await act(async () => releases[attempt - 1]({ kind: "proceed", provider: "model", turnId: `late-${attempt}`,
        summary: "迟到结果", confirmedPrompt: "分析当前表" }));
      expect(segmentSnapshots(container)).toEqual(frozen);
      expect(statusMock).not.toHaveBeenCalled();
      expect(resumeMock).not.toHaveBeenCalled();
      expect(writes).toEqual([]);
      if (attempt === 1) await act(async () => container.querySelector<HTMLButtonElement>('[aria-label="继续任务"], [aria-label="继续分析"]')!.click());
    }
  });

  it("ignores already consumed query events in a new segment and closes the attachment receipt", async () => {
    const { container, writes } = await reachPauseCard(pauseResult(), true);
    statusMock.mockResolvedValueOnce({ turnId: "turn-pause-1", status: "running", revision: 12,
      lastSequence: 30, canContinue: false } as never)
      .mockResolvedValue({ turnId: "turn-pause-1", status: "completed", revision: 13,
        lastSequence: 31, canContinue: false,
        result: { kind: "answer", provider: "model", message: "仅取回已有结果" } } as never);
    vi.mocked(getWpsPlanningTaskEvents).mockResolvedValue({ turnId: "turn-pause-1", active: false,
      status: "completed", lastSequence: 31, events: [
        { schemaVersion: 1, operationId: "turn-pause-1", sequence: 7, phase: "planning",
          stepId: "step-7", status: "completed", title: "旧查询不应重复出现" },
        { schemaVersion: 1, operationId: "turn-pause-1", sequence: 31, phase: "planning",
          stepId: "submit", status: "completed", title: "已有结果已提交" }
      ] } as never);
    await clickContinue(container);
    await until(() => Boolean(container.textContent?.includes("仅取回已有结果")), "取回结果");
    expect(vi.mocked(getWpsPlanningTaskEvents).mock.calls[0][1]).toBe(7);
    expect(lastSegment(container)?.textContent).not.toContain("旧查询不应重复出现");
    /*
     * 这一段正常结束 ⇒ 过程框收起成一行，步骤行不在 DOM 里。
     * 要考"每一步的状态"，先展开；收起是呈现，不是内容消失。
     */
    await act(async () => lastSegment(container)!.querySelector<HTMLButtonElement>(".task-progress-toggle")!.click());
    expect(lastSegment(container)?.querySelector('.activity-steps > li')?.getAttribute('data-status')).toBe('completed');
    expect(lastSegment(container)?.textContent).not.toContain("此步骤未收到结束确认");
    expect(lastSegment(container)?.getAttribute('data-status')).toBe('completed');
    expect(resumeMock).not.toHaveBeenCalled();
    expect(runTurnMock).toHaveBeenCalledOnce();
    expect(writes).toEqual([]);
  });

  it("composer pause stops a stalled recovery poll and ignores its late page", async () => {
    const { container, writes } = await reachPauseCard(pauseResult(), true);
    const requests: string[] = [];
    vi.stubGlobal('fetch', vi.fn(async (url: string) => {
      requests.push(String(url)); return new Response('{}', {status: 200});
    }));
    statusMock.mockResolvedValue({ turnId: "turn-pause-1", status: "running", revision: 12,
      lastSequence: 30, canContinue: false } as never);
    let finish!: (value: any) => void;
    vi.mocked(getWpsPlanningTaskEvents).mockImplementation(() => new Promise((resolve) => { finish = resolve; }));
    await clickContinue(container);
    await until(() => Boolean(finish), '补发请求开始');
    await act(async () => container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')!.click());
    await until(() => !container.querySelector('[aria-label="暂停任务"]'), '暂停不能等无响应的轮询结束');
    expect(requests.some((url) => url.endsWith('/api/turn/turn-pause-1/pause'))).toBe(true);
    const before = segmentSnapshots(container);
    await act(async () => finish({ turnId: 'turn-pause-1', active: false, status: 'completed', lastSequence: 31,
      events: [{ operationId: 'turn-pause-1', sequence: 31, stepId: 'late', title: '迟到页面不应出现', status: 'completed', phase: 'planning' }] }));
    expect(segmentSnapshots(container)).toEqual(before);
    expect(container.textContent).not.toContain('迟到页面不应出现');
    expect(resumeMock).not.toHaveBeenCalled();
    expect(writes).toEqual([]);
  });

  it.each([
    ["MODEL_QUOTA_EXHAUSTED", "模型 API 余额不足或调用额度已用尽", "模型设置"],
    ["MODEL_CONTEXT_LIMIT", "上下文长度上限", "拆分任务"],
    ["MODEL_OUTPUT_LIMIT", "未获得完整结果", "输出上限"]
  ])("shows actionable model limit %s without writing", async (code, title, action) => {
    const { host, writes } = writeCountingHost();
    getWpsApplicationMock.mockReturnValue(host);
    listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
    getModelSettingsMock.mockResolvedValue({} as never);
    vi.mocked(understandMessage).mockRejectedValueOnce(new ApiRequestError(502, code, "PRIVATE_PROVIDER_BODY", true));
    const { container } = await mount(panel());
    await until(() => Boolean(container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型")), "模型目录应就绪");
    await type(container.querySelector("textarea"), "检查当前表");
    await submit(container.querySelector("textarea"));
    await until(() => Boolean(container.textContent?.includes(title)), "额度提示必须对用户可见");
    expect(container.textContent).toContain(action);
    expect(container.textContent).toContain("请以执行结果卡为准");
    expect(container.textContent).not.toContain("PRIVATE_PROVIDER_BODY");
    expect(container.querySelector("textarea")?.value).toBe("检查当前表");
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(writes).toEqual([]);
  });
  it("点击继续时原任务已完成，结果落在新的回复段并结束那一段", async () => {
    const { container, writes } = await reachPauseCard(pauseResult(), true);
    const before = segmentSnapshots(container);
    statusMock.mockResolvedValue({ turnId: "turn-pause-1", status: "completed", revision: 13,
      lastSequence: 32, canContinue: false,
      result: { kind: "answer", provider: "model", message: "后台已完成的结果已接回" } } as never);
    await clickContinue(container);
    await until(() => (container.textContent ?? "").includes("后台已完成的结果已接回"), "应取回已完成结果");
    expect(resumeMock).not.toHaveBeenCalled();
    expect(runTurnMock).toHaveBeenCalledOnce();
    const after = segmentSnapshots(container);
    // 同任务、多段回复：暂停段原样保留，结果落在新段的收尾状态上。
    expect(after).toHaveLength(before.length + 1);
    expect(after.slice(0, before.length).map((segment) => segment.steps))
      .toEqual(before.map((segment) => segment.steps));
    expect(lastSegment(container)?.getAttribute("data-status")).toBe("completed");
    expect(liveSegmentCount(container)).toBeLessThanOrEqual(1);
    expect(writes).toEqual([]);
  });

  it.each(["confirmed", "unconfirmed", "changed-path", "changed-sheet-id"])(
    "checks original analysis consent when host wrappers change: %s", async (scenario) => {
      const first = writeCountingHost();
      getWpsApplicationMock.mockReturnValue(first.host);
      listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
      getModelSettingsMock.mockResolvedValue({} as never);
      runTurnMock.mockResolvedValue({ kind: "planning_pause", pause: pauseResult() } as WpsTurnResult);
      statusMock.mockResolvedValue({ turnId: "turn-pause-1", status: "waiting_user", revision: 12,
        lastSequence: 30, canContinue: true, pauseType: "budget", reason: null });
      resumeMock.mockResolvedValue({ status: {} as never,
        outcome: { kind: "answer", provider: "model", message: "已接续原分析" } });
      let confirmScope: Parameters<NonNullable<ComponentProps<typeof WpsAgentPanel>["onScopeConfirmationReady"]>>[0];
      const { container } = await mount(<WpsAgentPanel selectedSheetIndexes={[1]} scopeDock={null}
        onScopeConfirmationReady={(handler) => { confirmScope = handler; }} />);
      await until(() => Boolean(confirmScope) && Boolean(container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型")), "初始化完成");
      if (scenario !== "unconfirmed") {
        const { inspectWpsAuthorizedTarget } = await import("./wpsAuthorizedTarget");
        await act(async () => { expect(await confirmScope!(await inspectWpsAuthorizedTarget([1]))).toBe(true); });
      }
      await type(container.querySelector("textarea"), "生成统计预览");
      await submit(container.querySelector("textarea"));
      await until(() => Boolean(container.querySelector(".wps-planning-pause-card")), "应出现原任务暂停卡");
      const fresh = writeCountingHost();
      const book = (fresh.host as import("./wpsHost").WpsApplication).ActiveWorkbook!;
      if (scenario === "changed-path") Object.assign(book, { FullName: "D:\\Other\\测试.xlsx", Path: "D:\\Other" });
      if (scenario === "changed-sheet-id") Object.assign(book.Worksheets.Item(1), { CodeName: "ReplacementSheet" });
      // Native host reads may return new wrappers on every access, not just once.
      Object.defineProperty(fresh.host, "ActiveWorkbook", { get: () => ({ ...book,
        Worksheets: { ...book.Worksheets, Item: (index: number) => ({ ...book.Worksheets.Item(index) }) }
      }) });
      getWpsApplicationMock.mockReturnValue(fresh.host);
      await clickContinue(container);
      if (scenario === "confirmed") {
        await until(() => (container.textContent ?? "").includes("已接续原分析"), "同一范围已确认，不应被包装对象变化阻断");
        expect(resumeMock).toHaveBeenCalledOnce();
      } else {
        expect(resumeMock).not.toHaveBeenCalled();
        if (scenario === "unconfirmed") {
          expect(container.textContent).not.toContain("原任务的工作簿或授权范围已变化");
          expect(container.textContent).toContain("暂时无法核实原任务");
          const { inspectWpsAuthorizedTarget } = await import("./wpsAuthorizedTarget");
          await act(async () => { expect(await confirmScope!(await inspectWpsAuthorizedTarget([1]))).toBe(true); });
          await clickContinue(container);
          await until(() => (container.textContent ?? "").includes("已接续原分析"), "补充确认后应接回原任务");
          expect(resumeMock).toHaveBeenCalledOnce();
        }
      }
      expect(runTurnMock).toHaveBeenCalledOnce();
      expect([...first.writes, ...fresh.writes]).toEqual([]);
    });

  it.each(["button", "继续", "接着刚才暂停的任务往下做"])("resumes the seventh row via %s before appending step eight", async (via) => {
    const { container, writes } = await reachPauseCard(pauseResult(), true);
    let finish!: () => void;
    resumeMock.mockImplementation(({ onProgress }) => new Promise((resolve) => {
      onProgress?.({ stage: "answering", turnId: "turn-pause-1", sequence: 31,
        stepId: "step-7", title: "步骤7继续", message: "接续未完成步骤" });
      finish = () => {
        onProgress?.({ stage: "answering", turnId: "turn-pause-1", sequence: 32,
          stepId: "step-8", title: "步骤8", message: "下一步" });
        resolve({ status: {} as never, outcome: { kind: "answer", provider: "model", message: "接续完成" } });
      };
    }));
    const paused = segmentSnapshots(container);
    if (via === "button") await clickContinue(container);
    else {
      vi.mocked(understandMessage).mockResolvedValueOnce({ kind: "task_recovery", turnId: "turn-pause-1", action: "continue", provider: "model" });
      await type(container.querySelector("textarea"), via);
      await submit(container.querySelector("textarea"));
    }
    await until(() => Boolean(finish), "继续开始");
    // 续跑段自成一档：它不重排编号，所以既不显示"步骤 7"，也不是第 7 行。
    // 旧段保留暂停时的样子，只多一个"查看后续进展"的定位入口。
    const before = segmentSnapshots(container);
    expect(before).toHaveLength(paused.length + 1);
    expect(before.slice(0, paused.length).map((segment) => segment.steps))
      .toEqual(paused.map((segment) => segment.steps));
    const resumed = before[before.length - 1];
    expect(resumed.steps.length).toBe(2);
    expect(resumed.steps[0]).toContain("接回原任务进度");
    expect(resumed.steps[1]).toContain("步骤7继续");
    expect(lastSegment(container)?.querySelector(".activity-step-index")).toBeNull();
    expect(container.textContent).toContain("步骤7继续");
    expect(liveSegmentCount(container)).toBeLessThanOrEqual(1);
    expect(container.querySelector('[aria-label="任务执行过程"]')).toBeNull();
    await act(async () => finish());
    await until(() => (container.textContent ?? "").includes("接续完成"), "完成");
    /*
     * 正常结束 ⇒ 过程框收起成一行；要看"第八步真的追加了"，得先展开。
     * 收起是呈现，步骤没有被删。
     */
    expect(lastSegment(container)?.getAttribute("data-expanded")).toBe("false");
    await act(async () => lastSegment(container)!.querySelector<HTMLButtonElement>(".task-progress-toggle")!.click());
    const done = segmentSnapshots(container);
    expect(done).toHaveLength(paused.length + 1);
    expect(done[0].steps).toEqual(paused[0].steps);
    expect(done[done.length - 1].steps).toHaveLength(resumed.steps.length + 1);
    expect(done[done.length - 1].steps.at(-1)).toContain("步骤8");
    expect(runTurnMock).toHaveBeenCalledOnce();
    expect(writes).toEqual([]);
  });
  /**
   * 「原任务仍在服务端运行」这条恢复路径。
   *
   * 服务端在跑时「继续」会被 409 挡住（再发一次就是重复规划），但"什么都不发生"
   * 同样是错的：任务确实在跑，进度也确实写进了账本。正确做法是**只读补发**已落盘
   * 的事件，把它们落到**新的回复段**（同任务、多段回复），等它结束后再走同一条
   * 结果处理链。
   */
  it.each([false, true])("running 接回处理读表回执，包括暂停前已收到的指针 %s", async (alreadyReceived) => {
    const OP = "op_" + "b".repeat(32);
    let published = alreadyReceived;
    let receipted = false;
    const receipts: any[] = [];
    vi.stubGlobal("fetch", vi.fn(async (url: string, init?: RequestInit) => {
      const body = JSON.parse(String(init?.body ?? "{}"));
      if (String(url).includes("/poll")) return new Response(JSON.stringify({ generation: 1,
        calls: published && !receipted ? [{ operationId: OP, tool: "query_table",
          arguments: { mode: "aggregate", scope: "selected", metrics: [{ operation: "count", outputName: "rows" }] }, requestFingerprint: "d".repeat(64), expiresAt: 0 }] : [] }), { status: 200 });
      if (String(url).includes("/receipt")) {
        receipts.push(body); receipted = true;
        return new Response(JSON.stringify({ status: "accepted", code: null }), { status: 200 });
      }
      return new Response("{}", { status: 200 });
    }));
    const { container, writes } = await reachPauseCard(pauseResult(), true);
    statusMock.mockImplementation(async () => ({ turnId: "turn-pause-1", status: receipted ? "completed" : "running",
      canContinue: false, revision: 12, lastSequence: 31,
      ...(receipted ? { result: { kind: "answer", provider: "model", message: "收到真实读表回执后完成" } } : {}) } as never));
    vi.mocked(getWpsPlanningTaskEvents).mockImplementation(async () => {
      const first = !published; published = true;
      return { turnId: "turn-pause-1", active: !receipted, status: receipted ? "completed" : "running", lastSequence: 31,
        events: first ? [{ schemaVersion: 1, operationId: "turn-pause-1", sequence: 31, phase: "planning",
          stepId: "bridge.query_table:bbbb", status: "started", title: "读取授权表", timestamp: "2026-10-22T00:00:00Z",
          bridgeTool: { operationId: OP, tool: "query_table" } }] : [] } as never;
    });
    const before = segmentSnapshots(container);
    await clickContinue(container);
    await untilWithin(() => (container.textContent ?? "").includes("收到真实读表回执后完成"), "必须等到读表回执后才能取回结果");
    // 继续开新段；暂停段保留原样，同一时刻只有一个活动段。
    const after = segmentSnapshots(container);
    expect(after).toHaveLength(before.length + 1);
    expect(after.slice(0, before.length).map((segment) => segment.steps))
      .toEqual(before.map((segment) => segment.steps));
    expect(liveSegmentCount(container)).toBeLessThanOrEqual(1);
    expect(resumeMock).not.toHaveBeenCalled();
    expect(runTurnMock).toHaveBeenCalledOnce();
    expect(writes).toEqual([]);
    expect(receipts).toHaveLength(1);
    expect(receipts[0]).toMatchObject({ taskId: "turn-pause-1", operationId: OP });
    expect(receipts[0].receipt.errorCode).toBeUndefined();
    expect(receipts[0].receipt.result).toBeDefined();
  });

  it("理解消息中暂停后继续开新的回复段，旧段停在暂停那一刻", async () => {
    const { host, writes } = writeCountingHost();
    getWpsApplicationMock.mockReturnValue(host);
    listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
    getModelSettingsMock.mockResolvedValue({} as never);
    vi.mocked(understandMessage).mockImplementationOnce((_request, signal) => new Promise((_resolve, reject) => {
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    let finish: (() => void) | undefined;
    runTurnMock.mockImplementation(() => new Promise((resolve) => {
      finish = () => resolve({ kind: "answer", provider: "model", message: "完成" } as WpsTurnResult);
    }));
    const { container } = await mount(panel());
    await until(() => Boolean(container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型")), "模型就绪");
    const input = container.querySelector<HTMLTextAreaElement>("textarea");
    await type(input, "生成统计预览"); await submit(input);
    await until(() => Boolean(container.querySelector('[aria-label="暂停任务"]')), "理解中可以暂停");
    await act(async () => container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')!.click());
    await until(() => Boolean(findResume(container)), "暂停后可以继续");
    const paused = segmentSnapshots(container);
    expect(paused).toHaveLength(1);
    // 暂停那一刻：未收到结束确认，展示为中断而非业务警告。
    expect(paused[0].steps[0]).toContain("理解消息");
    expect(container.querySelector('.activity-steps > li')?.getAttribute("data-status")).toBe("interrupted");
    expect(container.querySelector('.activity-steps > li .activity-step-check')?.getAttribute("aria-label")).toBe("本段已中断");
    await clickContinue(container);
    await until(() => runTurnMock.mock.calls.length === 1, "接续到下一阶段");
    const after = segmentSnapshots(container);
    // 判别点 1（缺陷 1）：继续开**新的回复段**，旧段逐字停在暂停那一刻。
    expect(after).toHaveLength(paused.length + 1);
    expect(after.slice(0, paused.length)).toEqual(paused);
    // 旧段的结论仍是"等待你处理"，没有被这一轮改写成"已完成"。
    expect(after[0].status).toBe("awaiting_input");
    // 判别点 2：旧段给出「查看后续进展」，且只有当前段是活动段。
    expect(container.querySelectorAll(".task-progress-jump")).toHaveLength(after.length - 1);
    expect(liveSegmentCount(container)).toBeLessThanOrEqual(1);
    // 判别点 3：新段的第一行已经跑完（不是把旧段的 warning 搬过来）。
    const resumedStep = lastSegment(container)?.querySelector(".activity-steps > li");
    expect(resumedStep?.getAttribute("data-status")).toBe("completed");
    expect(resumedStep?.textContent).toContain("理解消息");
    // 计时只在新段上继续走；旧段的耗时不参与。
    const elapsed = lastSegment(container)?.querySelector(".activity-step-time")?.textContent;
    await settle();
    expect(lastSegment(container)?.querySelector(".activity-step-time")?.textContent).toBe(elapsed);
    await act(async () => finish?.());
    expect(writes).toEqual([]);
  });

  it("原任务仍在运行时把后续进度接到新的回复段，不重发需求也不换执行身份", async () => {
    const { container, writes } = await reachPauseCard(pauseResult(), true);
    statusMock
      .mockResolvedValueOnce({ turnId: "turn-pause-1", status: "running", revision: 12,
        lastSequence: 30, canContinue: false } as never)
      .mockResolvedValueOnce({ turnId: "turn-pause-1", status: "completed", revision: 13,
        lastSequence: 32, canContinue: false,
        result: { kind: "answer", provider: "model", message: "接续完成" } } as never);
    vi.mocked(getWpsPlanningTaskEvents)
      // 第 1 页：任务确实还在跑，但它已经把第 7 步的进度写进账本了。
      .mockResolvedValueOnce({ turnId: "turn-pause-1", status: "running", active: true,
        lastSequence: 31, events: [{
          schemaVersion: 1, operationId: "turn-pause-1", sequence: 31, phase: "planning",
          stepId: "step-7", status: "progress", title: "步骤7继续",
          timestamp: "2026-10-21T00:00:00Z"
        }] } as never)
      // 第 2 页：任务已经结束，最后一次补发带出第 8 步。
      .mockResolvedValueOnce({ turnId: "turn-pause-1", status: "completed", active: false,
        lastSequence: 32, events: [{
          schemaVersion: 1, operationId: "turn-pause-1", sequence: 32, phase: "planning",
          stepId: "step-8", status: "completed", title: "步骤8",
          timestamp: "2026-10-21T00:00:00Z"
        }] } as never);

    const pausedSegments = segmentSnapshots(container);
    await clickContinue(container);

    await untilWithin(() => (container.textContent ?? "").includes("接续完成"), "任务结束后应取回结果");
    // 判别点 1：继续产生**新的回复段**；暂停段保留它停止时的步骤行。
    const after = segmentSnapshots(container);
    expect(after).toHaveLength(pausedSegments.length + 1);
    expect(after.slice(0, pausedSegments.length).map((segment) => segment.steps))
      .toEqual(pausedSegments.map((segment) => segment.steps));
    // 判别点 2：补发出来的两步落在**最新那一段**里（续跑段不重排编号）。
    const resumed = lastSegment(container)!;
    // 同上：这一段已经正常结束并收起，展开后才能数到补发出来的步骤。
    expect(resumed.getAttribute("data-expanded")).toBe("false");
    await act(async () => resumed.querySelector<HTMLButtonElement>(".task-progress-toggle")!.click());
    const rows = resumed.querySelectorAll(".activity-steps > li");
    expect(rows).toHaveLength(3);
    expect(rows[0].textContent).toContain("接回原任务进度");
    expect(rows[1].textContent).toContain("步骤7继续");
    expect(rows[2].textContent).toContain("步骤8");
    expect(resumed.querySelector(".activity-step-index")).toBeNull();
    // 判别点 3：只读补发，**没有**再发一次继续请求，也没有新的规划轮次。
    expect(resumeMock).not.toHaveBeenCalled();
    expect(runTurnMock).toHaveBeenCalledOnce();
    expect(writes).toEqual([]);
  });

  async function manuallyPaused(waitForInspection = true) {
    const { host, writes } = writeCountingHost();
    getWpsApplicationMock.mockReturnValue(host);
    listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
    getModelSettingsMock.mockResolvedValue({} as never);
    runTurnMock.mockImplementation(({ signal, onProgress }) => new Promise((_resolve, reject) => {
      onProgress?.({ stage: "answering", turnId: "turn-manual-pause", sequence: 4, message: "正在规划" });
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    const entry = await mount(panel());
    const { container } = entry;
    await until(() => Boolean(container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型")), "目录就绪");
    const input = container.querySelector<HTMLTextAreaElement>("textarea");
    await type(input, "继续这个统计任务"); await submit(input);
    await until(() => runTurnMock.mock.calls.length === 1 && Boolean(container.querySelector('[aria-label="暂停任务"]')), "应进入规划");
    await act(async () => { container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')!.click(); });
    await until(() => statusMock.mock.calls.length > 0 && (!waitForInspection || !container.querySelector('[aria-label="暂停任务"]')), "暂停后自动核对原任务");
    return { ...entry, writes };
  }

  it.each(["running", "unavailable", "no_checkpoint", "completed"])("manual pause status %s never triggers another planner or write", async (kind) => {
    if (kind === "unavailable") statusMock.mockRejectedValue(new TypeError("Failed to fetch"));
    else statusMock.mockResolvedValue({ turnId: "turn-manual-pause", status: kind === "no_checkpoint" ? "interrupted" : kind,
      revision: 6, lastSequence: 8, canContinue: false, pauseType: "interrupted", reason: "没有可续跑检查点",
      ...(kind === "completed" ? { result: { kind: "answer", provider: "model", message: "已取回原任务结果" } } : {}) } as never);
    const { container, writes } = await manuallyPaused();
    expect(statusMock).toHaveBeenCalledOnce();
    expect(container.querySelector<HTMLTextAreaElement>("textarea")?.value).toBe("");
    expect(resumeMock).not.toHaveBeenCalled();
    expect(runTurnMock).toHaveBeenCalledOnce();
    expect(findResume(container)).toBeNull();
    if (kind === "running" || kind === "unavailable") expect(findButton(container, "核对任务状态")).not.toBeNull();
    if (kind === "completed") {
      expect(container.textContent).toContain("结果已保留");
      expect(container.textContent).not.toContain("已取回原任务结果");
    }
    expect(writes).toEqual([]);
  });

  it("WPS pause timeout is unconfirmed and ignores a late acknowledgement", async () => {
    const { container } = await reachPauseCard(pauseResult());
    let resolvePause!: (value: { turnId: string; pauseRequested: boolean }) => void;
    vi.mocked(pausePlanningTask).mockImplementationOnce(() => new Promise(resolve => { resolvePause = resolve; }));
    resumeMock.mockImplementation(({ signal, onProgress }) => new Promise((_resolve, reject) => {
      onProgress?.({ stage: "answering", turnId: "turn-pause-1", sequence: 14, message: "继续中" });
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    await clickContinue(container);
    await until(() => resumeMock.mock.calls.length === 1 && Boolean(container.querySelector('[aria-label="暂停任务"]')), "可暂停");
    vi.useFakeTimers();
    try {
      await act(async () => { container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')!.click(); });
      await act(async () => { await vi.advanceTimersByTimeAsync(7999); });
      expect(container.querySelector(".composer-pause-note")?.getAttribute("data-phase")).toBe("requesting");
      await act(async () => { await vi.advanceTimersByTimeAsync(1); });
      const note = container.querySelector(".composer-pause-note");
      expect(note?.getAttribute("data-phase")).toBe("unconfirmed");
      expect(note?.textContent).toContain("8 秒内没有确认");
      expect(liveSegmentCount(container)).toBe(0);
      await act(async () => { resolvePause({ turnId: "turn-pause-1", pauseRequested: true }); });
      expect(note?.getAttribute("data-phase")).toBe("unconfirmed");
    } finally { vi.useRealTimers(); }
  });

  it.each(["composer", "card"])("a resumed stream via %s can be paused again and keeps its latest received cursor", async (entry) => {
    const { container, writes } = await reachPauseCard(pauseResult());
    let resolvePause!: (value: { turnId: string; pauseRequested: boolean }) => void;
    vi.mocked(pausePlanningTask).mockImplementationOnce(() => new Promise((resolve) => { resolvePause = resolve; }));
    const pauseButton = () => entry === "composer"
      ? container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')
      : findButton(container, "暂停");
    let frozenSteps: string[][] = [];
    const pauseRound = async (round: number) => {
      await until(() => Boolean(pauseButton()) && liveSegmentCount(container) === 1, "本轮可暂停");
      const rows = [...container.querySelectorAll(".message-row")];
      const currentRow = lastSegment(container)?.closest(".message-row");
      const userRow = container.querySelectorAll(".message-row.user");
      expect(rows.indexOf(currentRow!)).toBeGreaterThan(rows.indexOf(userRow[userRow.length - 1]));
      expect(segmentSnapshots(container).slice(0, frozenSteps.length).map(segment => segment.steps)).toEqual(frozenSteps);
      await act(async () => { pauseButton()!.click(); });
      await until(() => liveSegmentCount(container) === 0, "本轮已收尾");
      /*
       * 暂停请求在飞时，右下角只能是**禁用的「正在暂停」**：这时候还能点"继续"，
       * 等于在暂停还没落地时又把它拉起来。以前那颗按钮长在卡面上、不受输入框状态
       * 约束，所以能点；现在它和主按钮是同一颗，就必须守这条。
       */
      const inFlight = container.querySelector<HTMLButtonElement>(".composer .composer-primary");
      expect(inFlight).not.toBeNull();
      // 暂停回执在飞 ⇒ 只能是禁用的「正在暂停」；回执已到 ⇒ 才是可点的「继续」。
      // 两种都只有**一颗**，不存在"卡面一颗、输入框一颗"。
      if (inFlight!.getAttribute("aria-label") === "正在暂停") {
        expect(inFlight!.disabled).toBe(true);
      } else {
        expect(inFlight!.getAttribute("aria-label")).toBe("继续分析");
        expect(inFlight!.disabled).toBe(false);
      }
      expect(container.querySelectorAll(".composer .composer-primary")).toHaveLength(1);
      expect(resumeMock.mock.calls[round - 1][0].signal?.aborted).toBe(true);
      frozenSteps = segmentSnapshots(container).map(segment => segment.steps);
      const resumes = resumeMock.mock.calls.length;
      const sends = runTurnMock.mock.calls.length;
      // 已暂停（或暂停请求在飞）时根本没有「发送」这颗按钮：空表单提交既不会发出去，
      // 也不会开新一轮；主按钮永远是 type=button，表单 submit 不能顺带触发它。
      expect(container.querySelector<HTMLButtonElement>('button[aria-label="发送"]')).toBeNull();
      expect(container.querySelector<HTMLButtonElement>(".composer-primary")?.getAttribute("type")).toBe("button");
      await act(async () => { container.querySelector("form.composer-box")!.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true })); });
      expect(resumeMock).toHaveBeenCalledTimes(resumes);
      expect(runTurnMock).toHaveBeenCalledTimes(sends);
      // 输入区上方只剩一句状态说明，没有第二颗按钮（动作只在右下角）。
      expect(container.querySelector('.task-execution-notice button')).toBeNull();
    };
    const pausedSegments = segmentSnapshots(container);
    resumeMock.mockImplementation(({ signal, onProgress }) => new Promise((_resolve, reject) => {
      onProgress?.({ stage: "answering", turnId: "turn-pause-1", sequence: 14, message: "继续中" });
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    await clickContinue(container);
    await until(() => resumeMock.mock.calls.length === 1, "继续开始");
    await pauseRound(1);
    /*
     * 暂停回执迟到：在飞期间不许继续（上面已断言），回执回来之后才出现继续入口。
     * 这条迟到回执也不能污染下一轮 —— 后面的轮次继续验这件事。
     */
    await act(async () => { resolvePause({ turnId: "turn-pause-1", pauseRequested: true }); });
    await until(() => Boolean(findResume(container)), "再次核对检查点");
    /*
     * 再走两轮"继续 → 暂停"，凑成**三停三续**。
     *
     * 两轮就够暴露链断裂，但三轮才能证明"每一段只指向自己的直接后继"不是
     * 巧合：A→B→C→D 里任何一次父段选错，最终都会少一个「查看后续进展」。
     */
    for (const round of [2, 3]) {
      await clickContinue(container);
      await until(() => resumeMock.mock.calls.length === round, `第 ${round} 次继续开始`);
      if (round === 2) {
        // 迟到的那份暂停回执只落在第 1 轮，不该在输入区留下任何提示。
        expect(container.querySelector(".composer-pause-note")).toBeNull();
      }
      await pauseRound(round);
      await until(() => Boolean(findResume(container)), `第 ${round} 次暂停后可再次继续`);
    }
    resumeMock.mockResolvedValue({ status: {} as never, outcome: { kind: "answer", provider: "model", message: "三次暂停后完成" } as never });
    await clickContinue(container);
    await until(() => (container.textContent ?? "").includes("三次暂停后完成"), "完成");
    expect(resumeMock.mock.calls[1][0]).toMatchObject({ turnId: "turn-pause-1", afterSequence: 14 });
    expect(resumeMock.mock.calls[2][0]).toMatchObject({ turnId: "turn-pause-1", afterSequence: 14 });
    expect(resumeMock.mock.calls[3][0]).toMatchObject({ turnId: "turn-pause-1", afterSequence: 14 });
    // 连续三次"暂停 → 继续"后：历史段各自保留，活动段仍只有一个。
    const after = segmentSnapshots(container);
    expect(after.length).toBeGreaterThanOrEqual(pausedSegments.length + 3);
    expect(after.slice(0, pausedSegments.length).map((segment) => segment.steps))
      .toEqual(pausedSegments.map((segment) => segment.steps));
    /*
     * 判别点（缺陷 2）：段链是 A→B→C，每一段的「查看后续进展」接的是**直接后继**。
     * 旧实现用 `messages.find(...)` 从前往后找父段 ⇒ 第二次继续仍然改写 A 的
     * `nextSegmentId`（A→C），B 没有后继 ⇒ 跳转入口只有 1 个而段有 3 段。
     */
    expect(container.querySelectorAll(".task-progress-jump")).toHaveLength(after.length - 1);
    expect(liveSegmentCount(container)).toBeLessThanOrEqual(1);
    expect(runTurnMock).toHaveBeenCalledOnce();
    expect(writes).toEqual([]);
  });

  /*
   * ===================================================================
   * 暂停后的状态一致性（2026-10-23）
   *
   * 用户实测：继续之后按暂停，请求确实 abort 了，但过程卡仍显示"进行中"、
   * 耗时继续增长；随后按钮还可能重新变回"可暂停"却点不动。
   *
   * 这三条用例把"请求取消 / 界面停止"两件事**分开**断言：abort 只能证明前者，
   * 界面必须自己收尾。
   * ===================================================================
   */
  it.each(["composer", "card"])(
    "pausing a resumed stream via %s leaves no live card and freezes the elapsed time",
    async (via) => {
      const { container } = await reachPauseCard(pauseResult());
      resumeMock.mockImplementation(({ signal, onProgress }) => new Promise((_resolve, reject) => {
        onProgress?.({ stage: "answering", turnId: "turn-pause-1", sequence: 14, message: "继续中" });
        signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
      }));
      await clickContinue(container);
      await until(() => resumeMock.mock.calls.length === 1, "继续开始");
      await until(() => liveSegmentCount(container) === 1, "继续之后应有一段活动卡");
      await until(() => Boolean(container.querySelector('[aria-label="暂停任务"]')), "应显示暂停按钮");

      const pause = via === "composer"
        ? container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')
        : findButton(container, "暂停");
      expect(pause, `${via} 入口应有暂停按钮`).toBeTruthy();
      await act(async () => { pause!.click(); });

      // A 请求取消：只证明"连接断了"，不证明界面停止。
      await until(() => Boolean(resumeMock.mock.calls[0][0].signal?.aborted), "继续请求应被 abort");

      // C 界面停止：必须为 0，不是"≤1"。
      await until(() => liveSegmentCount(container) === 0, "暂停后不得再有活动卡");

      // 耗时冻结：卡头的时间文本在 1.1s 真实时间后必须逐字不变。
      const frozenHeader = lastSegmentHeader(container);
      await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, 1100)); });
      expect(lastSegmentHeader(container)).toBe(frozenHeader);

      // 不再提供任何"可暂停"入口（含卡片上的）。
      expect(container.querySelector('[aria-label="暂停任务"]')).toBeNull();
      expect(findButton(container, "暂停")).toBeNull();
    }
  );

  it("pausing a resumed stream also ends the read bridge it opened", async () => {
    /*
     * 桥的归属与 controller 不同：按暂停时自动触发的那次状态核对会把
     * `controllerRef` 换成自己的 controller。旧实现用同一个 `ownsTurn` 布尔管三样
     * 资源，于是"controller 被接走 ⇒ 桥永远不收"，迟到的读表进度又能把按钮点亮。
     */
    const { container } = await reachPauseCard(pauseResult());
    const before = bridgeProbe.endCalls;
    resumeMock.mockImplementation(({ signal, onProgress }) => new Promise((_resolve, reject) => {
      onProgress?.({ stage: "answering", turnId: "turn-pause-1", sequence: 14, message: "继续中" });
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    await clickContinue(container);
    await until(() => resumeMock.mock.calls.length === 1, "继续开始");
    await until(() => bridgeProbe.progress !== null, "读表桥应登记过进度回调");
    await act(async () => {
      container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')!.click();
    });
    await until(() => liveSegmentCount(container) === 0, "暂停后不得再有活动卡");
    await until(() => bridgeProbe.endCalls > before, "暂停后必须收掉本次继续立起的读表桥");
  });

  it("pausing stops the local side even while the background status check never returns", async () => {
    /*
     * 判定点：**本端是否已经停下，不取决于核对回不回来。**
     *
     * 核对是一次真实的状态读取，它慢、它卡住，只说明"服务端状态未确认"，与
     * "我已经不接收这条流的进度了"是两件事。旧实现把收尾全挂在核对之后的
     * `finally` 里：核对不回来，活动卡就永远在转、耗时永远在涨、桥也一直活着。
     */
    const { container } = await reachPauseCard(pauseResult());
    const bridgeEndsBefore = bridgeProbe.endCalls;
    resumeMock.mockImplementation(({ signal, onProgress }) => new Promise((_resolve, reject) => {
      onProgress?.({ stage: "answering", turnId: "turn-pause-1", sequence: 14, message: "继续中" });
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    await clickContinue(container);
    await until(() => resumeMock.mock.calls.length === 1, "继续开始");
    await until(() => liveSegmentCount(container) === 1, "继续之后应有一段活动卡");

    // 从这里起，所有状态查询都**永不返回**：模拟后台卡住/网络挂起。
    statusMock.mockImplementation(() => new Promise(() => {}) as never);
    await act(async () => {
      container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')!.click();
    });

    await until(() => liveSegmentCount(container) === 0, "核对还没回来，活动卡也必须已经停下");
    // 耗时冻结：核对挂着也不再涨（活动已经收尾，不是"看起来停了"）。
    const frozenHeader = lastSegmentHeader(container);
    await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, 1100)); });
    expect(lastSegmentHeader(container)).toBe(frozenHeader);
    // 桥也必须已经收：它属于这次继续，不属于那次核对。
    expect(bridgeProbe.endCalls).toBeGreaterThan(bridgeEndsBefore);
  });

  it("a late bridge progress after pause never re-opens the running state", async () => {
    const { container } = await reachPauseCard(pauseResult());
    resumeMock.mockImplementation(({ signal, onProgress }) => new Promise((_resolve, reject) => {
      onProgress?.({ stage: "answering", turnId: "turn-pause-1", sequence: 14, message: "继续中" });
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    await clickContinue(container);
    await until(() => resumeMock.mock.calls.length === 1, "继续开始");
    await until(() => bridgeProbe.progress !== null, "读表桥应登记过进度回调");

    await act(async () => {
      container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')!.click();
    });
    await until(() => liveSegmentCount(container) === 0, "暂停后不得再有活动卡");

    /*
     * 迟到事件：暂停之后桥才把一条读取进度送达。
     *
     * 它既不能让"进行中"回来，也不能把输入框重新变成 ■（那条老路正是
     * "按钮显示可暂停、点下去却没有可中止对象"的来源）。
     */
    await act(async () => {
      bridgeProbe.progress?.({ scannedRows: 10, totalRows: 90, sheet: "Sheet1" });
    });
    await settle();
    expect(liveSegmentCount(container)).toBe(0);
    expect(container.querySelector('[aria-label="暂停任务"]')).toBeNull();
    expect(container.textContent ?? "").not.toContain("正在读取");
  });

  it("repeating pause and re-entering the same task keeps one live segment at most", async () => {
    const { container } = await reachPauseCard(pauseResult());
    resumeMock.mockImplementation(({ signal, onProgress }) => new Promise((_resolve, reject) => {
      onProgress?.({ stage: "answering", turnId: "turn-pause-1", sequence: 14, message: "继续中" });
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    for (const round of [1, 2, 3]) {
      await clickContinue(container);
      await until(() => resumeMock.mock.calls.length === round, `第 ${round} 次继续开始`);
      await until(() => liveSegmentCount(container) === 1, `第 ${round} 次继续后应有活动卡`);
      await act(async () => {
        container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')!.click();
      });
      // 每次暂停后都必须彻底停下：0 而不是"最多 1"。
      await until(() => liveSegmentCount(container) === 0, `第 ${round} 次暂停后不得再有活动卡`);
    }
    // 重复暂停/继续不得制造第二个同时活动的段。
    expect(liveSegmentCount(container)).toBe(0);
    expect(container.querySelector('[aria-label="暂停任务"]')).toBeNull();
  });

  it("reload queries a manually paused task but never automatically resumes it", async () => {
    statusMock.mockResolvedValue({ turnId: "turn-manual-pause", status: "running", revision: 4, lastSequence: 4, canContinue: false });
    const entry = await manuallyPaused();
    entry.unmount();
    mounted = mounted.filter((item) => item.container !== entry.container);
    statusMock.mockResolvedValue({ turnId: "turn-manual-pause", status: "interrupted", revision: 5, lastSequence: 10,
      canContinue: true, pauseType: "interrupted", reason: "检查点可用" });
    const { container } = await mount(panel());
    await until(() => Boolean(findResume(container)), "重载后核对可继续现场");
    expect(resumeMock).not.toHaveBeenCalled();
    resumeMock.mockResolvedValue({ status: {} as never, outcome: { kind: "answer", provider: "model", message: "重载后原断点完成" } as never });
    await clickContinue(container);
    await until(() => (container.textContent ?? "").includes("重载后原断点完成"), "原任务完成");
    expect(resumeMock.mock.calls[0][0]).toMatchObject({ turnId: "turn-manual-pause", afterSequence: 4 });
    expect(runTurnMock).toHaveBeenCalledOnce();
  });
  it("real continue request keeps the original model and replays only after the received cursor", async () => {
    const { container, writes } = await reachPauseCard(pauseResult());
    const actual = await vi.importActual<typeof import("./wpsReadonlyAgent")>("./wpsReadonlyAgent");
    resumeMock.mockImplementation(actual.resumeWpsPlanning);
    const fetchMock = vi.fn(async (_url: unknown, init?: RequestInit) => {
      const body = JSON.parse(String(init?.body));
      if ("modelId" in body || "confirmModelSwitch" in body) return new Response("{}", { status: 422 });
      return new Response('event: result\ndata: {"kind":"answer","provider":"model","message":"真实续跑完成"}\n\n', {
        headers: { "Content-Type": "text/event-stream" }
      });
    });
    vi.stubGlobal("fetch", fetchMock);
    await clickContinue(container);
    await until(() => (container.textContent ?? "").includes("真实续跑完成"), "普通续跑不应被模型切换校验拒绝");
    const body = JSON.parse(String(fetchMock.mock.calls[0][1]?.body));
    expect(body.afterSequence).toBe(0);
    expect(body.turnId).toBe("turn-pause-1");
    expect(body.prompt).toBeUndefined();
    expect(writes).toEqual([]);
  });

  it("manual pause retains the task, checks status, then continues without resending or writing", async () => {
    vi.mocked(pausePlanningTask).mockResolvedValueOnce({ turnId: "turn-manual-pause", pauseRequested: true });
    const { host, writes } = writeCountingHost();
    getWpsApplicationMock.mockReturnValue(host);
    listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
    getModelSettingsMock.mockResolvedValue({} as never);
    runTurnMock.mockImplementation(({ signal, onProgress }) => new Promise((_resolve, reject) => {
      onProgress?.({ stage: "answering", turnId: "turn-manual-pause", sequence: 4, message: "正在规划" });
      signal?.addEventListener("abort", () => reject(new DOMException("paused", "AbortError")));
    }));
    statusMock.mockResolvedValue({ turnId: "turn-manual-pause", status: "interrupted", revision: 3,
      lastSequence: 9, canContinue: true, pauseType: "interrupted", reason: "已找到检查点" });
    resumeMock.mockResolvedValue({ status: {} as never,
      outcome: { kind: "answer", provider: "model", message: "从原断点完成" } as never });
    const { container } = await mount(panel());
    await until(() => Boolean(container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型")), "目录就绪");
    const input = container.querySelector<HTMLTextAreaElement>("textarea");
    await type(input, "继续这个统计任务"); await submit(input);
    await until(() => runTurnMock.mock.calls.length === 1 && Boolean(container.querySelector('[aria-label="暂停任务"]')), "应有暂停入口");
    await act(async () => { container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')!.click(); });
    await until(() => Boolean(findResume(container)), "暂停后自动核对检查点");
    expect(statusMock).toHaveBeenCalledOnce();
    expect(resumeMock).not.toHaveBeenCalled();
    expect(input?.value).toBe("");
    await submit(input);
    expect(resumeMock).not.toHaveBeenCalled();
    expect(runTurnMock).toHaveBeenCalledOnce();
    expect(container.querySelectorAll(".message-row.user")).toHaveLength(1);
    await until(() => Boolean(container.querySelector<HTMLButtonElement>('[aria-label="继续任务"], [aria-label="继续分析"]')) &&
      !container.querySelector<HTMLButtonElement>('[aria-label="继续任务"], [aria-label="继续分析"]')!.disabled, "等待显式继续入口");
    const play = container.querySelector<HTMLButtonElement>('[aria-label="继续任务"], [aria-label="继续分析"]')!;
    expect(play.textContent).toBe("▶");
    expect(play.getAttribute("type")).toBe("button");
    await act(async () => { play.click(); });
    await until(() => (container.textContent ?? "").includes("从原断点完成"), "点击 ▶ 后续跑完成");
    expect(resumeMock).toHaveBeenCalledOnce();
    expect(resumeMock.mock.calls[0][0]).toMatchObject({ turnId: "turn-manual-pause", afterSequence: 4 });
    expect(runTurnMock).toHaveBeenCalledOnce();
    expect(writes).toEqual([]);
  });

  it("an automatic status check cannot overwrite a new draft or start the planner", async () => {
    let resolveStatus!: (status: never) => void;
    statusMock.mockImplementation(() => new Promise((resolve) => { resolveStatus = resolve; }));
    const { container, writes } = await manuallyPaused(false);
    const input = container.querySelector<HTMLTextAreaElement>("textarea");
    await type(input, "尚未发送的新问题");
    await act(async () => { resolveStatus({ turnId: "turn-manual-pause", status: "interrupted", revision: 6,
      lastSequence: 9, canContinue: true, pauseType: "interrupted", reason: "检查点可用" } as never); });
    /*
     * 有未发送草稿时，右下角那颗主按钮是「发送」—— 续跑入口让位，草稿一个字都不许被
     * 覆盖或丢弃（这正是这条用例要证的事）。清空草稿后，核对出来的续跑入口必须回来，
     * 说明"让位"只是优先级，不是把结论丢了。
     */
    await until(() => Boolean(findButton(container, "发送")), "有草稿时主按钮是发送");
    expect(input?.value).toBe("尚未发送的新问题");
    expect(findResume(container)).toBeNull();
    await type(input, "");
    await until(() => Boolean(findResume(container)), "清空草稿后应出现继续入口");
    expect(resumeMock).not.toHaveBeenCalled();
    expect(runTurnMock).toHaveBeenCalledOnce();
    expect(writes).toEqual([]);
  });
  /**
   * 分支 1：**计划**。
   *
   * 变异点：把续跑结果一律报"请重新发送需求"（旧实现）。那样这条用例会红 ——
   * 因为容器里什么都不会出现，用户被要求重发。
   */
  it("pausing plan validation does not publish the late preview or reapply the completed result", async () => {
    const { container, writes } = await reachPauseCard(pauseResult());
    const plan = { id: "pause-validation", title: "迟到预览不得出现", summary: "测试计划", assumptions: [], warnings: [],
      actions: [{ type: "writeValues", sheet: "Sheet1", range: "B2", values: [["测试"]] }], acceptanceCriteria: [] };
    resumeMock.mockResolvedValue({ status: {} as never,
      outcome: { kind: "plan", provider: "model", turnId: "turn-pause-1", plan } } as never);
    let finish!: (response: Response) => void;
    let validations = 0;
    vi.stubGlobal('fetch', vi.fn(async (url: string) => {
      if (String(url).includes('/api/plans/validate')) {
        validations++;
        return await new Promise<Response>((resolve) => { finish = resolve; });
      }
      return new Response('{}', {status: 200});
    }));
    await clickContinue(container);
    await until(() => Boolean(finish), '进入结果校验');
    statusMock.mockResolvedValue({ turnId: 'turn-pause-1', status: 'completed', revision: 13,
      lastSequence: 31, canContinue: false, result: { kind: 'plan', provider: 'model', plan } } as never);
    await act(async () => container.querySelector<HTMLButtonElement>('[aria-label="暂停任务"]')!.click());
    await until(() => !container.querySelector('[aria-label="暂停任务"]'), '校验等待可暂停');
    await act(async () => finish(new Response(JSON.stringify(plan), {status: 200})));
    expect(validations).toBe(1);
    expect(container.textContent).not.toContain('迟到预览不得出现');
    expect(container.textContent).toContain('结果已保留');
    expect(writes).toEqual([]);
  });

  it("继续返回 plan → 进入预览并生成待确认的计划卡，宿主写入为 0", async () => {
    const { container, writes } = await reachPauseCard(pauseResult());
    resumeMock.mockResolvedValue({
      status: {} as never,
      outcome: {
        kind: "plan",
        provider: "model",
        turnId: "turn-pause-1",
        plan: {
          id: "plan-1",
          title: "按类别汇总",
          summary: "在 B 列填归属",
          assumptions: [],
          warnings: [],
          actions: [],
          acceptanceCriteria: []
        }
      } as never
    });

    await clickContinue(container);
    await settle();
    await settle();
    /*
     * 计划分支若被静默丢弃（旧实现），控制台既没有预览、也没有可辨认的
     * 计划卡；这里断言"面板确实对这份计划做了处理"——要么落预览，
     * 要么给出**具体**的失败原因（而不是一句笼统的"请重新发送需求"）。
     *
     * 这条用例的核心判据在最下面：**写次数为 0**。
     */
    const text = container.textContent ?? "";
    const handled =
      text.includes("按类别汇总") ||
      text.includes("计划中没有可执行操作") ||
      text.includes("无法") ||
      text.includes("校验");
    expect(handled, `计划结果未被接回处理链：${text.slice(0, 200)}`).toBe(true);
    // 反向锚点：绝不能一句"请重新发送需求"把结果丢掉（旧行为）。
    expect(text).not.toContain("继续后的结果需要重新发起处理");
    expect(writes).toEqual([]);
  });

  /**
   * 分支 2：**澄清**。
   *
   * 变异点同上。另外钉住：澄清卡必须是**可填写**的（出现输入区），
   * 而不是一句"请重新描述需求"。
   */
  it("继续返回 clarification → 落成可填写的澄清卡，不要求重发", async () => {
    const { container, writes } = await reachPauseCard(pauseResult());
    resumeMock.mockResolvedValue({
      status: {} as never,
      outcome: {
        kind: "clarification",
        clarification: {
          response: {
            kind: "clarification",
            provider: "model",
            turnId: "turn-pause-1",
            clarification: {
              id: "q-1",
              summary: "需要确认归属口径",
              question: "「甲」应归到哪个大区？",
              reason: "规则未覆盖",
              scopeLabel: "Sheet1!A2:A3",
              options: [
                { label: "华东", description: "按门店地址", resolution: "华东" },
                { label: "华南", description: "按发货仓", resolution: "华南" }
              ]
            }
          }
        }
      } as never
    });

    await clickContinue(container);
    await until(
      () => (container.textContent ?? "").includes("「甲」应归到哪个大区？"),
      "继续返回的澄清应落成卡片"
    );
    /*
     * 断言必须**限定在活动卡的操作区**，不能扫全容器。
     *
     * 历史里那张旧暂停卡仍在，且它带着正确的"已被新的任务取代"说明 ——
     * 扫全容器会把那句正确的话误判成"丢弃了结果"。判别点要打在
     * "有没有进入对应界面"上：澄清题面出现、选项在同一张澄清卡里可点。
     */
    const optionsBox = container.querySelector(".clarification-options");
    expect(optionsBox).not.toBeNull();
    const labels = Array.from(optionsBox!.querySelectorAll("button")).map(
      (button) => (button.querySelector("strong")?.textContent ?? "").trim()
    );
    expect(labels).toEqual(["华东", "华南"]);
    expect(writes).toEqual([]);
  });

  /**
   * 分支 3：**本地计算规则**。
   *
   * 这一支最微妙：服务端只给规则、不给源值，必须**按当前授权**在本机读取
   * 并计算。变异点若把它当 `answer` 渲染，用户就看不到"需要补映射/已生成预览"，
   * 但底层其实什么都没做 —— 典型的假成功。
   */
  it("继续返回 local_plan_rule → 按当前授权读取计算（本地映射澄清卡，不是 answer）", async () => {
    const { container, writes } = await reachPauseCard(pauseResult());
    /*
     * 服务端判定"规则缺映射、算不出来" → 必须落本地映射澄清卡，
     * 而不是硬算一个猜的值矩阵，也不是报"请重新发送需求"。
     */
    resumeMock.mockResolvedValue({
      status: {} as never,
      outcome: {
        kind: "clarification",
        clarification: {
          localMapping: {
            rule: {
              kind: "columnMapping",
              sourceSheet: "Sheet1",
              targetSheet: "Sheet1",
              sourceRange: "A2:A3",
              targetRange: "B2:B3",
              mapping: [],
              blankLabel: "未分类"
            },
            context: {
              summary: "按映射规则写入归属列。",
              warnings: [],
              assumptions: []
            },
            values: ["甲", "乙"],
            reason: "unmapped",
            blankCount: 0,
            allowExactMapping: true
          }
        }
      } as never
    });

    await clickContinue(container);
    await until(
      () => (container.textContent ?? "").includes("本地映射需要你决定"),
      "本地规则应落成本地映射澄清卡"
    );
    expect(writes).toEqual([]);
  });

  /**
   * 反向锚点：续跑**被服务端当场拒绝**（outcome 为 null）时不能假装成功。
   *
   * 变异点：把 `outcome === null` 当"什么都没有发生"静默返回 —— 卡片上
   * 按钮还在、用户以为点了没反应。正确行为是把**真实原因**落到卡上。
   */
  it("服务端拒绝继续（已完成/没有现场）→ 如实说明，不假装续跑过", async () => {
    const { container, writes } = await reachPauseCard(pauseResult());
    resumeMock.mockResolvedValue({
      status: {
        turnId: "turn-pause-1",
        status: "completed",
        revision: 40,
        lastSequence: 90,
        canContinue: false,
        pauseType: null,
        reason: null
      } as never,
      outcome: null as never
    });

    await clickContinue(container);
    await until(
      () => (container.textContent ?? "").includes("未能继续该任务"),
      "应如实报告无法继续"
    );
    // 关键：没有发生任何写入，也没有假装进预览。
    expect(writes).toEqual([]);
  });
});

describe("CAP-02 续跑幂等身份（第 2 条）", () => {
  /**
   * 同一次继续操作固定 `requestId`。
   *
   * 变异点：每次点击都生成新 requestId（例如把 `requestId` 的生成放进
   * `resumeWpsPlanning` 内、或每次点击都新生成）—— 超时重试就会变成
   * "另一次继续"，服务端重复追加预算、重复跑一轮规划。
   */
  it("超时重试沿用一个 requestId，不是每次点击都换新的", async () => {
    const { container } = await reachPauseCard(pauseResult());
    const seen: string[] = [];
    let attempt = 0;
    resumeMock.mockImplementation(async (input: { requestId?: string }) => {
      seen.push(String(input.requestId ?? ""));
      attempt += 1;
      if (attempt === 1) throw new Error("network timeout");
      return {
        status: {} as never,
        outcome: { kind: "answer", provider: "model", message: "完成", resultContext: null } as never
      } as never;
    });

    await clickContinue(container);
    await settle();
    // 第一次失败后重试：沿用同一个 requestId。
    await clickContinue(container);
    await settle();

    expect(seen.length).toBeGreaterThanOrEqual(2);
    // 核心断言：两次请求的身份**完全相同**。
    expect(seen[0]).not.toBe("");
    expect(seen.every((id) => id === seen[0])).toBe(true);
  });

  /**
   * 重复点击不得产生重复请求。
   *
   * 变异点：去掉同步并发闸（只靠 `busy` 状态）—— 连点两下时两帧都读到
   * `busy === false`，于是两次请求都发出去。
   */
  it("同一张卡上连点两次只产生一次继续请求（同步并发闸）", async () => {
    const { container } = await reachPauseCard(pauseResult());
    let release: (() => void) | null = null;
    resumeMock.mockImplementation(async () => {
      await new Promise<void>((resolve) => { release = resolve; });
      return {
        status: {} as never,
        outcome: { kind: "answer", provider: "model", message: "完成", resultContext: null } as never
      } as never;
    });

    const button = findResume(container)!;
    await act(async () => {
      button.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      button.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });
    await until(() => resumeMock.mock.calls.length === 1, "接续开始");
    expect(container.querySelector(".pause-details summary")?.textContent ?? container.textContent).toContain("正在继续");
    expect(findResume(container)).toBeNull();
    expect(resumeMock).toHaveBeenCalledTimes(1);
    await act(async () => { release?.(); });
    await settle();
  });

  /**
   * 「原任务仍在运行」这条分支上的并发闸（证据 6 在**新分支**下的复验）。
   *
   * 这条路径不发续跑请求，所以"只发一次请求"这个判据不适用；它真正可能重复的
   * 是**只读补发**：每点一次就多轮询一页、多把同一批事件重新灌进时间线。两次
   * 点击在同一帧里发出（React 还没来得及重绘、按钮也还没禁用），因此挡住第二次
   * 的只可能是 `resumingPauseRef` 那道同步闸。
   */
  it("原任务仍在运行时连点两次「继续」只接回一次进度", async () => {
    const { container, writes } = await reachPauseCard(pauseResult(), true);
    statusMock
      .mockResolvedValueOnce({ turnId: "turn-pause-1", status: "running", revision: 12,
        lastSequence: 30, canContinue: false } as never)
      .mockResolvedValueOnce({ turnId: "turn-pause-1", status: "completed", revision: 13,
        lastSequence: 32, canContinue: false,
        result: { kind: "answer", provider: "model", message: "接续完成" } } as never);
    vi.mocked(getWpsPlanningTaskEvents)
      .mockResolvedValueOnce({ turnId: "turn-pause-1", status: "running", active: true, lastSequence: 31, events: [{
        schemaVersion: 1, operationId: "turn-pause-1", sequence: 31, phase: "planning",
        stepId: "step-7", status: "progress", title: "步骤7继续", timestamp: "2026-10-21T00:00:00Z"
      }] } as never)
      .mockResolvedValueOnce({ turnId: "turn-pause-1", status: "completed", active: false, lastSequence: 32, events: [{
        schemaVersion: 1, operationId: "turn-pause-1", sequence: 32, phase: "planning",
        stepId: "step-8", status: "completed", title: "步骤8", timestamp: "2026-10-21T00:00:00Z"
      }] } as never);

    const button = findResume(container);
    if (!button) throw new Error("没有继续按钮");
    const pausedSegments = segmentSnapshots(container);
    await act(async () => {
      button.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      button.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await untilWithin(() => (container.textContent ?? "").includes("接续完成"), "接续完成");
    // 一次恢复 = 一次状态查询（entry）+ 两页补发 + 一次收尾状态查询。
    expect(statusMock).toHaveBeenCalledTimes(2);
    expect(vi.mocked(getWpsPlanningTaskEvents)).toHaveBeenCalledTimes(2);
    expect(resumeMock).not.toHaveBeenCalled();
    // 连点只受理一次恢复 ⇒ 只多出**一段**回复；不是两段一样的补发段。
    expect(segmentSnapshots(container)).toHaveLength(pausedSegments.length + 1);
    expect(liveSegmentCount(container)).toBeLessThanOrEqual(1);
    expect(writes).toEqual([]);
  });
});

describe("CAP-02 续跑前授权复核（第 5 条）", () => {
  /**
   * 授权范围在**等待网络响应期间**收窄 → 必须拦住，不能把旧任务结果
   * 套到新的（更小/不同的）范围上。
   *
   * 变异点：把 `targetGuard.check()` 从续跑路径上删掉 —— 那样这条会红
   * （继续照样发出去、结果照样落到新范围）。
   *
   * 做法：让查状态那一步**挂起**，在挂起期间改掉勾选范围，再放行。
   */
  it.each(["waiting_user", "running"])("等待核对期间换成同名文件 → 停止继续，不写原表 (%s)", async (status) => {
    const { container, writes } = await reachPauseCard(pauseResult());
    const originalHost = getWpsApplicationMock();
    let releaseStatus: ((value: unknown) => void) | null = null;
    statusMock.mockReturnValue(
      new Promise((resolve) => { releaseStatus = resolve as (value: unknown) => void; })
    );
    resumeMock.mockResolvedValue({
      status: {} as never,
      outcome: { kind: "answer", provider: "model", message: "完成", resultContext: null } as never
    });

    await clickContinue(container);
    await until(() => statusMock.mock.calls.length > 0, "应进入核对状态");
    // 真正改变宿主对象/路径；派发一个无人监听的自定义事件不能证明守卫生效。
    await act(async () => {
      getWpsApplicationMock.mockReturnValue({ ...originalHost,
        ActiveWorkbook: { ...originalHost.ActiveWorkbook!, FullName: "C:\\Other\\测试.xlsx", Path: "C:\\Other" } });
    });
    await act(async () => {
      releaseStatus?.({ turnId: "turn-pause-1", status, revision: 12, lastSequence: 30, canContinue: true, pauseType: "budget", reason: null });
    });
    await settle();
    // 无论面板最终如何表达"范围变了"，都**不能**出现"继续成功"的迹象。
    expect(container.textContent ?? "").not.toContain("继续后的结果需要重新发起处理");
    expect(resumeMock).not.toHaveBeenCalled();
    expect(writes).toEqual([]);
  });

  /**
   * 服务端拒绝续跑时，**授权确认不能冒充澄清答案**。
   *
   * 这条钉住的是"两件事不能混"：重新授权只推进授权版本，不是对某个问题的回答。
   * 落盘记录里不应因为一次继续而产生伪造的 `answer`。
   */
  it("继续被拒时不产生伪造的澄清答案", async () => {
    const { container } = await reachPauseCard(pauseResult());
    resumeMock.mockResolvedValue({
      status: {
        turnId: "turn-pause-1",
        status: "completed",
        revision: 40,
        lastSequence: 90,
        canContinue: false,
        pauseType: null,
        reason: null
      } as never,
      outcome: null as never
    });

    await clickContinue(container);
    await until(
      () => (container.textContent ?? "").includes("未能继续该任务"),
      "应如实报告无法继续"
    );
    const stored = localStorage.getItem("excel-bro.wps.chat.v1");
    const serialized = stored ?? "";
    // 落盘历史里不能出现"已重新确认数据范围并继续"这种把授权当答案的文字。
    expect(serialized).not.toContain("已重新确认数据范围并继续");
  });
});

describe("CAP-02 断线恢复（第 4 条）", () => {
  /**
   * 网络错误后先查原 turnId 的状态。
   *
   * 变异点：直接提示"重新发送需求"（旧实现）—— 那会重复规划。
   */
  it("规划流中断 → 拿本轮 turnId 查状态，而不是直接建议重发", async () => {
    const { host } = writeCountingHost();
    listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
    getModelSettingsMock.mockResolvedValue({} as never);
    getWpsApplicationMock.mockReturnValue(host);
    // 意图阶段上报身份；规划阶段抛连接错误。
    runTurnMock.mockImplementation(async (input) => {
      input.onProgress?.({
        stage: "intent",
        stepId: "planning-identity",
        title: "已建立规划任务身份",
        message: "身份已建立",
        turnId: "turn-live-1"
      });
      const error = new TypeError("Failed to fetch");
      throw error;
    });
    statusMock.mockResolvedValue({
      turnId: "turn-live-1",
      status: "running",
      revision: 3,
      lastSequence: 6,
      canContinue: false,
      pauseType: null,
      reason: null
    });

    const { container } = await mount(panel());
    await until(
      () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false,
      "应加载模型目录"
    );
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "生成统计预览");
    await submit(textarea);
    await until(() => statusMock.mock.calls.length > 0, "断线后应查状态");

    expect(statusMock.mock.calls[0][0]).toBe("turn-live-1");
    // 关键反向断言：不能提示用户重复发送同一需求。
    expect(container.textContent ?? "").not.toContain("可以重新发送需求");
    expect(container.textContent ?? "").toContain("仍在服务端处理中");
  });

  /**
   * 只有服务端**明确说没有这条任务**（404）才允许提示重新发送。
   *
   * 与上一条互为锚点：上一条证明"有任务 → 不许重发"，这一条证明
   * "确实没任务 → 可以重发"。只留一条的话，"永远不许重发"和
   * "永远建议重发"都能骗过其中一半。
   */
  it("服务端明确 404「没有这条任务」时才允许建议重新发送", async () => {
    const { host } = writeCountingHost();
    listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
    getModelSettingsMock.mockResolvedValue({} as never);
    getWpsApplicationMock.mockReturnValue(host);
    runTurnMock.mockImplementation(async (input) => {
      input.onProgress?.({
        stage: "intent",
        stepId: "planning-identity",
        title: "已建立规划任务身份",
        message: "身份已建立",
        turnId: "turn-missing-1"
      });
      throw new TypeError("Failed to fetch");
    });
    const { ApiRequestError } = await import("../../excel-addin/src/api");
    statusMock.mockRejectedValue(
      new ApiRequestError(404, "PLANNING_TASK_NOT_FOUND", "未找到该规划任务。", false)
    );

    const { container } = await mount(panel());
    await until(
      () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false,
      "应加载模型目录"
    );
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "生成统计预览");
    await submit(textarea);
    await until(
      () => (container.textContent ?? "").includes("可以重新发送需求"),
      "404 后应允许重发"
    );
  });
});
