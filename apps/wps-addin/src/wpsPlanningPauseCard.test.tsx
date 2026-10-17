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
import { listModels, getModelSettings } from "../../excel-addin/src/api";
import { getWpsApplication } from "./wpsHost";
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
  return { ...actual, getWpsPlanningTaskStatus: vi.fn() };
});

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  return {
    ...actual,
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
 * 面板是默认导出，必须在 mock 之后动态引入 —— 静态 import 会被提升到
 * `vi.mock` 之前，于是面板拿到的是真实实现。
 */
const { default: WpsAgentPanel } = await import("./WpsAgentPanel");
const { runWpsReadonlyTurn } = await import("./wpsReadonlyAgent");
const { getWpsPlanningTaskStatus } = await import("./wpsPlanningStream");

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
  return Array.from(container.querySelectorAll("button")).find(
    (button) => (button.textContent ?? "").trim() === label
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

describe("CAP-02 规划暂停卡", () => {
  it("可继续时画按钮，且状态与操作分属不同节点（变异点：状态与按钮混在一处）", async () => {
    const { container } = await reachPauseCard(pauseResult());
    const card = container.querySelector(".wps-planning-pause-card");
    expect(card).not.toBeNull();

    // 状态独占 role=status 节点。
    const status = card!.querySelector('[role="status"]');
    expect(status).not.toBeNull();
    expect(status!.textContent).toContain("本轮处理时间已达上限");

    // 操作区里只有 <button>，没有把状态文字塞进去。
    const actions = card!.querySelector(".wps-planning-pause-actions");
    expect(actions).not.toBeNull();
    const buttons = actions!.querySelectorAll("button");
    expect(buttons).toHaveLength(1);
    expect(buttons[0].textContent).toBe("继续任务");
    // 判据：按钮自己的文本就是动作，不含状态描述。
    expect(status!.textContent).not.toContain("继续任务");
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
    // 关键：没有可点的"继续任务"。
    expect(findButton(container, "继续任务")).toBeNull();
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
    expect(findButton(container, "继续任务"), "核对完成前不该有可点按钮").toBeNull();
    expect(container.querySelector(".wps-planning-pause-actions")).toBeNull();
    // 放行，避免悬挂的 promise 泄漏到后续用例。
    await act(async () => {
      releaseStatus?.({ turnId: "turn-pause-1", status: "waiting_user", revision: 12, lastSequence: 30, canContinue: true, pauseType: "budget", reason: null });
    });
  });

  it("核对通过后按钮恢复可用（用的是服务端 revision 而不是落盘的那份）", async () => {
    await seedHistoryThenReload();
    const { container } = await mount(panel());
    await until(() => Boolean(findButton(container, "继续任务")), "核对通过后应恢复继续按钮");
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
    expect(findButton(container, "继续任务")).toBeNull();
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
    expect(findButton(container, "继续任务")).toBeNull();
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
    expect(findButton(container, "继续任务")).toBeNull();
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
    expect(findButton(container, "继续任务")).toBeNull();
  });
});
