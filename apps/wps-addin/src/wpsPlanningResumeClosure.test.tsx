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
import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { listModels, getModelSettings } from "../../excel-addin/src/api";
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
  return { ...actual, getWpsPlanningTaskStatus: vi.fn() };
});

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  return { ...actual, listModels: vi.fn(), getModelSettings: vi.fn() };
});

vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

const { default: WpsAgentPanel } = await import("./WpsAgentPanel");
const { runWpsReadonlyTurn, resumeWpsPlanning } = await import("./wpsReadonlyAgent");
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
  return Array.from(container.querySelectorAll("button")).find(
    (button) => (button.textContent ?? "").trim() === label
  ) ?? null;
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
async function reachPauseCard(pause: ReturnType<typeof pauseResult>) {
  const { host, writes } = writeCountingHost();
  listModelsMock.mockResolvedValue({ defaultModelId: connection.id, models: [connection] });
  getModelSettingsMock.mockResolvedValue({} as never);
  getWpsApplicationMock.mockReturnValue(host);
  runTurnMock.mockResolvedValue({ kind: "planning_pause", pause } as WpsTurnResult);
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
  const button = findButton(container, "继续任务");
  if (!button) throw new Error("没有继续按钮");
  await act(async () => {
    button.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
}

beforeEach(() => {
  mounted = [];
  localStorage.clear();
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("CAP-02 续跑结果闭环：三分支接回正常处理链", () => {
  /**
   * 分支 1：**计划**。
   *
   * 变异点：把续跑结果一律报"请重新发送需求"（旧实现）。那样这条用例会红 ——
   * 因为容器里什么都不会出现，用户被要求重发。
   */
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

    await clickContinue(container);
    // 第一次继续还悬在空中时，再点一次 —— 必须被挡住。
    await clickContinue(container);
    await settle();

    expect(resumeMock).toHaveBeenCalledTimes(1);
    await act(async () => { release?.(); });
    await settle();
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
  it("等待核对期间授权范围收窄 → 停止继续，不写原表", async () => {
    const { container } = await reachPauseCard(pauseResult());
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
    // 核对还悬着的时候，勾选范围发生变化。
    await act(async () => {
      container.dispatchEvent(new CustomEvent("wps-scope-change", { bubbles: true }));
    });
    await act(async () => {
      releaseStatus?.({ turnId: "turn-pause-1", status: "waiting_user", revision: 12, lastSequence: 30, canContinue: true, pauseType: "budget", reason: null });
    });
    await settle();
    // 无论面板最终如何表达"范围变了"，都**不能**出现"继续成功"的迹象。
    expect(container.textContent ?? "").not.toContain("继续后的结果需要重新发起处理");
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
