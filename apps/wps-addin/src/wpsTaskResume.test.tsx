/**
 * 定向测试：「任务中断后继续」第一阶段 —— WPS 侧。
 *
 * 之前的行为是：任务窗格一重开，待答澄清就被收敛成 `unconfirmed`，卡片只剩
 * 一句"已失效；如需继续，请重新发送需求"。用户的需求、原文件、以及他已经
 * 填过的回答全都留在那儿，却没有任何入口能把它们捡回来。
 *
 * 这里逐条钉住新的闭环：
 *  1. 重载后**任务仍在**：卡片给出「确认此范围并继续」与「取消任务」，
 *     并如实展示中断原因 / 已保留的信息 / 原表写入状态。
 *  2. 继续走的是**重新授权**（重新读取当前宿主元数据并推进授权版本），
 *     不是复用面板重开前那一版授权 —— 所以不会被同一个旧对象引用检查again卡住。
 *  3. 已提交的回答跨重载存活：继续时原需求与确认内容一起重新发出。
 *  4. 取消是终态：重开后不会再被摆成"可以继续"。
 *  5. 写入状态未知时**不得**显示"原表未写入"。
 *  6. 没有恢复记录的旧卡片不被冒充成可继续分析。
 */
import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { listModels, getModelSettings, validatePlan } from "../../excel-addin/src/api";
import WpsAgentPanel from "./WpsAgentPanel";
import { runWpsReadonlyTurn } from "./wpsReadonlyAgent";
import { getWpsApplication } from "./wpsHost";

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  const discovery = await import("./wpsPendingDiscovery.testHost");
  return { ...actual, ...discovery.pendingDiscoveryApiStubs(), understandMessage: vi.fn(async () => ({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" })), listModels: vi.fn(), getModelSettings: vi.fn(), validatePlan: vi.fn() };
});
vi.mock("./wpsReadonlyAgent", () => ({ runWpsReadonlyTurn: vi.fn() }));
vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const WPS_CHAT_STORAGE_KEY = "excel-bro.wps.chat.v1";

const listModelsMock = vi.mocked(listModels);
const getModelSettingsMock = vi.mocked(getModelSettings);
const validatePlanMock = vi.mocked(validatePlan);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);
const getWpsApplicationMock = vi.mocked(getWpsApplication);
void validatePlanMock;

const firstConnection = {
  id: "connection:first",
  label: "第一个连接",
  provider: "model" as const,
  available: true,
  baseUrl: "https://example.test/v1",
  model: "test-model",
  hasApiKey: true,
  supportsVision: false
};

let mounted: Array<{ unmount: () => void }> = [];

async function act_(fn: () => void | Promise<void>) {
  await act(async () => {
    await fn();
  });
}

async function mount(element: ReactElement) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  await act_(() => {
    root.render(element);
  });
  mounted.push({ unmount: () => root.unmount() });
  return { container: host };
}

async function until(check: () => boolean, message: string, timeout = 2000) {
  const started = Date.now();
  while (Date.now() - started < timeout) {
    if (check()) return;
    await act_(() => new Promise((resolve) => setTimeout(resolve, 10)));
  }
  throw new Error(`等待超时：${message}`);
}

function textOf(container: HTMLElement) {
  return container.textContent ?? "";
}

async function click(target: HTMLElement | null) {
  if (!target) throw new Error("要点击的元素不存在");
  await act_(() => {
    target.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
}


/**
 * 选中第 index 个快捷答案并**提交**。
 *
 * 新交互下点选项只选中（想改口径还有机会），提交是右下角那颗「发送回答」。
 * 旧用例写的是"点一下就等于作答"，语义等价于这里的两次点击。
 */
async function selectOption(container: HTMLElement, index: number) {
  const options = Array.from(container.querySelectorAll<HTMLButtonElement>(".clarification-options button"));
  const option = options[index];
  if (!option) throw new Error(`没有第 ${index} 个快捷答案`);
  await click(option);
  const submitButton = container.querySelector<HTMLButtonElement>('button[aria-label="发送回答"]');
  if (!submitButton) throw new Error("没有找到「发送回答」按钮");
  await click(submitButton);
}

async function type(target: HTMLTextAreaElement | null, value: string) {
  if (!target) throw new Error("要输入的元素不存在");
  const setter = Object.getOwnPropertyDescriptor(
    HTMLTextAreaElement.prototype,
    "value"
  )?.set;
  await act_(() => {
    setter?.call(target, value);
    target.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

async function submit(target: HTMLTextAreaElement | null) {
  if (!target) throw new Error("要提交的元素不存在");
  await act_(() => {
    target.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
  });
}

function panel() {
  return (
    <WpsAgentPanel
      selectedSheetIndexes={[1]}
      scopeDock={<div className="eb-scope-dock">当前表 Sheet1</div>}
    />
  );
}

/** 表头、尺寸与数据一致的最小替身；每次重建都返回全新对象（贴合真宿主语义）。 */
function hostStub(options: { workbookName?: string; sheetNames: string[] }) {
  const workbookName = options.workbookName ?? "测试.xlsx";
  const path = "C:\\数据";

  const makeRange = () => {
    const range: Record<string, unknown> = {
      Rows: { Count: 3 },
      Columns: { Count: 2 },
      MergeCells: false
    };
    Object.defineProperties(range, {
      Value2: {
        get: () => [
          ["类别", "归属"],
          ["办公用品", "华东"],
          ["家具", "华南"]
        ]
      }
    });
    return range;
  };

  const makeSheet = (name: string, position: number) => ({
    Name: name,
    ProtectContents: false,
    CodeName: `${name}__code__${position}`,
    get Index() {
      return position + 1;
    },
    get UsedRange() {
      return makeRange();
    },
    Range: () => makeRange()
  });

  const sheets = options.sheetNames.map(makeSheet);
  const workbook = {
    Name: workbookName,
    FullName: `${path}\\${workbookName}`,
    get Path() {
      return path;
    },
    ReadOnly: false,
    Worksheets: { Count: sheets.length, Item: (index: number) => sheets[Number(index) - 1] },
    ActiveSheet: sheets[0]
  };

  const application = {
    Version: "12.0",
    get ActiveWorkbook() {
      return workbook;
    },
    get ActiveSheet() {
      return sheets[0];
    },
    PluginStorage: { getItem: () => undefined, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never;

  return { application };
}

/** 模型返回一张需求澄清卡。 */
function clarificationCard() {
  runTurnMock.mockResolvedValue({
    kind: "clarification",
    clarification: {
      response: {
        kind: "clarification",
        provider: "model",
        clarification: {
          id: `clarify-${Math.random().toString(16).slice(2)}`,
          summary: "需要明确目标",
          question: "要查询还是修改？",
          reason: "两种操作不同",
          scopeLabel: "当前表",
          options: [
            { id: "read", label: "查询", description: "只读", resolution: "查询数据" }
          ]
        }
      }
    },
    headerRows: [["类别", "归属"]]
  } as never);
}

async function pendingClarification() {
  const { container } = await mount(panel());
  await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
  const composer = container.querySelector<HTMLTextAreaElement>(".composer textarea");
  await type(composer, "帮我处理表格");
  await submit(composer);
  await until(() => Boolean(container.querySelector(".clarification-card, .composer-box .clarification-composer")), "应显示澄清卡");
  await until(() => runTurnMock.mock.calls.length === 1, "应完成这一轮");
  return container;
}

async function reopen() {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
  const reopened = await mount(panel());
  await until(
    () => reopened.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false,
    "重开应加载模型"
  );
  await until(
    () => Boolean(reopened.container.querySelector(".clarification-card, .composer-box .clarification-composer")),
    "重开后应恢复出那张卡"
  );
  return reopened.container;
}

function recoveryButtons(container: HTMLElement): HTMLButtonElement[] {
  return Array.from(
    container.querySelectorAll<HTMLButtonElement>(".clarification-recovery-actions button")
  );
}

function recoveryButton(container: HTMLElement, label: string): HTMLButtonElement {
  const found = recoveryButtons(container).find((button) =>
    (button.textContent ?? "").includes(label)
  );
  if (!found) throw new Error(`找不到恢复按钮：${label}`);
  return found;
}

function clarificationOptions(container: HTMLElement): HTMLButtonElement[] {
  return Array.from(
    container.querySelectorAll<HTMLButtonElement>(".clarification-options button")
  );
}

/** 往宿主存储里塞一段历史（用于构造"旧记录"、"写入状态未知"等场景）。 */
function seedHistory(messages: unknown[]) {
  const now = new Date().toISOString();
  localStorage.setItem(
    WPS_CHAT_STORAGE_KEY,
    JSON.stringify({
      activeConversationId: "c1",
      conversations: [
        { id: "c1", title: "历史", messages, createdAt: now, updatedAt: now }
      ]
    })
  );
}

function storedHistoryText() {
  return localStorage.getItem(WPS_CHAT_STORAGE_KEY) ?? "";
}

/** 从落盘的对话历史里取出所有任务恢复记录。 */
function storedTaskResumes(): Array<Record<string, unknown>> {
  const saved = JSON.parse(localStorage.getItem(WPS_CHAT_STORAGE_KEY) ?? "null") as {
    conversations?: Array<{ messages?: Array<{ taskResume?: Record<string, unknown> }> }>;
  } | null;
  return (saved?.conversations ?? [])
    .flatMap((conversation) => conversation.messages ?? [])
    .map((message) => message.taskResume)
    .filter(Boolean) as Array<Record<string, unknown>>;
}

/** 造一条"写入结果未知、需要先核对"的历史记录。 */
function suspendedWriteMessage(overrides: Record<string, unknown> = {}) {
  const now = new Date().toISOString();
  return {
    id: "m1",
    role: "assistant",
    text: "",
    createdAt: now,
    taskResume: {
      schemaVersion: 1,
      taskId: "task-suspended",
      phase: "executing",
      blocker: "execution_unknown",
      resumability: "reconcile",
      writeState: "needs_review",
      originalPrompt: "把这批数据写回去",
      confirmations: [],
      authorizationRevision: 2,
      createdAt: now,
      updatedAt: now,
      ...overrides
    },
    clarification: {
      summary: "需要确认",
      question: "是否继续？",
      reason: "写入结果未知",
      scopeLabel: "当前表",
      options: [],
      status: "pending"
    }
  };
}

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  listModelsMock.mockResolvedValue({
    defaultModelId: firstConnection.id,
    models: [firstConnection]
  });
  getModelSettingsMock.mockResolvedValue({} as never);
  getWpsApplicationMock.mockReturnValue({} as never);
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
});

describe("待答澄清：面板重载后任务仍然可继续", () => {
  it("重载后保留任务，展示中断原因/已保留信息/写入状态，并可确认范围后继续", async () => {
    clarificationCard();
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单", "明细"] }).application);
    await pendingClarification();

    const reopened = await reopen();
    // 身份在内存里，重开后无从证明 —— 卡收敛为"无法确认"，但**任务没有被丢掉**。
    expect(reopened.querySelector(".clarification-card.unconfirmed")).not.toBeNull();
    expect(reopened.querySelector(".clarification-card.invalidated")).toBeNull();
    // 不得谎称"数据范围已经变化"：重开本身没有给出任何变化证据。
    expect(textOf(reopened)).not.toContain("数据范围已经变化");

    // 三件事分开说：中断原因 / 已保留的信息 / 原表写入状态。
    const facts = Array.from(reopened.querySelectorAll(".clarification-recovery-facts dt"))
      .map((node) => node.textContent);
    expect(facts).toEqual(["中断原因", "已保留的信息", "原表写入状态"]);
    expect(textOf(reopened)).toContain("需要你重新确认数据范围");
    expect(textOf(reopened)).toContain("原需求已保留");
    expect(textOf(reopened)).toContain("原表未写入");
    // 当前文件与待授权工作表必须摆出来，供用户确认。
    expect(textOf(reopened)).toContain("测试.xlsx");

    // 重新授权并继续 → 触发新的一轮规划，且带的是保留的原需求。
    await click(recoveryButton(reopened, "确认此范围并继续"));
    await until(() => runTurnMock.mock.calls.length === 2, "继续应发起新的一轮");
    expect((runTurnMock.mock.calls[1][0] as { prompt: string }).prompt).toBe("帮我处理表格");
    // 旧卡原位更新为"已继续"，不再摆着可点按钮。
    expect(recoveryButtons(reopened)).toHaveLength(0);
  });

  it("用户显式取消是终态：重开后不会再被摆成可以继续", async () => {
    clarificationCard();
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单", "明细"] }).application);
    await pendingClarification();
    const reopened = await reopen();

    await click(recoveryButton(reopened, "取消任务"));
    await until(
      () => textOf(reopened).includes("已取消本次澄清"),
      "取消后应显示终态文案"
    );
    expect(reopened.querySelector(".clarification-recovery-actions")).toBeNull();

    const again = await reopen();
    expect(again.querySelector(".clarification-recovery-actions")).toBeNull();
    expect(textOf(again)).toContain("已取消本次澄清");
    expect(runTurnMock).toHaveBeenCalledTimes(1);
  });

  it("写入状态未知时不得显示「原表未写入」", async () => {
    seedHistory([
      {
        id: "m1",
        role: "assistant",
        text: "",
        createdAt: new Date().toISOString(),
        taskResume: {
          schemaVersion: 1,
          taskId: "task-unknown",
          phase: "executing",
          blocker: "execution_unknown",
          resumability: "reconcile",
          writeState: "unknown",
          originalPrompt: "把这批数据写回去",
          confirmations: [],
          authorizationRevision: 3,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        },
        clarification: {
          summary: "需要确认",
          question: "是否继续？",
          reason: "写入结果未知",
          scopeLabel: "当前表",
          options: [],
          status: "pending"
        }
      }
    ]);
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);

    const { container } = await mount(panel());
    await until(
      () => Boolean(container.querySelector(".clarification-recovery")),
      "应渲染中断卡"
    );
    expect(textOf(container)).toContain("写入状态未知，需要核对");
    expect(textOf(container)).not.toContain("原表未写入");
    // 结果不确定时先核对，不提供"直接重跑"的入口。
    expect(recoveryButtons(container).map((button) => button.textContent)).toEqual([
      "核对待续跑状态",
      "取消任务"
    ]);
  });

  it("部分写入（需要核对原表）重开后只允许核对，不得退化成重新预览后直接写", async () => {
    seedHistory([
      {
        id: "m1",
        role: "assistant",
        text: "",
        createdAt: new Date().toISOString(),
        taskResume: {
          schemaVersion: 1,
          taskId: "task-partial",
          phase: "executing",
          blocker: "execution_unknown",
          resumability: "reconcile",
          writeState: "needs_review",
          originalPrompt: "把这批数据写回去",
          confirmations: [],
          authorizationRevision: 2,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        },
        clarification: {
          summary: "需要确认",
          question: "是否继续？",
          reason: "部分写入",
          scopeLabel: "当前表",
          options: [],
          status: "pending"
        }
      }
    ]);
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);

    const { container } = await mount(panel());
    await until(
      () => Boolean(container.querySelector(".clarification-recovery")),
      "应渲染中断卡"
    );
    expect(textOf(container)).toContain("需要核对原表");
    expect(textOf(container)).not.toContain("原表未写入");
    /*
     * 中断原因必须说"写入结果无法确认"，**不能**降级成笼统的"连接中断"。
     * 这一条是有判别力的：只把 `needs_review` 从重载收敛里拿掉时，
     * resumability 仍是 reconcile，按钮文案不变 —— 只有这句原因会变。
     */
    expect(textOf(container)).toContain("写入结果无法确认");
    // 只给核对入口：重开一次不等于"可以重新生成预览并再写一遍"。
    expect(recoveryButtons(container).map((button) => button.textContent)).toEqual([
      "核对待续跑状态",
      "取消任务"
    ]);
    // 也不许出现"重新生成预览"那套说法 —— 那会诱导用户直接再写一遍。
    expect(textOf(container)).not.toContain("重新生成预览");
  });

  it("已完成写入的任务重开后不得再出现可继续入口", async () => {
    seedHistory([
      {
        id: "m1",
        role: "assistant",
        text: "",
        createdAt: new Date().toISOString(),
        taskResume: {
          schemaVersion: 1,
          taskId: "task-done",
          phase: "executing",
          blocker: "none",
          resumability: "unavailable",
          writeState: "confirmed_result",
          originalPrompt: "把这批数据写回去",
          confirmations: [],
          authorizationRevision: 2,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        }
      }
    ]);
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);

    const { container } = await mount(panel());
    await until(
      () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false,
      "重开应加载模型"
    );
    expect(container.querySelector(".clarification-recovery")).toBeNull();
    expect(recoveryButtons(container)).toEqual([]);
    expect(textOf(container)).not.toContain("继续分析");
  });

  it("点「核对待续跑状态」不得重新规划：不调用规划、不生成新写入计划", async () => {
    clarificationCard();
    seedHistory([suspendedWriteMessage()]);
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);

    const { container } = await mount(panel());
    await until(
      () => Boolean(container.querySelector(".clarification-recovery")),
      "应渲染中断卡"
    );
    // 前提锚点：点之前一次规划都没有发起，否则下面的负向断言就是空过。
    expect(runTurnMock).not.toHaveBeenCalled();

    await click(recoveryButton(container, "核对待续跑状态"));
    await until(
      () => textOf(container).includes("无法自动核对"),
      "没有跨会话执行证据时应**明说**无法自动核对"
    );

    // 核心负向：核对分支不得读数据、不得规划、不得产出写入计划。
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(container.querySelector(".wps-plan-card")).toBeNull();
    expect(textOf(container)).not.toContain("写入前确认");
    // 而且必须保持挂起：卡还在、状态仍是"需要核对原表"，不伪装成已核对/已继续。
    expect(textOf(container)).toContain("需要核对原表");
    expect(container.querySelector(".clarification-recovery")).not.toBeNull();
    expect(recoveryButton(container, "核对待续跑状态")).not.toBeNull();
  });

  it("重开后再继续：新记录沿用原 taskId 与已确认信息，不换新身份", async () => {
    clarificationCard();
    const now = new Date().toISOString();
    seedHistory([
      {
        id: "m1",
        role: "assistant",
        text: "",
        createdAt: now,
        taskResume: {
          schemaVersion: 1,
          taskId: "task-keep",
          phase: "awaiting_clarification",
          blocker: "network",
          resumability: "replan",
          writeState: "not_written",
          originalPrompt: "帮我处理表格",
          confirmations: [{ label: "统计口径", value: "按金额汇总" }],
          authorizationRevision: 1,
          createdAt: now,
          updatedAt: now
        },
        clarification: {
          summary: "需要确认",
          question: "要查询还是修改？",
          reason: "两种操作不同",
          scopeLabel: "当前表",
          options: [],
          status: "pending"
        }
      }
    ]);
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);

    const { container } = await mount(panel());
    await until(
      () => Boolean(container.querySelector(".clarification-recovery")),
      "应渲染中断卡"
    );
    await click(recoveryButton(container, "重新分析"));
    await until(() => runTurnMock.mock.calls.length === 1, "继续应发起一轮");
    await until(
      () => storedTaskResumes().some((item) => item.phase === "answer_submitted"),
      "旧卡的记录应推进到已提交"
    );

    const kept = storedTaskResumes();
    expect(kept.length).toBeGreaterThanOrEqual(2);
    /*
     * `taskIdRef` 的初值是挂载时随机生成的；继续时若不从记录里取回 taskId，
     * 新记录会挂到一个新身份上 —— "同一件事"当场断掉。这两条断言正是抓它。
     */
    expect(kept.every((item) => item.taskId === "task-keep")).toBe(true);
    // 已确认信息也必须带过来，否则用户要重答同一个问题。
    expect(kept.every((item) => (item.confirmations as unknown[]).length === 1)).toBe(true);
  });

  it("没有恢复记录的旧卡片不被冒充成可继续分析", async () => {
    seedHistory([
      {
        id: "m1",
        role: "assistant",
        text: "",
        createdAt: new Date().toISOString(),
        clarification: {
          summary: "需要确认",
          question: "要查询还是修改？",
          reason: "旧记录",
          scopeLabel: "当前表",
          options: [{ label: "查询", description: "只读" }],
          status: "pending"
        }
      }
    ]);
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);

    const { container } = await mount(panel());
    await until(() => Boolean(container.querySelector(".clarification-card, .composer-box .clarification-composer")), "应恢复出卡片");
    expect(container.querySelector(".clarification-recovery")).toBeNull();
    expect(textOf(container)).toContain("没有可用的恢复记录");
  });
});

describe("已提交答案：断线后不丢答案、不重复询问", () => {
  it("续答那一轮断线后转为可恢复中断卡，继续时带上已确认的回答", async () => {
    clarificationCard();
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单", "明细"] }).application);
    const container = await pendingClarification();

    // 让续答那一轮断线（连接类错误）。
    runTurnMock.mockRejectedValueOnce(new Error("Failed to fetch"));
    await selectOption(container, 0);
    await until(
      () => container.querySelector(".clarification-card.unconfirmed") !== null,
      "断线后应转为可恢复的中断卡"
    );
    await until(
      () => Boolean(container.querySelector(".clarification-recovery")),
      "应给出恢复卡"
    );
    expect(textOf(container)).toContain("连接中断");
    // 已提交的回答被记进恢复记录，界面如实说保留了几项。
    expect(textOf(container)).toContain("已提交的 1 项回答");

    // 重开之后仍然带着这条回答，不需要用户再回答一次同一个问题。
    const reopened = await reopen();
    await until(
      () => Boolean(reopened.querySelector(".clarification-recovery")),
      "重开后恢复卡仍在"
    );
    expect(textOf(reopened)).toContain("已提交的 1 项回答");

    // 继续：原需求 + 已确认的回答一起重新发出。
    runTurnMock.mockResolvedValue({
      kind: "answer",
      message: "已按你的确认处理",
      resultContext: null,
      provider: "model"
    } as never);
    await click(recoveryButton(reopened, "重新分析"));
    await until(() => runTurnMock.mock.calls.length === 3, "继续应发起新的一轮");
    const prompt = (runTurnMock.mock.calls[2][0] as { prompt: string }).prompt;
    expect(prompt).toContain("帮我处理表格");
    expect(prompt).toContain("补充确认：查询数据");
  });
});

/**
 * 「重新确认数据范围」是**授权动作**，不是对那道题的回答。
 *
 * 早期实现把它写进 `clarification.answer`（"已重新确认数据范围并继续"）：
 * 一是把用户真正的回答覆盖掉，二是让"写入资格"看起来已经具备 —— 而重新
 * 授权其实只证明了"你现在授权这几张表"。下面三条分别钉住这两种后果。
 */
describe("重新授权不是澄清答案", () => {
  it("回答提交时身份变成 unknown：真实答案先保存，重新授权时带上答案而不是重问", async () => {
    const host = hostStub({ sheetNames: ["订单"] });
    getWpsApplicationMock.mockReturnValue(host.application);
    clarificationCard();
    const container = await pendingClarification();
    // 真正改变 workbook/worksheet 包装对象，不只是改变 Range 包装。
    const app = host.application as unknown as { ActiveWorkbook: { Worksheets: { Item: (index: number) => unknown } } };
    const book = app.ActiveWorkbook;
    Object.defineProperty(app, "ActiveWorkbook", { get: () => ({
      ...book, Worksheets: { ...book.Worksheets, Item: (index: number) => ({ ...(book.Worksheets.Item(index) as object) }) }
    }) });
    await type(container.querySelector<HTMLTextAreaElement>('.composer textarea'), "逐行映射，空白不写入");
    await click(container.querySelector<HTMLButtonElement>('button[aria-label="发送回答"]'));
    await until(() => Boolean(container.querySelector('.clarification-recovery')), "等待重新授权");
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    expect(storedTaskResumes().at(-1)?.confirmations).toEqual([
      expect.objectContaining({ value: "逐行映射，空白不写入" })
    ]);
    expect(textOf(container)).toContain("已提交的 1 项回答");
    await click(recoveryButton(container, "确认此范围并继续"));
    await until(() => Boolean(container.querySelector('.authorized-target-dialog')), "恢复先确认弱宿主目标证据");
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    expect(container.querySelector('.authorized-target-dialog input[type="radio"]')).toBeNull();
    await click(Array.from(container.querySelectorAll<HTMLButtonElement>('.authorized-target-dialog button'))
      .find((button) => button.textContent === "确认范围并继续") ?? null);
    await until(() => runTurnMock.mock.calls.length === 2, "确认后发起新一轮");
    expect(runTurnMock.mock.calls[1][0].prompt).toContain("补充确认：逐行映射，空白不写入");
    expect(storedTaskResumes().at(-1)?.confirmations).toHaveLength(1);
  });

  it("连续两次澄清再重开：新卡仍带全部确认信息与原始需求，不把确认重复拼进原需求", async () => {
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);
    clarificationCard();
    const container = await pendingClarification();
    clarificationCard(); // 下一轮是新问题，具有不同的 questionId。
    await selectOption(container, 0);
    await until(() => runTurnMock.mock.calls.length === 2 && clarificationOptions(container).length > 0, "第二张澄清卡");
    expect(storedTaskResumes().at(-1)?.originalPrompt).toBe("帮我处理表格");
    expect(storedTaskResumes().at(-1)?.confirmations).toHaveLength(1);
    await type(container.querySelector<HTMLTextAreaElement>('.composer textarea'), "空白保留原值");
    await click(container.querySelector<HTMLButtonElement>('button[aria-label="发送回答"]'));
    await until(() => runTurnMock.mock.calls.length === 3, "第三轮");
    expect(storedTaskResumes().at(-1)?.confirmations).toHaveLength(2);
    const reopened = await reopen();
    await click(recoveryButton(reopened, "确认此范围并继续"));
    await until(() => runTurnMock.mock.calls.length === 4, "重开继续");
    const prompt = runTurnMock.mock.calls[3][0].prompt;
    expect(prompt.match(/补充确认：查询数据/g)).toHaveLength(1);
    expect(prompt.match(/补充确认：空白保留原值/g)).toHaveLength(1);
  });
  /** 造一条"待继续"的历史：带一道题与一条已确认信息。 */
  function pendingTaskMessage(overrides: {
    clarification: Record<string, unknown>;
    confirmations?: Array<Record<string, unknown>>;
  }) {
    const now = new Date().toISOString();
    return {
      id: "m1",
      role: "assistant",
      text: "",
      createdAt: now,
      taskResume: {
        schemaVersion: 1,
        taskId: "task-reauthorize",
        phase: "awaiting_clarification",
        blocker: "network",
        resumability: "replan",
        writeState: "not_written",
        originalPrompt: "把类别列翻译成归属",
        confirmations: overrides.confirmations ?? [
          { label: "空白处理", value: "空白保留原值", questionId: "column-mapping:blankPolicyNeeded" }
        ],
        authorizationRevision: 1,
        createdAt: now,
        updatedAt: now
      },
      clarification: {
        summary: "需要确认",
        question: "空白格怎么处理？",
        reason: "源区域有空白",
        scopeLabel: "当前表",
        options: [],
        status: "unconfirmed",
        ...overrides.clarification
      }
    };
  }

  function storedClarifications(): Array<Record<string, unknown>> {
    const saved = JSON.parse(localStorage.getItem(WPS_CHAT_STORAGE_KEY) ?? "null") as {
      conversations?: Array<{
        messages?: Array<{ clarification?: Record<string, unknown> }>;
      }>;
    } | null;
    return (saved?.conversations ?? [])
      .flatMap((conversation) => conversation.messages ?? [])
      .map((message) => message.clarification)
      .filter(Boolean) as Array<Record<string, unknown>>;
  }

  it("已有真实答案时：答案原样保留，不会被套话覆盖", async () => {
    seedHistory([
      pendingTaskMessage({ clarification: { answer: "空白保留原值" } })
    ]);
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);

    const { container } = await mount(panel());
    await until(
      () => Boolean(container.querySelector(".clarification-recovery")),
      "应渲染中断卡"
    );
    await click(recoveryButton(container, "重新分析"));
    await until(() => runTurnMock.mock.calls.length === 1, "继续应发起一轮");

    const cards = storedClarifications();
    expect(cards.length).toBeGreaterThan(0);
    for (const card of cards) {
      // 负向：不许出现那句套话。
      expect(String(card.answer ?? "")).not.toContain("重新确认数据范围");
      // 正向锚点：真实答案还在（断言"没有套话"才不是空过）。
      if (card.answer !== undefined) {
        expect(card.answer).toBe("空白保留原值");
      }
    }
  });

  it("还没有真实答案时：不伪造'已回答'", async () => {
    seedHistory([pendingTaskMessage({ clarification: {} })]);
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);

    const { container } = await mount(panel());
    await until(
      () => Boolean(container.querySelector(".clarification-recovery")),
      "应渲染中断卡"
    );
    await click(recoveryButton(container, "重新分析"));
    await until(() => runTurnMock.mock.calls.length === 1, "继续应发起一轮");

    for (const card of storedClarifications()) {
      expect(card.answer).toBeUndefined();
    }
  });

  it("重新授权后已确认信息仍然带过去，不会让用户重答", async () => {
    seedHistory([
      pendingTaskMessage({
        clarification: { answer: "空白保留原值" },
        confirmations: [
          { label: "空白处理", value: "空白保留原值", questionId: "column-mapping:blankPolicyNeeded" },
          { label: "映射补充", value: "丙→费用", questionId: "column-mapping:matchedDelimiterMissingCategories" }
        ]
      })
    ]);
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);

    const { container } = await mount(panel());
    await until(
      () => Boolean(container.querySelector(".clarification-recovery")),
      "应渲染中断卡"
    );
    await click(recoveryButton(container, "重新分析"));
    await until(() => runTurnMock.mock.calls.length === 1, "继续应发起一轮");

    const prompt = (runTurnMock.mock.calls[0][0] as { prompt: string }).prompt;
    expect(prompt).toContain("把类别列翻译成归属");
    expect(prompt).toContain("补充确认：空白保留原值");
    expect(prompt).toContain("补充确认：丙→费用");
  });

  it("宿主每次访问都返回新的包装对象时，显式重授权后仍能到达新的一轮", async () => {
    // 替身的 `UsedRange` / `Range` 每次都新建对象 —— 正是"同一张表换了包装"。
    // 若把引用不相等当成"用户换了表"，这里会永远卡住、出不了新一轮。
    seedHistory([pendingTaskMessage({ clarification: { answer: "空白保留原值" } })]);
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单"] }).application);

    const { container } = await mount(panel());
    await until(
      () => Boolean(container.querySelector(".clarification-recovery")),
      "应渲染中断卡"
    );
    await click(recoveryButton(container, "重新分析"));
    await until(() => runTurnMock.mock.calls.length === 1, "重授权后应发起新一轮");
    // 不得谎称范围已变 —— 引用换了不等于表换了。
    expect(textOf(container)).not.toContain("数据范围已经变化");
  });
});

describe("隐私边界：恢复记录只带定位、阶段、原需求与用户确认", () => {
  it("落盘的历史里没有完整路径、也没有来源数据行", async () => {
    clarificationCard();
    getWpsApplicationMock.mockReturnValue(hostStub({ sheetNames: ["订单", "明细"] }).application);
    await pendingClarification();
    await reopen();

    const stored = storedHistoryText();
    // 正向锚点：恢复记录**确实**落盘了，"没有越界内容"才是有意义的断言。
    expect(stored).toContain("taskResume");
    expect(stored).toContain("task-");
    // 路径只在内存比对；恢复记录与对话历史都不落盘。
    expect(stored).not.toContain("C:\\\\数据");
    // 替身表里的单元格内容不得经由恢复记录落盘。
    expect(stored).not.toContain("华东");
    expect(stored).not.toContain("华南");
    expect(stored).not.toContain("办公用品");

    // 更强的一条：落盘的恢复记录键集必须**恰好**是白名单那一组，
    // 多一个键（矩阵、对象引用、路径）都会在这里变红。
    const parsed = JSON.parse(stored) as {
      conversations: Array<{ messages: Array<{ taskResume?: Record<string, unknown> }> }>;
    };
    const records = parsed.conversations
      .flatMap((conversation) => conversation.messages)
      .map((message) => message.taskResume)
      .filter((record): record is Record<string, unknown> => Boolean(record));
    expect(records.length).toBeGreaterThan(0);
    for (const record of records) {
      expect(Object.keys(record).sort()).toEqual(
        [
          "authorizationRevision",
          "blocker",
          "confirmations",
          "createdAt",
          "originalPrompt",
          "phase",
          "resumability",
          "schemaVersion",
          "taskId",
          "updatedAt",
          "writeState"
        ].sort()
      );
    }
  });
});
