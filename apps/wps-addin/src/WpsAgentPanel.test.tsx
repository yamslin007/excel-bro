import { readFileSync } from "node:fs";
import { act, useEffect, useState, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MODEL_STORAGE_KEY } from "../../excel-addin/src/conversation";
import { useModelManagement } from "../../excel-addin/src/hooks/useModelManagement";
import { getModelSettings, listModels, validatePlan } from "../../excel-addin/src/api";
import WpsAgentPanel from "./WpsAgentPanel";
import WpsModelSettingsPanel from "./WpsModelSettingsPanel";
import { runWpsReadonlyTurn } from "./wpsReadonlyAgent";
import { getWpsApplication } from "./wpsHost";

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("../../excel-addin/src/api")>();
  return {
    ...actual,
    listModels: vi.fn(),
    getModelSettings: vi.fn(),
    validatePlan: vi.fn()
  };
});

vi.mock("./wpsReadonlyAgent", () => ({
  runWpsReadonlyTurn: vi.fn()
}));

vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

// React 19 从 react 包导出 act；这里刻意不用 @testing-library/react ——
// 它只装在 excel-addin 工作区，会带进根目录那份 React（19.2.x），
// 与本工作区组件使用的 React（19.1.x）形成双实例，触发 Invalid hook call。
(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true;

// jsdom 不实现 scrollIntoView；面板只用它把新消息滚到底部，
// 在测试里补一个空实现即可，不需要断言滚动行为。
if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const styles = readFileSync("src/styles.css", "utf8");
const excelStyles = readFileSync("../excel-addin/src/styles.css", "utf8");
const listModelsMock = vi.mocked(listModels);
const getModelSettingsMock = vi.mocked(getModelSettings);
const validatePlanMock = vi.mocked(validatePlan);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);
const getWpsApplicationMock = vi.mocked(getWpsApplication);

const firstConnection = {
  id: "connection:first",
  label: "第一个连接",
  provider: "model" as const,
  available: true,
  supportsVision: false
};

const secondConnection = {
  id: "connection:second",
  label: "第二个连接",
  provider: "model" as const,
  available: true,
  supportsVision: true
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
      act(() => {
        root.unmount();
      });
      container.remove();
    }
  };
  mounted.push(entry);
  return entry;
}

/** 让已排队的 promise（模型目录读取等）落定，并让 React 处理由此产生的更新。 */
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

function textOf(container: HTMLElement): string {
  return container.textContent ?? "";
}

function findButton(
  container: HTMLElement,
  name: string
): HTMLButtonElement | null {
  return (
    Array.from(container.querySelectorAll("button")).find(
      (item) => (item.textContent ?? "").trim() === name
    ) ?? null
  );
}

async function click(element: Element | null): Promise<void> {
  if (!element) throw new Error("要点击的元素不存在");
  await act(async () => {
    element.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
  await settle();
}

/** 受控输入必须走原生 value setter，否则 React 收不到这次变更。 */
async function type(
  element: HTMLTextAreaElement | null,
  value: string
): Promise<void> {
  if (!element) throw new Error("没有找到输入框");
  const setter = Object.getOwnPropertyDescriptor(
    HTMLTextAreaElement.prototype,
    "value"
  )?.set;
  await act(async () => {
    setter?.call(element, value);
    element.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

async function submit(element: HTMLTextAreaElement | null): Promise<void> {
  const form = element?.closest("form");
  if (!form) throw new Error("没有找到输入区表单");
  await act(async () => {
    form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
  });
  await settle();
}

function panel() {
  return (
    <WpsAgentPanel
      selectedSheetIndexes={[1]}
      scopeDock={<div className="eb-scope-dock">当前表 Sheet1</div>}
    />
  );
}

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  getWpsApplicationMock.mockReturnValue({} as never);
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
});

describe("WPS Excel Bro chat UI", () => {
  it("shows the Excel-style chat shell and keeps guarded write boundaries visible", () => {
    const markup = renderToStaticMarkup(panel());

    expect(markup).toContain("wps-chat-shell chat-shell");
    expect(markup).toContain("model-picker-trigger");
    const modelTrigger =
      markup.match(/<button class="model-picker-trigger"([^>]*)>/)?.[1] ?? "";
    expect(modelTrigger).not.toContain("disabled");
    // 目录尚未加载时不得出现任何“基础模式”一类占位模型。
    expect(markup).not.toContain("基础模式");
    expect(markup).toContain("正在读取模型配置");
    expect(markup).toContain('class="composer"');
    expect(markup).toContain('class="app-logo"');
    expect(markup).toContain('aria-label="隐藏格仔"');
    expect(markup).toContain("pet-float mood-idle");
    expect(markup).toContain("当前表 Sheet1");
    expect(markup).toContain("新对话");
    expect(markup).toContain("工具");
    expect(markup).toContain("备份");
    expect(markup).toContain("当前仍可查询");
    expect(markup).toContain("theme-panel");
    expect(markup).toContain("预设主题");
    expect(markup).toContain("自定义背景");
    expect(markup).not.toContain("删除工作表");
  });

  it("offers the add-model entry when the catalog loads empty", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });

    const { container } = await mount(panel());
    await until(
      () => textOf(container).includes("尚未配置模型"),
      "目录加载成功后应显示未配置状态"
    );

    expect(textOf(container)).toContain("添加模型");
    expect(textOf(container)).not.toContain("基础模式");
  });

  it("keeps the typed task and guides to add a model instead of answering", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });

    const { container } = await mount(panel());
    await until(
      () => textOf(container).includes("尚未配置模型"),
      "目录加载成功后应显示未配置状态"
    );

    const textarea = container.querySelector("textarea");
    await type(textarea, "按部门统计平均分");
    await submit(textarea);

    expect(textOf(container)).toContain("尚未配置模型。添加一个模型连接后");
    // 任务文本必须留在输入框里，不能因为缺少模型就被清掉。
    expect(textarea?.value).toBe("按部门统计平均分");
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(container.querySelector(".message-row.system")).toBeNull();
  });

  it("explains a disconnected backend without exposing Failed to fetch and keeps the task", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockRejectedValue(new TypeError("Failed to fetch"));
    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");

    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "按部门统计平均分");
    await submit(textarea);
    await until(() => textOf(container).includes("无法连接 Excel Bro 后端"), "应解释连接失败");

    expect(textOf(container)).not.toContain("Failed to fetch");
    expect(textarea?.value).toBe("按部门统计平均分");
    expect(textOf(container)).toContain("本次未修改工作表");
  });

  it("does not restore an executable plan from saved chat history", async () => {
    const now = new Date().toISOString();
    localStorage.setItem("excel-bro.wps.chat.v1", JSON.stringify({
      activeConversationId: "old-chat",
      conversations: [{
        id: "old-chat", title: "修改 A1", createdAt: now, updatedAt: now,
        messages: [
          { id: "question", role: "user", text: "修改 A1", createdAt: now },
          { id: "old-plan", role: "assistant", text: "", createdAt: now,
            planPreview: { title: "更新 A1", status: "pending" } }
        ]
      }]
    }));
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    const { container } = await mount(panel());
    expect(textOf(container)).toContain("计划预览已失效");
    expect(container.querySelector(".wps-plan-card")).toBeNull();
    expect(findButton(container, "确认写入")).toBeNull();
  });

  it("does not present a refreshed in-flight WPS activity as still running", async () => {
    const now = new Date().toISOString();
    localStorage.setItem("excel-bro.wps.chat.v1", JSON.stringify({
      activeConversationId: "old-chat",
      conversations: [{
        id: "old-chat", title: "修改 A1", createdAt: now, updatedAt: now,
        messages: [{
          id: "activity", role: "assistant", text: "", createdAt: now,
          activityLog: { totalMs: 120, steps: [{ id: "writing:0", label: "执行第 1 个写入动作", elapsedMs: 120, status: "progress" }] }
        }]
      }]
    }));
    const { container } = await mount(panel());
    const timeline = container.querySelector(".task-progress-timeline");
    expect(timeline?.getAttribute("data-status")).toBe("interrupted");
    expect(timeline?.textContent).toContain("此过程未取得终态");
    expect(timeline?.classList.contains("is-live")).toBe(false);
  });

  it("shows task progress beside the question that started it", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    let finishTurn: ((value: Awaited<ReturnType<typeof runWpsReadonlyTurn>>) => void) | undefined;
    runTurnMock.mockImplementation((input) => {
      input.onProgress?.({ stage: "intent", stepId: "intent", title: "理解需求", message: "正在判断任务类型" });
      return new Promise((resolve) => { finishTurn = resolve; });
    });
    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");

    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计当前表");
    await submit(textarea);
    const question = Array.from(container.querySelectorAll(".message-row.user"))
      .find((item) => item.textContent?.includes("统计当前表"));
    const processMessage = question?.nextElementSibling;
    expect(processMessage?.classList.contains("message-row")).toBe(true);
    expect(processMessage?.querySelector(".task-progress-timeline.is-live")).not.toBeNull();
    expect(processMessage?.textContent).toContain("理解需求");

    await act(async () => {
      finishTurn?.({ kind: "answer", message: "统计完成", resultContext: null, provider: "model" });
    });
    expect(processMessage?.querySelector(".task-progress-timeline.is-live")).toBeNull();
    expect(processMessage?.querySelector(".task-progress-timeline.activity-log-card")).not.toBeNull();
    expect(processMessage?.textContent).toContain("已完成");
    const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(saved.conversations[0].messages.find((message: { activityLog?: { outcome?: string } }) =>
      message.activityLog?.outcome === "completed")).toBeDefined();
  });

  it("copies both sides of the conversation when the embedded clipboard API rejects", async () => {
    const clipboardDescriptor = Object.getOwnPropertyDescriptor(navigator, "clipboard");
    const commandDescriptor = Object.getOwnPropertyDescriptor(document, "execCommand");
    const writeText = vi.fn().mockRejectedValue(new Error("not allowed in WPS WebView"));
    const legacyCopy = vi.fn().mockReturnValue(true);
    Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText } });
    Object.defineProperty(document, "execCommand", { configurable: true, value: legacyCopy });
    try {
      listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
      runTurnMock.mockResolvedValue({ kind: "answer", message: "这是一条回答", resultContext: null, provider: "model" });
      const { container } = await mount(panel());
      await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
      const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
      await type(textarea, "这是一条提问");
      await submit(textarea);
      await until(() => textOf(container).includes("这是一条回答"), "应收到回答");

      await click(container.querySelector(".message-row.user .message-copy-button"));
      expect(legacyCopy).toHaveBeenCalledTimes(1);
      expect(container.querySelector(".message-row.user .message-copy-button")?.textContent).toContain("已复制");
      const answer = Array.from(container.querySelectorAll(".message-row.assistant"))
        .find((item) => item.textContent?.includes("这是一条回答"));
      await click(answer?.querySelector(".message-copy-button") ?? null);
      expect(writeText).toHaveBeenCalledTimes(2);
      expect(legacyCopy).toHaveBeenCalledTimes(2);
      expect(answer?.querySelector(".message-copy-button")?.textContent).toContain("已复制");
    } finally {
      if (clipboardDescriptor) Object.defineProperty(navigator, "clipboard", clipboardDescriptor);
      else Reflect.deleteProperty(navigator, "clipboard");
      if (commandDescriptor) Object.defineProperty(document, "execCommand", commandDescriptor);
      else Reflect.deleteProperty(document, "execCommand");
    }
  });

  it("shows query source, row count and warnings instead of dropping result metadata", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "answer", provider: "local", message: "查询完成",
      resultContext: {
        kind: "table", title: "按类别统计", headers: ["类别", "数量"],
        rows: [["A", 2]], sourceSheets: ["Sheet1"], warnings: ["仅统计非空类别"]
      }
    });
    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计类别");
    await submit(textarea);
    await until(() => textOf(container).includes("查询完成"), "应显示查询结果");
    expect(container.querySelector(".wps-query-result-meta")?.textContent).toContain("按类别统计");
    expect(container.querySelector(".wps-query-result-meta")?.textContent).toContain("显示 1 / 1 行");
    expect(container.querySelector(".wps-query-result-meta")?.textContent).toContain("数据来源：Sheet1");
    expect(container.querySelector(".wps-query-result-warnings")?.textContent).toContain("仅统计非空类别");
  });

  it("sends the model id shown in the picker and migrates a stale local record", async () => {
    localStorage.setItem(MODEL_STORAGE_KEY, "local");
    listModelsMock.mockResolvedValue({
      defaultModelId: secondConnection.id,
      models: [firstConnection, secondConnection]
    });
    runTurnMock.mockResolvedValue({
      kind: "answer",
      message: "已统计完成。",
      resultContext: null,
      provider: "model"
    });

    const { container } = await mount(panel());
    await until(
      () => textOf(container).includes("模型：第二个连接"),
      "应解析出服务端声明的默认真实连接"
    );

    const textarea = container.querySelector("textarea");
    await type(textarea, "按部门统计平均分");
    await submit(textarea);

    await until(
      () => runTurnMock.mock.calls.length === 1,
      "有可用模型时应发起只读查询"
    );
    // 请求携带的必须是目录里解析出来的真实连接：既不是旧 local，
    // 也不是列表首项（首项是 connection:first）。
    expect(runTurnMock.mock.calls[0][0].modelId).toBe(secondConnection.id);
    // 旧的 local 记录已被迁移出本地存储。
    expect(localStorage.getItem(MODEL_STORAGE_KEY)).toBe(secondConnection.id);
  });

  it("shows no card for ordinary chat and accepts free text only when clarification is needed", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock
      .mockResolvedValueOnce({ kind: "answer", message: "你好！", resultContext: null, provider: "model" })
      .mockResolvedValueOnce({
        kind: "clarification",
        response: {
          kind: "clarification",
          provider: "model",
          clarification: {
            id: "clarify-1", summary: "需要明确目标", question: "想怎样处理？",
            reason: "不同操作结果不同", scopeLabel: "当前表",
            options: [
              { id: "read", label: "查询", description: "只读分析", resolution: "查询数据" },
              { id: "write", label: "修改", description: "预览后修改", resolution: "修改数据" }
            ]
          }
        }
      })
      .mockResolvedValueOnce({ kind: "answer", message: "我来按你的补充说明处理。", resultContext: null, provider: "model" });

    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "你好");
    await submit(textarea);
    await until(() => textOf(container).includes("你好！"), "问候应直接得到回复");
    expect(container.querySelector(".clarification-card")).toBeNull();

    await type(textarea, "帮我处理表格");
    await submit(textarea);
    await until(() => Boolean(container.querySelector(".clarification-card")), "关键歧义才显示卡片");
    const custom = container.querySelector<HTMLTextAreaElement>('[aria-label="其他想法"]');
    await type(custom, "只统计记录数，不修改表格");
    await click(findButton(container, "发送"));
    await until(() => runTurnMock.mock.calls.length === 3, "补充说明应继续原任务");
    expect(runTurnMock.mock.calls[2][0].prompt).toContain("补充确认：只统计记录数，不修改表格");
    expect(runTurnMock.mock.calls[2][0].intentConfirmed).toBe(true);
    expect(container.querySelector(".clarification-card.resolved")).not.toBeNull();
    expect(textOf(container)).toContain("已按「只统计记录数，不修改表格」继续");
    const stored = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(stored.conversations[0].messages.some((message: { clarification?: { status: string } }) =>
      message.clarification?.status === "resolved"
    )).toBe(true);
    expect(JSON.stringify(stored)).not.toContain("修改数据");
  });

  it("keeps a cancelled clarification in history without leaving live choices", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "clarification",
      response: {
        kind: "clarification",
        provider: "model",
        clarification: {
          id: "clarify-cancel",
          summary: "需要明确目标",
          question: "要查询还是修改？",
          reason: "两种操作不同",
          scopeLabel: "当前表",
          options: [{ id: "read", label: "查询", description: "只读", resolution: "查询数据" }]
        }
      }
    });
    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const composer = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(composer, "帮我处理表格");
    await submit(composer);
    await until(() => Boolean(container.querySelector(".clarification-card.pending")), "应显示澄清卡");
    await click(findButton(container, "取消本次澄清"));
    expect(container.querySelector(".clarification-card.cancelled")).not.toBeNull();
    expect(container.querySelector(".clarification-card button")).toBeNull();
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    const stored = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(stored.conversations[0].messages.some((message: { clarification?: { status: string } }) =>
      message.clarification?.status === "cancelled"
    )).toBe(true);
  });

  it("shows a stop action in the composer only while a read-only turn is running", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockImplementation(({ signal }) => new Promise((_resolve, reject) => {
      signal?.addEventListener("abort", () => reject(new DOMException("stopped", "AbortError")));
    }));
    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const composer = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(composer, "统计当前表");
    await submit(composer);
    await until(() => Boolean(container.querySelector('[aria-label="停止任务"]')), "运行时应显示停止入口");
    await click(container.querySelector('[aria-label="停止任务"]'));
    await until(() => textOf(container).includes("任务已停止"), "停止后应有明确反馈");
    expect(container.querySelector('[aria-label="停止任务"]')).toBeNull();
  });

  it("resolves follow-current authorization again at send time", async () => {
    const resolveSelectedSheetIndexes = vi.fn(() => [2] as const);
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    runTurnMock.mockResolvedValue({
      kind: "answer",
      message: "已完成。",
      resultContext: null,
      provider: "model"
    });

    const { container } = await mount(
      <WpsAgentPanel
        selectedSheetIndexes={[1]}
        resolveSelectedSheetIndexes={resolveSelectedSheetIndexes}
        scopeDock={<div>当前表 Sheet1</div>}
      />
    );
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "统计当前表");
    await submit(textarea);
    await until(() => runTurnMock.mock.calls.length === 1, "应发起查询");

    expect(resolveSelectedSheetIndexes).toHaveBeenCalledTimes(1);
    expect(runTurnMock.mock.calls[0][0].selectedSheetIndexes).toEqual([2]);
  });

  it("keeps earlier conversations when a new chat is created", async () => {
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    runTurnMock.mockResolvedValue({
      kind: "answer",
      message: "历史回答",
      resultContext: null,
      provider: "model"
    });

    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "历史问题");
    await submit(textarea);
    await until(() => textOf(container).includes("历史回答"), "应完成首个对话");
    await click(container.querySelector<HTMLButtonElement>('[aria-label="新对话"]'));

    expect(textOf(container)).toContain("已新建对话");
    expect(textOf(container)).not.toContain("新对话已开始");
    await until(() => document.activeElement === textarea, "新对话后应聚焦输入框");

    // 已经是空白对话时再次点击只给反馈，不制造重复空历史。
    await click(container.querySelector<HTMLButtonElement>('[aria-label="新对话"]'));
    expect(textOf(container)).toContain("当前已是新对话");
    await click(container.querySelector<HTMLButtonElement>('[aria-label="更多"]'));
    await click(container.querySelector<HTMLButtonElement>('[aria-label="历史对话"]'));

    expect(textOf(container)).toContain("历史问题");
    expect(textOf(container)).toContain("新对话");
    expect(container.querySelectorAll(".history-item")).toHaveLength(2);

    await act(async () => {
      await new Promise<void>((resolve) => setTimeout(resolve, 1550));
    });
    expect(container.querySelector('[role="status"]')).toBeNull();
  });

  it("shows a normalized write plan without touching cells until explicit confirmation", async () => {
    let value = "旧值";
    let formula = "旧值";
    let writes = 0;
    const range = { Rows: { Count: 1 }, Columns: { Count: 1 }, MergeCells: false } as Record<string, unknown>;
    Object.defineProperties(range, {
      Value2: {
        get: () => value,
        set: (next: string) => { value = next; formula = next; writes += 1; },
        enumerable: true
      },
      Formula: {
        get: () => formula,
        set: (next: string) => { formula = next; value = next; },
        enumerable: true
      }
    });
    const sheet = { Name: "Sheet1", ProtectContents: false, Range: () => range };
    const storage = new Map<string, unknown>();
    getWpsApplicationMock.mockReturnValue({
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
      PluginStorage: {
        getItem: (key: string) => storage.get(key),
        setItem: (key: string, item: unknown) => { storage.set(key, item); }
      },
      CreateTaskPane: () => ({ ID: 1, Visible: true }),
      GetTaskPane: () => undefined
    } as never);
    const writePlan = {
      id: "write-1",
      title: "更新 A1",
      summary: "把 A1 更新为新值",
      assumptions: [],
      warnings: [],
      actions: [{ type: "writeValues" as const, sheet: "Sheet1", range: "A1", values: [["新值"]] }],
      acceptanceCriteria: []
    };
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({ kind: "plan", plan: writePlan });
    validatePlanMock.mockImplementation(async (submitted) => submitted);

    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "把 A1 改成新值");
    await submit(textarea);
    await until(() => textOf(container).includes("写入前确认"), "应展示写入计划");

    expect(writes).toBe(0);
    expect(container.querySelector(".wps-plan-card")?.closest(".message-row.assistant")).not.toBeNull();
    const savedHistory = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    const savedPreview = savedHistory.conversations[0].messages.find((message: { planPreview?: unknown }) => message.planPreview);
    expect(savedPreview.planPreview).toEqual({ title: "更新 A1", status: "pending" });
    expect(savedPreview.text).toBe("");
    expect(textOf(container)).toContain("当前 WPS JSAPI 版本仅在开发调试模式开放写入");
    expect(textOf(container)).toContain("Sheet1!A1");
    expect(textOf(container)).toContain("新值");
    await click(container.querySelector<HTMLButtonElement>('[aria-label="新对话"]'));
    expect(textOf(container)).toContain("放弃待确认的写入计划？");
    expect(textOf(container)).toContain("当前对话还有一份尚未执行的写入计划");
    await click(findButton(container, "继续当前对话"));
    expect(textOf(container)).toContain("写入前确认");
    await click(findButton(container, "编辑计划"));
    expect(findButton(container, "确认写入")).toBeNull();
    const editBox = container.querySelector<HTMLTextAreaElement>('[aria-label="编辑写入计划 JSON"]');
    await type(editBox, JSON.stringify({
      ...writePlan,
      actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [["修订值"]] }]
    }));
    await click(findButton(container, "应用修改并重新预览"));
    await until(() => textOf(container).includes("修订值") && findButton(container, "确认写入") !== null, "编辑后应重新预览最终值");
    expect(writes).toBe(0);
    await click(findButton(container, "确认写入"));
    await until(() => textOf(container).includes("已保存恢复点"), "确认后应完成写入并提供撤销");
    expect(writes).toBe(1);
    expect(value).toBe("修订值");
    expect(textOf(container)).toContain("计划已执行并回读确认");
    expect(container.querySelector(".wps-verification-card.verified")?.textContent).toContain("1 个动作 · 1 个单元格已回读核对");
    expect(container.querySelector(".wps-verification-card.verified")?.textContent).toContain("Sheet1!A1");
    expect(textOf(container)).toContain("保存写前恢复点");
    expect(textOf(container)).toContain("回读核对第 1 个动作");
    const completedTimeline = Array.from(container.querySelectorAll(".task-progress-timeline.activity-log-card"))
      .find((item) => item.textContent?.includes("执行第 1 个写入动作"));
    expect(completedTimeline).not.toBeNull();
    expect(completedTimeline?.textContent).toContain("Sheet1!A1");
    expect(completedTimeline?.textContent).toContain("writeValues");
  });

  it("本端算出的计划不送服务端校验", async () => {
    // 验收条件里带着完整值矩阵（RangeEqualsCriterion.expected），送校验等于把
    // 刚才刻意留在本地的源值上传 —— 单射映射可逆。这条用例把"不上传"钉住。
    let value = "旧值";
    const range = {
      Rows: { Count: 1 },
      Columns: { Count: 1 },
      MergeCells: false,
      Formula: "旧值"
    } as Record<string, unknown>;
    Object.defineProperty(range, "Value2", {
      get: () => value,
      set: (next: string) => { value = next; },
      enumerable: true
    });
    const sheet = { Name: "Sheet1", ProtectContents: false, Range: () => range };
    const storage = new Map<string, unknown>();
    getWpsApplicationMock.mockReturnValue({
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
      PluginStorage: {
        getItem: (key: string) => storage.get(key),
        setItem: (key: string, item: unknown) => { storage.set(key, item); }
      },
      CreateTaskPane: () => ({ ID: 1, Visible: true }),
      GetTaskPane: () => undefined
    } as never);
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    runTurnMock.mockResolvedValue({
      kind: "plan",
      computedLocally: true,
      sourceSnapshots: [{ sheetName: "Sheet1", range: "A1", values: [["旧值"]] }],
      plan: {
        id: "wps-map-column-abc",
        title: "按映射表填写 Sheet1!A1",
        summary: "按 1 条映射规则写入。",
        assumptions: [],
        warnings: [],
        actions: [
          { type: "writeValues" as const, sheet: "Sheet1", range: "A1", values: [["行政"]] }
        ]
      }
    });

    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "把类别列翻译成归属");
    await submit(textarea);
    await until(() => textOf(container).includes("写入前确认"), "应展示写入计划");

    expect(validatePlanMock).not.toHaveBeenCalled();
    expect(textOf(container)).toContain("Sheet1!A1");
    await click(findButton(container, "编辑计划"));
    await click(findButton(container, "应用修改并重新预览"));
    await until(() => findButton(container, "确认写入") !== null, "本地计划编辑后应重新预览");
    expect(validatePlanMock).not.toHaveBeenCalled();
  });

  it("keeps a secret leaked by an upstream error out of the persisted chat history", async () => {
    // 这条断言必须**先注入**再检查：如果不往流程里塞一个密钥，
    // `not.toContain("sk-")` 会因为压根没有密钥而空过 —— 那是假绿灯。
    // 通道是真实的：`turnErrorMessage` 会把非连接类错误的 message 原样写进
    // 对话历史，而服务端错误 message 源自 `str(error)`，可能带上请求里的凭据。
    const secret = "sk-live-leaked-from-upstream-9f3a";
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockRejectedValue(
      new Error(`模型连接校验失败：上游拒绝 Authorization: Bearer ${secret}`)
    );
    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计订单金额");
    await submit(textarea);
    await until(
      () => container.querySelector(".message-row.system") !== null,
      "非连接类错误应产生一条系统消息"
    );

    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    // 断言的是**密钥本体**不落盘。`Bearer` 这个词本身不是凭据 ——
    // 修复后的文本是「Authorization: Bearer 【已隐藏】」，方案名保留反而更可诊断，
    // 所以不能把「出现 Bearer 字面」当成泄漏。
    expect(persisted).not.toContain(secret);
    expect(persisted).not.toContain("sk-live");
    expect(persisted).not.toContain("sk-");
    expect(textOf(container)).not.toContain(secret);
    // 正向锚点：这条错误确实写进了历史（否则上面的负向断言可能空过），
    // 且密钥位置被换成了明确占位符 —— 是「被隐藏」而不是「整段消失」。
    expect(persisted).toContain("模型连接校验失败");
    expect(persisted).toContain("Bearer 【已隐藏】");
    // 用户仍应看到失败这件事与可诊断的部分，只是不带凭据。
    expect(textOf(container)).toContain("模型连接校验失败");
    expect(textOf(container)).toContain("已隐藏");
  });

  it("sanitizes an old record before it can be rendered, and blocks resuming it", async () => {
    // 读取路径的职责（本轮修正后）：原始存储 → **立刻脱敏** → 进入可渲染状态，
    // 同时登记一个持久的旧记录标记用于发送闸。旧密钥**不得**出现在界面、
    // 复制内容或内存里的可发送历史中。
    const secret = "sk-live-legacy-stored-2d77";
    localStorage.setItem(
      "excel-bro.wps.chat.v1",
      JSON.stringify({
        activeConversationId: "legacy",
        conversations: [
          {
            id: "legacy",
            title: `用 ${secret} 统计订单`,
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: `用 ${secret} 统计订单`,
                createdAt: "2026-09-27T00:00:00.000Z"
              }
            ]
          }
        ]
      })
    );
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });

    const { container } = await mount(panel());
    await until(
      () => container.querySelector(".stale-credentials-warning") !== null,
      "旧记录应给出「需新开会话」提示"
    );

    // 正向锚点：旧记录确实被读回来了、非敏感内容完整保留。
    expect(textOf(container)).toContain("统计订单");
    // 关键负向断言：界面里不得出现真密钥。
    expect(textOf(container)).not.toContain(secret);
    expect(textOf(container)).not.toContain("sk-live");

    // 落盘内容同样不含密钥，且带上了持久标记。
    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    expect(persisted).not.toContain(secret);
    expect(persisted).toContain("legacyCredentialConversationIds");
  });

  it("renders a restored placeholder and keeps an ordinary question untouched", async () => {
    // 已经落盘脱敏过的记录：读取时应把占位符还原成可读措辞；
    // 若它带着旧记录标记，则会继续提示，但文本里不含任何真凭据。
    localStorage.setItem(
      "excel-bro.wps.chat.v1",
      JSON.stringify({
        activeConversationId: "saved",
        conversations: [
          {
            id: "saved",
            title: "用 【已隐藏】 统计订单",
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: "用 【已隐藏】 统计订单",
                createdAt: "2026-09-27T00:00:00.000Z"
              },
              {
                id: "m2",
                role: "user",
                text: "Bearer 和 api key 有什么区别？",
                createdAt: "2026-09-27T00:00:01.000Z"
              }
            ]
          }
        ]
      })
    );
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });

    const { container } = await mount(panel());
    await until(
      () => container.querySelector(".message-row.user") !== null,
      "旧历史的用户消息应被读回并渲染"
    );

    const stream = textOf(container.querySelector(".message-stream") as HTMLElement);
    // 落盘标记被还原成更明确的措辞。
    expect(stream).toContain("此处原为凭据");
    expect(stream).not.toContain("【已隐藏】");
    // 用户正常提问一字不改 —— 读取时不做新的脱敏。
    expect(stream).toContain("Bearer 和 api key 有什么区别？");
    // 没有被标为旧记录的干净历史不触发提示（否则每次重开都会误锁续聊）。
    expect(container.querySelector(".stale-credentials-warning")).toBeNull();
  });

  it("blocks resuming a chat whose stored history still carries a real credential", async () => {
    // 核心行为：拦在**发送之前**，且拦的依据是持久标记而不是"还能不能匹配密钥"。
    const secret = "sk-live-resume-blocked-6c02";
    localStorage.setItem(
      "excel-bro.wps.chat.v1",
      JSON.stringify({
        activeConversationId: "legacy",
        conversations: [
          {
            id: "legacy",
            title: `用 ${secret} 统计订单`,
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: `用 ${secret} 统计订单`,
                createdAt: "2026-09-27T00:00:00.000Z"
              }
            ]
          }
        ]
      })
    );
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });

    const { container } = await mount(panel());
    await until(
      () => container.querySelector(".composer textarea") !== null,
      "面板应就绪"
    );

    await type(container.querySelector<HTMLTextAreaElement>(".composer textarea"), "继续统计");
    await submit(container.querySelector<HTMLTextAreaElement>(".composer textarea"));

    // 正向锚点：确实给出了反馈，而不是静默什么都不发生。
    await until(
      () => textOf(container).includes("已被拦下"),
      "应提示这段历史的凭据被拦下"
    );
    // 关键负向断言：旧记录没有被送给模型。
    expect(runTurnMock).not.toHaveBeenCalled();
    // 界面里也看不到真密钥。
    expect(textOf(container)).not.toContain(secret);
  });

  it("warns without leaking when the host storage refuses the history write", async () => {
    // 宿主禁止 localStorage（或配额用尽）时，落盘 effect 过去是**静默忽略**的：
    // 用户会以为"历史保存了"，关掉任务窗格才发现全没了。这条用例钉三件事：
    // 1. 界面给出**明确**提示；
    // 2. 提示**不含**错误对象原文，也不含任何凭据；
    // 3. 内存历史仍保持脱敏与陈旧标记 —— 落盘失败**不会**让发送闸松开。
    const secret = "sk-live-storage-refused-77c1";
    localStorage.setItem(
      "excel-bro.wps.chat.v1",
      JSON.stringify({
        activeConversationId: "legacy",
        conversations: [
          {
            id: "legacy",
            title: `用 ${secret} 统计订单`,
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: `用 ${secret} 统计订单`,
                createdAt: "2026-09-27T00:00:00.000Z"
              }
            ]
          }
        ]
      })
    );
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });

    // 只让**对话历史**那一笔写入失败；读取（getItem）与其他键不受影响。
    // 抛出的错误里刻意塞进凭据，用来证明提示不会把异常原文透出去。
    const originalSetItem = Storage.prototype.setItem;
    const spy = vi
      .spyOn(Storage.prototype, "setItem")
      .mockImplementation(function (this: Storage, key: string, value: string) {
        if (key === "excel-bro.wps.chat.v1") {
          throw new Error(`QuotaExceededError while writing ${secret}`);
        }
        originalSetItem.call(this, key, value);
      });

    try {
      const { container } = await mount(panel());
      await until(
        () => container.querySelector(".storage-write-warning") !== null,
        "落盘失败应给出明确提示"
      );

      const notice = container.querySelector(".storage-write-warning") as HTMLElement;
      // 正向锚点：提示确实说明了"不会被保留"这件事。
      expect(notice.textContent ?? "").toContain("不会被保留");
      // 关键负向断言 1：提示里不得出现错误对象原文，更不得出现凭据。
      expect(notice.textContent ?? "").not.toContain(secret);
      expect(notice.textContent ?? "").not.toContain("sk-live");
      expect(notice.textContent ?? "").not.toContain("QuotaExceededError");
      // 关键负向断言 2：整个界面都不能因为"写入失败"而把原文露出来。
      expect(textOf(container)).not.toContain(secret);
      expect(textOf(container)).not.toContain("sk-live");

      // 关键负向断言 3：落盘失败**不解除发送闸** —— 旧记录仍被拦在发送之前。
      await type(container.querySelector<HTMLTextAreaElement>(".composer textarea"), "继续统计");
      await submit(container.querySelector<HTMLTextAreaElement>(".composer textarea"));
      await until(
        () => textOf(container).includes("已被拦下"),
        "即使落盘失败，旧记录仍应被拦"
      );
      expect(runTurnMock).not.toHaveBeenCalled();
      // 陈旧提示也还在。
      expect(container.querySelector(".stale-credentials-warning")).not.toBeNull();
    } finally {
      spy.mockRestore();
    }
  });

  it("keeps the legacy flag across a reopen so the send gate stays closed", async () => {
    // 重开之后：文本已被脱敏，靠"匹配密钥"已无法判断，必须靠持久标记。
    const secret = "sk-live-reopen-wps-90af";
    localStorage.setItem(
      "excel-bro.wps.chat.v1",
      JSON.stringify({
        activeConversationId: "legacy",
        conversations: [
          {
            id: "legacy",
            title: `用 ${secret} 统计订单`,
            createdAt: "2026-09-27T00:00:00.000Z",
            updatedAt: "2026-09-27T00:00:00.000Z",
            messages: [
              {
                id: "m1",
                role: "user",
                text: `用 ${secret} 统计订单`,
                createdAt: "2026-09-27T00:00:00.000Z"
              }
            ]
          }
        ]
      })
    );
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });

    // 第一次挂载：脱敏 + 落盘（写入标记）。
    const first = await mount(panel());
    await until(
      () => first.container.querySelector(".stale-credentials-warning") !== null,
      "首次挂载应提示"
    );
    await settle();
    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    expect(persisted).not.toContain(secret);
    first.unmount();

    // 第二次挂载（等价于重开任务窗格）：文本已干净，但标记仍在。
    const second = await mount(panel());
    await until(
      () => second.container.querySelector(".composer textarea") !== null,
      "面板应就绪"
    );
    await until(
      () => second.container.querySelector(".stale-credentials-warning") !== null,
      "重开后仍应保持提示"
    );

    await type(
      second.container.querySelector<HTMLTextAreaElement>(".composer textarea"),
      "继续统计"
    );
    await submit(second.container.querySelector<HTMLTextAreaElement>(".composer textarea"));
    await until(
      () => textOf(second.container).includes("已被拦下"),
      "重开后发送仍应被拦"
    );
    expect(runTurnMock).not.toHaveBeenCalled();
  });

  it("keeps credentials and raw cell values out of the persisted chat history", async () => {
    // 持久历史进的是宿主 localStorage，用户可读可复制。
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "answer", message: "统计完成", provider: "model", resultContext: null
    });
    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计订单金额");
    await submit(textarea);
    await until(() => textOf(container).includes("统计完成"), "应完成一轮");

    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    // 正向锚点：历史确实写进去了（否则下面的负向断言又会空过）。
    expect(persisted).toContain("统计订单金额");
    expect(persisted).not.toContain("sk-");
    expect(persisted).not.toContain("apiKey");
    expect(persisted).not.toContain("Authorization");
    expect(persisted).not.toContain("bearer");
  });

  it("keeps a secret pasted into the prompt out of the persisted conversation title", async () => {
    // 标题 = 首条用户消息前 34 字。用户把密钥连同问题一起粘进来时，
    // 只脱敏 message.text 会让标题成为绕过通道。
    const secret = "sk-live-in-title-c41d";
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "answer", message: "已统计。", provider: "model", resultContext: null
    });
    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, `用 ${secret} 统计订单金额`);
    await submit(textarea);
    await until(() => textOf(container).includes("已统计。"), "应完成一轮");

    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    const parsed = JSON.parse(persisted) as {
      conversations: Array<{ title: string }>;
    };
    // 正向锚点：标题确实生成了（否则负向断言可能空过），且仍是可读句子。
    expect(parsed.conversations[0].title).toContain("统计订单金额");
    expect(parsed.conversations[0].title).not.toContain(secret);
    expect(parsed.conversations[0].title).not.toContain("sk-");
    expect(parsed.conversations[0].title).toContain("已隐藏");
    // 正文同样不能漏。持久化原文（含转义）整体也不该出现密钥。
    expect(persisted).not.toContain(secret);
  });

  it("keeps a JSON-shaped secret out of every persisted text field", async () => {
    // 用户可能整段粘贴配置片段。JSON / 类 JSON 形态走的是另一条正则分支，
    // 必须单独注入验证，不能因为「裸 apiKey= 已覆盖」就假定这里也覆盖。
    const secret = "testsecret123456";
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "answer", message: `收到配置 {"apiKey":"${secret}"}，已统计。`,
      provider: "model", resultContext: null
    });
    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, `按 {'api_key': '${secret}'} 统计订单金额`);
    await submit(textarea);
    await until(() => textOf(container).includes("已统计。"), "应完成一轮");

    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    const parsed = JSON.parse(persisted) as {
      conversations: Array<{ title: string; messages: Array<{ text: string }> }>;
    };
    const title = parsed.conversations[0].title;
    const answerText = parsed.conversations[0].messages
      .map((message) => message.text)
      .join("\n");

    expect(persisted).not.toContain(secret);
    // 键名与结构保留，说明是「脱敏」而不是「整条消息被丢弃」。
    expect(title).toContain("api_key");
    expect(answerText).toContain("apiKey");
    expect(title).toContain("已隐藏");
    expect(answerText).toContain("已隐藏");
  });

  it("does not persist the locally computed target matrix with the plan preview", async () => {
    // 本地映射计划的目标值矩阵与源值一一对应（单射可逆），既不上传也不落盘。
    // 源区域在预览时会被对着活宿主复核（prepareWpsPlan），所以替身单元格必须
    // 与快照值一致 —— 否则拒绝的原因是"源值变了"，测不到持久化这件事。
    let value = "办公用品";
    const range = {
      Rows: { Count: 1 }, Columns: { Count: 1 }, MergeCells: false, Formula: "办公用品"
    } as Record<string, unknown>;
    Object.defineProperty(range, "Value2", {
      get: () => value,
      set: (next: string) => { value = next; },
      enumerable: true
    });
    const sheet = { Name: "Sheet1", ProtectContents: false, Range: () => range };
    const storage = new Map<string, unknown>();
    getWpsApplicationMock.mockReturnValue({
      Version: "12.0",
      ActiveWorkbook: {
        Name: "测试.xlsx", FullName: "C:\\Data\\测试.xlsx", Path: "C:\\Data",
        ReadOnly: false, Worksheets: { Count: 1, Item: () => sheet }, ActiveSheet: sheet
      },
      ActiveSheet: sheet,
      PluginStorage: {
        getItem: (key: string) => storage.get(key),
        setItem: (key: string, item: unknown) => { storage.set(key, item); }
      },
      CreateTaskPane: () => ({ ID: 1, Visible: true }),
      GetTaskPane: () => undefined
    } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({
      kind: "plan",
      computedLocally: true,
      sourceSnapshots: [{ sheetName: "Sheet1", range: "A1", values: [["办公用品"]] }],
      plan: {
        id: "wps-map-column-xyz",
        title: "按映射表填写 Sheet1!A1",
        summary: "按 1 条映射规则写入。",
        assumptions: [], warnings: [],
        actions: [{ type: "writeValues" as const, sheet: "Sheet1", range: "A1", values: [["行政"]] }]
      }
    });

    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "把类别列翻译成归属");
    await submit(textarea);
    await until(() => textOf(container).includes("写入前确认"), "应展示写入计划");

    // 目标矩阵「行政」与源值「办公用品」都不能出现在持久历史里。
    const persisted = localStorage.getItem("excel-bro.wps.chat.v1") ?? "";
    expect(persisted).not.toContain("办公用品");
    expect(persisted).not.toContain("行政");
    expect(persisted).toContain("按映射表填写 Sheet1!A1");
    expect(validatePlanMock).not.toHaveBeenCalled();
  });

  it("keeps an initial catalog failure out of chat and offers contextual retry", async () => {
    listModelsMock.mockRejectedValueOnce(new Error("Failed to fetch"));
    const timeoutSpy = vi.spyOn(window, "setTimeout");

    const { container } = await mount(panel());
    await until(
      () => textOf(container).includes("点击添加模型开始"),
      "首次连接不可用时应保留中性的入口状态"
    );

    // 尚未配置与暂时无法取得目录是两种不同状态，首次打开不能抛出失败卡。
    expect(textOf(container)).not.toContain("尚未配置模型");
    expect(textOf(container)).not.toContain("Failed to fetch");
    expect(textOf(container)).toContain("点击上方「添加模型」开始连接");
    expect(container.querySelector(".message-row.system")).toBeNull();
    expect(timeoutSpy.mock.calls.some(([, delay]) => delay === 5000)).toBe(false);
    timeoutSpy.mockRestore();

    await act(async () => {
      window.dispatchEvent(new Event("focus"));
    });
    expect(listModelsMock).toHaveBeenCalledTimes(1);

    const textarea = container.querySelector("textarea");
    await type(textarea, "你好");
    await submit(textarea);
    await submit(textarea);
    expect(textarea?.value).toBe("你好");
    expect(container.querySelector(".message-row.system")).toBeNull();
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(textOf(container)).toContain("暂时无法取得已保存的模型连接");

    listModelsMock.mockResolvedValueOnce({
      defaultModelId: secondConnection.id,
      models: [secondConnection]
    });
    await click(findButton(container, "重试读取模型配置"));

    await until(
      () => textOf(container).includes("模型：第二个连接"),
      "重试成功后应显示真实连接"
    );
    expect(listModelsMock).toHaveBeenCalledTimes(2);
  });

  it("removes only the obsolete model-catalog notices from saved chat history", async () => {
    const obsolete = "暂时无法读取模型配置，请点击上方模型菜单里的「重试」。已输入的内容仍保留在输入框中。";
    localStorage.setItem("excel-bro.wps.chat.v1", JSON.stringify({
      activeConversationId: "saved-1",
      conversations: [{
        id: "saved-1", title: "你好", createdAt: "2026-09-26T00:00:00.000Z",
        updatedAt: "2026-09-26T00:00:00.000Z",
        messages: [
          { id: "user-1", role: "user", text: "你好", createdAt: "2026-09-26T00:00:00.000Z" },
          { id: "old-1", role: "system", text: obsolete, createdAt: "2026-09-26T00:00:01.000Z" }
        ]
      }]
    }));
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });

    const { container } = await mount(panel());
    await until(() => textOf(container).includes("尚未配置模型"), "应完成目录读取");
    expect(textOf(container)).toContain("你好");
    expect(textOf(container)).not.toContain(obsolete);
    expect(localStorage.getItem("excel-bro.wps.chat.v1")).not.toContain(obsolete);
  });

  it("opens the WPS add-model entry immediately while checking saved connections", async () => {
    listModelsMock.mockRejectedValueOnce(new Error("Failed to fetch"));
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });
    getModelSettingsMock.mockResolvedValue({
      baseUrl: null,
      defaultModel: null,
      apiKeyConfigured: false,
      apiKeyHint: null,
      formulaModelId: "",
      connections: []
    });

    const { container } = await mount(panel());
    await until(() => textOf(container).includes("点击添加模型开始"), "应完成首次模型目录读取");
    await click(container.querySelector(".model-picker-trigger"));
    await click(findButton(container, "＋ 添加模型连接"));

    expect(textOf(container)).toContain("添加模型连接");
    expect(textOf(container)).not.toContain("正在读取模型连接");
    expect(textOf(container)).not.toContain("后台暂时不可用");
    await until(() => getModelSettingsMock.mock.calls.length === 1, "新增表单应异步检查已有连接");
    expect(textOf(container)).toContain("已检查：这台电脑上暂无已保存的模型连接");
    expect(listModelsMock).toHaveBeenCalledTimes(2);

    await click(container.querySelector('.wps-model-settings-header [aria-label="关闭模型设置"]'));
    await click(container.querySelector(".model-picker-trigger"));
    await click(findButton(container, "管理模型连接"));
    await until(() => getModelSettingsMock.mock.calls.length === 2, "管理入口也应读取已有连接");
  });

  it("keeps the add form usable when saved-connection discovery fails", async () => {
    getModelSettingsMock.mockRejectedValue(new Error("Failed to fetch"));

    function Harness() {
      const management = useModelManagement({
        refreshServiceHealth: async () => null
      });
      const [open, setOpen] = useState(true);
      return (
        <WpsModelSettingsPanel
          open={open}
          models={[]}
          currentModelId=""
          management={management}
          onClose={() => setOpen(false)}
          onAddConnection={() => {
            management.openConnectionCreator();
            void management.checkExistingConnections();
          }}
        />
      );
    }

    const { container } = await mount(<Harness />);

    const addButton = findButton(container, "＋ 添加");
    expect(addButton?.hasAttribute("disabled")).toBe(false);
    await click(addButton);

    // 新建表单不等待后台；读取失败不是“没有已保存连接”。
    await until(
      () => textOf(container).includes("添加模型连接"),
      "后台不可读时也应能打开新增表单"
    );
    expect(
      findButton(container, "添加并使用")?.hasAttribute("disabled")
    ).toBe(true);
    await until(() => textOf(container).includes("暂时无法检查已有连接"), "应区分读取失败和空列表");
    expect(textOf(container)).not.toContain("已检查：这台电脑上暂无已保存的模型连接");
    expect(findButton(container, "重试检查")).not.toBeNull();
    expect(getModelSettingsMock).toHaveBeenCalledTimes(1);
  });

  it("renders the complete local model connection manager", () => {
    const noop = () => undefined;
    const management = {
      modelSettings: {
        baseUrl: null,
        defaultModel: null,
        apiKeyConfigured: false,
        apiKeyHint: null,
        formulaModelId: "",
        connections: []
      },
      apiKeyDraft: "",
      showApiKey: false,
      connectionDraft: {
        id: null,
        label: "",
        baseUrl: "",
        modelId: "",
        apiKey: "",
        clearApiKey: false,
        supportsVision: false
      },
      pendingDeleteConnectionId: null,
      settingsSaving: false,
      settingsTesting: false,
      settingsLoading: false,
      settingsFeedback: "",
      connectionTestFeedback: null,
      connectionVerified: false,
      saveApiKey: noop,
      editModelConnection: noop,
      verifyConnection: noop,
      saveConnection: noop,
      removeConnection: noop,
      saveFormulaModel: noop,
      setApiKeyDraft: noop,
      setShowApiKey: noop,
      setConnectionDraft: noop,
      setPendingDeleteConnectionId: noop,
      setSettingsFeedback: noop
    };

    const markup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        models={[]}
        currentModelId=""
        management={management as never}
        onClose={noop}
        onAddConnection={noop}
      />
    );

    expect(markup).toContain("模型设置");
    expect(markup).toContain("API Key 安全保存在这台电脑上");
    expect(markup).toContain("添加模型连接");
    expect(markup).toContain("测试连接");
    expect(markup).toContain("添加并使用");
    expect(markup).toContain("/function 公式模型");

    const existingConnectionMarkup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        models={[]}
        currentModelId=""
        management={{
          ...management,
          existingConnectionsStatus: "ready",
          connectionVerified: true,
          connectionDraft: {
            id: null,
            label: "再次添加",
            baseUrl: "https://api.example.com/",
            modelId: "GPT-4.1",
            apiKey: "",
            clearApiKey: false,
            supportsVision: false
          },
          modelSettings: {
            ...management.modelSettings,
            connections: [{
              id: "saved",
              catalogModelId: "connection:saved",
              label: "已有模型",
              baseUrl: "https://api.example.com",
              modelId: "gpt-4.1",
              supportsVision: false,
              apiKeyConfigured: true,
              apiKeyHint: "••••1234"
            }]
          },
          selectModel: noop
        } as never}
        onClose={noop}
        onAddConnection={noop}
      />
    );
    expect(existingConnectionMarkup).toContain("Excel 与 WPS 共用这些配置");
    expect(existingConnectionMarkup).not.toContain("无需重复添加");
    expect(existingConnectionMarkup.match(/<button class="model-primary-action"[^>]*>/)?.[0]).not.toContain("disabled");
    expect(existingConnectionMarkup.match(/<button type="button"[^>]*>测试连接<\/button>/)?.[0]).not.toContain("disabled");

    const unavailableMarkup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        models={[]}
        currentModelId=""
        management={{
          ...management,
          connectionDraft: null,
          modelSettings: {
            ...management.modelSettings,
            connections: [{
              id: "saved",
              catalogModelId: "connection:saved",
              label: "已有模型",
              baseUrl: "https://api.example.com",
              modelId: "gpt-4.1",
              supportsVision: false,
              apiKeyConfigured: true,
              apiKeyHint: "••••1234"
            }]
          }
        } as never}
        onClose={noop}
        onAddConnection={noop}
      />
    );
    expect(unavailableMarkup).toContain("部分连接暂不可切换");
    expect(unavailableMarkup).toContain("重试读取");
    expect(unavailableMarkup.match(/<button class="model-connection-select"[^>]*>/)?.[0]).toContain("disabled");
  });

  it("switches models in place and expands editing below the clicked card", async () => {
    const connections = [
      {
        id: "first",
        catalogModelId: firstConnection.id,
        label: "DS",
        baseUrl: "https://api.deepseek.com",
        modelId: "deepseek-chat",
        supportsVision: false,
        apiKeyConfigured: true,
        apiKeyHint: "••••1234"
      },
      {
        id: "second",
        catalogModelId: secondConnection.id,
        label: "vec",
        baseUrl: "https://api.example.com/v1",
        modelId: "claude-opus-4-6",
        supportsVision: false,
        apiKeyConfigured: true,
        apiKeyHint: "••••5678"
      }
    ];
    const onClose = vi.fn();
    function Harness() {
      const management = useModelManagement({ refreshServiceHealth: async () => null });
      useEffect(() => {
        management.setModelSettings({
          baseUrl: null,
          defaultModel: null,
          apiKeyConfigured: false,
          apiKeyHint: null,
          formulaModelId: "",
          connections
        });
      }, []);
      return (
        <WpsModelSettingsPanel
          open
          models={[firstConnection, secondConnection]}
          currentModelId={management.selectedModelId || firstConnection.id}
          management={management}
          onClose={onClose}
          onAddConnection={management.openConnectionCreator}
        />
      );
    }

    const { container } = await mount(<Harness />);
    await until(() => container.querySelectorAll(".model-connection-select").length === 2, "应列出两个模型");
    const cards = container.querySelectorAll<HTMLButtonElement>(".model-connection-select");
    expect(cards[0].disabled).toBe(true);
    expect(cards[1].disabled).toBe(false);
    expect(cards[1].textContent).toContain("点击切换使用");

    const editButtons = container.querySelectorAll<HTMLButtonElement>(".model-connection-actions button:first-child");
    await click(editButtons[1]);
    expect(textOf(container)).toContain("编辑 vec");
    expect(container.querySelector(".model-connection-list")).not.toBeNull();
    expect(container.querySelectorAll(".model-connection-card")[0].querySelector(".model-connection-editor")).toBeNull();
    expect(container.querySelectorAll(".model-connection-card")[1].querySelector(".model-connection-editor")).not.toBeNull();
    expect(textOf(container)).toContain("/function 公式模型");
    expect(textOf(container)).toContain("已保存 ••••5678");
    expect(textOf(container)).toContain("现有密钥不回显");
    const keyInput = container.querySelector<HTMLInputElement>('input[placeholder="输入新 Key 以替换；留空保持不变"]');
    expect(keyInput?.value).toBe("");
    expect(findButton(container, "显示新 Key")).toBeNull();
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
      setter?.call(keyInput, "replacement-test-key");
      keyInput?.dispatchEvent(new Event("input", { bubbles: true }));
    });
    expect(findButton(container, "显示新 Key")).not.toBeNull();
    await click(findButton(container, "显示新 Key"));
    expect(container.querySelector<HTMLInputElement>('input[value="replacement-test-key"]')?.type).toBe("text");

    await click(findButton(container, "取消"));
    expect(container.querySelector(".model-connection-editor")).toBeNull();
    expect(container.querySelectorAll(".model-connection-select")).toHaveLength(2);
    await click(container.querySelectorAll(".model-connection-select")[1]);
    expect(localStorage.getItem(MODEL_STORAGE_KEY)).toBe(secondConnection.id);
    expect(container.querySelectorAll(".model-connection-card")[1].classList.contains("current")).toBe(true);
    expect(container.querySelector(".wps-model-settings")).not.toBeNull();
    expect(container.querySelector(".wps-model-settings.tool-drawer.settings-drawer")).not.toBeNull();
    expect(container.querySelector(".model-settings-backdrop")).toBeNull();
    expect(onClose).not.toHaveBeenCalled();
  });

  it("keeps the connection form usable after an explicit management read failure", () => {
    const noop = () => undefined;
    const management = {
      modelSettings: null,
      apiKeyDraft: "",
      showApiKey: false,
      connectionDraft: null,
      pendingDeleteConnectionId: null,
      settingsSaving: false,
      settingsTesting: false,
      settingsLoading: false,
      settingsFeedback: "暂时无法连接 Excel Bro 后台：Failed to fetch。",
      connectionTestFeedback: null,
      connectionVerified: false,
      saveApiKey: noop,
      editModelConnection: noop,
      verifyConnection: noop,
      saveConnection: noop,
      removeConnection: noop,
      saveFormulaModel: noop,
      setApiKeyDraft: noop,
      setShowApiKey: noop,
      setConnectionDraft: noop,
      setPendingDeleteConnectionId: noop,
      setSettingsFeedback: noop
    };

    const markup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        models={[]}
        currentModelId=""
        management={management as never}
        onClose={noop}
        onAddConnection={noop}
      />
    );
    const addButton =
      markup.match(/<button type="button"([^>]*)>＋ 添加<\/button>/)?.[1] ?? "";

    expect(addButton).not.toContain("disabled");
    expect(markup).toContain("暂时无法连接 Excel Bro 后台");

    const editorMarkup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        models={[]}
        currentModelId=""
        management={
          {
            ...management,
            connectionDraft: {
              id: null,
              label: "DeepSeek",
              baseUrl: "https://api.deepseek.com",
              modelId: "deepseek-chat",
              apiKey: "test-key",
              clearApiKey: false,
              supportsVision: false
            }
          } as never
        }
        onClose={noop}
        onAddConnection={noop}
      />
    );
    const testButton =
      editorMarkup.match(
        /<button type="button"([^>]*)>测试连接<\/button>/
      )?.[1] ?? "";
    const saveButton =
      editorMarkup.match(
        /<button class="model-primary-action" type="button"([^>]*)>添加并使用<\/button>/
      )?.[1] ?? "";

    expect(testButton).not.toContain("disabled");
    expect(saveButton).toContain("disabled");
  });

  it("keeps the model menu above the conversation cards", () => {
    expect(excelStyles).toMatch(/\.chat-header\s*\{[\s\S]*?z-index:\s*4;/);
    expect(excelStyles).toMatch(
      /\.chat-header\.menu-open\s*\{[\s\S]*?z-index:\s*50;/
    );
    expect(excelStyles).toMatch(/\.model-menu\s*\{[\s\S]*?z-index:\s*45;/);
  });

  it("positions WPS model settings as an in-panel drawer below the header", () => {
    expect(styles).toMatch(/\.wps-model-settings\s*\{[\s\S]*?position:\s*absolute;[\s\S]*?top:\s*53px;/);
  });

  it("keeps the service status readable before collapsing the narrow header", () => {
    expect(styles).toMatch(
      /@media \(max-width: 620px\)[\s\S]*?\.chat-header \.brand-copy span\s*\{[\s\S]*?max-width:\s*none;[\s\S]*?overflow:\s*visible;/
    );
    expect(styles).toMatch(
      /@media \(max-width: 620px\)[\s\S]*?\.chat-header \.new-chat-entry > span:not\(\[aria-hidden="true"\]\)\s*\{[\s\S]*?display:\s*none;/
    );
  });
});
