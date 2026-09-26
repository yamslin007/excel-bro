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
    const custom = container.querySelector<HTMLInputElement>('[aria-label="其他想法"]');
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
    await act(async () => {
      setter?.call(custom, "只统计记录数，不修改表格");
      custom?.dispatchEvent(new Event("input", { bubbles: true }));
    });
    await click(findButton(container, "发送"));
    await until(() => runTurnMock.mock.calls.length === 3, "补充说明应继续原任务");
    expect(runTurnMock.mock.calls[2][0].prompt).toContain("补充确认：只统计记录数，不修改表格");
    expect(runTurnMock.mock.calls[2][0].intentConfirmed).toBe(true);
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
      Version: "12.1.0.28488",
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
    validatePlanMock.mockResolvedValue(writePlan);

    const { container } = await mount(panel());
    await until(() => textOf(container).includes("模型：第一个连接"), "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "把 A1 改成新值");
    await submit(textarea);
    await until(() => textOf(container).includes("写入前确认"), "应展示写入计划");

    expect(writes).toBe(0);
    expect(textOf(container)).toContain("Sheet1!A1");
    await click(container.querySelector<HTMLButtonElement>('[aria-label="新对话"]'));
    expect(textOf(container)).toContain("放弃待确认的写入计划？");
    expect(textOf(container)).toContain("当前对话还有一份尚未执行的写入计划");
    await click(findButton(container, "继续当前对话"));
    expect(textOf(container)).toContain("写入前确认");
    await click(findButton(container, "确认写入"));
    await until(() => textOf(container).includes("已保存恢复点"), "确认后应完成写入并提供撤销");
    expect(writes).toBe(1);
    expect(value).toBe("新值");
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

  it("keeps the service status readable before collapsing the narrow header", () => {
    expect(styles).toMatch(
      /@media \(max-width: 620px\)[\s\S]*?\.chat-header \.brand-copy span\s*\{[\s\S]*?max-width:\s*none;[\s\S]*?overflow:\s*visible;/
    );
    expect(styles).toMatch(
      /@media \(max-width: 620px\)[\s\S]*?\.chat-header \.new-chat-entry > span:not\(\[aria-hidden="true"\]\)\s*\{[\s\S]*?display:\s*none;/
    );
  });
});
