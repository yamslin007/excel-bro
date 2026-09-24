import { readFileSync } from "node:fs";
import { act, useState, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MODEL_STORAGE_KEY } from "../../excel-addin/src/conversation";
import { useModelManagement } from "../../excel-addin/src/hooks/useModelManagement";
import { getModelSettings, listModels } from "../../excel-addin/src/api";
import WpsAgentPanel from "./WpsAgentPanel";
import WpsModelSettingsPanel from "./WpsModelSettingsPanel";
import { runWpsReadonlyTurn } from "./wpsReadonlyAgent";

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("../../excel-addin/src/api")>();
  return {
    ...actual,
    listModels: vi.fn(),
    getModelSettings: vi.fn()
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
const listModelsMock = vi.mocked(listModels);
const getModelSettingsMock = vi.mocked(getModelSettings);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);

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
      selectedSheetIndexes={[0]}
      scopeDock={<div className="eb-scope-dock">当前表 Sheet1</div>}
    />
  );
}

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
});

describe("WPS Excel Bro chat UI", () => {
  it("shows the Excel-style chat shell and keeps WPS read-only boundaries visible", () => {
    const markup = renderToStaticMarkup(panel());

    expect(markup).toContain("wps-chat-shell chat-shell");
    expect(markup).toContain("eb-model-trigger");
    const modelTrigger =
      markup.match(/<button class="eb-model-trigger"([^>]*)>/)?.[1] ?? "";
    expect(modelTrigger).not.toContain("disabled");
    // 目录尚未加载时不得出现任何“基础模式”一类占位模型。
    expect(markup).not.toContain("基础模式");
    expect(markup).toContain("正在读取模型配置");
    expect(markup).toContain("eb-composer-zone");
    expect(markup).toContain('class="app-logo"');
    expect(markup).toContain('aria-label="隐藏格仔"');
    expect(markup).toContain("pet-float mood-idle");
    expect(markup).toContain("当前表 Sheet1");
    expect(markup).toContain("WPS 当前仅支持只读查询");
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

    expect(textOf(container)).toContain("尚未配置模型。请点击上方「添加模型」");
    // 任务文本必须留在输入框里，不能因为缺少模型就被清掉。
    expect(textarea?.value).toBe("按部门统计平均分");
    expect(runTurnMock).not.toHaveBeenCalled();
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

  it("reports a failed catalog read as a failure with retry, not as zero models", async () => {
    listModelsMock.mockRejectedValueOnce(new Error("Failed to fetch"));

    const { container } = await mount(panel());
    await until(
      () => textOf(container).includes("无法读取模型配置"),
      "读取失败应显示失败状态"
    );

    // 读取失败不等于“用户没有配置模型”，不能叠加成未配置结论。
    expect(textOf(container)).not.toContain("尚未配置模型");

    listModelsMock.mockResolvedValueOnce({
      defaultModelId: secondConnection.id,
      models: [secondConnection]
    });
    await click(findButton(container, "重试"));

    await until(
      () => textOf(container).includes("模型：第二个连接"),
      "重试成功后应显示真实连接"
    );
    expect(listModelsMock).toHaveBeenCalledTimes(2);
  });

  it("opens the add-connection form even when the backend cannot be read", async () => {
    getModelSettingsMock.mockRejectedValue(new Error("Failed to fetch"));

    function Harness() {
      const management = useModelManagement({
        refreshServiceHealth: async () => null
      });
      const [open, setOpen] = useState(true);
      return (
        <WpsModelSettingsPanel
          open={open}
          backendAvailable={false}
          models={[]}
          currentModelId=""
          management={management}
          onClose={() => setOpen(false)}
          onAddConnection={() => void management.openConnectionCreator()}
        />
      );
    }

    const { container } = await mount(<Harness />);

    const addButton = findButton(container, "＋ 添加");
    expect(addButton?.hasAttribute("disabled")).toBe(false);
    await click(addButton);

    // 后台读取失败仍要进入可填写的表单，并且能看到失败说明；
    // 未完成连接测试前不能保存到共享模型目录。
    await until(
      () => textOf(container).includes("添加模型连接"),
      "后台不可读时也应能打开新增表单"
    );
    expect(
      findButton(container, "添加并使用")?.hasAttribute("disabled")
    ).toBe(true);
    expect(textOf(container)).toContain("仍可填写配置");
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
        backendAvailable
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
    // 公式模型下拉只列真实连接，不需要再靠黑名单过滤掉基础模式。
    expect(markup).toContain("跟随全局选择（默认）");

    const duplicateFeedbackMarkup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        backendAvailable
        models={[]}
        currentModelId=""
        management={{
          ...management,
          settingsFeedback: "模型「DS · deepseek-chat」已经使用相同的服务地址。"
        } as never}
        onClose={noop}
        onAddConnection={noop}
      />
    );
    expect(duplicateFeedbackMarkup).toContain(
      'class="model-settings-feedback error"'
    );
  });

  it("allows opening the model connection form while the local service is offline", () => {
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
      settingsFeedback: "Failed to fetch",
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
        backendAvailable={false}
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
    expect(markup).toContain("你仍可填写配置");

    const editorMarkup = renderToStaticMarkup(
      <WpsModelSettingsPanel
        open
        backendAvailable={false}
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
    expect(styles).toMatch(/\.eb-chat-header\s*\{[\s\S]*?z-index:\s*4;/);
    expect(styles).toMatch(/\.eb-model-menu\s*\{[\s\S]*?z-index:\s*45;/);
  });

  it("keeps the service status readable before collapsing the narrow header", () => {
    expect(styles).toMatch(
      /@media \(max-width: 620px\)[\s\S]*?\.eb-brand-copy span\s*\{[\s\S]*?max-width:\s*none;[\s\S]*?overflow:\s*visible;/
    );
    expect(styles).toMatch(
      /@media \(max-width: 620px\)[\s\S]*?\.eb-new-chat\s*\{[^}]*width:\s*28px;[^}]*font-size:\s*0;/
    );
  });
});
