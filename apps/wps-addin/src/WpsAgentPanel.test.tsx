import { readFileSync } from "node:fs";
import { act, useEffect, useState, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MODEL_STORAGE_KEY } from "../../excel-addin/src/conversation";
import { useModelManagement } from "../../excel-addin/src/hooks/useModelManagement";
import { getModelSettings, listModels, validatePlan, registerWpsExecution, getWpsExecutionStatus, startWpsAction, completeWpsAction, finishWpsExecution, suspendWpsExecution, claimWpsExecutionResume, reconcileWpsActionReceipt } from "../../excel-addin/src/api";
import WpsAgentPanel from "./WpsAgentPanel";
import { matrixFingerprintV2 } from "./wpsPlanExecutor";
import WpsModelSettingsPanel from "./WpsModelSettingsPanel";
import { runWpsReadonlyTurn } from "./wpsReadonlyAgent";
import { getWpsApplication } from "./wpsHost";
import {
  COMPOSER_AUTO_GROW_MAX_HEIGHT,
  COMPOSER_RESIZE_LABEL
} from "../../excel-addin/src/hooks/useComposerResize";
import { compareCommittedWpsResultToCopy, stageWpsNativeComparison, WpsFormulaTrialError } from "./wpsNativeComparison";
import type { AnalysisPlan } from "../../excel-addin/src/contracts";

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("../../excel-addin/src/api")>();
  // CAP-03 服务端登记 API：面板测试不测真实网络，返回一个"已授予写入权"的
  // 成功状态，让写路径能走完动作级检查点。真实登记失败/中止由独立测试覆盖。
  const wpsRunningState = (
    operationId: string,
    revision = 1,
    overrides: Record<string, unknown> = {}
  ) => ({
    operationId,
    status: "running" as const,
    revision,
    workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
    taskId: "task-test",
    descriptor: { actions: [] },
    authorizedSheetNames: ["Sheet1"],
    nextActionIndex: 0,
    totalActions: 1,
    inFlightActionIndex: null,
    beforeSnapshot: null,
    checkpoint: null,
    result: null,
    error: null,
    canResume: true,
    ownerActive: false,
    requiresReconciliation: false,
    ownerToken: "owner-token-0123456789abcdef0123456789abcdef",
    ...overrides
  });
  return {
    ...actual,
    listModels: vi.fn(),
    getModelSettings: vi.fn(),
    validatePlan: vi.fn(),
    registerWpsExecution: vi.fn(async (input: { operationId: string }) =>
      wpsRunningState(input.operationId)),
    getWpsExecutionStatus: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    startWpsAction: vi.fn(async (operationId: string, input: { expectedRevision: number }) =>
      wpsRunningState(operationId, input.expectedRevision + 1)),
    completeWpsAction: vi.fn(async (operationId: string, input: { expectedRevision: number }) =>
      wpsRunningState(operationId, input.expectedRevision + 1)),
    suspendWpsExecution: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    claimWpsExecutionResume: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    reconcileWpsActionReceipt: vi.fn(async (operationId: string) => wpsRunningState(operationId)),
    finishWpsExecution: vi.fn(async (operationId: string) => ({
      ...wpsRunningState(operationId),
      status: "completed" as const
    }))
  };
});

vi.mock("./wpsReadonlyAgent", () => ({
  runWpsReadonlyTurn: vi.fn()
}));

// 面板测试只验证“两次确认”状态机；原生复制、着色及失效检测由
// wpsNativeComparison.test.ts 的严格宿主替身独立覆盖。
// 注意：替身必须给全 `WpsNativeComparison` 的**每个**字段。少给字段时
// `vi.fn` 的类型很宽松、tsc 也放过它，但侧栏一旦真的用到就会在运行期崩
// —— 这跟宿主替身少给参数是同一类"替身比真实现宽容"的缺口。
vi.mock("./wpsNativeComparison", () => ({
  WpsFormulaTrialError: class WpsFormulaTrialError extends Error {},
  stageWpsNativeComparison: vi.fn((_application: unknown, plan: unknown) => {
    const prepared = plan as { actions: { values: unknown[][] }[] };
    return {
      plan,
      changedCells: ["A1"],
      reviews: [],
      original: { Name: "测试.xlsx" },
      copy: { Name: "Book1" },
      copyIdentity: { name: "Book1", fullName: null, path: null },
      originalBefore: prepared.actions.map(() => ({ values: [[]], formulas: [[]], valuesCaptured: true })),
      copyFingerprints: prepared.actions.map(() => "copy-fingerprint"),
      // 真实路径里回读值已与计划校对一致，替身照同一语义给。
      copyAfter: prepared.actions.map((action) => ({
        values: action.values,
        formulas: [],
        valuesCaptured: true
      }))
    };
  }),
  compareCommittedWpsResultToCopy: vi.fn(() => "matched"),
  verifyWpsNativeComparison: vi.fn(),
  discardWpsNativeComparison: vi.fn()
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
const registerWpsExecutionMock = vi.mocked(registerWpsExecution);
const getWpsExecutionStatusMock = vi.mocked(getWpsExecutionStatus);
const startWpsActionMock = vi.mocked(startWpsAction);
const completeWpsActionMock = vi.mocked(completeWpsAction);
const finishWpsExecutionMock = vi.mocked(finishWpsExecution);
const suspendWpsExecutionMock = vi.mocked(suspendWpsExecution);
const claimWpsExecutionResumeMock = vi.mocked(claimWpsExecutionResume);
const reconcileWpsActionReceiptMock = vi.mocked(reconcileWpsActionReceipt);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);
const getWpsApplicationMock = vi.mocked(getWpsApplication);
const stageComparisonMock = vi.mocked(stageWpsNativeComparison);
const compareCommittedMock = vi.mocked(compareCommittedWpsResultToCopy);

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

/**
 * 「备份与恢复」抽屉里的撤销入口。抽屉仍保留这个快捷入口，所以先打开它，
 * 再按容器限定取按钮 —— 消息流里的执行结果卡现在也有同名按钮，
 * 只用文本查找会指错对象。
 */
async function undoButton(container: HTMLElement): Promise<HTMLButtonElement | null> {
  if (!container.querySelector(".wps-backup-card")) {
    await click(
      container.querySelector<HTMLButtonElement>('button[aria-label="备份与恢复"]')
    );
  }
  return container.querySelector<HTMLButtonElement>(".wps-backup-card button");
}

/** 执行结果卡里的撤销按钮（抽屉里那个不算）。 */
function cardUndoButton(container: HTMLElement): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>(
    '.wps-verification-card button[aria-label="撤销本次执行"]'
  );
}

/**
 * 消息流里所有名为「撤销本次执行」的按钮（抽屉里的不算）。每一个都必须落在
 * **某张卡**里 —— 早先那张无条件挂在消息流末尾的全局撤销卡会在这里暴露成
 * "没有卡片容器的裸按钮"。
 */
function flowUndoButtons(container: HTMLElement): HTMLButtonElement[] {
  return Array.from(container.querySelectorAll<HTMLButtonElement>("button")).filter(
    (item) =>
      (item.textContent ?? "").trim() === "↶ 撤销本次执行" &&
      !item.closest(".tool-drawer")
  );
}

async function stageAndSave(container: HTMLElement): Promise<void> {
  await click(findButton(container, "生成右侧副本并排审阅"));
  expect(findButton(container, "确认写回原表")).not.toBeNull();
  await click(findButton(container, "确认写回原表"));
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

/**
 * 最小可用宿主：只提供**澄清复核**需要的东西 —— 一个稳定的工作簿对象、
 * 一张稳定的工作表对象、以及一个结构不变的已用区域。
 *
 * 澄清复核现在额外要求"能证明还是建卡时那个对象"（见 `wpsObjectIdentity`），
 * 所以凡是要走完澄清作答的用例都得先有一个真实宿主：给 `{}` 会 fail-closed。
 * 这里刻意把同一份 `sheet`/`workbook` 对象复用于每次读取，模拟真宿主
 * "同一张表始终是同一个包装对象"的行为。
 */
function clarificationHost(options: { sheetNames?: string[]; workbookName?: string } = {}) {
  const sheetNames = options.sheetNames ?? ["Sheet1"];
  const usedRange = { Rows: { Count: 3 }, Columns: { Count: 2 }, Row: 1, Column: 1 } as Record<
    string,
    unknown
  >;
  Object.defineProperties(usedRange, {
    Value2: { get: () => [["类别", "归属"], ["甲", "华东"], ["乙", "华南"]] },
    Item: { value: () => usedRange },
    MergeCells: { value: false }
  });
  const range = { Rows: { Count: 1 }, Columns: { Count: 2 }, MergeCells: false } as Record<
    string,
    unknown
  >;
  Object.defineProperties(range, {
    Value2: { get: () => [["类别", "归属"]] }
  });
  const sheets = sheetNames.map((Name, position) => ({
    Name,
    Index: position + 1,
    CodeName: `SheetCode${position + 1}`,
    ProtectContents: false,
    UsedRange: usedRange,
    Range: () => range
  }));
  const workbook = {
    Name: options.workbookName ?? "测试.xlsx",
    FullName: `C:\\Data\\${options.workbookName ?? "测试.xlsx"}`,
    Path: "C:\\Data",
    ReadOnly: false,
    Worksheets: { Count: sheets.length, Item: (index: number) => sheets[Number(index) - 1] },
    ActiveSheet: sheets[0]
  };
  const storage = new Map<string, unknown>();
  return {
    Version: "12.0",
    ActiveWorkbook: workbook,
    ActiveSheet: sheets[0],
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, item: unknown) => { storage.set(key, item); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never;
}

/**
 * 一块替身单元格。默认写入生效（执行器据此判定回读一致）；
 * `staleWrite` 时写入不落盘，于是回读核对**真的**失败 ——
 * 这样才能断言界面在失败时不谎报"验证通过"，而不是靠 mock 抛错绕开真实链路。
 */function writeRangeHost(options: {
  staleWrite?: boolean;
  /** 每次宿主写入后回调（参数是累计写入次数），用来在"第一次写入发生的那一刻"取证。 */
  onWrite?: (writes: number) => void;
  /** 写入一旦发生就让 `PluginStorage.getItem` 抛错：模拟恢复点读取失败。 */
  breakStorageReadsAfterWrite?: boolean;
} = {}) {
  let value = "旧值";
  let formula = "旧值";
  let writes = 0;
  const range = { Rows: { Count: 1 }, Columns: { Count: 1 }, MergeCells: false } as Record<string, unknown>;
  Object.defineProperties(range, {
    Value2: {
      get: () => value,
      set: (next: string) => {
        writes += 1;
        if (!options.staleWrite) {
          value = next;
          formula = next;
        }
        options.onWrite?.(writes);
      },
      enumerable: true
    },
    Formula: {
      get: () => formula,
      set: (next: string) => {
        writes += 1;
        if (!options.staleWrite) {
          formula = next;
          value = next;
        }
        options.onWrite?.(writes);
      },
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
      getItem: (key: string) => {
        if (options.breakStorageReadsAfterWrite && writes > 0) {
          throw new Error("PluginStorage 读取失败");
        }
        return storage.get(key);
      },
      setItem: (key: string, item: unknown) => { storage.set(key, item); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never);
  return {
    writes: () => writes,
    cellValue: () => value,
    /** 模拟"执行之后用户又手工改了这块区域"：撤销前指纹复核必须能识别出来。 */
    setCell: (next: string) => {
      value = next;
      formula = next;
    }
  };
}

/** 走到"写入前确认"这一步：两端对照用的同一任务状态。 */
async function mountPlanPreview(plan: AnalysisPlan, task: string) {
  listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
  runTurnMock.mockResolvedValue({ kind: "plan", plan, turnId: "turn-preview-1" });
  validatePlanMock.mockImplementation(async (submitted) => submitted);
  const { container } = await mount(panel());
  await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");
  const textarea = container.querySelector("textarea");
  await type(textarea, task);
  await submit(textarea);
  await until(() => textOf(container).includes("写入前确认"), "应展示写入计划预览");
  return { container, unmount: () => unmountMounted(container) };
}

/**
 * 卸载历史里挂载过的某个面板，并从 `mounted` 里摘掉它 —— `mount` 的清理是
 * 按数组走的，重复卸载同一个 root 会报错。供"面板重开"这类用例使用。
 */
function unmountMounted(container: HTMLElement): void {
  const entry = mounted.find((item) => item.container === container);
  if (!entry) return;
  entry.unmount();
  mounted = mounted.filter((item) => item !== entry);
}

const singleWritePlan: AnalysisPlan = {
  id: "plan-single-write",
  title: "更新 A1",
  summary: "把 A1 更新为新值",
  assumptions: [],
  warnings: [],
  actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [["90"]] }],
  acceptanceCriteria: []
};

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
  it("shows a structured planning stop as failure, not completed or a write preview", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    runTurnMock.mockResolvedValue({ kind: "answer", provider: "system", outcome: "failed", message: "规划预算已用完" });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "模型加载");
    const textarea = container.querySelector("textarea");
    await type(textarea, "生成预览");
    await submit(textarea);
    await until(() => textOf(container).includes("规划预算已用完"), "失败提示");
    expect(textOf(container)).toContain("执行失败");
    expect(textOf(container)).not.toContain("已完成");
    expect(textOf(container)).not.toContain("生成右侧副本并排审阅");
    expect(validatePlanMock).not.toHaveBeenCalled();
  });
  it("launches the read-only wide-window probe from the real chat pane without replacing the conversation", async () => {
    const showDialog = vi.fn((..._args: unknown[]) => true);
    getWpsApplicationMock.mockReturnValue({
      Version: "12.0",
      ShowDialog: showDialog
    } as never);
    listModelsMock.mockResolvedValue({ defaultModelId: null, models: [] });

    const { container } = await mount(panel());
    const textarea = container.querySelector("textarea");
    await type(textarea, "当前对话仍在这里");
    await click(container.querySelector('[aria-label="更多"]'));
    const probeEntry = Array.from(container.querySelectorAll('[role="menuitem"]'))
      .find((item) => item.textContent?.includes("探测独立宽窗口"));
    await click(probeEntry ?? null);

    expect(showDialog).toHaveBeenCalledTimes(1);
    expect(showDialog.mock.calls[0][4]).toBe(false);
    expect(container.querySelector('[aria-label="宽窗口探测结果"]')).not.toBeNull();
    expect(container.querySelector("textarea")?.value).toBe("当前对话仍在这里");
    expect(container.querySelector(".wps-chat-shell")).not.toBeNull();
  });

  it("keeps header actions icon-only, named and connected to their existing surfaces", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "模型加载完成");
    expect(container.querySelector(".brand-copy")?.textContent?.trim()).toBe("Excel Bro");
    for (const [selector, glyph] of [
      [".model-picker-trigger", ""],
      [".new-chat-entry", "＋"],
      [".tools-entry", ""],
      [".folder-backup-entry", "↶"]
    ]) {
      const button = container.querySelector<HTMLButtonElement>(selector);
      expect(button?.textContent).toBe(glyph);
      expect(button?.getAttribute("aria-label")).toBeTruthy();
      expect(button?.title).toBeTruthy();
      expect(button?.querySelector('svg, [aria-hidden="true"]')).not.toBeNull();
    }
    await click(container.querySelector(".model-picker-trigger"));
    expect(container.querySelector('.model-menu[role="menu"]')).not.toBeNull();
    expect(container.querySelector('[role="menuitemradio"][aria-checked="true"]')?.textContent).toContain("第一个连接");
    await click(container.querySelector(".tools-entry"));
    expect(container.querySelector('aside[aria-label="我的工具"]')).not.toBeNull();
    await click(container.querySelector(".folder-backup-entry"));
    expect(container.querySelector('aside[aria-label="备份与恢复"]')).not.toBeNull();
    expect(runTurnMock).not.toHaveBeenCalled();
  });

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

    expect(container.querySelector('.model-picker-trigger')?.getAttribute("aria-label")).toBe("添加模型");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");

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
    expect(findButton(container, "生成右侧副本并排审阅")).toBeNull();
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");

    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
    await type(textarea, "统计当前表");
    await submit(textarea);
    const question = Array.from(container.querySelectorAll(".message-row.user"))
      .find((item) => item.textContent?.includes("统计当前表"));
    const processMessage = question?.nextElementSibling;
    expect(processMessage?.classList.contains("message-row")).toBe(true);
    expect(processMessage?.querySelector(".message-author")?.textContent?.trim()).toBe("Excel Bro");
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
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea");
      await type(textarea, "这是一条提问");
      await submit(textarea);
      await until(() => textOf(container).includes("这是一条回答"), "应收到回答");

      const userCopy = container.querySelector<HTMLButtonElement>(".message-row.user .message-copy-button");
      expect(userCopy?.getAttribute("aria-label")).toBe("复制这条消息");
      expect(userCopy?.title).toBe("复制这条消息");
      expect(userCopy?.querySelector("svg rect")).not.toBeNull();
      expect(userCopy?.textContent).toBe("");
      await click(userCopy);
      expect(legacyCopy).toHaveBeenCalledTimes(1);
      expect(container.querySelector(".message-row.user .message-copy-button")?.textContent).toContain("已复制");
      const answer = Array.from(container.querySelectorAll(".message-row.assistant"))
        .find((item) => item.textContent?.includes("这是一条回答"));
      const answerCopy = answer?.querySelector<HTMLButtonElement>(".message-copy-button");
      expect(answerCopy?.querySelector("svg rect")).not.toBeNull();
      expect(answerCopy?.textContent).toBe("");
      await click(answerCopy ?? null);
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
      () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第二个连接") ?? false,
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
    // 澄清复核要能证明"还是建卡时那个对象"，所以这里必须给一个真实宿主。
    getWpsApplicationMock.mockReturnValue(clarificationHost());
    runTurnMock
      .mockResolvedValueOnce({ kind: "answer", message: "你好！", resultContext: null, provider: "model" })
      .mockResolvedValueOnce({
        kind: "clarification",
        clarification: {
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
        }
      })
      .mockResolvedValueOnce({ kind: "answer", message: "我来按你的补充说明处理。", resultContext: null, provider: "model" });

    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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

  it("renders separate fields and continues the original request with all answers, without executing", async () => {
    listModelsMock.mockResolvedValue({ models: [firstConnection], defaultModelId: firstConnection.id });
    getWpsApplicationMock.mockReturnValue(clarificationHost());
    runTurnMock.mockResolvedValueOnce({ kind: "clarification", clarification: { response: {
      kind: "clarification", provider: "model", clarification: {
        id: "fields-1", summary: "补充口径", question: "请填写两项信息", reason: "尚缺口径",
        scopeLabel: "Sheet1", options: [], fields: [
          { id: "period", label: "时间范围", placeholder: "说明时间范围" }, { id: "unit", label: "统计单位" }
        ]
      }
    } } }).mockResolvedValueOnce({ kind: "answer", message: "收到补充，准备预览", provider: "model" });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "模型应就绪");
    const prompt = container.querySelector("textarea");
    await type(prompt, "按已选表统计并预览");
    await submit(prompt);
    await until(() => !!container.querySelector(".clarification-fields"), "规划追问应显示表单");
    const button = findButton(container, "提交补充并生成预览");
    if (!button) throw new Error("缺少表单提交按钮");
    expect(button.disabled).toBe(true);
    await type(container.querySelector('[aria-label="时间范围"]'), "本季度");
    expect(button.disabled).toBe(true);
    await type(container.querySelector('[aria-label="统计单位"]'), "件");
    expect(button.disabled).toBe(false);
    await click(button);
    await until(() => runTurnMock.mock.calls.length === 2, "提交应继续原任务");
    expect(runTurnMock.mock.calls[1][0]).toMatchObject({
      prompt: "按已选表统计并预览\n补充确认：时间范围：本季度\n统计单位：件", intentConfirmed: true
    });
    expect(stageComparisonMock).not.toHaveBeenCalled();
    expect(validatePlanMock).not.toHaveBeenCalled();
    expect(container.querySelector(".clarification-card.resolved")).not.toBeNull();
  });

  it("keeps a cancelled clarification in history without leaving live choices", async () => {
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    getWpsApplicationMock.mockReturnValue(clarificationHost());
    runTurnMock.mockResolvedValue({
      kind: "clarification",
      clarification: {
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
      }
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
    expect(findButton(container, "生成右侧副本并排审阅")).toBeNull();
    const editBox = container.querySelector<HTMLTextAreaElement>('[aria-label="编辑写入计划 JSON"]');
    await type(editBox, JSON.stringify({
      ...writePlan,
      actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [["修订值"]] }]
    }));
    await click(findButton(container, "应用修改并重新预览"));
    await until(() => textOf(container).includes("修订值") && findButton(container, "生成右侧副本并排审阅") !== null, "编辑后应重新预览最终值");
    expect(writes).toBe(0);
    await stageAndSave(container);
    await until(
      () => Boolean(container.querySelector('button[aria-label="备份与恢复"] b')),
      "确认后应完成写入并提供恢复入口"
    );
    expect(writes).toBe(1);
    expect(value).toBe("修订值");
    expect(textOf(container)).toContain("计划已执行并回读确认");
    const resultMessage = container.querySelector(".wps-verification-card.verified")?.closest(".message-row.assistant");
    expect(resultMessage?.querySelector(".message-author")?.textContent?.trim()).toBe("Excel Bro");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
    const textarea = container.querySelector("textarea");
    await type(textarea, "把类别列翻译成归属");
    await submit(textarea);
    await until(() => textOf(container).includes("写入前确认"), "应展示写入计划");

    expect(validatePlanMock).not.toHaveBeenCalled();
    expect(textOf(container)).toContain("Sheet1!A1");
    await click(findButton(container, "编辑计划"));
    await click(findButton(container, "应用修改并重新预览"));
    await until(() => findButton(container, "生成右侧副本并排审阅") !== null, "本地计划编辑后应重新预览");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
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
    expect(textOf(container)).toContain("点击上方模型图标开始连接");
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
      () => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第二个连接") ?? false,
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

  it("mirrors the Excel plan card hierarchy for the same pending task", async () => {
    // Excel 用「采用的信息」/「补充说明」两个区块承载服务端下发的口径；WPS
    // 拿到的是同一个 AnalysisPlan，这两个前提不该在预览里丢掉。
    writeRangeHost();
    const { container } = await mountPlanPreview({
      ...singleWritePlan,
      id: "plan-notes",
      assumptions: ["源列空白按「未分类」处理。"],
      warnings: ["汇总表已有数据，写入会覆盖原值。"]
    }, "按部门汇总平均分");

    const notes = Array.from(
      container.querySelector(".wps-plan-card")?.querySelectorAll(".wps-plan-notes") ?? []
    );
    expect(notes.map((note) => note.querySelector("strong")?.textContent?.trim()))
      .toEqual(["采用的信息", "补充说明"]);
    expect(notes[0]?.textContent).toContain("源列空白按「未分类」处理。");
    expect(notes[1]?.textContent).toContain("汇总表已有数据，写入会覆盖原值。");
    // 恢复能力必须写成"会话内"，不能给人永久可撤销的错觉。
    const meta = Array.from(container.querySelectorAll(".wps-plan-meta span"))
      .map((item) => item.textContent?.trim());
    expect(meta).toContain("会话内可撤销");
    // 还没确认，就不得出现任何执行结果。
    expect(container.querySelector(".wps-verification-card")).toBeNull();
    expect(textOf(container)).not.toContain("写入验证通过");
  });

  it("hides both note blocks when the plan states neither", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    expect(container.querySelector(".wps-plan-assumptions")).toBeNull();
    expect(container.querySelector(".wps-plan-warnings")).toBeNull();
    expect(textOf(container)).not.toContain("采用的信息");
  });

  it("marks the preview as edited only after the user rewrote the plan", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    expect(textOf(container)).not.toContain("已按你的修改更新");
    await click(findButton(container, "编辑计划"));
    const editBox = container.querySelector<HTMLTextAreaElement>('[aria-label="编辑写入计划 JSON"]');
    await type(editBox, JSON.stringify({ ...singleWritePlan, summary: "改成最大值" }));
    await click(findButton(container, "应用修改并重新预览"));
    await until(() => textOf(container).includes("已按你的修改更新"), "编辑后重新预览必须有可见标识");

    expect(container.querySelector(".wps-plan-card")?.textContent).toContain("写入前请重新核对");
    expect(host.writes()).toBe(0);
    // 标识跟着消息持久化：重新载入面板后仍能知道这份计划被改过。
    const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    const savedPreview = saved.conversations[0].messages.find(
      (message: { planPreview?: { revised?: boolean } }) => message.planPreview
    );
    expect(savedPreview.planPreview.revised).toBe(true);
  });

  it("keeps the original untouched during native review and discards the copy without committing", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "先并排审阅结果");
    await click(findButton(container, "生成右侧副本并排审阅"));
    expect(host.writes()).toBe(0);
    expect(host.cellValue()).toBe("旧值");
    expect(findButton(container, "确认写回原表")).not.toBeNull();
    expect(textOf(container)).toContain("左侧原表尚未写入");
    await click(findButton(container, "放弃副本"));
    expect(host.writes()).toBe(0);
    expect(host.cellValue()).toBe("旧值");
    expect(findButton(container, "确认写回原表")).toBeNull();
    expect(textOf(container)).toContain("已放弃右侧审阅副本");
  });

  it("公式副本计算值不符时重新预览固定值副本，原表仍待用户确认", async () => {
    const host = writeRangeHost();
    const formulaPlan: AnalysisPlan = {
      ...singleWritePlan,
      id: "wps-map-column-test",
      actions: [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=1+1"]] }],
      acceptanceCriteria: [{ type: "rangeEquals", sheet: "Sheet1", range: "A1", expected: [["2"]] }]
    };
    const { container } = await mountPlanPreview(formulaPlan, "用公式计算 A1");
    stageComparisonMock.mockImplementationOnce(() => { throw new WpsFormulaTrialError("计算不符"); });
    await click(findButton(container, "生成右侧副本并排审阅"));
    expect(stageComparisonMock).toHaveBeenCalledTimes(2);
    const secondPrepared = stageComparisonMock.mock.calls[1][1];
    expect(secondPrepared.plan.actions[0]).toEqual({
      type: "writeValues", sheet: "Sheet1", range: "A1", values: [["2"]]
    });
    expect(textOf(container)).toContain("改用固定值重新生成");
    expect(findButton(container, "确认写回原表")).not.toBeNull();
    expect(host.writes()).toBe(0);
  });

  it("回读不符时不谎报验证通过，但确有恢复点的部分写入仍必须给出卡与撤销入口", async () => {
    // 写入不落盘 → 回读与计划不同。执行器据此判定为「可恢复的部分写入」，
    // 界面要如实说"写入未完成、恢复资料已保留"，**绝不**说验证通过；
    // 也不能因为"不是干净成功"就把恢复点藏起来。
    const host = writeRangeHost({ staleWrite: true });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    await stageAndSave(container);
    await until(() => container.querySelector(".message-row.system") !== null, "失败应给出系统反馈");

    // 负向断言必须先有正向锚点，否则只是"什么都没渲染"的空过。
    expect(textOf(container)).not.toContain("写入验证通过");
    expect(container.querySelector(".wps-verification-card.verified")).toBeNull();
    expect(container.querySelector(".wps-verification-card.unverified")).toBeNull();

    const card = container.querySelector(".wps-verification-card.partial");
    expect(card).not.toBeNull();
    expect(card?.textContent).toContain("恢复资料已保留");
    expect(cardUndoButton(container)).not.toBeNull();

    expect(container.querySelector(".message-row.system")?.textContent)
      .toContain("回读结果与计划不一致");
    // 计划仍在等待确认：可以修正后重试，界面不能假装已经写完了。
    expect(container.querySelector(".wps-plan-card")).not.toBeNull();
    expect(textOf(container)).toContain("写入前确认");
    expect(host.writes()).toBeGreaterThan(0);
  });

  it("写回后计算值与副本不同，不发验证通过卡但结果卡内保留撤销入口", async () => {
    const host = writeRangeHost();
    compareCommittedMock.mockReturnValueOnce("different");
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    expect(host.writes()).toBeGreaterThan(0);
    expect(textOf(container)).toContain("计算结果与右侧副本不同");
    expect(textOf(container)).toContain("最终计算值未获证实");
    // 它不是 verified，但**是**一张带恢复点的结果卡 —— 入口必须在这张卡里。
    expect(container.querySelector(".wps-verification-card.verified")).toBeNull();
    const card = container.querySelector(".wps-verification-card.unverified");
    expect(card).not.toBeNull();
    expect(card?.textContent).not.toContain("写入验证通过");
    expect(cardUndoButton(container)).not.toBeNull();
    expect(await undoButton(container)).not.toBeNull();
  });

  it("switches the verified card to restored and keeps recovery session-scoped", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );
    expect(container.querySelector(".wps-verification-card")?.textContent)
      .toContain("1 个动作 · 1 个单元格已回读核对");

    await click(await undoButton(container));
    await until(() => textOf(container).includes("已撤销并核对恢复"), "撤销后应给出核对过的恢复结果");

    const undone = container.querySelector(".wps-verification-card.undone");
    expect(undone).not.toBeNull();
    expect(undone?.textContent).not.toContain("写入验证通过");
    expect(undone?.textContent).toContain("恢复点仅在当前加载项会话内可用");
    expect(container.querySelector(".wps-verification-card.verified")).toBeNull();
    expect(host.cellValue()).toBe("旧值");
  });

  it("uses native review before commit, then keeps only the verified result and undo summary", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "按计划更新 A1");

    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );

    // Successful WPS writes no longer reopen the obsolete narrow HTML grid.
    expect(container.querySelector(".change-review-summary")).toBeNull();
    expect(container.querySelector(".change-review-table")).toBeNull();
    expect(container.querySelector(".wps-chat-shell")).not.toBeNull();

    // 明细不进聊天历史：持久化下来的只有定位与计数，没有单元格内容。
    const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(saved).not.toBeNull();
    // "旧值"是单元格内容，不该出现在被持久化的对话里（时间戳之类的数字串不适合
    // 做这类断言，故另设哨兵用例，见下一条）。
    expect(JSON.stringify(saved)).not.toContain("旧值");

    await click(await undoButton(container));
    await until(() => textOf(container).includes("已撤销并核对恢复"), "撤销后应给出核对过的恢复结果");
    // 区域已回到执行前，不能继续挂着过期的变更明细。
    expect(container.querySelector(".change-review-summary")).toBeNull();
    expect(host.cellValue()).toBe("旧值");
    // 但"这次确实改过、后来撤销了"是操作记录，不该跟着明细一起被抹掉；
    // 留下来的形态只有定位与计数，不含任何单元格内容。
    const note = container.querySelector(".change-review-undo-note");
    expect(note).not.toBeNull();
    expect(note?.textContent).toContain("Sheet1!A1");
    expect(note?.textContent).toContain("改动 1 / 1");
    expect(note?.textContent).not.toContain("旧值");
    expect(note?.textContent).not.toContain("90");
  });

  it("keeps the 'changed then undone' note free of cell values in the UI and in history", async () => {
    const SENTINEL = "私密撤销值-Q7W2";
    writeRangeHost();
    const { container } = await mountPlanPreview({
      ...singleWritePlan,
      id: "plan-undo-sentinel",
      actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [[SENTINEL]] }]
    }, "按计划更新 A1");

    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );
    expect(container.querySelector(".change-review-summary")).toBeNull();

    await click(await undoButton(container));
    await until(
      () => container.querySelector(".change-review-undo-note") !== null,
      "撤销后应留下「曾修改、现已撤销」的摘要"
    );
    // 明细已释放，摘要里也不该有单元格内容。
    expect(container.querySelector(".change-review-summary")).toBeNull();
    expect(container.querySelector(".change-review-undo-note")?.textContent).not.toContain(SENTINEL);
    const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(JSON.stringify(saved)).not.toContain(SENTINEL);
  });

  it("attaches the change review to the failure message and never claims success", async () => {
    writeRangeHost({ staleWrite: true });
    const { container } = await mountPlanPreview(singleWritePlan, "按计划更新 A1");

    await stageAndSave(container);
    await until(() => container.querySelector(".message-row.system") !== null, "失败应给出系统反馈");

    const summary = container.querySelector(".message-row.system .change-review-summary");
    expect(summary).not.toBeNull();
    expect(summary!.textContent).toContain("回读与预期不一致");
    expect(summary!.querySelector("button")).toBeNull();
    expect(container.textContent).not.toContain("打开全幅双视图");
    expect(container.querySelector(".wps-review-shell")).toBeNull();
    expect(summary!.textContent).not.toContain("写入验证通过");
  });

  it("keeps cell values out of the persisted conversation even when a review was delivered", async () => {
    // 哨兵只存在于工作簿单元格与回读矩阵里：用户提示、计划标题、活动日志都不含它。
    // 这样才能证明"明细不进历史"是真的，而不是碰巧没撞上子串。
    const SENTINEL = "单元格私密值-K9Z3";
    writeRangeHost();
    const { container } = await mountPlanPreview({
      ...singleWritePlan,
      id: "plan-sentinel",
      actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [[SENTINEL]] }]
    }, "按计划更新 A1");

    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card.verified") !== null,
      "写入成功应给出回读核对结果"
    );
    // 成功结果不再提供旧的窄 HTML 对照；副本逐格审阅由严格原生宿主用例覆盖。
    expect(container.querySelector(".change-review-summary")).toBeNull();

    const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null");
    expect(saved).not.toBeNull();
    expect(JSON.stringify(saved)).not.toContain(SENTINEL);
  });

  describe("待答澄清的授权身份", () => {
    /**
     * 造一台可切换工作表身份的宿主。
     *
     * 澄清卡只有在**授权身份没变**时才允许继续回答。索引只是"第几张表"，
     * 重排/替换之后同一个索引会指向另一张表 —— 那时用户对着旧表回答的问题
     * 会被拿去算新表，授权边界形同虚设。所以这里的判据是**名称**，不是索引。
     */
    function ambiguityHost(sheetNames: string[]) {
      const sheets = sheetNames.map((name) => ({ Name: name, ProtectContents: false }));
      const storage = new Map<string, unknown>();
      const application = {
        Version: "12.0",
        ActiveWorkbook: {
          Name: "测试.xlsx",
          FullName: "C:\\Data\\测试.xlsx",
          Path: "C:\\Data",
          ReadOnly: false,
          Worksheets: {
            Count: sheets.length,
            Item: (index: number) => sheets[Number(index) - 1]
          },
          ActiveSheet: sheets[0]
        },
        ActiveSheet: sheets[0],
        PluginStorage: {
          getItem: (key: string) => storage.get(key),
          setItem: (key: string, item: unknown) => { storage.set(key, item); }
        },
        CreateTaskPane: () => ({ ID: 1, Visible: true }),
        GetTaskPane: () => undefined
      } as never;
      return {
        application,
        rename: (position: number, name: string) => { sheets[position].Name = name; },
        reorder: (next: string[]) => {
          const byName = new Map(sheets.map((sheet) => [sheet.Name, sheet]));
          next.forEach((name, position) => {
            const sheet = byName.get(name);
            if (sheet) sheets[position] = sheet;
          });
        }
      };
    }

    /** 让这一轮返回一张**需求澄清**卡（与列映射归因无关，专测范围失效判定）。 */
    function ambiguityCard() {
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
        }
      });
    }

    /** 走到待答澄清卡为止。 */
    async function pendingClarification() {
      const { container } = await mount(panel());
      await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
      const composer = container.querySelector<HTMLTextAreaElement>(".composer textarea");
      await type(composer, "帮我处理表格");
      await submit(composer);
      await until(
        () => Boolean(container.querySelector(".clarification-card")),
        "应显示澄清卡"
      );
      await until(() => runTurnMock.mock.calls.length === 1, "应完成这一轮");
      return container;
    }

    /** 卡片上的选项按钮 —— 与"取消本次澄清"区分开。 */
    function clarificationOptions(container: HTMLElement): HTMLButtonElement[] {
      return Array.from(
        container.querySelectorAll<HTMLButtonElement>(
          ".clarification-options button"
        )
      );
    }

    it("同索引换成另一张表：待答澄清失效，不拿旧答案继续算", async () => {
      listModelsMock.mockResolvedValue({
        defaultModelId: firstConnection.id,
        models: [firstConnection]
      });
      ambiguityCard();
      const host = ambiguityHost(["订单", "明细"]);
      getWpsApplicationMock.mockReturnValue(host.application);

      const container = await pendingClarification();
      expect(clarificationOptions(container)).toHaveLength(1);
      // 授权身份变了：第 1 张表从「订单」换成「客户」。
      host.rename(0, "客户");

      await click(clarificationOptions(container)[0]);
      await until(
        () => container.querySelector(".clarification-card.invalidated") !== null,
        "同索引换表后待答澄清必须失效"
      );
      // 没有第二轮：失效的澄清不能带着旧上下文继续跑。
      expect(runTurnMock).toHaveBeenCalledTimes(1);
      expect(textOf(container)).toContain("数据范围已经变化");
      // 原需求回到输入框，由用户重新确认范围后再发。
      expect(
        container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value
      ).toContain("帮我处理表格");
    });

    it("工作表重排：索引指向了另一张表，待答澄清失效", async () => {
      listModelsMock.mockResolvedValue({
        defaultModelId: firstConnection.id,
        models: [firstConnection]
      });
      ambiguityCard();
      const host = ambiguityHost(["订单", "明细"]);
      getWpsApplicationMock.mockReturnValue(host.application);

      const container = await pendingClarification();
      // 重排：原第 2 张「明细」挪到第 1 位，索引 1 现在指向另一张表。
      host.reorder(["明细", "订单"]);

      await click(clarificationOptions(container)[0]);
      await until(
        () => container.querySelector(".clarification-card.invalidated") !== null,
        "重排后待答澄清必须失效"
      );
      expect(runTurnMock).toHaveBeenCalledTimes(1);
    });

    it("已授权表改名（顺序与集合都不变）：待答澄清失效并要求重新授权", async () => {
      listModelsMock.mockResolvedValue({
        defaultModelId: firstConnection.id,
        models: [firstConnection]
      });
      ambiguityCard();
      const host = ambiguityHost(["订单", "明细"]);
      getWpsApplicationMock.mockReturnValue(host.application);

      const container = await pendingClarification();
      host.rename(0, "订单（新）");

      await click(clarificationOptions(container)[0]);
      await until(
        () => container.querySelector(".clarification-card.invalidated") !== null,
        "改名会让授权身份变化，待答澄清必须失效"
      );
      expect(runTurnMock).toHaveBeenCalledTimes(1);
    });

    it("授权范围与身份都没变：澄清可以正常继续", async () => {
      listModelsMock.mockResolvedValue({
        defaultModelId: firstConnection.id,
        models: [firstConnection]
      });
      runTurnMock
        .mockResolvedValueOnce({
          kind: "clarification",
          clarification: {
            response: {
              kind: "clarification",
              provider: "model",
              clarification: {
                id: "clarify-keep",
                summary: "需要明确目标",
                question: "要查询还是修改？",
                reason: "两种操作不同",
                scopeLabel: "当前表",
                options: [
                  { id: "read", label: "查询", description: "只读", resolution: "查询数据" }
                ]
              }
            }
          }
        })
        .mockResolvedValueOnce({
          kind: "answer",
          message: "已按补充说明处理。",
          resultContext: null,
          provider: "model"
        });
      const host = ambiguityHost(["订单", "明细"]);
      getWpsApplicationMock.mockReturnValue(host.application);

      const container = await pendingClarification();
      // 只做一次**无副作用**的读取类调用：身份不变，卡片必须保留。
      await click(clarificationOptions(container)[0]);
      await until(() => runTurnMock.mock.calls.length === 2, "身份未变时应继续原任务");
      expect(container.querySelector(".clarification-card.resolved")).not.toBeNull();
      expect(textOf(container)).not.toContain("数据范围已经变化");
    });
  });

  /**
   * 侧栏信息层级的**结构与状态**断言。
   *
   * 这里能覆盖的是"哪些元素在哪、什么时候出现、按钮叫什么"，**不是布局**：
   * jsdom 不计算真实排版，`.wps-native-review-actions` 里的 DOM 断言无法证明
   * 按钮在 350/430/700 CSS px 下没有被挤压。CSS 契约只锁住"声明本身还在"，
   * 真机视口表现必须另外用 WPS 截图验收（见
   * `docs/WPS_NATIVE_REVIEW_UI_GUIDE_2026-09-30.md` 第 4 节）。
   */
describe("撤销按钮归属：执行结果卡内，按 executionId 绑定", () => {
    /** 再跑一轮同样的计划（同一容器、同一宿主），用于"连续执行"的用例。 */
    async function runAnotherPlan(container: HTMLElement, task: string) {
      const textarea = container.querySelector("textarea");
      await type(textarea, task);
      await submit(textarea);
      await until(
        () => findButton(container, "生成右侧副本并排审阅") !== null,
        "应展示新一轮计划预览"
      );
      await stageAndSave(container);
    }

    it("写入成功后：按钮在对应结果卡内，消息流末尾不再有全局撤销节点", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(container);
      await until(
        () => container.querySelector(".wps-verification-card") !== null,
        "确认写入后应给出执行结果卡"
      );

      // 正向锚点：按钮确实在产生恢复点的那张卡里，并且卡上写明状态。
      const card = container.querySelector(".wps-verification-card");
      expect(cardUndoButton(container)).not.toBeNull();
      expect(card?.querySelector(".wps-undo-inline")?.textContent).toContain("可撤销");

      // 负向：那张无条件挂在消息流末尾的全局撤销卡不许回来……
      expect(container.querySelector(".wps-undo-card")).toBeNull();
      // ……而且消息流里任何同文案按钮都必须落在某张卡（结果卡/备份卡）里面，
      // 不允许存在"没有卡片容器的裸按钮"。
      const buttons = flowUndoButtons(container);
      expect(buttons.length).toBeGreaterThan(0);
      for (const button of buttons) {
        expect(button.closest(".wps-verification-card, .wps-backup-card")).not.toBeNull();
      }

      // 抽屉仍保留快捷入口，并标明它属于哪一次执行。
      const opener = container.querySelector<HTMLButtonElement>('button[aria-label="备份与恢复"]');
      await click(opener);
      expect(await undoButton(container)).not.toBeNull();
      expect(container.querySelector(".wps-backup-card")?.textContent).toContain("所属执行");
    });

    it("之后来一条无关回复，按钮仍留在旧结果卡里，不会被赶到最新回复下面", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(container);
      await until(
        () => container.querySelector(".wps-verification-card") !== null,
        "确认写入后应给出执行结果卡"
      );
      const card = container.querySelector(".wps-verification-card")!;

      runTurnMock.mockResolvedValue({
        kind: "answer",
        message: "好的，还有什么需要处理的？",
        provider: "model",
        resultContext: null
      });
      const textarea = container.querySelector("textarea");
      await type(textarea, "随便聊聊");
      await submit(textarea);
      await until(() => textOf(container).includes("还有什么需要处理的"), "应显示新的普通回复");

      // 旧卡（同一个 DOM 节点）连同按钮一起按时间顺序往上走，没有被重建或挪走。
      expect(container.contains(card)).toBe(true);
      expect(cardUndoButton(container)).not.toBeNull();

      // 最新那条普通回复里没有任何撤销入口。
      const rows = Array.from(container.querySelectorAll<HTMLElement>(".message-row"));
      const last = rows[rows.length - 1];
      expect(last.textContent).toContain("还有什么需要处理的");
      expect(last.querySelector(".wps-undo-inline")).toBeNull();
      expect(container.querySelector(".wps-undo-card")).toBeNull();
    });

    it("连续两次执行：只有当前活跃恢复点那一次可点，旧卡降级为「恢复点已不可用」", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "第一次更新 A1");
      await stageAndSave(container);
      await until(
        () => container.querySelectorAll(".wps-verification-card").length === 1,
        "第一次执行应给出一张结果卡"
      );

      await runAnotherPlan(container, "第二次更新 A1");
      await until(
        () => container.querySelectorAll(".wps-verification-card").length === 2,
        "第二次执行后应有两张结果卡"
      );

      const cards = Array.from(container.querySelectorAll<HTMLElement>(".wps-verification-card"));
      // 旧卡在前、新卡在后：它是按时间顺序待在自己那条消息里的。
      expect(cards[0].querySelector(".wps-undo-inline")?.textContent).toContain("恢复点已不可用");
      expect(cards[0].querySelector('button[aria-label="撤销本次执行"]')).toBeNull();
      expect(cards[1].querySelector(".wps-undo-inline")?.textContent).toContain("可撤销");
      expect(cards[1].querySelector('button[aria-label="撤销本次执行"]')).not.toBeNull();
      // 全局节点依旧不存在。
      expect(container.querySelector(".wps-undo-card")).toBeNull();
    });

    it("撤销冲突：执行后区域又被改动时，卡片改为「撤销冲突」且不再可点", async () => {
      const host = writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(container);
      await until(
        () => container.querySelector(".wps-verification-card") !== null,
        "确认写入后应给出执行结果卡"
      );
      expect(cardUndoButton(container)).not.toBeNull();

      // 模拟用户在执行之后手工改了这块区域。
      host.setCell("人工改过的值");

      await click(cardUndoButton(container));
      await until(() => textOf(container).includes("又被修改"), "应给出冲突提示");

      const card = container.querySelector(".wps-verification-card")!;
      expect(card.querySelector(".wps-undo-inline")?.textContent).toContain("撤销冲突");
      expect(card.querySelector('button[aria-label="撤销本次执行"]')).toBeNull();
      // 冲突不等于"已经被撤销"：区域里是用户后来写的内容，不能被覆盖。
      expect(card.className).not.toContain("undone");
      expect(host.cellValue()).toBe("人工改过的值");
    });

    it("面板重开（新的加载项会话）：旧结果卡只显示「恢复点已不可用」，没有可点的假按钮", async () => {
      writeRangeHost();
      const first = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(first.container);
      await until(
        () => cardUndoButton(first.container) !== null,
        "第一次会话里应给出可点的撤销入口"
      );
      first.unmount();
      mounted = mounted.filter((entry) => entry !== first);

      // 重开：聊天历史（localStorage）还在，但当前加载项会话里已没有恢复点。
      writeRangeHost();
      const second = await mount(panel());
      await until(
        () => second.container.querySelector(".wps-verification-card") !== null,
        "历史里的执行结果卡仍应显示"
      );

      const card = second.container.querySelector(".wps-verification-card")!;
      expect(card.querySelector(".wps-undo-inline")?.textContent).toContain("恢复点已不可用");
      expect(cardUndoButton(second.container)).toBeNull();
      // 也不能靠抽屉里的残留入口变相提供一个可点的假按钮。
      expect(await undoButton(second.container)).toBeNull();
    });

    it("新建独立对话：旧执行卡留在旧对话，不会移到新对话的消息末尾", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(container);
      await until(
        () => container.querySelector(".wps-verification-card") !== null,
        "确认写入后应给出执行结果卡"
      );

      await click(container.querySelector('button[aria-label="新对话"]'));
      await until(() => textOf(container).includes("已新建对话"), "应新建对话");

      // 新对话里既没有那张卡，也没有任何撤销入口。
      expect(container.querySelector(".wps-verification-card")).toBeNull();
      expect(flowUndoButtons(container).length).toBe(0);
      expect(container.querySelector(".wps-undo-card")).toBeNull();

      // 旧执行卡留在旧对话：切回去还在，入口仍在卡里。
      await click(container.querySelector('[aria-label="更多"]'));
      await click(container.querySelector('button[aria-label="历史对话"]'));
      const entry = Array.from(container.querySelectorAll<HTMLElement>(".history-open-button")).find(
        (item) => item.textContent?.includes("把 A1 改成 90")
      );
      await click(entry ?? null);
      await until(
        () => container.querySelector(".wps-verification-card") !== null,
        "旧对话里的执行结果卡应保留"
      );
      expect(cardUndoButton(container)).not.toBeNull();
      expect(container.querySelector(".wps-undo-card")).toBeNull();
    });

    it("成功撤销后：所属执行卡变为「已撤销」，且不再残留可点按钮", async () => {
      const host = writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await stageAndSave(container);
      await until(
        () => container.querySelector(".wps-verification-card.verified") !== null,
        "写入成功应给出回读核对结果"
      );

      await click(cardUndoButton(container));
      await until(
        () => container.querySelector(".wps-verification-card.undone") !== null,
        "撤销后所属执行卡应变为已撤销"
      );

      const card = container.querySelector(".wps-verification-card")!;
      expect(card.querySelector(".wps-undo-inline")?.textContent).toContain("已撤销");
      expect(cardUndoButton(container)).toBeNull();
      expect(container.querySelector(".wps-verification-card.verified")).toBeNull();
      expect(host.cellValue()).toBe("旧值");
      // 抽屉也不该再提供入口。
      expect(await undoButton(container)).toBeNull();
    });
  });

  describe("审阅侧栏的信息层级", () => {
    it("keeps the long review status out of the action row so buttons never fight it for width", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await click(findButton(container, "生成右侧副本并排审阅"));

      const status = container.querySelector(".wps-native-review-status");
      const actions = container.querySelector(".wps-native-review-actions");
      expect(status).not.toBeNull();
      expect(actions).not.toBeNull();
      // 正向锚点：状态说明确实渲染出来了，否则下面的负向断言就是空过。
      expect(status?.textContent).toContain("尚未写入");
      expect(textOf(container)).toContain("副本回读");
      // 负向：操作区里除 button 外不许有别的元素。长说明一旦搬回按钮容器，
      // 单行 flex 就会把按钮压窄到逐字折行 —— 真机竖排的直接原因。
      expect(
        Array.from(actions!.children).every((child) => child.tagName === "BUTTON")
      ).toBe(true);
      expect(actions!.querySelector("span, p")).toBeNull();
      // 用户先读状态，再决定点哪个：状态区必须排在操作区之前。
      expect(
        (status!.compareDocumentPosition(actions!) & Node.DOCUMENT_POSITION_FOLLOWING) !== 0
      ).toBe(true);
    });

    it("names the two commit actions unambiguously and hides them before staging", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      // 尚未试算：不得出现可写回的入口。
      expect(findButton(container, "确认写回原表")).toBeNull();
      expect(findButton(container, "放弃副本")).toBeNull();

      await click(findButton(container, "生成右侧副本并排审阅"));
      const actions = container.querySelector(".wps-native-review-actions")!;
      const names = Array.from(actions.querySelectorAll("button")).map(
        (item) => (item.textContent ?? "").trim()
      );
      expect(names).toEqual(["放弃副本", "确认写回原表"]);
      // 文案要说清"写回的是原表"，而不是含糊的"保存"。
      expect(findButton(container, "确认写回原表")?.getAttribute("class")).toContain("primary");
    });

    it("disables both commit actions while the commit is in flight", async () => {
      const host = writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await click(findButton(container, "生成右侧副本并排审阅"));

      // 不用 `click()` 助手：它会 settle 掉整条异步链，捕捉不到在途状态。
      // `busy` 由 `writeProgress` 驱动，而写回第一步就同步置位。
      act(() => {
        findButton(container, "确认写回原表")!.dispatchEvent(
          new MouseEvent("click", { bubbles: true })
        );
      });
      expect(findButton(container, "放弃副本")?.disabled).toBe(true);
      expect(findButton(container, "确认写回原表")?.disabled).toBe(true);
      await settle();
      expect(host.writes()).toBeGreaterThan(0);
    });

    it("collapses the long body by default while keeping summary and actions on screen", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");

      const disclosures = Array.from(container.querySelectorAll<HTMLDetailsElement>(".wps-plan-disclosure"));
      expect(disclosures.length).toBeGreaterThanOrEqual(2);
      // 默认全部收起；收起不等于删除，内容仍在 DOM 里可展开。
      expect(disclosures.every((item) => item.open === false)).toBe(true);
      const summaries = disclosures.map((item) => item.querySelector("summary")?.textContent ?? "");
      expect(summaries.some((text) => text.includes("查看预计结果"))).toBe(true);
      expect(summaries.some((text) => text.includes("查看公式与完整计划"))).toBe(true);
      // 首屏（details 之外）必须已经能看到目标、规模与操作按钮。
      const card = container.querySelector(".wps-plan-card")!;
      expect(card.querySelector(".wps-plan-target")?.textContent).toContain("Sheet1!A1");
      expect(findButton(container, "生成右侧副本并排审阅")).not.toBeNull();
    });

    it("labels expected values and copy read-back as different things", async () => {
      writeRangeHost();
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      // 试算前只有"预计"，绝不允许出现任何"已回读"的表述。
      const beforeStaging = textOf(container);
      expect(beforeStaging).toContain("尚未生成副本");
      expect(beforeStaging).toContain("预计");
      expect(beforeStaging).not.toContain("副本回读");

      await click(findButton(container, "生成右侧副本并排审阅"));
      const afterStaging = textOf(container);
      expect(afterStaging).toContain("副本回读");
      expect(afterStaging).toContain("不是预计值");
      const summaries = Array.from(container.querySelectorAll(".wps-plan-disclosure summary"))
        .map((item) => item.textContent ?? "");
      // 预计与回读各自有独立入口，标签不同。
      expect(summaries.some((text) => text.includes("查看预计结果（尚未写入）"))).toBe(true);
      expect(summaries.some((text) => text.includes("查看副本回读结果"))).toBe(true);
    });

    it("states formula trial and fixed-value fallback with different wording", async () => {
      writeRangeHost();
      const valuesPreview = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await click(findButton(valuesPreview.container, "生成右侧副本并排审阅"));
      expect(textOf(valuesPreview.container)).toContain("副本已写入固定值并回读");
      expect(textOf(valuesPreview.container)).not.toContain("副本已试算公式");

      writeRangeHost();
      const formulaPlan: AnalysisPlan = {
        ...singleWritePlan,
        id: "wps-map-column-test",
        actions: [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=1+1"]] }],
        acceptanceCriteria: [{ type: "rangeEquals", sheet: "Sheet1", range: "A1", expected: [["2"]] }]
      };
      const formulaPreview = await mountPlanPreview(formulaPlan, "用公式计算 A1");
      await click(findButton(formulaPreview.container, "生成右侧副本并排审阅"));
      expect(textOf(formulaPreview.container)).toContain("副本已试算公式并回读计算值");
      expect(textOf(formulaPreview.container)).not.toContain("副本已写入固定值");
    });

    it("keeps the CSS contract that two buttons stay readable at narrow and wide widths", () => {
      // 再次强调：这些是**声明层面**的契约，不能替代真机视口验收。
      // 真机 350/430 px 下"按钮没有逐字折行"只能靠 WPS 截图证明。
      //
      // 正则必须用 `[^}]*`（**不许跨 `}`**）：用 `[\s\S]*?` 会一路匹配到
      // 后面别的规则里的同名声明，断言就变成恒真的假绿灯 —— 这一点用一个
      // 删掉 `white-space: nowrap` 的反向变异验过。
      expect(styles).toMatch(/\.wps-native-review-actions button\s*\{[^}]*white-space:\s*nowrap;/);
      expect(styles).toMatch(/\.wps-native-review-actions button\s*\{[^}]*min-height:\s*44px;/);
      // 窄窗格整行纵排。
      expect(styles).toMatch(/@media \(max-width:\s*559px\)\s*\{[^}]*flex-basis:\s*100%/);
      // 旧的单行 flex 容器不得复活：说明和按钮再同排就会重现真机竖排。
      expect(styles).not.toMatch(/\.wps-plan-buttons\s*\{/);
      // 长公式走等宽 + 横向滚动，不再任意折断。
      expect(styles).toMatch(
        /\.wps-plan-values th,\s*\.wps-plan-values td\s*\{[^}]*white-space:\s*nowrap;/
      );
      expect(styles).not.toMatch(
        /\.wps-plan-values th,\s*\.wps-plan-values td\s*\{[^}]*overflow-wrap:\s*anywhere/
      );
    });
  });
});

/**
 * 输入框的「上拉」：与 Excel 端共用 `useComposerResize`。
 *
 * 这里能覆盖的是"把手在哪、叫什么、高度和滚动怎么变"；jsdom 不排版，所以
 * 真实的拖动手感与窄窗格表现仍须在 WPS 真机里验收。
 */
describe("输入框上拉：与 Excel 端同一套行为", () => {  it("把手挂在输入框**外面**的 region 上，文案来自共享常量", async () => {
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");

    const handle = container.querySelector<HTMLElement>(".composer-resize-handle");
    expect(handle).not.toBeNull();
    expect(handle?.getAttribute("role")).toBe("separator");
    expect(handle?.getAttribute("aria-label")).toBe(COMPOSER_RESIZE_LABEL);
    expect(handle?.getAttribute("aria-orientation")).toBe("horizontal");
    expect(handle?.getAttribute("tabindex")).toBe("0");
    // 必须待在 region 的上内边距里：挪进 composer-box 会盖住文本区、抢走光标。
    expect(handle?.closest(".composer-input-region")).not.toBeNull();
    expect(handle?.closest(".composer-box")).toBeNull();
  });

  it("内容超过上限时改为可滚动查看，而不是把多余的行切掉", async () => {
    listModelsMock.mockResolvedValue({
      defaultModelId: firstConnection.id,
      models: [firstConnection]
    });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");
    const textarea = container.querySelector<HTMLTextAreaElement>("textarea")!;

    // jsdom 不排版：直接给一个"内容很高"的宿主输入，等价于真机里输入了很多行。
    Object.defineProperty(textarea, "scrollHeight", { configurable: true, get: () => 900 });
    await type(textarea, "需要填很长的一段说明，长到超过输入框的自动高度上限为止……");

    // 正向前提：确实超过了上限，否则下面的断言就没意义。
    expect(textarea.scrollHeight).toBeGreaterThan(COMPOSER_AUTO_GROW_MAX_HEIGHT);
    expect(textarea.style.height).toBe(`${COMPOSER_AUTO_GROW_MAX_HEIGHT}px`);
    // 核心：溢出改为可滚动，而不是 `hidden` 把文字切掉。
    expect(textarea.style.overflowY).toBe("auto");
  });

  it("CSS 契约：不再依赖 WPS WebView 里拖不动的原生 resize 把手", () => {
    const block = styles.match(/\.wps-chat-shell \.composer-box textarea\s*\{[^}]*\}/)?.[0] ?? "";
    expect(block).not.toBe("");
    expect(block).not.toContain("resize: vertical");
    // 下限与 hook 的 `minHeight: 56` 一致，拖到底和实际渲染高度才对得上。
    expect(block).toContain("min-height: 56px");
  });
});

/**
 * 任务恢复记录里"原表写入状态"的生命周期。
 *
 * 这一段最容易说假话：任务在写入途中被打断时，记录里如果没有留下任何痕迹，
 * 卡片就会对用户说"原表未写入" —— 而原表可能已经被改过了。所以判据是：
 *  - **写前**先记 `unknown`（宁可说未知，也不假装没写）；
 *  - 写成功 → `confirmed_result` + 终态（已完成的任务不再兜售"继续"）；
 *  - 部分写入/回读失败 → `needs_review`，下次只能核对；
 *  - 用户显式放弃 → 终态，且如实写"未写入"（确实一个格子都没碰）。
 */
/**
 * 从落盘的对话历史里取出所有任务恢复记录。
 *
 * 读的是 **localStorage**（真正的持久面），不是组件 state —— 这样断言的是
 * "重开之后还能看到什么"，而不是"内存里现在是什么"。
 */
function storedRecords(): Array<Record<string, unknown>> {
  const saved = JSON.parse(localStorage.getItem("excel-bro.wps.chat.v1") ?? "null") as {
    conversations?: Array<{ messages?: Array<{ taskResume?: Record<string, unknown> }> }>;
  } | null;
  return (saved?.conversations ?? [])
    .flatMap((conversation) => conversation.messages ?? [])
    .map((message) => message.taskResume)
    .filter(Boolean) as Array<Record<string, unknown>>;
}

describe("任务恢复记录：写前/写后状态与终态", () => {

  it("预览阶段就有记录：写前状态如实是「未写入」", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    expect(container.querySelector(".wps-plan-card")).not.toBeNull();

    const records = storedRecords();
    expect(records.length).toBeGreaterThan(0);
    const record = records[records.length - 1];
    expect(record.writeState).toBe("not_written");
    expect(record.phase).toBe("preview_ready");
    expect(record.originalPrompt).toBe("把 A1 改成 90");
    // 规划身份要跟着记录走：断线后靠它查真实状态，而不是重发一遍需求。
    expect(record.turnId).toBe("turn-preview-1");
  });

  it("写入成功后：记录为「已确认写入结果」且是终态，不再出现可继续入口", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card") !== null,
      "写入成功应给出结果卡"
    );

    const record = storedRecords().find((item) => item.writeState === "confirmed_result");
    expect(record).toBeDefined();
    expect(record?.resumability).toBe("unavailable");
    // 已完成的任务不能重新被摆成"可以继续"。
    expect(container.querySelector(".clarification-recovery-actions")).toBeNull();
    expect(findButton(container, "继续任务")).toBeNull();
  });

  it("部分写入失败后：记录为「需要核对原表」，不得谎报未写入", async () => {
    // 写入不落盘 → 回读与计划不符 → 执行器判定为可恢复的部分写入。
    const host = writeRangeHost({ staleWrite: true });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".message-row.system") !== null,
      "失败应给出系统反馈"
    );
    expect(host.writes()).toBeGreaterThan(0);

    const record = storedRecords().find((item) => item.phase === "executing");
    expect(record).toBeDefined();
    expect(record?.writeState).toBe("needs_review");
    expect(record?.blocker).toBe("execution_unknown");
    expect(record?.resumability).toBe("reconcile");
    // 关键负向：绝不能是"未写入"——那时原表可能已经被改过。
    expect(record?.writeState).not.toBe("not_written");
  });

  it("用户显式放弃副本：终态，且如实记录「未写入」", async () => {
    writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await click(findButton(container, "生成右侧副本并排审阅"));
    await click(findButton(container, "放弃副本"));
    await until(() => textOf(container).includes("已放弃右侧审阅副本"), "应确认已放弃");

    const record = storedRecords().find((item) => item.resumability === "unavailable");
    expect(record).toBeDefined();
    // 一个格子都没写过，这里说"未写入"是真的。
    expect(record?.writeState).toBe("not_written");
    expect(container.querySelector(".clarification-recovery-actions")).toBeNull();
  });

  it("第一次宿主写入发生的那一刻，落盘记录已经是「未知」", async () => {
    let atFirstWrite: Array<Record<string, unknown>> = [];
    writeRangeHost({
      onWrite: (count) => {
        // 在写入**发生的那一刻**取证：读的是 localStorage，不是组件 state。
        if (count === 1) atFirstWrite = storedRecords();
      }
    });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card") !== null,
      "写入应完成"
    );

    expect(atFirstWrite.length).toBeGreaterThan(0);
    const record = atFirstWrite[atFirstWrite.length - 1];
    // 顺序保证：登记必须早于第一次写入，否则写入途中断线就没人知道原表动过。
    expect(record.writeState).toBe("unknown");
    expect(record.phase).toBe("executing");
  });

  it("执行前恢复标记写不进存储时：取消写入，原表写入次数为 0", async () => {
    const host = writeRangeHost();
    const original = Storage.prototype.setItem;
    const spy = vi
      .spyOn(Storage.prototype, "setItem")
      .mockImplementation(function (this: Storage, key: string, value: string) {
        if (key === "excel-bro.wps.chat.v1") throw new Error("QuotaExceededError");
        return original.call(this, key, value);
      });
    try {
      const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
      await click(findButton(container, "生成右侧副本并排审阅"));
      await click(findButton(container, "确认写回原表"));
      await until(
        () => textOf(container).includes("已取消本次写入"),
        "存不下标记时应明确取消写入"
      );

      // 核心：没有持久证据就**一个单元格都不写**。
      expect(host.writes()).toBe(0);
      expect(host.cellValue()).toBe("旧值");
    } finally {
      spy.mockRestore();
    }
  });

  it("写入之后恢复点读取失败：不得显示「原表未写入」", async () => {
    const host = writeRangeHost({ staleWrite: true, breakStorageReadsAfterWrite: true });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".message-row.system") !== null,
      "失败应给出系统反馈"
    );
    expect(host.writes()).toBeGreaterThan(0);

    const record = storedRecords().find((item) => item.phase === "executing");
    expect(record).toBeDefined();
    /*
     * 读取失败 ≠ 没写过。这里必须保持"未知"——把它折成 `not_written`，
     * 就是对着可能已被改过的原表说"原表未写入"。
     */
    expect(record?.writeState).toBe("unknown");
    expect(record?.writeState).not.toBe("not_written");
    expect(textOf(container)).not.toContain("原表未写入");
    expect(textOf(container)).toContain("无法确认");
  });
});

describe("WPS 写入耐久恢复（CAP-03）", () => {
  it("服务端登记失败：取消写入，原表写入次数为 0", async () => {
    const host = writeRangeHost();
    registerWpsExecutionMock.mockRejectedValueOnce(new Error("服务端不可用"));
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => textOf(container).includes("无法持久登记本次写入现场"),
      "登记失败应中止写入"
    );

    // 核心：拿不到持久证据就一个单元格都不写。
    expect(host.writes()).toBe(0);
    expect(host.cellValue()).toBe("旧值");
  });

  it("服务端未授予写入权（无 ownerToken）：取消写入", async () => {
    const host = writeRangeHost();
    registerWpsExecutionMock.mockResolvedValueOnce({
      operationId: "wps-execution-test",
      status: "running",
      revision: 1,
      workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
      taskId: "task-test",
      descriptor: { actions: [] },
      authorizedSheetNames: ["Sheet1"],
      nextActionIndex: 0,
      totalActions: 1,
      inFlightActionIndex: null,
      beforeSnapshot: null,
      checkpoint: null,
      result: null,
      error: null,
      canResume: true,
      ownerActive: false,
      requiresReconciliation: false,
      ownerToken: null
    });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => textOf(container).includes("没有授予本次 WPS 写入权"),
      "无写入权应中止"
    );

    expect(host.writes()).toBe(0);
  });

  it("写成功并提交动作级检查点与完成回执", async () => {
    const host = writeRangeHost();
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card") !== null,
      "写入应完成"
    );

    expect(host.writes()).toBeGreaterThan(0);
    expect(startWpsActionMock).toHaveBeenCalled();
    expect(completeWpsActionMock).toHaveBeenCalled();
    expect(finishWpsExecutionMock).toHaveBeenCalled();
  });

  it("本地发现索引只存白名单字段，不含计划、授权或矩阵", async () => {
    const host = writeRangeHost();
    registerWpsExecutionMock.mockResolvedValueOnce({
      operationId: "wps-execution-test",
      status: "running",
      revision: 1,
      workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
      taskId: "task-test",
      descriptor: { actions: [] },
      authorizedSheetNames: ["Sheet1"],
      nextActionIndex: 0,
      totalActions: 1,
      inFlightActionIndex: null,
      beforeSnapshot: null,
      checkpoint: null,
      result: null,
      error: null,
      canResume: true,
      ownerActive: false,
      requiresReconciliation: false,
      ownerToken: "owner-token-0123456789abcdef0123456789abcdef"
    });
    const { container } = await mountPlanPreview(singleWritePlan, "把 A1 改成 90");
    await stageAndSave(container);
    await until(
      () => container.querySelector(".wps-verification-card") !== null,
      "写入应完成"
    );

    const raw = localStorage.getItem("excel-bro.pending-wps-execution.v1");
    // 写成功后索引已被清除（完成回执释放）。
    expect(raw).toBeNull();
    void host;
  });

  /*
   * 重开面板后的恢复卡：**点真实按钮、看真实 API 调用与宿主写入次数**。
   * 不只看徽标文案 —— 上一轮的缺陷正是"显示可以继续却没有按钮"。
   */
  const PENDING_KEY = "excel-bro.pending-wps-execution.v1";

  function recoveryState(overrides: Record<string, unknown> = {}) {
    const total = Number(overrides.totalActions ?? 1);
    const fingerprint = matrixFingerprintV2([["旧值"]], [["旧值"]]);
    return {
      operationId: "wps-execution-test",
      status: "running" as const,
      revision: 7,
      workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
      taskId: "task-test",
      descriptor: { actions: Array.from({ length: total }, (_, actionIndex) => ({
        actionIndex, type: "writeValues", sheet: "Sheet1", range: `A${actionIndex + 1}`,
        rowCount: 1, columnCount: 1, cellCount: 1,
        previewFingerprint: fingerprint, expectedFingerprint: fingerprint
      })) },
      authorizedSheetNames: ["Sheet1"],
      nextActionIndex: 0,
      totalActions: 1,
      inFlightActionIndex: null,
      beforeSnapshot: null,
      checkpoint: null,
      result: null,
      error: null,
      canResume: true,
      ownerActive: false,
      requiresReconciliation: false,
      ownerToken: null,
      ...overrides
    };
  }

  /** 带 `taskResume` 的历史消息：重新预览要靠它找回原始需求。 */
  function historyWithTaskResume(): unknown[] {
    return [
      {
        id: "msg-resume",
        role: "assistant",
        text: "写入未完成",
        createdAt: new Date().toISOString(),
        taskResume: {
          schemaVersion: 1,
          taskId: "task-test",
          phase: "executing",
          blocker: "execution_unknown",
          resumability: "reconcile",
          writeState: "unknown",
          originalPrompt: "把 A1 改成 90",
          confirmations: [],
          authorizationRevision: 1,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        }
      }
    ];
  }

  async function mountRecoveryCard(options: {
    state?: Record<string, unknown>;
    fullName?: string;
    withHistory?: boolean;
    /** 在挂载**之前**塞一条现场冲突的恢复点，让只读核对拿到 conflict。 */
    seedUndo?: "conflict";
  } = {}) {
    const host = writeRangeHost();
    if (options.seedUndo === "conflict") {
      seedConflictingUndoRecord("wps-execution-test");
      host.setCell("人工修改");
    }
    if (options.withHistory) {
      localStorage.setItem(
        "excel-bro.wps.chat.v1",
        JSON.stringify({
          activeConversationId: "conv-1",
          conversations: [
            {
              id: "conv-1",
              title: "历史",
              messages: historyWithTaskResume(),
              createdAt: new Date().toISOString(),
              updatedAt: new Date().toISOString()
            }
          ]
        })
      );
    }
    if (options.fullName !== undefined) {
      const application = getWpsApplicationMock() as unknown as {
        ActiveWorkbook: { FullName: string };
      };
      application.ActiveWorkbook.FullName = options.fullName;
    }
    localStorage.setItem(
      PENDING_KEY,
      JSON.stringify({
        schemaVersion: 1,
        operationId: "wps-execution-test",
        conversationId: "conv-1",
        activityMessageId: "msg-1",
        workbookIdentity: "wps-path:c:\\data\\测试.xlsx",
        taskId: "task-test",
        createdAt: new Date().toISOString()
      })
    );
    getWpsExecutionStatusMock.mockResolvedValue(recoveryState(options.state) as never);
    listModelsMock.mockResolvedValue({ defaultModelId: firstConnection.id, models: [firstConnection] });
    const { container } = await mount(panel());
    await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型目录");
    await until(
      () => container.querySelector(".wps-write-recovery") !== null,
      "重载后应出现恢复核对卡"
    );
    await until(
      () => !textOf(container).includes("正在查询服务端写入现场"),
      "应查完服务端状态"
    );
    return { container, host };
  }

  function recoveryButtons(container: HTMLElement): string[] {
    const card = container.querySelector(".wps-write-recovery");
    if (!card) return [];
    // Only mutation/dismiss actions; the separate refresh button is read-only.
    return Array.from(card.querySelectorAll(".wps-recovery-action")).map((item) =>
      (item.textContent ?? "").trim()
    );
  }

  it("零写入证据的锚点：纯测试替身计数包含 Value2 和 Formula 两种 setter", () => {
    const host = writeRangeHost();
    // getWpsApplicationMock was explicitly replaced by writeRangeHost; this is
    // only a JS fixture, never the real WPS application or a runtime write path.
    const fixture = getWpsApplicationMock();
    const range = fixture.ActiveWorkbook!.Worksheets.Item(1).Range!("A1");
    expect(host.writes()).toBe(0);
    range.Value2 = "TEST_VALUE";
    expect(host.writes()).toBe(1);
    range.Formula = "=1+1";
    expect(host.writes()).toBe(2);
  });

  it("全部动作已完成、只丢最终回执：点补记完成 → 领取执行权 + 提交完成回执，宿主写入 0", async () => {
    const { container, host } = await mountRecoveryCard({
      state: { nextActionIndex: 2, totalActions: 2 }
    });
    claimWpsExecutionResumeMock.mockResolvedValueOnce(recoveryState({ nextActionIndex: 2,
      totalActions: 2, ownerToken: "claimed-owner", revision: 8 }) as never);

    expect(recoveryButtons(container)).toEqual([
      "补记完成（不重写原表）",
      "清除这条恢复提示"
    ]);

    await click(findButton(container, "补记完成（不重写原表）"));
    await until(
      () => textOf(container).includes("已补记这次写入的完成回执"),
      "补记应成功"
    );

    expect(claimWpsExecutionResumeMock).toHaveBeenCalledWith("wps-execution-test", 7);
    expect(finishWpsExecutionMock).toHaveBeenCalledTimes(1);
    expect(finishWpsExecutionMock).toHaveBeenCalledWith("wps-execution-test", {
      expectedRevision: 8, ownerToken: "claimed-owner", result: { backfilled: true, actionCount: 2 }
    });
    // 补记只改账本：一次宿主写入都没有。
    expect(host.writes()).toBe(0);
    expect(localStorage.getItem(PENDING_KEY)).toBeNull();
    expect(container.querySelector(".wps-write-recovery")).toBeNull();
  });

  it("还有动作没写：只给「重新读取并生成预览」，不给任何继续/写入按钮", async () => {
    const { container } = await mountRecoveryCard({
      state: { nextActionIndex: 1, totalActions: 3 }
    });

    const buttons = recoveryButtons(container);
    expect(buttons).toEqual(["重新读取并生成预览", "清除这条恢复提示"]);
    // 关键：不许出现"可以继续"这类冒充断点续跑的入口。
    expect(buttons.some((name) => name.includes("继续"))).toBe(false);
    expect(textOf(container)).toContain("无法精确续跑剩余动作");
  });

  it("点「重新读取并生成预览」→ 用原需求重新发起一轮，不是续跑", async () => {
    const { container, host } = await mountRecoveryCard({
      state: { nextActionIndex: 1, totalActions: 3 },
      withHistory: true
    });
    runTurnMock.mockResolvedValue({ kind: "plan", plan: singleWritePlan, turnId: "turn-replan" });
    expect(findButton(container, "重新读取并生成预览")?.disabled).toBe(true);
    await click(container.querySelector('.wps-write-recovery input[type="checkbox"]') as HTMLInputElement);
    await click(findButton(container, "重新读取并生成预览"));
    await until(
      () => textOf(container).includes("重新读取并生成预览 —— 那是新的一轮"),
      "应说明这是新的一轮"
    );

    // 重新走一轮规划，且**没有**对原表做任何写入。
    expect(runTurnMock).toHaveBeenCalled();
    expect(host.writes()).toBe(0);
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(claimWpsExecutionResumeMock).toHaveBeenCalled();
    expect(suspendWpsExecutionMock).toHaveBeenCalled();
    expect(localStorage.getItem(PENDING_KEY)).not.toBeNull();
  });

  it.each([
    ["canResume 变 false", { canResume: false }],
    ["另一个执行者领取", { ownerActive: true }],
    ["剩余动作重新出现", { nextActionIndex: 0 }]
  ])("点补记前重新查询：%s 不调用领取或完成", async (_label, patch) => {
    const { container, host } = await mountRecoveryCard({ state: { nextActionIndex: 1 } });
    getWpsExecutionStatusMock.mockResolvedValue(recoveryState({ nextActionIndex: 1, ...patch }) as never);
    await click(findButton(container, "补记完成（不重写原表）"));
    await until(() => getWpsExecutionStatusMock.mock.calls.length >= 2, "必须重新查询");
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(host.writes()).toBe(0);
  });

  it("卡片显示后人工改值：点补记重新读现场，拦住回执", async () => {
    const { container, host } = await mountRecoveryCard({ state: { nextActionIndex: 1 } });
    host.setCell("人工修改后");
    await click(findButton(container, "补记完成（不重写原表）"));
    await until(() => textOf(container).includes("疑似人工修改"), "应识别现场冲突");
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(host.writes()).toBe(0);
  });

  it("领取期间人工改值：第二次现场核对仍拦住完成回执", async () => {
    const { container, host } = await mountRecoveryCard({ state: { nextActionIndex: 1 } });
    claimWpsExecutionResumeMock.mockImplementationOnce(async () => {
      host.setCell("领取期间修改");
      return recoveryState({ nextActionIndex: 1, ownerToken: "new-owner", revision: 8 }) as never;
    });
    await click(findButton(container, "补记完成（不重写原表）"));
    await until(() => textOf(container).includes("领取执行权后现场发生变化"), "第二次核对应拦截");
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(suspendWpsExecutionMock).toHaveBeenCalled();
    expect(host.writes()).toBe(0);
  });

  it("最后动作回执丢失：点击核对提交指纹后只补记完成，零宿主写入", async () => {
    const { container, host } = await mountRecoveryCard({ state: {
      status: "needs_reconciliation", requiresReconciliation: true, inFlightActionIndex: 0
    }});
    reconcileWpsActionReceiptMock.mockResolvedValueOnce(recoveryState({ nextActionIndex: 1,
      ownerToken: "new-owner", revision: 8 }) as never);
    await click(findButton(container, "核对在途动作"));
    await until(() => textOf(container).includes("已核对最后一个动作并补记完成"), "应只补记");
    expect(reconcileWpsActionReceiptMock).toHaveBeenCalledWith("wps-execution-test", expect.objectContaining({
      outcome: "applied", checkpoint: expect.objectContaining({ type: "writeValues",
        sheetName: "Sheet1", range: "A1", afterFingerprint: expect.any(String) })
    }));
    expect(finishWpsExecutionMock).toHaveBeenCalledTimes(1);
    expect(runTurnMock).not.toHaveBeenCalled();
    expect(host.writes()).toBe(0);
  });

  it("canResume=false：只给清除提示，不给补记或核对按钮", async () => {
    const { container } = await mountRecoveryCard({
      state: { canResume: false, nextActionIndex: 2, totalActions: 2 }
    });
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
  });

  it("执行权仍活跃时仅只读刷新，释放后才出现补记入口", async () => {
    const { container, host } = await mountRecoveryCard({ state: { nextActionIndex: 1, ownerActive: true } });
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
    getWpsExecutionStatusMock.mockResolvedValue(recoveryState({ nextActionIndex: 1 }) as never);
    await click(findButton(container, "重新核对（只读）"));
    await until(() => recoveryButtons(container).includes("补记完成（不重写原表）"), "刷新后才开放补记");
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
    expect(host.writes()).toBe(0);
  });

  it("身份不一致（另一个工作簿）：只给清除提示，不给写入按钮", async () => {
    const { container } = await mountRecoveryCard({
      state: { nextActionIndex: 2, totalActions: 2 },
      fullName: "C:\\other\\测试.xlsx"
    });
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
    expect(textOf(container)).toContain("不是同一个文件");
  });

  /**
   * 往 PluginStorage 塞一条**现场与写前、写后都不一致**的恢复点。
   * 指纹故意用不可能匹配上的值，让只读核对必然得出 `conflict`。
   */
  function seedConflictingUndoRecord(executionId: string) {
    const application = getWpsApplicationMock() as unknown as {
      PluginStorage: { setItem: (key: string, value: unknown) => void };
    };
    application.PluginStorage.setItem(
      "excel-bro.wps.undo.v1",
      JSON.stringify([
        {
          version: 2,
          executionId,
          planId: "plan-old",
          planTitle: "旧计划",
          createdAt: new Date().toISOString(),
          workbookIdentity: "c:\\data\\测试.xlsx",
          authorizedSheetNames: ["Sheet1"],
          state: "applying",
          actions: [
            {
              actionIndex: 0,
              type: "writeValues",
              sheetIndex: 1,
              sheetName: "Sheet1",
              range: "A1",
              rowCount: 1,
              columnCount: 1,
              cellCount: 1,
              values: [["90"]],
              preview: [["90"]],
              previewFingerprint: "1:00000000",
              rect: { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 },
              beforeValues: [["旧值"]],
              beforeFormulas: [["旧值"]],
              beforeFingerprint: "1:aaaaaaaa",
              afterFingerprint: "1:bbbbbbbb",
              appliedRows: 1
            }
          ],
          error: null
        }
      ])
    );
  }

  it("现场冲突（疑似人工修改）：不给补记/核对按钮，也不提交任何回执", async () => {
    const { container } = await mountRecoveryCard({
      state: {
        status: "needs_reconciliation",
        requiresReconciliation: true,
        inFlightActionIndex: 0
      },
      seedUndo: "conflict"
    });
    expect(textOf(container)).toContain("疑似人工修改或部分写入");
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);

    await click(findButton(container, "清除这条恢复提示"));
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(claimWpsExecutionResumeMock).not.toHaveBeenCalled();
    expect(finishWpsExecutionMock).not.toHaveBeenCalled();
  });

  it("全部动作已完成但本机现场是写前（两者矛盾）：不给补记按钮", async () => {
    const { container } = await mountRecoveryCard({
      state: { nextActionIndex: 2, totalActions: 2 },
      seedUndo: "conflict"
    });
    // 服务端说写完了、本机却与写后不一致 → 矛盾，挂起等人工核对。
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
  });

  it("旧描述缺少写后指纹：不给核对按钮，也不提交回执", async () => {
    const { container } = await mountRecoveryCard({
      state: {
        status: "needs_reconciliation",
        requiresReconciliation: true,
        inFlightActionIndex: 0,
        descriptor: { actions: [{ actionIndex: 0, type: "writeValues", sheet: "Sheet1",
          range: "A1", rowCount: 1, columnCount: 1, cellCount: 1, previewFingerprint: "legacy" }] }
      }
    });
    expect(recoveryButtons(container)).toEqual(["清除这条恢复提示"]);
    await click(findButton(container, "清除这条恢复提示"));
    expect(reconcileWpsActionReceiptMock).not.toHaveBeenCalled();
    expect(container.querySelector(".wps-write-recovery")).toBeNull();
  });
});
