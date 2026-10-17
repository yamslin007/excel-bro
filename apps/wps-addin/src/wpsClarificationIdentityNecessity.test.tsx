/**
 * 定向测试：**身份判据的必要性**。
 *
 * 上一轮暴露的问题：只禁用身份判据、保留表头判据，204 条测试仍全绿。
 * 根因是「表头指纹」当时并不是真正的表头 —— 面板把
 * `listWpsWorksheets().sheets.map((s) => [s.index, s.name])` 喂给了
 * `wpsHeaderFingerprint`，即**工作表元数据**。于是两条判据吃的是同一份数据，
 * 表头判据只是身份判据的副本，禁掉身份判据自然看不出差别。
 *
 * 这里的替身**刻意让表头、行列尺寸、数据值全都保持一致**，只有身份不同：
 *  - 同一张表改名（表头内容不变）
 *  - 重排（每张表的表头与尺寸都不变）
 *  - 同索引替换成另一张表，但新表**表头与尺寸完全相同**（内容上无从区分）
 *  - 换一个工作簿（表头与尺寸相同）
 *
 * 这样"表头判据"无论如何都无法触发，只有**身份判据**能拦住 —— 用例必须变红，
 * 否则说明身份判据没有被咬住。
 */
import { readFileSync } from "node:fs";
import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { listModels, getModelSettings, validatePlan } from "../../excel-addin/src/api";
import WpsAgentPanel from "./WpsAgentPanel";
import { runWpsReadonlyTurn } from "./wpsReadonlyAgent";
import { getWpsApplication } from "./wpsHost";

// Isolate the pre-existing stronger clarification proof from the new early metadata guard.
// Full real-guard integration is exercised in WpsAgentPanel.test.tsx.
vi.mock("../../excel-addin/src/hooks/useAuthorizedTarget", () => ({
  useAuthorizedTarget: () => ({ check: async () => true, checkForAnalysis: async () => true,
    clearAnalysisConsent: () => {}, hasAnalysisConsent: () => false,
    requestConfirmation: async () => false, notice: null, checking: false, binding: null })
}));

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  return { ...actual, understandMessage: vi.fn(async () => ({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" })), listModels: vi.fn(), getModelSettings: vi.fn(), validatePlan: vi.fn() };
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
readFileSync("src/styles.css", "utf8");

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

/**
 * 一台**表头与尺寸在所有工作表上完全一致**的替身。
 *
 * 每张表的已用区域都是 3 行 × 2 列，第 1 行都是 `["类别", "归属"]`。
 * 唯一能区分"这是不是原来那张表"的信息，只有工作表名称与顺序 ——
 * 也就是**授权身份**。表头判据在这里天然失效。
 */
function uniformHost(options: {
  workbookName?: string;
  sheetNames: string[];
}) {
  const workbookName = options.workbookName ?? "测试.xlsx";
  const sheets = options.sheetNames.map((name) => ({
    Name: name,
    ProtectContents: false
  }));

  /** 表头行内容 —— 可被 `setHeader` 改写，用来单独驱动**表头判据**。 */
  let headerRow: string[][] = [["类别", "归属"]];

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
      },
      Item: {
        // `Item(row, col)` 返回一个"单格区域"替身，供矩形端点使用。
        value: () => makeRange()
      }
    });
    return range;
  };

  /**
   * 真宿主语义：`sheet.Range(first, last)`（只接受两个参数，多传会抛
   * `Too many parameters.`，见 2026-09-30 真机证据）返回由两个端点框定的区域。
   * 结构扫描传的是 `Item(1,1)..Item(1,columns)`，所以这里只需返回**第 1 行**的
   * 单元格值 —— 数据行不进这次读取，与生产侧"结构扫描只取表头"一致。
   */
  const rangeFrom = () => {
    const range: Record<string, unknown> = {
      Rows: { Count: headerRow.length },
      Columns: { Count: headerRow[0]?.length ?? 0 },
      MergeCells: false
    };
    Object.defineProperties(range, {
      Value2: { get: () => headerRow.map((row) => [...row]) }
    });
    return range;
  };

  sheets.forEach((sheet, position) => {
    Object.defineProperties(sheet, {
      UsedRange: { get: makeRange },
      Index: { get: () => position + 1 },
      Range: { value: () => rangeFrom() }
    });
  });

  const storage = new Map<string, unknown>();
  const application = {
    Version: "12.0",
    ActiveWorkbook: {
      Name: workbookName,
      FullName: `C:\\Data\\${workbookName}`,
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
      setItem: (key: string, item: unknown) => {
        storage.set(key, item);
      }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never;

  return {
    application,
    sheets,
    rename: (position: number, name: string) => {
      sheets[position].Name = name;
    },
    reorder: (next: string[]) => {
      const byName = new Map(sheets.map((sheet) => [sheet.Name, sheet]));
      next.forEach((name, position) => {
        const sheet = byName.get(name);
        if (sheet) sheets[position] = sheet;
      });
    },
    setWorkbookName: (name: string) => {
      (application as { ActiveWorkbook: { Name: string } }).ActiveWorkbook.Name = name;
    },
    /** 只改表头内容：身份（表名/顺序/工作簿）一个字不动。 */
    setHeader: (next: string[][]) => {
      headerRow = next;
    }
  };
}

function clarificationCard(headerRows: string[][] = [["类别", "归属"]]) {
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
    // `runWpsTurn` 结构扫描真实读到的表头 —— 面板据此做表头指纹。
    headerRows
  } as never);
}

async function pendingClarification() {
  const { container } = await mount(panel());
  await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
  const composer = container.querySelector<HTMLTextAreaElement>(".composer textarea");
  await type(composer, "帮我处理表格");
  await submit(composer);
  await until(() => Boolean(container.querySelector(".clarification-card")), "应显示澄清卡");
  await until(() => runTurnMock.mock.calls.length === 1, "应完成这一轮");
  return container;
}

function clarificationOptions(container: HTMLElement): HTMLButtonElement[] {
  return Array.from(
    container.querySelectorAll<HTMLButtonElement>(".clarification-options button")
  );
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

describe("身份判据的必要性（表头/尺寸/数据值全部相同的替身）", () => {
  it("已授权表改名：表头与尺寸完全不变，只有身份变化 —— 必须失效", async () => {
    clarificationCard();
    const host = uniformHost({ sheetNames: ["订单", "明细"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    expect(clarificationOptions(container)).toHaveLength(1);
    // 只有名称变：表头仍是 ["类别","归属"]，仍是 3 行 × 2 列，数据行一字未动。
    host.rename(0, "订单（新）");

    await click(clarificationOptions(container)[0]);
    await until(
      () => container.querySelector(".clarification-card.invalidated") !== null,
      "改名后待答澄清必须失效（表头判据在这里天然触发不了）"
    );
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    expect(textOf(container)).toContain("数据范围已经变化");
    expect(
      container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value
    ).toContain("帮我处理表格");
  });

  it("工作表重排：每张表的表头与尺寸都不变 —— 必须失效", async () => {
    clarificationCard();
    const host = uniformHost({ sheetNames: ["订单", "明细"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    host.reorder(["明细", "订单"]);

    await click(clarificationOptions(container)[0]);
    await until(
      () => container.querySelector(".clarification-card.invalidated") !== null,
      "重排后待答澄清必须失效"
    );
    expect(runTurnMock).toHaveBeenCalledTimes(1);
  });

  it("同索引换成另一张表，且新表表头/尺寸/数据完全相同 —— 必须失效", async () => {
    clarificationCard();
    // 两张表内容完全一样，只有名字不同：表头判据无从分辨。
    const host = uniformHost({ sheetNames: ["订单", "订单（副本）"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    // 索引 1 指向的表被替换成同一份结构的另一张表。
    host.sheets[0].Name = "订单（副本归入）";
    host.sheets[1].Name = "订单";

    await click(clarificationOptions(container)[0]);
    await until(
      () => container.querySelector(".clarification-card.invalidated") !== null,
      "同索引换表后待答澄清必须失效"
    );
    expect(runTurnMock).toHaveBeenCalledTimes(1);
  });

  it("换一本工作簿，表头与尺寸不变 —— 必须失效", async () => {
    clarificationCard();
    const host = uniformHost({ sheetNames: ["订单", "明细"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    host.setWorkbookName("另一本.xlsx");

    await click(clarificationOptions(container)[0]);
    await until(
      () => container.querySelector(".clarification-card.invalidated") !== null,
      "换工作簿后待答澄清必须失效"
    );
    expect(runTurnMock).toHaveBeenCalledTimes(1);
  });

  it("身份与表头都没变：澄清可以正常继续（防误伤）", async () => {
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
      } as never)
      .mockResolvedValueOnce({
        kind: "answer",
        message: "已按补充说明处理。",
        resultContext: null,
        provider: "model"
      } as never);
    const host = uniformHost({ sheetNames: ["订单", "明细"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    await click(clarificationOptions(container)[0]);
    await until(() => runTurnMock.mock.calls.length === 2, "身份未变时应继续原任务");
    expect(container.querySelector(".clarification-card.resolved")).not.toBeNull();
    expect(textOf(container)).not.toContain("数据范围已经变化");
  });

  it("只有表头被改名（表名/顺序/工作簿都不变）：表头判据必须独立生效", async () => {
    clarificationCard();
    const host = uniformHost({ sheetNames: ["订单", "明细"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    // 身份一个字没动，只有第 1 行的列名被改写。
    host.setHeader([["类型", "归属"]]);

    await click(clarificationOptions(container)[0]);
    await until(
      () => container.querySelector(".clarification-card.invalidated") !== null,
      "表头改写后待答澄清必须失效（这条只有表头判据能拦）"
    );
    expect(runTurnMock).toHaveBeenCalledTimes(1);
  });
});
