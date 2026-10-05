/**
 * 定向测试：WPS 数据范围面板的**草稿 / 已提交**分离。
 *
 * 之前 `WpsProbePanel` 里勾选就是提交：`toggleSheet` 立刻改 `selectedIndexes`，
 * 而面板把它当授权范围传给 `WpsAgentPanel`。于是用户只要"顺手点一下另一张表
 * 看看"，就会触发面板里按授权 key 的失效副作用，把正在等回答的任务打掉 ——
 * 他并没有改变任何授权意图，只是想看一眼。
 *
 * 这里钉住三条：
 *  1. 勾选只改草稿；取消勾选再选回，已提交范围一字未动。
 *  2. 「确认范围」才提交，且**提交时重新读取真实宿主元数据**（不是拿开弹层时的快照）。
 *  3. 「放弃修改」与关闭不改变已提交授权。
 */
import { act, type ReactElement, type ReactNode } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import WpsProbePanel from "./WpsProbePanel";
import { getWpsApplication } from "./wpsHost";
import { listWpsWorksheets, type WpsWorkbookMetadata } from "./wpsWorkbookReader";

vi.mock("./ribbon", () => ({ initializeRestoredTaskPaneWidth: () => undefined }));
vi.mock("./devProbe", () => ({ reportDevProbe: () => undefined }));
vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});
vi.mock("./wpsWorkbookReader", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsWorkbookReader")>();
  return { ...actual, listWpsWorksheets: vi.fn() };
});
/**
 * 面板本体换成只暴露"它收到的已提交范围"的替身。
 *
 * 这样断言落在**授权的提交面**上，而不是被面板内部的渲染细节干扰 ——
 * 我们要证的是"勾选有没有提前提交授权"，不是面板长什么样。
 */
vi.mock("./WpsAgentPanel", () => ({
  default: ({
    selectedSheetIndexes,
    scopeDock
  }: {
    selectedSheetIndexes: readonly number[];
    scopeDock?: ReactNode;
  }) => (
    <div>
      <div className="panel-committed">{selectedSheetIndexes.join(",")}</div>
      {scopeDock}
    </div>
  )
}));

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const getWpsApplicationMock = vi.mocked(getWpsApplication);
const listWpsWorksheetsMock = vi.mocked(listWpsWorksheets);

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

async function click(target: HTMLElement | null) {
  if (!target) throw new Error("要点击的元素不存在");
  await act_(() => {
    target.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
}

function byText(container: HTMLElement, selector: string, text: string): HTMLButtonElement {
  const found = Array.from(container.querySelectorAll<HTMLButtonElement>(selector)).find(
    (button) => (button.textContent ?? "").includes(text)
  );
  if (!found) throw new Error(`找不到「${text}」`);
  return found;
}

function sheetButton(container: HTMLElement, name: string): HTMLButtonElement {
  return byText(container, ".wps-sheet-picker-options button", name);
}

function committed(container: HTMLElement): string {
  return container.querySelector(".panel-committed")?.textContent ?? "";
}

/** 两张表的最小宿主替身（只有元数据走 mock，探测仍读真实宿主对象）。 */
function hostApplication() {
  const makeSheet = (name: string, position: number) => ({
    Name: name,
    get Index() {
      return position + 1;
    }
  });
  const sheets = [makeSheet("订单", 0), makeSheet("明细", 1)];
  return {
    Version: "12.0",
    get ActiveWorkbook() {
      return {
        Name: "测试.xlsx",
        FullName: "C:\\数据\\测试.xlsx",
        Worksheets: {
          Count: sheets.length,
          Item: (index: number) => sheets[Number(index) - 1]
        },
        ActiveSheet: sheets[0]
      };
    },
    get ActiveSheet() {
      return sheets[0];
    },
    Selection: undefined,
    PluginStorage: { getItem: () => undefined, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never;
}

function metadataWith(sheets: Array<{ index: number; name: string; isActive?: boolean }>): WpsWorkbookMetadata {
  return {
    name: "测试.xlsx",
    path: "C:\\数据\\测试.xlsx",
    fullName: "C:\\数据\\测试.xlsx",
    sheets: sheets.map((sheet) => ({ ...sheet, isActive: sheet.isActive ?? false }))
  } as never;
}

beforeEach(() => {
  vi.clearAllMocks();
  getWpsApplicationMock.mockReturnValue(hostApplication());
  listWpsWorksheetsMock.mockReturnValue(
    metadataWith([
      { index: 1, name: "订单", isActive: true },
      { index: 2, name: "明细" }
    ])
  );
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
});

describe("数据范围：勾选只改草稿，确认才提交授权", () => {
  it("取消勾选再选回，已提交范围一字未动（净变化为零时不算草稿）", async () => {
    const { container } = await mount(<WpsProbePanel />);
    expect(committed(container)).toBe("1");

    await click(byText(container, ".scope-trigger", ""));
    // 切到「选择多个」才能手动勾表。
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    // 勾选只是草稿：授权范围没有提前改变。
    expect(committed(container)).toBe("1");
    expect(container.querySelector(".scope-confirm")).not.toBeNull();

    // 再取消选回 —— 勾选集合回到已提交范围。
    await click(sheetButton(container, "明细"));
    expect(sheetButton(container, "明细").className).not.toContain("selected");
    // 全程已提交范围一字未动：用户只是"看一眼另一张表"，任务不该因此被销毁。
    expect(committed(container)).toBe("1");
  });

  it("确认范围才提交，且提交时重新读取真实宿主元数据", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    expect(committed(container)).toBe("1");

    // 弹层开着的时候，宿主那边发生了变化：索引 2 已经不存在了。
    listWpsWorksheetsMock.mockReturnValue(
      metadataWith([{ index: 1, name: "订单", isActive: true }])
    );

    await click(byText(container, ".wps-scope-actions button", "确认范围"));
    // 用**当前真实元数据**解析草稿：已经消失的索引不会被提交成授权。
    expect(committed(container)).toBe("1");
    expect(container.querySelector(".scope-popover")).toBeNull();
  });

  it("确认范围提交新增的工作表，并且只在确认之后才生效", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    await click(byText(container, ".wps-scope-actions button", "确认范围"));
    expect(committed(container)).toBe("1,2");
  });

  it("放弃修改不改变已提交授权，重开弹层时草稿回到已提交范围", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    await click(byText(container, ".wps-scope-actions button", "放弃修改"));
    expect(committed(container)).toBe("1");

    // 重开：草稿必须是已提交范围（明细未被选中）。
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    expect(sheetButton(container, "明细").className).not.toContain("selected");
  });

  it("手动模式下没有任何勾选时不能确认范围", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(byText(container, ".sheet-picker-toolbar button", "清空"));
    const confirm = byText(container, ".wps-scope-actions button", "确认范围");
    expect(confirm.disabled).toBe(true);
    // 已提交范围不受草稿清空影响。
    expect(committed(container)).toBe("1");
  });
});
