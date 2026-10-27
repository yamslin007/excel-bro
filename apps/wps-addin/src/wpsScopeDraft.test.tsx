import { act, useEffect, type ReactElement, type ReactNode } from "react";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { SCOPE_ADMISSION_MESSAGES } from "@excel-bro/core/scopeAdmission";
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
 * 我们要证的是"勾选是否立即提交授权"，不是面板长什么样。
 */
vi.mock("./WpsAgentPanel", () => ({
  default: function WpsScopeHarness({
    selectedSheetIndexes,
    scopeDock,
    scopeOverlay,
    onScopeConfirmationReady
  }: {
    selectedSheetIndexes: readonly number[];
    scopeDock?: ReactNode;
    scopeOverlay?: ReactNode;
    onScopeConfirmationReady?: (handler: ((target: AuthorizedTarget) => Promise<boolean>) | null) => void;
  }) {
    useEffect(() => { onScopeConfirmationReady?.(scopeAuthorization); return () => onScopeConfirmationReady?.(null); }, [onScopeConfirmationReady]);
    return (
    <div>
      <div className="panel-committed">{selectedSheetIndexes.join(",")}</div>
      {scopeDock}
      {/* 面板本体住在独立呈现层里；替身要把它也渲染出来，否则判据查不到它。 */}
      {scopeOverlay ? <div className="scope-layer">{scopeOverlay}</div> : null}
    </div>
    );
  }
}));

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const getWpsApplicationMock = vi.mocked(getWpsApplication);
const listWpsWorksheetsMock = vi.mocked(listWpsWorksheets);
const scopeAuthorization = vi.fn(async (_target: AuthorizedTarget) => true);

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
function hostApplication(activeIndex = 1) {
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
        ActiveSheet: sheets[activeIndex - 1]
      };
    },
    get ActiveSheet() {
      return sheets[activeIndex - 1];
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
  localStorage.clear();
  scopeAuthorization.mockResolvedValue(true);
  getWpsApplicationMock.mockReturnValue(hostApplication());
  listWpsWorksheetsMock.mockReturnValue(
    metadataWith([
      { index: 1, name: "订单", isActive: true },
      { index: 2, name: "明细" }
    ])
  );
});

afterEach(() => {
  act(() => { mounted.forEach((entry) => entry.unmount()); });
  mounted = [];
  document.body.innerHTML = "";
});

describe("数据范围：勾选即授权", () => {
  it("a rejected footer confirmation keeps the picker open and the old authorization", async () => {
    const { container } = await mount(<WpsProbePanel />);
    getWpsApplicationMock.mockReturnValue(hostApplication(2));
    listWpsWorksheetsMock.mockReturnValue(metadataWith([{ index: 1, name: "订单" }, { index: 2, name: "明细", isActive: true }]));
    await click(byText(container, ".scope-trigger", ""));
    scopeAuthorization.mockResolvedValueOnce(false);
    await click(byText(container, ".scope-close", "确认"));
    expect(committed(container)).toBe("1");
    expect(container.querySelector(".scope-popover")).not.toBeNull();
    expect(container.textContent).toContain("暂时无法核实文件与工作表");
    await click(byText(container, ".scope-close", "确认"));
    expect(committed(container)).toBe("2");
    expect(container.querySelector(".scope-popover")).toBeNull();
  });
  it("footer confirmation verifies the changed active sheet without a mode gesture", async () => {
    const { container } = await mount(<WpsProbePanel />);
    getWpsApplicationMock.mockReturnValue(hostApplication(2));
    listWpsWorksheetsMock.mockReturnValue(metadataWith([
      { index: 1, name: "订单" }, { index: 2, name: "明细", isActive: true }
    ]));
    await click(byText(container, ".scope-trigger", ""));
    expect(scopeAuthorization).not.toHaveBeenCalled();
    await click(byText(container, ".scope-close", "确认"));
    expect(scopeAuthorization.mock.lastCall?.[0].selected).toEqual(["2"]);
    expect(committed(container)).toBe("2");
    expect(container.querySelector(".scope-popover")).toBeNull();
  });

  it("restores fixed selections after unmount without following the active sheet", async () => {
    const first = await mount(<WpsProbePanel />);
    await click(byText(first.container, ".scope-trigger", ""));
    await click(byText(first.container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(first.container, "明细"));
    expect(committed(first.container)).toBe("1,2");
    act(() => mounted.pop()!.unmount());
    const calls = scopeAuthorization.mock.calls.length;
    const second = await mount(<WpsProbePanel />);
    expect(committed(second.container)).toBe("1,2");
    expect(second.container.textContent).toContain("已固定 2 个工作表");
    expect(scopeAuthorization.mock.calls.length).toBe(calls);
  });
  it("fails closed when a saved selected sheet is replaced at the same index", async () => {
    const first = await mount(<WpsProbePanel />);
    await click(byText(first.container, ".scope-trigger", ""));
    await click(byText(first.container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(first.container, "明细"));
    act(() => mounted.pop()!.unmount());
    listWpsWorksheetsMock.mockReturnValue(metadataWith([{index:1,name:"订单",isActive:true},{index:2,name:"新表"}]));
    const second = await mount(<WpsProbePanel />);
    expect(committed(second.container)).toBe("");
    expect(second.container.textContent).toContain("原固定工作表无法核实");
  });
  it("opening and refreshing an unchanged picker never grants new access", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-refresh", "刷新"));
    expect(scopeAuthorization).not.toHaveBeenCalled();
    expect(committed(container)).toBe("1");
    expect(container.querySelector(".scope-confirm")).toBeNull();
    expect(container.querySelector(".scope-discard")).toBeNull();
    expect(container.querySelector(".scope-popover > .scope-popover-header .scope-authorization-hint")?.textContent).toContain("勾选即授权");
    expect(container.querySelector(".wps-scope-actions")?.textContent).not.toContain("勾选即授权");
  });

  it("a checkbox commits the exact live target without confirmation or closing the picker", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    expect(scopeAuthorization.mock.lastCall?.[0].selected).toEqual(["1", "2"]);
    expect(committed(container)).toBe("1,2");
    expect(container.querySelector(".scope-popover")).not.toBeNull();
    await click(sheetButton(container, "订单"));
    expect(committed(container)).toBe("2");
  });

  it("confirmation rechecks the selection and reopening unchanged does not reauthorize", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    const calls = scopeAuthorization.mock.calls.length;
    await click(byText(container, ".scope-close", "确认"));
    expect(committed(container)).toBe("1,2");
    await click(byText(container, ".scope-trigger", ""));
    expect(sheetButton(container, "明细").getAttribute("aria-pressed")).toBe("true");
    expect(scopeAuthorization.mock.calls.length).toBe(calls + 1);
  });

  it("clearing the last checkbox revokes all sheets", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(byText(container, ".sheet-picker-toolbar button", "清空"));
    expect(committed(container)).toBe("");
    expect(scopeAuthorization.mock.lastCall?.[0].selected).toEqual([]);
    await click(byText(container, ".scope-close", "确认"));
    await click(byText(container, ".scope-trigger", ""));
    expect(sheetButton(container, "订单").getAttribute("aria-pressed")).toBe("false");
  });

  it("a failed check retains the chosen boxes, blocks access, and refresh retries them", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    scopeAuthorization.mockResolvedValueOnce(false);
    await click(sheetButton(container, "明细"));
    expect(committed(container)).toBe("1");
    expect(container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("true");
    expect(sheetButton(container, "明细").getAttribute("aria-pressed")).toBe("true");
    await click(byText(container, ".scope-refresh", "刷新"));
    expect(committed(container)).toBe("1,2");
    expect(container.querySelector(".scope-close")?.getAttribute("data-scope-pending")).toBe("false");
  });

  it("a missing selected sheet is reported without silently shrinking consent", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    await click(sheetButton(container, "明细"));
    const calls = scopeAuthorization.mock.calls.length;
    listWpsWorksheetsMock.mockReturnValue(metadataWith([{ index: 1, name: "订单", isActive: true }]));
    await click(byText(container, ".scope-refresh", "刷新"));
    expect(committed(container)).toBe("1,2");
    expect(container.querySelector(".scope-admission-summary")?.textContent).toContain("已不存在");
    expect(scopeAuthorization.mock.calls.length).toBe(calls);
  });

  it("host failure never publishes a new scope or exposes private error text", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    getWpsApplicationMock.mockImplementation(() => { throw new Error("PRIVATE_HOST_FAILURE"); });
    const calls = scopeAuthorization.mock.calls.length;
    await click(sheetButton(container, "明细"));
    expect(committed(container)).toBe("1");
    expect(scopeAuthorization.mock.calls.length).toBe(calls);
    expect(container.textContent).not.toContain("PRIVATE_HOST_FAILURE");
    expect(container.querySelector(".scope-admission-summary")?.textContent).toContain("暂时无法核实");
  });

  it("shows compact eligibility outside checkbox rows and keeps the footer outside scrolling content", async () => {
    const { container } = await mount(<WpsProbePanel />);
    await click(byText(container, ".scope-trigger", ""));
    await click(byText(container, ".scope-mode-list button", "选择多个"));
    const row = sheetButton(container, "明细");
    for (const message of Object.values(SCOPE_ADMISSION_MESSAGES)) expect(row.textContent).not.toContain(message);
    const actions = container.querySelector(".wps-scope-actions")!;
    expect(container.querySelector(".scope-popover-scroll")!.contains(actions)).toBe(false);
    expect(actions.querySelector(".scope-close")).not.toBeNull();
    expect(actions.querySelector(".scope-confirm")).toBeNull();
    const glyph = container.querySelector(".scope-refresh-glyph");
    await click(byText(container, ".scope-refresh", "刷新"));
    expect(container.querySelector(".scope-refresh-glyph")).not.toBe(glyph);
  });
});
