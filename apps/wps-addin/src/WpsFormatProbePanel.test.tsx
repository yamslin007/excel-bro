import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import WpsFormatProbePanel from "./WpsFormatProbePanel";
import { getWpsApplication } from "./wpsHost";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";

vi.mock("./devProbe", () => ({
  reportDevProbe: () => undefined,
  reportDevProbeDetail: () => undefined
}));
vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const getWpsApplicationMock = vi.mocked(getWpsApplication);

function fakeRange(overrides: Partial<WpsRange> = {}): WpsRange {
  const font: Record<string, unknown> = { Bold: false, Italic: false, Name: "宋体", Size: 11, Color: 0, Underline: false };
  const interior: { Color: number; Pattern: unknown; PatternColor: unknown } = { Color: 0xffffff, Pattern: -4142, PatternColor: 0 };
  const border: { LineStyle: unknown; Weight: unknown; ColorIndex: unknown } = { LineStyle: -4142, Weight: 1, ColorIndex: 1 };
  const state: Record<string, unknown> = { HorizontalAlignment: -4131, VerticalAlignment: -4107, WrapText: false, RowHeight: 15, ColumnWidth: 8.43 };
  return {
    Value2: "",
    Formula: undefined,
    Font: {
      get Bold() { return font.Bold; }, set Bold(v: unknown) { font.Bold = v; },
      get Italic() { return font.Italic; }, set Italic(v: unknown) { font.Italic = v; },
      get Name() { return font.Name; }, set Name(v: unknown) { font.Name = v; },
      get Size() { return font.Size; }, set Size(v: unknown) { font.Size = v; },
      get Color() { return font.Color; }, set Color(v: unknown) { font.Color = v; },
      get Underline() { return font.Underline; }, set Underline(v: unknown) { font.Underline = v; }
    },
    Interior: {
      get Color() { return interior.Color; }, set Color(v: number) { interior.Color = v; },
      get Pattern() { return interior.Pattern; }, set Pattern(v: unknown) { interior.Pattern = v; }
    },
    Borders: {
      Count: 6,
      Item: () => ({
        get LineStyle() { return border.LineStyle; }, set LineStyle(v: unknown) { border.LineStyle = v; },
        get Weight() { return border.Weight; }, set Weight(v: unknown) { border.Weight = v; },
        get ColorIndex() { return border.ColorIndex; }, set ColorIndex(v: unknown) { border.ColorIndex = v; }
      })
    },
    get HorizontalAlignment() { return state.HorizontalAlignment; }, set HorizontalAlignment(v: unknown) { state.HorizontalAlignment = v; },
    get VerticalAlignment() { return state.VerticalAlignment; }, set VerticalAlignment(v: unknown) { state.VerticalAlignment = v; },
    get WrapText() { return state.WrapText; }, set WrapText(v: unknown) { state.WrapText = v; },
    get RowHeight() { return state.RowHeight; }, set RowHeight(v: unknown) { state.RowHeight = v; },
    get ColumnWidth() { return state.ColumnWidth; }, set ColumnWidth(v: unknown) { state.ColumnWidth = v; },
    EntireRow: { AutoFit: () => undefined },
    EntireColumn: { AutoFit: () => undefined },
    ...overrides
  };
}

function fakeApplication(range: WpsRange = fakeRange()): WpsApplication {
  const sheet: WpsWorksheet = { Name: "探测表", Range: () => range };
  return {
    Version: "12.0",
    PluginStorage: { getItem: () => null, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined,
    ActiveWorkbook: {
      Name: "探测.xlsx",
      FullName: "C:\\Data\\探测.xlsx",
      Worksheets: { Count: 1, Item: () => sheet },
      ActiveSheet: sheet
    },
    ActiveSheet: sheet
  };
}

let mounted: Array<{ unmount: () => void }> = [];

async function mount(element: ReactElement): Promise<HTMLElement> {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  await act(async () => { root.render(element); });
  mounted.push({ unmount: () => root.unmount() });
  return host;
}

function click(host: HTMLElement, selector: string) {
  const el = host.querySelector(selector) as HTMLElement;
  el?.click();
  return el;
}

afterEach(async () => {
  await act(async () => {
    for (const entry of mounted.splice(0)) entry.unmount();
    document.body.innerHTML = "";
  });
});

describe("WPS 格式能力探测面板", () => {
  it("初始状态展示工作簿/工作表/目标，未确认时无写入", async () => {
    getWpsApplicationMock.mockReturnValue(fakeApplication());
    const host = await mount(<WpsFormatProbePanel />);
    expect(host.textContent).toContain("探测.xlsx");
    expect(host.textContent).toContain("探测表");
    expect(host.textContent).toContain("选择目标并确认");
  });

  it('点确认按钮后，出现「确认这是可丢弃的独立测试工作簿」对话框，需再点确认才运行', async () => {
    getWpsApplicationMock.mockReturnValue(fakeApplication());
    const host = await mount(<WpsFormatProbePanel />);
    // 第一步：点"选择目标并确认"。
    await act(async () => { click(host, "button.primary"); });
    expect(host.textContent).toContain("确认这是可丢弃的独立测试工作簿");
    // 未点"确认并运行探测"前，不应有报告表格。
    expect(host.querySelector(".format-probe-table")).toBeNull();
  });

  it("确认后运行探测，逐项报告展示，AutoFit 标未执行", async () => {
    getWpsApplicationMock.mockReturnValue(fakeApplication());
    const host = await mount(<WpsFormatProbePanel />);
    await act(async () => { click(host, "button.primary"); });
    await act(async () => { click(host, ".poc-panel-confirm button.primary"); });
    const table = host.querySelector(".format-probe-table");
    expect(table).not.toBeNull();
    expect(host.textContent).toContain("整轮基线恢复通过");
  });

  it("非法目标（区域地址）时按钮禁用，不运行探测", async () => {
    getWpsApplicationMock.mockReturnValue(fakeApplication());
    const host = await mount(<WpsFormatProbePanel />);
    const input = host.querySelector("input[aria-label='探测单元格']") as HTMLInputElement;
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
      setter?.call(input, "A1:B2");
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });
    const button = host.querySelector("button.primary") as HTMLButtonElement;
    expect(button.disabled).toBe(true);
    expect(host.textContent).toContain("不是单个单元格");
  });

  it("换簿后旧确认失效：确认后切换工作簿，回到需重新确认状态", async () => {
    const app = fakeApplication();
    getWpsApplicationMock.mockReturnValue(app);
    const host = await mount(<WpsFormatProbePanel />);
    await act(async () => { click(host, "button.primary"); });
    expect(host.textContent).toContain("确认这是可丢弃的独立测试工作簿");
    // 模拟切换工作簿：FullName 变化。
    app.ActiveWorkbook!.FullName = "C:\\Data\\另一个.xlsx";
    app.ActiveWorkbook!.Name = "另一个.xlsx";
    // 重新触发一次身份读取（通过再次点击按钮进入确认流程，会因身份变化而失效）。
    await act(async () => { click(host, ".poc-panel-confirm button:not(.primary)"); }); // 取消
    await act(async () => { click(host, "button.primary"); });
    expect(host.textContent).toContain("确认这是可丢弃的独立测试工作簿");
  });

  it("能力组选择：取消所有组后按钮禁用，至少选一组才能确认", async () => {
    getWpsApplicationMock.mockReturnValue(fakeApplication());
    const host = await mount(<WpsFormatProbePanel />);
    // 初始全选，按钮可用。
    const button = () => host.querySelector("button.primary") as HTMLButtonElement;
    expect(button().disabled).toBe(false);
    // 取消所有组勾选。
    const checkboxes = Array.from(
      host.querySelectorAll<HTMLInputElement>(".probe-group-select input[type='checkbox']")
    );
    await act(async () => {
      for (const checkbox of checkboxes) {
        checkbox.click();
      }
    });
    expect(button().disabled).toBe(true);
    // 重新勾选一组后按钮恢复可用。
    await act(async () => { checkboxes[0].click(); });
    expect(button().disabled).toBe(false);
  });

  it("只选 font 组运行，报告只含字体成员，不含对齐/尺寸", async () => {
    getWpsApplicationMock.mockReturnValue(fakeApplication());
    const host = await mount(<WpsFormatProbePanel />);
    // 取消所有，只勾选 font 组（第一项）。
    const checkboxes = Array.from(
      host.querySelectorAll<HTMLInputElement>(".probe-group-select input[type='checkbox']")
    );
    await act(async () => {
      for (const checkbox of checkboxes) checkbox.click();
      checkboxes[0].click();
    });
    await act(async () => { click(host, "button.primary"); });
    await act(async () => { click(host, ".poc-panel-confirm button.primary"); });
    const cells = Array.from(host.querySelectorAll(".format-probe-table tbody tr td:first-child")).map(
      (td) => td.textContent
    );
    expect(cells.length).toBe(6);
    expect(cells.every((name) => name?.startsWith("Font."))).toBe(true);
    expect(host.textContent).not.toContain("HorizontalAlignment");
  });
});
