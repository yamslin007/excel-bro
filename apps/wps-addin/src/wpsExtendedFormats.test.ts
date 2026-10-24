import { describe, expect, it } from "vitest";
import type { AnalysisPlan } from "../../excel-addin/src/contracts";
import { assertWpsRecoveryDescriptor } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import { executeWpsPlan, prepareWpsPlan, undoWpsExecution, inspectWpsUndoRecord } from "./wpsPlanExecutor";
import { wpsRecoveryDescriptor } from "./wpsExecutionLedger";
import { wpsOleColorFromHex } from "./wpsFormatCapability";
function formatHost(rows: number, columns: number) {
  const grid: Array<Array<{
    bold: boolean; italic: boolean; name: string; size: number; color: number;
    horizontalAlignment: number; wrapText: boolean; rowHeight: number; fillColor: number; fillPattern: number; fillPatternColor: number; numberFormat: string;
  }>> = [];
  for (let r = 0; r < rows; r += 1) {
    const rowCells: any[] = [];
    for (let c = 0; c < columns; c += 1) {
      rowCells.push({
        bold: false, italic: false, name: "宋体", size: 11, color: 0,
        horizontalAlignment: 1, wrapText: false, rowHeight: 13.5, fillColor: 0x123456, fillPattern: -4142, fillPatternColor: 0, numberFormat: "General"
      });
    }
    grid.push(rowCells);
  }

  const cellCache = new Map<string, WpsRange>();
  const makeCell = (r: number, c: number): WpsRange => {
    const key = `${r},${c}`;
    const cached = cellCache.get(key);
    if (cached) return cached;
    const state = grid[r][c];
    const cell: WpsRange = {
      Interior: {
        get Color() { return state.fillColor; }, set Color(v: number) { state.fillColor = v; },
        get Pattern() { return state.fillPattern; }, set Pattern(v: unknown) { state.fillPattern = v as number; },
        get PatternColor() { return state.fillPatternColor; }, set PatternColor(v: unknown) { state.fillPatternColor = v as number; }
      },
      get NumberFormat() { return state.numberFormat; }, set NumberFormat(v: unknown) { state.numberFormat = v as string; },
      Font: {
        get Bold() { return state.bold; }, set Bold(v: unknown) { state.bold = v as boolean; },
        get Italic() { return state.italic; }, set Italic(v: unknown) { state.italic = v as boolean; },
        get Name() { return state.name; }, set Name(v: unknown) { state.name = v as string; },
        get Size() { return state.size; }, set Size(v: unknown) { state.size = v as number; },
        get Color() { return state.color; }, set Color(v: unknown) { state.color = v as number; }
      },
      get HorizontalAlignment() { return state.horizontalAlignment; },
      set HorizontalAlignment(v: unknown) { state.horizontalAlignment = v as number; },
      get WrapText() { return state.wrapText; }, set WrapText(v: unknown) { state.wrapText = v as boolean; },
      get RowHeight() { return state.rowHeight; }, set RowHeight(v: unknown) { state.rowHeight = v as number; },
      Rows: { Count: 1 },
      Columns: { Count: 1 }
    };
    cellCache.set(key, cell);
    return cell;
  };

  const range: WpsRange = {
    MergeCells: false,
    Rows: { Count: rows },
    Columns: { Count: columns },
    Item: (r: number, c: number) => makeCell(r - 1, c - 1),
    // 单格区域（Range("A1")）在真实 WPS 里**本身就是单元格**，`Font`/对齐/换行/
    // 行高直接挂在 range 上（`captureWpsFormatSnapshot` 单格路径读 `range` 本身，
    // 不走 `Item`）。替身把这些 getter/setter 委托到第一个格，模拟真实单格 range。
    Font: {
      get Bold() { return grid[0][0].bold; }, set Bold(v: unknown) { grid[0][0].bold = v as boolean; },
      get Italic() { return grid[0][0].italic; }, set Italic(v: unknown) { grid[0][0].italic = v as boolean; },
      get Name() { return grid[0][0].name; }, set Name(v: unknown) { grid[0][0].name = v as string; },
      get Size() { return grid[0][0].size; }, set Size(v: unknown) { grid[0][0].size = v as number; },
      get Color() { return grid[0][0].color; }, set Color(v: unknown) { grid[0][0].color = v as number; }
    },
    get HorizontalAlignment() { return grid[0][0].horizontalAlignment; },
    set HorizontalAlignment(v: unknown) { grid[0][0].horizontalAlignment = v as number; },
    get WrapText() { return grid[0][0].wrapText; }, set WrapText(v: unknown) { grid[0][0].wrapText = v as boolean; },
    get RowHeight() { return grid[0][0].rowHeight; }, set RowHeight(v: unknown) { grid[0][0].rowHeight = v as number; },
    // 格式动作的 `prepareFormatAction` 现在要读值矩阵作为"值不变"基线
    // （副本审阅 `previewFingerprint` 用真实值指纹，而非空矩阵指纹）。
    // 替身给全 null 矩阵即可 —— 格式测试不关心值内容，只关心 Formula 可读。
    Value2: Array.from({ length: rows }, () => Array.from({ length: columns }, () => null)),
    Formula: Array.from({ length: rows }, () => Array.from({ length: columns }, () => null))
  };
  const sheet: WpsWorksheet = {
    Name: "明细",
    ProtectContents: false,
    Range: () => range
  };
  const workbook = {
    Name: "格式测试.xlsx",
    FullName: "C:\\Data\\格式测试.xlsx",
    Path: "C:\\Data",
    ReadOnly: false,
    Worksheets: { Count: 1, Item: () => sheet },
    ActiveSheet: sheet
  };
  const storage = new Map<string, unknown>();
  const application: WpsApplication = {
    Version: "12.0",
    ActiveWorkbook: workbook,
    ActiveSheet: sheet,
    PluginStorage: {
      getItem: (key) => storage.get(key),
      setItem: (key, value) => { storage.set(key, value); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  };
  return {
    application,
    grid,
    cell: (r: number, c: number) => grid[r][c],
    storedRecords: () => JSON.parse(String(storage.get("excel-bro.wps.undo.v1") ?? "[]"))
  };
}

function plan(actions: AnalysisPlan["actions"]): AnalysisPlan {
  return {
    id: "plan-fmt-1",
    title: "设置格式",
    summary: "设置字体与对齐",
    assumptions: [],
    warnings: [],
    actions,
    acceptanceCriteria: []
  };
}


describe("WPS added cell formats preserve mixed originals", () => {
  it("fill uses OLE color, restores original per-cell patterns and preserves numbers", async () => {
    const host = formatHost(2, 2);
    host.grid[1][1].fillColor = 0xabcdef;
    host.grid[1][1].fillPattern = 3;
    const before = structuredClone(host.grid);
    const prepared = prepareWpsPlan(host.application, [1], plan([
      { type: "setFill", sheet: "明细", range: "A1:B2", color: "#FF0012" }
    ]));
    expect(host.grid).toEqual(before);
    const evidence = wpsRecoveryDescriptor(prepared);
    expect(evidence.actions[0].formatFields).toEqual(["fillColor", "fillPattern"]);
    expect(() => assertWpsRecoveryDescriptor(evidence)).not.toThrow();
    const record = await executeWpsPlan(host.application, prepared);
    expect(host.grid[1][1].fillColor).toBe(0x1200ff);
    expect(host.grid[1][1].fillPattern).toBe(1);
    expect(inspectWpsUndoRecord(host.application, record.executionId).kind).toBe("completed");
    await expect(undoWpsExecution(host.application, record.executionId)).resolves.toMatchObject({state:"restored"});
    expect(host.grid).toEqual(before);
  });
  it("number format restores mixed formats without overwriting later font edits", async () => {
    const host = formatHost(2, 2);
    host.grid[0][1].numberFormat = "0%";
    const record = await executeWpsPlan(host.application, prepareWpsPlan(host.application, [1], plan([
      { type: "setNumberFormat", sheet: "明细", range: "A1:B2", formatCode: "0.00" }
    ])));
    expect(host.grid.every(row => row.every(cell => cell.numberFormat === "0.00"))).toBe(true);
    host.grid[0][0].bold = true;
    await expect(undoWpsExecution(host.application, record.executionId)).resolves.toMatchObject({state:"restored"});
    expect(host.grid[0][1].numberFormat).toBe("0%");
    expect(host.grid[1][1].numberFormat).toBe("General");
    expect(host.grid[0][0].bold).toBe(true);
  });
  it("overlapping fills and number formats restore to the original, not an intermediate state", async () => {
    const host = formatHost(2, 2), before = structuredClone(host.grid);
    const prepared = prepareWpsPlan(host.application, [1], plan([
      { type: "setFill", sheet: "明细", range: "A1:B2", color: "#FF0000" },
      { type: "setNumberFormat", sheet: "明细", range: "A1:B2", formatCode: "0%" },
      { type: "setFill", sheet: "明细", range: "A1:B2", color: "#0000FF" }
    ]));
    const descriptor = wpsRecoveryDescriptor(prepared);
    const record = await executeWpsPlan(host.application, prepared, undefined, undefined, async (phase, checkpoint) => {
      const action = descriptor.actions[checkpoint.actionIndex];
      expect(checkpoint.beforeFingerprint).toBe(action.previewFingerprint);
      if (phase === "after") expect(checkpoint.afterFingerprint).toBe(action.expectedFingerprint);
    });
    expect(host.grid[0][0].fillColor).toBe(0xff0000);
    await expect(undoWpsExecution(host.application, record.executionId)).resolves.toMatchObject({state:"restored"});
    expect(host.grid).toEqual(before);
  });
  it("missing fill metadata rejects before creating a recovery record or changing cells", () => {
    const host = formatHost(2, 2), before = structuredClone(host.grid);
    delete host.application.ActiveWorkbook!.Worksheets.Item(1).Range!("A1:B2").Item!(1, 1).Interior;
    expect(() => prepareWpsPlan(host.application, [1], plan([
      { type: "setFill", sheet: "明细", range: "A1:B2", color: "#FF0000" }
    ]))).toThrow(/无法完整采集/);
    expect(host.grid).toEqual(before);
    expect(host.storedRecords()).toEqual([]);
  });
  it("external fill edits are a conflict and must never be overwritten", async () => {
    const host = formatHost(2, 2);
    const record = await executeWpsPlan(host.application, prepareWpsPlan(host.application, [1], plan([
      { type: "setFill", sheet: "明细", range: "A1:B2", color: "#FF0000" }
    ])));
    host.grid[1][1].fillColor = 0x555555;
    const edited = structuredClone(host.grid);
    await expect(undoWpsExecution(host.application, record.executionId)).rejects.toThrow(/修改/);
    expect(host.grid).toEqual(edited);
  });
  it("a host ignoring NumberFormat never receives a verified receipt", async () => {
    const host = formatHost(2, 2);
    const cell = host.application.ActiveWorkbook!.Worksheets.Item(1).Range!("A1:B2").Item!(1, 1);
    Object.defineProperty(cell, "NumberFormat", { get: () => "General", set: () => undefined });
    await expect(executeWpsPlan(host.application, prepareWpsPlan(host.application, [1], plan([
      { type: "setNumberFormat", sheet: "明细", range: "A1:B2", formatCode: "0%" }
    ])))).rejects.toThrow(/回读/);
    expect(host.storedRecords()[0].state).not.toBe("verified");
  });
  it("unavailable optional fill properties do not disable existing font operations", async () => {
    const host = formatHost(2, 2);
    const cell = host.application.ActiveWorkbook!.Worksheets.Item(1).Range!("A1:B2").Item!(1, 1);
    Object.defineProperty(cell, "Interior", {get: () => {throw new Error("unsupported host property");}});
    expect(() => prepareWpsPlan(host.application,[1],plan([{type:"setFill",sheet:"明细",range:"A1:B2",color:"#123456"}]))).toThrow("无法完整采集");
    const record = await executeWpsPlan(host.application,prepareWpsPlan(host.application,[1],plan([{type:"setFont",sheet:"明细",range:"A1:B2",bold:true}])));
    expect(host.grid[0][0].bold).toBe(true);
    await expect(undoWpsExecution(host.application,record.executionId)).resolves.toMatchObject({state:"restored"});
  });
  it("RGB channels are not swapped or inferred from symmetric colors", () => {
    expect(wpsOleColorFromHex("#123456")).toBe(0x563412);
    expect(wpsOleColorFromHex("#FF0000")).toBe(255);
    expect(wpsOleColorFromHex("#0000FF")).toBe(16711680);
  });
});
