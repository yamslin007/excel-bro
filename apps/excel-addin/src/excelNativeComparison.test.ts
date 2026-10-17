import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AnalysisPlan, CellValue } from "./contracts";
import { assertExcelNativeComparisonUnchanged, closeCompatReviewCopy, closeExcelNativeComparison, compareExcelNativeCommittedResult, excelNativeReviewMode, nativeComparisonPlanIssue, stageExcelNativeComparison } from "./excelNativeComparison";
import { nativeColumnName } from "./nativeReviewFile";
import { parseExplicitA1Rect } from "./planCore";
import { readSeedEntries as unzipSeed } from "./nativeReviewTestHelpers";
import { executePlan } from "./excel";
import { ExcelReviewLayoutVersionError } from "./api";

const layout = vi.hoisted(() => ({prepare: vi.fn(), arrange: vi.fn(), close: vi.fn()}));
vi.mock('./excelWindowLayout', () => ({prepareLocalExcelWindows: layout.prepare, arrangeLocalExcelWindows: layout.arrange, closeLocalExcelWindows: layout.close}));
beforeEach(() => { layout.prepare.mockReset().mockRejectedValue(new Error('unavailable')); layout.arrange.mockReset().mockResolvedValue({status:'arranged',layoutVersion:3,copySide:'left',originalSide:'right',synchronousScrolling:'enabled'});
  layout.close.mockReset().mockResolvedValue({status:'closed', layoutVersion:4, copyClosed:true, recoveryCopyPath:'E:/recovery/review-copy.xlsx'}); });

const plan = (actions: AnalysisPlan["actions"] = [{ type: "writeValues", sheet: "Ledger", range: "C2:C3", values: [["LOCAL_PRIVATE_SENTINEL"], ["done"]] }]): AnalysisPlan => ({ id: "p", title: "Update", summary: "Update", assumptions: [], warnings: [], actions });
const load = () => undefined;
function host() {
  const writes: string[] = [], colors: string[] = [], windows: any[] = [];
  function sheet(name: string, id: string, initial: CellValue[][]) {
    const cells = new Map<string, { value: CellValue; formula: CellValue; format: string; props: any }>();
    const put = (r: number, c: number, value: CellValue, formula: CellValue = value) => cells.set(`${r},${c}`, { value, formula, format: "General", props: {} });
    initial.forEach((row, r) => row.forEach((value, c) => put(r + 1, c + 1, value)));
    const s: any = { name, id, load, protection: { protected: false, load }, calculate: vi.fn(),
      tables: { getCount: () => ({ value: 0 }) }, charts: { getCount: () => ({ value: 0 }) },
      shapes: { getCount: () => ({ value: 0 }) }, pivotTables: { getCount: () => ({ value: 0 }) }, names: { getCount: () => ({ value: 0 }) } };
    const cell = (r: number, c: number) => cells.get(`${r},${c}`) ?? { value: "", formula: "", format: "General", props: {} };
    function range(address: string): any {
      const rect = parseExplicitA1Rect(address)!;
      const matrix = (field: "value" | "formula" | "format" | "props") => Array.from({ length: rect.rows }, (_, r) => Array.from({ length: rect.columns }, (_, c) => cell(rect.startRow + r, rect.startColumn + c)[field]));
      const write = (values: CellValue[][], formulas: boolean) => {
        writes.push(`${id}:${address}`);
        values.forEach((row, r) => row.forEach((value, c) => {
          const cr = rect.startRow + r, cc = rect.startColumn + c;
          const calculated = formulas && typeof value === "string" && /^=B\d+\+1$/.test(value) ? Number(cell(cr, 2).value) + 1 : value;
          put(cr, cc, calculated, value);
        }));
      };
      const rangeFormat = { load, rowHeight: 16, columnWidth: 60, horizontalAlignment: "General", verticalAlignment: "Bottom", wrapText: false, indentLevel: 0,
        font: { load, bold: false, color: "#000000", name: "Arial", size: 11, italic: false, underline: "None" },
        borders: { getItem: () => ({ load, style: "None", color: "#000000", weight: "Thin" }) },
        fill: { get color() { return "#FFFFFF"; }, set color(value: string) { colors.push(`${id}:${address}:${value}`); } } };
      return { load, isNullObject: false, rowIndex: rect.startRow - 1, columnIndex: rect.startColumn - 1, rowCount: rect.rows, columnCount: rect.columns,
        get values() { return matrix("value"); }, set values(value) { write(value, false); },
        get formulas() { return matrix("formula"); }, set formulas(value) { write(value, true); },
        get numberFormat() { return matrix("format"); }, set numberFormat(value) { value.forEach((row: string[], r: number) => row.forEach((v, c) => { const old = cell(rect.startRow + r, rect.startColumn + c); cells.set(`${rect.startRow + r},${rect.startColumn + c}`, { ...old, format: v }); })); },
        format: rangeFormat,
        clear() { write(Array.from({ length: rect.rows }, () => Array(rect.columns).fill("")), false); },
        conditionalFormats: { getCount: () => ({ value: 0 }) },
        getMergedAreasOrNullObject: () => ({ isNullObject: true, load }),
        getSpecialCellsOrNullObject(type: string) {
          const items = type === "Formulas" ? [...cells.entries()].filter(([, value]) => typeof value.formula === "string" && value.formula.startsWith("=") && value.formula !== value.value).map(([key]) => {
            const [r, c] = key.split(",").map(Number); return { rowIndex: r - 1, columnIndex: c - 1, rowCount: 1, columnCount: 1 };
          }) : [];
          return { isNullObject: items.length === 0, load, areas: { items, load } };
        },
        getCellProperties: () => ({ value: matrix("props") }),
        getDisplayedCellProperties: () => ({
          value: Array.from({ length: rect.rows }, () =>
            Array.from({ length: rect.columns }, () => ({
              format: {
                fill: { pattern: "None" },
                font: { name: "Arial", size: 11, bold: false, italic: false, color: "#000000", underline: "None" },
                horizontalAlignment: "General",
                verticalAlignment: "Bottom",
                wrapText: false,
                indentLevel: 0,
                borders: {
                  top: { style: "None", color: "#000000", weight: "Thin" },
                  bottom: { style: "None", color: "#000000", weight: "Thin" },
                  left: { style: "None", color: "#000000", weight: "Thin" },
                  right: { style: "None", color: "#000000", weight: "Thin" }
                }
              }
            }))
          )
        }),
        setCellProperties(value: any[][]) { value.forEach((row, r) => row.forEach((v, c) => { const old = cell(rect.startRow + r, rect.startColumn + c); cells.set(`${rect.startRow + r},${rect.startColumn + c}`, { ...old, props: v }); })); },
        getRowProperties: () => ({ value: Array.from({ length: rect.rows }, () => ({ format: { rowHeight: 16 } })) }),
        getColumnProperties: () => ({ value: Array.from({ length: rect.columns }, () => ({ format: { columnWidth: 60 } })) }),
        setRowProperties: () => undefined,
        setColumnProperties: () => undefined,
        getRow: (r: number) => range(`${nativeColumnName(rect.startColumn)}${rect.startRow + r}:${nativeColumnName(rect.endColumn)}${rect.startRow + r}`),
        getColumn: (c: number) => range(`${nativeColumnName(rect.startColumn + c)}${rect.startRow}:${nativeColumnName(rect.startColumn + c)}${rect.endRow}`)
      };
    }
    s.getRange = range;
    s.getUsedRangeOrNullObject = () => {
      const coords = [...cells.keys()].map(key => key.split(",").map(Number));
      return range(`A1:${nativeColumnName(Math.max(...coords.map(p => p[1])))}${Math.max(...coords.map(p => p[0]))}`);
    };
    return s;
  }
  const original = sheet("Ledger", "original", [["Key", "Input", "Result"], ["a", 7, 99], ["b", 9, ""]]);
  const left: any = { name: "Original.xlsx", activeWorksheet: original, load, activate: vi.fn(), close: vi.fn(() => ({ value: false })) };
  windows.push(left);
  const collection = { load, items: windows, compareCurrentSideBySideWith: vi.fn(), resetPositionsSideBySide: vi.fn() };
  const getItem = (name: string) => { if (name !== "Ledger") throw new Error("UNAUTHORIZED_READ"); return original; };
  const context = { sync: vi.fn(async () => undefined), workbook: { worksheets: { load, items: [original], getItem, getItemOrNullObject: getItem },
    tables: { load, items: [] }, pivotTables: { load, items: [] }, names: { getCount: () => ({ value: 0 }), load, items: [] }, application: { windows: collection, activeWindow: left } } };
  const createWorkbook = vi.fn(async (base64: string) => {
    const files = unzipSeed(base64), name = files["xl/workbook.xml"].match(/sheet name="([^"]+)"/)![1];
    const copied = sheet(name, "copy", original.getRange("A1:C3").values);
    windows.push({ name: "Book1", activeWorksheet: copied, load, activate: vi.fn(), close: vi.fn(() => { windows.splice(1, 1); return { value: true }; }) });
  });
  vi.stubGlobal("Office", { context: { document: { url: "file:///E:/Original.xlsx" }, requirements: { isSetSupported: vi.fn(() => true) } } });
  vi.stubGlobal("Excel", { run: async (job: any) => job(context), createWorkbook, ClearApplyTo: { contents: "Contents" }, SpecialCellType: { formulas: "Formulas", dataValidations: "DataValidations" }, BorderIndex: { edgeTop: "EdgeTop", edgeBottom: "EdgeBottom", edgeLeft: "EdgeLeft", edgeRight: "EdgeRight", insideHorizontal: "InsideHorizontal", insideVertical: "InsideVertical" } });
  return { writes, colors, original, windows, collection, context, createWorkbook };
}
afterEach(() => vi.unstubAllGlobals());
describe("native authorized copy", () => {
  it('uses the bounded bridge to explicitly enable scrolling even on the Desktop 1.1 path', async () => {
    const h = host(), ticket = {token:'T'.repeat(43),layoutVersion:3};
    layout.prepare.mockResolvedValue(ticket);
    const session = await stageExcelNativeComparison(plan(), ['Ledger']);
    expect(layout.prepare).toHaveBeenCalledWith(session.workbookIdentity, 'Ledger', session.copySheetName);
    expect(layout.arrange).toHaveBeenCalledWith(ticket);
    expect(session.windowLayout).toBe('arranged');
    expect(session.synchronousScrolling).toBe('enabled');
    expect(h.writes.every(write => write.startsWith('copy:'))).toBe(true);
  });
  it("creates a separate copy, applies the real plan there, colors changes and requests native compare; discard never writes original", async () => {
    const h = host(), fetch = vi.fn(); vi.stubGlobal("fetch", fetch);
    const session = await stageExcelNativeComparison(plan(), ["Ledger"]);
    expect(JSON.stringify(session.plan)).toContain("LOCAL_PRIVATE_SENTINEL"); // positive privacy anchor
    expect(unzipSeed(h.createWorkbook.mock.calls[0][0])["xl/worksheets/sheet1.xml"]).toContain("Result");
    expect(h.writes.length).toBeGreaterThan(0); expect(h.writes.every(w => w.startsWith("copy:"))).toBe(true);
    expect(h.original.getRange("C2").values).toEqual([[99]]);
    expect(session.changedCells).toBe(2); expect(h.colors).toHaveLength(2);
    expect(h.collection.compareCurrentSideBySideWith).toHaveBeenCalledWith("Book1");
    expect(h.collection.resetPositionsSideBySide).toHaveBeenCalledOnce();
    expect(fetch).not.toHaveBeenCalled();
    expect(await closeExcelNativeComparison(session)).toBe('closed');
    expect(h.writes.some(w => w.startsWith("original:"))).toBe(false);
  });
  it.each(["original", "copy target", "copy source", "authorization", "plan", "workbook", "closed copy", "copy identity"])("blocks save when %s changed", async reason => {
    const h = host(), p = plan(), session = await stageExcelNativeComparison(p, ["Ledger"]);
    let scope = ["Ledger"];
    if (reason === "original") h.original.getRange("B2").values = [[8]];
    if (reason === "copy target") h.windows[1].activeWorksheet.getRange("C2").values = [["edited"]];
    if (reason === "copy source") h.windows[1].activeWorksheet.getRange("B2").values = [["edited"]];
    if (reason === "authorization") scope = ["Other"];
    if (reason === "plan") p.actions = [{ type: "clearRange", sheet: "Ledger", range: "C2", applyTo: "contents" }];
    if (reason === "workbook") Office.context.document.url = "file:///E:/Other.xlsx";
    if (reason === "closed copy") h.windows.splice(1, 1);
    if (reason === "copy identity") h.windows[1].activeWorksheet.id = "replacement";
    const before = h.writes.length;
    await expect(assertExcelNativeComparisonUnchanged(h.context as never, session, scope, p)).rejects.toThrow();
    expect(h.writes).toHaveLength(before); // guard itself is read-only
  });
  it("preserves actual formulas and does not turn literal '=' text into a formula", async () => {
    const h = host(); h.original.getRange("A2").values = [["=LITERAL"]]; h.original.getRange("C2").formulas = [["=B2+1"]]; h.writes.length = 0;
    const session = await stageExcelNativeComparison(plan([{ type: "writeFormulas", sheet: "Ledger", range: "C3", formulas: [["=B3+1"]] }]), ["Ledger"]);
    expect(h.writes).not.toContain("copy:A1:C3");
    expect(h.writes).toContain("copy:C2");
    expect(h.windows[1].activeWorksheet.getRange("C3").values).toEqual([[10]]);
    expect(session.original.formulaCells).toEqual(["C2"]);
  });
  it("does not write if the new native window cannot be uniquely identified", async () => {
    const h = host(); h.createWorkbook.mockImplementation(async () => undefined);
    await expect(stageExcelNativeComparison(plan(), ["Ledger"])).rejects.toThrow("无法唯一识别");
    expect(h.writes).toEqual([]);
  });
  it("rejects unsupported hosts before opening or reading, and reports an unconfirmed close rather than a closed one", async () => {
    const h = host(); (Office.context.requirements.isSetSupported as ReturnType<typeof vi.fn>).mockReturnValue(false);
    await expect(stageExcelNativeComparison(plan(), ["Ledger"])).rejects.toThrow("缺少");
    expect(h.createWorkbook).not.toHaveBeenCalled(); expect(h.context.sync).not.toHaveBeenCalled();
    (Office.context.requirements.isSetSupported as ReturnType<typeof vi.fn>).mockReturnValue(true);
    const session = await stageExcelNativeComparison(plan(), ["Ledger"]);
    h.windows[1].close.mockReturnValue({ value: false });
    // close() 返回 false 是**歧义**信号（被取消 / 被推迟 / 被拒），我们又没有对话框证据
    // ⇒ 只能说"关闭尚未确认"，绝不说"已关闭"，也不推断宿主在问是否保存。
    expect(await closeExcelNativeComparison(session)).toBe('close_unconfirmed');
    // 窗口已不在 ⇒ 重复清理幂等，算成功收尾。
    h.windows.splice(1, 1);
    expect(await closeExcelNativeComparison(session)).toBe('already_closed');
  });
  it("runs the actual executor guard AFTER the durable callback and BEFORE any target write", async () => {
    const h = host(), p = plan(), session = await stageExcelNativeComparison(p, ["Ledger"]);
    const order: string[] = [];
    h.writes.length = 0;
    const guard = vi.fn(async (context: Excel.RequestContext) => {
      order.push("guard"); await assertExcelNativeComparisonUnchanged(context, session, ["Ledger"], p);
    });
    await expect(executePlan(p, {
      authorizedSheetNames: ["Ledger"], beforeFirstWrite: guard,
      recovery: { startActionIndex: 0, onBeforeAction: async () => {
        order.push("registered"); h.original.getRange("B2").values = [[100]]; // human edit while registration is in flight
      }, onActionCompleted: vi.fn() }
    })).rejects.toThrow("原表在审阅期间已变化");
    expect(order).toEqual(["registered", "guard"]);
    expect(guard).toHaveBeenCalledOnce();
    expect(h.writes).toEqual(["original:B2"]); // only simulated human edit, ZERO plan target writes
    expect(h.original.getRange("C2").values).toEqual([[99]]);
  });
  it("compares committed actual values AND formulas with the reviewed copy, without further writes", async () => {
    const h = host(), p = plan([{ type: "writeFormulas", sheet: "Ledger", range: "C2", formulas: [["=B2+1"]] }]);
    const session = await stageExcelNativeComparison(p, ["Ledger"]);
    expect(await compareExcelNativeCommittedResult(session)).toBe("different");
    h.original.getRange("C2").formulas = [["=B2+1"]];
    expect(await compareExcelNativeCommittedResult(session)).toBe("matched");
    h.original.getRange("B2").values = [[100]]; h.original.getRange("C2").formulas = [["=B2+1"]]; // same formula, different calculated value
    const before = h.writes.length;
    expect(await compareExcelNativeCommittedResult(session)).toBe("different");
    expect(h.writes).toHaveLength(before);
    Office.context.document.url = "file:///E:/Other.xlsx";
    expect(await compareExcelNativeCommittedResult(session)).toBe("unavailable");
  });
});
describe("Office 2024 manual native review", () => {
  function manualHost() {
    const h = host();
    (Office.context.requirements.isSetSupported as ReturnType<typeof vi.fn>).mockImplementation((set: string) => set !== 'ExcelApiDesktop');
    // A strict host, not a permissive mock: even touching windows is illegal.
    Object.defineProperty(h.context.workbook.application, 'windows', { get: () => { throw new Error('WINDOW_API_UNSUPPORTED'); } });
    Object.defineProperty(h.context.workbook.application, 'activeWindow', { get: () => { throw new Error('WINDOW_API_UNSUPPORTED'); } });
    return h;
  }
  it.each(['enabled','different_instance','unavailable'] as const)('propagates the scrolling receipt through the actual compatibility staging chain: %s', async synchronousScrolling => {
    const h = manualHost();
    layout.prepare.mockResolvedValue({token:'T'.repeat(43),layoutVersion:3});
    layout.arrange.mockResolvedValue({status:'arranged',layoutVersion:3,copySide:'left',originalSide:'right',synchronousScrolling});
    const session = await stageExcelNativeComparison(plan(), ['Ledger']);
    expect(session.windowLayout).toBe('arranged');
    expect(session.synchronousScrolling).toBe(synchronousScrolling);
    expect(h.writes).toEqual([]);
  });
  it('prepares before opening, arranges only after opening and leaves all original cell setters untouched', async () => {
    const h = manualHost(), order: string[] = [], ticket = {token:'T'.repeat(43),layoutVersion:3};
    layout.prepare.mockImplementation(async () => { order.push('prepare'); return ticket; });
    h.createWorkbook.mockImplementationOnce(async () => { order.push('open'); });
    layout.arrange.mockImplementation(async () => { order.push('arrange'); return {status:'arranged',layoutVersion:3,copySide:'left',originalSide:'right',synchronousScrolling:'enabled'}; });
    const session = await stageExcelNativeComparison(plan(), ['Ledger']);
    expect(order).toEqual(['prepare','open','arrange']);
    expect(layout.prepare).toHaveBeenCalledWith(session.workbookIdentity, 'Ledger', session.copySheetName);
    expect(layout.arrange).toHaveBeenCalledWith(ticket);
    expect(session.windowLayout).toBe('arranged'); expect(h.writes).toEqual([]);
  });
  it('layout failure keeps a real copy available and explicitly records manual fallback', async () => {
    const h = manualHost(); layout.prepare.mockResolvedValue({token:'T'.repeat(43)});
    layout.arrange.mockRejectedValue(new Error('window ambiguous'));
    const session = await stageExcelNativeComparison(plan(), ['Ledger']);
    expect(h.createWorkbook).toHaveBeenCalledOnce(); expect(layout.arrange).toHaveBeenCalledOnce();
    expect(session.windowLayout).toBe('manual'); expect(h.writes).toEqual([]);
  });
  it.each(['prepare','arrange'] as const)('outdated %s service cannot produce a false arranged result or write the original', async when => {
    const h = manualHost();
    layout.prepare.mockResolvedValue({token:'T'.repeat(43),layoutVersion:3});
    layout[when].mockRejectedValue(new ExcelReviewLayoutVersionError());
    const session = await stageExcelNativeComparison(plan(),['Ledger']);
    expect(session.windowLayout).toBe('service_outdated');
    expect(h.createWorkbook).toHaveBeenCalledOnce(); expect(h.writes).toEqual([]);
    if (when === 'prepare') expect(layout.arrange).not.toHaveBeenCalled();
  });
  it.each(['prepare','arrange'] as const)('source changes during %s cannot produce a ready review', async when => {
    const h = manualHost(); layout.prepare.mockResolvedValue({token:'T'.repeat(43)});
    layout[when].mockImplementation(async () => {
      h.original.getRange('B2').values = [[8]];
      return when === 'prepare' ? {token:'T'.repeat(43)} : {status:'arranged'};
    });
    await expect(stageExcelNativeComparison(plan(), ['Ledger'])).rejects.toThrow('原表在审阅期间已变化');
    if (when === 'prepare') expect(h.createWorkbook).not.toHaveBeenCalled();
    expect(h.writes).toEqual(['original:B2']);
  });
  it("projects the plan into an actual opened workbook, without touching window APIs, original setters, network or storage", async () => {
    const h = manualHost(), fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
    expect(excelNativeReviewMode()).toBe('manual');
    const p = plan([{ type:'writeFormulas',sheet:'Ledger',range:'C2',formulas:[['=B2+1']] }]);
    const session = await stageExcelNativeComparison(p, ['Ledger']);
    expect(session.mode).toBe('manual'); expect(h.createWorkbook).toHaveBeenCalledOnce();
    const files = unzipSeed(h.createWorkbook.mock.calls[0][0]);
    expect(files['xl/worksheets/sheet1.xml']).toContain('<f>B2+1</f>');
    expect(files['xl/styles.xml']).toContain('FFE3F1E8');
    expect(files['xl/workbook.xml']).toContain('fullCalcOnLoad="1"');
    expect(h.writes).toEqual([]); expect(fetch).not.toHaveBeenCalled();
    expect(h.original.getRange('C2').values).toEqual([[99]]);
    expect(session.copyEvidence).toBe(''); expect(session.actionEvidence).toBe('');
    expect(await compareExcelNativeCommittedResult(session)).toBe('unavailable');
    expect(await closeExcelNativeComparison(session)).toBe('closed');
    expect(session.recoveryCopyPath).toBe('E:/recovery/review-copy.xlsx');
    expect(layout.close).toHaveBeenCalledWith(session.workbookIdentity, 'Ledger', session.copySheetName);
  });
  it("maps the local bridge's close receipt onto the shared outcome without ever faking a close", async () => {
    const h = manualHost();
    const session = await stageExcelNativeComparison(plan(), ['Ledger']);
    for (const [status, copyClosed, expected] of [
      ['closed', true, 'closed'],
      ['already_closed', true, 'already_closed'],
      ['close_unconfirmed', false, 'close_unconfirmed'],
      ['user_saved', false, 'user_saved'],
      ['identity_lost', false, 'identity_lost'],
      ['ambiguous', false, 'ambiguous'],
      ['unsupported', false, 'unsupported'],
      ['failed', false, 'failed'],
      // 自相矛盾的收据（说关掉了但 copyClosed=false）一律按失败处理，绝不宣称已关闭。
      ['closed', false, 'failed'],
      ['already_closed', false, 'failed']
    ] as const) {
      layout.close.mockResolvedValueOnce({ status, layoutVersion: 4, recoveryCopyPath: "E:/recovery/review-copy.xlsx", copyClosed });
      expect(await closeCompatReviewCopy(session)).toBe(expected);
    }
    // 桥不可达 / 服务过旧 ⇒ 没有可核实的答复，降级为手动关闭。
    layout.close.mockRejectedValueOnce(new Error('offline'));
    expect(await closeCompatReviewCopy(session)).toBe('unsupported');
    expect(h.writes).toEqual([]);   // 关不掉副本不改写原表，也不产生任何原表写入
    layout.close.mockResolvedValueOnce({ status: 'closed', layoutVersion: 4, recoveryCopyPath: 'E:/recovery/review-copy.xlsx', copyClosed: true });
    const callsBefore = layout.close.mock.calls.length;
    expect(await closeExcelNativeComparison(session)).toBe('closed');
    expect(layout.close.mock.calls.length).toBe(callsBefore + 1);

  });
  /*
   * 第三项判断：**副本身份可信 + 没被另存**都成立时，用户仍可能在副本上敲过字。
   * 关掉之前必须能证明"副本还是我们写入后的样子"，否则不关 —— 静默丢弃别人的改动
   * 和静默导入别人的改动一样不可接受。
   */
  it("副本被人工修改后结束审阅 ⇒ 不关副本、报 extra_edits，修改不被静默丢弃", async () => {
    const h = host(), session = await stageExcelNativeComparison(plan(), ["Ledger"]);
    const copyWindow = h.windows[1];
    h.writes.length = 0;
    copyWindow.activeWorksheet.getRange("B2").values = [["HUMAN_EDIT"]];  // 用户在副本上改了一格
    expect(await closeExcelNativeComparison(session)).toBe('extra_edits');
    // 关闭一旦发生，用户刚敲进去的内容就没了 ⇒ 这里必须一个关闭调用都没有。
    expect(copyWindow.close).not.toHaveBeenCalled();
    expect(h.windows[1]).toBe(copyWindow);
    // 副本里的人工修改**不导入原表**：原表一格未写。
    expect(h.writes.some((write: string) => write.startsWith('original:'))).toBe(false);
    expect(h.original.getRange("B2").values).toEqual([[7]]);
  });
  it("副本读数拿不到 ⇒ edits_unverified，同样不关副本（不把'没读到'当成'没改过'）", async () => {
    const h = host(), session = await stageExcelNativeComparison(plan(), ["Ledger"]);
    const copyWindow = h.windows[1];
    Object.defineProperty(copyWindow.activeWorksheet, 'getUsedRangeOrNullObject',
      { get: () => () => { throw new Error('READ_FAILED'); } });
    expect(await closeExcelNativeComparison(session)).toBe('edits_unverified');
    expect(copyWindow.close).not.toHaveBeenCalled();
  });
  it("副本没被改过 ⇒ 正常关闭（正向锚点：前两条不是'永远拒绝关闭'）", async () => {
    const h = host(), session = await stageExcelNativeComparison(plan(), ["Ledger"]);
    const copyWindow = h.windows[1];
    expect(await closeExcelNativeComparison(session)).toBe('closed');
    expect(copyWindow.close).toHaveBeenCalledOnce();
    expect(h.windows).toHaveLength(1);   // 副本窗口确实消失了，只剩原表
  });
  it("副本被人工修改后确认写回 ⇒ 阻断写回，绝不导入副本里的修改", async () => {
    const h = host(), p = plan(), session = await stageExcelNativeComparison(p, ["Ledger"]);
    h.windows[1].activeWorksheet.getRange("B2").values = [["HUMAN_EDIT"]];
    await expect(assertExcelNativeComparisonUnchanged(h.context as never, session, ["Ledger"], p))
      .rejects.toThrow("副本数据或公式已变化");
    expect(h.writes.some((write: string) => write.startsWith('original:'))).toBe(false);
    expect(h.original.getRange("B2").values).toEqual([[7]]);
  });
  it.each(['original', 'authorization', 'plan', 'workbook', 'sheet replacement'])("manual save still blocks %s drift without any target write", async reason => {
    const h = manualHost(), p = plan(), session = await stageExcelNativeComparison(p, ['Ledger']);
    let scope = ['Ledger'];
    if (reason === 'original') h.original.getRange('B2').values = [[8]];
    if (reason === 'authorization') scope = ['Other'];
    if (reason === 'plan') p.actions = [{type:'clearRange',sheet:'Ledger',range:'C2',applyTo:'contents'}];
    if (reason === 'workbook') Office.context.document.url = 'file:///E:/Different.xlsx';
    if (reason === 'sheet replacement') h.original.id = 'replacement';
    const count = h.writes.length;
    await expect(assertExcelNativeComparisonUnchanged(h.context as never, session, scope, p)).rejects.toThrow();
    expect(h.writes).toHaveLength(count);
  });
  it("keeps actual private plan values in the createWorkbook argument only, never fetch or browser persistence", async () => {
    const h = manualHost(), p = plan(), fetch = vi.fn(), setItem = vi.fn();
    vi.stubGlobal('fetch', fetch);
    vi.stubGlobal('localStorage', {setItem, getItem: vi.fn()});
    vi.stubGlobal('sessionStorage', {setItem, getItem: vi.fn()});
    expect(JSON.stringify(p)).toContain('LOCAL_PRIVATE_SENTINEL');
    await stageExcelNativeComparison(p, ['Ledger']);
    expect(unzipSeed(h.createWorkbook.mock.calls[0][0])['xl/worksheets/sheet1.xml']).toContain('LOCAL_PRIVATE_SENTINEL');
    expect(fetch).not.toHaveBeenCalled(); expect(setItem).not.toHaveBeenCalled();
    expect(h.writes).toEqual([]);
  });
  it("rejects a source edit during durable registration at the manual executor's last write boundary", async () => {
    const h = manualHost(), p = plan(), session = await stageExcelNativeComparison(p, ['Ledger']);
    const order: string[] = [];
    await expect(executePlan(p, {
      authorizedSheetNames: ['Ledger'],
      beforeFirstWrite: async context => { order.push('guard'); await assertExcelNativeComparisonUnchanged(context, session, ['Ledger'], p); },
      recovery: {startActionIndex: 0, onBeforeAction: async () => {
        order.push('registered'); h.original.getRange('B2').values = [[100]];
      }, onActionCompleted: vi.fn()}
    })).rejects.toThrow('原表在审阅期间已变化');
    expect(order).toEqual(['registered', 'guard']);
    expect(h.writes).toEqual(['original:B2']); // simulated human edit, zero plan writes
    expect(h.original.getRange('C2').values).toEqual([[99]]);
  });
  it("rechecks the original after the copy-open promise, and never treats an open failure as a ready review", async () => {
    const h = manualHost(); h.createWorkbook.mockImplementationOnce(async () => { h.original.getRange('B2').values = [[8]]; });
    await expect(stageExcelNativeComparison(plan(), ['Ledger'])).rejects.toThrow('原表在审阅期间已变化');
    expect(h.writes).toEqual(['original:B2']); // simulated human edit only
    h.createWorkbook.mockRejectedValueOnce(new Error('COPY_OPEN_DENIED'));
    await expect(stageExcelNativeComparison(plan(), ['Ledger'])).rejects.toThrow('COPY_OPEN_DENIED');
    expect(h.original.getRange('C2').values).toEqual([[99]]);
  });
});
describe("native plan boundaries", () => {
  it("does not widen authorization or accept overlapping actions, mismatched shapes or dangerous/dependent formulas", () => {
    expect(nativeComparisonPlanIssue(plan(), ["Ledger"])).toBeNull();
    expect(nativeComparisonPlanIssue(plan(), ["Ledger", "Other"])).not.toBeNull();
    expect(nativeComparisonPlanIssue(plan(), ["Other"])).not.toBeNull();
    expect(nativeComparisonPlanIssue(plan([plan().actions[0], plan().actions[0]]), ["Ledger"])).toContain("重叠");
    expect(nativeComparisonPlanIssue(plan([{ type: "writeValues", sheet: "Ledger", range: "C2:C3", values: [[1]] }]), ["Ledger"])).not.toBeNull();
    for (const formula of ['=WEBSERVICE("https://example.invalid")', '=Other!A1', '=INDIRECT("A1")', '=SUM(Table1[Amount])', 'text']) {
      expect(nativeComparisonPlanIssue(plan([{ type: "writeFormulas", sheet: "Ledger", range: "C2", formulas: [[formula]] }]), ["Ledger"])).not.toBeNull();
    }
  });
});
