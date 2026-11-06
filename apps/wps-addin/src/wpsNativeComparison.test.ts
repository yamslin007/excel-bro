import { describe, expect, it, vi } from "vitest";
import type { AnalysisPlan } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorkbook, WpsWorksheet } from "./wpsHost";
import { executeWpsPlan, latestWpsUndo, prepareWpsPlan, undoWpsExecution } from "./wpsPlanExecutor";
import {
  changedWpsCellAddresses,
  compareCommittedWpsResultToCopy,
  discardWpsNativeComparison,
  rearrangeWpsComparison,
  stageWpsNativeComparison,
  captureWpsBeforeResult,
  stageWpsBeforeComparison,
  suspendWpsComparison,
  switchWpsBeforeComparison,
  activateWpsBeforeComparison,
  verifyWpsNativeComparison,
  wpsCopyContentVerdict,
  WpsFormulaTrialError
} from "./wpsNativeComparison";
import type { WpsNativeComparison } from "./wpsNativeComparison";

function fakeHost(options: { formulaDisplayOnOriginal?: string; formulaDisplayOnCopy?: string; freshWorkbookWrappers?: boolean } = {}) {
  const originalCell = { value: "原值" as string | null | undefined, formula: "原值" as string | null, color: undefined as number | undefined };
  const copiedCell = { value: "原值" as string | null | undefined, formula: "原值" as string | null, color: undefined as number | undefined };
  let active: WpsWorkbook;
  let copyClosed = false;
  const originalRange = () => ({
    get Value2(): string | null | undefined { return originalCell.value; },
    set Value2(value: string) { originalCell.value = value; originalCell.formula = value; },
    get Formula(): string | null { return originalCell.formula; },
    set Formula(value: string) {
      originalCell.formula = value;
      originalCell.value = value.startsWith("=") ? options.formulaDisplayOnOriginal ?? "/" : value;
    },
    MergeCells: false,
    Interior: {
      get Color() { return originalCell.color; },
      set Color(value: number | undefined) { originalCell.color = value; }
    }
  });
  const copyRange = () => ({
    get Value2(): string | null | undefined { return copiedCell.value; },
    set Value2(value: string) { copiedCell.value = value; copiedCell.formula = value; },
    get Formula(): string | null { return copiedCell.formula; },
    set Formula(value: string) { copiedCell.formula = value; copiedCell.value = value.startsWith("=") ? options.formulaDisplayOnCopy ?? "/" : value; },
    MergeCells: false,
    Interior: {
      get Color() { return copiedCell.color; },
      set Color(value: number | undefined) { copiedCell.color = value; }
    }
  });
  const copySheet = {
    Name: "Sheet1", ProtectContents: false,
    Range: vi.fn((...args: unknown[]) => {
      if (args.length !== 1) throw new Error("Too many parameters.");
      return copyRange();
    })
  };
  const copy: WpsWorkbook = {
    Name: "Book1", ReadOnly: false,
    Worksheets: { Count: 1, Item: () => copySheet },
    Windows: { Count: 1, Item: () => ({ Caption: "Book1" }) },
    Activate: () => { active = copy; },
    Close: vi.fn((saveChanges?: boolean) => {
      if (saveChanges !== false) throw new Error("副本不得保存");
      copyClosed = true;
      active = original;
      return true;
    })
  };
  const originalSheet = {
    Name: "Sheet1", ProtectContents: false,
    Range: vi.fn((...args: unknown[]) => {
      if (args.length !== 1) throw new Error("Too many parameters.");
      return originalRange();
    }),
    Copy: vi.fn((...args: unknown[]) => {
      if (args.length !== 0) throw new Error("Too many parameters.");
      copiedCell.value = originalCell.value;
      copiedCell.formula = originalCell.formula;
      active = copy;
    })
  };
  const original: WpsWorkbook = {
    Name: "测试.xlsx", FullName: "C:\\Data\\测试.xlsx", Path: "C:\\Data", ReadOnly: false,
    Worksheets: { Count: 1, Item: () => originalSheet },
    ActiveSheet: originalSheet,
    Activate: () => { active = original; }
  };
  active = original;
  const storage = new Map<string, unknown>();
  const compare = vi.fn((...args: unknown[]) => {
    if (args.length !== 1 || args[0] !== "Book1") throw new Error("并排参数不符");
    if (active !== original) throw new Error("必须从原表发起并排");
    return true;
  });
  const application = {
    Version: "12.0",
    get ActiveWorkbook() { return options.freshWorkbookWrappers ? { ...active } : active; },
    Windows: { Count: 2, Item: () => ({ Caption: "Book1" }), CompareSideBySideWith: compare },
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, value: unknown) => { storage.set(key, value); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as WpsApplication;
  const plan: AnalysisPlan = {
    id: "native-compare", title: "更新单元格", summary: "测试副本与原表分离",
    assumptions: [], warnings: [], acceptanceCriteria: [],
    actions: [{ type: "writeValues", sheet: "Sheet1", range: "A1", values: [["结果"]] }]
  };
  return {
    application, plan, original, originalSheet, copy, compare,
    setActiveWorkbook: (workbook: WpsWorkbook) => { active = workbook; },
    originalCell, copiedCell, copyClosed: () => copyClosed
  };
}

type MultiMatrix = (string | number | null)[][];

const MULTI_MEMBERS: MultiMatrix = [
  ["会员编号", "会员", "等级"],
  ["7", "甲", "金卡"],
  ["007", "乙", "银卡"],
  ["", "丙", "铜卡"],
  ["9", "丁", "金卡"],
  ["404", "戊", "金卡"]
];
const MULTI_ORDERS: MultiMatrix = [
  ["会员编号", "金额"],
  ["7", 100],
  ["007", 200],
  ["9", 300],
  ["404", 400]
];
const TARGET_MATRIX: MultiMatrix = [
  ["会员等级", "金额合计", "明细"],
  ["未匹配", 400, ""],
  ["金卡", 400, ""],
  ["银卡", 200, ""],
  ["空键数", 1, ""],
  ["重复键数", 1, ""]
];

/**
 * **多表授权**的宿主替身：三张已授权表（两张只读来源 + 一张写入目标），
 * `Worksheet.Copy()` 真的只复制被调用的那一张表。
 */
function multiSheetHost(transformWrittenValue: (value: unknown, format: string) => unknown = (value) => value) {
  const names = ["会员表", "订单表", "汇总表"];
  const grids: Record<string, MultiMatrix> = {
    会员表: MULTI_MEMBERS.map((row) => [...row]),
    订单表: MULTI_ORDERS.map((row) => [...row]),
    汇总表: []
  };
  const copiedSheets: string[] = [];
  const copyBooks: WpsWorkbook[] = [];
  const copyTabs = new WeakMap<WpsWorkbook, WpsWorksheet[]>();
  const storage = new Map<string, unknown>();
  let active: WpsWorkbook;
  let copyClosed = false;

  const cellFormats = new WeakMap<MultiMatrix, Map<string, string>>();
  function formatsFor(grid: MultiMatrix) {
    let formats = cellFormats.get(grid);
    if (!formats) { formats = new Map(); cellFormats.set(grid, formats); }
    return formats;
  }

  function rectRange(read: () => MultiMatrix, rect: { startRow: number; startColumn: number; endRow: number; endColumn: number }) {
    const rows = rect.endRow - rect.startRow + 1;
    const columns = rect.endColumn - rect.startColumn + 1;
    const slice = () =>
      Array.from({ length: rows }, (_, rowOffset) =>
        Array.from({ length: columns }, (_, columnOffset) =>
          read()[rect.startRow - 1 + rowOffset]?.[rect.startColumn - 1 + columnOffset] ?? null
        )
      );
    const range: Record<string, unknown> = {
      Row: rect.startRow, Column: rect.startColumn,
      Rows: { Count: rows },
      Columns: { Count: columns },
      MergeCells: false,
      Interior: { Color: 0 }
    };
    range.Item = (r: number, c: number) => rectRange(read, {
      startRow: rect.startRow + r - 1, endRow: rect.startRow + r - 1,
      startColumn: rect.startColumn + c - 1, endColumn: rect.startColumn + c - 1
    });
    Object.defineProperty(range, "NumberFormat", {
      get: () => formatsFor(read()).get(`${rect.startRow},${rect.startColumn}`) ?? "General",
      set: (format: string) => {
        for (let r = rect.startRow; r <= rect.endRow; r++) {
          for (let c = rect.startColumn; c <= rect.endColumn; c++) formatsFor(read()).set(`${r},${c}`, format);
        }
      }
    });
    const store = (incoming: unknown) => {
      const normalized = Array.isArray(incoming) ? (incoming as unknown[][]) : [[incoming]];
      const grid = read();
      for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
        const rowIndex = rect.startRow - 1 + rowOffset;
        if (!grid[rowIndex]) grid[rowIndex] = [];
        for (let columnOffset = 0; columnOffset < columns; columnOffset += 1) {
          const value = normalized[rowOffset]?.[columnOffset];
          const format = formatsFor(grid).get(`${rowIndex + 1},${rect.startColumn + columnOffset}`) ?? "General";
          if (value !== undefined) grid[rowIndex][rect.startColumn - 1 + columnOffset] = transformWrittenValue(value, format) as never;
        }
      }
    };
    Object.defineProperties(range, {
      Value2: { get: () => slice(), set: store, enumerable: true },
      Formula: { get: () => slice(), set: store, enumerable: true }
    });
    return range as unknown as WpsRange;
  }

  function parseA1(address: string) {
    const [start, end = start] = address.split(":");
    const cell = (text: string) => {
      const match = /^([A-Z]+)(\d+)$/i.exec(text.trim());
      if (!match) throw new Error(`无法解析区域：${address}`);
      let column = 0;
      for (const char of match[1].toUpperCase()) column = column * 26 + (char.charCodeAt(0) - 64);
      return { column, row: Number(match[2]) };
    };
    const from = cell(start);
    const to = cell(end);
    return { startRow: from.row, startColumn: from.column, endRow: to.row, endColumn: to.column };
  }

  function buildSheet(name: string, read: () => MultiMatrix, onCopy?: (before?: WpsWorksheet, after?: WpsWorksheet) => void): WpsWorksheet {
    const sheet: Record<string, unknown> = {
      Name: name,
      CodeName: `Code_${name}`,
      ProtectContents: false,
      Range: (address: unknown) => rectRange(read, parseA1(String(address)))
    };
    Object.defineProperty(sheet, "UsedRange", {
      get: () =>
        rectRange(read, {
          startRow: 1,
          startColumn: 1,
          endRow: Math.max(1, read().length),
          endColumn: Math.max(1, ...read().map((row) => row.length))
        }),
      enumerable: true
    });
    if (onCopy) sheet.Copy = onCopy;
    return sheet as unknown as WpsWorksheet;
  }

  const sheets = new Map<string, WpsWorksheet>();
  function sheet(name: string): WpsWorksheet {
    const cached = sheets.get(name);
    if (cached) return cached;
    const built = buildSheet(name, () => grids[name] ?? [], (_before, after) => {
      copiedSheets.push(name);
      const snapshot = (grids[name] ?? []).map((row) => [...row]);
      const copyName = name;
      const copySheet = buildSheet(copyName, () => snapshot);
      if (after) {
        const destination = copyBooks.find(book => copyTabs.get(book)?.includes(after));
        if (!destination) throw new Error('unknown copy destination');
        copyTabs.get(destination)!.push(copySheet);
        copySheet.Activate = () => { active = destination; destination.ActiveSheet = copySheet; };
        copySheet.Activate();
        return;
      }
      const tabs = [copySheet];
      const copyWindow = {Caption: 'Book1', Visible: true};
      const copy: WpsWorkbook = {
        Name: "Book1",
        FullName: "",
        Path: "",
        ReadOnly: false,
        Worksheets: { get Count() {return tabs.length;}, Item: (i: number | string) => typeof i === 'number' ? tabs[i-1] : tabs.find(s => s.Name === i)! },
        ActiveSheet: copySheet,
        Windows: { Count: 1, Item: () => copyWindow },
        Activate: () => {
          active = copy;
        },
        Close: (saveChanges?: boolean) => {
          if (saveChanges !== false) throw new Error("副本不得保存");
          copyClosed = true;
          if (active === copy) active = original;
          return true;
        }
      };
      copyBooks.push(copy); copyTabs.set(copy, tabs);
      copySheet.Activate = () => { active = copy; copy.ActiveSheet = copySheet; };
      active = copy;
    });
    sheets.set(name, built);
    built.Activate = () => { active = original; original.ActiveSheet = built; };
    return built;
  }

  const original: WpsWorkbook = {
    Name: "会员.xlsx",
    FullName: "C:\\Data\\会员.xlsx",
    Path: "C:\\Data",
    ReadOnly: false,
    // 真宿主对不存在的名字/越界序号会抛而不是返回 undefined；替身按"读不到就是
    // 授权前提不成立"（`authorizeWpsSheets` 会抛）处理，故整体收窄成集合类型。
    Worksheets: {
      get Count() {
        return names.length;
      },
      Item: (key: number | string) => {
        if (typeof key === "string" && !/^\d+$/.test(key)) {
          return names.includes(key) ? sheet(key) : undefined;
        }
        const name = names[Number(key) - 1];
        return name ? sheet(name) : undefined;
      }
    } as unknown as WpsWorkbook["Worksheets"],
    ActiveSheet: sheet("会员表"),
    Windows: { Count: 1, Item: () => ({ Caption: "会员.xlsx" }) },
    Activate: () => {
      active = original;
    }
  };
  active = original;

  const compare = vi.fn(() => true);
  const application = {
    Version: "12.0",
    get ActiveWorkbook() {
      return active;
    },
    Windows: { Count: 1, Item: () => ({ Caption: "Book1" }), CompareSideBySideWith: compare, SyncScrollingSideBySide: false },
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, value: unknown) => {
        storage.set(key, value);
      }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as unknown as WpsApplication;

  const plan: AnalysisPlan = {
    id: "native-compare-multi",
    title: "会员等级汇总",
    summary: "两表关联汇总后写到汇总表",
    assumptions: [],
    warnings: [],
    acceptanceCriteria: [],
    actions: [{ type: "writeValues", sheet: "汇总表", range: "A1:C6", values: TARGET_MATRIX.map((row) => [...row]) }]
  };

  return {
    application,
    plan,
    copiedSheets: () => [...copiedSheets],
    matrix: (name: string) => (grids[name] ?? []).map((row) => [...row]),
    copyClosed: () => copyClosed
  };
}

describe("WPS 原生双窗口的隔离审阅", () => {
  it('三张目标表只建立一个副本；往返切表不重排，非当前子表有人工修改就保留', () => {
    const host = multiSheetHost(), app = host.application;
    const original = app.ActiveWorkbook!;
    original.Worksheets.Item(3).Range!('A1').Value2 = '旧汇总';
    const plan = {...host.plan, actions: ['会员表', '订单表', '汇总表'].map(sheet => ({
      type: 'writeValues' as const, sheet, range: 'A1', values: [['新值']]}))};
    const before = captureWpsBeforeResult(app, prepareWpsPlan(app, [1,2,3], plan));
    const oldValues = [1,2,3].map(i => original.Worksheets.Item(i).Range!('A1').Value2);
    [1,2,3].forEach(i => { original.Worksheets.Item(i).Range!('A1').Value2 = '新值'; });
    const session = stageWpsBeforeComparison(app, before, undefined, '会员表');
    const copy = session.copy;
    expect(copy.Worksheets.Count).toBe(3);
    expect([1,2,3].map(i => copy.Worksheets.Item(i).Name)).toEqual(['会员表', '订单表', '汇总表']);
    expect([1,2,3].map(i => copy.Worksheets.Item(i).Range!('A1').Value2)).toEqual(oldValues);
    activateWpsBeforeComparison(app, session);
    for (const name of ['汇总表', '会员表', '订单表']) {
      expect(switchWpsBeforeComparison(app, session, name)).toBe(true);
      expect(copy.ActiveSheet!.Name).toBe(name);
      expect(original.ActiveSheet!.Name).toBe(name);
      expect(session.copy).toBe(copy);
    }
    expect(app.Windows!.CompareSideBySideWith).toHaveBeenCalledTimes(1);
    app.Windows!.BreakSideBySide = vi.fn(() => true);
    expect(suspendWpsComparison(app, session)).toBeNull();
    expect(switchWpsBeforeComparison(app, session, '汇总表')).toBe(true);
    expect(app.Windows!.CompareSideBySideWith).toHaveBeenCalledTimes(2);
    expect(host.copiedSheets()).toHaveLength(3);
    expect(wpsCopyContentVerdict(session)).toBe('unchanged');
    copy.Worksheets.Item(1).Range!('B1').Value2 = '副本人工修改';
    expect(discardWpsNativeComparison(session)).toBe('extra_edits');
    expect(host.copyClosed()).toBe(false);
    expect([1,2,3].map(i => original.Worksheets.Item(i).Range!('A1').Value2)).toEqual([[['新值']],[['新值']],[['新值']]]);
  });
  it('执行前保存整表上下文，执行后其他区域变化不会混入副本', () => {
    const host = multiSheetHost();
    const sheet = host.application.ActiveWorkbook!.Worksheets.Item(3);
    sheet.Range!('A1:B1').Value2 = [['旧目标', '旧表头']];
    const plan = {...host.plan, actions:[{type:'writeValues' as const, sheet:'汇总表',range:'A1',values:[['新目标']]}]};
    const before = captureWpsBeforeResult(host.application, prepareWpsPlan(host.application, [1,2,3], plan));
    expect(before.contextSheets?.['汇总表'].values).toEqual([['旧目标','旧表头']]);
    sheet.Range!('A1:B1').Value2 = [['新目标', '后来改的表头']];
    const staged = stageWpsBeforeComparison(host.application, before);
    expect(staged.snapshotNotice).toBeUndefined();
    expect(staged.snapshotScope).toContain('不作为历史对比依据');
    expect(staged.copy.Worksheets.Item(1).Range!('A1:B1').Value2).toEqual([['旧目标','旧表头']]);
    expect(sheet.Range!('A1:B1').Value2).toEqual([['新目标','后来改的表头']]);
    staged.copy.Worksheets.Item(1).Range!('B1').Value2 = '副本人工修改';
    expect(discardWpsNativeComparison(staged)).toBe('extra_edits');
    expect(host.copyClosed()).toBe(false);
  });

  it('退出隐藏而不关闭，重开复用同一副本', () => {
    const host = fakeHost();
    const window = {Caption:'Book1', Visible:true};
    host.copy.Windows = {Count:1, Item:()=>window};
    host.application.Windows!.BreakSideBySide = vi.fn(() => true);
    const before = captureWpsBeforeResult(host.application, prepareWpsPlan(host.application, [1], host.plan));
    const staged = stageWpsBeforeComparison(host.application, before);
    expect(suspendWpsComparison(host.application,staged)).toBeNull();
    expect(window.Visible).toBe(false);
    expect(host.copyClosed()).toBe(false);
    expect(rearrangeWpsComparison(host.application, staged)).toBe('arranged');
    expect(window.Visible).toBe(true);
    expect(host.originalSheet.Copy).toHaveBeenCalledTimes(1);
    expect(host.originalCell.value).toBe('原值');
  });
  it("数字文本在副本、原表写回与撤销中保持类型和原数字格式", async () => {
    const host = multiSheetHost((value, format) =>
      typeof value === "string" && /^\d+$/.test(value) && format !== "@" ? Number(value) : value);
    const sheet = host.application.ActiveWorkbook!.Worksheets.Item(3);
    const cell = sheet.Range!("D8");
    cell.NumberFormat = "@";
    cell.Value2 = "0009";
    cell.NumberFormat = "General";
    host.plan.actions = [{ type: "writeValues", sheet: "汇总表", range: "D8:F8", values: [["7", "007", 400]] }];
    const prepared = prepareWpsPlan(host.application, [1, 2, 3], host.plan);
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(staged.copyAfter[0].values).toEqual([["7", "007", 400]]);
    expect(host.matrix("汇总表")[7][3]).toBe("0009");
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(host.matrix("汇总表")[7].slice(3, 6)).toEqual(["7", "007", 400]);
    expect(sheet.Range!("D8").NumberFormat).toBe("General");
    await undoWpsExecution(host.application, record.executionId);
    expect(host.matrix("汇总表")[7].slice(3, 6)).toEqual(["0009", null, null]);
    expect(sheet.Range!("D8").NumberFormat).toBe("General");
  });
  it.each([
    { expected: "007", actual: 7, diagnostic: "计划为文本，回读为数值" },
    { expected: null, actual: "", diagnostic: "计划为空单元格，回读为空字符串" },
    { expected: "private-before", actual: "private-after", diagnostic: "计划为文本，回读为文本（内容不同）" }
  ])("回读失败定位到真实坐标，不泄露内容或放松拒绝: $diagnostic", ({ expected, actual, diagnostic }) => {
    const host = multiSheetHost((value) => value === expected ? actual : value);
    host.plan.actions = [{ type: "writeValues", sheet: "汇总表", range: "C4:D5", values: [[1, 2], [3, expected]] }];
    const prepared = prepareWpsPlan(host.application, [1, 2, 3], host.plan);
    let message = "";
    try { stageWpsNativeComparison(host.application, prepared); } catch (error) {
      message = (error as Error).message;
    }
    expect(message).toContain(`首个差异 D5：${diagnostic}`);
    expect(message).toContain("原表未写入");
    expect(message).not.toContain("private-before");
    expect(message).not.toContain("private-after");
    expect(host.matrix("汇总表")).toEqual([]);
    expect(host.copyClosed()).toBe(true);
    expect(latestWpsUndo(host.application)).toBeNull();
  });
  it("只标真正变化的同坐标单元格，公式变化即使显示值相同也标记", () => {
    const host = fakeHost();
    const action = prepareWpsPlan(host.application, [1], host.plan).actions[0];
    const twoRows = {
      ...action,
      range: "A1:A2", rowCount: 2, cellCount: 2,
      rect: { startRow: 1, endRow: 2, startColumn: 1, endColumn: 1 },
      beforeValues: [["旧"], ["相同"]],
      beforeFormulas: [["旧"], ["相同"]],
      beforeFingerprint: "", afterFingerprint: null, appliedRows: 2
    };
    expect(changedWpsCellAddresses(twoRows, {
      values: [["新"], ["相同"]], formulas: [["新"], ["相同"]], valuesCaptured: true
    })).toEqual(["A1"]);
    expect(changedWpsCellAddresses(twoRows, {
      values: [["旧"], ["相同"]], formulas: [["=1+1"], ["相同"]], valuesCaptured: true
    })).toEqual(["A1"]);
  });
  it("左原表不写、右副本回读后只标记变化位置，保存前仍待确认", () => {
    const host = fakeHost();
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(host.originalSheet.Copy).toHaveBeenCalledWith();
    expect(host.originalCell.value).toBe("原值");
    expect(host.originalCell.color).toBeUndefined();
    expect(latestWpsUndo(host.application)).toBeNull();
    expect(host.copiedCell.value).toBe("结果");
    expect(host.copiedCell.color).toBe(0xE8F5EA);
    expect(staged.changedCells).toEqual(["A1"]);
    expect(staged.reviews[0].status).toBe("applied");
    expect(host.compare).toHaveBeenCalledWith("Book1");
    verifyWpsNativeComparison(host.application, staged);
    expect(host.originalCell.value).toBe("原值");
    discardWpsNativeComparison(staged);
    expect(host.copyClosed()).toBe(true);
    expect(host.copy.Close).toHaveBeenCalledWith(false);
    expect(host.originalCell.value).toBe("原值");
  });

  it("空白单格的 null 是已取得的原值，允许副本预览但不写原表", () => {
    const host = fakeHost();
    host.originalCell.value = null;
    host.originalCell.formula = null;
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    expect(prepared.actions[0].beforeValuesCaptured).toBe(true);
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(staged.originalBefore[0].values).toEqual([[null]]);
    expect(staged.originalBefore[0].valuesCaptured).toBe(true);
    expect(host.originalCell.value).toBeNull();
    expect(host.copiedCell.value).toBe("结果");
    discardWpsNativeComparison(staged);
    expect(host.originalCell.value).toBeNull();
  });

  it("缺失 Value2 仍拒绝审阅，不把未知值当作空白", () => {
    const host = fakeHost();
    host.originalCell.value = undefined;
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    expect(prepared.actions[0].beforeValuesCaptured).toBe(false);
    expect(() => stageWpsNativeComparison(host.application, prepared)).toThrow("宿主未提供原表目标区域的显示值");
    expect(host.originalSheet.Copy).not.toHaveBeenCalled();
  });

  it("同一本已保存工作簿返回新的 JSAPI 包装对象时仍能生成、复核并提交", async () => {
    const host = fakeHost({ freshWorkbookWrappers: true });
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(host.originalSheet.Copy).toHaveBeenCalledOnce();
    expect(host.originalCell.value).toBe("原值");
    verifyWpsNativeComparison(host.application, staged);
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(host.originalCell.value).toBe("结果");
    expect(compareCommittedWpsResultToCopy(staged)).toBe("matched");
  });

  it("同名同内容但不同路径的工作簿不应被误当成原表", () => {
    const host = fakeHost({ freshWorkbookWrappers: true });
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    host.setActiveWorkbook({
      ...host.original,
      FullName: "D:\\Other\\测试.xlsx",
      Path: "D:\\Other"
    });
    expect(() => stageWpsNativeComparison(host.application, prepared))
      .toThrow("当前活动工作簿已不是预览时的原表");
    expect(host.originalSheet.Copy).not.toHaveBeenCalled();
    expect(host.originalCell.value).toBe("原值");
  });

  it("副本被保存为磁盘文件后拒绝当作未保存的审阅副本写回", () => {
    const host = fakeHost({ freshWorkbookWrappers: true });
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    host.copy.Path = "C:\\Data";
    host.copy.FullName = "C:\\Data\\Book1.xlsx";
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow("审阅副本身份已变化");
    expect(host.originalCell.value).toBe("原值");
  });

  it("原表在并排审阅期间改变时拒绝写回", () => {
    const host = fakeHost();
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    host.originalCell.value = "他人新值";
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow("原表或授权范围在预览后发生变化");
    expect(host.originalCell.value).toBe("他人新值");
  });

  it("原表在生成副本前已变或计划被篡改时，连副本都不生成", () => {
    const host = fakeHost();
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    host.originalCell.value = "预览后外部修改";
    expect(() => stageWpsNativeComparison(host.application, prepared)).toThrow("原表或授权范围在预览后发生变化");
    expect(host.originalSheet.Copy).not.toHaveBeenCalled();
    host.originalCell.value = "原值";
    prepared.actions[0].values[0][0] = "未经预览的值";
    expect(() => stageWpsNativeComparison(host.application, prepared)).toThrow("写入计划或预览已变化");
    expect(host.originalSheet.Copy).not.toHaveBeenCalled();
  });

  it("多表授权不会退化为复制整本工作簿：副本只承载写入目标那一张表", () => {
    /*
     * 判据是"**写入**只落在一张已授权表上"，不是"授权表只有一张"。
     *
     * 两表关联后汇总到第三张表时，授权表里两张是只读来源、一张是写入目标；
     * 把授权张数当成写入张数会把这类计划整片挡在确认写回之外（Excel 端有等价的
     * 显式执行入口，两端就此分叉）。安全性质**没有放宽**：副本仍然只有一张表，
     * 绝不把来源表或任何未授权的表复制出去。
     */
    const host = multiSheetHost();
    const prepared = prepareWpsPlan(host.application, [1, 2, 3], host.plan);
    expect(prepared.authorization.sheets.map((sheet) => sheet.name)).toEqual(["会员表", "订单表", "汇总表"]);
    const staged = stageWpsNativeComparison(host.application, prepared);
    // 复制出去的只有**写入目标**那一张。
    expect(host.copiedSheets()).toEqual(["汇总表"]);
    expect(staged.writeSheet.name).toBe("汇总表");
    expect(staged.copy.Worksheets.Count).toBe(1);
    expect(staged.copy.Worksheets.Item(1).Name).toBe("汇总表");
    // 副本上试写了目标矩阵，原表一格未写；来源表也一格未动。
    expect(staged.copyAfter[0].values).toEqual(TARGET_MATRIX);
    expect(host.matrix("汇总表").flat().every((value) => value === null)).toBe(true);
    expect(host.matrix("会员表")).toEqual(MULTI_MEMBERS);
    expect(host.matrix("订单表")).toEqual(MULTI_ORDERS);
    verifyWpsNativeComparison(host.application, staged);
    discardWpsNativeComparison(staged);
    expect(host.copyClosed()).toBe(true);
    expect(host.matrix("汇总表").flat().every((value) => value === null)).toBe(true);
  });

  it("跨表写入仍拒绝并排审阅：不会把整本工作簿复制成副本", () => {
    const host = multiSheetHost();
    host.plan.actions = [
      { type: "writeValues", sheet: "汇总表", range: "A1:C6", values: TARGET_MATRIX },
      { type: "writeValues", sheet: "会员表", range: "A1:C6", values: TARGET_MATRIX }
    ];
    const prepared = prepareWpsPlan(host.application, [1, 2, 3], host.plan);
    expect(() => stageWpsNativeComparison(host.application, prepared))
      .toThrow("只支持单张工作表上的写入计划");
    expect(host.copiedSheets()).toHaveLength(0);
    expect(host.matrix("汇总表").flat().every((value) => value === null)).toBe(true);
    expect(host.matrix("会员表")).toEqual(MULTI_MEMBERS);
  });

  it("右侧副本被手动改动后拒绝把未确认内容写回", () => {
    const host = fakeHost();
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    host.copiedCell.value = "手工改动";
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow("审阅副本在生成对照后发生变化");
    expect(host.originalCell.value).toBe("原值");
  });

  it("常量单元格在预览后被改动，拒绝写回且新增写入为 0", () => {
    // Constant cell: only the displayed value exists (no formula linkage), so a drift is a real
    // external change and must be intercepted. Which guard fires first is not the point — the
    // assertion only requires that the write-back is rejected and no new write happens.
    const host = fakeHost();
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    const before = host.originalCell.value;
    host.originalCell.value = "他人改的值";
    expect(() => verifyWpsNativeComparison(host.application, staged)).toThrow();
    // 新增写入为 0：原表仍是外部改动的值，没有被覆盖成计划值
    expect(host.originalCell.value).toBe("他人改的值");
    expect(host.originalCell.value).not.toBe(before);
  });

  it("原表公式文本未变但计算值漂移时拒绝写回（公式重算，不是常量改动）", () => {
    // Real formula baseline on the original, established *before* the plan is prepared so the
    // preview snapshot binds the formula. Write the formula through the range setter (the same
    // path a real host uses), giving formula "=1+1" with calculated value "2".
    //
    // 原表侧漂移由两道合法守卫共同保证：`assertOriginalUnchanged`（重新读原表、比对快照 →
    // 「原表或授权范围在预览后发生变化」）在后道复核之前先触发。本用例断言**前一道**的确切
    // 文案，从而针对 `assertOriginalUnchanged` 形成独立判别力；后道快照守卫是同性质的冗余层，
    // 不在此用例中冒充独立证据（拆任一道仍可能被另一道拒绝，见 docs/DEVELOPMENT.md 记录）。
    const host = fakeHost({ formulaDisplayOnOriginal: "2" });
    host.originalSheet.Range("A1").Formula = "=1+1";
    expect(host.originalCell.formula).toBe("=1+1");
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    // 基线：公式文本 + 计算值都拿到了
    expect(host.originalCell.formula).toBe("=1+1");
    expect(host.originalCell.value).toBe("2");
    // 公式文本不变，仅计算值漂移（真实公式重算，非常量编辑）
    host.originalCell.value = "99";
    expect(host.originalCell.formula).toBe("=1+1");
    expect(host.originalCell.value).toBe("99");
    expect(() => verifyWpsNativeComparison(host.application, staged))
      .toThrow("原表或授权范围在预览后发生变化");
    // 新增写入为 0：漂移的计算值保留，未被计划值覆盖
    expect(host.originalCell.value).toBe("99");
  });

  it("副本公式文本未变但计算值漂移时拒绝审阅写回（副本侧指纹绑定公式）", () => {
    // Formula cell on the disposable copy: after staging the formula text stays "=1+1" while the
    // calculated display value drifts. The copy-side check reads this before the original-side
    // snapshot, so it is reached without any original change.
    const host = fakeHost();
    host.plan.actions = [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=1+1"]] }];
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    host.copiedCell.formula = "=1+1";
    host.copiedCell.value = "3"; // formula text unchanged, calculated value drifted
    expect(() => verifyWpsNativeComparison(host.application, staged))
      .toThrow("审阅副本在生成对照后发生变化");
    // 原表未被写入
    expect(host.originalCell.value).toBe("原值");
  });

  it("明确保存后才用原执行器写回，并逐格核对副本与原表结果", async () => {
    const host = fakeHost();
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(host.originalCell.value).toBe("原值");
    verifyWpsNativeComparison(host.application, staged);
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(host.originalCell.value).toBe("结果");
    expect(host.originalCell.color).toBeUndefined();
    expect(compareCommittedWpsResultToCopy(staged)).toBe("matched");
    host.originalCell.value = "计算结果发生偏差";
    expect(compareCommittedWpsResultToCopy(staged)).toBe("different");
  });

  it("公式文本相同但原表计算值不同，明确报告副本与提交结果不一致", async () => {
    const host = fakeHost({ formulaDisplayOnOriginal: "另一计算值" });
    host.plan.actions = [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=1+1"]] }];
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    const staged = stageWpsNativeComparison(host.application, prepared);
    expect(host.originalCell.value).toBe("原值");
    expect(host.copiedCell.formula).toBe("=1+1");
    expect(host.copiedCell.color).toBe(0xE8F5EA);
    verifyWpsNativeComparison(host.application, staged);
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(compareCommittedWpsResultToCopy(staged)).toBe("different");
  });

  it("公式文本回读相同但副本计算值错误时拒绝审阅与写回，并关闭试写副本", () => {
    const host = fakeHost();
    host.plan.id = "wps-map-column-test";
    host.plan.actions = [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=IF(EXACT(B1,\"甲\"),\"正确\",NA())"]] }];
    host.plan.acceptanceCriteria = [{ type: "rangeEquals", sheet: "Sheet1", range: "A1", expected: [["正确"]] }];
    expect(() => stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan)))
      .toThrow(WpsFormulaTrialError);
    expect(host.copyClosed()).toBe(true);
    expect(host.originalCell.value).toBe("原值");
    expect(host.compare).not.toHaveBeenCalled();
  });

  it("公式与本地预期计算值都相同才开放右侧副本审阅", () => {
    const host = fakeHost({ formulaDisplayOnCopy: "正确" });
    host.plan.id = "wps-map-column-test";
    host.plan.actions = [{ type: "writeFormulas", sheet: "Sheet1", range: "A1", formulas: [["=IF(EXACT(B1,\"甲\"),\"正确\",NA())"]] }];
    host.plan.acceptanceCriteria = [{ type: "rangeEquals", sheet: "Sheet1", range: "A1", expected: [["正确"]] }];
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    expect(staged.copyAfter[0].values).toEqual([["正确"]]);
    expect(staged.reviews[0].status).toBe("applied");
    expect(host.originalCell.value).toBe("原值");
    expect(host.compare).toHaveBeenCalledOnce();
  });

  it("宿主并排失败时关闭自己创建的副本且不碰原表", () => {
    const host = fakeHost();
    host.compare.mockReturnValueOnce(false);
    expect(() => stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan)))
      .toThrow("WPS 未能打开原生并排比较");
    expect(host.copyClosed()).toBe(true);
    expect(host.originalCell.value).toBe("原值");
  });
});

/*
 * 「重新排列」：WPS 只能重排**两张表窗口**（CompareSideBySideWith），
 * 不能像 Excel 本地桥那样设置精确宽度。所以这里只断言"窗口已重排 / 未确认"，
 * 并且任何一条路径都**不得**写单元格。
 */
describe("WPS 审阅「重新排列」的能力边界", () => {
  type Host = ReturnType<typeof fakeHost>;
  type BreakIt = (host: Host, staged: WpsNativeComparison) => void;
  function stagedReview() {
    const host = fakeHost();
    const staged = stageWpsNativeComparison(host.application, prepareWpsPlan(host.application, [1], host.plan));
    host.compare.mockClear();
    return { host, staged };
  }
  it("requests the native side-by-side pair and never writes a cell", () => {
    const { host, staged } = stagedReview();
    const before = latestWpsUndo(host.application);
    expect(rearrangeWpsComparison(host.application, staged)).toBe("arranged");
    expect(host.compare).toHaveBeenCalledWith("Book1");
    expect(host.application.Windows!.SyncScrollingSideBySide).toBe(true);
    // 重排只是几何操作：原表未写、副本未改、没有新的撤销记录。
    expect(host.originalCell.value).toBe("原值");
    expect(latestWpsUndo(host.application)).toBe(before);
    expect(host.copyClosed()).toBe(false);
  });
  it("turns an optional synchronous-scrolling failure into a geometry-only result, not a failure", () => {
    const { host, staged } = stagedReview();
    Object.defineProperty(host.application.Windows!, "SyncScrollingSideBySide", {
      configurable: true, set() { throw new Error("不支持同步滚动"); }
    });
    expect(rearrangeWpsComparison(host.application, staged)).toBe("arranged");
    expect(host.originalCell.value).toBe("原值");
  });
  it("also issues the reset when the host offers it", () => {
    const { host, staged } = stagedReview();
    const reset = vi.fn();
    host.application.Windows!.ResetPositionsSideBySide = reset;
    expect(rearrangeWpsComparison(host.application, staged)).toBe("arranged");
    expect(reset).toHaveBeenCalledOnce();
  });
  it("falls back to the copy workbook name when the host exposes no caption", () => {
    const { host, staged } = stagedReview();
    staged.copy.Windows = undefined;
    // 回退用副本工作簿名，而不是编一个窗口标题。
    expect(rearrangeWpsComparison(host.application, staged)).toBe("arranged");
    expect(host.compare).toHaveBeenCalledWith("Book1");
  });
  it.each([
    ["no caption and no name", (_host: Host, staged: WpsNativeComparison) => { staged.copy.Windows = undefined; staged.copy.Name = ""; }],
    ["copy identity drifted", (_host: Host, staged: WpsNativeComparison) => { staged.copy.Name = "Other.xlsx"; }],
    ["original sheet lost", (host: Host, _staged: WpsNativeComparison) => { host.original.Worksheets = { Count: 0, Item: () => host.originalSheet }; }],
    ["host lacks side-by-side", (host: Host, _staged: WpsNativeComparison) => { (host.application.Windows as { CompareSideBySideWith?: unknown }).CompareSideBySideWith = undefined; }],
    ["side-by-side refused", (host: Host, _staged: WpsNativeComparison) => { host.compare.mockReturnValueOnce(false); }],
    ["side-by-side threw", (host: Host, _staged: WpsNativeComparison) => { host.compare.mockImplementationOnce(() => { throw new Error("宿主拒绝"); }); }]
  ] satisfies Array<[string, BreakIt]>)("reports unavailable without touching cells when %s", (_label: string, breakIt: BreakIt) => {
    const { host, staged } = stagedReview();
    breakIt(host, staged);
    expect(rearrangeWpsComparison(host.application, staged)).toBe("unavailable");
    expect(host.originalCell.value).toBe("原值");
    expect(host.copiedCell.value).toBe("结果");
    expect(latestWpsUndo(host.application)).toBeNull();
  });
  it("checks identity after the native call too, so a mid-call swap is not reported as arranged", () => {
    const { host, staged } = stagedReview();
    host.compare.mockImplementationOnce(() => { staged.copy.Name = "Other.xlsx"; return true; });
    expect(rearrangeWpsComparison(host.application, staged)).toBe("unavailable");
  });
});

/*
 * 审阅副本的生命周期：关闭结果**不抛异常**，而是返回共享状态用的取值。
 * 关键判据：
 *  - 关不掉不许说成"已关闭"；
 *  - 用户另存/替换过的副本**不自动丢弃**；
 *  - 副本已由用户关闭 ⇒ 幂等收尾（already_closed 是成功，不是失败）；
 *  - 清理副本**不等于**撤销：原表的撤销记录一格不动。
 */
describe("WPS 审阅副本的结束与清理", () => {
  function staged() {
    const host = fakeHost();
    const prepared = prepareWpsPlan(host.application, [1], host.plan);
    const review = stageWpsNativeComparison(host.application, prepared);
    return { host, prepared, review };
  }
  it("正常关闭 ⇒ closed，且原表值、原表焦点与撤销记录都不受影响", async () => {
    const { host, prepared, review } = staged();
    verifyWpsNativeComparison(host.application, review);
    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    const undoBefore = latestWpsUndo(host.application);
    expect(undoBefore).not.toBeNull();
    expect(discardWpsNativeComparison(review)).toBe("closed");
    expect(host.copy.Close).toHaveBeenCalledWith(false);
    expect(host.copyClosed()).toBe(true);
    // 清理副本 ≠ 撤销：恢复点仍在，而且真的能撤销。
    expect(latestWpsUndo(host.application)).toEqual(undoBefore);
    expect(host.originalCell.value).toBe("结果");
    await undoWpsExecution(host.application, record.executionId);
    expect(host.originalCell.value).toBe("原值");
  });
  it("宿主拒绝关闭 / 返回 false ⇒ 既不谎报已关闭，也不把它说成业务失败", () => {
    const rejected = staged();
    (rejected.host.copy.Close as unknown as ReturnType<typeof vi.fn>).mockImplementationOnce(() => { throw new Error("宿主拒绝关闭"); });
    expect(discardWpsNativeComparison(rejected.review)).toBe("failed");
    // 抛过异常之后副本仍在：不能报 closed。
    expect(rejected.host.copyClosed()).toBe(false);
    const refused = staged();
    (refused.host.copy.Close as unknown as ReturnType<typeof vi.fn>).mockReturnValueOnce(false);
    // Close(false) 返回 false 是**歧义**信号（被取消/被推迟/被拒），而 Close 不返回任何
    // 对话框信息 ⇒ 只能说"关闭尚未确认"，绝不推断宿主正在询问是否保存。
    expect(discardWpsNativeComparison(refused.review)).toBe("close_unconfirmed");
    expect(refused.host.copyClosed()).toBe(false);
    const unsupported = staged();
    Object.defineProperty(unsupported.host.copy, "Close", { value: undefined, configurable: true });
    expect(discardWpsNativeComparison(unsupported.review)).toBe("unsupported");
    expect(unsupported.host.copyClosed()).toBe(false);
  });
  /*
   * 第三项判断：副本身份可信、没被另存 **都推不出**"副本里没有额外人工修改"。
   * 关掉（= 不保存地丢掉）之前必须证明副本还是我们写入后的样子，否则一个 Close 都不发。
   */
  it("副本里被人工改过 ⇒ 不关副本、报 extra_edits，改动不被静默丢弃", () => {
    const { host, review } = staged();
    // 用户在并排审阅期间往副本里敲了新内容。
    (host.copy.Worksheets.Item(1).Range!("A1") as unknown as { Value2: string }).Value2 = "人工改的";
    expect(wpsCopyContentVerdict(review)).toBe("changed");
    expect(discardWpsNativeComparison(review)).toBe("extra_edits");
    // 关掉就等于把用户刚敲的内容静默丢掉 ⇒ 一个 Close 调用都不许有。
    expect(host.copy.Close).not.toHaveBeenCalled();
    expect(host.copyClosed()).toBe(false);
    // 副本里的人工修改**不导入原表**。
    expect(host.originalCell.value).toBe("原值");
  });
  it("副本读数拿不到 ⇒ edits_unverified，同样不关副本（不把'没读到'当成'没改过'）", () => {
    const { host, review } = staged();
    host.copy.Worksheets.Item = () => { throw new Error("副本不可读"); };
    expect(wpsCopyContentVerdict(review)).toBe("unknown");
    expect(discardWpsNativeComparison(review)).toBe("edits_unverified");
    expect(host.copy.Close).not.toHaveBeenCalled();
    expect(host.copyClosed()).toBe(false);
  });
  it("副本没被改过 ⇒ 正常关闭（正向锚点：上面两条不是'永远拒绝关闭'）", () => {
    const { host, review } = staged();
    expect(wpsCopyContentVerdict(review)).toBe("unchanged");
    expect(discardWpsNativeComparison(review)).toBe("closed");
    expect(host.copy.Close).toHaveBeenCalledWith(false);
    expect(host.copyClosed()).toBe(true);
  });
  it("写回侧与放弃侧对副本的第三项判断必须一致：副本被改过时写回同样阻断", () => {
    const { host, review } = staged();
    (host.copy.Worksheets.Item(1).Range!("A1") as unknown as { Value2: string }).Value2 = "人工改的";
    expect(() => verifyWpsNativeComparison(host.application, review)).toThrow(/发生变化/);
    // 写回被阻断 ⇒ 原表一格未写；副本里的人工修改绝不导入原表。
    expect(host.originalCell.value).toBe("原值");
  });
  it("副本已被用户关闭 ⇒ already_closed（幂等）；被另存/替换 ⇒ user_saved 且不动它", () => {
    const gone = staged();
    Object.defineProperty(gone.host.copy, "Name", { get() { throw new Error("副本已不在"); }, configurable: true });
    expect(discardWpsNativeComparison(gone.review)).toBe("already_closed");
    expect(gone.host.copy.Close).not.toHaveBeenCalled();
    // 另存/替换：身份变了 ⇒ 那是用户自己的文件，绝不自动丢弃。
    const savedAs = staged();
    savedAs.host.copy.Name = "另存为的名字.xlsx";
    expect(discardWpsNativeComparison(savedAs.review)).toBe("user_saved");
    expect(savedAs.host.copy.Close).not.toHaveBeenCalled();
    expect(savedAs.host.copyClosed()).toBe(false);
    // 第二次清理：已经不是我们那一份了，仍然只提示、不动手。
    expect(discardWpsNativeComparison(savedAs.review)).toBe("user_saved");
  });
  it("身份核对发生在关闭之后不成立：先复核身份，再决定动不动手", () => {
    const { host, review } = staged();
    // 名字相同但 FullName/Path 变了（被另存到磁盘）——身份不同即为 user_saved。
    host.copy.FullName = "E:/saved/Book1.xlsx";
    expect(discardWpsNativeComparison(review)).toBe("user_saved");
    expect(host.copy.Close).not.toHaveBeenCalled();
  });
});
