/**
 * WPS sortRange：真实状态变化（不是"配置里有这一项"）。
 *
 * 这一组用例全部跑在**带真实网格状态**的替身宿主上：`Range(address)` 返回真实的
 * 子区域视图，`Item(r,c)` 定位到真实单元格，写入会改变退回去再读到的值。
 * 所以"排序完成"必须由**回读到的行顺序**证明，而不是"写方法被调用过"。
 *
 * 覆盖：正常执行 → 回读 → 撤销；空白恒末位；表头不动；含公式 / 格式不一致 /
 * 表格对象重叠 / 宿主缺 ListObjects 四条写入前拒绝；预览后目标变化零写入；
 * 执行后人工修改使撤销冲突；未授权工作表拒绝。
 */
import { describe, expect, it } from "vitest";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import { executeWpsPlan, prepareWpsPlan, undoWpsExecution } from "./wpsPlanExecutor";

type Style = {
  bold: boolean;
  italic: boolean;
  name: string;
  size: number;
  color: number;
  horizontal: number;
  vertical: number;
  wrap: boolean;
  fillColor: number;
  fillPattern: number;
  fillPatternColor: number;
  numberFormat: string;
};

const BASE_STYLE: Style = {
  bold: false, italic: false, name: "Calibri", size: 11, color: 0,
  horizontal: -4131, vertical: -4108, wrap: false,
  fillColor: 16777215, fillPattern: -4142, fillPatternColor: 0, numberFormat: "General"
};

type Cell = { value: CellValue; formula: CellValue; style: Style };

function columnName(index: number): string {
  let value = index;
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function columnIndex(name: string): number {
  return [...name.toUpperCase()].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}

function parseAddress(address: string) {
  const match = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(address.trim());
  if (!match) throw new Error(`测试宿主持有不支持的地址：${address}`);
  const startColumn = columnIndex(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnIndex(match[3] ?? match[1]);
  const endRow = Number(match[4] ?? match[2]);
  return { startRow, startColumn, endRow, endColumn,
    rows: endRow - startRow + 1, columns: endColumn - startColumn + 1 };
}

type GridHostOptions = {
  /** 每行的格式覆盖（下标 0 是表头行）。 */
  styleByRow?: (rowIndex: number) => Partial<Style>;
  /** 表格对象（ListObject）声明。 */
  listObjects?: Array<{ name: string; address: string }>;
  /** false = 宿主不提供 `ListObjects`（模拟旧版本）。 */
  provideListObjects?: boolean;
};

function gridHost(initial: CellValue[][], options: GridHostOptions = {}) {
  const rows = initial.length;
  const columns = initial[0].length;
  const cells = new Map<string, Cell>();
  for (let r = 1; r <= rows; r += 1) {
    for (let c = 1; c <= columns; c += 1) {
      const value = initial[r - 1]?.[c - 1] ?? null;
      cells.set(`${r},${c}`, {
        value,
        formula: typeof value === "string" && value.startsWith("=")
          ? value
          : (value === null ? "" : String(value)),
        style: { ...BASE_STYLE, ...(options.styleByRow?.(r - 1) ?? {}) }
      });
    }
  }
  const storage = new Map<string, unknown>();
  let writeCount = 0;

  const cellAt = (row: number, column: number): Cell => {
    const key = `${row},${column}`;
    let cell = cells.get(key);
    // 真宿主对"区域里还没被写过的格子"返回空单元格；替身同样按需创建，
    // 否则 A1:B4 指向一行的表时会在读取阶段抛错，掩盖真正的业务守卫。
    if (!cell) {
      cell = { value: null, formula: "", style: { ...BASE_STYLE } };
      cells.set(key, cell);
    }
    return cell;
  };

  const makeRange = (rect: ReturnType<typeof parseAddress>): WpsRange => {
    const single = rect.rows === 1 && rect.columns === 1;
    const view: Record<string, unknown> = {
      MergeCells: false,
      Rows: { Count: rect.rows },
      Columns: { Count: rect.columns },
      Address: `${columnName(rect.startColumn)}${rect.startRow}:${columnName(rect.endColumn)}${rect.endRow}`
    };
    Object.defineProperties(view, {
      Value2: {
        get: () => {
          const matrix: CellValue[][] = [];
          for (let r = 0; r < rect.rows; r += 1) {
            const row: CellValue[] = [];
            for (let c = 0; c < rect.columns; c += 1) row.push(cellAt(rect.startRow + r, rect.startColumn + c).value);
            matrix.push(row);
          }
          return single ? matrix[0][0] : matrix;
        },
        set: (next: CellValue[][] | CellValue) => {
          writeCount += 1;
          const matrix = Array.isArray(next)
            ? (single ? [[next]] as never : next as CellValue[][])
            : [[next]];
          for (let r = 0; r < rect.rows; r += 1) {
            for (let c = 0; c < rect.columns; c += 1) {
              const cell = cellAt(rect.startRow + r, rect.startColumn + c);
              const value = (matrix as CellValue[][])[r]?.[c] ?? null;
              cell.value = value;
              cell.formula = value === null ? "" : value;
            }
          }
        },
        enumerable: true
      },
      Formula: {
        get: () => {
          const matrix: CellValue[][] = [];
          for (let r = 0; r < rect.rows; r += 1) {
            const row: CellValue[] = [];
            for (let c = 0; c < rect.columns; c += 1) row.push(cellAt(rect.startRow + r, rect.startColumn + c).formula);
            matrix.push(row);
          }
          return single ? matrix[0][0] : matrix;
        },
        set: (next: CellValue[][] | CellValue) => {
          const matrix = Array.isArray(next) ? next as CellValue[][] : [[next]];
          for (let r = 0; r < rect.rows; r += 1) {
            for (let c = 0; c < rect.columns; c += 1) {
              const cell = cellAt(rect.startRow + r, rect.startColumn + c);
              const formula = matrix[r]?.[c] ?? "";
              cell.formula = formula;
              cell.value = formula;
            }
          }
        },
        enumerable: true
      }
    });
    if (!single) {
      view.Item = (rowOffset: number, columnOffset = 1) =>
        makeRange(parseAddress(`${columnName(rect.startColumn + columnOffset - 1)}${rect.startRow + rowOffset - 1}`));
    } else {
      const cell = cellAt(rect.startRow, rect.startColumn);
      const style = cell.style;
      view.Font = {
        get Bold() { return style.bold; }, set Bold(next: boolean) { style.bold = next; },
        get Italic() { return style.italic; }, set Italic(next: boolean) { style.italic = next; },
        get Name() { return style.name; }, set Name(next: string) { style.name = next; },
        get Size() { return style.size; }, set Size(next: number) { style.size = next; },
        get Color() { return style.color; }, set Color(next: number) { style.color = next; }
      };
      view.Interior = {
        get Color() { return style.fillColor; }, set Color(next: number) { style.fillColor = next; },
        get Pattern() { return style.fillPattern; }, set Pattern(next: number) { style.fillPattern = next; },
        get PatternColor() { return style.fillPatternColor; }, set PatternColor(next: number) { style.fillPatternColor = next; }
      };
      Object.defineProperties(view, {
        HorizontalAlignment: {
          get: () => style.horizontal, set: (next: number) => { style.horizontal = next; }, enumerable: true
        },
        VerticalAlignment: {
          get: () => style.vertical, set: (next: number) => { style.vertical = next; }, enumerable: true
        },
        WrapText: {
          get: () => style.wrap, set: (next: boolean) => { style.wrap = next; }, enumerable: true
        },
        NumberFormat: {
          get: () => style.numberFormat, set: (next: string) => { style.numberFormat = next; }, enumerable: true
        },
        RowHeight: { get: () => 15, enumerable: true }
      });
    }
    return view as unknown as WpsRange;
  };

  const sheet: WpsWorksheet = {
    Name: "明细",
    Index: 1,
    ProtectContents: false,
    Range: (address: string) => makeRange(parseAddress(address))
  };
  if (options.provideListObjects !== false) {
    const objects = options.listObjects ?? [];
    sheet.ListObjects = {
      Count: objects.length,
      Item: (index: number) => ({
        Name: objects[index - 1].name,
        Range: makeRange(parseAddress(objects[index - 1].address))
      })
    };
  }
  const otherSheet: WpsWorksheet = { Name: "未授权", Index: 2, ProtectContents: false, Range: sheet.Range };
  const sheets = [sheet, otherSheet];
  const workbook = {
    Name: "测试.xlsx",
    FullName: "C:\\Data\\测试.xlsx",
    Path: "C:\\Data",
    ReadOnly: false,
    Worksheets: { Count: sheets.length, Item: (index: number) => sheets[index - 1] },
    ActiveSheet: sheet
  };
  const application = {
    Version: "12.0",
    ActiveWorkbook: workbook,
    ActiveSheet: sheet,
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, value: unknown) => { storage.set(key, value); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as unknown as WpsApplication;

  return {
    application,
    /** 逐行的第一列值 —— 用它断言"顺序真的变了"。 */
    column: (index: number) => {
      const result: CellValue[] = [];
      for (let r = 1; r <= rows; r += 1) result.push(cellAt(r, index + 1).value);
      return result;
    },
    setCell: (row: number, column: number, value: CellValue) => {
      const cell = cellAt(row, column);
      cell.value = value;
      cell.formula = value;
    },
    setRowStyle: (row: number, patch: Partial<Style>) => {
      for (let c = 1; c <= columns; c += 1) Object.assign(cellAt(row, c).style, patch);
    },
    writes: () => writeCount,
    records: () => JSON.parse(String(storage.get("excel-bro.wps.undo.v1") ?? "[]"))
  };
}

const DATA: CellValue[][] = [
  ["名称", "数量"],
  ["甲", 30],
  ["乙", 10],
  ["丙", 20]
];

function sortPlan(overrides: Partial<Extract<AnalysisPlan["actions"][number], { type: "sortRange" }>> = {}): AnalysisPlan {
  return {
    id: "plan-sort",
    title: "按数量升序",
    summary: "把明细按数量从小到大排序",
    assumptions: [],
    warnings: [],
    actions: [{
      type: "sortRange",
      sheet: "明细",
      range: "A1:B4",
      keys: [{ column: 1, ascending: true }],
      hasHeaders: true,
      ...overrides
    }],
    acceptanceCriteria: []
  };
}

describe("WPS sortRange：真实网格上的执行、回读与撤销", () => {
  it("升序：执行后行顺序真的变了，回读确认，撤销后逐格回到执行前", async () => {
    const host = gridHost(DATA);
    const prepared = prepareWpsPlan(host.application, [1], sortPlan());
    expect(host.column(0)).toEqual(["名称", "甲", "乙", "丙"]); // 预览阶段零写入

    const record = await executeWpsPlan(host.application, prepared);
    expect(record.state).toBe("verified");
    expect(host.column(0)).toEqual(["名称", "乙", "丙", "甲"]);
    expect(host.column(1)).toEqual(["数量", 10, 20, 30]);

    await undoWpsExecution(host.application, record.executionId);
    expect(host.column(0)).toEqual(["名称", "甲", "乙", "丙"]);
    expect(host.column(1)).toEqual(["数量", 30, 10, 20]);
  });

  it("表头不动：只排数据行，表头留在第一行", async () => {
    // 用 ASCII 名字断言，避免文本排序依赖运行环境的本地化规则；
    // 中文文本的顺序由共享层 `localeCompare` 决定，不在这一条里钉死。
    const host = gridHost([
      ["名称", "数量"],
      ["b", 30],
      ["a", 10],
      ["c", 20]
    ]);
    const prepared = prepareWpsPlan(host.application, [1], sortPlan());
    await executeWpsPlan(host.application, prepared);
    expect(host.column(0)[0]).toBe("名称");
    expect(host.column(0).slice(1)).toEqual(["a", "c", "b"]);
  });

  it("空白恒排末位：降序也不把空行顶到第一屏", async () => {
    const host = gridHost([
      ["名称", "数量"],
      ["甲", 10],
      ["乙", null],
      ["丙", 30]
    ]);
    const prepared = prepareWpsPlan(host.application, [1], sortPlan({ keys: [{ column: 1, ascending: false }] }));
    await executeWpsPlan(host.application, prepared);
    expect(host.column(0)).toEqual(["名称", "丙", "甲", "乙"]);
  });

  it("执行后把真实的逐格前后对照交给审阅流程（不是伪造摘要）", async () => {
    const host = gridHost(DATA);
    const prepared = prepareWpsPlan(host.application, [1], sortPlan());
    const captured: Array<{ sheet: string; status: string; rows: unknown[]; changed: number }> = [];
    await executeWpsPlan(host.application, prepared, undefined, (reviews) => {
      for (const review of reviews) {
        captured.push({ sheet: review.sheet, status: review.status, rows: review.rows, changed: review.changedCount });
      }
    });
    expect(captured).toHaveLength(1);
    expect(captured[0].sheet).toBe("明细");
    // 排序必然改变多行内容：对照里必须有真实行，且状态不是"未变化"。
    expect(captured[0].status).not.toBe("unchanged");
    expect(captured[0].changed).toBeGreaterThan(0);
    expect(captured[0].rows.length).toBeGreaterThan(0);
  });

  it("预览后目标内容被改：执行零写入", async () => {
    const host = gridHost(DATA);
    const prepared = prepareWpsPlan(host.application, [1], sortPlan());
    host.setCell(2, 2, 999);
    const before = host.column(0);
    await expect(executeWpsPlan(host.application, prepared)).rejects.toThrow(/在预览后发生变化/);
    expect(host.column(0)).toEqual(before);
    expect(host.column(1)).toEqual(["数量", 999, 10, 20]);
  });

  it("执行后人工改过目标：撤销报冲突，不覆盖用户改动", async () => {
    const host = gridHost(DATA);
    const prepared = prepareWpsPlan(host.application, [1], sortPlan());
    const record = await executeWpsPlan(host.application, prepared);
    host.setCell(2, 1, "用户手改");
    await expect(undoWpsExecution(host.application, record.executionId)).rejects.toThrow(/执行后又被修改/);
    expect(host.column(0)).toEqual(["名称", "用户手改", "丙", "甲"]);
  });

  it("未授权工作表：prepare 阶段就拒绝，零写入", async () => {
    const host = gridHost(DATA);
    expect(() => prepareWpsPlan(host.application, [1], sortPlan({
      sheet: "未授权"
    }))).toThrow(/未授权|授权/);
    expect(host.writes()).toBe(0);
  });
});

describe("WPS sortRange：不能可靠处理的场景必须在写入前拒绝", () => {
  it("区域含公式：拒绝，并说明相对引用会失效", () => {
    const host = gridHost([
      ["名称", "数量", "小计"],
      ["甲", 10, "=B2*2"],
      ["乙", 20, "=B3*2"]
    ]);
    expect(() => prepareWpsPlan(host.application, [1], sortPlan({ range: "A1:C3" })))
      .toThrow(/包含 \d+ 个公式/);
    expect(host.writes()).toBe(0);
  });

  it("数据行格式不一致：拒绝，因为只重排值会让格式与数据错位", () => {
    const host = gridHost(DATA, {
      styleByRow: (row) => (row === 1 ? { bold: true, fillColor: 0xff0000 } : {})
    });
    expect(() => prepareWpsPlan(host.application, [1], sortPlan()))
      .toThrow(/数据行存在不同格式/);
    expect(host.writes()).toBe(0);
  });

  it("与表格对象重叠：拒绝", () => {
    const host = gridHost(DATA, { listObjects: [{ name: "明细表", address: "A1:B4" }] });
    expect(() => prepareWpsPlan(host.application, [1], sortPlan()))
      .toThrow(/表格对象「明细表」重叠/);
    expect(host.writes()).toBe(0);
  });

  it("宿主没有 ListObjects：按「无法确认」拒绝，而不是当成「没有表格对象」", () => {
    const host = gridHost(DATA, { provideListObjects: false });
    expect(() => prepareWpsPlan(host.application, [1], sortPlan()))
      .toThrow(/未提供 ListObjects/);
    expect(host.writes()).toBe(0);
  });

  it("只有表头没有数据行：拒绝", () => {
    const host = gridHost([["名称", "数量"]]);
    expect(() => prepareWpsPlan(host.application, [1], sortPlan({ range: "A1:B1" })))
      .toThrow(/没有可排序的数据行/);
  });

  it("removeDuplicates：明确拒绝并说明原因，不静默降级成清空尾行", () => {
    const host = gridHost(DATA);
    const plan: AnalysisPlan = {
      id: "plan-dedupe",
      title: "去重",
      summary: "去掉重复行",
      assumptions: [],
      warnings: [],
      actions: [{
        type: "removeDuplicates",
        sheet: "明细",
        range: "A1:B4",
        columns: [0],
        hasHeaders: true
      }],
      acceptanceCriteria: []
    };
    expect(() => prepareWpsPlan(host.application, [1], plan))
      .toThrow(/暂不支持 removeDuplicates：它需要删除整行并把下方数据上移/);
    expect(host.writes()).toBe(0);
    expect(host.column(0)).toEqual(["名称", "甲", "乙", "丙"]);
  });

  it("排序列越界：拒绝", () => {
    const host = gridHost(DATA);
    expect(() => prepareWpsPlan(host.application, [1], sortPlan({ keys: [{ column: 5, ascending: true }] })))
      .toThrow(/超出「明细!A1:B4」的 2 列范围/);
  });
});
