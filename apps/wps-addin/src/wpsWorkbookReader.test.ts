import { describe, expect, it } from "vitest";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import {
  authorizeWpsSheets,
  listWpsWorksheets,
  normalizeWpsMatrix,
  readAuthorizedWpsSheetSchemas,
  readAuthorizedWpsSheets,
  WpsAuthorizationError
} from "./wpsWorkbookReader";

/** 结构扫描的列宽上限（`capabilities.snapshot.structureColumns`），用例据此算预期宽度。 */
const STRUCTURE_COLUMNS = 100;
/** 结构扫描的行上限（`capabilities.queryTable.headerScanRows`）。 */
const HEADER_SCAN_ROWS = 12;

function makeSheet(name: string, values: unknown[][]): WpsWorksheet {
  const rowCount = values.length;
  const columnCount = Math.max(0, ...values.map((row) => row.length));
  const usedRange: WpsRange = {
    Row: 1,
    Column: 1,
    Rows: { Count: rowCount },
    Columns: { Count: columnCount },
    Item: (rowIndex, columnIndex = 1) => ({ Row: rowIndex, Column: columnIndex })
  };
  return {
    Name: name,
    UsedRange: usedRange,
    /**
     * `...extra` 是**真宿主签名契约**：真宿主的 `Worksheet.Range` 只接受两个
     * 参数，多传会抛 `Too many parameters.`（2026-09-30 真机证据）。替身在这里
     * 模拟同一限制 —— 生产代码若再往宿主塞多余实参，任何经过 `readChunkRange`
     * 的测试都会当场红，而不是等真机爆炸。
     */
    Range: (first: WpsRange, last?: WpsRange, ...extra: unknown[]) => {
      if (extra.length > 0) throw new Error("Too many parameters.");
      const startRow = Number(first.Row ?? 1);
      const startColumn = Number(first.Column ?? 1);
      const endRow = Number(last?.Row ?? startRow);
      const endColumn = Number(last?.Column ?? startColumn);
      return {
        Value2: values
          .slice(startRow - 1, endRow)
          .map((row) => row.slice(startColumn - 1, endColumn))
      };
    }
  };
}

/**
 * 在**真实 Value2 访问面**上记账的替身宿主。
 *
 * 为什么不能按 `purpose` 标签记账：`purpose` 只是调用方贴上去的标记，
 * 它**不影响**替身取到哪些格。按"取数据目的才计数"来算就会得出
 * "数据读取 = 0"这种与实际访问范围不符的结论。这里改成：谁被取过值就记谁，
 * 然后由用例去判断这些格子**算不算数据区**。
 *
 * `tracker.reached(row, col)` 是判据：被取过值的格子里有没有数据行。
 */
function makeTrackingSheet(
  name: string,
  values: unknown[][],
  tracker: { valueReads: string[]; ranges: Array<{ first: string; last: string }> }
): WpsWorksheet {
  const base = makeSheet(name, values);
  return {
    ...base,
    /** 同上：多传第 3 个实参 → 抛 `Too many parameters.`（真宿主签名契约）。 */
    Range: (first: WpsRange, last?: WpsRange, ...extra: unknown[]) => {
      if (extra.length > 0) throw new Error("Too many parameters.");
      const describe = (cell: unknown): string =>
        typeof cell === "string"
          ? cell
          : `${(cell as WpsRange)?.Row ?? "?"},${(cell as WpsRange)?.Column ?? "?"}`;
      tracker.ranges.push({ first: describe(first), last: describe(last) });
      const range = base.Range!(first, last);
      const startRow = Number((first as WpsRange).Row ?? 1);
      const startColumn = Number((first as WpsRange).Column ?? 1);
      const endRow = Number((last as WpsRange).Row ?? startRow);
      const endColumn = Number((last as WpsRange).Column ?? startColumn);
      const flat = range.Value2 as unknown;
      // `Value2` 在真实宿主上是**属性读取**：取到哪块，那块的值就进了内存。
      Object.defineProperty({}, "unused", { value: undefined });
      const rows = Array.isArray(flat) && Array.isArray((flat as unknown[])[0])
        ? (flat as unknown[][])
        : [[flat]];
      rows.forEach((row, rowOffset) => {
        (Array.isArray(row) ? row : [row]).forEach((_cell, columnOffset) => {
          tracker.valueReads.push(`${startRow + rowOffset},${startColumn + columnOffset}`);
        });
      });
      void endRow;
      void endColumn;
      return range;
    }
  };
}

function makeApplication(sheets: WpsWorksheet[]): WpsApplication {
  const workbook = {
    Name: "测试.xlsx",
    FullName: "C:\\Data\\测试.xlsx",
    Path: "C:\\Data",
    Worksheets: {
      Count: sheets.length,
      Item: (index: number | string) => sheets[Number(index) - 1]
    },
    ActiveSheet: sheets[0]
  };
  return {
    PluginStorage: { getItem: () => null, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined,
    ActiveWorkbook: workbook,
    ActiveSheet: sheets[0]
  };
}

describe("WPS workbook reader", () => {
  it("normalizes scalar, row and column return shapes", () => {
    expect(normalizeWpsMatrix("A", 1, 1)).toEqual([["A"]]);
    expect(normalizeWpsMatrix(["A", "B"], 1, 2)).toEqual([["A", "B"]]);
    expect(normalizeWpsMatrix(["A", "B"], 2, 1)).toEqual([["A"], ["B"]]);
  });

  it("lists metadata without reading cell values", () => {
    const application = makeApplication([makeSheet("明细", [[1]]), makeSheet("汇总", [[2]])]);
    expect(listWpsWorksheets(application).sheets).toEqual([
      { index: 1, name: "明细", isActive: true },
      { index: 2, name: "汇总", isActive: false }
    ]);
  });

  it("hides EB system sheets while preserving native WPS indexes", () => {
    const application = makeApplication([
      makeSheet("明细", [[1]]),
      makeSheet("#EB_RULES", [["内部规则"]]),
      makeSheet("汇总", [[2]]),
      makeSheet("#eb_rules_backup", [["备份"]])
    ]);
    expect(listWpsWorksheets(application).sheets).toEqual([
      { index: 1, name: "明细", isActive: true },
      { index: 3, name: "汇总", isActive: false }
    ]);
    expect(() => authorizeWpsSheets(application, [2])).toThrow("内部系统表");
    expect(() => authorizeWpsSheets(application, [4])).toThrow("内部系统表");
  });

  it("does not expand an explicit empty authorization", () => {
    const application = makeApplication([makeSheet("明细", [[1]])]);
    expect(() => authorizeWpsSheets(application, [])).toThrow(WpsAuthorizationError);
  });

  it("reads only authorized sheets in chunks without dropping rows", async () => {
    const application = makeApplication([
      makeSheet("明细", [["标题"], [1], [2], [3], [4]]),
      makeSheet("未授权", [["秘密"]])
    ]);
    const authorization = authorizeWpsSheets(application, [1]);
    const progress: number[] = [];
    const results = await readAuthorizedWpsSheets(application, authorization, {
      chunkRows: 2,
      yieldControl: async () => undefined,
      onProgress: ({ rowsRead }) => progress.push(rowsRead)
    });

    expect(results).toHaveLength(1);
    expect(results[0].sheet.name).toBe("明细");
    expect(results[0].values).toEqual([["标题"], [1], [2], [3], [4]]);
    expect(results[0].chunksRead).toBe(3);
    expect(progress).toEqual([2, 4, 5]);
  });

  it("reads candidate headers without scanning all data rows", async () => {
    const values = [["影院", "数量"], ...Array.from({ length: 50 }, (_, index) => [`影院${index}`, index])];
    const sheet = makeSheet("明细", values);
    const range = sheet.Range!;
    const reads: number[] = [];
    sheet.Range = (first, last) => {
      reads.push(Number((last as WpsRange).Row ?? 0));
      return range(first, last);
    };
    const application = makeApplication([sheet]);
    const authorization = authorizeWpsSheets(application, [1]);
    const results = await readAuthorizedWpsSheetSchemas(application, authorization, {
      yieldControl: async () => undefined
    });

    expect(results[0].rowCount).toBe(51);
    expect(results[0].values[0]).toEqual(["影院", "数量"]);
    expect(results[0].values.length).toBeLessThan(51);
    expect(Math.max(...reads)).toBeLessThan(51);
  });

  it("stops when the active workbook changes", async () => {
    const application = makeApplication([makeSheet("明细", [[1], [2], [3]])]);
    const authorization = authorizeWpsSheets(application, [1]);
    let yielded = false;
    await expect(readAuthorizedWpsSheets(application, authorization, {
      chunkRows: 1,
      yieldControl: async () => {
        if (yielded) return;
        yielded = true;
        const other = makeApplication([makeSheet("其他", [[9]])]).ActiveWorkbook!;
        other.Name = "其他.xlsx";
        other.FullName = "C:\\Data\\其他.xlsx";
        application.ActiveWorkbook = other;
      }
    })).rejects.toThrow("工作簿身份发生变化");
  });

  it("accepts a refreshed host proxy for the same saved workbook", async () => {
    const sheets = [makeSheet("明细", [[1], [2]])];
    const application = makeApplication(sheets);
    const authorization = authorizeWpsSheets(application, [1]);
    application.ActiveWorkbook = {
      ...application.ActiveWorkbook!,
      Worksheets: {
        Count: sheets.length,
        Item: (index) => sheets[Number(index) - 1]
      }
    };

    await expect(readAuthorizedWpsSheets(application, authorization, {
      chunkRows: 1,
      yieldControl: async () => undefined
    })).resolves.toHaveLength(1);
  });

  it("stops when an authorized worksheet is renamed", async () => {
    const sheet = makeSheet("明细", [[1], [2]]);
    const application = makeApplication([sheet]);
    const authorization = authorizeWpsSheets(application, [1]);
    await expect(readAuthorizedWpsSheets(application, authorization, {
      chunkRows: 1,
      yieldControl: async () => { sheet.Name = "已改名"; }
    })).rejects.toThrow("被移动、重命名或删除");
  });

  it("结构扫描的真实取值范围只覆盖表头行，数据行的值不进内存", async () => {
    // 已用区域 A1:B3：第 1 行是表头，第 2/3 行是**数据行**。这条用例按真实
    // `Value2` 访问范围取证 —— 不认 `purpose` 标签，只看哪些格真的被取过值。
    const tracker = { valueReads: [] as string[], ranges: [] as Array<{ first: string; last: string }> };
    const sheet = makeTrackingSheet(
      "明细",
      [["类别", "归属"], ["办公用品", "A"], ["清洁用品", "B"]],
      tracker
    );
    const application = makeApplication([sheet]);
    const authorization = authorizeWpsSheets(application, [1]);

    const results = await readAuthorizedWpsSheetSchemas(application, authorization, {
      yieldControl: async () => undefined
    });

    // 正向锚点：结构扫描确实取到了**表头**的值（否则下面的负向断言是空过）。
    expect(results[0].values[0]).toEqual(["类别", "归属"]);
    // 反例判据：数据行（第 2 行及以后）的值**一次都没有**被取走。
    const dataRowReads = tracker.valueReads.filter((cell) => Number(cell.split(",")[0]) > 1);
    expect(dataRowReads).toEqual([]);
    // 表头行的每一格则确实被取过。
    expect(tracker.valueReads).toContain("1,1");
    expect(tracker.valueReads).toContain("1,2");
    // 尺寸仍然报告整块已用区域 —— 授权边界与上限判定不受读取范围影响。
    expect(results[0].rowCount).toBe(3);
    expect(results[0].columnCount).toBe(2);
  });

  it("结构扫描的 Range 参数里不出现数据行地址", async () => {
    const tracker = { valueReads: [] as string[], ranges: [] as Array<{ first: string; last: string }> };
    const sheet = makeTrackingSheet(
      "明细",
      [["类别", "归属"], ["办公用品", "A"], ["清洁用品", "B"]],
      tracker
    );
    const application = makeApplication([sheet]);
    const authorization = authorizeWpsSheets(application, [1]);

    await readAuthorizedWpsSheetSchemas(application, authorization, {
      yieldControl: async () => undefined
    });

    expect(tracker.ranges.length).toBeGreaterThan(0);
    for (const range of tracker.ranges) {
      // 终点必须落在表头扫描行内（第 1..headerScanRows 行），不能伸进数据区。
      const endRow = Number(range.last.split(",")[0]);
      expect(Number.isFinite(endRow)).toBe(true);
      expect(endRow).toBeLessThanOrEqual(HEADER_SCAN_ROWS);
      // 起点也不该是数据行。
      expect(Number(range.first.split(",")[0])).toBeLessThanOrEqual(HEADER_SCAN_ROWS);
    }
  });

  it("结构扫描的列宽不取决于已用区域的列数，不靠多读列去凑结构", async () => {
    // 已用区域是 3 行 × 300 列：结构扫描只需取表头行，且列宽收敛到
    // `structureColumns`（100），不因为表很宽就多取一倍的值。
    const header = Array.from({ length: 300 }, (_, index) => `字段${index + 1}`);
    const tracker = { valueReads: [] as string[], ranges: [] as Array<{ first: string; last: string }> };
    const sheet = makeTrackingSheet(
      "宽表",
      [header, [...header], [...header]],
      tracker
    );
    const application = makeApplication([sheet]);
    const authorization = authorizeWpsSheets(application, [1]);

    const results = await readAuthorizedWpsSheetSchemas(application, authorization, {
      yieldControl: async () => undefined
    });

    expect(results[0].columnCount).toBe(300);
    expect(tracker.valueReads.length).toBe(STRUCTURE_COLUMNS);
    expect(tracker.valueReads.every((cell) => cell.startsWith("1,"))).toBe(true);
  });
});
