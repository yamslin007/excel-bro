import { describe, expect, it } from "vitest";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import {
  authorizeWpsSheets,
  listWpsWorksheets,
  normalizeWpsMatrix,
  readAuthorizedWpsSheets,
  WpsAuthorizationError
} from "./wpsWorkbookReader";

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
    Range: (first, last) => {
      const startRow = Number((first as WpsRange).Row ?? 1);
      const startColumn = Number((first as WpsRange).Column ?? 1);
      const endRow = Number((last as WpsRange).Row ?? startRow);
      const endColumn = Number((last as WpsRange).Column ?? startColumn);
      return {
        Value2: values
          .slice(startRow - 1, endRow)
          .map((row) => row.slice(startColumn - 1, endColumn))
      };
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
});
