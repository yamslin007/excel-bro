// @vitest-environment jsdom
/**
 * `prepare_python_data` 的**双端共同契约**（WPS 侧）。
 *
 * 夹具与 Excel 侧 `apps/excel-addin/src/pythonPrepareShared.test.ts` **同一份**
 * `config/python_prepare_cases.json`。两端算法不同（WPS 从已用区域的**相对**
 * 矩阵切片，Excel 用 `getRange` 读**绝对**区域），答案必须一致。
 */

import { describe, expect, it } from "vitest";

import shared from "../../../config/python_prepare_cases.json";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import { sliceWpsPythonSources } from "./wpsSdkBridgeHost";
import {
  authorizeWpsSheets,
  readAuthorizedWpsSheets,
  type WpsCellValue,
  type WpsSheetReadResult
} from "./wpsWorkbookReader";

const fixture = shared as {
  sheet: {
    name: string;
    startRow: number;
    startColumn: number;
    matrix: (string | number)[][];
    texts: string[][];
    numberFormats: string[][];
  };
  slice: {
    name: string;
    range: string;
    expected?: (string | number)[][];
    refuse?: boolean;
  }[];
  values: {
    name: string;
    value: string | number | boolean;
    text: string;
    numberFormat: string;
    expected: string | number | boolean;
  }[];
  ragged: {
    name: string;
    matrix: (string | number)[][];
    rowCount: number;
    columnCount: number;
    range: string;
    expected?: (string | number)[][];
    refuse?: boolean;
  }[];
};

const SHEET = fixture.sheet.name;

function makeRead(matrix: (string | number)[][], texts?: string[][], numberFormats?: string[][]): WpsSheetReadResult {
  return {
    sheet: { index: 1, name: SHEET },
    startRow: fixture.sheet.startRow,
    startColumn: fixture.sheet.startColumn,
    rowCount: matrix.length,
    columnCount: Math.max(0, ...matrix.map((row) => row.length)),
    chunksRead: 1,
    values: matrix as WpsCellValue[][],
    ...(texts && numberFormats ? { texts, numberFormats } : {})
  };
}

function fullRead(): WpsSheetReadResult {
  return makeRead(fixture.sheet.matrix, fixture.sheet.texts, fixture.sheet.numberFormats);
}

describe("prepare_python_data 双端共同契约（WPS 侧）", () => {
  for (const item of fixture.slice) {
    it(item.name, () => {
      const run = () =>
        sliceWpsPythonSources([{ sheet: SHEET, range: item.range }], [fullRead()], [SHEET]);
      if (item.refuse === true) {
        expect(run).toThrow("数据登记区域超出已读取范围");
        return;
      }
      const blocks = run();
      expect(blocks).toHaveLength(1);
      expect(blocks[0].sourceIndex).toBe(0);
      expect(blocks[0].values).toEqual(item.expected);
    });
  }

  /**
   * WPS 独有的失败面：已用区域声明了 N 行，实际读回来的矩阵却缺行/短行。
   * 这里**不补空白** —— 补了就把"读了但没读到"伪造成"读到了空值"。
   */
  for (const item of fixture.ragged) {
    it(`不完整矩阵：${item.name}`, () => {
      const run = () =>
        sliceWpsPythonSources(
          [{ sheet: SHEET, range: item.range }],
          [makeRead(item.matrix)],
          [SHEET]
        );
      if (item.refuse === true) {
        expect(run).toThrow("数据登记区域超出已读取范围");
        return;
      }
      expect(run()[0].values).toEqual(item.expected);
    });
  }

  it("真实空白与人为补齐必须可区分：区域内空格是空串，区域外直接拒绝", () => {
    expect(
      sliceWpsPythonSources([{ sheet: SHEET, range: "E6:E6" }], [fullRead()], [SHEET])[0]
        .values
    ).toEqual([[""]]);
    expect(() =>
      sliceWpsPythonSources([{ sheet: SHEET, range: "E7:E7" }], [fullRead()], [SHEET])
    ).toThrow("数据登记区域超出已读取范围");
  });

  it("授权外的表一律拒绝", () => {
    expect(() =>
      sliceWpsPythonSources(
        [{ sheet: "没授权的表", range: "A1:A1" }],
        [fullRead()],
        [SHEET]
      )
    ).toThrow("数据登记请求超出了用户选择的工作表范围");
  });

  it("多个来源的 sourceIndex 按声明顺序给，不重排", () => {
    const blocks = sliceWpsPythonSources(
      [
        { sheet: SHEET, range: "C4:D5" },
        { sheet: SHEET, range: "B3:B3" }
      ],
      [fullRead()],
      [SHEET]
    );
    expect(blocks.map((block) => block.sourceIndex)).toEqual([0, 1]);
    expect(blocks[0].values).toEqual([
      ["苹果", 10],
      ["梨", 20]
    ]);
  });
});

describe("值归一共同契约（WPS 侧）", () => {
  for (const item of fixture.values) {
    it(item.name, () => {
      const read: WpsSheetReadResult = {
        sheet: { index: 1, name: "值契表" },
        startRow: 1,
        startColumn: 1,
        rowCount: 1,
        columnCount: 1,
        chunksRead: 1,
        values: [[item.value as WpsCellValue]],
        texts: [[item.text]],
        numberFormats: [[item.numberFormat]]
      };
      const blocks = sliceWpsPythonSources(
        [{ sheet: "值契表", range: "A1:A1" }],
        [read],
        ["值契表"]
      );
      expect(blocks[0].values).toEqual([[item.expected]]);
    });
  }
});

/** 取消：真实读取路径在信号已 abort 时不得返回任何矩阵。 */
describe("取消（WPS 侧）", () => {
  function makeApplication(values: unknown[][]): WpsApplication {
    const usedRange: WpsRange = {
      Row: 1,
      Column: 1,
      Rows: { Count: values.length },
      Columns: { Count: Math.max(0, ...values.map((row) => (row as unknown[]).length)) },
      Item: (rowIndex: number, columnIndex = 1) => ({ Row: rowIndex, Column: columnIndex })
    };
    const sheet: WpsWorksheet = {
      Name: "取消表",
      UsedRange: usedRange,
      Range: (first: WpsRange, last?: WpsRange) => ({
        Value2: values
          .slice(Number(first.Row ?? 1) - 1, Number(last?.Row ?? first.Row ?? 1))
          .map((row) =>
            (row as unknown[]).slice(
              Number(first.Column ?? 1) - 1,
              Number(last?.Column ?? first.Column ?? 1)
            )
          ),
        Text: values.map((row) => (row as unknown[]).map((value) => String(value ?? ""))),
        NumberFormat: values.map((row) => (row as unknown[]).map(() => "General"))
      })
    };
    return {
      PluginStorage: { getItem: () => null, setItem: () => undefined },
      CreateTaskPane: () => ({ ID: 1, Visible: true }),
      GetTaskPane: () => undefined,
      ActiveWorkbook: {
        Name: "取消.xlsx",
        FullName: "C:\\Data\\取消.xlsx",
        Path: "C:\\Data",
        Worksheets: { Count: 1, Item: () => sheet },
        ActiveSheet: sheet
      },
      ActiveSheet: sheet
    } as unknown as WpsApplication;
  }

  it("信号已 abort 时读取直接失败，不产出任何数据块", async () => {
    const application = makeApplication([["编号"], ["001"]]);
    const authorization = authorizeWpsSheets(application, [1]);
    const controller = new AbortController();
    controller.abort();
    await expect(
      readAuthorizedWpsSheets(application, authorization, {
        signal: controller.signal,
        includeDisplay: true
      })
    ).rejects.toThrow("读取已取消");
  });

  it("includeDisplay 会同时取回文本与数字格式，且形状与值矩阵一致", async () => {
    const application = makeApplication([
      ["编号", "名称"],
      ["001", "苹果"]
    ]);
    const authorization = authorizeWpsSheets(application, [1]);
    const reads = await readAuthorizedWpsSheets(application, authorization, {
      includeDisplay: true
    });
    expect(reads).toHaveLength(1);
    expect(reads[0].texts).toEqual([
      ["编号", "名称"],
      ["001", "苹果"]
    ]);
    expect(reads[0].numberFormats).toEqual([
      ["General", "General"],
      ["General", "General"]
    ]);
    expect(reads[0].texts).toHaveLength(reads[0].values.length);
  });
});
