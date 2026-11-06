import { describe, expect, it } from "vitest";
import {
  getWpsApplication,
  maxFocusWindowSize,
  readWpsHostProbe,
  WpsHostUnavailableError,
  type WpsApplication,
  type WpsWorksheet
} from "./wpsHost";

function fakeApplication(overrides: Partial<WpsApplication> = {}): WpsApplication {
  const sheets: WpsWorksheet[] = [{ Name: "明细" }, { Name: "汇总" }];
  return {
    PluginStorage: {
      getItem: () => null,
      setItem: () => undefined
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined,
    ActiveWorkbook: {
      Name: "测试.xlsx",
      FullName: "C:\\Data\\测试.xlsx",
      Worksheets: {
        Count: sheets.length,
        Item: (index) => sheets[Number(index) - 1]
      },
      ActiveSheet: sheets[0]
    },
    ActiveSheet: sheets[0],
    Selection: {
      Address: "$A$1:$B$2",
      Text: [["影院", "影片"], ["甲", "乙"]]
    },
    ...overrides
  };
}

describe("WPS host probe", () => {
  it("the scope picker metadata probe does not touch selected cell contents", () => {
    const application = fakeApplication();
    Object.defineProperty(application, "Selection", { get() { throw new Error("no business data before consent"); } });
    const result = readWpsHostProbe(application, { includeSelectionPreview: false });
    expect(result.workbookName).toBe("测试.xlsx");
    expect(result.selectionPreview).toBeNull();
  });
  it("rejects pages that are not connected to WPS", () => {
    expect(() => getWpsApplication({ Application: undefined })).toThrow(WpsHostUnavailableError);
  });

  it("reads workbook, sheets and selection without writing", () => {
    const result = readWpsHostProbe(fakeApplication());
    expect(result).toEqual({
      host: "wps-et",
      workbookName: "测试.xlsx",
      workbookPath: "C:\\Data\\测试.xlsx",
      activeSheetName: "明细",
      sheetNames: ["明细", "汇总"],
      selectionAddress: "$A$1:$B$2",
      selectionPreview: "影院 · 影片 / 甲 · 乙"
    });
  });

  it("calls the WPS Address method instead of rendering the function", () => {
    const application = fakeApplication();
    application.Selection = {
      Address: () => "$C$7",
      Text: "测试"
    };

    expect(readWpsHostProbe(application).selectionAddress).toBe("$C$7");
  });

  it("requires an open workbook", () => {
    expect(() => readWpsHostProbe(fakeApplication({ ActiveWorkbook: undefined }))).toThrow(
      "当前没有打开的工作簿"
    );
  });
});

describe("maxFocusWindowSize", () => {
  it("uses the full available screen area when the WebView reports it", () => {
    expect(maxFocusWindowSize({ availWidth: 1920, availHeight: 1040 })).toEqual({
      width: 1920,
      height: 1040
    });
  });

  it("falls back to the conservative target when screen metrics are unavailable", () => {
    expect(maxFocusWindowSize({ availWidth: 0, availHeight: 0 })).toEqual({ width: 1280, height: 860 });
    expect(maxFocusWindowSize({ availWidth: NaN, availHeight: NaN })).toEqual({ width: 1280, height: 860 });
  });
});
