import { describe, expect, it } from "vitest";
import {
  extractWorkbookDataPeriod,
  workbookNameFromDocumentUrl,
  workbookFolderFromDocumentUrl
} from "./workbookIdentity";

describe("workbook identity", () => {
  it("extracts and decodes the workbook name from an Office document URL", () => {
    expect(
      workbookNameFromDocumentUrl(
        "file:///C:/Users/jia/Desktop/%E5%BD%B1%E9%99%A2%E6%98%8E%E7%BB%86.xlsx"
      )
    ).toBe("影院明细.xlsx");
    expect(workbookNameFromDocumentUrl("")).toBe("未保存的工作簿");
  });

  it("extracts a single report date when both ends are the same", () => {
    expect(
      extractWorkbookDataPeriod(
        "影院放映成绩实时报表(2026-07-26到2026-07-26明细).xlsx"
      )
    ).toBe("2026-07-26");
  });

  it("extracts a report date range", () => {
    expect(
      extractWorkbookDataPeriod("影院报表_2026年7月17日至2026年7月26日.xlsx")
    ).toBe("2026-07-17 至 2026-07-26");
    expect(extractWorkbookDataPeriod("普通工作簿.xlsx")).toBeNull();
  });

  it("derives local folder path from drive-letter file URLs", () => {
    expect(
      workbookFolderFromDocumentUrl("C:\\Users\\jia\\Desktop\\报表.xlsx")
    ).toEqual({ kind: "local", folderPath: "C:\\Users\\jia\\Desktop" });
    expect(
      workbookFolderFromDocumentUrl("file:///C:/work/data/影院.xlsx")
    ).toEqual({ kind: "local", folderPath: "C:/work/data" });
  });

  it("handles root directory paths correctly", () => {
    expect(workbookFolderFromDocumentUrl("D:\\报表.xlsx")).toEqual({
      kind: "local",
      folderPath: "D:\\"
    });
  });

  it("recognizes UNC paths as local", () => {
    expect(
      workbookFolderFromDocumentUrl("\\\\server\\share\\folder\\file.xlsx")
    ).toEqual({ kind: "local", folderPath: "\\\\server\\share\\folder" });
  });

  it("classifies cloud URLs as remote", () => {
    expect(
      workbookFolderFromDocumentUrl(
        "https://example.sharepoint.com/sites/team/documents/报表.xlsx"
      )
    ).toEqual({ kind: "remote" });
  });

  it("classifies empty or unsaved workbooks", () => {
    expect(workbookFolderFromDocumentUrl("")).toEqual({ kind: "unsaved" });
    expect(workbookFolderFromDocumentUrl(null)).toEqual({ kind: "unsaved" });
  });
});
