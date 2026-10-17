import { describe, expect, it } from "vitest";
import {
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
