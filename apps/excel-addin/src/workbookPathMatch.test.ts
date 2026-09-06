import { describe, expect, it } from "vitest";
import {
  folderFileFullPath,
  normalizeWorkbookPath
} from "./workbookPathMatch";

describe("workbook path matching", () => {
  it("strips the file URL prefix and decodes the real workbook path", () => {
    expect(
      normalizeWorkbookPath(
        "file:///C:/Users/jia/Desktop/%E5%BD%B1%E9%99%A2%E6%98%8E%E7%BB%86.xlsx"
      )
    ).toBe("c:/users/jia/desktop/影院明细.xlsx");
  });

  it("normalizes separators, suffixes, casing, and empty values", () => {
    expect(normalizeWorkbookPath("C:\\Data\\Book.xlsx?version=1#sheet"))
      .toBe("c:/data/book.xlsx");
    expect(normalizeWorkbookPath("c:/DATA/book.XLSX"))
      .toBe("c:/data/book.xlsx");
    expect(normalizeWorkbookPath("")).toBeNull();
    expect(normalizeWorkbookPath(undefined)).toBeNull();
  });

  it("preserves the UNC host and share", () => {
    expect(normalizeWorkbookPath("file://server/share/a.xlsx"))
      .toBe("//server/share/a.xlsx");
  });

  it("does not confuse workbooks with the same name in different folders", () => {
    expect(normalizeWorkbookPath("file:///C:/A/ss.xlsx"))
      .not.toBe(normalizeWorkbookPath("C:/B/ss.xlsx"));
  });

  it("joins a folder path with a nested relative workbook path", () => {
    expect(folderFileFullPath("C:\\Data", "sub\\ss.xlsx"))
      .toBe("c:/data/sub/ss.xlsx");
  });
});