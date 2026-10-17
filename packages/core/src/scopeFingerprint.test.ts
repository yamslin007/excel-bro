import { describe, expect, it } from "vitest";
import { folderRootIdentity, intentScopeFingerprint, type ConversationScope } from "./scopeFingerprint";

const scope: ConversationScope = { workbookName: "文件夹：Data", sourceMode: "folder",
  selectionMode: "folder", activeWorksheet: "book.xlsx › Selected", sheets: [{ name: "book.xlsx › Selected" }] };

describe("local folder conversation binding", () => {
  it("separates identical display names and relative IDs under different roots", () => {
    expect(intentScopeFingerprint(scope, "C:/One/Data")).not.toBe(intentScopeFingerprint(scope, "D:/Two/Data"));
    expect(intentScopeFingerprint(scope, "C:/")).not.toBe(intentScopeFingerprint(scope, "D:/"));
  });
  it("accepts Windows path spelling changes, without folding POSIX case", () => {
    expect(intentScopeFingerprint(scope, "C:\\One\\Data\\")).toBe(intentScopeFingerprint(scope, "c:/one/data"));
    expect(folderRootIdentity("/home/Data")).not.toBe(folderRootIdentity("/home/data"));
    expect(folderRootIdentity("/")).toBe("/");
    expect(folderRootIdentity("\\\\one\\share\\Data")).not.toBe(folderRootIdentity("\\\\two\\share\\Data"));
  });
  it("fails closed when no absolute folder identity is available", () => {
    for (const root of [undefined, "", "Data", "https://external.example/Data"]) {
      expect(() => intentScopeFingerprint(scope, root)).toThrow("无法核实工作目录身份");
    }
  });
  it("does not accept an old folder binding based only on display names", () => {
    expect(intentScopeFingerprint(scope, "C:/One/Data")).not.toBe(JSON.stringify({
      workbookName: scope.workbookName, sourceMode: scope.sourceMode, selectionMode: scope.selectionMode, sheets: scope.sheets.map((sheet) => sheet.name)
    }));
  });
  it("does not bind cursor changes in manual workbook mode", () => {
    const workbookScope = { ...scope, sourceMode: "workbook" as const, selectionMode: "manual" as const };
    expect(intentScopeFingerprint(workbookScope)).toBe(intentScopeFingerprint({ ...workbookScope, activeWorksheet: "Other" }));
    expect(intentScopeFingerprint({ ...workbookScope, selectionMode: "auto" })).not.toBe(
      intentScopeFingerprint({ ...workbookScope, selectionMode: "auto", activeWorksheet: "Other" }));
  });
});
