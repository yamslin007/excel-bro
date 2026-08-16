// @vitest-environment jsdom
import { describe, it, expect } from "vitest";
import { renderHook, act } from "@testing-library/react";
import type {
  FolderCatalog,
  WorkbookSnapshot
} from "../contracts";
import { useScopeSelection } from "./useScopeSelection";

const workbook: WorkbookSnapshot = {
  name: "book.xlsx",
  capturedAt: "2026-08-16T00:00:00.000Z",
  activeWorksheet: "Sheet1",
  worksheets: [
    {
      name: "Sheet1",
      usedRange: "A1:B2",
      rowCount: 2,
      columnCount: 2,
      headers: ["Name", "Amount"],
      dataRows: [],
      truncated: false
    },
    {
      name: "Sheet2",
      usedRange: "A1:C3",
      rowCount: 3,
      columnCount: 3,
      headers: ["Name", "Amount", "Qty"],
      dataRows: [],
      truncated: false
    }
  ]
};

const folderCatalog: FolderCatalog = {
  sessionId: "session-1",
  folderName: "Data",
  folderPath: "C:/Data",
  files: [
    {
      id: "file-1",
      name: "a.xlsx",
      relativePath: "a.xlsx",
      worksheets: [
        { name: "Data", rowCount: 10, columnCount: 3 }
      ]
    },
    {
      id: "file-2",
      name: "b.xlsx",
      relativePath: "b.xlsx",
      worksheets: [
        { name: "Summary", rowCount: 5, columnCount: 2 }
      ]
    }
  ],
  totalFiles: 2,
  truncated: false,
  expiresAt: "2026-08-16T01:00:00.000Z"
};

describe("useScopeSelection", () => {
  it("should initialize scope and tool fields with defaults", () => {
    const { result } = renderHook(() => useScopeSelection());

    expect(result.current.toolName).toBe("");
    expect(result.current.toolDescription).toBe("");
    expect(result.current.selectedSheetNames).toEqual([]);
    expect(result.current.selectionConfirmed).toBe(false);
    expect(result.current.sourceMode).toBe("workbook");
    expect(result.current.workbookScopeMode).toBe("auto");
    expect(result.current.folderCatalog).toBeNull();
    expect(result.current.folderSheetKeys).toEqual([]);
  });

  it("should update tool name and description", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => {
      result.current.setToolName("My Tool");
      result.current.setToolDescription("Description");
    });

    expect(result.current.toolName).toBe("My Tool");
    expect(result.current.toolDescription).toBe("Description");
  });

  it("should apply workbook snapshot selection in auto mode", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.applyWorkbookSnapshotSelection(workbook));

    expect(result.current.selectedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.selectionConfirmed).toBe(true);
  });

  it("should toggle workbook sheets and unconfirm selection", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.chooseManualScope(workbook));
    act(() => result.current.toggleSheet("Sheet2"));

    expect(result.current.selectedSheetNames).toContain("Sheet2");
    expect(result.current.selectionConfirmed).toBe(false);
  });

  it("should choose automatic scope and use active worksheet", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.chooseAutomaticScope(workbook));

    expect(result.current.sourceMode).toBe("workbook");
    expect(result.current.workbookScopeMode).toBe("auto");
    expect(result.current.selectedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.selectionConfirmed).toBe(true);
  });

  it("should choose manual scope and preserve selected sheets", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.setSelectedSheetNames(["Sheet2"]));
    act(() => result.current.chooseManualScope(workbook));

    expect(result.current.sourceMode).toBe("workbook");
    expect(result.current.workbookScopeMode).toBe("manual");
    expect(result.current.selectedSheetNames).toEqual(["Sheet2"]);
    expect(result.current.selectionConfirmed).toBe(true);
  });

  it("should choose folder scope based on selected folder sheets", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.applyFolderCatalog(folderCatalog));
    act(() => result.current.toggleFolderSheet("file-1", "Data"));
    act(() => result.current.chooseFolderScope());

    expect(result.current.sourceMode).toBe("folder");
    expect(result.current.selectionConfirmed).toBe(true);
  });

  it("should reset folder selection when applying a new catalog", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.setFolderCatalog(folderCatalog));
    act(() => result.current.setFolderSheetKeys(["file-1\0Data"]));
    act(() => result.current.applyFolderCatalog(folderCatalog));

    expect(result.current.folderCatalog).toBe(folderCatalog);
    expect(result.current.folderSheetKeys).toEqual([]);
    expect(result.current.selectionConfirmed).toBe(false);
  });

  it("should derive folder selections from selected keys", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.setFolderCatalog(folderCatalog));
    act(() =>
      result.current.setFolderSheetKeys(["file-1\0Data", "file-2\0Summary"])
    );

    expect(result.current.folderSelections()).toEqual([
      { fileId: "file-1", sheets: ["Data"] },
      { fileId: "file-2", sheets: ["Summary"] }
    ]);
  });

  it("should select names for a workbook snapshot", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.chooseAutomaticScope(workbook));
    expect(result.current.selectedNamesFor(workbook)).toEqual(["Sheet1"]);

    act(() => result.current.chooseManualScope(workbook));
    expect(result.current.selectedNamesFor(workbook)).toEqual(["Sheet1"]);
  });

  it("should select all and clear selected sheets", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.selectAllSheets(["Sheet1", "Sheet2"]));
    expect(result.current.selectedSheetNames).toEqual(["Sheet1", "Sheet2"]);
    expect(result.current.selectionConfirmed).toBe(false);

    act(() => result.current.clearSelectedSheets());
    expect(result.current.selectedSheetNames).toEqual([]);
  });
});
