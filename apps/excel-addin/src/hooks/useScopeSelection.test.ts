// @vitest-environment jsdom
import { describe, it, expect } from "vitest";
import { renderHook, act } from "@testing-library/react";
import type {
  FolderCatalog,
  WorkbookSnapshot
} from "../contracts";
import {
  addedFolderSheetKeys,
  resolveFolderViewAfterFileClear,
  useScopeSelection
} from "./useScopeSelection";

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
  it("should leave selected-only view when clearing the last selected file", () => {
    expect(
      resolveFolderViewAfterFileClear(
        ["file-1\0Data"],
        "file-1",
        ["file-1", "file-2"],
        true
      )
    ).toEqual({
      remainingKeys: [],
      showSelected: false,
      activeFileId: "file-1"
    });
  });

  it("should keep selected-only view and move to the next selected file", () => {
    expect(
      resolveFolderViewAfterFileClear(
        ["file-1\0Data", "file-2\0Summary"],
        "file-1",
        ["file-1", "file-2"],
        true
      )
    ).toEqual({
      remainingKeys: ["file-2\0Summary"],
      showSelected: true,
      activeFileId: "file-2"
    });
  });

  it("should preserve the all-files view when clearing one file", () => {
    expect(
      resolveFolderViewAfterFileClear(
        ["file-1\0Data", "file-2\0Summary"],
        "file-1",
        ["file-1", "file-2"],
        false
      )
    ).toEqual({
      remainingKeys: ["file-2\0Summary"],
      showSelected: false,
      activeFileId: "file-1"
    });
  });

  it("should identify worksheets added by a folder refresh", () => {
    const refreshed: FolderCatalog = {
      ...folderCatalog,
      files: folderCatalog.files.map((file) =>
        file.id === "file-1"
          ? {
              ...file,
              worksheets: [
                ...file.worksheets,
                { name: "新增", rowCount: 1, columnCount: 1 },
                { name: "#EB_RULES", rowCount: 1, columnCount: 1 }
              ]
            }
          : file
      )
    };

    expect(addedFolderSheetKeys(folderCatalog, refreshed)).toEqual([
      "file-1\0新增"
    ]);
    expect(addedFolderSheetKeys(null, refreshed)).toEqual([]);
  });

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

  it("should atomically adopt selected sheets from the current workbook", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.chooseFolderScope());
    act(() =>
      result.current.adoptCurrentWorkbookScope(workbook, ["Sheet1", "Sheet2"])
    );

    expect(result.current.sourceMode).toBe("workbook");
    expect(result.current.workbookScopeMode).toBe("manual");
    expect(result.current.selectedSheetNames).toEqual(["Sheet1", "Sheet2"]);
    expect(result.current.selectionConfirmed).toBe(true);
  });

  it("should confirm the adopted scope regardless of the previous selection", () => {
    // 这条锁住闭包旧值：若把 adoptCurrentWorkbookScope 退化成
    // chooseManualScope() + setSelectedSheetNames() 的复合调用，
    // chooseManualScope 会读到闭包里的空 selectedSheetNames，
    // selectionConfirmed 就会按旧值算错。
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.setSelectedSheetNames([]));
    act(() =>
      result.current.adoptCurrentWorkbookScope(workbook, ["Sheet2"])
    );

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

  it("should keep still-valid folder selection when applying a new catalog", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.setFolderCatalog(folderCatalog));
    act(() => result.current.setFolderSheetKeys(["file-1\0Data"]));
    act(() => result.current.applyFolderCatalog(folderCatalog));

    expect(result.current.folderCatalog).toBe(folderCatalog);
    // 勾选在新 catalog 中依然存在，保留 —— 支撑刷新页面后的会话恢复。
    expect(result.current.folderSheetKeys).toEqual(["file-1\0Data"]);
    expect(result.current.selectionConfirmed).toBe(false);
  });

  it("should drop folder selection that no longer exists in the new catalog", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.setFolderCatalog(folderCatalog));
    act(() =>
      result.current.setFolderSheetKeys([
        "file-1\0Data",
        "file-1\0Gone",
        "file-404\0Data"
      ])
    );
    act(() => result.current.applyFolderCatalog(folderCatalog));

    expect(result.current.folderSheetKeys).toEqual(["file-1\0Data"]);
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

describe("useScopeSelection：草稿与已提交授权分开", () => {
  it("勾选只改草稿：取消勾选再选回，已提交授权一字未动", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.chooseAutomaticScope(workbook));
    expect(result.current.committedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.scopeDraftDirty).toBe(false);

    // 用户"顺手点一下另一张表看看" —— 这不该被当成改了授权。
    act(() => result.current.toggleSheet("Sheet2"));
    expect(result.current.selectedSheetNames).toEqual(["Sheet1", "Sheet2"]);
    expect(result.current.committedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.scopeDraftDirty).toBe(true);

    // 取消选回：草稿回到已提交范围，且重新算成"没有未确认修改"。
    act(() => result.current.toggleSheet("Sheet2"));
    expect(result.current.selectedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.committedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.scopeDraftDirty).toBe(false);
  });

  it("取消范围编辑把草稿放回已提交授权", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.chooseAutomaticScope(workbook));
    act(() => result.current.toggleSheet("Sheet2"));
    expect(result.current.scopeDraftDirty).toBe(true);

    act(() => result.current.cancelScopeEdit());
    expect(result.current.selectedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.committedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.scopeDraftDirty).toBe(false);
    expect(result.current.selectionConfirmed).toBe(true);
  });

  it("确认范围按当前真实工作簿复核：已经不存在的工作表不会被提交成授权", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.chooseAutomaticScope(workbook));
    act(() => result.current.setSelectedSheetNames(["Sheet1", "Gone"]));
    act(() => {
      // 弹层可能开了很久：确认时按**当前**工作簿清单复核，而不是开弹层时的快照。
      const committed = result.current.confirmScopeSelection(workbook);
      expect(committed).toEqual(["Sheet1"]);
    });

    expect(result.current.committedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.selectedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.scopeDraftDirty).toBe(false);
  });

  it("确认范围时草稿里的表全部不存在则回落到活动工作表", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.setSelectedSheetNames(["Gone", "AlsoGone"]));
    act(() => {
      expect(result.current.confirmScopeSelection(workbook)).toEqual([
        "Sheet1"
      ]);
    });

    expect(result.current.committedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.selectionConfirmed).toBe(true);
  });

  it("确认范围不带快照时原样提交草稿", () => {
    const { result } = renderHook(() => useScopeSelection());

    act(() => result.current.chooseManualScope(workbook));
    act(() => result.current.toggleSheet("Sheet2"));
    act(() => {
      expect(result.current.confirmScopeSelection()).toEqual([
        "Sheet1",
        "Sheet2"
      ]);
    });

    expect(result.current.committedSheetNames).toEqual(["Sheet1", "Sheet2"]);
    expect(result.current.scopeDraftDirty).toBe(false);
  });
});
