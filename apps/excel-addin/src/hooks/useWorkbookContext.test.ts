// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import type {
  FolderCatalog,
  WorkbookSnapshot
} from "../contracts";
import {
  captureWorkbookStructure,
  isRunningInExcel,
  watchWorkbookStructureChanges
} from "../excel";
import { selectFolder } from "../api";
import { folderSheetKey } from "../utils";
import { demoWorkbook } from "../demo";
import { extractWorkbookDataPeriod } from "../workbookIdentity";
import { useWorkbookContext } from "./useWorkbookContext";

vi.mock("../excel", () => ({
  captureWorkbookStructure: vi.fn(),
  isRunningInExcel: vi.fn(),
  watchWorkbookStructureChanges: vi.fn()
}));

vi.mock("../api", () => ({
  selectFolder: vi.fn()
}));

vi.mock("../utils", () => ({
  folderSheetKey: vi.fn((fileId: string, sheetName: string) =>
    `${fileId}\u0000${sheetName}`
  )
}));

vi.mock("../demo", () => ({
  demoWorkbook: {
    name: "demo.xlsx",
    capturedAt: "2026-08-16T00:00:00.000Z",
    activeWorksheet: "Demo",
    worksheets: [
      {
        name: "Demo",
        usedRange: "A1:B2",
        rowCount: 2,
        columnCount: 2,
        headers: ["Name", "Amount"],
        dataRows: [],
        truncated: false
      }
    ]
  }
}));

vi.mock("../workbookIdentity", () => ({
  extractWorkbookDataPeriod: vi.fn(() => "2026")
}));

const captureWorkbookStructureMock = vi.mocked(captureWorkbookStructure);
const isRunningInExcelMock = vi.mocked(isRunningInExcel);
const watchWorkbookStructureChangesMock = vi.mocked(
  watchWorkbookStructureChanges
);
const selectFolderMock = vi.mocked(selectFolder);
const folderSheetKeyMock = vi.mocked(folderSheetKey);
const extractWorkbookDataPeriodMock = vi.mocked(extractWorkbookDataPeriod);

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
      worksheets: [{ name: "Data", rowCount: 10, columnCount: 3 }]
    }
  ],
  totalFiles: 1,
  truncated: false,
  expiresAt: "2026-08-16T01:00:00.000Z"
};

let readyCallback: (() => Promise<void> | void) | undefined;

describe("useWorkbookContext", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    readyCallback = undefined;
    captureWorkbookStructureMock.mockResolvedValue(workbook);
    isRunningInExcelMock.mockReturnValue(true);
    watchWorkbookStructureChangesMock.mockResolvedValue(vi.fn());
    selectFolderMock.mockResolvedValue(folderCatalog);
    folderSheetKeyMock.mockImplementation((fileId, sheetName) =>
      `${fileId}\u0000${sheetName}`
    );
    extractWorkbookDataPeriodMock.mockReturnValue("2026");
    vi.stubGlobal("Office", {
      onReady: vi.fn((callback: () => Promise<void> | void) => {
        readyCallback = callback;
      })
    });
  });

  async function flushOfficeReady() {
    await act(async () => {
      await readyCallback?.();
      await Promise.resolve();
    });
  }

  it("should initialize with null workbook and empty selections", () => {
    const { result } = renderHook(() => useWorkbookContext());

    expect(result.current.workbook).toBeNull();
    expect(result.current.selectedSheetNames).toEqual([]);
    expect(result.current.selectionConfirmed).toBe(false);
    expect(result.current.sourceMode).toBe("workbook");
    expect(result.current.workbookScopeMode).toBe("auto");
  });

  it("should use demo workbook when Office is unavailable", async () => {
    vi.stubGlobal("Office", undefined);
    const { result } = renderHook(() => useWorkbookContext());
    await act(async () => {
      await Promise.resolve();
    });

    expect(result.current.workbook?.name).toBe("demo.xlsx");
    expect(result.current.selectedSheetNames).toEqual(["Demo"]);
  });

  it("should scan workbook and watch structure changes on ready", async () => {
    const dispose = vi.fn();
    watchWorkbookStructureChangesMock.mockResolvedValue(dispose);
    const { result } = renderHook(() => useWorkbookContext());

    await flushOfficeReady();

    expect(captureWorkbookStructureMock).toHaveBeenCalled();
    expect(result.current.workbook).toEqual(workbook);
    expect(result.current.selectedSheetNames).toEqual(["Sheet1"]);
    expect(result.current.selectionConfirmed).toBe(true);
    expect(watchWorkbookStructureChangesMock).toHaveBeenCalled();
  });

  it("should derive workbook data period", async () => {
    const { result } = renderHook(() => useWorkbookContext());

    await flushOfficeReady();

    expect(extractWorkbookDataPeriodMock).toHaveBeenCalledWith("book.xlsx");
    expect(result.current.workbookDataPeriod).toBe("2026");
  });

  it("should toggle workbook sheets and unconfirm selection", async () => {
    const { result } = renderHook(() => useWorkbookContext());
    await flushOfficeReady();

    act(() => result.current.toggleSheet("Sheet2"));

    expect(result.current.selectedSheetNames).toContain("Sheet2");
    expect(result.current.selectionConfirmed).toBe(false);
  });

  it("should toggle folder sheets and unconfirm selection", () => {
    const { result } = renderHook(() => useWorkbookContext());

    act(() => result.current.setFolderCatalog(folderCatalog));
    act(() => result.current.toggleFolderSheet("file-1", "Data"));

    expect(result.current.folderSheetKeys).toEqual(["file-1\0Data"]);
    expect(result.current.selectionConfirmed).toBe(false);
  });

  it("should browse and apply folder catalog", async () => {
    const { result } = renderHook(() => useWorkbookContext());

    await act(async () => {
      await result.current.browseFolder({
        onMessage: vi.fn(),
        onServerOnline: vi.fn()
      });
    });

    expect(result.current.folderCatalog).toEqual(folderCatalog);
    expect(result.current.folderSheetKeys).toEqual([]);
    expect(result.current.selectionConfirmed).toBe(false);
  });

  it("should choose automatic, manual and folder scope", async () => {
    const { result } = renderHook(() => useWorkbookContext());
    await flushOfficeReady();

    act(() => result.current.chooseAutomaticScope());
    expect(result.current.sourceMode).toBe("workbook");
    expect(result.current.workbookScopeMode).toBe("auto");

    act(() => result.current.chooseManualScope());
    expect(result.current.workbookScopeMode).toBe("manual");

    act(() => result.current.chooseFolderScope());
    expect(result.current.sourceMode).toBe("folder");
  });

  it("should derive folder selections", async () => {
    const { result } = renderHook(() => useWorkbookContext());
    await flushOfficeReady();

    act(() => result.current.setFolderCatalog(folderCatalog));
    act(() =>
      result.current.setFolderSheetKeys(["file-1\0Data"])
    );

    expect(result.current.getFolderSelections()).toEqual([
      { fileId: "file-1", sheets: ["Data"] }
    ]);
  });

  it("should not leak watcher when unmounted before ready", async () => {
    const dispose = vi.fn();
    watchWorkbookStructureChangesMock.mockResolvedValue(dispose);
    const view = renderHook(() => useWorkbookContext());

    view.unmount();
    await act(async () => {
      await readyCallback?.();
      await Promise.resolve();
    });

    expect(captureWorkbookStructureMock).not.toHaveBeenCalled();
    expect(watchWorkbookStructureChangesMock).not.toHaveBeenCalled();
  });
});
