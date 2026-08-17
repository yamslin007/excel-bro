import { useCallback, useState } from "react";
import type {
  FolderCatalog,
  FolderSelection,
  WorkbookSnapshot
} from "../contracts";
import type { SourceMode, WorkbookScopeMode } from "../types/workbook";
import { folderSheetKey } from "../utils";

/**
 * 数据范围与工具保存字段 Hook
 *
 * 职责：
 * - 管理数据源模式（workbook/folder）和工作簿范围模式（auto/manual）
 * - 管理工作簿/文件夹中的工作表选择状态
 * - 管理工具保存对话框中的名称与用途说明
 * - 提供范围切换、工作表勾选和文件夹选择等操作
 */
export function useScopeSelection() {
  const [toolName, setToolName] = useState("");
  const [toolDescription, setToolDescription] = useState("");
  const [selectedSheetNames, setSelectedSheetNames] = useState<string[]>([]);
  const [selectionConfirmed, setSelectionConfirmed] = useState(false);
  const [sourceMode, setSourceMode] = useState<SourceMode>("workbook");
  const [workbookScopeMode, setWorkbookScopeMode] =
    useState<WorkbookScopeMode>("auto");
  const [folderCatalog, setFolderCatalog] = useState<FolderCatalog | null>(null);
  const [folderSheetKeys, setFolderSheetKeys] = useState<string[]>([]);

  const applyWorkbookSnapshotSelection = useCallback(
    (snapshot: WorkbookSnapshot) => {
      setSelectedSheetNames((current) => {
        if (sourceMode === "workbook" && workbookScopeMode === "auto") {
          return [snapshot.activeWorksheet];
        }
        const available = new Set(
          snapshot.worksheets.map((sheet) => sheet.name)
        );
        const preserved = current.filter((name) => available.has(name));
        return preserved.length > 0 ? preserved : [snapshot.activeWorksheet];
      });
      setSelectionConfirmed(true);
    },
    [sourceMode, workbookScopeMode]
  );

  const toggleSheet = useCallback((sheetName: string) => {
    setSelectedSheetNames((current) =>
      current.includes(sheetName)
        ? current.filter((name) => name !== sheetName)
        : [...current, sheetName]
    );
    setSelectionConfirmed(false);
  }, []);

  const toggleFolderSheet = useCallback((fileId: string, sheetName: string) => {
    const key = folderSheetKey(fileId, sheetName);
    setFolderSheetKeys((current) =>
      current.includes(key)
        ? current.filter((item) => item !== key)
        : [...current, key]
    );
    setSelectionConfirmed(false);
  }, []);

  const applyFolderCatalog = useCallback((catalog: FolderCatalog) => {
    setFolderCatalog(catalog);
    setFolderSheetKeys([]);
    setSelectionConfirmed(false);
  }, []);

  const chooseAutomaticScope = useCallback(
    (workbook: WorkbookSnapshot | null) => {
      setSourceMode("workbook");
      setWorkbookScopeMode("auto");
      if (workbook) {
        setSelectedSheetNames([workbook.activeWorksheet]);
      }
      setSelectionConfirmed(true);
    },
    []
  );

  const chooseManualScope = useCallback(
    (workbook: WorkbookSnapshot | null) => {
      setSourceMode("workbook");
      setWorkbookScopeMode("manual");
      if (selectedSheetNames.length === 0 && workbook) {
        setSelectedSheetNames([workbook.activeWorksheet]);
      }
      setSelectionConfirmed(selectedSheetNames.length > 0 || workbook !== null);
    },
    [selectedSheetNames]
  );

  const chooseFolderScope = useCallback(() => {
    setSourceMode("folder");
    setSelectionConfirmed(folderSheetKeys.length > 0);
  }, [folderSheetKeys]);

  const folderSelections = useCallback((): FolderSelection[] => {
    if (!folderCatalog) return [];
    return folderCatalog.files
      .map((file) => ({
        fileId: file.id,
        sheets: file.worksheets
          .filter((sheet) =>
            folderSheetKeys.includes(folderSheetKey(file.id, sheet.name))
          )
          .map((sheet) => sheet.name)
      }))
      .filter((selection) => selection.sheets.length > 0);
  }, [folderCatalog, folderSheetKeys]);

  const selectedNamesFor = useCallback(
    (snapshot: WorkbookSnapshot): string[] => {
      if (sourceMode === "workbook" && workbookScopeMode === "auto") {
        return [snapshot.activeWorksheet];
      }
      const available = new Set(
        snapshot.worksheets.map((sheet) => sheet.name)
      );
      const selected = selectedSheetNames.filter((name) =>
        available.has(name)
      );
      return selected.length > 0 ? selected : [snapshot.activeWorksheet];
    },
    [sourceMode, workbookScopeMode, selectedSheetNames]
  );

  const selectAllSheets = useCallback((sheetNames: string[]) => {
    setSelectedSheetNames((current) => [
      ...new Set([...current, ...sheetNames])
    ]);
    setSelectionConfirmed(false);
  }, []);

  const clearSelectedSheets = useCallback(() => {
    setSelectedSheetNames([]);
    setSelectionConfirmed(false);
  }, []);

  return {
    // 工具保存字段
    toolName,
    setToolName,
    toolDescription,
    setToolDescription,

    // 范围状态
    selectedSheetNames,
    setSelectedSheetNames,
    selectionConfirmed,
    setSelectionConfirmed,
    sourceMode,
    setSourceMode,
    workbookScopeMode,
    setWorkbookScopeMode,
    folderCatalog,
    setFolderCatalog,
    folderSheetKeys,
    setFolderSheetKeys,

    // 操作
    applyWorkbookSnapshotSelection,
    toggleSheet,
    toggleFolderSheet,
    applyFolderCatalog,
    chooseAutomaticScope,
    chooseManualScope,
    chooseFolderScope,
    folderSelections,
    selectedNamesFor,
    selectAllSheets,
    clearSelectedSheets
  };
}
