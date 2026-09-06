import { useCallback, useEffect, useState } from "react";
import type {
  FolderCatalog,
  FolderSelection,
  WorkbookSnapshot
} from "../contracts";
import type { SourceMode, WorkbookScopeMode } from "../types/workbook";
import { folderSheetKey } from "../utils";
import { isEBSystemSheet } from "../excel";

const FOLDER_CATALOG_KEY = "excelBro.folderCatalog";
const FOLDER_SHEET_KEYS_KEY = "excelBro.folderSheetKeys";

/** 比较两次文件夹扫描，返回新出现且允许用户选择的工作表。 */
export function addedFolderSheetKeys(
  previous: FolderCatalog | null,
  current: FolderCatalog
): string[] {
  if (!previous) return [];
  const previousKeys = new Set(
    previous.files.flatMap((file) =>
      file.worksheets
        .filter((sheet) => !isEBSystemSheet(sheet.name))
        .map((sheet) => folderSheetKey(file.id, sheet.name))
    )
  );
  return current.files.flatMap((file) =>
    file.worksheets
      .filter((sheet) => !isEBSystemSheet(sheet.name))
      .map((sheet) => folderSheetKey(file.id, sheet.name))
      .filter((key) => !previousKeys.has(key))
  );
}

export interface FolderViewAfterFileClear {
  remainingKeys: string[];
  showSelected: boolean;
  activeFileId: string;
}

/**
 * 清空当前文件后修正文件夹选择视图，避免「只看已选」落入无法退出的空结果。
 */
export function resolveFolderViewAfterFileClear(
  selectedKeys: string[],
  clearedFileId: string,
  orderedFileIds: string[],
  showSelected: boolean
): FolderViewAfterFileClear {
  const clearedPrefix = folderSheetKey(clearedFileId, "");
  const remainingKeys = selectedKeys.filter(
    (key) => !key.startsWith(clearedPrefix)
  );
  if (!showSelected) {
    return {
      remainingKeys,
      showSelected: false,
      activeFileId: clearedFileId
    };
  }
  const nextSelectedFileId = orderedFileIds.find((fileId) => {
    const prefix = folderSheetKey(fileId, "");
    return remainingKeys.some((key) => key.startsWith(prefix));
  });
  return {
    remainingKeys,
    showSelected: Boolean(nextSelectedFileId),
    activeFileId: nextSelectedFileId ?? clearedFileId
  };
}

/**
 * 数据范围与工具保存字段 Hook
 *
 * 职责：
 * - 管理数据源模式（workbook/folder）和工作簿范围模式（auto/manual）
 * - 管理工作簿/文件夹中的工作表选择状态
 * - 管理工具保存对话框中的名称与用途说明
 * - 提供范围切换、工作表勾选和文件夹选择等操作
 * - 持久化文件夹会话到 localStorage（30分钟内刷新页面不丢失）
 */
export function useScopeSelection() {
  const [toolName, setToolName] = useState("");
  const [toolDescription, setToolDescription] = useState("");
  const [selectedSheetNames, setSelectedSheetNames] = useState<string[]>([]);
  const [selectionConfirmed, setSelectionConfirmed] = useState(false);
  const [sourceMode, setSourceMode] = useState<SourceMode>("workbook");
  const [workbookScopeMode, setWorkbookScopeMode] =
    useState<WorkbookScopeMode>("auto");
  const [folderCatalog, setFolderCatalog] = useState<FolderCatalog | null>(() => {
    try {
      const stored = localStorage.getItem(FOLDER_CATALOG_KEY);
      return stored ? JSON.parse(stored) : null;
    } catch {
      return null;
    }
  });
  const [folderSheetKeys, setFolderSheetKeys] = useState<string[]>(() => {
    try {
      const stored = localStorage.getItem(FOLDER_SHEET_KEYS_KEY);
      return stored ? JSON.parse(stored) : [];
    } catch {
      return [];
    }
  });

  // 持久化 folderCatalog 到 localStorage
  useEffect(() => {
    try {
      if (folderCatalog) {
        localStorage.setItem(FOLDER_CATALOG_KEY, JSON.stringify(folderCatalog));
      } else {
        localStorage.removeItem(FOLDER_CATALOG_KEY);
      }
    } catch {
      // localStorage 写入失败（隐私模式/配额满）：静默忽略
    }
  }, [folderCatalog]);

  // 持久化 folderSheetKeys 到 localStorage
  useEffect(() => {
    try {
      if (folderSheetKeys.length > 0) {
        localStorage.setItem(FOLDER_SHEET_KEYS_KEY, JSON.stringify(folderSheetKeys));
      } else {
        localStorage.removeItem(FOLDER_SHEET_KEYS_KEY);
      }
    } catch {
      // localStorage 写入失败：静默忽略
    }
  }, [folderSheetKeys]);

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

  const selectAllSheetsInFile = useCallback((fileId: string) => {
    if (!folderCatalog) return;
    const file = folderCatalog.files.find(
      (candidate) => candidate.id === fileId
    );
    if (!file) return;
    const keys = file.worksheets
      .filter((sheet) => !isEBSystemSheet(sheet.name))
      .map((sheet) => folderSheetKey(fileId, sheet.name));
    setFolderSheetKeys((current) => [...new Set([...current, ...keys])]);
    setSelectionConfirmed(false);
  }, [folderCatalog]);

  const clearSheetsInFile = useCallback((fileId: string) => {
    const prefix = folderSheetKey(fileId, "");
    setFolderSheetKeys((current) =>
      current.filter((key) => !key.startsWith(prefix))
    );
    setSelectionConfirmed(false);
  }, []);

  const applyFolderCatalog = useCallback((catalog: FolderCatalog) => {
    setFolderCatalog(catalog);
    // 保留仍然有效的勾选（文件 ID 和表名都能在新 catalog 中找到）
    const validFileSheets = new Set<string>();
    for (const file of catalog.files) {
      for (const sheet of file.worksheets) {
        validFileSheets.add(folderSheetKey(file.id, sheet.name));
      }
    }
    setFolderSheetKeys((current) =>
      current.filter((key) => validFileSheets.has(key))
    );
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

  const adoptCurrentWorkbookScope = useCallback(
    (_snapshot: WorkbookSnapshot, sheetNames: string[]) => {
      setSourceMode("workbook");
      setWorkbookScopeMode("manual");
      setSelectedSheetNames(sheetNames);
      setSelectionConfirmed(true);
    },
    []
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
    selectAllSheetsInFile,
    clearSheetsInFile,
    applyFolderCatalog,
    chooseAutomaticScope,
    chooseManualScope,
    adoptCurrentWorkbookScope,
    chooseFolderScope,
    folderSelections,
    selectedNamesFor,
    selectAllSheets,
    clearSelectedSheets
  };
}
