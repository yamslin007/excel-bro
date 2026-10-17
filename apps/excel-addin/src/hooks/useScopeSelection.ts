import { useCallback, useEffect, useState, type SetStateAction } from "react";
import type {
  FolderCatalog,
  FolderSelection,
  WorkbookSnapshot
} from "../contracts";
import type { SourceMode, WorkbookScopeMode } from "../types/workbook";
import {
  addedFolderSheetKeys,
  canPreserveFolderAuthorization,
  folderScopeDraftDirty,
  folderSelectionsFromKeys,
  folderSheetKey,
  folderSheetKeysOf,
  isEBSystemSheet,
  resolveFolderViewAfterFileClear,
  type FolderViewAfterFileClear
} from "@excel-bro/core/folderScope";

// 这些判据全在共享层（`@excel-bro/core/folderScope`）。这里原样转发，
// 让 App.tsx 与既有用例的导入路径（`./useScopeSelection`）不用动，
// 同时保证两端读到的是**同一份**声明而不是"行为相似的副本"。
export { addedFolderSheetKeys, resolveFolderViewAfterFileClear };
export type { FolderViewAfterFileClear };

const FOLDER_CATALOG_KEY = "excelBro.folderCatalog";
const FOLDER_SHEET_KEYS_KEY = "excelBro.folderSheetKeys";

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
  /**
   * 当前草稿里的工作表选择。勾选、取消勾选、全选/清空都只改它 ——
   * 这些动作**不构成授权**，所以它们不该顺手作废一件正在等待回答的任务。
   */
  const [selectedSheetNames, setDraftSheetNames] = useState<string[]>([]);
  /**
   * **已提交**的数据授权。只有点「确认范围」或明确选用某个范围时才变。
   *
   * 它与草稿刻意分开：用户为了看一眼别的表而勾上又取消，不该被当成
   * "他改了授权"。真正的授权变更必须有一个人明确的确认动作。
   */
  const [committedSheetNames, setCommittedSheetNames] = useState<string[]>([]);
  const [selectionConfirmed, setSelectionConfirmed] = useState(false);
  // Exposed selection setters edit a draft, never silently advance authorization.
  const setSelectedSheetNames = useCallback((update: SetStateAction<string[]>) => {
    setDraftSheetNames(update);
    setSelectionConfirmed(false);
  }, []);
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

  /**
   * 显式**提交**当前草稿为这一版数据授权。
   *
   * 给出 `snapshot` 时按**当前真实工作表清单**复核草稿：已经不存在的工作表
   * 不会被提交成授权（弹层可能开了很久，期间用户可能删过表或换了工作簿）。
   * 返回真正被提交的名称，调用方据此展示"这次用的是哪些表"。
   */
  const confirmScopeSelection = useCallback(
    (snapshot?: WorkbookSnapshot | null): string[] => {
      let next = selectedSheetNames;
      if (snapshot) {
        const available = new Set(snapshot.worksheets.map((sheet) => sheet.name));
        if (!next.length || next.some((name) => !available.has(name))) {
          setSelectionConfirmed(false);
          return [];
        }
      }
      setCommittedSheetNames(next);
      setSelectionConfirmed(next.length > 0);
      return next;
    },
    [selectedSheetNames]
  );

  /**
   * 放弃本次范围编辑：草稿回到已提交授权，**已提交授权一字不改**。
   *
   * 这是"我点开只是想看看"的出口 —— 它不该把任何正在等待回答的任务打掉。
   */
  const cancelScopeEdit = useCallback(() => {
    setSelectedSheetNames(committedSheetNames);
    setSelectionConfirmed(committedSheetNames.length > 0);
  }, [committedSheetNames]);

  /**
   * 草稿与已提交授权是否已经不一致。
   *
   * 界面据此提示"还有未确认的范围修改"，而不是让用户以为改了就已经生效。
   */
  const scopeDraftDirty = folderScopeDraftDirty(
    selectedSheetNames,
    committedSheetNames
  );

  /**
   * 凡是把范围标成"已确认"的路径，都同步把当前选择记成**已提交授权**。
   *
   * 这样既有行为（选自动范围 / 采用当前工作簿 / 选文件夹）都自动获得
   * "已提交"的语义，不必逐个回调去改；而只改草稿的路径（toggleSheet 等）
   * 会把 `selectionConfirmed` 置回 false，因此不会走到这里。
   */
  useEffect(() => {
    if (selectionConfirmed) setCommittedSheetNames(selectedSheetNames);
  }, [selectionConfirmed, selectedSheetNames]);

  const toggleSheet = useCallback((sheetName: string) => {
    setSelectedSheetNames((current) =>
      current.includes(sheetName)
        ? current.filter((name) => name !== sheetName)
        : [...current, sheetName]
    );
    // 只改草稿：`committedSheetNames` 不动，重新确认前它仍是原来的授权。
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

  const applyFolderCatalog = useCallback((
    catalog: FolderCatalog,
    options: { preserveAuthorization?: boolean } = {}
  ) => {
    setFolderCatalog(catalog);
    // 保留仍然有效的勾选（文件 ID 和表名都能在新 catalog 中找到）。
    // 这里用的是**存在性**名单（不过滤 EB 表）：一个真存在的键被过滤掉
    // 会让下一次刷新把它误判成「目标已消失」。
    const validFileSheets = new Set(folderSheetKeysOf(catalog));
    setFolderSheetKeys((current) =>
      current.filter((key) => validFileSheets.has(key))
    );
    // Refresh is not consent. Preserve an existing confirmation only; never
    // turn a draft or a restored localStorage selection into authorization.
    // 判据本身在共享层：同根 + 同会话 + 至少勾一张 + 每张都还在。
    const sameTargets = options.preserveAuthorization === true &&
      canPreserveFolderAuthorization(folderCatalog, catalog, folderSheetKeys);
    if (!sameTargets) setSelectionConfirmed(false);
  }, [folderCatalog, folderSheetKeys]);

  const chooseAutomaticScope = useCallback(
    (workbook: WorkbookSnapshot | null) => {
      setSourceMode("workbook");
      setWorkbookScopeMode("auto");
      if (workbook) {
        setSelectedSheetNames([workbook.activeWorksheet]);
      }
      setSelectionConfirmed(false);
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
      setSelectionConfirmed(false);
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
    if (sourceMode !== "folder") setSelectionConfirmed(false);
  }, [sourceMode]);

  const folderSelections = useCallback((): FolderSelection[] => {
    if (!folderCatalog) return [];
    return folderSelectionsFromKeys(folderCatalog, folderSheetKeys);
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

    // 已提交授权与显式确认
    committedSheetNames,
    scopeDraftDirty,
    confirmScopeSelection,
    cancelScopeEdit,

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
