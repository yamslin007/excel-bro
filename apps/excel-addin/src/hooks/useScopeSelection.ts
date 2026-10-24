import { recordDiagnosticEvent } from "../diagnostics";
import { useCallback, useEffect, useRef, useState, type SetStateAction } from "react";
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
  /** 界面勾选。选择事件由 App 自动核实并提交；后台刷新仅更新目录，不构成授权。 */
  const [selectedSheetNames, setDraftSheetNames] = useState<string[]>([]);
  /** 已核实的授权。与界面选择分开，防止异步读取期间沿用被取消的勾选。 */
  const [committedSheetNames, setCommittedSheetNames] = useState<string[]>([]);
  /** 已核实的范围模式，与授权表名同批提交。自动与固定模式不能混用。 */
  const [committedScopeMode, setCommittedScopeMode] =
    useState<WorkbookScopeMode>("auto");
  const [selectionConfirmed, setSelectionConfirmed] = useState(false);
  /** 首读只播种一次；此后"已提交"只能由显式选择的核验结果写。 */
  const committedSeededRef = useRef(false);
  /** 现读已提交名单（供"要不要播种"这类"当下事实"判断用，不吃闭包旧值）。 */
  const committedNamesRef = useRef(committedSheetNames);
  committedNamesRef.current = committedSheetNames;
  // Exposed selection setters edit a draft, never silently advance authorization.
  const setSelectedSheetNames = useCallback((update: SetStateAction<string[]>) => {
    setDraftSheetNames(update);
    setSelectionConfirmed(false);
  }, []);
  const [sourceMode, setSourceMode] = useState<SourceMode>("workbook");
  const [workbookScopeMode, setWorkbookScopeMode] =
    useState<WorkbookScopeMode>("auto");
  // Local diagnostics contain counts/modes only, never workbook names or rows.
  useEffect(() => {
    recordDiagnosticEvent({ timestamp: new Date().toISOString(), phase: 'scope_selection',
      durationMs: 0, modelCalls: 0, status: 'succeeded',
      scopeSelection: { mode: workbookScopeMode, selectedCount: selectedSheetNames.length,
        committedMode: committedScopeMode, committedCount: committedSheetNames.length } });
  }, [workbookScopeMode, selectedSheetNames, committedScopeMode, committedSheetNames]);
  // Async task cleanup may retain an older scan callback. Read the current mode.
  const liveScope = useRef({ sourceMode, workbookScopeMode });
  liveScope.current = { sourceMode, workbookScopeMode };
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

  /**
   * 提交一次数据授权：**范围与口径同批写进"已提交"**。
   *
   * 这是唯一一条写 `committedSheetNames` / `committedScopeMode` 的路（首读播种除外）。
   * 守卫（`App` 的 `inspect`）读的就是这两个值：它们必须与守卫基线在**同一次**
   * 交互里一起推进。分成两步写——先 bind 基线、再等一个 effect 把草稿抄成已提交——
   * 中间那一帧读到的是"新基线 + 旧范围"，于是用户刚确认完就被判成范围变化。
   */
  const commitAuthorizedScope = useCallback(
    (sheetNames: string[], mode?: WorkbookScopeMode) => {
      setCommittedSheetNames([...sheetNames]);
      if (mode) setCommittedScopeMode(mode);
      setSelectionConfirmed(true);
    },
    []
  );

  /**
   * 重读宿主后刷新草稿。
   *
   * **重读不是授权动作**：它只改草稿。唯一的例外是**第一次**读到工作簿
   * （此前没有任何已提交范围）—— 那时把默认范围记为已提交，否则"打开面板直接
   * 提问"会被"请先确认数据范围"拦在门口。之后无论是刷新、结构变化还是目录
   * 刷新，已提交范围都只能由 `commitAuthorizedScope` 改 —— 后台读一次就把
   * 用户的授权换掉，正是"刚确认完又要求确认"的来源。
   */
  const applyWorkbookSnapshotSelection = useCallback(
    (snapshot: WorkbookSnapshot) => {
      const fallback = [snapshot.activeWorksheet];
      const mode = liveScope.current;
      if (mode.sourceMode !== 'workbook') return;
      // A catalog refresh is not a selection gesture. Preserve manual choices,
      // including missing sheets; admission/target guards explain missing targets.
      if (mode.workbookScopeMode === 'auto') setSelectedSheetNames(fallback);
      if (!committedSeededRef.current && committedNamesRef.current.length === 0) {
        committedSeededRef.current = true;
        setCommittedSheetNames(fallback);
        setCommittedScopeMode(mode.workbookScopeMode);
      }
    },
    []
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
      if (!next.length) {
        setSelectionConfirmed(false);
        return next;
      }
      commitAuthorizedScope(next, workbookScopeMode);
      return next;
    },
    [selectedSheetNames, workbookScopeMode, commitAuthorizedScope]
  );

  /**
   * 放弃本次范围编辑：草稿回到已提交授权（**包括口径**），已提交授权一字不改。
   *
   * 这是"我点开只是想看看"的出口 —— 它不该把任何正在等待回答的任务打掉，
   * 也不该把"跟随当前 / 选择多个"的口径留在草稿里：口径是授权的一部分，
   * 留在草稿里会让守卫拿另一套口径去比已提交的范围。
   */
  const cancelScopeEdit = useCallback(() => {
    setSelectedSheetNames(committedSheetNames);
    setWorkbookScopeMode(committedScopeMode);
    setSelectionConfirmed(committedSheetNames.length > 0);
  }, [committedSheetNames, committedScopeMode]);

  /**
   * 草稿与已提交授权是否已经不一致 —— **口径也算**。
   *
   * 只比表名是不够的：把"固定多表"改成"跟随当前"（或反过来）同样是换了范围，
   * 只是换法不一样。口径留在草稿里最危险的一处是**读取**：发送时按草稿口径
   * 决定读活动表还是读已提交名单，若那时已被放行，就会读一张并没有被授权过的
   * 活动表。WPS 侧早就是这条判据（`draftMode !== scopeMode`），两端同源。
   */
  const scopeDraftDirty =
    folderScopeDraftDirty(selectedSheetNames, committedSheetNames) ||
    (sourceMode === "workbook" && workbookScopeMode !== committedScopeMode);

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
    (
      _snapshot: WorkbookSnapshot,
      sheetNames: string[],
      /**
       * 这次采用用的是什么口径。默认「固定多表」（旧行为），但"跟随当前"的确认
       * 也必须把 `auto` 写进已提交 —— 否则守卫会按 `manual` 去读已提交的名字，
       * 而用户确认的是"跟随活动表"。
       */
      mode: WorkbookScopeMode = "manual"
    ) => {
      setSourceMode("workbook");
      setWorkbookScopeMode(mode);
      setSelectedSheetNames(sheetNames);
      // 草稿与已提交**同批**落地：守卫在这一帧之后立刻就会读到它们。
      commitAuthorizedScope(sheetNames, mode);
    },
    [commitAuthorizedScope]
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
      return selected;
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
    /** 已提交的**口径**：守卫按它决定"当前范围"读活动表还是读已提交名单。 */
    committedScopeMode,
    scopeDraftDirty,
    confirmScopeSelection,
    cancelScopeEdit,
    /** 唯一一条写"已提交授权（范围 + 口径）"的路。 */
    commitAuthorizedScope,

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
