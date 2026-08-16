import { useState, useEffect, useCallback, useMemo } from "react";
import type {
  WorkbookSnapshot,
  FolderCatalog,
  FolderSelection
} from "../contracts";
import {
  captureWorkbookStructure,
  watchWorkbookStructureChanges,
  isRunningInExcel
} from "../excel";
import { selectFolder } from "../api";
import { folderSheetKey } from "../utils";
import { demoWorkbook } from "../demo";
import { extractWorkbookDataPeriod } from "../workbookIdentity";
import type { DiagnosticEvent } from "../diagnostics";

export type SourceMode = "workbook" | "folder";
export type WorkbookScopeMode = "auto" | "manual";

/**
 * 工作簿上下文管理 Hook
 *
 * 职责：
 * - 管理工作簿快照（捕获、刷新、结构监听）
 * - 管理数据源模式（workbook vs folder）
 * - 管理工作表选择（自动/手动模式）
 * - 管理文件夹批量操作
 */
export function useWorkbookContext() {
  // 工作簿相关状态
  const [workbook, setWorkbook] = useState<WorkbookSnapshot | null>(null);
  const [selectedSheetNames, setSelectedSheetNames] = useState<string[]>([]);
  const [selectionConfirmed, setSelectionConfirmed] = useState(false);

  // 数据源模式
  const [sourceMode, setSourceMode] = useState<SourceMode>("workbook");
  const [workbookScopeMode, setWorkbookScopeMode] = useState<WorkbookScopeMode>("auto");

  // 文件夹模式相关
  const [folderCatalog, setFolderCatalog] = useState<FolderCatalog | null>(null);
  const [folderSheetKeys, setFolderSheetKeys] = useState<string[]>([]);

  /**
   * 派生状态：工作簿数据周期（从文件名提取）
   */
  const workbookDataPeriod = useMemo(
    () => (workbook ? extractWorkbookDataPeriod(workbook.name) : null),
    [workbook]
  );

  /**
   * 捕获工作簿快照
   *
   * @param options.announce - 是否在消息流中通知
   * @param options.onMessage - 消息回调
   * @param options.onDiagnostic - 诊断事件回调
   */
  const scan = useCallback(async (options?: {
    announce?: boolean;
    onMessage?: (text: string) => void;
    onDiagnostic?: (event: DiagnosticEvent) => void;
  }) => {
    const diagnosticStartedAt = performance.now();

    try {
      const snapshot = isRunningInExcel()
        ? await captureWorkbookStructure(
            sourceMode === "workbook" && workbookScopeMode === "manual"
              ? selectedSheetNames
              : undefined
          )
        : demoWorkbook;

      setWorkbook(snapshot);

      setSelectedSheetNames((current) => {
        if (sourceMode === "workbook" && workbookScopeMode === "auto") {
          return [snapshot.activeWorksheet];
        }
        const available = new Set(snapshot.worksheets.map((sheet) => sheet.name));
        const preserved = current.filter((name) => available.has(name));
        return preserved.length > 0 ? preserved : [snapshot.activeWorksheet];
      });

      setSelectionConfirmed(true);

      if (options?.announce && options.onMessage) {
        options.onMessage(
          `已重新读取「${snapshot.name}」：${snapshot.worksheets.length} 个工作表。`
        );
      }

      if (options?.onDiagnostic) {
        options.onDiagnostic({
          timestamp: new Date().toISOString(),
          phase: "scan",
          durationMs: performance.now() - diagnosticStartedAt,
          modelCalls: 0,
          status: "succeeded"
        });
      }

      return snapshot;
    } catch (reason) {
      if (options?.onDiagnostic) {
        options.onDiagnostic({
          timestamp: new Date().toISOString(),
          phase: "scan",
          durationMs: performance.now() - diagnosticStartedAt,
          modelCalls: 0,
          status: "failed",
          errorCategory: "data"
        });
      }

      const errorMessage = reason instanceof Error ? reason.message : "读取工作簿失败";

      if (options?.onMessage) {
        options.onMessage(errorMessage);
      }

      throw new Error(errorMessage);
    }
  }, [sourceMode, workbookScopeMode, selectedSheetNames]);

  /**
   * 切换工作表选择（workbook 模式）
   */
  const toggleSheet = useCallback((sheetName: string) => {
    setSelectedSheetNames((current) =>
      current.includes(sheetName)
        ? current.filter((name) => name !== sheetName)
        : [...current, sheetName]
    );
    setSelectionConfirmed(false);
  }, []);

  /**
   * 切换文件夹工作表选择（folder 模式）
   */
  const toggleFolderSheet = useCallback((fileId: string, sheetName: string) => {
    const key = folderSheetKey(fileId, sheetName);
    setFolderSheetKeys((current) =>
      current.includes(key)
        ? current.filter((item) => item !== key)
        : [...current, key]
    );
    setSelectionConfirmed(false);
  }, []);

  /**
   * 浏览并选择文件夹
   */
  const browseFolder = useCallback(async (options?: {
    onMessage?: (text: string) => void;
    onServerOnline?: () => void;
  }) => {
    try {
      const catalog = await selectFolder();
      if (!catalog) return null;

      setFolderCatalog(catalog);
      setFolderSheetKeys([]);
      setSelectionConfirmed(false);

      if (options?.onServerOnline) {
        options.onServerOnline();
      }

      return catalog;
    } catch (reason) {
      const errorMessage = reason instanceof Error ? reason.message : "读取文件夹失败";

      if (options?.onMessage) {
        options.onMessage(errorMessage);
      }

      throw new Error(errorMessage);
    }
  }, []);

  /**
   * 选择自动范围模式（当前活动工作表）
   */
  const chooseAutomaticScope = useCallback(() => {
    setSourceMode("workbook");
    setWorkbookScopeMode("auto");
    if (workbook) {
      setSelectedSheetNames([workbook.activeWorksheet]);
    }
    setSelectionConfirmed(true);
  }, [workbook]);

  /**
   * 选择手动范围模式（用户选择工作表）
   */
  const chooseManualScope = useCallback(() => {
    setSourceMode("workbook");
    setWorkbookScopeMode("manual");
    if (selectedSheetNames.length === 0 && workbook) {
      setSelectedSheetNames([workbook.activeWorksheet]);
    }
    setSelectionConfirmed(selectedSheetNames.length > 0 || workbook !== null);
  }, [workbook, selectedSheetNames]);

  /**
   * 选择文件夹模式
   */
  const chooseFolderScope = useCallback(() => {
    setSourceMode("folder");
    setSelectionConfirmed(folderSheetKeys.length > 0);
  }, [folderSheetKeys]);

  /**
   * 获取文件夹选择列表
   */
  const getFolderSelections = useCallback((): FolderSelection[] => {
    if (!folderCatalog) return [];

    return folderCatalog.files
      .map((file) => ({
        fileId: file.id,
        sheets: file.worksheets
          .filter((sheet) => folderSheetKeys.includes(folderSheetKey(file.id, sheet.name)))
          .map((sheet) => sheet.name)
      }))
      .filter((selection) => selection.sheets.length > 0);
  }, [folderCatalog, folderSheetKeys]);

  /**
   * 初始化：捕获工作簿 + 监听结构变化
   */
  useEffect(() => {
    if (typeof Office === "undefined") {
      setWorkbook(demoWorkbook);
      setSelectedSheetNames([demoWorkbook.activeWorksheet]);
      return;
    }

    let disposed = false;
    let dispose: (() => void) | undefined;

    Office.onReady(async () => {
      if (disposed) return;
      try {
        await scan();
      } catch {
        // scan 内部已处理错误
      }

      if (disposed) return;

      if (isRunningInExcel()) {
        const watcherDispose = await watchWorkbookStructureChanges(() => {
          setSelectionConfirmed(false);
          // 通知父组件：结构变化，需要清除撤销快照
          // 这部分逻辑由父组件通过监听 workbook 变化来处理
        });

        if (disposed) {
          watcherDispose?.();
          return;
        }

        dispose = watcherDispose;
      }
    });

    return () => {
      disposed = true;
      dispose?.();
    };
  }, []); // 只在挂载时运行

  return {
    // 状态
    workbook,
    selectedSheetNames,
    selectionConfirmed,
    sourceMode,
    workbookScopeMode,
    folderCatalog,
    folderSheetKeys,

    // 派生状态
    workbookDataPeriod,

    // 操作
    scan,
    toggleSheet,
    toggleFolderSheet,
    browseFolder,
    chooseAutomaticScope,
    chooseManualScope,
    chooseFolderScope,
    getFolderSelections,

    // 状态设置（给特殊场景使用）
    setWorkbook,
    setSelectedSheetNames,
    setSelectionConfirmed,
    setSourceMode,
    setWorkbookScopeMode,
    setFolderCatalog,
    setFolderSheetKeys,
  };
}
