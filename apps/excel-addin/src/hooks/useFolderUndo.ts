import { useCallback, useState } from "react";
import capabilities from "../../../../config/capabilities.json";
import { undoFolderExecution } from "../api";
import type {
  FolderUndoReference,
  FolderUndoResult
} from "../contracts";

const MAX_UNDO_ENTRIES = capabilities.undo.maxEntries;

export interface FolderUndoEntry extends FolderUndoReference {
  planId: string;
}

interface UseFolderUndoOptions {
  isBusy: boolean;
  onStatusChange: (status: "executing" | "idle") => void;
  onMessage: (text: string) => void;
  onAfterUndo?: (
    entry: FolderUndoEntry,
    result: FolderUndoResult
  ) => Promise<void> | void;
}

export function useFolderUndo({
  isBusy,
  onStatusChange,
  onMessage,
  onAfterUndo
}: UseFolderUndoOptions) {
  const [folderUndoStack, setFolderUndoStack] = useState<FolderUndoEntry[]>([]);

  const pushFolderUndo = useCallback((entry: FolderUndoEntry | null) => {
    if (!entry) return;
    setFolderUndoStack((stack) => {
      const sameSession =
        stack.length === 0 ||
        stack[stack.length - 1].sessionId === entry.sessionId;
      return (sameSession ? [...stack, entry] : [entry]).slice(
        -MAX_UNDO_ENTRIES
      );
    });
  }, []);

  const clearFolderUndo = useCallback(() => {
    setFolderUndoStack([]);
  }, []);

  const reconcileFolderUndoIds = useCallback((remainingUndoIds: string[]) => {
    const remaining = new Set(remainingUndoIds);
    setFolderUndoStack((stack) =>
      stack.filter((entry) => remaining.has(entry.undoId))
    );
  }, []);

  const undoFolderExecutionTo = useCallback(
    async (undoIdOrPlanId?: string, recovery?: FolderUndoReference) => {
      if (isBusy) return;
      const index =
        undoIdOrPlanId === undefined
          ? folderUndoStack.length - 1
          : folderUndoStack.findIndex(
              (entry) =>
                entry.undoId === undoIdOrPlanId ||
                entry.planId === undoIdOrPlanId
            );
      if (index < 0 && !recovery) return;

      const entry: FolderUndoEntry =
        index >= 0
          ? folderUndoStack[index]
          : {
              ...recovery!,
              planId: undoIdOrPlanId ?? recovery!.undoId
            };
      onStatusChange("executing");
      try {
        const result = await undoFolderExecution(
          entry.sessionId,
          entry.undoId
        );
        setFolderUndoStack((stack) =>
          index >= 0
            ? stack.slice(0, index)
            : stack.slice(0, Math.min(stack.length, result.remainingUndoCount))
        );
        onMessage(
          result.executionsUndone > 1
            ? `已撤销最近 ${result.executionsUndone} 次文件夹执行，并恢复到执行前状态：${result.filesRestored.join("、")}。`
            : `已撤销上一次文件夹执行，并恢复到执行前状态：${result.filesRestored.join("、")}。`
        );
        try {
          await onAfterUndo?.(entry, result);
        } catch (reason) {
          onMessage(
            `撤销已经完成，但文件夹列表刷新失败：${
              reason instanceof Error ? reason.message : String(reason)
            }`
          );
        }
      } catch (reason) {
        onMessage(
          reason instanceof Error ? reason.message : "撤销文件夹执行失败"
        );
      } finally {
        onStatusChange("idle");
      }
    },
    [
      folderUndoStack,
      isBusy,
      onStatusChange,
      onMessage,
      onAfterUndo
    ]
  );

  return {
    folderUndoStack,
    pushFolderUndo,
    clearFolderUndo,
    reconcileFolderUndoIds,
    undoFolderExecutionTo
  };
}
