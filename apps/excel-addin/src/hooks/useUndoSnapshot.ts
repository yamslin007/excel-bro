import { useCallback, useState } from "react";
import type { ExecutionUndoSnapshot } from "../contracts";
import { isRunningInExcel, undoExecution } from "../excel";

interface UseUndoSnapshotOptions {
  isBusy: boolean;
  onStatusChange: (status: "executing" | "idle") => void;
  onMessage: (text: string) => void;
  onAfterUndo?: () => Promise<void> | void;
}

/**
 * 撤销快照管理 Hook
 *
 * 职责：
 * - 保存最近一次成功执行产生的可撤销快照
 * - 在工作簿结构变化等场景下清除失效快照
 * - 执行最近一次撤销并通知上层刷新/提示
 */
export function useUndoSnapshot({
  isBusy,
  onStatusChange,
  onMessage,
  onAfterUndo
}: UseUndoSnapshotOptions) {
  const [lastUndoSnapshot, setLastUndoSnapshot] =
    useState<ExecutionUndoSnapshot | null>(null);

  const clearUndoSnapshot = useCallback(() => {
    setLastUndoSnapshot(null);
  }, []);

  const undoLastExecution = useCallback(async () => {
    if (!lastUndoSnapshot || isBusy || !isRunningInExcel()) return;

    onStatusChange("executing");
    try {
      await undoExecution(lastUndoSnapshot);
      setLastUndoSnapshot(null);
      onMessage("已撤销上一次执行中记录的单元格值、公式和常用格式。");
      await onAfterUndo?.();
    } catch (reason) {
      onMessage(
        reason instanceof Error ? reason.message : "撤销上一次执行失败"
      );
    } finally {
      onStatusChange("idle");
    }
  }, [
    lastUndoSnapshot,
    isBusy,
    onStatusChange,
    onMessage,
    onAfterUndo
  ]);

  return {
    lastUndoSnapshot,
    setLastUndoSnapshot,
    clearUndoSnapshot,
    undoLastExecution
  };
}
