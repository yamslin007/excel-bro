import { useCallback, useState } from "react";
import capabilities from "../../../../config/capabilities.json";
import type { ExecutionUndoSnapshot } from "../contracts";
import { isRunningInExcel, undoExecution } from "../excel";

const MAX_UNDO_ENTRIES = capabilities.undo.maxEntries;
const MAX_UNDO_TOTAL_CELLS = capabilities.undo.maxTotalCells;

interface UseUndoSnapshotOptions {
  isBusy: boolean;
  onStatusChange: (status: "executing" | "idle") => void;
  onMessage: (text: string) => void;
  onAfterUndo?: () => Promise<void> | void;
  onRestored?: (snapshot: ExecutionUndoSnapshot) => void;
  beforeOperation?: () => Promise<boolean>;
}

function snapshotCellCount(snapshot: ExecutionUndoSnapshot): number {
  let total = 0;
  for (const range of snapshot.ranges) {
    total += range.formulas.length * (range.formulas[0]?.length ?? 0);
  }
  for (const sheet of snapshot.deletedWorksheets) {
    total += sheet.formulas.length * (sheet.formulas[0]?.length ?? 0);
  }
  return total;
}

/**
 * 把撤销栈裁剪到容量之内。栈顶（末尾）是最近一次执行，超限时从栈底丢弃 ——
 * 最久远的执行最不可能被撤销，且丢弃它不影响较新条目的可撤销性。
 */
function trimStack(stack: ExecutionUndoSnapshot[]): ExecutionUndoSnapshot[] {
  let trimmed = stack.slice(-MAX_UNDO_ENTRIES);
  let total = trimmed.reduce((sum, entry) => sum + snapshotCellCount(entry), 0);
  while (trimmed.length > 1 && total > MAX_UNDO_TOTAL_CELLS) {
    total -= snapshotCellCount(trimmed[0]);
    trimmed = trimmed.slice(1);
  }
  return trimmed;
}

/**
 * 撤销栈管理 Hook
 *
 * 职责：
 * - 维护最近若干次成功执行的可撤销快照（栈顶为最近一次）
 * - 在工作簿结构变化等场景下清除全部失效快照
 * - 按 LIFO 逐级撤销：撤销某一条会连带撤销它之后的所有执行
 *
 * 为什么是 LIFO 而不能任选一条：多次执行的区域可能重叠，跳过中间某次执行
 * 的话，被跳过的那次写入无法在已变化的上下文里重建，恢复结果不可预测。
 */
export function useUndoSnapshot({
  isBusy,
  onStatusChange,
  onMessage,
  onAfterUndo,
  onRestored,
  beforeOperation
}: UseUndoSnapshotOptions) {
  const [undoStack, setUndoStack] = useState<ExecutionUndoSnapshot[]>([]);
  const lastUndoSnapshot =
    undoStack.length > 0 ? undoStack[undoStack.length - 1] : null;

  const pushUndoSnapshot = useCallback(
    (snapshot: ExecutionUndoSnapshot | null) => {
      if (!snapshot) return;
      setUndoStack((stack) => trimStack([...stack, snapshot]));
    },
    []
  );

  const clearUndoSnapshot = useCallback(() => {
    setUndoStack([]);
  }, []);

  /** 撤销到指定 executionId；旧消息没有 executionId 时兼容 planId。 */
  const undoExecutionTo = useCallback(
    async (executionIdOrPlanId?: string) => {
      if (isBusy || !isRunningInExcel()) return false;
      const index =
        executionIdOrPlanId === undefined
          ? undoStack.length - 1
          : undoStack.findIndex(
              (entry) =>
                entry.executionId === executionIdOrPlanId ||
                entry.planId === executionIdOrPlanId
            );
      if (index < 0) return false;
      if (beforeOperation && !(await beforeOperation())) return false;
      // 逆序撤销：先撤最近的，再往前退，保证每一步的上下文都是它执行完的状态。
      const targets = undoStack.slice(index).reverse();

      onStatusChange("executing");
      let undone = 0;
      let restored = false;
      try {
        for (const snapshot of targets) {
          await undoExecution(snapshot);
          undone += 1;
          onRestored?.(snapshot);
        }
        restored = true;
        onMessage(
          undone > 1
            ? `已撤销最近 ${undone} 次执行中记录的单元格值、公式和格式。`
            : "已撤销上一次执行中记录的单元格值、公式和格式。"
        );
        await onAfterUndo?.();
      } catch (reason) {
        onMessage(
          reason instanceof Error ? reason.message : "撤销执行失败"
        );
      } finally {
        // 已成功撤销的条目移出栈；失败的那条保留，供用户重试。
        if (undone > 0) {
          setUndoStack((stack) => stack.slice(0, stack.length - undone));
        }
        onStatusChange("idle");
      }
      return restored;
    },
    [undoStack, isBusy, onStatusChange, onMessage, onAfterUndo, onRestored, beforeOperation]
  );

  const undoLastExecution = useCallback(
    () => undoExecutionTo(),
    [undoExecutionTo]
  );

  return {
    undoStack,
    lastUndoSnapshot,
    pushUndoSnapshot,
    clearUndoSnapshot,
    undoLastExecution,
    undoExecutionTo
  };
}
