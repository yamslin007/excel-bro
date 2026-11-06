import { useCallback, useRef, useState } from "react";
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
  onRestored?: (snapshot: ExecutionUndoSnapshot) => Promise<void> | void;
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
  /*
   * 栈的**同步**副本。
   *
   * 存在理由（2026-10-31）：恢复卡的「撤销本次执行」是"先挂快照、紧接着就撤"——
   * `ensureUndoSnapshot(...)` 之后**同一个 tick** 调 `undoExecutionTo(...)`。React 的
   * 状态更新要等下一次渲染，闭包里那份 `undoStack` 在那一刻还是旧的，于是刚挂进去的
   * 那一条找不到，撤销静默不执行（按钮点了没反应，连失败文案都没有）。
   * 所有入栈/出栈都先写 ref 再写 state，同一次点击里就能看到自己刚放进去的东西。
   */
  const stackRef = useRef<ExecutionUndoSnapshot[]>([]);
  const commitStack = useCallback(
    (update: (stack: ExecutionUndoSnapshot[]) => ExecutionUndoSnapshot[]) => {
      const next = update(stackRef.current);
      stackRef.current = next;
      setUndoStack(next);
      return next;
    },
    []
  );
  const lastUndoSnapshot =
    undoStack.length > 0 ? undoStack[undoStack.length - 1] : null;

  const pushUndoSnapshot = useCallback(
    (snapshot: ExecutionUndoSnapshot | null) => {
      if (!snapshot) return;
      /*
       * 同一执行**只保留最新的一份**（新的替换旧的）。
       *
       * 存在理由（2026-10-31）：一次执行可能先后产生两份快照 —— 断点恢复时从账本材料
       * 重建的"已完成前缀"那一份，和续跑完成后整份执行那一份。两份都留在栈里，LIFO
       * 会把同一批区域恢复两遍：先按整份还原，再按前缀那份核对"写后指纹"必然对不上，
       * 于是**放弃变更/撤销被误判成冲突**（真机上表现为"恢复未成功，当前结果仍保留"）。
       * 覆盖语义才是对的：同一执行的撤销点只有一个，越晚的越完整。
       */
      commitStack((stack) =>
        trimStack([
          ...stack.filter((entry) => !snapshot.executionId || entry.executionId !== snapshot.executionId),
          snapshot
        ])
      );
    },
    [commitStack]
  );

  const clearUndoSnapshot = useCallback(() => {
    commitStack(() => []);
  }, [commitStack]);

  /**
   * 把一份**从账本材料重建**的快照放进栈里，同一执行只放一次。
   *
   * 与 `pushUndoSnapshot` 的差别只有一条：这里按 `executionId` 去重。重开窗格后，
   * 账本状态可能被读不止一次（发现、核对、续跑各一次），每次都盲 push 就会在栈里
   * 攒出同一执行的多个副本 —— 撤销时会把同一块区域恢复两遍，栈长也被污染。
   * 快照内容一律按账本原样，这里不改写任何区域。
   */
  const ensureUndoSnapshot = useCallback(
    (snapshot: ExecutionUndoSnapshot | null | undefined) => {
      if (!snapshot) return;
      commitStack((stack) => {
        const id = snapshot.executionId;
        const exists = id
          ? stack.some((entry) => entry.executionId === id)
          : stack.some((entry) => entry.planId === snapshot.planId);
        return exists ? stack : trimStack([...stack, snapshot]);
      });
    },
    [commitStack]
  );

  /** 撤销到指定 executionId；旧消息没有 executionId 时兼容 planId。 */
  const undoExecutionTo = useCallback(
    async (executionIdOrPlanId?: string, onlyThisExecution = false, silentSuccess = false) => {
      if (isBusy || !isRunningInExcel()) return false;
      // 读 ref 而不是闭包里的 state：同一 tick 里刚 `ensureUndoSnapshot` 的那一条也要算数。
      const stack = stackRef.current;
      const index =
        executionIdOrPlanId === undefined
          ? stack.length - 1
          : stack.findIndex(
              (entry) =>
                entry.executionId === executionIdOrPlanId ||
                entry.planId === executionIdOrPlanId
            );
      if (index < 0) return false;
      if (onlyThisExecution && (index !== stack.length - 1 || stack[index].executionId !== executionIdOrPlanId)) {
        throw new Error('本次执行不是当前恢复栈顶，未撤销其他任务。');
      }
      if (beforeOperation && !(await beforeOperation())) return false;
      if (onlyThisExecution && stackRef.current !== stack) {
        throw new Error('核对期间恢复栈已变化，未撤销任何任务，请重新核对。');
      }
      // 逆序撤销：先撤最近的，再往前退，保证每一步的上下文都是它执行完的状态。
      const targets = stack.slice(index).reverse();

      onStatusChange("executing");
      let undone = 0;
      let restored = false;
      try {
        for (const snapshot of targets) {
          await undoExecution(snapshot);
          undone += 1;
          await onRestored?.(snapshot);
        }
        restored = true;
        if (!silentSuccess) onMessage(
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
          commitStack((current) => current.slice(0, current.length - undone));
        }
        onStatusChange("idle");
      }
      return restored;
    },
    [isBusy, onStatusChange, onMessage, onAfterUndo, onRestored, beforeOperation, commitStack]
  );

  /** Discard one reviewed execution. A remounted pane may have no in-memory
   * stack: attach only its bound recovery receipt, never step over another task. */
  const undoReviewedExecution = useCallback(
    async (executionId: string, recoverySnapshot?: ExecutionUndoSnapshot | null) => {
      if (isBusy || !isRunningInExcel()) return false;
      if (!executionId || (recoverySnapshot && recoverySnapshot.executionId !== executionId)) {
        throw new Error('本次审阅与恢复材料的执行身份不一致，未撤销任何任务。');
      }
      const top = stackRef.current.at(-1);
      if (top && top.executionId !== executionId) {
        throw new Error('本次执行不是当前恢复栈顶，未撤销其他任务。');
      }
      if (!top) {
        if (!recoverySnapshot) throw new Error('本次执行缺少恢复材料，未放弃变更。');
        ensureUndoSnapshot(recoverySnapshot);
      }
      return undoExecutionTo(executionId, true, true);
    },
    [isBusy, ensureUndoSnapshot, undoExecutionTo]
  );

  const undoLastExecution = useCallback(
    () => undoExecutionTo(),
    [undoExecutionTo]
  );

  return {
    undoStack,
    lastUndoSnapshot,
    pushUndoSnapshot,
    ensureUndoSnapshot,
    clearUndoSnapshot,
    undoLastExecution,
    undoReviewedExecution,
    undoExecutionTo
  };
}
