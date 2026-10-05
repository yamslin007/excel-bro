import { SESSION_RECOVERY_LABEL, VERIFICATION_PASSED_LABEL } from "@excel-bro/core";
import type {
  TaskBlocker,
  TaskPhase,
  TaskResumability,
  TaskWriteState
} from "@excel-bro/core";

/**
 * WPS 执行结果卡的**归属与状态判定**。
 *
 * 为什么单独成模块：撤销按钮的归属曾经按"消息位置"推断 —— 一张全局撤销卡
 * 无条件挂在消息流末尾，于是**上一轮**写入留下的恢复点看起来像**本轮**唯一
 * 能做的下一步。修法不能只是换个位置，判据本身必须是显式的：
 *
 * - 绑定键只有 `executionId`。不按消息位置、不按计划标题、也不按"最近一次回复"。
 * - 恢复点必须**当前会话里真的还活着**，且**允许撤销**，按钮才可点。
 * - 卡片必须如实区分四种状态，谁都不许冒充谁。
 *
 * 这里只做纯判定与文案；渲染由 `WpsAgentPanel.tsx` 负责，测试从 DOM 上验。
 */

/** 这次执行到底发生了什么。三态刻意分开，不许合并成"成功/失败"。 */
export type WpsExecutionOutcome = "readback_passed" | "unverified" | "partial";

/** 恢复点在**此刻**的可操作性。 */
export type WpsRecoveryCardState = "undoable" | "conflict" | "undone" | "unavailable";

/** 当前会话里活跃的那个恢复点（只带判定需要的最小信息）。 */
export type WpsLiveRecoveryPoint = {
  executionId: string;
  state: string;
};

/**
 * 一张执行结果卡当前该显示成什么状态。
 *
 * - `undone`：已撤销。**终端态**，优先于一切 —— 撤销之后即便宿主里还留着
 *   记录，也不许重新变回可点。
 * - `undoable`：活跃恢复点就是这一次，且它当前允许撤销。
 * - `conflict`：活跃恢复点仍是这一次，但宿主报告执行后被改动过，不能自动覆盖。
 * - `unavailable`：当前会话里没有匹配 `executionId` 的可撤销恢复点 ——
 *   可能被后续执行取代了、可能所在加载项会话已经结束（面板重开）、
 *   也可能这条记录本来就没留下恢复点。
 */
export function wpsRecoveryCardState(input: {
  executionId: string;
  undone: boolean;
  live: WpsLiveRecoveryPoint | null;
}): WpsRecoveryCardState {
  if (input.undone) return "undone";
  const live = input.live;
  if (!live || live.executionId !== input.executionId) return "unavailable";
  if (live.state === "undo_conflict") return "conflict";
  if (live.state === "verified" || live.state === "partial_write_recoverable") return "undoable";
  return "unavailable";
}

/**
 * 恢复点状态的用户可见文案。四种状态互不重复、互不冒充。
 *
 * `undoable` 刻意**不**写成裸的"可撤销" —— 那会让人以为恢复点是持久的。
 * 它直接引用共享的 `SESSION_RECOVERY_LABEL`，与两端的结果卡口径同源，
 * 因此这里也**不得**出现任何硬编码的会话内措辞。
 */
export const WPS_RECOVERY_CARD_LABEL: Record<WpsRecoveryCardState, string> = {
  undoable: SESSION_RECOVERY_LABEL,
  conflict: "撤销冲突",
  undone: "已撤销",
  unavailable: "恢复点已不可用"
};

/**
 * 宿主里这个执行状态是否**真的留下了恢复点**。
 *
 * `undo_conflict` 也算：恢复资料被保留下来了，只是不能自动覆盖。
 * `restored` / `ready` / `applying` / `undoing` 都不算 —— 前者说明区域已回到
 * 执行前、后者是中间态。拿不到记录（`null`/`undefined`）同样不算。
 */
export function wpsExecutionHasRecoveryPoint(state: string | null | undefined): boolean {
  return state === "verified" || state === "partial_write_recoverable" || state === "undo_conflict";
}

/**
 * 三种执行结果的标题。
 *
 * 只有**完全回读通过**才允许出现共享的 `VERIFICATION_PASSED_LABEL`；
 * 另外两种各有自己的说法，不得借用"验证通过"这四个字。
 */
export const WPS_EXECUTION_OUTCOME_LABEL: Record<WpsExecutionOutcome, string> = {
  readback_passed: `写入${VERIFICATION_PASSED_LABEL}`,
  unverified: "已写入，计算值未获证实",
  partial: "写入未完成，恢复资料已保留"
};

/**
 * 结果卡第二行的后缀（前面拼 `${actionCount} 个动作 · ${cellCount} `）。
 * 只有完全回读通过才写得出"已回读核对"。
 */
export const WPS_EXECUTION_OUTCOME_DETAIL_SUFFIX: Record<WpsExecutionOutcome, string> = {
  readback_passed: "个单元格已回读核对",
  unverified: "个单元格已写入，未能核对计算值",
  partial: "个单元格写入未完成，恢复资料已保留"
};

/** 结果卡的色调类名。**不得**把未证实或部分写入伪装成 `verified`。 */
export function wpsExecutionCardTone(input: {
  outcome: WpsExecutionOutcome;
  undone: boolean;
}): "verified" | "unverified" | "partial" | "undone" {
  if (input.undone) return "undone";
  if (input.outcome === "unverified") return "unverified";
  if (input.outcome === "partial") return "partial";
  return "verified";
}

/** 旧历史里的结果卡没有 `outcome` 字段；那时只有完全回读通过才会画卡。 */
export function normalizeWpsExecutionOutcome(value: unknown): WpsExecutionOutcome {
  return value === "unverified" || value === "partial" ? value : "readback_passed";
}

/**
 * 一次写入尝试到底留下了什么**证据**。
 *
 * 为什么必须四态而不是"成功/失败"：用户看到的"原表写入状态"是据此写的，
 * 而"找不到记录""读取失败""回调丢失"三者**都不能**推出"没写过"。把任何一个
 * 折成 `not_started`，就是对着可能已被改过的原表说"原表未写入"。
 */
export type WpsWriteEvidence =
  | { kind: "completed"; executionId: string }
  | { kind: "not_started"; executionId: string }
  | { kind: "needs_review"; executionId: string }
  | { kind: "unknown" };

/**
 * 由"执行回调给的 executionId + 宿主里读到的原始记录 + 是否读失败"判定证据。
 *
 * 判据只认**正面证据**：
 * - `restored` 是执行器自己的结论（没有写进任何一块），才算"未开始写入"；
 * - `verified` 才是"写完并回读过"；
 * - `partial_write_recoverable` / `undo_conflict` 是"写了但没走完 / 现场有冲突"；
 * - 其余（读失败、回调丢失、找不到记录、`ready`/`applying`/`undoing` 中间态）
 *   一律 `unknown` —— **不猜**。
 */
export function classifyWpsWriteEvidence(input: {
  /** 执行回调给的执行身份；`null` 表示回调没拿到（回执丢失）。 */
  executionId: string | null;
  /** 按 executionId 读到的**原始**记录（未过滤可撤销性）。 */
  record: { executionId: string; state: string } | null;
  /** 读取宿主存储本身失败了（与"读到但没有"是两件事）。 */
  readFailed: boolean;
}): WpsWriteEvidence {
  if (input.readFailed) return { kind: "unknown" };
  // 回执丢失：无法把任何一条记录归属到这次执行，因此不能下任何结论。
  if (!input.executionId) return { kind: "unknown" };
  const record = input.record;
  // 找不到恢复点：可能是存储被清、键损坏，也可能是从未登记 —— 无法区分，只能说未知。
  if (!record || record.executionId !== input.executionId) return { kind: "unknown" };
  if (record.state === "restored") return { kind: "not_started", executionId: record.executionId };
  if (record.state === "verified") return { kind: "completed", executionId: record.executionId };
  if (record.state === "partial_write_recoverable" || record.state === "undo_conflict") {
    return { kind: "needs_review", executionId: record.executionId };
  }
  return { kind: "unknown" };
}

/**
 * 把证据翻译成任务恢复记录的四个字段。面板只用这一处映射，
 * 避免"某条分支忘了改 writeState"这种人眼容易漏的错。
 */
export function taskResumePatchForEvidence(evidence: WpsWriteEvidence): {
  phase: TaskPhase;
  blocker: TaskBlocker;
  resumability: TaskResumability;
  writeState: TaskWriteState;
} {
  switch (evidence.kind) {
    case "completed":
      return {
        phase: "executing",
        blocker: "none",
        resumability: "unavailable",
        writeState: "confirmed_result"
      };
    case "not_started":
      return {
        phase: "preview_ready",
        blocker: "none",
        resumability: "replan",
        writeState: "not_written"
      };
    case "needs_review":
      return {
        phase: "executing",
        blocker: "execution_unknown",
        resumability: "reconcile",
        writeState: "needs_review"
      };
    default:
      return {
        phase: "executing",
        blocker: "execution_unknown",
        resumability: "reconcile",
        writeState: "unknown"
      };
  }
}
