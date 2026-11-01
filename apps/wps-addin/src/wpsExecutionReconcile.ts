import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import {
  EXECUTION_RECOVERY_ACTION_LABEL,
  EXECUTION_RECOVERY_STATE_LABEL,
  executionRecoveryActions,
  executionRecoveryDetail,
  type ExecutionRecoveryAction,
  type ExecutionRecoveryVerdict
} from "@excel-bro/core/executionRecoveryView";

/** Read-only recovery verdict. Resume requires independently checked local material;
 * an in-flight action must first be reconciled, never guessed or replayed.
 * Historical records without compatible source evidence remain replan-only.
 *
 * 形状由共享层 `ExecutionRecoveryVerdict` 定义（Excel 读同一份），这里不再自己声明一遍。
 */

/** 身份核对失败的几种具体原因。 */
export type WpsIdentityVerdict =
  | { kind: "match" }
  | { kind: "mismatch"; detail: string }
  | { kind: "unavailable"; detail: string };

/**
 * 核对当前工作簿身份与服务端记录是否一致。
 *
 * `currentIdentity` 是本次读到的 `wps-path:` 字符串（可能为 `null` = 未保存 /
 * 无法稳定识别）；`recordedIdentity` 是服务端记录里的身份。判据是**完整字符串
 * 相等** —— 不按工作簿名、不按前缀、不按"看着像"。
 */
export function wpsIdentityVerdict(input: {
  currentIdentity: string | null;
  recordedIdentity: string;
}): WpsIdentityVerdict {
  if (!input.currentIdentity) {
    return { kind: "unavailable", detail: "当前工作簿未保存或无法稳定识别，不能核对这次写入。" };
  }
  if (input.currentIdentity !== input.recordedIdentity) {
    return { kind: "mismatch", detail: "当前工作簿与这次写入记录的不是同一个文件。" };
  }
  return { kind: "match" };
}

/** Server ownership/cursor and current target inspection remain authoritative.
 * resumeAvailable only adds a suffix-write option after those checks pass.
 */
export function wpsReconcileVerdict(input: {
  state: WpsExecutionState | null;
  resumeAvailable?: boolean;
  resumeUnavailableReason?: string | null;
  identityMatch: boolean;
  localInspection: "completed" | "not_started" | "conflict" | "unknown" | null;
}): WpsReconcileVerdict {
  if (!input.state) {
    return { kind: "insufficient", detail: "服务端没有这次执行的记录，无法判断写入现场。" };
  }
  if (!input.identityMatch) {
    return {
      kind: "insufficient",
      detail: "当前工作簿与这次写入记录的身份不一致；请打开原工作簿后再核对，不猜是不是同一个文件。"
    };
  }
  /*
   * 终态优先（2026-10-31）：撤销/放弃之后账本 `status` 与游标都不变（`record_office_disposition`
   * 只写两列时刻），宿主缓存里 `record.state="restored"` 又随重开消失 —— 不在这里先判，
   * 跨重启的续跑判据会把一次已经回退掉的执行重新开放续跑/撤销。
   *
   * 与 Excel 侧 `officeExecutionTerminalReason` 同一条纪律：`undone` 是终态（不给任何按钮），
   * `abandoned` 是"放弃变更"（回退过了，需要人工核对，给 blocked 不给写入）。
   */
  if (input.state.disposition?.undoneAt) {
    return {
      kind: "undone",
      // 账本已留处置痕（常为另一台机器撤的）：本机没回退，现场未知，不用默认"已回到执行前"。
      detail: "这次执行已经撤销过（账本已留处置痕），不会再执行或补记任何动作。"
    };
  }
  if (input.state.disposition?.abandonedAt) {
    return {
      kind: "blocked",
      detail: "这次执行已按「放弃变更」处置过，原表已按恢复点还原；不能再继续剩余动作或补记完成。",
      /*
       * `abandoned: true` = 用户已经做过决定，结论不会再变（终态）。仍是 `blocked`
       * （不给写入/补记），但与"现场冲突 / 执行权被占 / 来源变化"这类**还要人核对**
       * 的阻断分开 —— 只有它可以退出固定区域，换成消息里一条可展开的终态摘要。
       */
      abandoned: true
    };
  }
  if (input.state.status === "completed") {
    return { kind: "completed", recorded: true };
  }
  if (input.state.ownerActive !== false || !input.state.canResume) {
    return { kind: "blocked", detail: "原执行者仍持有执行权，或服务端未允许恢复；不能接管或重复规划写入。" };
  }
  if (!Number.isInteger(input.state.totalActions) || input.state.totalActions < 1 ||
      !Number.isInteger(input.state.nextActionIndex) || input.state.nextActionIndex < 0 ||
      input.state.nextActionIndex > input.state.totalActions) {
    return { kind: "insufficient", detail: "动作总数或进度游标无效，无法安全补记。" };
  }
  if (input.localInspection === "conflict") {
    return { kind: "blocked", detail: "现场与登记状态不一致，疑似人工修改或部分写入，需人工核对。" };
  }
  if (input.localInspection == null || input.localInspection === "unknown") {
    return { kind: "insufficient", detail: "无法读取或证明当前目标现场，不能补记或重新规划写入。" };
  }
  // 字段缺失与显式 null 同义：都表示"服务端没给出在途动作位置"。
  const inFlight = input.state.inFlightActionIndex ?? null;
  if (inFlight != null && (!Number.isInteger(inFlight) || inFlight !== input.state.nextActionIndex || inFlight >= input.state.totalActions)) {
    return { kind: "insufficient", detail: "在途动作与进度游标不一致，无法安全补记。" };
  }
  if (input.state.requiresReconciliation || inFlight != null) {
    if (inFlight === null) {
      // 服务端说"需要核对"却没给在途动作号：状态自相矛盾，不许自己挑一个去补记。
      return {
        kind: "insufficient",
        detail: "服务端标记需要核对，但没有给出在途动作位置，无法安全补记。"
      };
    }
    if (input.localInspection === "completed") {
      return {
        kind: "reconcile",
        actionIndex: inFlight,
        outcome: "applied",
        remaining: Math.max(0, input.state.totalActions - inFlight - 1)
      };
    }
    if (input.localInspection === "not_started") {
      return {
        kind: "reconcile",
        actionIndex: inFlight,
        outcome: "not_applied",
        remaining: Math.max(0, input.state.totalActions - inFlight - 1)
      };
    }
    return {
      kind: "insufficient",
      detail: "有在途动作，但本机恢复点读不到现场，无法自动核对；请人工确认原表后再决定。"
    };
  }
  const remaining = Math.max(0, input.state.totalActions - input.state.nextActionIndex);
  if (remaining === 0) {
    // 全部动作都已被服务端记为"回读一致后完成"，缺的只是最终回执。
    // 补记不产生任何宿主写入；若本机现场与"已完成"矛盾，仍然拦住。
    if (input.localInspection === "not_started") {
      return {
        kind: "blocked",
        detail: "服务端记录显示全部动作已完成，但本机现场仍是写前状态，两者矛盾，需人工核对。"
      };
    }
    return { kind: "completed" };
  }
  return input.resumeAvailable ? {kind: "resume", remaining} : {kind: "replan", remaining, ...(input.resumeUnavailableReason ? {reason: input.resumeUnavailableReason} : {})};
}

/**
 * 结论与措辞的**唯一出处**在共享层 `@excel-bro/core/executionRecoveryView`：
 * Excel 端 2026-10-29 起读同一份。这里只做两件事 —— 用宿主自己的名字转发，
 * 以及把结论收窄成"一定读过现场"（WPS 的判定不会停在 `pending`）。
 *
 * 下面这些别名不是"行为相似的副本"：改文案只改共享层一处，两端同时生效。
 */
export type WpsReconcileVerdict = ExecutionRecoveryVerdict<"applied" | "not_applied">;

/** 核对结论的用户可见文案（标题级）。 */
export const WPS_RECONCILE_LABEL = EXECUTION_RECOVERY_STATE_LABEL;

/** 结论的正文说明。冲突与证据不足直接复用判定给出的 `detail`，不另写一套措辞。 */
export const WPS_RECOVERY_DETAIL = executionRecoveryDetail;

/** 各结论允许出现的操作。返回空数组 = 不给任何写入/补记按钮。 */
export type WpsReconcileAction = ExecutionRecoveryAction;

export const WPS_RECOVERY_ACTION_LABEL = EXECUTION_RECOVERY_ACTION_LABEL;

export const wpsReconcileActions = executionRecoveryActions;
