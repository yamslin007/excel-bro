import type { WpsExecutionState } from "../../excel-addin/src/contracts";

/** Read-only recovery verdict. Resume requires independently checked local material;
 * an in-flight action must first be reconciled, never guessed or replayed.
 * Historical records without compatible source evidence remain replan-only.
 */
export type WpsReconcileVerdict =
  | { kind: "completed"; recorded?: true }
  | { kind: "reconcile"; actionIndex: number; outcome: "applied" | "not_applied" }
  | { kind: "replan"; remaining: number; reason?: string }
  | { kind: "resume"; remaining: number }
  | { kind: "blocked"; detail: string }
  | { kind: "insufficient"; detail: string };

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
      return { kind: "reconcile", actionIndex: inFlight, outcome: "applied" };
    }
    if (input.localInspection === "not_started") {
      return { kind: "reconcile", actionIndex: inFlight, outcome: "not_applied" };
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

/** 核对结论的用户可见文案（标题级）。 */
export const WPS_RECONCILE_LABEL: Record<WpsReconcileVerdict["kind"], string> = {
  completed: "已完成，仅缺回执",
  reconcile: "待核对在途动作",
  replan: "需重新规划",
  resume: "可继续剩余动作",
  blocked: "需要人工核对",
  insufficient: "证据不足"
};

/** 结论的正文说明。冲突与证据不足直接复用判定给出的 `detail`，不另写一套措辞。 */
export function WPS_RECOVERY_DETAIL(verdict: WpsReconcileVerdict): string {
  switch (verdict.kind) {
    case "completed":
      return verdict.recorded ? "账本已保存历史完成结果，无需重复写入；不代表当前单元格仍未被修改。" : "动作与当前现场已核对，缺的只是最终回执；补记只改账本状态，不会重写原表。";
    case "reconcile":
      return verdict.outcome === "applied"
        ? `第 ${verdict.actionIndex + 1} 个动作的现场与写后一致，可补记为已写入，不重复写。`
        : `第 ${verdict.actionIndex + 1} 个动作的现场仍是写前状态，可补记为未写入。`;
    case "resume":
      return `原计划与当前授权、来源和目标已核对，可继续剩余 ${verdict.remaining} 个动作；已完成动作不会重写，完成后审阅结果。`;
    case "replan":
      return (
        `还有 ${verdict.remaining} 个动作没写。本机材料缺失、过期或不符合精确续跑条件。` +
        (verdict.reason ? `具体原因：${verdict.reason} ` : "") +
        "可在当前授权范围内重新读取并规划；通过校验后先执行，再审阅。"
      );
    default:
      return verdict.detail;
  }
}

/** 各结论允许出现的操作。返回空数组 = 不给任何写入/补记按钮。 */
export type WpsReconcileAction = "backfill" | "reconcile" | "replan" | "resume" | "dismiss";

export const WPS_RECOVERY_ACTION_LABEL: Record<WpsReconcileAction, string> = {
  resume: "继续剩余动作",
  backfill: "补记完成（不重写原表）",
  reconcile: "核对在途动作",
  replan: "重新读取并规划",
  dismiss: "清除这条恢复提示"
};

export function wpsReconcileActions(verdict: WpsReconcileVerdict): WpsReconcileAction[] {
  switch (verdict.kind) {
    case "completed":
      return verdict.recorded ? ["dismiss"] : ["backfill", "dismiss"];
    case "reconcile":
      return ["reconcile", "dismiss"];
    case "resume":
      return ["resume", "dismiss"];
    case "replan":
      return ["replan", "dismiss"];
    default:
      // `blocked` / `insufficient`：证据不够或现场冲突，一律不给写入或补记按钮，
      // 只留"已人工核对，清除提示"。
      return ["dismiss"];
  }
}
