import type { WpsExecutionState } from "../../excel-addin/src/contracts";

/**
 * WPS 写入耐久恢复的**重载核对判定**（CAP-03）。
 *
 * 权威现场是服务端账本（`WpsExecutionState`）。面板重载后：读本地发现索引 →
 * 查服务端真实状态 → 身份核对（`wps-path:` 完整字符串一致才继续）→ 依据
 * 服务端状态、剩余动作与本机现场把这次写入归成五类。判定只认正面证据，
 * 拿不到证据就挂起，不猜"写没写"。
 *
 * 纯函数，不碰宿主、不碰网络；渲染与动作由 `WpsAgentPanel.tsx` 负责。
 */

/**
 * 重载后这次写入的核对结论。五态互不冒充。
 *
 * - `completed`：全部动作已在服务端记为完成（缺的只是最终回执）→ **只补记，不重写**。
 * - `reconcile`：有在途动作，且本机现场能给出正面结论 → 可提交核对回执。
 * - `replan`：还有动作没写。确认快照只活在会话内内存，跨关闭拿不回来，
 *   **无法精确续跑剩余动作** → 只能重新读取并生成预览、重新授权确认；
 *   这里绝不能冒充"断点续跑"。
 * - `blocked`：现场与写前/写后都不一致，或状态矛盾 → 人工核对，不给写入按钮。
 * - `insufficient`：查不到状态、身份不可用、本机现场读不到 → 挂起，不给假按钮。
 */
export type WpsReconcileVerdict =
  | { kind: "completed"; recorded?: true }
  | { kind: "reconcile"; actionIndex: number; outcome: "applied" | "not_applied" }
  | { kind: "replan"; remaining: number }
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

/**
 * 依据服务端状态 + 本机现场 + 身份，判断重载后能安全做什么。
 *
 * 每一档都要**正面证据**，逐条收紧：
 *
 * 1. 服务端没有这次执行 → `insufficient`（查询不创建执行，也不猜）；
 * 2. 身份不match → `insufficient`（同名异路径、未保存一律不放行）；
 * 3. 服务端已 `completed` → `completed`（安全重放，不重写）；
 * 4. 有在途动作：本机现场给出正面结论才 `reconcile`；冲突 → `blocked`；
 *    读不到 → `insufficient`（不许拿"读不到"当成"没写过"）；
 * 5. `canResume === false` → `blocked`（服务端明确说这次不能续）；
 * 6. 剩余动作 = 0 → `completed`（只补记最终回执，不产生任何宿主写入）；
 *    但若本机现场与其矛盾（仍是写前 / 与写后冲突）→ `blocked`；
 * 7. 剩余动作 > 0 → `replan`。**这里不给"继续"**：目标值矩阵与确认快照只活在
 *    会话内内存，重载后拼不回用户当初确认的那份计划，精确续跑未实现。
 *
 * `localInspection` 来自当前授权内登记目标区域的真实只读回读
 * （`inspectWpsRecoveryState`），不依赖 PluginStorage 仍保存完整矩阵。
 * 旧描述没有写后指纹、现场读不到或授权不完整时只能判 unknown。
 */
export function wpsReconcileVerdict(input: {
  state: WpsExecutionState | null;
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
  /*
   * 还有动作没写。**不给"继续"** —— 精确续跑剩余动作未实现：
   * 目标值矩阵、来源快照、确认快照与宿主版本都只活在会话内内存，重载后拼不回
   * 用户当初确认的那份计划。硬要"续跑"只能拿重载后重新读到的现场拼一份新计划，
   * 那不是续跑，是换个东西写进原表。所以这里如实归成 `replan`。
   */
  return { kind: "replan", remaining };
}

/** 核对结论的用户可见文案（标题级）。 */
export const WPS_RECONCILE_LABEL: Record<WpsReconcileVerdict["kind"], string> = {
  completed: "已完成，仅缺回执",
  reconcile: "待核对在途动作",
  replan: "需重新预览",
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
    case "replan":
      return (
        `还有 ${verdict.remaining} 个动作没写。目标矩阵与确认快照只存在会话内内存，` +
        "重开后无法精确续跑剩余动作 —— 只能按原需求重新读取并生成预览，重新授权确认后再写入。"
      );
    default:
      return verdict.detail;
  }
}

/** 各结论允许出现的操作。返回空数组 = 不给任何写入/补记按钮。 */
export type WpsReconcileAction = "backfill" | "reconcile" | "replan" | "dismiss";

export const WPS_RECOVERY_ACTION_LABEL: Record<WpsReconcileAction, string> = {
  backfill: "补记完成（不重写原表）",
  reconcile: "核对在途动作",
  replan: "重新读取并生成预览",
  dismiss: "清除这条恢复提示"
};

export function wpsReconcileActions(verdict: WpsReconcileVerdict): WpsReconcileAction[] {
  switch (verdict.kind) {
    case "completed":
      return verdict.recorded ? ["dismiss"] : ["backfill", "dismiss"];
    case "reconcile":
      return ["reconcile", "dismiss"];
    case "replan":
      return ["replan", "dismiss"];
    default:
      // `blocked` / `insufficient`：证据不够或现场冲突，一律不给写入或补记按钮，
      // 只留"已人工核对，清除提示"。
      return ["dismiss"];
  }
}
