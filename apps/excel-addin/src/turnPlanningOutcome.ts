/**
 * 一轮规划结束后的活动终态判定（宿主无关的纯逻辑）。
 *
 * 为什么单独抽出来：`analyzeConfirmedIntent` 内部有多条出口（澄清、暂停、
 * 回答、计划、本地映射失败……），调用方随后还要不要补一次 `finishActivity()`
 * 取决于这一轮到底有没有产出计划。此前无条件补一次，导致**规划失败也被记成
 * "已完成"** —— 本地列映射读取/计算失败时尤其明显：错误也报了，进度条却写成
 * "已完成"，用户以为写完了。
 *
 * 判定规则只有一条：**只有真正产出计划（或澄清/回答这类正常终态）才允许
 * 用默认的"已完成"收尾**；失败、打断、跳过一律不补，交由失败路径自己收尾。
 */

/** 一轮规划的结果分类。 */
export type TurnPlanningOutcome =
  | "planned"
  | "failed"
  | "aborted"
  | "skipped";

/**
 * 这一轮结束后，调用方是否应当再调用一次默认的 `finishActivity()`
 * （即把当前步骤标成"已完成"）。
 *
 * 只有 `planned` 才为真。失败已由失败路径以 `failed`/`warning` 收尾，
 * 打断已收尾，跳过意味着没有产出计划（比如澄清待答），都不该被标成完成。
 */
export function shouldFinishActivityAsCompleted(
  outcome: TurnPlanningOutcome
): boolean {
  return outcome === "planned";
}
