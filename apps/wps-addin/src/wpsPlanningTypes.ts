/**
 * WPS 规划通道的结构化结果类型（CAP-02）。
 *
 * 存在的唯一理由：**流式与非流式必须共用同一个出口**。
 *
 * 一次性通道（`runWpsTurn` 里的 `createAssistantResponse`）和流式通道
 * （`streamWpsPlanning` / `continueWpsPlanning`）在服务端是两条传输路径，
 * 但产出的必须是**同一个结构化联合**：`answer` / `plan` / `clarification` /
 * `local_plan_rule`，外加暂停态的 `planning_pause`。如果两条通道各写一个类型，
 * 分支处理就会在两处漂移 —— 典型症状是流式路径把 `local_plan_rule` 当 `answer`
 * 渲染，用户看不到"需要补映射"这张卡。
 *
 * 因此这里**不重定义** `AssistantResponse`，只是把 `PlanningPauseResult`
 * 拼上来，形成 `AssistantResponseOrPause`。所有分支处理（澄清、本地规则、
 * 计划、回答）对两条通道只有一份实现。
 */
import type {
  AssistantResponse,
  PlanningPauseResult
} from "../../excel-addin/src/contracts";

export type { AssistantResponse, PlanningPauseResult };

/**
 * 规划通道的完整结构化出口。
 *
 * 消费方必须逐项覆盖：`answer`、`plan`、`clarification`、`local_plan_rule`、
 * `planning_pause`，以及服务端以 `error` 帧 / 非 2xx 表达失败、以终态
 * `cancelled` 表达取消、以 404 `PLANNING_TASK_NOT_FOUND` 表达过期。
 */
export type AssistantResponseOrPause = AssistantResponse | PlanningPauseResult;

/** 是否为暂停态。用 `kind` 判别，**不看助手文本**。 */
export function isPlanningPause(
  value: AssistantResponseOrPause
): value is PlanningPauseResult {
  return value.kind === "planning_pause";
}

/**
 * 结构化结果的种类清单。新增种类时必须同步这里，否则
 * `assertKnownResultKind` 会在运行期报错而不是静默漏处理。
 */
export const WPS_RESULT_KINDS = [
  "answer",
  "plan",
  "clarification",
  "local_plan_rule",
  "planning_pause"
] as const;

export type WpsResultKind = (typeof WPS_RESULT_KINDS)[number];

/**
 * 校验结果种类在已知集合内。
 *
 * **不得把未知种类统一当 `answer` 处理** —— 那正是"表面成功"的来源：
 * 卡片画出来了，但底层是一次压根没被识别的结果。
 */
export function assertKnownResultKind(value: AssistantResponseOrPause): WpsResultKind {
  const kind = (value as { kind?: unknown }).kind;
  if (typeof kind !== "string" || !(WPS_RESULT_KINDS as readonly string[]).includes(kind)) {
    throw new Error(`规划返回了本端尚不支持的结果类型：${String(kind)}`);
  }
  return kind as WpsResultKind;
}
