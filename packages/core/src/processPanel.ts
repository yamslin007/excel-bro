/**
 * 执行过程框的**呈现判据**（Excel / WPS 两端共用，纯函数）。
 *
 * 这一层只回答"框该多高、要不要收起、收起后那一行写什么"，**不碰 React、
 * 不碰宿主、不碰滚动**。为什么必须只有一份：这套规则一旦两端各写一次，
 * 就会出现"Excel 收起了、WPS 还挂着"这种同一件事两种观感，而用户来回切换
 * 两端时看到的其实是同一条任务。
 *
 * 三条不变量（改动前先读）：
 *
 * A. **收起不是删除。** 收起的只是详细动作区；阶段结论、耗时、以及在折叠区
 *    **外**的错误原因 / 阻断说明 / 澄清问题 / 操作按钮一律照旧可见。
 *    把原因藏进折叠区等于让用户点了才知道失败了。
 *
 * B. **用户的选择优先于任何自动收起。** 自动收起只发生在"这一次从进行中变成
 *    正常结束"的那个瞬间，而且只在用户**没有明确表过态**时发生。一旦用户点过
 *    展开，后续任何状态变化（新步骤、重渲染、结束）都不许把它收回去。
 *
 * C. **限高不是截断。** 进行中限高是为了不让过程框把输入框挤出视口，内容永远
 *    可以在框内滚到；所以"内容比限高长"只影响要不要给收起/跳转入口，
 *    不影响是否显示全部步骤。
 */

/** 「进行中」的默认限高（px）。约 160px = 4~5 行步骤，够看出在做什么。 */
export const PROCESS_PANEL_LIVE_HEIGHT = 160;
/** 用户主动要求"多看点"时，进行中的限高放宽到（px）。 */
export const PROCESS_PANEL_LIVE_EXPANDED_HEIGHT = 420;
/** 已经结束的过程，未主动放宽时的限高（px）。 */
export const PROCESS_PANEL_SETTLED_HEIGHT = 320;
/** 已经结束的过程，用户主动放宽后的限高（px）。 */
export const PROCESS_PANEL_SETTLED_EXPANDED_HEIGHT = 560;

/** 用户对这一个过程框**明确表过**的展开/收起意图；`null` = 还没表态。 */
export type ProcessPanelUserChoice = "expanded" | "collapsed" | null;

/** 只压缩已结束的单步消息理解；有工具、现场证据或异常时保留完整记录。 */
export function isCompactMessageProcess(log?: {
  outcome?: string;
  connectionMessage?: string;
  steps: readonly {
    id?: string; status?: string; phase?: string;
    toolName?: string | null; actionType?: string | null; target?: unknown;
    facts?: readonly unknown[] | null; trace?: unknown;
  }[];
} | null): boolean {
  if (!log || log.outcome !== "completed" || log.connectionMessage || log.steps.length !== 1) return false;
  const step = log.steps[0];
  return step.id === "message-intent" && step.status === "completed" && step.phase === "intent"
    && !step.toolName && !step.actionType && !step.target && !step.facts?.length && !step.trace;
}

/**
 * 这一次渲染该不该展开详细动作区。
 *
 * 进行中、正常结束和等待回答默认收起；等待回答时问题在对话中直接展示。
 * 失败、中断、取消保留详情，因为用户需要知道"停在哪一步"。
 * 把失败原因折起来，用户看到的就只剩一行"执行失败"，
 * 然后必须自己想到去点展开。
 */
export function processPanelExpanded(input: {
  live: boolean;
  outcome?: string | null;
  userChoice: ProcessPanelUserChoice;
  awaitingAnswer?: boolean;
}): boolean {
  if (input.userChoice === "expanded") return true;
  if (input.userChoice === "collapsed") return false;
  if (input.live) return false;
  // Waiting for an answer shows the question in the conversation, not an expanded log.
  return input.outcome !== "completed" && !(input.outcome === "awaiting_input" && input.awaitingAnswer);
}

/**
 * 详细动作区的限高上限（px）。返回值交给 CSS 变量，再由 CSS 用
 * `min(var(--…), 46vh)` 与窗格高度取小 —— 窗格适配留在 CSS，这里只给业务上限。
 *
 * 参数是 `enlarged`（用户明确要求看更多）而**不是** `expanded`（详情是否显示）：
 * 这两件事在"进行中"里是分开的 —— 进行中默认就显示详情，但只给 160px，
 * 内部自己滚。把两者混成一个布尔的结果是"进行中一上来就占 420px"，
 * 输入框被顶下去，而用户并没有要求过。
 */
export function processPanelBodyMaxHeight(input: {
  live: boolean;
  enlarged: boolean;
}): number {
  if (input.live) {
    return input.enlarged ? PROCESS_PANEL_LIVE_EXPANDED_HEIGHT : PROCESS_PANEL_LIVE_HEIGHT;
  }
  return input.enlarged ? PROCESS_PANEL_SETTLED_EXPANDED_HEIGHT : PROCESS_PANEL_SETTLED_HEIGHT;
}

/**
 * 内容比限高长 ⇒ 需要给出"收起/跳到最新"之类的入口。
 *
 * 单独成函数而不是在组件里写 `scrollHeight > 160`：那个 160 一旦改了，
 * 判据就悄悄失配（要么永远不出入口，要么一直在出）。
 */
export function processPanelHasOverflow(input: {
  live: boolean;
  enlarged: boolean;
  contentHeight: number;
}): boolean {
  return input.contentHeight > processPanelBodyMaxHeight(input);
}

/** 耗时文案。与过程框内部的步骤计时同一口径（秒，一位小数），不再另写一套。 */
export function formatProcessDuration(ms: number): string {
  const value = Math.max(0, ms);
  if (value < 1000) return Math.round(value) + "ms";
  if (value < 60_000) return (value / 1000).toFixed(1) + "s";
  const minutes = Math.floor(value / 60_000);
  const seconds = Math.round((value % 60_000) / 1000);
  return seconds > 0 ? `${minutes}分${seconds}秒` : `${minutes}分`;
}

/**
 * 收起之后那一行：**阶段结论 + 耗时**。
 *
 * 刻意只说这两件靠得住的事：结论来自活动级回执（`outcome`），耗时来自计时分账。
 * 不写"已完成 3 处修改"这种需要另找证据的话 —— 那是结果摘要的活，
 * 由 `processOutcomeSummary` 从结构化回执派生，不在这里猜。
 */
export function foldedProcessLine(input: {
  conclusion: string;
  activeMs: number;
}): string {
  const conclusion = input.conclusion.trim() || "已结束";
  return `${conclusion} · ${formatProcessDuration(input.activeMs)}`;
}

/**
 * 折叠区外必须可见的东西。调用方（两端面板）据此渲染，**不据此决定是否收起**。
 *
 * 这里给的是"哪几类必须露在外面"，具体文案由各端已有的卡片提供；
 * 之所以列成常量而不是函数：它是一条纪律，不是计算。
 */
export const PROCESS_PANEL_OUTSIDE_FOLD = [
  "错误原因",
  "阻断或核对结论",
  "澄清问题",
  "操作按钮"
] as const;
