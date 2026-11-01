/**
 * 「贴底跟随 / 上翻保持」的唯一判据（Excel / WPS 两端共用）。
 *
 * 这套规则同时管**两个**滚动容器，语义必须一模一样：
 *  - 消息流（`.message-stream`）—— 新消息来了要不要跟着走；
 *  - 执行过程框内部（`.task-progress-content`）—— 新步骤来了要不要跟着走。
 *
 * 为什么抽出来：这两个容器各写过一次"到底了没有"，结果一处用"位置"判、另一处
 * 用"有没有新内容"判。后者的表现是：用户上翻看历史、期间恰好没有新内容，
 * 界面上什么都不出现 —— 想回到最新那条只能自己滚回去。判据只能有一份。
 *
 * 三条不变量：
 *  1. **判据看位置，不看有没有新内容**。滚离底部就该出现入口，无论期间是否有更新。
 *  2. **跟随是有状态的**：一旦用户主动上翻就不再跟随，直到他滚回底部或点了入口。
 *     新的内容不许把"跟随"这个意图改回来。
 *  3. **阈值只影响"算不算到底"**，不改变上面两条。
 */

export interface ScrollMetrics {
  scrollHeight: number;
  scrollTop: number;
  clientHeight: number;
}

/** 距底部多少像素以内算"贴底"。两端共用，不再各写一个数。 */
export const SCROLL_BOTTOM_THRESHOLD = 120;

export function isNearScrollBottom(
  metrics: ScrollMetrics,
  threshold = SCROLL_BOTTOM_THRESHOLD
): boolean {
  return (
    Math.max(0, metrics.scrollHeight - metrics.scrollTop - metrics.clientHeight) <=
    threshold
  );
}

/**
 * 一次布局变化之后，这个容器该不该继续跟随。
 *
 * 与 `isNearScrollBottom` 分开是因为"应不应该"和"现在在哪"不是一回事：
 * 用户的意图（`following`）优先；只有从没定过意图时才回落到当前位置。
 * 合并成一个判断的后果是：每次内容变化都会把用户的"我上翻了"覆盖成"你在底部"。
 */
export function shouldFollowAfterLayout(input: {
  /** 用户上一次的意图；`true` = 跟随。 */
  following: boolean;
  metrics: ScrollMetrics;
  threshold?: number;
}): boolean {
  if (input.following) return true;
  return isNearScrollBottom(input.metrics, input.threshold);
}

/** 用户滚动事件 → 新的意图。滚回底部就是重新跟随，不需要别的信号。 */
export function followingAfterScroll(
  metrics: ScrollMetrics,
  threshold = SCROLL_BOTTOM_THRESHOLD
): boolean {
  return isNearScrollBottom(metrics, threshold);
}

/**
 * "回到底部"类入口的文案。消息流与过程框共用一套措辞，只有范围不同。
 *
 * 刻意区分"有新进展"与"回到底部"：
 *  - 还在跑 ⇒ 下面确实在长东西，说"有新进展"是事实；
 *  - 已经停了 ⇒ 下面不再增长，只是位置在别处，说"回到底部"才是事实。
 * 把两者混成一句话会让用户以为点了会触发什么。
 */
export function scrollJumpLabel(input: {
  running: boolean;
  scope: "stream" | "process";
}): { text: string; ariaLabel: string; title: string } {
  const target = input.scope === "stream" ? "最新内容" : "最新步骤";
  if (input.running) {
    return {
      text: "有新进展",
      ariaLabel: `查看${target}`,
      title: `回到底部查看${target}`
    };
  }
  return {
    text: "回到底部",
    ariaLabel: `回到底部查看${target}`,
    title: `回到底部查看${target}`
  };
}
