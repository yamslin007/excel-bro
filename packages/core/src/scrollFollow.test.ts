import { describe, expect, it } from "vitest";
import {
  SCROLL_BOTTOM_THRESHOLD,
  followingAfterScroll,
  isNearScrollBottom,
  scrollJumpLabel,
  shouldFollowAfterLayout
} from "./scrollFollow";

const metrics = (scrollHeight: number, scrollTop: number, clientHeight: number) => ({
  scrollHeight,
  scrollTop,
  clientHeight
});

describe("isNearScrollBottom", () => {
  it("贴底与阈值边界都算到底", () => {
    expect(isNearScrollBottom(metrics(1000, 400, 600))).toBe(true);
    expect(isNearScrollBottom(metrics(1000, 400 - SCROLL_BOTTOM_THRESHOLD, 600))).toBe(true);
  });

  it("超出阈值一步就不算到底", () => {
    expect(isNearScrollBottom(metrics(1000, 400 - SCROLL_BOTTOM_THRESHOLD - 1, 600))).toBe(false);
  });

  it("内容比视口短（无处可滚）也算贴底", () => {
    expect(isNearScrollBottom(metrics(300, 0, 600))).toBe(true);
  });
});

describe("shouldFollowAfterLayout", () => {
  it("用户已表态跟随 ⇒ 无论滚到哪都跟随", () => {
    expect(shouldFollowAfterLayout({ following: true, metrics: metrics(2000, 0, 600) })).toBe(true);
  });

  it("用户已上翻 ⇒ 内容变长也不把意图改回跟随", () => {
    // 关键：这是"新内容不许抢滚动"的唯一依据。位置看起来离底部很远，
    // 但真正的原因是用户自己翻上去了 —— 不能因为内容增长就把它拉下来。
    expect(shouldFollowAfterLayout({ following: false, metrics: metrics(4000, 100, 600) })).toBe(false);
  });

  it("从没定过意图时回落到当前位置", () => {
    expect(shouldFollowAfterLayout({ following: false, metrics: metrics(1000, 400, 600) })).toBe(true);
  });
});

describe("followingAfterScroll", () => {
  it("滚回底部就是重新跟随（不要求期间有新内容）", () => {
    expect(followingAfterScroll(metrics(1200, 600, 600))).toBe(true);
    expect(followingAfterScroll(metrics(1200, 200, 600))).toBe(false);
  });
});

describe("scrollJumpLabel", () => {
  it("运行中说的是有新进展，停下说的是回到底部", () => {
    const running = scrollJumpLabel({ running: true, scope: "stream" });
    const idle = scrollJumpLabel({ running: false, scope: "stream" });
    expect(running.text).toBe("有新进展");
    expect(idle.text).toBe("回到底部");
    // 两者必须不同：停下之后下面不再增长，说"有新进展"就是假的。
    expect(running.text).not.toBe(idle.text);
  });

  it("过程框与消息流指向不同的东西", () => {
    const stream = scrollJumpLabel({ running: true, scope: "stream" });
    const process = scrollJumpLabel({ running: true, scope: "process" });
    expect(stream.ariaLabel).toContain("最新内容");
    expect(process.ariaLabel).toContain("最新步骤");
    expect(stream.ariaLabel).not.toBe(process.ariaLabel);
  });
});
