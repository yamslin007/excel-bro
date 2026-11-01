// @vitest-environment jsdom
import { cleanup, render } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { selectCurrentAction } from "@excel-bro/core/currentActionCard";
import { ComposerActionDock, type CurrentActionView } from "./ComposerActionDock";

afterEach(cleanup);

const views = {
  review: <div data-testid="review-view">审阅</div>,
  recovery: <div data-testid="recovery-view">恢复</div>,
  clarification: <div data-testid="clarification-view">回答</div>
};

const dock = (
  input: Parameters<typeof selectCurrentAction>[0],
  pinned?: CurrentActionView | null,
  onSwitch?: (kind: CurrentActionView) => void
) => {
  const selection = selectCurrentAction(input, pinned ?? null);
  return render(<ComposerActionDock selection={selection} views={views} onSwitch={onSwitch} />);
};

const all = { review: true, recovery: true, clarification: true, running: false };

describe("固定区域只展开一块", () => {
  it("审阅与恢复同时在场时，只展开审阅那一块", () => {
    const { container } = dock(all, "review");
    expect(container.querySelector('[data-dock-view="review"]')).not.toBeNull();
    // 判别点：另一块**不能**同时展开 —— 那正是"只算出名字、四块照样全开"的旧形状。
    expect(container.querySelector('[data-dock-view="recovery"]')).toBeNull();
    expect(container.querySelectorAll("[data-dock-view]")).toHaveLength(1);
  });

  it("显式切到恢复后，展开的是恢复，审阅退成 chip（不是消失）", () => {
    const { container } = dock(all, "recovery", vi.fn());
    expect(container.querySelector('[data-dock-view="recovery"]')).not.toBeNull();
    expect(container.querySelector('[data-dock-view="review"]')).toBeNull();
    // 被换下来的那件事仍然列在 chip 行里 —— 否则"切过去"等于把原来那件丢了。
    expect(container.querySelector('[data-dock-switch="review"]')).not.toBeNull();
  });

  it("没有待处理卡片时整块容器不渲染", () => {
    const { container } = dock({ review: false, recovery: false, clarification: false, running: false });
    expect(container.firstChild).toBeNull();
  });
});

describe("其余事项：紧凑提示 + 显式切换入口", () => {
  it("可切的那两件是按钮，点一下就报出要切到哪一件", () => {
    const onSwitch = vi.fn();
    const { container } = dock(
      { review: true, recovery: true, clarification: true, running: false },
      "review",
      onSwitch
    );
    const chip = container.querySelector<HTMLButtonElement>('[data-dock-switch="recovery"]');
    expect(chip).not.toBeNull();
    expect(chip!.textContent).toContain("恢复核对");
    chip!.click();
    expect(onSwitch).toHaveBeenCalledWith("recovery");
  });

  it("澄清不是可切的 chip：它就在输入框下方，折起来会让输入框在说另一件事", () => {
    const onSwitch = vi.fn();
    const { container } = dock(all, "review", onSwitch);
    const inplace = container.querySelector('[data-dock-inplace="clarification"]');
    expect(inplace).not.toBeNull();
    expect(inplace!.tagName).not.toBe("BUTTON");
    expect(inplace!.textContent).toContain("就在下方");
    onSwitch.mockClear();
    (inplace as HTMLElement).click();
    expect(onSwitch).not.toHaveBeenCalled();
  });

  it("「在跑」只说明控制在右下角，不在固定区域再画一个运行按钮", () => {
    const { container } = dock({ review: true, recovery: false, clarification: false, running: true }, "review", vi.fn());
    const inplace = container.querySelector('[data-dock-inplace="runtime"]');
    expect(inplace).not.toBeNull();
    expect(inplace!.textContent).toContain("右下角");
    expect(container.querySelectorAll("button")).toHaveLength(0);
  });

  it("只有一件事时不产生 chip 行（不重复造入口）", () => {
    const { container } = dock({ review: true, recovery: false, clarification: false, running: false });
    expect(container.querySelector(".composer-action-dock-others")).toBeNull();
    expect(container.querySelector('[data-dock-others="0"]')).not.toBeNull();
  });

  it("其余待办逐条点名，不是只给一个数字（原来那块摘要的唯一价值）", () => {
    const { container } = dock(all, "review", vi.fn());
    const row = container.querySelector<HTMLElement>(".composer-action-dock-others")!;
    // 两件其余待办各占一个 chip；计数与条数必须一致，否则就是"算了个数却没说是什么"。
    expect(container.querySelectorAll("[data-dock-switch], [data-dock-inplace]")).toHaveLength(2);
    expect(container.querySelector('[data-dock-others="2"]')).not.toBeNull();
    expect(row.textContent).toContain("恢复核对");
    expect(row.textContent).toContain("回答提问");
  });

  it("没有给切换回调时不画假按钮，只说明", () => {
    const { container } = dock(all, "review");
    expect(container.querySelector('[data-dock-switch]')).toBeNull();
    expect(container.querySelector('[data-dock-inplace="recovery"]')).not.toBeNull();
  });
});
