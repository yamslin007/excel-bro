// @vitest-environment jsdom
import { cleanup, fireEvent, render } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { ComposerPendingBar } from "./ComposerPendingBar";
import type { ComposerPendingItem } from "@excel-bro/core/composerPending";

afterEach(cleanup);

const statusItems: ComposerPendingItem[] = [
  { id: "a", label: "授权范围待确认", kind: "status" },
  { id: "b", label: "恢复核对", kind: "recovery" }
];

describe("底部常驻事项栏（整个输入容器上方）", () => {
  it("没有事项时整块不渲染，不占底部高度", () => {
    const { container } = render(
      <ComposerPendingBar items={[]}><p>不该出现</p></ComposerPendingBar>
    );
    expect(container.querySelector(".composer-pending")).toBeNull();
    expect(container.textContent).toBe("");
  });

  it("默认只给一行紧凑摘要，详细内容仍在 DOM 里可展开", () => {
    const { container } = render(
      <ComposerPendingBar items={statusItems}>
        <p data-testid="detail">恢复核对详情</p>
      </ComposerPendingBar>
    );
    const pending = container.querySelector<HTMLElement>(".composer-pending")!;
    expect(pending.getAttribute("data-pending-count")).toBe("2");
    expect(pending.getAttribute("data-pending-open")).toBe("false");
    expect(pending.querySelector(".composer-pending-text")?.textContent).toContain("待处理 2 项");
    expect(pending.querySelector(".composer-pending-text")?.textContent).toContain("恢复核对");
    // 收起不等于删除：详细内容仍在 DOM，点开就能用。
    expect(pending.querySelector('[data-testid="detail"]')).not.toBeNull();
  });

  it("点摘要可展开，再点可收起", () => {
    const { container } = render(
      <ComposerPendingBar items={statusItems}><p>详情</p></ComposerPendingBar>
    );
    const summary = container.querySelector<HTMLElement>(".composer-pending-summary")!;
    const pending = container.querySelector<HTMLElement>(".composer-pending")!;
    fireEvent.click(summary);
    expect(pending.getAttribute("data-pending-open")).toBe("true");
    fireEvent.click(summary);
    expect(pending.getAttribute("data-pending-open")).toBe("false");
  });

  it("需要用户做决定的事项默认展开，不把主入口藏起来", () => {
    const { container } = render(
      <ComposerPendingBar items={[
        { id: "d", label: "待审阅决定", kind: "decision" },
        ...statusItems
      ]}>
        <button type="button">确认全部变更</button>
      </ComposerPendingBar>
    );
    expect(container.querySelector(".composer-pending")?.getAttribute("data-pending-open")).toBe("true");
    expect(container.textContent).toContain("确认全部变更");
  });
});
