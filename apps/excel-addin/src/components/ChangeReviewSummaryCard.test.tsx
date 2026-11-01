// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { buildRangeChangeReview } from "@excel-bro/core/changeReview";
import ChangeReviewSummaryCard from "./ChangeReviewSummaryCard";
import { ChangeReviewUndoNote } from "./ChangeReviewPresentation";
import { changeReviewSummary } from "@excel-bro/core/changeReview";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const mounted: Array<() => void> = [];

afterEach(() => {
  while (mounted.length) mounted.pop()?.();
});

async function mount(node: React.ReactElement) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  mounted.push(() => {
    act(() => root.unmount());
    container.remove();
  });
  await act(async () => root.render(node));
  return container;
}

function snapshot(values: unknown[][], formulas: unknown[][], valuesCaptured = true) {
  return { values: values as never, formulas: formulas as never, valuesCaptured };
}

function review(overrides: Partial<Parameters<typeof buildRangeChangeReview>[0]> = {}) {
  return buildRangeChangeReview({
    sheet: "Sheet1",
    range: "A1:A2",
    phase: "actual",
    status: "applied",
    before: snapshot([["旧"], ["保留"]], [[null], [null]]),
    after: snapshot([["新"], ["保留"]], [[null], [null]]),
    ...overrides
  });
}

describe("变更摘要卡 · 侧栏", () => {
  it("只显示定位与计数，不显示任何单元格内容", async () => {
    const container = await mount(
      <ChangeReviewSummaryCard reviews={[review()]} />
    );
    const text = container.textContent ?? "";
    expect(text).toContain("Sheet1!A1:A2");
    expect(text).toContain("改动 1 / 2 个单元格");
    // 侧栏摘要不渲染具体单元格值。
    expect(text).not.toContain("旧");
    expect(text).not.toContain("新");
    expect(container.querySelector("table")).toBeNull();
  });

  it("没有网页对比或写入按钮", async () => {
    const container = await mount(
      <ChangeReviewSummaryCard reviews={[review()]} />
    );
    expect(container.querySelector("button")).toBeNull();
    expect(container.textContent).toContain("变更摘要");
    expect(container.textContent).not.toContain("打开");
  });

  it("预计状态不得冒充已写入", async () => {
    const container = await mount(
      <ChangeReviewSummaryCard reviews={[review({ phase: "expected", status: "ready" })]} />
    );
    expect(container.textContent).toContain("预计，尚未写入");
    expect(container.textContent).not.toContain("已写入并回读");
  });

  it("状态不是 applied 时用「含未核对区域」的措辞，且不出现成功字样", async () => {
    const container = await mount(
      <ChangeReviewSummaryCard
        reviews={[review({ status: "readback_failed" })]}
      />
    );
    const text = container.textContent ?? "";
    expect(text).toContain("含未核对区域");
    expect(text).toContain("回读失败");
    expect(text).not.toContain("已写入并回读");
  });

  it("撤销说明仍保留历史定位和计数，不包含单元格值", async () => {
    const container = await mount(<ChangeReviewUndoNote summaries={[changeReviewSummary(review())]} />);
    expect(container.textContent).toContain("Sheet1!A1:A2");
    expect(container.textContent).toContain("共 2 个单元格");
    expect(container.textContent).toContain("经回读确认改动 1 个");
    expect(container.textContent).toContain("现已撤销");
    expect(container.textContent).not.toContain("旧");
    expect(container.textContent).not.toContain("新");
    expect(container.querySelector("button")).toBeNull();
  });

  it("中断过的执行：三分之一未执行时标出阶段，且不把未核对算成改动", async () => {
    // 真机现场：B2 写入并回读、B3 写后回执丢失（写后快照缺失）、B4 未执行。
    // 旧文案会把这 3 格报成「改动 2 / 3 个单元格」—— 多出来的那 1 格是
    // 「写前非空 → 写后快照缺失」被当成"改成了空"。
    const summaries = [
      { sheet: "目标表", range: "B2", status: "applied", phase: "actual",
        changedCount: 1, totalCount: 1, truncated: false },
      { sheet: "目标表", range: "B3", status: "readback_failed", phase: "actual",
        changedCount: 1, totalCount: 1, truncated: false },
      { sheet: "目标表", range: "B4", status: "not_executed", phase: "actual",
        changedCount: 0, totalCount: 1, truncated: false }
    ] as const;
    const container = await mount(<ChangeReviewUndoNote summaries={summaries} />);
    const text = container.textContent ?? "";
    expect(text).toContain("这次执行中断时");
    expect(text).toContain("共 3 个单元格");
    expect(text).toContain("经回读确认改动 1 个");
    expect(text).toContain("写入但未通过回读核对 1 个");
    expect(text).toContain("未执行 1 个");
    // 旧口径的假改动不许复活。
    expect(text).not.toContain("改动 2 / 3");
    expect(text).not.toContain("2 / 3");
  });

  it("全部回读通过时不标「中断」，也不补出多余的分类", async () => {
    const summaries = [
      { sheet: "目标表", range: "B2", status: "applied", phase: "actual",
        changedCount: 1, totalCount: 1, truncated: false }
    ] as const;
    const container = await mount(<ChangeReviewUndoNote summaries={summaries} />);
    const text = container.textContent ?? "";
    expect(text).toContain("这次执行涉及");
    expect(text).not.toContain("中断");
    expect(text).not.toContain("未执行");
    expect(text).toContain("经回读确认改动 1 个");
  });
});
