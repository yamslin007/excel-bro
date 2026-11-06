// @vitest-environment jsdom
import { cleanup, render } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { ExecutionFeedbackSummary } from "./ExecutionFeedbackSummary";
import type { OperationFeedbackRecord } from "@excel-bro/core/operationFeedback";

afterEach(cleanup);

const records: OperationFeedbackRecord[] = [
  { executionId: "exec-1", kind: "discarded", text: "已放弃并恢复：工作簿已按恢复点还原。" },
  { executionId: "exec-1", kind: "discarded", text: "重复的一条" },
  {
    executionId: "exec-1", kind: "copy_cleanup_failed", text: "展示副本未清理",
    issues: [
      { identity: "s1:EB_aaa", label: "副本1", reason: "窗口仍在" },
      { identity: "s1:EB_bbb", label: "副本2", reason: "窗口仍在" }
    ]
  }
];

describe("归属到执行记录的操作反馈", () => {
  it("同一次执行的多条反馈合并成一行，副本按真实身份去重给数量", () => {
    const { container } = render(<ExecutionFeedbackSummary records={records} />);
    const line = container.querySelector(".execution-feedback-line")?.textContent ?? "";
    expect(line).toContain("已放弃并恢复");
    expect(line).toContain("展示副本未清理 · 2 份");
    expect(line.match(/已放弃并恢复/g)).toHaveLength(1);
    expect(container.querySelector(".execution-feedback")?.getAttribute("data-copy-issue-count")).toBe("2");
  });

  it("逐份原因可展开，且不说恢复失败、不假称全部清理完成", () => {
    const { container } = render(<ExecutionFeedbackSummary records={records} />);
    const details = container.querySelector(".execution-feedback-details")!;
    expect(details.querySelector("summary")?.textContent).toContain("2 份");
    expect(details.querySelectorAll("li")).toHaveLength(2);
    expect(container.textContent).not.toContain("恢复失败");
    expect(container.textContent).not.toContain("已全部清理");
  });

  it("没有副本问题时不画展开面板", () => {
    const { container } = render(
      <ExecutionFeedbackSummary records={[records[0]]} />
    );
    expect(container.querySelector(".execution-feedback-details")).toBeNull();
    expect(container.textContent).toContain("已放弃并恢复");
  });
});
