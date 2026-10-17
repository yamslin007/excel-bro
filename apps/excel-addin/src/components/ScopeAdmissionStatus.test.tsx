// @vitest-environment jsdom
import { cleanup, render } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { admissionItem } from "@excel-bro/core/scopeAdmission";
import { ScopeAdmissionSummary } from "./ScopeAdmissionStatus";

afterEach(cleanup);

describe("scope admission summary", () => {
  it("shows just one normal status line without another action or consent explanation", () => {
    const ui = render(<ScopeAdmissionSummary items={[admissionItem("one")]} count={1} checking={false} />);
    expect(ui.getByRole("status").textContent).toBe("已选 1 张表 · 可分析");
    expect(ui.queryByRole("button")).toBeNull();
    expect(ui.container.querySelector("small")).toBeNull();
  });
  it("does not announce availability while the check is pending or incomplete", () => {
    const ui = render(<ScopeAdmissionSummary items={[]} count={1} checking={true} />);
    expect(ui.getByRole("status").textContent).toBe("已选 1 张表 · 正在检查…");
    ui.rerender(<ScopeAdmissionSummary items={[]} count={1} checking={false} />);
    expect(ui.getByRole("status").textContent).not.toContain("可分析");
  });
  it("keeps a specific failure reason, directing retry to refresh instead of a second button", () => {
    const ui = render(<ScopeAdmissionSummary items={[admissionItem("one", "identity_unverified")]} count={1} checking={false} />);
    expect(ui.getByRole("status").textContent).toContain("1 张不可用");
    expect(ui.getByRole("status").textContent).toContain("请刷新后重试");
    expect(ui.queryByRole("button")).toBeNull();
  });
  it("retains real write restrictions without claiming unknown write access is a proven restriction", () => {
    const ui = render(<ScopeAdmissionSummary items={[admissionItem("one", "protected")]} count={1} checking={false} />);
    expect(ui.getByRole("status").textContent).toContain("1 张写入受限");
    expect(ui.getByRole("status").textContent).toContain("工作表受保护");
    ui.rerender(<ScopeAdmissionSummary items={[admissionItem("one", "write_unverified")]} count={1} checking={false} />);
    expect(ui.getByRole("status").textContent).toContain("写入状态待核实");
    expect(ui.getByRole("status").textContent).not.toContain("1 张写入受限");
  });
  it("does not show stale permission reasons after clearing or changing the selection", () => {
    const ui = render(<ScopeAdmissionSummary items={[admissionItem("old", "unreadable")]} count={0} checking={false} />);
    expect(ui.getByRole("status").textContent).toBe("请先选择工作表");
  });
});
