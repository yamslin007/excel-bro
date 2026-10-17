// @vitest-environment jsdom
import { cleanup, fireEvent, render } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { TargetScopeNotice } from "./TargetScopeNotice";
import type { AuthorizedTarget, TargetNoticeReason } from "@excel-bro/core/authorizedTarget";

afterEach(cleanup);
const target: AuthorizedTarget = { name: "同名.xlsx", fileKey: "PRIVATE_PATH", sheets: [{ id: "1", name: "Sheet1" }], selected: ["1"] };
describe("one scope selector, no second file-choice dialog", () => {
  it("opens the existing selector instead of offering duplicate same-name choices", () => {
    const open = vi.fn();
    const ui = render(<TargetScopeNotice notice={{ bound: target, current: target, reason: "scope_changed", round: 1 }}
      error="" busy={false} onOpenScope={open} />);
    expect(ui.queryByRole("dialog")).toBeNull();
    expect(ui.queryAllByRole("radio")).toHaveLength(0);
    expect(ui.queryByText(/继续使用原授权/)).toBeNull();
    expect(ui.queryByText(/改用当前文件/)).toBeNull();
    fireEvent.click(ui.getByText("检查数据范围"));
    expect(open).toHaveBeenCalledOnce();
    expect(ui.container.textContent).not.toContain("PRIVATE_PATH");
  });
  it("unknown identity opens the same selector without claiming that the file changed", () => {
    const open = vi.fn();
    const ui = render(<TargetScopeNotice notice={{ bound: target, current: null, reason: "unavailable", round: 1 }}
      error="" busy={false} onOpenScope={open} />);
    expect(ui.container.textContent).toContain("暂时无法核实");
    expect(ui.container.textContent).not.toContain("发生变化");
    expect(ui.getAllByRole("button")).toHaveLength(1);
    expect(ui.queryByText("重新检查")).toBeNull();
    fireEvent.click(ui.getByRole("button", { name: "检查数据范围" }));
    expect(open).toHaveBeenCalledOnce();
  });
  it.each<TargetNoticeReason>(["file_changed", "scope_changed", "unavailable", "confirmation_required"])(
    "%s has one action and does nothing until the user opens the selector", (reason) => {
      const open = vi.fn();
      const ui = render(<TargetScopeNotice notice={{ bound: target, current: target, reason, round: 1 }}
        error="" busy={false} onOpenScope={open} />);
      expect(ui.getAllByRole("button")).toHaveLength(1);
      expect(ui.queryByText("暂不操作")).toBeNull();
      expect(open).not.toHaveBeenCalled();
      expect(ui.container.textContent).toContain("需求与回答仍保留");
      expect(ui.container.textContent).toContain("暂停");
    }
  );
  it("shows the original and current sheet ranges without exposing private file identity", () => {
    const current: AuthorizedTarget = { ...target, sheets: [{ id: "5", name: "Sheet5" }], selected: ["5"] };
    const ui = render(<TargetScopeNotice notice={{ bound: target, current, reason: "scope_changed", round: 1 }}
      error="" busy={false} onOpenScope={vi.fn()} />);
    expect(ui.container.textContent).toContain("原范围：同名.xlsx · Sheet1");
    expect(ui.container.textContent).toContain("当前：同名.xlsx · Sheet5");
    expect(ui.container.textContent).not.toContain("PRIVATE_PATH");
  });
  it("disables its only action during confirmation and retains validation errors", () => {
    const open = vi.fn();
    const ui = render(<TargetScopeNotice notice={{ bound: target, current: target, reason: "scope_changed", round: 1 }}
      error="目标暂不可用" busy onOpenScope={open} />);
    const button = ui.getByRole("button", { name: "检查数据范围" }) as HTMLButtonElement;
    expect(button.disabled).toBe(true);
    expect(ui.getAllByRole("button")).toHaveLength(1);
    expect(ui.getByRole("alert").textContent).toBe("目标暂不可用");
    fireEvent.click(button);
    expect(open).not.toHaveBeenCalled();
  });
});
