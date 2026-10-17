// @vitest-environment jsdom
import { useLayoutEffect } from "react";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { AuthorizedTargetConfirmation } from "./AuthorizedTargetConfirmation";

afterEach(cleanup);
const target: AuthorizedTarget = {
  fileKey: "synthetic-original", name: "Original.xlsx",
  sheets: [{ id: "1", name: "Sheet1" }], selected: ["1"]
};
const current = { ...target, fileKey: "synthetic-current", name: "Current.xlsx" };

describe("target dialog first interaction", () => {
  it("confirms one current scope without duplicate original/current choices", () => {
    const confirm = vi.fn();
    render(<AuthorizedTargetConfirmation
      notice={{ bound: target, current: target, reason: "confirmation_required", round: 1 }}
      error="" busy={false} onConfirm={confirm} onCancel={vi.fn()} />);
    expect(screen.getByRole("dialog").textContent).toContain("确认本次操作范围");
    expect(screen.queryAllByRole("radio")).toHaveLength(0);
    expect(screen.getAllByText("Original.xlsx · Sheet1")).toHaveLength(1);
    expect(screen.getByRole("dialog").textContent).not.toContain("重新确认");
    fireEvent.click(screen.getByText("确认范围并继续"));
    expect(confirm).toHaveBeenCalledWith("original", []);
  });
  it("unavailable identity offers recheck, never an authorization override", () => {
    const confirm = vi.fn(); const recheck = vi.fn();
    render(<AuthorizedTargetConfirmation
      notice={{ bound: target, current: target, reason: "unavailable", round: 1 }}
      error="" busy={false} onConfirm={confirm} onCancel={vi.fn()} onRecheck={recheck} />);
    expect(screen.getByRole("dialog").textContent).toContain("暂时无法核实操作目标");
    expect(screen.queryAllByRole("radio")).toHaveLength(0);
    expect(screen.queryByText("确认目标并继续")).toBeNull();
    fireEvent.click(screen.getByText("重新检查"));
    expect(recheck).toHaveBeenCalledOnce();
    expect(confirm).not.toHaveBeenCalled();
  });
  it.each(["original", "current"] as const)("does not erase an immediate %s choice with a mounting reset", (choice) => {
    const confirm = vi.fn();
    function Harness() {
      // Click after the child DOM is committed, before passive effects run.
      // A delayed initial reset used to erase this interaction.
      useLayoutEffect(() => {
        (screen.getByLabelText(choice === "original" ? /继续使用原授权/ : /改用当前文件/) as HTMLInputElement).click();
      }, []);
      return <AuthorizedTargetConfirmation
        notice={{ bound: target, current, reason: "file_changed", round: 1 }}
        error="" busy={false} onConfirm={confirm} onCancel={vi.fn()} />;
    }
    render(<Harness />);
    expect((screen.getByLabelText(choice === "original" ? /继续使用原授权/ : /改用当前文件/) as HTMLInputElement).checked).toBe(true);
    if (choice === "current") fireEvent.click(screen.getByLabelText("Sheet1"));
    fireEvent.click(screen.getByText("确认目标并继续"));
    expect(confirm).toHaveBeenCalledWith(choice, choice === "current" ? ["1"] : []);
  });
});
