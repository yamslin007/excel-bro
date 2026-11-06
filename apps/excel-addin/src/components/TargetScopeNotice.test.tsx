// @vitest-environment jsdom
import { cleanup, fireEvent, render } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { TargetScopeNotice } from "./TargetScopeNotice";
import type { AuthorizedTarget, TargetNoticeReason } from "@excel-bro/core/authorizedTarget";

afterEach(cleanup);
const target: AuthorizedTarget = { name: "同名.xlsx", fileKey: "PRIVATE_PATH", sheets: [{ id: "1", name: "Sheet1" }], selected: ["1"] };
describe("one scope selector, no second file-choice dialog", () => {
  it("rechecks unavailable evidence without asking the user to select the same sheets again", () => {
    const recheck = vi.fn(); const open = vi.fn();
    const ui = render(<TargetScopeNotice notice={{ bound: target, current: null, reason: "unavailable", round: 1 }}
      error="" busy={false} onOpenScope={open} onRecheck={recheck} />);
    expect(ui.container.textContent).toContain("已保留勾选范围");
    expect(ui.container.textContent).not.toContain("检查并确认");
    fireEvent.click(ui.getByRole("button", { name: "重新检查" }));
    expect(recheck).toHaveBeenCalledOnce();
    expect(open).not.toHaveBeenCalled();
  });
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
  /*
   * `confirmation_required` 是"读取范围"的确认，不是写入前的第二道确认。
   * 有原样确认处理器时，卡片**直接**给这个按钮（复用同一个确认处理器），
   * 不需要用户去面板里取消勾选再勾回。
   */
  it("confirmation_required 直接给「确认范围并继续」，点击只走原样确认", () => {
    const open = vi.fn(); const confirmOriginal = vi.fn();
    const ui = render(<TargetScopeNotice notice={{ bound: target, current: target, reason: "confirmation_required", round: 1 }}
      error="" busy={false} onOpenScope={open} onConfirmOriginal={confirmOriginal} />);
    fireEvent.click(ui.getByRole("button", { name: "确认范围并继续" }));
    expect(confirmOriginal).toHaveBeenCalledOnce();
    expect(open).not.toHaveBeenCalled();
    // 选择面板仍然可用（用户确实想改范围时才去），但不是唯一出路。
    expect(ui.getByRole("button", { name: "检查数据范围" })).toBeTruthy();
    // 这是读取范围确认：文案里不出现"写入"之外的第二道确认入口。
    expect(ui.container.textContent).not.toContain("确认写入");
  });
  it("窗格重载后确需确认：说明已重新加载与已选表数，不暗示用户改过勾选", () => {
    const four: AuthorizedTarget = {
      name: "台账.xlsx", fileKey: "PRIVATE_PATH",
      sheets: ["来源", "结果1", "结果2", "结果3"].map((name, index) => ({ id: String(index + 1), name })),
      selected: ["1", "2", "3", "4"]
    };
    const ui = render(<TargetScopeNotice notice={{ bound: four, current: four, reason: "confirmation_required", round: 1 }}
      error="" busy={false} onOpenScope={vi.fn()} onConfirmOriginal={vi.fn()} restoredScope />);
    const text = ui.container.textContent ?? "";
    expect(text).toContain("窗格已重新加载，请确认继续使用已选的 4 张工作表。");
    expect(text).toContain("工作表勾选没有变动");
    expect(text).not.toContain("工作表范围发生变化");
    expect(text).not.toContain("当前文件与已授权文件不同");
    expect(text).not.toContain("PRIVATE_PATH");
  });
  it("没有重载线索时说通用文案，不假装窗格重载过", () => {
    const ui = render(<TargetScopeNotice notice={{ bound: target, current: target, reason: "confirmation_required", round: 1 }}
      error="" busy={false} onOpenScope={vi.fn()} onConfirmOriginal={vi.fn()} />);
    const text = ui.container.textContent ?? "";
    expect(text).not.toContain("窗格已重新加载");
    expect(text).toContain("同一会话内范围未变时不再重复询问");
  });
  it("真正换文件时不给原样确认，只说换了什么", () => {
    const confirmOriginal = vi.fn();
    const ui = render(<TargetScopeNotice notice={{ bound: target, current: { ...target, fileKey: "OTHER" }, reason: "file_changed", round: 1 }}
      error="" busy={false} onOpenScope={vi.fn()} onConfirmOriginal={confirmOriginal} />);
    expect(ui.queryByRole("button", { name: "确认范围并继续" })).toBeNull();
    expect(ui.container.textContent).toContain("当前文件与已授权文件不同");
    expect(ui.container.textContent).not.toContain("窗格已重新加载");
  });
});
