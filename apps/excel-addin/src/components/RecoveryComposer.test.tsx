// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";

import { RecoveryComposer, RecoveryPendingEntry } from "./RecoveryComposer";
import type { RecoveryComposerView } from "@excel-bro/core/recoveryComposer";

afterEach(cleanup);

function view(overrides: Partial<RecoveryComposerView> = {}): RecoveryComposerView {
  return {
    action: "reconcile",
    headline: "任务中断 · 还剩 1 步",
    lines: [
      "已写入：结果1!B2、结果2!B2",
      "尚未写入：结果3!B2",
      "先检查已写入结果，不会修改表格。",
      "检查完成后还需要你再点一次，才会执行剩下的步骤。"
    ],
    primaryLabel: "检查已写入结果",
    primaryTitle: "只读核对第 2 步的结果并补上完成记录：不修改任何单元格，也不会自动继续。",
    progress: "已完成 1 / 3 步 · 在途第 2 步",
    detail: "第 2 个动作可能在停止前已经写入，也可能没有；核对只读当前目标与写前快照。",
    note: "恢复只使用当前已授权工作表；续跑通过复核后执行剩余动作，再审阅结果。",
    ...overrides
  };
}

describe("输入框里的恢复操作区", () => {
  it("标题/正文/主操作都来自共享视图，不在这里另写一句", () => {
    render(<RecoveryComposer view={view()} busy={false} detailsOpen={false}
      onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} />);
    expect(screen.getByText("任务中断 · 还剩 1 步")).toBeTruthy();
    for (const line of view().lines) expect(screen.getByText(line)).toBeTruthy();
    const primary = screen.getByRole("button", { name: "检查已写入结果" });
    expect(primary.getAttribute("title")).toContain("不会自动继续");
  });

  it("同一时刻只有一颗主操作，且三颗按钮都不是提交按钮", () => {
    const { container } = render(
      <RecoveryComposer view={view()} busy={false} detailsOpen={false}
        onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} />
    );
    expect(container.querySelectorAll("[data-recovery-primary]")).toHaveLength(1);
    expect(screen.getAllByRole("button")).toHaveLength(3);
    for (const button of screen.getAllByRole("button")) {
      expect(button.getAttribute("type")).toBe("button");
    }
  });

  it("忙碌时主操作仍占位、只是不可点，并说明正在做什么", () => {
    const { container } = render(
      <RecoveryComposer view={view()} busy detailsOpen={false}
        onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} />
    );
    const primary = screen.getByRole("button", { name: "正在检查…" });
    expect((primary as HTMLButtonElement).disabled).toBe(true);
    expect(container.querySelector(".recovery-composer")).not.toBeNull();
  });

  it("点一次只调一次处理器（检查不自动续跑）", async () => {
    const onPrimary = vi.fn();
    const onNewRequest = vi.fn();
    render(<RecoveryComposer view={view()} busy={false} detailsOpen={false}
      onPrimary={onPrimary} onNewRequest={onNewRequest} onToggleDetails={() => {}} />);
    fireEvent.click(screen.getByRole("button", { name: "检查已写入结果" }));
    expect(onPrimary).toHaveBeenCalledTimes(1);
    expect(onNewRequest).not.toHaveBeenCalled();
  });

  it("「另提需求」是显式逃生口：只通知调用方，不发检查/继续", async () => {
    const onPrimary = vi.fn();
    const onNewRequest = vi.fn();
    render(<RecoveryComposer view={view()} busy={false} detailsOpen={false}
      onPrimary={onPrimary} onNewRequest={onNewRequest} onToggleDetails={() => {}} />);
    fireEvent.click(screen.getByRole("button", { name: "另提需求" }));
    expect(onNewRequest).toHaveBeenCalledTimes(1);
    expect(onPrimary).not.toHaveBeenCalled();
  });

  it("查看详情才展开技术记录：进度/结论/边界，不是又一张大卡", async () => {
    const onToggleDetails = vi.fn();
    const { rerender } = render(
      <RecoveryComposer view={view()} busy={false} detailsOpen={false}
        onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={onToggleDetails} />
    );
    expect(screen.queryByText(view().progress)).toBeNull();
    fireEvent.click(screen.getByRole("button", { name: "查看详情" }));
    expect(onToggleDetails).toHaveBeenCalledTimes(1);
    rerender(<RecoveryComposer view={view()} busy={false} detailsOpen
      onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={onToggleDetails} />);
    expect(screen.getByText(view().progress)).toBeTruthy();
    expect(screen.getByText(view().detail)).toBeTruthy();
    expect(screen.getByText(view().note)).toBeTruthy();
  });

  it("「清除恢复提示」只在详情里、且只在提供时出现；说明它不等于撤销", async () => {
    const onDismiss = vi.fn();
    const { rerender } = render(
      <RecoveryComposer view={view()} busy={false} detailsOpen
        onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} />
    );
    expect(screen.queryByRole("button", { name: "清除这条恢复提示" })).toBeNull();
    rerender(
      <RecoveryComposer view={view()} busy={false} detailsOpen
        dismiss={{ label: "清除这条恢复提示", hint: "只收起这条提示，不撤销已经写入的内容。" }}
        onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} onDismiss={onDismiss} />
    );
    const dismiss = screen.getByRole("button", { name: "清除这条恢复提示" });
    expect(screen.getByText("只收起这条提示，不撤销已经写入的内容。")).toBeTruthy();
    fireEvent.click(dismiss);
    expect(onDismiss).toHaveBeenCalledTimes(1);
  });

  it("撤销行**不收进详情**：始终可见（安全入口），点一次只调一次；阻断时只给一句话不给按钮", () => {
    const onUndo = vi.fn();
    const { rerender } = render(
      <RecoveryComposer view={view()} busy={false} detailsOpen={false}
        undo={{ state: "available", label: "↶ 撤销本次执行", detail: "当前可安全撤销", onUndo }}
        onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} />
    );
    // 收起时也画：撤销是安全入口，收进折叠详情等于把它藏起来（真机 B 现场就靠它兜底）。
    const undoButton = screen.getByRole("button", { name: "↶ 撤销本次执行" });
    fireEvent.click(undoButton);
    expect(onUndo).toHaveBeenCalledTimes(1);
    rerender(
      <RecoveryComposer view={view()} busy={false} detailsOpen
        undo={{ state: "blocked", detail: "目标已被改过，不能再撤销回退。" }}
        onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} />
    );
    expect(screen.getByText("目标已被改过，不能再撤销回退。")).toBeTruthy();
    expect(screen.queryByRole("button", { name: /撤销本次执行/ })).toBeNull();
  });

  it("技术诊断收在详情里，收起时不出现", () => {
    const diagnostics = <pre>{"verdict=reconcile code=ok"}</pre>;
    const { container, rerender } = render(
      <RecoveryComposer view={view()} busy={false} detailsOpen={false} diagnostics={diagnostics}
        onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} />
    );
    expect(container.textContent).not.toContain("verdict=reconcile");
    rerender(<RecoveryComposer view={view()} busy={false} detailsOpen diagnostics={diagnostics}
      onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} />);
    expect(container.textContent).toContain("技术诊断（仅开发）");
    expect(container.textContent).toContain("verdict=reconcile");
  });

  it("回车不会触发恢复操作（这里根本没有提交键位）", () => {
    /*
     * 恢复操作区**没有** textarea，也没有任何 keydown 处理：回车只表达"把这段话
     * 发出去"，而这里没有"这段话"。为此留一条反向用例 —— 将来有人给它加个
     * `onKeyDown` 图省事，这条会红。
     */
    const onPrimary = vi.fn();
    const { container } = render(
      <RecoveryComposer view={view()} busy={false} detailsOpen={false}
        onPrimary={onPrimary} onNewRequest={() => {}} onToggleDetails={() => {}} />
    );
    const section = container.querySelector(".recovery-composer")!;
    fireEvent.keyDown(section, { key: "Enter" });
    fireEvent.keyDown(screen.getByRole("button", { name: "检查已写入结果" }), { key: "Enter" });
    expect(onPrimary).not.toHaveBeenCalled();
    expect(container.querySelector("textarea")).toBeNull();
  });

  it("续跑态的主操作带剩余步数，标题说检查完成", () => {
    render(<RecoveryComposer
      view={view({
        action: "resume",
        headline: "检查完成，可以继续",
        lines: ["下一步：向结果3!B2写入公式。", "前 2 步不会重复执行。"],
        primaryLabel: "继续剩余 1 步",
        primaryTitle: "只执行剩下的 1 步；已完成的部分不会重写，结束后再审阅结果。"
      })}
      busy={false} detailsOpen={false}
      onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} />);
    expect(screen.getByRole("button", { name: "继续剩余 1 步" })).toBeTruthy();
    expect(screen.getByText("检查完成，可以继续")).toBeTruthy();
    expect(screen.getByText("下一步：向结果3!B2写入公式。")).toBeTruthy();
  });
});

describe("防连点：一次点击只派发一次，且不跨越「检查 → 继续」", () => {
  it("双击的第二下（detail>1）不派发；第一下照常派发一次", () => {
    const onPrimary = vi.fn();
    render(<RecoveryComposer view={view()} busy={false} detailsOpen={false}
      onPrimary={onPrimary} onNewRequest={() => {}} onToggleDetails={() => {}} />);
    const primary = screen.getByRole("button", { name: "检查已写入结果" });

    // 双击：浏览器对同一次双击派发 detail=1 与 detail=2 两下。
    fireEvent.click(primary, { detail: 1 });
    fireEvent.click(primary, { detail: 2 });
    expect(onPrimary).toHaveBeenCalledTimes(1);

    // 键盘激活给的是 detail=0（用户明确按下去），必须放行 —— 这是正向锚点。
    onPrimary.mockClear();
    fireEvent.click(primary, { detail: 0 });
    expect(onPrimary).toHaveBeenCalledTimes(1);
  });

  it("核对 → 继续 换的是**节点**，旧节点上的第二下没有落点", () => {
    const onPrimary = vi.fn();
    const { container, rerender } = render(
      <RecoveryComposer view={view()} busy={false} detailsOpen={false}
        onPrimary={onPrimary} onNewRequest={() => {}} onToggleDetails={() => {}} />
    );
    const before = container.querySelector("[data-recovery-primary]")!;
    expect(before.getAttribute("data-recovery-primary")).toBe("reconcile");

    // 核对完成后视图变成"继续剩余 N 步"。
    rerender(<RecoveryComposer
      view={view({
        action: "resume",
        headline: "检查完成，可以继续",
        lines: ["下一步：向结果3!B2写入公式。", "前 2 步不会重复执行。"],
        primaryLabel: "继续剩余 1 步"
      })}
      busy={false} detailsOpen={false}
      onPrimary={onPrimary} onNewRequest={() => {}} onToggleDetails={() => {}} />);
    const after = container.querySelector("[data-recovery-primary]")!;
    expect(after).not.toBe(before);
    expect(after.getAttribute("data-recovery-primary")).toBe("resume");
    /*
     * 这一条是"旧点击事件不许跨越检查→继续"的可测形态：
     * 节点已被替换 ⇒ 落在旧节点上的那一下不会派发；换成新节点后用户必须**再点一次**。
     */
    fireEvent.click(before, { detail: 1 });
    expect(onPrimary).not.toHaveBeenCalled();
    fireEvent.click(after, { detail: 1 });
    expect(onPrimary).toHaveBeenCalledTimes(1);
  });

  it("撤销按钮同样只认第一下（它也是一次真实宿主写入）", () => {
    const onUndo = vi.fn();
    render(<RecoveryComposer
      view={view({ action: "resume", primaryLabel: "继续剩余 1 步" })}
      busy={false} detailsOpen={false}
      undo={{ state: "available", detail: "撤销已写入的两步。", onUndo }}
      onPrimary={() => {}} onNewRequest={() => {}} onToggleDetails={() => {}} />);
    const undo = screen.getByRole("button", { name: /撤销本次执行/ });
    fireEvent.click(undo, { detail: 1 });
    fireEvent.click(undo, { detail: 2 });
    expect(onUndo).toHaveBeenCalledTimes(1);
  });
});

describe("另提需求之后的紧凑待办入口", () => {
  it("只把输入框要回来：按钮文案来自入口判据，不承诺撤销", async () => {
    const onReturn = vi.fn();
    render(<RecoveryPendingEntry busy={false} onReturn={onReturn}
      entry={{ action: "resume", label: "恢复待办 · 还剩 1 步", title: "回到恢复处理：继续剩下的 1 步。" }} />);
    const button = screen.getByRole("button", { name: "恢复待办 · 还剩 1 步" });
    expect(button.getAttribute("title")).toContain("回到恢复处理");
    fireEvent.click(button);
    expect(onReturn).toHaveBeenCalledTimes(1);
    expect(screen.queryByText(/撤销/)).toBeNull();
  });
});
