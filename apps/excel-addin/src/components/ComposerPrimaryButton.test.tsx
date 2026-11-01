// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";
import { flushSync } from "react-dom";
import { useState } from "react";
import { afterEach, expect, it, vi } from "vitest";
import {
  composerPrimaryAction,
  type ComposerPrimaryAction,
  type ComposerPrimaryInput
} from "@excel-bro/core/composerPrimaryAction";
import { ComposerPrimaryButton } from "./ComposerPrimaryButton";

afterEach(cleanup);

/**
 * 判据在共享层（`composerPrimaryAction`），组件只负责按它画。
 * 所以这里造数走真实派生函数，而不是手写一个"看起来像"的对象。
 */
const action = (extra: Partial<ComposerPrimaryInput> = {}): ComposerPrimaryAction =>
  composerPrimaryAction({
    hasInput: false,
    paused: false,
    taskRunning: false,
    controllerAlive: false,
    ...extra
  });

/*
 * 这一条是**真机双击**复现过的缺陷：点击 ■ 之后立刻变成 ▶，如果还是同一个 DOM 节点，
 * 浏览器那次点击的 click 计数（`event.detail`）会留到下一帧 —— 双击的第二下被当成
 * "点继续"，于是一次暂停连出一次续跑。
 *
 * 这里用 `flushSync` 把状态改掉再回来，与真机"点下去 → 立即重渲染"的时序一致。
 * （普通 `setState` 在 `act` 里是批量的，覆盖不到"同一拍换节点"这一面。）
 */
it("pause then continue uses two distinct DOM nodes even when state settles synchronously", () => {
  const invokes: string[] = [];
  function Harness() {
    const [running, setRunning] = useState(true);
    return (
      <ComposerPrimaryButton
        action={action({ taskRunning: running, controllerAlive: running, paused: !running })}
        onInvoke={(invoke) => {
          invokes.push(invoke);
          if (invoke === "pause") flushSync(() => setRunning(false));
        }}
      />
    );
  }
  render(<Harness />);
  const pause = screen.getByRole("button", { name: "暂停任务" });
  act(() => { pause.click(); });
  const play = screen.getByRole("button", { name: "继续任务" });
  expect(play).not.toBe(pause);
  expect(invokes).toEqual(["pause"]);
  // 双击的第二下（`detail === 2`）不产生第二个动作；单击照常。
  fireEvent.click(play, { detail: 2 });
  expect(invokes).toEqual(["pause"]);
  fireEvent.click(play, { detail: 1 });
  expect(invokes).toEqual(["pause", "resume"]);
  // 按钮永远不是 submit：在表单里按它不能顺带把输入框提交一遍。
  expect(play.getAttribute("type")).toBe("button");
});

it("does not permit a second action while the pause request is settling", () => {
  const invoke = vi.fn();
  render(<ComposerPrimaryButton action={action({ pausing: true })} onInvoke={invoke} />);
  const button = screen.getByRole("button", { name: "正在暂停" });
  expect(button.hasAttribute("disabled")).toBe(true);
  expect(button.getAttribute("aria-busy")).toBe("true");
  expect(button.textContent).toBe("…");
  fireEvent.click(button);
  expect(invoke).not.toHaveBeenCalled();
});

it("写入/回读期间画的是禁用的 … 与原因，而不是一颗能点的 ■", () => {
  const invoke = vi.fn();
  render(
    <ComposerPrimaryButton
      action={action({ taskRunning: true, controllerAlive: true, uninterruptible: true })}
      onInvoke={invoke}
    />
  );
  const button = screen.getByRole("button", { name: /写入与回读期间不能中断/ });
  expect(button.textContent).toBe("…");
  expect(button.hasAttribute("disabled")).toBe(true);
  // 原因不能只藏在 tooltip 里：`title` 与 `aria-label` 都得说。
  expect(button.getAttribute("title")).toContain("写入与回读期间不能中断");
  fireEvent.click(button);
  expect(invoke).not.toHaveBeenCalled();
});

it("现场不明时主按钮是「核对」而不是「继续」", () => {
  const invoke = vi.fn();
  render(
    <ComposerPrimaryButton
      action={action({ paused: true, needsReconcile: true })}
      onInvoke={invoke}
    />
  );
  expect(screen.queryByRole("button", { name: "继续任务" })).toBeNull();
  const button = screen.getByRole("button", { name: "核对在途动作" });
  expect(button.getAttribute("data-primary-action")).toBe("reconcile");
  fireEvent.click(button);
  expect(invoke).toHaveBeenCalledWith("reconcile");
});

it("宿主侧的额外禁用条件必须给出原因，不能只是变灰", () => {
  const invoke = vi.fn();
  render(
    <ComposerPrimaryButton
      action={action({ paused: true })}
      onInvoke={invoke}
      disabled
      disabledHint="等待审阅结果：可确认变更、放弃变更或打开双视图"
    />
  );
  const button = screen.getByRole("button", { name: "继续任务" });
  expect(button.hasAttribute("disabled")).toBe(true);
  expect(button.getAttribute("title")).toBe("等待审阅结果：可确认变更、放弃变更或打开双视图");
  fireEvent.click(button);
  expect(invoke).not.toHaveBeenCalled();
});

it("回答澄清题时按钮发的是回答，不是新需求", () => {
  const invoke = vi.fn();
  render(
    <ComposerPrimaryButton
      action={action({ answeringClarification: true, hasAnswerContent: true })}
      onInvoke={invoke}
    />
  );
  const button = screen.getByRole("button", { name: "发送回答" });
  fireEvent.click(button);
  expect(invoke).toHaveBeenCalledWith("answer");
});

it("没有选中/没有写内容时按钮禁用，并说清要先做什么", () => {
  const invoke = vi.fn();
  render(
    <ComposerPrimaryButton
      action={action({ answeringClarification: true, hasAnswerContent: false })}
      onInvoke={invoke}
    />
  );
  const button = screen.getByRole("button", { name: "发送回答" });
  expect(button.hasAttribute("disabled")).toBe(true);
  expect(button.getAttribute("title")).toContain("先选一项");
  fireEvent.click(button);
  expect(invoke).not.toHaveBeenCalled();
});
