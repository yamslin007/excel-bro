// @vitest-environment jsdom
import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, expect, it, vi } from "vitest";
import {
  composerPrimaryAction,
  type ComposerPrimaryInput
} from "@excel-bro/core/composerPrimaryAction";
import { TaskExecutionNotice, taskExecutionNoticeText } from "./TaskExecutionNotice";

afterEach(cleanup);

const action = (extra: Partial<ComposerPrimaryInput> = {}) =>
  composerPrimaryAction({
    hasInput: false,
    paused: false,
    taskRunning: false,
    controllerAlive: false,
    ...extra
  });

/*
 * 这一行**不许有按钮**。
 *
 * 它以前是"运行控制条"：自己长了一颗 ■/▶，和输入框右下角那颗主按钮抢同一个动作。
 * 同一个阶段两个入口、其中一个还不受输入框内容影响 —— "我只是想补充一句话"会被
 * 当成"点继续"。现在动作全在右下角，这里只说话。
 */
it("does not render any button, even while a task is running", () => {
  const { container } = render(
    <TaskExecutionNotice action={action({ taskRunning: true, controllerAlive: true })} />
  );
  expect(container.querySelectorAll("button")).toHaveLength(0);
  expect(screen.queryByRole("button", { name: "暂停任务" })).toBeNull();
  expect(screen.queryByRole("button", { name: "继续任务" })).toBeNull();
});

it("reviewing says it is waiting for review and offers no resume", () => {
  render(<TaskExecutionNotice action={action({ paused: true })} reviewing />);
  expect(screen.getByText("等待审阅结果")).toBeTruthy();
  expect(screen.queryByRole("button", { name: "继续任务" })).toBeNull();
});

/*
 * 写入/回读期间不能中断：原因必须是**看得见的文字**，不能只挂在 tooltip 上
 * （触屏与键盘用户看不到 tooltip）。按钮本身在右下角（见 `ComposerPrimaryButton`）。
 */
it("写入/回读期间把不能中断的原因露在外面", () => {
  render(
    <TaskExecutionNotice
      action={action({ taskRunning: true, controllerAlive: true, uninterruptible: true })}
    />
  );
  expect(screen.getByText(/写入与回读期间不能中断/)).toBeTruthy();
  expect(screen.getByRole("status").getAttribute("data-notice")).toBe("pause");
});

it("pausing 与真正可继续是两句不同的话", () => {
  const { rerender } = render(
    <TaskExecutionNotice action={action({ taskRunning: true, pausing: true })} />
  );
  expect(screen.getByRole("status").textContent).toBe("正在停止当前处理…");

  rerender(<TaskExecutionNotice action={action({ paused: true })} />);
  expect(screen.getByRole("status").textContent).toBe("已暂停，可从右下角继续");
});

/*
 * "已暂停"不是随便说的：`continue` 形态的前提是 `paused` 为真，而 `paused` 只由
 * "存在待续跑消息（真检查点）"给出。仅仅 abort 掉一次请求**不会**走到这里 ——
 * 那条路径由下面的用例守住：`taskRunning: false, paused: false` 什么都不说。
 */
it("没有可恢复状态时不冒充「已暂停」", () => {
  const { container, rerender } = render(<TaskExecutionNotice action={action()} />);
  expect(container.firstChild).toBeNull();
  rerender(<TaskExecutionNotice action={action({ hasInput: true })} />);
  expect(container.firstChild).toBeNull();
  rerender(
    <TaskExecutionNotice
      action={action({ paused: true, needsReconcile: true })}
    />
  );
  // 现场不明 ⇒ 是"去核对"，不是"已暂停"。这里不说话，说明在恢复卡上。
  expect(container.firstChild).toBeNull();
});

it('父容器只登记实际有文字的任务状态，普通聊天不产生空待办', () => {
  for (const input of [{}, {hasInput:true}, {taskRunning:true,controllerAlive:true}, {needsReconcile:true}]) {
    const current = action(input);
    expect(taskExecutionNoticeText(current)).toBeNull();
    const {container,unmount}=render(<TaskExecutionNotice action={current}/>);
    expect(container.firstChild).toBeNull();unmount();
  }
  expect(taskExecutionNoticeText(action({paused:true}))).toContain('已暂停');
  expect(taskExecutionNoticeText(action({pausing:true,taskRunning:true}))).toContain('正在停止');
  expect(taskExecutionNoticeText(action(),true)).toBe('等待审阅结果');
});
