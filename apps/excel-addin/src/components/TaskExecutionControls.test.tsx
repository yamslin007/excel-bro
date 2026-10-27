// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";
import { flushSync } from "react-dom";
import { useState } from "react";
import { afterEach, expect, it, vi } from "vitest";
import { pauseButtonState } from "@excel-bro/core/pauseControl";
import { TaskExecutionControls } from "./TaskExecutionControls";
afterEach(cleanup);
const state = (running: boolean) => pauseButtonState({ hasInput: false, taskRunning: running, controllerAlive: running, paused: !running });

it("a native pause click cannot become a submit or a resume after a synchronous render", () => {
  const resume = vi.fn(); const submit = vi.fn();
  function Harness() {
    const [running, setRunning] = useState(true);
    return <form onSubmit={event => { event.preventDefault(); submit(); }}>
      <TaskExecutionControls state={state(running)} pausing={false}
        onPause={() => flushSync(() => setRunning(false))} onResume={resume} />
      <button type="submit" disabled>发送</button>
    </form>;
  }
  render(<Harness />);
  const pause = screen.getByRole("button", { name: "暂停任务" });
  act(() => pause.click());
  const play = screen.getByRole("button", { name: "继续任务" });
  expect(play).not.toBe(pause);
  expect(play.getAttribute("type")).toBe("button");
  expect(submit).not.toHaveBeenCalled(); expect(resume).not.toHaveBeenCalled();
  fireEvent.click(play, { detail: 2 });
  expect(resume).not.toHaveBeenCalled();
  fireEvent.click(play, { detail: 1 });
  expect(resume).toHaveBeenCalledOnce(); expect(submit).not.toHaveBeenCalled();
});
it("does not permit resume while pause is settling", () => {
  const resume = vi.fn();
  render(<TaskExecutionControls state={state(false)} pausing onPause={vi.fn()} onResume={resume} />);
  const button = screen.getByRole("button", { name: "正在暂停" });
  expect(button.hasAttribute("disabled")).toBe(true);
  fireEvent.click(button);
  expect(resume).not.toHaveBeenCalled();
});

it("reviewing suppresses historical resume without consuming the independent paused task", () => {
  const resume = vi.fn();
  const view = render(<TaskExecutionControls state={state(false)} reviewing pausing={false} onPause={vi.fn()} onResume={resume} />);
  expect(screen.getByText("等待审阅结果")).toBeTruthy();
  expect(screen.queryByRole("button", { name: "继续任务" })).toBeNull();
  view.rerender(<TaskExecutionControls state={state(false)} pausing={false} onPause={vi.fn()} onResume={resume} />);
  fireEvent.click(screen.getByRole("button", { name: "继续任务" }));
  expect(resume).toHaveBeenCalledOnce();
});
