import { describe, expect, it, vi } from "vitest";
import {
  PAUSE_UNAVAILABLE_REASON,
  PAUSE_UNINTERRUPTIBLE_REASON,
  PauseTimeoutError,
  controllerOwnerOf,
  executionPauseState,
  newExecutionId,
  pauseButtonState,
  pauseDiagnosticSnapshot,
  planExecutionTeardown,
  recordPauseDiagnostic,
  resetPauseDiagnostics,
  tagControllerOwner,
  tailOf,
  withPauseTimeout
} from "./pauseControl";

describe("暂停按钮判据（三件事分开）", () => {
  it("只有真的存在可中止对象时才显示 ■", () => {
    const state = pauseButtonState({
      hasInput: false, paused: false, taskRunning: true, controllerAlive: true
    });
    expect(state.kind).toBe("pause");
    expect(state.glyph).toBe("■");
    expect(state.canPause).toBe(true);
  });

  it("有任务在跑但没有可中止对象时不许显示 ■，并给出原因", () => {
    const state = pauseButtonState({
      hasInput: false, paused: false, taskRunning: true, controllerAlive: false
    });
    expect(state.kind).toBe("unavailable");
    expect(state.glyph).not.toBe("■");
    expect(state.canPause).toBe(false);
    expect(state.reason).toBe(PAUSE_UNAVAILABLE_REASON);
  });

  it("写入/回读期间是明确的不可中断，不是可暂停", () => {
    const state = pauseButtonState({
      hasInput: false, paused: false, taskRunning: true, controllerAlive: true, uninterruptible: true
    });
    expect(state.kind).toBe("unavailable");
    expect(state.reason).toBe(PAUSE_UNINTERRUPTIBLE_REASON);
    expect(state.canPause).toBe(false);
  });

  it("暂停且输入为空显示 ▶，点它是继续", () => {
    const state = pauseButtonState({
      hasInput: false, paused: true, taskRunning: false, controllerAlive: false
    });
    expect(state.kind).toBe("continue");
    expect(state.glyph).toBe("▶");
    expect(state.label).toBe("继续任务");
  });

  it("输入了新内容时按钮回到发送，不被暂停态劫持", () => {
    expect(pauseButtonState({
      hasInput: true, paused: true, taskRunning: false, controllerAlive: false
    }).kind).toBe("send");
    expect(pauseButtonState({
      hasInput: true, paused: false, taskRunning: true, controllerAlive: true
    }).kind).toBe("send");
  });

  it("有输入且有任务在跑时，那个 ↑ 是「打断并补充」而不再是「发送」", () => {
    const state = pauseButtonState({
      hasInput: true, paused: false, taskRunning: true, controllerAlive: true
    });
    expect(state.kind).toBe("send");
    expect(state.label).toBe("打断并补充");
    expect(state.canPause).toBe(false);
    expect(pauseButtonState({
      hasInput: true, paused: false, taskRunning: false, controllerAlive: false
    }).label).toBe("发送");
  });

  it("没有任何任务时是发送，不是暂停", () => {
    const state = pauseButtonState({
      hasInput: false, paused: false, taskRunning: false, controllerAlive: false
    });
    expect(state.kind).toBe("send");
  });

  it("过程卡上的「停止」不受输入影响：运行中打了字，卡上仍然停得下来", () => {
    const base = { taskRunning: true, controllerAlive: true };
    expect(executionPauseState(base).canPause).toBe(true);
    expect(executionPauseState(base).glyph).toBe("■");
    // 同一个现场，输入框那个按钮变成"打断并补充"，但不是把卡上的入口也收走。
    expect(pauseButtonState({ ...base, hasInput: true, paused: false }).kind).toBe("send");
    expect(executionPauseState({ taskRunning: true, controllerAlive: false }).canPause).toBe(false);
  });

  it("任何输入下画出的 ■ 都必须自带 canPause=true（不许有死的暂停按钮）", () => {
    const cases = [
      { hasInput: false, paused: false, taskRunning: true, controllerAlive: true },
      { hasInput: false, paused: false, taskRunning: true, controllerAlive: false },
      { hasInput: false, paused: false, taskRunning: false, controllerAlive: false },
      { hasInput: true, paused: false, taskRunning: true, controllerAlive: true },
      { hasInput: false, paused: true, taskRunning: false, controllerAlive: false }
    ];
    for (const input of cases) {
      const state = pauseButtonState(input);
      if (state.glyph === "■") expect(state.canPause).toBe(true);
    }
  });
});

describe("收尾按资源分项归属", () => {
  const base = { owner: "exec-A", activitySettled: false, hasActivity: true };

  it("自己开的活动自己收尾，哪怕 controller 已被别人接走", () => {
    const plan = planExecutionTeardown({
      ...base, controllerOwner: "inspect-B", bridgeOwner: "inspect-B"
    });
    expect(plan.settleActivity).toBe(true);
    expect(plan.clearController).toBe(false);
    expect(plan.endBridge).toBe(false);
  });

  it("仍持有 controller 时清 controller", () => {
    expect(planExecutionTeardown({
      ...base, controllerOwner: "exec-A", bridgeOwner: null
    }).clearController).toBe(true);
  });

  it("只关自己立的桥", () => {
    expect(planExecutionTeardown({
      ...base, controllerOwner: "exec-A", bridgeOwner: "exec-A"
    }).endBridge).toBe(true);
    expect(planExecutionTeardown({
      ...base, controllerOwner: "exec-A", bridgeOwner: "exec-B"
    }).endBridge).toBe(false);
  });

  it("旧执行不得动新执行的活动/controller/桥", () => {
    const plan = planExecutionTeardown({
      owner: "exec-old",
      controllerOwner: "exec-new",
      bridgeOwner: "exec-new",
      hasActivity: false,
      activitySettled: true
    });
    expect(plan).toEqual({ settleActivity: false, clearController: false, endBridge: false });
  });

  it("重复收尾安全：已收过的活动不再收第二遍", () => {
    const plan = planExecutionTeardown({
      owner: "exec-A", controllerOwner: null, bridgeOwner: null,
      hasActivity: true, activitySettled: true
    });
    expect(plan.settleActivity).toBe(false);
  });
});

describe("执行身份跟着 controller 走", () => {
  it("身份从对象上读回来，不靠调用点自己记变量", () => {
    const controller = new AbortController();
    expect(controllerOwnerOf(controller)).toBeNull();
    tagControllerOwner(controller, "exec-7");
    expect(controllerOwnerOf(controller)).toBe("exec-7");
    expect(newExecutionId()).not.toBe(newExecutionId());
  });

  it("核对把 ref 换掉之后，旧执行看到的是新主人（于是它不抢 controller，但仍该收自己的活动）", () => {
    const resume = new AbortController();
    tagControllerOwner(resume, "exec-resume");
    const inspection = new AbortController();
    tagControllerOwner(inspection, "exec-inspect");
    // 核对接管：ref 现在指向 inspection。
    const plan = planExecutionTeardown({
      owner: "exec-resume",
      controllerOwner: controllerOwnerOf(inspection),
      bridgeOwner: "exec-resume",
      hasActivity: true,
      activitySettled: true
    });
    expect(plan.clearController).toBe(false);
    // 但桥是本次执行立的，必须由它收 —— 旧实现里这两件事被同一个布尔绑死了。
    expect(plan.endBridge).toBe(true);
  });
});

describe("暂停请求与核对都要有上界", () => {
  it("超时给出可识别的错误，而不是永远挂着", async () => {
    vi.useFakeTimers();
    try {
      const pending = withPauseTimeout(new Promise<never>(() => {}), 50, "暂停请求");
      const assertion = expect(pending).rejects.toBeInstanceOf(PauseTimeoutError);
      await vi.advanceTimersByTimeAsync(60);
      await assertion;
    } finally {
      vi.useRealTimers();
    }
  });

  it("按时返回的结果不受超时影响", async () => {
    await expect(withPauseTimeout(Promise.resolve("ok"), 1000)).resolves.toBe("ok");
  });

  it("底层失败照原样抛出，不被超时包装成成功", async () => {
    await expect(withPauseTimeout(Promise.reject(new Error("boom")), 1000)).rejects.toThrow("boom");
  });
});

describe("诊断记录不许带正文", () => {
  it("只收结构化字段，turnId 只留尾部", () => {
    resetPauseDiagnostics();
    recordPauseDiagnostic({
      kind: "pause-click",
      source: "composer",
      executionId: "turn-1",
      turnIdTail: tailOf("turn-abcdef123456"),
      controller: "owned"
    });
    const [event] = pauseDiagnosticSnapshot();
    expect(event.kind).toBe("pause-click");
    expect(event.turnIdTail).toBe("123456");
    expect(Object.keys(event).sort()).toEqual(
      ["at", "controller", "executionId", "kind", "source", "turnIdTail"].sort()
    );
  });

  it("环形缓冲有上限，不会无限增长", () => {
    resetPauseDiagnostics();
    for (let index = 0; index < 60; index += 1) {
      recordPauseDiagnostic({ kind: "pause-request", executionId: `e-${index}` });
    }
    const snapshot = pauseDiagnosticSnapshot();
    expect(snapshot).toHaveLength(40);
    expect(snapshot[snapshot.length - 1].executionId).toBe("e-59");
  });
});
