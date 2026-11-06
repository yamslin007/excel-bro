import { describe, expect, it } from "vitest";
import {
  PRIMARY_PAUSE_UNINTERRUPTIBLE,
  PRIMARY_PAUSE_UNAVAILABLE,
  composerPrimaryAction,
  type ComposerPrimaryInput
} from "./composerPrimaryAction";

const base: ComposerPrimaryInput = {
  hasInput: false,
  paused: false,
  taskRunning: false,
  controllerAlive: false
};

describe("composerPrimaryAction 六态", () => {
  it("任务结束 / 待输入 ⇒ 发送", () => {
    expect(composerPrimaryAction(base).kind).toBe("send");
    expect(composerPrimaryAction({ ...base, hasInput: true }).kind).toBe("send");
  });

  it("已暂停且输入为空 ⇒ 继续", () => {
    const action = composerPrimaryAction({ ...base, paused: true });
    expect(action.kind).toBe("continue");
    expect(action.invoke).toBe("resume");
    expect(action.glyph).toBe("▶");
  });

  it("已暂停但有输入 ⇒ 发送补充要求，而不是继续", () => {
    const action = composerPrimaryAction({ ...base, paused: true, hasInput: true });
    expect(action.kind).toBe("send");
    // 判别点：这条链路必须走 send 而不是 resume —— 走 resume 就是把旧计划原样重放。
    expect(action.invoke).toBe("send");
    expect(action.title).toContain("重新判断");
    expect(action.title).toContain("不会重放");
  });

  it("运行中且可中止 ⇒ 暂停", () => {
    const action = composerPrimaryAction({ ...base, taskRunning: true, controllerAlive: true });
    expect(action.kind).toBe("pause");
    expect(action.glyph).toBe("■");
    expect(action.invoke).toBe("pause");
    expect(action.disabled).toBe(false);
  });

  it("运行中但没有可中止对象 ⇒ 不可暂停，且给出原因", () => {
    const action = composerPrimaryAction({ ...base, taskRunning: true });
    expect(action.glyph).toBe("…");
    expect(action.disabled).toBe(true);
    expect(action.invoke).toBeNull();
    expect(action.reason).toBe(PRIMARY_PAUSE_UNAVAILABLE);
  });

  it("写入/回读期间不画可点的 ■", () => {
    const action = composerPrimaryAction({
      ...base, taskRunning: true, controllerAlive: true, uninterruptible: true
    });
    expect(action.glyph).not.toBe("■");
    expect(action.disabled).toBe(true);
    expect(action.reason).toBe(PRIMARY_PAUSE_UNINTERRUPTIBLE);
  });

  it("暂停请求处理中 ⇒ 等待图标且禁用，重复点击不产生第二个请求", () => {
    const action = composerPrimaryAction({
      ...base, taskRunning: true, controllerAlive: true, pausing: true
    });
    expect(action.kind).toBe("pausing");
    expect(action.glyph).toBe("…");
    expect(action.busy).toBe(true);
    expect(action.disabled).toBe(true);
    // 关键：invoke 为 null ⇒ 连点也不会再发一次暂停请求。
    expect(action.invoke).toBeNull();
  });

  it("运行中有字 ⇒ 打断并补充（不是排队发送）", () => {
    const action = composerPrimaryAction({
      ...base, taskRunning: true, controllerAlive: true, hasInput: true
    });
    expect(action.kind).toBe("send-supplement");
    expect(action.label).toBe("打断并补充");
    expect(action.title).toContain("不是排队发送");
  });

  it("现场不明 ⇒ 核对，且与继续是两个不同的形态", () => {
    const action = composerPrimaryAction({ ...base, paused: true, needsReconcile: true });
    expect(action.kind).toBe("reconcile");
    expect(action.invoke).toBe("reconcile");
    expect(action.kind).not.toBe("continue");
    expect(action.label).not.toBe(composerPrimaryAction({ ...base, paused: true }).label);
  });

  it("核对优先于暂停：任务在跑时也不把核对塞进 ■ 里", () => {
    const action = composerPrimaryAction({
      ...base, taskRunning: true, controllerAlive: true, needsReconcile: true
    });
    expect(action.kind).toBe("pause");
  });
});

describe("composerPrimaryAction：回答澄清题时的发送", () => {
  it("有 pending 问题且回答有内容 ⇒ 发送回答（不是新需求）", () => {
    const action = composerPrimaryAction({
      ...base, hasInput: true, answeringClarification: true, hasAnswerContent: true
    });
    expect(action.kind).toBe("answer");
    expect(action.invoke).toBe("answer");
    expect(action.label).toBe("发送回答");
    expect(action.disabled).toBe(false);
  });

  it("回答没内容 ⇒ 禁用并说明，而不是让按钮消失", () => {
    const action = composerPrimaryAction({ ...base, answeringClarification: true });
    expect(action.kind).toBe("answer");
    expect(action.disabled).toBe(true);
    expect(action.invoke).toBeNull();
    expect(action.reason).toContain("请选择一项");
  });

  it("没有 pending 问题时同一句话走新需求，绝不走 answer", () => {
    // 反向锚点：把这一档写成"有字就 answer"，用户的普通需求会被当成旧问题的答案。
    const action = composerPrimaryAction({ ...base, hasInput: true });
    expect(action.kind).toBe("send");
    expect(action.invoke).toBe("send");
  });

  it("任务仍在跑时不抢暂停：回答要等停下来再发", () => {
    const action = composerPrimaryAction({
      ...base, taskRunning: true, controllerAlive: true,
      answeringClarification: true, hasAnswerContent: true
    });
    expect(action.kind).toBe("pause");
  });
});

describe("composerPrimaryAction：恢复卡的动作由右下角这颗按钮承接", () => {
  it("待核对 ⇒ 核对（只读语义），且不是继续", () => {
    const action = composerPrimaryAction({ ...base, needsReconcile: true });
    expect(action.kind).toBe("reconcile");
    expect(action.invoke).toBe("reconcile");
    expect(action.reason).toContain("先核对");
    expect(action.invoke).not.toBe("resume");
  });

  it("核对完、还差几步 ⇒ 继续剩余动作（文案不许写成「继续任务」）", () => {
    const action = composerPrimaryAction({ ...base, recoveryResume: true });
    expect(action.kind).toBe("continue");
    expect(action.invoke).toBe("resume");
    expect(action.label).toBe("继续剩余动作");
    expect(action.title).toContain("剩余步骤");
    // 与"继续分析"必须能分辨：同一个 ▶ 下是两件不同的事。
    expect(action.label).not.toBe(composerPrimaryAction({ ...base, paused: true }).label);
  });

  it("核对优先于恢复续跑：两个事实同时成立时先核对，不许跳过", () => {
    const action = composerPrimaryAction({ ...base, needsReconcile: true, recoveryResume: true });
    expect(action.kind).toBe("reconcile");
    expect(action.invoke).toBe("reconcile");
  });

  it("输入框有字 ⇒ 永远发送；既不核对也不续跑（不许吞掉用户正在写的话）", () => {
    for (const input of [
      { ...base, hasInput: true, needsReconcile: true },
      { ...base, hasInput: true, recoveryResume: true },
      { ...base, hasInput: true, needsReconcile: true, recoveryResume: true, paused: true }
    ]) {
      const action = composerPrimaryAction(input);
      expect(action.kind).toBe("send");
      expect(action.invoke).toBe("send");
    }
  });

  it("任务仍在跑时不把核对/续跑塞进暂停按钮", () => {
    for (const extra of [{ needsReconcile: true }, { recoveryResume: true }]) {
      const action = composerPrimaryAction({
        ...base, taskRunning: true, controllerAlive: true, ...extra
      });
      expect(action.kind).toBe("pause");
      expect(action.invoke).toBe("pause");
    }
  });
});

describe("composerPrimaryAction 的可辨识性", () => {
  it("每个形态都有非空的 aria-label 与 tooltip，且暂停/继续/发送各不相同", () => {
    const kinds = [
      composerPrimaryAction({ ...base }),
      composerPrimaryAction({ ...base, hasInput: true }),
      composerPrimaryAction({ ...base, paused: true }),
      composerPrimaryAction({ ...base, paused: true, needsReconcile: true }),
      composerPrimaryAction({ ...base, taskRunning: true, controllerAlive: true }),
      composerPrimaryAction({ ...base, taskRunning: true, controllerAlive: true, pausing: true })
    ];
    for (const action of kinds) {
      expect(action.label.trim().length).toBeGreaterThan(0);
      expect(action.title.trim().length).toBeGreaterThan(0);
    }
    expect(kinds[2].label).toBe("继续任务");
    expect(kinds[4].label).toBe("暂停任务");
    expect(kinds[5].label).toBe("正在暂停");
  });

  it("禁用一定伴随原因，可点一定没有原因", () => {
    const cases: ComposerPrimaryInput[] = [
      base,
      { ...base, hasInput: true },
      { ...base, paused: true },
      { ...base, taskRunning: true },
      { ...base, taskRunning: true, controllerAlive: true },
      { ...base, taskRunning: true, controllerAlive: true, uninterruptible: true },
      { ...base, taskRunning: true, controllerAlive: true, pausing: true }
    ];
    for (const input of cases) {
      const action = composerPrimaryAction(input);
      if (action.disabled) {
        // pausing 是"正在发生"，它不需要解释"为什么不能点"；其余禁用必须给原因。
        if (action.kind !== "pausing") expect(action.reason).not.toBeNull();
        expect(action.invoke).toBeNull();
      } else {
        expect(action.reason).toBeNull();
        expect(action.invoke).not.toBeNull();
      }
    }
  });
});


describe("running draft safety", () => {
  it.each([false, true])("write/readback blocks dispatch even with draft=%s", (hasInput) => {
    const action = composerPrimaryAction({hasInput, taskRunning:true, controllerAlive:true, uninterruptible:true});
    expect(action.disabled).toBe(true); expect(action.invoke).toBeNull();
  });
  it("a draft cannot restart a task whose controller is already gone", () => {
    const action = composerPrimaryAction({hasInput:true, taskRunning:true, controllerAlive:false});
    expect(action.disabled).toBe(true); expect(action.invoke).toBeNull();
  });
});
