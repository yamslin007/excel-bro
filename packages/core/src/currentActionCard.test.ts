import { describe, expect, it } from "vitest";
import {
  CURRENT_ACTION_PRIORITY,
  currentActionLabel,
  currentActionOtherCount,
  currentActionShortLabel,
  currentActionSwitchTargets,
  selectCurrentAction,
  type CurrentActionInput
} from "./currentActionCard";

const base: CurrentActionInput = {
  review: false,
  recovery: false,
  clarification: false,
  running: false
};

describe("selectCurrentAction", () => {
  it("没有待处理事项时是 none，而且不给任何标签", () => {
    const selection = selectCurrentAction(base);
    expect(selection.kind).toBe("none");
    expect(selection.others).toEqual([]);
    expect(currentActionLabel("none")).toBe("");
  });

  /*
   * 优先级不是"谁先出现谁第一"，而是**按代价**：已经写进工作簿的东西在等你确认，
   * 放着不管的代价最大；核对是只读的，可以等审阅收尾。
   */
  it("审阅压过恢复、澄清与运行", () => {
    const selection = selectCurrentAction({
      review: true, recovery: true, clarification: true, running: true
    });
    expect(selection.kind).toBe("review");
    // 其余三件都要如实列出来，不隐藏。
    expect(selection.others).toEqual(["recovery", "clarification", "runtime"]);
    expect(currentActionOtherCount(selection)).toBe(3);
  });

  it("现场不明（恢复）压过澄清，即使澄清先发生", () => {
    const selection = selectCurrentAction({ ...base, recovery: true, clarification: true });
    expect(selection.kind).toBe("recovery");
    expect(selection.others).toEqual(["clarification"]);
  });

  it("只是「在跑」时不会在固定区域再占一张卡", () => {
    const selection = selectCurrentAction({ ...base, running: true });
    // 判据仍然给出 runtime，但它是**最后**一档：有正式事项时一定排在后面。
    expect(selection.kind).toBe("runtime");
    expect(CURRENT_ACTION_PRIORITY.at(-1)).toBe("runtime");

    const withClarification = selectCurrentAction({ ...base, running: true, clarification: true });
    expect(withClarification.kind).toBe("clarification");
    expect(withClarification.others).toEqual(["runtime"]);
  });

  it("单件事时 others 为空（界面因此不画「还有 N 项」）", () => {
    for (const key of ["review", "recovery", "clarification", "running"] as const) {
      const selection = selectCurrentAction({ ...base, [key]: true });
      expect(selection.others, key).toEqual([]);
      expect(currentActionOtherCount(selection)).toBe(0);
    }
  });

  /*
   * 判据顺序必须与"理由"一致：改顺序等于改产品决定，`review` 不能被排到后面。
   */
  it("判据顺序与优先级表逐项一致", () => {
    expect(CURRENT_ACTION_PRIORITY).toEqual(["review", "recovery", "clarification", "runtime"]);
  });

  it("每个非空形态都有面向用户的一句话", () => {
    for (const kind of ["review", "recovery", "clarification", "runtime"] as const) {
      expect(currentActionLabel(kind).length).toBeGreaterThan(0);
      // 短名只用在紧凑提示里，但同样不许为空 —— 否则"还有 1 项："后面什么都没有。
      expect(currentActionShortLabel(kind).length).toBeGreaterThan(0);
    }
  });
});

describe("显式切换：优先级不许永久遮住另一条任务", () => {
  const allFour: CurrentActionInput = {
    review: true, recovery: true, clarification: true, running: true
  };

  it("显式选中较低优先级的那件时，它成为主视图，其余（含原来的第一件）都还在", () => {
    const selection = selectCurrentAction(allFour, "clarification");
    expect(selection.kind).toBe("clarification");
    // 反向锚点：被换下来的 review 不能消失 —— 消失就等于"切过去以后原来的事没人管"。
    expect(selection.others).toContain("review");
    expect(selection.others).toContain("recovery");
    expect(currentActionOtherCount(selection)).toBe(3);
  });

  it("切过去的那件事已经不在了 ⇒ 退回默认顺序，不留一张空卡", () => {
    const selection = selectCurrentAction({ ...base, review: true }, "recovery");
    expect(selection.kind).toBe("review");
    expect(selection.others).toEqual([]);
  });

  it("切换候选与 others 同源：提示里列几项，就能切过去几项", () => {
    const selection = selectCurrentAction(allFour, "recovery");
    const targets = currentActionSwitchTargets(selection);
    expect(targets.map((target) => target.kind)).toEqual(selection.others);
    for (const target of targets) {
      expect(target.label.length).toBeGreaterThan(0);
      expect(target.short.length).toBeGreaterThan(0);
      // 每一项切过去都真的能成为主视图（不是画了一个点不动的入口）。
      expect(selectCurrentAction(allFour, target.kind).kind).toBe(target.kind);
    }
  });

  it("多个恢复候选不会被自动挑一个：判据只吃一个布尔", () => {
    // 这条是"不许替用户选"的锚点：`recovery` 是单一事实，选择在
    // `selectExecutionRecoveryEntry`（按 executionId 取最后一条），不在这里。
    const selection = selectCurrentAction({ ...base, recovery: true });
    expect(selection.kind).toBe("recovery");
    expect(Object.keys(selection).sort()).toEqual(["kind", "others"]);
  });
});
