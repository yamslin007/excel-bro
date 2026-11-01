import { describe, expect, it } from "vitest";
import {
  PROCESS_PANEL_LIVE_EXPANDED_HEIGHT,
  PROCESS_PANEL_LIVE_HEIGHT,
  PROCESS_PANEL_SETTLED_EXPANDED_HEIGHT,
  PROCESS_PANEL_SETTLED_HEIGHT,
  foldedProcessLine,
  isCompactMessageProcess,
  formatProcessDuration,
  processPanelBodyMaxHeight,
  processPanelExpanded,
  processPanelHasOverflow
} from "./processPanel";

describe("单步消息理解的简洁展示", () => {
  const step = { id: "message-intent", status: "completed", phase: "intent" };
  it("只压缩完成的单步理解，异常、多步和工具证据保留", () => {
    expect(isCompactMessageProcess({ outcome: "completed", steps: [step] })).toBe(true);
    for (const outcome of ["failed", "interrupted", "awaiting_input", undefined]) {
      expect(isCompactMessageProcess({ outcome, steps: [step] })).toBe(false);
    }
    expect(isCompactMessageProcess({ outcome: "completed", steps: [step, step] })).toBe(false);
    expect(isCompactMessageProcess({ outcome: "completed", steps: [{ ...step, toolName: "read_range" }] })).toBe(false);
    expect(isCompactMessageProcess({ outcome: "completed", steps: [{ ...step, facts: [{}] }] })).toBe(false);
  });
});

describe("processPanelExpanded", () => {
  it("进行中默认展开，正常结束后自动收起", () => {
    expect(processPanelExpanded({ live: true, outcome: null, userChoice: null })).toBe(true);
    expect(processPanelExpanded({ live: false, outcome: "completed", userChoice: null })).toBe(false);
  });

  it("失败/中断/取消/等待回答都不自动收起", () => {
    // 反向锚点：把这条判据写成 `!live` 就会让"执行失败"只剩一行标题，
    // 用户必须自己想到去展开才知道停在哪一步。
    for (const outcome of ["failed", "interrupted", "cancelled", "warning", "awaiting_input"]) {
      expect(processPanelExpanded({ live: false, outcome, userChoice: null })).toBe(true);
    }
  });

  it("用户表过态就压过一切自动收起", () => {
    expect(processPanelExpanded({ live: true, outcome: null, userChoice: "collapsed" })).toBe(false);
    expect(processPanelExpanded({ live: false, outcome: "completed", userChoice: "expanded" })).toBe(true);
    // 同一条路径第二次渲染（用户展开后活动又变成进行中）也不许被拉回收起。
    expect(processPanelExpanded({ live: true, outcome: "failed", userChoice: "expanded" })).toBe(true);
  });
});

describe("processPanelBodyMaxHeight / processPanelHasOverflow", () => {
  it("进行中默认就是 160px，不靠用户先点一下", () => {
    expect(PROCESS_PANEL_LIVE_HEIGHT).toBe(160);
    expect(processPanelBodyMaxHeight({ live: true, enlarged: false })).toBe(160);
  });

  it("用户要求看更多时才放宽，且必须真的更高", () => {
    expect(processPanelBodyMaxHeight({ live: true, enlarged: true }))
      .toBe(PROCESS_PANEL_LIVE_EXPANDED_HEIGHT);
    expect(PROCESS_PANEL_LIVE_EXPANDED_HEIGHT).toBeGreaterThan(PROCESS_PANEL_LIVE_HEIGHT);
  });

  it("结束后的展开比进行中的展开更宽松", () => {
    expect(processPanelBodyMaxHeight({ live: false, enlarged: false }))
      .toBe(PROCESS_PANEL_SETTLED_HEIGHT);
    expect(processPanelBodyMaxHeight({ live: false, enlarged: true }))
      .toBe(PROCESS_PANEL_SETTLED_EXPANDED_HEIGHT);
    expect(PROCESS_PANEL_SETTLED_EXPANDED_HEIGHT)
      .toBeGreaterThan(PROCESS_PANEL_LIVE_EXPANDED_HEIGHT);
  });

  it("超出限高才需要给出收起/跳最新入口", () => {
    // 反向锚点：把阈值写成常量 480，160px 的框里塞 300px 内容就不会出入口。
    expect(processPanelHasOverflow({ live: true, enlarged: false, contentHeight: 300 })).toBe(true);
    expect(processPanelHasOverflow({ live: true, enlarged: false, contentHeight: 120 })).toBe(false);
    expect(processPanelHasOverflow({ live: true, enlarged: true, contentHeight: 300 })).toBe(false);
  });
});

describe("formatProcessDuration", () => {
  it("毫秒 → 秒 → 分", () => {
    expect(formatProcessDuration(320)).toBe("320ms");
    expect(formatProcessDuration(8_300)).toBe("8.3s");
    expect(formatProcessDuration(65_000)).toBe("1分5秒");
    expect(formatProcessDuration(120_000)).toBe("2分");
  });

  it("负数按 0 处理，不产出 -1.0s 这种文案", () => {
    expect(formatProcessDuration(-5)).toBe("0ms");
  });
});

describe("foldedProcessLine", () => {
  it("收起后一行 = 阶段结论 + 耗时", () => {
    expect(foldedProcessLine({ conclusion: "已完成", activeMs: 8_300 })).toBe("已完成 · 8.3s");
  });

  it("空结论有兜底，不产出只有耗时的一行", () => {
    expect(foldedProcessLine({ conclusion: "   ", activeMs: 1_000 })).toBe("已结束 · 1.0s");
  });
});
