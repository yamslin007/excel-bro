import { describe, expect, it } from "vitest";
import {
  canArmResultRefFault,
  canCancelResultRefFault,
  describeResultRefFault,
  resultRefFaultTarget,
  resultRefFaultVisible,
  type ResultRefFaultPhase,
  type ResultRefFaultStatus,
  type ResultRefFaultTarget
} from "./resultRefFault";

const TARGET: ResultRefFaultTarget = {
  hostId: "excel-office",
  workbookName: "测试.xlsx",
  workbookIdentity: "url:https://example/测试.xlsx"
};

function statusOf(
  phase: ResultRefFaultPhase,
  overrides: Partial<ResultRefFaultStatus> = {}
): ResultRefFaultStatus {
  return {
    enabled: phase !== "off",
    phase,
    expiresInSeconds: phase === "armed" || phase === "bound" ? 590 : 0,
    taskBound: phase === "bound",
    refBound: phase === "bound",
    ...overrides
  };
}

const PHASES: ResultRefFaultPhase[] = ["off", "idle", "armed", "bound", "consumed", "expired"];

describe("入口可见性（两端读同一份判据）", () => {
  it("只有显式开发模式才可见", () => {
    expect(resultRefFaultVisible(true)).toBe(true);
    expect(resultRefFaultVisible(false)).toBe(false);
  });
});

describe("启用目标", () => {
  it("工作簿名缺失时不给目标 —— 不退化成全局开关", () => {
    // 没有名字就无从绑定"这本工作簿的下一次任务"；给一个空目标出去，
    // 服务端会按 fallback 身份认到一本叫空串的簿上，等于谁都能被它命中。
    expect(resultRefFaultTarget("excel-office", "", "url:x")).toBeNull();
    expect(resultRefFaultTarget("excel-office", "   ", "url:x")).toBeNull();
    expect(resultRefFaultTarget("excel-office", null, "url:x")).toBeNull();
    expect(resultRefFaultTarget("excel-office", undefined, undefined)).toBeNull();
  });

  it("请求体只有宿主与工作簿三项，不含任务或引用", () => {
    const target = resultRefFaultTarget("excel-office", "  测试.xlsx  ", " url:x ");
    expect(target).toEqual({
      hostId: "excel-office",
      workbookName: "测试.xlsx",
      workbookIdentity: "url:x"
    });
    // 键集合钉死：多一个 taskId/resultRef 就等于把"让哪个引用失效"交给调用方。
    expect(Object.keys(target!).sort()).toEqual([
      "hostId",
      "workbookIdentity",
      "workbookName"
    ]);
  });

  it("身份可以缺省（服务端按工作簿名兜底），但名字不能", () => {
    expect(resultRefFaultTarget("wps-et", "影院.xlsx", null)).toEqual({
      hostId: "wps-et",
      workbookName: "影院.xlsx",
      workbookIdentity: ""
    });
  });
});

describe("状态文案", () => {
  it("六种 phase 各有各的说法，互不冒充", () => {
    const texts = PHASES.map((phase) => describeResultRefFault(statusOf(phase), TARGET));
    expect(new Set(texts).size).toBe(PHASES.length);
    // "已触发"是终局事实，绝不能被说成"未启用"或"等待"。
    expect(describeResultRefFault(statusOf("consumed"), TARGET)).toContain("已触发");
    expect(describeResultRefFault(statusOf("bound"), TARGET)).toContain("已绑定");
    expect(describeResultRefFault(statusOf("armed"), TARGET)).toContain("下一次规划任务");
    expect(describeResultRefFault(statusOf("expired"), TARGET)).toContain("未触发");
  });

  it("服务端没开（enabled=false）与 phase=off 说同一件事", () => {
    expect(describeResultRefFault(statusOf("idle", { enabled: false }), TARGET)).toBe(
      describeResultRefFault(statusOf("off"), TARGET)
    );
    expect(describeResultRefFault(statusOf("off"), TARGET)).toContain("开发模式");
  });

  it("目标不可用与状态未读取是两件事，不许合并成一句", () => {
    const noTarget = describeResultRefFault(statusOf("idle"), null);
    const noStatus = describeResultRefFault(null, TARGET);
    expect(noTarget).not.toBe(noStatus);
    expect(noTarget).toContain("未启用");
    expect(noStatus).toContain("尚未读取");
  });

  it("已启用文案里带上剩余秒数（限时是一次性的一部分）", () => {
    expect(describeResultRefFault(statusOf("armed"), TARGET)).toContain("590");
  });
});

describe("按钮可按性", () => {
  it("忙的时候两个按钮都按不动", () => {
    expect(canArmResultRefFault(statusOf("idle"), TARGET, true)).toBe(false);
    expect(canCancelResultRefFault(statusOf("armed"), true)).toBe(false);
  });

  it("没有目标就不能启用（但取消只跟状态有关）", () => {
    expect(canArmResultRefFault(statusOf("idle"), null, false)).toBe(false);
    expect(canCancelResultRefFault(statusOf("armed"), false)).toBe(true);
  });

  it("已绑定到任务之后不能再改主意", () => {
    // "单次"的语义就是绑上就不能换：否则用户可以在任务开跑后把它挪到别的任务上。
    expect(canArmResultRefFault(statusOf("bound"), TARGET, false)).toBe(false);
    expect(canCancelResultRefFault(statusOf("bound"), false)).toBe(false);
  });

  it("未启用、待命、已过期、已触发之后都还能重新启用", () => {
    for (const phase of ["idle", "armed", "expired", "consumed"] as const) {
      expect(canArmResultRefFault(statusOf(phase), TARGET, false)).toBe(true);
    }
  });

  it("服务端没开时启用按钮不可按，取消也不需要", () => {
    expect(canArmResultRefFault(statusOf("off"), TARGET, false)).toBe(false);
    expect(canCancelResultRefFault(statusOf("off"), false)).toBe(false);
  });

  it("状态还没读回来时可以点启用（首次读取失败不该把人挡在外面）", () => {
    expect(canArmResultRefFault(null, TARGET, false)).toBe(true);
    expect(canArmResultRefFault(null, TARGET, true)).toBe(false);
    expect(canCancelResultRefFault(null, false)).toBe(false);
  });
});
