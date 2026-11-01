import { describe, expect, it } from "vitest";
import {
  describeChangedWhere,
  describeNextStep,
  describeVerification,
  processOutcomePhaseFrom,
  processOutcomeSummary,
  type ProcessOutcomePhase
} from "./processOutcome";

const applied = (sheet: string, range: string, changed: number) => ({
  sheet,
  range,
  status: "applied",
  rows: Array.from({ length: changed }, (_, index) => ({ changed: true, key: String(index) }))
});

describe("describeChangedWhere", () => {
  it("说得出具体位置与格数", () => {
    expect(describeChangedWhere([applied("汇总", "B2:D4", 9)]))
      .toBe("改了：汇总!B2:D4（9 格）");
  });

  it("超过三处只列前三处并给总数", () => {
    const text = describeChangedWhere([
      applied("A", "A1", 1),
      applied("B", "B1", 1),
      applied("C", "C1", 1),
      applied("D", "D1", 1)
    ])!;
    expect(text.startsWith("改了 4 处：")).toBe(true);
    expect(text).toContain("A!A1");
    expect(text).not.toContain("D!D1");
  });

  it("没有证据就什么都不说（不是给「零写入」）", () => {
    // 反向锚点：把 null 改成 "本次没有写入" 就是把"不知道"说成"确定没写"。
    expect(describeChangedWhere([])).toBeNull();
    expect(describeChangedWhere(null)).toBeNull();
    expect(describeChangedWhere(undefined)).toBeNull();
  });

  it("读数失败与未执行不算改动", () => {
    expect(describeChangedWhere([
      { sheet: "S", range: "A1:A2", status: "readback_failed", rows: [{ changed: true }] }
    ])).toBeNull();
    expect(describeChangedWhere([
      { sheet: "S", range: "A1:A2", status: "not_executed", rows: [{ changed: true }] }
    ])).toBeNull();
  });

  it("写入了但没有任何格子变化 ⇒ 不宣称改了哪里", () => {
    expect(describeChangedWhere([
      { sheet: "S", range: "A1:A2", status: "applied", rows: [{ changed: false }, {}] }
    ])).toBeNull();
  });
});

describe("describeVerification", () => {
  it("三档各自说清楚", () => {
    expect(describeVerification({ status: "verified", total: 6, passed: 6, unverified: 0 }))
      .toBe("6/6 项符合预期");
    expect(describeVerification({ status: "executed_unverified", total: 6, passed: 5, unverified: 1 }))
      .toBe("已执行；1 步缺少独立验收");
    expect(describeVerification({ status: "failed", total: 4, passed: 2, unverified: 0 }))
      .toBe("2/4 项不符合预期");
  });

  it("没有回执就不给结论（不拿「已执行」冒充「已验证」）", () => {
    expect(describeVerification(null)).toBeNull();
    expect(describeVerification(undefined)).toBeNull();
  });
});

describe("describeNextStep", () => {
  it("每个阶段都有可执行的下一步，且互不相同", () => {
    const phases: ProcessOutcomePhase[] = [
      "reviewing", "confirmed", "discarded", "undone",
      "needs_review", "resumable", "blocked", "executing", "idle"
    ];
    const texts = phases.map((phase) => describeNextStep(phase));
    for (const text of texts) expect(text.trim().length).toBeGreaterThan(0);
    // 审阅 / 待核对 / 可续跑 三句必须不同：它们指向三个不同的按钮。
    expect(describeNextStep("reviewing")).not.toBe(describeNextStep("needs_review"));
    expect(describeNextStep("needs_review")).not.toBe(describeNextStep("resumable"));
  });

  it("撤销与放弃说的是恢复，不是「继续」", () => {
    expect(describeNextStep("undone")).toContain("撤销");
    expect(describeNextStep("discarded")).toContain("还原");
  });
});

describe("processOutcomePhaseFrom", () => {
  it("终态优先：撤销过就是撤销过，不会被旧的审阅阶段改判", () => {
    expect(processOutcomePhaseFrom({ undone: true, reviewPhase: "reviewing" })).toBe("undone");
    expect(processOutcomePhaseFrom({ abandoned: true, reviewPhase: "confirmed" })).toBe("discarded");
  });

  it("审阅阶段优先于恢复结论", () => {
    expect(processOutcomePhaseFrom({ reviewPhase: "reviewing", recoveryKind: "resume" })).toBe("reviewing");
    expect(processOutcomePhaseFrom({ reviewPhase: "confirmed" })).toBe("confirmed");
    expect(processOutcomePhaseFrom({ reviewPhase: "discarded" })).toBe("discarded");
  });

  it("恢复结论各归各位：能续跑的别写成要核对，要核对的别写成能续跑", () => {
    expect(processOutcomePhaseFrom({ recoveryKind: "resume" })).toBe("resumable");
    expect(processOutcomePhaseFrom({ recoveryKind: "reconcile" })).toBe("needs_review");
    expect(processOutcomePhaseFrom({ recoveryKind: "completed" })).toBe("needs_review");
    expect(processOutcomePhaseFrom({ recoveryKind: "undone" })).toBe("undone");
    for (const kind of ["replan", "blocked", "insufficient"]) {
      expect(processOutcomePhaseFrom({ recoveryKind: kind })).toBe("blocked");
    }
  });

  it("什么证据都没有就是 idle，不猜一个阶段", () => {
    expect(processOutcomePhaseFrom({})).toBe("idle");
    expect(processOutcomePhaseFrom({ recoveryKind: "unknown-kind" })).toBe("idle");
  });
});

describe("processOutcomeSummary", () => {
  it("进行中不出结论", () => {
    expect(processOutcomeSummary({
      live: true,
      phase: "executing",
      reviews: [applied("汇总", "B2", 1)],
      verification: { status: "verified", total: 1, passed: 1, unverified: 0 }
    })).toBeNull();
  });

  it("正常结束：三段齐全", () => {
    const summary = processOutcomeSummary({
      live: false,
      phase: "reviewing",
      reviews: [applied("汇总", "B2:D4", 9)],
      verification: { status: "verified", total: 9, passed: 9, unverified: 0 }
    })!;
    expect(summary.changedWhere).toBe("改了：汇总!B2:D4（9 格）");
    expect(summary.verification).toBe("9/9 项符合预期");
    expect(summary.nextStep).toContain("确认变更");
  });

  it("已撤销：只说恢复事实，不复述旧验证结论", () => {
    const summary = processOutcomeSummary({
      live: false,
      phase: "undone",
      reviews: [applied("汇总", "B2:D4", 9)],
      verification: { status: "verified", total: 9, passed: 9, unverified: 0 }
    })!;
    expect(summary.changedWhere).toContain("还原");
    // 关键：撤销之后再说"9/9 项符合预期"是在描述一段已经不存在的状态。
    expect(summary.verification).toBeNull();
    expect(summary.nextStep).toContain("撤销");
  });

  it("没有验收回执时，验证结果那行留空而不是给一句肯定的结论", () => {
    const summary = processOutcomeSummary({
      live: false,
      phase: "confirmed",
      reviews: [applied("汇总", "B2", 1)],
      verification: null
    })!;
    expect(summary.verification).toBeNull();
    expect(summary.changedWhere).toBe("改了：汇总!B2（1 格）");
  });

  it("档位型回执（WPS）用调用方那句已有结论，且不覆盖计数型", () => {
    expect(processOutcomeSummary({
      live: false,
      phase: "reviewing",
      verificationNote: "已写入，计算值未获证实"
    })!.verification).toBe("已写入，计算值未获证实");

    // 有计数型回执时以计数为准，避免同一行出现两种口径。
    expect(processOutcomeSummary({
      live: false,
      phase: "reviewing",
      verification: { status: "verified", total: 2, passed: 2, unverified: 0 },
      verificationNote: "已写入，计算值未获证实"
    })!.verification).toBe("2/2 项符合预期");
  });

  it("无回执无待办不显示空摘要，实际回执及阻断仍显示", () => {
    expect(processOutcomeSummary({ live: false, phase: "idle" })).toBeNull();
    expect(processOutcomeSummary({ live: false, phase: "idle", verificationNote: "回读失败" })?.verification).toBe("回读失败");
    expect(processOutcomeSummary({ live: false, phase: "blocked" })?.nextStep).toContain("不能继续");
  });

  it("撤销之后连档位结论也不再复述", () => {
    expect(processOutcomeSummary({
      live: false,
      phase: "undone",
      verificationNote: "已写入，计算值未获证实"
    })!.verification).toBeNull();
  });
});
