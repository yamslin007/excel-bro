import { describe, expect, it } from "vitest";
import {
  PLAN_SOURCE_MISSING_MESSAGE,
  PLAN_SOURCE_SCOPE_MISMATCH_MESSAGE,
  PLAN_SOURCE_STALE_MESSAGE,
  PLAN_SOURCE_UNVERIFIABLE_MESSAGE,
  judgePlanSourceGate,
  planHasBakedValues,
  planSourceComparisonNeeded,
  planSourceAnchor,
  planSourceGateMessage,
  sourceAnchorFromBinding
} from "./planSourceAnchor";

const ANCHOR = { fingerprint: "fp-read-1", sheets: ["会员表", "订单表"] };

/** 计划只关心动作类型；这里给最小形状，避免把"动作内容"与"来源"混在一起断言。 */
const plan = (...types: string[]) => ({
  actions: types.map((type) => ({ type })),
  sourceFingerprint: null as string | null,
  sourceFingerprintSheets: [] as string[]
});

const anchored = (types: string[], anchor = ANCHOR) => ({
  ...plan(...types),
  sourceFingerprint: anchor.fingerprint,
  sourceFingerprintSheets: [...anchor.sheets]
});

describe("planHasBakedValues：有没有'算好的固定值'", () => {
  it("写值/写表是烘焙，公式与复制、排序、格式不是", () => {
    expect(planHasBakedValues(plan("writeValues"))).toBe(true);
    expect(planHasBakedValues(plan("writeTable"))).toBe(true);
    // 执行时求值：公式由宿主算，复制/排序读的是当时的现场。
    expect(planHasBakedValues(plan("writeFormulas"))).toBe(false);
    expect(planHasBakedValues(plan("copyRange"))).toBe(false);
    expect(planHasBakedValues(plan("sortRange"))).toBe(false);
    expect(planHasBakedValues(plan("setFont"))).toBe(false);
    // 混装：只要有一处固定值就算。
    expect(planHasBakedValues(plan("writeFormulas", "writeValues"))).toBe(true);
    expect(planHasBakedValues(plan())).toBe(false);
    expect(planHasBakedValues(null)).toBe(false);
  });
});

describe("judgePlanSourceGate：'不需要锚点'与'本应有但缺失'要分得开", () => {
  it("执行时求值的计划不需要锚点，也不会因为没锚点被拒", () => {
    const verdict = judgePlanSourceGate({
      plan: plan("writeFormulas", "sortRange"),
      anchor: null,
      authorizedSheets: ["会员表", "订单表"],
      liveFingerprint: null
    });
    expect(verdict).toEqual({ verdict: "ok", basis: "runtime-evaluated" });
  });

  it("本机已独立复核的本地计算计划不需要计划级锚点", () => {
    const verdict = judgePlanSourceGate({
      plan: plan("writeValues"),
      anchor: null,
      authorizedSheets: ["会员表", "订单表"],
      liveFingerprint: null,
      independentSourceReview: true
    });
    expect(verdict).toEqual({ verdict: "ok", basis: "independent-review" });
  });

  it("有固定值却没有锚点：拒绝，且成因说得出来", () => {
    expect(
      judgePlanSourceGate({
        plan: plan("writeValues"),
        anchor: null,
        authorizedSheets: ["会员表"],
        liveFingerprint: "fp-now"
      })
    ).toEqual({ verdict: "refused", problem: "no-fingerprint" });

    // 有指纹、没有覆盖表：不知道它按哪些表算的，同样算缺失。
    expect(
      judgePlanSourceGate({
        plan: { ...plan("writeValues"), sourceFingerprint: "fp-read-1" },
        anchor: null,
        authorizedSheets: ["会员表"],
        liveFingerprint: "fp-now"
      })
    ).toEqual({ verdict: "refused", problem: "no-sheets" });

    // 装配方知道更精确的成因（例如锚点来自别的轮次）时，用它。
    expect(
      judgePlanSourceGate({
        plan: plan("writeValues"),
        anchor: null,
        problem: "foreign-task",
        authorizedSheets: ["会员表"],
        liveFingerprint: "fp-now"
      })
    ).toEqual({ verdict: "refused", problem: "foreign-task" });
  });

  it("锚点覆盖的表已不在授权范围：拒绝（写前复核读不到它们，也就比不了）", () => {
    expect(
      judgePlanSourceGate({
        plan: anchored(["writeValues"]),
        anchor: ANCHOR,
        // 用户收窄了授权，或这枚锚点本来就是别一轮的范围：写前复核读不到那张表。
        authorizedSheets: ["会员表"],
        liveFingerprint: ANCHOR.fingerprint
      })
    ).toEqual({ verdict: "refused", problem: "scope-mismatch" });
    expect(
      judgePlanSourceGate({
        plan: anchored(["writeValues"]),
        anchor: ANCHOR,
        authorizedSheets: [],
        liveFingerprint: ANCHOR.fingerprint
      })
    ).toEqual({ verdict: "refused", problem: "scope-mismatch" });
  });

  it("锚点比授权范围**更窄**：那是更精确的那版读数，照比", () => {
    // 值复制计划只按它真读过的那张表记账，授权却是两张 —— 不是"不一致"，是更准。
    expect(
      judgePlanSourceGate({
        plan: anchored(["writeValues"]),
        anchor: ANCHOR,
        authorizedSheets: ["会员表", "订单表", "汇总表"],
        liveFingerprint: ANCHOR.fingerprint
      })
    ).toEqual({ verdict: "ok", basis: "anchor" });
  });

  it("锚点在、现场读不到：unverifiable（不等于没变）", () => {
    expect(
      judgePlanSourceGate({
        plan: anchored(["writeValues"]),
        anchor: ANCHOR,
        authorizedSheets: ANCHOR.sheets,
        liveFingerprint: null
      })
    ).toEqual({ verdict: "unverifiable" });
    expect(
      judgePlanSourceGate({
        plan: anchored(["writeValues"]),
        anchor: ANCHOR,
        authorizedSheets: ANCHOR.sheets,
        liveFingerprint: "   "
      })
    ).toEqual({ verdict: "unverifiable" });
  });

  it("锚点在、现场一致/不一致：ok(anchor) / changed", () => {
    expect(
      judgePlanSourceGate({
        plan: anchored(["writeValues"]),
        anchor: ANCHOR,
        authorizedSheets: ["订单表", "会员表"],
        liveFingerprint: ANCHOR.fingerprint
      })
    ).toEqual({ verdict: "ok", basis: "anchor" });
    expect(
      judgePlanSourceGate({
        plan: anchored(["writeValues"]),
        anchor: ANCHOR,
        authorizedSheets: ANCHOR.sheets,
        liveFingerprint: "fp-now"
      })
    ).toEqual({ verdict: "changed" });
  });
});

describe("planSourceComparisonNeeded：什么时候才值得去读一次宿主", () => {
  it("执行时求值、本机已复核、没有锚点 —— 三种都不读", () => {
    expect(
      planSourceComparisonNeeded({ plan: plan("writeFormulas"), anchor: ANCHOR })
    ).toBe(false);
    expect(
      planSourceComparisonNeeded({
        plan: plan("writeValues"),
        anchor: ANCHOR,
        independentSourceReview: true
      })
    ).toBe(false);
    expect(planSourceComparisonNeeded({ plan: plan("writeValues"), anchor: null })).toBe(false);
  });

  it("有固定值、有锚点、没有独立复核 —— 必须读", () => {
    expect(planSourceComparisonNeeded({ plan: plan("writeValues"), anchor: ANCHOR })).toBe(true);
  });
});

describe("planSourceGateMessage：每种结论说自己的话", () => {
  it("四条拒绝/比对结论各有固定文案，放行没有话", () => {
    expect(planSourceGateMessage({ verdict: "ok", basis: "anchor" })).toBeNull();
    expect(planSourceGateMessage({ verdict: "changed" })).toBe(PLAN_SOURCE_STALE_MESSAGE);
    expect(planSourceGateMessage({ verdict: "unverifiable" })).toBe(
      PLAN_SOURCE_UNVERIFIABLE_MESSAGE
    );
    expect(planSourceGateMessage({ verdict: "refused", problem: "no-fingerprint" })).toBe(
      PLAN_SOURCE_MISSING_MESSAGE
    );
    expect(planSourceGateMessage({ verdict: "refused", problem: "foreign-task" })).toBe(
      PLAN_SOURCE_MISSING_MESSAGE
    );
    expect(planSourceGateMessage({ verdict: "refused", problem: "scope-mismatch" })).toBe(
      PLAN_SOURCE_SCOPE_MISMATCH_MESSAGE
    );
    // 三条文案必须互相区分：把"没记录过"说成"变了"会误导用户。
    expect(
      new Set([
        PLAN_SOURCE_STALE_MESSAGE,
        PLAN_SOURCE_UNVERIFIABLE_MESSAGE,
        PLAN_SOURCE_MISSING_MESSAGE,
        PLAN_SOURCE_SCOPE_MISMATCH_MESSAGE
      ]).size
    ).toBe(4);
  });
});

describe("planSourceAnchor / sourceAnchorFromBinding：锚点只认成对且同本轮", () => {
  it("指纹与覆盖表缺一不可", () => {
    expect(planSourceAnchor({ sourceFingerprint: "fp", sourceFingerprintSheets: [] })).toBeNull();
    expect(planSourceAnchor({ sourceFingerprint: "", sourceFingerprintSheets: ["A"] })).toBeNull();
    expect(planSourceAnchor({ sourceFingerprint: "fp", sourceFingerprintSheets: ["A"] })).toEqual({
      fingerprint: "fp",
      sheets: ["A"]
    });
  });

  it("任务不是本轮、指纹为空、覆盖表为空 —— 都不给锚点", () => {
    const binding = {
      taskId: "turn-1",
      sourceFingerprint: "fp-read-1",
      authorizedSheetNames: ["会员表"]
    };
    expect(sourceAnchorFromBinding(binding, "turn-2")).toBeNull();
    expect(sourceAnchorFromBinding(binding, "")).toBeNull();
    expect(sourceAnchorFromBinding({ ...binding, sourceFingerprint: null }, "turn-1")).toBeNull();
    expect(
      sourceAnchorFromBinding({ ...binding, authorizedSheetNames: [] }, "turn-1")
    ).toBeNull();
    expect(sourceAnchorFromBinding(binding, "turn-1")).toEqual({
      fingerprint: "fp-read-1",
      sheets: ["会员表"]
    });
  });
});

it("missing trusted binding removes response-supplied fingerprints", async () => {
  const { stampPlanSourceAnchor } = await import("./planSourceAnchor");
  const response = anchored(["writeValues"]);
  const result = stampPlanSourceAnchor(response, sourceAnchorFromBinding({
    taskId: "other-turn", sourceFingerprint: ANCHOR.fingerprint,
    authorizedSheetNames: ANCHOR.sheets
  }, "current-turn"));
  expect(planSourceAnchor(result)).toBeNull();
  expect(planSourceAnchor(response)).toEqual(ANCHOR);
  expect(planSourceGateMessage(judgePlanSourceGate({plan: result, anchor: planSourceAnchor(result),
    authorizedSheets: ANCHOR.sheets, liveFingerprint: ANCHOR.fingerprint}))).toBe(PLAN_SOURCE_MISSING_MESSAGE);
});
