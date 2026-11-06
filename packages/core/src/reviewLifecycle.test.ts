/**
 * 结果审阅的**持久生命周期**判据（2026-11-03）。
 *
 * 缺陷：审阅状态只活在 `ResultReview` 的内存里。执行完成、回读通过，用户还没点
 * 「确认全部变更 / 放弃全部变更」时刷新窗格，待审阅界面就消失 —— 而原表停在一个
 * 没人收尾的写入现场上。
 *
 * 这组用例钉的是"什么不算审阅状态"：
 * - `status === 'completed'` 不算（那是动作跑完了，不是用户看过）；
 * - 标记、双视图、组件挂载都不算；
 * - `review_state IS NULL` 的老记录既不是待审阅、也不是已确认。
 */
import { describe, expect, it } from "vitest";
import {
  reviewBlockerMessage,
  reviewDecisionAllowed,
  reviewIsTerminal,
  reviewNeedsAttention,
  reviewRestoreVerdict,
  reviewStateAfterDecision,
  type ReviewRestoreInput
} from "./reviewLifecycle";

function input(overrides: Partial<ReviewRestoreInput> = {}): ReviewRestoreInput {
  return {
    reviewState: "pending",
    recordedIdentity: "wps-path:c:\\work\\a.xlsx",
    currentIdentity: "wps-path:c:\\work\\a.xlsx",
    recordedSheets: ["结果", "来源"],
    currentSheets: ["来源", "结果"],
    status: "completed",
    resultPresent: true,
    resultVerified: true,
    materialAvailable: true,
    ...overrides
  };
}

describe("终态与待处理是两回事", () => {
  it("已确认/已放弃/已撤销不再需要用户处理，也要能被单独认出来", () => {
    for (const state of ["pending", "blocked"] as const) {
      expect(reviewNeedsAttention(state)).toBe(true);
      expect(reviewIsTerminal(state)).toBe(false);
    }
    for (const state of ["confirmed", "discarded", "undone"] as const) {
      expect(reviewNeedsAttention(state)).toBe(false);
    }
    expect(reviewIsTerminal("discarded")).toBe(true);
    expect(reviewIsTerminal("undone")).toBe(true);
    // 已确认不是终态：确认之后还能撤销（那是真实存在的一条路）。
    expect(reviewIsTerminal("confirmed")).toBe(false);
  });

  it("`null`（老记录）不参与「要不要处理」的判断 —— 既不是待审阅也不是已确认", () => {
    expect(reviewNeedsAttention(null)).toBe(false);
  });

  it("确认之后只允许撤销，终态不接收任何新决定", () => {
    expect(reviewDecisionAllowed("pending", "confirmed")).toBe(true);
    expect(reviewDecisionAllowed("pending", "discarded")).toBe(true);
    expect(reviewDecisionAllowed("blocked", "confirmed")).toBe(true);
    expect(reviewDecisionAllowed("confirmed", "undone")).toBe(true);
    expect(reviewDecisionAllowed("confirmed", "discarded")).toBe(false);
    expect(reviewDecisionAllowed("confirmed", "confirmed")).toBe(false);
    for (const decision of ["confirmed", "discarded", "undone"] as const) {
      expect(reviewDecisionAllowed("discarded", decision)).toBe(false);
      expect(reviewDecisionAllowed("undone", decision)).toBe(false);
    }
  });

  it("决定落账后的状态与决定同名，两种「还原」不互相冒充", () => {
    expect(reviewStateAfterDecision("discarded")).toBe("discarded");
    expect(reviewStateAfterDecision("undone")).toBe("undone");
    expect(reviewStateAfterDecision("confirmed")).toBe("confirmed");
  });
});

describe("刷新后「这条审阅还能不能重建」", () => {
  it("身份与授权都对得上、结果在且已验证 ⇒ 可重建，且能确认能放弃", () => {
    const verdict = reviewRestoreVerdict(input());
    expect(verdict.kind).toBe("restorable");
    expect(verdict.blockers).toEqual([]);
    expect(verdict.canConfirm).toBe(true);
    expect(verdict.canDiscard).toBe(true);
    expect(verdict.limitations).toEqual({});
  });

  it("授权按集合比较，不看勾选顺序", () => {
    expect(reviewRestoreVerdict(input({ currentSheets: ["结果", "来源"] })).kind).toBe("restorable");
    // 反方向（顺序不同）也一样：集合相同就没有限制。
    expect(
      reviewRestoreVerdict(input({ currentSheets: ["来源", "结果"] })).limitations.scope
    ).toBeUndefined();
  });

  it("当前选区与登记范围不同**不阻断**恢复：那只是提示，不是判据", () => {
    /*
     * 2026-11-03 实测踩到的错：拿刷新后的当前选区去判"这条审阅还在不在"，
     * 结果每次刷新都因为"选区没恢复"把待审阅弄丢，用户再也回不到收尾现场。
     * 恢复一律按账本登记的授权范围核对与恢复，当前选区既不参与也不扩大原范围。
     */
    for (const currentSheets of [["结果"], ["结果", "来源", "其他"], []]) {
      const verdict = reviewRestoreVerdict(input({ currentSheets }));
      expect(verdict.kind).toBe("restorable");
      expect(verdict.blockers).toEqual([]);
      expect(verdict.limitations.scope).toBeTruthy();
      // 提示归提示，两个操作都不受限。
      expect(verdict.canConfirm).toBe(true);
      expect(verdict.canDiscard).toBe(true);
    }
  });

  it("终态优先：已放弃/已撤销/已确认的记录刷新后不许再变回待审阅", () => {
    for (const state of ["discarded", "undone", "confirmed"] as const) {
      const verdict = reviewRestoreVerdict(input({ reviewState: state }));
      expect(verdict.kind).toBe(state === "confirmed" ? "restorable" : "settled");
    }
    // 已确认仍是"可以再撤销"的活记录，但它不再是待审阅。
    expect(reviewNeedsAttention("confirmed")).toBe(false);
  });

  it("老记录（review_state IS NULL）只报「无法判断」，两种假装都不做", () => {
    const verdict = reviewRestoreVerdict(input({ reviewState: null }));
    expect(verdict.kind).toBe("legacy_unknown");
    expect(verdict.canConfirm).toBe(false);
    expect(verdict.canDiscard).toBe(false);
  });

  it("身份不符时先拦身份，不拿「结果缺失」掩盖身份问题", () => {
    const verdict = reviewRestoreVerdict(
      input({ currentIdentity: "wps-path:c:\\work\\b.xlsx", resultPresent: false })
    );
    expect(verdict.kind).toBe("blocked");
    expect(verdict.blockers).toEqual(["identity_mismatch", "result_missing"]);
    expect(verdict.blockers.map(reviewBlockerMessage).every((text) => text.length > 0)).toBe(true);
  });

  it("结果没通过独立验证 ⇒ 挡住确认，但仍如实说出原因", () => {
    const verdict = reviewRestoreVerdict(input({ resultVerified: false }));
    expect(verdict.kind).toBe("blocked");
    expect(verdict.blockers).toEqual(["result_unverified"]);
    expect(verdict.canConfirm).toBe(false);
  });

  it("材料缺失**不**等于没有待审阅：只限制「放弃」与「双视图」", () => {
    const verdict = reviewRestoreVerdict(input({ materialAvailable: false }));
    expect(verdict.kind).toBe("restorable");
    expect(verdict.limitations.discard).toBeTruthy();
    expect(verdict.limitations.comparison).toBeTruthy();
    expect(verdict.canDiscard).toBe(false);
    // 确认只需要已有结果，不需要恢复材料。
    expect(verdict.canConfirm).toBe(true);
  });

  it("`blocked` 档（结果未验证）不允许「确认」，但仍是待处理记录", () => {
    const verdict = reviewRestoreVerdict(input({ reviewState: "blocked" }));
    expect(verdict.kind).toBe("restorable");
    expect(verdict.canConfirm).toBe(false);
    expect(reviewNeedsAttention("blocked")).toBe(true);
  });
});
