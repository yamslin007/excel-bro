import { describe, expect, it } from "vitest";
import {
  pendingResumeReadiness,
  pendingResumeTargetSheets,
  type PendingResumeInput
} from "./pendingExecutionResume";

const BASE: PendingResumeInput = {
  currentIdentity: "url:https://contoso.sharepoint.com/book.xlsx",
  recordedIdentity: "url:https://contoso.sharepoint.com/book.xlsx",
  status: "running",
  canResume: true,
  inFlightActionIndex: null,
  hasBeforeSnapshot: false,
  recordedAuthorizedSheets: ["来源表", "目标表"],
  // 默认"本会话还没确认过范围"：这时不做比对（宿主挂载时的播种不是用户手势）。
  currentAuthorizedSheets: null,
  planTargetSheets: ["目标表"],
  workbookSheetNames: ["来源表", "目标表", "备用表"]
};

function readiness(overrides: Partial<PendingResumeInput>) {
  return pendingResumeReadiness({ ...BASE, ...overrides });
}

describe("unfinished execution resume readiness", () => {
  it("身份、授权与现场都对得上时才允许摆出可续跑的卡", () => {
    expect(readiness({})).toEqual({ ready: true, requiresReconciliation: false });
    expect(readiness({ status: "needs_reconciliation", inFlightActionIndex: 2, hasBeforeSnapshot: true }))
      .toEqual({ ready: true, requiresReconciliation: true });
  });

  it("身份读不出来就阻断，不退回文件名或登记值", () => {
    const result = readiness({ currentIdentity: null });
    expect(result).toMatchObject({ ready: false, reason: "identity_unavailable" });
  });

  it("另一个工作簿的记录不能摆到当前工作簿上", () => {
    const result = readiness({ currentIdentity: "url:https://contoso.sharepoint.com/other.xlsx" });
    expect(result).toMatchObject({ ready: false, reason: "identity_mismatch" });
    // 同名不同路径也算另一个簿：身份是完整字符串比较，不做文件名相似度。
    const sameName = readiness({ currentIdentity: "url:https://contoso.sharepoint.com/a/book.xlsx" });
    expect(sameName).toMatchObject({ ready: false, reason: "identity_mismatch" });
  });

  it("已完成的记录不再作为待续写出现", () => {
    expect(readiness({ status: "completed" })).toMatchObject({ ready: false, reason: "completed" });
  });

  it("服务端没释放执行权时阻断，不重复登记新写入", () => {
    expect(readiness({ canResume: false })).toMatchObject({ ready: false, reason: "not_resumable" });
  });

  it("没有用户确认过的范围时不做比对，也不把 null 当成“一张都没授权”", () => {
    expect(readiness({ currentAuthorizedSheets: null })).toEqual({ ready: true, requiresReconciliation: false });
  });

  it("用户确认后把目标表收窄掉时点名阻断，不自动把表加回去", () => {
    const result = readiness({ currentAuthorizedSheets: ["来源表"] });
    expect(result).toMatchObject({
      ready: false,
      reason: "plan_target_not_authorized_now",
      sheets: ["目标表"]
    });
    if (result.ready) throw new Error("unreachable");
    expect(result.detail).toContain("目标表");
    expect(result.detail).toContain("不会自行扩大授权");
  });

  it("计划要写的表不在登记授权里时点名阻断（记录自相矛盾，fail-closed）", () => {
    const result = readiness({ recordedAuthorizedSheets: ["来源表"] });
    expect(result).toMatchObject({
      ready: false,
      reason: "plan_target_never_authorized",
      sheets: ["目标表"]
    });
    if (result.ready) throw new Error("unreachable");
    expect(result.detail).toContain("目标表");
  });

  it("记录里从来没有的目标表属于记录自相矛盾，与“现在没勾选”分开报", () => {
    const result = readiness({ recordedAuthorizedSheets: ["来源表"] });
    expect(result).toMatchObject({ ready: false, reason: "plan_target_never_authorized", sheets: ["目标表"] });
  });

  it("目标表已从工作簿里消失时单独报出来", () => {
    const result = readiness({ workbookSheetNames: ["来源表", "备用表"] });
    expect(result).toMatchObject({ ready: false, reason: "plan_target_missing", sheets: ["目标表"] });
  });

  it("在途动作缺写前快照时只能人工核对", () => {
    const result = readiness({ inFlightActionIndex: 1, hasBeforeSnapshot: false });
    expect(result).toMatchObject({ ready: false, reason: "in_flight_without_snapshot" });
    // 有快照就是可核对，不是"不可续跑"。
    expect(readiness({ inFlightActionIndex: 1, hasBeforeSnapshot: true }))
      .toEqual({ ready: true, requiresReconciliation: true });
  });

  it("计划目标表名去重去空，但不做名字归一化", () => {
    expect(pendingResumeTargetSheets({
      actions: [
        { sheet: "目标表" },
        { sheet: " 目标表 " },
        { sheet: "来源表" },
        { sheet: "" },
        { sheet: null },
        { sheet: "目标表 " }
      ]
    })).toEqual(["目标表", "来源表"]);
  });
});
