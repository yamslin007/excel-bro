import { describe, it, expect } from "vitest";
import {
  isWorksheetDeletion,
  reconcileWorksheetDeletionScope,
  worksheetDeletionIssue
} from "./worksheetDeletion";
const plan = { actions: [{ type: "deleteWorksheet", sheet: "B" }] };
describe("worksheet deletion routing", () => {
  it("routes a pure plan and leaves additional archive permission to the checkpoint service", () => {
    expect(isWorksheetDeletion(plan)).toBe(true);
    expect(worksheetDeletionIssue(plan, ["A", "B"], ["A", "B"])).toBeNull();
    expect(worksheetDeletionIssue(plan, ["B"], ["A", "B"])).toBeNull();
  });
  it("rejects mixed, duplicate, last-sheet and unauthorized deletion", () => {
    expect(isWorksheetDeletion({ actions: [] })).toBe(false);
    expect(isWorksheetDeletion({ actions: [...plan.actions, { type: "setFont" }] })).toBe(false);
    expect(worksheetDeletionIssue({ actions: [...plan.actions, ...plan.actions] }, ["A", "B"], ["A", "B"])).toContain("重复");
    expect(worksheetDeletionIssue(plan, ["A"], ["A"])).toContain("授权");
    expect(worksheetDeletionIssue(plan, ["B"], ["B"])).toContain("保留");
  });
});

/**
 * 删表后的"预期收窄"。
 *
 * 这条判据决定的是：**能不能把授权范围悄悄同步掉、不再提示"尚未写入"**。
 * 所以每一条失败分支都必须能独立把同步拦下来，不能靠"看起来很合理"。
 */
describe("worksheet deletion scope reconciliation", () => {
  const base = {
    beforeIdentity: "url:c:/book.xlsx",
    afterIdentity: "url:c:/book.xlsx",
    authorizedSheets: ["A", "B", "C"],
    receipt: { deletedSheets: ["B"], remainingSheets: ["A", "C"] },
    actualSheets: ["A", "C"]
  };

  it("shrinks only the authorized scope when the receipt matches the live sheet set", () => {
    const outcome = reconcileWorksheetDeletionScope(base);
    expect(outcome).toEqual({ ok: true, scope: ["A", "C"] });
  });

  it("keeps the authorized order and never adds a sheet back", () => {
    const outcome = reconcileWorksheetDeletionScope({
      ...base,
      authorizedSheets: ["C", "A", "B"],
      actualSheets: ["A", "C", "D"]
    });
    // D 从未被授权，不得因为"现在存在"就进入范围。
    expect(outcome).toEqual({ ok: false, reason: "receipt_mismatch" });
    const kept = reconcileWorksheetDeletionScope({
      ...base,
      authorizedSheets: ["C", "A", "B"],
      receipt: { deletedSheets: ["B"], remainingSheets: ["A", "C", "D"] },
      actualSheets: ["A", "C", "D"]
    });
    expect(kept).toEqual({ ok: true, scope: ["C", "A"] });
  });

  it("refuses a different workbook or a missing identity", () => {
    expect(reconcileWorksheetDeletionScope({ ...base, afterIdentity: "url:c:/other.xlsx" }))
      .toEqual({ ok: false, reason: "workbook_changed" });
    expect(reconcileWorksheetDeletionScope({ ...base, afterIdentity: null }))
      .toEqual({ ok: false, reason: "identity_missing" });
    expect(reconcileWorksheetDeletionScope({ ...base, beforeIdentity: "" }))
      .toEqual({ ok: false, reason: "identity_missing" });
  });

  it("refuses when the live sheet set drifts from the receipt", () => {
    // 外部又删了一张：回执里还有 C，实际已经没了。
    expect(reconcileWorksheetDeletionScope({ ...base, actualSheets: ["A"] }))
      .toEqual({ ok: false, reason: "receipt_mismatch" });
    // 外部新增了一张：回执里没有它。
    expect(reconcileWorksheetDeletionScope({ ...base, actualSheets: ["A", "C", "D"] }))
      .toEqual({ ok: false, reason: "receipt_mismatch" });
    // 回执自报的删除清单与实际差一张。
    expect(reconcileWorksheetDeletionScope({
      ...base,
      receipt: { deletedSheets: ["B"], remainingSheets: ["A", "C"] },
      actualSheets: ["A", "C"],
      authorizedSheets: ["A", "B", "C", "D"]
    })).toEqual({ ok: false, reason: "receipt_mismatch" });
  });

  it("refuses a deletion that claims an unauthorized sheet or a still-present sheet", () => {
    expect(reconcileWorksheetDeletionScope({
      ...base,
      receipt: { deletedSheets: ["Z"], remainingSheets: ["A", "C"] }
    })).toEqual({ ok: false, reason: "receipt_mismatch" });
    expect(reconcileWorksheetDeletionScope({
      ...base,
      actualSheets: ["A", "B", "C"],
      receipt: { deletedSheets: ["B"], remainingSheets: ["A", "B", "C"] }
    })).toEqual({ ok: false, reason: "receipt_mismatch" });
  });

  it("never claims a shrink when nothing authorized survives", () => {
    expect(reconcileWorksheetDeletionScope({
      ...base,
      authorizedSheets: ["B"],
      receipt: { deletedSheets: ["B"], remainingSheets: ["A", "C"] },
      actualSheets: ["A", "C"]
    })).toEqual({ ok: false, reason: "nothing_authorized_left" });
  });

  it("ignores app-owned system sheets on both sides of the comparison", () => {
    const isSystemSheet = (name: string) => name.startsWith("#EB_");
    // Excel 的目录快照不含系统表，回执是全簿口径 —— 两边都要按同一判据过滤。
    expect(reconcileWorksheetDeletionScope({
      ...base,
      receipt: { deletedSheets: ["B"], remainingSheets: ["A", "C", "#EB_RULES"] },
      isSystemSheet
    })).toEqual({ ok: true, scope: ["A", "C"] });
    // WPS 的目录会把系统表列出来。
    expect(reconcileWorksheetDeletionScope({
      ...base,
      receipt: { deletedSheets: ["B"], remainingSheets: ["A", "C", "#EB_RULES"] },
      actualSheets: ["A", "C", "#EB_RULES"],
      isSystemSheet
    })).toEqual({ ok: true, scope: ["A", "C"] });
  });
});
