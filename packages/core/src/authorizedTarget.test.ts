import { describe, expect, it } from "vitest";
import {
  authorizedTargetEvidenceKey,
  canonicalTargetSelection,
  checkAuthorizedTarget,
  selectedTargetSheetNames,
  type AuthorizedTarget
} from "./authorizedTarget";
const bound: AuthorizedTarget = { fileKey: "file:a", name: "Same.xlsx", sheets: [{ id: "one", name: "Sheet1" }], selected: ["one"] };
describe("early authorized target detection", () => {
  it("does not confuse identical file names with identical targets", () => {
    expect(checkAuthorizedTarget(bound, { ...bound, fileKey: "file:b" })).toBe("file_changed");
  });
  it("detects unavailable host without asserting that data changed", () => {
    expect(checkAuthorizedTarget(bound, null)).toBe("unavailable");
  });
  it.each([
    { ...bound, selected: ["two"] },
    { ...bound, sheets: [{ id: "two", name: "Sheet1" }] },
    { ...bound, sheets: [{ id: "one", name: "Renamed" }] },
    { ...bound, sheets: [] }
  ])("detects selected sheet change/removal/replacement", (current) => {
    expect(checkAuthorizedTarget(bound, current)).toBe("scope_changed");
  });
  it("does not include unselected new sheets in authorization", () => {
    expect(checkAuthorizedTarget(bound, { ...bound, sheets: [...bound.sheets, { id: "two", name: "Extra" }] })).toBe("same");
  });
  it("requires same-session reference for unsaved identity", () => {
    const ref = {};
    const unsaved = { ...bound, fileKey: null, fileRef: ref };
    expect(checkAuthorizedTarget(unsaved, { ...unsaved })).toBe("same");
    expect(checkAuthorizedTarget(unsaved, { ...unsaved, fileRef: {} })).toBe("unavailable");
  });
});

/*
 * 授权集合的身份 = **成员的集合**，不是点选的先后顺序。
 *
 * 现场：用户按「Sheet11 → Sheet1 → Sheet10」勾了三张表（不是工作簿目录顺序），
 * 确认后发送需求，被拦在「暂时无法核实当前目标」。原因不是范围真的变了，而是
 * 两处**顺序敏感**的写法：已提交名单被按目录顺序重写、证据键直接
 * `JSON.stringify(selected)`。同一组表于是产出两个证据键，`checkForAnalysis`
 * 把它读成"证据冲突"。
 *
 * 规范化只消除顺序。它**不**去重、**不**合并成员、**不**改写成员 ——
 * 重复 ID 与缺失表仍然是真实差异，不能被"排序无关"吞掉。
 */
describe("authorized selection has one deterministic canonical form", () => {
  const sheets = [
    { id: "1", name: "Sheet1" },
    { id: "10", name: "Sheet10" },
    { id: "11", name: "Sheet11" }
  ];
  const pickedInUserOrder: AuthorizedTarget = {
    fileKey: "file:a", name: "Book.xlsx", sheets, selected: ["11", "1", "10"]
  };
  const writtenInDirectoryOrder: AuthorizedTarget = { ...pickedInUserOrder, selected: ["1", "10", "11"] };

  it("the same set of sheet ids is the same scope whatever the click order was", () => {
    expect(checkAuthorizedTarget(pickedInUserOrder, writtenInDirectoryOrder)).toBe("same");
  });

  it("canonicalises a copy and leaves the caller's array untouched", () => {
    const selected = ["11", "1", "10"];
    expect(canonicalTargetSelection(selected)).toEqual(["1", "10", "11"]);
    // 原地 sort 会把调用方"按点选顺序映射名字"的那份数组改掉。
    expect(selected).toEqual(["11", "1", "10"]);
    expect(canonicalTargetSelection(selected)).not.toBe(selected);
  });

  it("order-insensitivity never merges members: duplicates stay a real difference", () => {
    expect(canonicalTargetSelection(["10", "1", "10"])).toEqual(["1", "10", "10"]);
    expect(checkAuthorizedTarget(
      { ...pickedInUserOrder, selected: ["1", "10"] },
      { ...pickedInUserOrder, selected: ["10", "1", "10"] })).toBe("scope_changed");
  });

  it("still detects added, removed and replaced members", () => {
    expect(checkAuthorizedTarget(pickedInUserOrder, { ...pickedInUserOrder, selected: ["11", "1", "10", "2"] }))
      .toBe("scope_changed");
    expect(checkAuthorizedTarget(pickedInUserOrder, { ...pickedInUserOrder, selected: ["11", "10"] }))
      .toBe("scope_changed");
    // 缺表：id 还在名单里、目录里却没有了 —— 不是"顺序问题"。
    expect(checkAuthorizedTarget(pickedInUserOrder, {
      ...pickedInUserOrder, sheets: sheets.filter((sheet) => sheet.id !== "1")
    })).toBe("scope_changed");
  });

  it("builds the same evidence key for the same set in a different order, and a different one otherwise", () => {
    expect(authorizedTargetEvidenceKey("id", ["11", "1", "10"]))
      .toBe(authorizedTargetEvidenceKey("id", ["1", "10", "11"]));
    expect(authorizedTargetEvidenceKey("id", ["1", "10"]))
      .not.toBe(authorizedTargetEvidenceKey("id", ["1", "10", "11"]));
    expect(authorizedTargetEvidenceKey("other-id", ["1", "10", "11"]))
      .not.toBe(authorizedTargetEvidenceKey("id", ["1", "10", "11"]));
  });

  it("keeps the authorised sheet names in the order the user picked them", () => {
    expect(selectedTargetSheetNames(pickedInUserOrder)).toEqual(["Sheet11", "Sheet1", "Sheet10"]);
    // 目录里多一张未授权的表不改变"已授权的是哪些"
    expect(selectedTargetSheetNames({
      ...pickedInUserOrder, sheets: [...sheets, { id: "12", name: "Sheet12" }]
    })).toEqual(["Sheet11", "Sheet1", "Sheet10"]);
  });
});
