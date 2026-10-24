import { describe, it, expect } from "vitest";
import { reconcileWorksheetRestoreScope } from "./worksheetRestore";
import { isEBSystemSheet } from "./folderScope";

/**
 * 原地撤销之后的"预期扩张"。
 *
 * 与删表时的 `reconcileWorksheetDeletionScope` 对称：撤销会把表**放回**工作簿，
 * 于是目录比授权时多几张。判据要回答两个问题，且必须分开回答 ——
 *   ① 这次目录变化能不能被这次撤销解释（不然下一轮守卫会误报"目标已变化"）；
 *   ② 放回来的表**没有**被偷偷加进授权（撤销 ≠ 扩大授权）。
 *
 * 每条失败分支都要能独立把"放过"拦下来，否则等于全局关掉了目标守卫。
 */
describe("worksheet restore scope reconciliation", () => {
  const base = {
    beforeIdentity: "url:c:/book.xlsx",
    afterIdentity: "url:c:/book.xlsx",
    // 删表后已提交的授权：只剩 A、C。
    authorizedSheets: ["A", "C"],
    restoredSheets: ["B"],
    // 撤销后宿主真实目录：B 回到原位。
    actualSheets: ["A", "B", "C"]
  };

  it("accepts the restored sheets back into the directory but never into the authorized set", () => {
    expect(reconcileWorksheetRestoreScope(base)).toEqual({ ok: true, authorized: ["A", "C"] });
  });

  it("keeps the authorized order exactly as submitted", () => {
    // 顺序按提交的那一份保留，不按目录顺序重排（否则授权基线会无谓地变一次）。
    expect(reconcileWorksheetRestoreScope({
      ...base,
      authorizedSheets: ["C", "A"],
      actualSheets: ["A", "B", "C"]
    })).toEqual({ ok: true, authorized: ["C", "A"] });
  });

  it("restores several sheets at once", () => {
    expect(reconcileWorksheetRestoreScope({
      ...base,
      restoredSheets: ["B", "D"],
      actualSheets: ["A", "B", "C", "D"]
    })).toEqual({ ok: true, authorized: ["A", "C"] });
  });

  it("refuses a different workbook or a missing identity", () => {
    expect(reconcileWorksheetRestoreScope({ ...base, afterIdentity: "url:c:/other.xlsx" }))
      .toEqual({ ok: false, reason: "workbook_changed" });
    expect(reconcileWorksheetRestoreScope({ ...base, afterIdentity: null }))
      .toEqual({ ok: false, reason: "identity_missing" });
    expect(reconcileWorksheetRestoreScope({ ...base, beforeIdentity: "" }))
      .toEqual({ ok: false, reason: "identity_missing" });
  });

  it("refuses when the receipt claims a sheet the directory does not have", () => {
    // 回执说 B 放回去了，实际目录里没有 B：这份回执不能当证据。
    expect(reconcileWorksheetRestoreScope({ ...base, actualSheets: ["A", "C"] }))
      .toEqual({ ok: false, reason: "restored_missing" });
    // 空回执同理（没有可核对的放回事实）。
    expect(reconcileWorksheetRestoreScope({ ...base, restoredSheets: [] }))
      .toEqual({ ok: false, reason: "restored_missing" });
  });

  it("refuses when an authorized sheet disappeared", () => {
    // 撤销只加表，不该让任何已授权表消失；少了就是别的东西在动工作簿。
    expect(reconcileWorksheetRestoreScope({ ...base, actualSheets: ["A", "B"] }))
      .toEqual({ ok: false, reason: "authorization_changed" });
  });

  it("refuses sheets that this restore cannot explain", () => {
    // 用户自己新建的表 / 外部程序加的表：不能借撤销的名义被一并放过。
    expect(reconcileWorksheetRestoreScope({ ...base, actualSheets: ["A", "B", "C", "新表"] }))
      .toEqual({ ok: false, reason: "unexpected_sheets" });
  });

  it("ignores EB system sheets when judging the directory", () => {
    // 系统表不进授权、也不该被算成"多出来的表"。
    expect(reconcileWorksheetRestoreScope({
      ...base,
      actualSheets: ["A", "B", "C", "#EB_RULES"],
      isSystemSheet: isEBSystemSheet
    })).toEqual({ ok: true, authorized: ["A", "C"] });
  });
});
