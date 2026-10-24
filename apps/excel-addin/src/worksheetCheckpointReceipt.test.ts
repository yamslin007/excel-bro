import { describe, it, expect } from "vitest";
import { assertWorksheetCheckpointResult, assertWorksheetRestoreResult } from "./contracts";
const receipt = { operationId: "a".repeat(32), status: "completed", deletedSheets: ["B"], remainingSheets: ["A"], recoveryAvailable: true, message: "done" };
describe("worksheet checkpoint receipt", () => {
  it("accepts verified completion and unknown outcomes without inventing success", () => {
    expect(() => assertWorksheetCheckpointResult(receipt, receipt.operationId)).not.toThrow();
    expect(() => assertWorksheetCheckpointResult({ ...receipt, status: "partial_or_unknown", deletedSheets: [], remainingSheets: [] }, receipt.operationId)).not.toThrow();
  });
  it.each([{ recoveryAvailable: false }, { remainingSheets: [] }, { remainingSheets: ["B"] }, { deletedSheets: null }, { operationId: "b".repeat(32) }, { status: "success" }])("rejects incomplete or mismatched completion %j", change => {
    expect(() => assertWorksheetCheckpointResult({ ...receipt, ...change }, receipt.operationId)).toThrow("尚未确认");
  });
});

it("requires a complete backup-only consent receipt with no prior mutation", () => {
  const pending = { ...receipt, status: "needs_backup_consent", recoveryAvailable: false, deletedSheets: [], remainingSheets: [], backupSheets: ["Hidden"] };
  expect(() => assertWorksheetCheckpointResult(pending, receipt.operationId)).not.toThrow();
  for (const change of [{ backupSheets: [] }, { backupSheets: undefined }, { recoveryAvailable: true }, { deletedSheets: ["B"] }]) {
    expect(() => assertWorksheetCheckpointResult({ ...pending, ...change }, receipt.operationId)).toThrow();
  }
});

/**
 * 撤销回执也要**先自证**再进结果卡。
 *
 * 三条硬约束（见 `contracts.ts` 的注释）：归属必须是这次执行；`completed` 必须覆盖
 * 本次删除的全部表且没有失败/未核对；`partial` 必须真的"只完成一部分"。
 * 任何一条被放宽，界面都会在没核对完的时候说"已撤销"。
 */
describe("worksheet restore receipt", () => {
  const restore = {
    operationId: "a".repeat(32),
    status: "completed",
    restoredSheets: ["B", "C"],
    failedSheets: [] as string[],
    unverifiedSheets: [] as string[],
    repairedFormulas: 2,
    repairedNames: 1,
    limitations: [] as string[],
    message: "已把 2 张工作表放回原工作簿。"
  };
  const deletedSheets = ["B", "C"];

  it("accepts a receipt that covers exactly the deleted sheets", () => {
    expect(() => assertWorksheetRestoreResult(restore, restore.operationId, deletedSheets)).not.toThrow();
  });

  it("accepts a partial receipt with both lists filled", () => {
    expect(() => assertWorksheetRestoreResult({
      ...restore, status: "partial", restoredSheets: ["B"], failedSheets: ["C"]
    }, restore.operationId, deletedSheets)).not.toThrow();
  });

  it.each([
    // 归属不是这次执行
    { operationId: "b".repeat(32) },
    // 声称完成，却少放回一张
    { restoredSheets: ["B"] },
    // 声称完成，却有放回失败 / 未核对的
    { failedSheets: ["C"] },
    { unverifiedSheets: ["C"] },
    // 结构不完整：撤销成功的话术不能建立在半份回执上
    { restoredSheets: null },
    { limitations: null },
    { repairedFormulas: "2" },
    { message: undefined },
    { status: "success" }
  ])("rejects an untrustworthy completion %j", change => {
    expect(() => assertWorksheetRestoreResult({ ...restore, ...change }, restore.operationId, deletedSheets)).toThrow();
  });

  it("rejects a partial receipt that actually restored everything or nothing", () => {
    expect(() => assertWorksheetRestoreResult(
      { ...restore, status: "partial", failedSheets: [] }, restore.operationId, deletedSheets)).toThrow("清单与状态不一致");
    expect(() => assertWorksheetRestoreResult(
      { ...restore, status: "partial", restoredSheets: [], failedSheets: ["B", "C"] }, restore.operationId, deletedSheets))
      .toThrow("清单与状态不一致");
  });
});
