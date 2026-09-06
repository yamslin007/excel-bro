import { describe, expect, it } from "vitest";

import {
  evaluateBulkWorksheetDeletion,
  parseBulkWorksheetDeletion
} from "./structuralIntent";

describe("parseBulkWorksheetDeletion", () => {
  it.each([
    "清除sheet-dc1外的所有子表",
    "删除除 sheet-dc1 外的所有工作表",
    "只保留 sheet-dc1，删除其他工作表",
    "保留 sheet-dc1，删除其他所有工作表",
    "僅保留『sheet-dc1』，刪除其餘工作表",
    "将本表中除了sheet-dc1外的sheet表都删除",
    "把除了 sheet-dc1 外的所有工作表都删除",
    "将工作簿中除 sheet-dc1 外的其他工作表删除",
    "请将当前工作簿里除了sheet-dc1外的所有其他sheet表全部删除掉"
  ])("识别独立的保留一表批量删除命令：%s", (prompt) => {
    expect(parseBulkWorksheetDeletion(prompt)).toEqual({
      keeperName: "sheet-dc1"
    });
  });

  it.each([
    "清空 sheet-dc1 的数据",
    "删除 A2:A20",
    "按影片名称拆分工作表",
    "把 sheet-dc1 外的数据清空"
  ])("不把其他清空或拆分请求识别成删表：%s", (prompt) => {
    expect(parseBulkWorksheetDeletion(prompt)).toBeNull();
  });
});

describe("evaluateBulkWorksheetDeletion", () => {
  const workbookSheets = ["sheet-dc1", "恐怖游轮", "痴迷"];

  it("要求用户先选择所有涉及的工作表，不能自动扩大授权范围", () => {
    expect(
      evaluateBulkWorksheetDeletion(
        "清除sheet-dc1外的所有子表",
        workbookSheets,
        ["sheet-dc1"]
      )
    ).toEqual({
      kind: "selectionRequired",
      keeperName: "sheet-dc1",
      missingSheetNames: ["恐怖游轮", "痴迷"],
      targetSheetNames: ["恐怖游轮", "痴迷"]
    });
  });

  it("授权完整后只把保留表之外的工作表作为删除目标", () => {
    expect(
      evaluateBulkWorksheetDeletion(
        "删除除 SHEET-DC1 外的所有工作表",
        workbookSheets,
        workbookSheets
      )
    ).toEqual({
      kind: "ready",
      keeperName: "sheet-dc1",
      targetSheetNames: ["恐怖游轮", "痴迷"]
    });
  });

  it("保留表不存在时阻止生成删除计划", () => {
    expect(
      evaluateBulkWorksheetDeletion(
        "删除除 sheet-missing 外的所有工作表",
        workbookSheets,
        workbookSheets
      )
    ).toEqual({
      kind: "keeperMissing",
      requestedKeeperName: "sheet-missing"
    });
  });
});
