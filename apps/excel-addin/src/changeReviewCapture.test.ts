import { describe, expect, it } from "vitest";
import { canClaimWriteSuccess } from "@excel-bro/core/changeReview";
import { changeReviewStatusFromVerification, toChangeFormulas } from "./excel";
import capabilities from "../../../config/capabilities.json";

/**
 * Excel 侧双栏对照的可纯测部分。
 *
 * `executePlan` 本身依赖 Office.js 运行时，本仓没有可落地的 stub，
 * 所以这里只锁住两件最容易被"写顺手了"破坏的事：
 * 1. 什么才算"原公式"；2. 验收没过时不许让双栏落到可宣称成功的状态。
 */

describe("Excel 双栏 · 公式识别", () => {
  it("只有以 = 开头的字符串才算原公式，其余一律 null", () => {
    const formulas = toChangeFormulas([
      ["=SUM(A1:A2)", "普通文本", 5, null] as never
    ]);
    expect(formulas[0]).toEqual(["=SUM(A1:A2)", null, null, null]);
  });

  it("撤销快照里非公式单元格存的是值，不得被当成公式展示", () => {
    // 5 是数字常量、""2026-01-01"" 是文本常量，都不是公式。
    const formulas = toChangeFormulas([
      ["5", "2026-01-01"] as never
    ]);
    expect(formulas[0]).toEqual([null, null]);
  });
});

describe("Excel 双栏 · 验收结论决定能否宣称成功", () => {
  it("verified → applied（唯一允许宣称成功的状态）", () => {
    const status = changeReviewStatusFromVerification("verified");
    expect(status).toBe("applied");
    expect(canClaimWriteSuccess(status)).toBe(true);
  });

  it("failed → readback_mismatch，不得宣称成功", () => {
    const status = changeReviewStatusFromVerification("failed");
    expect(status).toBe("readback_mismatch");
    expect(canClaimWriteSuccess(status)).toBe(false);
  });

  it("executed_unverified 保留自身语义，也不得宣称成功", () => {
    // 前后矩阵是真的，但独立验收没核完 —— 双栏展示真值，但不替验收背书。
    const status = changeReviewStatusFromVerification("executed_unverified");
    expect(status).toBe("executed_unverified");
    expect(canClaimWriteSuccess(status)).toBe(false);
  });
});

describe("Excel 双栏 · 限额来自配置而不是硬编码", () => {
  it("capabilities 里存在双栏限额且为正数", () => {
    expect(capabilities.changeReview.maxRows).toBeGreaterThan(0);
    expect(capabilities.changeReview.maxTotalCells).toBeGreaterThan(0);
    expect(capabilities.changeReview.captureMaxCells).toBeGreaterThan(0);
  });
});
