/**
 * 列映射澄清 → 澄清卡的接线锁（Office 通道）。
 *
 * 这里钉三件接线上的事，它们都是"看起来能跑、但结果是错的"：
 *
 * 1. **归因决定选项**。分隔符已对上的情形若冒出「指定其他分隔符」，
 *    用户会去改一个本来没问题的设置。
 * 2. **`resolutionMode` 必须是 `append`**。选完动作后是往原需求后面
 *    追加一句补充，不是替换 —— 替换掉原需求会丢失映射表、区域、
 *    空白标签这些第二轮还要用的上下文。
 * 3. **选项 id 就是协议枚举值**。回传路径直接把它当 `resolvedBy` 用，
 *    若这里改成序号，回传就会带一个服务端认不出的值。
 *
 * 文案不逐字钉：它随措辞调整，钉死会让每次改字都要改测试。
 */

import { describe, expect, it } from "vitest";

import type { ColumnMappingClarification } from "./contracts";
import { columnMappingClarificationMessage } from "./columnMappingClarification";

function message(clarification: ColumnMappingClarification) {
  return columnMappingClarificationMessage({
    clarification,
    scopeLabel: "订单!B2:B200",
    originalPrompt: "把类别列翻译成归属",
    scopeFingerprint: "scope-1",
    round: 1
  });
}

describe("列映射澄清接进既有澄清卡", () => {
  it("分隔符已对上：只给补类别映射一个选项，且带上缺的类别", () => {
    const card = message({
      reason: "matchedDelimiterMissingCategories",
      message: "以下类别不在映射表内。",
      missingCategories: ["丙", "丁"],
      unresolvedCellCount: 0,
      resolvedBy: ["addCategories"]
    });
    expect(card.options.map((option) => option.id)).toEqual(["addCategories"]);
    // 缺的类别出现在 resolution 里，用户点了才知道要补什么。
    expect(card.options[0].resolution).toContain("丙");
    expect(card.options[0].resolution).toContain("丁");
  });

  it("归因不确定：三个选项都在，且不声称缺哪些类别", () => {
    const card = message({
      reason: "unresolvedSeparatorOrMapping",
      message: "有 2 个来源格无法归类。",
      missingCategories: [],
      unresolvedCellCount: 2,
      resolvedBy: [
        "specifyOtherDelimiter",
        "addCategories",
        "treatWholeCellAsCategory"
      ]
    });
    expect(card.options.map((option) => option.id)).toEqual([
      "specifyOtherDelimiter",
      "addCategories",
      "treatWholeCellAsCategory"
    ]);
    for (const option of card.options) {
      expect(option.resolution).not.toContain("需要补充的类别");
    }
  });

  it("每个选项都以追加方式补充原需求，不替换它", () => {
    const card = message({
      reason: "unresolvedSeparatorOrMapping",
      message: "有 1 个来源格无法归类。",
      missingCategories: [],
      unresolvedCellCount: 1,
      resolvedBy: [
        "specifyOtherDelimiter",
        "addCategories",
        "treatWholeCellAsCategory"
      ]
    });
    for (const option of card.options) {
      expect(option.resolutionMode).toBe("append");
    }
  });

  it("卡片带齐续聊需要的上下文，且初始为待答", () => {
    const card = message({
      reason: "matchedDelimiterMissingCategories",
      message: "有类别不在映射表内。",
      missingCategories: ["丙"],
      unresolvedCellCount: 0,
      resolvedBy: ["addCategories"]
    });
    expect(card.status).toBe("pending");
    expect(card.originalPrompt).toBe("把类别列翻译成归属");
    expect(card.scopeFingerprint).toBe("scope-1");
    expect(card.round).toBe(1);
    expect(card.scopeLabel).toBe("订单!B2:B200");
    expect(card.summary.trim().length).toBeGreaterThan(0);
    expect(card.question.trim().length).toBeGreaterThan(0);
    expect(card.reason.trim().length).toBeGreaterThan(0);
  });

  it("卡片里不含单元格原文", () => {
    const card = message({
      reason: "unresolvedSeparatorOrMapping",
      message: "有 1 个来源格无法归类。",
      missingCategories: [],
      unresolvedCellCount: 1,
      resolvedBy: [
        "specifyOtherDelimiter",
        "addCategories",
        "treatWholeCellAsCategory"
      ]
    });
    expect(JSON.stringify(card)).not.toContain("甲/乙");
  });
});
