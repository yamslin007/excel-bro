/**
 * 澄清卡的**归因 → 可选项**映射锁。
 *
 * 这里钉的不是文案，而是两件会静默出错的语义：
 *
 * 1. **归因确定时不给"指定其他分隔符"**。分隔符已经声明且对上了，
 *    还让用户去改分隔符设置，等于把一个不存在的问题指给他 ——
 *    用户改完之后情况不会变好，只会更困惑。
 * 2. **卡片里不带单元格原文**。卡片内容会进聊天历史/导出，
 *    而原文属于本端数据。
 *
 * 文案本身（label/description）刻意不做逐字断言：它们会随措辞调整，
 * 钉死会让每次改字都要改测试，而真正不能变的是"给了哪几个选项"。
 */

import { describe, expect, it } from "vitest";

import type { ColumnMappingClarification } from "./columnMapping";
import {
  columnMappingClarificationCard,
  columnMappingResolutionsFor
} from "./columnMappingClarification";

function clarification(
  overrides: Partial<ColumnMappingClarification> = {}
): ColumnMappingClarification {
  return {
    reason: "matchedDelimiterMissingCategories",
    message: "有类别不在映射表内。",
    missingCategories: ["丙"],
    unresolvedCellCount: 0,
    resolvedBy: ["addCategories"],
    ...overrides
  };
}

describe("归因决定界面能提供哪些补救动作", () => {
  it("分隔符已对上时，只给补类别映射这一条", () => {
    const card = columnMappingClarificationCard(clarification());
    expect(card.choices.map((choice) => choice.resolution)).toEqual([
      "addCategories"
    ]);
    // 关键反面断言：不能出现"指定其他分隔符"。
    expect(card.choices.map((choice) => choice.resolution)).not.toContain(
      "specifyOtherDelimiter"
    );
  });

  it("归因不确定时，三个选项都给全且按推荐顺序", () => {
    const card = columnMappingClarificationCard(
      clarification({
        reason: "unresolvedSeparatorOrMapping",
        missingCategories: [],
        unresolvedCellCount: 2,
        resolvedBy: [
          "specifyOtherDelimiter",
          "addCategories",
          "treatWholeCellAsCategory"
        ]
      })
    );
    expect(card.choices.map((choice) => choice.resolution)).toEqual([
      "specifyOtherDelimiter",
      "addCategories",
      "treatWholeCellAsCategory"
    ]);
    expect(card.unresolvedCellCount).toBe(2);
  });

  it("可选项只由归因推导，不听信传入的 resolvedBy", () => {
    // 若界面直接信这个字段，一个错写的值就能让"分隔符已对上"的卡片
    // 冒出"指定其他分隔符"—— 正是本模块要防的误导。
    const card = columnMappingClarificationCard(
      clarification({
        reason: "matchedDelimiterMissingCategories",
        resolvedBy: [
          "specifyOtherDelimiter",
          "addCategories",
          "treatWholeCellAsCategory"
        ]
      })
    );
    expect(card.choices.map((choice) => choice.resolution)).toEqual([
      "addCategories"
    ]);
  });

  it("归因确定时把无法处理的格数归零，避免夸大规模", () => {
    const card = columnMappingClarificationCard(
      clarification({ unresolvedCellCount: 99 })
    );
    expect(card.unresolvedCellCount).toBe(0);
  });

  it("卡片里不含单元格原文，只带类别名与计数", () => {
    const card = columnMappingClarificationCard(
      clarification({
        reason: "unresolvedSeparatorOrMapping",
        message: "有 1 个来源格无法归类：可能是用了未声明的分隔符。",
        missingCategories: [],
        unresolvedCellCount: 1
      })
    );
    const rendered = JSON.stringify(card);
    // 原文不出现在卡片任何字段里。
    expect(rendered).not.toContain("甲/乙");
  });

  it("每个选项都带用户可读的文案，且回传值与协议枚举一致", () => {
    const card = columnMappingClarificationCard(
      clarification({
        reason: "unresolvedSeparatorOrMapping",
        missingCategories: [],
        unresolvedCellCount: 1,
        resolvedBy: [
          "specifyOtherDelimiter",
          "addCategories",
          "treatWholeCellAsCategory"
        ]
      })
    );
    for (const choice of card.choices) {
      expect(choice.label.trim().length).toBeGreaterThan(0);
      expect(choice.description.trim().length).toBeGreaterThan(0);
    }
    expect(card.choices.map((choice) => choice.resolution)).toEqual(
      columnMappingResolutionsFor("unresolvedSeparatorOrMapping")
    );
  });

  it("回传列表与卡片可选项同源，不会出现界面给了却回传不了的选项", () => {
    for (const reason of [
      "matchedDelimiterMissingCategories",
      "unresolvedSeparatorOrMapping"
    ] as const) {
      const card = columnMappingClarificationCard(clarification({ reason }));
      expect(card.choices.map((choice) => choice.resolution)).toEqual(
        columnMappingResolutionsFor(reason)
      );
    }
  });
});
