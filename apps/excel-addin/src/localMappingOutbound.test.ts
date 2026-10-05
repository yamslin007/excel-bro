/**
 * 本地列映射通道的出站面：源值不上传。
 *
 * 与 WPS 的 `wpsLocalPlanOutbound.test.tsx`（fetch 层拦截真实字节）不同，这里
 * 钉的是 Excel 侧**构造规划请求的那一刻**：`planningReadPolicy: "none"` 时，
 * 送出去的快照必须已经剥掉数据行。
 *
 * 为什么这两件事不能只留一个：端到端拦截证明"某一次真实跑动没泄漏"，但它依赖
 * 完整 UI 与 Office 宿主替身，跑得慢、也容易因为别的原因提前失败而空过。这里
 * 直接对"将要出站的请求体"做判定，快、稳，且与前者互补 —— 前者管真实链路，
 * 后者管构造规则本身。
 *
 * 正向锚点与负向断言并存：表头（字段名）**允许**出站，否则"不含源值"就会因为
 * 请求体为空而恒真。
 */

import { describe, expect, it } from "vitest";

import { stripDataRows, buildColumnMappingPlan } from "./columnMappingPlan";
import type { ColumnMappingRule, WorkbookSnapshot } from "./contracts";

/** 一条只存在于数据行里的私有哨兵：绝不允许出现在出站快照中。 */
const PRIVATE_SENTINEL = "内部代号-9K2W-仅存在于数据行";
const SECOND_SENTINEL = "内部代号-4H8N-仅存在于数据行";

function snapshot(): WorkbookSnapshot {
  return {
    name: "book.xlsx",
    capturedAt: "2026-09-28T00:00:00Z",
    activeWorksheet: "数据",
    selectedRange: null,
    worksheets: [
      {
        name: "数据",
        usedRange: "A1:B3",
        rowCount: 3,
        columnCount: 2,
        headers: ["来源字段", "目标字段"],
        dataRows: [
          [PRIVATE_SENTINEL, null],
          [SECOND_SENTINEL, null]
        ],
        truncated: false
      }
    ]
  };
}

describe("planningReadPolicy=none：出站快照必须剥掉数据行", () => {
  it("数据行清空、结构保留（表头/区域/行列数照旧）", () => {
    const stripped = stripDataRows(snapshot());
    const sheet = stripped.worksheets[0];
    expect(sheet.dataRows).toEqual([]);
    // 结构不能一起被抹掉：服务端要靠它判越界、产规则。
    expect(sheet.headers).toEqual(["来源字段", "目标字段"]);
    expect(sheet.usedRange).toBe("A1:B3");
    expect(sheet.rowCount).toBe(3);
    expect(sheet.columnCount).toBe(2);
  });

  it("私有哨兵不出现在出站快照的序列化结果里，但表头会出现", () => {
    const outbound = JSON.stringify(stripDataRows(snapshot()));
    // 负向：数据行里的私有值不上传。
    expect(outbound).not.toContain(PRIVATE_SENTINEL);
    expect(outbound).not.toContain(SECOND_SENTINEL);
    expect(outbound).not.toContain("9K2W");
    expect(outbound).not.toContain("4H8N");
    // 正向锚点：表头是字段名，允许出站 —— 证明上面不是"整个请求体为空"。
    expect(outbound).toContain("来源字段");
    expect(outbound).toContain("目标字段");
  });

  it("不改动原快照（纯函数，原对象仍带数据行）", () => {
    const original = snapshot();
    stripDataRows(original);
    expect(original.worksheets[0].dataRows).toHaveLength(2);
  });
});

/**
 * 澄清路径的出站面。
 *
 * 服务端判定"这条规则算不出正确答案"时不会产计划，只回一份带 `clarification`
 * 的规则。这条路径**在构造上就不可能**把源值送出去 —— 本端根本不进入
 * `buildColumnMappingPlan`（那是要读源值才会走到的分支）。这里用真实序列化
 * 结果钉住它，免得日后有人为了"顺手也展示一下"把源值读出来再回传。
 */
describe("澄清路径：不读源值、也不产出计划", () => {
  function clarificationRule(): ColumnMappingRule {
    return {
      kind: "columnMapping",
      sourceSheet: "数据",
      targetSheet: "数据",
      sourceRange: "A2:A2",
      targetRange: "B2:B2",
      mapping: [{ source: "已知类别", target: "目标值" }],
      blankLabel: "未分类",
      delimiters: [" "],
      valueRank: [{ value: "已知类别", rank: 1 }],
      clarification: {
        reason: "unresolvedSeparatorOrMapping",
        message: "有 1 个来源格无法归类。",
        missingCategories: [],
        unresolvedCellCount: 1,
        resolvedBy: [
          "specifyOtherDelimiter",
          "addCategories",
          "treatWholeCellAsCategory"
        ]
      }
    };
  }

  it("带着澄清的规则永远算不出计划，因此没有任何矩阵可被发送", () => {
    const sheet = {
      name: "数据",
      startRow: 1,
      startColumn: 1,
      columnCount: 2,
      values: [
        ["来源", "目标"],
        [PRIVATE_SENTINEL, null]
      ]
    };
    // 归因已确定"算不出正确答案"：算出来的矩阵必然是猜的。
    expect(() =>
      buildColumnMappingPlan(clarificationRule(), sheet, {
        summary: "s",
        warnings: [],
        assumptions: []
      })
    ).toThrow();
  });

  it("澄清内容序列化后不含来源格原文", () => {
    const outbound = JSON.stringify(clarificationRule().clarification);
    expect(outbound).not.toContain(PRIVATE_SENTINEL);
    expect(outbound).not.toContain("9K2W");
    // 正向锚点：澄清本身确实被序列化了，上面的负向断言不是"对象为空"。
    expect(outbound).toContain("unresolvedSeparatorOrMapping");
    expect(outbound).toContain("specifyOtherDelimiter");
  });

  it("剥掉数据行的快照里也没有哨兵，同时表头仍然出站（正向锚点）", () => {
    const stripped = stripDataRows(snapshot());
    const outbound = JSON.stringify(stripped);
    expect(outbound).not.toContain(PRIVATE_SENTINEL);
    expect(outbound).not.toContain(SECOND_SENTINEL);
    // 证明不是"整个请求体为空"。
    expect(outbound).toContain("来源字段");
    expect(outbound).toContain("目标字段");
  });
});
