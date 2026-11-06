import { describe, expect, it } from "vitest";
import {
  completenessOf,
  describeIncompleteReason,
  describeIncompleteReasons,
  describeTruncation
} from "./queryCompleteness";

/**
 * "为什么不完整"的文案与派生。
 *
 * 原来结果卡把任何 `complete===false` 都写成「结果不完整（扫描已达上限）」——
 * 2026-11-01 真机上三张表缺「列2」被说成扫描触顶，用户当然以为是读崩了。
 * 这里钉住：文案只复述事实，`complete` 只由原因列表决定，截断是另一件事。
 */
describe("不完整原因", () => {
  it("缺字段说清哪张表、缺哪个字段", () => {
    expect(
      describeIncompleteReason({
        kind: "missing_field",
        sheet: "结果1",
        fields: ["列2"]
      })
    ).toBe("「结果1」没有字段：列2");
  });

  it("没有数据行是另一种原因，不能和缺字段混成一句", () => {
    expect(describeIncompleteReason({ kind: "empty_sheet", sheet: "汇总表" })).toBe(
      "「汇总表」没有数据行"
    );
  });

  it("多张表缺字段时逐条列出，并说明一共几张表没计入", () => {
    const text = describeIncompleteReasons([
      { kind: "missing_field", sheet: "结果1", fields: ["列2"] },
      { kind: "missing_field", sheet: "结果2", fields: ["列2"] },
      { kind: "missing_field", sheet: "结果3", fields: ["列2"] }
    ]);
    expect(text).toContain("「结果1」没有字段：列2");
    expect(text).toContain("「结果3」没有字段：列2");
    expect(text).toContain("共 3 张");
    // 这句是本次缺陷的核心：缺字段不等于扫描触顶。
    expect(text).not.toContain("扫描");
  });

  it("没有原因就没有文案：不许替完整结果编一句提示", () => {
    expect(describeIncompleteReasons([])).toBeNull();
  });

  it("截断只说数量，不冒充不完整", () => {
    expect(
      describeTruncation({ shownRows: 20, totalRows: 137, limit: 20 })
    ).toBe("仅显示前 20 行，共 137 行（limit=20）。");
  });
});

describe("完整性事实", () => {
  it("complete 只由原因决定；截断不改变它", () => {
    expect(completenessOf({}).complete).toBe(true);
    expect(
      completenessOf({
        reasons: [{ kind: "missing_field", sheet: "结果1", fields: ["列2"] }]
      }).complete
    ).toBe(false);
    const truncated = completenessOf({
      truncation: { shownRows: 20, totalRows: 137, limit: 20 }
    });
    expect(truncated.complete).toBe(true);
    expect(truncated.truncated).toBe(true);
  });
});
