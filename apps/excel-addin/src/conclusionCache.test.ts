import { describe, expect, it } from "vitest";
import { normalizeIntentKey, normalizePrompt } from "./conversation";
import type { IntentScopeContext, QueryTableArguments } from "./contracts";

function scope(sheets: string[]): IntentScopeContext {
  return {
    workbookName: "book.xlsx",
    sourceMode: "workbook",
    selectionMode: "auto",
    activeWorksheet: sheets[0] ?? "Sheet1",
    totalWorksheetCount: sheets.length,
    worksheetNames: sheets,
    sheets: sheets.map((name) => ({
      name,
      usedRange: `${name}!A1:B2`,
      rowCount: 2,
      columnCount: 2,
      headers: ["影片名称", "预售票房"]
    }))
  };
}

describe("normalizePrompt", () => {
  it("folds whitespace and case so identical asks share a key", () => {
    expect(normalizePrompt("  按 影片名称   分组 ")).toBe(
      normalizePrompt("按 影片名称 分组")
    );
  });

  it("distinguishes genuinely different asks", () => {
    expect(normalizePrompt("按影片分组")).not.toBe(
      normalizePrompt("按城市分组")
    );
  });
});

describe("normalizeIntentKey", () => {
  const base: QueryTableArguments = {
    mode: "aggregate",
    groupBy: ["影片名称"],
    metrics: [
      { operation: "sum", field: "预售票房", outputName: "总预售" }
    ],
    filters: [
      { field: "城市", operator: "equals", value: "北京" },
      { field: "状态", operator: "equals", value: "上映" }
    ],
    sortBy: "总预售",
    sortDirection: "desc",
    limit: 10
  };

  it("is idempotent for the same args and scope", () => {
    expect(normalizeIntentKey(base, scope(["票房"]))).toBe(
      normalizeIntentKey(base, scope(["票房"]))
    );
  });

  it("normalizes filter ordering to the same key", () => {
    const reordered: QueryTableArguments = {
      ...base,
      filters: [base.filters![1], base.filters![0]]
    };
    expect(normalizeIntentKey(reordered, scope(["票房"]))).toBe(
      normalizeIntentKey(base, scope(["票房"]))
    );
  });

  it("normalizes field case to the same key", () => {
    const cased: QueryTableArguments = {
      ...base,
      groupBy: ["影片名称"],
      sortBy: "总预售"
    };
    expect(normalizeIntentKey(cased, scope(["票房"]))).toBe(
      normalizeIntentKey(base, scope(["票房"]))
    );
  });

  it("differs when the limit changes", () => {
    expect(
      normalizeIntentKey({ ...base, limit: 20 }, scope(["票房"]))
    ).not.toBe(normalizeIntentKey(base, scope(["票房"])));
  });

  it("differs when the source sheets change", () => {
    expect(normalizeIntentKey(base, scope(["其他表"]))).not.toBe(
      normalizeIntentKey(base, scope(["票房"]))
    );
  });

  it("ignores source sheet ordering", () => {
    expect(normalizeIntentKey(base, scope(["A", "B"]))).toBe(
      normalizeIntentKey(base, scope(["B", "A"]))
    );
  });
});
