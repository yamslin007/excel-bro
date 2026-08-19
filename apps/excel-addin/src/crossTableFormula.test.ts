import { describe, expect, it } from "vitest";
import type { FormulaExtraSheet } from "./contracts";
import {
  buildCrossTableFormulas,
  buildCrossTableProposal,
  columnsFromRange
} from "./crossTableFormula";

const sheetB: FormulaExtraSheet = {
  sourceFile: "B.xlsx",
  sourcePath: "子目录/B.xlsx",
  sheetName: "Sheet2",
  headers: ["编号", "名称", "单价"],
  columns: ["A", "B", "C"],
  sampleRows: [["1", "苹果", "3.5"]],
  rowCount: 7
};

describe("columnsFromRange", () => {
  it("maps headers to column letters from usedRange start", () => {
    expect(columnsFromRange(["编号", "名称"], "A1:D20")).toEqual([
      { name: "编号", letter: "A" },
      { name: "名称", letter: "B" }
    ]);
    expect(columnsFromRange(["编号", "名称"], "C2:E9")).toEqual([
      { name: "编号", letter: "C" },
      { name: "名称", letter: "D" }
    ]);
    expect(columnsFromRange([], null)).toEqual([]);
  });
});

describe("buildCrossTableProposal", () => {
  const mainColumns = [
    { name: "编号", letter: "A" },
    { name: "名称", letter: "B" }
  ];

  it("detects cross-table match with description-mentioned value", () => {
    const proposal = buildCrossTableProposal(
      "把B表的单价按编号匹配过来",
      mainColumns,
      [sheetB]
    );
    expect(proposal).not.toBeNull();
    expect(proposal!.externalFile).toBe("B.xlsx");
    expect(proposal!.externalSheet).toBe("Sheet2");
    expect(proposal!.selectedKey).toBe("编号");
    expect(proposal!.selectedValue).toBe("单价");
    expect(proposal!.keyCandidates[0]).toEqual({
      name: "编号",
      mainLetter: "A",
      externalLetter: "A"
    });
  });

  it("falls back to first external-only column when description has no header", () => {
    const proposal = buildCrossTableProposal("把单价匹配过来", mainColumns, [
      sheetB
    ]);
    expect(proposal).not.toBeNull();
    expect(proposal!.selectedValue).toBe("单价");
  });

  it("returns null without a common key column", () => {
    const proposal = buildCrossTableProposal(
      "匹配一下",
      [{ name: "项目", letter: "A" }],
      [sheetB]
    );
    expect(proposal).toBeNull();
  });

  it("returns null without extra sheets", () => {
    expect(buildCrossTableProposal("匹配单价", mainColumns, [])).toBeNull();
  });

  it("picks the strongest signal among multiple sheets", () => {
    const sheetC: FormulaExtraSheet = {
      ...sheetB,
      sourceFile: "C.xlsx",
      sheetName: "Sheet3",
      headers: ["编号", "库存"],
      columns: ["A", "B"],
      sampleRows: [],
      rowCount: 3
    };
    const proposal = buildCrossTableProposal(
      "把B表的单价按编号匹配过来",
      mainColumns,
      [sheetC, sheetB]
    );
    expect(proposal!.externalFile).toBe("B.xlsx");
    expect(proposal!.selectedValue).toBe("单价");
  });
});

describe("buildCrossTableFormulas", () => {
  const proposal = buildCrossTableProposal(
    "把B表的单价按编号匹配过来",
    [
      { name: "编号", letter: "A" },
      { name: "名称", letter: "B" }
    ],
    [sheetB]
  )!;

  it("builds VLOOKUP and XLOOKUP for key before value", () => {
    const pair = buildCrossTableFormulas(proposal, "E2");
    expect(pair.compatFormula).toBe(
      "=VLOOKUP(A2,'[B.xlsx]Sheet2'!$A$2:$C$7,3,FALSE)"
    );
    expect(pair.modernFormula).toBe(
      "=XLOOKUP(A2,'[B.xlsx]Sheet2'!$A$2:$A$7,'[B.xlsx]Sheet2'!$C$2:$C$7,\"\")"
    );
    expect(pair.compatFormula.startsWith("=")).toBe(true);
    expect(pair.modernFormula.startsWith("=")).toBe(true);
  });

  it("uses INDEX+MATCH when value column is left of key column", () => {
    const reversed: FormulaExtraSheet = {
      ...sheetB,
      headers: ["单价", "编号", "名称"],
      columns: ["A", "B", "C"]
    };
    const reversedProposal = buildCrossTableProposal(
      "把B表的单价按编号匹配过来",
      [
        { name: "编号", letter: "A" },
        { name: "名称", letter: "B" }
      ],
      [reversed]
    )!;
    const pair = buildCrossTableFormulas(reversedProposal, "E2");
    expect(pair.compatFormula).toBe(
      "=INDEX('[B.xlsx]Sheet2'!$A$2:$A$7,MATCH(A2,'[B.xlsx]Sheet2'!$B$2:$B$7,0))"
    );
    expect(pair.modernFormula).toBe(
      "=XLOOKUP(A2,'[B.xlsx]Sheet2'!$B$2:$B$7,'[B.xlsx]Sheet2'!$A$2:$A$7,\"\")"
    );
  });

  it("throws when proposal is empty", () => {
    expect(() =>
      buildCrossTableFormulas(
        {
          ...proposal,
          keyCandidates: [],
          valueCandidates: []
        },
        "E2"
      )
    ).toThrow(/参数不完整/);
  });
});
