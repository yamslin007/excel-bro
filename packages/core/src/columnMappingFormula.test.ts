import { describe, expect, it } from "vitest";
import { columnMappingFormulaCandidate } from "./columnMappingFormula";

const base = {
  sourceRange: "D2:D4",
  sourceValues: [["甲"], ["乙"], ["甲，乙"]],
  expected: [["低"], ["高"], ["高"]],
  mapping: [{ source: "甲", target: "低" }, { source: "乙", target: "高" }],
  maxCases: 32,
  maxCharacters: 2000
};

describe("columnMappingFormulaCandidate", () => {
  it("引用对应行的源格，覆盖已知整格与当前组合；未知组合明确返回 NA", () => {
    const candidate = columnMappingFormulaCandidate(base);
    expect(candidate?.expected).toEqual(base.expected);
    expect(candidate?.formulas[0][0]).toContain('EXACT(D2,"甲")');
    expect(candidate?.formulas[1][0]).toContain('EXACT(D3,"乙")');
    expect(candidate?.formulas[2][0]).toContain('EXACT(D4,"甲，乙")');
    expect(candidate?.formulas[2][0]).toContain('"高",NA()');
  });

  it("文字字面量转义，不把引号变成公式语法", () => {
    const candidate = columnMappingFormulaCandidate({
      ...base, sourceRange: "AA7", sourceValues: [['a"b']], expected: [['x"y']],
      mapping: [{ source: 'a"b', target: 'x"y' }]
    });
    expect(candidate?.formulas).toEqual([['=IF(EXACT(AA7,"a""b"),"x""y",NA())']]);
  });

  it("空白、数值、控制字符、过多分支或超长公式一律退回固定值", () => {
    expect(columnMappingFormulaCandidate({ ...base, sourceValues: [[null], ["乙"], ["甲，乙"]] })).toBeNull();
    expect(columnMappingFormulaCandidate({ ...base, sourceValues: [[1], ["乙"], ["甲，乙"]] })).toBeNull();
    expect(columnMappingFormulaCandidate({ ...base, sourceValues: [["甲\n"], ["乙"], ["甲，乙"]] })).toBeNull();
    expect(columnMappingFormulaCandidate({ ...base, maxCases: 2 })).toBeNull();
    expect(columnMappingFormulaCandidate({ ...base, maxCharacters: 20 })).toBeNull();
    expect(columnMappingFormulaCandidate({
      ...base, mapping: [{ source: "甲", target: "=1+1" }, { source: "乙", target: "高" }]
    })).toBeNull();
  });
});
