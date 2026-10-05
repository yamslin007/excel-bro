import type { CellValue } from "./contracts";
import { parseExplicitA1Rect } from "./planCore";

export type ColumnMappingFormulaCandidate = {
  formulas: string[][];
  expected: string[][];
};

function columnName(index: number): string {
  let name = "";
  for (let value = index; value > 0; value = Math.floor((value - 1) / 26)) {
    name = String.fromCharCode(65 + (value - 1) % 26) + name;
  }
  return name;
}

function literal(value: string): string {
  return `"${value.replaceAll('"', '""')}"`;
}

/**
 * A deliberately bounded formula candidate, not a translation of the entire mapping language.
 * It covers every declared whole-cell category and the current cell's deterministic result.
 * Unknown future combinations return #N/A instead of silently guessing. The host must still
 * calculate this formula in a disposable copy and read back `expected` before it is offered.
 */
export function columnMappingFormulaCandidate(input: {
  sourceRange: string;
  sourceValues: readonly (readonly CellValue[])[];
  expected: readonly (readonly string[])[];
  mapping: readonly { source: string; target: string }[];
  maxCases: number;
  maxCharacters: number;
}): ColumnMappingFormulaCandidate | null {
  const rect = parseExplicitA1Rect(input.sourceRange);
  if (!rect || input.mapping.length > input.maxCases || input.maxCases < 1 || input.maxCharacters < 1 ||
      input.sourceValues.length !== rect.endRow - rect.startRow + 1 ||
      input.sourceValues.length !== input.expected.length ||
      input.sourceValues.some((row, index) =>
        row.length !== rect.endColumn - rect.startColumn + 1 || row.length !== input.expected[index]?.length
      )) return null;
  const cases = new Map<string, string>();
  for (const pair of input.mapping) {
    if (!pair.source || pair.target.trimStart().startsWith("=") ||
        /[\u0000-\u001f]/u.test(pair.source) || /[\u0000-\u001f]/u.test(pair.target) ||
        (cases.has(pair.source) && cases.get(pair.source) !== pair.target)) return null;
    cases.set(pair.source, pair.target);
  }
  const formulas: string[][] = [];
  for (let row = 0; row < input.sourceValues.length; row += 1) {
    const formulaRow: string[] = [];
    for (let column = 0; column < input.sourceValues[row].length; column += 1) {
      const source = input.sourceValues[row][column];
      // Blank and non-text coercion differ between JS and spreadsheet formulas; use fixed values.
      if (typeof source !== "string" || source.trim() === "" || /[\u0000-\u001f]/u.test(source)) return null;
      const current = input.expected[row][column];
      if (typeof current !== "string") return null;
      const cell = `${columnName(rect.startColumn + column)}${rect.startRow + row}`;
      const branches = new Map(cases);
      if (branches.has(source) && branches.get(source) !== current) return null;
      branches.set(source, current);
      if (branches.size > input.maxCases) return null;
      let formula = "NA()";
      for (const [key, target] of [...branches].reverse()) {
        formula = `IF(EXACT(${cell},${literal(key)}),${literal(target)},${formula})`;
      }
      formula = `=${formula}`;
      if (formula.length > input.maxCharacters) return null;
      formulaRow.push(formula);
    }
    formulas.push(formulaRow);
  }
  return { formulas, expected: input.expected.map((row) => [...row]) };
}
