// 跨表匹配确定性公式：不调模型、不发数据，仅凭表头结构与描述关键词本地构造。
import type { CellValue, FormulaExtraSheet } from "./contracts";

export interface MainColumn {
  name: string;
  letter: string;
}

export interface CrossTableColumnCandidate {
  name: string;
  letter: string;
}

export interface CrossTableKeyCandidate {
  name: string;
  mainLetter: string;
  externalLetter: string;
}

export interface CrossTableMatchProposal {
  externalFile: string;
  externalPath: string;
  externalSheet: string;
  externalRowCount: number;
  keyCandidates: CrossTableKeyCandidate[];
  valueCandidates: CrossTableColumnCandidate[];
  selectedKey: string;
  selectedValue: string;
}

export interface GeneratedFormulaPair {
  modernFormula: string;
  modernExplanation: string;
  compatFormula: string;
  compatExplanation: string;
}

function columnIndexFromLetter(letter: string): number {
  let index = 0;
  for (const ch of letter.toUpperCase()) {
    index = index * 26 + (ch.charCodeAt(0) - 64);
  }
  return index - 1;
}

function columnLetterFromIndex(index: number): string {
  let n = index;
  let letters = "";
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return letters;
}

function startColumnOfRange(address: string | null): string | null {
  if (!address) return null;
  const bare = address.includes("!") ? address.split("!").pop()! : address;
  const match = bare.match(/^\$?([A-Za-z]{1,3})/);
  return match ? match[1].toUpperCase() : null;
}

/** 从表头 + usedRange 推导"列头 → 列字母"映射。 */
export function columnsFromRange(
  headers: CellValue[],
  usedRange: string | null
): MainColumn[] {
  const startCol = startColumnOfRange(usedRange);
  const startIndex = startCol ? columnIndexFromLetter(startCol) : 0;
  return headers.map((header, index) => ({
    name: String(header ?? ""),
    letter: columnLetterFromIndex(startIndex + index)
  }));
}

/**
 * 本地判断是否命中跨表匹配，并生成预填提案。
 * 键列 = 两张表共有的列头；取值列 = 外部表列头（描述里提到的最优先）。
 * 命中至少一个键列 + 一个取值列才返回提案；多张外部表取信号最强的。
 */
export function buildCrossTableProposal(
  description: string,
  mainColumns: MainColumn[],
  extraSheets: FormulaExtraSheet[]
): CrossTableMatchProposal | null {
  const mainNames = new Set(mainColumns.map((column) => column.name));
  const mainLetterByName = new Map(
    mainColumns.map((column) => [column.name, column.letter] as const)
  );
  let best: { score: number; proposal: CrossTableMatchProposal } | null = null;

  for (const sheet of extraSheets) {
    const externalLetterByName = new Map(
      sheet.headers.map(
        (header, index) => [header, sheet.columns[index]] as const
      )
    );
    const keyCandidates: CrossTableKeyCandidate[] = sheet.headers
      .map((header) => ({
        name: header,
        mainLetter: mainLetterByName.get(header) ?? "",
        externalLetter: externalLetterByName.get(header) ?? ""
      }))
      .filter(
        (candidate) =>
          mainNames.has(candidate.name) &&
          candidate.mainLetter !== "" &&
          candidate.externalLetter !== ""
      );
    if (keyCandidates.length === 0) continue;

    const selectedKey =
      keyCandidates.find((candidate) =>
        description.includes(candidate.name)
      )?.name ?? keyCandidates[0].name;
    // 取值列候选必须排除当前匹配键列（键列不能同时是返回值）。
    const allValues: CrossTableColumnCandidate[] = sheet.headers
      .map((header) => ({
        name: header,
        letter: externalLetterByName.get(header) ?? ""
      }))
      .filter(
        (candidate) =>
          candidate.letter !== "" && candidate.name !== selectedKey
      );
    if (allValues.length === 0) continue;

    const mentionedValues = allValues.filter((candidate) =>
      description.includes(candidate.name)
    );
    // 取值列候选排序：描述提到的 → 外部表独有 → 两表共有。
    const valueCandidates = [
      ...mentionedValues,
      ...allValues.filter(
        (candidate) =>
          !mainNames.has(candidate.name) &&
          !mentionedValues.includes(candidate)
      ),
      ...allValues.filter(
        (candidate) =>
          mainNames.has(candidate.name) &&
          !mentionedValues.includes(candidate)
      )
    ];
    const selectedValue = valueCandidates[0]?.name;
    if (!selectedValue) continue;

    const score =
      mentionedValues.length * 2 +
      keyCandidates.filter((candidate) =>
        description.includes(candidate.name)
      ).length +
      (valueCandidates.some((candidate) => !mainNames.has(candidate.name))
        ? 1
        : 0);
    const proposal: CrossTableMatchProposal = {
      externalFile: sheet.sourceFile,
      externalPath: sheet.sourcePath,
      externalSheet: sheet.sheetName,
      externalRowCount: sheet.rowCount,
      keyCandidates,
      valueCandidates,
      selectedKey,
      selectedValue
    };
    if (!best || score > best.score) best = { score, proposal };
  }

  return best?.proposal ?? null;
}

/** 外部工作表引用串：统一带引号（'[B.xlsx]Sheet2'），兼容空格等特殊字符。 */
function externalSheetRef(file: string, sheet: string): string {
  return `'[${file}]${sheet}'`;
}

/** 本地拼两版跨表公式：兼容版 VLOOKUP / INDEX+MATCH，现代版 XLOOKUP。 */
export function buildCrossTableFormulas(
  match: CrossTableMatchProposal,
  firstCell: string
): GeneratedFormulaPair {
  const key =
    match.keyCandidates.find(
      (candidate) => candidate.name === match.selectedKey
    ) ?? match.keyCandidates[0];
  const value =
    match.valueCandidates.find(
      (candidate) => candidate.name === match.selectedValue
    ) ?? match.valueCandidates[0];
  if (!key || !value) {
    throw new Error("跨表匹配参数不完整，请重新确认匹配键与取值列。");
  }
  const rowMatch = firstCell.match(/\d+$/);
  const row = rowMatch ? rowMatch[0] : "2";
  const keyCell = `${key.mainLetter}${row}`;
  const sheetRef = externalSheetRef(match.externalFile, match.externalSheet);
  const endRow = Math.max(2, match.externalRowCount);
  const keyIndex = columnIndexFromLetter(key.externalLetter);
  const valueIndex = columnIndexFromLetter(value.letter);

  // VLOOKUP 要求键在区域首列；取值列在键右侧才用 VLOOKUP，否则用 INDEX+MATCH。
  const compatFormula =
    valueIndex >= keyIndex
      ? `=VLOOKUP(${keyCell},${sheetRef}!$${key.externalLetter}$2:$${value.letter}$${endRow},${valueIndex - keyIndex + 1},FALSE)`
      : `=INDEX(${sheetRef}!$${value.letter}$2:$${value.letter}$${endRow},MATCH(${keyCell},${sheetRef}!$${key.externalLetter}$2:$${key.externalLetter}$${endRow},0))`;
  const modernFormula = `=XLOOKUP(${keyCell},${sheetRef}!$${key.externalLetter}$2:$${key.externalLetter}$${endRow},${sheetRef}!$${value.letter}$2:$${value.letter}$${endRow},"")`;

  return {
    modernFormula,
    modernExplanation: `按「${key.name}」从 ${match.externalFile} 的「${match.externalSheet}」匹配「${value.name}」；用 XLOOKUP 精确匹配，查不到返回空。`,
    compatFormula,
    compatExplanation: `按「${key.name}」从 ${match.externalFile} 的「${match.externalSheet}」匹配「${value.name}」；用 VLOOKUP/INDEX+MATCH 精确匹配，查不到返回 #N/A。`
  };
}
