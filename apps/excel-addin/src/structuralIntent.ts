export interface BulkWorksheetDeletionIntent {
  keeperName: string;
}

export type BulkWorksheetDeletionEvaluation =
  | { kind: "notMatched" }
  | { kind: "keeperMissing"; requestedKeeperName: string }
  | {
      kind: "selectionRequired";
      keeperName: string;
      missingSheetNames: string[];
      targetSheetNames: string[];
    }
  | { kind: "nothingToDelete"; keeperName: string }
  | { kind: "ready"; keeperName: string; targetSheetNames: string[] };

const variantCharacters: Record<string, string> = {
  "刪": "删",
  "餘": "余",
  "僅": "仅",
  "頁": "页",
  "錶": "表"
};

const sheetNoun = String.raw`(?:工作表|子表|sheet(?:页|表)?)`;
const deleteVerb = String.raw`(?:删除|清除|移除)`;
const trailingPunctuation = String.raw`[。.!！?？]*`;
const exceptKeeperPattern = new RegExp(
  String.raw`^\s*(?:请\s*)?${deleteVerb}\s*(?:除(?:了)?\s*)?(.+?)\s*(?:以外|之外|外)\s*(?:的)?\s*(?:所有|全部)?\s*(?:其他|其它|其余|剩余)?\s*(?:所有|全部)?\s*${sheetNoun}\s*${trailingPunctuation}\s*$`,
  "iu"
);
const keepThenDeletePattern = new RegExp(
  String.raw`^\s*(?:请\s*)?(?:只|仅)?\s*(?:保留|留下)\s*(.+?)\s*[,，;；]?\s*(?:并|然后|再)?\s*${deleteVerb}\s*(?:所有|全部)?\s*(?:其余|其他|其它|剩余)\s*(?:所有|全部)?\s*(?:的)?\s*${sheetNoun}\s*${trailingPunctuation}\s*$`,
  "iu"
);
const contextualDeletePattern = new RegExp(
  String.raw`^\s*(?:请\s*)?(?:将|把)?\s*(?:本表|当前工作簿|工作簿)?\s*(?:中|里)?\s*(?:除(?:了)?\s*)?(.+?)\s*(?:以外|之外|外)\s*(?:的)?\s*(?:所有|全部)?\s*(?:其他|其它)?\s*${sheetNoun}\s*(?:都|全部|全)?\s*${deleteVerb}\s*(?:掉)?\s*${trailingPunctuation}\s*$`,
  "iu"
);

function normalizeVariants(value: string): string {
  return [...value]
    .map((character) => variantCharacters[character] ?? character)
    .join("");
}

function cleanKeeperName(value: string): string {
  return value
    .trim()
    .replace(/^["'“”‘’`「」『』【】，,;；\s]+|["'“”‘’`「」『』【】，,;；\s]+$/gu, "");
}

function sheetKey(value: string): string {
  return value.trim().toLocaleLowerCase();
}

export function parseBulkWorksheetDeletion(
  prompt: string
): BulkWorksheetDeletionIntent | null {
  const normalized = normalizeVariants(prompt);
  for (const pattern of [exceptKeeperPattern, keepThenDeletePattern, contextualDeletePattern]) {
    const match = pattern.exec(normalized);
    if (!match) continue;
    const keeperName = cleanKeeperName(match[1] ?? "");
    if (keeperName) return { keeperName };
  }
  return null;
}

export function evaluateBulkWorksheetDeletion(
  prompt: string,
  workbookSheetNames: string[],
  authorizedSheetNames: string[]
): BulkWorksheetDeletionEvaluation {
  const intent = parseBulkWorksheetDeletion(prompt);
  if (!intent) return { kind: "notMatched" };

  const keeperKey = sheetKey(intent.keeperName);
  const keeperName = workbookSheetNames.find(
    (name) => sheetKey(name) === keeperKey
  );
  if (!keeperName) {
    return {
      kind: "keeperMissing",
      requestedKeeperName: intent.keeperName
    };
  }

  const targetSheetNames = workbookSheetNames.filter(
    (name) => sheetKey(name) !== sheetKey(keeperName)
  );
  if (targetSheetNames.length === 0) {
    return { kind: "nothingToDelete", keeperName };
  }

  const authorized = new Set(authorizedSheetNames.map(sheetKey));
  const missingSheetNames = workbookSheetNames.filter(
    (name) => !authorized.has(sheetKey(name))
  );
  if (missingSheetNames.length > 0) {
    return {
      kind: "selectionRequired",
      keeperName,
      missingSheetNames,
      targetSheetNames
    };
  }
  return { kind: "ready", keeperName, targetSheetNames };
}
