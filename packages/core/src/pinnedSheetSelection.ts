/** Saved selection metadata only; never a substitute for live authorization checks. */
export type SheetIdentity = { index: number; name: string };
export type RestoredSelection = { kind: "none" } | { kind: "invalid" } | { kind: "fixed"; indexes: number[] };
export function restorePinnedSelection(raw: string | null, identity: string, sheets: readonly SheetIdentity[]): RestoredSelection {
  if (raw === null) return { kind: "none" };
  try {
    const record = JSON.parse(raw);
    if (record.version !== 1 || record.identity !== identity || !Array.isArray(record.sheets)) return { kind: "invalid" };
    const indexes: number[] = [];
    for (const s of record.sheets) {
      if (!s || !Number.isInteger(s.index) || typeof s.name !== "string" || indexes.includes(s.index) ||
          sheets.filter(live => live.index === s.index && live.name === s.name).length !== 1) return { kind: "invalid" };
      indexes.push(s.index);
    }
    return { kind: "fixed", indexes };
  } catch { return { kind: "invalid" }; }
}
export function serializePinnedSelection(identity: string, sheets: readonly SheetIdentity[], indexes: readonly number[]): string {
  const selected = indexes.map(index => {
    const matches = sheets.filter(s => s.index === index);
    if (matches.length !== 1) throw new Error("工作表已变化，不能保存固定选择。");
    return { index, name: matches[0].name };
  });
  return JSON.stringify({ version: 1, identity, sheets: selected });
}
