import type { WpsObjectIdentity } from "./wpsObjectIdentity";

/** Private key for explicit pre-analysis target consent, NOT proof of native object identity. */
export function wpsConfirmationTargetKey(target: WpsObjectIdentity): string | null {
  if (!target.workbookFullName || !/^(?:[a-z]:[\\/]|\\\\)/i.test(target.workbookFullName) || !target.sheets.length) return null;
  // This key is private, memory only. Never render, persist or send it.
  return JSON.stringify([target.workbookFullName, target.workbookPath, target.workbookName,
    target.sheets.map(({ index, name, id }) => [index, name, id ?? null])]);
}

export function wpsConfirmationTargetLabel(target: WpsObjectIdentity): string {
  return `${target.workbookName} · ${target.sheets.map((sheet) => sheet.name).join("、")}`;
}
