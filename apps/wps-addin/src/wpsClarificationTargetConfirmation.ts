import type { WpsObjectIdentity } from "./wpsObjectIdentity";
import { authorizedTargetEvidenceKey } from "@excel-bro/core/authorizedTarget";

/**
 * Private key for explicit pre-analysis target consent, NOT proof of native object identity.
 *
 * 与 Excel 侧共用共享层的键构造：授权那一维只比**集合**，不比点选顺序。
 * 工作表身份（索引 / 名称 / 稳定 ID）仍然逐个参与，一张变了就换键。
 */
export function wpsConfirmationTargetKey(target: WpsObjectIdentity): string | null {
  if (!target.workbookFullName || !/^(?:[a-z]:[\\/]|\\\\)/i.test(target.workbookFullName) || !target.sheets.length) return null;
  // This key is private, memory only. Never render, persist or send it.
  return authorizedTargetEvidenceKey(
    JSON.stringify([target.workbookFullName, target.workbookPath, target.workbookName]),
    target.sheets.map(({ index, name, id }) => JSON.stringify([index, name, id ?? null]))
  );
}

export function wpsConfirmationTargetLabel(target: WpsObjectIdentity): string {
  return `${target.workbookName} · ${target.sheets.map((sheet) => sheet.name).join("、")}`;
}
