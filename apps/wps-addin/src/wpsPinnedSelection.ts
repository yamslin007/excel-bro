import { restorePinnedSelection, serializePinnedSelection, type RestoredSelection } from "@excel-bro/core/pinnedSheetSelection";
import { wpsWorkbookIdentity } from "./wpsExecutionLedger";
import type { WpsWorkbookMetadata } from "./wpsWorkbookReader";
const prefix = "excel-bro.wps.pinned-sheets.v1:";
export function loadWpsPinnedSelection(metadata: WpsWorkbookMetadata): RestoredSelection {
  const identity = wpsWorkbookIdentity(metadata.fullName);
  if (!identity) return { kind: "none" };
  try { return restorePinnedSelection(localStorage.getItem(prefix + identity), identity, metadata.sheets); }
  catch { return { kind: "invalid" }; }
}
export function saveWpsPinnedSelection(metadata: WpsWorkbookMetadata, mode: "auto" | "manual", indexes: readonly number[]): void {
  const identity = wpsWorkbookIdentity(metadata.fullName);
  if (!identity) return;
  if (mode === "auto") localStorage.removeItem(prefix + identity);
  else localStorage.setItem(prefix + identity, serializePinnedSelection(identity, metadata.sheets, indexes));
}
