import capabilities from '../../../config/capabilities.json';
import type { WpsApplication } from './wpsHost';

export const TASK_PANE_STORAGE_KEY = 'excel_bro_taskpane_id';
export const REVIEW_PANE_LEASE_KEY = 'excel_bro_review_pane_width_v1';
type Lease = { paneId: string; owner: string; expiresAt: number };
export function readReviewPaneLease(app: WpsApplication, includeExpired = false): Lease | undefined {
  try {
    const raw = app.PluginStorage.getItem(REVIEW_PANE_LEASE_KEY);
    const value = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (value && typeof value.paneId === 'string' && typeof value.owner === 'string' &&
        Number.isFinite(value.expiresAt) && (includeExpired || value.expiresAt > Date.now())) return value;
  } catch { /* A broken marker cannot grant pane ownership. */ }
  return undefined;
}
export function refreshReviewPaneLease(app: WpsApplication, paneId: string, owner: string): void {
  const live = readReviewPaneLease(app);
  if (live && live.owner !== owner) throw new Error('Pane belongs to another review');
  // Only geometry ownership; never workbook identity, plan, formulas or cell data.
  app.PluginStorage.setItem(REVIEW_PANE_LEASE_KEY, JSON.stringify({
    paneId, owner, expiresAt: Date.now() + capabilities.wpsReviewPane.leaseTtlMs
  }));
  if (readReviewPaneLease(app)?.owner !== owner) throw new Error('Pane marker could not be saved');
}
