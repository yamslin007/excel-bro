import capabilities from '../../../../config/capabilities.json';

/** Local-only display evidence; never sent to the model or used as undo material. */
type Entry<T> = { version: 1; host: string; identity: string; executionId: string; sheets: string[]; before: T };
const key = (host: string, identity: string, executionId: string) => JSON.stringify([host, identity, executionId]);
async function access<T>(mode: IDBTransactionMode, action: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  if (typeof indexedDB === 'undefined') throw new Error('本机快照存储不可用');
  const db = await new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open('excel_bro_review_before', 1);
    let blocked = false;
    request.onupgradeneeded = () => request.result.createObjectStore('snapshots');
    request.onblocked = () => { blocked = true; reject(new Error('本机快照存储被另一窗口占用')); };
    request.onsuccess = () => { if (blocked) request.result.close(); else resolve(request.result); };
    request.onerror = () => reject(request.error);
  });
  return new Promise((resolve, reject) => {
    const tx = db.transaction('snapshots', mode);
    const request = action(tx.objectStore('snapshots'));
    tx.oncomplete = () => { db.close(); resolve(request.result); };
    tx.onerror = tx.onabort = () => { db.close(); reject(tx.error ?? new Error('本机快照读写失败')); };
  });
}
export async function saveReviewBefore<T>(host: string, identity: string, executionId: string, sheets: readonly string[], before: T): Promise<boolean> {
  try {
    const entry: Entry<T> = { version: 1, host, identity, executionId, sheets: [...sheets].sort(), before };
    if (new TextEncoder().encode(JSON.stringify(entry)).length > capabilities.excelNativeComparison.maxBytes) return false;
    await access('readwrite', store => store.put(entry, key(host, identity, executionId)));
    return true;
  } catch { return false; }
}
export async function loadReviewBefore<T>(host: string, identity: string, executionId: string, sheets: readonly string[]): Promise<T | null> {
  try {
    const entry = await access('readonly', store => store.get(key(host, identity, executionId))) as Entry<T> | undefined;
    return entry?.version === 1 && entry.host === host && entry.identity === identity && entry.executionId === executionId &&
      JSON.stringify(entry.sheets) === JSON.stringify([...sheets].sort()) ? entry.before : null;
  } catch { return null; }
}
export async function deleteReviewBefore(host: string, identity: string, executionId: string): Promise<void> {
  await access('readwrite', store => store.delete(key(host, identity, executionId)));
}
