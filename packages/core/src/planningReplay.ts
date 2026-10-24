/** Recovery cursors describe consumed events, never the server's latest event. */
export function planningReceivedSequence(
  taskId: string,
  records: readonly { taskId?: string | null; sequence?: number | null }[]
): number {
  return records.reduce((cursor, record) => record.taskId === taskId &&
    Number.isSafeInteger(record.sequence) && record.sequence! >= 0
    ? Math.max(cursor, record.sequence!) : cursor, 0);
}

/** Share the same identity/duplicate gate for polling and streamed recovery. */
export function acceptsPlanningReplay(
  taskId: string, cursor: number,
  event: { operationId?: string | null; sequence?: number }
): boolean {
  return event.operationId === taskId && Number.isSafeInteger(event.sequence) && event.sequence! > cursor;
}

/** Cancel the wait as well as the underlying fetch; some host transports ignore abort. */
export function waitForPlanning<T>(work: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return work;
  return new Promise<T>((resolve, reject) => {
    const abort = () => reject(new DOMException("Aborted", "AbortError"));
    const cleanup = () => signal.removeEventListener("abort", abort);
    // Always observe work, including when already aborted (no unhandled rejection).
    work.then((value) => { cleanup(); signal.aborted ? abort() : resolve(value); },
      (error) => { cleanup(); reject(error); });
    if (signal.aborted) abort();
    else signal.addEventListener("abort", abort, { once: true });
  });
}
