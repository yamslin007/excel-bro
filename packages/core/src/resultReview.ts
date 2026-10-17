/** Review an already executed result. Confirmation never executes the plan again.
 * Host data and undo handles stay in memory; only host adapters touch workbooks.
 */
export type ResultReviewPhase = 'idle' | 'executing' | 'reviewing' | 'confirmed' | 'discarded' | 'failed';
export type ResultReviewState = {
  phase: ResultReviewPhase;
  key: string | null;
  executionId: string | null;
  pending: boolean;
  comparing: boolean;
  notice: string | null;
  marked?: boolean;
};
export type ResultReviewReceipt = { executionId: string; verified: boolean; undoable: boolean };
export interface ResultReviewOperations<Before, Receipt extends ResultReviewReceipt, Marker, Copy> {
  /** Includes plan revision, authorization and the original workbook identity. */
  isCurrent(): boolean;
  prepare(): Promise<Before>;
  execute(before: Before): Promise<Receipt>;
  mark(receipt: Receipt): Promise<Marker>;
  clearMarker(marker: Marker): Promise<void>;
  showBefore(before: Before, receipt: Receipt): Promise<Copy>;
  closeBefore(copy: Copy): Promise<string | null>;
  /** A host may need a native before-copy even before dual view is opened. */
  disposePrepared?(before: Before): Promise<string | null>;
  restore(receipt: Receipt): Promise<void>;
  validateResult?(receipt: Receipt): Promise<void>;
  accepted(receipt: Receipt): void;
}

export class ResultReview<Before, Receipt extends ResultReviewReceipt, Marker, Copy> {
  private state: ResultReviewState = { phase: 'idle', key: null, executionId: null, pending: false, comparing: false, notice: null };
  private listeners = new Set<() => void>();
  private attempted = new Set<string>();
  private before?: Before;
  private receipt?: Receipt;
  private marker?: Marker;
  private copy?: Copy;
  private operations?: ResultReviewOperations<Before, Receipt, Marker, Copy>;
  getSnapshot = () => this.state;
  subscribe = (listener: () => void) => { this.listeners.add(listener); return () => { this.listeners.delete(listener); }; };
  private update(change: Partial<ResultReviewState>) {
    this.state = { ...this.state, ...change };
    this.listeners.forEach(listener => listener());
  }
  private assertCurrent() {
    if (!this.operations?.isCurrent()) throw new Error('工作簿、授权或计划已变化，请返回本次执行的原工作簿核对。');
  }
  private reason(reason: unknown) { return reason instanceof Error ? reason.message : '操作未完成，请核对当前工作簿。'; }
  async start(key: string, operations: ResultReviewOperations<Before, Receipt, Marker, Copy>): Promise<boolean> {
    if (!key || this.state.pending || this.state.phase === 'reviewing' || this.attempted.has(key)) return false;
    // Reserve before the first await, including failures: never silently replay writes.
    this.attempted.add(key);
    this.operations = operations;
    this.before = undefined; this.receipt = undefined; this.marker = undefined; this.copy = undefined;
    this.update({ phase: 'executing', key, pending: true, executionId: null, comparing: false, notice: null, marked: false });
    try {
      this.assertCurrent();
      this.before = await operations.prepare();
      this.assertCurrent();
      this.receipt = await operations.execute(this.before);
      this.update({ executionId: this.receipt.executionId || null });
      if (!this.receipt.executionId || !this.receipt.verified || !this.receipt.undoable) {
        throw new Error('执行结果或恢复点未完整验证，请查看执行记录；不会自动重试或显示已确认。');
      }
      // A marker failure must not hide a successful write or remove its undo entry.
      try { this.assertCurrent(); this.marker = await operations.mark(this.receipt); this.update({ marked: true }); }
      catch (reason) { this.update({ notice: `结果已写入，颜色标记未完整建立：${this.reason(reason)}` }); }
      this.update({ phase: 'reviewing' });
      return true;
    } catch (reason) {
      this.update({ phase: 'failed', notice: this.reason(reason) });
      return false;
    } finally { this.update({ pending: false }); }
  }
  private async clearMarker() {
    if (this.marker !== undefined) {
      await this.operations!.clearMarker(this.marker);
      this.marker = undefined;
      this.update({ marked: false });
    }
  }
  private async closeBefore() {
    if (this.copy === undefined && !this.operations!.disposePrepared) return;
    // Cleanup failure never changes the accepted/restored business outcome.
    try { this.update({ notice: this.copy !== undefined
      ? await this.operations!.closeBefore(this.copy)
      : await this.operations!.disposePrepared!(this.before!) }); }
    catch (reason) { this.update({ notice: `对比副本尚未关闭：${this.reason(reason)}` }); }
  }
  async compare(key: string): Promise<boolean> {
    if (key !== this.state.key || this.state.phase !== 'reviewing' || this.state.pending || this.state.comparing || !this.receipt) return false;
    this.update({ pending: true });
    try {
      this.assertCurrent();
      if (this.copy === undefined) this.copy = await this.operations!.showBefore(this.before!, this.receipt);
      this.assertCurrent();
      this.update({ comparing: true });
      return true;
    } catch (reason) { this.update({ notice: this.reason(reason) }); return false; }
    finally { this.update({ pending: false }); }
  }
  async confirm(key: string): Promise<boolean> {
    if (key !== this.state.key || this.state.phase !== 'reviewing' || this.state.pending || !this.receipt) return false;
    this.update({ pending: true });
    try {
      this.assertCurrent();
      await this.clearMarker();
      this.assertCurrent();
      await this.operations!.validateResult?.(this.receipt);
      this.assertCurrent();
      this.operations!.accepted(this.receipt);
      this.update({ phase: 'confirmed', notice: null });
      await this.closeBefore();
      return true;
    } catch (reason) { this.update({ notice: this.reason(reason) }); return false; }
    finally { this.update({ pending: false }); }
  }
  async discard(key: string): Promise<boolean> {
    if (key !== this.state.key || this.state.phase !== 'reviewing' || this.state.pending || !this.receipt) return false;
    this.update({ pending: true });
    try {
      this.assertCurrent();
      await this.clearMarker();
      this.assertCurrent();
      // Only the real undo adapter may declare restoration; a closed copy is no evidence.
      await this.operations!.restore(this.receipt);
      this.update({ phase: 'discarded', notice: null });
      await this.closeBefore();
      return true;
    } catch (reason) { this.update({ notice: `未能放弃本次变更：${this.reason(reason)}` }); return false; }
    finally { this.update({ pending: false }); }
  }
}
