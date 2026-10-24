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
  /**
   * 标记还在，但上一次清理失败了。
   *
   * 它是"可重试状态"的唯一来源：界面据此只显示「重试清理标记」，
   * 而不是把整件事说成"确认失败、请重做"，更不会重新执行计划。
   */
  markerPending?: boolean;
};
export type ResultReviewReceipt = { executionId: string; verified: boolean; undoable: boolean };
/**
 * **当前性校验失败**（工作簿/授权/计划已变），与"标记清理失败"是两回事。
 *
 * 分开的意义：身份不成立的唯一出路是回到本次执行的原表再操作，此时谈
 * "重试清理标记"是误导——重试也清不掉"目标不对"这件事，还会让人以为
 * 确认只是收尾没做完。所以界面遇到这一类不显示「重试清理标记」。
 */
export class ReviewNotCurrentError extends Error {
  constructor(message = '工作簿、授权或计划已变化，请返回本次执行的原工作簿核对。') {
    super(message);
    this.name = 'ReviewNotCurrentError';
  }
}
export interface ResultReviewOperations<Before, Receipt extends ResultReviewReceipt, Marker, Copy> {
  /** Includes plan revision, authorization and the original workbook identity. */
  isCurrent(): boolean;
  prepare(): Promise<Before>;
  execute(before: Before): Promise<Receipt>;
  mark(receipt: Receipt): Promise<Marker>;
  clearMarker(marker: Marker): Promise<void>;
  showBefore(before: Before, receipt: Receipt): Promise<Copy>;
  comparisonNotice?(copy: Copy): string | null;
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
  private operations?: Omit<ResultReviewOperations<Before, Receipt, Marker, Copy>, "prepare" | "execute">;
  getSnapshot = () => this.state;
  subscribe = (listener: () => void) => { this.listeners.add(listener); return () => { this.listeners.delete(listener); }; };
  private update(change: Partial<ResultReviewState>) {
    this.state = { ...this.state, ...change };
    this.listeners.forEach(listener => listener());
  }
  private assertCurrent() {
    if (!this.operations?.isCurrent()) throw new ReviewNotCurrentError();
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
  /** Adopt an already completed, verified execution. No prepare/execute callback exists on this path. */
  async adoptCompleted(key: string, before: Before, receipt: Receipt,
    operations: Omit<ResultReviewOperations<Before, Receipt, Marker, Copy>, 'prepare' | 'execute'>): Promise<boolean> {
    if (!key || this.state.pending || this.state.phase === 'reviewing' || this.attempted.has(key)) return false;
    this.attempted.add(key);
    this.operations = operations; this.before = before; this.receipt = receipt;
    this.marker = undefined; this.copy = undefined;
    this.update({phase:'executing', key, executionId:receipt.executionId, pending:true,
      comparing:false, notice:null, marked:false, markerPending:false});
    try {
      this.assertCurrent();
      if (!receipt.executionId || !receipt.verified || !receipt.undoable) throw new Error('恢复结果或恢复点尚未验证，不能进入审阅。');
      await operations.validateResult?.(receipt);
      this.assertCurrent();
      try {this.marker = await operations.mark(receipt);this.update({marked:true});}
      catch(reason) {this.update({notice:`结果已写入，颜色标记未完整建立：${this.reason(reason)}`});}
      this.update({phase:'reviewing'}); return true;
    } catch(reason) {this.update({phase:'failed',notice:this.reason(reason)});return false;}
    finally {this.update({pending:false});}
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
      this.update({ comparing: true, notice: this.operations!.comparisonNotice?.(this.copy) ?? this.state.notice });
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
      this.update({ markerPending: false });
      this.assertCurrent();
      await this.operations!.validateResult?.(this.receipt);
      this.assertCurrent();
      this.operations!.accepted(this.receipt);
      this.update({ phase: 'confirmed', notice: null });
      await this.closeBefore();
      return true;
    } catch (reason) {
      /*
       * 确认这一步里有两类失败，**不能混成一句话**：
       *
       *  1. 当前性不成立（工作簿/授权/计划已变）——唯一的出路是回到本次执行的原表，
       *     再点一次确认。这时提"重试清理标记"是误导：重试只清标记，清不掉"目标不对"。
       *  2. 移除本次标记失败——这才是可重试状态：计划已经写完了，用户要的只是把标记
       *     撤下。不能报成"确认失败，请重新执行"。
       *
       * 所以只有第 2 类给出「重试清理标记」（`markerPending`），第 1 类只报原因。
       */
      const notCurrent = reason instanceof ReviewNotCurrentError;
      this.update({
        notice: notCurrent
          ? `确认未完成：${this.reason(reason)}`
          : `确认未完成：${this.reason(reason)} 可重试清理标记；重试只移除本次标记，不会再次写入。`,
        markerPending: notCurrent ? false : this.marker !== undefined
      });
      return false;
    }
    finally { this.update({ pending: false }); }
  }
  /**
   * 只重试**移除本次标记**。
   *
   * 与 `confirm()` 的区别：这里不重新校验结果、不再调用 `accepted`，也不会
   * 触发任何写入或撤销。它只处理"标记清理失败"这一种可重试状态。
   */
  async retryMarkerCleanup(key: string): Promise<boolean> {
    if (key !== this.state.key || this.state.pending || !this.receipt) return false;
    const phase = this.state.phase;
    if (phase !== 'reviewing' && phase !== 'confirmed' && phase !== 'discarded') return false;
    this.update({ pending: true });
    try {
      await this.clearMarker();
      // 已确认的保持 confirmed；已放弃/审阅中的保持在原阶段，只把提示清掉。
      this.update({ notice: null, markerPending: false });
      return true;
    } catch (reason) {
      this.update({ notice: `审阅标记仍未移除：${this.reason(reason)}` });
      return false;
    } finally { this.update({ pending: false }); }
  }
  /**
   * **外部入口**已经撤销了这次执行（例如结果卡上的「撤销本次执行」）。
   *
   * 为什么需要一条单独的通道：那个入口走的是宿主自己的恢复点，不是这里的
   * `discard()`。若不同步，固定操作条会继续显示「请查看主表中的执行结果」并
   * 提供「确认变更」——用户刚撤销完，界面却劝他确认一份已被还原的结果；
   * 「保存为工具」资格也可能仍然开着，等于把一次已撤销的写入固化成工具。
   *
   * 所以这里只做**状态同步**：不调用 `restore`（撤销已经发生，再调一次会二次恢复），
   * 也不再写表；清掉本次标记并把阶段推进到终态。
   */
  async settleExternally(executionId: string): Promise<boolean> {
    if (!executionId || this.state.executionId !== executionId) return false;
    if (this.state.phase !== 'reviewing' && this.state.phase !== 'confirmed') return false;
    if (this.state.pending) return false;
    this.update({ pending: true });
    try {
      try { await this.clearMarker(); this.update({ markerPending: false }); }
      catch (reason) { this.update({ notice: `审阅标记未能全部移除：${this.reason(reason)}`, markerPending: true }); }
      this.update({ phase: 'discarded', notice: this.state.notice ?? null });
      await this.closeBefore();
      return true;
    } finally { this.update({ pending: false }); }
  }
  async discard(key: string): Promise<boolean> {
    if (key !== this.state.key || this.state.phase !== 'reviewing' || this.state.pending || !this.receipt) return false;
    this.update({ pending: true });
    try {
      this.assertCurrent();
      /*
       * **先恢复，再清标记。**
       *
       * 顺序反过来会让"标记清理失败"把恢复整个拦下：用户点了放弃，主表却还是执行后
       * 的状态，而他看到的原因是"标记没清掉"。那两件事必须分开：
       *  - 恢复是业务结论，只能由真实撤销适配器宣告；
       *  - 标记清理是收尾，失败要单独说，并且**不能**把已完成的恢复改成失败。
       */
      await this.operations!.restore(this.receipt);
      this.update({ phase: 'discarded', notice: null });
      try { await this.clearMarker(); }
      catch (reason) {
        this.update({ notice: `主表已恢复到执行前；审阅标记未能全部移除：${this.reason(reason)} 可重试清理，不会重复恢复。`,
          markerPending: true });
      }
      await this.closeBefore();
      return true;
    } catch (reason) { this.update({ notice: `未能放弃本次变更：${this.reason(reason)}` }); return false; }
    finally { this.update({ pending: false }); }
  }
}
