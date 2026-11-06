/** Review an already executed result. Confirmation never executes the plan again.
 * Host data and undo handles stay in memory; only host adapters touch workbooks.
 */
import {
  copyCleanupNotice,
  copyCleanupSummary,
  dedupeReviewCopyCleanupIssues,
  type ReviewCopyCleanupIssue
} from './operationFeedback';
export type ResultReviewPhase = 'idle' | 'executing' | 'reviewing' | 'confirmed' | 'discarded' | 'not_started' | 'failed';
export type ResultReviewState = {
  phase: ResultReviewPhase;
  key: string | null;
  executionId: string | null;
  pending: boolean;
  comparing: boolean;
  comparisonSheets?: string[];
  comparisonSheet?: string;
  notice: string | null;
  /** Presentation only: routine scope and layout assistance, never operation errors. */
  comparisonHelp?: { scope?: string; layout?: string };
  marked?: boolean;
  /**
   * 标记还在，但上一次清理失败了。
   *
   * 它是"可重试状态"的唯一来源：界面据此只显示「重试清理标记」，
   * 而不是把整件事说成"确认失败、请重做"，更不会重新执行计划。
   */
  markerPending?: boolean;
  /**
   * 宿主侧的决定**已经成立**，但"这次审阅已经处理过"没有写进执行记录。
   *
   * 与 `markerPending` 分开：那个说的是颜色标记没撤下（纯收尾），这个说的是
   * **持久事实缺失** —— 此刻刷新页面，这条审阅会重新变成待处理。界面据此只说
   * "重试记录审阅结果"，既不改口说"确认失败"，也不再碰任何单元格。
   */
  decisionPending?: ResultReviewDecision | null;
  /**
   * **展示副本**没清理掉的结构化清单（按真实副本身份去重）。
   *
   * 与业务结论完全分开：这里非空只说明"对照用的展示副本还开着"，
   * **不表示**工作簿没恢复、也不表示确认/放弃失败。逐份原因由界面展开显示。
   * 判据与去重只在 `operationFeedback.ts`，本类只按引用写入。
   */
  copyCleanup?: {
    count: number;
    headline: string;
    issues: ReviewCopyCleanupIssue[];
  } | null;
};
/** 一次审阅决定：确认接受现有结果 / 放弃并恢复 / 确认之后又撤销。 */
export type ResultReviewDecision = 'confirmed' | 'discarded' | 'undone';
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
/** Host adapter proved no action started; never infer this from a missing receipt. */
export class ReviewNotStartedError extends Error {
  constructor(message: string) { super(message); this.name = 'ReviewNotStartedError'; }
}
export interface ResultReviewOperations<Before, Receipt extends ResultReviewReceipt, Marker, Copy> {
  /** Includes plan revision, authorization and the original workbook identity. */
  isCurrent(): boolean;
  prepare(): Promise<Before>;
  execute(before: Before): Promise<Receipt>;
  mark(receipt: Receipt): Promise<Marker>;
  clearMarker(marker: Marker): Promise<void>;
  comparisonSheets?(before: Before): string[];
  showBefore(before: Before, receipt: Receipt, sheet?: string): Promise<Copy>;
  activateComparison?(copy: Copy): Promise<void>;
  /** Switch tabs inside an existing multi-sheet copy. False means legacy single-sheet copy. */
  switchComparison?(copy: Copy, sheet: string): Promise<boolean>;
  /** Leave viewing without discarding the owned copy; re-entry reuses it. */
  suspendComparison?(copy: Copy): Promise<string | null>;
  saveComparisons?(receipt: Receipt, copies: Array<[string, Copy]>): Promise<void>;
  loadComparisons?(receipt: Receipt): Promise<Array<[string, Copy]>>;
  comparisonNotice?(copy: Copy): string | null;
  comparisonHelp?(copy: Copy): ResultReviewState["comparisonHelp"];
  /**
   * 这份副本的**真实身份**（本次会话随机 nonce + 会话标识，或宿主侧唯一副本名）。
   * 清理失败去重用它，绝不拿错误文本当身份。没给就退回"执行身份 + 该副本的登记键"，
   * 仍然是按副本去重，不是按原因去重。
   */
  copyIdentity?(copy: Copy): string | null;
  closeBefore(copy: Copy): Promise<string | null>;
  /** A host may need a native before-copy even before dual view is opened. */
  disposePrepared?(before: Before): Promise<string | null>;
  restore(receipt: Receipt): Promise<void>;
  validateResult?(receipt: Receipt): Promise<void>;
  /**
   * 把这次决定**落到持久账本**（服务端 `review_state`）。没有它，审阅状态就只活在
   * 这个类的内存里 —— 刷新/重开窗格后"还有一次审阅没处理"这件事会整段消失。
   *
   * 契约：抛错 = **没写成功**。调用方据此保留 `decisionPending` 并允许重试，
   * 绝不能在没写成功的情况下宣告"处理完了"；重试只重发这一条决定，不碰单元格。
   */
  settle?(receipt: Receipt, decision: ResultReviewDecision): Promise<void>;
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
  private copies = new Map<string, Copy>();
  /**
   * 本次审阅里"还没关掉的展示副本"：key = 真实副本身份。
   * 跨 confirm / retryDecision / settleExternally / discard 多次调用持久保留，
   * 重试成功才移除 —— 所以同一份副本不会被每个入口各报一次。
   */
  private cleanupIssues = new Map<string, ReviewCopyCleanupIssue>();
  /** 本类自己写下的那条副本清理提示，用来在收尾成功后只清自己的提示。 */
  private cleanupNotice?: string;
  private comparisonRestoreIssue?: string;
  private operations?: Omit<ResultReviewOperations<Before, Receipt, Marker, Copy>, "prepare" | "execute">;
  getSnapshot = () => this.state;
  /** The active in-memory copy is for host layout adapters only; never serialize it. */
  getComparisonCopy = (): Copy | undefined =>
    this.state.phase === 'reviewing' && this.state.comparing ? this.copy : undefined;

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
    this.copies.clear();
    this.cleanupIssues.clear();
    this.cleanupNotice = undefined;
    this.comparisonRestoreIssue = undefined;
    this.update({ phase: 'executing', key, pending: true, executionId: null, comparing: false, comparisonHelp: undefined, notice: null, marked: false, comparisonSheets: [], comparisonSheet: undefined, copyCleanup: null });
    let executionStarted = false;
    try {
      this.assertCurrent();
      this.before = await operations.prepare();
      this.update({ comparisonSheets: operations.comparisonSheets?.(this.before) ?? [] });
      this.assertCurrent();
      executionStarted = true;
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
      this.update({ phase: executionStarted && !(reason instanceof ReviewNotStartedError) ? 'failed' : 'not_started', notice: this.reason(reason) });
      return false;
    } finally { this.update({ pending: false }); }
  }
  /**
   * Adopt an already completed, verified execution. No prepare/execute callback exists on this path.
   *
   * `restored`（2026-11-03）：这次审阅是**刷新/重开后找回的**，不是刚跑完的。两处差别：
   *  - 结果与恢复材料的可用性已经由调用方用共享判据（`reviewRestoreVerdict`）核过一遍，
   *    这里不再用"回执自检"把它二次拦下 —— 否则一条真实存在的待审阅会因为"恢复材料
   *    缺一份"而在界面上彻底消失（正是要修的那个缺陷）；
   *  - 但 `executionId` 仍然必须存在：没有执行身份的审阅不能摆出来。
   * 允许进入审阅**不等于**允许确认：能不能确认由调用方按同一份判据决定并只暴露对应操作。
   */
  async adoptCompleted(key: string, before: Before, receipt: Receipt,
    operations: Omit<ResultReviewOperations<Before, Receipt, Marker, Copy>, 'prepare' | 'execute'>,
    options: { restored?: boolean } = {}): Promise<boolean> {
    if (!key || this.state.pending || this.state.phase === 'reviewing' || this.attempted.has(key)) return false;
    this.attempted.add(key);
    this.operations = operations; this.before = before; this.receipt = receipt;
    this.marker = undefined; this.copy = undefined;
    this.copies.clear();
    this.cleanupIssues.clear();
    this.cleanupNotice = undefined;
    this.comparisonRestoreIssue = undefined;
    this.update({phase:'executing', key, executionId:receipt.executionId, pending:true,
      comparing:false, comparisonHelp:undefined, notice:null, marked:false, markerPending:false, decisionPending:null,
      comparisonSheets: operations.comparisonSheets?.(before) ?? [], comparisonSheet: undefined, copyCleanup:null});
    try {
      this.assertCurrent();
      if (!receipt.executionId) throw new Error('这次审阅没有执行身份，不能进入审阅。');
      if (!options.restored && (!receipt.verified || !receipt.undoable)) {
        throw new Error('恢复结果或恢复点尚未验证，不能进入审阅。');
      }
      await operations.validateResult?.(receipt);
      this.assertCurrent();
      if (operations.loadComparisons) {
        try { this.copies = new Map(await operations.loadComparisons(receipt)); }
        catch (reason) {
          this.comparisonRestoreIssue = this.reason(reason);
          this.update({notice: this.comparisonRestoreIssue});
        }
      }
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
  /**
   * 一份展示副本的登记身份。宿主给了真实 nonce 就用它；没给则退回
   * 「执行身份 + 该副本在登记 Map 里的键」，仍然是**按副本**去重，不是按错误文本。
   */
  private cleanupIdentity(copy: Copy, sheet?: string): { identity: string; label?: string } {
    const fromHost = this.operations?.copyIdentity?.(copy)?.trim();
    if (fromHost) return { identity: fromHost };
    const execution = this.state.executionId ?? this.receipt?.executionId ?? 'unknown-execution';
    return { identity: `${execution}:${sheet ?? 'primary'}`, label: sheet ?? '本次展示副本' };
  }
  private recordCleanupIssue(identity: string, reason: string, label?: string) {
    this.cleanupIssues.set(identity, { identity, label, reason: reason || '未说明原因' });
  }
  /**
   * 把当前未关闭副本集合折成**一条**紧凑提示 + 结构化清单。
   * 多次入口重复触发只更新同一份清单，不会把同一句原因越拼越长。
   */
  private publishCleanup(): boolean {
    const issues = dedupeReviewCopyCleanupIssues([...this.cleanupIssues.values()]);
    if (issues.length === 0) {
      const ownNotice = this.cleanupNotice !== undefined && this.state.notice === this.cleanupNotice;
      this.cleanupNotice = undefined;
      this.update({ copyCleanup: null, ...(ownNotice ? { notice: null } : {}) });
      return false;
    }
    const summary = copyCleanupSummary({
      executionId: this.state.executionId ?? this.receipt?.executionId ?? '',
      issues
    });
    const notice = copyCleanupNotice({ ...summary, issues: summary.issues });
    this.cleanupNotice = notice;
    this.update({
      copyCleanup: { count: summary.count, headline: summary.headline, issues: summary.issues },
      notice
    });
    return true;
  }
  private async closeBefore() {
    if (this.state.decisionPending) return;
    // Cleanup failure never changes the accepted/restored business outcome.
    if (this.operations?.suspendComparison) {
      const groups = new Map<string, { copy: Copy; label?: string; keys: string[] }>();
      for (const [sheet, copy] of this.copies) {
        const { identity, label } = this.cleanupIdentity(copy, sheet);
        const group = groups.get(identity);
        if (group) group.keys.push(sheet);
        else groups.set(identity, { copy, label, keys: [sheet] });
      }
      for (const [identity, {copy, label, keys}] of groups) {
        try {
          const issue = await this.operations.closeBefore(copy);
          if (issue) this.recordCleanupIssue(identity, issue, label);
          else { this.cleanupIssues.delete(identity); keys.forEach(key => this.copies.delete(key)); }
        } catch (reason) { this.recordCleanupIssue(identity, this.reason(reason), label); }
      }
      try {
        await this.operations.saveComparisons?.(this.receipt!, [...this.copies]);
        this.cleanupIssues.delete('comparison-registry');
      } catch (reason) { this.recordCleanupIssue('comparison-registry', this.reason(reason), '副本登记更新'); }
      if (this.publishCleanup()) return;
      this.copy = undefined;
      this.update({ comparing: false });
      return;
    }
    if (this.copy === undefined && !this.operations!.disposePrepared) return;
    const hasCopy = this.copy !== undefined;
    const { identity, label } = this.cleanupIdentity(
      (hasCopy ? this.copy : this.before) as Copy,
      hasCopy ? undefined : 'prepared'
    );
    try {
      const issue = hasCopy
        ? await this.operations!.closeBefore(this.copy!)
        : await this.operations!.disposePrepared!(this.before!);
      if (issue) this.recordCleanupIssue(identity, issue, label);
      else this.cleanupIssues.delete(identity);
    } catch (reason) { this.recordCleanupIssue(identity, this.reason(reason), label); }
    if (!this.publishCleanup()) { this.copy = undefined; this.update({ comparing: false }); }
  }
  /** 只写"这次决定"这一条持久事实；没有适配器 = 没有持久账本（不假装写过）。 */
  private async persistDecision(decision: ResultReviewDecision): Promise<void> {
    if (!this.operations?.settle || !this.receipt) return;
    await this.operations.settle(this.receipt, decision);
  }
  async compare(key: string, sheet?: string): Promise<boolean> {
    if (key !== this.state.key || this.state.phase !== 'reviewing' || this.state.pending || !this.receipt) return false;
    if (this.comparisonRestoreIssue) { this.update({notice: this.comparisonRestoreIssue}); return false; }
    const target = sheet ?? this.state.comparisonSheet ?? this.state.comparisonSheets?.[0];
    if (sheet && !this.state.comparisonSheets?.includes(sheet)) {
      this.update({ notice: '所选工作表不属于本次结果对比。' }); return false;
    }
    if (this.state.comparing && target === this.state.comparisonSheet && !this.operations?.activateComparison) return false;
    this.update({ pending: true });
    try {
      this.assertCurrent();
      if (target && this.operations!.switchComparison) {
        const candidates = new Set([...(this.copy === undefined ? [] : [this.copy]), ...this.copies.values()]);
        for (const retained of candidates) {
          if (!await this.operations!.switchComparison(retained, target)) continue;
          this.copy = retained;
          this.assertCurrent();
          this.update({ comparing: true, comparisonSheet: target,
            comparisonHelp: this.operations!.comparisonHelp?.(retained), notice: this.operations!.comparisonNotice?.(retained) ?? null });
          return true;
        }
      }
      if (this.copy !== undefined && (target !== this.state.comparisonSheet || this.operations!.switchComparison)) {
        const issue = await (this.operations!.suspendComparison ?? this.operations!.closeBefore)(this.copy);
        if (issue) throw new Error(`暂未切换对比：${issue}`);
        this.copy = undefined;
        this.update({ comparing: false, comparisonSheet: undefined });
        this.assertCurrent();
      }
      if (this.copy === undefined) {
        // A new adapter upgrades legacy per-sheet copies from the saved before-image.
        // Retain old handles until settlement; never overwrite and lose ownership.
        const cached = this.operations!.switchComparison ? undefined : this.copies.get(target ?? '');
        this.copy = cached ?? await this.operations!.showBefore(this.before!, this.receipt, target);
        if (this.operations!.suspendComparison) {
          const previous = this.copies.get(target ?? '');
          if (previous !== undefined && previous !== this.copy)
            this.copies.set(`__legacy__:${this.copies.size}:${target ?? ''}`, previous);
          this.copies.set(target ?? '', this.copy);
        }
        if (cached === undefined) await this.operations!.saveComparisons?.(this.receipt, [...this.copies]);
        this.update({ comparisonSheet: target });
      }
      // Keep ownership even when arranging windows fails, so retry/close can still handle the copy.
      await this.operations!.activateComparison?.(this.copy);
      this.assertCurrent();
      this.update({ comparing: true, comparisonHelp: this.operations!.comparisonHelp?.(this.copy), notice: this.operations!.comparisonNotice?.(this.copy) ?? null });
      return true;
    } catch (reason) { this.update({ notice: this.reason(reason) }); return false; }
    finally { this.update({ pending: false }); }
  }
  async exitComparison(key: string): Promise<boolean> {
    if (this.state.pending || this.state.phase !== 'reviewing') return false;
    if (key !== this.state.key || this.copy === undefined) {
      this.update({ notice: key !== this.state.key
        ? '未退出双视图：审阅任务已变化，请回到原任务后重试。'
        : '未退出双视图：找不到本次对比副本的安全关闭记录，请检查副本后手动关闭。' });
      return false;
    }
    this.update({ pending: true });
    try {
      this.assertCurrent();
      const issue = await (this.operations!.suspendComparison ?? this.operations!.closeBefore)(this.copy);
      if (issue) { this.update({ notice: `未退出双视图：${issue}` }); return false; }
      if (!this.operations!.suspendComparison) this.copy = undefined;
      this.update({ comparing: false, notice: null });
      return true;
    } catch (reason) {
      this.update({ notice: `未退出双视图：${this.reason(reason)}` });
      return false;
    } finally { this.update({ pending: false }); }
  }
  async confirm(key: string): Promise<boolean> {
    if (this.state.pending) return false;
    if (this.state.phase !== 'reviewing') return false;
    if (key !== this.state.key || !this.receipt) {
      this.update({ notice: key !== this.state.key
        ? '确认未完成：此按钮对应的任务已变化，请使用当前任务的审阅入口。'
        : '确认未完成：本次执行回执不可用，请保留当前工作簿并核对执行记录，不要重复执行。' });
      return false;
    }
    this.update({ pending: true });
    let cleaningMarker = false;
    try {
      this.assertCurrent();
      cleaningMarker = true;
      await this.clearMarker();
      cleaningMarker = false;
      this.update({ markerPending: false });
      this.assertCurrent();
      await this.operations!.validateResult?.(this.receipt);
      this.assertCurrent();
      this.operations!.accepted(this.receipt);
      this.update({ phase: 'confirmed', notice: null });
      /*
       * 宿主侧已经接受了结果，**从这一刻起单元格就是"已确认"的**，不再回退。
       * 但"这次审阅处理过"还得落进执行记录：写不进去时只能诚实说"记录没写成"，
       * 并留一个只重发这条决定的入口（`retryDecision`）。绝不能把它说成
       * "确认失败、请重新执行"，也绝不能假装已经落盘。
       */
      try {
        await this.persistDecision('confirmed');
        this.update({ decisionPending: null });
      } catch (reason) {
        this.update({
          decisionPending: 'confirmed',
          notice: `结果已接受，但审阅状态没有写进执行记录：${this.reason(reason)}`
        });
        await this.closeBefore();
        return false;
      }
      await this.closeBefore();
      return true;
    } catch (reason) {
      /*
       * 按实际失败阶段分类；结果核对或接受失败不能冒充标记清理失败。
       *
       *  1. 当前性不成立（工作簿/授权/计划已变）——唯一的出路是回到本次执行的原表，
       *     再点一次确认。这时提"重试清理标记"是误导：重试只清标记，清不掉"目标不对"。
       *  2. 移除本次标记失败——这才是可重试状态：计划已经写完了，用户要的只是把标记
       *     撤下。不能报成"确认失败，请重新执行"。
       *
       * 只有清理标记阶段给出「重试清理标记」（`markerPending`），其他阶段保留原始原因。
       */
      const markerFailed = cleaningMarker && !(reason instanceof ReviewNotCurrentError);
      this.update({
        notice: markerFailed
          ? `确认未完成：${this.reason(reason)} 可重试清理标记；重试只移除本次标记，不会再次写入。`
          : `确认未完成：${this.reason(reason)}`,
        markerPending: markerFailed && this.marker !== undefined
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
   * 只重试**把这次决定写进执行记录**。
   *
   * 与 `confirm()` / `discard()` 的区别：宿主侧的决定已经成立，这里不再清标记、
   * 不再校验结果、不再执行、也**不再恢复**。重试的永远只是那一条持久事实。
   */
  async retryCopyCleanup(key: string): Promise<boolean> {
    if (key !== this.state.key || this.state.pending || this.state.decisionPending ||
        !['confirmed', 'discarded'].includes(this.state.phase)) return false;
    this.update({ pending: true });
    try { await this.closeBefore(); return !this.state.copyCleanup; }
    finally { this.update({ pending: false }); }
  }
  async retryDecision(key: string): Promise<boolean> {
    const decision = this.state.decisionPending;
    if (!decision || key !== this.state.key || this.state.pending) return false;
    this.update({ pending: true });
    try {
      await this.persistDecision(decision);
      this.update({ decisionPending: null, notice: null });
      await this.closeBefore();
      return true;
    } catch (reason) {
      this.update({ notice: `审阅状态仍未写进执行记录：${this.reason(reason)}` });
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
      /*
       * 「撤销」也是审阅的一种终态：确认过再撤销，这条审阅照样不该在刷新后
       * 重新冒出来。写失败只留可重试项，不改判已经发生的撤销。
       */
      try {
        await this.persistDecision('undone');
        this.update({ decisionPending: null });
      } catch (reason) {
        this.update({
          decisionPending: 'undone',
          notice: `本次执行已撤销；审阅状态没有写进执行记录：${this.reason(reason)} 可重试记录，不会重复撤销。`
        });
      }
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
      /*
       * 恢复是业务结论，**已经发生**；"放弃"这条持久事实写不进去是另一件事。
       * 后者失败只留可重试的 `decisionPending` —— 重试只重发这条决定，
       * 不会回到宿主再恢复一次（那会把用户此后的编辑一起抹掉）。
       */
      try {
        await this.persistDecision('discarded');
        this.update({ decisionPending: null });
      } catch (reason) {
        this.update({
          decisionPending: 'discarded',
          notice: `主表已恢复到执行前；审阅状态没有写进执行记录：${this.reason(reason)} 可重试记录，不会重复恢复。`
        });
      }
      await this.closeBefore();
      return true;
    } catch (reason) { this.update({ notice: `未能放弃本次变更：${this.reason(reason)}` }); return false; }
    finally { this.update({ pending: false }); }
  }
}
