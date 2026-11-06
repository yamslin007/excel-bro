import { useEffect, useRef, useState } from "react";
import { checkAuthorizedTarget, type AuthorizedTarget, type TargetNoticeReason } from "@excel-bro/core/authorizedTarget";
import {
  targetDiagnostic,
  type TargetDiagnostic,
  type TargetDiagnosticHost,
  type TargetDiagnosticInput
} from "@excel-bro/core/targetDiagnostics";
import capabilities from "../../../../config/capabilities.json";

export function useAuthorizedTarget(options: {
  enabled: boolean;
  inspect: () => Promise<AuthorizedTarget>;
  adopt: (target: AuthorizedTarget) => void;
  /** A known review copy must not be adopted as an original by this guard. */
  reviewActive?: boolean;
  monitorPaused?: boolean;
  selectionPending?: () => boolean;
  isReviewCopy?: (target: AuthorizedTarget) => boolean;
  /** Host evidence for analysis/read authorization only, never permission to write.
   * Keys are private memory bindings. Consistent weak evidence requires consent;
   * missing or conflicting evidence remains blocked until a successful recheck.
   * `documentBound` is a second grade for hosts that cannot keep object wrappers
   * stable across a reload (WPS): same file path + host-declared sheet ids. It is
   * accepted here only — write/recovery/data-read gates still read `conclusive`. */
  analysisEvidence?: (target: AuthorizedTarget) => Promise<{
    key: string; conclusive: boolean; documentBound?: boolean;
    /** 只用于诊断的宿主观测事实（稳定 ID 可读程度 / 有无路径身份），不参与判定。 */
    diagnostic?: TargetDiagnosticHost;
  } | null>;
  /**
   * 每次范围判定把**为什么**交出去一次（只有布尔值/原因码）。
   *
   * 不是判据：拿掉它，行为一字不变。存在是因为真机上提示标题只有几种，
   * 分不清"证据不足""键冲突""调用方强制确认"，而 WPS 任务窗格开不了 devtools。
   */
  onTargetDiagnostic?: (diagnostic: TargetDiagnostic) => void;
}) {
  const live = useRef(options);
  live.current = options;
  const bound = useRef<AuthorizedTarget | null>(null);
  const retry = useRef<(() => void) | undefined>(undefined);
  const analysisRetries = useRef(new WeakSet<() => void>());
  const flight = useRef<Promise<AuthorizedTarget | null> | null>(null);
  const mounted = useRef(true);
  const checksInFlight = useRef(0);
  /**
   * 基线代数：每推进一次基线 +1。
   *
   * 读数与"用这份读数做判断"之间隔着一个 `await`，中间完全可能有人推进基线
   * （删表收尾的 `bindVerifiedShrink`、原地撤销的 `bindVerifiedRestore`、
   * 从文件入口采用当前工作簿的 `bind` —— 它们都**不经过** `read()`）。
   * 拿旧读数去比新基线，得到的 `scope_changed` 是假的：用户刚确认完就被要求
   * 再确认一次。所以每次读数都记下发起时的代数，回来时代数变了就作废重读。
   */
  const bindGeneration = useRef(0);
  const confirmLock = useRef(false);
  const analysisConsent = useRef<string | null>(null);
  const [binding, setBinding] = useState<AuthorizedTarget | null>(null);
  const [notice, setNotice] = useState<{
    bound: AuthorizedTarget; current: AuthorizedTarget | null; reason: TargetNoticeReason;
    /** Stable across a rejected confirmation retry; changes only when a genuinely new round starts. */
    round: number;
    /** Sheet ids to re-apply after a rejected retry dropped ids the new target lacks. */
    retrySelection?: string[];
  } | null>(null);
  const [error, setError] = useState("");
  const [confirming, setConfirming] = useState(false);
  const [checking, setChecking] = useState(false);
  const noticeRef = useRef(notice);
  noticeRef.current = notice;
  const roundRef = useRef(0);
  /** 本页面内第几次判定（单调递增）。用来区分"新请求又被拦了一次"与"旧提示没清掉"。 */
  const diagnosticSeq = useRef(0);
  /**
   * 交出一条诊断。**任何异常都不许影响判定** —— 它不是判据的一部分。
   * `round` 取当前提示轮次：同一轮里反复判定说明还是同一条提示，没有开新轮。
   */
  const emitDiagnostic = (input: Omit<TargetDiagnosticInput, "round" | "seq">) => {
    const sink = live.current.onTargetDiagnostic;
    if (!sink) return;
    diagnosticSeq.current += 1;
    try {
      sink(targetDiagnostic({ ...input, round: roundRef.current, seq: diagnosticSeq.current }));
    } catch {
      // 诊断永远不许打断守卫。
    }
  };
  /**
   * 判定**通过**之后，把"已经不再成立"的提示清掉。
   *
   * 不清的后果实测过：先前留下的提示只由 `bind()` / 确认 / 取消清除，而定时巡检
   * 在提示存在时又主动跳过 —— 于是目标早已恢复正常，界面还挂着那条提示，
   * 用户看到的是"又拦了我一次"，实际是`旧提示没人清`。这两件事在界面上长得一样。
   *
   * 只清**本道闸能解释**的提示（理由由调用方给）：分析闸通过不能替元数据闸的
   * "换了文件"背书，反之亦然。
   */
  const clearStaleNotice = (reasons: readonly TargetNoticeReason[]) => {
    const shown = noticeRef.current;
    if (!shown || !reasons.includes(shown.reason)) return;
    noticeRef.current = null;
    setNotice(null);
    setError("");
    // 提示没了，它挂着的"待继续请求"也不再有入口：留着只会在下一次确认时被重放。
    retry.current = undefined;
  };
  const read = () => {
    if (!flight.current) {
      flight.current = live.current.inspect().catch(() => null).finally(() => { flight.current = null; });
    }
    return flight.current;
  };
  const readWithRetry = async () => {
    for (let attempt = 0; attempt < capabilities.authorizedTarget.verificationAttempts; attempt++) {
      const target = await read();
      if (!mounted.current || target) return target;
    }
    return null;
  };
  const bind = (target: AuthorizedTarget) => {
    if (!bound.current || checkAuthorizedTarget(bound.current, target) !== "same") analysisConsent.current = null;
    bound.current = target;
    bindGeneration.current += 1;
    setBinding(target);
    setNotice(null);
    setError("");
  };
  const pause = (current: AuthorizedTarget | null, reason: TargetNoticeReason, onRetry?: () => void,
    revokeConsent = reason !== "unavailable") => {
    // Readability is not authorization. Keep explicit consent while evidence is
    // unavailable; use it only after fresh reads prove the original key again.
    // Positive changes and conflicting proofs still revoke it immediately.
    if (revokeConsent) analysisConsent.current = null;
    if (onRetry) retry.current = onRetry;
    const missing: AuthorizedTarget = { fileKey: null, name: "目标身份暂不可读取", sheets: [], selected: [] };
    const shown = noticeRef.current;
    const shownCurrent = shown?.current ?? null;
    // Compare identity, not a stale selection carried by an open dialog.
    const sameShownFile = Boolean(shownCurrent && current && checkAuthorizedTarget(
      { ...shownCurrent, selected: current.selected }, current) === "same");
    if (shown && sameShownFile) setNotice({ ...shown, current, reason });
    else { roundRef.current += 1; setNotice({ bound: bound.current ?? missing, current, reason, round: roundRef.current }); }
    return false;
  };
  /**
   * 拦下并记录。
   *
   * 顺序不能反过来：`pause` 才是"这一轮提示"的起点（它可能沿用旧轮、也可能开新轮），
   * 先 emit 再 pause 会把**上一轮**的编号写进诊断，于是"同一 round = 同一条提示"
   * 这个判读规则当场失效。
   */
  const stop = (current: AuthorizedTarget | null, reason: TargetNoticeReason,
    diagnostic: Omit<TargetDiagnosticInput, "round" | "seq">,
    onRetry?: () => void, revokeConsent?: boolean) => {
    const blocked = pause(current, reason, onRetry, revokeConsent);
    emitDiagnostic(diagnostic);
    return blocked;
  };
  const check = async (onRetry?: () => void, requireConfirmation = false): Promise<boolean> => {
    if (live.current.selectionPending?.()) return false;
    if (!live.current.enabled) return true;
    checksInFlight.current += 1;
    setChecking(true);
    let lastCurrent: AuthorizedTarget | null = null;
    try {
      /*
       * 读数期间基线可能被推进（见 `bindGeneration`）。这种情况作废这一次读数、
       * 重读一次；**不**拿旧读数去比新基线（那是假的范围变化），也**不**直接
       * 放行（那会把新基线当成旧读数的结论）。
       *
       * 上限 = 正常重试次数 + 2 次"作废"，保证一条"读到稳定读数就算"的路，
       * 同时不会因为持续推进基线而转圈：都在作废就按"读不出来"处理，交给用户确认。
       */
      const maxReads = capabilities.authorizedTarget.verificationAttempts + 2;
      for (let attempt = 0; attempt < maxReads; attempt++) {
        const generation = bindGeneration.current;
        const current = await readWithRetry();
        if (!mounted.current || live.current.selectionPending?.()) return false;
        lastCurrent = current;
        if (bindGeneration.current !== generation) continue;
        if (!bound.current && current) bind(current);
        if (!bound.current) {
          return stop(current, "unavailable",
            { outcome: "blocked", gate: "metadata", branch: "metadata_unavailable" }, onRetry);
        }
        // Only a positively identified owned copy is exempt, never an arbitrary active file.
        if (current && !onRetry && !requireConfirmation && live.current.isReviewCopy?.(current)) return true;
        const detected = checkAuthorizedTarget(bound.current, current);
        const reason = requireConfirmation && detected === "same" ? "confirmation_required" : detected;
        if (reason === "same") {
          emitDiagnostic({ outcome: "allowed", gate: "metadata", pathIdentity: Boolean(current?.fileKey) });
          // 元数据闸通过：由**它**产生的旧提示现在不成立了，清掉。
          clearStaleNotice(["file_changed", "scope_changed", "unavailable"]);
          return true;
        }
        return stop(current, reason, {
          outcome: "blocked", gate: "metadata",
          branch: reason === "confirmation_required" ? "forced_confirmation"
            : reason === "file_changed" ? "metadata_file_changed"
            : reason === "scope_changed" ? "metadata_scope_changed"
            : "metadata_unavailable",
          pathIdentity: Boolean(current?.fileKey)
        }, onRetry);
      }
      if (!mounted.current) return false;
      return stop(lastCurrent, "unavailable",
        { outcome: "blocked", gate: "metadata", branch: "metadata_unavailable" }, onRetry);
    } finally {
      checksInFlight.current -= 1;
      if (mounted.current && checksInFlight.current === 0) setChecking(false);
    }
  };
  const checkForAnalysis = async (onRetry?: () => void, options: { requireExplicitConsent?: boolean } = {}): Promise<boolean> => {
    if (onRetry) analysisRetries.current.add(onRetry);
    checksInFlight.current += 1;
    setChecking(true);
    try {
      if (!(await check(onRetry))) return false;
      if (!live.current.enabled || !live.current.analysisEvidence) return true;
      let lastCurrent: AuthorizedTarget | null = null;
      for (let attempt = 0; attempt < capabilities.authorizedTarget.verificationAttempts; attempt++) {
        try {
          const generation = bindGeneration.current;
          const current = await read();
          if (!mounted.current || live.current.selectionPending?.()) return false;
          if (generation !== bindGeneration.current) continue;
          lastCurrent = current;
          if (!bound.current) {
            return stop(current, "unavailable",
              { outcome: "blocked", gate: "analysis", branch: "evidence_missing" }, onRetry);
          }
          const detected = checkAuthorizedTarget(bound.current, current);
          if (detected === "file_changed" || detected === "scope_changed") {
            return stop(current, detected, {
              outcome: "blocked", gate: "analysis",
              branch: detected === "file_changed" ? "metadata_file_changed" : "metadata_scope_changed"
            }, onRetry);
          }
          if (detected !== "same") continue;
          const boundEvidence = await live.current.analysisEvidence(bound.current);
          const evidence = current && await live.current.analysisEvidence(current);
          if (!mounted.current || live.current.selectionPending?.()) return false;
          if (generation !== bindGeneration.current) continue;
          // Inspect again after asynchronous evidence. Positive changes are never retried away.
          const after = await read();
          if (!mounted.current || live.current.selectionPending?.()) return false;
          if (generation !== bindGeneration.current) continue;
          lastCurrent = after;
          const afterCheck = checkAuthorizedTarget(bound.current, after);
          if (afterCheck === "file_changed" || afterCheck === "scope_changed") {
            return stop(after, afterCheck, {
              outcome: "blocked", gate: "analysis",
              branch: afterCheck === "file_changed" ? "metadata_file_changed" : "metadata_scope_changed"
            }, onRetry);
          }
          if (!current || afterCheck !== "same" || checkAuthorizedTarget(current, after) !== "same") continue;
          const repeated = after && await live.current.analysisEvidence(after);
          if (!mounted.current || live.current.selectionPending?.()) return false;
          if (generation !== bindGeneration.current) continue;
          // A missing third proof must not erase a conflict already established by two.
          const keys = [boundEvidence?.key, evidence?.key, repeated?.key].filter((key) => key !== undefined);
          if (new Set(keys).size > 1) {
            return stop(after, "unavailable", {
              outcome: "blocked", gate: "analysis", branch: "evidence_conflict",
              conclusive: evidence?.conclusive === true,
              documentBound: evidence?.documentBound === true,
              stableId: evidence?.diagnostic?.stableId ?? "unknown",
              pathIdentity: evidence?.diagnostic?.pathIdentity === true
            }, onRetry, true);
          }
          if (!boundEvidence || !evidence || !repeated) continue;
          /*
           * 这一处的事实摊平成布尔值，供诊断与判定共用 —— 诊断里写的必须是
           * **判定真正读到的那几个值**，不是另算一份"看起来一样"的结论。
           */
          const facts = {
            conclusive: evidence.conclusive && repeated.conclusive,
            documentBound: evidence.documentBound === true && repeated.documentBound === true,
            hasConsent: analysisConsent.current === evidence.key,
            explicitRequired: options.requireExplicitConsent === true,
            stableId: evidence.diagnostic?.stableId ?? "unknown" as const,
            pathIdentity: evidence.diagnostic?.pathIdentity === true
          };
          if (facts.conclusive || facts.hasConsent) {
            emitDiagnostic({
              outcome: "allowed", gate: "analysis",
              via: facts.conclusive ? "conclusive" : "consent", ...facts
            });
            // 分析闸通过：`confirmation_required` 这条提示已经不成立，清掉它。
            clearStaleNotice(["confirmation_required"]);
            return true;
          }
          /*
           * 文档级证据（路径身份 + 宿主声明的工作表稳定 ID）**只在宿主当次真的给出
           * 稳定 ID 时**才成立 —— 它不假设 `CodeName`/`SId` 可读，读不到就是 false。
           * 成立时它跨刷新仍有效，因此不必再让用户确认一次；它**只**用在这道分析闸，
           * 写入、恢复与数据读取仍各自只认 `conclusive`（对象级）。
           *
           * `requireExplicitConsent`：只**禁止用文档级证据替调用方跳过显式确认**
           * （恢复一张挂起的任务卡时，调用方自己判定旧授权不再适用）。它**不**是
           * "无条件再问一次" —— 对象级证明或本会话同键同意已经放行的话，走不到这里。
           */
          if (!facts.explicitRequired && facts.documentBound) {
            emitDiagnostic({ outcome: "allowed", gate: "analysis", via: "documentBound", ...facts });
            clearStaleNotice(["confirmation_required"]);
            return true;
          }
          // Consistent weak evidence permits explicit fresh-read consent, not automatic access.
          return stop(after, "confirmation_required", {
            outcome: "blocked", gate: "analysis",
            branch: facts.documentBound ? "document_bound_rejected" : "weak_evidence", ...facts
          }, onRetry);
        } catch {
          // Retry transient host failures, without clearing the consent between attempts.
        }
      }
      if (!mounted.current) return false;
      return stop(lastCurrent, "unavailable",
        { outcome: "blocked", gate: "analysis", branch: "evidence_missing" }, onRetry);
    } finally {
      checksInFlight.current -= 1;
      if (mounted.current && checksInFlight.current === 0) setChecking(false);
    }
  };
  /** Explicit scope selection is the authorization action, not another analysis request.
   * Re-read metadata/evidence before binding; never replay a blocked request or write. */
  const confirmScope = async (expected: AuthorizedTarget, options: {
    continuePendingAnalysis?: boolean; isCurrent?: () => boolean; allowEmpty?: boolean;
  } = {}): Promise<boolean> => {
    if (!live.current.enabled || confirmLock.current) return false;
    confirmLock.current = true;
    setConfirming(true);
    setError("");
    try {
      if (live.current.reviewActive) {
        setError("请先保存或放弃当前审阅副本，再切换任务授权。");
        return false;
      }
      if (flight.current) await flight.current;
      const fresh = await read();
      if (!mounted.current || options.isCurrent?.() === false) return false;
      const actual = fresh && { ...fresh, selected: [...expected.selected] };
      if (!actual || (!expected.selected.length && !options.allowEmpty) || new Set(expected.selected).size !== expected.selected.length ||
          checkAuthorizedTarget(expected, actual) !== "same") {
        setError("文件或工作表在确认期间发生变化，请重新检查；未变更授权。");
        return false;
      }
      let key: string | null = null;
      if (expected.selected.length && live.current.analysisEvidence) {
        // Compare the clicked snapshot as well as the fresh one: the same URL alone
        // must not authorize a replacement Office document.
        const before = await live.current.analysisEvidence(expected);
        const evidence = await live.current.analysisEvidence(actual);
        const after = await read();
        const repeated = after && await live.current.analysisEvidence({ ...after, selected: actual.selected });
        if (!mounted.current || options.isCurrent?.() === false) return false;
        if (!before || !evidence || !repeated || before.key !== evidence.key || evidence.key !== repeated.key ||
            !after || checkAuthorizedTarget(actual, { ...after, selected: actual.selected }) !== "same") {
          setError("暂时无法核实目标身份，请重新检查；未变更授权。");
          return false;
        }
        key = evidence.key;
      }
      if (!mounted.current || !live.current.enabled || options.isCurrent?.() === false) return false;
      if (live.current.reviewActive) {
        setError("请先保存或放弃当前审阅副本，再切换任务授权。");
        return false;
      }
      // Only callbacks registered by analysis preflight may resume from the picker.
      // A stopped write/preview action is never made replayable by confirming scope.
      const pending = options.continuePendingAnalysis && noticeRef.current && retry.current &&
        analysisRetries.current.has(retry.current) ? retry.current : undefined;
      retry.current = undefined;
      live.current.adopt(actual);
      bind(actual);
      analysisConsent.current = key;
      if (pending) setTimeout(() => { if (mounted.current) pending(); }, 0);
      return true;
    } catch {
      if (mounted.current) setError("暂时无法核实目标身份，请重新检查；未变更授权。");
      return false;
    } finally {
      confirmLock.current = false;
      if (mounted.current) setConfirming(false);
    }
  };
  const confirm = async (choice: "original" | "current", selected: string[]) => {
    if (!notice || notice.reason === "unavailable" || confirmLock.current ||
        (notice.reason === "confirmation_required" && choice !== "original")) return;
    confirmLock.current = true;
    setConfirming(true);
    setError("");
    try {
      // Do not reuse an inspection started before the user clicked confirmation.
      if (flight.current) await flight.current;
      const fresh = await read();
      if (!mounted.current) return;
      const expected = choice === "original" ? notice.bound : notice.current;
      if (!fresh || !expected) {
        setError("暂时无法读取文件身份，请切回目标文件后重新检查。");
        return;
      }
      const next = choice === "original" ? expected : { ...expected, selected };
      // Re-read at confirmation: an open dialog is not an authorization token.
      const actual = { ...fresh, selected: next.selected };
      if (!next.selected.length || checkAuthorizedTarget(next, actual) !== "same") {
        // Same round as the rejected attempt: keep the user's choice so a retry is one click,
        // not a full re-selection. Drop only sheet ids the re-read target no longer has.
        const alive = new Set(fresh.sheets.map((sheet) => sheet.id));
        const kept = selected.filter((id) => alive.has(id));
        const detected = checkAuthorizedTarget(notice.bound, fresh);
        setNotice({ ...notice, current: fresh, reason: detected === "same" ? "confirmation_required" : detected, round: notice.round,
          ...(choice === "current" && kept.length !== selected.length ? { retrySelection: kept } : {}) });
        setError("文件或工作表在确认期间又发生变化，请重新选择；尚未进行操作。");
        return;
      }
      if (live.current.reviewActive) {
        setError("请先保存或放弃当前审阅副本，再切换任务授权。");
        return;
      }
      let analysisKey: string | null = null;
      if (live.current.analysisEvidence) {
        const before = await live.current.analysisEvidence(next);
        const evidence = await live.current.analysisEvidence(actual);
        const after = await read();
        const repeated = after && await live.current.analysisEvidence({ ...after, selected: actual.selected });
        if (!mounted.current) return;
        // Inspection still carries the old committed selection until adopt() below.
        // Validate the newly checked ids against live metadata, not the stale selection.
        if (!before || !evidence || !repeated || before.key !== evidence.key || evidence.key !== repeated.key || !after ||
            checkAuthorizedTarget(actual, { ...after, selected: actual.selected }) !== "same") {
          setError("无法确认目标身份，请先保存目标文件并重新检查；尚未进行需求分析或写入。");
          return;
        }
        analysisKey = evidence.key;
      }
      if (!mounted.current || !live.current.enabled) return;
      if (live.current.reviewActive) {
        setError("请先保存或放弃当前审阅副本，再切换任务授权。");
        return;
      }
      live.current.adopt(actual);
      bind(actual);
      analysisConsent.current = analysisKey;
      // Continue only a fresh user request; never replay an old write/preview action.
      const pending = retry.current;
      retry.current = undefined;
      if (pending) setTimeout(() => { if (mounted.current) pending(); }, 0);
    } catch {
      if (mounted.current) setError("确认目标未完成，请检查当前文件后重试；尚未执行写入。");
    } finally {
      confirmLock.current = false;
      if (mounted.current) setConfirming(false);
    }
  };
  const recheck = async () => {
    if (!noticeRef.current || noticeRef.current.reason !== "unavailable" || confirmLock.current) return;
    confirmLock.current = true;
    setConfirming(true);
    setError("");
    try {
      // Classify before checkForAnalysis registers retries. A write callback from
      // check() must never become an analysis callback merely by rechecking.
      const pending = retry.current && analysisRetries.current.has(retry.current) ? retry.current : undefined;
      if (!(await checkForAnalysis(pending)) || !mounted.current) return;
      setNotice(null);
      retry.current = undefined;
      // Only a queued fresh analysis request can resume. Old writes are never queued here.
      if (pending) setTimeout(() => { if (mounted.current) pending(); }, 0);
    } finally {
      confirmLock.current = false;
      if (mounted.current) setConfirming(false);
    }
  };
  /**
   * 只接受**同一次删表造成的收窄**，用它替掉一次会弹"范围发生变化"的确认。
   *
   * 调用方必须先按执行回执核对过（共享层 `reconcileWorksheetDeletionScope`）；
   * 这里再独立卡一遍"只减不增"，因为这是唯一一条不走 `confirmScope` 就推进
   * 基线的路：
   * - 文件身份必须相同（换簿一律拒绝）；
   * - 收窄后的选中表名必须是原选中表名的**真子集**（不能新增、不能换名）；
   * - 目标里的表不能出现原来没有的名字（外部新增工作表一律拒绝）。
   * 任一条不成立就返回 false，基线不动，原有的范围提示照常出现。
   */
  const bindVerifiedShrink = (target: AuthorizedTarget): boolean => {
    const current = bound.current;
    if (!mounted.current || !current) return false;
    if (!current.fileKey || current.fileKey !== target.fileKey) return false;
    const names = (item: AuthorizedTarget) =>
      item.sheets.filter((sheet) => item.selected.includes(sheet.id)).map((sheet) => sheet.name);
    const before = names(current);
    const after = names(target);
    if (!after.length || after.length >= before.length) return false;
    const allowed = new Set(before);
    if (after.some((name) => !allowed.has(name))) return false;
    const known = new Set(current.sheets.map((sheet) => sheet.name));
    if (target.sheets.some((sheet) => !known.has(sheet.name))) return false;
    bind(target);
    return true;
  };
  /**
   * 只接受**同一次原地撤销**造成的目录增加。
   *
   * 撤销会把表放回工作簿，于是目录比基线多了几张。若不认下来，下一轮守卫会把它
   * 读成"范围变了"，弹一句"尚未写入"——与刚撤销成功的事实冲突。
   *
   * 放过的条件同样要证明（调用方已用共享层 `reconcileWorksheetRestoreScope` 核对过）：
   * - 文件身份必须相同；
   * - 授权集合**一字不差**（撤销不扩大授权，恢复出来的表也不会自动被授权）；
   * - 目录里新增的名字只能来自这次撤销，且回执里的表必须真的出现。
   */
  const bindVerifiedRestore = (target: AuthorizedTarget, restoredSheets: readonly string[]): boolean => {
    const current = bound.current;
    if (!mounted.current || !current || !restoredSheets.length) return false;
    if (!current.fileKey || current.fileKey !== target.fileKey) return false;
    const names = (item: AuthorizedTarget) =>
      item.selected
        .map((id) => item.sheets.find((sheet) => sheet.id === id)?.name)
        .filter((name): name is string => Boolean(name));
    const before = names(current);
    const after = names(target);
    if (!before.length || before.length !== after.length) return false;
    const beforeSet = new Set(before);
    const afterSet = new Set(after);
    if (before.some((name) => !afterSet.has(name)) || after.some((name) => !beforeSet.has(name))) return false;
    const restored = new Set(restoredSheets);
    const knownBefore = new Set(current.sheets.map((sheet) => sheet.name));
    const knownAfter = new Set(target.sheets.map((sheet) => sheet.name));
    if (target.sheets.some((sheet) => !knownBefore.has(sheet.name) && !restored.has(sheet.name))) return false;
    if (restoredSheets.some((name) => !knownAfter.has(name))) return false;
    bind(target);
    return true;
  };
  const checkRef = useRef(check);  checkRef.current = check;  useEffect(() => {
    mounted.current = true;
    const inspect = () => {
      if (live.current.enabled && !live.current.reviewActive && !live.current.monitorPaused && !live.current.selectionPending?.() &&
          !noticeRef.current && document.visibilityState !== "hidden" &&
          !confirmLock.current && !flight.current) void checkRef.current();
    };
    const visible = () => { if (document.visibilityState !== "hidden") inspect(); };
    inspect();
    window.addEventListener("focus", inspect);
    document.addEventListener("visibilitychange", visible);
    const timer = window.setInterval(inspect, capabilities.authorizedTarget.monitorIntervalMs);
    return () => {
      mounted.current = false;
      window.removeEventListener("focus", inspect);
      document.removeEventListener("visibilitychange", visible);
      window.clearInterval(timer);
      retry.current = undefined;
      analysisConsent.current = null;
    };
  }, [options.enabled]);
  return { binding, notice, error, confirming, checking, check, checkForAnalysis, confirm, confirmScope, recheck,
    hasAnalysisConsent: (key: string) => analysisConsent.current === key,
    clearAnalysisConsent: () => { analysisConsent.current = null; },
    /** A stronger host guard can require explicit consent even if metadata alone looks unchanged. */
    requestConfirmation: (onRetry?: () => void) => check(onRetry, true),
    cancel: () => { retry.current = undefined; setNotice(null); setError(""); },
    /** Only explicit scope confirmation may advance this baseline. */
    bind,
    /** 同一次删表的预期收窄；调用方必须先用回执核对过。 */
    bindVerifiedShrink,
    /** 同一次原地撤销造成的目录增加；授权集合不变。 */
    bindVerifiedRestore };
}
