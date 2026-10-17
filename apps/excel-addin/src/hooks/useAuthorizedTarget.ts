import { useEffect, useRef, useState } from "react";
import { checkAuthorizedTarget, type AuthorizedTarget, type TargetNoticeReason } from "@excel-bro/core/authorizedTarget";
import capabilities from "../../../../config/capabilities.json";

export function useAuthorizedTarget(options: {
  enabled: boolean;
  inspect: () => Promise<AuthorizedTarget>;
  adopt: (target: AuthorizedTarget) => void;
  /** A known review copy must not be adopted as an original by this guard. */
  reviewActive?: boolean;
  monitorPaused?: boolean;
  isReviewCopy?: (target: AuthorizedTarget) => boolean;
  /** Host evidence for analysis/read authorization only, never permission to write.
   * Keys are private memory bindings. Consistent weak evidence requires consent;
   * missing or conflicting evidence remains blocked until a successful recheck. */
  analysisEvidence?: (target: AuthorizedTarget) => Promise<{ key: string; conclusive: boolean } | null>;
}) {
  const live = useRef(options);
  live.current = options;
  const bound = useRef<AuthorizedTarget | null>(null);
  const retry = useRef<(() => void) | undefined>(undefined);
  const analysisRetries = useRef(new WeakSet<() => void>());
  const flight = useRef<Promise<AuthorizedTarget | null> | null>(null);
  const mounted = useRef(true);
  const checksInFlight = useRef(0);
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
    setBinding(target);
    setNotice(null);
    setError("");
  };
  const pause = (current: AuthorizedTarget | null, reason: TargetNoticeReason, onRetry?: () => void) => {
    analysisConsent.current = null;
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
  const check = async (onRetry?: () => void, requireConfirmation = false): Promise<boolean> => {
    if (!live.current.enabled) return true;
    checksInFlight.current += 1;
    setChecking(true);
    try {
      const current = await readWithRetry();
      if (!mounted.current) return false;
      if (!bound.current && current) bind(current);
      if (!bound.current) {
        return pause(current, "unavailable", onRetry);
      }
      // Only a positively identified owned copy is exempt, never an arbitrary active file.
      if (current && !onRetry && !requireConfirmation && live.current.isReviewCopy?.(current)) return true;
      const detected = checkAuthorizedTarget(bound.current, current);
      const reason = requireConfirmation && detected === "same" ? "confirmation_required" : detected;
      if (reason === "same") return true;
      return pause(current, reason, onRetry);
    } finally {
      checksInFlight.current -= 1;
      if (mounted.current && checksInFlight.current === 0) setChecking(false);
    }
  };
  const checkForAnalysis = async (onRetry?: () => void): Promise<boolean> => {
    if (onRetry) analysisRetries.current.add(onRetry);
    checksInFlight.current += 1;
    setChecking(true);
    try {
      if (!(await check(onRetry))) return false;
      if (!live.current.enabled || !live.current.analysisEvidence) return true;
      let lastCurrent: AuthorizedTarget | null = null;
      for (let attempt = 0; attempt < capabilities.authorizedTarget.verificationAttempts; attempt++) {
        try {
          const current = await read();
          if (!mounted.current) return false;
          lastCurrent = current;
          if (!bound.current) return pause(current, "unavailable", onRetry);
          const detected = checkAuthorizedTarget(bound.current, current);
          if (detected === "file_changed" || detected === "scope_changed") return pause(current, detected, onRetry);
          if (detected !== "same") continue;
          const boundEvidence = await live.current.analysisEvidence(bound.current);
          const evidence = current && await live.current.analysisEvidence(current);
          if (!mounted.current) return false;
          // Inspect again after asynchronous evidence. Positive changes are never retried away.
          const after = await read();
          if (!mounted.current) return false;
          lastCurrent = after;
          const afterCheck = checkAuthorizedTarget(bound.current, after);
          if (afterCheck === "file_changed" || afterCheck === "scope_changed") return pause(after, afterCheck, onRetry);
          if (!current || afterCheck !== "same" || checkAuthorizedTarget(current, after) !== "same") continue;
          const repeated = after && await live.current.analysisEvidence(after);
          if (!mounted.current) return false;
          // A missing third proof must not erase a conflict already established by two.
          const keys = [boundEvidence?.key, evidence?.key, repeated?.key].filter((key) => key !== undefined);
          if (new Set(keys).size > 1) return pause(after, "unavailable", onRetry);
          if (!boundEvidence || !evidence || !repeated) continue;
          if ((evidence.conclusive && repeated.conclusive) || analysisConsent.current === evidence.key) return true;
          // Consistent weak evidence permits explicit fresh-read consent, not automatic access.
          return pause(after, "confirmation_required", onRetry);
        } catch {
          // Retry transient host failures, without clearing the consent between attempts.
        }
      }
      if (!mounted.current) return false;
      return pause(lastCurrent, "unavailable", onRetry);
    } finally {
      checksInFlight.current -= 1;
      if (mounted.current && checksInFlight.current === 0) setChecking(false);
    }
  };
  /** Explicit scope selection is the authorization action, not another analysis request.
   * Re-read metadata/evidence before binding; never replay a blocked request or write. */
  const confirmScope = async (expected: AuthorizedTarget, options: { continuePendingAnalysis?: boolean } = {}): Promise<boolean> => {
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
      if (!mounted.current) return false;
      const actual = fresh && { ...fresh, selected: [...expected.selected] };
      if (!actual || !expected.selected.length || new Set(expected.selected).size !== expected.selected.length ||
          checkAuthorizedTarget(expected, actual) !== "same") {
        setError("文件或工作表在确认期间发生变化，请重新检查；未变更授权。");
        return false;
      }
      let key: string | null = null;
      if (live.current.analysisEvidence) {
        // Compare the clicked snapshot as well as the fresh one: the same URL alone
        // must not authorize a replacement Office document.
        const before = await live.current.analysisEvidence(expected);
        const evidence = await live.current.analysisEvidence(actual);
        const after = await read();
        const repeated = after && await live.current.analysisEvidence({ ...after, selected: actual.selected });
        if (!mounted.current) return false;
        if (!before || !evidence || !repeated || before.key !== evidence.key || evidence.key !== repeated.key ||
            !after || checkAuthorizedTarget(actual, { ...after, selected: actual.selected }) !== "same") {
          setError("暂时无法核实目标身份，请重新检查；未变更授权。");
          return false;
        }
        key = evidence.key;
      }
      if (!mounted.current || !live.current.enabled) return false;
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
      if (!(await checkForAnalysis(retry.current)) || !mounted.current) return;
      setNotice(null);
      const pending = retry.current;
      retry.current = undefined;
      // Only a queued fresh analysis request can resume. Old writes are never queued here.
      if (pending) setTimeout(() => { if (mounted.current) pending(); }, 0);
    } finally {
      confirmLock.current = false;
      if (mounted.current) setConfirming(false);
    }
  };
  const checkRef = useRef(check);
  checkRef.current = check;
  useEffect(() => {
    mounted.current = true;
    const inspect = () => {
      if (live.current.enabled && !live.current.reviewActive && !live.current.monitorPaused &&
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
    bind };
}
