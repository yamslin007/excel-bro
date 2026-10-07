import { useEffect, useRef, useState } from "react";
import { checkAuthorizedTarget, type AuthorizedTarget, type TargetCheck } from "@excel-bro/core/authorizedTarget";
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
   * Keys are private memory bindings. Unknown evidence requires consent BEFORE analysis. */
  analysisEvidence?: (target: AuthorizedTarget) => Promise<{ key: string; conclusive: boolean } | null>;
}) {
  const live = useRef(options);
  live.current = options;
  const bound = useRef<AuthorizedTarget | null>(null);
  const retry = useRef<(() => void) | undefined>(undefined);
  const flight = useRef<Promise<AuthorizedTarget | null> | null>(null);
  const mounted = useRef(true);
  const checksInFlight = useRef(0);
  const confirmLock = useRef(false);
  const analysisConsent = useRef<string | null>(null);
  const [binding, setBinding] = useState<AuthorizedTarget | null>(null);
  const [notice, setNotice] = useState<{
    bound: AuthorizedTarget; current: AuthorizedTarget | null; reason: TargetCheck;
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
  const bind = (target: AuthorizedTarget) => {
    if (!bound.current || checkAuthorizedTarget(bound.current, target) !== "same") analysisConsent.current = null;
    bound.current = target;
    setBinding(target);
    setNotice(null);
    setError("");
  };
  const check = async (onRetry?: () => void, requireConfirmation = false): Promise<boolean> => {
    if (!live.current.enabled) return true;
    checksInFlight.current += 1;
    setChecking(true);
    try {
      const current = await read();
      if (!mounted.current) return false;
      if (!bound.current && current) bind(current);
      if (!bound.current) {
        const missing: AuthorizedTarget = { fileKey: null, name: "原文件身份暂不可读取", sheets: [], selected: [] };
        if (onRetry) retry.current = onRetry;
        roundRef.current += 1;
        setNotice({ bound: missing, current, reason: "unavailable", round: roundRef.current });
        return false;
      }
      // Only a positively identified owned copy is exempt, never an arbitrary active file.
      if (current && !onRetry && !requireConfirmation && live.current.isReviewCopy?.(current)) return true;
      const detected = checkAuthorizedTarget(bound.current, current);
      const reason = requireConfirmation && detected === "same" ? "unavailable" : detected;
      if (reason === "same") return true;
      if (onRetry) retry.current = onRetry;
      // Keep the round only while the shown file identity is unchanged, so a rejected
      // confirmation retry keeps the user's choice. A different file is a genuinely new
      // round and must not inherit a selection made against another file. Compare file
      // identity only: the shown target carries a stale `selected` the user has not re-made.
      const shown = noticeRef.current;
      const shownCurrent = shown?.current ?? null;
      // Compare file identity only: the shown target carries a stale `selected` the user has not re-made.
      const sameShownFile = Boolean(shownCurrent && current && checkAuthorizedTarget(
        { ...shownCurrent, selected: current.selected }, current) === "same");
      if (shown && sameShownFile) setNotice({ ...shown, current, reason });
      else { roundRef.current += 1; setNotice({ bound: bound.current, current, reason, round: roundRef.current }); }
      return false;
    } finally {
      checksInFlight.current -= 1;
      if (mounted.current && checksInFlight.current === 0) setChecking(false);
    }
  };
  const checkForAnalysis = async (onRetry?: () => void): Promise<boolean> => {
    checksInFlight.current += 1;
    setChecking(true);
    try {
      if (!(await check(onRetry))) return false;
      if (!live.current.enabled || !live.current.analysisEvidence) return true;
      try {
        const current = await read();
        const evidence = current && await live.current.analysisEvidence(current);
        if (!mounted.current) return false;
        // Inspect again after async host evidence; do not reuse a target from before the read.
        const after = await read();
        if (!current || checkAuthorizedTarget(current, after) !== "same" ||
            !bound.current || checkAuthorizedTarget(bound.current, after) !== "same") return await check(onRetry, true);
        const repeated = after && await live.current.analysisEvidence(after);
        if (evidence && repeated && evidence.key === repeated.key &&
            ((evidence.conclusive && repeated.conclusive) || analysisConsent.current === evidence.key)) return true;
      } catch {
        // Missing host evidence is a pause, not a model/execution failure.
      }
      return await check(onRetry, true);
    } finally {
      checksInFlight.current -= 1;
      if (mounted.current && checksInFlight.current === 0) setChecking(false);
    }
  };
  const confirm = async (choice: "original" | "current", selected: string[]) => {
    if (!notice || confirmLock.current) return;
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
        setNotice({ ...notice, current: fresh, reason: checkAuthorizedTarget(notice.bound, fresh), round: notice.round,
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
        const evidence = await live.current.analysisEvidence(actual);
        const repeated = await live.current.analysisEvidence(actual);
        const after = await read();
        if (!evidence || !repeated || evidence.key !== repeated.key || checkAuthorizedTarget(actual, after) !== "same") {
          setError("无法确认目标身份，请先保存目标文件并重新检查；尚未进行需求分析或写入。");
          return;
        }
        analysisKey = evidence.key;
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
    };
  }, [options.enabled]);
  return { binding, notice, error, confirming, checking, check, checkForAnalysis, confirm,
    hasAnalysisConsent: (key: string) => analysisConsent.current === key,
    clearAnalysisConsent: () => { analysisConsent.current = null; },
    /** A stronger host guard can require explicit consent even if metadata alone looks unchanged. */
    requestConfirmation: (onRetry?: () => void) => check(onRetry, true),
    cancel: () => { retry.current = undefined; setNotice(null); setError(""); },
    /** Only explicit scope confirmation may advance this baseline. */
    bind };
}
