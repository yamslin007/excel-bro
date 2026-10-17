import { useEffect, useRef, useState } from "react";
import { admissionComplete, admissionItem, sameAdmissionTarget, SCOPE_ADMISSION_MESSAGES,
  type ScopeAdmissionReport } from "@excel-bro/core/scopeAdmission";
import capabilities from "../../../../config/capabilities.json";

/** Shared by Office, WPS and folder selection. Checking never commits authorization. */
export function useScopeAdmission(options: {
  enabled: boolean;
  scopeKey: string;
  ids: readonly string[];
  inspect: () => Promise<ScopeAdmissionReport>;
}) {
  const live = useRef(options); live.current = options;
  const epoch = useRef(0);
  const mounted = useRef(false);
  const [state, setState] = useState<{ key: string; report: ScopeAdmissionReport | null; checking: boolean }>({
    key: "", report: null, checking: false
  });
  const stateRef = useRef(state); stateRef.current = state;
  const identity = JSON.stringify([options.scopeKey, options.ids]);
  const identityRef = useRef(identity); identityRef.current = identity;
  const run = async (verify = false): Promise<ScopeAdmissionReport | null> => {
    const input = { ...live.current, ids: [...live.current.ids] };
    if (!mounted.current || !input.enabled || input.ids.length === 0) return null;
    const key = identityRef.current;
    const previous = stateRef.current.key === key ? stateRef.current.report : null;
    if (verify && !admissionComplete(input.ids, previous)) return null;
    const turn = ++epoch.current;
    setState({ key, report: null, checking: true });
    let timer: ReturnType<typeof setTimeout> | undefined;
    let report: ScopeAdmissionReport;
    try {
      if (input.ids.length > capabilities.scopeAdmission.maxSelections) {
        throw new Error("selection budget exceeded");
      }
      report = await Promise.race([input.inspect(), new Promise<never>((_, reject) => {
        timer = setTimeout(() => reject(new Error("selection check timed out")), capabilities.scopeAdmission.timeoutMs);
      })]);
      // Reject incomplete, duplicate, foreign or forged status rows; never silently drop an item.
      if (!report || !Array.isArray(report.items) || report.items.length !== input.ids.length ||
          new Set(report.items.map((item) => item.id)).size !== input.ids.length ||
          !input.ids.every((id) => report.items.some((item) => item.id === id &&
            Object.prototype.hasOwnProperty.call(SCOPE_ADMISSION_MESSAGES, item.reason) &&
            item.status === admissionItem(id, item.reason).status))) throw new Error("invalid selection check");
      if (verify && previous && admissionComplete(input.ids, report) && !sameAdmissionTarget(previous, report)) {
        report = { items: input.ids.map((id) => admissionItem(id, "target_changed")) };
      }
    } catch {
      report = { items: input.ids.map((id) => admissionItem(id, "identity_unverified")) };
    } finally { if (timer) clearTimeout(timer); }
    if (!mounted.current || turn !== epoch.current || key !== identityRef.current || !live.current.enabled) return null;
    setState({ key, report, checking: false });
    return admissionComplete(input.ids, report) ? report : null;
  };
  const runRef = useRef(run); runRef.current = run;
  useEffect(() => {
    mounted.current = true;
    if (options.enabled && options.ids.length) void runRef.current();
    else { epoch.current++; setState({ key: identity, report: null, checking: false }); }
    return () => { mounted.current = false; epoch.current++; };
  }, [identity, options.enabled]);
  const current = state.key === identity && options.enabled ? state : null;
  const checking = options.enabled && options.ids.length > 0 && (!current || current.checking);
  return {
    report: current?.report ?? null,
    checking,
    canCommit: !checking && options.enabled && admissionComplete(options.ids, current?.report ?? null),
    item: (id: string) => current?.report?.items.find((item) => item.id === id),
    retry: () => run(false),
    /** Re-read before committing; the visible green check is not a perpetual permission. */
    verify: () => run(true)
  };
}
