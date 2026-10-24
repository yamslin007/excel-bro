import { useEffect, useRef, useState } from "react";

/** Only an explicit picker gesture requests authorization. Latest gesture wins;
 * catalog refreshes and restored UI state cannot grant access by themselves. */
export function useImmediateScopeAuthorization(options: {
  ready: boolean;
  apply: (isCurrent: () => boolean) => Promise<boolean>;
}) {
  const live = useRef(options); live.current = options;
  const generation = useRef(0);
  const attempted = useRef(0);
  const applied = useRef(0);
  const running = useRef(false);
  const mounted = useRef(true);
  const [revision, setRevision] = useState(0);
  useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
  const request = () => { generation.current++; setRevision((value) => value + 1); };
  useEffect(() => {
    const turn = generation.current;
    if (!turn || attempted.current === turn || running.current || !live.current.ready) return;
    attempted.current = turn;
    running.current = true;
    const isCurrent = () => mounted.current && generation.current === turn;
    void live.current.apply(isCurrent).then((ok) => {
      if (ok && isCurrent()) applied.current = turn;
    }).catch(() => { /* The picker retains the selection and reports its failure. */ }).finally(() => {
      running.current = false;
      if (mounted.current) setRevision((value) => value + 1);
    });
  }, [revision, options.ready]);
  return { request, pending: generation.current > applied.current,
    isPending: () => generation.current > applied.current };
}
