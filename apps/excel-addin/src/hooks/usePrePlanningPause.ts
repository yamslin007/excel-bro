import { useEffect, useRef, useState } from "react";
import { checkAuthorizedTarget, type AuthorizedTarget } from "@excel-bro/core/authorizedTarget";

/** Local request retention, not a server checkpoint. Never persisted or sent to the model. */
export function usePrePlanningPause<T>(contextKey: string) {
  type Entry = { id: string; messageId: string; conversationId: string; target: AuthorizedTarget | null;
    mode: string; data: T; error?: string };
  const [entries, setEntries] = useState<Entry[]>([]);
  const live = useRef<Entry[]>([]);
  const claimed = useRef(new Set<string>());
  const activeContext = useRef(contextKey);
  activeContext.current = contextKey;
  const mounted = useRef(true);
  useEffect(() => {
    mounted.current = true;
    return () => { mounted.current = false; };
  }, []);
  const publish = (next: Entry[]) => { live.current = next; setEntries(next); };
  const remember = (entry: Omit<Entry, "id">) => {
    const id = `preplanning-${entry.messageId}`;
    publish([...live.current.filter((item) => item.id !== id), { ...entry, id,
      target: entry.target ? { ...entry.target, selected: [...entry.target.selected],
        sheets: entry.target.sheets.map((sheet) => ({ ...sheet })) } : null }]);
    return id;
  };
  const restart = async (id: string, conversationId: string, mode: string,
    inspect: () => Promise<AuthorizedTarget | null>, run: (data: T) => Promise<void>,
    stillCurrent: () => boolean = () => true) => {
    const entry = live.current.find((item) => item.id === id && item.conversationId === conversationId);
    if (!entry || claimed.current.has(id)) return;
    const issuedContext = activeContext.current;
    claimed.current.add(id);
    try {
      const target = await inspect();
      if (!mounted.current || !stillCurrent() || issuedContext !== activeContext.current) return;
      if (entry.mode !== mode || !entry.target || checkAuthorizedTarget(entry.target, target) !== "same") {
        publish(live.current.map((item) => item.id === id ? { ...item,
          error: "原需求的工作簿或授权范围无法核实；请检查数据范围。未重新分析、未写入。" } : item));
        return;
      }
      // Consume before dispatch: double clicks and late semantic responses cannot replay it.
      publish(live.current.filter((item) => item.id !== id));
      await run(entry.data);
    } catch {
      if (!mounted.current) return;
      publish(live.current.map((item) => item.id === id ? { ...item,
        error: "暂时无法核实原需求；请稍后重试。未写入工作表。" } : item));
    } finally { claimed.current.delete(id); }
  };
  // Imperative dispatch can run before React renders publish(). Never route
  // another request using the old render's list of already consumed entries.
  const pendingEntries = () => live.current.filter((entry) => !claimed.current.has(entry.id));
  return { entries, pendingEntries, remember, restart };
}
