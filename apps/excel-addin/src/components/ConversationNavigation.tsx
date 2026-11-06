import { useLayoutEffect, useMemo, useRef, useState, type RefObject } from "react";
import "./ConversationNavigation.css";

type Message = { id: string; role: string; text?: string };
type Props = {
  messages: readonly Message[];
  streamRef: RefObject<HTMLElement | null>;
  onNavigate: (id: string) => void;
};

/** Local navigation only: no model calls, history mutation or execution callbacks. */
export default function ConversationNavigation({ messages, streamRef, onNavigate }: Props) {
  const turns = useMemo(() => messages.filter((message) => message.role === "user"), [messages]);
  const [active, setActive] = useState<string>();
  const [expanded, setExpanded] = useState(false);
  const ticksRef = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    const ticks = ticksRef.current;
    const selected = ticks?.querySelector<HTMLElement>("[aria-current]");
    if (!ticks || !selected) return;
    // Keep the current tick visible in long histories; never scroll the page.
    if (selected.offsetTop < ticks.scrollTop) ticks.scrollTop = selected.offsetTop;
    else if (selected.offsetTop + selected.offsetHeight > ticks.scrollTop + ticks.clientHeight) {
      ticks.scrollTop = selected.offsetTop + selected.offsetHeight - ticks.clientHeight;
    }
  }, [active, expanded]);

  useLayoutEffect(() => {
    const stream = streamRef.current;
    if (!stream || turns.length < 2) return;
    const rows = Array.from(stream.querySelectorAll<HTMLElement>("[data-message-id]"));
    const byId = new Map(rows.map((row) => [row.dataset.messageId, row]));
    const update = () => {
      const readingLine = stream.getBoundingClientRect().top + 32;
      let current = turns[0].id;
      for (const turn of turns) {
        const row = byId.get(turn.id);
        if (row && row.getBoundingClientRect().top <= readingLine) current = turn.id;
      }
      setActive(current);
    };
    update();
    stream.addEventListener("scroll", update, { passive: true });
    const observer = typeof ResizeObserver === "undefined" ? null : new ResizeObserver(update);
    observer?.observe(stream);
    rows.forEach((row) => observer?.observe(row));
    return () => {
      stream.removeEventListener("scroll", update);
      observer?.disconnect();
    };
  }, [turns, streamRef]);

  if (turns.length < 2) return null;
  return (
    <nav className="conversation-navigation" data-expanded={expanded} aria-label="对话定位"
      onMouseEnter={() => setExpanded(true)}
      onMouseLeave={() => setExpanded(false)}
      onFocus={() => setExpanded(true)}
      onBlur={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setExpanded(false);
      }}
      onKeyDown={(event) => {
        if (event.key === "Escape") { event.preventDefault(); setExpanded(false); }
      }}>
      <div ref={ticksRef} className="conversation-navigation-ticks">
        {turns.map((turn, index) => {
          const label = turn.text?.trim().replace(/\s+/g, " ").slice(0, 160) || "图片或附件消息";
          return (
            <button key={turn.id} type="button" aria-label={`第 ${index + 1} 轮：${label}`}
              aria-current={active === turn.id ? "location" : undefined}
              onClick={() => {
                if (!expanded) { setExpanded(true); return; }
                onNavigate(turn.id); setActive(turn.id);
              }}>
              <span className="conversation-navigation-label" aria-hidden="true">{label}</span>
              <i className="conversation-navigation-tick" aria-hidden="true" />
            </button>
          );
        })}
      </div>
    </nav>
  );
}
