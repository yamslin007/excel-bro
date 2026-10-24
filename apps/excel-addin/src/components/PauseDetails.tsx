import type { ReactNode } from "react";

/** Routine pause is one line; recovery diagnostics remain available on demand. */
export function PauseDetails({ children, blocked = false, active = false, resolved = false }: {
  children: ReactNode; blocked?: boolean; active?: boolean; resolved?: boolean;
}) {
  return <details className="pause-details" open={blocked || undefined}>
    <summary>{resolved ? "已接续" : active ? "正在继续…" : blocked ? "暂时无法继续" : "已暂停"}
      <span>详情</span>
    </summary>
    {children}
  </details>;
}
