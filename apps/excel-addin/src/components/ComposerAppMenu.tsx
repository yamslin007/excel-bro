import { useEffect, useRef, type ReactNode } from "react";

type Props = {
  logoSrc: string;
  open: boolean;
  onToggle: () => void;
  onClose: () => void;
  children: ReactNode;
};

export function ComposerAppMenu({ logoSrc, open, onToggle, onClose, children }: Props) {
  const root = useRef<HTMLDivElement>(null);
  const trigger = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    if (!open) return;
    const closeOutside = (event: PointerEvent) => {
      if (!root.current?.contains(event.target as Node)) onClose();
    };
    document.addEventListener("pointerdown", closeOutside);
    return () => document.removeEventListener("pointerdown", closeOutside);
  }, [open, onClose]);

  return <div className="composer-app-menu view-menu-wrap" ref={root}
    onBlur={(event) => {
      if (event.relatedTarget && !event.currentTarget.contains(event.relatedTarget as Node)) onClose();
    }}
    onKeyDown={(event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        onClose();
        trigger.current?.focus();
      }
      if (open && ["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) {
        const items = [...event.currentTarget.querySelectorAll<HTMLButtonElement>('[role="menuitem"]:not(:disabled), [role="menuitemcheckbox"]:not(:disabled)')];
        if (!items.length) return;
        event.preventDefault();
        const index = items.indexOf(document.activeElement as HTMLButtonElement);
        const next = event.key === "Home" ? 0 : event.key === "End" ? items.length - 1
          : index < 0 ? (event.key === "ArrowUp" ? items.length - 1 : 0)
          : (index + (event.key === "ArrowUp" ? -1 : 1) + items.length) % items.length;
        items[next].focus();
      }
    }}>
    <button ref={trigger} className="composer-app-trigger" type="button"
      aria-label="EB 菜单" title="EB 菜单" aria-haspopup="menu" aria-expanded={open} onClick={onToggle}>
      <img src={logoSrc} alt="" />
    </button>
    {open && <div className="view-menu" role="menu" aria-label="EB 菜单">{children}</div>}
  </div>;
}
