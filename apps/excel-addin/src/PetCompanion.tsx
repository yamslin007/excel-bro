import { useMemo, useState } from "react";
import { PET_TIPS, tipAt } from "./petTips";
import { brandLogoUrl } from "./brandLogo";

interface PetCompanionProps {
  busy: boolean;
  onHide(): void;
}

// EB logo 作为浮动助手；点击展开技巧，显示开关位于 EB 菜单。
export default function PetCompanion({ busy, onHide }: PetCompanionProps) {
  const [tipIndex, setTipIndex] = useState(() =>
    Math.floor(Math.random() * PET_TIPS.length)
  );
  const [open, setOpen] = useState(false);
  const tip = useMemo(() => tipAt(tipIndex), [tipIndex]);
  const mood = busy ? "busy" : "idle";

  return (
    <div className={`pet-float mood-${mood}`}>
      <button
        type="button"
        className="pet-avatar"
        onClick={() => setOpen((current) => !current)}
        onDoubleClick={() => { setOpen(false); onHide(); }}
        aria-label={open ? "收起技巧" : "查看技巧"}
        aria-expanded={open}
        title={open ? "收起技巧；双击隐藏助手" : "查看一条 Excel 技巧；双击隐藏助手"}
      >
        <img className="pet-brand-logo" src={brandLogoUrl} alt="EB 助手" draggable={false} />
      </button>
      {open && (
        <div className="pet-bubble" role="dialog" aria-label="EB 小提示">
          <div className="pet-bubble-head">
            {tip.category && (
              <span className="pet-bubble-eyebrow">{tip.category}</span>
            )}
            <button
              type="button"
              className="pet-next"
              onClick={() => setTipIndex((current) => current + 1)}
            >
              换一条
            </button>
          </div>
          <strong>{tip.title}</strong>
          <p>{tip.detail}</p>
          {tip.shortcut && <kbd className="pet-shortcut">{tip.shortcut}</kbd>}
        </div>
      )}
    </div>
  );
}
