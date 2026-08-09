import { useMemo, useState } from "react";
import { PET_TIPS, tipAt } from "./petTips";

interface PetCompanionProps {
  busy: boolean;
}

// 格仔：一块带单元格网格纹的绿色果冻，作为左上角浮窗常驻。
// 表情随状态变化：idle 好奇、busy 眯眼专注。
// 隐藏格仔走头部按钮，点格仔本身只展开/收起技巧气泡。
// 点击格仔展开/收起技巧气泡；气泡里可换技巧或隐藏格仔。
export default function PetCompanion({ busy }: PetCompanionProps) {
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
        aria-label={open ? "收起技巧" : "查看技巧"}
        aria-expanded={open}
        title={open ? "收起技巧" : "查看一条 Excel 技巧"}
      >
        <PetJelly mood={mood} />
      </button>
      {open && (
        <div className="pet-bubble" role="dialog" aria-label="格仔小提示">
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

function PetJelly({ mood }: { mood: "idle" | "busy" }) {
  return (
    <svg viewBox="0 0 72 72" role="img" aria-hidden="true" className="pet-jelly">
      <defs>
        <linearGradient id="pet-jelly-body" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#a6e6bd" />
          <stop offset="100%" stopColor="#6fc98c" />
        </linearGradient>
        <clipPath id="pet-jelly-clip">
          <rect x="8" y="10" width="56" height="52" rx="16" />
        </clipPath>
      </defs>

      <g className="pet-jelly-shape">
        <rect
          x="8"
          y="10"
          width="56"
          height="52"
          rx="16"
          fill="url(#pet-jelly-body)"
          stroke="#4fae72"
          strokeWidth="1.5"
        />
        {/* 单元格网格纹 */}
        <g clipPath="url(#pet-jelly-clip)" stroke="#57b87c" strokeWidth="1" opacity="0.5">
          <line x1="26" y1="10" x2="26" y2="62" />
          <line x1="44" y1="10" x2="44" y2="62" />
          <line x1="8" y1="28" x2="64" y2="28" />
          <line x1="8" y1="45" x2="64" y2="45" />
        </g>
        {/* 高光 */}
        <path
          className="pet-shine"
          d="M18 18 l3 -3 M16 22 q4 -6 10 -7"
          stroke="#ffffff"
          strokeWidth="2.4"
          strokeLinecap="round"
          fill="none"
          opacity="0.85"
        />
        {/* 腮红 */}
        <circle cx="22" cy="44" r="4" fill="#f4a9b8" opacity="0.55" />
        <circle cx="50" cy="44" r="4" fill="#f4a9b8" opacity="0.55" />
        {/* 眼睛 */}
        <g className="pet-eyes" fill="#28402f">
          {mood === "busy" ? (
            <>
              <path d="M24 37 q4 -3 8 0" stroke="#28402f" strokeWidth="2.4" strokeLinecap="round" fill="none" />
              <path d="M40 37 q4 -3 8 0" stroke="#28402f" strokeWidth="2.4" strokeLinecap="round" fill="none" />
            </>
          ) : (
            <>
              <circle cx="28" cy="38" r="3.4" />
              <circle cx="44" cy="38" r="3.4" />
              <circle cx="29.2" cy="36.8" r="1" fill="#fff" />
              <circle cx="45.2" cy="36.8" r="1" fill="#fff" />
            </>
          )}
        </g>
        {/* 嘴巴 */}
        {mood === "busy" ? (
          <path d="M33 48 q3 3 6 0" stroke="#28402f" strokeWidth="2" strokeLinecap="round" fill="none" />
        ) : (
          <ellipse cx="36" cy="48" rx="2.4" ry="3" fill="#28402f" />
        )}
      </g>
    </svg>
  );
}
