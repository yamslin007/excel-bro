type ScrollToLatestButtonProps = {
  /** EB 是否正在跑。运行中显示三点动效，停下后换成锚图标。 */
  running: boolean;
  onClick: () => void;
};

/**
 * 「有新进度 / 回到底部」浮标 —— Excel 与 WPS 双端共用。
 *
 * 同一个控件、同一个位置、同一种点击结果（回到底部），只有字形随运行状态变：
 *
 * - `running`：EB 正在跑 —— 三颗点聚合/环绕（位形与动效见 `.progress-dots`）。
 *   它表达的是"下面还在长东西"，所以必须动。
 * - `idle`：EB 已经停下 —— 换成**实心剪影锚**图标。此时下面不再增长，是一个已经落定的
 *   位置；锚表达"内容停在那里、点我下去"，与三点形成"动 / 静"的对照。
 *
 * 字形只由 `data-mode` 区分，两者共用同一层品牌浅色底与同一个前景令牌
 * `--color-progress-glyph`（跟随主题）。这里**不写任何色值**：
 * 宿主引擎对 `color-mix()` 之类运行时色彩函数的支持不可假设，写在 CSS 里的
 * 令牌才是确定值。
 */
export default function ScrollToLatestButton({ running, onClick }: ScrollToLatestButtonProps) {
  return (
    <button
      type="button"
      className="new-progress-button"
      data-mode={running ? "running" : "idle"}
      onClick={onClick}
      aria-label={running ? "查看最新执行进度" : "回到底部查看最新内容"}
      title={running ? "查看最新进度" : "回到底部"}
    >
      {running ? (
        <span className="progress-dots" aria-hidden="true">
          <span />
          <span />
          <span />
        </span>
      ) : (
        <svg
          className="progress-anchor-icon"
          viewBox="0 0 24 24"
          aria-hidden="true"
          focusable="false"
        >
          {/*
            **实心剪影锚**（2026-10-12 用户裁决：换成参考图那种实心锚，不再是一根
            细线描出来的）。形状完全由这里的 `d` 与几何决定，CSS 只负责填色 ——
            所以 `.progress-anchor-icon` 那边是 `fill`，**没有** `stroke`。

            为什么是五段而不是一条 path：
              · 锚眼环**中空**，靠 `fillRule="evenodd"` 用外圆减掉内圆。拆成两条
                独立 path 会各填各的，内圆反而被填实、锚眼糊死。
              · 横杆是一根**两端带球的短杆**：杆用 `<rect>`（比 path 弧线稳），
                两端球各一个 `<circle>`，就是参考图里那两颗端点。
              · 锚杆 + 底部尖角是一条 path（向下收成龙骨尖）。
              · 锚臂是一条**闭合的月牙**：两端自然收成尖爪，不需要单独的爪路径。
          */}
          <path
            fillRule="evenodd"
            d="M12 1.4a2.8 2.8 0 1 1 0 5.6 2.8 2.8 0 0 1 0-5.6Zm0 1.5a1.3 1.3 0 1 0 0 2.6 1.3 1.3 0 0 0 0-2.6Z"
          />
          <rect x="7.3" y="8.3" width="9.4" height="1.4" rx="0.7" />
          <circle cx="6" cy="9" r="1.5" />
          <circle cx="18" cy="9" r="1.5" />
          <path d="M11.15 6.3h1.7l.3 13.2L12 23l-1.15-3.5Z" />
          <path d="M1.7 10.1C3.7 17.7 6.7 21.8 12 22c5.3-.2 8.3-4.3 10.3-11.9-2.8 6.3-5.4 8.9-10.3 9.1-4.9-.2-7.5-2.8-10.3-9.1Z" />
        </svg>
      )}
    </button>
  );
}
