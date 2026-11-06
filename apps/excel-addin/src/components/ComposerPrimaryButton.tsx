import type {
  ComposerPrimaryAction,
  PrimaryActionInvoke
} from "@excel-bro/core/composerPrimaryAction";
import "./composerPrimaryButton.css";

/**
 * 输入框右下角那**一颗**主操作按钮（Excel / WPS 共用同一个组件）。
 *
 * 改造前这里按阶段散着三套东西：输入框按钮（↑/■/▶）、输入区上方的运行控制条、
 * 以及"暂停请求在飞"时的临时态。同一个位置在不同阶段长出不同的按钮，用户只能靠试。
 * 现在它的形态、文案、禁用与"按下去走哪条链路"**全部**来自共享层的
 * `composerPrimaryAction`（纯判据），本组件只负责把它画出来、并把 `invoke`
 * 原样交回调用方 —— 两端都不许在这里另写一套状态判断。
 *
 * 两条被真机验证过的细节，不要顺手改掉：
 *
 *  1. `key={action.kind}`：形态切换必须换一个 DOM 节点。否则同一个按钮会从 ■ 变成 ▶，
 *     而浏览器那次点击的 click 计数（`event.detail`）还留着 —— 双击的第二下会被
 *     当成"点继续"，于是一次暂停连出一次续跑。
 *  2. `event.detail > 1` 直接返回：双击不产生第二个动作。（① 保证换节点后计数归零，
 *     ② 是在换节点之前的最后一道闸；两条各自独立，不能只留一条。）
 *
 * 禁用与"正在发生"分开：`action.busy`（暂停请求在飞）既不可点也要显示 `…`，
 * 而 `action.disabled` 可能是"写入/回读期间不能中断"——后者必须在 `title` 里
 * **说清原因**，绝不画一颗点了没反应的按钮。
 */
export function ComposerPrimaryButton({
  action,
  onInvoke,
  className = "send-button",
  disabled = false,
  disabledHint
}: {
  action: ComposerPrimaryAction;
  onInvoke: (invoke: NonNullable<PrimaryActionInvoke>) => void;
  /** 宿主侧的额外类名（WPS 的 `.composer-send-button` 仍要保留）。 */
  className?: string;
  /** 宿主侧额外禁用条件（没有工作簿、图片还在处理、模型不支持视觉…）。 */
  disabled?: boolean;
  /** 上面那条禁用条件的**原因**；有它时覆盖 `title`，别让人猜为什么点不动。 */
  disabledHint?: string;
}) {
  const blocked = disabled || action.disabled || action.busy;
  const stateClass =
    action.kind === "send-supplement"
      ? "is-steer"
      : action.kind === "pause"
        ? "is-stop"
        : action.kind === "continue"
          ? "is-continue"
          : action.kind === "reconcile"
            ? "is-reconcile"
            : action.kind === "pausing"
              ? "is-pausing"
              : "";
  const title = !action.disabled && disabled && disabledHint ? disabledHint : action.title;
  return (
    <button
      key={action.kind}
      type="button"
      className={[className, "composer-primary", stateClass].filter(Boolean).join(" ")}
      data-primary-action={action.kind}
      aria-label={action.label}
      title={title}
      aria-busy={action.busy || undefined}
      disabled={blocked}
      onClick={(event) => {
        event.preventDefault();
        if (event.detail > 1) return;
        if (!action.invoke) return;
        onInvoke(action.invoke);
      }}
    >
      {action.kind === "answer" ? (
        "提交回答"
      ) : action.glyph === "↑" ? (
        /* 发送/补充：向上的实心箭头；操作判据和防双击行为保持不变。 */
        <svg
          className="composer-send-arrow"
          viewBox="0 0 24 24"
          aria-hidden="true"
          focusable="false"
        >
          <path d="M10.7 3.9a1.45 1.45 0 0 1 2.6 0l7.1 15.3c.6 1.2-.7 2.5-1.9 1.9L12 17.9l-6.5 3.2c-1.2.6-2.5-.7-1.9-1.9Z" />
        </svg>
      ) : (
        action.glyph
      )}
    </button>
  );
}
