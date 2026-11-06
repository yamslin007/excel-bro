import { useEffect, useState, type ReactNode } from "react";
import {
  composerPendingSummary,
  type ComposerPendingItem
} from "@excel-bro/core/composerPending";

/**
 * 输入容器**上方**的常驻待处理事项（Excel / WPS 共用）。
 *
 * 结构纪律：`.composer-dock`（工具栏 + 输入框）里不许夹通知、错误、恢复摘要、
 * 历史执行结果。所有这类内容都作为 `children` 放进这块可展开区域，由本组件
 * 统一给出一行紧凑摘要。底部高度因此只随"有没有事项"变一次，不会每来一条
 * 短反馈就把工具栏从输入框旁边挤开。
 *
 * 用原生 `<details>` 而不是 `overflow:hidden` / 负边距 / 绝对定位：收起时详细
 * 内容仍在 DOM 与可访问语义里，点开就能看到，不存在"为了布局把入口藏掉"。
 * 需要用户做决定的事项（正式待审阅的决定、恢复候选）默认展开，不藏主入口。
 */
export function ComposerPendingBar({
  items,
  children
}: {
  items: readonly ComposerPendingItem[];
  children: ReactNode;
}) {
  const summary = composerPendingSummary(items);
  const [open, setOpen] = useState(summary.defaultOpen);
  // 一旦出现必须点的事项（例如进入正式待审阅），立刻展开，不等用户去找。
  useEffect(() => {
    if (summary.defaultOpen) setOpen(true);
  }, [summary.defaultOpen]);
  if (summary.count === 0) return null;
  return (
    <details
      className="composer-pending"
      open={open}
      data-pending-count={summary.count}
      data-pending-open={open ? "true" : "false"}
    >
      {/*
        `open` 由 React 控制，点击摘要**自己**翻转状态并 `preventDefault()`：
        原生 `<summary>` 的展开是宿主默认行为，jsdom 不实现它 —— 只靠 `onToggle`
        会让"点一下能不能展开"既测不到也控不住。键盘 Enter/Space 同样走 click，
        所以这一条覆盖鼠标与键盘两条路。
      */}
      <summary
        className="composer-pending-summary"
        aria-label={`${summary.text}，可展开查看详情`}
        onClick={(event) => {
          event.preventDefault();
          setOpen((current) => !current);
        }}
      >
        <span className="composer-pending-dot" aria-hidden="true" />
        <span className="composer-pending-text">{summary.text}</span>
        <span className="composer-pending-toggle" aria-hidden="true">{open ? "收起" : "展开"}</span>
      </summary>
      <div className="composer-pending-details">{children}</div>
    </details>
  );
}

export default ComposerPendingBar;
