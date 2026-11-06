/**
 * 底部「常驻待处理事项」的紧凑摘要（Excel / WPS 共用）。
 *
 * 底部固定容器（工具栏 + 输入框）里**只准有这两块**；所有通知、状态、历史摘要都
 * 归到它上方这块可展开区域。摘要只给数量与最要紧的一两项，不把详细内容铺开，
 * 这样底部高度不会随每一条短反馈来回跳。
 *
 * `decision` 类事项（正式待审阅的决定、恢复候选选择）默认展开：那是"要用户点的
 * 那一下"，不能为了紧凑把它藏起来。其余状态/通知默认收起，点一下即可展开。
 */
export type ComposerPendingKind = "decision" | "recovery" | "status" | "notice";

export type ComposerPendingItem = {
  id: string;
  label: string;
  kind: ComposerPendingKind;
};

export type ComposerPendingSummary = {
  count: number;
  text: string;
  /** 摘要里写下的事项名（最多两条），供无障碍标签与布局探针读取。 */
  labels: string[];
  defaultOpen: boolean;
};

export function composerPendingSummary(
  items: readonly ComposerPendingItem[]
): ComposerPendingSummary {
  const count = items.length;
  if (count === 0) return { count: 0, text: "", labels: [], defaultOpen: false };
  // 决定 > 恢复 > 状态 > 通知：摘要里先说真正要用户处理的那件。
  const rank: Record<ComposerPendingKind, number> = { decision: 0, recovery: 1, status: 2, notice: 3 };
  const ordered = [...items].sort((a, b) => rank[a.kind] - rank[b.kind]);
  const labels = ordered.slice(0, 2).map((item) => item.label);
  const rest = count > labels.length ? ` 等 ${count} 项` : "";
  return {
    count,
    text: `待处理 ${count} 项 · ${labels.join("、")}${rest}`,
    labels,
    defaultOpen: items.some((item) => item.kind === "decision")
  };
}
