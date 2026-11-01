import type { ReactNode } from "react";
import {
  currentActionOtherCount,
  currentActionSwitchTargets,
  type CurrentActionKind,
  type CurrentActionSelection
} from "@excel-bro/core/currentActionCard";
import "./composerActionDock.css";

export type CurrentActionView = Exclude<CurrentActionKind, "none">;

/**
 * 输入区上方的**固定区域容器**（Excel / WPS 共用）。
 *
 * 改造前这里是四块互不知情的界面各自决定自己出不出现：结果审阅栏、写入恢复卡
 * （含多候选选择卡）、澄清回答区、当前事项说明。四者可以同时在场，视觉分量又都差不多，
 * "现在该做什么"要靠用户自己从四块里挑。`selectCurrentAction` 当时只算出一个名字，
 * 四块照样全展开 —— 名字算对了，界面没变。
 *
 * 现在这个容器只让**一块卡片**展开，其余用一行紧凑提示 + **显式切换入口**保留：
 *
 * - 展开的是 `selection.kind` 对应的那一块（`review` / `recovery`）；
 * - 其余的写成 chip：能切的给按钮（点了立刻换主视图），不能切的如实说明它在哪
 *   —— 澄清回答区**永远不折叠**（它贴着输入框，见下）、运行控制本来就在右下角；
 * - 用户的显式切换（`pinned`）压过默认优先级，但**只认仍然存在的**那件事：
 *   被处理掉之后马上退回默认顺序，不留一张空卡。
 *
 * 为什么澄清不参与折叠：它改变**输入框里那句话的意思**（发送 = 提交回答）。
 * 把它折起来而输入框仍是"回答模式"，用户会对着一道看不见的题作答；
 * 反过来把它折起来又切回"新需求模式"，等于替用户改了他正在写的东西的语义。
 * 所以它和输入框同层，只以 chip 的形式在同一行里说明"下面就是它"。
 */
export function ComposerActionDock({
  selection,
  views,
  onSwitch
}: {
  selection: CurrentActionSelection;
  /** 各视图的节点；`null` = 这件事现在没有界面可展开。澄清不在这里（见上）。 */
  views: { review: ReactNode; recovery: ReactNode };
  /** 切到另一件事。没给 = 只说明不切（例如澄清就在下方）。 */
  onSwitch?: (kind: CurrentActionView) => void;
}) {
  const primary: CurrentActionView | null =
    selection.kind === "review" || selection.kind === "recovery" ? selection.kind : null;
  /*
   * 其余事项只从 `currentActionSwitchTargets` 取，**不在这里另算一份**：
   * 另算就会出现"提示里说还有一项，切过去却没有"（`others` 与提示必须同源）。
   */
  const others = currentActionSwitchTargets(selection);
  if (primary === null && others.length === 0) return null;

  const switchable = (kind: CurrentActionView): boolean =>
    Boolean(onSwitch) && (kind === "review" || kind === "recovery");

  return (
    <section
      className="composer-action-dock"
      aria-label="待处理事项"
      data-dock-primary={selection.kind}
      data-dock-others={currentActionOtherCount(selection)}
    >
      {primary !== null && views[primary] ? (
        <div className="composer-action-dock-view" data-dock-view={primary}>
          {views[primary]}
        </div>
      ) : null}
      {others.length > 0 ? (
        <p className="composer-action-dock-others" role="status">
          <span>另有 {others.length} 项：</span>
          {others.map(({ kind, label, short }) => {
            if (switchable(kind)) {
              return (
                <button
                  key={kind}
                  type="button"
                  className="composer-action-dock-chip"
                  data-dock-switch={kind}
                  /* 完整说法当 title / 可访问名，短名当可见文案：一行放得下，读屏也听得懂。 */
                  title={`先处理这一项：${label}`}
                  aria-label={`先处理这一项：${label}`}
                  onClick={() => onSwitch?.(kind)}
                >
                  {short}
                </button>
              );
            }
            /*
             * 不能切的两种，都**如实说明它在哪**，而不是画一个点了没反应的按钮：
             * 澄清回答区就贴在输入框上方（本来就能看见），运行控制就在右下角。
             */
            return (
              <span key={kind} className="composer-action-dock-chip is-inplace" data-dock-inplace={kind}>
                {short}
                {kind === "clarification" ? "（就在下方）" : kind === "runtime" ? "（控制在右下角）" : ""}
              </span>
            );
          })}
        </p>
      ) : null}
    </section>
  );
}

export default ComposerActionDock;
