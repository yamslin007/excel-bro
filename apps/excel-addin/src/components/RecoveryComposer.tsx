import type { ReactNode } from "react";
import type { ExecutionRecoveryAction } from "@excel-bro/core/executionRecoveryView";
import {
  RECOVERY_COMPOSER_MODE_LABEL,
  type RecoveryComposerEntry,
  type RecoveryComposerView
} from "@excel-bro/core/recoveryComposer";
import "./recoveryComposer.css";

/**
 * 恢复待办**所在的那块输入区**（Excel / WPS 共用，嵌在 composer-box 内部）。
 *
 * 与澄清问卷同一套形态：普通输入框被它**置换**，而不是在输入框上再叠一张卡。
 * 这样"现在该点哪一下"只有一个答案 —— 输入框里那颗文字按钮。
 *
 * 三条边界（与 `recoveryComposer.ts` 的判据配套）：
 *
 * 1. **它不自己判断**：标题、正文、按钮文案全部来自共享层给的 `view`。这里只排版。
 *    WPS 与 Excel 因此不会各说一套"第几步/还剩几步/写了哪几格"。
 * 2. **它不自己发请求**：`onPrimary` 就是调用方原来的核对 / 续跑处理器，一下点击一次。
 *    没有回车、没有自动续跑：核对完成后这个组件会被换成"继续剩余 N 步"，等用户再点。
 * 3. **技术记录不在这里复述**：进度与结论正文折叠在「查看详情」里，与消息流里的
 *    只读执行记录同一份来源（`executionRecoveryProgress` / `executionRecoveryDetail`）。
 *
 * 因为输入区接管之后固定区域那张恢复大卡会被收走（同一个结论不画两遍），大卡上
 * **只有它能承接**的两样东西也挪进「查看详情」：撤销行（Excel 端大卡的唯一撤销入口）
 * 与开发诊断。它们不是主操作，所以不进动作行，也不跟主按钮抢位置。
 *
 * 布局纪律：动作行是滚动正文的**兄弟**，不是后代 —— 320px 窄窗格里正文自己滚，
 * 按钮永远留在原地（与数据范围面板、恢复卡同一条）。
 */
export function RecoveryComposer({
  view,
  busy,
  acting,
  detailsOpen,
  dismiss,
  undo,
  diagnostics,
  onPrimary,
  onNewRequest,
  onToggleDetails,
  onDismiss,
  onUndo
}: {
  view: RecoveryComposerView;
  /** 正在查询或提交：按钮禁用（不是隐藏 —— 隐藏会让人以为点错了）。 */
  busy: boolean;
  /** 正在执行的具体动作（决定等待文案）。 */
  acting?: ExecutionRecoveryAction | null;
  detailsOpen: boolean;
  /** 「清除恢复提示」：只有本端真能落地时才给（Excel 没有处置端点 ⇒ 传 null）。 */
  dismiss?: { label: string; hint: string } | null;
  /**
   * 「撤销本次执行」：大卡被收走之后它唯一的入口（WPS 的撤销在「备份与恢复」抽屉里，
   * 那边不传）。`state` 不是 `available` 时只显示一句话，不画一颗点下去只会报错的按钮。
   */
  undo?: {
    state: "available" | "blocked" | "done";
    label?: string;
    detail: string;
    onUndo?: () => void;
  } | null;
  /** 开发诊断（仅 DEV 由调用方提供）：只放原因码与计数，不放单元格内容。 */
  diagnostics?: ReactNode;
  onPrimary: () => void;
  onNewRequest: () => void;
  onToggleDetails: () => void;
  onDismiss?: () => void;
  onUndo?: () => void;
}) {
  const waiting = busy || acting === view.action;
  /*
   * `acting` 缺省是 `undefined`，而"有没有动作在跑"是 `!== null` —— 直接比会把
   * `undefined` 当成"在跑"，撤销按钮就永远点不动（组件测试逮到过）。这里归一成 null。
   */
  const actingNow = acting ?? null;
  const primaryLabel = waiting
    ? view.action === "reconcile" ? "正在检查…" : "正在继续…"
    : view.primaryLabel;
  return (
    <section
      className="recovery-composer"
      aria-label="恢复待办"
      data-recovery-composer={view.action}
      data-recovery-busy={waiting ? "true" : "false"}
    >
      <header className="recovery-composer-heading">
        <span className="recovery-composer-mode">{RECOVERY_COMPOSER_MODE_LABEL}</span>
        <p className="recovery-composer-headline">{view.headline}</p>
      </header>
      <div className="recovery-composer-scroll">
        {view.lines.map((line) => (
          <p key={line} className="recovery-composer-line">{line}</p>
        ))}
        {/*
          撤销行：**不**收进「查看详情」，它是安全相关的可见入口。
          大卡被收走之后，撤销是这张卡上唯一不能由主按钮承接的东西 —— 它回答
          "已经写进去的怎么退回来"，与"这次任务怎么收尾"是两件事。收进折叠的详情里，
          就等于把一个"能退回这次写入"的入口藏起来（真机 B 现场就靠它兜底）。
        */}
        {undo && undo.state !== "available" && (
          <p
            className={
              undo.state === "done"
                ? "recovery-composer-undo is-done"
                : "recovery-composer-undo is-blocked"
            }
            role="status"
            data-recovery-undo={undo.state}
          >
            {undo.detail}
          </p>
        )}
        {undo?.state === "available" && (
          <div className="recovery-composer-undo-row" data-recovery-undo="available">
            <button
              type="button"
              className="recovery-composer-secondary"
              data-recovery-action="undo"
              disabled={busy || actingNow !== null}
              onClick={(event) => {
                /* 双击的第二下不派发：撤销也是一次真实的宿主写入，不能连点两遍。 */
                if (event.detail > 1) return;
                (onUndo ?? undo.onUndo)?.();
              }}
            >
              {busy ? "正在撤销…" : undo.label ?? "撤销本次执行"}
            </button>
            <span className="recovery-composer-dismiss-hint">{undo.detail}</span>
          </div>
        )}
        {/*
          详情：技术口径（进度 / 结论 / 边界）+ 两样收进来的次要东西
          （「清除恢复提示」/ 开发诊断）。消息流里的只读执行记录一条不少，
          这里只把"要不要展开"交给用户。
        */}
        {detailsOpen && (
          <div className="recovery-composer-details" data-recovery-details="open">
            <p className="recovery-composer-progress">{view.progress}</p>
            <p className="recovery-composer-detail">{view.detail}</p>
            <p className="recovery-composer-note">{view.note}</p>
            {dismiss && onDismiss && (
              <div className="recovery-composer-dismiss">
                <button
                  type="button"
                  className="recovery-composer-secondary"
                  data-recovery-action="dismiss"
                  disabled={busy}
                  onClick={onDismiss}
                >
                  {dismiss.label}
                </button>
                <span className="recovery-composer-dismiss-hint">{dismiss.hint}</span>
              </div>
            )}
            {diagnostics && (
              <details className="recovery-composer-diagnostics">
                <summary>技术诊断（仅开发）</summary>
                {diagnostics}
              </details>
            )}
          </div>
        )}
      </div>
      <div className="recovery-composer-actions">
        <button
          type="button"
          className="recovery-composer-ghost"
          data-recovery-composer-action="new-request"
          onClick={onNewRequest}
        >
          另提需求
        </button>
        <button
          type="button"
          className="recovery-composer-ghost"
          data-recovery-composer-action="details"
          aria-expanded={detailsOpen}
          onClick={onToggleDetails}
        >
          查看详情
        </button>
        <button
          type="button"
          className="recovery-composer-primary"
          data-recovery-primary={view.action}
          /*
           * `key={view.action}`：核对 → 继续换的是**节点**，不只是文案。
           * 双击/连点时，第二下若落在同一个节点上，就会在"检查"刚结束、视图已变成
           * "继续剩余 N 步"之后触发一次**真实写入** —— 用户以为自己在检查，结果写了表。
           * 换节点让那一下没有可落的目标（与 `ComposerPrimaryButton` 的 `key={kind}` 同一条纪律）。
           */
          key={view.action}
          title={view.primaryTitle}
          aria-label={waiting ? primaryLabel : view.primaryLabel}
          disabled={waiting}
          onClick={(event) => {
            /*
             * 双击的第二下（`detail > 1`）不派发：一次点击一次调用。
             * 键盘激活（回车/空格）的 `detail` 是 0，仍然放行 —— 那是用户明确按下去的。
             */
            if (event.detail > 1) return;
            onPrimary();
          }}
        >
          {primaryLabel}
        </button>
      </div>
    </section>
  );
}

/**
 * 「另提需求」之后留在固定区域的那一枚**紧凑待办入口**。
 *
 * 它与 `RecoveryComposer` 同一时刻**只有一个**：占据输入框时不需要入口，
 * 让出输入框时只要这一行 —— 这就是"当前主操作只保留一处"的落地方式。
 *
 * 它**不是**撤销、不是放弃：点它只是把输入框换回恢复操作区，任务、材料、账本一个字不动。
 */
export function RecoveryPendingEntry({
  entry,
  busy,
  onReturn
}: {
  entry: RecoveryComposerEntry;
  busy: boolean;
  onReturn: () => void;
}) {
  return (
    <p className="recovery-pending-entry" role="status">
      <span>这件事还等着你处理：</span>
      <button
        type="button"
        className="recovery-pending-entry-button"
        data-recovery-entry={entry.action}
        title={entry.title}
        disabled={busy}
        onClick={onReturn}
      >
        {entry.label}
      </button>
    </p>
  );
}

export default RecoveryComposer;
