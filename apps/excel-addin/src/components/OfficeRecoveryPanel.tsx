import { RecoveryTerminalRecord } from "./RecoveryTerminalRecord";
import {
  EXECUTION_RECOVERY_ACTION_LABEL,
  EXECUTION_RECOVERY_ENTRY_HINT,
  EXECUTION_RECOVERY_TITLE,
  executionRecoveryActions,
  executionRecoveryBadge,
  executionRecoveryDetail,
  executionRecoveryNote,
  executionRecoveryProgress,
  executionRecoveryTerminalReason,
  executionRecoveryTerminalSummary,
  type ExecutionRecoveryAction,
  type ExecutionRecoveryKind
} from "@excel-bro/core/executionRecoveryView";
import type { ChatMessage } from "../types/chat";

type OfficeRecoveryCard = NonNullable<ChatMessage["officeResume"]>;

/**
 * Excel 端**真的能执行**的恢复动作。
 *
 * 结论与按钮集合来自共享层（与 WPS 同一份），但"能不能落地"是宿主事实：
 * `replan`（重新读取并规划）与 `dismiss`（把"已核对"写进服务端的处置元数据）
 * 目前只有 WPS 有 —— Excel 的 `/api/office/executions` 没有处置端点，重新规划也要
 * 另起一轮任务。共享结论给出这几个动作时，卡片改给「重新核对（只读）」，
 * 不画一个点了没反应的按钮。
 */
export const EXCEL_RECOVERY_ACTIONS: readonly ExecutionRecoveryAction[] = [
  "reconcile",
  "resume",
  "backfill"
];

/** 动作进行中的按钮文案（空闲时用共享层的 `EXECUTION_RECOVERY_ACTION_LABEL`）。 */
const EXECUTION_RECOVERY_ACTIVE_LABEL: Record<ExecutionRecoveryAction, string> = {
  reconcile: "正在核对…",
  resume: "正在继续…",
  backfill: "正在补记…",
  replan: "正在重新规划…",
  dismiss: "正在清除…"
};

/** 卡片上真正可点的动作：共享结论 ∩ 本端能落地。 */
export function excelRecoveryActions(card: OfficeRecoveryCard): ExecutionRecoveryAction[] {
  return executionRecoveryActions(card.verdict).filter((action) =>
    EXCEL_RECOVERY_ACTIONS.includes(action)
  );
}

/**
 * 卡面上**保留**的次要动作。
 *
 * `reconcile` 与 `resume` 已搬进**输入框里的恢复操作区**（2026-11-02 决定，
 * 前一版搬到了右下角那颗主按钮上）：同一个动作在 40px 内画两颗按钮，用户只能靠试。
 * 留在这里的是"补记完成回执"—— 它既不是核对也不是续跑，是一次独立的收尾动作。
 */
const OFFICE_RECOVERY_SECONDARY_ACTIONS: readonly ExecutionRecoveryAction[] = ["backfill"];

/**
 * 主操作现在承接哪一步（`null` = 这一步不由操作区承接）。
 *
 * 文案里的**按钮名**来自共享判据算出来的 `primaryLabel`（宿主按同一份 facts 算，
 * 不在这里另写一遍"核对在途动作"）—— 按钮上写的是「检查已写入结果」，指路句却叫
 * 另一个名字，等于让用户去找一颗不存在的按钮。位置也照实说：它在输入框里，不在卡片上。
 */
function officeRecoveryPrimaryHint(
  kind: ExecutionRecoveryKind,
  primaryLabel: string,
  primaryEntryVisible: boolean
): string | null {
  if (kind !== "reconcile" && kind !== "resume") return null;
  if (!primaryEntryVisible) {
    /*
     * 输入区归别的正事（审阅在等决定 / 澄清在等回答 / 正在跑）时，恢复操作区**不接管**
     * 输入框，那里只剩一块普通 textarea —— 再说"在输入框里"就是让用户去点一颗不存在的
     * 按钮。所以照实说"还没到、等上面收尾后出现"，**不给假入口，也不把恢复待办藏起来**
     * （它仍在固定区域的 chip 里，点一下就能看到这张卡）。
     */
    return kind === "reconcile"
      ? `下一步「${primaryLabel}」：等当前事项处理完后，入口会出现在输入框里 —— 它只读核对，不写任何单元格。`
      : `下一步「${primaryLabel}」：等当前事项处理完后，入口会出现在输入框里 —— 它只执行剩下的步骤，已写过的不会重写。`;
  }
  return kind === "reconcile"
    ? `下一步「${primaryLabel}」在输入框里：只读核对，不写任何单元格。`
    : `下一步「${primaryLabel}」在输入框里：只执行剩下的步骤，已写过的不会重写。`;
}

/**
 * 「上次写入的恢复核对」—— **唯一**可点击的入口，住在输入区上方。
 *
 * 为什么不在消息流里（2026-10-29 真机）：恢复卡原本画在产生它的那条消息下面。
 * 点「核对在途动作」会追加一行日志，消息区跟着自动滚到底，而更新成
 * 「继续剩余动作」的那张卡还停在旧消息的位置 —— 用户看到"已核对，未写单元格"，
 * 下一步入口却在视口之外。把可点击的那一份搬到输入区上方之后，它**不在滚动容器
 * 里**，日志再多也挤不掉它；上翻历史也不会被打断（不需要任何"再滚一次"）。
 *
 * 结构上刻意分三层：标题（结论 badge）/ 可滚动的正文 / 动作区。
 * 动作区是正文的**兄弟**而不是后代：正文超长时自己滚，按钮永远留在原地 ——
 * 与数据范围面板「动作条必须是滚动容器的兄弟」是同一条纪律。
 */
/**
 * 「这次已写入的内容还能不能撤销」的呈现状态。
 *
 * 与恢复动作**刻意分开**：核对/续跑回答"这次任务怎么收尾"，撤销回答"已经写进去的
 * 东西怎么退回来"。两者可以同时成立（来源变了不能续跑，但撤销照样可用），也可以
 * 同时不成立（材料不够 ⇒ 两个入口都不给）。
 */
export interface OfficeRecoveryUndoNotice {
  state: "available" | "blocked" | "done";
  detail: string;
  label?: string;
  onUndo?: () => void;
}

export function OfficeRecoveryPanel({
  card,
  busy,
  undo,
  primaryLabel,
  onAction,
  onReverify,
  primaryEntryVisible
}: {
  card: OfficeRecoveryCard;
  busy: boolean;
  undo?: OfficeRecoveryUndoNotice;
  /** 主操作此刻叫什么（宿主用同一份 facts 从共享判据算）。指路句里引用它，不另起名字。 */
  primaryLabel: string;
  onAction: (action: ExecutionRecoveryAction) => void;
  onReverify: () => void;
  /**
   * 输入区此刻**真的**摆着那颗主操作吗（恢复模式，或至少留着那行待办入口）。
   * `false` 时指路句改说"还没到"，不写一句用户点不到的话。
   */
  primaryEntryVisible: boolean;
}) {
  const actions = excelRecoveryActions(card);
  const secondary = actions.filter((action) => OFFICE_RECOVERY_SECONDARY_ACTIONS.includes(action));
  const primaryHint = officeRecoveryPrimaryHint(
    card.verdict.kind,
    primaryLabel,
    primaryEntryVisible
  );
  const acting = card.acting ?? null;
  return (
    <article
      className="office-recovery-panel"
      role="status"
      aria-label="上次 Office 写入的恢复核对"
      data-operation-id={card.operationId}
      data-recovery-kind={card.verdict.kind}
      data-recovery-acting={acting ?? "idle"}
    >
      <div className="office-resume-head">
        <strong>{EXECUTION_RECOVERY_TITLE}</strong>
        <span className={`office-resume-badge office-resume-${card.verdict.kind}`}>
          {executionRecoveryBadge(card.verdict)}
        </span>
      </div>
      <div className="office-recovery-panel-body">
        <p className="office-resume-detail">{executionRecoveryDetail(card.verdict)}</p>
        <p className="office-resume-detail">{executionRecoveryNote(card.verdict)}</p>
        <details className="office-resume-record">
          <summary>执行记录</summary>
          <p>{executionRecoveryProgress(card.verdict, card)}</p>
          <p>只支持已保存工作簿中的受控区域写入；核对与补记都不写单元格。</p>
        </details>
        {/*
          撤销行放在**正文里**（不动动作条）：它回答的是"已经写进去的怎么退回来"，
          与"这次任务怎么收尾"是两件事。可撤销时给一颗与结果卡同名的按钮；
          阻断时只给原因 —— 不画一个点下去只会报错的按钮；已撤销时只留结论。
        */}
        {undo && undo.state !== "available" && (
          <p
            className={undo.state === "done" ? "office-undo-result is-done" : "office-undo-result is-blocked"}
            role="status"
            data-recovery-undo={undo.state}
          >
            {undo.detail}
          </p>
        )}
        {undo?.state === "available" && (
          <div className="office-undo-row" data-recovery-undo="available">
            <button
              type="button"
              className="secondary-button"
              data-recovery-action="undo"
              disabled={busy || acting !== null}
              onClick={() => undo.onUndo?.()}
            >
              {busy ? "正在撤销…" : undo.label}
            </button>
            <span className="office-undo-hint">{undo.detail}</span>
          </div>
        )}
      </div>
      <div className="office-recovery-panel-actions">
        {/*
          主按钮（核对 / 继续执行）**不在这里**：它在输入框内部那块恢复操作区里，
          用同一份结论、同一个处理器。卡面只留一句"下一步去哪"，以及补记/重读这类
          次要动作 —— 否则同一个动作就有两个入口，而其中一颗不受输入框状态约束。
        */}
        {primaryHint && (
          <p className="office-recovery-primary-hint" data-recovery-primary-hint={card.verdict.kind}>
            {primaryHint}
          </p>
        )}
        {secondary.map((action) => (
          <button
            key={action}
            type="button"
            className="secondary-button"
            data-recovery-action={action}
            disabled={busy || acting !== null}
            onClick={() => onAction(action)}
          >
            {acting === action
              ? EXECUTION_RECOVERY_ACTIVE_LABEL[action]
              : EXECUTION_RECOVERY_ACTION_LABEL[action]}
          </button>
        ))}
        {actions.length === 0 && (
          /*
           * 没有可执行的写入动作时才给"重新核对（只读）"：用户改正授权、
           * 打开正确的工作簿之后，同一条判据重算结论，按钮自己回来。
           */
          <button
            type="button"
            className="secondary-button"
            data-recovery-action="reverify"
            disabled={busy}
            onClick={onReverify}
          >
            重新核对（只读）
          </button>
        )}
      </div>
    </article>
  );
}

/**
 * 历史消息里的**只读**执行记录：没有按钮、不重复结论 badge。
 *
 * 它是"这件事发生过"的证据，不是入口。可点击的那一份在输入框内部
 * （`RecoveryComposer`），所以 `selectExecutionRecoveryEntry` 挑中哪一条，
 * 哪一条才多一行"入口在输入框里"的指路 —— 其余历史记录不给这句，
 * 免得让人去找一个已经不在那儿的上方入口。
 *
 * **终态**（撤销确认 / 放弃 / 收尾）反过来：固定区域已经收走那张卡，这里改成
 * **可展开的终态摘要**（`<details>`）。折叠时只有一行"已结束 + 去哪看记录"，
 * 展开后是结论正文、进度与边界说明 —— 记录一条不少，但不再占着输入区的版面。
 * `active` 指路句对终态**不适用**（上方已经没有任何入口可指）。
 */
export function OfficeRecoveryRecord({
  card,
  active
}: {
  card: OfficeRecoveryCard;
  active: boolean;
}) {
  const terminal = executionRecoveryTerminalReason(card.verdict);
  if (terminal) {
    return (
      <RecoveryTerminalRecord className="office-recovery-record" kind={terminal} operationId={card.operationId}
        title={executionRecoveryBadge(card.verdict)} summary={executionRecoveryTerminalSummary(card.verdict)}>

        <p>{executionRecoveryDetail(card.verdict)}</p>
        <p>{executionRecoveryProgress(card.verdict, card)}</p>
        <p>{executionRecoveryNote(card.verdict)}</p>
      </RecoveryTerminalRecord>
    );
  }
  return (
    <section
      className="office-recovery-record"
      aria-label="上次 Office 写入的执行记录"
      data-operation-id={card.operationId}
    >
      <strong>执行记录</strong>
      <p>{executionRecoveryProgress(card.verdict, card)}</p>
      {active && <p>{EXECUTION_RECOVERY_ENTRY_HINT}</p>}
    </section>
  );
}
