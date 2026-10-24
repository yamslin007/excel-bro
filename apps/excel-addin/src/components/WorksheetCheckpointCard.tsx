import { useEffect, useRef, useState } from "react";
import type { WorksheetCheckpointSummary } from "../contracts";
import type { WorksheetCheckpointEntry, useWorksheetCheckpoint } from "../hooks/useWorksheetCheckpoint";
import {
  WORKSHEET_CHECKPOINT_BACKUP_FALLBACK_HINT,
  WORKSHEET_CHECKPOINT_RECOVERY_HINT,
  WORKSHEET_CHECKPOINT_RECOVERY_LABEL,
  WORKSHEET_CHECKPOINT_UNDO_LABEL,
  describeWorksheetCheckpoint,
  worksheetCheckpointStatusLabel
} from "@excel-bro/core/worksheetCheckpoint";
import { ExecutionResultCard, ExecutionResultDetails } from "./ExecutionResultCard";
import "./worksheetCheckpointCard.css";

type Controller = ReturnType<typeof useWorksheetCheckpoint>;

/**
 * 一条请求自己的删除结果卡。
 *
 * 它只画在**发起这次删除的那条消息**下面：没有"全局最新一条"的语义，也不会
 * 被追加到聊天气泡流末尾。详细清单与诊断默认折叠，首屏只留一句结论。
 *
 * 右下角的入口与普通写入结果卡同名（「↶ 撤销本次执行」），它现在**真的**做原地撤销：
 * 把本次删掉的工作表放回当前打开的原工作簿，文件身份不变、后续编辑保留。
 * 只有"本次删除完整完成 + 有可信归档 + 还没有撤销成功"时按钮才出现；
 * 否则只给原因，并指路「备份与恢复」里那个"打开执行前副本"的备用入口。
 */
export function WorksheetCheckpointCard({
  entry,
  controller,
  onRestored
}: {
  entry: WorksheetCheckpointEntry;
  controller: Controller;
  /**
   * 原地撤销**核验通过后**回调：宿主据此刷新目录、把授权与目标基线对齐。
   *
   * 只在 `restore()` 返回 ok 时触发；失败/冲突不碰宿主的目录与授权。
   */
  onRestored?: (result: import("../contracts").WorksheetRestoreResult, entry: WorksheetCheckpointEntry) => void;
}) {
  const view = describeWorksheetCheckpoint({
    receipt: entry.result,
    error: entry.error,
    deferred: entry.deferred,
    opened: entry.opened
  });
  const restoring = entry.restoring === true;
  const hasDetails = view.deletedSheets.length > 0 || view.remainingSheets.length > 0 ||
    view.restoredSheets.length > 0 || view.diagnostic;
  return (
    <ExecutionResultCard
      tone={view.tone}
      title={view.title}
      summary={view.summary}
      className="worksheet-checkpoint-card"
    >
      {hasDetails && (
        <ExecutionResultDetails>
          {view.deletedSheets.length > 0 && <p>已删除：{view.deletedSheets.join("、")}</p>}
          {view.remainingSheets.length > 0 && <p>剩余工作表：{view.remainingSheets.join("、")}</p>}
          {view.restoredSheets.length > 0 && <p>已放回：{view.restoredSheets.join("、")}</p>}
          {view.diagnostic && <p>{view.diagnostic}</p>}
        </ExecutionResultDetails>
      )}
      {view.consentSheets && (
        <>
          <p className="worksheet-checkpoint-consent">
            仅加入本地备份：{view.consentSheets.join("、")}
          </p>
          {controller.canResume(entry.operationId) ? (
            <div className="worksheet-checkpoint-actions">
              <button
                className="primary-button"
                disabled={controller.pending}
                onClick={() => void controller.confirmBackup(entry.operationId)}
              >
                允许仅本地备份并继续
              </button>
              <button
                className="secondary-button"
                disabled={controller.pending}
                onClick={() => controller.defer(entry.operationId)}
              >
                暂不删除
              </button>
            </div>
          ) : (
            <p>本次授权现场已结束，请在原工作簿重新发起任务；不会自动继续删除。</p>
          )}
        </>
      )}
      {/* 没有回执时，这句原文已经是首屏结论；有回执时它才作为独立告警出现，不重复。 */}
      {entry.error && entry.result && <p role="alert">{entry.error}</p>}
      {view.openedNotice && <p role="status">{view.openedNotice}</p>}
      {view.restoredNotice && <p role="status">{view.restoredNotice}</p>}
      {entry.restoreError && <p role="alert">{entry.restoreError}</p>}
      {view.undoAvailable && (
        <div className="execution-recovery-footer">
          {/*
            与普通写入结果卡同一个入口文案与排布：按钮在右下，说明在左侧。
            这个按钮**真的**在原工作簿里撤销本次删除，所以说明句写的是"保留后续编辑"，
            不再说"打开副本"。副本入口在「备份与恢复」，两者不能互相冒充。
          */}
          <button
            type="button"
            className="secondary-button"
            data-worksheet-checkpoint-undo={entry.operationId}
            disabled={controller.pending || restoring}
            onClick={async () => {
              const outcome = await controller.restore(entry.operationId);
              if (outcome.ok && outcome.result && onRestored) onRestored(outcome.result, entry);
            }}
          >
            {restoring ? "正在撤销本次删除…" : WORKSHEET_CHECKPOINT_UNDO_LABEL}
          </button>
          <span className="undo-row-hint">{view.undoHint}</span>
        </div>
      )}
      {!view.undoAvailable && view.undoReason && (
        <p className="worksheet-checkpoint-undo-blocked" role="status">
          {view.undoReason}
          {view.recoveryAvailable ? ` ${WORKSHEET_CHECKPOINT_BACKUP_FALLBACK_HINT}。` : ""}
        </p>
      )}
      {view.canQueryStatus && (
        <button
          type="button"
          className="secondary-button"
          disabled={controller.pending}
          onClick={() => void controller.refresh(entry.operationId)}
        >
          查询执行状态
        </button>
      )}
    </ExecutionResultCard>
  );
}

function formatStamp(value: string | null): string {
  if (!value) return "时间未知";
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? value : parsed.toLocaleString("zh-CN");
}

/**
 * 「备份与恢复」里的工作表删除恢复点。
 *
 * 列表来自服务端按**工作簿身份**的查询：旧版只存 operationId 的记录也从这里
 * 找回，而不是被猜进当前会话。这里只打开副本，不覆盖原表。
 */
export function WorksheetCheckpointRecoverySection({
  controller,
  workbookPath
}: {
  controller: Controller;
  workbookPath: string | null;
}) {
  const [items, setItems] = useState<WorksheetCheckpointSummary[]>([]);
  const [state, setState] = useState<"idle" | "loading" | "ready" | "failed">("idle");
  const [error, setError] = useState<string | null>(null);
  const [busyId, setBusyId] = useState<string | null>(null);
  const [notice, setNotice] = useState<{ operationId: string; text: string; failed: boolean } | null>(null);
  // controller 每次渲染都是新对象；只取稳定语义，避免因依赖它而反复重新查询。
  const listRecoverable = useRef(controller.listRecoverable);
  listRecoverable.current = controller.listRecoverable;
  const openCheckpoint = useRef(controller.open);
  openCheckpoint.current = controller.open;

  useEffect(() => {
    if (!workbookPath) {
      setState("idle");
      setItems([]);
      return;
    }
    let alive = true;
    setState("loading");
    setError(null);
    void listRecoverable.current(workbookPath).then(
      checkpoints => {
        if (!alive) return;
        setItems(checkpoints);
        setState("ready");
      },
      reason => {
        if (!alive) return;
        setError(reason instanceof Error ? reason.message : "恢复点列表读取失败。");
        setState("failed");
      }
    );
    return () => { alive = false; };
  }, [workbookPath]);

  return (
    <section className="worksheet-checkpoint-recovery" aria-label="工作表删除恢复点">
      <header>
        <strong>工作表删除恢复点</strong>
        <span>按当前工作簿显示；打开的是执行前副本，原表及后续编辑保留。</span>
      </header>
      {state === "idle" && (
        <p className="worksheet-checkpoint-recovery-state">
          当前工作簿身份不可用（尚未保存），暂时无法列出恢复点。
        </p>
      )}
      {state === "loading" && <p className="worksheet-checkpoint-recovery-state">正在读取恢复点…</p>}
      {state === "failed" && <p className="folder-backup-feedback error" role="alert">{error}</p>}
      {state === "ready" && items.length === 0 && (
        <p className="worksheet-checkpoint-recovery-state">这个工作簿还没有删除恢复点。</p>
      )}
      {items.length > 0 && (
        <ul className="worksheet-checkpoint-recovery-list">
          {items.map(item => {
            const current = notice?.operationId === item.operationId ? notice : null;
            return (
              <li key={item.operationId}>
                <div className="worksheet-checkpoint-recovery-head">
                  <strong>{item.planTitle ?? "工作表删除"}</strong>
                  <em className={`status-${item.status}`}>
                    {worksheetCheckpointStatusLabel(item.status, item.restoreStatus ?? null)}
                  </em>
                </div>
                <small>{formatStamp(item.createdAt)}</small>
                <details>
                  <summary>查看删除清单与诊断</summary>
                  {item.deletedSheets.length > 0 && <p>已删除：{item.deletedSheets.join("、")}</p>}
                  {item.remainingSheets.length > 0 && <p>剩余工作表：{item.remainingSheets.join("、")}</p>}
                  <p>{item.message}</p>
                </details>
                {item.recoveryAvailable ? (
                  <button
                    type="button"
                    className="secondary-button"
                    disabled={busyId !== null}
                    onClick={async () => {
                      setBusyId(item.operationId);
                      const outcome = await openCheckpoint.current(item.operationId);
                      setBusyId(null);
                      setNotice({
                        operationId: item.operationId,
                        text: outcome.ok ? "已打开恢复副本。" : outcome.error ?? "恢复副本未打开；当前原表保留。",
                        failed: !outcome.ok
                      });
                    }}
                    title={WORKSHEET_CHECKPOINT_RECOVERY_HINT}
                  >
                    {busyId === item.operationId ? "正在打开恢复副本…" : WORKSHEET_CHECKPOINT_RECOVERY_LABEL}
                  </button>
                ) : (
                  <small>这份记录没有可用副本。</small>
                )}
                {current && (
                  <p role={current.failed ? "alert" : "status"}>{current.text}</p>
                )}
              </li>
            );
          })}
        </ul>
      )}
    </section>
  );
}
