import { useId, useLayoutEffect, useRef, useState } from "react";
import type {
  TaskProgressPhase,
  TaskProgressStatus,
  TaskProgressTarget
} from "../contracts";
import type {
  ActivityLog,
  ActivityOutcome,
  ActivityStep
} from "../hooks/useActivityProgress";

interface TaskProgressTimelineProps {
  log: ActivityLog;
  live?: boolean;
  onStop?: () => void;
}

const LONG_CONTENT_HEIGHT = 480;

const PHASE_LABELS: Record<TaskProgressPhase, string> = {
  intent: "理解",
  planning: "规划",
  preflight: "检查",
  backup: "备份",
  executing: "执行",
  saving: "保存",
  verifying: "验证",
  refreshing: "刷新"
};

const STATUS_LABELS: Record<TaskProgressStatus, string> = {
  started: "正在处理",
  progress: "进行中",
  completed: "已完成",
  warning: "需要注意",
  failed: "执行失败",
  cancelled: "已取消"
};

const OUTCOME_LABELS: Record<ActivityOutcome, string> = {
  completed: "已完成",
  awaiting_input: "等待确认",
  warning: "需要注意",
  failed: "执行失败",
  cancelled: "已取消",
  interrupted: "连接中断"
};

const STATUS_ICONS: Record<TaskProgressStatus, string> = {
  started: "●",
  progress: "●",
  completed: "✓",
  warning: "!",
  failed: "×",
  cancelled: "■"
};

function formatElapsed(ms: number): string {
  if (ms < 1000) return Math.max(0, Math.round(ms)) + "ms";
  return (ms / 1000).toFixed(1) + "s";
}

function unitLabel(step: ActivityStep): string {
  if (step.unit === "rows") return "行";
  if (step.unit === "fields") return "个字段";
  if (step.unit === "actions") return "步";
  if (step.unit === "sheets") return "张表";
  if (step.unit === "files") return "个文件";
  if (step.unit === "checks") return "项检查";
  return "";
}

function formatTarget(target?: TaskProgressTarget | null): string {
  if (!target) return "";
  const workbook = target.file?.trim() ?? "";
  const sheet = target.sheet?.trim() ?? "";
  const range = target.range?.trim() ?? "";
  let location = workbook;
  if (sheet) location += (location ? " › " : "") + sheet;
  if (range) location += (location ? "!" : "") + range;
  return location;
}

function stepStatus(step: ActivityStep): TaskProgressStatus {
  return step.status ?? "completed";
}

function hasStepExplanation(step: ActivityStep): boolean {
  return Boolean(
    step.note ||
      step.detail ||
      step.toolName ||
      step.actionType ||
      formatTarget(step.target) ||
      (step.facts && step.facts.length > 0) ||
      step.trace
  );
}

export function TaskProgressTimeline({
  log,
  live = false,
  onStop
}: TaskProgressTimelineProps) {
  const [expanded, setExpanded] = useState(true);
  const [hasLongContent, setHasLongContent] = useState(false);
  const contentId = useId();
  const sectionRef = useRef<HTMLElement>(null);
  const contentBodyRef = useRef<HTMLDivElement>(null);
  const headerToggleRef = useRef<HTMLButtonElement>(null);
  const visibleTraceIds = new Set(
    log.steps
      .map((step) => step.trace?.id)
      .filter((id): id is string => Boolean(id))
  );
  const unmatchedTraces = (log.traces ?? []).filter(
    (trace) => !visibleTraceIds.has(trace.id)
  );
  const lastStep = log.steps.length > 0 ? log.steps[log.steps.length - 1] : undefined;
  const overallStatus: ActivityOutcome | TaskProgressStatus =
    log.outcome ?? (lastStep ? stepStatus(lastStep) : "progress");
  const overallLabel =
    log.outcome != null
      ? OUTCOME_LABELS[log.outcome]
      : STATUS_LABELS[overallStatus as TaskProgressStatus];
  const processLabel =
    log.outcome === "awaiting_input" ? "需求分析过程" : "执行过程";

  useLayoutEffect(() => {
    if (!expanded) {
      setHasLongContent(false);
      return;
    }
    const body = contentBodyRef.current;
    if (!body) return;
    const measure = () => {
      setHasLongContent(body.scrollHeight > LONG_CONTENT_HEIGHT);
    };
    measure();
    if (typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver(measure);
    observer.observe(body);
    return () => observer.disconnect();
  }, [expanded, log.steps.length, unmatchedTraces.length]);

  const collapseFromBottom = () => {
    setExpanded(false);
    const returnToHeader = () => {
      sectionRef.current?.scrollIntoView?.({ block: "nearest" });
      headerToggleRef.current?.focus({ preventScroll: true });
    };
    if (typeof window.requestAnimationFrame === "function") {
      window.requestAnimationFrame(returnToHeader);
    } else {
      window.setTimeout(returnToHeader, 0);
    }
  };

  return (
    <section
      ref={sectionRef}
      className={
        "task-progress-timeline " +
        (live ? "activity-card is-live" : "activity-log-card")
      }
      data-status={overallStatus}
      aria-live={live ? "polite" : undefined}
    >
      <div className="task-progress-header">
        <div className="task-progress-heading">
          {live && (
            <span className="activity-pulse" aria-hidden="true">
              <i />
              <i />
              <i />
            </span>
          )}
          <strong>
            {processLabel} · {log.steps.length} 步 · {formatElapsed(log.totalMs)}
          </strong>
          <span className="task-progress-overall">
            {overallLabel}
          </span>
        </div>
        <div className="task-progress-actions">
          {onStop && (
            <button
              type="button"
              className="activity-stop-button"
              onClick={onStop}
            >
              停止
            </button>
          )}
          <button
            ref={headerToggleRef}
            type="button"
            className="task-progress-toggle"
            aria-expanded={expanded}
            aria-controls={contentId}
            onClick={() => setExpanded((current) => !current)}
          >
            {expanded ? "收起" : "展开"}
          </button>
        </div>
      </div>

      {expanded && (
        <div id={contentId} className="task-progress-content">
          <div ref={contentBodyRef} className="task-progress-content-body">
          {log.connectionState &&
            log.connectionState !== "connected" &&
            log.connectionMessage && (
              <div
                className="task-progress-connection"
                data-state={log.connectionState}
              >
                {log.connectionMessage}
              </div>
            )}
          <ol className="activity-steps">
            {log.steps.map((step, index) => {
              const status = stepStatus(step);
              const target = formatTarget(step.target);
              const fallback = hasStepExplanation(step)
                ? null
                : STATUS_LABELS[status];
              return (
                <li
                  key={step.id ?? step.label + "-" + index}
                  data-status={status}
                >
                  <span
                    className="activity-step-check"
                    data-status={status}
                    aria-label={STATUS_LABELS[status]}
                  >
                    {STATUS_ICONS[status]}
                  </span>
                  <div className="activity-step-body">
                    <div className="activity-step-title-row">
                      <span className="activity-step-index">{index + 1}</span>
                      <strong className="activity-step-label">
                        {step.label}
                      </strong>
                      <span className="activity-step-time">
                        {formatElapsed(step.elapsedMs)}
                      </span>
                    </div>
                    <div className="activity-step-badges">
                      {step.phase && (
                        <span className="activity-step-badge is-phase">
                          {PHASE_LABELS[step.phase]}
                        </span>
                      )}
                      {step.toolName && (
                        <span className="activity-step-badge is-technical">
                          工具 <code>{step.toolName}</code>
                        </span>
                      )}
                      {step.actionType && (
                        <span className="activity-step-badge is-technical">
                          动作 <code>{step.actionType}</code>
                        </span>
                      )}
                    </div>
                    {step.note && (
                      <p className="activity-step-note">{step.note}</p>
                    )}
                    {target && (
                      <div className="activity-step-fact">
                        <span>目标</span>
                        <code title={target}>{target}</code>
                      </div>
                    )}
                    {step.facts?.map((fact, factIndex) => (
                      <div
                        className="activity-step-fact"
                        key={fact.label + "-" + factIndex}
                      >
                        <span>{fact.label}</span>
                        {fact.style === "code" ? (
                          <code>{fact.value}</code>
                        ) : (
                          <b>{fact.value}</b>
                        )}
                      </div>
                    ))}
                    {step.detail && (
                      <p className="activity-step-detail-text">{step.detail}</p>
                    )}
                    {step.trace && (
                      <div
                        className="activity-step-trace"
                        data-kind={step.trace.kind}
                      >
                        <strong>判断：{step.trace.summary}</strong>
                        {step.trace.basis && (
                          <span>依据：{step.trace.basis}</span>
                        )}
                        {step.trace.changedFrom && (
                          <span>纠偏前：{step.trace.changedFrom}</span>
                        )}
                      </div>
                    )}
                    {fallback && (
                      <p className="activity-step-note">{fallback}</p>
                    )}
                    {step.total != null && (
                      <div
                        className="activity-measure"
                        role="progressbar"
                        aria-valuemin={0}
                        aria-valuemax={step.total}
                        aria-valuenow={step.current ?? 0}
                      >
                        <span>
                          {step.current ?? 0} / {step.total} {unitLabel(step)}
                        </span>
                        <i
                          style={{
                            width:
                              Math.min(
                                100,
                                step.total > 0
                                  ? ((step.current ?? 0) / step.total) * 100
                                  : 0
                              ) + "%"
                          }}
                        />
                      </div>
                    )}
                  </div>
                </li>
              );
            })}
          </ol>
          {unmatchedTraces.length > 0 && (
            <div className="activity-traces persisted">
              <strong>判断依据与纠偏记录</strong>
              <ul>
                {unmatchedTraces.map((trace) => (
                  <li key={trace.id} data-kind={trace.kind}>
                    <strong>{trace.summary}</strong>
                    {trace.basis && <span>依据：{trace.basis}</span>}
                    {trace.changedFrom && (
                      <span>纠偏前：{trace.changedFrom}</span>
                    )}
                  </li>
                ))}
              </ul>
            </div>
          )}
          </div>
          {hasLongContent && (
            <div className="task-progress-bottom-actions">
              <button
                type="button"
                className="task-progress-toggle task-progress-bottom-toggle"
                aria-expanded="true"
                aria-controls={contentId}
                onClick={collapseFromBottom}
              >
                ↑ 收起执行过程
              </button>
            </div>
          )}
        </div>
      )}
    </section>
  );
}

export default TaskProgressTimeline;
