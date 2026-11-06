import { timelineDurations, isTerminalStatus } from "@excel-bro/core/activityTimeline";
import {
  foldedProcessLine,
  isCompactMessageProcess,
  processPanelBodyMaxHeight,
  processPanelExpanded,
  processPanelHasOverflow,
  type ProcessPanelUserChoice
} from "@excel-bro/core/processPanel";
import { latestPythonCode } from "@excel-bro/core/pythonCodeArtifact";
import {
  followingAfterScroll,
  isNearScrollBottom,
  scrollJumpLabel,
  shouldFollowAfterLayout
} from "@excel-bro/core/scrollFollow";
import { useEffect, useId, useLayoutEffect, useRef, useState } from "react";
import { PrePlanningPauseCard } from "./PrePlanningPauseCard";
import PythonCodeBlock from "./PythonCodeBlock";
import { PYTHON_CODE_LIMITS } from "../contracts";
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

/**
 * 这一段回复在对话里的位置信息。
 *
 * 它**只是展示**：续跑段不重排编号，历史段提供"查看后续进展"的定位入口。
 * 任务身份、步骤身份、事件序号与写入回执都不经过这里。
 */
export interface ActivitySegmentView {
  /** 这一段是接着原任务继续的（不显示"步骤 1"这种重排编号）。 */
  continuesTask?: boolean;
  /** 这一段已被后续回复接续：给出轻量定位入口。 */
  onShowLater?: () => void;
  /** 定位入口的文案，例如「查看后续进展」。 */
  laterLabel?: string;
  /** 这一段是当前活动的展示位置。 */
  live?: boolean;
}

interface TaskProgressTimelineProps {
  log: ActivityLog;
  live?: boolean;
  onStop?: () => void;
  onContinue?: () => void;
  continueBusy?: boolean;
  continuationError?: string;
  segment?: ActivitySegmentView;
  /**
   * 这一段正在等用户回答澄清问题。
   *
   * 为什么单独一个入参而不是从 `outcome` 推：`awaiting_input` 既可能是"等你回答"，
   * 也可能是"预览已生成，等你确认写入" —— 两者该说的是不同的话，而过程卡本身
   * 看不到澄清状态。由面板照实传，卡只说事实。
   */
  awaitingAnswer?: boolean;
}

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
  awaiting_input: "等待下一步操作",
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

/**
 * 最近的可滚动祖先。
 *
 * 用来在**展开/收起**时补偿外层滚动位置：卡片高度变了，卡片下面的内容会整体
 * 上下跳，用户正在读的那一行就跑了。这里只找容器、不自己造滚动。
 * jsdom 里 `overflowY` 通常拿不到 `auto`，因此测试环境自然走不到补偿分支 ——
 * 那是刻意的：结构测试不该依赖真实布局数值。
 */
function scrollableAncestor(from: HTMLElement | null): HTMLElement | null {
  let node = from?.parentElement ?? null;
  while (node) {
    const overflowY = window.getComputedStyle?.(node).overflowY;
    if (overflowY === "auto" || overflowY === "scroll") return node;
    node = node.parentElement;
  }
  return null;
}

export function TaskProgressTimeline({
  log,
  live = false,
  onStop,
  onContinue,
  continueBusy = false,
  continuationError,
  segment,
  awaitingAnswer = false
}: TaskProgressTimelineProps) {
  const [clock, setClock] = useState(Date.now);
  useEffect(() => {
    if (!live) return;
    setClock(Date.now());
    const timer = window.setInterval(() => setClock(Date.now()), 1000);
    return () => window.clearInterval(timer);
  }, [live]);
  const origin = log.timelineStartedAt ?? clock - log.totalMs;
  const durations = timelineDurations(log, origin,
    log.timelineStartedAt !== undefined && live ? clock : origin + log.totalMs);
  /*
   * 展开状态只由三样东西决定：是否在跑、活动级结论、**用户自己的选择**。
   * 判据在共享层（`processPanelExpanded`），两端读同一份。
   *
   * 历史实现是 `setExpanded(!collapseCompletedChat)` 配一个 `useLayoutEffect`：
   * 每次 `defaultCollapsed` 变化都会把用户的展开选择覆盖掉。所以这里改成记
   * "用户表过什么态"，表过就不再让任何自动收起改回去。
   */
  const [userChoice, setUserChoice] = useState<ProcessPanelUserChoice>(null);
  const expanded = processPanelExpanded({
    live,
    outcome: log.outcome ?? null,
    userChoice,
    awaitingAnswer
  });
  /*
   * 「多显示一些」与「展开/收起」是**两个**意愿，不能共用同一个布尔：
   *  - 展开/收起回答"要不要看详细动作"；
   *  - 放大回答"这个框能占多高"。
   * 进行中默认就是"显示细节但只给 160px"，用户想多看几行时才放宽。
   * 混用会让进行中一上来就占 420px，把输入框顶下去。
   */
  const [enlarged, setEnlarged] = useState(false);
  const bodyCap = processPanelBodyMaxHeight({ live, enlarged });
  const [hasLongContent, setHasLongContent] = useState(false);
  const contentId = useId();
  const sectionRef = useRef<HTMLElement>(null);
  const contentBodyRef = useRef<HTMLDivElement>(null);
  const headerToggleRef = useRef<HTMLButtonElement>(null);
  /*
   * 过程框**内部**的跟随意图：与消息流读同一份判据（`scrollFollow`）。
   * 两个容器各写一套的结果是：消息流贴底跟随、过程框里的最新步骤却在外面。
   */
  const innerFollowingRef = useRef(true);
  const [awayFromInnerBottom, setAwayFromInnerBottom] = useState(false);
  const pendingAnchorRef = useRef<{ bottom: number; scroller: HTMLElement | null } | null>(null);
  const visibleTraceIds = new Set(
    log.steps
      .map((step) => step.trace?.id)
      .filter((id): id is string => Boolean(id))
  );
  const unmatchedTraces = (log.traces ?? []).filter(
    (trace) => !visibleTraceIds.has(trace.id)
  );
  const lastStep = log.steps.length > 0 ? log.steps[log.steps.length - 1] : undefined;
  const activeStep = [...log.steps].reverse().find(step => !isTerminalStatus(step.status));
  const overallStatus: ActivityOutcome | TaskProgressStatus =
    log.outcome ?? (live ? "progress" : lastStep ? stepStatus(lastStep) : "progress");
  const overallLabel =
    awaitingAnswer ? "等待你的回答" :
    !live && segment?.onShowLater ? "本段已结束 · 已接续" :
    !live && log.outcome === "awaiting_input" && log.steps.some(step =>
      step.id === "preflight-plan" && step.status === "completed") ? "预览已生成 · 待确认写入" :
    continueBusy && onContinue ? "正在继续" : onContinue ? "已暂停" : log.outcome != null
      ? OUTCOME_LABELS[log.outcome]
      : STATUS_LABELS[overallStatus as TaskProgressStatus];
  const processLabel =
    segment?.continuesTask ? "接续过程" : log.outcome === "awaiting_input"
      ? "需求分析过程"
      : segment?.continuesTask ? "接续过程" : "执行过程";
  /**
   * 这一段展示的 Python 代码工件。
   *
   * 从步骤里取（最后一条带工件的步骤胜出），不是从文案里猜：代码块要回答的是
   * "这次真的跑了哪段代码"，只能来自事件携带的工件身份与版本。
   */
  const pythonCode = latestPythonCode(log.steps);
  /**
   * 收起的详细动作区只留一行：**阶段结论 + 耗时**（判据在共享层）。
   * 展开时的标题仍然是技术信息（步数 + 前台耗时），它属于"过程详情"。
   */
  const headingText = expanded
    ? `${processLabel} · ${log.steps.length} 步 · ${formatElapsed(durations.activeMs)}`
    : foldedProcessLine({ conclusion: live && !onContinue && activeStep
      ? `${activeStep.label}…` : overallLabel, activeMs: durations.activeMs });

  useLayoutEffect(() => {
    if (!expanded) {
      setHasLongContent(false);
      return;
    }
    const body = contentBodyRef.current;
    if (!body) return;
    const measure = () => {
      setHasLongContent(processPanelHasOverflow({ live, enlarged, contentHeight: body.scrollHeight }));
    };
    measure();
    if (typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver(measure);
    observer.observe(body);
    return () => observer.disconnect();
  }, [expanded, live, enlarged, log.steps.length, unmatchedTraces.length]);

  /*
   * 新步骤进来：只在用户**没有上翻**时把内部滚动带到最新一步。
   * 依赖里刻意没有 `clock` —— 每秒重绘不该动滚动位置。
   */
  useLayoutEffect(() => {
    const body = contentBodyRef.current;
    if (!body || !expanded) return;
    if (shouldFollowAfterLayout({ following: innerFollowingRef.current, metrics: body })) {
      body.scrollTop = body.scrollHeight;
      setAwayFromInnerBottom(false);
    }
  }, [expanded, log.steps.length]);

  /*
   * 展开/收起之后把**卡片下面**的内容钉回原位。
   *
   * 为什么不用 `scrollIntoView`：那个调用会替用户决定"现在该看哪"，卡片一收起
   * 就把正在读的位置顶走 —— 用户明确要求过"结束收起不打断阅读"。这里只补偿
   * 高度差，不动阅读焦点；滚动位置如果本来就在底部（跟随中），交给跟随逻辑处理。
   */
  useLayoutEffect(() => {
    const pending = pendingAnchorRef.current;
    pendingAnchorRef.current = null;
    const section = sectionRef.current;
    const scroller = pending?.scroller ?? null;
    if (!pending || !section || !scroller) return;
    if (isNearScrollBottom(scroller)) return;
    const delta = section.getBoundingClientRect().bottom - pending.bottom;
    if (Math.abs(delta) < 1) return;
    scroller.scrollTop += delta;
  }, [expanded]);

  const applyExpanded = (next: boolean) => {
    const section = sectionRef.current;
    pendingAnchorRef.current = section
      ? { bottom: section.getBoundingClientRect().bottom, scroller: scrollableAncestor(section) }
      : null;
    setUserChoice(next ? "expanded" : "collapsed");
  };

  const collapseFromBottom = () => {
    applyExpanded(false);
    const returnToHeader = () => {
      // 只把焦点送回标题按钮，**不**滚动页面（`preventScroll`）。
      headerToggleRef.current?.focus({ preventScroll: true });
    };
    if (typeof window.requestAnimationFrame === "function") {
      window.requestAnimationFrame(returnToHeader);
    } else {
      window.setTimeout(returnToHeader, 0);
    }
  };

  const handleInnerScroll = () => {
    const body = contentBodyRef.current;
    if (!body) return;
    const following = followingAfterScroll(body);
    innerFollowingRef.current = following;
    setAwayFromInnerBottom(!following);
  };

  const scrollInnerToLatest = () => {
    const body = contentBodyRef.current;
    if (!body) return;
    innerFollowingRef.current = true;
    body.scrollTop = body.scrollHeight;
    setAwayFromInnerBottom(false);
  };

  const innerJump = scrollJumpLabel({ running: live, scope: "process" });

  return (
    <section
      ref={sectionRef}
      className={
        "task-progress-timeline " +
        (live ? "activity-card is-live" : "activity-log-card")
      }
      data-status={overallStatus}
      data-expanded={expanded ? "true" : "false"}
      data-folded={expanded ? "false" : "true"}
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
          <strong title="本段实际用时；阶段与工具可能重叠，步骤耗时不能相加。">
            {headingText}
          </strong>
          {(durations.pausedMs > 0 || durations.waitingMs > 0) && (
            <small className="task-progress-duration-breakdown" title="本段停止时的计时快照；等待进度仅表示未接收进度的时间，不证明后台正在运行或已暂停。">
              暂停 {formatElapsed(durations.pausedMs)} · 等待进度 {formatElapsed(durations.waitingMs)}
            </small>
          )}
          {/*
            收起之后**视觉上**去掉结论 badge：结论已经在标题那一行里，
            同一行写两遍会让人以为是两件事。
            但节点保留（`data-folded` 下由 CSS 隐藏），因为它是"这一段处于什么状态"
            的唯一稳定标记 —— 无障碍查询与既有断言都读它，删节点等于把状态标记一起删掉。
          */}
          <span className="task-progress-overall">
            {overallLabel}
          </span>
        </div>
        <div className="task-progress-actions">
          {segment?.onShowLater && (
            <button
              type="button"
              className="task-progress-jump"
              onClick={segment.onShowLater}
            >
              {segment.laterLabel ?? "查看后续进展"} →
            </button>
          )}
          {/*
            过程框内部的"回到最新步骤"。只在用户自己翻上去之后出现 ——
            贴底时它是噪音，而贴底正是绝大多数时候。
          */}
          {expanded && awayFromInnerBottom && (
            <button
              type="button"
              className="task-progress-inline-action"
              data-process-jump="true"
              onClick={scrollInnerToLatest}
              aria-label={innerJump.ariaLabel}
              title={innerJump.title}
            >
              {innerJump.text}
            </button>
          )}
          {/*
            「放大」只在内容真的超出限高时出现：一个永远在、点了没变化的按钮
            等于假的。它改的是**框能占多高**，不改变"要不要看详细动作"。
          */}
          {expanded && hasLongContent && (
            <button
              type="button"
              className="task-progress-inline-action"
              data-process-enlarge={enlarged ? "on" : "off"}
              aria-pressed={enlarged}
              onClick={() => setEnlarged((current) => !current)}
              title={enlarged ? "回到默认高度，框内仍可滚动" : "把过程框放高一些，少滚几下"}
            >
              {enlarged ? "还原高度" : "多显示几行"}
            </button>
          )}
          {onContinue && <PrePlanningPauseCard busy={continueBusy} error={continuationError} onContinue={onContinue} />}
          {onStop && (
            <button
              type="button"
              className="activity-stop-button"
              onClick={onStop}
            >
                暂停
            </button>
          )}
          <button
            ref={headerToggleRef}
            type="button"
            className="task-progress-toggle"
            aria-expanded={expanded}
            aria-controls={contentId}
            aria-label={expanded ? "收起" : "查看过程"}
            onClick={() => applyExpanded(!expanded)}
          >
            {expanded ? "收起" : "查看过程"}
          </button>
        </div>
      </div>

      {expanded && (
        <div id={contentId} className="task-progress-content">
          <div
            ref={contentBodyRef}
            className="task-progress-content-body"
            onScroll={handleInnerScroll}
            tabIndex={0}
            role="group"
            aria-label={`${processLabel}详细步骤`}
            style={{
              // 业务上限交给 CSS 变量，窗格适配（vh）留在样式表里。
              ["--process-panel-max-height" as string]: `${bodyCap}px`
            }}
          >
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
              if (!live && !awaitingAnswer && isCompactMessageProcess(log)) {
                return (
                  <li key={step.id} data-status={status}>
                    <span className="activity-step-check" aria-label={STATUS_LABELS[status]}>{STATUS_ICONS[status]}</span>
                    <div className="activity-step-body">
                      <div className="activity-step-title-row">
                        <strong className="activity-step-label">{step.note || step.label}</strong>
                        <span className="activity-step-time">{formatElapsed(step.elapsedMs)}</span>
                      </div>
                    </div>
                  </li>
                );
              }
              const segmentStopped = status === "warning" && !!step.id &&
                log.closedByFinish?.includes(step.id) === true;
              const target = formatTarget(step.target);
              /*
               * "当前步骤" = 最后一个未结束的步骤。它需要被突出，而已完成的行
               * 要弱化 —— 否则一屏 12 行里到底哪一action 在跑得靠读时间猜。
               * 判据只用状态：没收到结束事件的那一行就是当前行。
               */
              const isCurrent = !isTerminalStatus(step.status) &&
                log.steps.slice(index + 1).every((later) => isTerminalStatus(later.status));
              const fallback = hasStepExplanation(step)
                ? null
                : segmentStopped ? "本段已停止接收进度，完成情况以后续进展为准。" : STATUS_LABELS[status];
              return (
                <li
                  key={step.id ?? step.label + "-" + index}
                  data-status={segmentStopped ? "interrupted" : status}
                  data-current={isCurrent ? "true" : undefined}
                >
                  <span
                    className="activity-step-check"
                    data-status={segmentStopped ? "interrupted" : status}
                    aria-label={segmentStopped ? "本段已中断" : STATUS_LABELS[status]}
                  >
                    {segmentStopped ? "Ⅱ" : STATUS_ICONS[status]}
                  </span>
                  <div className="activity-step-body">
                    <div className="activity-step-title-row">
                      {!segment?.continuesTask && (
                        <span className="activity-step-index">{index + 1}</span>
                      )}
                      <strong className="activity-step-label">
                        {step.label}
                        {step.id === "planning.model" && "（整体阶段，包含工具调用）"}
                      </strong>
                      <span className="activity-step-time">
                        {formatElapsed(live && !log.excludedKind && !isTerminalStatus(step.status) && step.startedAt !== undefined
                          ? Math.max(0, clock - step.startedAt) : step.elapsedMs)}
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
          {pythonCode && (
            <PythonCodeBlock
              artifact={pythonCode}
              taskId={log.operationId ?? null}
              limits={PYTHON_CODE_LIMITS}
            />
          )}
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
          {/*
            底部收起入口是滚动容器的**兄弟**（不在 `-body` 里）：内容再长、
            内部滚到哪，它都留在原地 —— 与数据范围面板"动作条必须是滚动容器的
            兄弟"是同一条布局纪律。
          */}
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
