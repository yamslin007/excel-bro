/**
 * 「暂停」这件事的三个独立概念（2026-10-23）。
 *
 * 过去的实现把它们混成了一两个：
 *  - **交互锁定**（`busy`）：结果在等用户、正在确认数据范围……都可能锁定输入；
 *  - **任务正在执行**（`progress` / `taskRunning`）：界面上有一条进度；
 *  - **当前可暂停**（有没有一个能 abort 的对象）：只有这个才该显示 ■。
 *
 * 混用的后果是实测到的那个缺陷：按钮显示 ■，点下去却因为没有可中止对象而**静默
 * 返回**，用户看到的是"点了没反应"。这里把判据抽成纯函数，两端读同一份。
 */

/** 按钮形态。`unavailable` = 有任务在跑、但本端没有可中止对象（必须给反馈）。 */
export type PauseButtonKind = "pause" | "send" | "continue" | "unavailable";

export type PauseButtonState = {
  kind: PauseButtonKind;
  /** 按钮上的字形：■ 暂停 / ↑ 发送 / ▶ 继续 / … 不可暂停。 */
  glyph: string;
  /** 无障碍标签与悬停说明。 */
  label: string;
  title: string;
  /** 点击是否真的能中止当前执行。 */
  canPause: boolean;
  /** `unavailable` 时的具体原因；由调用方如实展示，不许吞掉。 */
  reason: string | null;
};

export type PauseButtonInput = {
  /** 输入框里有内容（有内容时按钮是发送/打断，不是暂停）。 */
  hasInput: boolean;
  /** 已经暂停、等待用户继续。 */
  paused: boolean;
  /** 本端认为有任务在执行（进度/状态非空）。 */
  taskRunning: boolean;
  /** 存在一个未 abort 的 controller —— **唯一**的"可暂停"依据。 */
  controllerAlive: boolean;
  /** 写入/回读等明确不允许中断的阶段。 */
  uninterruptible?: boolean;
};

export const PAUSE_UNAVAILABLE_REASON =
  "本端已经没有可中止的执行；正在等状态核对结果，不会静默。";
export const PAUSE_UNINTERRUPTIBLE_REASON =
  "写入与回读期间不能中断，以免结果状态不明。";

/**
 * 决定按钮该长什么样、以及点下去会不会真的生效。
 *
 * 顺序是有意的：输入优先（用户在打字时按钮就是发送），其次是"已暂停 → 继续"，
 * 然后才轮到暂停/不可暂停。**任何情况下都不允许"显示 ■ 但 canPause 为假"**。
 */
export function pauseButtonState(input: PauseButtonInput): PauseButtonState {
  if (input.hasInput && input.taskRunning) {
    /*
     * 有输入 + 有任务在跑：这个 ↑ 不是"再发一条"，而是**打断并补充**。
     * 文案必须分开，否则用户会以为是排队发送，实际却把当前这一轮掐掉了。
     */
    return {
      kind: "send",
      glyph: "↑",
      label: "打断并补充",
      title: "打断当前处理，带上这句话重新开始",
      canPause: false,
      reason: null
    };
  }
  if (input.hasInput) {
    return {
      kind: "send",
      glyph: "↑",
      label: "发送",
      title: "发送这条需求",
      canPause: false,
      reason: null
    };
  }
  if (input.paused && !input.taskRunning) {
    return {
      kind: "continue",
      glyph: "▶",
      label: "继续任务",
      title: "继续上次暂停的任务",
      canPause: false,
      reason: null
    };
  }
  if (input.taskRunning && input.uninterruptible) {
    return {
      kind: "unavailable",
      glyph: "…",
      label: PAUSE_UNINTERRUPTIBLE_REASON,
      title: PAUSE_UNINTERRUPTIBLE_REASON,
      canPause: false,
      reason: PAUSE_UNINTERRUPTIBLE_REASON
    };
  }
  if (input.taskRunning && input.controllerAlive) {
    return {
      kind: "pause",
      glyph: "■",
      label: "暂停任务",
      title: "暂停接收规划进度，保留需求；核对检查点后才可继续",
      canPause: true,
      reason: null
    };
  }
  if (input.taskRunning) {
    return {
      kind: "unavailable",
      glyph: "…",
      label: PAUSE_UNAVAILABLE_REASON,
      title: PAUSE_UNAVAILABLE_REASON,
      canPause: false,
      reason: PAUSE_UNAVAILABLE_REASON
    };
  }
  return {
    kind: "send",
    glyph: "↑",
    label: "发送",
    title: "发送这条需求",
    canPause: false,
    reason: null
  };
}

/**
 * 过程卡/进度条上的「停止」。
 *
 * 与输入框那个按钮读**同一条**"能不能停"的判据，但**不受输入影响**：用户在运行中
 * 打了半句话，不该让卡上的停下入口消失（那正是"卡片上的暂停按不到了"的来源）。
 * 输入只决定输入框按钮是 ↑ 还是 ■。
 */
export function executionPauseState(
  input: Omit<PauseButtonInput, "hasInput" | "paused">
): PauseButtonState {
  return pauseButtonState({ ...input, hasInput: false, paused: false });
}

/**
 * 一次执行收尾时的**分项**归属判断。
 *
 * 三样资源各有各的所有者，不能塞进同一个 `ownsTurn`：
 *  - 活动（卡与耗时）由**开它的那次执行**收尾 —— 谁开谁关，哪怕 controller 已经被
 *    状态核对接走；
 *  - `controllerRef` 只由**仍持有它**的那次执行清空；
 *  - 本机读表桥只由**最后一次把它立起来**的那次执行关闭。
 *
 * 只用一个布尔判断的后果是实测到的：核对把 controller 接走之后，恢复流的
 * `finally` 认为"这一轮不归我"，于是活动没结束（卡一直"进行中"、耗时继续涨）、
 * 桥也没收（迟到的进度又把按钮点亮）。
 */
export type ExecutionTeardownPlan = {
  /** 结束属于本次执行的活动（卡 + 计时）。 */
  settleActivity: boolean;
  /** 清空 `controllerRef`（只有仍归本次执行时才清）。 */
  clearController: boolean;
  /** 关闭本机读表桥（只有本执行仍是桥主人时才关）。 */
  endBridge: boolean;
};

export function planExecutionTeardown(input: {
  /** 本次执行的身份（每次执行唯一、可重复收尾）。 */
  owner: string;
  /** 当前 controller 的主人；null 表示已经没有人持有。 */
  controllerOwner: string | null;
  /** 当前桥的主人；null 表示桥已收或从未建立。 */
  bridgeOwner: string | null;
  /** 本次执行的活动是否已经收过尾（重复收尾必须安全）。 */
  activitySettled: boolean;
  /** 本次执行是否真的开过活动。 */
  hasActivity: boolean;
}): ExecutionTeardownPlan {
  return {
    settleActivity: input.hasActivity && !input.activitySettled,
    clearController: input.controllerOwner === input.owner,
    endBridge: input.bridgeOwner === input.owner
  };
}

// ---------------------------------------------------------------------------
// 执行身份：跟着对象走，而不是靠一堆散落的 `===` 比较
// ---------------------------------------------------------------------------

/**
 * 一次执行的身份。唯一、短、不含任何业务内容（不带 turnId、不带正文）。
 *
 * 存在的理由：`controllerRef.current === controller` 这种比较只在"没人接管"时成立。
 * 状态核对会把同一个 ref 换成自己的 controller，之后所有 `===` 判断都变成"不归我"，
 * 于是本来该由本次执行收尾的活动/桥就永远留在原地。
 */
let executionCounter = 0;

export function newExecutionId(prefix = "exec"): string {
  executionCounter += 1;
  return `${prefix}-${executionCounter}`;
}

/**
 * controller → 执行身份。
 *
 * 用 WeakMap 而不是在每个 `controllerRef.current = x` 旁边再记一个变量：赋值点有
 * 七八处，多记一个字段就一定有一处忘掉，而"忘掉"的表现恰好是"某个入口又不收尾"。
 * 身份贴着对象走，取的时候不可能取到别人的。
 */
const controllerOwners = new WeakMap<AbortController, string>();

export function tagControllerOwner(
  controller: AbortController | null | undefined,
  owner: string
): void {
  if (!controller) return;
  controllerOwners.set(controller, owner);
}

/** 当前 controller 的主人；`null` = 没有对象，或这个对象不是本层创建的。 */
export function controllerOwnerOf(
  controller: AbortController | null | undefined
): string | null {
  if (!controller) return null;
  return controllerOwners.get(controller) ?? null;
}

// ---------------------------------------------------------------------------
// 有界等待
// ---------------------------------------------------------------------------

/** 暂停请求/状态核对共用：超时不是失败结论，只是"暂时没结论"。 */
export class PauseTimeoutError extends Error {
  readonly timeoutMs: number;

  constructor(timeoutMs: number, what = "暂停请求") {
    super(`${what}在 ${timeoutMs}ms 内没有结果`);
    this.name = "PauseTimeoutError";
    this.timeoutMs = timeoutMs;
  }
}

/**
 * 给"暂停请求"和"状态核对"加一个上界。
 *
 * 刻意**不**声称取消了底层请求：这里的 fetch 没有接 signal，超时只是让界面能
 * 如实说出"未确认"，而不是把不确定说成已暂停。
 */
export function withPauseTimeout<T>(
  work: Promise<T>,
  timeoutMs: number,
  what = "暂停请求"
): Promise<T> {
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) return work;
  return new Promise<T>((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      reject(new PauseTimeoutError(timeoutMs, what));
    }, timeoutMs);
    const finish = (callback: () => void) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      callback();
    };
    work.then(
      (value) => finish(() => resolve(value)),
      (error) => finish(() => reject(error))
    );
  });
}

// ---------------------------------------------------------------------------
// 最小诊断记录
// ---------------------------------------------------------------------------

/**
 * 暂停链路的最小诊断事件。
 *
 * **只允许结构化字段**：没有自由文本，因此不可能带出提示词正文、单元格值或密钥。
 * `turnId` 只留尾部若干位，够把同一条任务的多条记录串起来，又不把完整身份留在日志里。
 */
export type PauseDiagnosticKind =
  | "pause-click"
  | "pause-unavailable"
  | "pause-request"
  | "pause-ack"
  | "pause-failed"
  | "pause-timeout"
  | "pause-requested"
  | "teardown"
  | "late-progress-ignored";

export type PauseDiagnosticEvent = {
  kind: PauseDiagnosticKind;
  /** 毫秒时间戳（`Date.now()`），用于对齐"两个入口谁先点"。 */
  at: number;
  /** 入口来源：输入框 / 过程卡 / 系统内部。 */
  source?: "composer" | "card" | "system" | "core";
  /** 本次执行身份（调用方保证唯一；不收提示词、不收工作簿内容）。 */
  executionId?: string;
  /** 任务身份尾部（只留 6 位）。 */
  turnIdTail?: string;
  controller?: "owned" | "stale" | "aborted" | "none";
  bridge?: "owned" | "other" | "none";
  outcome?: "settled" | "skipped" | "cleared" | "ended" | "confirmed" | "unconfirmed";
  reasonCode?: string;
};

const DIAGNOSTIC_LIMIT = 40;
const diagnostics: PauseDiagnosticEvent[] = [];

/** 只留尾部若干位：够串联、不落完整身份。 */
export function tailOf(value: string | null | undefined, keep = 6): string | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  return trimmed.length <= keep ? trimmed : trimmed.slice(-keep);
}

export function recordPauseDiagnostic(event: Omit<PauseDiagnosticEvent, "at"> & { at?: number }): void {
  diagnostics.push({ ...event, at: event.at ?? Date.now() });
  if (diagnostics.length > DIAGNOSTIC_LIMIT) diagnostics.splice(0, diagnostics.length - DIAGNOSTIC_LIMIT);
}

/** 供测试与"导出诊断"使用的只读快照。 */
export function pauseDiagnosticSnapshot(): readonly PauseDiagnosticEvent[] {
  return diagnostics.map((event) => ({ ...event }));
}

export function resetPauseDiagnostics(): void {
  diagnostics.length = 0;
}
