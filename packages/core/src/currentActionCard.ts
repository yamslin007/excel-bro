/**
 * 输入框上方那块**固定区域**里，现在到底该强调哪一件事（Excel / WPS 共用，纯函数）。
 *
 * 改造前这里是四条互不知情的通道各自决定自己出不出现：结果审阅栏、写入恢复卡、
 * 澄清回答区、任务状态行。四者可以同时在场，于是"现在最该做什么"要靠用户自己
 * 从四块 UI 里挑 —— 而它们的视觉分量还都差不多。
 *
 * 本模块只回答一个问题：**按任务生命周期，哪一件更需要用户现在处理**。
 * 它不隐藏任何东西（其余的会出现在 `others` 里，由界面如实说明"还有 N 项"），
 * 也不把多个恢复候选合并成一个 —— 恢复候选的挑法在 `selectExecutionRecoveryEntry`
 * （按 executionId 取最后一条），两者是两件事，不许互相顶替。
 *
 * 优先级与理由（顺序即判据）：
 *
 *  1. `review`    **已经写进工作簿的东西**在等你确认/放弃。放着不管的代价最大：
 *                服务端还在等你确认，副本与标记都悬着，而且它有时间成本。
 *  2. `recovery`  上一次执行留下了在途动作，现场不明。**先核对再续跑**；
 *                但核对是只读的，等审阅收尾再点也不会多写一个单元格。
 *  3. `clarification` 一道问题在等回答。它只阻塞"这一轮规划"，没有写入悬空。
 *  4. `runtime`   只是"在跑"。运行控制就在输入框右下角，不需要在固定区域再占一张卡。
 *  5. `none`      没有需要用户处理的事。
 */

export type CurrentActionKind =
  | "review"
  | "recovery"
  | "clarification"
  | "runtime"
  | "none";

/** 判据顺序，也是"其余待办"的展示顺序。 */
export const CURRENT_ACTION_PRIORITY: readonly Exclude<CurrentActionKind, "none">[] = [
  "review",
  "recovery",
  "clarification",
  "runtime"
];

export interface CurrentActionInput {
  /** 结果在等确认 / 放弃（`resultReviewState.phase === "reviewing"` 那一类事实）。 */
  review: boolean;
  /** 有未收尾的执行需要**先核对**（不是"有历史执行"，是"现在有待处理动作"）。 */
  recovery: boolean;
  /** 有一道问题在等回答，且用户没有显式选择"另提新需求"。 */
  clarification: boolean;
  /** 任务在跑（不含"在等你看结果"—— 那是 `review`）。 */
  running: boolean;
}

export interface CurrentActionSelection {
  kind: CurrentActionKind;
  /** 同样待处理、但优先级更低的其余事项。**用于如实说明，不用于隐藏。** */
  others: Exclude<CurrentActionKind, "none">[];
}

const CURRENT_ACTION_LABEL: Record<CurrentActionKind, string> = {
  review: "确认或放弃这次写入的结果",
  recovery: "核对上次写入留下的在途动作",
  clarification: "回答待确认的问题",
  runtime: "任务正在运行",
  none: ""
};

/** 紧凑提示用的短名（"另有 2 项：审阅结果、回答提问"）。 */
const CURRENT_ACTION_SHORT_LABEL: Record<CurrentActionKind, string> = {
  review: "审阅结果",
  recovery: "恢复核对",
  clarification: "回答提问",
  runtime: "正在运行",
  none: ""
};

/** 一句话说明"现在这件事是什么"。空字符串 = 没有需要用户处理的事。 */
export function currentActionLabel(kind: CurrentActionKind): string {
  return CURRENT_ACTION_LABEL[kind];
}

/** 短名：只用在紧凑提示里，不替代 `currentActionLabel` 的完整说法。 */
export function currentActionShortLabel(kind: CurrentActionKind): string {
  return CURRENT_ACTION_SHORT_LABEL[kind];
}

/**
 * 每个形态对应哪一个**事实**输入。
 *
 * 单独一张表而不是 `input[kind]`：`runtime` 的事实叫 `running`（"在跑"是状态，
 * `runtime` 是这块固定区域里的形态名）。用同一串 key 去查，`runtime` 会永远查不到，
 * 于是"在跑"这件事**永远不出现**在"还有 N 项"里 —— 那不是少一行字，是少说一件事。
 */
const PRESENT: Record<Exclude<CurrentActionKind, "none">, (input: CurrentActionInput) => boolean> = {
  review: (input) => input.review,
  recovery: (input) => input.recovery,
  clarification: (input) => input.clarification,
  runtime: (input) => input.running
};

export function selectCurrentAction(
  input: CurrentActionInput,
  /**
   * 用户**显式**切到了哪一件（点击紧凑提示里的切换入口）。
   *
   * 为什么需要它：只按固定优先级选，另一条任务就被永久压在下面 —— 那正是
   * "用优先级遮住任务"。这里让用户的显式选择压过默认顺序，但**只认仍然存在的**：
   * `pinned` 指的那件事已经不在（被处理掉了），立刻退回默认顺序，不留一张空卡。
   */
  pinned?: Exclude<CurrentActionKind, "none"> | null
): CurrentActionSelection {
  const present = CURRENT_ACTION_PRIORITY.filter((kind) => PRESENT[kind](input));
  if (present.length === 0) return { kind: "none", others: [] };
  if (pinned && present.includes(pinned)) {
    return { kind: pinned, others: present.filter((kind) => kind !== pinned) };
  }
  return { kind: present[0], others: present.slice(1) };
}

/**
 * 紧凑提示要列出哪些"其余事项"，以及它们各自切过去之后会发生什么。
 *
 * 与 `others` 同源，不另算一遍 —— 另算一份就会出现"提示里说还有一项，
 * 切过去却没有"。
 */
export function currentActionSwitchTargets(
  selection: CurrentActionSelection
): { kind: Exclude<CurrentActionKind, "none">; label: string; short: string }[] {
  return selection.others.map((kind) => ({
    kind,
    label: CURRENT_ACTION_LABEL[kind],
    short: CURRENT_ACTION_SHORT_LABEL[kind]
  }));
}

/**
 * 固定区域里"除了当前这件事，还有几件事"。
 *
 * 单列出来是因为界面要用它说一句实话：**没有隐藏**。以前的做法是不说 ——
 * 用户从别的入口（消息里的记录、过程卡）偶然发现自己还有一件事没处理。
 */
export function currentActionOtherCount(selection: CurrentActionSelection): number {
  return selection.others.length;
}
