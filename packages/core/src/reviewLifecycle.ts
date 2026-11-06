/**
 * 结果审阅的**持久生命周期**（Excel / WPS / 服务端同一份口径）。
 *
 * 存在理由（2026-11-03）：审阅状态以前只活在 `ResultReview` 的内存里。执行完成、
 * 回读通过，用户还没点「确认全部变更 / 放弃全部变更」时刷新窗格，待审阅界面就消失，
 * 输入框回到普通状态，甚至写「本次没有需要你处理的后续操作」——而原表停在一个没人
 * 收尾的写入现场上。
 *
 * 三条不能用代替品的判据：
 *
 * - **已完成的执行不等于待审阅**：账本 `status === 'completed'` 只说明动作跑完了。
 *   "用户看过没有"是另一个事实，必须单独持久化（`reviewState`），不许由 status 推。
 * - **标记/双视图/组件挂载都不等于审阅状态**：颜色标记是展示层收尾，双视图是可选
 *   展示材料，二者缺失都不能说明用户确认过。
 * - **`null` 不等于已确认**：本能力上线前的记录没有审阅列，那是"无法判断"，不是
 *   "用户处理过了"。老记录既不批量推成待审阅，也不推成已确认。
 *
 * 本模块只做**确定性判断**，不碰网络、不碰宿主：状态机、恢复核验结论、候选要不要
 * 出现在启动发现里，都从这里出，两端只按引用转发。
 */

/** 账本里持久化的审阅状态。 */
export type ReviewPersistenceState =
  /** 执行完成且结果已验证，等用户决定。 */
  | "pending"
  /** 执行完成但结果未能独立验证：要用户核对，但**不允许**直接当"确认过的结果"接受。 */
  | "blocked"
  /** 用户接受了已有结果（只清标记，未再写一次）。 */
  | "confirmed"
  /** 用户按恢复点放弃，或外部撤销生效。 */
  | "discarded"
  /** 确认之后又被撤销（保留"确认过"与"已撤销"两件事各自的时刻）。 */
  | "undone";

/** 用户能做的一次决定。 */
export type ReviewDecision = "confirmed" | "discarded" | "undone";

const TERMINAL: ReadonlySet<ReviewPersistenceState> = new Set(["discarded", "undone"]);

/** 还要用户处理的两种状态：`blocked` 同样要人看，只是不能直接接受。 */
const NEEDS_ATTENTION: ReadonlySet<ReviewPersistenceState> = new Set(["pending", "blocked"]);

/**
 * 允许的状态迁移。
 *
 * 为什么不是"任意覆盖"：同一份执行可能同时被两个窗格操作。若后到的旧决定能覆盖
 * 已经成立的终态，用户会看到"确认过的结果又变回待审阅"，而原表可能已经被放弃恢复过。
 *
 * - `pending` / `blocked` → 三种决定都可以；
 * - `confirmed` → 只允许 `undone`（撤销就是"确认之后又反悔"，这是真实存在的路）；
 * - `discarded` / `undone` → 终态，任何新决定都不算数（幂等重复不算冲突）。
 */
export function reviewDecisionAllowed(
  current: ReviewPersistenceState,
  decision: ReviewDecision
): boolean {
  if (TERMINAL.has(current)) return false;
  if (current === "confirmed") return decision === "undone";
  return true;
}

export function reviewIsTerminal(state: ReviewPersistenceState): boolean {
  return TERMINAL.has(state);
}

/** 这条记录还要不要让用户看见（启动发现用）。`null` 是"无法判断"，不在此列。 */
export function reviewNeedsAttention(state: ReviewPersistenceState | null): boolean {
  return state !== null && NEEDS_ATTENTION.has(state);
}

/**
 * 一次审阅决定落账后应该是什么状态。
 *
 * `discarded` 与 `undone` 都算"现场已被还原"，但它们是两件事：前者是审阅阶段的
 * 放弃，后者是确认之后（或外部）撤销。分开存是因为撤销入口与「保存为工具」资格
 * 都按它判断，合并会让人以为"放弃"和"撤销"是同一条路。
 */
export function reviewStateAfterDecision(decision: ReviewDecision): ReviewPersistenceState {
  return decision;
}

/**
 * 恢复受阻的原因。每一种都要能单独说清"哪件事不成立"。
 *
 * 注意这里**没有** `scope_mismatch`：授权范围不是"这条审阅还在不在"的判据。
 * 刷新之后，面板里的勾选是**当下**的选择，不是这次执行登记的那一组；拿它去拦
 * 恢复审阅，等于每次刷新都因为"选区没恢复"而把待审阅弄丢（2026-11-03 实测如此）。
 * 恢复出来的审阅一律按**账本登记的授权范围**核对与恢复，当前选区既不参与、
 * 也不会被用来扩大原范围 —— 那是"恢复审阅不授予新任务"的含义，不是"选区必须一致"。
 */
export type ReviewRestoreBlocker =
  | "identity_mismatch"
  | "execution_not_completed"
  | "result_missing"
  | "result_unverified"
  | "record_missing";

export interface ReviewRestoreInput {
  /** 账本里的审阅状态（`null` = 本能力上线前的老记录，无法判断）。 */
  reviewState: ReviewPersistenceState | null;
  /** 账本记录的工作簿身份。 */
  recordedIdentity: string | null;
  /** 当前宿主的稳定身份。 */
  currentIdentity: string | null;
  /** 登记时的授权表名。 */
  recordedSheets: readonly string[];
  /** 当前授权（用户此刻勾选）。**不得**因它变化扩大原范围。 */
  currentSheets: readonly string[];
  /** 账本 `status`。 */
  status: string;
  /** 是否有可解析的完成结果（`result_json`）。 */
  resultPresent: boolean;
  /** 结果是否通过独立验证（`verified === true`）。 */
  resultVerified: boolean;
  /** 本机恢复材料是否可读（决定"放弃"能不能做，不决定要不要显示待审阅）。 */
  materialAvailable: boolean;
}

export interface ReviewRestoreVerdict {
  /** `restorable` 才重建审阅；`settled` = 终态，刷新后不该再出现；其余为受阻。 */
  kind: "restorable" | "settled" | "blocked" | "legacy_unknown";
  blockers: ReviewRestoreBlocker[];
  /** 每个受阻操作各自的说明；能做的操作不出现在这里。 */
  limitations: { discard?: string; comparison?: string; scope?: string };
  /** 只有 `kind==='restorable'` 时有效。 */
  canConfirm: boolean;
  canDiscard: boolean;
}

const NO_LIMITATIONS: ReviewRestoreVerdict["limitations"] = {};

/**
 * 启动/刷新时"这次执行还能不能恢复出审阅"。
 *
 * 顺序固定：先身份、再授权、再看执行与结果、最后看材料。每一层失败都要能单独说清，
 * 否则用户拿到的永远是同一句"不能恢复"。
 */
export function reviewRestoreVerdict(input: ReviewRestoreInput): ReviewRestoreVerdict {
  const blocked = (...blockers: ReviewRestoreBlocker[]): ReviewRestoreVerdict => ({
    kind: "blocked",
    blockers,
    limitations: NO_LIMITATIONS,
    canConfirm: false,
    canDiscard: false
  });

  // 终态优先：已确认/已放弃/已撤销的记录，刷新后不许再变回待审阅。
  if (input.reviewState !== null && reviewIsTerminal(input.reviewState)) {
    return {kind: "settled", blockers: [], limitations: NO_LIMITATIONS, canConfirm: false, canDiscard: false};
  }
  /*
   * 老记录：没有审阅列。**不推成待审阅，也不推成已确认** —— 那是同一份"无法判断"
   * 的两种假装。调用方据此不显示审阅入口，也不声称用户处理过。
   */
  if (input.reviewState === null) {
    return {kind: "legacy_unknown", blockers: [], limitations: NO_LIMITATIONS, canConfirm: false, canDiscard: false};
  }

  const blockers: ReviewRestoreBlocker[] = [];
  if (!input.currentIdentity || !input.recordedIdentity || input.currentIdentity !== input.recordedIdentity)
    blockers.push("identity_mismatch");
  if (input.status !== "completed") blockers.push("execution_not_completed");
  if (!input.resultPresent) blockers.push("result_missing");
  else if (!input.resultVerified) blockers.push("result_unverified");
  if (blockers.length) return blocked(...blockers);

  /*
   * 材料只影响**放弃**与**双视图**，不影响"这条审阅还在"。
   * 混在一起会让材料丢了的现场被报成"没有待审阅"，用户就再也不会回来收尾。
   */
  const limitations: ReviewRestoreVerdict["limitations"] = {};
  const recorded = [...new Set(input.recordedSheets)].sort();
  const current = [...new Set(input.currentSheets)].sort();
  if (JSON.stringify(recorded) !== JSON.stringify(current)) {
    /*
     * 选区与登记范围不同**不阻断**恢复：恢复出来的审阅只按登记范围核对与恢复，
     * 当前选区不会给它加表，也不会被当成新的授权。
     */
    limitations.scope =
      "当前勾选的工作表与这次执行登记的范围不同；恢复审阅只按登记范围核对与恢复，不会因当前选区扩大或缩小原范围。";
  }
  if (!input.materialAvailable) {
    limitations.discard = "本机读不到这次执行的恢复材料，暂时不能放弃变更；已写入的结果仍按原样保留，确认变更不受影响。";
    limitations.comparison = "本机读不到执行前快照，双视图不可用；这不影响确认或放弃。";
  }
  return {
    kind: "restorable",
    blockers: [],
    limitations,
    // `blocked` 状态（结果未通过独立验证）不允许"确认"，但仍要让人看见并处理。
    canConfirm: input.reviewState === "pending" && input.resultVerified,
    canDiscard: input.materialAvailable
  };
}

/** 受阻原因的人话（两端共用，避免各写一份）。 */
export function reviewBlockerMessage(blocker: ReviewRestoreBlocker): string {
  switch (blocker) {
    case "identity_mismatch":
      return "当前工作簿与这次执行登记的身份不一致；请打开原工作簿后再处理。";
    case "execution_not_completed":
      return "这次执行在账本里还没有完成，不能当作已完成结果来审阅。";
    case "result_missing":
      return "这次执行没有完整的完成回执，无法重建审阅所需的结果证据。";
    case "result_unverified":
      return "这次执行的结果没有通过独立验证，不能直接接受；请先核对目标区域。";
    case "record_missing":
    default:
      return "找不到这次执行的账本记录。";
  }
}
