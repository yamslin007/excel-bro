/**
 * 「重开之后还能不能撤销这次未完成的写入」的共享判据（宿主无关）。
 *
 * 存在理由（2026-10-31 真机）：三步公式计划在第二步回读后中断 → 用户核对在途动作 →
 * 改来源 → 续跑被拒（**这条拒绝是对的**）。但用户重开 Excel 后，恢复卡仍在、写入的
 * 两个公式仍在表里，界面却**没有**「撤销本次执行」——因为撤销栈只在
 * `status === "completed"` 时才用 `result.undoSnapshot` 重建，未完成/被阻断的分支一条
 * 都没重建，而账本里其实**已经持久保存着**写前快照（`checkpoint`：每个已完成动作的
 * 写前值/公式 + 写后指纹）。
 *
 * 核心是**两种资格必须分开**：
 *
 * - **继续执行**（`pendingResumeReadiness`）：来源变了就必须拒绝 —— 拿旧结果继续写，
 *   等于把过期数据写进表里。
 * - **撤销已写入内容**（本模块）：它不需要来源一致，它只需要「**已验证的写前恢复点**
 *   ＋ **目标现场仍是这次执行留下的样子**」。来源被用户改了，不该连带把撤销也收走 ——
 *   那会让用户既不能继续、也不能回退，只能手工一格一格擦。
 *
 * 三条不可放宽的边界：
 *
 * 1. **只判不写**：没有任何网络/宿主调用，输入是账本里已持久化的材料，输出是纯结论。
 *    身份、授权、材料完整性、目标写后状态全部照旧由宿主侧再核一次
 *    （Excel：`assertUndoPostState`；撤销前还会核对执行后指纹）。
 * 2. **不现读补造**：这里的快照只来自账本里登记的 `checkpoint`，绝不为了"凑一份能撤销的
 *    快照"去读当前单元格当写前值 —— 那会把用户后来的修改当成"执行前状态"写回去。
 * 3. **拿不准就阻断**：在途动作没核对（结果不明）、材料缺口、进度与材料对不上、
 *    授权收窄、身份不符，一律不给入口，并给出**一个**具体原因码。
 *
 * 与 WPS 的关系：WPS 的撤销恢复点活在 `PluginStorage`（关闭加载项即失效），恢复包里的
 * 写前值目前只用于续跑核对，**没有**接到"重开后撤销"上。本模块是纯判据，WPS 若日后
 * 用恢复包里的 `beforeValues/beforeFormulas` 接这条能力，判定应复用这里，不另写一份。
 */

/** 账本里持久保存的写前快照区域（只取判定用得到的字段）。 */
export interface UndoRebuildRange {
  actionIndex: number;
  sheet: string;
  /** A1 地址，如 `"B2"`。 */
  range: string;
  /** 写后指纹 / 范围摘要；三者至少有一个才能证明"这一段写后核对过"。 */
  postFingerprint?: string;
  scopeDigest?: string;
  extendedFormatDigest?: string;
}

/** 账本里的 `checkpoint`：已完成动作的写前快照累积。 */
export interface UndoRebuildCheckpoint {
  planId: string;
  executionId?: string;
  workbookIdentity?: string;
  coverage?: "full" | "partial";
  completedActionCount?: number;
  /** `complete` 才允许自动覆盖（`assertUndoPostState` 同口径）。 */
  postStateCoverage?: "complete" | "partial";
  ranges: readonly UndoRebuildRange[];
}

/** 计划里某个动作的**可恢复目标**（宿主白名单解析后的单个矩形）。 */
export interface UndoRebuildTarget {
  actionIndex: number;
  sheet: string;
  range: string;
}

/** 用户对这次执行做过的终态处置（首次时刻；`null` = 没有这条留痕）。 */
export interface UndoRebuildDisposition {
  undoneAt?: number | null;
  abandonedAt?: number | null;
}

/**
 * 阻断原因码。每个码只在**一处**返回点上产生，界面按码给文案，不按字符串猜。
 *
 * `nothing_written` 不是异常：一个动作都没写成的执行本来就没有可撤销的东西，
 * 界面**不该**为它画任何撤销说明（画了就成了"又有东西要处理"的噪音）。
 */
export type UndoRebuildBlockReason =
  | "nothing_written"
  | "identity_unavailable"
  | "identity_mismatch"
  | "disposition_recorded"
  | "already_completed"
  | "in_flight_unknown"
  | "material_missing"
  | "coverage_incomplete"
  | "plan_mismatch"
  | "plan_target_never_authorized"
  | "plan_target_not_authorized_now";

export interface UndoRebuildInput<Checkpoint extends UndoRebuildCheckpoint> {
  operationId: string;
  status: "running" | "needs_reconciliation" | "completed" | "failed";
  planId: string;
  /** 登记时的工作簿身份（`url:…`）。 */
  workbookIdentity: string;
  nextActionIndex: number;
  inFlightActionIndex: number | null;
  requiresReconciliation: boolean;
  checkpoint: Checkpoint | null | undefined;
  /** 登记时授权的工作表名。 */
  recordedAuthorization: readonly string[];
  /** 用户**此刻**确认过的范围；`null` = 本会话没确认过（不比对，与续跑判据同口径）。 */
  currentAuthorization: readonly string[] | null;
  currentIdentity: string | null;
  /** 计划每个动作的可恢复目标；解析不出单个矩形的动作给 `null`。 */
  targets: readonly (UndoRebuildTarget | null)[];
  disposition?: UndoRebuildDisposition | null;
}

export type UndoRebuildOutcome<Checkpoint extends UndoRebuildCheckpoint> =
  | {
    available: true;
    /** 账本里登记的写前快照原样返回（调用方只补身份字段，**不许**改写区域内容）。 */
    checkpoint: Checkpoint;
    /** 已写入的动作数（= 这次的进度，而不是计划总数）。 */
    writtenActionCount: number;
    /** 写前快照覆盖的区域数（界面说"几项"用它）。 */
    rangeCount: number;
  }
  | { available: false; reason: UndoRebuildBlockReason; detail: string };

const address = (value: string) =>
  value.replaceAll("$", "").replace(/\s+/g, "").toUpperCase();

/** 处置留痕的一句话（`undone` 与 `abandoned` **不互冒充**）。 */
function dispositionDetail(disposition: UndoRebuildDisposition): string {
  if (disposition.undoneAt) return "这次执行已经撤销过；不再重复撤销。";
  return "这次执行已按「放弃」处置过；不再重复撤销。";
}

/**
 * 判定「重开后能不能撤销这次已写入的内容」。
 *
 * 顺序是刻意的：**先身份、再处置、再状态、再材料、最后授权**。
 * 身份不符时任何一句关于材料的话都不可信；已经处置过的执行不该再被"重建"出一个入口。
 */
export function rebuildExecutionUndo<Checkpoint extends UndoRebuildCheckpoint>(
  input: UndoRebuildInput<Checkpoint>
): UndoRebuildOutcome<Checkpoint> {
  if (!input.currentIdentity) {
    return {
      available: false,
      reason: "identity_unavailable",
      detail: "当前工作簿尚未保存，无法确认它就是原来那个文件，已停止重建撤销。"
    };
  }
  if (input.currentIdentity !== input.workbookIdentity) {
    return {
      available: false,
      reason: "identity_mismatch",
      detail: "当前不是产生这次写入的工作簿，已停止；撤销不会跨工作簿进行。"
    };
  }
  const disposition = input.disposition ?? null;
  if (disposition && (disposition.undoneAt || disposition.abandonedAt)) {
    return { available: false, reason: "disposition_recorded", detail: dispositionDetail(disposition) };
  }
  if (input.status === "completed") {
    return {
      available: false,
      reason: "already_completed",
      detail: "账本已记录这次执行完成；撤销入口跟随它的结果卡，不在恢复卡上重复给一个。"
    };
  }
  if (input.inFlightActionIndex != null || input.requiresReconciliation) {
    const index = input.inFlightActionIndex;
    return {
      available: false,
      reason: "in_flight_unknown",
      detail: index == null
        ? "服务端标记需要核对在途动作，但没有给出位置；先核对，再谈撤销。"
        : `第 ${index + 1} 个动作可能在停止前已经写入，也可能没有；先核对在途动作（只补记回执、不写单元格），再撤销。`
    };
  }
  if (input.nextActionIndex <= 0) {
    return {
      available: false,
      reason: "nothing_written",
      detail: "这次执行没有已完成写入的动作，没有可撤销的内容。"
    };
  }
  const checkpoint = input.checkpoint ?? null;
  if (!checkpoint) {
    return {
      available: false,
      reason: "material_missing",
      detail:
        "账本没有保存这次执行的写前快照，无法知道每个目标执行前是什么样子；不会用当前单元格反推，也不会直接清空目标。"
    };
  }
  if (checkpoint.postStateCoverage !== "complete") {
    return {
      available: false,
      reason: "coverage_incomplete",
      detail: "写前快照缺少完整的执行后核对，不能自动覆盖当前目标；请先人工核对。"
    };
  }
  if (
    checkpoint.planId !== input.planId ||
    (checkpoint.executionId && checkpoint.executionId !== input.operationId) ||
    (checkpoint.workbookIdentity && checkpoint.workbookIdentity !== input.workbookIdentity)
  ) {
    return {
      available: false,
      reason: "plan_mismatch",
      detail: "写前快照不属于这次执行（计划或执行身份对不上），不能当作它的撤销点。"
    };
  }
  const completed = checkpoint.completedActionCount ?? checkpoint.ranges.length;
  if (completed !== input.nextActionIndex) {
    return {
      available: false,
      reason: "plan_mismatch",
      detail: `账本进度（已完成 ${input.nextActionIndex} 个动作）与写前快照（覆盖 ${completed} 个动作）对不上，不能猜。`
    };
  }
  const covered = new Set<number>();
  for (const range of checkpoint.ranges) {
    if (!Number.isInteger(range.actionIndex) || range.actionIndex < 0 || range.actionIndex >= completed) {
      return {
        available: false,
        reason: "plan_mismatch",
        detail: "写前快照里出现不属于已完成动作的区域，不能确定它该不该被恢复。"
      };
    }
    if (!range.postFingerprint && !range.scopeDigest && !range.extendedFormatDigest) {
      return {
        available: false,
        reason: "coverage_incomplete",
        detail: `「${range.sheet}」${range.range} 缺少执行后校验，撤销会无法判断目标是否被改过。`
      };
    }
    const target = input.targets[range.actionIndex] ?? null;
    if (!target || target.sheet !== range.sheet || address(target.range) !== address(range.range)) {
      return {
        available: false,
        reason: "plan_mismatch",
        detail: `写前快照的区域「${range.sheet}」${range.range} 不是第 ${range.actionIndex + 1} 个动作的目标，不能恢复。`
      };
    }
    covered.add(range.actionIndex);
  }
  for (let index = 0; index < completed; index += 1) {
    if (!covered.has(index)) {
      return {
        available: false,
        reason: "coverage_incomplete",
        detail: `第 ${index + 1} 个动作没有留下写前快照，撤销会漏掉它写过的单元格。`
      };
    }
  }
  const recorded = new Set(input.recordedAuthorization);
  for (const range of checkpoint.ranges) {
    if (!recorded.has(range.sheet)) {
      return {
        available: false,
        reason: "plan_target_never_authorized",
        detail: `这次执行的登记授权里没有工作表「${range.sheet}」，记录自相矛盾，不能据此写回。`
      };
    }
  }
  if (input.currentAuthorization) {
    const current = new Set(input.currentAuthorization);
    const missing = [...new Set(checkpoint.ranges.map((range) => range.sheet))]
      .filter((sheet) => !current.has(sheet));
    if (missing.length > 0) {
      return {
        available: false,
        reason: "plan_target_not_authorized_now",
        detail: `当前授权里没有工作表「${missing.join("」「")}」；撤销会写回这些表，请先重新勾选（不会自行扩大授权）。`
      };
    }
  }
  return {
    available: true,
    checkpoint,
    writtenActionCount: completed,
    rangeCount: checkpoint.ranges.length
  };
}

/** 撤销入口按钮文案（与结果卡同名）：这是**同一件事**，不另起一个说法。 */
export const UNDO_RECOVERY_LABEL = "↶ 撤销本次执行";

/**
 * 撤销说明的**唯一**出处。
 *
 * 必须说清两件事，否则用户不敢点：① 只写回本次执行碰过的目标；② 来源保持用户现在的
 * 修改 —— 撤销不是"把整次任务回滚到发起前"，它不碰来源表，也不重新执行公式。
 */
export function undoRecoveryDetail(writtenActionCount: number, rangeCount: number, sourceSheets: readonly string[]): string {
  const source = sourceSheets.length > 0
    ? `来源工作表「${sourceSheets.join("」「")}」保持你现在的修改，不会被写回。`
    : "只写回本次执行碰过的目标区域。";
  return `写前快照已登记：可把这次写入的 ${writtenActionCount} 个动作（${rangeCount} 个区域）恢复到执行前。${source}撤销不会重新执行计划。`;
}
