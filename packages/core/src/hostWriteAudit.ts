/**
 * 「实际写入了几次」的**唯一口径**（A1-E3）。
 *
 * 为什么需要一份共享口径：验收说明过去把 `/executions/{id}/actions/start` 的
 * 请求行数当成"实际写入次数"。那只是**请求**次数，甚至只是"有人申请过写这次动作"：
 * 重复请求会被账本拒（409），认领成功也可能在碰宿主之前就失败，一次动作还可能
 * 分块写多次。拿它当写入次数，等于把"申请"说成"事实"。
 *
 * 本模块只定义**能观测到的事实**，并按"能不能下结论"分层：
 *
 * | 事实 | 来源 | 能证明什么 |
 * | --- | --- | --- |
 * | 请求次数 / 认领结果 | 服务端请求日志与账本 | 有人申请过、申请被接受还是被拒 |
 * | `hostWriteCalls` | 执行器在**真的调用宿主写 API 处**累加 | 这一次动作向宿主发起了几次写调用 |
 * | `outcome` | 执行器 | 写前失败 / 写后失败 / 已写入 / 结果未知 |
 * | `readback` | 执行器的回读核对 | 目标区域是否真的等于计划值 |
 * | 完成回执 | 账本 | 服务端收到了这次动作的写后检查点 |
 *
 * 纪律：
 * - `unknown` **不是** 0，也不是成功：宿主抛异常、响应丢失、只写了一半，都只能
 *   记"结果未知"，据实说"无法确认"，不许推测；
 * - 只有请求日志时，只能称"请求次数"，不能称为写入次数；
 * - 只有"动作数"（每动作一次）也不等于"底层写 API 调用数"：Excel 一次动作的写入
 *   可能按批提交多次，WPS 的分块写每个块各调用一次。
 */

/** 一次动作的宿主写入事实（不含任何单元格值、公式正文或工具参数）。 */
export interface HostWriteAudit {
  /** 计划里的动作序号（0 起，与账本 `nextActionIndex` 同一口径）。 */
  actionIndex: number;
  /** 动作类型（白名单里的名字，不含参数）。 */
  actionType: string;
  /**
   * 进入宿主写 API 的调用次数。
   *
   * - Excel：该动作窗口内**承载写入**的 `context.sync()` 次数（Office.js 把排队的
   *   写操作按批提交，一次刷新可能承载多个写操作 ⇒ 这是**批次数**）；
   * - WPS：逐次赋值调用次数（分块写每块一次）。
   */
  hostWriteCalls: number;
  /** 该动作排队的写操作个数（Excel 记为本动作执行的写 API 应用次数）。 */
  mutatingCalls: number;
  /**
   * 结果四态：
   * - `written`：宿主调用已返回；
   * - `failed_before_write`：还没进入宿主写 API 就失败了（**不是**已写入）；
   * - `failed_after_write`：已进入写 API 但随后失败（可能写了一半，需回读核对）；
   * - `unknown`：没有拿到确定的宿主结果（异常/断线/超时）。
   */
  outcome: "written" | "failed_before_write" | "failed_after_write" | "unknown";
  /** 回读核对结果。`not_applicable` = 这次没做回读（例如写前就失败）。 */
  readback: "verified" | "failed" | "not_applicable";
  /** 固定错误码（不是宿主原文），可选。 */
  errorCode?: string | null;
}

const OUTCOMES: ReadonlySet<string> = new Set([
  "written",
  "failed_before_write",
  "failed_after_write",
  "unknown"
]);
const READBACKS: ReadonlySet<string> = new Set(["verified", "failed", "not_applicable"]);

/** 计数必须是**非负整数**：`-1`、`1.5`、`NaN` 一律拒绝，不静默取整。 */
function cleanCount(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  if (!Number.isInteger(value) || value < 0) return null;
  return value;
}

/**
 * 规范化一条上报。返回 `null` 表示这条**不可采信**（调用方应当丢弃并如实写
 * "没有采到宿主写调用数"，而不是补一个 0）。
 */
export function normalizeHostWriteAudit(raw: unknown): HostWriteAudit | null {
  if (!raw || typeof raw !== "object") return null;
  const item = raw as Record<string, unknown>;
  const actionIndex = cleanCount(item.actionIndex);
  const hostWriteCalls = cleanCount(item.hostWriteCalls);
  const mutatingCalls = cleanCount(item.mutatingCalls);
  const actionType = typeof item.actionType === "string" ? item.actionType.trim() : "";
  const outcome = typeof item.outcome === "string" ? item.outcome : "";
  const readback = typeof item.readback === "string" ? item.readback : "";
  if (actionIndex === null || hostWriteCalls === null || mutatingCalls === null) return null;
  if (!actionType || !OUTCOMES.has(outcome) || !READBACKS.has(readback)) return null;
  const errorCode =
    typeof item.errorCode === "string" && item.errorCode.trim()
      ? item.errorCode.trim().slice(0, 60)
      : null;
  return {
    actionIndex,
    actionType,
    hostWriteCalls,
    mutatingCalls,
    outcome: outcome as HostWriteAudit["outcome"],
    readback: readback as HostWriteAudit["readback"],
    errorCode
  };
}

/** 汇总只做加法，不做推断：`unknown` 单独计数，绝不并进"已写入"。 */
export interface HostWriteSummary {
  /** 有记库的条目数（= 有事实支撑的动作数）。 */
  reportedActions: number;
  /** 宿主写调用总数（只统计 `outcome !== "failed_before_write"` 的条目）。 */
  hostWriteCalls: number;
  actionsWritten: number;
  actionsFailedBeforeWrite: number;
  actionsFailedAfterWrite: number;
  actionsUnknown: number;
  readbackVerified: number;
  readbackFailed: number;
}

export function summarizeHostWrites(entries: readonly HostWriteAudit[]): HostWriteSummary {
  const summary: HostWriteSummary = {
    reportedActions: entries.length,
    hostWriteCalls: 0,
    actionsWritten: 0,
    actionsFailedBeforeWrite: 0,
    actionsFailedAfterWrite: 0,
    actionsUnknown: 0,
    readbackVerified: 0,
    readbackFailed: 0
  };
  for (const entry of entries) {
    if (entry.outcome !== "failed_before_write") summary.hostWriteCalls += entry.hostWriteCalls;
    if (entry.outcome === "written") summary.actionsWritten += 1;
    else if (entry.outcome === "failed_before_write") summary.actionsFailedBeforeWrite += 1;
    else if (entry.outcome === "failed_after_write") summary.actionsFailedAfterWrite += 1;
    else summary.actionsUnknown += 1;
    if (entry.readback === "verified") summary.readbackVerified += 1;
    else if (entry.readback === "failed") summary.readbackFailed += 1;
  }
  return summary;
}

/**
 * **观测覆盖**：这批数字能不能当成"零写入"的证据（A1-E3 第三批）。
 *
 * 补报是 best-effort，可能一条都到不了。所以空数组必须与"明确 0 次调用"分开：
 *
 * | `observation` | `callEvidence` | 能说的话 |
 * | --- | --- | --- |
 * | `not_collected` | `none` | 没收到观测 —— 对宿主一无所知 |
 * | `partial` | `nonzero` | 已经看到非零调用（覆盖还不全，同看 `missingActions`） |
 * | `partial` | `unknown` | 只覆盖了一部分 ⇒ 不能说零 |
 * | `complete` | `zero` | **明确观测到零写调用** |
 * | `complete` | `unknown` | 有结果未知的条目 ⇒ 不能说零 |
 *
 * 判据本身由**服务端**给出（只有它有账本里的动作认领边界）。这里只做类型与
 * 文案，不再另写一份"覆盖率推断"—— 两份推断一定会漂移。
 *
 * 唯一例外是**自洽复核**：`zero` 与 `missingActions`/`unexpectedReports` 矛盾时
 * 拒绝升级为"零写入"（见 `describeHostWriteObservation`）。那不是第二份推断，
 * 是拒绝把矛盾数据当成更强的结论。
 */
export interface HostWriteObservation {
  observation: "not_collected" | "partial" | "complete";
  callEvidence: "none" | "zero" | "nonzero" | "unknown";
  /** 账本判定"有可能写过"的动作数（认领过 = 已完成前缀 ∪ 在途那一个）。 */
  attemptedActions: number;
  reportedActions: number;
  missingActions: number;
  unexpectedReports: number;
  /** **计量单位不同，不许互相折算**。 */
  unit: "excel_sync_batches" | "wps_api_calls";
  lastReportedAt: number | null;
}

export interface HostWriteObservationVerdict {
  /** 一句话（界面/验收记录直接用，不要各写一句）。 */
  label: string;
  /** 只有这样才允许说"明确没有向宿主发起写调用"。 */
  provesZeroWrites: boolean;
  /** 这批数字能不能下结论（无论结论是零还是非零）。 */
  conclusive: boolean;
}

/**
 * 把服务端给的覆盖判据折成一句人话。
 *
 * `provesZeroWrites` 只在 `complete + zero` 时为真：**空数组、部分覆盖、有结果
 * 未知的条目一律为假** —— 这正是"没有观测"与"明确零写入"的分界。
 *
 * 另外还要自洽：`zero` 却带着 `missingActions > 0` 或 `unexpectedReports > 0`
 * 是矛盾的线上数据（服务端不会产生），这里**拒绝升级**成"零写入"，而不是
 * 照着 `callEvidence` 字面读。
 */
export function describeHostWriteObservation(
  observation: HostWriteObservation
): HostWriteObservationVerdict {
  const unitLabel =
    observation.unit === "excel_sync_batches" ? "宿主刷新批次" : "宿主写 API 调用";
  if (observation.observation === "not_collected") {
    return {
      label:
        observation.attemptedActions > 0
          ? `未采集到宿主观测（账本显示 ${observation.attemptedActions} 个动作可能写过）`
          : "未采集到宿主观测（账本也没有动作被认领）",
      provesZeroWrites: false,
      conclusive: false
    };
  }
  if (observation.callEvidence === "nonzero") {
    return {
      label:
        observation.missingActions > 0
          ? `已观测到非零${unitLabel}，但还缺 ${observation.missingActions} 个动作的观测`
          : `已观测到非零${unitLabel}`,
      provesZeroWrites: false,
      conclusive: observation.missingActions === 0
    };
  }
  if (observation.callEvidence === "zero") {
    /*
     * `zero` 只是**调用计数**这一面的事实，不是覆盖面的保证。
     *
     * 「明确零写入」需要两件事同时成立：边界内每个动作都有观测（`complete`），
     * 而且没有落在边界外的可疑上报。少任何一件，"零"就只是"我看到的那部分里
     * 没有写"—— 那不能推出"没有写"。这里的复核不是第二份覆盖推断（覆盖判定仍然
     * 只有服务端一份），而是拒绝把一条**自相矛盾**的线上数据升级成更强的结论：
     * `partial + zero` 在服务端不会产生，但展示层不能因此假定它永远不会出现。
     */
    const selfConsistent =
      observation.observation === "complete" &&
      observation.missingActions === 0 &&
      observation.unexpectedReports === 0;
    if (!selfConsistent) {
      return {
        label:
          observation.unexpectedReports > 0
            ? `有 ${observation.unexpectedReports} 条上报落在执行边界之外，不能据此判断零写入`
            : `观测只覆盖 ${observation.reportedActions} 个动作（应覆盖 ${observation.attemptedActions} 个），不能说零写入`,
        provesZeroWrites: false,
        conclusive: false
      };
    }
    return {
      label: `已观测到零${unitLabel}（覆盖 ${observation.attemptedActions} 个动作）`,
      provesZeroWrites: true,
      conclusive: true
    };
  }
  if (observation.callEvidence === "unknown") {
    return {
      label:
        observation.missingActions > 0
          ? `观测不完整（缺 ${observation.missingActions} 个动作），结果未知`
          : "观测结果未知（有动作没拿到确定的宿主结果）",
      provesZeroWrites: false,
      conclusive: false
    };
  }
  return {
    label: "未采集到宿主观测",
    provesZeroWrites: false,
    conclusive: false
  };
}
