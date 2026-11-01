/**
 * 「这次执行是用户放弃的，还是撤销的」—— 服务端留痕的唯一判据。
 *
 * 为什么需要它：`office_executions` 里只有进度（`status` / `next_action_index` /
 * `in_flight_action_index`）。用户点了「放弃变更」或「撤销本次执行」之后，账本**一个字
 * 都不变** —— 于是"这次执行到底是被用户撤掉的，还是根本没跑完"只能靠"恢复材料没了 +
 * `revision` 没动 + 截图"三方推断。留痕要解决的正是这一格。
 *
 * 三条纪律（**判据写在这里，两个宿主只按引用调用，不许各写一份"行为相似"的副本**）：
 *
 * 1. 只有两种取值，`abandoned` 与 `undone` 不许互相冒充：
 *    - `abandoned` = 审阅阶段点「放弃变更」（还没有确认过结果）；
 *    - `undone` = 确认之后点「撤销本次执行」。
 * 2. **没有执行身份就不上报**：空白 id 一律 `skipped`，绝不去编一个（留痕不能造事实）。
 * 3. **上报失败不许影响用户已经做完的动作**：恢复已经发生、单元格已经回退，留痕只是观测。
 *    失败返回 `failed`，调用方**不要**弹错误、不要重试成阻塞 —— 与
 *    `office_host_writes` 的 best-effort 口径一致。
 */

/** 用户对一次执行的终态处置。 */
export type ExecutionDisposition = "abandoned" | "undone";

/** 两个取值的用户可读名称（与审阅卡上的按钮同名，别另起一套叫法）。 */
export const EXECUTION_DISPOSITION_LABEL: Readonly<Record<ExecutionDisposition, string>> = {
  abandoned: "放弃变更",
  undone: "撤销本次执行"
};

/** 一次上报的结论。`skipped` 与 `failed` **不是**同一件事，别合成一个布尔。 */
export type ExecutionDispositionOutcome = "recorded" | "skipped" | "failed";

/** 把外部输入收敛成合法取值；不认识就是 `null`（不猜、不默认成 `undone`）。 */
export function normalizeExecutionDisposition(value: unknown): ExecutionDisposition | null {
  if (value === "abandoned" || value === "undone") return value;
  return null;
}

/**
 * 执行身份归一：去首尾空白，空串/非字符串一律 `null`。
 *
 * 不做大小写或前缀归一 —— 账本按**完整字符串**比身份，任何"看起来像"的改写都会让
 * 一条记录被当成另一次执行。
 */
export function executionDispositionTarget(executionId: string | null | undefined): string | null {
  if (typeof executionId !== "string") return null;
  const trimmed = executionId.trim();
  return trimmed.length > 0 ? trimmed : null;
}

/**
 * 上报一次处置留痕。**永不抛错**：返回 `skipped` / `failed` 让调用方知情，但不打断收尾。
 *
 * @param send 端点适配层（各宿主自己的 `recordExecutionDisposition`），抛错由这里吞掉。
 */
export async function recordExecutionDisposition(
  send: (executionId: string, disposition: ExecutionDisposition) => Promise<unknown>,
  executionId: string | null | undefined,
  disposition: ExecutionDisposition
): Promise<ExecutionDispositionOutcome> {
  const target = executionDispositionTarget(executionId);
  if (target === null) return "skipped";
  const kind = normalizeExecutionDisposition(disposition);
  if (kind === null) return "skipped";
  try {
    await send(target, kind);
    return "recorded";
  } catch {
    return "failed";
  }
}
