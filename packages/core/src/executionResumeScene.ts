/**
 * 「剩余动作还能不能继续」的**现场**判据（宿主无关，纯函数）。
 *
 * 存在理由（2026-10-31 真机）：撤销成功之后服务端处置留痕那次上报失败了（网络剪断），
 * 于是账本里没有 `undone_at`，只有「已完成 2 / 共 3 步」。重开窗格时账本说"可续跑"，
 * 而**表已经不是那个样子了** —— 撤销把两个公式清空了。此刻若照账本续跑，就会在
 * 一次已经回退掉的执行后面补写第 3 步，并把整次执行记成完成。
 *
 * 所以续跑不能只看账本进度，还要看「这次执行写入过的目标现在还是不是它留下的样子」。
 * 证据来自账本里已持久化的写前快照（`checkpoint.ranges[*].postFingerprint` 等写后
 * 摘要），现场由宿主只读一次；本模块只做**政策**（谁必须被覆盖、看到什么才算数），
 * 读取与"两个摘要算不算一样"留在宿主（Excel：`observeOfficeExecutionPostState`）。
 *
 * 三条边界：
 *
 * - **只判不写**：没有网络/宿主调用，输入是宿主读回来的观测结果。
 * - **没证据就不放行**：某个已完成动作没有写后摘要 ⇒ 无法证明它还在，按 `no_evidence`
 *   拒绝 —— 这不算"更严的新闸门"，它就是撤销判据（`rebuildExecutionUndo` 的
 *   `coverage_incomplete`）在续跑这一侧的同一条事实。
 * - **与来源判据分开说**：现场对不上（本轮）和来源变了
 *   （`verifyFormulaResumeSources`）是两种事实，用户要做的事也不同 —— 一个去撤销，
 *   一个去改回来源。
 */

/** 宿主对某个动作目标的单次只读观测。 */
export type ResumeSceneObservationState = "matched" | "changed" | "unreadable";

export interface ResumeSceneObservation {
  actionIndex: number;
  state: ResumeSceneObservationState;
}

export type ResumeSceneReason = "no_evidence" | "unreadable" | "changed";

export type ResumeSceneOutcome =
  | { ok: true; checked: number }
  | { ok: false; reason: ResumeSceneReason; detail: string };

export interface ResumeSceneInput {
  /** 账本进度：已完成几个动作（= 需要被覆盖的目标个数）。 */
  completedActionCount: number;
  /** 写前快照里有写后摘要、可被核对的已完成动作号。 */
  coveredActionIndexes: readonly number[];
  /** 宿主读回来的观测（可能只覆盖一部分）。 */
  observations: readonly ResumeSceneObservation[];
}

/**
 * 判定「这次执行写入过的目标是否仍是执行后的样子」。
 *
 * 顺序刻意如此：先看证据覆盖（没有摘要就没什么可比的），再看现场。
 * 只有**每一个**已完成动作都是 `matched` 才返回 ok —— 少一个都不算。
 */
export function verifyResumeScene(input: ResumeSceneInput): ResumeSceneOutcome {
  const completed = Math.max(0, Math.trunc(input.completedActionCount));
  if (completed === 0) return { ok: true, checked: 0 };
  const covered = new Set(input.coveredActionIndexes.filter((index) => Number.isInteger(index)));
  for (let index = 0; index < completed; index += 1) {
    if (!covered.has(index)) {
      return {
        ok: false,
        reason: "no_evidence",
        detail:
          `第 ${index + 1} 个动作没有留下执行后校验，无法确认这次写入还在表里；` +
          "不能继续剩余动作，也不能按旧进度补记。"
      };
    }
  }
  const byIndex = new Map<number, ResumeSceneObservationState>();
  for (const observation of input.observations) {
    if (Number.isInteger(observation.actionIndex)) byIndex.set(observation.actionIndex, observation.state);
  }
  for (let index = 0; index < completed; index += 1) {
    const state = byIndex.get(index);
    if (state === "matched") continue;
    if (state === "changed") {
      return {
        ok: false,
        reason: "changed",
        detail:
          `第 ${index + 1} 个动作的目标已经不是这次执行写入后的样子（可能已被撤销或被改过）；` +
          "不能继续剩余动作。要回退已写入的内容，请用本次执行的撤销。"
      };
    }
    return {
      ok: false,
      reason: "unreadable",
      detail:
        `读不到第 ${index + 1} 个动作的目标现在是什么样子，无法确认这次写入还在；` +
        "不能继续剩余动作，请稍后重新核对（只读）。"
    };
  }
  return { ok: true, checked: completed };
}
