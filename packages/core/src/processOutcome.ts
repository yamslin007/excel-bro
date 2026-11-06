/**
 * 「结果摘要」的共享派生层（Excel / WPS 两端共用，纯函数）。
 *
 * 它回答三件事，且**只**从结构化事实派生：
 *  - **改了哪里** ← 变更对照里真实读回的前后差异（`RangeChangeReview`），
 *    不是模型叙述的"我修改了 A 列"；
 *  - **验证结果** ← 独立验收回执（`VerificationReport`），没有回执就什么都不说；
 *  - **下一步操作** ← 当前所处的阶段（审阅 / 待核对 / 可续跑 / 已撤销…）。
 *
 * 为什么必须这样：这三句是用户唯一会认真读的结论。让模型"总结一下刚才做了什么"，
 * 就会得到"已完成 3 处修改""本次零写入"这类听起来确定、实际没有证据的句子。
 * 本模块的每条输出都能指回一个字段；指不回去的就不输出。
 *
 * 三条不变量：
 *  A. **进行中不出结论。** `live` 时返回 `null`，过程中只由过程框说"在做什么"。
 *  B. **没有证据就不说。** 没有验收回执 ⇒ 不给"验证结果"这一行（不是给"已验证"）；
 *     没有变更对照 ⇒ 不给"改了哪里"（不是给"零写入"）。
 *  C. **撤销/放弃是终态。** 二者只描述恢复，不复述旧的验证结论 —— 旧结论证明的是
 *     那次写入当时的结果，而此刻单元格已经被还原。
 */

/** 变更对照里本模块**只用**的字段；两端的完整类型可以更宽。 */
export interface ProcessOutcomeReviewLike {
  sheet: string;
  range: string;
  /** `applied` / `executed_unverified` 才算"真的改了"；其余是没执行或读数失败。 */
  status: string;
  rows?: readonly { changed?: boolean }[];
}

/** 独立验收回执里本模块**只用**的字段。 */
export interface ProcessOutcomeVerificationLike {
  status: "verified" | "executed_unverified" | "failed";
  total: number;
  passed: number;
  unverified: number;
}

/**
 * 当前阶段。刻意用**语义名**而不是各端的 phase 字符串：
 * 两端内部枚举不同，但"用户在等什么"只有这几种。
 */
export type ProcessOutcomePhase =
  | "executing"
  | "reviewing"
  | "confirmed"
  | "discarded"
  | "undone"
  | "needs_review"
  | "resumable"
  | "blocked"
  | "idle";

export interface ProcessOutcomeInput {
  /** 还在跑：不出结论。 */
  live: boolean;
  phase: ProcessOutcomePhase;
  reviews?: readonly ProcessOutcomeReviewLike[] | null;
  verification?: ProcessOutcomeVerificationLike | null;
  /**
   * 输入区上方**此刻**还摆着一件恢复待办（待核对 / 待续跑），与这条消息本身无关。
   *
   * 存在理由（2026-11-01）：`phase` 是按"这条消息自己那次执行"折算的。跨重启、
   * 换会话或经发现流程找回来的恢复卡，往往对不上任何一条消息的 `officeResume` ——
   * 于是消息里写「本次没有需要你处理的后续操作」，而输入框上方正摆着「核对结果」。
   * 同一屏里两句话互相打脸。
   *
   * 所以这里多给一条**全局**事实：只要它成立，`idle` 就不许输出"没有需要处理的
   * 后续操作"，改成指路那句。这不是新判据，只是让"没有待办"这句话在还有待办时闭嘴。
   */
  pendingRecovery?: boolean;
  /**
   * 已经由调用方渲染好的一句验收结论。
   *
   * 存在的理由：两端的结果回执形状不同 —— Excel 是"N 项检查里通过了多少"，
   * WPS 是"回读通过 / 计算值未获证实 / 部分写入"这种档位，没有项数可数。
   * 硬把它们塞进同一个计数模型就必须编一个分母，那比不显示更糟。
   * 所以 WPS 传自己**已有**的那句结论（`WPS_EXECUTION_OUTCOME_LABEL`），
   * 本函数在 `verification` 缺席时才用它。
   */
  verificationNote?: string | null;
}

export interface ProcessOutcomeSummary {
  /** 改了哪里；`null` = 没有可用证据，**不**等于"没改"。 */
  changedWhere: string | null;
  /** 验证结果；`null` = 没有验收回执。 */
  verification: string | null;
  /** 下一步操作。永远有一句，否则用户不知道现在该做什么。 */
  nextStep: string;
}

/** 摘要里最多列出几处改动位置；超出只给数量。 */
const MAX_LISTED_TARGETS = 3;

/** 只有这两种状态能说"这次真的写进去了"。 */
const APPLIED_STATUSES: ReadonlySet<string> = new Set([
  "applied",
  "executed_unverified"
]);

function changedCellCount(review: ProcessOutcomeReviewLike): number {
  if (!review.rows) return 0;
  return review.rows.filter((row) => row.changed === true).length;
}

/**
 * 「改了哪里」。
 *
 * 只认 `applied` / `executed_unverified` 且确实有格子发生变化的对照 —— 一次
 * "读到执行前快照但没写"的对照（`not_executed`）或读数失败（`readback_failed`）
 * 都不该被算成"改了这里"。`readback_failed` 尤其不能：它连写没写成都不知道。
 */
export function describeChangedWhere(
  reviews: readonly ProcessOutcomeReviewLike[] | null | undefined
): string | null {
  if (!reviews || reviews.length === 0) return null;
  const applied = reviews.filter(
    (review) => APPLIED_STATUSES.has(review.status) && changedCellCount(review) > 0
  );
  if (applied.length === 0) return null;
  const labels = applied.map((review) => {
    const range = review.range?.trim();
    const sheet = review.sheet?.trim();
    const where = sheet && range ? `${sheet}!${range}` : sheet || range || "未知区域";
    return `${where}（${changedCellCount(review)} 格）`;
  });
  if (labels.length <= MAX_LISTED_TARGETS) return `改了：${labels.join("、")}`;
  return `改了 ${applied.length} 处：${labels.slice(0, MAX_LISTED_TARGETS).join("、")} 等`;
}

/** 「验证结果」。没有回执就是 `null` —— 不拿"执行完成"冒充"验证通过"。 */
export function describeVerification(
  verification: ProcessOutcomeVerificationLike | null | undefined
): string | null {
  if (!verification) return null;
  if (verification.status === "verified") {
    return `${verification.passed}/${verification.total} 项符合预期`;
  }
  if (verification.status === "executed_unverified") {
    return verification.unverified > 0
      ? `已执行；${verification.unverified} 步缺少独立验收`
      : "已执行；缺少独立验收";
  }
  return `${verification.total - verification.passed}/${verification.total} 项不符合预期`;
}

/** 「下一步操作」。每一种阶段都必须给一句人话，不能空着。 */
export function describeNextStep(phase: ProcessOutcomePhase): string {
  switch (phase) {
    case "reviewing":
      return "在主表核对结果后点「确认变更」或「放弃变更」。";
    case "confirmed":
      return "已确认变更；确认不等于保存文件，需要落盘请在宿主里保存。";
    case "discarded":
      return "已按恢复点还原主表，本次没有保留任何写入。";
    case "undone":
      return "本次执行已撤销；如需重做请重新发送需求。";
    case "needs_review":
      /*
       * 不写死按钮名（2026-11-02）：主操作按证据分档（「检查已写入结果」/
       * 「检查写入结果」/「检查当前结果」），写死一个名字就会指到一颗不存在的按钮上。
       * 所以照实说"在输入框里"，让人自己看见那颗按钮。
       */
      return "这次执行需要先核对：在输入框里点核对按钮（只读，不写单元格），再决定继续或撤销。";
    case "resumable":
      /* 同理：续跑那档的按钮带剩余步数（「继续剩余 N 步」），名字不固定。 */
      return "在输入框里点继续按钮执行尚未完成的步骤；已完成的部分不会重跑。";
    case "blocked":
      return "当前不能继续；按上方提示处理后重新核对。";
    case "executing":
      return "正在执行；结束后可在此核对结果。";
    case "idle":
    default:
      return "本次没有需要你处理的后续操作。";
  }
}

/**
 * 还有恢复待办时替代 `idle` 那句的指路话（唯一出处）。
 *
 * 只在这条消息自己**说不出**下一步、而输入区上方确实还摆着一件事时使用：
 * 那句话必须说"去哪处理"，而不是硬把别人的待办算到这次执行头上。
 */
export const PENDING_RECOVERY_NEXT_STEP = "还有一次写入没有收尾；处理入口在输入框里。";

/**
 * 把两端各自的内部枚举折成**语义阶段**（`ProcessOutcomePhase`）。
 *
 * 只有一个入口，是因为"下一步该点哪个按钮"这句话必须和真实入口一致：
 * 审阅中不能写"点继续剩余动作"，可续跑也不能写"点确认变更"。两端内部
 * 各有一套 phase 字符串与恢复结论枚举，折一次、共用，就不会出现某端
 * 说着一句只有另一端才有的按钮的话。
 *
 * 优先级按"离用户当前动作最近"排：撤销/放弃是终态，压过一切；其次是
 * 正在进行的审阅；最后才是恢复结论。
 */
export function processOutcomePhaseFrom(input: {
  undone?: boolean;
  abandoned?: boolean;
  /** 结果审阅阶段（`resultReview` 的 phase）。 */
  reviewPhase?: string | null;
  /** 恢复结论（`executionRecoveryView` 的 verdict.kind）。 */
  recoveryKind?: string | null;
}): ProcessOutcomePhase {
  if (input.undone) return "undone";
  if (input.abandoned) return "discarded";
  if (input.reviewPhase === "reviewing") return "reviewing";
  if (input.reviewPhase === "confirmed") return "confirmed";
  if (input.reviewPhase === "discarded") return "discarded";
  switch (input.recoveryKind) {
    case "undone":
      return "undone";
    case "resume":
      return "resumable";
    case "reconcile":
      /* 在途动作还没核对清楚 ⇒ 先核对，不能直接续跑。 */
      return "needs_review";
    case "completed":
      /* 动作都做完了、只缺回执 ⇒ 需要补记（它自己不写单元格）。 */
      return "needs_review";
    case "replan":
    case "blocked":
    case "insufficient":
      return "blocked";
    default:
      return "idle";
  }
}

/**
 * 三段摘要一次派生。`live` 时返回 `null`：进行中只由过程框描述"在做什么"，
 * 提前给结论就是在替还没有的结果下判断。
 */
export function processOutcomeSummary(
  input: ProcessOutcomeInput
): ProcessOutcomeSummary | null {
  if (input.live) return null;
  const terminal = input.phase === "undone" || input.phase === "discarded";
  const summary: ProcessOutcomeSummary = {
    /*
     * 撤销/放弃之后不再展示"改了哪里"：那块区域此刻已经是执行前的样子，
     * 继续写"改了 X"会让人以为改动还在。只说明恢复事实。
     */
    changedWhere: terminal
      ? input.phase === "undone"
        ? "已按恢复点还原本次写入的内容与格式。"
        : "已放弃变更并按恢复点还原。"
      : describeChangedWhere(input.reviews),
    verification: terminal
      ? null
      : describeVerification(input.verification) ??
        (input.verificationNote?.trim() ? input.verificationNote.trim() : null),
    nextStep:
      input.phase === "idle" && input.pendingRecovery
        ? PENDING_RECOVERY_NEXT_STEP
        : describeNextStep(input.phase)
  };
  // 普通聊天没有结果回执或待办，不额外输出一条空的操作提示。
  // 例外：输入区上方还摆着恢复待办时，"没有待办"这句话不成立，必须给指路那句。
  if (
    input.phase === "idle" &&
    !summary.changedWhere &&
    !summary.verification &&
    !input.pendingRecovery
  ) {
    return null;
  }
  return summary;
}
