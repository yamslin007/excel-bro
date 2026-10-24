/**
 * 「继续分析」之前必须交出的现场证据。
 *
 * 为什么需要这一层：检查点只说明"服务端存过一份可续的现场"，它**不是**"此刻的
 * 工作簿与授权还是那份"的证据。等待网络、切换文件、撤销勾选、改数据都发生在
 * 服务端看不见的地方，所以客户端必须在点下去的那一刻重新读一次宿主，并把读到
 * 的东西随请求交上去。
 *
 * 三条硬纪律，两端共用**这一份**实现：
 *
 * 1. **不补足**。工作簿身份读不到就是读不到，不许退到文件名、旧快照或上次的
 *    记录值。`name:<名字>` 这类兜底身份在检查点契约里本来就不合法。
 * 2. **不混态**。"尚未采集""读取失败""漏传"是三件事：前者在纯结构任务里合法，
 *    后两者必须停下来。把它们合并成一个"没证据"会让"读不到"被当成"不需要读"。
 * 3. **来源按依赖判断**。原任务依赖来源数据时必须重读并比对；确实没读过来源的
 *    结构任务不该被凭空要求读数据 —— 那会把合法任务挡在门外。
 */

/** 来源核实的三个状态，与协议里的 `sourceVerification` 一一对应。 */
export type PlanningSourceVerification = "verified" | "not_collected" | "failed";

export interface PlanningContinuationEvidence {
  workbookIdentity: string;
  authorizedSheetNames: string[];
  sourceVerification: PlanningSourceVerification;
  /** 仅 `sourceVerification === "verified"` 时存在。 */
  sourceFingerprint?: string;
}

/**
 * 拿不到证据的三种原因。三者都是**停止继续**，但文案不同：
 * 用户该做的事分别是"换个能识别身份的文件""恢复至少一张授权表""重新读取一次"。
 */
export type PlanningEvidenceFailure =
  | "identity_unavailable"
  | "scope_unavailable"
  | "source_unreadable";

export type PlanningEvidenceInput = {
  /** 宿主此刻给出的工作簿身份；读不到就是 null/undefined，**不要**在这之前兜底。 */
  workbookIdentity: string | null | undefined;
  /** 当前已提交授权的工作表名（用户勾选并确认过的那一组）。 */
  authorizedSheetNames: readonly string[];
  /**
   * 原任务是否依赖来源数据。
   *
   * 由**服务端**判定（检查点里有没有来源指纹 / 来源表 / 结果引用），客户端不自己猜：
   * 猜错的两种方向都危险 —— 猜"需要"会挡住纯结构任务，猜"不需要"会放过改了数据的现场。
   */
  sourceRequired: boolean;
  /** 重读来源指纹。抛错或返回空都算**读取失败**，不是"尚未采集"。 */
  readSourceFingerprint: () => string | null | undefined | Promise<string | null | undefined>;
};

export type PlanningEvidenceResult =
  | { ok: true; evidence: PlanningContinuationEvidence }
  | { ok: false; failure: PlanningEvidenceFailure };

function cleanName(name: string): string {
  return typeof name === "string" ? name.trim() : "";
}

export async function buildPlanningContinuationEvidence(
  input: PlanningEvidenceInput
): Promise<PlanningEvidenceResult> {
  const identity = (input.workbookIdentity ?? "").trim();
  if (!identity) {
    // 硬停：不拿文件名、旧快照或记录里的值补一个"看起来能用"的身份。
    return { ok: false, failure: "identity_unavailable" };
  }
  const names = [...new Set(input.authorizedSheetNames.map(cleanName).filter(Boolean))];
  if (names.length === 0) {
    return { ok: false, failure: "scope_unavailable" };
  }
  if (!input.sourceRequired) {
    return {
      ok: true,
      evidence: {
        workbookIdentity: identity,
        authorizedSheetNames: names,
        sourceVerification: "not_collected"
      }
    };
  }
  let fingerprint: string | null | undefined;
  try {
    fingerprint = await input.readSourceFingerprint();
  } catch {
    // 读取失败**不能**退化成"尚未采集"：那等于用一次异常换一张通行证。
    return { ok: false, failure: "source_unreadable" };
  }
  const value = (fingerprint ?? "").trim();
  if (!value) {
    return { ok: false, failure: "source_unreadable" };
  }
  return {
    ok: true,
    evidence: {
      workbookIdentity: identity,
      authorizedSheetNames: names,
      sourceVerification: "verified",
      sourceFingerprint: value
    }
  };
}

/** 三种失败的固定文案。两端共用，避免一边说"读不到"一边说"没有检查点"。 */
export const PLANNING_EVIDENCE_FAILURE_MESSAGES: Record<PlanningEvidenceFailure, string> = {
  identity_unavailable:
    "读不到当前工作簿的身份（未保存的文件无法核对），已停止本次继续；原需求仍保留。",
  scope_unavailable:
    "当前没有已确认的授权工作表，已停止本次继续；请重新勾选数据范围后再继续。",
  source_unreadable:
    "重新读取当前数据时失败（或没有读到内容），无法确认数据是否仍是原任务那一份，已停止本次继续。"
};
