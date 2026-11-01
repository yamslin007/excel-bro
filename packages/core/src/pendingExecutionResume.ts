/**
 * 「能不能把一条未完成执行摆成可续跑的卡」的共享判据（宿主无关）。
 *
 * 存在理由（2026-10-29）：Excel 端此前只有**一条** `localStorage` 索引把用户带回断点，
 * 索引不在（换机器、清缓存、窗格重开）时就再也看不到这次执行。补发现入口的同时必须
 * 把"发现到了 ≠ 可以续跑"说清楚，否则新入口会把一张"能继续"的卡递给用户，而实际
 * 身份不符、工作表已收窄、在途动作没有写前快照。
 *
 * 三条纪律：
 *
 * - **只判不写**：这里没有任何网络、宿主或账本调用，输入是纯数据，输出是纯结论；
 *   认领、补记、续跑仍走各自宿主的既有路径。
 * - **授权按"登记过的那一份 + 用户这次真的确认过的那一份"判**：
 *   计划要写的每张表必须在**登记授权**里（否则是记录自相矛盾），且此刻仍存在于工作簿；
 *   如果用户**在本次会话里确认过范围**，还必须在那个已确认范围里 —— 用户把表收窄后，
 *   卡面必须跟着变红，而不是停在按旧范围算出的"可继续"。
 *   刻意**不**拿"宿主的当前勾选"去比对：Excel 窗格重新挂载时会把已提交范围播种成
 *   "当前活动表"，那不是用户的手势，拿它比对只会制造假阳性（真机上表现为"什么都没动，
 *   却说你把表取消了"）。所以只有调用方明确给出"这次真的确认过"的集合才比对，
 *   给不出就是 `null` = 不比对。
 * - **拿不准就阻断**：身份读不出、记录自相矛盾、缺写前快照，一律 fail-closed。
 *
 * 与 WPS 的关系：WPS 端的对应判定含**宿主材料级**检查（`prepareWpsResume` 校验动作
 * 描述与写前指纹），目前仍在 `apps/wps-addin/src/wpsExecutionResume.ts` 里，**没有**
 * 接到本模块上。所以本轮只能说"Excel 用了共享判据"，不宣称两端已经统一。
 */

/** 阻断原因码。每个码只在**一处**返回点上产生，界面按码给文案，不按字符串猜。 */
export type PendingResumeBlockReason =
  | "identity_unavailable"
  | "identity_mismatch"
  | "completed"
  | "not_resumable"
  | "plan_target_never_authorized"
  | "plan_target_missing"
  | "plan_target_not_authorized_now"
  | "in_flight_without_snapshot";

export type PendingResumeReadiness =
  | { ready: true; requiresReconciliation: boolean }
  | {
    ready: false;
    reason: PendingResumeBlockReason;
    /** 用户可读的一句话；点名到具体工作表，方便用户回去重新勾选。 */
    detail: string;
    /** 与原因相关的工作表名（只有授权/存在性两类原因会带）。 */
    sheets?: string[];
  };

export interface PendingResumeInput {
  /** 此刻宿主读到的稳定身份；读不出来是 `null`（不是空串、不是文件名）。 */
  currentIdentity: string | null;
  /** 账本登记的身份。 */
  recordedIdentity: string;
  status: "running" | "needs_reconciliation" | "completed" | "failed";
  /** 服务端是否允许续跑（执行权已释放等）。 */
  canResume: boolean;
  /** 在途动作位置；`null` 表示没有卡在写一半的动作。 */
  inFlightActionIndex: number | null;
  /** 在途动作是否带写前快照（没有就只能人工核对）。 */
  hasBeforeSnapshot: boolean;
  /** 登记时的授权工作表名。 */
  recordedAuthorizedSheets: readonly string[];
  /**
   * **用户在本会话里确认过**的授权工作表名；`null` = 没有可比对的手势（不比对）。
   *
   * 传 `null` 不等于"一张都没授权" —— 那会把"刚打开窗格"说成"你把表取消了"。
   */
  currentAuthorizedSheets: readonly string[] | null;
  /** 计划要写的工作表名（去重后由 `pendingResumeTargetSheets` 给出）。 */
  planTargetSheets: readonly string[];
  /** 工作簿当前**存在**的工作表名（可选；给了才能区分"被删了"与"被取消授权"）。 */
  workbookSheetNames?: readonly string[] | null;
}

/** 计划要写的表名（去空、去重、保序；不做名字归一化 —— 表名就是表名）。 */
export function pendingResumeTargetSheets(plan: {
  actions: readonly { sheet?: string | null }[];
}): string[] {
  const seen = new Set<string>();
  const names: string[] = [];
  for (const action of plan.actions) {
    const name = typeof action.sheet === "string" ? action.sheet.trim() : "";
    if (!name || seen.has(name)) continue;
    seen.add(name);
    names.push(name);
  }
  return names;
}

/**
 * 结论。顺序不是随意排的：身份 → 状态 → 授权 → 在途材料。
 *
 * 身份放最前，是因为"另一个工作簿"会让后面每一条都失去意义（同名异路径的两个簿
 * 最容易在这里串台）。状态放第二，是因为已完成记录不该被当成"待续写"再摆一次卡。
 */
export function pendingResumeReadiness(input: PendingResumeInput): PendingResumeReadiness {
  if (!input.currentIdentity) {
    return {
      ready: false,
      reason: "identity_unavailable",
      detail: "当前工作簿尚未保存，无法确认它就是原工作簿。请先保存并重新打开原文件。"
    };
  }
  if (input.currentIdentity !== input.recordedIdentity) {
    return {
      ready: false,
      reason: "identity_mismatch",
      detail: "当前不是产生该断点的工作簿。请打开原工作簿后再继续。"
    };
  }
  if (input.status === "completed") {
    return {
      ready: false,
      reason: "completed",
      detail: "服务端记录这次写入已经完成，不需要续跑；请在结果卡核对已写入的内容。"
    };
  }
  if (!input.canResume) {
    return {
      ready: false,
      reason: "not_resumable",
      detail: "服务端尚未释放这次写入的执行权，或这条记录不能继续；不要重复登记新的写入。"
    };
  }
  const targets = [...new Set(input.planTargetSheets)];
  const recordedAuthorized = input.recordedAuthorizedSheets;
  const neverAuthorized = targets.filter((name) => !recordedAuthorized.includes(name));
  if (neverAuthorized.length > 0) {
    return {
      ready: false,
      reason: "plan_target_never_authorized",
      detail: `这次写入的记录里没有授权工作表「${neverAuthorized.join("」「")}」，不能按原计划续跑。`,
      sheets: neverAuthorized
    };
  }
  const authorizedNow = input.currentAuthorizedSheets;
  if (authorizedNow) {
    const notAuthorizedNow = targets.filter((name) => !authorizedNow.includes(name));
    if (notAuthorizedNow.length > 0) {
      return {
        ready: false,
        reason: "plan_target_not_authorized_now",
        detail: `当前授权里没有工作表「${notAuthorizedNow.join("」「")}」；请重新勾选后再核对，不会自行扩大授权。`,
        sheets: notAuthorizedNow
      };
    }
  }
  const workbookSheets = input.workbookSheetNames;
  if (workbookSheets) {
    const missing = targets.filter((name) => !workbookSheets.includes(name));
    if (missing.length > 0) {
      return {
        ready: false,
        reason: "plan_target_missing",
        detail: `工作表「${missing.join("」「")}」已不在当前工作簿里，无法按原计划续跑。`,
        sheets: missing
      };
    }
  }
  if (input.inFlightActionIndex != null && !input.hasBeforeSnapshot) {
    return {
      ready: false,
      reason: "in_flight_without_snapshot",
      detail: "在途动作缺少写前快照，只能人工核对，不能自动续跑。"
    };
  }
  return {
    ready: true,
    requiresReconciliation:
      input.inFlightActionIndex != null || input.status === "needs_reconciliation"
  };
}
