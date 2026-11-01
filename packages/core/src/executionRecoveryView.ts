/**
 * 「上次写入没收尾」的恢复结论 —— **宿主无关**的呈现契约。
 *
 * 存在理由（2026-10-29）：WPS 侧早已把结论分成六种（已完成 / 待核对在途动作 /
 * 可继续剩余动作 / 需重新规划 / 需要人工核对 / 证据不足），Excel 侧却只有一个
 * `blocked | pending` 两态，并把「核对在途动作」和「执行剩余动作」塞进同一次点击。
 * 于是同一件事在两端是两个说法：WPS 上"第 2 步已写入、只等回执"，在 Excel 上被
 * 描述成"核对并继续"——用户既看不出第 3 步还没跑，也不知道这一下会不会再写一次。
 *
 * 所以本模块只放**结论与措辞**，不放任何判定输入：
 *
 * - 判定（读服务端状态、读宿主现场、比对本机材料）留在各宿主的适配层：
 *   WPS 是 `wpsExecutionReconcile.ts`，Excel 是 `officeRecoveryVerdictFor`。
 * - 执行（claim / 补记回执 / 执行剩余动作）也留在各宿主：两端 API 与写前快照
 *   的形状不同，硬合并只会让其中一端失去自己的 fail-closed 判据。
 *
 * 三条纪律：
 *
 * - **一个结论一个动作集合**：`executionRecoveryActions` 是按钮的**唯一**来源。
 *   结论说证据不足就一个写入/补记按钮都不给 —— 画一个点不动的按钮等于骗人。
 * - **核对与续跑不同结论**：`reconcile` 只补记回执（不写单元格），`resume` 才执行
 *   剩余动作。两者互斥，不许在同一个按钮里先补记再续跑。
 * - **同一句话只有一个出处**：文案在这里，宿主只按引用转发。
 */

/** 恢复卡允许出现的动作。返回空数组 = 不给任何按钮。 */
export type ExecutionRecoveryAction =
  | "backfill"
  | "reconcile"
  | "replan"
  | "resume"
  | "dismiss";

/** 在途动作的现场结论；`pending` = 还没核对过（只出现在"待核对"这一步之前）。 */
export type ExecutionReconcileOutcome = "applied" | "not_applied" | "pending";

/**
 * 「剩余动作能不能沿用旧结果继续」里**来源**那一环的结论。
 *
 * 存在理由（2026-10-31 真机）：来源变化时续跑被正确拒绝，可**重开之后**恢复卡却写
 * “来源和目标已核对，可继续剩余动作”—— 因为发现阶段只读账本、根本没读过来源表，
 * 而那句正文把“来源已核对”写成了既成事实。判据（`verifyFormulaResumeSources`）一直
 * 在，只是没人在**出结论的地方**问过它。
 *
 * 所以把这一环单独表达出来，让“没复核”与“复核过且一致”**不能混为一谈**：
 *
 * - `matched` / `not-required` 才允许出 `resume`（前者真的比过指纹，后者这份计划
 *   压根不读别的表，没有来源可核）；
 * - `unverified` 只能出“待复核”，并给一次**只读**复核 —— 它不是错误，是“还没问”；
 * - `refused` 是复核之后的**事实**（来源变了 / 没登记证据 / 读不到），`reason` 直接
 *   用共享判据（`formulaWriteback.ts`）给的原文，不另写一套措辞。
 *
 * 结论**不缓存成"永久事实"**：每次复核都按当时的现场与登记基线重新算（见宿主侧
 * `officeSourceVerdictFor`），所以用户把来源改回登记时的值，卡面会自己变回可继续。
 */
export type ExecutionRecoverySourceVerdict =
  | { state: "matched" }
  | { state: "not-required" }
  | { state: "unverified" }
  | { state: "refused"; reason: string };

/**
 * 恢复结论。
 *
 * `Outcome` 是给宿主收窄用的：WPS 的判定一定会现场回读，只会产出
 * `applied | not_applied`；Excel 的发现阶段只做只读状态查询，点核对之前是 `pending`。
 */
export type ExecutionRecoveryVerdict<
  Outcome extends ExecutionReconcileOutcome = ExecutionReconcileOutcome
> =
  | {
    kind: "completed";
    recorded?: true;
  }
  | {
    kind: "reconcile";
    actionIndex: number;
    outcome: Outcome;
    /** 在途动作**之后**还有多少个动作没跑。核对不会代跑它们，卡片要写清楚。 */
    remaining?: number;
  }
  | { kind: "replan"; remaining: number; reason?: string }
  /**
   * `source` 是这句正文的**依据**：`matched` 才允许说“来源已核对”。缺省表示
   * “这次没有来源可核”（例如计划只写固定值），正文也**不**声称比过来源。
   */
  | { kind: "resume"; remaining: number; source?: "matched" | "not-required" }
  /**
   * `sceneChanged` 表示"这次执行写入过的目标已经不再是写后状态"（例如已被撤销
   * 清空、或被人改过）。它同时是续跑和撤销**两个入口**的失效原因：目标不在写后
   * 状态，既不能继续，也不能再"撤销回退"（因为已经回退过了）。宿主据此把撤销行
   * 也一并收走，避免"卡面说阻断、撤销按钮却还亮着"。
   */
  | { kind: "blocked"; detail: string; sceneChanged?: true; abandoned?: true }
  | { kind: "insufficient"; detail: string }
  /**
   * 终态：这次执行已经被撤销（本机确认恢复成功，或账本已留处置痕）。
   *
   * 与 `blocked` 分开的理由：撤销之后**没有第二个动作可做** —— 不给继续、不给补记、
   * 也不给“再撤一次”。用一个专门的结论表达，卡片的 badge 才能是“已撤销”，而不是
   * 含糊的“需要人工核对”；`executionRecoveryActions` 也能干净地返回空数组。
   *
   * `detail` 区分两种来源，默认文案 `EXECUTION_RECOVERY_UNDONE_DETAIL` 说的是“本机
   * 撤销成功、写入已回到执行前”。但“账本已留处置痕”（例如**另一台机器**已经撤过）
   * 时本机并没有真的把写入回退，表格现场未知，不能说“已回到执行前”，宿主应传自己的
   * 措辞（“这次执行已经撤销过（账本已留处置痕）…”），`executionRecoveryDetail` 会优先用它。
   */
  | { kind: "undone"; detail?: string };

export type ExecutionRecoveryKind = ExecutionRecoveryVerdict["kind"];

/** 结论的状态名（卡片 badge / 标题级）。 */
export const EXECUTION_RECOVERY_STATE_LABEL: Record<ExecutionRecoveryKind, string> = {
  completed: "已完成，仅缺回执",
  reconcile: "待核对在途动作",
  replan: "需重新规划",
  resume: "可继续剩余动作",
  blocked: "需要人工核对",
  insufficient: "证据不足",
  undone: "已撤销本次执行"
};

/** 动作按钮的文案。 */
export const EXECUTION_RECOVERY_ACTION_LABEL: Record<ExecutionRecoveryAction, string> = {
  resume: "继续剩余动作",
  backfill: "补记完成（不重写原表）",
  reconcile: "核对在途动作",
  replan: "重新读取并规划",
  dismiss: "清除这条恢复提示"
};

/** 恢复卡标题（两端同一句）。 */
export const EXECUTION_RECOVERY_TITLE = "上次写入的恢复核对";

/** 恢复卡的固定边界说明（两端同一句）。 */
export const EXECUTION_RECOVERY_NOTE =
  "恢复只使用当前已授权工作表；续跑通过复核后执行剩余动作，再审阅结果。";

/**
 * 按**结论**派生卡片上的边界说明（`EXECUTION_RECOVERY_NOTE` 的按状态版本）。
 *
 * 存在理由（2026-10-31 真机）：`EXECUTION_RECOVERY_NOTE` 是一句写死的"续跑引导"，
 * 卡片无论什么结论都原样显示。于是已撤销的终态卡上，"已撤销本次执行"下面还跟着
 * "续跑通过复核后执行剩余动作，再审阅结果" —— 与终态自相矛盾（撤销之后**没有**
 * 续跑，也没有审阅）。终态说明必须**替换**这句引导，而不是在它后面再补一句。
 *
 * 只有 `undone` 走终态说明，其余结论原句照旧：这不是放宽任何判据，只是把"下一步
 * 要做什么"的说明和"这次已经结束了"的事实对齐。
 */
export const EXECUTION_RECOVERY_UNDONE_NOTE =
  "这次执行已撤销，不会再提供补记或审阅入口；来源工作表保持你现在的修改。";

export function executionRecoveryNote(verdict: ExecutionRecoveryVerdict): string {
  return verdict.kind === "undone" ? EXECUTION_RECOVERY_UNDONE_NOTE : EXECUTION_RECOVERY_NOTE;
}

/**
 * 「当前这张卡放在哪儿」—— 入口**只有一个**。
 *
 * 存在理由（2026-10-29 真机）：Excel 的恢复卡原本画在**产生它的那条消息**下面。
 * 点「核对在途动作」会追加一行日志，`useMessageAutoFollow` 于是把消息区滚到底，
 * 而已经换成「继续剩余动作」的那张卡仍停在旧消息的位置 —— 用户看到的是
 * "已核对，未写单元格"，下一步入口却已被顶出视口。
 *
 * 修法不是"再滚一次"（那样上翻历史就会被打断），而是把**可点击的那一份**从
 * 消息流里搬到输入区上方那块稳定区域；消息里只留只读的执行记录。于是"到底有
 * 几个入口"变成一个可以被断言的问题，判据就是下面这个选择器：
 *
 * - **取最后一条**带卡的记录：更新的执行覆盖更旧的，绝不把两次执行合成一张卡
 *   （合并就会出现"卡上的按钮属于另一次执行"）。
 * - 没有 `operationId` 的卡直接跳过：绑不上具体执行就不能给按钮。
 * - 返回 `null` = 这次一条卡都不该画（宿主必须什么都不渲染，而不是画个空壳）。
 *
 * 宿主只做两件事：按它自己的字段取卡（`recoveryOf`）、把返回的那一条渲染到
 * 稳定位置。判定与文案仍各归各家。
 */
export function selectExecutionRecoveryEntry<T, C extends { operationId: string }>(
  rows: readonly T[],
  recoveryOf: (row: T) => C | null | undefined
): { index: number; row: T; card: C; operationId: string } | null {
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const card = recoveryOf(rows[index]);
    if (!card) continue;
    const operationId = typeof card.operationId === "string" ? card.operationId.trim() : "";
    if (!operationId) continue;
    return { index, row: rows[index], card, operationId };
  }
  return null;
}

/**
 * 执行记录的进度一句话：恢复卡与消息里的只读记录**共用同一句**。
 *
 * 以前这句话只在卡片里拼，消息那边没有；于是"第几步"的事实要么两处各写一份、
 * 要么一处根本不写。进度是详情（可折叠），结论在 badge 与正文里。
 */
export function executionRecoveryProgress(
  verdict: ExecutionRecoveryVerdict,
  progress: { nextActionIndex: number; totalActions: number }
): string {
  const head = `已完成 ${progress.nextActionIndex} / ${progress.totalActions} 步`;
  if (verdict.kind === "reconcile") return `${head} · 在途第 ${verdict.actionIndex + 1} 步`;
  if (verdict.kind === "resume") return `${head} · 剩余 ${verdict.remaining} 步`;
  if (verdict.kind === "replan") return `${head} · 剩余 ${verdict.remaining} 步待重新规划`;
  return head;
}

/** 历史消息里的只读记录要指向**唯一**入口，否则用户会以为按钮丢了。 */
export const EXECUTION_RECOVERY_ENTRY_HINT = "操作入口在输入框上方，这条消息只保留执行记录。";

/** badge 文案：`completed` 且账本本来就有完成结果时，别说成"仅缺回执"。 */
export function executionRecoveryBadge(verdict: ExecutionRecoveryVerdict): string {
  if (verdict.kind === "completed" && verdict.recorded) return "已记录完成";
  return EXECUTION_RECOVERY_STATE_LABEL[verdict.kind];
}

/**
 * 「已撤销」的正文（**唯一**出处）。
 *
 * 撤销是终态，这句话必须一次说清三件事，否则用户会去找一个不存在的下一步：
 * ① 写入已经回到执行前；② 来源表**没有**被写回（用户的修改保留）；
 * ③ 不会再执行这次计划的任何动作。
 */
export const EXECUTION_RECOVERY_UNDONE_DETAIL =
  "这次执行已经撤销：本次写入已回到执行前，来源工作表保持你现在的修改；不会重做任何写入，也不再提供继续或补记。";

/**
 * 「还有动作没跑」的结论 —— **来源复核结论决定能不能给出续跑按钮**。
 *
 * 存在理由（2026-10-31 真机）：`resume` 的正文曾经写死“原计划与当前授权、来源和目标
 * 已核对”，而结论只由账本 + 授权算出来，来源一次都没读过。于是重开之后卡片说
 * “可继续剩余动作”，点下去却被 `verifyFormulaResumeSources` 拒绝 —— 显示与守卫
 * 各说一套。这里把两者绑在一起：**没复核 ⇒ 不出 `resume`**。
 *
 * 三种“还不能继续”的正文分别对应三种处置，不合并成一句“不能续跑”：
 * 待复核（去点只读复核）／来源变了（把来源改回去再复核）／没有证据（这条计划本就
 * 不支持精确续跑）。三者都**不**给写入按钮 —— 画一个点了只会报错的按钮等于骗人。
 */
export function executionRecoveryResumeVerdict(
  remaining: number,
  source: ExecutionRecoverySourceVerdict
): ExecutionRecoveryVerdict {
  if (remaining <= 0) return { kind: "completed" };
  if (source.state === "matched") return { kind: "resume", remaining, source: "matched" };
  if (source.state === "not-required") {
    return { kind: "resume", remaining, source: "not-required" };
  }
  if (source.state === "unverified") {
    return {
      kind: "blocked",
      detail:
        `还有 ${remaining} 个动作没跑；这份计划的公式要读其他工作表作为来源，` +
        "重开或刷新后还没有和登记时的来源核对过，所以现在不能继续。" +
        "点「重新核对（只读）」先比来源，确认一致后才会给出续跑；核对不写任何单元格。"
    };
  }
  return {
    kind: "blocked",
    detail:
      `${source.reason}还有 ${remaining} 个动作没跑，不能沿用旧结果继续。` +
      "把来源改回登记时的值之后，点「重新核对（只读）」会按当时的现场重新判断，不会自动重新规划，也不会放宽。"
  };
}

/**
 * 结论的一句话正文。冲突与证据不足直接复用判定给出的 `detail`，不另写一套措辞。
 */
export function executionRecoveryDetail(verdict: ExecutionRecoveryVerdict): string {
  switch (verdict.kind) {
    case "completed":
      return verdict.recorded
        ? "账本已保存历史完成结果，无需重复写入；不代表当前单元格仍未被修改。"
        : "动作与当前现场已核对，缺的只是最终回执；补记只改账本状态，不会重写原表。";
    case "reconcile": {
      const pending = typeof verdict.remaining === "number" && verdict.remaining > 0
        ? `还有 ${verdict.remaining} 个动作没执行；核对只补记这一个，不会代跑它们。`
        : "";
      if (verdict.outcome === "applied") {
        return `第 ${verdict.actionIndex + 1} 个动作的现场与写后一致，可补记为已写入，不重复写。${pending}`;
      }
      if (verdict.outcome === "not_applied") {
        return `第 ${verdict.actionIndex + 1} 个动作的现场仍是写前状态，可补记为未写入。${pending}`;
      }
      return `第 ${verdict.actionIndex + 1} 个动作可能在停止前已经写入，也可能没有；核对只读当前目标与写前快照，不写任何单元格。${pending}`;
    }
    case "resume":
      /*
       * 这句话只在 `source` 真的**比过**之后才提来源。
       *
       * 旧写法把“来源已核对”写死在正文里：发现阶段只读账本、一次都没读来源表，
       * 用户重开后看到的就是“来源和目标已核对，可继续”—— 而他刚刚才因为来源变化
       * 被拒绝过一次（2026-10-31 真机）。措辞必须能被判据支撑，不能是祝愿。
       */
      return verdict.source === "matched"
        ? `原计划与当前授权、来源和目标已核对，可继续剩余 ${verdict.remaining} 个动作；已完成动作不会重写，完成后审阅结果。`
        : `原计划与当前授权和目标已核对（这份计划不读其他工作表作为来源），可继续剩余 ${verdict.remaining} 个动作；已完成动作不会重写，完成后审阅结果。`;
    case "undone":
      return verdict.detail ?? EXECUTION_RECOVERY_UNDONE_DETAIL;
    case "replan":
      return (
        `还有 ${verdict.remaining} 个动作没写。本机材料缺失、过期或不符合精确续跑条件。` +
        (verdict.reason ? `具体原因：${verdict.reason} ` : "") +
        "可在当前授权范围内重新读取并规划；通过校验后先执行，再审阅。"
      );
    default:
      return verdict.detail;
  }
}

/**
 * 这条恢复记录**还有没有下一步可做** —— 决定它要不要占用输入区上方那块固定区域。
 *
 * 存在理由（2026-11-01）：撤销确认、放弃、收尾之后，恢复卡已经没有任何动作，却仍然
 * 占着固定区域的主视图（或退成一枚 chip），把"还有 N 项"里那句话留给一件**已经结束**
 * 的事；换任务、发普通聊天之后它还在，看起来像还有一个待办。
 *
 * 三种终态**分开命名**，因为它们对用户说的不是同一件事：
 *
 * - `undone`   本次写入已被撤销（`kind === "undone"`）；
 * - `abandoned` 用户按「放弃变更」处置过（`blocked` + `abandoned: true`）。它仍然是
 *   `blocked`（不给写入/补记按钮），但**结论本身已经定了**，等的是"知道这件事结束了"，
 *   不是"再核对一次"；
 * - `settled`  账本已经记完成（`completed` + `recorded: true`），没有剩余动作。
 *
 * **不许扩大的边界**：`reconcile`（待核对 / 在途未知）、`resume`、`replan`、普通
 * `blocked`（现场冲突、执行权被占、来源变化…）与 `insufficient`（证据不足）都**不是**
 * 终态 —— 它们要么还要用户点一下，要么现场还没查清。把"没查清"说成"结束了"就是隐藏待办。
 */
export type ExecutionRecoveryTerminalReason = "undone" | "abandoned" | "settled";

export function executionRecoveryTerminalReason(
  verdict: ExecutionRecoveryVerdict
): ExecutionRecoveryTerminalReason | null {
  if (verdict.kind === "undone") return "undone";
  if (verdict.kind === "blocked" && verdict.abandoned === true) return "abandoned";
  if (verdict.kind === "completed" && verdict.recorded === true) return "settled";
  return null;
}

export function executionRecoveryIsTerminal(verdict: ExecutionRecoveryVerdict): boolean {
  return executionRecoveryTerminalReason(verdict) !== null;
}

/**
 * 终态摘要的那句话 —— 只在**所属任务的消息**里出现（可展开），不再画成可点的卡。
 *
 * 一句话必须同时说清"结束了"和"还留下了什么记录"：否则用户会去输入区上方找一个
 * 已经被收走的入口。展开后的正文用 `executionRecoveryDetail`，两边同一份。
 */
export const EXECUTION_RECOVERY_TERMINAL_SUMMARY: Record<
  ExecutionRecoveryTerminalReason,
  string
> = {
  undone: "这次写入已撤销，不再占用输入区；展开可看执行记录。",
  abandoned: "这次执行已按「放弃变更」处置，不再占用输入区；展开可看执行记录。",
  settled: "这次执行已记录完成，没有剩余动作；展开可看执行记录。"
};

export function executionRecoveryTerminalSummary(verdict: ExecutionRecoveryVerdict): string | null {
  const reason = executionRecoveryTerminalReason(verdict);
  return reason ? EXECUTION_RECOVERY_TERMINAL_SUMMARY[reason] : null;
}

/** 各结论允许出现的操作。返回空数组 = 不给任何写入/补记按钮。 */
export function executionRecoveryActions(verdict: ExecutionRecoveryVerdict): ExecutionRecoveryAction[] {
  switch (verdict.kind) {
    case "completed":
      return verdict.recorded ? ["dismiss"] : ["backfill", "dismiss"];
    case "reconcile":
      return ["reconcile", "dismiss"];
    case "resume":
      return ["resume", "dismiss"];
    case "replan":
      return ["replan", "dismiss"];
    case "undone":
      /*
       * 终态：**一个动作都不给**。不给“继续”“补记”，也不给“再撤一次”——
       * 撤销已经完成，再点一次只会试图把已经回到执行前的目标再写一遍。
       */
      return [];
    default:
      // `blocked` / `insufficient`：证据不够或现场冲突，一律不给写入或补记按钮，
      // 只留"已人工核对，清除提示"。
      return ["dismiss"];
  }
}
