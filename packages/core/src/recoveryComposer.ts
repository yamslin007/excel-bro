/**
 * 「恢复待办」融进输入框之后的那份**呈现判据**（Excel / WPS 共用，纯函数）。
 *
 * 背景：待核对 / 待续跑时，输入框上方那张恢复大卡 + 右下角那颗主按钮 + 消息里的
 * 技术记录，三处都在说同一件事，用户要先读三遍才知道"现在该点哪"。这一轮把**要用户
 * 点的那一下**搬进输入框内部，像澄清问卷那样把普通输入置换掉：一个标题、两三句人话、
 * 一颗文字主操作。
 *
 * 本模块只回答三个问题，且都从真实判据派生：
 *
 * 1. **现在该不该进入恢复操作区？**（`recoveryComposerState`）
 *    - 终态（撤销 / 放弃 / 收尾）⇒ 退出，回到普通输入框；
 *    - 结论给不出 `reconcile` / `resume` ⇒ 不进入（`blocked` / `replan` / 证据不足
 *      没有"一个主操作"可言，卡片留在固定区域，不在这里编一颗按钮）；
 *    - 有多条候选还没选定 ⇒ 不进入（替用户挑一条就是替他决定"哪次写入可以不管"）；
 *    - 用户按了「另提需求」，或输入框里已有草稿 ⇒ **不抢占**，退成一行紧凑待办入口，
 *      由用户自己决定什么时候切回来。
 * 2. **这次执行到底写了哪些目标、还剩哪些？**（`recoveryComposerTargets`）
 * 3. **里面写什么、按钮按下去会发生什么？**（`recoveryComposerView`）
 *
 * 四条不变量：
 *
 * A. **不硬编码任何步数 / 目标 / 完成状态。** 步数来自账本游标与动作总数，目标来自
 *    这次计划的动作清单（宿主从计划/账本读出来再传进来）；拿不到就**不说那一句**，
 *    绝不编一个「结果3!B2」出来。
 * B. **只有证据支持才说"已写入"。** 账本游标之前（`nextActionIndex`）是确认写入的；
 *    在途那一步只有 `outcome === "applied"` 才算写入，`not_applied` 说"还没有写入"，
 *    `pending` 一个字都不许往"已写入"上靠 —— 三句话不能合并。
 * C. **核对与继续是两件事。** 这里只给一颗按钮：待核对时是「检查…」那一档，核对完
 *    （结论变成 `resume`）才是「继续剩余 N 步」。文案必须说清"点它不会顺手把剩下的也跑了"。
 * D. **按钮不说做不到的事**：在途那一步还没确认时，按钮不能叫「检查已写入结果」
 *    （那时没有任何证据说它写进去了）—— 主操作文案按 `outcome` 分档，见
 *    `recoveryComposerPrimaryLabel`。
 *
 * 文案纪律（用户明确要求）：主要说明只讲"已经做了什么、还差什么、点按钮会发生什么"，
 * 不把「在途动作 / 现场 / 回执 / 补记」这类术语放在正文里 —— 那些词留在「查看详情」
 * 展开的执行记录里（`detail` / `progress` / `note`，与消息里的只读记录同一份来源）。
 */

import {
  executionRecoveryActions,
  executionRecoveryDetail,
  executionRecoveryIsTerminal,
  executionRecoveryNote,
  executionRecoveryProgress,
  type ExecutionRecoveryVerdict
} from "./executionRecoveryView";

/** 输入框里能承接的两步。其余动作（补记 / 重新规划 / 清除）都不是"一句话的主操作"。 */
export type RecoveryComposerAction = "reconcile" | "resume";

/** 下一步要写的位置。`range` 缺失时只报工作表名，不拼一个假的区域。 */
export interface RecoveryComposerTarget {
  sheet: string;
  range?: string | null;
  /** 动作类型（`writeFormulas` / `writeValues` / …）：只让「下一步」说具体，不参与判定。 */
  type?: string | null;
}

export interface RecoveryComposerFacts {
  verdict: ExecutionRecoveryVerdict;
  /** 账本里已确认完成的动作数（`nextActionIndex`）。 */
  nextActionIndex: number;
  /** 这次执行的动作总数。 */
  totalActions: number;
  /** 下一条要执行的动作位置；拿不到就是 `null`（**不猜**）。 */
  nextTarget?: RecoveryComposerTarget | null;
  /**
   * 这次计划的**全部动作**（顺序即执行顺序），由宿主从计划 / 账本读出来。
   *
   * 只用于把"已写入 / 尚未写入"两句说具体，本身**不是**判定输入：哪些动作算写入，
   * 由 `nextActionIndex` 与结论里的 `outcome` 决定（见 `recoveryComposerTargets`）。
   * 拿不到就省略 —— 正文退回"前 N 步已完成 / 后面还有 N 步"的说法，绝不编位置。
   */
  actions?: readonly RecoveryComposerTarget[] | null;
  /** 正在查询或提交（按钮禁用并显示等待文案）。 */
  busy?: boolean;
  /**
   * 用户**显式**选了哪一边。
   *
   * 两枚入口互相对望，缺一边就会有一个"点了没反应"的按钮：
   * - `"composer"` = 操作区里按了「另提需求」⇒ 把输入框要回去，旧任务退成待办入口；
   * - `"recovery"` = 待办入口上点了回来 ⇒ 重新占据输入框。
   *
   * `"recovery"` 能压过 `hasDraft`（草稿留在状态里，下次切回去还在），所以"有草稿就只能
   * 二选一"这种死角不存在。它压不过 `composerBusyElsewhere` / `ambiguous`：那两种情形下
   * 输入框确实不归它，硬抢会造出一个假的入口。
   */
  choice?: "composer" | "recovery" | null;
  /** 输入框里有草稿（文字或附件）：**不抢占**，退成待办入口。 */
  hasDraft?: boolean;
  /** 同一工作簿有多条候选、还没选定一条：不进入操作区。 */
  ambiguous?: boolean;
  /**
   * 输入框此刻归**别的正事**：任务在跑 / 审阅在等你决定 / 一道澄清题在等回答。
   *
   * **不含"恢复动作自己在跑"**：点「检查…」之后宿主会把全局状态置成执行中，那一下
   * 若是把输入框抢回去（换成普通 textarea），用户看到的就是"点了没反应、框还变了"。
   * 自己的请求在飞由 `busy` 表达（按钮禁用 + 等待文案），不在这里赶人。
   *
   * 三种情形都**连待办入口都不给**：它们各自在固定区域或澄清区里有自己的入口，
   * 不会因此失联；而在输入框里再挂一行"还有件事等着你"只会跟那颗"暂停 / 发送 /
   * 提交回答"抢同一块版面（2026-11-02 决定，与 `selectCurrentAction` 的优先级同向：
   * 运行与审阅都不该被一条待办提示顶掉）。
   */
  composerBusyElsewhere?: boolean;
}

export interface RecoveryComposerView {
  action: RecoveryComposerAction;
  /** 第一行：现在处于哪一步。 */
  headline: string;
  /** 正文两三句：已做什么、还差什么、点按钮会发生什么。 */
  lines: string[];
  /** 主操作文字（按证据分档，见 `recoveryComposerPrimaryLabel`）。 */
  primaryLabel: string;
  /** 主操作的悬停说明：说清"不会自动继续"。 */
  primaryTitle: string;
  /** 详情里展示的执行记录（与消息里的只读记录同一句）。 */
  progress: string;
  /** 详情正文（技术口径，允许出现"回执/补记"这类词）。 */
  detail: string;
  /** 详情里的边界说明。 */
  note: string;
}

/** 退成紧凑入口时那行字。 */
export interface RecoveryComposerEntry {
  action: RecoveryComposerAction;
  /** 一行放得下的短标签（入口按钮的可见文案）。 */
  label: string;
  /** 悬停 / 可访问名：说明切过去会看到什么。 */
  title: string;
}

export type RecoveryComposerState =
  | { mode: "recovery"; view: RecoveryComposerView }
  | { mode: "composer"; entry: RecoveryComposerEntry | null };

/** 「另提需求」之后，旧任务留在固定区域的那一枚入口（不是撤销、不是放弃）。 */
export const RECOVERY_COMPOSER_ENTRY_LABEL = "恢复待办";

/** 输入框里那块操作区的可访问名 / 模式名。 */
export const RECOVERY_COMPOSER_MODE_LABEL = "恢复待办";

/** 最多列几个位置；再多就说"等 N 处"，免得正文变成一张表。 */
const TARGET_LIST_LIMIT = 4;

/**
 * 结论 → 输入框能承接哪一步。
 *
 * 只认共享结论给的按钮集合（`executionRecoveryActions`），不在这里另议一遍：
 * 结论说证据不足就一个动作都没有，输入框里也就不该出现主操作。
 */
export function recoveryComposerAction(
  verdict: ExecutionRecoveryVerdict
): RecoveryComposerAction | null {
  if (executionRecoveryIsTerminal(verdict)) return null;
  const actions = executionRecoveryActions(verdict);
  if (actions.includes("reconcile")) return "reconcile";
  if (actions.includes("resume")) return "resume";
  return null;
}

/**
 * 这次执行还剩几步。
 *
 * 只认判据自己给的数：`resume` 的 `remaining` 是"核对后还差几步"，`reconcile` 的
 * `remaining` 是"**在途那一步之后**还差几步"（见 `executionRecoveryView` 的字段注释）。
 * 两者都不许拿 `total - nextActionIndex` 重新算一遍 —— 在途那一步既不算完成、也不算
 * 剩余，用游标减出来的数会把它再算一次（真机 B：三条动作、第 2 步在途，减出来是 2，
 * 而实际只剩第 3 步）。
 *
 * 只有在结论没给 `remaining` 时才退回游标，并且明确按"这一步之后"算。
 */
function remainingSteps(facts: RecoveryComposerFacts): number {
  const verdict = facts.verdict;
  if (verdict.kind === "resume" || verdict.kind === "reconcile") {
    if (typeof verdict.remaining === "number") return Math.max(0, verdict.remaining);
  }
  const total = Math.max(0, facts.totalActions);
  if (verdict.kind === "reconcile") {
    return Math.max(0, total - (verdict.actionIndex + 1));
  }
  const done = Math.min(Math.max(0, facts.nextActionIndex), total);
  return Math.max(0, total - done);
}

/**
 * 一个位置的人话（`结果3!B2`）。
 *
 * 只有工作表名时只报工作表名 —— 补一个假区域比不报更糟（用户会照着它去找）。
 */
export function recoveryComposerTargetLabel(
  target: RecoveryComposerTarget | null | undefined
): string | null {
  const sheet = typeof target?.sheet === "string" ? target.sheet.trim() : "";
  if (!sheet) return null;
  const range = typeof target?.range === "string" ? target.range.trim() : "";
  return range ? `${sheet}!${range}` : sheet;
}

/** 一串位置；超过 4 个就折成"等 N 处"（正文不是台账）。 */
function targetList(targets: readonly RecoveryComposerTarget[]): string {
  const labels = targets
    .map((target) => recoveryComposerTargetLabel(target))
    .filter((label): label is string => label !== null);
  if (labels.length <= TARGET_LIST_LIMIT) return labels.join("、");
  return `${labels.slice(0, TARGET_LIST_LIMIT).join("、")} 等 ${labels.length} 处`;
}

/** 这次执行的目标清单：谁写了、谁没写、在途那一步算什么。 */
export interface RecoveryComposerTargetSplit {
  /** 有证据支持"已经写进去了"的目标（账本游标之前 + 结论说 applied 的那一步）。 */
  written: RecoveryComposerTarget[];
  /** 有证据支持"还没写"的目标；`pending` 的在途那一步**不在**这里。 */
  unwritten: RecoveryComposerTarget[];
  /** 在途那一步：只有 `reconcile` 才有；`pending` 表示还没确认。 */
  inFlight: { target: RecoveryComposerTarget; outcome: "applied" | "not_applied" | "pending" } | null;
}

/**
 * 把"这次写了哪些目标"从**证据**里拆出来。
 *
 * 判据只有两个：账本游标（`nextActionIndex` = 前 N 步确认写入）与结论里的 `outcome`
 * （在途那一步）。二者之外一律不猜：
 *
 * - `pending` ⇒ 在途那一步既不进 `written`，也不进 `unwritten`（它可能写了，也可能没写），
 *   只在 `inFlight` 里挂着，界面据此说"还没确认"；
 * - `not_applied` ⇒ 在途那一步进 `unwritten`（结论说现场还是写前状态）；
 * - `resume` ⇒ 游标之前全算写入，游标之后全算没写（核对已经确认过在途那一步）。
 *
 * `actions` 缺失或长度对不上时，返回空清单 —— 界面退回"只报步数"的说法，不编位置。
 */
export function recoveryComposerTargets(
  facts: RecoveryComposerFacts
): RecoveryComposerTargetSplit {
  const all = Array.isArray(facts.actions) ? facts.actions : [];
  const split: RecoveryComposerTargetSplit = { written: [], unwritten: [], inFlight: null };
  if (all.length === 0) return split;
  const clamp = (value: number, max: number) =>
    Math.max(0, Math.min(Number.isFinite(value) ? Math.trunc(value) : 0, max));
  if (facts.verdict.kind === "resume") {
    const boundary = clamp(facts.nextActionIndex, all.length);
    split.written = all.slice(0, boundary);
    split.unwritten = all.slice(boundary);
    return split;
  }
  if (facts.verdict.kind !== "reconcile") return split;
  const index = clamp(facts.verdict.actionIndex, all.length - 1);
  const confirmed = clamp(Math.min(facts.nextActionIndex, index), all.length);
  split.written = all.slice(0, confirmed);
  const target = all[index];
  const outcome = facts.verdict.outcome;
  if (target) split.inFlight = { target, outcome };
  if (target && outcome === "applied") split.written.push(target);
  if (target && outcome === "not_applied") split.unwritten.push(target);
  split.unwritten.push(...all.slice(index + 1));
  return split;
}

/** 「下一步：向结果3!B2写入公式。」里的那个动词短语 —— 只按动作类型分档，不猜内容。 */
function nextStepPhrase(target: RecoveryComposerTarget | null | undefined): string | null {
  const label = recoveryComposerTargetLabel(target);
  if (!label) return null;
  switch (typeof target?.type === "string" ? target.type : "") {
    case "writeFormulas":
      return `向${label}写入公式`;
    case "writeValues":
      return `向${label}写入数值`;
    case "writeTable":
      return `向${label}写入表格`;
    case "clearRange":
      return `清空${label}`;
    case "copyRange":
      return `把内容复制到${label}`;
    case "setFill":
    case "setFont":
    case "setNumberFormat":
    case "autofit":
    case "setAlignment":
      return `调整${label}的格式`;
    default:
      return `向${label}写入`;
  }
}

/**
 * 第一行：只有**证据支持**的事实。
 *
 * `outcome` 三档分开写，因为它们的下一步完全不同：
 *  - `applied`（现场与写后一致）⇒ 这步其实已经写进去了，核对只是补那条记录；
 *    这时"还剩 N 步"是**能算准**的（在途之后还有几个），所以照实报；
 *  - `not_applied` ⇒ 这步没写，不报"还剩几步"，先说自己没写；
 *  - `pending`（还没核对）⇒ 只能说"待核对"，**不能说已写入**，也不报剩余步数
 *    （这一格如果最后判定没写，剩余步数会多一步，先报就是预支结论）。
 */
export function recoveryComposerHeadline(facts: RecoveryComposerFacts): string {
  const verdict = facts.verdict;
  if (verdict.kind === "resume") return "检查完成，可以继续";
  if (verdict.kind !== "reconcile") return "";
  const step = verdict.actionIndex + 1;
  if (verdict.outcome === "applied") return `任务中断 · 还剩 ${remainingSteps(facts)} 步`;
  if (verdict.outcome === "not_applied") return `任务中断 · 第 ${step} 步没有写入`;
  return `任务中断 · 第 ${step} 步结果待核对`;
}

/**
 * 正文：已写什么 / 还没写什么 / 点按钮会发生什么。
 *
 * 每句都对着一个字段，缺字段就不写那一句 —— 这正是"步数、目标、完成状态都来自
 * 真实判据"的落地方式。
 */
export function recoveryComposerLines(facts: RecoveryComposerFacts): string[] {
  const verdict = facts.verdict;
  const split = recoveryComposerTargets(facts);
  if (verdict.kind === "reconcile") {
    const step = verdict.actionIndex + 1;
    const lines: string[] = [];
    /* 说不清位置的动作（`sheet` 为空）在 `targetList` 里被过滤掉：整句不成立就不写。 */
    const writtenText = targetList(split.written);
    if (writtenText) {
      lines.push(`已写入：${writtenText}`);
    }
    /*
     * 在途那一步还没确认时，单独说它一句：它不是"已写入"，也不是"尚未写入" ——
     * 把两种情况压成一句话，就等于替现场下结论。
     */
    if (split.inFlight && split.inFlight.outcome === "pending") {
      const label = recoveryComposerTargetLabel(split.inFlight.target);
      lines.push(label ? `第 ${step} 步（${label}）是否已写入还没确认。` : `第 ${step} 步是否已写入还没确认。`);
    }
    const unwrittenText = targetList(split.unwritten);
    if (unwrittenText) {
      lines.push(`尚未写入：${unwrittenText}`);
    }
    if (lines.length === 0) {
      /*
       * 宿主没给出动作清单时的退路：只报得出来的那几件事（步数 + outcome），
       * 不编任何位置。判据仍是 `nextActionIndex` / `totalActions`。
       */
      const total = Math.max(0, facts.totalActions);
      const done = Math.max(0, Math.min(facts.nextActionIndex, verdict.actionIndex));
      const written =
        verdict.outcome === "applied"
          ? `第 ${step} 步已经写入，只差把结果记下来`
          : verdict.outcome === "not_applied"
            ? `第 ${step} 步还没有写入`
            : `第 ${step} 步写没写完还没确认`;
      lines.push(
        `${done > 0 ? `前 ${done} 步已完成` : "还没有已完成的步骤"}；${written}；共 ${total} 步。`
      );
    }
    lines.push(
      writtenText ? "先检查已写入结果，不会修改表格。" : "先检查当前结果，不会修改表格。"
    );
    lines.push("检查完成后还需要你再点一次，才会执行剩下的步骤。");
    return lines;
  }
  if (verdict.kind === "resume") {
    const remaining = remainingSteps(facts);
    const done = Math.max(0, Math.min(facts.nextActionIndex, facts.totalActions));
    const next =
      split.unwritten.find((target) => recoveryComposerTargetLabel(target)) ??
      facts.nextTarget ??
      null;
    const phrase = nextStepPhrase(next);
    return [
      phrase ? `下一步：${phrase}。` : `继续后将执行剩余 ${remaining} 步。`,
      done > 0 ? `前 ${done} 步不会重复执行。` : "还没有已完成的步骤。"
    ];
  }
  return [];
}

/**
 * 主操作文案 —— **按证据分档**，另加一档"继续"。
 *
 * 为什么不是一句固定的「核对结果」（2026-11-02 用户要求"主按钮使用明确文字"）：
 * 按钮上写的就是"点下去会发生什么"，所以它必须说得比"核对"更具体：
 *  - 有已写入的目标 ⇒「检查已写入结果」；
 *  - 结论说这一步没写 ⇒「检查当前结果」（说"已写入"就是假的）；
 *  - 还没确认 ⇒「检查写入结果」（不说"已写入"，也不说"没有"）。
 *
 * 续跑那一档带上剩余步数：「继续剩余 1 步」比「继续执行」更能让人核对一遍要动几格。
 */
export function recoveryComposerPrimaryLabel(facts: RecoveryComposerFacts): string {
  const action = recoveryComposerAction(facts.verdict);
  if (action === "reconcile") {
    const verdict = facts.verdict;
    if (verdict.kind === "reconcile" && verdict.outcome === "not_applied") return "检查当前结果";
    if (verdict.kind === "reconcile" && verdict.outcome === "pending") return "检查写入结果";
    return "检查已写入结果";
  }
  if (action === "resume") return `继续剩余 ${remainingSteps(facts)} 步`;
  return "";
}

export function recoveryComposerPrimaryTitle(
  action: RecoveryComposerAction,
  facts: RecoveryComposerFacts
): string {
  if (action === "reconcile") {
    const step = facts.verdict.kind === "reconcile" ? facts.verdict.actionIndex + 1 : null;
    return step === null
      ? "只读核对当前结果并补上完成记录：不修改任何单元格，也不会自动继续。"
      : `只读核对第 ${step} 步的结果并补上完成记录：不修改任何单元格，也不会自动继续。`;
  }
  const remaining = remainingSteps(facts);
  return `只执行剩下的 ${remaining} 步；已完成的部分不会重写，结束后再审阅结果。`;
}

/** 主操作放在这里，是因为界面与测试都要读同一份文案；判据仍然只有上面那几个函数。 */
export function recoveryComposerView(facts: RecoveryComposerFacts): RecoveryComposerView {
  const action = recoveryComposerAction(facts.verdict);
  if (!action) {
    /*
     * 走到这里说明调用方没先问 `recoveryComposerState`。给一个空壳而不是抛错：
     * 抛错会让"结论刚好在这一帧变化"变成整块界面崩掉，而空壳只是不显示。
     */
    return {
      action: "reconcile",
      headline: "",
      lines: [],
      primaryLabel: "",
      primaryTitle: "",
      progress: "",
      detail: "",
      note: ""
    };
  }
  return {
    action,
    headline: recoveryComposerHeadline(facts),
    lines: recoveryComposerLines(facts),
    primaryLabel: recoveryComposerPrimaryLabel(facts),
    primaryTitle: recoveryComposerPrimaryTitle(action, facts),
    /*
     * 三句技术口径**原样转发**共享层（消息里的只读记录读同一份）：
     * 详情不与正文抢说法，正文说人话，详情给依据。
     */
    progress: executionRecoveryProgress(facts.verdict, {
      nextActionIndex: facts.nextActionIndex,
      totalActions: facts.totalActions
    }),
    detail: executionRecoveryDetail(facts.verdict),
    note: executionRecoveryNote(facts.verdict)
  };
}

/** 退成入口时那枚按钮的文案：短、带步数，一眼看出还有几件事没做。 */
export function recoveryComposerEntry(facts: RecoveryComposerFacts): RecoveryComposerEntry | null {
  const action = recoveryComposerAction(facts.verdict);
  if (!action) return null;
  const remaining = remainingSteps(facts);
  const label = action === "reconcile"
    ? `${RECOVERY_COMPOSER_ENTRY_LABEL} · 待核对`
    : `${RECOVERY_COMPOSER_ENTRY_LABEL} · 还剩 ${remaining} 步`;
  return {
    action,
    label,
    title: action === "reconcile"
      ? "回到恢复处理：只读核对当前结果，不修改单元格。"
      : `回到恢复处理：继续剩下的 ${remaining} 步，已完成的部分不会重写。`
  };
}

/**
 * 判据入口：现在该把普通输入框换成恢复操作区，还是只留一枚待办入口。
 *
 * 顺序即判据，每条都对应一类真实状态：
 *  1. 终态 ⇒ 退回普通输入框（`entry: null`）。撤销/放弃之后没有下一步可做，
 *     留着入口就成了"还有件事没处理"的假象。
 *  2. 结论给不出核对/继续 ⇒ 同上（`replan` 的复选框、`blocked` 的原因都留在固定区域，
 *     塞进一行操作区只会把它们压成一句话）。
 *  3. 有多条候选 ⇒ 不进入，也**不给入口**（**不替用户挑**；候选卡自己就在固定区域里）。
 *  4. 输入框归别的正事（在跑 / 等审阅 / 在答题）⇒ 不进入、不给入口（见字段注释）。
 *  5. 用户按了「另提需求」⇒ 只留入口，把输入框还给他。
 *  6. 用户从入口点了回来 ⇒ 进入操作区（草稿留着，切回去就能看见）。
 *  7. 有草稿 ⇒ 只留入口，不抢他正在写的东西。
 *  8. 其余 ⇒ 进入操作区。
 */
export function recoveryComposerState(facts: RecoveryComposerFacts): RecoveryComposerState {
  const entry = recoveryComposerEntry(facts);
  if (!entry) return { mode: "composer", entry: null };
  if (facts.ambiguous || facts.composerBusyElsewhere) {
    return { mode: "composer", entry: null };
  }
  if (facts.choice === "composer") return { mode: "composer", entry };
  if (facts.choice === "recovery") return { mode: "recovery", view: recoveryComposerView(facts) };
  if (facts.hasDraft) return { mode: "composer", entry };
  return { mode: "recovery", view: recoveryComposerView(facts) };
}
