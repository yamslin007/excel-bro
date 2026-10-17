/**
 * 计划预览卡的**执行方式选择**（Excel / WPS **共用语义**，宿主操作留在各自适配器）。
 *
 * 用户在计划预览卡上选的是「怎么落地」，不是「写不写」：
 *
 *  - `direct`（直接执行到原表）：跳过**副本比对**这一步，确认后直接写当前目标工作簿。
 *  - `copy-compare`（先生成副本比对）：生成审阅副本，核对后再写回原表。
 *
 * **"直接执行"只跳过副本比对**，不跳过任何别的东西：授权边界、计划校验、
 * 执行前复核、恢复数据采集、回读验证、撤销保障在两条路径上一模一样。
 * 把"直接执行"实现成"少几道守卫"是这个模块存在的全部理由 ——
 * `directExecutionStillRuns()` 是这条不变式的唯一判据入口，任何调用点想省掉其中
 * 一项时都必须先过它。
 *
 * **"文件"的口径**：用户说的"文件"是**当前目标工作簿**，不是磁盘上新文件。
 * 所以 `direct` 不 `Save As`、不建副本文件、不换工作簿 —— 它就是原地写。
 * 按钮与文案不许出现"保存到文件/另存"，那会让用户以为目标换了地方。
 *
 * **切换规则**（不可越过）：
 *  - 副本生成**之前**可以自由切换；
 *  - 副本**已存在**时切到 `direct`，必须让旧审阅的确认与回调失效，并明确交代
 *    旧副本怎么办 —— 不能留下两个有效写回入口（`modeSwitchInvalidatesReview()`）。
 *  - 执行在途时禁止切换、禁止重复提交（`executionModeSwitchBlockReason()`）。
 *
 * 这里只做**确定性判定与文案**。生成副本、关副本、真正写表都在各端。
 */

/** 用户可选的两种执行方式。 */
export type PlanExecutionMode = 'direct' | 'copy-compare';

/** 副本比对在某宿主上当前**不可用**（缺基础 API / 能力探测失败）。 */
export type ExecutionModeAvailability = {
  'copy-compare': { available: boolean; /** 不可用时说清为什么；不是"功能没做"，而是缺什么。 */ reason: string | null };
};

export const PLAN_EXECUTION_MODE_LABELS: Record<PlanExecutionMode, string> = {
  direct: '直接执行到原表',
  'copy-compare': '先生成副本比对'
};

export type ExecutionModeOption = {
  mode: PlanExecutionMode;
  label: string;
  /** 这条方式到底做什么 / 不做什么 —— 决策必需信息，放在按钮附近而不是折叠里。 */
  detail: string;
  available: boolean;
  /** 不可用原因；可用时为 null。 */
  blockedReason: string | null;
};

/**
 * 直接执行的**逐项说明**。
 *
 * 关键是把"跳过了什么"和"没跳过什么"同时写出来：只写"跳过副本比对"不够，
 * 用户会以为可以绕过校验。措辞已按"当前目标工作簿"而不是"磁盘文件"。
 */
export const DIRECT_EXECUTION_DETAIL =
  '确认后按本计划直接写入当前工作簿，不生成副本、不新建文件。仍会执行授权范围与计划校验、执行前复核、恢复点采集和写入后回读，并保留撤销入口。';

export const COPY_COMPARE_EXECUTION_DETAIL =
  '先生成独立审阅副本并试算，核对后由你确认写回当前工作簿；写回并验证成功后关闭本次副本，原表始终保持打开。';

/** 两种方式的**共同前置**：都还没落到原表。 */
export const EXECUTION_MODE_PREFACE = '两种方式都会先展示白名单计划、目标范围与变更摘要，再由你确认。';

/**
 * 预览卡上那排选择项。
 *
 * `selected` 是当前生效的方式。**只做展示与文案**：能不能点由
 * `executionModeSwitchBlockReason()` 判（执行在途、已有副本要切走等），
 * 不在这里偷偷把不可点的项藏起来 —— 藏起来等于用户不知道为什么少了一个选择。
 */
export function executionModeOptions(
  selected: PlanExecutionMode,
  availability: ExecutionModeAvailability
): ExecutionModeOption[] {
  const options: ExecutionModeOption[] = [
    {
      mode: 'direct',
      label: PLAN_EXECUTION_MODE_LABELS.direct,
      detail: DIRECT_EXECUTION_DETAIL,
      available: true,
      blockedReason: null
    },
    {
      mode: 'copy-compare',
      label: PLAN_EXECUTION_MODE_LABELS['copy-compare'],
      detail: COPY_COMPARE_EXECUTION_DETAIL,
      available: availability['copy-compare'].available,
      blockedReason: availability['copy-compare'].reason
    }
  ];
  // `selected` 只用于调用方对照；这里原样返回两项，让界面自己标出当前选中项。
  void selected;
  return options;
}

/**
 * 「直接执行」到底省了什么、没省什么。
 *
 * 存在的唯一理由是让"直接执行 = 少一层预览"这句话**可测**：
 * 副本比对是唯一被允许跳过的一步，其余每一项都必须为 `true`。
 * 任何一端想把其中某项改成 `false`（比如为了少跑一次回读），这个判据立刻变红。
 */
export function directExecutionStillRuns(): {
  authorizationCheck: true;
  planValidation: true;
  preExecutionRecheck: true;
  recoverySnapshot: true;
  readbackVerification: true;
  undoAvailable: true;
} {
  return {
    authorizationCheck: true,
    planValidation: true,
    preExecutionRecheck: true,
    recoverySnapshot: true,
    readbackVerification: true,
    undoAvailable: true
  };
}

/** 唯一被 `direct` 跳过的一步。 */
export const DIRECT_EXECUTION_SKIPS: readonly string[] = ['copy-comparison'];

/** 副本在不在场 —— 决定"能不能直接切到 `direct`"的关键事实。 */
export type ModeSwitchState = {
  /** 当前选中的方式。 */
  current: PlanExecutionMode;
  /** 目标方式。 */
  next: PlanExecutionMode;
  /** 本次审阅副本是否已经生成（无论当前选哪种方式）。 */
  copyExists: boolean;
  /** 是否有一次写入/关闭正在跑。 */
  inFlight: boolean;
  /** 计划是否处于编辑态（编辑中不提供切换）。 */
  editingPlan: boolean;
};

export type ModeSwitchBlock = 'in_flight' | 'editing_plan' | 'no_change' | 'copy_compare_unavailable';

export type ModeSwitchDecision = {
  /** 能不能切。 */
  allowed: boolean;
  reason: ModeSwitchBlock | null;
  /**
   * 切换到 `direct` 且副本已在场时，**旧审阅必须失效**。
   *
   * 这一条是安全边界而不是清理细节：旧审阅的确认框与回调若还有效，就同时存在
   * "直接写回"与"确认后写回"两个有效入口 —— 用户点哪个都可能写表，
   * 而其中一个写的是用户根本没在核对的那份计划。
   */
  invalidatesReview: boolean;
  /** 旧副本需要用户知道的处理说明（失效不等于删掉副本）。 */
  reviewNotice: string | null;
};

/** 副本路径当前不可用时，切到它要拦下。 */
export function modeSwitchBlockReason(
  state: ModeSwitchState,
  availability: ExecutionModeAvailability
): ModeSwitchBlock | null {
  if (state.inFlight) return 'in_flight';
  if (state.editingPlan) return 'editing_plan';
  if (state.current === state.next) return 'no_change';
  if (state.next === 'copy-compare' && !availability['copy-compare'].available) return 'copy_compare_unavailable';
  return null;
}

/**
 * 切换的完整判定：能不能切、切了要不要让旧审阅失效、旧副本怎么交代。
 *
 * 判据**只读入参**（现场事实），不读任何闭包变量 —— 渲染期闭包里的
 * `copyExists` 是它诞生那一刻的事实，放弃之后那个值还是旧的。
 */
export function executionModeSwitchDecision(
  state: ModeSwitchState,
  availability: ExecutionModeAvailability
): ModeSwitchDecision {
  const reason = modeSwitchBlockReason(state, availability);
  if (reason) return { allowed: false, reason, invalidatesReview: false, reviewNotice: null };
  const invalidatesReview = state.next === 'direct' && state.copyExists;
  return {
    allowed: true,
    reason: null,
    invalidatesReview,
    reviewNotice: invalidatesReview
      ? '已切换为直接执行：上一份副本的确认与写回入口同时失效。请核对窗口名后自行处理那份副本，它不会被自动关闭。'
      : null
  };
}

/** 两种方式在执行在途时都不许再次提交。 */
export function executionSubmitBlocked(state: { inFlight: boolean }): boolean {
  return state.inFlight;
}

/**
 * 一条写回请求发起时记下的身份，与回来时的现场比对。
 *
 * 与 `reviewWriteBackGate` 的三元组同构，但**多一项执行方式**：
 * 副本还在、用户把方式切成了 `direct` 时，"同一次审阅 / 同一计划版本 /
 * 同一原工作簿"三项全同 —— 三项都成立，但这一次请求的**前提**（先核对副本）
 * 已经不成立了。所以方式必须进身份，否则切换后的旧回调还能写表。
 */
export type ExecutionModeWriteBackTarget = {
  review: unknown;
  plan: unknown;
  workbookIdentity: string;
  mode: PlanExecutionMode;
};

export function executionModeTargetChanged(
  issued: ExecutionModeWriteBackTarget | null,
  current: ExecutionModeWriteBackTarget | null
): boolean {
  if (!issued || !current) return true;
  return issued.review !== current.review ||
    issued.plan !== current.plan ||
    issued.workbookIdentity !== current.workbookIdentity ||
    issued.mode !== current.mode;
}