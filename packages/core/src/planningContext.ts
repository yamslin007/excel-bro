/**
 * 「规划中断后可继续」的受控现场（Excel / WPS / 服务端共用同一份判据）。
 *
 * 边界（与项目既有纪律一致，不得在此放宽）：
 *
 * 1. **这不是 SDK 会话恢复。** 它只记录"按当前现场重新分析"所需的最小事实：
 *    原需求、已确认的澄清、授权范围身份、来源指纹、有归属的结果引用、已完成工具步骤。
 *    续跑时**重新读取、重新计算、重新规划**，不复用模型上下文。
 *
 * 2. **不是"有检查点就能继续"。** `checkpoint_json` 非空只是必要条件。必须同时满足
 *    版本支持、任务归属一致、阶段受支持、继续资格为 `continue`、任务版本不回退、
 *    无任何值矩阵字段。任一条不成立都明确拒绝，并给出固定错误码。
 *
 * 3. **旧循环的检查点不得被当成 planning_context。** 判别点是 `kind`：
 *    旧结构没有 `kind`，`validatePlanningContext` 会以 `PLANNING_CONTEXT_KIND_MISMATCH` 拒绝。
 *
 * 4. **值矩阵结构上进不来。** 字段白名单 + 递归禁用键（`values` / `rows` / `formulas` /
 *    `expected` …）双保险：即使某天有人在服务端误把计划塞进检查点，也会被判 `SENSITIVE`。
 */

/** 检查点类型判别值。缺它的一律不是本契约（含旧循环检查点）。 */
export const PLANNING_CONTEXT_KIND = "planning_context";

/** 契约版本。字段语义变化必须升版本，并让旧版本走"不支持"分支。 */
export const PLANNING_CONTEXT_SCHEMA_VERSION = 1;

/** 允许出现的阶段。首版只支持"尚未进入宿主写入阶段"的规划中断。 */
export type PlanningContextStage = "before_model" | "after_tool";

/** 允许的继续资格。`replan` 表示"现场已不足以接续"，不得当 `continue` 用。 */
export type PlanningContextResumability = "continue" | "replan";

export interface PlanningContextConfirmation {
  questionId: string;
  /** 字段名（"关联方式"之类）。用于让模型知道这个答案在回答什么。 */
  label?: string | null;
  answerLabel: string;
}

export interface PlanningContextTarget {
  /** 工作簿身份依据（Office：`url:`；WPS：`wps-path:`）。命名空间是硬判据。 */
  identity: string;
  /**
   * 授权集合的证据键（共享层 `authorizedTargetEvidenceKey` 产出）。
   *
   * **可以为空**，但空不等于"免检"：客户端比较时按 `identity` + `sheetNames`
   * **重新计算**证据键，因此伪造或省略它最多导致"不能继续"，不会导致错误放行。
   * 存下来只是为了审计时能直接看到当时那一份。
   */
  evidenceKey?: string | null;
  /** 授权成员表名。 */
  sheetNames: string[];
}

export interface PlanningContextAuthorization {
  /**
   * 当时授权成员。**身份依据就是"身份 + 成员集合"**，不含点选顺序 ——
   * 比较时按共享层 `authorizedTargetEvidenceKey` 重算，不在这里另存一份版本号
   * （服务端并不持有授权版本，填一个假值比不填更糟）。
   */
  sheetNames: string[];
}

export interface PlanningContextResultRef {
  resultRef: string;
  /** 该引用产生时的来源指纹；`null` 表示当时就没有指纹可记。 */
  sourceFingerprint: string | null;
}

export interface PlanningContextToolStep {
  tool: string;
  status: "completed";
}

export interface PlanningContextRecord {
  kind: typeof PLANNING_CONTEXT_KIND;
  schemaVersion: typeof PLANNING_CONTEXT_SCHEMA_VERSION;
  /** 任务身份。必须与 `turnId` 相同（同一条账本记录）。 */
  taskId: string;
  turnId: string;
  /** 会话归属：写这份检查点的服务进程。与账本 `executor_id` 一致性由服务端复核。 */
  executorId: string;
  /** 写检查点时的任务版本。必须 `>= 1` 且不超过当前版本。 */
  taskRevision: number;
  stage: PlanningContextStage;
  resumability: PlanningContextResumability;
  /** ISO 时间戳。 */
  savedAt: string;
  /** 原需求（已截断；不含数据行）。 */
  prompt: string;
  planningGoal: string | null;
  confirmations: PlanningContextConfirmation[];
  target: PlanningContextTarget;
  authorization: PlanningContextAuthorization;
  sourceFingerprint: string | null;
  sourceFingerprintSheets: string[];
  resultRefs: PlanningContextResultRef[];
  completedTools: PlanningContextToolStep[];
  /** 明确失败的步骤。`null` 表示没有。 */
  failedTool: { tool: string; errorCode: string } | null;
  /** 结果未知的步骤。首版恒为 `null`：规划阶段不写宿主，出现即拒绝继续。 */
  unknownTool: { tool: string } | null;
}

/** 记录里**允许**出现的顶层键。多一个都算伪造。 */
export const PLANNING_CONTEXT_FIELDS: readonly string[] = [
  "kind",
  "schemaVersion",
  "taskId",
  "turnId",
  "executorId",
  "taskRevision",
  "stage",
  "resumability",
  "savedAt",
  "prompt",
  "planningGoal",
  "confirmations",
  "target",
  "authorization",
  "sourceFingerprint",
  "sourceFingerprintSheets",
  "resultRefs",
  "completedTools",
  "failedTool",
  "unknownTool"
];

/** 必须存在且非空的键。 */
export const PLANNING_CONTEXT_REQUIRED_FIELDS: readonly string[] = [
  "kind",
  "schemaVersion",
  "taskId",
  "turnId",
  "executorId",
  "taskRevision",
  "stage",
  "resumability",
  "savedAt",
  "prompt",
  "confirmations",
  "target",
  "authorization",
  "sourceFingerprintSheets",
  "resultRefs",
  "completedTools"
];

/**
 * 值矩阵键名。**按名字拦，不按动作类型枚举** —— 新增动作类型忘记登记也不会漏值。
 * 与 `server/app/ledger_privacy.MATRIX_FIELD_NAMES` 同一份口径。
 */
export const PLANNING_CONTEXT_FORBIDDEN_KEYS: readonly string[] = [
  "values",
  "rows",
  "headers",
  "formulas",
  "data",
  "preview",
  "matrix",
  "expected",
  "plan",
  "actions",
  "acceptanceCriteria",
  "snapshots",
  "before"
];

export const PLANNING_CONTEXT_LIMITS = {
  maxPromptChars: 4000,
  maxPlanningGoalChars: 400,
  maxConfirmations: 8,
  maxResultRefs: 12,
  maxCompletedTools: 32,
  maxSheetNames: 256,
  maxLabelChars: 80,
  maxToolNameChars: 80,
  maxErrorCodeChars: 80
} as const;

/** 固定拒绝码。用户可见文案由 `planningContextMessage` 统一给出。 */
export const PLANNING_CONTEXT_CODES = {
  missing: "PLANNING_CONTEXT_MISSING",
  kindMismatch: "PLANNING_CONTEXT_KIND_MISMATCH",
  unsupportedVersion: "PLANNING_CONTEXT_UNSUPPORTED_VERSION",
  invalid: "PLANNING_CONTEXT_INVALID",
  sensitive: "PLANNING_CONTEXT_SENSITIVE",
  taskMismatch: "PLANNING_CONTEXT_TASK_MISMATCH",
  stageUnsupported: "PLANNING_CONTEXT_STAGE_UNSUPPORTED",
  notContinuable: "PLANNING_CONTEXT_NOT_CONTINUABLE",
  staleRevision: "PLANNING_CONTEXT_STALE_REVISION"
} as const;

export type PlanningContextCode =
  (typeof PLANNING_CONTEXT_CODES)[keyof typeof PLANNING_CONTEXT_CODES];

export interface PlanningContextExpectation {
  turnId: string;
  /** 当前账本版本。检查点版本不能高于它。 */
  revision: number;
}

export type PlanningContextVerdict =
  | { ok: true; record: PlanningContextRecord }
  | { ok: false; code: PlanningContextCode; detail: string };

const RESULT_REF_PATTERN = /^res_[0-9a-f]{32}$/;

/** 工作簿身份必须是既有命名空间（`url:` / `wps-path:`）。 */
export const TARGET_IDENTITY_PATTERN = /^(?:url:.+|wps-path:.+)$/;

function reject(code: PlanningContextCode, detail: string): PlanningContextVerdict {
  return { ok: false, code, detail };
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

/** 递归查找禁用键。返回命中的键名，未命中返回 `null`。 */
export function findForbiddenKey(
  value: unknown,
  depth = 0
): string | null {
  if (depth > 8) return null;
  if (Array.isArray(value)) {
    for (const item of value) {
      const hit = findForbiddenKey(item, depth + 1);
      if (hit) return hit;
    }
    return null;
  }
  if (!isPlainObject(value)) return null;
  for (const [key, item] of Object.entries(value)) {
    if (PLANNING_CONTEXT_FORBIDDEN_KEYS.includes(key)) return key;
    const hit = findForbiddenKey(item, depth + 1);
    if (hit) return hit;
  }
  return null;
}

function validateStringList(
  value: unknown,
  max: number,
  label: string
): { ok: true; list: string[] } | { ok: false; detail: string } {
  if (!Array.isArray(value)) return { ok: false, detail: `${label} 必须是数组` };
  if (value.length > max) return { ok: false, detail: `${label} 超出上限` };
  const list: string[] = [];
  for (const item of value) {
    if (!isNonEmptyString(item)) return { ok: false, detail: `${label} 含空成员` };
    list.push(item.trim());
  }
  return { ok: true, list };
}

/**
 * 校验一份检查点能否用作"继续分析"的现场。
 *
 * 顺序是刻意的：先判"是不是这种检查点"，再判版本，再看有没有值，
 * 最后才判归属与版本号。这样"旧循环检查点"与"损坏记录"不会都被糊成"没有现场"。
 */
export function validatePlanningContext(
  value: unknown,
  expect: PlanningContextExpectation
): PlanningContextVerdict {
  if (value === null || value === undefined) {
    return reject(PLANNING_CONTEXT_CODES.missing, "没有可用的规划现场");
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return reject(PLANNING_CONTEXT_CODES.missing, "规划现场为空");
    return reject(PLANNING_CONTEXT_CODES.invalid, "规划现场不是结构化记录");
  }
  if (!isPlainObject(value)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "规划现场不是结构化记录");
  }
  if (value.kind !== PLANNING_CONTEXT_KIND) {
    // 旧循环检查点没有 kind；别的 kind 也不是本契约。
    return reject(
      PLANNING_CONTEXT_CODES.kindMismatch,
      "这不是可继续的规划现场（可能来自旧版本或别的恢复通道）"
    );
  }
  if (value.schemaVersion !== PLANNING_CONTEXT_SCHEMA_VERSION) {
    return reject(
      PLANNING_CONTEXT_CODES.unsupportedVersion,
      "规划现场版本不受支持，不能继续"
    );
  }
  const forbidden = findForbiddenKey(value);
  if (forbidden) {
    return reject(
      PLANNING_CONTEXT_CODES.sensitive,
      `规划现场含不应持久化的内容（${forbidden}）`
    );
  }
  for (const key of Object.keys(value)) {
    if (!PLANNING_CONTEXT_FIELDS.includes(key)) {
      return reject(PLANNING_CONTEXT_CODES.invalid, `规划现场含未知字段（${key}）`);
    }
  }
  for (const key of PLANNING_CONTEXT_REQUIRED_FIELDS) {
    if (!(key in value)) {
      return reject(PLANNING_CONTEXT_CODES.invalid, `规划现场缺少字段（${key}）`);
    }
  }
  if (!isNonEmptyString(value.taskId) || !isNonEmptyString(value.turnId)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "规划现场缺少任务身份");
  }
  if (value.taskId !== value.turnId) {
    return reject(PLANNING_CONTEXT_CODES.taskMismatch, "任务身份不一致");
  }
  if (value.taskId !== expect.turnId) {
    return reject(PLANNING_CONTEXT_CODES.taskMismatch, "规划现场属于另一个任务");
  }
  if (!isNonEmptyString(value.executorId)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "规划现场缺少会话归属");
  }
  const revision = value.taskRevision;
  if (typeof revision !== "number" || !Number.isInteger(revision) || revision < 1) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "规划现场版本号无效");
  }
  if (revision > expect.revision) {
    return reject(
      PLANNING_CONTEXT_CODES.staleRevision,
      "规划现场比当前任务更新，不能据此继续"
    );
  }
  if (value.stage !== "before_model" && value.stage !== "after_tool") {
    return reject(PLANNING_CONTEXT_CODES.stageUnsupported, "该中断阶段暂不支持继续");
  }
  if (value.resumability !== "continue") {
    return reject(PLANNING_CONTEXT_CODES.notContinuable, "该现场不足以继续，需要重新分析");
  }
  if (value.unknownTool !== null && value.unknownTool !== undefined) {
    // 规划阶段不写宿主。出现"结果未知步骤"说明现场不干净，必须走执行核对。
    return reject(
      PLANNING_CONTEXT_CODES.notContinuable,
      "存在结果未知的步骤，不能直接继续规划"
    );
  }
  if (!isNonEmptyString(value.savedAt)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "规划现场缺少保存时间");
  }
  if (!isNonEmptyString(value.prompt)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "规划现场缺少原需求");
  }
  if (value.prompt.length > PLANNING_CONTEXT_LIMITS.maxPromptChars) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "原需求超出可继续长度");
  }
  if (
    value.planningGoal !== null &&
    value.planningGoal !== undefined &&
    (typeof value.planningGoal !== "string" ||
      value.planningGoal.length > PLANNING_CONTEXT_LIMITS.maxPlanningGoalChars)
  ) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "规划目标无效");
  }

  if (!Array.isArray(value.confirmations)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "澄清记录无效");
  }
  if (value.confirmations.length > PLANNING_CONTEXT_LIMITS.maxConfirmations) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "澄清记录过多");
  }
  const confirmations: PlanningContextConfirmation[] = [];
  for (const item of value.confirmations) {
    if (!isPlainObject(item) || !isNonEmptyString(item.questionId)) {
      return reject(PLANNING_CONTEXT_CODES.invalid, "澄清记录缺少问题标识");
    }
    if (typeof item.answerLabel !== "string") {
      return reject(PLANNING_CONTEXT_CODES.invalid, "澄清记录缺少回答");
    }
    if (
      item.label !== null &&
      item.label !== undefined &&
      (typeof item.label !== "string" ||
        item.label.length > PLANNING_CONTEXT_LIMITS.maxLabelChars)
    ) {
      return reject(PLANNING_CONTEXT_CODES.invalid, "澄清记录字段名无效");
    }
    confirmations.push({
      questionId: item.questionId.trim(),
      label: (item.label as string | null | undefined) ?? null,
      answerLabel: item.answerLabel
    });
  }

  const target = value.target;
  if (!isPlainObject(target)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "缺少工作簿与授权范围身份");
  }
  if (!isNonEmptyString(target.identity)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "缺少工作簿身份依据");
  }
  if (!TARGET_IDENTITY_PATTERN.test(target.identity)) {
    // 只认既有命名空间。未保存工作簿会退化成 `name:` 兜底，不能启用断点续跑。
    return reject(PLANNING_CONTEXT_CODES.invalid, "工作簿身份不是可核对的命名空间身份");
  }
  if (
    target.evidenceKey !== null &&
    target.evidenceKey !== undefined &&
    typeof target.evidenceKey !== "string"
  ) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "授权证据键类型无效");
  }
  const targetSheets = validateStringList(
    target.sheetNames,
    PLANNING_CONTEXT_LIMITS.maxSheetNames,
    "授权成员"
  );
  if (!targetSheets.ok) {
    return reject(PLANNING_CONTEXT_CODES.invalid, targetSheets.detail);
  }
  if (targetSheets.list.length === 0) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "授权成员为空");
  }

  const authorization = value.authorization;
  if (!isPlainObject(authorization)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "缺少授权版本");
  }
  const authSheets = validateStringList(
    authorization.sheetNames,
    PLANNING_CONTEXT_LIMITS.maxSheetNames,
    "授权成员"
  );
  if (!authSheets.ok) {
    return reject(PLANNING_CONTEXT_CODES.invalid, authSheets.detail);
  }

  const fingerprintSheets = validateStringList(
    value.sourceFingerprintSheets,
    PLANNING_CONTEXT_LIMITS.maxSheetNames,
    "来源表清单"
  );
  if (!fingerprintSheets.ok) {
    return reject(PLANNING_CONTEXT_CODES.invalid, fingerprintSheets.detail);
  }
  if (value.sourceFingerprint !== null && value.sourceFingerprint !== undefined) {
    if (!isNonEmptyString(value.sourceFingerprint)) {
      return reject(PLANNING_CONTEXT_CODES.invalid, "来源指纹无效");
    }
  }
  if (
    (value.sourceFingerprint === null || value.sourceFingerprint === undefined) &&
    fingerprintSheets.list.length > 0
  ) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "有来源表却没有来源指纹");
  }

  if (!Array.isArray(value.resultRefs)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "结果引用无效");
  }
  if (value.resultRefs.length > PLANNING_CONTEXT_LIMITS.maxResultRefs) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "结果引用过多");
  }
  const resultRefs: PlanningContextResultRef[] = [];
  for (const item of value.resultRefs) {
    if (!isPlainObject(item) || typeof item.resultRef !== "string") {
      return reject(PLANNING_CONTEXT_CODES.invalid, "结果引用缺少标识");
    }
    if (!RESULT_REF_PATTERN.test(item.resultRef)) {
      return reject(PLANNING_CONTEXT_CODES.invalid, "结果引用格式无效");
    }
    const fingerprint = item.sourceFingerprint;
    if (
      fingerprint !== null &&
      fingerprint !== undefined &&
      !isNonEmptyString(fingerprint)
    ) {
      return reject(PLANNING_CONTEXT_CODES.invalid, "结果引用的来源指纹无效");
    }
    resultRefs.push({
      resultRef: item.resultRef,
      sourceFingerprint: (fingerprint as string | null | undefined) ?? null
    });
  }

  if (!Array.isArray(value.completedTools)) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "已完成步骤无效");
  }
  if (value.completedTools.length > PLANNING_CONTEXT_LIMITS.maxCompletedTools) {
    return reject(PLANNING_CONTEXT_CODES.invalid, "已完成步骤过多");
  }
  const completedTools: PlanningContextToolStep[] = [];
  for (const item of value.completedTools) {
    if (!isPlainObject(item) || !isNonEmptyString(item.tool)) {
      return reject(PLANNING_CONTEXT_CODES.invalid, "已完成步骤缺少工具名");
    }
    if (item.tool.length > PLANNING_CONTEXT_LIMITS.maxToolNameChars) {
      return reject(PLANNING_CONTEXT_CODES.invalid, "工具名过长");
    }
    if (item.status !== "completed") {
      return reject(PLANNING_CONTEXT_CODES.invalid, "已完成步骤状态无效");
    }
    completedTools.push({ tool: item.tool, status: "completed" });
  }

  let failedTool: { tool: string; errorCode: string } | null = null;
  if (value.failedTool !== null && value.failedTool !== undefined) {
    if (!isPlainObject(value.failedTool) || !isNonEmptyString(value.failedTool.tool)) {
      return reject(PLANNING_CONTEXT_CODES.invalid, "失败步骤无效");
    }
    if (
      typeof value.failedTool.errorCode !== "string" ||
      value.failedTool.errorCode.length > PLANNING_CONTEXT_LIMITS.maxErrorCodeChars
    ) {
      return reject(PLANNING_CONTEXT_CODES.invalid, "失败原因无效");
    }
    failedTool = {
      tool: value.failedTool.tool.trim(),
      errorCode: value.failedTool.errorCode
    };
  }

  const record: PlanningContextRecord = {
    kind: PLANNING_CONTEXT_KIND,
    schemaVersion: PLANNING_CONTEXT_SCHEMA_VERSION,
    taskId: value.taskId,
    turnId: value.turnId,
    executorId: value.executorId,
    taskRevision: revision,
    stage: value.stage,
    resumability: "continue",
    savedAt: value.savedAt,
    prompt: value.prompt,
    planningGoal: (value.planningGoal as string | null | undefined) ?? null,
    confirmations,
    target: {
      identity: target.identity,
      evidenceKey: (target.evidenceKey as string | null | undefined) ?? null,
      sheetNames: targetSheets.list
    },
    authorization: { sheetNames: authSheets.list },
    sourceFingerprint: (value.sourceFingerprint as string | null | undefined) ?? null,
    sourceFingerprintSheets: fingerprintSheets.list,
    resultRefs,
    completedTools,
    failedTool,
    unknownTool: null
  };
  return { ok: true, record };
}

/** 拒绝码 → 用户可见的中文说明。两端共用，避免同一件事两种说法。 */
export function planningContextMessage(code: PlanningContextCode): string {
  switch (code) {
    case PLANNING_CONTEXT_CODES.missing:
      return "这个任务没有保存到足以继续的现场，需要重新发起需求。";
    case PLANNING_CONTEXT_CODES.kindMismatch:
      return "这个任务保存的现场不是可继续的分析现场，需要重新发起需求。";
    case PLANNING_CONTEXT_CODES.unsupportedVersion:
      return "这个任务保存的现场来自不受支持的版本，需要重新发起需求。";
    case PLANNING_CONTEXT_CODES.invalid:
      return "这个任务的现场不完整或已损坏，不能据此继续。";
    case PLANNING_CONTEXT_CODES.sensitive:
      return "这个任务的现场含有不应保存的内容，不能据此继续。";
    case PLANNING_CONTEXT_CODES.taskMismatch:
      return "这个现场属于另一个任务，不能拿来继续当前任务。";
    case PLANNING_CONTEXT_CODES.stageUnsupported:
      return "这个任务停在暂不支持继续的阶段，请重新发起需求。";
    case PLANNING_CONTEXT_CODES.notContinuable:
      return "这个任务的结果还不确定，先核对原表状态，不能直接继续分析。";
    case PLANNING_CONTEXT_CODES.staleRevision:
      return "这个任务的现场版本与当前记录不一致，不能继续。";
    default:
      return "这个任务暂时不能继续，需要重新发起需求。";
  }
}

/** 客户端预检：本地只判"值不值得发这一枪"，结论仍以服务端为准。 */
export function canRequestPlanningContinuation(
  value: unknown,
  expect: PlanningContextExpectation
): { ok: boolean; code: PlanningContextCode | null; message: string } {
  const verdict = validatePlanningContext(value, expect);
  if (verdict.ok) return { ok: true, code: null, message: "" };
  return {
    ok: false,
    code: verdict.code,
    message: planningContextMessage(verdict.code)
  };
}
