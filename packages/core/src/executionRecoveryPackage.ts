/**
 * 「未完成任务恢复包」的共享协议（Excel / WPS 两端与服务端同一份判据的口径）。
 *
 * 存在理由（2026-10-27，真实 WPS 验收）：
 * 宿主侧 `PluginStorage` 在本机实测中**不能**为"完整退出 WPS 后再打开"提供可靠保证
 * （退出后本机恢复记录读不到），于是"已核对第二步、仍可继续第三步"在真实机器上断链。
 * 恢复包因此改由**本机后端应用数据目录**持久保存：PluginStorage / localStorage 只保留
 * 索引、缓存与兼容读取，不再是跨进程恢复的唯一材料来源。
 *
 * 三条不可放宽的边界：
 *
 * 1. **恢复包不是账本。** 账本（`office_executions`）保存可变执行进度（`nextActionIndex`、
 *    在途动作、租约），恢复包只保存**不可变的原始恢复材料**（原计划动作顺序、授权范围、
 *    来源证据、每个动作的写前值/公式与预期写后指纹）。两份材料各管一件事，不许互相
 *    复制对方的真相。
 * 2. **恢复包不上云、不进模型、不进普通日志/遥测。** 它含单元格原值，属于本机受控材料；
 *    序列化只允许走本机后端接口与 `PluginStorage` 缓存，任何事件、异常消息、访问日志
 *    都不得拼接它的正文。
 * 3. **白名单结构。** `assertExecutionRecoveryPackage` 是唯一的准入闸：未知字段、矩阵
 *    形状不符、指纹不是本项目的 `v2:` 口径，一律拒绝；"多带了几个字段"不算兼容。
 */

import type { CellValue } from "./contracts";
import { formulaWriteMatrixIssue } from "./formulaWriteback";

/** 恢复包格式版本。**读到不认识的版本必须如实说"无法用于恢复"**，不许猜。 */
export const EXECUTION_RECOVERY_PACKAGE_VERSION = 1;

/** 单个任务的恢复包动作上限（与 `wpsWrite.maxExpandedActions` 同量级，超出即拒绝）。 */
export const EXECUTION_RECOVERY_MAX_ACTIONS = 500;

/**
 * 恢复包登记的动作类型。
 *
 * `writeFormulas` 与 `writeValues` 同构（行×列矩阵），差别只在矩阵里装的是**公式文本**。
 * 公式动作的目标矩阵必须是公式文本，**不许**用计算值顶替 —— 两条公式可以算出同一个数，
 * 拿值当目标会把公式写成常量。判据见 `./formulaWriteback`。
 */
export const EXECUTION_RECOVERY_ACTION_TYPES = ["writeValues", "writeFormulas"] as const;

export type ExecutionRecoveryActionType = (typeof EXECUTION_RECOVERY_ACTION_TYPES)[number];

/** 来源证据：结构与 `WpsResumeContext` 同形（本层不反向依赖宿主端类型）。 */
export interface ExecutionRecoverySource {
  sheetName: string;
  range: string;
  rowCount: number;
  columnCount: number;
  previewFingerprint: string;
  valueFingerprint: string;
}

export interface ExecutionRecoverySourceEvidence {
  /** v1 has sources; v2 records a trusted literal-only plan without data dependencies. */
  version: 1 | 2;
  hostVersion: string;
  sources: ExecutionRecoverySource[];
}

/**
 * 单个动作的不可变恢复材料。
 *
 * `values` 是**预期写入的目标值**，`beforeValues` / `beforeFormulas` 是执行前现场：
 * 三者缺一不可 —— 少了目标值就不知道该继续写什么，少了写前值就既不能核对"是否已经
 * 写过"，也不能把原表恢复回去。
 */
export interface ExecutionRecoveryPackageAction {
  actionIndex: number;
  type: ExecutionRecoveryActionType;
  sheet: string;
  range: string;
  rowCount: number;
  columnCount: number;
  cellCount: number;
  /** 写前逻辑指纹（值与公式的 v2 指纹），执行器核对"是否已经写过"的依据。 */
  previewFingerprint: string;
  /** 预期写后指纹，核对"这一步是否已经写成"的依据。 */
  expectedFingerprint: string;
  /**
   * `writeValues` 的目标值矩阵；**公式动作必须为 `null`**。
   * 公式动作带值矩阵 = 拿计算值顶替公式，一律拒绝（不是兼容问题，是写错东西）。
   */
  values?: CellValue[][] | null;
  /** `writeFormulas` 的目标公式矩阵（逐格 `=` 开头）；**值动作必须为 `null`**。 */
  formulas?: string[][] | null;
  beforeValues: CellValue[][];
  beforeFormulas: CellValue[][];
}

export interface ExecutionRecoveryPackage {
  schemaVersion: typeof EXECUTION_RECOVERY_PACKAGE_VERSION;
  host: "wps";
  operationId: string;
  taskId: string;
  /** 宿主身份（WPS 为 `wps-path:` 前缀的小写磁盘完整路径），不按文件名比对。 */
  workbookIdentity: string;
  authorizedSheetNames: string[];
  planId: string;
  planTitle: string;
  /** 计划指纹：由动作顺序与定位/形状/指纹算出，用来证明"包与账本登记的是同一版计划"。 */
  planFingerprint: string;
  hostVersion: string;
  /** 原始执行创建时间（不是"最后一次保存时间"，读取不得刷新它）。 */
  createdAt: string;
  actions: ExecutionRecoveryPackageAction[];
  sourceEvidence: ExecutionRecoverySourceEvidence;
}

const FINGERPRINT = /^v2:\d+:[0-9a-f]{8}$/;
const A1_RANGE = /^\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/;
const OPERATION_ID = /^[A-Za-z0-9][A-Za-z0-9_.:-]{7,99}$/;

function isCell(value: unknown): value is CellValue {
  return value === null || typeof value === "string" || typeof value === "boolean" ||
    (typeof value === "number" && Number.isFinite(value));
}

function isMatrix(value: unknown, rows: number, columns: number): value is CellValue[][] {
  return Array.isArray(value) && value.length === rows &&
    value.every((row) => Array.isArray(row) && row.length === columns && row.every(isCell));
}

/**
 * 计划指纹：对**动作顺序 + 定位 + 形状 + 写前/写后指纹**做稳定摘要。
 *
 * 用 FNV-1a 32 位是刻意的：它只用来发现"包与账本不是同一版计划"，不是密码学认证，
 * 与仓库既有 `v2:` 指纹口径一致。参与计算的字段在两端都能从**同一个入口**得到
 * （客户端从恢复包、服务端/宿主从 `WpsRecoveryDescriptor`），因此不需要额外落盘。
 */
export function executionRecoveryPlanFingerprint(
  actions: readonly { actionIndex: number; type: string; sheet: string; range: string;
    rowCount: number; columnCount: number; cellCount: number;
    previewFingerprint: string; expectedFingerprint?: string | null }[]
): string {
  const parts = actions.map((action) => [
    action.actionIndex, action.type, action.sheet, action.range,
    action.rowCount, action.columnCount, action.cellCount,
    action.previewFingerprint, action.expectedFingerprint ?? null
  ]);
  const text = JSON.stringify(parts);
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return `plan:v1:${actions.length}:${hash.toString(16).padStart(8, "0")}`;
}

function hasOnlyKeys(value: object, keys: readonly string[]): boolean {
  return Object.keys(value).every((key) => keys.includes(key));
}

/**
 * 恢复包准入闸。**未知字段一律拒绝**：这是"白名单"而不是"至少包含"。
 */
export function assertExecutionRecoveryPackage(value: unknown): asserts value is ExecutionRecoveryPackage {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("恢复包不是有效对象。");
  const candidate = value as Record<string, unknown>;
  if (!hasOnlyKeys(candidate, ["schemaVersion", "host", "operationId", "taskId", "workbookIdentity",
    "authorizedSheetNames", "planId", "planTitle", "planFingerprint", "hostVersion", "createdAt",
    "actions", "sourceEvidence"]))
    throw new Error("恢复包含未知字段，已拒绝使用。");
  if (candidate.schemaVersion !== EXECUTION_RECOVERY_PACKAGE_VERSION)
    throw new Error("恢复包格式版本不受支持，已拒绝使用。");
  if (candidate.host !== "wps") throw new Error("恢复包宿主类型不受支持。");
  if (typeof candidate.operationId !== "string" || !OPERATION_ID.test(candidate.operationId) ||
      typeof candidate.taskId !== "string" || !candidate.taskId || candidate.taskId.length > 100 ||
      typeof candidate.workbookIdentity !== "string" || !candidate.workbookIdentity.startsWith("wps-path:") ||
      typeof candidate.planId !== "string" || !candidate.planId || candidate.planId.length > 200 ||
      typeof candidate.planTitle !== "string" || candidate.planTitle.length > 500 ||
      typeof candidate.planFingerprint !== "string" || !/^plan:v1:\d+:[0-9a-f]{8}$/.test(candidate.planFingerprint) ||
      typeof candidate.hostVersion !== "string" || !candidate.hostVersion || candidate.hostVersion.length > 80 ||
      typeof candidate.createdAt !== "string" || !Number.isFinite(Date.parse(candidate.createdAt)))
    throw new Error("恢复包身份或计划字段无效。");
  const names = candidate.authorizedSheetNames;
  if (!Array.isArray(names) || !names.length || names.length > 200 ||
      names.some((name) => typeof name !== "string" || !name || name.length > 31) ||
      new Set(names).size !== names.length)
    throw new Error("恢复包授权范围无效。");

  const evidence = candidate.sourceEvidence as Record<string, unknown> | undefined;
  if (!evidence || typeof evidence !== "object" || !hasOnlyKeys(evidence, ["version", "hostVersion", "sources"]) ||
      ![1,2].includes(evidence.version as number) || evidence.hostVersion !== candidate.hostVersion || !Array.isArray(evidence.sources) ||
      evidence.sources.length > 500 || (evidence.version === 1 ? !evidence.sources.length : evidence.sources.length !== 0))
    throw new Error("恢复包来源证据无效。");
  const sourceKeys = ["sheetName", "range", "rowCount", "columnCount", "previewFingerprint", "valueFingerprint"];
  for (const source of evidence.sources as Record<string, unknown>[]) {
    if (!source || typeof source !== "object" || !hasOnlyKeys(source, sourceKeys) ||
        typeof source.sheetName !== "string" || !source.sheetName || !names.includes(source.sheetName) ||
        typeof source.range !== "string" || !A1_RANGE.test(source.range) ||
        !Number.isInteger(source.rowCount) || (source.rowCount as number) < 1 ||
        !Number.isInteger(source.columnCount) || (source.columnCount as number) < 1 ||
        typeof source.previewFingerprint !== "string" || !FINGERPRINT.test(source.previewFingerprint) ||
        typeof source.valueFingerprint !== "string" || !FINGERPRINT.test(source.valueFingerprint))
      throw new Error("恢复包来源证据含未知字段或无效指纹。");
  }

  const actions = candidate.actions;
  if (!Array.isArray(actions) || !actions.length || actions.length > EXECUTION_RECOVERY_MAX_ACTIONS)
    throw new Error("恢复包动作数量无效。");
  const actionKeys = ["actionIndex", "type", "sheet", "range", "rowCount", "columnCount", "cellCount",
    "previewFingerprint", "expectedFingerprint", "values", "formulas", "beforeValues", "beforeFormulas"];
  for (const [index, action] of (actions as Record<string, unknown>[]).entries()) {
    if (!action || typeof action !== "object" || !hasOnlyKeys(action, actionKeys))
      throw new Error("恢复包动作含未知字段。");
    if (action.actionIndex !== index ||
        !(EXECUTION_RECOVERY_ACTION_TYPES as readonly string[]).includes(String(action.type)) ||
        typeof action.sheet !== "string" || !names.includes(action.sheet) ||
        typeof action.range !== "string" || !A1_RANGE.test(action.range) ||
        !Number.isInteger(action.rowCount) || (action.rowCount as number) < 1 ||
        !Number.isInteger(action.columnCount) || (action.columnCount as number) < 1 ||
        action.cellCount !== (action.rowCount as number) * (action.columnCount as number) ||
        typeof action.previewFingerprint !== "string" || !FINGERPRINT.test(action.previewFingerprint) ||
        typeof action.expectedFingerprint !== "string" || !FINGERPRINT.test(action.expectedFingerprint))
      throw new Error("恢复包动作定位、形状或指纹无效。");
    const rows = action.rowCount as number;
    const columns = action.columnCount as number;
    /*
     * 目标矩阵按类型二选一，缺一不可、互不顶替：
     * - `writeValues` 只认 `values`，带了 `formulas` 就是形状不对；
     * - `writeFormulas` 只认 `formulas`，带了 `values` 说明有人想把计算值当公式写下去。
     *
     * 老包没有 `formulas` 这个键 —— `undefined` 与 `null` 一视同仁，所以历史材料照旧可读。
     */
    if (action.type === "writeValues") {
      if (!isMatrix(action.values, rows, columns) || (action.formulas ?? null) !== null)
        throw new Error("恢复包动作矩阵形状无效。");
    } else {
      if ((action.values ?? null) !== null)
        throw new Error("恢复包公式动作不得携带值矩阵：计算值不能顶替公式。");
      const formulas = action.formulas;
      if (!Array.isArray(formulas) || formulas.length !== rows ||
          formulas.some((row) => !Array.isArray(row) || row.length !== columns))
        throw new Error("恢复包动作矩阵形状无效。");
      const issue = formulaWriteMatrixIssue(formulas as string[][]);
      if (issue)
        throw new Error(`恢复包公式动作第 ${issue.row + 1} 行第 ${issue.column + 1} 列不是公式文本。`);
    }
    for (const key of ["beforeValues", "beforeFormulas"]) {
      if (!isMatrix(action[key], rows, columns)) throw new Error("恢复包动作矩阵形状无效。");
    }
  }
  // 动作不得落在来源表上（源/目标重叠时，写入会改变"来源"，续跑判据自相矛盾）。
  const sheets = new Set((actions as Record<string, unknown>[]).map((action) => action.sheet));
  if ((evidence.sources as Record<string, unknown>[]).some((source) => sheets.has(source.sheetName)))
    throw new Error("恢复包的来源与写入目标重叠。");
  if (candidate.planFingerprint !== executionRecoveryPlanFingerprint(
    (actions as unknown as ExecutionRecoveryPackageAction[])))
    throw new Error("恢复包计划指纹与动作不一致。");
}

export function parseExecutionRecoveryPackage(text: string): ExecutionRecoveryPackage {
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch {
    throw new Error("恢复包正文不是有效 JSON。");
  }
  assertExecutionRecoveryPackage(parsed);
  return parsed;
}

export type ExecutionRecoveryPackageVerdict =
  | { kind: "ok" }
  | { kind: "expired"; detail: string }
  | { kind: "incompatible"; detail: string }
  | { kind: "mismatch"; detail: string };

/**
 * 恢复包对**这一次**恢复请求是否成立。
 *
 * 只回答"材料能不能用"，不回答"现在该不该写" —— 后者仍由执行器的来源/目标/授权复核
 * 与账本认领决定。四类结论必须分开说，不能用一个 `false` 抹平：
 * 过期（等太久）、不兼容（版本/宿主不同）、对不上（身份/计划/授权不符）、可用。
 */
export function executionRecoveryPackageVerdict(input: {
  package: ExecutionRecoveryPackage;
  operationId: string;
  taskId: string;
  workbookIdentity: string;
  authorizedSheetNames: readonly string[];
  hostVersion: string;
  /** 服务端给出的到期时间；缺省时按**包内 createdAt + 保留天数**由调用方算好再传。 */
  expiresAt: string;
  now?: number;
}): ExecutionRecoveryPackageVerdict {
  const now = input.now ?? Date.now();
  const expires = Date.parse(input.expiresAt);
  if (!Number.isFinite(expires)) return { kind: "incompatible", detail: "恢复包到期时间无效。" };
  /*
   * 到期口径三处必须一致：服务端 `execution_recovery.load`（`deadline <= now` 即过期）、
   * 宿主侧 `wpsResumeWindow`（`now >= deadline`）、以及这里。恰好到期那一刻就算过期，
   * 不给自己留"差一毫秒还算有效"的灰区。
   */
  if (now >= expires) return { kind: "expired", detail: "恢复包已超过保留期，不能用于继续执行。" };
  const pkg = input.package;
  if (pkg.operationId !== input.operationId || pkg.taskId !== input.taskId)
    return { kind: "mismatch", detail: "恢复包与这次执行的任务身份不一致。" };
  if (pkg.workbookIdentity !== input.workbookIdentity)
    return { kind: "mismatch", detail: "恢复包属于另一个工作簿，不能套用。" };
  if (JSON.stringify([...pkg.authorizedSheetNames].sort()) !==
      JSON.stringify([...input.authorizedSheetNames].sort()))
    return { kind: "mismatch", detail: "恢复包登记时的授权范围与当前记录不一致。" };
  if (pkg.hostVersion !== input.hostVersion)
    return { kind: "incompatible", detail: "宿主版本已变化，恢复包不再适用。" };
  return { kind: "ok" };
}
