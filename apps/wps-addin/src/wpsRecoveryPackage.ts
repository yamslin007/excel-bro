/**
 * WPS 未完成任务恢复包：构建、写前落盘、跨进程读回与水合。
 *
 * 为什么不能只靠 `PluginStorage`（2026-10-27 真实 WPS 验收）：完整退出 WPS 再打开后，
 * 本机 `PluginStorage` 里读不到原执行记录，于是"已核对第二步、仍可继续第三步"在真机上
 * 断链。恢复材料因此改存**本机后端应用数据目录**；`PluginStorage` 退化为索引/缓存。
 *
 * 分工（不得混淆）：
 * - **恢复包**（本文件）：不可变的原始材料 —— 原计划动作顺序、授权范围、来源证据、
 *   每一步的目标值与写前值/公式、写前/写后指纹、创建时间与格式版本。
 * - **账本**（`wpsExecutionLedger`）：可变进度 —— `nextActionIndex`、在途动作、租约。
 *   恢复时以账本为准推进游标，绝不把包里的"当时进度"当成现在的进度。
 *
 * 隐私：恢复包含单元格原值，只允许经本机后端接口与宿主缓存流转；不写日志、不进事件、
 * 不进模型上下文，也不拼进异常消息（异常只说状态）。
 */

import { formulaWriteMatrixIssue } from "@excel-bro/core/formulaWriteback";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import {
  ApiRequestError,
  deleteExecutionRecoveryPackage,
  loadExecutionRecoveryPackage,
  saveExecutionRecoveryPackage
} from "../../excel-addin/src/api";
import {
  assertExecutionRecoveryPackage,
  executionRecoveryPackageVerdict,
  executionRecoveryPlanFingerprint,
  parseExecutionRecoveryPackage,
  type ExecutionRecoveryPackage
} from "@excel-bro/core/executionRecoveryPackage";
import { parseExplicitA1Rect } from "../../excel-addin/src/planCore";
import {
  wpsAppliedActionEvidence,
  wpsCheckpointAfterFingerprint,
  type WpsAppliedActionEvidence
} from "./wpsAppliedActions";
import type { WpsApplication } from "./wpsHost";
import { wpsReadWorkbook } from "./wpsResultReviewTarget";
import {
  adoptWpsUndoRecord,
  matrixFingerprintV2,
  wpsUndoRecordFor,
  type WpsUndoRecord
} from "./wpsPlanExecutor";

/** 只有带 `resume` 的登记执行才需要恢复包；没有它本来就不支持精确续跑。 */
export function requiresWpsRecoveryPackage(record: WpsUndoRecord): boolean {
  return Boolean(record.resume);
}

/**
 * 从本机撤销记录构建恢复包。
 *
 * 只接受当前**真正支持精确续跑**的形状：`writeValues` / `writeFormulas` 动作 + 完整来源
 * 证据 + 矩阵齐全。公式动作落盘的是**公式文本**，不是计算值 —— 两者在真机上可以"数对公式错"，
 * 拿值当目标续写会把公式变成常量。格式/结构动作的写前材料是执行期逐动作采集的，
 * 本批仍不承诺跨进程恢复：与其存一份不完整的材料让人误以为能恢复，不如在这里返回
 * `null` 并让界面明说"需要重新规划"。
 */
export function buildWpsRecoveryPackage(record: WpsUndoRecord): ExecutionRecoveryPackage | null {
  if (!record.resume || record.version !== 2) return null;
  if (!record.actions.length ||
      record.actions.some(action => action.type !== "writeValues" && action.type !== "writeFormulas")) return null;
  // "宿主这次没交出 Value2" 的动作不能冒充有写前值：缺了它既判不出是否写过，也撤不回。
  if (record.actions.some(action => action.beforeValuesCaptured === false)) return null;
  /*
   * 公式动作的写后矩阵必须仍是公式文本。记录里若只剩计算值（或被人替换成常量），
   * 写下去就是把公式变成常量 —— 哪怕算出来的数一模一样。判据与续跑闸共用
   * `@excel-bro/core/formulaWriteback`。
   */
  for (const action of record.actions) {
    if (action.type !== "writeFormulas") continue;
    if (formulaWriteMatrixIssue(action.values)) return null;
  }
  const packageJson: ExecutionRecoveryPackage = {
    schemaVersion: 1,
    host: "wps",
    operationId: record.executionId,
    taskId: record.resume.taskId,
    workbookIdentity: `wps-path:${record.workbookIdentity}`,
    authorizedSheetNames: [...record.authorizedSheetNames],
    planId: record.planId,
    planTitle: record.planTitle,
    planFingerprint: executionRecoveryPlanFingerprint(record.actions.map((action, index) => ({
      actionIndex: index,
      type: action.type,
      sheet: action.sheetName,
      range: action.range,
      rowCount: action.rowCount,
      columnCount: action.columnCount,
      cellCount: action.cellCount,
      previewFingerprint: action.beforeFingerprint,
      // 计划的目标指纹只能由**目标值**算出（`afterFingerprint` 是写后回读的现场，
      // 崩溃时可能还没有）。用回读值顶替会让恢复包与账本登记的计划对不上。
      expectedFingerprint: matrixFingerprintV2(action.values, action.values)
    }))),
    hostVersion: record.resume.context.hostVersion,
    createdAt: record.createdAt,
    actions: record.actions.map((action, index) => {
      const formula = action.type === "writeFormulas";
      return {
        actionIndex: index,
        type: formula ? ("writeFormulas" as const) : ("writeValues" as const),
        sheet: action.sheetName,
        range: action.range,
        rowCount: action.rowCount,
        columnCount: action.columnCount,
        cellCount: action.cellCount,
        previewFingerprint: action.beforeFingerprint,
        // 写后指纹只能由**目标矩阵**算出：恢复时用它证明"这一步写成没有"。
        // `afterFingerprint` 是写后回读的现场，崩溃时可能还没有，不能拿来顶替登记的目标。
        // 公式动作的目标矩阵就是公式文本，指纹口径（formula-first）与账本登记一致。
        expectedFingerprint: matrixFingerprintV2(action.values, action.values),
        values: formula ? null : action.values.map(row => [...row]),
        formulas: formula ? (action.values as string[][]).map(row => [...row]) : null,
        beforeValues: action.beforeValues.map(row => [...row]),
        beforeFormulas: action.beforeFormulas.map(row => [...row])
      };
    }),
    sourceEvidence: {
      version: record.resume.context.version,
      hostVersion: record.resume.context.hostVersion,
      sources: record.resume.context.sources.map(source => ({ ...source }))
    }
  };
  // 自校验：构建出来就必须过协议闸，别等读回时才在另一个进程里发现形状不对。
  assertExecutionRecoveryPackage(packageJson);
  return packageJson;
}

/**
 * 首次宿主写入**之前**落盘，并做读回校验。
 *
 * 失败必须抛出，让调用方中止执行 —— "只存内存也继续写"会在真实机器上留下
 * "写了一半、却没有任何可核对材料"的现场。读回比对的是**逐字节相同**：
 * 服务端说保存成功不等于磁盘上的就是这一份。
 */
export async function persistWpsRecoveryPackage(record: WpsUndoRecord): Promise<{ expiresAt: string }> {
  const packageJson = buildWpsRecoveryPackage(record);
  if (!packageJson) throw new Error("这次执行没有可持久化的恢复材料，已停止写入。");
  const text = JSON.stringify(packageJson);
  const saved = await saveExecutionRecoveryPackage(record.executionId, text);
  const readBack = await loadExecutionRecoveryPackage(record.executionId);
  if (!readBack || readBack.packageJson !== text || readBack.digest !== saved.digest)
    throw new Error("恢复包落盘后读回不一致，已停止写入；原表未被修改。");
  return { expiresAt: readBack.expiresAt };
}

/**
 * 终态回收：任务真的结束（撤销成功、原表已回到执行前）后删掉本机恢复材料。
 *
 * 只在**面板**的撤销收尾处调用（执行器不碰网络）。失败不上抛：回收失败最多让材料按
 * 保留期自然到期，绝不能因为"删不掉"而把已经成功的撤销说成失败，更不能因此重放写入。
 */
export function retireWpsRecoveryPackage(executionId: string): void {
  void deleteExecutionRecoveryPackage(executionId).catch(() => {
    /* 材料留在盘上，随保留期到期由服务端清理。 */
  });
}

/** 账本里与"这次到底写了哪些动作"有关的三件事。 */
type WpsLedgerProgressFacts = {
  nextActionIndex: number;
  inFlightActionIndex?: number | null;
  checkpoint?: Record<string, unknown> | null;
};

function wpsLedgerProgress(
  facts: WpsLedgerProgressFacts,
  action: {
    actionIndex: number;
    rowCount: number;
    sheetName: string;
    range: string;
    /** 登记的**目标**指纹。水合时来自恢复包；给已存在的记录补进度时不传。 */
    expectedFingerprint?: string | null;
  }
): WpsAppliedActionEvidence {
  return wpsAppliedActionEvidence(
    {
      nextActionIndex: facts.nextActionIndex,
      inFlightActionIndex: facts.inFlightActionIndex ?? null,
      // 归属按"表名 + 矩形"认，不按"离谁近"猜：`checkpoint` 只带定位与指纹。
      checkpointAfterFingerprint: wpsCheckpointAfterFingerprint(facts.checkpoint, action)
    },
    {
      actionIndex: action.actionIndex,
      rowCount: action.rowCount,
      expectedFingerprint: action.expectedFingerprint
    }
  );
}

/**
 * 把**账本进度**重新投影到已经存在的宿主缓存记录上（只补不撤）。
 *
 * 为什么要重投影（2026-11-02 WPS 公式 B 真机）：缓存是在重开那一刻按当时的游标建的。
 * 之后用户点「核对」、账本把丢失的完成回执补上（`nextActionIndex` 前进），缓存没有任何
 * 人更新 —— 撤销于是继续以为"那一步没写过"，跳过它、把"漏了一格"报成"已全部恢复"。
 *
 * **只补不撤**：执行器在进入 setter 之前就把 `appliedRows` 记成了"尝试写入的行数"，
 * 那是执行期事实；账本这边只负责补上"完成回执缺失但确实写了"的那一步。
 * 若反过来允许把 `appliedRows` 抹回 0，一次"回执还没送到"的时序就会让撤销悄悄漏掉
 * 一格 —— 正是这个缺陷本身。
 */
export function wpsRecordWithLedgerProgress(
  record: WpsUndoRecord,
  state: WpsExecutionState
): WpsUndoRecord {
  const facts: WpsLedgerProgressFacts = {
    nextActionIndex: state.nextActionIndex,
    inFlightActionIndex: state.inFlightActionIndex ?? null,
    checkpoint: state.checkpoint
  };
  let changed = false;
  const actions = record.actions.map((action) => {
    /*
     * 只补"矩阵类值动作"：恢复包只装 `writeValues` / `writeFormulas`（`buildWpsRecoveryPackage`
     * 会拒绝别的形状），所以只有这两种可能来自"重开后水合出来的缓存"。
     * 格式动作的"已写入"由逐格格式快照证明，`clearRange` 不进恢复包、执行期已按
     * 实际影响记进记录 —— 两者都不在这里补算。
     */
    if (action.type !== "writeValues" && action.type !== "writeFormulas") return action;
    const evidence = wpsLedgerProgress(facts, action);
    const appliedRows = Math.max(action.appliedRows, evidence.appliedRows);
    // 写后指纹只补空缺：已有的是执行期回读的真实现场，不被"账本里的登记值"顶替。
    const afterFingerprint = action.afterFingerprint ?? evidence.afterFingerprint;
    if (appliedRows === action.appliedRows && afterFingerprint === action.afterFingerprint) return action;
    changed = true;
    return { ...action, appliedRows, afterFingerprint };
  });
  return changed ? { ...record, actions } : record;
}

function sheetIndexByName(application: WpsApplication, name: string): number {
  const workbook = wpsReadWorkbook(application);
  const count = Number(workbook?.Worksheets?.Count ?? 0);
  for (let index = 1; index <= count; index += 1) {
    if (String(workbook!.Worksheets.Item(index)?.Name) === name) return index;
  }
  throw new Error(`工作表「${name}」当前不存在，不能用这份恢复材料继续。`);
}

export type WpsRecoveryMaterialResult =
  | { kind: "none" }
  | { kind: "host" }
  | { kind: "package"; record: WpsUndoRecord; expiresAt: string }
  | { kind: "unavailable"; detail: string };

/**
 * 材料不可用时给用户看的原因，**按状态分开**。
 *
 * 以前这里直接把底层异常文本透传出去，于是"过期""损坏""根本没有"在界面上退化成同一种
 * 说法（甚至落到"找不到恢复记录"）。到期和缺失是两件不同的事：一个还能靠重新规划补救、
 * 一个说明这份任务的材料已经不可信。所以按状态给固定措辞，不猜、不合并。
 */
const WPS_RECOVERY_FAILURE_TEXT: Record<string, string> = {
  expired: "本机恢复包已超过保留期，不能继续执行。",
  corrupt: "本机恢复材料已损坏，不能继续执行。",
  missing: "本机没有这次执行的恢复包。",
  conflict: "本机恢复材料与账本登记不一致。",
  quota: "本机恢复材料容量不足，无法保存或读取。"
};

const WPS_RECOVERY_FAILURE_CODES: Record<string, string> = {
  RECOVERY_PACKAGE_EXPIRED: "expired",
  RECOVERY_PACKAGE_CORRUPT: "corrupt",
  RECOVERY_PACKAGE_MISSING: "missing",
  RECOVERY_PACKAGE_CONFLICT: "conflict",
  RECOVERY_PACKAGE_INVALID: "conflict",
  RECOVERY_PACKAGE_QUOTA: "quota"
};

function recoveryMaterialFailure(error: unknown): string {
  const code = error instanceof ApiRequestError ? error.code : "";
  const marker = typeof (error as { recoveryState?: unknown })?.recoveryState === "string"
    ? String((error as { recoveryState?: unknown }).recoveryState)
    : "";
  const state = WPS_RECOVERY_FAILURE_CODES[code] ?? marker;
  if (state && WPS_RECOVERY_FAILURE_TEXT[state]) return WPS_RECOVERY_FAILURE_TEXT[state];
  return error instanceof Error ? error.message : "本机恢复材料不可用。";
}

/**
 * 取回本机恢复记录：宿主缓存优先，缺失时从**持久恢复包**水合。
 *
 * 水合把包里的不可变材料写回宿主缓存，于是既有的 `prepareWpsResume` /
 * `prepareWpsCompletedReview` / `inspectWpsUndoRecord` / `undoWpsExecution`
 * 全部按原样工作 —— 不另开一套"从包里读"的并行实现（两套判据必然漂移）。
 *
 * 进度不取自恢复包：`appliedRows` 与 `state` 一律按**账本**给出的进度事实推导 ——
 * 游标 + 在途动作的写后证据（见 `wpsAppliedActions.ts`），不按包里的"当时写到第几步"。
 * 完整退出未保存时原表可能已回到旧值，因此这里只描述"账本认为写了哪些"，真正的现场
 * 核对仍由 `inspectWpsRecoveryState` 读宿主完成。
 */
export async function ensureWpsRecoveryMaterial(
  application: WpsApplication,
  state: WpsExecutionState,
  options: { load?: typeof loadExecutionRecoveryPackage } = {}
): Promise<WpsRecoveryMaterialResult> {
  const existing = wpsUndoRecordFor(application, state.operationId);
  const load = options.load ?? loadExecutionRecoveryPackage;
  let material;
  try {
    material = await load(state.operationId);
  } catch (error) {
    return { kind: "unavailable", detail: recoveryMaterialFailure(error) };
  }
  if (!material) {
    return existing ? { kind: "host" } : { kind: "none" };
  }
  let packageJson: ExecutionRecoveryPackage;
  try {
    packageJson = parseExecutionRecoveryPackage(material.packageJson);
  } catch {
    return { kind: "unavailable", detail: WPS_RECOVERY_FAILURE_TEXT.corrupt };
  }
  const verdict = executionRecoveryPackageVerdict({
    package: packageJson,
    operationId: state.operationId,
    taskId: state.taskId,
    workbookIdentity: state.workbookIdentity,
    authorizedSheetNames: state.authorizedSheetNames,
    hostVersion: String(application.Version ?? "").trim(),
    expiresAt: material.expiresAt
  });
  if (verdict.kind !== "ok") return { kind: "unavailable", detail: verdict.detail };
  // 包里的计划必须与账本登记的是同一版：指纹不同说明材料与任务不是一对。
  const registered = executionRecoveryPlanFingerprint(state.descriptor.actions.map((action, index) => ({
    actionIndex: index,
    type: action.type,
    sheet: action.sheet,
    range: action.range,
    rowCount: action.rowCount,
    columnCount: action.columnCount,
    cellCount: action.cellCount,
    previewFingerprint: action.previewFingerprint,
    expectedFingerprint: action.expectedFingerprint ?? null
  })));
  if (registered !== packageJson.planFingerprint || packageJson.actions.length !== state.totalActions)
    return { kind: "unavailable", detail: "恢复包与账本登记的计划不是同一版。" };

  if (existing) {
    // 宿主缓存还在：补上"持久材料的到期时间"，让长会话/跨天也能续跑。
    // 这不是延长有效期 —— 时间点来自服务端记录的原始保留期。
    //
    // 顺带把**账本进度**重新投影一次（2026-11-02 真机）：这份缓存可能是重开那一刻
    // 按"当时"的游标建起来的，之后用户点了核对、账本 `nextActionIndex` 前进了，
    // 缓存却没人更新 —— 于是撤销仍以为"第二步没写过"，跳过它并把漏掉一格报成成功。
    const refreshed = wpsRecordWithLedgerProgress(existing, state);
    const next =
      refreshed.resumeExpiresAt !== material.expiresAt
        ? { ...refreshed, resumeExpiresAt: material.expiresAt }
        : refreshed;
    if (next !== existing) adoptWpsUndoRecord(application, next);
    return { kind: "host" };
  }
  const actions = packageJson.actions.map(action => {
    const rect = parseExplicitA1Rect(action.range);
    if (!rect) throw new Error(`恢复包里的区域「${action.range}」不是有效矩形。`);
    /*
     * 水合时把"目标矩阵"放回 `values` 这一个字段里：`prepareWpsResume` 只读 `values`
     * 并按 `type` 决定它装的是值还是公式。公式动作的 `values` 在包里是 `null`、
     * 公式在 `formulas` 里 —— 这里如果忘了搬，续跑就会拿空矩阵去写。
     */
    const target = action.type === "writeFormulas" ? action.formulas : action.values;
    if (!target) throw new Error(`恢复包动作 ${action.actionIndex} 缺少目标矩阵。`);
    const evidence = wpsLedgerProgress(
      {
        nextActionIndex: state.nextActionIndex,
        inFlightActionIndex: state.inFlightActionIndex ?? null,
        checkpoint: state.checkpoint
      },
      { actionIndex: action.actionIndex, rowCount: action.rowCount, sheetName: action.sheet, range: action.range,
        expectedFingerprint: action.expectedFingerprint }
    );
    return {
      actionIndex: action.actionIndex,
      type: action.type === "writeFormulas" ? ("writeFormulas" as const) : ("writeValues" as const),
      sheetIndex: sheetIndexByName(application, action.sheet),
      sheetName: action.sheet,
      range: action.range,
      rowCount: action.rowCount,
      columnCount: action.columnCount,
      cellCount: action.cellCount,
      values: target.map(row => [...row]),
      preview: target.slice(0, 5).map(row => row.slice(0, 5)),
      beforeValues: action.beforeValues.map(row => [...row]),
      beforeValuesCaptured: true,
      beforeFormulas: action.beforeFormulas.map(row => [...row]),
      previewFingerprint: action.previewFingerprint,
      beforeFingerprint: action.previewFingerprint,
      expectedFingerprint: action.expectedFingerprint,
      /*
       * 进度取自**账本**，但账本要比"游标"多一点：第 N 步写完、完成回执没落时，
       * 游标仍停在 N、`inFlightActionIndex === N`，而账本 `checkpoint` 里留着这一步的
       * 写后指纹。只按游标算就会把这一步当成"没写过"，撤销随即漏掉它。
       * 判据只此一份：`wpsAppliedActions.ts`。
       */
      afterFingerprint: evidence.afterFingerprint,
      appliedRows: evidence.appliedRows,
      rect
    };
  });
  /*
   * 跨重启终态闸（2026-10-31）：撤销/放弃之后账本游标与 status 都不变，只有
   * disposition 两列记了时刻。水合若只按游标推 `state`，会把一次已经回退掉的执行
   * 推回 `partial_write_recoverable` —— 于是 `latestWpsUndo` 重新把它当成"还能撤销"
   * 交出来（撤销入口重新开放），`prepareWpsResume` 也放行续跑。所以这里先读
   * disposition：任一终态都水合成 `restored`（`latestWpsUndo` 会过滤它，续跑闸
   * `prepareWpsResume` 也不认它），不再推回可续跑/可撤销的中间态。
   */
  const disposed = Boolean(state.disposition?.undoneAt || state.disposition?.abandonedAt);
  const record: WpsUndoRecord = {
    version: 2,
    executionId: packageJson.operationId,
    planId: packageJson.planId,
    planTitle: packageJson.planTitle,
    createdAt: packageJson.createdAt,
    workbookIdentity: packageJson.workbookIdentity.replace(/^wps-path:/, ""),
    authorizedSheetNames: [...packageJson.authorizedSheetNames],
    state: disposed
      ? "restored"
      : (state.nextActionIndex === state.totalActions && state.inFlightActionIndex == null
        ? "verified" : "partial_write_recoverable"),
    actions: actions as WpsUndoRecord["actions"],
    error: null,
    resume: { taskId: packageJson.taskId, context: packageJson.sourceEvidence },
    resumeExpiresAt: material.expiresAt
  };
  adoptWpsUndoRecord(application, record);
  return { kind: "package", record, expiresAt: material.expiresAt };
}
