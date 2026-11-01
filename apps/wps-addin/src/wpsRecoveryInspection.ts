import capabilities from "../../../config/capabilities.json";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import { parseExplicitA1Rect } from "../../excel-addin/src/planCore";
import type { WpsApplication } from "./wpsHost";
import { wpsReadWorkbook } from "./wpsResultReviewTarget";
import {
  isFormatActionType,
  matrixFingerprintV2,
  wpsFormatEvidenceFingerprint
} from "./wpsPlanExecutor";
import {
  ALL_WPS_FORMAT_FIELDS,
  captureWpsFormatSnapshot,
  type WpsFormatField
} from "./wpsFormatSnapshot";
import { authorizeWpsSheets, normalizeWpsMatrix } from "./wpsWorkbookReader";
import { wpsWorkbookIdentity } from "./wpsExecutionLedger";
import {
  safeDiagnosticErrorName,
  type WpsInspectionOutcome,
  type WpsInspectionPhase,
  type WpsInspectionResult,
  type WpsRecoveryDiagnosticCode
} from "./wpsRecoveryDiagnostics";

/** Reads only registered target rectangles within today's explicit sheet authorization.
 * No PluginStorage, matrices, plan reconstruction or host writes are involved.
 * Formula fingerprints prove formula text, NOT calculation results.
 *
 * 现场指纹的**命名空间与登记端一致**：值动作比值矩阵指纹，格式动作比格式证据指纹
 * （`f1:`）。格式动作不改值，值矩阵指纹在写前写后完全一样 —— 拿它判"在途动作写没写成"
 * 永远判不出结果，只会把已经写完的现场报成冲突。
 *
 * 返回两种形态：
 * - `inspectWpsRecoveryState`：只有结论（`completed | not_started | conflict | unknown`），
 *   既有调用点全部走它，判据一字未改；
 * - `inspectWpsRecoveryStateDetail`：结论 + **原因码 + 阶段 + 动作序号 + 表计数**，
 *   供 DEV 诊断使用。它只多带标量，不读第二遍、不改变任何提前返回的顺序。
 *
 * 为什么要有原因码：2026-10-28 夜间真机上，重开后第一次显示「证据不足」、点一次
 * 「重新核对（只读）」就好，界面只有四个字，`unknown` 的十几条产生路径一条都分不出来。
 */
export function inspectWpsRecoveryState(application: WpsApplication, state: WpsExecutionState,
  selectedIndexes: readonly number[]): WpsInspectionOutcome {
  return inspectWpsRecoveryStateDetail(application, state, selectedIndexes).outcome;
}

export function inspectWpsRecoveryStateDetail(application: WpsApplication, state: WpsExecutionState,
  selectedIndexes: readonly number[]): WpsInspectionResult {
  /*
   * `phase` / `actionIndex` 是**给异常用的**：抛出去以后只有它们能说明"断在哪一步"。
   * 每个会抛的调用前先更新 `phase`，这样 `catch` 里的原因码才有位置信息。
   */
  let phase: WpsInspectionPhase = "authorize";
  let actionIndex: number | undefined;
  let selectedSheetCount = selectedIndexes.length;
  let authorizedSheetCount = 0;

  /** 统一的提前返回：带上当前阶段与已经数出来的计数。 */
  const stop = (
    outcome: WpsInspectionOutcome,
    code: WpsRecoveryDiagnosticCode
  ): WpsInspectionResult => ({
    outcome,
    code,
    phase,
    selectedSheetCount,
    authorizedSheetCount,
    ...(actionIndex === undefined ? {} : { actionIndex })
  });

  try {
    /*
     * 已选表为空：直接给独立原因码，不再去调 `authorizeWpsSheets`。
     *
     * 它不是"额外判据"——`authorizeWpsSheets` 对空集**必然**抛
     * `尚未选择工作表`，结论同样是 `unknown`。差别只在诊断上：
     * "重开后授权范围还没解析出来"与"选了表但授权解析失败"必须分开报，
     * 否则两种完全不同的现场又会被写成同一句「证据不足」。
     */
    if (!selectedSheetCount) return stop("unknown", "authorize_no_selection");
    /*
     * 授权解析单独包一层：刚重开时工作表索引可能还没解析出来，`authorizeWpsSheets`
     * 会直接抛。这是"读不到现场"最常见的一支，必须与别的异常分开报，
     * 否则又会退回到"证据不足"四个字。
     */
    let authorization: ReturnType<typeof authorizeWpsSheets>;
    try {
      authorization = authorizeWpsSheets(application, selectedIndexes);
    } catch (error) {
      return {
        ...stop("unknown", "authorize_threw"),
        ...(safeDiagnosticErrorName(error) ? { errorName: safeDiagnosticErrorName(error)! } : {})
      };
    }
    authorizedSheetCount = authorization.sheets.length;
    phase = "identity";
    const currentIdentity = wpsWorkbookIdentity(authorization.workbookFullName);
    // 身份读不出来（未保存 / 后缀不合规）与"读出来了但不是这一本"是两件事：
    // 前者重试可能就好了，后者重试一百次也一样。
    if (!currentIdentity) return stop("unknown", "workbook_identity_unreadable");
    if (currentIdentity !== state.workbookIdentity) return stop("unknown", "workbook_identity_mismatch");
    phase = "descriptor";
    if (!state.descriptor.actions.length) return stop("unknown", "descriptor_empty");
    if (state.descriptor.actions.length !== state.totalActions) return stop("unknown", "descriptor_action_count");
    phase = "sheets";
    const allowed = new Map(authorization.sheets.map((sheet) => [sheet.name, sheet.index]));
    if (state.authorizedSheetNames.some((name) => !allowed.has(name))) {
      return stop("unknown", "authorized_sheet_not_allowed");
    }
    phase = "actions";
    let cells = 0;
    // Validate EVERY target before any Value2/Formula read, so a later unauthorized
    // action cannot cause partial unauthorized inspection.
    for (const [index, action] of state.descriptor.actions.entries()) {
      actionIndex = index;
      const rect = parseExplicitA1Rect(action.range);
      if (!allowed.has(action.sheet)) return stop("unknown", "action_sheet_not_allowed");
      if (!state.authorizedSheetNames.includes(action.sheet)) return stop("unknown", "action_sheet_not_authorized");
      if (action.actionIndex !== index) return stop("unknown", "action_index_mismatch");
      if (!rect) return stop("unknown", "action_range_unparsable");
      if (rect.endRow - rect.startRow + 1 !== action.rowCount ||
          rect.endColumn - rect.startColumn + 1 !== action.columnCount) {
        return stop("unknown", "action_range_shape_mismatch");
      }
      if (action.cellCount !== action.rowCount * action.columnCount) {
        return stop("unknown", "action_cell_count_mismatch");
      }
      if (!action.expectedFingerprint) return stop("unknown", "action_fingerprint_missing");
      cells += action.cellCount;
    }
    if (cells > capabilities.wpsWrite.maxCells) return stop("unknown", "cell_budget_exceeded");
    actionIndex = undefined;
    let inFlight: "completed" | "not_started" | null = null;
    for (const action of state.descriptor.actions) {
      actionIndex = action.actionIndex;
      phase = "read";
      // 恢复核对的是**原表**：副本激活时读活动簿会拿到执行前的副本数据，结论必错。
      const book = wpsReadWorkbook(application);
      if (!book) return stop("unknown", "workbook_unreadable");
      const sheet = book.Worksheets.Item(allowed.get(action.sheet)!);
      const range = sheet.Range?.(action.range);
      if (!range) return stop("unknown", "host_range_unreadable");
      // 宿主交出的区域形状必须与登记一致；不一致就不能拿它去算指纹。
      if ((range.Rows && range.Rows.Count !== action.rowCount) ||
          (range.Columns && range.Columns.Count !== action.columnCount)) {
        return stop("unknown", "host_range_shape_mismatch");
      }
      let fingerprint: string;
      if (isFormatActionType(action.type)) {
        /*
         * 格式动作：按**登记声明的属性名**采集当前格式快照，用与登记端同一套序列化
         * 口径算格式证据指纹。只读 —— `captureWpsFormatSnapshot` 一个 setter 都不碰。
         * 属性名不可识别、或快照不完整（有格读不到）时判"未知"，不用缺项凑一个哈希。
         */
        phase = "format";
        const declared = action.formatFields ?? [];
        if (!declared.length) return stop("unknown", "format_fields_undeclared");
        const fields = declared.filter((field): field is WpsFormatField =>
          (ALL_WPS_FORMAT_FIELDS as readonly string[]).includes(field));
        if (fields.length !== declared.length) return stop("unknown", "format_fields_unrecognized");
        const snapshot = captureWpsFormatSnapshot(range, action.rowCount, action.columnCount);
        if (!snapshot.captured) return stop("unknown", "format_snapshot_incomplete");
        fingerprint = wpsFormatEvidenceFingerprint(action.type, fields, snapshot);
      } else {
        phase = "values";
        const rawValues = range.Value2;
        const rawFormulas = range.Formula;
        if (rawValues === undefined || rawFormulas === undefined) return stop("unknown", "value_read_undefined");
        const captured = (raw: unknown): boolean => {
          const cell = (value: unknown) => value === null || ["string", "number", "boolean"].includes(typeof value);
          if (!Array.isArray(raw)) return action.cellCount === 1 && cell(raw);
          if (raw.some(Array.isArray)) return raw.length === action.rowCount &&
            raw.every((row) => Array.isArray(row) && row.length === action.columnCount && row.every(cell));
          return raw.length === action.cellCount && raw.every(cell);
        };
        if (!captured(rawValues) || !captured(rawFormulas)) return stop("unknown", "value_shape_unexpected");
        fingerprint = matrixFingerprintV2(normalizeWpsMatrix(rawValues, action.rowCount, action.columnCount),
          normalizeWpsMatrix(rawFormulas, action.rowCount, action.columnCount));
      }
      if (action.actionIndex === state.inFlightActionIndex) {
        if (fingerprint === action.expectedFingerprint) inFlight = "completed";
        else if (fingerprint === action.previewFingerprint) inFlight = "not_started";
        else return stop("conflict", "in_flight_unexpected");
      } else if (fingerprint !== (action.actionIndex < state.nextActionIndex
        ? action.expectedFingerprint : action.previewFingerprint)) {
        return stop("conflict", "completed_action_mismatch");
      }
    }
    actionIndex = undefined;
    const outcome: WpsInspectionOutcome =
      inFlight ?? (state.nextActionIndex === state.totalActions ? "completed" : "not_started");
    return { outcome, phase, selectedSheetCount, authorizedSheetCount };
  } catch (error) {
    /*
     * 以前这里只回一句 `unknown`，异常本身被彻底吞掉 —— 真机上就只剩"证据不足"四个字。
     * 现在把**阶段**和异常 `name`（白名单正则，不收 message：WPS 的异常消息常把完整
     * 文件路径带出来）带出去。结论仍是 `unknown`，判据一个字都不放松。
     */
    return {
      outcome: "unknown",
      code: "inspection_threw",
      phase,
      selectedSheetCount,
      authorizedSheetCount,
      ...(actionIndex === undefined ? {} : { actionIndex }),
      ...(safeDiagnosticErrorName(error) ? { errorName: safeDiagnosticErrorName(error)! } : {})
    };
  }
}
