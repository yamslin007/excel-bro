import capabilities from "../../../config/capabilities.json";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import { parseExplicitA1Rect } from "../../excel-addin/src/planCore";
import type { WpsApplication } from "./wpsHost";
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

/** Reads only registered target rectangles within today's explicit sheet authorization.
 * No PluginStorage, matrices, plan reconstruction or host writes are involved.
 * Formula fingerprints prove formula text, NOT calculation results.
 *
 * 现场指纹的**命名空间与登记端一致**：值动作比值矩阵指纹，格式动作比格式证据指纹
 * （`f1:`）。格式动作不改值，值矩阵指纹在写前写后完全一样 —— 拿它判"在途动作写没写成"
 * 永远判不出结果，只会把已经写完的现场报成冲突。
 */
export function inspectWpsRecoveryState(application: WpsApplication, state: WpsExecutionState,
  selectedIndexes: readonly number[]): "completed" | "not_started" | "conflict" | "unknown" {
  try {
    const authorization = authorizeWpsSheets(application, selectedIndexes);
    if (wpsWorkbookIdentity(authorization.workbookFullName) !== state.workbookIdentity) return "unknown";
    if (!state.descriptor.actions.length || state.descriptor.actions.length !== state.totalActions) return "unknown";
    const allowed = new Map(authorization.sheets.map((sheet) => [sheet.name, sheet.index]));
    if (state.authorizedSheetNames.some((name) => !allowed.has(name))) return "unknown";
    let cells = 0;
    // Validate EVERY target before any Value2/Formula read, so a later unauthorized
    // action cannot cause partial unauthorized inspection.
    for (const [index, action] of state.descriptor.actions.entries()) {
      const rect = parseExplicitA1Rect(action.range);
      if (!allowed.has(action.sheet) || !state.authorizedSheetNames.includes(action.sheet) ||
          action.actionIndex !== index || !rect ||
          rect.endRow - rect.startRow + 1 !== action.rowCount ||
          rect.endColumn - rect.startColumn + 1 !== action.columnCount ||
          action.cellCount !== action.rowCount * action.columnCount || !action.expectedFingerprint) return "unknown";
      cells += action.cellCount;
    }
    if (cells > capabilities.wpsWrite.maxCells) return "unknown";
    let inFlight: "completed" | "not_started" | null = null;
    for (const action of state.descriptor.actions) {
      const sheet = application.ActiveWorkbook!.Worksheets.Item(allowed.get(action.sheet)!);
      const range = sheet.Range?.(action.range);
      if (!range) return "unknown";
      // 宿主交出的区域形状必须与登记一致；不一致就不能拿它去算指纹。
      if ((range.Rows && range.Rows.Count !== action.rowCount) ||
          (range.Columns && range.Columns.Count !== action.columnCount)) return "unknown";
      let fingerprint: string;
      if (isFormatActionType(action.type)) {
        /*
         * 格式动作：按**登记声明的属性名**采集当前格式快照，用与登记端同一套序列化
         * 口径算格式证据指纹。只读 —— `captureWpsFormatSnapshot` 一个 setter 都不碰。
         * 属性名不可识别、或快照不完整（有格读不到）时判"未知"，不用缺项凑一个哈希。
         */
        const declared = action.formatFields ?? [];
        const fields = declared.filter((field): field is WpsFormatField =>
          (ALL_WPS_FORMAT_FIELDS as readonly string[]).includes(field));
        if (!fields.length || fields.length !== declared.length) return "unknown";
        const snapshot = captureWpsFormatSnapshot(range, action.rowCount, action.columnCount);
        if (!snapshot.captured) return "unknown";
        fingerprint = wpsFormatEvidenceFingerprint(action.type, fields, snapshot);
      } else {
        const rawValues = range.Value2;
        const rawFormulas = range.Formula;
        if (rawValues === undefined || rawFormulas === undefined) return "unknown";
        const captured = (raw: unknown): boolean => {
          const cell = (value: unknown) => value === null || ["string", "number", "boolean"].includes(typeof value);
          if (!Array.isArray(raw)) return action.cellCount === 1 && cell(raw);
          if (raw.some(Array.isArray)) return raw.length === action.rowCount &&
            raw.every((row) => Array.isArray(row) && row.length === action.columnCount && row.every(cell));
          return raw.length === action.cellCount && raw.every(cell);
        };
        if (!captured(rawValues) || !captured(rawFormulas)) return "unknown";
        fingerprint = matrixFingerprintV2(normalizeWpsMatrix(rawValues, action.rowCount, action.columnCount),
          normalizeWpsMatrix(rawFormulas, action.rowCount, action.columnCount));
      }
      if (action.actionIndex === state.inFlightActionIndex) {
        if (fingerprint === action.expectedFingerprint) inFlight = "completed";
        else if (fingerprint === action.previewFingerprint) inFlight = "not_started";
        else return "conflict";
      } else if (fingerprint !== (action.actionIndex < state.nextActionIndex
        ? action.expectedFingerprint : action.previewFingerprint)) return "conflict";
    }
    return inFlight ?? (state.nextActionIndex === state.totalActions ? "completed" : "not_started");
  } catch { return "unknown"; }
}
