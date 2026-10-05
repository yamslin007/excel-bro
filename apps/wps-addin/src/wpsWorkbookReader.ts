import capabilities from "../../../config/capabilities.json";
import type { WpsApplication, WpsRange, WpsWorkbook, WpsWorksheet } from "./wpsHost";

export type WpsCellValue = string | number | boolean | null;

export type WpsWorksheetMetadata = {
  index: number;
  name: string;
  isActive: boolean;
};

export type WpsWorkbookMetadata = {
  name: string;
  fullName: string | null;
  sheets: WpsWorksheetMetadata[];
};

export type WpsAuthorizedSheet = {
  index: number;
  name: string;
};

export type WpsReadAuthorization = {
  workbookName: string;
  workbookFullName: string | null;
  workbookPath: string | null;
  sheets: WpsAuthorizedSheet[];
  /** Retaining the original host reference makes workbook switches fail closed. */
  workbookRef: WpsWorkbook;
};

export type WpsSheetReadResult = {
  sheet: WpsAuthorizedSheet;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
  chunksRead: number;
  values: WpsCellValue[][];
};

export type WpsReadProgress = {
  sheet: WpsAuthorizedSheet;
  rowsRead: number;
  totalRows: number;
  chunkIndex: number;
};

export type WpsReadOptions = {
  signal?: AbortSignal;
  chunkRows?: number;
  onProgress?: (progress: WpsReadProgress) => void;
  yieldControl?: () => Promise<void>;
};

export class WpsAuthorizationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "WpsAuthorizationError";
  }
}

export class WpsReadLimitError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "WpsReadLimitError";
  }
}

const WPS_SYSTEM_SHEETS = new Set(["#EB_RULES", "#EB_RULES_BACKUP"]);

export function isWpsSystemSheet(name: string): boolean {
  return WPS_SYSTEM_SHEETS.has(name.trim().toUpperCase());
}

export function workbookFullName(workbook: WpsWorkbook): string | null {
  if (workbook.FullName) return String(workbook.FullName);
  if (workbook.Path) {
    return `${String(workbook.Path).replace(/[\\/]$/, "")}\\${String(workbook.Name)}`;
  }
  return null;
}

function workbookPath(workbook: WpsWorkbook): string | null {
  const path = String(workbook.Path ?? "").trim();
  return path || null;
}

function normalizeCell(value: unknown): WpsCellValue {
  if (value == null || value === "") return value === "" ? "" : null;
  if (typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") return Number.isFinite(value) ? value : String(value);
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

export function normalizeWpsMatrix(
  raw: unknown,
  rowCount: number,
  columnCount: number
): WpsCellValue[][] {
  if (rowCount <= 0 || columnCount <= 0) return [];

  let rows: unknown[][];
  if (Array.isArray(raw) && raw.some(Array.isArray)) {
    rows = raw.map((row) => Array.isArray(row) ? row : [row]);
  } else if (Array.isArray(raw)) {
    rows = rowCount === 1
      ? [raw]
      : columnCount === 1
        ? raw.map((value) => [value])
        : Array.from({ length: rowCount }, (_, rowIndex) =>
            raw.slice(rowIndex * columnCount, (rowIndex + 1) * columnCount)
          );
  } else {
    rows = [[raw]];
  }

  return Array.from({ length: rowCount }, (_, rowIndex) =>
    Array.from({ length: columnCount }, (_, columnIndex) =>
      normalizeCell(rows[rowIndex]?.[columnIndex])
    )
  );
}

export function listWpsWorksheets(application: WpsApplication): WpsWorkbookMetadata {
  const workbook = application.ActiveWorkbook;
  if (!workbook) throw new WpsAuthorizationError("WPS 表格中当前没有打开的工作簿。");

  const activeSheet = application.ActiveSheet ?? workbook.ActiveSheet;
  const sheets: WpsWorksheetMetadata[] = [];
  const count = Number(workbook.Worksheets?.Count ?? 0);
  for (let index = 1; index <= count; index += 1) {
    const sheet = workbook.Worksheets.Item(index);
    if (!sheet?.Name) continue;
    if (isWpsSystemSheet(String(sheet.Name))) continue;
    sheets.push({
      index,
      name: String(sheet.Name),
      isActive: sheet === activeSheet || String(sheet.Name) === String(activeSheet?.Name ?? "")
    });
  }

  return {
    name: String(workbook.Name || "未命名工作簿"),
    fullName: workbookFullName(workbook),
    sheets
  };
}

export function authorizeWpsSheets(
  application: WpsApplication,
  selectedSheetIndexes: readonly number[]
): WpsReadAuthorization {
  const workbook = application.ActiveWorkbook;
  if (!workbook) throw new WpsAuthorizationError("WPS 表格中当前没有打开的工作簿。");
  const uniqueIndexes = [...new Set(selectedSheetIndexes.map(Number))];
  if (uniqueIndexes.length === 0) {
    throw new WpsAuthorizationError("尚未选择工作表；不会自动改为读取当前表或全部工作表。");
  }

  const sheets = uniqueIndexes.map((index) => {
    if (!Number.isInteger(index) || index < 1 || index > Number(workbook.Worksheets.Count)) {
      throw new WpsAuthorizationError(`工作表序号 ${index} 已失效，请重新选择数据范围。`);
    }
    const sheet = workbook.Worksheets.Item(index);
    if (!sheet?.Name) throw new WpsAuthorizationError(`无法识别第 ${index} 张工作表。`);
    if (isWpsSystemSheet(String(sheet.Name))) {
      throw new WpsAuthorizationError(`工作表「${String(sheet.Name)}」是 Excel Bro 内部系统表，不能授权读取或修改。`);
    }
    return { index, name: String(sheet.Name) };
  });

  return {
    workbookName: String(workbook.Name || "未命名工作簿"),
    workbookFullName: workbookFullName(workbook),
    workbookPath: workbookPath(workbook),
    sheets,
    workbookRef: workbook
  };
}

export function assertWpsAuthorization(
  application: WpsApplication,
  authorization: WpsReadAuthorization,
  authorizedSheet: WpsAuthorizedSheet
): WpsWorksheet {
  const workbook = application.ActiveWorkbook;
  if (!workbook) {
    throw new WpsAuthorizationError("读取期间活动工作簿发生变化，已停止读取；请重新选择数据范围。");
  }
  const requiresReferenceMatch = authorization.workbookPath == null;
  if (requiresReferenceMatch && workbook !== authorization.workbookRef) {
    throw new WpsAuthorizationError("未保存工作簿的会话身份发生变化，已停止读取；请重新选择数据范围。");
  }
  if (
    String(workbook.Name || "") !== authorization.workbookName ||
    workbookFullName(workbook) !== authorization.workbookFullName ||
    workbookPath(workbook) !== authorization.workbookPath
  ) {
    throw new WpsAuthorizationError("工作簿身份发生变化，已停止读取；请重新选择数据范围。");
  }

  const sheet = workbook.Worksheets.Item(authorizedSheet.index);
  if (!sheet || String(sheet.Name) !== authorizedSheet.name) {
    throw new WpsAuthorizationError(
      `已授权工作表「${authorizedSheet.name}」被移动、重命名或删除，已停止读取。`
    );
  }
  if (isWpsSystemSheet(String(sheet.Name))) {
    throw new WpsAuthorizationError("Excel Bro 内部系统表不能授权读取或修改。");
  }
  return sheet;
}

function readChunkRange(
  sheet: WpsWorksheet,
  usedRange: WpsRange,
  offsetRows: number,
  chunkRowCount: number,
  columnsToRead: number,
  /**
   * 这次读取的**目的**。
   *
   * 它不只是个标签：取值面的构造方式据此不同 —— 结构扫描只取**第 1 行**
   * （见下面 `first`/`last` 的分支），所以"结构扫描没有读到数据行"是可以
   * 被証伪的事实，而不是调用方的自我声明。
   */
  purpose: "schema" | "values"
): WpsRange {
  if (typeof usedRange.Item !== "function" || typeof sheet.Range !== "function") {
    throw new WpsAuthorizationError("当前 WPS 版本没有提供所需的分块读取接口。");
  }
  /**
   * 结构扫描用 `UsedRange.Item(1, columns)` 取**一整行**，而不是 `Item(row, col)`
   * 那样的单格 —— 后者取到的对象本身就是一个只含一格的区域，拼出来的矩形列数
   * 会退化成 1，表头字段被丢掉。
   *
   * 为什么整行不会越界：`UsedRange` 的 `rowIndex`/`columnIndex` 未必是 1，
   * 但 `Item(1, N)` 走的正是 `UsedRange.Rows.Item(1).Cells(1, N)`，
   * 由宿主自己按**绝对**行列解释，所以从第 1 列起取满宽度是安全的。
   * 它只到第 1 行，不碰数据行。
   */
  const headerScanRows = Math.max(1, Number(capabilities.queryTable.headerScanRows));
  const first = purpose === "schema" && offsetRows === 0
    ? usedRange.Item(1, 1)
    : usedRange.Item(offsetRows + 1, 1);
  const last = purpose === "schema" && offsetRows === 0
    ? usedRange.Item(1, columnsToRead)
    : usedRange.Item(offsetRows + chunkRowCount, columnsToRead);
  /**
   * 第二道防线：即使传参被改坏，结构扫描的终点也不得越过表头扫描行。
   * 这里只做断言，不改写 `first`/`last` —— 取值面由上面两行的构造决定。
   */
  if (purpose === "schema") {
    const requestedEnd = offsetRows + chunkRowCount;
    if (requestedEnd > headerScanRows || offsetRows > headerScanRows) {
      throw new WpsAuthorizationError(
        `结构扫描请求了第 ${requestedEnd} 行，超过表头扫描上限 ${headerScanRows} 行。`
      );
    }
  }
  /**
   * 只传两个参数 —— 真宿主的 `Worksheet.Range` 只接受 `cell1`/`cell2`，
   * 多传任何自有实参都会直接抛 `Too many parameters.`（2026-09-30 真机证据，
   * 曾因假设"真宿主会忽略多余实参"而把 `purpose` 塞进这里，导致结构扫描
   * 在真机上第一步就失败）。
   *
   * "读表结构 vs 读数据"的区分**不靠给宿主传标签**：`purpose` 只在本函数内
   * 决定取值面的构造与断言（见上），测试替身按替身收到的矩形坐标与
   * `Value2` 访问面自行记账。
   */
  return sheet.Range(first, last);
}

function defaultYieldControl(): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, 0));
}

async function readAuthorizedWpsSheetsInternal(
  application: WpsApplication,
  authorization: WpsReadAuthorization,
  options: WpsReadOptions,
  schemaOnly: boolean
): Promise<WpsSheetReadResult[]> {
  const configuredChunkRows = Number(capabilities.wpsRead.chunkRows);
  const requestedChunkRows = Number(options.chunkRows ?? configuredChunkRows);
  const chunkRows = Math.max(1, Math.min(configuredChunkRows, requestedChunkRows));
  const yieldControl = options.yieldControl ?? defaultYieldControl;
  const results: WpsSheetReadResult[] = [];

  /**
   * 标记"当前这次读取的**目的**"。
   *
   * - `"schema"`：只取**表头行**判字段类型/结构。真实取值范围 = 第 1 行 ×
   *   前 `structureColumns` 列，**不覆盖任何数据行**；尺寸另行报告为整块已用区域。
   * - `"values"`：取完整授权数据，用于本地确定性计算。
   *
   * 两者的隐私含义完全不同，所以既要在宿主侧可区分（宿主实现据此收窄取值面），
   * 也要在测试里按**真实 `Value2` 访问范围**取证 —— 只看这个标记是不够的。
   */
  const readPurpose = schemaOnly ? "schema" : "values";
  const headerRows = Math.max(1, Number(capabilities.queryTable.headerScanRows));

  for (const authorizedSheet of authorization.sheets) {
    if (options.signal?.aborted) throw new DOMException("读取已取消。", "AbortError");
    const sheet = assertWpsAuthorization(application, authorization, authorizedSheet);
    const usedRange = sheet.UsedRange;
    if (!usedRange) {
      throw new WpsAuthorizationError(`无法读取工作表「${authorizedSheet.name}」的使用区域。`);
    }

    const rowCount = Math.max(0, Number(usedRange.Rows?.Count ?? 0));
    const columnCount = Math.max(0, Number(usedRange.Columns?.Count ?? 0));
    const cellCount = rowCount * columnCount;
    if (!schemaOnly && rowCount > capabilities.wpsRead.maxRows) {
      throw new WpsReadLimitError(`工作表「${authorizedSheet.name}」有 ${rowCount} 行，超过 WPS 只读上限。`);
    }
    if (!schemaOnly && columnCount > capabilities.wpsRead.maxColumns) {
      throw new WpsReadLimitError(`工作表「${authorizedSheet.name}」有 ${columnCount} 列，超过 WPS 只读上限。`);
    }
    if (!schemaOnly && cellCount > capabilities.wpsRead.maxCells) {
      throw new WpsReadLimitError(`工作表「${authorizedSheet.name}」包含 ${cellCount} 个单元格，超过 WPS 只读上限。`);
    }

    const values: WpsCellValue[][] = [];
    const rowsToRead = schemaOnly ? Math.min(rowCount, headerRows) : rowCount;
    // 结构扫描可以按 `structureColumns` 取宽（表可能很宽，字段也很多）；
    // 数据读取必须严守授权上限。
    const columnsToRead = Math.min(
      columnCount,
      Number(
        schemaOnly ? capabilities.snapshot.structureColumns : capabilities.wpsRead.maxColumns
      )
    );
    let chunksRead = 0;
    for (let offsetRows = 0; offsetRows < rowsToRead; offsetRows += chunkRows) {
      if (options.signal?.aborted) throw new DOMException("读取已取消。", "AbortError");
      const currentSheet = assertWpsAuthorization(application, authorization, authorizedSheet);
      const currentUsedRange = currentSheet.UsedRange;
      if (!currentUsedRange) {
        throw new WpsAuthorizationError("使用区域在读取期间失效，已停止读取。");
      }
      const currentRows = Number(currentUsedRange.Rows?.Count ?? 0);
      const currentColumns = Number(currentUsedRange.Columns?.Count ?? 0);
      if (currentRows !== rowCount || currentColumns !== columnCount) {
        throw new WpsAuthorizationError("工作表尺寸在读取期间发生变化，已停止读取；请重新查询。");
      }

      const currentChunkRows = Math.min(chunkRows, rowsToRead - offsetRows);
      const range = readChunkRange(
        currentSheet,
        currentUsedRange,
        offsetRows,
        currentChunkRows,
        columnsToRead,
        readPurpose
      );
      values.push(...normalizeWpsMatrix(range.Value2, currentChunkRows, columnsToRead));
      chunksRead += 1;
      options.onProgress?.({
        sheet: authorizedSheet,
        rowsRead: offsetRows + currentChunkRows,
        totalRows: rowsToRead,
        chunkIndex: chunksRead
      });
      await yieldControl();
    }

    results.push({
      sheet: authorizedSheet,
      startRow: Number(usedRange.Row ?? 1),
      startColumn: Number(usedRange.Column ?? 1),
      rowCount,
      columnCount,
      chunksRead,
      values
    });
  }

  return results;
}

/** Read only candidate header rows. The reported dimensions remain the full used range. */
export async function readAuthorizedWpsSheetSchemas(
  application: WpsApplication,
  authorization: WpsReadAuthorization,
  options: WpsReadOptions = {}
): Promise<WpsSheetReadResult[]> {
  return readAuthorizedWpsSheetsInternal(application, authorization, options, true);
}

export async function readAuthorizedWpsSheets(
  application: WpsApplication,
  authorization: WpsReadAuthorization,
  options: WpsReadOptions = {}
): Promise<WpsSheetReadResult[]> {
  return readAuthorizedWpsSheetsInternal(application, authorization, options, false);
}
