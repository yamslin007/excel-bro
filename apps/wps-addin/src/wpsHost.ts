export type WpsCollection<T> = {
  Count: number;
  Item(index: number | string): T;
};

export type WpsRange = {
  Address?: string | ((...args: unknown[]) => unknown);
  Text?: unknown;
  Value2?: unknown;
  /** 数字格式。读取归一与写入时的临时文本保真使用；写后恢复原格式。 */
  NumberFormat?: unknown;
  Formula?: unknown;
  MergeCells?: unknown;
  Interior?: { Color?: number; Pattern?: unknown; PatternColor?: unknown };
  /**
   * 字体对象。官方文档（open.wps.cn Range）确认 `Range.Font` 存在，成员含
   * `Bold`/`Italic`/`Name`/`Size`/`Color`/`Underline`。**全部标可选**：这是类型
   * 补齐，不等于已实机验证这些成员在目标 WPS 版本可用 —— 用前仍需受控探测。
   */
  Font?: {
    Bold?: unknown;
    Italic?: unknown;
    Name?: unknown;
    Size?: unknown;
    Color?: unknown;
    Underline?: unknown;
  };
  /**
   * 边框集合。官方文档确认 `Range.Borders.Item(xlEdgeBottom)` 返回单边，成员含
   * `LineStyle`/`Weight`/`ColorIndex`。标可选：未实机验证。
   */
  Borders?: {
    Count?: number;
    Item?(index: unknown): {
      LineStyle?: unknown;
      Weight?: unknown;
      Color?: unknown;
      ColorIndex?: unknown;
    };
  };
  /** 水平对齐。官方文档：xlCenter/xlLeft/xlRight/xlJustify/xlDistributed。 */
  HorizontalAlignment?: unknown;
  /** 垂直对齐。官方文档：xlVAlign 常量。 */
  VerticalAlignment?: unknown;
  WrapText?: unknown;
  IndentLevel?: unknown;
  ShrinkToFit?: unknown;
  Row?: number;
  Column?: number;
  RowHeight?: unknown;
  ColumnWidth?: unknown;
  /** 整行 / 整列。`AutoFit` 常通过 `EntireColumn.AutoFit()` / `EntireRow.AutoFit()` 调用。 */
  EntireRow?: WpsRange & { AutoFit?(): void };
  EntireColumn?: WpsRange & { AutoFit?(): void };
  Rows?: { Count: number };
  Columns?: { Count: number };
  Item?(rowIndex: number, columnIndex?: number): WpsRange;
};

export type WpsWorksheet = {
  Name: string;
  Index?: number;
  /**
   * 宿主提供的稳定工作表标识 —— **文档写了但本项目未实机验证**，所以一律标可选，
   * 读不到就回退到"对象引用 + 结构"判据，不得凭空造 ID。
   *
   * - `CodeName`：WPS 基础接口文档写明其**只读**、**改名不影响**、**运行期不可改**，
   *   是桌面 ET 上与表名解耦的标识。
   * - `SId`：WPS 内核分配的**静态** ID（与"同一工作表在不同会话可能不同"的动态
   *   `Id` 区分）。
   */
  CodeName?: unknown;
  SId?: unknown;
  ProtectContents?: boolean;
  UsedRange?: WpsRange;
  /**
   * 取一个矩形区域。
   *
   * **只接受两个参数。** 真宿主（WPS ET）对多余实参不宽容：传第 3 个参数会
   * 直接抛 `Too many parameters.`（2026-09-30 真机证据，曾因假设"真宿主会忽略
   * 多余实参"而把本加载项自有的 `purpose` 标签塞进这里，结构扫描第一步就失败）。
   *
   * "读表结构 vs 读数据"的区分**不能**靠给宿主传标签：测试替身按替身收到的
   * 矩形坐标与 `Value2` 访问面自行记账（`wpsWorkbookReader.test.ts`）。
   */
  Range?(cell1: WpsRange | string, cell2?: WpsRange | string): WpsRange;
  /** 无参数复制到新的未保存工作簿；原生并排审阅仅在实机确认支持时使用。 */
  Copy?(): void;
};

export type WpsNativeWindow = {
  Caption?: string;
  Visible?: boolean;
  Activate?(): void;
  Close?(saveChanges?: boolean): boolean | void;
};

export type WpsNativeWindows = WpsCollection<WpsNativeWindow> & {
  CompareSideBySideWith?(windowName: string): boolean;
  SyncScrollingSideBySide?: boolean;
  BreakSideBySide?(): boolean;
  ResetPositionsSideBySide?(): void;
};

export type WpsWorkbook = {
  Name: string;
  FullName?: string;
  Path?: string;
  ReadOnly?: boolean;
  Worksheets: WpsCollection<WpsWorksheet>;
  ActiveSheet?: WpsWorksheet;
  Windows?: WpsCollection<WpsNativeWindow>;
  Activate?(): void;
  Close?(saveChanges?: boolean): boolean | void;
};

export type WpsTaskPane = {
  ID: string | number;
  Visible: boolean;
  DockPosition?: number;
  Width?: number;
};

export type WpsPluginStorage = {
  getItem(key: string): unknown;
  setItem(key: string, value: unknown): void;
  removeItem?(key: string): void;
};

/**
 * Web 对话框句柄。字段按 WPS 开放平台「Web对话框」文档保守声明：
 * 当前项目只依赖 `Close`，其余全部标可选 —— 未实机验证过的成员不得当成已有能力。
 */
export type WpsWebDialog = {
  ID?: string | number;
  Close?: () => void;
  Show?: () => void;
  Visible?: boolean;
  Url?: string;
  Width?: number;
  Height?: number;
};

/**
 * `ShowDialog(url, caption, width, height, bModal)` —— 参数顺序取自
 * WPS 开放平台 Application.ShowDialog 文档。当前文档的宽高是**逻辑像素**、
 * 返回值为 boolean；保留句柄联合类型只为兼容可能存在的旧版宿主实现。
 */
export type WpsShowDialog = (
  url: string,
  caption: string,
  width: number,
  height: number,
  bModal: boolean,
  hasCaption?: boolean,
  resizeEdge?: number
) => boolean | WpsWebDialog | undefined;

export type WpsApplication = {
  Version?: string;
  /** ET documented main application width in points; do not use a sheet subwindow width. */
  Width?: number;
  Hwnd?: number;
  Workbooks?: WpsCollection<WpsWorkbook>;
  ActiveWorkbook?: WpsWorkbook;
  ActiveSheet?: WpsWorksheet;
  Windows?: WpsNativeWindows;
  Selection?: WpsRange;
  PluginStorage: WpsPluginStorage;
  CreateTaskPane(url: string, title?: string): WpsTaskPane;
  GetTaskPane(id: string | number): WpsTaskPane | undefined;
  /** 可能不存在 —— 是否可用必须实机探测，不能因为类型里写了就认为有。 */
  ShowDialog?: WpsShowDialog;
  CreateWebDialog?: WpsShowDialog;
  GetWebDialog?: (id: string | number) => WpsWebDialog | undefined;
  ribbonUI?: object;
};

export type WpsHostProbe = {
  host: "wps-et";
  workbookName: string;
  workbookPath: string | null;
  activeSheetName: string;
  sheetNames: string[];
  selectionAddress: string | null;
  selectionPreview: string | null;
};

declare global {
  interface Window {
    Application?: WpsApplication;
  }
}

export class WpsHostUnavailableError extends Error {
  constructor(message = "当前页面没有连接到 WPS 表格宿主。") {
    super(message);
    this.name = "WpsHostUnavailableError";
  }
}

export function getWpsApplication(scope: Pick<Window, "Application"> = window): WpsApplication {
  const application = scope.Application;
  if (!application?.PluginStorage || typeof application.CreateTaskPane !== "function") {
    throw new WpsHostUnavailableError();
  }
  return application;
}

function normalizeSelectionPreview(value: unknown): string | null {
  if (value == null || value === "") return null;
  if (Array.isArray(value)) {
    return value
      .slice(0, 3)
      .map((row) => Array.isArray(row) ? row.slice(0, 3).join(" · ") : String(row))
      .join(" / ");
  }
  return String(value);
}

function readRangeAddress(range: WpsRange | undefined): string | null {
  const address = range?.Address;
  if (typeof address === "function") {
    const value = address.call(range);
    return value == null || value === "" ? null : String(value);
  }
  return address == null || address === "" ? null : String(address);
}

export function readWpsHostProbe(application: WpsApplication, options: { includeSelectionPreview?: boolean } = {}): WpsHostProbe {
  const workbook = application.ActiveWorkbook;
  if (!workbook) {
    throw new WpsHostUnavailableError("WPS 表格中当前没有打开的工作簿。");
  }

  const sheetNames: string[] = [];
  for (let index = 1; index <= Number(workbook.Worksheets?.Count ?? 0); index += 1) {
    const sheet = workbook.Worksheets.Item(index);
    if (sheet?.Name) sheetNames.push(String(sheet.Name));
  }

  const activeSheet = application.ActiveSheet ?? workbook.ActiveSheet;
  const selection = options.includeSelectionPreview === false ? undefined : application.Selection;
  const selectionText = selection?.Text ?? selection?.Value2;

  return {
    host: "wps-et",
    workbookName: String(workbook.Name || "未命名工作簿"),
    workbookPath: workbook.FullName ? String(workbook.FullName) : workbook.Path ? String(workbook.Path) : null,
    activeSheetName: activeSheet?.Name ? String(activeSheet.Name) : "未识别",
    sheetNames,
    selectionAddress: readRangeAddress(selection),
    selectionPreview: normalizeSelectionPreview(selectionText)
  };
}
