/**
 * 测试专用：**有真实状态**的 Office.js（Excel）宿主替身。
 *
 * 为什么需要它：写侧的验收（预览零写入 → 确认 → 写前复核 → 原表写入 → 回读）
 * 以前只在纯函数层被证明，**没有任何一处断言真实宿主被写过**。要拿到"真实宿主
 * 写入"级证据，就必须先有一个真的会记住写入、并且后续读取能读到新状态的宿主。
 *
 * 纪律（与用户本轮要求逐条对应）：
 * - `Excel.run` / `context.sync` / 工作表 / Range 都是替身自己实现的，**不是**
 *   mock 掉执行器或预览流程；被调用的仍是产品代码（`executePlan`,
 *   `stageExcelNativeComparison`, `assertExcelNativeComparisonUnchanged` …）。
 * - 写入是**真的改状态**：`range.values = …` 落在替身的单元格存储上，之后任何
 *   `range.values` / `formulas` / `getUsedRangeOrNullObject()` 都从**同一份状态**
 *   读，不再回放预置结果。
 * - 原表（`original`）与审阅副本（`createWorkbook` 打开的每一本 `copy`）**分账
 *   存储**，写入日志按 store 分开计数，于是"预览期间原表零写入"是可数的。
 *
 * 已知替身与真实宿主的差异（**不得**把替身证据当成实机证据）：
 * 1. 不解析种子 ZIP 里的单元格内容：`Excel.createWorkbook(base64)` 会校验它是
 *    一份合法 stored ZIP，并从 `xl/workbook.xml` 读出真实生成的副本表名；副本的
 *    初始内容由替身**从原表当前状态克隆**得到（真实宿主由 Excel 打开种子）。
 *    两种做法对"副本初始内容 = 预览时的原表"这一点等价，但字节级路径不同。
 * 2. `context.sync()` 是空操作：替身的 getter 直接读当前状态，所以不存在
 *    "load 了但没 sync"的真实批量语义，也不模拟 Excel 的重算时机。
 * 3. 只实现本仓库实际用到的那部分宿主 API；遇到未实现的成员会**抛错**而不是
 *    静默返回 undefined —— 以免替身的缺口被当成产品行为。
 */
import type { CellValue } from "./contracts";

export type TestHostStore = "original" | string;

export interface TestHostWrite {
  store: TestHostStore;
  sheet: string;
  /** 裸 A1，例如 `D1:E3`。 */
  address: string;
  kind: "values" | "formulas" | "clear" | "format";
  cells: number;
}

export interface TestHostOptions {
  /** `Office.context.document.url`；决定 `currentOfficeWorkbookIdentity()`。 */
  url?: string;
  /** 初始工作表。第一张是活动表。 */
  sheets: Array<{
    name: string;
    values: ReadonlyArray<ReadonlyArray<CellValue>>;
  }>;
  activeSheet?: string;
  selectedRange?: string;
  /** 声明支持的 ExcelApi 版本；默认给到 1.18（原生审阅要求 1.13）。 */
  excelApiVersions?: string[];
  excelApiDesktopVersions?: string[];
  readOnly?: boolean;
  /**
   * 模拟"格式采集失败"（矩阵截断/行列数量不足/缺项），供执行前拒绝用例取证。
   *
   * 真实 Excel 里 `getCellProperties` 等几乎不会返回形状不匹配的矩阵，但产品代码
   * 必须对这种"不可完整捕获"的情形 fail-closed。替身默认永远返回正确形状，所以
   * 这里留一个注入点：对 `getCellProperties`/`getRowProperties`/`getColumnProperties`
   * 的返回值做后处理（截断行/列、置空元素），让测试能构造出真实的失败输入。
   */
  corruptFormatRead?: (value: unknown, kind: "cell" | "row" | "column") => unknown;
  /**
   * 模拟"目标区域含条件格式"：`range.conditionalFormats.getCount().value` 返回该值。
   *
   * 兼容路径（宿主无 ExcelApi 1.19）下条件格式支配显示、且可能覆盖本次写入的显示效果，
   * 生产代码必须在写入前明确拒绝 —— 这里给测试一个构造该情形的注入点。
   */
  conditionalFormatCount?: number;
  /** 模拟"目标区域含合并单元格"：`range.getMergedAreasOrNullObject()` 返回非空对象。 */
  mergedAreas?: boolean;
  /**
   * 模拟"目标区域含数据验证"：这些单元格地址（如 `"A1"`、`"B2:C3"`）上附着数据验证。
   *
   * 兼容路径（宿主无 ExcelApi 1.19）的恢复要靠 `clear(Formats)`，而它会把数据验证一并
   * 清除（官方 `removeHyperlinks` 的说明反证了这一点），且当前实现既不采集也不恢复
   * 数据验证 ⇒ 生产代码必须在写入前明确拒绝。这里给测试一个构造该情形的注入点。
   */
  dataValidationCells?: string[];
}

/** 写入拦截：`predicate` 命中时抛错（模拟宿主拒绝/部分失败）。 */
export interface TestHostWriteEvent {
  store: TestHostStore;
  sheet: string;
  address: string;
  kind: TestHostWrite["kind"];
}

/**
 * `text`/`format` 是**显示层**信息，只在需要值归一的时候才用（隔离 Python 的
 * 数据登记要拿它还原前导零编号与日期）。替身默认不给，等价于 `General`。
 */
type Cell = {
  value: CellValue;
  formula: string | null;
  text?: string;
  format?: string;
  /**
   * 单元格**命名样式**（真实宿主的 `SettableCellProperties.style`）。
   *
   * 与 `style`（逐属性直接格式）是两回事：命名样式是一层"命名的属性集合"，
   * `clear(Formats)` 会把它退回 `Normal`。不忠实模拟这一点，"非 Normal 样式 +
   * 直接格式覆盖"的恢复就永远测不出来。
   */
  styleName?: string;
  /**
   * 逐格格式（混合格式测试用）。**只存"直接设置在单元格上"的属性** ——
   * 与真实 `getCellProperties`（1.9）同口径：没设过的属性在这里**不存在**，
   * 读取方要么走 `mergedCellStyle()` 补默认（= 显示值），要么原样看部分集（= 直接格式）。
   * 用"整份默认值"当存储会让"直接设置"与"继承/默认"再也分不开。
   */
  style?: Partial<CellStyle>;
};
/**
 * 逐格格式：字段名与 `contracts.DirectCellFormatValues` / `UndoCellFormat` 同形，
 * 供 `getCellProperties`/`setCellProperties` 读写。
 *
 * **字段清单对齐 `CellPropertiesFormatLoadOptions` 的可读全集**（不是"用到哪算哪"）：
 * 兼容路径的撤销要先 `clear(Formats)`，替身若存不下某个属性，就等于替生产代码
 * "假装它不存在"，测试会恒真通过。
 */
export type CellStyle = {
  /** 填充样式：`"None"` = 无填充（合法空态），`"Solid"` 等 = 有填充。 */  fillPattern: string;
  fillColor: string;
  /** 图案色（斜线/网格图案的线条色）。 */
  fillPatternColor: string;
  fillPatternTintAndShade: number;
  fillTintAndShade: number;
  fontName: string;
  fontSize: number;
  fontBold: boolean;
  fontItalic: boolean;
  fontColor: string;
  fontUnderline: string;
  /** 删除线。 */
  fontStrikethrough: boolean;
  /** 下标 / 上标。 */
  fontSubscript: boolean;
  fontSuperscript: boolean;
  fontTintAndShade: number;
  horizontalAlignment: string;
  verticalAlignment: string;
  wrapText: boolean;
  indentLevel: number;
  /** 缩小字体填充（`clear(Formats)` 会重置它，必须能被捕获）。 */
  shrinkToFit: boolean;
  /** 文字方向（度，`-90`–`90`）。 */
  textOrientation: number;
  /** 阅读顺序。 */
  readingOrder: string;
  /** 自动缩进。 */
  autoIndent: boolean;
  /** 单元格保护：锁定 / 隐藏公式。 */
  locked: boolean;
  formulaHidden: boolean;
  /** 八条边（`Excel.CellBorderCollection`）：含内部边与斜线。 */
  borders: Record<string, { style: string; color: string; weight: string; tintAndShade?: number }>;
};
type SheetModel = {
  reviewRules?: Array<{ id: string; address: string; type: string; custom: {
    rule: { formula: string }; load(): void; format: { borders: { getItem(edge: string): Record<string, unknown> } }
  }; load(): void; delete(): void }>;
  id: string;
  name: string;
  position: number;
  protected: boolean;
  cells: Map<string, Cell>;
  rowHeights: Map<number, number>;
  columnWidths: Map<number, number>;
  /**
   * 样式/条件格式背景（显示值的一部分）。`getDisplayedCellProperties` 用它表达
   * "未直接设置、但显示上受表格样式/条件格式影响"的格式 —— 这是真实 Office.js 里
   * `getCellProperties` 返回 `undefined`、而 `getDisplayedCellProperties` 返回具体值
   * 的那一类场景。缺省为空（无样式背景）。
   */
  displayedDefaults: Partial<CellStyle>;
  /**
   * 本表条件格式规则条数（工作表级状态，缺省回落到宿主选项）。
   *
   * 刻意做成**表状态**而不是只读常量：断言"撤销没有破坏条件格式规则"才有判别力 ——
   * 常量永远等于 1，什么都证明不了。`clear(Formats)` **不会**删除条件格式规则
   * （真实 Excel 亦然），所以它只该被显式删除操作改动。
   */
  conditionalFormatCount?: number;
  /**
   * 附着数据验证的单元格键集合（工作表级状态）。`range.dataValidation.type` 根据
   * 区域内是否有任一格含验证，返回 `"List"`（示例类型）或 `"None"` —— 与真实宿主
   * 对"区域内部分有验证"返回 `"MixedCriteria"` 的语义对齐（有验证 = 非 None）。
   */
  dataValidatedCells?: Set<string>;
};
type BookModel = { id: string; sheets: SheetModel[] };
type WindowModel = {
  name: string;
  book: BookModel;
  activeSheetId: string;
  scrollRow: number;
  scrollColumn: number;
};

export interface TestOfficeHost {
  Excel: unknown;
  Office: unknown;
  /** 每次写入的原始记录，按 store 分账。 */
  writes(): TestHostWrite[];
  originalWrites(): TestHostWrite[];
  copyWrites(): TestHostWrite[];
  resetWrites(): void;
  /** 读单元格当前值（store 默认 original）。 */
  cell(sheet: string, address: string, store?: TestHostStore): CellValue;
  /** 读整块矩阵。 */
  matrix(
    sheet: string,
    address: string,
    store?: TestHostStore
  ): CellValue[][];
  /** 模拟"等待期间外部改了源值/公式"。 */
  setCell(
    sheet: string,
    address: string,
    value: CellValue,
    options?: { store?: TestHostStore; asFormula?: boolean }
  ): void;
  /** 给单元格挂显示文本/数字格式：只有值归一（`normalizeCellValue`）用得上。 */
  setCellDisplay(
    sheet: string,
    address: string,
    display: { text?: string; format?: string },
    options?: { store?: TestHostStore }
  ): void;
  /** 模拟写入被宿主改写（回读不符）。 */
  transformWrites(
    predicate: (event: TestHostWriteEvent) => boolean,
    transform: (matrix: CellValue[][], event: TestHostWriteEvent) => CellValue[][]
  ): void;
  /** 模拟写入被宿主拒绝（部分失败）。 */
  failWrites(
    predicate: (event: TestHostWriteEvent) => boolean,
    message: string
  ): void;
  /** 模拟写入后立刻被外部改回/改坏：只对**下一次**匹配写入生效。 */
  tamperNextWrite(
    predicate: (event: TestHostWriteEvent) => boolean,
    transform: (matrix: CellValue[][], event: TestHostWriteEvent) => CellValue[][]
  ): void;
  /** 换掉文件身份（新的 document 对象，`url` 不同）。 */
  setDocumentUrl(url: string): void;
  /** 已打开的窗口名（第一张是原表窗口）。 */
  windowNames(): string[];
  /** 已创建的副本 store 标识（按打开顺序），可与 `cell`/`matrix` 的 store 参数配合。 */
  copyStores(): string[];
  /** `Excel.createWorkbook` 收到的 base64 种子（按调用顺序）。 */
  createWorkbookCalls(): string[];
  /** 解析种子 ZIP 得到的文件条目，供断言"副本确实是从原表内容生成的"。 */
  seedEntries(index: number): Record<string, string>;
  /** 清空拦截规则与日志，保留单元格状态。 */
  resetInterceptors(): void;
  /** 设置单元格逐格格式（混合格式测试构造初始状态用）。 */
  setCellStyle(sheet: string, address: string, style: Partial<CellStyle>, options?: { store?: TestHostStore }): void;
  /**
   * 设置单元格的**命名样式**（真实宿主的 `SettableCellProperties.style`）。
   * 与 `setCellStyle`（逐属性直接格式）正交：命名样式是一层"属性集合"，
   * `clear(Formats)` 会把它退回 `Normal`。用于构造"非 Normal 样式 + 直接格式覆盖"。
   */
  setCellStyleName(
    sheet: string,
    address: string,
    styleName: string,
    options?: { store?: TestHostStore }
  ): void;
  /** 读单元格的命名样式（`getCellProperties({ style: true })` 的回读口径）。 */
  cellStyleName(sheet: string, address: string, store?: TestHostStore): string | undefined;
  /** 读单元格的数字格式（`range.numberFormat` 的单元格口径，便于断言 clear 是否清掉了它）。 */
  cellNumberFormat(sheet: string, address: string, store?: TestHostStore): string | undefined;
  /** 读本表条件格式规则条数（断言"撤销没动条件格式规则"用）。 */
  conditionalFormatCount(sheet: string, store?: TestHostStore): number;
  /** 显式增删条件格式规则（构造初始状态 / 模拟外部改动）。 */
  setConditionalFormatCount(sheet: string, count: number, store?: TestHostStore): void;
  /** 读本表带数据验证的单元格数量（断言"写入前拒绝"没动数据验证用）。 */
  dataValidationCount(sheet: string, store?: TestHostStore): number;
  /** 显式给一个单元格地址加/摘数据验证（构造初始状态 / 模拟外部改动）。 */
  setDataValidation(sheet: string, address: string, enabled: boolean, store?: TestHostStore): void;
  /** 读单元格逐格格式（断言恢复结果用）。 */
  /**
   * 读单元格的**部分**样式（= 直接设置在单元格上的属性）。
   * 缺省属性**不出现**在此对象里 —— 与 `getCellProperties` 同口径，
   * 用于区分"直接设置成默认值"与"根本没设置"。
   */
  cellStyle(sheet: string, address: string, store?: TestHostStore): Partial<CellStyle> | undefined;
  /**
   * 读单元格的**显示**样式（= 硬默认 ← 样式背景 ← 直接设置，与
   * `getDisplayedCellProperties` 同口径）。
   *
   * 与 `cellStyle`（只含**直接设置**的属性）刻意分开：这样才能分别断言
   * "用户看到的仍是加粗"与"单元格并未被直接写上加粗"—— 把显示一致当成来源语义恢复
   * 正是本轮要拒绝的口径，测试也必须能区分二者。
   */
  cellDisplayedStyle(sheet: string, address: string, store?: TestHostStore): CellStyle;
  /** 读某行行高（断言恢复结果用）。 */
  rowHeight(sheet: string, rowIndex: number, store?: TestHostStore): number;
  /** 读某列列宽（断言恢复结果用）。 */
  columnWidth(sheet: string, columnIndex: number, store?: TestHostStore): number;
  /** 设置某行行高（构造混合初始状态用）。 */
  setRowHeight(sheet: string, rowIndex: number, value: number, store?: TestHostStore): void;
  /** 设置某列列宽（构造混合初始状态用）。 */
  setColumnWidth(sheet: string, columnIndex: number, value: number, store?: TestHostStore): void;
  /**
   * 设置整表的"样式/条件格式背景"（显示值的一部分）。用于构造"未直接设置、但显示上
   * 受表格样式/条件格式影响"的场景 —— 例如样式给了加粗或填充色。`getCellProperties`
   * 对这些格仍返回 `undefined`，`getDisplayedCellProperties` 返回该背景值。
   */
  setDisplayedDefault(
    sheet: string,
    defaults: Partial<{
      fillPattern: string;
      fillColor: string;
      fontBold: boolean;
      fontItalic: boolean;
      fontName: string;
      fontSize: number;
      fontColor: string;
      horizontalAlignment: string;
      verticalAlignment: string;
      wrapText: boolean;
      indentLevel: number;
    }>,
    store?: TestHostStore
  ): void;
}

const COLUMN_LIMIT = 16_384;
const ROW_LIMIT = 1_048_576;

function columnIndex(name: string): number {
  let value = 0;
  for (const character of name.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value - 1;
}

function columnName(index: number): string {
  let value = index + 1;
  let text = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    text = String.fromCharCode(65 + remainder) + text;
    value = Math.floor((value - 1) / 26);
  }
  return text;
}

const cellKey = (row: number, column: number) => `${row},${column}`;

const CELL_PART = /^([A-Z]{1,3})([1-9]\d*)$/;
const COLUMN_PART = /^([A-Z]{1,3})$/;
const ROW_PART = /^([1-9]\d*)$/;

/** 解析裸 A1 / 区域地址（容忍 `$`、工作表前缀、整列整行）。 */
function parseAddress(
  raw: string
): { r0: number; c0: number; r1: number; c1: number } {
  const local = raw.includes("!") ? raw.slice(raw.lastIndexOf("!") + 1) : raw;
  const text = local.replace(/\$/g, "").trim().toUpperCase();
  const [start, end] = text.includes(":") ? text.split(":") : [text, text];
  const startCell = CELL_PART.exec(start);
  const endCell = CELL_PART.exec(end);
  if (startCell && endCell) {
    return rect(
      Number(startCell[2]) - 1,
      columnIndex(startCell[1]),
      Number(endCell[2]) - 1,
      columnIndex(endCell[1])
    );
  }
  const startColumn = COLUMN_PART.exec(start);
  const endColumn = COLUMN_PART.exec(end);
  if (startColumn && endColumn) {
    return rect(0, columnIndex(startColumn[1]), ROW_LIMIT - 1, columnIndex(endColumn[1]));
  }
  const startRow = ROW_PART.exec(start);
  const endRow = ROW_PART.exec(end);
  if (startRow && endRow) {
    return rect(Number(startRow[1]) - 1, 0, Number(endRow[1]) - 1, COLUMN_LIMIT - 1);
  }
  throw new Error(`宿主替身无法解析区域地址「${raw}」`);
}

function rect(
  r0: number,
  c0: number,
  r1: number,
  c1: number
): { r0: number; c0: number; r1: number; c1: number } {
  if (r1 < r0 || c1 < c0) {
    throw new Error(`宿主替身收到倒置的区域地址 ${r0},${c0}:${r1},${c1}`);
  }
  return { r0, c0, r1, c1 };
}

const bareAddress = (r0: number, c0: number, r1: number, c1: number) =>
  r0 === r1 && c0 === c1
    ? `${columnName(c0)}${r0 + 1}`
    : `${columnName(c0)}${r0 + 1}:${columnName(c1)}${r1 + 1}`;

/**
 * `Range.format.borders.getItem(...)` 的入参 → 存储键（`Excel.CellBorderCollection` 的边名）。
 *
 * **必须同时接受数值与字符串**：真实宿主里 `Excel.BorderIndex` 是数值枚举，而
 * `applyFormatAction` 走的是 `"EdgeTop"` / `"InsideHorizontal"` 这类字符串常量 ——
 * 替身两边都要认。存储键统一成 `CellBorderCollection` 的名字（`horizontal`/`vertical`/
 * `diagonalDown`/`diagonalUp`），这样 `getCellProperties` 的 `format.borders` 输出
 * 与生产代码按协议边名读的 `CELL_BORDER_SIDES` 天然对齐。
 *
 * 曾经这里只 switch 数值，字符串入参一律落到 `default: "top"` ⇒ **所有边都写进 `top`**，
 * 而断言只查 `top` 就看不出来。
 */
const BORDER_SIDE_KEYS: Record<string, string> = {
  "0": "top",
  "1": "bottom",
  "2": "left",
  "3": "right",
  "4": "horizontal",
  "5": "vertical",
  "6": "diagonalDown",
  "7": "diagonalUp",
  edgetop: "top",
  edgebottom: "bottom",
  edgeleft: "left",
  edgeright: "right",
  insidehorizontal: "horizontal",
  insidevertical: "vertical",
  diagonaldown: "diagonalDown",
  diagonalup: "diagonalUp",
  top: "top",
  bottom: "bottom",
  left: "left",
  right: "right",
  horizontal: "horizontal",
  vertical: "vertical"
};

function borderSideKey(sideIndex: number | string): string {
  const key = typeof sideIndex === "number" ? String(sideIndex) : String(sideIndex).toLowerCase();
  const mapped = BORDER_SIDE_KEYS[key];
  if (!mapped) throw new Error(`宿主替身：未知的边框位置「${String(sideIndex)}」`);
  return mapped;
}

function defaultCellStyle(): CellStyle {
  return {
    fillPattern: "None",
    fillColor: "",
    fillPatternColor: "",
    fillPatternTintAndShade: 0,
    fillTintAndShade: 0,
    fontName: "Calibri",
    fontSize: 11,
    fontBold: false,
    fontItalic: false,
    fontColor: "#000000",
    fontUnderline: "None",
    fontStrikethrough: false,
    fontSubscript: false,
    fontSuperscript: false,
    fontTintAndShade: 0,
    horizontalAlignment: "General",
    verticalAlignment: "Bottom",
    wrapText: false,
    indentLevel: 0,
    shrinkToFit: false,
    textOrientation: 0,
    readingOrder: "Context",
    autoIndent: false,
    locked: true,
    formulaHidden: false,
    // 八条边：`Excel.CellBorderCollection` 的全集（含内部边与斜线）。
    borders: Object.fromEntries(
      [
        "top",
        "bottom",
        "left",
        "right",
        "horizontal",
        "vertical",
        "diagonalDown",
        "diagonalUp"
      ].map((side) => [side, { style: "None", color: "#000000", weight: "Thin" }])
    )
  };
}

/**
 * 把**部分**样式补成完整样式（= 真实 Excel 生效的显示值）。
 *
 * 存储层只装"直接设置过"的属性；任何需要"看到完整值"的地方（`format.*` 读取、
 * 显示值路径）都先过一次这个函数。少了这一步，`format.font.bold` 之类会读成
 * `undefined`，与真实宿主不符。
 */
function mergedCellStyle(style: Partial<CellStyle> | undefined): CellStyle {
  const base = defaultCellStyle();
  if (!style) return base;
  return {
    ...base,
    ...style,
    borders: { ...base.borders, ...(style.borders ?? {}) }
  };
}

/** 去掉部分样式里"没有直接设置任何东西"的空壳（不留下无意义的 style 对象）。 */
function isEmptyPartialStyle(style: Partial<CellStyle>): boolean {
  return Object.keys(style).every(
    (key) =>
      key === "borders" &&
      Object.keys((style.borders ?? {}) as Record<string, unknown>).length === 0
  ) && (Object.keys(style).length <= 1);
}

/**
 * 计算"显示值"：硬默认（`defaultCellStyle()`）← 样式背景（`displayedDefaults`）← 直接
 * 设置值（`direct`）。直接设置覆盖样式背景，样式背景覆盖硬默认。
 */
function mergeDisplayedStyle(
  displayedDefaults: Partial<CellStyle>,
  direct: Partial<CellStyle> | undefined
): CellStyle {
  const base = defaultCellStyle();
  const merged: CellStyle = {
    ...base,
    ...displayedDefaults,
    ...(direct ?? {}),
    borders: {
      ...base.borders,
      ...(displayedDefaults.borders ?? {}),
      ...(direct?.borders ?? {})
    }
  };
  return merged;
}

/**
 * `setCellProperties` 把一条单元格属性写进 style。`borders` 位于 `format.borders`（与 `CellPropertiesFormat` 一致）。
 *
 * `current` 是**部分**样式：只往里写"入参里确实给了"的字段，其余保持原样（可能仍是"未直接设置"）。
 */
function applyCellProperties(
  current: Partial<CellStyle>,
  properties: { format?: Record<string, unknown>; borders?: Record<string, unknown> }
): Partial<CellStyle> {
  const style: Partial<CellStyle> = {
    ...current,
    borders: { ...(current.borders ?? {}) }
  };
  const format = properties?.format;
  if (format) {
    const fill = format.fill as Record<string, unknown> | undefined;
    const font = format.font as Record<string, unknown> | undefined;
    if (fill) {
      // 忠实模拟真实 Office.js：`fill.pattern: "None"` 清除填充（回到无填充空态），
      // `fill.pattern: "Solid"` + `fill.color` 设置填充色。只写 color 不写 pattern 时
      // 保持原 pattern；只写 pattern 不写 color 时保持原色（但 None 应清空色）。
      if (fill.pattern != null) {
        style.fillPattern = String(fill.pattern);
        if (String(fill.pattern) === "None") style.fillColor = "";
      }
      if (fill.color != null) {
        style.fillColor = String(fill.color);
        // 真实 Excel：设置了填充色即意味着有填充；pattern 未显式给时视为 Solid。
        if (style.fillColor !== "" && (style.fillPattern ?? "None") === "None") {
          style.fillPattern = "Solid";
        }
      }
      if (fill.patternColor != null) style.fillPatternColor = String(fill.patternColor);
      if (fill.patternTintAndShade != null) {
        style.fillPatternTintAndShade = Number(fill.patternTintAndShade);
      }
      if (fill.tintAndShade != null) style.fillTintAndShade = Number(fill.tintAndShade);
    }
    if (font) {
      if (font.name != null) style.fontName = String(font.name);
      if (font.size != null) style.fontSize = Number(font.size);
      if (font.bold != null) style.fontBold = Boolean(font.bold);
      if (font.italic != null) style.fontItalic = Boolean(font.italic);
      if (font.color != null) style.fontColor = String(font.color);
      if (font.underline != null) style.fontUnderline = String(font.underline);
      if (font.strikethrough != null) style.fontStrikethrough = Boolean(font.strikethrough);
      if (font.subscript != null) style.fontSubscript = Boolean(font.subscript);
      if (font.superscript != null) style.fontSuperscript = Boolean(font.superscript);
      if (font.tintAndShade != null) style.fontTintAndShade = Number(font.tintAndShade);
    }
    if (format.horizontalAlignment != null) style.horizontalAlignment = String(format.horizontalAlignment);
    if (format.verticalAlignment != null) style.verticalAlignment = String(format.verticalAlignment);
    if (format.wrapText != null) style.wrapText = Boolean(format.wrapText);
    if (format.indentLevel != null) style.indentLevel = Number(format.indentLevel);
    if (format.shrinkToFit != null) style.shrinkToFit = Boolean(format.shrinkToFit);
    if (format.textOrientation != null) style.textOrientation = Number(format.textOrientation);
    if (format.readingOrder != null) style.readingOrder = String(format.readingOrder);
    if (format.autoIndent != null) style.autoIndent = Boolean(format.autoIndent);
    const protection = format.protection as Record<string, unknown> | undefined;
    if (protection) {
      if (protection.locked != null) style.locked = Boolean(protection.locked);
      if (protection.formulaHidden != null) style.formulaHidden = Boolean(protection.formulaHidden);
    }
    const formatBorders = format.borders as Record<string, Record<string, unknown>> | undefined;
    if (formatBorders) {
      for (const [side, border] of Object.entries(formatBorders)) {
        if (!border) continue;
        const key = borderSideKey(side);
        const existing = style.borders![key] ?? { style: "None", color: "#000000", weight: "Thin" };
        const next: { style: string; color: string; weight: string; tintAndShade?: number } = {
          style: border.style != null ? String(border.style) : existing.style,
          color: border.color != null ? String(border.color) : existing.color,
          weight: border.weight != null ? String(border.weight) : existing.weight
        };
        if (border.tintAndShade != null) next.tintAndShade = Number(border.tintAndShade);
        else if (existing.tintAndShade != null) next.tintAndShade = existing.tintAndShade;
        style.borders![key] = next;
      }
    }
  }
  if (Object.keys(style.borders ?? {}).length === 0) delete style.borders;
  return style;
}

/** 只认 stored（不压缩）ZIP —— 与 `nativeReviewPackage` 的产出格式一致。 */
function readStoredZip(base64: string): Record<string, string> {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  const view = new DataView(bytes.buffer);
  const entries: Record<string, string> = {};
  const decode = (from: number, to: number) =>
    new TextDecoder().decode(bytes.subarray(from, to));
  let at = 0;
  while (view.getUint32(at, true) === 0x04034b50) {
    if (view.getUint16(at + 8, true) !== 0) {
      throw new Error("宿主替身：副本种子不是 stored ZIP");
    }
    const size = view.getUint32(at + 18, true);
    const nameSize = view.getUint16(at + 26, true);
    const extraSize = view.getUint16(at + 28, true);
    const start = at + 30 + nameSize + extraSize;
    entries[decode(at + 30, at + 30 + nameSize)] = decode(start, start + size);
    at = start + size;
  }
  if (view.getUint32(at, true) !== 0x02014b50) {
    throw new Error("宿主替身：副本种子缺少中央目录");
  }
  return entries;
}

export function createOfficeTestHost(options: TestHostOptions): TestOfficeHost {  let sheetSeq = 0;
  let bookSeq = 0;
  let windowSeq = 0;

  const makeSheet = (
    name: string,
    position: number,
    values?: ReadonlyArray<ReadonlyArray<CellValue>>
  ): SheetModel => {
    const sheet: SheetModel = {
      id: `sheet-${(sheetSeq += 1)}`,
      name,
      position,
      protected: false,
      cells: new Map(),
      rowHeights: new Map(),
      columnWidths: new Map(),
      displayedDefaults: {},
      dataValidatedCells: new Set()
    };
    values?.forEach((row, rowIndex) =>
      row.forEach((value, columnIndex) => {
        if (value === null || value === "") return;
        sheet.cells.set(cellKey(rowIndex, columnIndex), {
          value,
          formula: null
        });
      })
    );
    return sheet;
  };

  const originalBook: BookModel = {
    id: `book-${(bookSeq += 1)}`,
    sheets: options.sheets.map((sheet, index) =>
      makeSheet(sheet.name, index, sheet.values)
    )
  };
  const activeSheetName = options.activeSheet ?? options.sheets[0].name;
  const windows: WindowModel[] = [
    {
      name: `${options.url ? options.url.split("/").pop() : "Book1.xlsx"}`,
      book: originalBook,
      activeSheetId:
        originalBook.sheets.find((sheet) => sheet.name === activeSheetName)?.id ??
        originalBook.sheets[0].id,
      scrollRow: 0,
      scrollColumn: 0
    }
  ];
  // 种子数据验证：`dataValidationCells` 里的地址（作用于活动表）标记为"有数据验证"。
  for (const address of options.dataValidationCells ?? []) {
    const bounds = parseAddress(address);
    const target = originalBook.sheets.find((sheet) => sheet.name === activeSheetName) ?? originalBook.sheets[0];
    for (let row = bounds.r0; row <= bounds.r1; row += 1) {
      for (let column = bounds.c0; column <= bounds.c1; column += 1) {
        target.dataValidatedCells?.add(cellKey(row, column));
      }
    }
  }
  let activeWindowName = windows[0].name;
  const selectionAddress = options.selectedRange ?? "A1";

  const writes: TestHostWrite[] = [];
  const seeds: Array<Record<string, string>> = [];
  let writeInterceptors: Array<{
    predicate: (event: TestHostWriteEvent) => boolean;
    fail?: string;
    transform?: (
      matrix: CellValue[][],
      event: TestHostWriteEvent
    ) => CellValue[][];
    once: boolean;
  }> = [];

  /** 替身内部写：不走拦截（供测试自己改状态用）。 */
  const writeCell = (sheet: SheetModel, row: number, column: number, value: CellValue, asFormula = false) => {
    const key = cellKey(row, column);
    const existing = sheet.cells.get(key);
    const formula = asFormula && typeof value === "string" ? value : null;
    /**
     * 真实宿主：写值/公式**不动这个单元格已有的格式**（直接格式、数字格式、命名样式都不动）。
     *
     * 旧替身每次写值都新建 `{value, formula}`，顺手把 `style`/`format` 丢掉 —— 于是
     * "只改值的撤销不碰格式"这条性质在替身里**无法验证**（写值本身就把格式清了）。
     * 所以这里显式保留格式字段。
     */
    const carryOver = (): Pick<Cell, "format" | "style" | "styleName"> => ({
      ...(existing?.format != null ? { format: existing.format } : {}),
      ...(existing?.style != null ? { style: existing.style } : {}),
      ...(existing?.styleName != null ? { styleName: existing.styleName } : {})
    });
    if (value === null || value === "") {
      if (!existing) return;
      // 空值只清**内容**，不删格式（真实 Excel 里"清内容"与"清格式"是两回事）。
      // 完全没有任何格式的格子仍然删掉，保持"空位 = 无 cell"的读取约定不变。
      const carried = carryOver();
      if (Object.keys(carried).length === 0) {
        sheet.cells.delete(key);
        return;
      }
      sheet.cells.set(key, { value: "", formula, ...carried });
      return;
    }
    sheet.cells.set(key, { value, formula, ...carryOver() });
  };

  const matrixOf = (
    sheet: SheetModel,
    rect: { r0: number; c0: number; r1: number; c1: number },
    kind: "values" | "formulas" | "text"
  ): CellValue[][] => {
    const rows: CellValue[][] = [];
    for (let row = rect.r0; row <= rect.r1; row += 1) {
      const cells: CellValue[] = [];
      for (let column = rect.c0; column <= rect.c1; column += 1) {
        const cell = sheet.cells.get(cellKey(row, column));
        if (!cell) {
          cells.push(kind === "text" ? "" : "");
          continue;
        }
        if (kind === "formulas") cells.push(cell.formula ?? cell.value);
        else if (kind === "text") cells.push(cell.text ?? String(cell.value));
        else cells.push(cell.value);
      }
      rows.push(cells);
    }
    return rows;
  };

  /** 写日志 + 拦截。返回真正要落盘的矩阵。 */
  const commitWrite = (
    store: TestHostStore,
    sheet: SheetModel,
    rect: { r0: number; c0: number; r1: number; c1: number },
    matrix: CellValue[][],
    kind: TestHostWrite["kind"]
  ): CellValue[][] => {
    const event: TestHostWriteEvent = {
      store,
      sheet: sheet.name,
      address: bareAddress(rect.r0, rect.c0, rect.r1, rect.c1),
      kind
    };
    let effective = matrix;
    const pending: typeof writeInterceptors = [];
    for (const rule of writeInterceptors) {
      if (!rule.predicate(event)) {
        pending.push(rule);
        continue;
      }
      if (rule.fail) throw new Error(rule.fail);
      if (rule.transform) effective = rule.transform(effective, event);
      if (!rule.once) pending.push(rule);
    }
    writeInterceptors = pending;
    writes.push({
      ...event,
      cells: effective.length * (effective[0]?.length ?? 0)
    });
    effective.forEach((row, rowIndex) =>
      row.forEach((value, columnIndex) =>
        writeCell(sheet, rect.r0 + rowIndex, rect.c0 + columnIndex, value, kind === "formulas")
      )
    );
    return effective;
  };

  const usedBounds = (sheet: SheetModel) => {
    let minRow = Infinity;
    let minColumn = Infinity;
    let maxRow = -1;
    let maxColumn = -1;
    for (const key of sheet.cells.keys()) {
      const [row, column] = key.split(",").map(Number);
      if (row < minRow) minRow = row;
      if (column < minColumn) minColumn = column;
      if (row > maxRow) maxRow = row;
      if (column > maxColumn) maxColumn = column;
    }
    if (maxRow < 0) return null;
    return { r0: minRow, c0: minColumn, r1: maxRow, c1: maxColumn };
  };

  const wrapRange = (
    book: BookModel,
    store: TestHostStore,
    sheet: SheetModel,
    rect: { r0: number; c0: number; r1: number; c1: number }
  ) => {
    const bounds = () => rect;
    const range: Record<string, unknown> = {
      load: () => undefined,
      get rowIndex() {
        return rect.r0;
      },
      get columnIndex() {
        return rect.c0;
      },
      get rowCount() {
        return rect.r1 - rect.r0 + 1;
      },
      get columnCount() {
        return rect.c1 - rect.c0 + 1;
      },
      get address() {
        return `${sheet.name}!${bareAddress(rect.r0, rect.c0, rect.r1, rect.c1)}`;
      },
      get values() {
        return matrixOf(sheet, bounds(), "values");
      },
      set values(matrix: CellValue[][]) {
        assertMatrixShape(matrix, rect);
        commitWrite(store, sheet, rect, matrix, "values");
      },
      get formulas() {
        return matrixOf(sheet, bounds(), "formulas");
      },
      set formulas(matrix: CellValue[][]) {
        assertMatrixShape(matrix, rect);
        commitWrite(store, sheet, rect, matrix, "formulas");
      },
      get formulasR1C1() {
        return matrixOf(sheet, bounds(), "formulas");
      },
      set formulasR1C1(matrix: CellValue[][]) {
        assertMatrixShape(matrix, rect);
        commitWrite(store, sheet, rect, matrix, "formulas");
      },
      get text() {
        return matrixOf(sheet, bounds(), "text");
      },
      get numberFormat() {
        return Array.from({ length: rect.r1 - rect.r0 + 1 }, (_unused, rowIndex) =>
          Array.from({ length: rect.c1 - rect.c0 + 1 }, (_also, columnIndex) => {
            const cell = sheet.cells.get(
              cellKey(rect.r0 + rowIndex, rect.c0 + columnIndex)
            );
            return cell?.format ?? "General";
          })
        );
      },
      set numberFormat(matrix: string[][]) {
        matrix.forEach((row, rowIndex) =>
          row.forEach((format, columnIndex) => {
            const cell = sheet.cells.get(
              cellKey(rect.r0 + rowIndex, rect.c0 + columnIndex)
            );
            if (cell) cell.format = format;
          })
        );
      },
      get worksheet() {
        return wrapSheet(book, store, sheet);
      },
      get format() {
        return makeFormat(sheet, rect, store);
      },
      getCellProperties: (loadOptions?: {
        style?: boolean;
        format?: Record<string, unknown>;
      }) => ({
        load: () => undefined,
        get value() {
          // 忠实模拟真实 Office.js：`getCellProperties` 只返回**直接设置在单元格上**
          // 的属性。没有任何直接设置格式的单元格，返回空对象（`format` 缺失），而不是
          // 塞进一套 `defaultCellStyle()` 的默认值 —— 否则生产代码会把"未设置"当成
          // "默认值"，测试永远抓不到"未采集 ≠ 默认值"这条缺陷。
          //
          // **并且只返回被请求的字段**（真实宿主如此）。这一点很重要：生产代码若漏请求
          // 某个会被 `clear(Formats)` 重置的属性，替身必须让它读不到 —— 否则"不请求
          // 就没有、不写回就等于丢"这条纪律在测试里恒真通过。
          const wantFormat = (loadOptions?.format ?? {}) as Record<string, unknown>;
          const sub = (key: string): Record<string, unknown> =>
            (wantFormat[key] ?? {}) as Record<string, unknown>;
          const matrix = Array.from({ length: rect.r1 - rect.r0 + 1 }, (_unused, rowIndex) =>
            Array.from({ length: rect.c1 - rect.c0 + 1 }, (_also, columnIndex) => {
              const cell = sheet.cells.get(cellKey(rect.r0 + rowIndex, rect.c0 + columnIndex));
              const style = cell?.style;
              const output: Record<string, unknown> = {};
              // 命名样式是 `CellPropertiesLoadOptions.style`（与 `format` 同级）。
              if (loadOptions?.style === true) output.style = cell?.styleName ?? "Normal";
              if (!style) return output;
              // 只输出**确实存在**的字段：缺席 = 未直接设置（继承/默认）。这是"直接格式"
              // 兼容路径能区分来源的前提。
              const format: Record<string, unknown> = {};
              const fillWant = sub("fill");
              const fill: Record<string, unknown> = {};
              if (fillWant.pattern === true && style.fillPattern != null) {
                fill.pattern = style.fillPattern;
              }
              // 无填充（pattern "None"）时真实 Office.js 不返回 color（undefined），
              // 而不是返回 `#FFFFFF`；有填充时才返回填充色。
              if (
                fillWant.color === true &&
                style.fillColor != null &&
                style.fillPattern !== "None"
              ) {
                fill.color = style.fillColor;
              }
              if (fillWant.patternColor === true && style.fillPatternColor != null) {
                fill.patternColor = style.fillPatternColor;
              }
              if (fillWant.patternTintAndShade === true && style.fillPatternTintAndShade != null) {
                fill.patternTintAndShade = style.fillPatternTintAndShade;
              }
              if (fillWant.tintAndShade === true && style.fillTintAndShade != null) {
                fill.tintAndShade = style.fillTintAndShade;
              }
              if (Object.keys(fill).length > 0) format.fill = fill;
              const fontWant = sub("font");
              const font: Record<string, unknown> = {};
              if (fontWant.name === true && style.fontName != null) font.name = style.fontName;
              if (fontWant.size === true && style.fontSize != null) font.size = style.fontSize;
              if (fontWant.bold === true && style.fontBold != null) font.bold = style.fontBold;
              if (fontWant.italic === true && style.fontItalic != null) font.italic = style.fontItalic;
              if (fontWant.color === true && style.fontColor != null) font.color = style.fontColor;
              if (fontWant.underline === true && style.fontUnderline != null) {
                font.underline = style.fontUnderline;
              }
              if (fontWant.strikethrough === true && style.fontStrikethrough != null) {
                font.strikethrough = style.fontStrikethrough;
              }
              if (fontWant.subscript === true && style.fontSubscript != null) {
                font.subscript = style.fontSubscript;
              }
              if (fontWant.superscript === true && style.fontSuperscript != null) {
                font.superscript = style.fontSuperscript;
              }
              if (fontWant.tintAndShade === true && style.fontTintAndShade != null) {
                font.tintAndShade = style.fontTintAndShade;
              }
              if (Object.keys(font).length > 0) format.font = font;
              if (wantFormat.horizontalAlignment === true && style.horizontalAlignment != null) {
                format.horizontalAlignment = style.horizontalAlignment;
              }
              if (wantFormat.verticalAlignment === true && style.verticalAlignment != null) {
                format.verticalAlignment = style.verticalAlignment;
              }
              if (wantFormat.wrapText === true && style.wrapText != null) {
                format.wrapText = style.wrapText;
              }
              if (wantFormat.indentLevel === true && style.indentLevel != null) {
                format.indentLevel = style.indentLevel;
              }
              if (wantFormat.shrinkToFit === true && style.shrinkToFit != null) {
                format.shrinkToFit = style.shrinkToFit;
              }
              if (wantFormat.textOrientation === true && style.textOrientation != null) {
                format.textOrientation = style.textOrientation;
              }
              if (wantFormat.readingOrder === true && style.readingOrder != null) {
                format.readingOrder = style.readingOrder;
              }
              if (wantFormat.autoIndent === true && style.autoIndent != null) {
                format.autoIndent = style.autoIndent;
              }
              if (
                wantFormat.protection === true &&
                (style.locked != null || style.formulaHidden != null)
              ) {
                format.protection = {
                  ...(style.locked != null ? { locked: style.locked } : {}),
                  ...(style.formulaHidden != null ? { formulaHidden: style.formulaHidden } : {})
                };
              }
              const borderWant = sub("borders");
              if (style.borders && Object.keys(style.borders).length > 0) {
                const borders: Record<string, unknown> = {};
                for (const [side, border] of Object.entries(style.borders)) {
                  const single: Record<string, unknown> = {};
                  if (borderWant.style === true && border.style != null) single.style = border.style;
                  if (borderWant.color === true && border.color != null) single.color = border.color;
                  if (borderWant.weight === true && border.weight != null) single.weight = border.weight;
                  if (borderWant.tintAndShade === true && border.tintAndShade != null) {
                    single.tintAndShade = border.tintAndShade;
                  }
                  if (Object.keys(single).length > 0) borders[side] = single;
                }
                if (Object.keys(borders).length > 0) format.borders = borders;
              }
              if (Object.keys(format).length > 0) output.format = format;
              return output;
            })
          );
          return options.corruptFormatRead
            ? options.corruptFormatRead(matrix, "cell")
            : matrix;
        }
      }),
      // 忠实模拟 `getDisplayedCellProperties`（ExcelApi 1.19）：返回单元格**实际显示**的
      // 格式 —— 无论来源是直接设置还是表格样式 / 命名样式 / 条件格式。显示值 =
      // 直接设置值（cell.style）覆盖到样式背景（sheet.displayedDefaults）再覆盖到硬默认
      // （defaultCellStyle()）。这样：
      //   - "明确无填充"：fill 未直接设置、样式背景也无 ⇒ 显示 pattern "None"；
      //   - "来源于样式的填充"：fill 未直接设置、但 displayedDefaults 给了 fillColor ⇒
      //     显示 pattern "Solid" + 该色（getCellProperties 此时仍返回 undefined）。
      getDisplayedCellProperties: () => ({
        load: () => undefined,
        get value() {
          const matrix = Array.from({ length: rect.r1 - rect.r0 + 1 }, (_unused, rowIndex) =>
            Array.from({ length: rect.c1 - rect.c0 + 1 }, (_also, columnIndex) => {
              const cell = sheet.cells.get(cellKey(rect.r0 + rowIndex, rect.c0 + columnIndex));
              const style = cell?.style;
              // 显示值 = 硬默认 <- 样式背景 <- 直接设置值。
              const displayed = mergeDisplayedStyle(sheet.displayedDefaults, style);
              const fill =
                displayed.fillPattern === "None"
                  ? { pattern: "None" }
                  : { pattern: displayed.fillPattern, color: displayed.fillColor };
              return {
                format: {
                  fill,
                  font: {
                    name: displayed.fontName,
                    size: displayed.fontSize,
                    bold: displayed.fontBold,
                    italic: displayed.fontItalic,
                    color: displayed.fontColor,
                    underline: displayed.fontUnderline
                  },
                  horizontalAlignment: displayed.horizontalAlignment,
                  verticalAlignment: displayed.verticalAlignment,
                  wrapText: displayed.wrapText,
                  indentLevel: displayed.indentLevel,
                  borders: {
                    top: { ...(displayed.borders.top ?? { style: "None", color: "#000000", weight: "Thin" }) },
                    bottom: { ...(displayed.borders.bottom ?? { style: "None", color: "#000000", weight: "Thin" }) },
                    left: { ...(displayed.borders.left ?? { style: "None", color: "#000000", weight: "Thin" }) },
                    right: { ...(displayed.borders.right ?? { style: "None", color: "#000000", weight: "Thin" }) }
                  }
                }
              };
            })
          );
          return options.corruptFormatRead
            ? options.corruptFormatRead(matrix, "cell")
            : matrix;
        }
      }),
      setCellProperties: (
        matrix: Array<Array<{ format?: Record<string, unknown>; style?: string }>>
      ) => {
        let touched = 0;
        matrix.forEach((row, rowIndex) =>
          row.forEach((properties, columnIndex) => {
            const key = cellKey(rect.r0 + rowIndex, rect.c0 + columnIndex);
            const existing = sheet.cells.get(key);
            // 命名样式（`SettableCellProperties.style`）与 `format` 同级，单独落。
            const namedStyle =
              typeof properties?.style === "string" ? properties.style : undefined;
            // 部分样式合并：只写"入参里确实给了"的字段。
            const style = applyCellProperties(existing?.style ?? {}, properties ?? {});
            const empty = isEmptyPartialStyle(style);
            const nextStyle = empty ? undefined : style;
            if (!existing) {
              if (nextStyle == null && namedStyle == null) return;
              // 格式操作不该凭空造出一个"有值"的单元格：用空串与"无 cell"的读取约定
              // （`matrixOf` 对空位返回 `""`）保持一致，避免把 `""` 变成 `null`。
              sheet.cells.set(key, {
                value: "",
                formula: null,
                style: nextStyle,
                ...(namedStyle != null ? { styleName: namedStyle } : {})
              });
            } else {
              existing.style = nextStyle;
              if (namedStyle != null) existing.styleName = namedStyle;
            }
            touched += 1;
          })
        );
        // 逐格写入也记一次（`clear(Formats)` 一直有记）。少了它，"撤销期间只写了值"
        // 这类取证就只能看见 clear、看不见 setCellProperties，判别力是残的。
        if (touched > 0) {
          writes.push({
            store,
            sheet: sheet.name,
            address: bareAddress(rect.r0, rect.c0, rect.r1, rect.c1),
            kind: "format",
            cells: touched
          });
        }
      },
      getRowProperties: () => ({
        load: () => undefined,
        get value() {
          // 忠实模拟真实 Office.js：`getRowProperties` 返回 `format.rowHeight`（直接设置）
          // 或 `format.useStandardHeight: true`（未直接设置、使用标准行高）。未直接设置的
          // 行不返回 `rowHeight`，而是返回 `useStandardHeight: true` —— 这是"确定状态"
          // （标准行高），不是"无法确定"。
          const matrix = Array.from({ length: rect.r1 - rect.r0 + 1 }, (_unused, rowIndex) => {
            const height = sheet.rowHeights.get(rect.r0 + rowIndex);
            return height == null
              ? { format: { useStandardHeight: true } }
              : { format: { rowHeight: height } };
          });
          return options.corruptFormatRead
            ? options.corruptFormatRead(matrix, "row")
            : matrix;
        }
      }),
      setRowProperties: (matrix: Array<{ format?: { rowHeight?: number } }>) => {
        matrix.forEach((row, rowIndex) => {
          if (row?.format?.rowHeight != null) {
            sheet.rowHeights.set(rect.r0 + rowIndex, row.format.rowHeight);
          }
        });
      },
      getColumnProperties: () => ({
        load: () => undefined,
        get value() {
          // 与 getRowProperties 同口径：未直接设置列宽的列返回 `useStandardWidth: true`。
          const matrix = Array.from({ length: rect.c1 - rect.c0 + 1 }, (_unused, columnIndex) => {
            const width = sheet.columnWidths.get(rect.c0 + columnIndex);
            return width == null
              ? { format: { useStandardWidth: true } }
              : { format: { columnWidth: width } };
          });
          return options.corruptFormatRead
            ? options.corruptFormatRead(matrix, "column")
            : matrix;
        }
      }),
      setColumnProperties: (matrix: Array<{ format?: { columnWidth?: number } }>) => {
        matrix.forEach((column, columnIndex) => {
          if (column?.format?.columnWidth != null) {
            sheet.columnWidths.set(rect.c0 + columnIndex, column.format.columnWidth);
          }
        });
      },
      clear: (applyTo?: unknown) => {
        // 忠实模拟 `Range.clear(applyTo)`：`Formats` 只清**直接格式**、**不动值**，
        // 这正是兼容路径"先清空直接格式、再写回直接格式集合"所依赖的语义。
        // 旧实现无差别清值 —— 会让兼容路径的撤销把整片数据抹掉，且测不出来。
        const normalized = typeof applyTo === "string" ? applyTo.toLowerCase() : "all";
        const clearsFormats = normalized === "formats" || normalized === "all";
        const clearsContents = normalized === "contents" || normalized === "all";
        if (clearsFormats) {
          for (let row = rect.r0; row <= rect.r1; row += 1) {
            for (let column = rect.c0; column <= rect.c1; column += 1) {
              const cell = sheet.cells.get(cellKey(row, column));
              if (!cell) continue;
              // 真实 `clear(Formats)` 清掉的是"这个单元格的**全部**格式"，含：
              //   ① 逐属性直接格式（`style`）；
              //   ② **数字格式**（退回 General）；
              //   ③ **单元格命名样式**（退回 Normal）。
              // 只 `delete cell.style` 是"只删除某个局部对象"，与真实宿主不符 ——
              // 而这正是"采集清单不全也照样通过"的根因，必须一并清掉。
              delete cell.style;
              delete cell.format;
              cell.styleName = undefined;
            }
          }
          writes.push({
            store,
            sheet: sheet.name,
            address: bareAddress(rect.r0, rect.c0, rect.r1, rect.c1),
            kind: "format",
            cells: (rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1)
          });
        }
        if (clearsContents) {
          commitWrite(
            store,
            sheet,
            rect,
            Array.from({ length: rect.r1 - rect.r0 + 1 }, () =>
              Array.from({ length: rect.c1 - rect.c0 + 1 }, () => "")
            ),
            "clear"
          );
        }
      },
      calculate: () => undefined,
      select: () => undefined,
      getRow: (index: number) => ({
        format: {
          load: () => undefined,
          get rowHeight() {
            return sheet.rowHeights.get(rect.r0 + index) ?? 15;
          },
          set rowHeight(value: number) {
            sheet.rowHeights.set(rect.r0 + index, value);
          },
          columnWidth: 8.43
        },
        clear: () => undefined,
        get rowIndex() {
          return rect.r0 + index;
        }
      }),
      getColumn: (index: number) => ({
        format: {
          load: () => undefined,
          rowHeight: 15,
          get columnWidth() {
            return sheet.columnWidths.get(rect.c0 + index) ?? 8.43;
          },
          set columnWidth(value: number) {
            sheet.columnWidths.set(rect.c0 + index, value);
          }
        },
        get columnIndex() {
          return rect.c0 + index;
        }
      }),
      getMergedAreasOrNullObject: () =>
        options.mergedAreas ? { isNullObject: false, load: () => undefined } : nullObject(),
      getSpecialCellsOrNullObject: () => nullObject(),
      getUsedRangeOrNullObject: () => {
        const used = usedBounds(sheet);
        return used ? wrapRange(book, store, sheet, used) : nullObject();
      },
      get conditionalFormats() {
        const rules = sheet.reviewRules ??= [];
        return {
          load: () => undefined,
          get items() { return rules.filter(rule => rule.address === bareAddress(rect.r0, rect.c0, rect.r1, rect.c1)); },
          add: (type: string) => {
            const borders = new Map<string, Record<string, unknown>>();
            const rule = { id: crypto.randomUUID(), address: bareAddress(rect.r0, rect.c0, rect.r1, rect.c1), type,
              custom: { rule: { formula: '' }, load: () => undefined,
                format: { borders: { getItem: (edge: string) => {
                  if (!borders.has(edge)) borders.set(edge, {});
                  return borders.get(edge)!;
                } } } }, load: () => undefined,
              delete: () => { const index = rules.indexOf(rule); if (index >= 0) rules.splice(index, 1); }
            };
            rules.push(rule); return rule;
          },
          getCount: () => ({
            load: () => undefined,
            value: (sheet.conditionalFormatCount ?? options.conditionalFormatCount ?? 0) + rules.length
          })
        };
      },
      get dataValidation() {
        return {
          // 真实 `range.dataValidation.type` 是 readonly 枚举：区域内无验证 ⇒ "None"，
          // 有验证 ⇒ 具体类型（"List"/"WholeNumber"…），部分有 ⇒ "MixedCriteria"。
          // 这里用"是否有任一格带验证"表达"有验证 = 非 None"这一判据。
          load: () => undefined,
          get type() {
            const hasAny = Array.from(
              { length: rect.r1 - rect.r0 + 1 },
              (_u, rowIndex) => rowIndex
            ).some((rowIndex) =>
              Array.from(
                { length: rect.c1 - rect.c0 + 1 },
                (_v, columnIndex) => columnIndex
              ).some((columnIndex) =>
                sheet.dataValidatedCells?.has(cellKey(rect.r0 + rowIndex, rect.c0 + columnIndex))
              )
            );
            return hasAny ? "List" : "None";
          }
        };
      },
      get tables() {
        return emptyCountCollection();
      },
      get charts() {
        return emptyCountCollection();
      },
      get shapes() {
        return emptyCountCollection();
      }
    };
    return range;
  };

  const makeFormat = (
    sheet: SheetModel,
    rect: { r0: number; c0: number; r1: number; c1: number },
    store: TestHostStore
  ) => {
    // 真实 Excel 的 `range.format.*` 写入作用于**区域内每个单元格**（行高作用于
    // 每行、列宽作用于每列）。替身必须忠实复刻这个语义，否则：
    //   - `setFill` 等格式动作执行后，`getCellProperties` 逐格回读读不到变化，
    //     回读判据形同虚设；
    //   - 撤销的逐格恢复（`setCellProperties` 写回捕获值）也失去判别力 ——
    //     执行前后逐格值一样，测试恒真。
    // 所以这里把区域级写入**广播**到逐格 style / 逐行 rowHeights / 逐列 columnWidths。
    const cellKeyFor = () => cellKey(rect.r0, rect.c0);
    /** 读取时补默认（= 真实生效的显示值）；写入时只往部分样式里落字段。 */
    const styleOf = (): CellStyle => mergedCellStyle(sheet.cells.get(cellKeyFor())?.style);
    /** 把一次格式变更广播到区域内所有单元格；返回命中单元格数。 */
    const applyToAll = (mutate: (style: Partial<CellStyle>) => void): number => {
      let touched = 0;
      for (let row = rect.r0; row <= rect.r1; row += 1) {
        for (let column = rect.c0; column <= rect.c1; column += 1) {
          const key = cellKey(row, column);
          const existing = sheet.cells.get(key);
          const style: Partial<CellStyle> = {
            ...(existing?.style ?? {}),
            borders: { ...(existing?.style?.borders ?? {}) }
          };
          mutate(style);
          if (Object.keys(style.borders ?? {}).length === 0) delete style.borders;
          if (!existing) sheet.cells.set(key, { value: "", formula: null, style });
          else existing.style = style;
          touched += 1;
        }
      }
      if (touched > 0) {
        writes.push({
          store,
          sheet: sheet.name,
          address: bareAddress(rect.r0, rect.c0, rect.r1, rect.c1),
          kind: "format",
          cells: touched
        });
      }
      return touched;
    };
    const commitStyle = (mutate: (style: Partial<CellStyle>) => void) => {
      applyToAll(mutate);
    };
    // 行高是行级、列宽是列级属性：写入广播到区域内每一行/每一列。
    const applyToRows = (mutate: (row: number) => void) => {
      for (let row = rect.r0; row <= rect.r1; row += 1) mutate(row);
    };
    const applyToColumns = (mutate: (column: number) => void) => {
      for (let column = rect.c0; column <= rect.c1; column += 1) mutate(column);
    };
    return {
      load: () => undefined,
      get horizontalAlignment() {
        return styleOf().horizontalAlignment;
      },
      set horizontalAlignment(value: string) {
        commitStyle((style) => { style.horizontalAlignment = value; });
      },
      get verticalAlignment() {
        return styleOf().verticalAlignment;
      },
      set verticalAlignment(value: string) {
        commitStyle((style) => { style.verticalAlignment = value; });
      },
      get wrapText() {
        return styleOf().wrapText;
      },
      set wrapText(value: boolean) {
        commitStyle((style) => { style.wrapText = value; });
      },
      get rowHeight() {
        return sheet.rowHeights.get(rect.r0) ?? 15;
      },
      set rowHeight(value: number) {
        applyToRows((row) => sheet.rowHeights.set(row, value));
      },
      get columnWidth() {
        return sheet.columnWidths.get(rect.c0) ?? 8.43;
      },
      set columnWidth(value: number) {
        applyToColumns((column) => sheet.columnWidths.set(column, value));
      },
      // 真实 Office.js 的 `RangeFormat.useStandardHeight`/`useStandardWidth`（1.7）：设为
      // true 时把区域内所有行/列重置为标准行高/列宽（即"未直接设置"）。restoreRangeFormat
      // 用它恢复"执行前未直接设置行高/列宽"的状态。
      get useStandardHeight() {
        return false;
      },
      set useStandardHeight(value: boolean) {
        if (value) applyToRows((row) => sheet.rowHeights.delete(row));
      },
      get useStandardWidth() {
        return false;
      },
      set useStandardWidth(value: boolean) {
        if (value) applyToColumns((column) => sheet.columnWidths.delete(column));
      },
      get indentLevel() {
        return styleOf().indentLevel;
      },
      set indentLevel(value: number) {
        commitStyle((style) => { style.indentLevel = value; });
      },
      get fill() {
        return {
          load: () => undefined,
          get color() {
            // 区域级单值摘要：格式动作统一写入后，区域内所有格一致，读首格即正确；
            // 混合态读首格是"混合值"的近似，`verifyPlan` 只在执行后读（此时已统一）。
            return styleOf().fillPattern === "None" ? "" : styleOf().fillColor;
          },
          set color(value: string) {
            commitStyle((style) => {
              style.fillColor = value;
              // 设置填充色即意味着有填充：pattern 从 None 变为 Solid（真实 Excel 语义）。
              if (value !== "") style.fillPattern = "Solid";
            });
          }
        };
      },
      get font() {
        return {
          load: () => undefined,
          get name() { return styleOf().fontName; },
          set name(value: string) { commitStyle((style) => { style.fontName = value; }); },
          get size() { return styleOf().fontSize; },
          set size(value: number) { commitStyle((style) => { style.fontSize = value; }); },
          get italic() { return styleOf().fontItalic; },
          set italic(value: boolean) { commitStyle((style) => { style.fontItalic = value; }); },
          get underline() { return styleOf().fontUnderline; },
          set underline(value: string) { commitStyle((style) => { style.fontUnderline = value; }); },
          get bold() { return styleOf().fontBold; },
          set bold(value: boolean) { commitStyle((style) => { style.fontBold = value; }); },
          get color() { return styleOf().fontColor; },
          set color(value: string) { commitStyle((style) => { style.fontColor = value; }); }
        };
      },
      get borders() {
        return {
          getItem: (sideIndex: number) => {
            const sideName = borderSideKey(sideIndex);
            return {
              load: () => undefined,
              get style() { return styleOf().borders[sideName]?.style ?? "None"; },
              set style(value: string) { commitStyle((style) => { style.borders![sideName] = { ...style.borders![sideName], style: value }; }); },
              get color() { return styleOf().borders[sideName]?.color ?? "#000000"; },
              set color(value: string) { commitStyle((style) => { style.borders![sideName] = { ...style.borders![sideName], color: value }; }); },
              get weight() { return styleOf().borders[sideName]?.weight ?? "Thin"; },
              set weight(value: string) { commitStyle((style) => { style.borders![sideName] = { ...style.borders![sideName], weight: value }; }); }
            };
          }
        };
      }
    };
  };

  const emptyCountCollection = () => ({
    load: () => undefined,
    getCount: () => ({ load: () => undefined, value: 0 }),
    items: [] as unknown[]
  });

  const nullObject = () => ({
    isNullObject: true,
    load: () => undefined
  });

  const assertMatrixShape = (
    matrix: CellValue[][],
    rect: { r0: number; c0: number; r1: number; c1: number }
  ) => {
    const rows = rect.r1 - rect.r0 + 1;
    const columns = rect.c1 - rect.c0 + 1;
    if (!Array.isArray(matrix) || matrix.length !== rows) {
      throw new Error(
        `宿主替身：写入矩阵行数 ${matrix?.length} 与区域 ${rows} 行不符`
      );
    }
    matrix.forEach((row, index) => {
      if (!Array.isArray(row) || row.length !== columns) {
        throw new Error(
          `宿主替身：写入矩阵第 ${index + 1} 行列数与区域 ${columns} 列不符`
        );
      }
    });
  };

  const wrapSheet = (
    book: BookModel,
    store: TestHostStore,
    sheet: SheetModel,
    nullFlag = false
  ): Record<string, unknown> => ({
    load: () => undefined,
    get isNullObject() {
      return nullFlag;
    },
    get id() {
      return sheet.id;
    },
    get name() {
      return sheet.name;
    },
    get position() {
      return sheet.position;
    },
    get protection() {
      return {
        load: () => undefined,
        get protected() {
          return sheet.protected;
        }
      };
    },
    getUsedRangeOrNullObject: () => {
      const used = usedBounds(sheet);
      return used ? wrapRange(book, store, sheet, used) : nullObject();
    },
    getUsedRange: () => {
      const used = usedBounds(sheet) ?? { r0: 0, c0: 0, r1: 0, c1: 0 };
      return wrapRange(book, store, sheet, used);
    },
    getRange: (address: string | Record<string, number>, to?: Record<string, number>) => {
      if (typeof address === "string") {
        return wrapRange(book, store, sheet, parseAddress(address));
      }
      const rect = {
        r0: address.Row,
        c0: address.Column,
        r1: to?.Row ?? address.Row,
        c1: to?.Column ?? address.Column
      };
      return wrapRange(book, store, sheet, rect);
    },
    getRangeByIndexes: (rowIndex: number, columnIndex: number, rowCount: number, columnCount: number) =>
      wrapRange(book, store, sheet, {
        r0: rowIndex,
        c0: columnIndex,
        r1: rowIndex + rowCount - 1,
        c1: columnIndex + columnCount - 1
      }),
    getCell: (row: number, column: number) =>
      wrapRange(book, store, sheet, { r0: row, c0: column, r1: row, c1: column }),
    get tables() {
      return emptyCountCollection();
    },
    get charts() {
      return emptyCountCollection();
    },
    get shapes() {
      return emptyCountCollection();
    },
    get pivotTables() {
      return emptyCountCollection();
    },
    get names() {
      return emptyCountCollection();
    },
    get autoFilter() {
      return {
        load: () => undefined,
        enabled: false,
        isDataFiltered: false,
        criteria: [] as unknown[],
        getRangeOrNullObject: () => nullObject()
      };
    },
    get onChanged() {
      return { add: () => undefined, remove: () => undefined };
    },
    activate: () => undefined,
    delete: () => undefined,
    calculate: () => undefined
  });

  const wrapWorksheetCollection = (book: BookModel, store: TestHostStore) => ({
    load: () => undefined,
    get items() {
      return book.sheets
        .slice()
        .sort((left, right) => left.position - right.position)
        .map((sheet) => wrapSheet(book, store, sheet));
    },
    /*
     * 真实 `WorksheetCollection.getItem` **既接受表名也接受表 id**
     * （产品里 `prepareOfficePythonResult` 的写前读取就是用 `sheetId` 定位的）。
     * 替身宿主只按名字匹配会让"用 id 定位"在这条路径上永远找不到表 —— 那是
     * 替身不够真，不是产品缺陷。两种都认，替身才与实机同语义。
     */
    getItem: (nameOrId: string) => {
      const found = book.sheets.find(
        (sheet) => sheet.name === nameOrId || sheet.id === nameOrId
      );
      if (!found) throw new Error(`宿主替身：工作表「${nameOrId}」不存在`);
      return wrapSheet(book, store, found);
    },
    getItemOrNullObject: (nameOrId: string) => {
      const found = book.sheets.find(
        (sheet) => sheet.name === nameOrId || sheet.id === nameOrId
      );
      return found
        ? wrapSheet(book, store, found)
        : wrapSheet(
            book,
            store,
            { id: "missing", name: nameOrId, position: -1, protected: false, cells: new Map(), rowHeights: new Map(), columnWidths: new Map(), displayedDefaults: {} },
            true
          );
    },
    getActiveWorksheet: () => {
      const window = windows.find((item) => item.name === activeWindowName)!;
      const sheet =
        window.book === book
          ? book.sheets.find((item) => item.id === window.activeSheetId)!
          : book.sheets[0];
      return wrapSheet(book, store, sheet);
    },
    add: (name: string) => {
      const created = makeSheet(name, book.sheets.length, undefined);
      book.sheets.push(created);
      return wrapSheet(book, store, created);
    },
    get onChanged() {
      return { add: () => undefined, remove: () => undefined };
    },
    get onAdded() {
      return { add: () => undefined, remove: () => undefined };
    },
    get onDeleted() {
      return { add: () => undefined, remove: () => undefined };
    },
    get onNameChanged() {
      return { add: () => undefined, remove: () => undefined };
    }
  });

  const wrapWindow = (window: WindowModel) => {
    const store: TestHostStore = window.book === originalBook ? "original" : window.book.id;
    return {
      load: () => undefined,
      get name() {
        return window.name;
      },
      get activeWorksheet() {
        const sheet = window.book.sheets.find((item) => item.id === window.activeSheetId)!;
        return wrapSheet(window.book, store, sheet);
      },
      activate: () => {
        activeWindowName = window.name;
      },
      close: () => {
        const index = windows.indexOf(window);
        if (index >= 0) windows.splice(index, 1);
        return { value: true };
      },
      get scrollRow() {
        return window.scrollRow;
      },
      set scrollRow(value: number) {
        window.scrollRow = value;
      },
      get scrollColumn() {
        return window.scrollColumn;
      },
      set scrollColumn(value: number) {
        window.scrollColumn = value;
      }
    };
  };

  const makeContext = () => {
    const book = originalBook;
    const store: TestHostStore = "original";
    const worksheetCollection = wrapWorksheetCollection(book, store);
    return {
      sync: async () => undefined,
      workbook: {
        worksheets: worksheetCollection,
        tables: { load: () => undefined, items: [] as unknown[] },
        names: {
          load: () => undefined,
          items: [] as unknown[],
          getCount: () => ({ load: () => undefined, value: 0 }),
          add: () => undefined
        },
        pivotTables: { load: () => undefined, items: [] as unknown[] },
        comments: { add: () => undefined },
        getSelectedRange: () =>
          wrapRange(
            book,
            store,
            book.sheets.find((sheet) => sheet.name === activeSheetName)!,
            parseAddress(selectionAddress)
          ),
        getActiveWorksheet: () => worksheetCollection.getActiveWorksheet(),
        application: {
          get windows() {
            return {
              load: () => undefined,
              get items() {
                return windows.map((window) => wrapWindow(window));
              },
              compareCurrentSideBySideWith: () => undefined,
              resetPositionsSideBySide: () => undefined
            };
          },
          get activeWindow() {
            const window = windows.find((item) => item.name === activeWindowName)!;
            return wrapWindow(window);
          }
        }
      }
    };
  };

  const apiVersionAtLeast = (
    set: string,
    required: string,
    available: string[]
  ): boolean => {
    if (set !== "ExcelApi" && set !== "ExcelApiDesktop") return false;
    const [major, minor] = required.split(".").map(Number);
    return available.some((version) => {
      const [availableMajor, availableMinor] = version.split(".").map(Number);
      return (
        availableMajor > major ||
        (availableMajor === major && availableMinor >= minor)
      );
    });
  };

  const excelApiVersions =
    options.excelApiVersions ??
    Array.from({ length: 18 }, (_, index) => `1.${index + 1}`);
  const excelApiDesktopVersions = options.excelApiDesktopVersions ?? ["1.1"];

  const documentStub: Record<string, unknown> = {
    url: options.url ?? "",
    mode: "ReadWrite"
  };
  let documentRef = documentStub;

  const OfficeStub = {
    HostType: { Excel: "Excel", Word: "Word", PowerPoint: "PowerPoint" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: {
      DialogEventReceived: "dialogEventReceived",
      DialogMessageReceived: "dialogMessageReceived",
      DialogParentMessageReceived: "dialogParentMessageReceived"
    },
    onReady: (callback: () => void) => {
      callback();
      return Promise.resolve({ host: "Excel", platform: "PC" });
    },
    context: {
      get document() {
        return documentRef;
      },
      host: "Excel",
      diagnostics: { version: "test-host" },
      requirements: {
        isSetSupported: (set: string, version?: string) => {
          if (!version) return true;
          if (set === "ExcelApi") {
            return apiVersionAtLeast(set, version, excelApiVersions);
          }
          if (set === "ExcelApiDesktop") {
            return apiVersionAtLeast(set, version, excelApiDesktopVersions);
          }
          return false;
        }
      },
      ui: {
        displayDialogAsync: () => {
          throw new Error("宿主替身未实现 displayDialogAsync");
        },
        addHandlerAsync: () => undefined,
        messageParent: () => undefined
      }
    }
  };

  const ExcelStub = {
    // 真实 Office.js 的 `Excel.run` 直接兑现回调的返回值（不是 {status,value} 信封）。
    run: async <T>(callback: (context: unknown) => Promise<T>): Promise<T> =>
      callback(makeContext()),
    createWorkbook: async (base64: string) => {
      const entries = readStoredZip(base64);
      seeds.push(entries);
      const workbookXml = entries["xl/workbook.xml"];
      if (!workbookXml) {
        throw new Error("宿主替身：副本种子缺少 xl/workbook.xml");
      }
      const sheetName = /<sheet[^>]*name="([^"]+)"/.exec(workbookXml)?.[1];
      if (!sheetName) {
        throw new Error("宿主替身：副本种子未声明工作表名");
      }
      // 副本初始内容 = 预览那一刻的原表状态（真实宿主由 Excel 打开种子得到）。
      const source = originalBook.sheets.find(
        (sheet) => sheet.name === activeSheetName
      )!;
      const copy: BookModel = {
        id: `book-${(bookSeq += 1)}`,
        sheets: [makeSheet(sheetName, 0, undefined)]
      };
      copy.sheets[0].cells = new Map(
        [...source.cells.entries()].map(([key, cell]) => [key, { ...cell }])
      );
      const window: WindowModel = {
        name: `Book${(windowSeq += 1)}`,
        book: copy,
        activeSheetId: copy.sheets[0].id,
        scrollRow: 0,
        scrollColumn: 0
      };
      windows.push(window);
      return { status: "succeeded", value: undefined };
    },
    SpecialCellType: {
      formulas: "Formulas",
      dataValidations: "DataValidations",
      blanks: "Blanks"
    },
    ClearApplyTo: {
      all: "All",
      contents: "Contents",
      formats: "Formats",
      hyperlinks: "Hyperlinks"
    },
    ClearOptions: { contentsOnly: "Contents" },
    BorderIndex: {
      edgeTop: "EdgeTop",
      edgeBottom: "EdgeBottom",
      edgeLeft: "EdgeLeft",
      edgeRight: "EdgeRight",
      insideHorizontal: "InsideHorizontal",
      insideVertical: "InsideVertical",
      diagonalDown: "DiagonalDown",
      diagonalUp: "DiagonalUp"
    },
    BorderLineStyle: { continuous: "Continuous", none: "None" },
    BorderWeight: { thin: "Thin", medium: "Medium", thick: "Thick" },
    RangeUnderlineStyle: { none: "None", single: "Single" },
    HorizontalAlignment: { general: "General", left: "Left", center: "Center", right: "Right" },
    VerticalAlignment: { bottom: "Bottom", center: "Center", top: "Top" },
    RangeCopyType: { all: "All", values: "Values", formats: "Formats" },
    WorksheetPositionType: { beginning: "Beginning", end: "End" },
    ChartType: { columnClustered: "ColumnClustered", line: "Line" },
    AggregationFunction: { sum: "Sum", average: "Average", count: "Count" },
    ConditionalCellValueOperator: { greaterThan: "GreaterThan" },
    DataValidationOperator: { between: "Between", equalTo: "EqualTo" },
    GeometricShapeType: { rectangle: "Rectangle" }
  };

  const storeOf = (sheetName: string, store: TestHostStore) => {
    if (store === "original") {
      const sheet = originalBook.sheets.find((item) => item.name === sheetName);
      if (!sheet) throw new Error(`宿主替身：原表没有工作表「${sheetName}」`);
      return sheet;
    }
    const book = windows.find((item) => item.book.id === store)?.book;
    const sheet = book?.sheets.find((item) => item.name === sheetName);
    if (!sheet) throw new Error(`宿主替身：副本 ${store} 没有工作表「${sheetName}」`);
    return sheet;
  };

  return {
    Excel: ExcelStub as unknown,
    Office: OfficeStub as unknown,
    writes: () => [...writes],
    originalWrites: () => writes.filter((write) => write.store === "original"),
    copyWrites: () => writes.filter((write) => write.store !== "original"),
    resetWrites: () => {
      writes.length = 0;
    },
    cell: (sheetName, address, store = "original") => {
      const sheet = storeOf(sheetName, store);
      const rect = parseAddress(address);
      const cell = sheet.cells.get(cellKey(rect.r0, rect.c0));
      return cell ? cell.value : "";
    },
    matrix: (sheetName, address, store = "original") => {
      const sheet = storeOf(sheetName, store);
      return matrixOf(sheet, parseAddress(address), "values");
    },
    setCell: (sheetName, address, value, cellOptions) => {
      const sheet = storeOf(sheetName, cellOptions?.store ?? "original");
      const rect = parseAddress(address);
      writeCell(sheet, rect.r0, rect.c0, value, cellOptions?.asFormula ?? false);
    },
    /** 给已存在的单元格挂上显示文本/数字格式（值归一契约用）。 */
    setCellDisplay: (sheetName, address, display, cellOptions) => {
      const sheet = storeOf(sheetName, cellOptions?.store ?? "original");
      const rect = parseAddress(address);
      const cell = sheet.cells.get(cellKey(rect.r0, rect.c0));
      if (!cell) throw new Error(`宿主替身：单元格「${sheetName}!${address}」不存在`);
      if (display.text !== undefined) cell.text = display.text;
      if (display.format !== undefined) cell.format = display.format;
    },
    transformWrites: (predicate, transform) => {
      writeInterceptors.push({ predicate, transform, once: false });
    },
    failWrites: (predicate, message) => {
      writeInterceptors.push({ predicate, fail: message, once: false });
    },
    tamperNextWrite: (predicate, transform) => {
      writeInterceptors.push({ predicate, transform, once: true });
    },
    setDocumentUrl: (url: string) => {
      documentRef = { url, mode: "ReadWrite" };
    },
    windowNames: () => windows.map((window) => window.name),
    copyStores: () =>
      windows
        .filter((window) => window.book !== originalBook)
        .map((window) => window.book.id),
    createWorkbookCalls: () => seeds.map((entries) => entries["xl/workbook.xml"] ?? ""),
    seedEntries: (index: number) => seeds[index],
    resetInterceptors: () => {
      writeInterceptors = [];
    },
    setCellStyle: (sheetName, address, style, cellOptions) => {
      const sheet = storeOf(sheetName, cellOptions?.store ?? "original");
      const rect = parseAddress(address);
      const key = cellKey(rect.r0, rect.c0);
      const existing = sheet.cells.get(key);
      // 只合并"这次确实给了"的字段：其余保持未直接设置，`getCellProperties` 才读得出
      // "直接设置 vs 继承"的区别。
      const next: Partial<CellStyle> = {
        ...(existing?.style ?? {}),
        ...style,
        borders: { ...(existing?.style?.borders ?? {}), ...(style.borders ?? {}) }
      };
      // 设置了非空填充色即意味着有填充：pattern 从 None 变为 Solid。
      if (style.fillColor != null && style.fillColor !== "" && (next.fillPattern ?? "None") === "None") {
        next.fillPattern = "Solid";
      }
      if (style.fillPattern === "None") next.fillColor = "";
      if (Object.keys(next.borders ?? {}).length === 0) delete next.borders;
      if (!existing) sheet.cells.set(key, { value: "", formula: null, style: next });
      else existing.style = next;
    },
    cellStyle: (sheetName, address, store = "original") => {
      const sheet = storeOf(sheetName, store);
      const rect = parseAddress(address);
      const cell = sheet.cells.get(cellKey(rect.r0, rect.c0));
      return cell?.style;
    },
    cellDisplayedStyle: (sheetName, address, store = "original") => {
      const sheet = storeOf(sheetName, store);
      const rect = parseAddress(address);
      const cell = sheet.cells.get(cellKey(rect.r0, rect.c0));
      return mergeDisplayedStyle(sheet.displayedDefaults, cell?.style);
    },
    setCellStyleName: (sheetName, address, styleName, cellOptions) => {
      const sheet = storeOf(sheetName, cellOptions?.store ?? "original");
      const rect = parseAddress(address);
      const key = cellKey(rect.r0, rect.c0);
      const existing = sheet.cells.get(key);
      if (!existing) sheet.cells.set(key, { value: "", formula: null, styleName });
      else existing.styleName = styleName;
    },
    cellStyleName: (sheetName, address, store = "original") => {
      const sheet = storeOf(sheetName, store);
      const rect = parseAddress(address);
      return sheet.cells.get(cellKey(rect.r0, rect.c0))?.styleName;
    },
    cellNumberFormat: (sheetName, address, store = "original") => {
      const sheet = storeOf(sheetName, store);
      const rect = parseAddress(address);
      return sheet.cells.get(cellKey(rect.r0, rect.c0))?.format;
    },
    conditionalFormatCount: (sheetName, store = "original") => {
      const sheet = storeOf(sheetName, store);
      return (sheet.conditionalFormatCount ?? options.conditionalFormatCount ?? 0) + (sheet.reviewRules?.length ?? 0);
    },
    setConditionalFormatCount: (sheetName, count, store = "original") => {
      storeOf(sheetName, store).conditionalFormatCount = count;
    },
    dataValidationCount: (sheetName, store = "original") => {
      const sheet = storeOf(sheetName, store);
      return sheet.dataValidatedCells?.size ?? 0;
    },
    setDataValidation: (sheetName, address, enabled, store = "original") => {
      const sheet = storeOf(sheetName, store);
      const rect = parseAddress(address);
      for (let row = rect.r0; row <= rect.r1; row += 1) {
        for (let column = rect.c0; column <= rect.c1; column += 1) {
          if (enabled) sheet.dataValidatedCells?.add(cellKey(row, column));
          else sheet.dataValidatedCells?.delete(cellKey(row, column));
        }
      }
    },
    rowHeight: (sheetName, rowIndex, store = "original") => {
      const sheet = storeOf(sheetName, store);
      return sheet.rowHeights.get(rowIndex) ?? 15;
    },
    columnWidth: (sheetName, columnIndex, store = "original") => {
      const sheet = storeOf(sheetName, store);
      return sheet.columnWidths.get(columnIndex) ?? 8.43;
    },
    setRowHeight: (sheetName, rowIndex, value, store = "original") => {
      const sheet = storeOf(sheetName, store);
      sheet.rowHeights.set(rowIndex, value);
    },
    setColumnWidth: (sheetName, columnIndex, value, store = "original") => {
      const sheet = storeOf(sheetName, store);
      sheet.columnWidths.set(columnIndex, value);
    },
    setDisplayedDefault: (sheetName, defaults, store = "original") => {
      const sheet = storeOf(sheetName, store);
      const fillColor = defaults.fillColor;
      // 样式背景给了非空填充色即意味着显示有填充：pattern 从 None 变为 Solid。
      let fillPattern = defaults.fillPattern;
      if (fillColor != null && fillColor !== "" && fillPattern == null) {
        fillPattern = "Solid";
      }
      sheet.displayedDefaults = {
        ...sheet.displayedDefaults,
        ...defaults,
        ...(fillPattern != null ? { fillPattern } : {})
      };
    }
  };
}

/** 便捷：把替身装到全局（测试 beforeEach 里用）。 */
export function installOfficeTestHost(host: TestOfficeHost): void {
  (globalThis as Record<string, unknown>).Excel = host.Excel;
  (globalThis as Record<string, unknown>).Office = host.Office;
}
