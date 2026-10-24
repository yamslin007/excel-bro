import { BORDER_FIELDS, readWpsBorder, restoreWpsBorder } from "./wpsBorders";
/**
 * WPS 格式属性的**逐格 / 逐行快照**采集与恢复（Task #53 的宿主适配层）。
 *
 * 这一层只做 WPS JSAPI 的读写，不做「该属性是否已验证」的判定（那是
 * `wpsFormatCapability.ts` 的职责），也不做「改了哪些属性」的语义（那是共享层
 * `formatActions.ts` 的职责）。
 *
 * 为什么必须逐格/逐行，而不是区域级单值：
 * - WPS 区域级 `range.Font.Bold` 对**混合格式**（区域内有的格加粗有的不加粗）返回的
 *   值不确定（可能是 null / 第一个格 / 折中值）。用它做撤销快照会把混合格式拍平成
 *   单一值，恢复时把「有的加粗有的不加粗」写坏成「全加粗或全不加粗」。
 * - **单格探测（Sheet1!A1）不证明混合格式恢复**。因此这里一律逐格读 `range.Item(r,c)`
 *   的字体/对齐/换行，逐行读行高，恢复时逐格/逐行写回。
 *
 * 关键语义（与 Excel 端 `captureRangeFormat` 同一条规矩，对象从单元格值换成属性）：
 * - **未采集 ≠ 默认值**：读不到（undefined/null）的属性存成 `null`，恢复时**跳过**该
 *   字段，绝不写回一个"看起来像默认值"的具体值（否则会把"继承的加粗"写成"显式不加粗"）。
 * - 只要**任一格**的**任一目标属性**读不到，整份快照判 `captured:false`，由调用方对
 *   纯格式动作在**修改前**拒绝执行（不产生不可撤销的修改）。
 *
 * **采集粒度铁律（Task #53 阻断点 3）**：多格（或多行）区域**必须**能逐格 / 逐行定位
 * （`range.Item(r,c)` 可用）。缺 `Item` 时**绝不**回退到整块区域读取 —— 那会把整块的
 * 单一格式值伪装成"逐格已采集"，恢复时把混合格式写坏。多格/多行缺 `Item` ⇒
 * `captured:false`，调用方据此在写入前拒绝。
 */

import type { WpsRange } from "./wpsHost";

/** 快照覆盖的全部格式字段名（用于属性级恢复的字段集合）。 */
export const ALL_WPS_FORMAT_FIELDS = [
  ...BORDER_FIELDS,
  "bold",
  "italic",
  "name",
  "size",
  "color",
  "horizontalAlignment",
  "verticalAlignment",
  "wrapText",
  "rowHeight",
  "fillColor", "fillPattern", "fillPatternColor", "numberFormat"
] as const;

export type WpsFormatField = (typeof ALL_WPS_FORMAT_FIELDS)[number];

/** 单格的字体/对齐/换行快照。`null` = 未直接设置或未采集，恢复时跳过。 */
export type WpsCellFormatSnapshot = {
  borderTop?: string | null;
  borderBottom?: string | null;
  borderLeft?: string | null;
  borderRight?: string | null;
  bold: boolean | null;
  italic: boolean | null;
  name: string | null;
  size: number | null;
  color: number | null;
  horizontalAlignment: number | null;
  verticalAlignment?: number | null;
  wrapText: boolean | null;
  fillColor?: number | null;
  fillPattern?: number | null;
  fillPatternColor?: number | null;
  numberFormat?: string | null;
};

/** 整块区域的格式快照：逐格字体/对齐/换行 + 逐行行高。 */
export type WpsFormatSnapshot = {
  /** [row][column] 的逐格格式。 */
  cells: WpsCellFormatSnapshot[][];
  /** 逐行行高（行级属性）。`null` = 未采集。 */
  rowHeights: Array<number | null>;
  /** 整份快照是否完整采集（所有目标属性都读到了）。false = 不可用于撤销。 */
  captured: boolean;
};

function readNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function readBoolean(value: unknown): boolean | null {
  return typeof value === "boolean" ? value : null;
}

function readString(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

/** New optional host properties may throw on older wrappers. Missing evidence blocks only actions that require it. */
function optionalFormat<T>(read: () => unknown, normalize: (value: unknown) => T | null): T | null {
  try { return normalize(read()); } catch { return null; }
}

/** 读单格（`range.Item(r+1, c+1)`）的格式快照。**不回退到整块**。 */
function readCellFormat(cell: WpsRange): WpsCellFormatSnapshot {
  const font = cell.Font;
  return {
    ...Object.fromEntries(BORDER_FIELDS.map(field=>[field,readWpsBorder(cell,field)])),
    bold: readBoolean(font?.Bold),
    italic: readBoolean(font?.Italic),
    name: readString(font?.Name),
    size: readNumber(font?.Size),
    color: readNumber(font?.Color),
    horizontalAlignment: readNumber(cell.HorizontalAlignment),
    verticalAlignment: optionalFormat(() => cell.VerticalAlignment, readNumber),
    wrapText: readBoolean(cell.WrapText),
    fillColor: optionalFormat(() => cell.Interior?.Color, readNumber),
    fillPattern: optionalFormat(() => cell.Interior?.Pattern, readNumber),
    fillPatternColor: optionalFormat(() => cell.Interior?.PatternColor, readNumber),
    numberFormat: optionalFormat(() => cell.NumberFormat, readString)
  };
}

/**
 * 取单格。**单格区域（rows===1 && columns===1）时 `range` 本身就是单元格**，
 * 直接返回 range（单格区域没有 `Item` 是正常的）；多格/多行时调用方已保证
 * `range.Item` 可用，才用 `Item(r+1, c+1)` 逐格定位。
 */
function cellAt(range: WpsRange, row: number, column: number, singleCell: boolean): WpsRange {
  return singleCell ? range : range.Item!(row + 1, column + 1);
}

/**
 * 采集整块区域的格式快照（逐格字体/对齐/换行 + 逐行行高）。
 *
 * 采集粒度铁律：多格（`rows * columns > 1`）或多行（`rows > 1`，行高是行级属性）
 * 时，`range.Item` 缺失 ⇒ 无法逐格 / 逐行定位 ⇒ `captured:false`，**不回退到整块**。
 *
 * 任一格的目标属性读不到（`null`）时，该格的快照仍保留，但整份 `captured` 置 false
 * —— 调用方据此对纯格式动作执行前拒绝。行高逐行读（行级属性，读该行第一个单元格
 * 的 `RowHeight` 即可）。
 */
export function captureWpsFormatSnapshot(range: WpsRange, rows: number, columns: number): WpsFormatSnapshot {
  const singleCell = rows === 1 && columns === 1;
  const multiCell = rows * columns > 1;
  const multiRow = rows > 1;
  // 多格需要逐格定位、多行需要逐行定位；缺 `Item` 一律 captured=false，不回退整块。
  // （单格时 `range` 本身就是单元格，不需要 `Item`。）
  if ((multiCell || multiRow) && typeof range.Item !== "function") {
    return { cells: [], rowHeights: [], captured: false };
  }

  const cells: WpsCellFormatSnapshot[][] = [];
  let captured = true;

  for (let r = 0; r < rows; r += 1) {
    const rowCells: WpsCellFormatSnapshot[] = [];
    for (let c = 0; c < columns; c += 1) {
      const cell = cellAt(range, r, c, singleCell);
      const cellSnapshot = readCellFormat(cell);
      // 字体五个属性 + 对齐 + 换行，任一为 null 即该格未完整采集。
      if (
        cellSnapshot.bold === null ||
        cellSnapshot.italic === null ||
        cellSnapshot.name === null ||
        cellSnapshot.size === null ||
        cellSnapshot.color === null ||
        cellSnapshot.horizontalAlignment === null ||
        cellSnapshot.wrapText === null
      ) {
        captured = false;
      }
      rowCells.push(cellSnapshot);
    }
    cells.push(rowCells);
  }

  const rowHeights: Array<number | null> = [];
  for (let r = 0; r < rows; r += 1) {
    const cell = cellAt(range, r, 0, singleCell);
    const height = readNumber(cell.RowHeight);
    if (height === null) captured = false;
    rowHeights.push(height);
  }

  return { cells, rowHeights, captured };
}

/** 逐格写回字体/对齐/换行（恢复执行前快照）。`null` 字段跳过，不写默认值。 */
function restoreCellFormat(
  cell: WpsRange,
  saved: WpsCellFormatSnapshot,
  fields: ReadonlySet<WpsFormatField>
): void {
  for(const field of BORDER_FIELDS) if(fields.has(field) && saved[field]!=null) restoreWpsBorder(cell,field,saved[field]);
  if (fields.has("numberFormat") && saved.numberFormat != null) cell.NumberFormat = saved.numberFormat;
  if ((fields.has("fillColor") || fields.has("fillPattern") || fields.has("fillPatternColor")) && cell.Interior) {
    if (fields.has("fillColor") && saved.fillColor != null) cell.Interior.Color = saved.fillColor;
    if (fields.has("fillPatternColor") && saved.fillPatternColor != null) cell.Interior.PatternColor = saved.fillPatternColor;
    // Pattern last: assigning Color can implicitly enable a solid fill.
    if (fields.has("fillPattern") && saved.fillPattern != null) cell.Interior.Pattern = saved.fillPattern;
  }
  const font = cell.Font;
  if (font) {
    if (fields.has("bold") && saved.bold !== null) font.Bold = saved.bold;
    if (fields.has("italic") && saved.italic !== null) font.Italic = saved.italic;
    if (fields.has("name") && saved.name !== null) font.Name = saved.name;
    if (fields.has("size") && saved.size !== null) font.Size = saved.size;
    if (fields.has("color") && saved.color !== null) font.Color = saved.color;
  }
  if (fields.has("horizontalAlignment") && saved.horizontalAlignment !== null) {
    cell.HorizontalAlignment = saved.horizontalAlignment;
  }
  if (fields.has("verticalAlignment") && saved.verticalAlignment != null) cell.VerticalAlignment = saved.verticalAlignment;
  if (fields.has("wrapText") && saved.wrapText !== null) {
    cell.WrapText = saved.wrapText;
  }
}

/**
 * 恢复整块区域的格式快照（逐格 + 逐行）。
 *
 * `fields` 控制**只恢复哪些属性**（属性级恢复，Task #53 阻断点 2）：撤销时传
 * "本次动作实际改动过的属性"，从而**保留用户对其他属性的后续修改**，不整块覆盖。
 * 缺省恢复全部字段（`ALL_WPS_FORMAT_FIELDS`），用于快照测试与"整块恢复"语义。
 */
export function restoreWpsFormatSnapshot(
  range: WpsRange,
  snapshot: WpsFormatSnapshot,
  fields: readonly WpsFormatField[] = ALL_WPS_FORMAT_FIELDS
): void {
  const fieldSet = new Set<WpsFormatField>(fields);
  const singleCell = snapshot.cells.length === 1 && snapshot.cells[0]?.length === 1;
  for (let r = 0; r < snapshot.cells.length; r += 1) {
    for (let c = 0; c < snapshot.cells[r].length; c += 1) {
      restoreCellFormat(cellAt(range, r, c, singleCell), snapshot.cells[r][c], fieldSet);
    }
  }
  if (fieldSet.has("rowHeight")) {
    for (let r = 0; r < snapshot.rowHeights.length; r += 1) {
      const height = snapshot.rowHeights[r];
      if (height === null) continue;
      cellAt(range, r, 0, singleCell).RowHeight = height;
    }
  }
}

/** 两份格式快照是否等价（用于回读比对 / 撤销后核对）。逐格逐行严格比对。 */
export function sameWpsFormatSnapshot(left: WpsFormatSnapshot, right: WpsFormatSnapshot): boolean {
  if (left.cells.length !== right.cells.length) return false;
  for (let r = 0; r < left.cells.length; r += 1) {
    if (left.cells[r].length !== right.cells[r].length) return false;
    for (let c = 0; c < left.cells[r].length; c += 1) {
      const a = left.cells[r][c];
      const b = right.cells[r][c];
      if (
        BORDER_FIELDS.some(field=>a[field]!==b[field]) ||
        a.bold !== b.bold ||
        a.italic !== b.italic ||
        a.name !== b.name ||
        a.size !== b.size ||
        a.color !== b.color ||
        a.horizontalAlignment !== b.horizontalAlignment ||
        a.verticalAlignment !== b.verticalAlignment ||
        a.wrapText !== b.wrapText ||
        a.fillColor !== b.fillColor || a.fillPattern !== b.fillPattern ||
        a.fillPatternColor !== b.fillPatternColor || a.numberFormat !== b.numberFormat
      ) {
        return false;
      }
    }
  }
  if (left.rowHeights.length !== right.rowHeights.length) return false;
  for (let r = 0; r < left.rowHeights.length; r += 1) {
    if (left.rowHeights[r] !== right.rowHeights[r]) return false;
  }
  return true;
}

/**
 * 「区域内**数据行**的内容级格式是否一致」探针 —— 排序 / 去重的前置守卫。
 *
 * 为什么需要它：排序只重排**值**，单元格格式留在原位。如果同一个区域里有的行
 * 加粗、有的行是红色、有的行是百分比格式，那么"值排序"之后格式就与数据脱钩了，
 * 用户看到的是"数据换了位置、格式没跟着走"。这类区域必须先被拒绝，或由宿主原生
 * 排序把格式一起搬 —— 本项目不冒这个险，直接拒绝并说清原因。
 *
 * 判据边界（刻意排除的两类属性）：
 *  - **边框**：边框是网格属性，相邻单元格共用一条物理边；它本来就不随行移动
 *    （见 `wpsBorders` 的归属检查），把它算进"格式一致性"会让几乎所有真实表被拒。
 *  - **行高**：行高属于物理行，不属于数据行，同样不随排序移动。
 *
 * 只读，不写任何属性。`fromRow` 用于跳过表头行（表头格式与数据行不同是正常的）。
 */
export function wpsUniformCellStyle(
  range: WpsRange,
  rows: number,
  columns: number,
  fromRow = 0
): { captured: boolean; uniform: boolean; detail: string } {
  const singleCell = rows === 1 && columns === 1;
  if (rows * columns > 1 && typeof range.Item !== "function") {
    return { captured: false, uniform: false, detail: "宿主的区域对象不支持逐格定位，无法核对格式一致性" };
  }
  const signatures = new Set<string>();
  for (let r = fromRow; r < rows; r += 1) {
    for (let c = 0; c < columns; c += 1) {
      const cell = cellAt(range, r, c, singleCell);
      const snapshot = readCellFormat(cell);
      // 与 `captureWpsFormatSnapshot` 同一条纪律：读不到就不是"默认值"。
      if (
        snapshot.bold === null || snapshot.italic === null || snapshot.name === null ||
        snapshot.size === null || snapshot.color === null ||
        snapshot.horizontalAlignment === null || snapshot.wrapText === null
      ) {
        return {
          captured: false,
          uniform: false,
          detail: `第 ${r + 1} 行第 ${c + 1} 列的字体/对齐无法完整读取`
        };
      }
      signatures.add(JSON.stringify([
        snapshot.bold, snapshot.italic, snapshot.name, snapshot.size, snapshot.color,
        snapshot.horizontalAlignment, snapshot.verticalAlignment, snapshot.wrapText,
        snapshot.fillColor ?? null, snapshot.fillPattern ?? null,
        snapshot.fillPatternColor ?? null, snapshot.numberFormat ?? null
      ]));
      if (signatures.size > 1) {
        return { captured: true, uniform: false, detail: "区域内存在两种以上内容级格式" };
      }
    }
  }
  return { captured: true, uniform: true, detail: "区域内数据行格式一致" };
}
