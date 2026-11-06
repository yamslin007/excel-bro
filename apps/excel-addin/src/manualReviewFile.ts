import type { AnalysisPlan, CellValue, ExcelAction } from "./contracts";
import { isFormatAction } from "./contracts";
import type { SheetSnapshot } from "./excelNativeComparison";
import { nativeColumnName, nativeReviewPackage, nativeReviewWorkbookPackage, nativeReviewXml as xml } from "./nativeReviewFile";
import { parseExplicitA1Rect } from "./planCore";
import { excelFormatActionFields } from "./excelFormatCapability";
import { directFormatFieldValue } from "./excel";

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const color = (value?: string) => value && /^#[0-9a-f]{6}$/i.test(value) ? `FF${value.slice(1).toUpperCase()}` : undefined;
function attribute(name: string, value: string | number | boolean | undefined): string {
  return value == null ? "" : ` ${name}="${xml(String(typeof value === 'boolean' ? Number(value) : value))}"`;
}

/** 深拷贝一份 `CellPropertiesFormat`：投影格式动作时会逐格改写嵌套的 font/fill/borders，不能共享引用。 */
function cloneCellFormat(source: Excel.CellPropertiesFormat | undefined): Excel.CellPropertiesFormat {
  if (!source) return {};
  return {
    ...source,
    ...(source.font ? { font: { ...source.font } } : {}),
    ...(source.fill ? { fill: { ...source.fill } } : {}),
    ...(source.borders
      ? {
          borders: Object.fromEntries(
            Object.entries(source.borders).map(([side, state]) => [
              side,
              state ? { ...(state as object) } : state
            ])
          )
        }
      : {}),
    ...(source.protection ? { protection: { ...source.protection } } : {})
  } as Excel.CellPropertiesFormat;
}

/** 格式动作里"线型/线宽"在 OOXML 侧用首字母大写的词（`Continuous` / `Thin`）。 */
const capitalize = (value: string) => value.charAt(0).toUpperCase() + value.slice(1);

/**
 * 把一个**格式动作**投影到某格的直接格式上（纯数据变换，不碰宿主）。
 *
 * 与执行器 `applyFormatAction` 的语义逐字对齐：`setFill` 设色即视为有填充（`Solid`）；
 * `setFont`/`setAlignment` 只改显式给出的属性；`setBorders` 补齐缺省值后写指定边；
 * `setNumberFormat` 由调用方另行按格记录。
 */
function projectFormatAction(target: Excel.CellPropertiesFormat, action: ExcelAction): void {
  switch (action.type) {
    case "setFill":
      target.fill = { pattern: "Solid", color: action.color } as Excel.CellPropertiesFormat["fill"];
      return;
    case "setFont": {
      const font = { ...(target.font ?? {}) } as NonNullable<Excel.CellPropertiesFormat["font"]>;
      if (action.bold != null) font.bold = action.bold;
      if (action.italic != null) font.italic = action.italic;
      if (action.name != null) font.name = action.name;
      if (action.size != null) font.size = action.size;
      if (action.color != null) font.color = action.color;
      target.font = font;
      return;
    }
    case "setAlignment": {
      if (action.horizontal != null) target.horizontalAlignment = capitalize(action.horizontal) as Excel.CellPropertiesFormat["horizontalAlignment"];
      if (action.vertical != null) target.verticalAlignment = capitalize(action.vertical) as Excel.CellPropertiesFormat["verticalAlignment"];
      if (action.wrapText != null) target.wrapText = action.wrapText;
      return;
    }
    case "setBorders": {
      const style = capitalize(action.style ?? "continuous");
      const weight = capitalize(action.weight ?? "thin");
      const borderColor = action.color ?? "#000000";
      const borders = { ...(target.borders ?? {}) } as Record<string, { style: string; color: string; weight: string }>;
      for (const side of action.sides) borders[side] = { style, color: borderColor, weight };
      target.borders = borders as Excel.CellPropertiesFormat["borders"];
      return;
    }
    default:
      return;
  }
}

function styleTable() {
  const fonts = ['<font><sz val="11"/><name val="Calibri"/></font>'];
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const formats: string[] = [], xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const add = (items: string[], entry: string) => { const old = items.indexOf(entry); if (old >= 0) return old; items.push(entry); return items.length - 1; };
  function index(properties: Excel.CellProperties | undefined, numberFormat: string, changed: boolean): number {
    const f = properties?.format ?? {}, font = f.font ?? {};
    const underline = ({Single:'single',Double:'double',SingleAccountant:'singleAccounting',DoubleAccountant:'doubleAccounting'} as Record<string,string>)[font.underline ?? 'None'];
    const fontId = add(fonts, `<font><sz val="${Number.isFinite(font.size) && font.size! > 0 ? font.size : 11}"/><name val="${xml(font.name ?? 'Calibri')}"/>${font.bold ? '<b/>' : ''}${font.italic ? '<i/>' : ''}${color(font.color) ? `<color rgb="${color(font.color)}"/>` : ''}${underline ? `<u val="${underline}"/>` : ''}</font>`);
    const rgb = changed ? 'FFE3F1E8' : color(f.fill?.color);
    const fillId = rgb ? add(fills, `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const borderStyles: Record<string, string> = { Continuous:'thin', Dash:'dashed', DashDot:'dashDot', DashDotDot:'dashDotDot', Dot:'dotted', Double:'double', SlantDashDot:'slantDashDot' };
    const borderXml = (['left','right','top','bottom'] as const).map(side => {
      const b = f.borders?.[side], base = b && borderStyles[b.style ?? 'None'];
      const style = base === 'thin' ? ({Hairline:'hair',Thin:'thin',Medium:'medium',Thick:'thick'} as Record<string,string>)[b?.weight ?? 'Thin'] ?? 'thin' : base;
      return style ? `<${side} style="${style}">${color(b?.color) ? `<color rgb="${color(b?.color)}"/>` : ''}</${side}>` : `<${side}/>`;
    }).join('');
    const borderId = add(borders, `<border>${borderXml}<diagonal/></border>`);
    const numFmtId = !numberFormat || numberFormat === 'General' ? 0 : 164 + add(formats, numberFormat);
    const horizontal = ({ General:'general', Left:'left', Center:'center', Right:'right', Fill:'fill', Justify:'justify', CenterAcrossSelection:'centerContinuous', Distributed:'distributed' } as Record<string,string>)[f.horizontalAlignment ?? 'General'];
    const vertical = ({ Top:'top', Center:'center', Bottom:'bottom', Justify:'justify', Distributed:'distributed' } as Record<string,string>)[f.verticalAlignment ?? 'Bottom'];
    const rotation = f.textOrientation === undefined ? undefined : f.textOrientation < 0 ? 90 - f.textOrientation : f.textOrientation;
    const alignment = `<alignment${attribute('horizontal',horizontal)}${attribute('vertical',vertical)}${attribute('wrapText',f.wrapText)}${attribute('indent',f.indentLevel)}${attribute('shrinkToFit',f.shrinkToFit)}${attribute('textRotation',rotation)}/>`;
    return add(xfs, `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">${alignment}</xf>`);
  }
  return { index, xml: () => `<styleSheet xmlns="${NS}"><numFmts count="${formats.length}">${formats.map((f,i) => `<numFmt numFmtId="${164+i}" formatCode="${xml(f)}"/>`).join('')}</numFmts><fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>` };
}

// Projection only: never evaluates formulas or calls host setters. The plan has already
// passed nativeComparisonPlanIssue. Excel recalculates this standalone workbook on open.
export function buildManualReviewFile(sheetName: string, source: SheetSnapshot, plan: AnalysisPlan, maxCells: number, maxBytes: number, reviewCopyId?: string): { base64: string; markedCells: number } {
  const styles = styleTable();
  const sheet = manualReviewSheet(source, plan, maxCells, styles);
  return {base64: nativeReviewPackage(sheetName, sheet.xml, maxBytes, styles.xml(), reviewCopyId), markedCells: sheet.markedCells};
}
/** One immutable before-workbook for the whole execution; all formulas use captured values. */
export function buildBeforeReviewWorkbook(sources: SheetSnapshot[], plan: AnalysisPlan, maxCells: number, maxBytes: number, reviewCopyId: string, selected: string) {
  if (sources.reduce((sum, s) => sum + s.values.reduce((n, row) => n + row.length, 0), 0) > maxCells)
    throw new Error('执行前比对工作簿超过本机范围限额。');
  const styles = styleTable();
  const sheets = sources.map(source => ({name: source.name,
    xml: manualReviewSheet({...source, formulaCells: []}, {...plan, actions: []}, maxCells, styles).xml}));
  const activeTab = sources.findIndex(s => s.name === selected);
  if (activeTab < 0) throw new Error('对比目标不在执行前快照内。');
  return {base64: nativeReviewWorkbookPackage(sheets, maxBytes, styles.xml(), reviewCopyId, activeTab)};
}
function manualReviewSheet(source: SheetSnapshot, plan: AnalysisPlan, maxCells: number, styles: ReturnType<typeof styleTable>) {
  const sourceRect = parseExplicitA1Rect(source.address)!;
  const rects = [sourceRect, ...plan.actions.map(action => parseExplicitA1Rect('range' in action ? action.range : '')!)];
  const firstRow = Math.min(...rects.map(r => r.startRow)), lastRow = Math.max(...rects.map(r => r.endRow));
  const firstColumn = Math.min(...rects.map(r => r.startColumn)), lastColumn = Math.max(...rects.map(r => r.endColumn));
  if ((lastRow-firstRow+1)*(lastColumn-firstColumn+1) > maxCells) throw new Error('兼容审阅副本超过本机范围限额。');
  const formulas = new Map(source.formulaCells.map(address => {
    const r = parseExplicitA1Rect(address)!;
    return [address, source.formulas[r.startRow-source.startRow][r.startColumn-source.startColumn] as string];
  }));
  const updates = new Map<string, CellValue>(), marked = new Set<string>();
  // 格式动作的投影结果（逐格）：与值更新分开存，因为格式是"叠加"而不是"替换"。
  const projectedFormat = new Map<string, Excel.CellPropertiesFormat>();
  const projectedNumberFormat = new Map<string, string>();
  // 行高/列宽投影：先复制源的逐行/逐列尺寸，再按 resizeRange 覆盖命中行列。
  const projectedRowHeights = [...source.rowHeights];
  const projectedColumnWidths = [...source.columnWidths];
  const sourceFormatAt = (row: number, column: number): Excel.CellPropertiesFormat | undefined =>
    source.properties[row - source.startRow]?.[column - source.startColumn]?.format;
  const projectedFormatAt = (row: number, column: number, address: string): Excel.CellPropertiesFormat => {
    const existing = projectedFormat.get(address);
    if (existing) return existing;
    const clone = cloneCellFormat(sourceFormatAt(row, column));
    projectedFormat.set(address, clone);
    return clone;
  };
  for (const action of plan.actions) {
    if (isFormatAction(action)) {
      const rect = parseExplicitA1Rect(action.range)!;
      if (action.type === "resizeRange") {
        if (action.rowHeight != null) {
          for (let row = rect.startRow; row <= rect.endRow; row += 1) {
            projectedRowHeights[row - source.startRow] = action.rowHeight;
          }
        }
        if (action.columnWidth != null) {
          for (let column = rect.startColumn; column <= rect.endColumn; column += 1) {
            projectedColumnWidths[column - source.startColumn] = action.columnWidth;
          }
        }
        continue;
      }
      const fields = excelFormatActionFields(action);
      for (let r = 0; r < rect.rows; r++) for (let c = 0; c < rect.columns; c++) {
        const row = rect.startRow + r, column = rect.startColumn + c;
        const address = `${nativeColumnName(column)}${row}`;
        const before = cloneCellFormat(projectedFormatAt(row, column, address));
        const after = cloneCellFormat(before);
        projectFormatAction(after, action);
        // 只比**该动作实际修改的属性**：未涉及的属性变化不该被算成"本次改动"。
        const digest = (format: Excel.CellPropertiesFormat) =>
          JSON.stringify(fields.map((field) => directFormatFieldValue(format, field)));
        if (digest(before) !== digest(after)) marked.add(address);
        projectedFormat.set(address, after);
        if (action.type === "setNumberFormat") projectedNumberFormat.set(address, action.formatCode);
      }
      continue;
    }
    if (action.type !== 'writeValues' && action.type !== 'writeFormulas' && action.type !== 'clearRange') throw new Error('不支持的兼容审阅动作。');
    const rect = parseExplicitA1Rect(action.range)!;
    for (let r=0; r<rect.rows; r++) for (let c=0; c<rect.columns; c++) {
      const row = rect.startRow+r, column = rect.startColumn+c, address = `${nativeColumnName(column)}${row}`;
      const old = source.values[row-source.startRow]?.[column-source.startColumn];
      if (action.type === 'writeFormulas') {
        formulas.set(address, action.formulas![r][c] as string);
        updates.set(address, null); // No cached value may masquerade as a host-calculated result.
        marked.add(address);
      } else {
        const value = action.type === 'writeValues' ? action.values[r][c] : null;
        if ((old ?? '') !== (value ?? '') || formulas.has(address)) marked.add(address);
        formulas.delete(address); updates.set(address, value);
      }
    }
  }
  let rows = '';
  for (let row=firstRow; row<=lastRow; row++) {
    let cells = '';
    for (let column=firstColumn; column<=lastColumn; column++) {
      const address = `${nativeColumnName(column)}${row}`, r = row-source.startRow, c = column-source.startColumn;
      const value = updates.has(address) ? updates.get(address)! : source.values[r]?.[c] ?? null;
      const formula = formulas.get(address);
      // 投影后的格式优先：格式动作改过的格子必须按**动作生效后**的样子渲染，
      // 否则副本上看不出"加粗/居中/行高"到底会变成什么。
      const properties: Excel.CellProperties | undefined = projectedFormat.has(address)
        ? { format: projectedFormat.get(address) }
        : source.properties[r]?.[c];
      const numberFormat = projectedNumberFormat.get(address) ?? source.numberFormat[r]?.[c] ?? 'General';
      // Keep format previews visually faithful; a green marker must not overwrite
      // either the requested fill or the cell's unchanged background.
      const style = styles.index(properties, numberFormat, marked.has(address) && !projectedFormat.has(address));
      const head = `<c r="${address}" s="${style}"`;
      if (formula) cells += `${head}><f>${xml(formula.slice(1))}</f></c>`;
      else if (value === null) cells += `${head}/>`;
      else if (typeof value === 'number') {
        if (!Number.isFinite(value)) throw new Error('副本包含非有限数值。');
        cells += `${head}><v>${value}</v></c>`;
      } else if (typeof value === 'boolean') cells += `${head} t="b"><v>${Number(value)}</v></c>`;
      else cells += `${head} t="inlineStr"><is><t xml:space="preserve">${xml(value)}</t></is></c>`;
    }
    const height = projectedRowHeights[row-source.startRow];
    rows += `<row r="${row}"${height > 0 && Number.isFinite(height) ? ` ht="${height}" customHeight="1"` : ''}>${cells}</row>`;
  }
  // Office widths are points; OOXML widths are font character units (approximate).
  const columns = projectedColumnWidths.map((width,c) => width > 0 && Number.isFinite(width) ? `<col min="${source.startColumn+c}" max="${source.startColumn+c}" width="${Math.max(0,(width*96/72-5)/7).toFixed(4)}" customWidth="1"/>` : '').join('');
  const topLeft = `${nativeColumnName(firstColumn)}${firstRow}`;
  const focus = source.focusAddress && parseExplicitA1Rect(source.focusAddress) ? source.focusAddress : topLeft;
  const merged = (source.mergedAreas ?? []).filter(area => parseExplicitA1Rect(area));
  const merges = merged.length ? `<mergeCells count="${merged.length}">${merged.map(area => `<mergeCell ref="${xml(area)}"/>`).join('')}</mergeCells>` : '';
  const sheetXml = `<worksheet xmlns="${NS}"><dimension ref="${topLeft}:${nativeColumnName(lastColumn)}${lastRow}"/><sheetViews><sheetView workbookViewId="0" topLeftCell="${focus}"><selection activeCell="${focus}" sqref="${focus}"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>${columns}</cols><sheetData>${rows}</sheetData>${merges}</worksheet>`;
  return { xml: sheetXml, markedCells: marked.size };
}
