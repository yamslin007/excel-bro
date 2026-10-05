import type { AnalysisPlan, CellValue } from "./contracts";
import type { SheetSnapshot } from "./excelNativeComparison";
import { nativeColumnName, nativeReviewPackage, nativeReviewXml as xml } from "./nativeReviewFile";
import { parseExplicitA1Rect } from "./planCore";

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const color = (value?: string) => value && /^#[0-9a-f]{6}$/i.test(value) ? `FF${value.slice(1).toUpperCase()}` : undefined;
function attribute(name: string, value: string | number | boolean | undefined): string {
  return value == null ? "" : ` ${name}="${xml(String(typeof value === 'boolean' ? Number(value) : value))}"`;
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
export function buildManualReviewFile(sheetName: string, source: SheetSnapshot, plan: AnalysisPlan, maxCells: number, maxBytes: number): { base64: string; markedCells: number } {
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
  for (const action of plan.actions) {
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
  const styles = styleTable();
  let rows = '';
  for (let row=firstRow; row<=lastRow; row++) {
    let cells = '';
    for (let column=firstColumn; column<=lastColumn; column++) {
      const address = `${nativeColumnName(column)}${row}`, r = row-source.startRow, c = column-source.startColumn;
      const value = updates.has(address) ? updates.get(address)! : source.values[r]?.[c] ?? null;
      const formula = formulas.get(address), style = styles.index(source.properties[r]?.[c], source.numberFormat[r]?.[c] ?? 'General', marked.has(address));
      const head = `<c r="${address}" s="${style}"`;
      if (formula) cells += `${head}><f>${xml(formula.slice(1))}</f></c>`;
      else if (value === null) cells += `${head}/>`;
      else if (typeof value === 'number') {
        if (!Number.isFinite(value)) throw new Error('副本包含非有限数值。');
        cells += `${head}><v>${value}</v></c>`;
      } else if (typeof value === 'boolean') cells += `${head} t="b"><v>${Number(value)}</v></c>`;
      else cells += `${head} t="inlineStr"><is><t xml:space="preserve">${xml(value)}</t></is></c>`;
    }
    const height = source.rowHeights[row-source.startRow];
    rows += `<row r="${row}"${height > 0 && Number.isFinite(height) ? ` ht="${height}" customHeight="1"` : ''}>${cells}</row>`;
  }
  // Office widths are points; OOXML widths are font character units (approximate).
  const columns = source.columnWidths.map((width,c) => width > 0 && Number.isFinite(width) ? `<col min="${source.startColumn+c}" max="${source.startColumn+c}" width="${Math.max(0,(width*96/72-5)/7).toFixed(4)}" customWidth="1"/>` : '').join('');
  const topLeft = `${nativeColumnName(firstColumn)}${firstRow}`;
  const sheetXml = `<worksheet xmlns="${NS}"><dimension ref="${topLeft}:${nativeColumnName(lastColumn)}${lastRow}"/><sheetViews><sheetView workbookViewId="0" topLeftCell="${topLeft}"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>${columns}</cols><sheetData>${rows}</sheetData></worksheet>`;
  return { base64: nativeReviewPackage(sheetName, sheetXml, maxBytes, styles.xml()), markedCells: marked.size };
}
