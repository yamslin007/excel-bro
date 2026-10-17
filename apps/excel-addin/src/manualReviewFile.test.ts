// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import type { AnalysisPlan } from './contracts';
import type { SheetSnapshot } from './excelNativeComparison';
import { buildManualReviewFile } from './manualReviewFile';
import { readSeedEntries } from './nativeReviewTestHelpers';

const source = (): SheetSnapshot => ({id:'original',name:'Ledger',address:'B5:D6',startRow:5,startColumn:2,
  values:[['LOCAL_PRIVATE_SENTINEL',7,8],['=LITERAL',9,99]],formulas:[['LOCAL_PRIVATE_SENTINEL',7,'=C5+1'],['=LITERAL',9,99]],
  formulaCells:['D5'],numberFormat:[['General','0.00','0.00'],['@','General','General']],
  properties:[[{format:{font:{name:'Arial',bold:true,size:12},fill:{color:'#FFF000'},wrapText:true,borders:{bottom:{style:'Continuous',weight:'Medium',color:'#000000'}}}},{},{}],[{},{},{}]],
  rowHeights:[24,30],columnWidths:[60,90,120]});
const p = (actions: AnalysisPlan['actions']): AnalysisPlan => ({id:'plan',title:'Review',summary:'Review',assumptions:[],warnings:[],actions});
function read(plan: AnalysisPlan, snapshot=source()) {
  const output = buildManualReviewFile('EB_Review',snapshot,plan,20000,8388608);
  const files = readSeedEntries(output.base64), parser = new DOMParser();
  const sheet = parser.parseFromString(files['xl/worksheets/sheet1.xml'],'application/xml');
  const styles = parser.parseFromString(files['xl/styles.xml'],'application/xml');
  expect(sheet.querySelector('parsererror')).toBeNull(); expect(styles.querySelector('parsererror')).toBeNull();
  return {output,files,sheet,styles};
}
describe('compatibility OOXML projection', () => {
  it('keeps real formulas as formulas, constants as constants and requests actual Excel recalculation without fake cached values', () => {
    const {files,sheet,output}=read(p([{type:'writeFormulas',sheet:'Ledger',range:'D6',formulas:[['=C6+1']]}]));
    expect(JSON.stringify(source())).toContain('LOCAL_PRIVATE_SENTINEL');
    expect(sheet.querySelector('c[r="B5"] t')?.textContent).toBe('LOCAL_PRIVATE_SENTINEL');
    expect(sheet.querySelector('c[r="B6"]')?.getAttribute('t')).toBe('inlineStr');
    expect(sheet.querySelector('c[r="B6"] f')).toBeNull();
    expect(sheet.querySelector('c[r="D5"] f')?.textContent).toBe('C5+1');
    expect(sheet.querySelector('c[r="D6"] f')?.textContent).toBe('C6+1');
    expect(sheet.querySelector('c[r="D6"] v')).toBeNull();
    expect(files['xl/workbook.xml']).toContain('fullCalcOnLoad="1"');
    expect(files['[Content_Types].xml']).toContain('/xl/styles.xml');
    expect(files['xl/_rels/workbook.xml.rels']).toContain('Target="styles.xml"');
    expect(output.markedCells).toBe(1);
  });
  it('preserves number formats, font, borders, wrap and row heights; only projected target cells receive the green marker', () => {
    const {sheet,styles}=read(p([{type:'writeValues',sheet:'Ledger',range:'D6',values:[['done']]}]));
    expect(sheet.querySelector('row[r="5"]')?.getAttribute('ht')).toBe('24');
    expect(sheet.querySelector('col[min="3"]')?.getAttribute('width')).toBeTruthy();
    expect(styles.querySelector('numFmt')?.getAttribute('formatCode')).toBe('0.00');
    expect(styles.querySelector('font b')).toBeTruthy();
    expect(styles.querySelector('font name[val="Arial"]')).toBeTruthy();
    expect(styles.querySelector('bottom[style="medium"]')).toBeTruthy();
    const fill = (address:string) => {
      const xf = styles.querySelector('cellXfs')!.children[Number(sheet.querySelector(`c[r="${address}"]`)!.getAttribute('s'))];
      return styles.querySelector('fills')!.children[Number(xf.getAttribute('fillId'))].querySelector('fgColor')?.getAttribute('rgb');
    };
    expect(fill('B5')).toBe('FFFFF000'); expect(fill('D6')).toBe('FFE3F1E8');
    expect(fill('C6')).not.toBe('FFE3F1E8');
  });
  it('clearing a formula truly removes it, writes outside the used range retain absolute addresses, and the original snapshot is untouched', () => {
    const snapshot = source(), before = JSON.stringify(snapshot);
    const {sheet,output}=read(p([{type:'clearRange',sheet:'Ledger',range:'D5',applyTo:'contents'},
      {type:'writeValues',sheet:'Ledger',range:'F8',values:[['<&"']]}]),snapshot);
    expect(sheet.querySelector('c[r="D5"] f')).toBeNull(); expect(sheet.querySelector('c[r="D5"] v')).toBeNull();
    expect(sheet.querySelector('c[r="F8"] t')?.textContent).toBe('<&"');
    expect(output.markedCells).toBe(2); expect(JSON.stringify(snapshot)).toBe(before);
  });
  it('does not mark identical constant writes, and rejects excessive merged range or bytes before opening', () => {
    const snapshot=source(), plan=p([{type:'writeValues',sheet:'Ledger',range:'C5',values:[[7]]}]);
    expect(read(plan).output.markedCells).toBe(0);
    expect(()=>buildManualReviewFile('EB_Review',snapshot,plan,5,8388608)).toThrow('限额');
    expect(()=>buildManualReviewFile('EB_Review',snapshot,plan,20000,10)).toThrow('限额');
  });
});
// 格式动作在兼容副本里的**投影**：这是批量许可永久版（ExcelApi ≤ 1.17，无原生并排窗口）
// 实际走的路径 —— 用户在打开的副本里必须先看到"加粗/居中/行高"会变成什么样。
describe('compatibility OOXML projection · format actions', () => {
  /** 某格地址对应的 cellXfs 条目（样式投影的落点）。 */
  function xfFor(sheet: Document, styles: Document, address: string): Element {
    const index = Number(sheet.querySelector(`c[r="${address}"]`)!.getAttribute('s'));
    return styles.querySelector('cellXfs')!.children[index];
  }
  it('setFont{bold} 投影成该格字体的 <b/>，且只有该格被标绿', () => {
    const {sheet,styles,output}=read(p([{type:'setFont',sheet:'Ledger',range:'D6',bold:true}]));
    const fontId = xfFor(sheet, styles, 'D6').getAttribute('fontId')!;
    expect(styles.querySelectorAll('fonts')[0].children[Number(fontId)].querySelector('b')).toBeTruthy();
    // 未受影响的格不该被标绿。
    const neighbour = xfFor(sheet, styles, 'C6').getAttribute('fillId')!;
    expect(styles.querySelectorAll('fills')[0].children[Number(neighbour)].querySelector('fgColor')?.getAttribute('rgb')).not.toBe('FFE3F1E8');
    expect(output.markedCells).toBe(1);
  });
  it('setAlignment{horizontal:center} 投影成 alignment horizontal="center"', () => {
    const {sheet,styles}=read(p([{type:'setAlignment',sheet:'Ledger',range:'D6',horizontal:'center'}]));
    expect(xfFor(sheet, styles, 'D6').querySelector('alignment')?.getAttribute('horizontal')).toBe('center');
  });
  function readSnapshot(plan: AnalysisPlan, snapshot: SheetSnapshot) {
    const output = buildManualReviewFile('EB_Review', snapshot, plan, 20000, 8388608);
    const files = readSeedEntries(output.base64), parser = new DOMParser();
    const sheet = parser.parseFromString(files['xl/worksheets/sheet1.xml'], 'application/xml');
    const styles = parser.parseFromString(files['xl/styles.xml'], 'application/xml');
    expect(sheet.querySelector('parsererror')).toBeNull();
    return { output, files, sheet, styles };
  }
  it('setFill 如实保留动作颜色，改动标记不得覆盖审阅结果', () => {
    // 已是同色（Solid + #FF0000）：投影结果与执行前一致 ⇒ 不标记，填充色如实渲染成动作色。
    const same = source();
    same.properties[1][2] = { format: { fill: { pattern: 'Solid', color: '#FF0000' } } };
    const unchanged = readSnapshot(p([{ type: 'setFill', sheet: 'Ledger', range: 'D6', color: '#FF0000' }]), same);
    const sameFillId = xfFor(unchanged.sheet, unchanged.styles, 'D6').getAttribute('fillId')!;
    expect(
      unchanged.styles.querySelectorAll('fills')[0].children[Number(sameFillId)].querySelector('fgColor')?.getAttribute('rgb')
    ).toBe('FFFF0000');
    expect(unchanged.output.markedCells).toBe(0);

    // 新着色仍计入变化，但必须显示计划颜色。
    const fresh = readSnapshot(p([{ type: 'setFill', sheet: 'Ledger', range: 'D6', color: '#FF0000' }]), source());
    const freshFillId = xfFor(fresh.sheet, fresh.styles, 'D6').getAttribute('fillId')!;
    expect(
      fresh.styles.querySelectorAll('fills')[0].children[Number(freshFillId)].querySelector('fgColor')?.getAttribute('rgb')
    ).toBe('FFFF0000');
    expect(fresh.output.markedCells).toBe(1);
  });
  it('setNumberFormat 投影成 numFmt 格式码', () => {
    const { files } = read(p([{ type: 'setNumberFormat', sheet: 'Ledger', range: 'D6', formatCode: '0.000' }]));
    expect(files['xl/styles.xml']).toContain('formatCode="0.000"');
  });
  it('resizeRange 投影成行高/列宽，且不把任何单元格算作"格式改动"', () => {
    const {sheet,output}=read(p([{type:'resizeRange',sheet:'Ledger',range:'D6',rowHeight:40,columnWidth:25}]));
    expect(sheet.querySelector('row[r="6"]')?.getAttribute('ht')).toBe('40');
    expect(sheet.querySelector('col[min="4"]')?.getAttribute('customWidth')).toBe('1');
    // 行高/列宽是行级/列级属性，不是逐格格式改动。
    expect(output.markedCells).toBe(0);
  });
  it('投影结果与执行前一致时不标记（幂等格式动作不制造假改动）', () => {
    // B5 执行前已经是 bold:true，再 setFont{bold:true} 不产生实际变化。
    const {output}=read(p([{type:'setFont',sheet:'Ledger',range:'B5',bold:true}]));
    expect(output.markedCells).toBe(0);
  });
});
