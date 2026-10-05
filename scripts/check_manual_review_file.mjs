// Offline verification only. Builds a synthetic in-memory workbook with the
// production browser serializer and opens it with an independent OOXML reader.
// No real Office host, workbook, service, or on-disk spreadsheet is involved.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { build } from 'esbuild';

const root = path.resolve(import.meta.dirname, '..');
const bundle = await build({ entryPoints: [path.join(root, 'apps/excel-addin/src/manualReviewFile.ts')],
  bundle: true, platform: 'node', format: 'cjs', write: false });
const module = {exports: {}};
new Function('module', 'exports', 'require', bundle.outputFiles[0].text)(module, module.exports, createRequire(import.meta.url));
const source = {id: 'synthetic', name: 'Synthetic', address: 'B5:D6', startRow: 5, startColumn: 2,
  values: [['LOCAL_PRIVATE_SENTINEL', 7, 8], ['=LITERAL', true, 99]],
  formulas: [['LOCAL_PRIVATE_SENTINEL', 7, '=C5+1'], ['=LITERAL', true, 99]],
  formulaCells: ['D5'], numberFormat: [['General', '0.00', '0.00'], ['@', 'General', 'General']],
  properties: [[{format: {font: {name: 'Arial', bold: true, size: 12}, fill: {color: '#FFF000'},
    wrapText: true, borders: {bottom: {style: 'Continuous', weight: 'Medium', color: '#000000'}}}}, {}, {}], [{}, {}, {}]],
  rowHeights: [24, 30], columnWidths: [60, 90, 120]};
const plan = {id: 'synthetic-plan', title: 'Synthetic', summary: 'Synthetic', assumptions: [], warnings: [],
  actions: [{type: 'writeFormulas', sheet: 'Synthetic', range: 'D6', formulas: [['=IF(C6,1,0)']]},
    {type: 'writeValues', sheet: 'Synthetic', range: 'F8', values: [['示例<&"']]}]};
const before = JSON.stringify(source);
const artifact = module.exports.buildManualReviewFile('EB_Synthetic', source, plan, 20000, 8388608);
assert.equal(JSON.stringify(source), before);
assert.equal(artifact.markedCells, 2);
const check = spawnSync(process.argv[2] || path.join(root, '.venv/Scripts/python.exe'), ['-c', `
import base64, io, sys, zipfile
from openpyxl import load_workbook
raw = base64.b64decode(sys.stdin.read())
book = load_workbook(io.BytesIO(raw), data_only=False)
assert book.sheetnames == ['EB_Synthetic']
sheet = book.active
assert sheet['B5'].value == 'LOCAL_PRIVATE_SENTINEL'
assert sheet['B6'].value == '=LITERAL' and sheet['B6'].data_type == 's'
assert sheet['C6'].value is True
assert sheet['D5'].value == '=C5+1' and sheet['D5'].data_type == 'f'
assert sheet['D6'].value == '=IF(C6,1,0)' and sheet['D6'].data_type == 'f'
assert sheet['F8'].value == '示例<&"'
assert sheet['B5'].font.name == 'Arial' and sheet['B5'].font.bold
assert sheet['B5'].alignment.wrap_text
assert sheet['B5'].border.bottom.style == 'medium'
assert sheet['B5'].fill.fgColor.rgb == 'FFFFF000'
assert sheet['D6'].fill.fgColor.rgb == 'FFE3F1E8'
assert sheet['C5'].number_format == '0.00'
assert sheet.row_dimensions[5].height == 24
assert sheet.column_dimensions['C'].width > 0
assert book.calculation.fullCalcOnLoad and book.calculation.forceFullCalc
assert load_workbook(io.BytesIO(raw), data_only=True).active['D6'].value is None
assert not any('external' in n.lower() or 'vba' in n.lower() for n in zipfile.ZipFile(io.BytesIO(raw)).namelist())
print('Independent OOXML reader: formulas, literal text, values, formats, markers and recalculation passed. No real workbook accessed.')
`], {input: artifact.base64, encoding: 'utf8', cwd: root});
assert.equal(check.status, 0, check.stderr || check.stdout || check.error?.message);
process.stdout.write(check.stdout);
