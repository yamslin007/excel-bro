import type { AnalysisPlan, CellValue, ExcelAction, VerificationCriterion } from './contracts';
import { copyRangePlan } from '@excel-bro/core/copyRangePlan';
import { parseExplicitA1Rect } from './planCore';
import { nativeColumnName } from './nativeReviewFile';
import { captureWorkbookSourceFingerprint, currentOfficeWorkbookIdentity } from './excel';

/** Geometry only: never execute this projection or use it as a write receipt. */
export function valueCopyReviewGeometry(plan: AnalysisPlan, authorization: readonly string[]): AnalysisPlan {
  const targets = new Set(plan.actions.map(action => action.sheet));
  return { ...plan, actions: plan.actions.map(action => {
    if (action.type !== 'copyRange') return action;
    if (!authorization.includes(action.sourceSheet) || !authorization.includes(action.sheet))
      throw new Error('复制来源或目标超出本次授权，未执行。');
    if (action.copyType !== 'values' || action.skipBlanks || action.filters?.length)
      throw new Error('该复制方式尚未接入自动结果审阅，未执行。');
    if (targets.has(action.sourceSheet)) throw new Error('复制来源与写入目标必须分离，未执行。');
    const source = parseExplicitA1Rect(action.sourceRange), target = parseExplicitA1Rect(action.targetRange);
    if (!source || !target) throw new Error('复制来源和目标必须是明确的 A1 区域。');
    const rows = action.transpose ? source.columns : source.rows;
    const columns = action.transpose ? source.rows : source.columns;
    if ((target.rows !== 1 || target.columns !== 1) && (target.rows !== rows || target.columns !== columns))
      throw new Error('复制目标尺寸与来源不符。');
    const endRow = target.startRow + rows - 1, endColumn = target.startColumn + columns - 1;
    if (endRow > 1048576 || endColumn > 16384) throw new Error('复制目标超出工作表边界。');
    return { type: 'clearRange', sheet: action.sheet, applyTo: 'contents',
      range: `${nativeColumnName(target.startColumn)}${target.startRow}:${nativeColumnName(endColumn)}${endRow}` };
  }) };
}

/** Read authorized source cells locally and lower using the same deterministic compiler as WPS. */
export async function prepareExcelValueCopies(plan: AnalysisPlan, authorization: readonly string[],
  identity: string, isCurrent: () => boolean): Promise<AnalysisPlan> {
  if (!plan.actions.some(action => action.type === 'copyRange')) return plan;
  valueCopyReviewGeometry(plan, authorization);
  const check = () => {
    if (!isCurrent() || currentOfficeWorkbookIdentity() !== identity) throw new Error('复制计划、授权或工作簿已变化，未执行。');
  };
  check();
  // Retain the planning baseline check; do not bless stale values by replacing it.
  if (plan.sourceFingerprint) {
    const old = await captureWorkbookSourceFingerprint([...(plan.sourceFingerprintSheets ?? authorization)]);
    check();
    if (old !== plan.sourceFingerprint) throw new Error('复制来源自规划后发生变化，请重新读取；未执行。');
  }
  const before = await captureWorkbookSourceFingerprint([...authorization]);
  check();
  const criteria: VerificationCriterion[] = [...(plan.acceptanceCriteria ?? [])];
  const sources = new Set(plan.sourceDependencySheets ?? []);
  for (const action of plan.actions) if (action.type === 'copyRange') sources.add(action.sourceSheet);
  // A recovery baseline must track inputs, not targets changed by our own earlier actions.
  // Keep the broad planning/target checks above and below; do not refresh a stale baseline on resume.
  if ([...sources].some(sheet => !authorization.includes(sheet))) throw new Error('复制来源超出本次授权，未执行。');
  const sourceBefore = await captureWorkbookSourceFingerprint([...sources]);
  check();
  const actions = await Excel.run(async context => {
    const output: ExcelAction[] = [];
    for (const action of plan.actions) {
      if (action.type !== 'copyRange') { output.push(action); continue; }
      check();
      const range = context.workbook.worksheets.getItem(action.sourceSheet).getRange(action.sourceRange);
      range.load('values,formulas'); await context.sync(); check();
      if (plan.actions.length > 1 && (range.formulas as CellValue[][]).some(row => row.some(value =>
        typeof value === 'string' && value.startsWith('='))))
        throw new Error('多动作值复制的来源含公式依赖，无法保证预读与逐步复制等价；未写入。');
      const writes = copyRangePlan(action, range.values as CellValue[][]);
      output.push(...writes);
      for (const write of writes) {
        if (write.type !== 'writeValues') throw new Error('值复制产生了非值写入动作。');
        criteria.push({type:'rangeEquals',sheet:write.sheet,range:write.range,expected:write.values});
      }
      sources.add(action.sourceSheet);
    }
    return output;
  });
  const sourceAfter = await captureWorkbookSourceFingerprint([...sources]);
  check();
  const after = await captureWorkbookSourceFingerprint([...authorization]);
  check();
  if (before !== after || sourceBefore !== sourceAfter) throw new Error('复制读取期间来源或目标发生变化，未执行。');
  return { ...plan, actions, acceptanceCriteria: criteria, sourceFingerprint: sourceAfter, sourceFingerprintSheets: [...sources],
    sourceDependencySheets: [...sources] };
}
