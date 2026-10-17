import type { AnalysisPlan } from './contracts';
import { currentOfficeWorkbookIdentity } from './excel';

export type ExcelResultMarks = { identity: string; formula: string; entries: Array<{ sheet: string; range: string; id: string }> };

/** A separate conditional-format rule, never an overwrite of the user's direct fill. */
export async function markExcelResult(plan: AnalysisPlan, identity: string): Promise<ExcelResultMarks> {
  if (!identity || currentOfficeWorkbookIdentity() !== identity) throw new Error('标记目标工作簿已变化。');
  const formula = `=N("EB_REVIEW_${crypto.randomUUID().replace(/-/g, '')}")=0`;
  const marks: ExcelResultMarks = { identity, formula, entries: [] };
  try {
    await Excel.run(async context => {
      const seen = new Set<string>();
      for (const action of plan.actions) {
        if (!('range' in action) || !action.range) continue;
        const key = JSON.stringify([action.sheet, action.range]);
        if (seen.has(key)) continue;
        seen.add(key);
        if (currentOfficeWorkbookIdentity() !== identity) throw new Error('标记目标工作簿已变化。');
        const range = context.workbook.worksheets.getItem(action.sheet).getRange(action.range);
        const rule = range.conditionalFormats.add('Custom');
        rule.custom.rule.formula = formula;
        // Border-only marks let the actual fill/font remain visible during review.
        for (const edge of ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'] as const) {
          const border = rule.custom.format.borders.getItem(edge);
          border.color = '#2F7650'; border.style = 'Continuous';
        }
        rule.load('id'); await context.sync();
        marks.entries.push({ sheet: action.sheet, range: action.range, id: rule.id });
      }
    });
    return marks;
  } catch (reason) {
    // Remove successfully registered marks, not other conditional-format rules.
    if (marks.entries.length) await clearExcelResultMarks(marks);
    throw reason;
  }
}

export async function clearExcelResultMarks(marks: ExcelResultMarks): Promise<void> {
  if (currentOfficeWorkbookIdentity() !== marks.identity) throw new Error('请返回本次执行的工作簿，再清除审阅标记。');
  await Excel.run(async context => {
    const removals: Excel.ConditionalFormat[] = [];
    for (const entry of marks.entries) {
      const rules = context.workbook.worksheets.getItem(entry.sheet).getRange(entry.range).conditionalFormats;
      rules.load('items/id,items/type'); await context.sync();
      const rule = rules.items.find(item => item.id === entry.id);
      if (!rule) continue; // The user already removed this mark.
      if (rule.type !== 'Custom') throw new Error('审阅标记已被修改，未删除该规则。');
      rule.custom.load('rule'); await context.sync();
      if (rule.custom.rule.formula !== marks.formula) throw new Error('审阅标记已被修改，未删除该规则。');
      removals.push(rule);
    }
    if (currentOfficeWorkbookIdentity() !== marks.identity) throw new Error('标记目标工作簿已变化。');
    removals.forEach(rule => rule.delete());
    await context.sync();
  });
}
