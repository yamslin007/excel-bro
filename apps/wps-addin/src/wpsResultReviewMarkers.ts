import type { WpsApplication, WpsWorkbook } from './wpsHost';
import type { WpsPreparedPlan } from './wpsPlanExecutor';

type Rule = { Formula1?: unknown; Borders?: { Item(index: number): { Color?: number; LineStyle?: number } }; Delete?(): void };
type Conditions = { Add(type: number, operator?: unknown, formula?: string): Rule };
export type WpsResultMarks = { workbook: WpsWorkbook; workbookFullName: string; formula: string; rules: Rule[] };

export function markWpsResult(application: WpsApplication, plan: WpsPreparedPlan): WpsResultMarks {
  const workbook = application.ActiveWorkbook;
  if (!workbook || String(workbook.FullName ?? '') !== plan.authorization.workbookFullName) throw new Error('标记目标工作簿已变化。');
  const formula = `=N("EB_REVIEW_${crypto.randomUUID().replace(/-/g, '')}")=0`;
  const marks: WpsResultMarks = { workbook, workbookFullName: plan.authorization.workbookFullName, formula, rules: [] };
  const ranges = new Map<string, Conditions>();
  // Check availability for every target before adding the first rule.
  for (const action of plan.actions) {
    const range = workbook.Worksheets.Item(action.sheetName).Range?.(action.range) as unknown as { FormatConditions?: Conditions };
    if (typeof range?.FormatConditions?.Add !== 'function') throw new Error('当前 WPS 不提供独立审阅标记接口，业务格式未被覆盖。');
    ranges.set(JSON.stringify([action.sheetName, action.range]), range.FormatConditions);
  }
  try {
    for (const conditions of ranges.values()) {
      const rule = conditions.Add(2, undefined, formula); // xlExpression
      marks.rules.push(rule);
      if (rule.Formula1 !== formula || typeof rule.Delete !== 'function' || !rule.Borders?.Item) throw new Error('WPS 未能建立可独立移除的审阅标记。');
      for (const edge of [7, 8, 9, 10]) { // xlEdgeLeft/Top/Bottom/Right
        const border = rule.Borders.Item(edge);
        border.Color = 0x50762F; border.LineStyle = 1;
      }
    }
    return marks;
  } catch (reason) {
    clearWpsResultMarks(application, marks);
    throw reason;
  }
}
export function clearWpsResultMarks(application: WpsApplication, marks: WpsResultMarks): void {
  if (!marks.workbookFullName || String(application.ActiveWorkbook?.FullName ?? '') !== marks.workbookFullName ||
      String(marks.workbook.FullName ?? '') !== marks.workbookFullName) throw new Error('请返回本次执行的工作簿，再移除标记。');
  // Validate the entire set first. Never call Range.FormatConditions.Delete().
  for (const rule of marks.rules) {
    if (rule.Formula1 !== marks.formula || typeof rule.Delete !== 'function') throw new Error('审阅标记已被修改，未删除该规则。');
  }
  for (const rule of marks.rules) rule.Delete!();
}
