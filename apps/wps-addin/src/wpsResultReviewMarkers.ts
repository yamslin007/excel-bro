import type { WpsApplication, WpsRange, WpsWorkbook } from './wpsHost';
import type { WpsPreparedPlan } from './wpsPlanExecutor';
import { wpsReadWorkbook } from './wpsResultReviewTarget';
import { reviewMarkStyleForPlan, REVIEW_MARK_COLORS } from '@excel-bro/core/reviewMarkStyle';

type BorderEdge = { Color?: number; LineStyle?: number; Weight?: unknown };
type Rule = {
  Formula1?: unknown;
  Borders?: { Item(index: number): BorderEdge };
  /** WPS 条件格式规则自身的 Interior（独立于单元格 `Range.Interior`）。可选：旧宿主可能没有。 */
  Interior?: { Color?: number; Pattern?: unknown; PatternColor?: unknown };
  Delete?(): void;
};
type Conditions = { Add(type: number, operator?: unknown, formula?: string): Rule };
export type WpsResultMarks = { workbook: WpsWorkbook; workbookFullName: string; formula: string; rules: Rule[] };

/**
 * 标记归属的**持久记录**（刷新/重载后用来认领遗留规则）。
 *
 * WPS 的 `FormatCondition` 没有可读 id，唯一可靠的归属凭据是本次随机生成的
 * `Formula1`（`=N("EB_REVIEW_<nonce>")=0`）。所以记录里存的是
 * **表名 + 区域 + 这个公式**，不存任何单元格内容。
 */
const STORAGE_KEY = 'excel-bro.wps.reviewMarks.v1';
type StoredWpsMarks = { v: 1; workbookFullName: string; formula: string; ranges: Array<{ sheet: string; range: string }> };

function readStored(): StoredWpsMarks | null {
  try {
    const raw = globalThis.localStorage?.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as StoredWpsMarks;
    if (!parsed || parsed.v !== 1 || typeof parsed.workbookFullName !== 'string' ||
        typeof parsed.formula !== 'string' || !Array.isArray(parsed.ranges)) return null;
    return parsed;
  } catch { return null; }
}

function writeStored(marks: WpsResultMarks, ranges: Array<{ sheet: string; range: string }>): void {
  try {
    const record: StoredWpsMarks = { v: 1, workbookFullName: marks.workbookFullName, formula: marks.formula, ranges };
    globalThis.localStorage?.setItem(STORAGE_KEY, JSON.stringify(record));
  } catch { /* 存不下不影响本次标记；只是重载后需要用户自己看图清理。 */ }
}

export function forgetWpsResultMarks(): void {
  try { globalThis.localStorage?.removeItem(STORAGE_KEY); } catch { /* ignore */ }
}

/**
 * `#RRGGBB` → WPS `Interior.Color` / `Border.Color` 用的 BGR 整型。
 * 失败返回 0（黑色）而不抛：颜色给错不能把整份标记打失败，宁可退回可见的边框。
 */
function parseColor(hex: string): number {
  const match = /^#?([0-9a-fA-F]{6})$/.exec(hex.trim());
  if (!match) return 0;
  const value = Number.parseInt(match[1], 16);
  const r = (value >> 16) & 0xff;
  const g = (value >> 8) & 0xff;
  const b = value & 0xff;
  return (b << 16) | (g << 8) | r;
}

/**
 * 从一张**副本**工作表上清掉被 `Copy()` 带过来的**本次**审阅标记。
 *
 * 为什么需要：`stageWpsBeforeComparison` 用原生 `Worksheet.Copy()` 复制当前原表，
 * 而原表此刻已经打上了审阅标记（条件格式规则）。不清理的话，执行前副本也会
 * 高亮同样的区域，用户看到两边都黄，误以为"执行前也是这个颜色"，比对失去意义。
 *
 * 归属判据是**本次实际创建标记的完整随机公式**（`formula`），不是 `EB_REVIEW_`
 * 前缀：另一轮审阅/另一任务的标记即便同前缀（nonce 不同）也**一律不碰**。
 * 用户自己的条件格式规则（及其顺序、语义）更是一个不碰。不写持久记录、
 * 也不认领原表 —— 纯副本侧清理。
 */
export function clearReviewMarksFromCopySheet(
  sheet: { Range?(address: string): WpsRange },
  addresses: readonly string[],
  formula: string
): void {
  for (const address of [...new Set(addresses)]) {
    try {
      const range = sheet.Range?.(address) as unknown as {
        FormatConditions?: { Count?: number; Item?(index: number): Rule };
      };
      const conditions = range?.FormatConditions;
      if (!conditions || typeof conditions.Item !== 'function') continue;
      const count = Number(conditions.Count ?? 0);
      for (let index = count; index >= 1; index -= 1) {
        const rule = conditions.Item(index);
        // 归属判据只有一个：本次完整随机公式。前缀/颜色/区域都不构成归属。
        if (rule?.Formula1 === formula && typeof rule?.Delete === 'function') rule.Delete();
      }
    } catch { /* 清不动就保留：副本多一圈高亮，好过误删用户规则。 */ }
  }
}

export function markWpsResult(application: WpsApplication, plan: WpsPreparedPlan): WpsResultMarks {
  // 标记永远打在**原表**上：副本获得焦点不代表目标是副本。
  const workbook = wpsReadWorkbook(application);
  if (!workbook || String(workbook.FullName ?? '') !== plan.authorization.workbookFullName) throw new Error('标记目标工作簿已变化。');
  const formula = `=N("EB_REVIEW_${crypto.randomUUID().replace(/-/g, '')}")=0`;
  const marks: WpsResultMarks = { workbook, workbookFullName: plan.authorization.workbookFullName, formula, rules: [] };
  // 按整份计划的重叠关系，一次算清每个区域该用哪种标记（不是逐个动作判）。
  // WPS 侧动作字段叫 `sheetName`，映射成共享判据要的 `sheet`（两端读同一份函数）。
  const styleByRange = reviewMarkStyleForPlan(
    plan.actions.map(action => ({ type: action.type, sheet: action.sheetName, range: action.range }))
  );
  const ranges = new Map<string, Conditions>();
  // Check availability for every target before adding the first rule.
  for (const action of plan.actions) {
    if (styleByRange.get(`${action.sheetName}\u0000${action.range}`) === 'none') continue;
    const range = workbook.Worksheets.Item(action.sheetName).Range?.(action.range) as unknown as { FormatConditions?: Conditions };
    if (typeof range?.FormatConditions?.Add !== 'function') throw new Error('当前 WPS 不提供独立审阅标记接口，业务格式未被覆盖。');
    ranges.set(JSON.stringify([action.sheetName, action.range]), range.FormatConditions);
  }
  const tracked: Array<{ sheet: string; range: string }> = [];
  try {
    for (const [key, conditions] of ranges.entries()) {
      const rule = conditions.Add(2, undefined, formula); // xlExpression
      marks.rules.push(rule);
      const [sheet, range] = JSON.parse(key) as [string, string];
      tracked.push({ sheet, range });
      // 逐条落盘：创建到一半失败时，已创建的部分仍有归属记录可追。
      writeStored(marks, tracked);
      if (rule.Formula1 !== formula || typeof rule.Delete !== 'function') throw new Error('WPS 未能建立可独立移除的审阅标记。');
      const style = styleByRange.get(`${sheet}\u0000${range}`) ?? 'fill+border';
      /*
       * 醒目边框：两种样式都要（绿色细边框在实机里像普通网格，用户根本看不见）。
       * 用金黄色 + 加粗，和"普通网格线"区分开。
       */
      if (style !== 'fill-only' && !rule.Borders?.Item) throw new Error('WPS 缺少审阅边框接口。');
      if (style !== 'fill-only') for (const edge of [7, 8, 9, 10]) { // xlEdgeLeft/Top/Bottom/Right
        const border = rule.Borders!.Item(edge);
        border.Color = parseColor(REVIEW_MARK_COLORS.border);
        border.LineStyle = 1; // xlContinuous
        border.Weight = 2;    // xlThin
      }
      /*
       * 浅黄临时底色：只给**不修改填充**的区域。
       *
       * 底色写在条件格式规则自己的 `Interior` 上，**不是**单元格 `Range.Interior` ——
       * 删除该规则时底色随标记一起消失，不覆盖业务填充、也不需要"先存旧底色再还原"。
       * 涉及填充变化的区域（setFill）只画边框，保证实际填充结果可见。
       */
      if (style === 'fill+border' || style === 'fill-only') {
        if (!rule.Interior) throw new Error('WPS 缺少审阅底色接口。');
        rule.Interior.Pattern = 1; // xlSolid
        rule.Interior.Color = parseColor(REVIEW_MARK_COLORS.fill);
      }
    }
    return marks;
  } catch (reason) {
    try {
      clearWpsResultMarks(application, marks);
      forgetWpsResultMarks();
    } catch {
      // 清理失败：保留归属记录，下次打开仍能认领这些规则。
      throw new Error(`${reason instanceof Error ? reason.message : '标记创建失败'}；已创建的 ${marks.rules.length} 处标记暂未移除，可在下次打开时重试清理。`);
    }
    throw reason;
  }
}

export function clearWpsResultMarks(application: WpsApplication, marks: WpsResultMarks): void {
  /*
   * 清理同样按原表：双视图会把副本激活成 `ActiveWorkbook`，若这里读活动簿，
   * 确认/放弃就会报"请返回本次执行的工作簿" —— 那正是实机里"清标记失败"的来源。
   */
  if (!marks.workbookFullName || String(wpsReadWorkbook(application)?.FullName ?? '') !== marks.workbookFullName ||
      String(marks.workbook.FullName ?? '') !== marks.workbookFullName) throw new Error('请返回本次执行的工作簿，再移除标记。');
  // Validate the entire set first. Never call Range.FormatConditions.Delete().
  for (const rule of marks.rules) {
    if (rule.Formula1 !== marks.formula || typeof rule.Delete !== 'function') throw new Error('审阅标记已被修改，未删除该规则。');
  }
  for (const rule of marks.rules) rule.Delete!();
  forgetWpsResultMarks();
}

/**
 * 刷新/重载后处理上一次遗留的标记。
 *
 * 只按记录的 `Formula1`（本次随机 nonce）认领：
 *  - 同一区域里 Formula1 不等于它的规则**一律不碰**（别的任务 / 用户自建）；
 *  - 工作簿身份不符 ⇒ 什么都不做，只说明情况；
 *  - 删除失败的规则保留在记录里，下次再试，绝不改成"按区域清空规则"。
 */
export function recoverWpsResultMarks(application: WpsApplication): string | null {
  const stored = readStored();
  if (!stored) return null;
  const active = wpsReadWorkbook(application);
  if (!active || String(active.FullName ?? '') !== stored.workbookFullName) {
    return '上次审阅留下的标记属于另一本工作簿，未做任何改动；请回到那本工作簿后再清理。';
  }
  let removed = 0;
  const remaining: Array<{ sheet: string; range: string }> = [];
  for (const entry of stored.ranges) {
    try {
      const holder = active.Worksheets.Item(entry.sheet).Range?.(entry.range) as unknown as {
        FormatConditions?: { Count?: number; Item?(index: number): Rule };
      };
      const conditions = holder?.FormatConditions;
      if (!conditions || typeof conditions.Item !== 'function') { remaining.push(entry); continue; }
      const count = Number(conditions.Count ?? 1);
      const matches: Rule[] = [];
      for (let index = 1; index <= count; index += 1) {
        const rule = conditions.Item(index);
        // 归属判据只有一个：本次随机公式。颜色 / 区域 / 前缀都不构成归属。
        if (rule?.Formula1 === stored.formula) matches.push(rule);
      }
      for (const rule of matches) {
        if (typeof rule.Delete !== 'function') { remaining.push(entry); continue; }
        rule.Delete();
        removed += 1;
      }
    } catch { remaining.push(entry); }
  }
  if (remaining.length === 0) {
    forgetWpsResultMarks();
    return removed > 0 ? '已清理上次审阅遗留的标记。' : '上次审阅的标记已不存在，记录已清除。';
  }
  try {
    globalThis.localStorage?.setItem(STORAGE_KEY, JSON.stringify({ ...stored, ranges: remaining }));
  } catch { /* ignore */ }
  return `上次审阅遗留的 ${remaining.length} 处标记未清理；可稍后重试。`;
}