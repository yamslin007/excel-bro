import type { AnalysisPlan } from './contracts';
import { currentOfficeWorkbookIdentity } from './excel';
import { reviewMarkStyleForPlan, REVIEW_MARK_COLORS } from '@excel-bro/core/reviewMarkStyle';

export type ExcelResultMarks = { identity: string; formula: string; entries: Array<{ sheet: string; range: string; id: string }> };

/**
 * 标记归属的**持久记录**。
 *
 * 刷新/重载之后内存里的 `marks` 就没了，但工作簿里的条件格式规则还在。没有这份
 * 记录，那些规则就成了"无主标记"：用户只能看到一圈边框，加载项既认不出也不敢清。
 *
 * **只存归属**：工作表名、区域、规则 id、以及本次随机公式（真正的归属凭据）。
 * 不存单元格值、公式或任何业务内容 —— 这份记录会进 localStorage，属于本机可读数据。
 */
const STORAGE_KEY = 'excel-bro.reviewMarks.v1';
type StoredMarks = { v: 1; identity: string; formula: string; entries: ExcelResultMarks['entries'] };

function readStored(): StoredMarks | null {
  try {
    const raw = globalThis.localStorage?.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as StoredMarks;
    if (!parsed || parsed.v !== 1 || typeof parsed.identity !== 'string' || typeof parsed.formula !== 'string' ||
        !Array.isArray(parsed.entries)) return null;
    return parsed;
  } catch { return null; }
}

function writeStored(marks: ExcelResultMarks): void {
  try {
    const record: StoredMarks = { v: 1, identity: marks.identity, formula: marks.formula, entries: marks.entries };
    globalThis.localStorage?.setItem(STORAGE_KEY, JSON.stringify(record));
  } catch { /* 存不下不影响本次标记本身；只是重载后需要用户自己看图清理。 */ }
}

/** 清掉持久记录（标记已确认移除，或确认不属于本机可处理的范围）。 */
export function forgetExcelResultMarks(): void {
  try { globalThis.localStorage?.removeItem(STORAGE_KEY); } catch { /* ignore */ }
}

/** 读取遗留归属记录（不触碰工作簿）。 */
export function loadExcelResultMarks(): ExcelResultMarks | null {
  const stored = readStored();
  return stored ? { identity: stored.identity, formula: stored.formula, entries: stored.entries } : null;
}

/** A separate conditional-format rule, never an overwrite of the user's direct fill. */
export async function markExcelResult(plan: AnalysisPlan, identity: string): Promise<ExcelResultMarks> {
  if (!identity || currentOfficeWorkbookIdentity() !== identity) throw new Error('标记目标工作簿已变化。');
  const formula = `=N("EB_REVIEW_${crypto.randomUUID().replace(/-/g, '')}")=0`;
  const marks: ExcelResultMarks = { identity, formula, entries: [] };
  // 按整份计划的重叠关系一次算清每个区域该用哪种标记（不是逐个动作判）。
  // 只统计"落在单元格区域上"的动作；无 range 的动作（如建表）不参与标记。
  const styleByRange = reviewMarkStyleForPlan(
    plan.actions.filter((action): action is typeof action & { sheet: string; range: string } =>
      'range' in action && typeof action.range === 'string' && typeof action.sheet === 'string'
    )
  );
  try {
    await Excel.run(async context => {
      const seen = new Set<string>();
      for (const action of plan.actions) {
        if (!('range' in action) || !action.range) continue;
        const key = JSON.stringify([action.sheet, action.range]);
        if (seen.has(key)) continue;
        seen.add(key);
        if (currentOfficeWorkbookIdentity() !== identity) throw new Error('标记目标工作簿已变化。');
        const style = styleByRange.get(`${action.sheet}\u0000${action.range}`) ?? 'fill+border';
        if (style === 'none') continue;
        const range = context.workbook.worksheets.getItem(action.sheet).getRange(action.range);
        const rule = range.conditionalFormats.add('Custom');
        rule.custom.rule.formula = formula;

        /*
         * 醒目边框：两种样式都要。原来的绿色细边框在实机里像普通网格，用户看不见；
         * 换金黄色、加粗线型，与网格线区分开。条件格式边框只有 color/style 两态，
         * 没有 weight —— 醒目靠颜色 + `Continuous`，不靠不存在的属性。
         */
        if (style !== 'fill-only') for (const edge of ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'] as const) {
          const border = rule.custom.format.borders.getItem(edge);
          border.color = REVIEW_MARK_COLORS.border;
          border.style = 'Continuous';
        }
        /*
         * 浅黄临时底色：只给**不修改填充**的区域。
         *
         * 底色写在**独立条件格式规则**的 `format.fill` 上，不是单元格 `format.fill` ——
         * 删除该规则时底色随标记一起消失，不覆盖业务填充、也不需要"先存旧底色再还原"。
         * 涉及填充变化的区域（setFill）只画边框，保证实际填充结果可见。
         */
        if (style === 'fill+border' || style === 'fill-only') {
          rule.custom.format.fill.color = REVIEW_MARK_COLORS.fill;
        }
        rule.load('id'); await context.sync();
        marks.entries.push({ sheet: action.sheet, range: action.range, id: rule.id });
        /*
         * **逐条落盘**：创建到第 N 条时失败，前 N-1 条必须留下可追踪的归属。
         * 只在整个循环结束后才写，部分成功就会变成没有任何记录的无主标记。
         */
        writeStored(marks);
      }
    });
    return marks;
  } catch (reason) {
    // Remove successfully registered marks, not other conditional-format rules.
    if (marks.entries.length) {
      try {
        await clearExcelResultMarks(marks);
      } catch {
        /*
         * 清理也失败：**保留**持久记录，让下一次打开面板/开始新审阅时还能认领它们。
         * 这里不能"顺手把记录删掉"——那才是真正把标记变成无主。
         */
        throw new Error(`${reason instanceof Error ? reason.message : '标记创建失败'}；已创建的 ${marks.entries.length} 处标记暂未移除，可在下次打开时重试清理。`);
      }
    }
    throw reason;
  }
}

export async function clearExcelResultMarks(marks: ExcelResultMarks): Promise<void> {
  if (currentOfficeWorkbookIdentity() !== marks.identity) throw new Error('请返回本次执行的工作簿，再清除审阅标记。');
  await Excel.run(async context => {
    const removals = new Map<string, Excel.ConditionalFormat>();
    const inRecordedRange = async (entry: ExcelResultMarks['entries'][number]) => {
      const rules = context.workbook.worksheets.getItem(entry.sheet).getRange(entry.range).conditionalFormats;
      rules.load('items/id,items/type'); await context.sync();
      for (const rule of rules.items) if (rule.type === 'Custom') rule.custom.load('rule');
      await context.sync();
      return rules.items;
    };
    for (const entry of marks.entries) {
      const rules = await inRecordedRange(entry);
      let rule = rules.find(item => item.id === entry.id);
      if (!rule) {
        // A stale ID is not evidence that the mark disappeared. Claim only a unique,
        // exact per-review nonce in its recorded range, never by colour or prefix.
        if (!/^=N\("EB_REVIEW_[a-f0-9]{32}"\)=0$/.test(marks.formula))
          throw new Error('审阅标记归属公式无效，未删除任何规则。');
        const matches = rules.filter(item => item.type === 'Custom' && item.custom.rule.formula === marks.formula);
        if (matches.length > 1) throw new Error('审阅标记归属不唯一，保留规则等待核对。');
        rule = matches[0];
        if (!rule) continue;
      }
      if (rule.type !== 'Custom' || rule.custom.rule.formula !== marks.formula)
        throw new Error('审阅标记已被修改，未删除该规则。');
      removals.set(`${entry.sheet}\u0000${rule.id}`, rule);
    }
    if (currentOfficeWorkbookIdentity() !== marks.identity) throw new Error('标记目标工作簿已变化。');
    removals.forEach(rule => rule.delete());
    await context.sync();
    for (const entry of marks.entries) {
      if (currentOfficeWorkbookIdentity() !== marks.identity) throw new Error('标记目标工作簿已变化。');
      const remaining = await inRecordedRange(entry);
      if (remaining.some(rule => rule.type === 'Custom' && rule.custom.rule.formula === marks.formula))
        throw new Error('审阅颜色标记尚未移除，可重试清理；不会再次恢复单元格。');
    }
  });
  // A concurrent newer review must not lose its ownership record.
  if (loadExcelResultMarks()?.formula === marks.formula) forgetExcelResultMarks();
}

/**
 * 刷新/重载后处理**上一次**遗留的标记。
 *
 * 三条纪律：
 *  1. 优先按记录的规则 id + 完整随机公式认领；id 缺失时，仅接受原记录区域内
 *     唯一的完整随机公式匹配。绝不按颜色/前缀批量删，归属歧义一律保留。
 *  2. 工作簿身份不符 ⇒ **什么都不做**，只如实说明"是另一本工作簿的标记"。
 *  3. 清理失败 ⇒ 保留记录，返回可读原因；下一次仍可重试。
 *
 * 返回 `null` 表示没有遗留、或已清理干净；返回字符串表示需要让用户知道的情况。
 */
export async function recoverExcelResultMarks(): Promise<string | null> {
  const marks = loadExcelResultMarks();
  if (!marks) return null;
  if (currentOfficeWorkbookIdentity() !== marks.identity) {
    // 不碰别的簿，也不删记录：等用户回到那本工作簿时还能认领。
    return '上次审阅留下的标记属于另一本工作簿，未做任何改动；请回到那本工作簿后再清理。';
  }
  try {
    await clearExcelResultMarks(marks);
    return '已清理上次审阅遗留的标记。';
  } catch (reason) {
    return `上次审阅遗留的标记未清理：${reason instanceof Error ? reason.message : '原因未知'}；可稍后重试。`;
  }
}
