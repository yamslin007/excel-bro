import type { WpsApplication, WpsRange, WpsWorkbook } from './wpsHost';
import type { WpsPreparedPlan } from './wpsPlanExecutor';
import { wpsReadWorkbook } from './wpsResultReviewTarget';
import { reviewMarkStyleForPlan, REVIEW_MARK_COLORS } from '@excel-bro/core/reviewMarkStyle';

type BorderEdge = { Color?: number; LineStyle?: number; Weight?: unknown };
type Rule = {
  Formula1?: unknown;
  /** 规则类型（`XlFormatConditionType`）。宿主可能不给，读不出来时是 `undefined`。 */
  Type?: unknown;
  Borders?: { Item(index: number): BorderEdge };
  /** WPS 条件格式规则自身的 Interior（独立于单元格 `Range.Interior`）。可选：旧宿主可能没有。 */
  Interior?: { Color?: number; Pattern?: unknown; PatternColor?: unknown };
  Delete?(): void;
};
type RuleCollection = { Count?: number; Item?(index: number): Rule | null | undefined };
type Conditions = RuleCollection & { Add(type: number, operator?: unknown, formula?: string): Rule | null | undefined };

/**
 * `XlFormatConditionType` 里**结构上就没有 `Formula1`** 的那些类型：
 * 色阶 3、数据条 4、前 10 项 5、图标集 6、唯一值 7、空值 9、时间段 10、
 * 高于/低于平均值 11、非空值 12、错误值 13、无错误值 14。
 *
 * **唯一一份白名单**，两个消费者读同一份判据：归属性扫描（`ownedRules`：Add 回空后
 * 找回本次规则、重载清理遗留标记）与副本清理（`clearOneCopyRange`）。
 *
 * 本次标记永远是 `xlExpression`(2) + 已知随机串，所以这些类型**一定不是本次的**，
 * 可以安全跳过 —— 否则用户只要放过一个色阶，每次开双视图/每次重开窗格都会收到
 * 一句"无法核实/未清理"的假警报。反过来，类型读不出来或是 1/2/8（带公式的类型）、
 * 或是不认识的值时，一律不许替宿主"认定"它是色阶，只能算"无法核实"。
 */
const NON_EXPRESSION_RULE_TYPES: ReadonlySet<number> = new Set([3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14]);

function isKnownNonExpressionRule(rule: Rule): boolean {
  const type = rule.Type;
  // 文本型的 Type 不能 `Number()` 一下就当枚举用：那不是"宿主明确返回了类型"。
  return typeof type === 'number' && NON_EXPRESSION_RULE_TYPES.has(type);
}

/**
 * 在集合里找**本次**规则（只认完整随机公式）。
 *
 * 返回值是空数组也算读数成功（本次规则确实不在）；**读不出来一律抛错**，由调用方
 * 保留归属记录，绝不能用"没找到"冒充"已经清掉"：
 *  - 集合/条数读不出、句柄是空的：这一下没法判断，抛；
 *  - 规则读不出公式：只有宿主**明确**给出可识别的非表达式类型（色阶/数据条…，见上）
 *    才安全跳过；未知类型或表达式规则读不出公式，仍算无法核实，抛。
 */
function ownedRules(conditions: RuleCollection | undefined, formula: string): Rule[] {
  const count = conditions?.Count;
  if (!conditions || typeof conditions.Item !== 'function' || typeof count !== 'number' ||
      !Number.isSafeInteger(count) || count < 0) throw new Error('WPS 审阅标记集合暂不可读取，归属记录已保留。');
  const matches: Rule[] = [];
  for (let index = 1; index <= count; index += 1) {
    const rule = conditions.Item(index);
    // 空句柄不是"这条不存在"：集合已不可信，继续找下去就是拿运气当证据。
    if (!rule) throw new Error('WPS 审阅标记句柄暂不可读取，归属记录已保留。');
    const ruleFormula = rule.Formula1;
    if (ruleFormula == null) {
      if (isKnownNonExpressionRule(rule)) continue;
      throw new Error('WPS 审阅标记句柄暂不可读取，归属记录已保留。');
    }
    if (ruleFormula === formula) matches.push(rule);
  }
  return matches;
}

function deleteOwnedRules(rules: Rule[], formula: string): void {
  // Validate all handles before deleting; never delete the collection itself.
  for (const rule of rules) {
    if (!rule || rule.Formula1 !== formula || typeof rule.Delete !== 'function') {
      throw new Error('审阅标记归属无法确认，未删除该规则。');
    }
  }
  for (const rule of rules) rule.Delete!();
}
export type WpsResultMarks = { workbook: WpsWorkbook; workbookFullName: string; formula: string; rules: Rule[]; ranges?: MarkRange[] };

/**
 * 标记归属的**持久记录**（刷新/重载后用来认领遗留规则）。
 *
 * WPS 的 `FormatCondition` 没有可读 id，唯一可靠的归属凭据是本次随机生成的
 * `Formula1`（`=N("EB_REVIEW_<nonce>")=0`）。所以记录里存的是
 * **簿身份 + 表名 + 区域 + 这个公式**，不存任何单元格内容。
 *
 * 2026-10-25 起记录是**多条**：旧任务没清完就开新审阅时，两条必须并存、各自
 * 清理。旧版（`v: 1`）是单条、字段平铺，读的时候照样按同一条记录认领，
 * 下一次写回时升级成 `v: 2` 的列表 —— 旧记录一条都不丢。
 */
const STORAGE_KEY = 'excel-bro.wps.reviewMarks.v1';
const STORAGE_VERSION = 2;
type MarkRange = { sheet: string; range: string };
type StoredRecord = { workbookFullName: string; formula: string; ranges: MarkRange[] };
type StoredMarks = { v: typeof STORAGE_VERSION; records: readonly StoredRecord[] };

/**
 * 只接受**这四个字段**：簿身份、随机公式、表名、区域。
 *
 * 反过来说，任何业务单元格值（甚至多出来的字段）都不会被读进来，也不会被写出去。
 */
function parseRecord(value: unknown): StoredRecord | null {
  if (!value || typeof value !== 'object') return null;
  const holder = value as { workbookFullName?: unknown; formula?: unknown; ranges?: unknown };
  if (typeof holder.workbookFullName !== 'string' || !holder.workbookFullName ||
      typeof holder.formula !== 'string' || !holder.formula || !Array.isArray(holder.ranges)) return null;
  const ranges: MarkRange[] = [];
  for (const entry of holder.ranges) {
    const item = entry as { sheet?: unknown; range?: unknown } | null;
    if (!item || typeof item.sheet !== 'string' || typeof item.range !== 'string') return null;
    ranges.push({ sheet: item.sheet, range: item.range });
  }
  return { workbookFullName: holder.workbookFullName, formula: holder.formula, ranges };
}

function readStored(): StoredRecord[] {
  try {
    const raw = globalThis.localStorage?.getItem(STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!parsed || typeof parsed !== 'object') return [];
    const holder = parsed as { v?: unknown; records?: unknown };
    if (holder.v === STORAGE_VERSION && Array.isArray(holder.records)) {
      return holder.records.map(parseRecord).filter((record): record is StoredRecord => record !== null);
    }
    // v1：单条、字段平铺。仍然认领，写回时升级成列表。
    if (holder.v === 1) {
      const single = parseRecord(parsed);
      return single ? [single] : [];
    }
    return [];
  } catch { return []; }
}

function writeStored(records: readonly StoredRecord[]): void {
  try {
    if (records.length === 0) { globalThis.localStorage?.removeItem(STORAGE_KEY); return; }
    globalThis.localStorage?.setItem(STORAGE_KEY, JSON.stringify({ v: STORAGE_VERSION, records } satisfies StoredMarks));
  } catch { /* 存不下不影响本次标记；只是重载后需要用户自己看图清理。 */ }
}

/** 只登记/更新**同一个公式**那一条；别的任务、别的簿的记录原样保留。 */
function upsertRecord(record: StoredRecord): void {
  writeStored([...readStored().filter((item) => item.formula !== record.formula), record]);
}

/** 只删**这一个公式**那一条；清不掉的那些由调用方重新写回。 */
function removeRecord(formula: string): void {
  writeStored(readStored().filter((item) => item.formula !== formula));
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
 * 副本侧清理的**结果**。
 *
 * `unconfirmed` 是"我删了但没法确认这张区域已经干净"（集合读不出来、删除被宿主
 * 拒绝）。副本是临时表，这里不重试、也不按区域批量删 —— 只把事实交出去，
 * 由界面如实提示"两张表可能仍然同色"。
 */
export type WpsCopyMarkCleanup = { cleared: number; unconfirmed: number };

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
 *
 * 标记**建到一半失败**时原表会留下本次残留；调用方只要把当时的完整公式传进来
 * （`markWpsResult` 的 `onFormula`），副本照样能精确清掉它。
 *
 * **读不出来不等于没有**（2026-10-25 补）：集合读不出条数、句柄是空的、规则读不出
 * 公式 —— 每一种都可能是"本次残留还在"，一律算进 `unconfirmed`，绝不当成
 * "这个区域本来就没有本次标记"。每个区域最多计一次：一处"没核实干净"按区域算，
 * 不按读不出来的规则条数放大。
 */
export function clearReviewMarksFromCopySheet(
  sheet: { Range?(address: string): WpsRange },
  addresses: readonly string[],
  formula: string
): WpsCopyMarkCleanup {
  const report: WpsCopyMarkCleanup = { cleared: 0, unconfirmed: 0 };
  for (const address of [...new Set(addresses)]) {
    if (!clearOneCopyRange(sheet, address, formula, report)) report.unconfirmed += 1;
  }
  return report;
}

/**
 * 清一个区域。返回 `true` = 集合读过了、该删的都删了（可以放心说这个区域干净）；
 * `false` = 有东西没核实（条数读不出、句柄是空的、公式读不出且类型不认识…）。
 *
 * 类型判据用的是**同一份** `isKnownNonExpressionRule`（定义在 `ownedRules` 上方），
 * 副本清理与归属性扫描不各维一份白名单 —— 否则两边迟早对"色阶算不算本次"给出不同答案。
 */
function clearOneCopyRange(
  sheet: { Range?(address: string): WpsRange },
  address: string,
  formula: string,
  report: WpsCopyMarkCleanup
): boolean {
  let confirmed = true;
  try {
    const range = sheet.Range?.(address) as unknown as {
      FormatConditions?: { Count?: unknown; Item?(index: number): Rule | null | undefined };
    };
    const conditions = range?.FormatConditions;
    if (!conditions || typeof conditions.Item !== 'function') return false;
    /*
     * `Count` 缺失 / null / 非数字 **不是 0 条**：读不出来只说明这次读数不可信。
     * 当成 0 就等于断定"这个区域没有本次规则" —— 副本会带着本次高亮，却报告
     * "已清理、无需提示"。
     */
    const count = conditions.Count;
    if (typeof count !== 'number' || !Number.isSafeInteger(count) || count < 0) return false;
    for (let index = count; index >= 1; index -= 1) {
      const rule = conditions.Item(index);
      // 空句柄同样不是"这条不存在"：集合已不可信，继续往下删就成盲删了。
      if (!rule) { confirmed = false; continue; }
      const ruleFormula = rule.Formula1;
      if (ruleFormula == null) {
        /*
         * 公式读不出来有两种可能：色阶/数据条这类本来就没有公式的规则，或者
         * **是我们的、只是这一下读不出公式**。只有宿主明确给出可识别的非表达式
         * 类型才敢跳过；其余一律保留规则并记 unconfirmed —— 宁可多提示一次
         * "可能没清干净"，也不能把本次残留留在副本上粉饰太平。
         */
        if (!isKnownNonExpressionRule(rule)) confirmed = false;
        continue;
      }
      // 归属判据只有一个：本次完整随机公式。前缀/颜色/区域都不构成归属。
      if (ruleFormula !== formula) continue;
      if (typeof rule.Delete !== 'function') { confirmed = false; continue; }
      rule.Delete();
      report.cleared += 1;
    }
  } catch { return false; }
  return confirmed;
}

/**
 * 给本次执行的结果打审阅标记。
 *
 * `onFormula` 在公式生成后**立刻**回调，先于任何一次 `Add`：标记建到一半失败时，
 * 调用方仍拿得到本次的完整随机串，双视图副本据此精确清掉本次残留（不是"没有标记"
 * 就只能任由副本黄着）。公式只是归属凭据，不代表清理成功。
 */
export function markWpsResult(
  application: WpsApplication,
  plan: WpsPreparedPlan,
  onFormula?: (formula: string) => void
): WpsResultMarks {
  // 标记永远打在**原表**上：副本获得焦点不代表目标是副本。
  const workbook = wpsReadWorkbook(application);
  if (!workbook || String(workbook.FullName ?? '') !== plan.authorization.workbookFullName) throw new Error('标记目标工作簿已变化。');
  const formula = `=N("EB_REVIEW_${crypto.randomUUID().replace(/-/g, '')}")=0`;
  onFormula?.(formula);
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
  marks.ranges = tracked;
  const handles = new Map<string, Rule>();
  try {
    for (const [key, conditions] of ranges.entries()) {
      const [sheet, range] = JSON.parse(key) as [string, string];
      tracked.push({ sheet, range });
      // Persist before Add: the host may create a rule and then return null or throw.
      upsertRecord({ workbookFullName: marks.workbookFullName, formula, ranges: tracked });
      let rule = conditions.Add(2, undefined, formula); // xlExpression; never retry Add
      if (!rule) {
        const matches = ownedRules(conditions, formula);
        if (matches.length !== 1) throw new Error('WPS 未能建立可独立移除的审阅标记。');
        rule = matches[0];
      }
      if (rule.Formula1 !== formula || typeof rule.Delete !== 'function') throw new Error('WPS 未能建立可独立移除的审阅标记。');
      handles.set(key, rule);
      marks.rules.push(rule);
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
    const remaining: typeof tracked = [];
    for (const entry of tracked) {
      try {
        if (String(wpsReadWorkbook(application)?.FullName ?? '') !== marks.workbookFullName ||
            String(workbook.FullName ?? '') !== marks.workbookFullName) throw new Error('标记目标工作簿已变化。');
        const key = JSON.stringify([entry.sheet, entry.range]);
        const handle = handles.get(key);
        // A verified Add handle is sufficient even on older hosts without Item.
        // Otherwise enumerate the exact nonce, including Add-then-throw side effects.
        deleteOwnedRules(handle ? [handle] : ownedRules(ranges.get(key), formula), formula);
      } catch { remaining.push(entry); }
    }
    if (remaining.length > 0) {
      // 只写回**本次**这一条（区域收窄到真正没清掉的那些），别的任务/别的簿的记录不动。
      upsertRecord({ workbookFullName: marks.workbookFullName, formula, ranges: remaining });
      throw new Error(`${reason instanceof Error ? reason.message : '标记创建失败'}；${remaining.length} 处标记状态待核对、暂未移除，可在下次打开时重试清理。`);
    }
    // 本次的规则已经全部清掉（或根本没建出来）：只撤掉本次的归属记录。
    removeRecord(formula);
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
  deleteOwnedRules(marks.rules, marks.formula);
  const ranges = marks.ranges ?? readStored().find(record => record.formula === marks.formula && record.workbookFullName === marks.workbookFullName)?.ranges ?? [];
  for (const entry of ranges) {
    const range = marks.workbook.Worksheets.Item(entry.sheet).Range?.(entry.range) as unknown as { FormatConditions?: RuleCollection };
    if (ownedRules(range?.FormatConditions, marks.formula).length)
      throw new Error('审阅颜色标记尚未移除，可重试清理；不会再次恢复单元格。');
  }

  // 只撤掉本次的归属记录：同一本簿里别的任务（新旧两轮）的记录必须留着。
  removeRecord(marks.formula);
}

/**
 * 刷新/重载后处理上一次遗留的标记。
 *
 * 记录是**多条**（新旧两轮可能同时在等清理）。只按每条记录自己的 `Formula1`
 * （本次随机 nonce）认领：
 *  - 同一区域里 Formula1 不等于它的规则**一律不碰**（别的任务 / 用户自建）；
 *  - 只处理**当前这本簿**的记录；别的簿的记录原样保留（不串用、不顺手清）；
 *  - 某一条清不掉的区域留在**它自己**那条记录里，下次再试；绝不改成"按区域清空规则"。
 */
export function recoverWpsResultMarks(application: WpsApplication): string | null {
  const stored = readStored();
  if (stored.length === 0) return null;
  const active = wpsReadWorkbook(application);
  const activeFullName = String(active?.FullName ?? '');
  if (!active || !activeFullName) {
    return '当前没有可读取的工作簿，未做任何改动；请在原工作簿中打开后再清理。';
  }
  const mine = stored.filter((record) => record.workbookFullName === activeFullName);
  if (mine.length === 0) {
    return '上次审阅留下的标记属于另一本工作簿，未做任何改动；请回到那本工作簿后再清理。';
  }
  let removed = 0;
  const survivors: StoredRecord[] = [];
  for (const record of mine) {
    const remaining: MarkRange[] = [];
    for (const entry of record.ranges) {
      try {
        const holder = active.Worksheets.Item(entry.sheet).Range?.(entry.range) as unknown as {
          FormatConditions?: { Count?: number; Item?(index: number): Rule };
        };
        const matches = ownedRules(holder?.FormatConditions, record.formula);
        deleteOwnedRules(matches, record.formula);
        if (ownedRules(holder?.FormatConditions, record.formula).length) throw new Error("审阅颜色标记尚未移除");
        removed += matches.length;
      } catch { remaining.push(entry); }
    }
    if (remaining.length > 0) survivors.push({ ...record, ranges: remaining });
  }
  writeStored([...stored.filter((record) => record.workbookFullName !== activeFullName), ...survivors]);
  const pending = survivors.reduce((total, record) => total + record.ranges.length, 0);
  if (pending === 0) {
    return removed > 0 ? '已清理上次审阅遗留的标记。' : '上次审阅的标记已不存在，记录已清除。';
  }
  return removed > 0
    ? `已清理上次审阅遗留的 ${removed} 处标记，另有 ${pending} 处未清理；可稍后重试。`
    : `上次审阅遗留的 ${pending} 处标记未清理；可稍后重试。`;
}