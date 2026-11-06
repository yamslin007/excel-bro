// @vitest-environment jsdom
/**
 * 审阅标记的生命周期（WPS 端）。
 *
 * 模拟范围：只替身 **WPS 的 FormatConditions 与工作簿对象**；标记模块是真的。
 * 这里**不**走"宿主不支持标记"的降级分支：降级只证明"没有标记"，证明不了
 * 标记存在时的归属判定、部分失败与刷新认领。
 *
 * WPS 没有可读的规则 id，归属凭据是本次随机 `Formula1`。所以"同一区域里
 * 另一条规则"必须活下来 —— 那是"不按区域/颜色批量删"的可观测后果。
 */
import { ResultReview, type ResultReviewReceipt } from "@excel-bro/core/resultReview";
import type { WpsResultMarks } from "./wpsResultReviewMarkers";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { WpsApplication, WpsWorkbook } from "./wpsHost";
import {
  clearWpsResultMarks,
  markWpsResult,
  recoverWpsResultMarks,
  clearReviewMarksFromCopySheet
} from "./wpsResultReviewMarkers";

type RuleState = {
  /** `null` = 宿主这一下读不出公式（我们自己的规则或非表达式规则都可能是这样）。 */
  formula: string | null;
  /** `XlFormatConditionType`；替身里 `undefined` 表示"类型读不出来"。 */
  type?: number;
  deleted: boolean;
  interior: { color: number; pattern: unknown } | null;
  borders: Record<number, { color: number; weight: unknown }>;
};

function installHost(options: {
  failOnAdd?: number; failOnDelete?: boolean; ignoreDelete?: boolean; ranges?: string[];
  /** 宿主"建出来了却回空句柄"：规则真的在集合里，`Add` 的返回值是 null。 */
  addReturnsNull?: boolean;
} = {}) {
  const byRange = new Map<string, RuleState[]>();
  const ranges = options.ranges ?? ["A1:B1"];
  for (const range of ranges) byRange.set(range, []);
  let seq = 0;

  /**
   * 一条规则的**完整句柄**。`Item` 也交这个，而不是只给 `Formula1` + `Delete` ——
   * 找回（Add 回空）之后紧接着要给这条规则上底色与边框，句柄缺接口就会被误判成
   * "宿主不支持标记"。
   */
  const handleOf = (state: RuleState) => ({
    get Formula1() { return state.formula; },
    set Formula1(next: string | null) { state.formula = next; },
    get Type() { return state.type; },
    Borders: { Item: (edge: number) => ({
      get Color() { return state.borders[edge]?.color ?? 0; },
      set Color(v: number) { state.borders[edge] = { color: v, weight: state.borders[edge]?.weight }; },
      get Weight() { return state.borders[edge]?.weight; },
      set Weight(v: unknown) { state.borders[edge] = { color: state.borders[edge]?.color ?? 0, weight: v }; },
      get LineStyle() { return 1; },
      set LineStyle(_v: unknown) {}
    }) },
    Interior: {
      get Color() { return state.interior?.color ?? 0; },
      set Color(v: number) { state.interior = { color: v, pattern: state.interior?.pattern }; },
      get Pattern() { return state.interior?.pattern; },
      set Pattern(v: unknown) { state.interior = { color: state.interior?.color ?? 0, pattern: v }; }
    },
    Delete: () => {
      if (options.failOnDelete) throw new Error("删除被宿主拒绝");
      if (!options.ignoreDelete) state.deleted = true;
    }
  });

  const rangeObject = (range: string) => {
    const list = byRange.get(range) ?? [];
    const conditions = {
      get Count() { return list.filter(rule => !rule.deleted).length; },
      // 真宿主在 Add 时就带上公式；替身必须同样回填，否则归属判据永远对不上。
      Add: (_type: number, _operator?: unknown, formula?: string) => {
        seq += 1;
        if (options.failOnAdd === seq) throw new Error("第 2 处标记创建失败");
        // 本次标记永远是表达式规则（xlExpression = 2）。
        const state: RuleState = { formula: formula ?? "", type: 2, deleted: false, interior: null, borders: {} };
        list.push(state);
        return options.addReturnsNull ? null : handleOf(state);
      },
      Item: (index: number) => {
        const live = list.filter(rule => !rule.deleted)[index - 1];
        return live ? handleOf(live) : { get Formula1() { return undefined; } };
      }
    };
    return { FormatConditions: conditions };
  };

  const sheet = { Name: "数据", Range: (range: string) => rangeObject(range) };
  const workbook = {
    Name: "测试.xlsx", FullName: "C:\\Data\\测试.xlsx", Path: "C:\\Data", ReadOnly: false,
    Worksheets: { Count: 1, Item: () => sheet },
    ActiveSheet: sheet
  } as unknown as WpsWorkbook;
  const application = { Version: "12.0", ActiveWorkbook: workbook, ActiveSheet: sheet } as unknown as WpsApplication;
  return {
    application,
    addCount: () => seq,
    live: (range = ranges[0]) => (byRange.get(range) ?? []).filter(rule => !rule.deleted).map(rule => rule.formula),
    state: (range = ranges[0]) => (byRange.get(range) ?? []).filter(rule => !rule.deleted),
    /** 同区域里别人（用户/别的任务）的规则；`formula` 传 null 表示"读不出公式的规则"。 */
    pushForeign: (range: string, formula: string | null, type?: number) => {
      (byRange.get(range) ?? []).push({ formula, type, deleted: false, interior: null, borders: {} });
    },
    /** 用户自建的色阶规则（xlColorScale = 3）：结构上就没有 Formula1。 */
    pushColorScale: (range: string) => {
      (byRange.get(range) ?? []).push({ formula: null, type: 3, deleted: false, interior: null, borders: {} });
    }
  };
}

const prepared = {
  authorization: { workbookFullName: "C:\\Data\\测试.xlsx", sheets: [{ name: "数据", index: 1 }] },
  actions: [
    { sheetName: "数据", range: "A1:B1", type: "writeValues" },
    { sheetName: "数据", range: "C1:D1", type: "writeValues" }
  ]
} as never;

beforeEach(() => { localStorage.clear(); });

describe("WPS 审阅标记", () => {
  it("删除调用静默失效时保留归属记录，重载清理也不得报告成功", () => {
    const host = installHost({ ignoreDelete: true });
    const marks = markWpsResult(host.application, single("A1:B1"));
    expect(() => clearWpsResultMarks(host.application, marks)).toThrow("审阅颜色标记尚未移除");
    expect(host.live()).toContain(marks.formula);
    expect(storedRecords()).toHaveLength(1);
    expect(recoverWpsResultMarks(host.application)).toContain("未清理");
    expect(storedRecords()).toHaveLength(1);
  });
  it("逐条落盘：创建部分成功后失败，已创建的可追踪并清理", () => {
    const host = installHost({ ranges: ["A1:B1", "C1:D1"], failOnAdd: 2 });
    expect(() => markWpsResult(host.application, prepared)).toThrow(/第 2 处标记创建失败/);
    // 第一条被清掉，记录不留。
    expect(host.live("A1:B1")).toHaveLength(0);
    expect(localStorage.getItem("excel-bro.wps.reviewMarks.v1")).toBeNull();
  });

  it("创建失败且清理也失败：保留归属记录（只有表名/区域/公式，没有业务值）", () => {
    const host = installHost({ ranges: ["A1:B1", "C1:D1"], failOnAdd: 2, failOnDelete: true });
    expect(() => markWpsResult(host.application, prepared)).toThrow(/暂未移除/);
    const stored = JSON.parse(localStorage.getItem("excel-bro.wps.reviewMarks.v1") ?? "null");
    expect(stored.records[0].ranges).toEqual([{ sheet: "数据", range: "A1:B1" }]);
    expect(JSON.stringify(stored)).not.toContain('"values"');
  });

  it("部分清理失败只留未清掉的那一处：宿主恢复后重开只清它，用户的同区规则仍在", () => {
    // 选项对象按引用被宿主持有，改它就是"宿主恢复"—— 不是另起一个替身再断言一次。
    const options: { failOnAdd?: number; failOnDelete?: boolean; ranges?: string[] } =
      { ranges: ["A1:B1", "C1:D1"], failOnAdd: 2, failOnDelete: true };
    const host = installHost(options);
    host.pushForeign("A1:B1", "=A1>0");
    expect(() => markWpsResult(host.application, prepared)).toThrow(/1 处标记状态待核对/);
    // A1 的本次规则还在（Delete 被拒），C1 从未创建 —— 记录里只该有 A1 这一处。
    expect(host.state("A1:B1")).toHaveLength(2);
    expect(host.state("C1:D1")).toHaveLength(0);
    const stored = JSON.parse(localStorage.getItem("excel-bro.wps.reviewMarks.v1") ?? "null");
    expect(stored.records).toHaveLength(1);
    expect(stored.records[0].ranges).toEqual([{ sheet: "数据", range: "A1:B1" }]);

    options.failOnDelete = false;
    expect(recoverWpsResultMarks(host.application)).toContain("已清理");
    // 只删本次公式那一条，用户的规则留在原地。
    expect(host.live("A1:B1")).toEqual(["=A1>0"]);
    expect(localStorage.getItem("excel-bro.wps.reviewMarks.v1")).toBeNull();
  });

  it("刷新后只按本次随机公式认领，同区域其它规则不碰", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    const marks = markWpsResult(host.application, prepared);
    // 用户/别的任务在同一区域另建了一条规则。
    host.pushForeign("A1:B1", "=A1>0");
    expect(host.live("A1:B1")).toHaveLength(2);

    const note = recoverWpsResultMarks(host.application);
    expect(note).toContain("已清理");
    // 只剩别人的那条 —— 证明"不按区域批量删"。
    expect(host.live("A1:B1")).toEqual(["=A1>0"]);
    expect(localStorage.getItem("excel-bro.wps.reviewMarks.v1")).toBeNull();
    expect(marks.formula).toContain("EB_REVIEW_");
  });

  it("工作簿身份不符：不碰规则、不删记录", () => {
    const host = installHost();
    markWpsResult(host.application, prepared);
    const other = { ...host.application, ActiveWorkbook: { ...(host.application.ActiveWorkbook as object), FullName: "C:\\Data\\Other.xlsx", Name: "Other.xlsx" } } as unknown as WpsApplication;
    const note = recoverWpsResultMarks(other);
    expect(note).toContain("另一本工作簿");
    expect(host.live()).toHaveLength(1);
    expect(localStorage.getItem("excel-bro.wps.reviewMarks.v1")).not.toBeNull();
  });

  it("确认清理只删本次公式对应的规则", () => {
    const host = installHost();
    const marks = markWpsResult(host.application, prepared);
    host.pushForeign("A1:B1", "=B1>0");
    clearWpsResultMarks(host.application, marks);
    expect(host.live()).toEqual(["=B1>0"]);
  });
});
describe("WPS 审阅标记呈现：底色与边框", () => {
  it("不修改填充的动作：浅黄底色 + 醒目边框", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    markWpsResult(host.application, prepared);
    const [rule] = host.state("A1:B1");
    // 底色 BGR：0xFFF3B0 → b=0xFF, g=0xF3, r=0xB0 → (0xFF<<16)|(0xF3<<8)|0xB0
    expect(rule.interior?.pattern).toBe(1);
    expect(rule.interior?.color).toBe((0xb0 << 16) | (0xf3 << 8) | 0xff);
    // 边框醒目金黄（0xB8860B → BGR 0x0B86B8）
    expect(rule.borders[7]?.color).toBe((0x0b << 16) | (0x86 << 8) | 0xb8);
    expect(rule.borders[7]?.weight).toBe(2);
  });

  it("涉及填充变化（setFill）的动作：只画边框，不叠底色", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    const fillPlan = {
      authorization: { workbookFullName: "C:\\Data\\测试.xlsx", sheets: [{ name: "数据", index: 1 }] },
      actions: [{ sheetName: "数据", range: "A1:B1", type: "setFill", color: "#FF0000" }]
    } as never;
    markWpsResult(host.application, fillPlan);
    const [rule] = host.state("A1:B1");
    expect(rule.interior).toBeNull();
    expect(rule.borders[7]?.color).toBe((0x0b << 16) | (0x86 << 8) | 0xb8);
  });

  it("混合计划：写值区域叠底色、setFill 区域只边框，各按各的", () => {
    const host = installHost({ ranges: ["A1:B1", "C1:D1"] });
    const mixed = {
      authorization: { workbookFullName: "C:\\Data\\测试.xlsx", sheets: [{ name: "数据", index: 1 }] },
      actions: [
        { sheetName: "数据", range: "A1:B1", type: "writeValues" },
        { sheetName: "数据", range: "C1:D1", type: "setFill", color: "#00FF00" }
      ]
    } as never;
    markWpsResult(host.application, mixed);
    const [writeRule] = host.state("A1:B1");
    const [fillRule] = host.state("C1:D1");
    expect(writeRule.interior?.color).toBe((0xb0 << 16) | (0xf3 << 8) | 0xff);
    expect(fillRule.interior).toBeNull();
  });

  it("副本清理只删本次完整公式对应的规则：另一轮审阅的 EB_REVIEW_ 规则也保留", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    const marks = markWpsResult(host.application, prepared);
    // 用户规则 + **另一轮审阅**的规则（同 EB_REVIEW_ 前缀、不同 nonce）。
    host.pushForeign("A1:B1", "=A1>0");
    host.pushForeign("A1:B1", '=N("EB_REVIEW_another_run_123")=0');
    expect(host.live("A1:B1")).toHaveLength(3);

    // 模拟一张通过 Copy() 带来的副本表：同一区域有本次规则、用户规则、另一轮审阅规则。
    const copySheet = {
      Range: (range: string) => {
        const obj = (host.application.ActiveWorkbook as unknown as {
          Worksheets: { Item(): { Range(r: string): unknown } }
        }).Worksheets.Item().Range(range);
        return obj as never;
      }
    };
    // 只给本次完整公式：另一轮审阅的规则即使同前缀也不删。
    clearReviewMarksFromCopySheet(copySheet as never, ["A1:B1"], marks.formula);
    expect(host.live("A1:B1")).toEqual(["=A1>0", '=N("EB_REVIEW_another_run_123")=0']);
  });
});

it("does not overlay business borders, and skips overlays when both border and fill are changed",()=>{
 const host=installHost({ranges:["A1:B1"]});
 const base={authorization:{workbookFullName:"C:\\Data\\测试.xlsx"}};
 const border={type:"setBorders",sheetName:"数据",range:"A1:B1"};
 const marks=markWpsResult(host.application,{...base,actions:[border]} as never);
 expect(host.state()[0].interior?.color).toBe(0xB0F3FF);
 expect(host.state()[0].borders).toEqual({});
 clearWpsResultMarks(host.application,marks);
 const combined=markWpsResult(host.application,{...base,actions:[border,{...border,type:"setFill"}]} as never);
 expect(host.live()).toHaveLength(0);expect(combined.rules).toHaveLength(0);
});


// Add may create a rule while the WPS bridge returns no object.
function faultHost() {
  const host = installHost({ ranges: ['A1:B1', 'C1:D1'] });
  const faults = { nullAdd: false, skipCreate: false, throwAfterAdd: false, nullItem: false, invalidCount: false };
  const sheet = host.application.ActiveWorkbook!.Worksheets.Item('数据');
  const originalRange = sheet.Range!.bind(sheet);
  let adds = 0;
  sheet.Range = ((address: string) => {
    const range = originalRange(address) as unknown as { FormatConditions: {
      Count: number; Add(...args: unknown[]): object; Item(index: number): { Formula1?: unknown }
    } };
    const original = range.FormatConditions;
    const handles = new Map<unknown, object>();
    return { FormatConditions: {
      get Count() { return faults.invalidCount ? Number.NaN : original.Count; },
      Add(...args: unknown[]) {
        adds += 1;
        if (faults.skipCreate) return null;
        const rule = original.Add(...args);
        handles.set(args[2], rule);
        if (faults.throwAfterAdd) throw new Error('创建后桥接失败');
        return faults.nullAdd ? null : rule;
      },
      Item(index: number) {
        if (faults.nullItem) return null;
        const rule = original.Item(index);
        return handles.get(rule.Formula1) ?? rule;
      }
    } };
  }) as typeof sheet.Range;
  return { ...host, faults, adds: () => adds };
}

describe('WPS 空规则句柄与可追踪清理', () => {
  it('Add 返回 null 但规则已创建：按完整公式找回、着色、每区只 Add 一次', () => {
    const host = faultHost();
    host.faults.nullAdd = true;
    host.pushForeign('A1:B1', '=N("EB_REVIEW_other")=0');
    const marks = markWpsResult(host.application, prepared);
    expect(host.adds()).toBe(2);
    expect(marks.rules).toHaveLength(2);
    expect(host.state('A1:B1')[1].interior?.color).toBe(0xB0F3FF);
    clearWpsResultMarks(host.application, marks);
    expect(host.live('A1:B1')).toEqual(['=N("EB_REVIEW_other")=0']);
    expect(host.live('C1:D1')).toEqual([]);
    expect(localStorage.getItem('excel-bro.wps.reviewMarks.v1')).toBeNull();
  });

  it('Add 返回 null 且没有创建：明确报错、不留下虚构标记', () => {
    const host = faultHost();
    host.faults.skipCreate = true;
    expect(() => markWpsResult(host.application, prepared)).toThrow(/未能建立/);
    expect(host.adds()).toBe(1);
    expect(host.live()).toEqual([]);
    expect(localStorage.getItem('excel-bro.wps.reviewMarks.v1')).toBeNull();
  });

  it('Add 已产生副作用再抛错：本次规则仍能清理，外来规则保留', () => {
    const host = faultHost();
    host.faults.throwAfterAdd = true;
    host.pushForeign('A1:B1', '=A1>0');
    expect(() => markWpsResult(host.application, prepared)).toThrow('创建后桥接失败');
    expect(host.live()).toEqual(['=A1>0']);
    expect(localStorage.getItem('excel-bro.wps.reviewMarks.v1')).toBeNull();
  });

  it.each(['nullItem', 'invalidCount'] as const)('集合暂不可读 (%s)：保留归属记录，恢复可读后只清理本次规则', fault => {
    const host = faultHost();
    host.faults.nullAdd = true;
    host.faults[fault] = true;
    host.pushForeign('A1:B1', '=A1>0');
    expect(() => markWpsResult(host.application, prepared)).toThrow(/暂未移除/);
    const stored = localStorage.getItem('excel-bro.wps.reviewMarks.v1');
    expect(JSON.parse(stored!).records[0].ranges).toEqual([{ sheet: '数据', range: 'A1:B1' }]);
    expect(recoverWpsResultMarks(host.application)).toContain('未清理');
    expect(JSON.parse(localStorage.getItem('excel-bro.wps.reviewMarks.v1')!).records[0].ranges)
      .toEqual([{ sheet: '数据', range: 'A1:B1' }]);
    expect(host.live()).toHaveLength(2);
    host.faults[fault] = false;
    expect(recoverWpsResultMarks(host.application)).toContain('已清理');
    expect(host.live()).toEqual(['=A1>0']);
    expect(localStorage.getItem('excel-bro.wps.reviewMarks.v1')).toBeNull();
  });
});


it.each(['confirm', 'discard'] as const)('空句柄无法找回时仍能 %s，审阅与比较不会重执行', async action => {
  const host = faultHost();
  host.faults.skipCreate = true;
  const review = new ResultReview<string, ResultReviewReceipt, WpsResultMarks, string>();
  const execute = vi.fn(async () => ({ executionId: 'written-once', verified: true, undoable: true }));
  const restore = vi.fn(async () => {});
  const accepted = vi.fn();
  const operations = {
    isCurrent: () => true, prepare: async () => 'verified-before', execute,
    mark: async () => markWpsResult(host.application, prepared),
    clearMarker: async (marks: WpsResultMarks) => clearWpsResultMarks(host.application, marks),
    showBefore: vi.fn(async () => 'before-copy'), closeBefore: async () => null,
    restore, accepted
  };
  expect(await review.start('plan-1', operations)).toBe(true);
  expect(review.getSnapshot()).toMatchObject({ phase: 'reviewing', marked: false, notice: expect.stringContaining('结果已写入') });
  expect(await review.compare('plan-1')).toBe(true);
  expect(operations.showBefore).toHaveBeenCalledOnce();
  expect(await review[action]('plan-1')).toBe(true);
  expect(review.getSnapshot().phase).toBe(action === 'confirm' ? 'confirmed' : 'discarded');
  expect(execute).toHaveBeenCalledOnce();
  expect(restore).toHaveBeenCalledTimes(action === 'discard' ? 1 : 0);
  expect(accepted).toHaveBeenCalledTimes(action === 'confirm' ? 1 : 0);
  expect(await review.start('plan-1', operations)).toBe(false);
  expect(execute).toHaveBeenCalledOnce();
});

/**
 * A1-F1 的两处残留（2026-10-25）。
 *
 * 1. 标记**建到一半失败**时，本次标记的完整随机公式必须仍然交得出去 —— 双视图
 *    副本要靠它精确清掉本次残留；否则"执行前副本"也黄成一片，对比失去意义。
 * 2. 归属记录**不再是一个槽**：旧任务未清理时开始新审阅，两条记录必须并存、
 *    各自清理。单槽会把旧记录覆盖掉，旧规则从此没有任何自动清理线索。
 */
const STORAGE_KEY = "excel-bro.wps.reviewMarks.v1";
type StoredRecord = { workbookFullName: string; formula: string; ranges: Array<{ sheet: string; range: string }> };
function storedRecords(): StoredRecord[] {
  const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "null");
  if (!parsed) return [];
  // 读得出 v1 单条（历史记录）也读得出 v2 列表：兼容是产品要求，不是测试细节。
  return Array.isArray(parsed.records) ? parsed.records : [{ ...parsed }];
}
const single = (range: string) => ({
  authorization: { workbookFullName: "C:\\Data\\测试.xlsx" },
  actions: [{ sheetName: "数据", range, type: "writeValues" }]
}) as never;

describe("A1-F1 残留：本次公式与多轮归属", () => {
  it("标记部分失败仍交出本次公式，副本只按它清残留，用户/其他任务规则不动", () => {
    const options: { failOnAdd?: number; failOnDelete?: boolean; ranges?: string[] } =
      { ranges: ["A1:B1", "C1:D1"], failOnAdd: 2, failOnDelete: true };
    const host = installHost(options);
    host.pushForeign("A1:B1", "=A1>0");
    host.pushForeign("A1:B1", '=N("EB_REVIEW_other")=0');
    const seen: string[] = [];
    expect(() => markWpsResult(host.application, prepared, formula => seen.push(formula)))
      .toThrow(/暂未移除/);
    // 失败也交出了公式，而且是**完整随机**那一条（不是前缀）。
    expect(seen).toHaveLength(1);
    expect(seen[0]).toMatch(/^=N\("EB_REVIEW_[0-9a-f]{32}"\)=0$/);
    // 归属记录里存的就是它 + 区域，没有任何业务值。
    const records = storedRecords();
    expect(records.map(record => record.formula)).toEqual(seen);
    expect(Object.keys(records[0]).sort()).toEqual(["formula", "ranges", "workbookFullName"]);
    expect(JSON.stringify(records)).not.toContain('"values"');

    // 双视图副本：Copy() 把三条规则一起带过去，只清本次那一条，并如实报告结果。
    const copyRules = [
      { formula: seen[0], deleted: false },
      { formula: "=A1>0", deleted: false },
      { formula: '=N("EB_REVIEW_other")=0', deleted: false }
    ];
    const copySheet = {
      Range: () => ({ FormatConditions: {
        get Count() { return copyRules.filter(rule => !rule.deleted).length; },
        Item: (index: number) => {
          const live = copyRules.filter(rule => !rule.deleted)[index - 1];
          return live
            ? { get Formula1() { return live.formula; }, Delete: () => { live.deleted = true; } }
            : null;
        }
      } })
    };
    expect(clearReviewMarksFromCopySheet(copySheet as never, ["A1:B1"], seen[0]))
      .toEqual({ cleared: 1, unconfirmed: 0 });
    expect(copyRules.filter(rule => !rule.deleted).map(rule => rule.formula))
      .toEqual(["=A1>0", '=N("EB_REVIEW_other")=0']);
    // 原表上的残留**没有**被顺手删掉：副本清理不碰原表，也不改归属记录。
    expect(host.live("A1:B1")).toHaveLength(3);
    expect(storedRecords()).toHaveLength(1);
  });

  it("副本集合读不出来时如实报告 unconfirmed，不声称已清理", () => {
    const copySheet = { Range: () => ({ FormatConditions: undefined }) };
    expect(clearReviewMarksFromCopySheet(copySheet as never, ["A1:B1"], '=N("EB_REVIEW_x")=0'))
      .toEqual({ cleared: 0, unconfirmed: 1 });
  });

  it("旧记录未清理时开始新审阅：两条记录并存，确认只清本次", () => {
    const options: { failOnAdd?: number; failOnDelete?: boolean; ranges?: string[] } =
      { ranges: ["A1:B1", "C1:D1"], failOnAdd: 2, failOnDelete: true };
    const host = installHost(options);
    host.pushForeign("A1:B1", "=A1>0");
    let first = "";
    expect(() => markWpsResult(host.application, prepared, formula => { first = formula; }))
      .toThrow(/暂未移除/);
    expect(storedRecords().map(record => record.formula)).toEqual([first]);

    // 宿主恢复正常后，用户在旧标记还在的情况下开始了一次**新**审阅。
    options.failOnDelete = false;
    const second = markWpsResult(host.application, single("A1:B1"));
    expect(second.formula).not.toBe(first);
    // 阈值：两条记录分别登记；单槽实现会在这里只剩 1 条（旧记录被覆盖）。
    expect(storedRecords().map(record => record.formula)).toEqual([first, second.formula]);
    // 用户规则在最前（先建的），两轮本次规则各一条，互不覆盖。
    expect(host.live("A1:B1")).toEqual(["=A1>0", first, second.formula]);

    // 确认这一轮：只删本轮规则、只删本轮记录，旧规则与旧记录都还在。
    clearWpsResultMarks(host.application, second);
    expect(host.live("A1:B1")).toEqual(["=A1>0", first]);
    expect(storedRecords().map(record => record.formula)).toEqual([first]);
    // 旧记录仍可被后续重开认领（不是"丢了线索"）。
    expect(recoverWpsResultMarks(host.application)).toContain("已清理");
    expect(host.live("A1:B1")).toEqual(["=A1>0"]);
    expect(localStorage.getItem(STORAGE_KEY)).toBeNull();
  });

  it("两轮记录在重载后分别清理：一处清不掉只留它自己的记录", () => {
    const host = installHost({ ranges: ["A1:B1", "C1:D1"] });
    host.pushForeign("A1:B1", "=A1>0");
    // 两轮各自成功（第 1 轮还没确认就开始了第 2 轮）。
    const first = markWpsResult(host.application, single("A1:B1"));
    const second = markWpsResult(host.application, single("C1:D1"));
    expect(storedRecords().map(record => record.formula)).toEqual([first.formula, second.formula]);

    // 重开：C1:D1 的集合这次读不出来 —— 只保留它的记录，A1:B1 照常清掉。
    const sheet = host.application.ActiveWorkbook!.Worksheets.Item("数据");
    const originalRange = sheet.Range!.bind(sheet);
    sheet.Range = ((address: string) => address === "C1:D1" ? {} : originalRange(address)) as never;
    const partial = recoverWpsResultMarks(host.application);
    expect(partial).toContain("已清理");
    expect(partial).toContain("未清理");
    expect(host.live("A1:B1")).toEqual(["=A1>0"]);
    expect(host.live("C1:D1")).toEqual([second.formula]);
    expect(storedRecords().map(record => record.formula)).toEqual([second.formula]);

    // 宿主恢复后重试：这一轮的记录被清掉，用户的规则仍在。
    sheet.Range = originalRange as never;
    expect(recoverWpsResultMarks(host.application)).toContain("已清理");
    expect(host.live("C1:D1")).toEqual([]);
    expect(host.live("A1:B1")).toEqual(["=A1>0"]);
    expect(localStorage.getItem(STORAGE_KEY)).toBeNull();
  });

  it("别的簿的记录不串用：本簿只清自己的，另一簿的记录与规则都不碰", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    const mine = markWpsResult(host.application, single("A1:B1"));
    expect(mine.formula).not.toBe("=N(\"EB_REVIEW_otherbook\")=0");
    const other: StoredRecord = { workbookFullName: "C:\\Data\\其他.xlsx",
      formula: '=N("EB_REVIEW_otherbook")=0', ranges: [{ sheet: "数据", range: "A1:B1" }] };
    // 危险的真实形态：另一本簿的表被复制到本簿，带着它那条 EB_REVIEW_ 规则。
    // 按"公式相同就删"会顺手删掉别人的标记 —— 记录说它属于**另一本簿**。
    host.pushForeign("A1:B1", other.formula);
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ v: 2, records: [...storedRecords(), other] }));

    expect(recoverWpsResultMarks(host.application)).toContain("已清理");
    expect(host.live("A1:B1")).toEqual([other.formula]);
    // 另一本簿的记录与它的规则都不能被"顺手清掉"或"顺手删掉"。
    expect(storedRecords()).toEqual([other]);
  });

  it("v1 单槽记录照样被认领：读得出、清得掉；下一次登记把它迁移进列表", () => {
    const host = installHost({ ranges: ["A1:B1", "C1:D1"] });
    const legacy = markWpsResult(host.application, single("A1:B1"));
    // 手工降级成旧版单条形状（顶层 formula/ranges + v:1），模拟历史记录。
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ v: 1,
      workbookFullName: "C:\\Data\\测试.xlsx", formula: legacy.formula,
      ranges: [{ sheet: "数据", range: "A1:B1" }] }));
    expect(storedRecords().map(record => record.formula)).toEqual([legacy.formula]);

    // 新审阅开在另一个区域：v1 那条必须被带进列表，而不是被覆盖。
    const fresh = markWpsResult(host.application, single("C1:D1"));
    expect(storedRecords().map(record => record.formula)).toEqual([legacy.formula, fresh.formula]);
    expect(JSON.parse(localStorage.getItem(STORAGE_KEY)!).v).toBe(2);

    // 两条分别清理，清完整个 key 消失。
    expect(recoverWpsResultMarks(host.application)).toContain("已清理");
    expect(host.live("A1:B1")).toEqual([]);
    expect(host.live("C1:D1")).toEqual([]);
    expect(localStorage.getItem(STORAGE_KEY)).toBeNull();
  });

  it.each(['confirm', 'discard'] as const)(
    '部分失败留下的残留不改变业务结论：%s 只执行一次、残留记录不被丢弃',
    async action => {
      const options: { failOnAdd?: number; failOnDelete?: boolean; ranges?: string[] } =
        { ranges: ["A1:B1", "C1:D1"], failOnAdd: 2, failOnDelete: true };
      const host = installHost(options);
      const review = new ResultReview<string, ResultReviewReceipt, WpsResultMarks, string>();
      const execute = vi.fn(async () => ({ executionId: 'written-once', verified: true, undoable: true }));
      const restore = vi.fn(async () => {});
      const accepted = vi.fn();
      let formula = "";
      const operations = {
        isCurrent: () => true, prepare: async () => 'verified-before', execute,
        mark: async () => markWpsResult(host.application, prepared, value => { formula = value; }),
        clearMarker: async (marks: WpsResultMarks) => clearWpsResultMarks(host.application, marks),
        showBefore: vi.fn(async () => 'before-copy'), closeBefore: async () => null,
        restore, accepted
      };
      expect(await review.start('plan-1', operations)).toBe(true);
      expect(review.getSnapshot()).toMatchObject({ phase: 'reviewing', marked: false });
      expect(formula).toContain('EB_REVIEW_');
      // 残留仍在、记录仍在：确认/放弃都不是"清理残留"的时机，也不能把它删掉。
      expect(host.live("A1:B1")).toHaveLength(1);
      expect(storedRecords().map(record => record.formula)).toEqual([formula]);

      await review.compare('plan-1');
      expect(await review[action]('plan-1')).toBe(true);
      expect(review.getSnapshot().phase).toBe(action === 'confirm' ? 'confirmed' : 'discarded');
      expect(execute).toHaveBeenCalledOnce();
      expect(restore).toHaveBeenCalledTimes(action === 'discard' ? 1 : 0);
      expect(accepted).toHaveBeenCalledTimes(action === 'confirm' ? 1 : 0);
      expect(storedRecords().map(record => record.formula)).toEqual([formula]);
    });
});

/**
 * A1-F1-a 副本清理的**判据**：副本上"读不出来"不等于"没有"。
 *
 * 副本清理的结论只有两种：清掉了（可以放心说副本干净）、或者没法核实。
 * 把读数失败当成"0 条规则"或"这条不存在"，就会在副本明明还黄着的时候报
 * `unconfirmed: 0` —— 用户看到的是"两张表同样高亮"，却收不到任何提示。
 */
describe("A1-F1-a 副本清理判据：读不出来不等于没有", () => {
  const RUN = '=N("EB_REVIEW_thisrun")=0';
  const otherRun = '=N("EB_REVIEW_otherrun")=0';
  /** 把一种宿主读数形态包成副本表的 `Range()`；集合原样交给清理函数。 */
  const copySheetWith = (conditions: unknown) => ({
    Range: () => ({ FormatConditions: conditions })
  });

  it("Count 缺失 / null / 非数字：一律无法核实，绝不当成 0 条规则", () => {
    const rule = { formula: RUN, deleted: false };
    // 旧实现 `Number(undefined ?? 0)` ⇒ 0 条 ⇒ 一声不响地"清理成功"。
    expect(clearReviewMarksFromCopySheet(copySheetWith({ Item: () => rule }) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
    expect(clearReviewMarksFromCopySheet(copySheetWith({ Count: null, Item: () => rule }) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
    // 宿主有时把 Count 当文本给：那也不是"确定有 2 条"。
    expect(clearReviewMarksFromCopySheet(copySheetWith({ Count: "2", Item: () => rule }) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
    // 读数不可信时一个都不许删。
    expect(rule.deleted).toBe(false);
  });

  it("Item 返回空：那一处无法核实，不当成\"这条规则不存在\"", () => {
    const foreign = { formula: "=A1>0", deleted: false };
    // 索引 2 读得出（用户规则），索引 1 是空句柄 —— 集合已经不可信。
    const conditions = { Count: 2, Item: (index: number) => (index === 1 ? null : foreign) };
    expect(clearReviewMarksFromCopySheet(copySheetWith(conditions) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
    expect(foreign.deleted).toBe(false);

    expect(clearReviewMarksFromCopySheet(
      copySheetWith({ Count: 1, Item: () => undefined }) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
  });

  it("表达式规则却读不出公式：保留规则并记为无法核实，不许猜成色阶/数据条", () => {
    // 宿主明确说是表达式规则（xlExpression=2），只是这一下拿不到 Formula1。
    const expression = { Type: 2, Formula1: null, deleted: false };
    expect(clearReviewMarksFromCopySheet(copySheetWith({ Count: 1, Item: () => expression }) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
    expect(expression.deleted).toBe(false);

    // 连 Type 都读不出来：更不许替宿主"认定"它是色阶。
    const anonymous = { Formula1: undefined, deleted: false };
    expect(clearReviewMarksFromCopySheet(copySheetWith({ Count: 1, Item: () => anonymous }) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
    expect(anonymous.deleted).toBe(false);
  });

  it("类型不认识（未知枚举 / 类型本身读不成数字）且读不出公式：仍记为无法核实", () => {
    const unknownType = { Type: 99, Formula1: null, deleted: false };
    expect(clearReviewMarksFromCopySheet(copySheetWith({ Count: 1, Item: () => unknownType }) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
    // 类型是文本：不能 `Number()` 一下就当枚举值用。
    const textType = { Type: "3", Formula1: null, deleted: false };
    expect(clearReviewMarksFromCopySheet(copySheetWith({ Count: 1, Item: () => textType }) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
    expect(unknownType.deleted || textType.deleted).toBe(false);
  });

  it("宿主明确给出非表达式类型（色阶/数据条/图标集）：安全跳过，不报无法核实", () => {
    // 这三种类型结构上就没有 Formula1，而且一定不是本次标记（本次永远是 xlExpression）。
    // 若把它们当"读不出来"，用户只要放过一个色阶，每次开双视图都会收到假警报。
    const rules = [{ Type: 3 }, { Type: 4 }, { Type: 6 }];
    const conditions = { Count: rules.length, Item: (index: number) => rules[index - 1] };
    expect(clearReviewMarksFromCopySheet(copySheetWith(conditions) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 0 });

    // 和本次规则混在一起时：本次的清掉、色阶留着，且不产生假警报。
    const mine = { Type: 2, formula: RUN, deleted: false };
    const mixed = {
      Count: 2,
      Item: (index: number) => (index === 1
        ? { Type: 3 }
        : { Type: 2, get Formula1() { return mine.formula; }, Delete: () => { mine.deleted = true; } })
    };
    expect(clearReviewMarksFromCopySheet(copySheetWith(mixed) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 1, unconfirmed: 0 });
    expect(mine.deleted).toBe(true);
  });

  it("每个区域最多计一次：同一区域多处读不出、重复地址都不放大计数", () => {
    // 一个区域里三条规则全读不出 —— 仍是"这一处没法核实"，不是 3 处。
    const allNull = copySheetWith({ Count: 3, Item: () => null });
    expect(clearReviewMarksFromCopySheet(allNull as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
    // 同一个地址重复给出也只算一次。
    expect(clearReviewMarksFromCopySheet(allNull as never, ["A1:B1", "A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 1 });
    // 两个不同区域各读不出一次 —— 按区域计 2 次。
    const twoBroken = { Range: () => ({ FormatConditions: { Count: 1, Item: () => null } }) };
    expect(clearReviewMarksFromCopySheet(twoBroken as never, ["A1:B1", "C1:D1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 2 });
  });

  it("一处读不出不影响另一处的正常清理：清掉的照实计，没核实的按区域计一次", () => {
    const mine = { formula: RUN, deleted: false };
    const sheet = {
      Range: (address: string) => address === "C1:D1"
        ? { FormatConditions: { Item: () => null } }
        : { FormatConditions: {
          get Count() { return mine.deleted ? 0 : 1; },
          Item: () => ({ get Formula1() { return mine.formula; }, Delete: () => { mine.deleted = true; } })
        } }
    };
    expect(clearReviewMarksFromCopySheet(sheet as never, ["A1:B1", "C1:D1"], RUN))
      .toEqual({ cleared: 1, unconfirmed: 1 });
    // 别人的规则一条都不碰（清理只按本次完整公式认领）。
    expect(clearReviewMarksFromCopySheet(
      copySheetWith({ Count: 1, Item: () => ({ Type: 2, Formula1: otherRun }) }) as never, ["A1:B1"], RUN))
      .toEqual({ cleared: 0, unconfirmed: 0 });
  });
});

/**
 * A1-F1 收尾：`ownedRules`（标记找回 + 重载清理）与副本清理必须读**同一份**类型判据。
 *
 * 同区域里有用户自建的色阶/数据条时，旧实现把"这条规则没有 Formula1"直接当成
 * "句柄读不出来"抛错，于是：①`Add` 回空但规则其实建好了 → 找不回来，报"未能建立
 * 可独立移除的审阅标记"；②重载清理 → 记录永远留着，每次都提示"未清理"。
 * 反过来也不能松：未知类型、表达式规则读不出公式、空句柄仍然必须阻断。
 */
describe("A1-F1 收尾：同区域非表达式规则不再阻断归属", () => {
  it("同区域有色阶、Add 建出规则却回空：找得回、着得上色，且只 Add 一次", () => {
    const host = installHost({ ranges: ["A1:B1"], addReturnsNull: true });
    host.pushColorScale("A1:B1");
    const marks = markWpsResult(host.application, single("A1:B1"));
    // 找回不是"再建一条"：色阶不该被当成失败去重试 Add。
    expect(host.addCount()).toBe(1);
    expect(host.state("A1:B1").filter(rule => rule.formula === marks.formula)).toHaveLength(1);
    // 找回的那条**真的被着色**了（不是"找到就算完"）。
    const mine = host.state("A1:B1").find(rule => rule.formula === marks.formula)!;
    expect(mine.interior?.pattern).toBe(1);
    expect(mine.borders[7]?.weight).toBe(2);
    // 色阶一个字节都没动。
    const scale = host.state("A1:B1").filter(rule => rule.type === 3);
    expect(scale).toHaveLength(1);
    expect(scale[0].interior).toBeNull();
    expect(host.live("A1:B1")).toContain(null);
  });

  it("同区域有色阶，重载清理只删本次标记；清理成功后再次恢复不重复提示", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    const marks = markWpsResult(host.application, single("A1:B1"));
    host.pushColorScale("A1:B1");
    host.pushForeign("A1:B1", '=N("EB_REVIEW_other")=0', 2);
    expect(host.live("A1:B1")).toEqual([marks.formula, null, '=N("EB_REVIEW_other")=0']);

    // 色阶不是本次的、也没有公式 —— 跳过它，但绝不删它。
    expect(recoverWpsResultMarks(host.application)).toContain("已清理");
    expect(host.live("A1:B1")).toEqual([null, '=N("EB_REVIEW_other")=0']);
    expect(host.state("A1:B1").filter(rule => rule.type === 3)).toHaveLength(1);
    expect(localStorage.getItem(STORAGE_KEY)).toBeNull();
    // 记录已经清掉：再恢复一次什么都不做，也不会再提示"未清理"。
    expect(recoverWpsResultMarks(host.application)).toBeNull();
  });

  it("混入未知类型或读不出公式的规则：仍保留待清理记录，不宣称已清理", () => {
    const host = installHost({ ranges: ["A1:B1"] });
    const marks = markWpsResult(host.application, single("A1:B1"));
    // 类型未知（undefined）、公式也读不出：这一条可能是我们的，不能猜着跳过。
    host.pushForeign("A1:B1", null);
    const note = recoverWpsResultMarks(host.application);
    expect(note).toContain("未清理");
    expect(note).not.toContain("已清理");
    // 本次规则没有被删（读数不可信时一次都不删），记录也还在。
    expect(host.live("A1:B1")).toContain(marks.formula);
    expect(storedRecords()).toHaveLength(1);
    expect(storedRecords()[0].formula).toBe(marks.formula);
  });

  it("表达式规则读不出公式（Type=2）时同样阻断；类型是文本也不算\"可识别\"", () => {
    for (const unreadable of [
      { formula: null as string | null, type: 2 },
      { formula: null as string | null, type: Number("3x") } // NaN：类型读不成数字
    ]) {
      const host = installHost({ ranges: ["A1:B1"] });
      const marks = markWpsResult(host.application, single("A1:B1"));
      host.pushForeign("A1:B1", unreadable.formula, unreadable.type);
      expect(recoverWpsResultMarks(host.application)).toContain("未清理");
      expect(host.live("A1:B1")).toContain(marks.formula);
      expect(storedRecords()).toHaveLength(1);
      localStorage.clear();
    }
  });

  it("多任务、多工作簿：色阶不影响各自的归属，另一簿的记录与规则都不碰", () => {
    const host = installHost({ ranges: ["A1:B1", "C1:D1"] });
    host.pushColorScale("A1:B1");
    const first = markWpsResult(host.application, single("A1:B1"));
    const second = markWpsResult(host.application, single("C1:D1"));
    const foreignRule = '=N("EB_REVIEW_otherbook")=0';
    const otherBook: StoredRecord = { workbookFullName: "C:\\Data\\其他.xlsx",
      formula: foreignRule, ranges: [{ sheet: "数据", range: "A1:B1" }] };
    // 危险形态：另一本簿的规则被复制到本簿，且它的记录也在。
    host.pushForeign("A1:B1", foreignRule, 2);
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ v: 2, records: [...storedRecords(), otherBook] }));

    expect(recoverWpsResultMarks(host.application)).toContain("已清理");
    // 本簿两条各自清掉；色阶、另一任务规则、另一簿规则都留着。
    expect(host.live("A1:B1")).toEqual([null, foreignRule]);
    expect(host.live("C1:D1")).toEqual([]);
    expect(storedRecords()).toEqual([otherBook]);
    expect(first.formula).not.toBe(second.formula);
  });

  /*
   * 两个入口必须给出**同一个**结论：归属性扫描（重载清理）与副本清理读同一份类型判据。
   * 各写一份白名单的后果不是"报错"，而是两条路径对"色阶算不算本次"给出不同答案 ——
   * 一处说清干净了、另一处说没清，或者反过来把用户的色阶删了。所以这里拿同一组类型
   * 分别喂给两个入口，断言分类一致。
   */
  it("同一组规则类型，两个入口的分类一致（色阶跳过、未知类型阻断）", () => {
    const RUN = '=N("EB_REVIEW_samejudge")=0';
    for (const type of [3, 4, 6]) { // 色阶 / 数据条 / 图标集
      const copy = clearReviewMarksFromCopySheet(
        { Range: () => ({ FormatConditions: { Count: 1, Item: () => ({ Type: type, Formula1: null }) } }) } as never,
        ["A1:B1"], RUN);
      expect(copy).toEqual({ cleared: 0, unconfirmed: 0 });

      const host = installHost({ ranges: ["A1:B1"] });
      const marks = markWpsResult(host.application, single("A1:B1"));
      host.pushForeign("A1:B1", null, type);
      expect(recoverWpsResultMarks(host.application)).toContain("已清理");
      expect(host.live("A1:B1")).toEqual([null]);
      expect(host.live("A1:B1")).not.toContain(marks.formula);
      localStorage.clear();
    }
    for (const type of [undefined, 99, Number("3x")]) { // 未知 / 不认识 / 读不成数字
      const copy = clearReviewMarksFromCopySheet(
        { Range: () => ({ FormatConditions: { Count: 1, Item: () => ({ Type: type, Formula1: null }) } }) } as never,
        ["A1:B1"], RUN);
      expect(copy).toEqual({ cleared: 0, unconfirmed: 1 });

      const host = installHost({ ranges: ["A1:B1"] });
      const marks = markWpsResult(host.application, single("A1:B1"));
      host.pushForeign("A1:B1", null, type);
      expect(recoverWpsResultMarks(host.application)).toContain("未清理");
      expect(host.live("A1:B1")).toContain(marks.formula);
      expect(storedRecords()).toHaveLength(1);
      localStorage.clear();
    }
  });
});
