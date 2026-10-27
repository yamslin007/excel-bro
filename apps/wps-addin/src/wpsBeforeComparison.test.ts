/**
 * 执行前副本的**真实重建**（`stageWpsBeforeComparison`）。
 *
 * 为什么要单独测：面板层把 `stageWpsBeforeComparison` 替身掉了，那里只能证明
 * "调用过、没再执行"。真正的问题是它**重建出来的内容对不对** —— 执行已经发生，
 * 复制当前表只能拿到"执行后状态"；若不再写回执行前快照，用户点双视图看到的
 * 就是"执行后的表 vs 执行后的表"，对比毫无意义。
 *
 * 这里用最小宿主替身：值、公式、加粗、行高四项，够覆盖"值 / 公式 / 混合格式 /
 * 本次涉及的行尺寸"四类，且不依赖面板。
 */
import { describe, expect, it } from "vitest";
import type { AnalysisPlan } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorkbook, WpsWorksheet } from "./wpsHost";
import { prepareWpsPlan } from "./wpsPlanExecutor";
import { captureWpsBeforeResult, stageWpsBeforeComparison } from "./wpsNativeComparison";

type Cell = { value: string | null; formula: string | null; bold: boolean };

/** 条件格式规则替身：只够"标记带过去又被清掉"这一件事。 */
type CondRule = { formula: string; deleted: boolean };

/** 把一份规则列表包成 WPS 的 FormatConditions 形状。 */
function conditionsOf(rules: CondRule[]) {
  return {
    get Count() { return rules.filter((rule) => !rule.deleted).length; },
    Item(index: number) {
      const live = rules.filter((rule) => !rule.deleted)[index - 1];
      return live
        ? { get Formula1() { return live.formula; }, Delete: () => { live.deleted = true; } }
        : undefined;
    }
  };
}

function makeHost(options: {
  values: string[][];
  formulas?: Record<string, string>;
  bold?: boolean;
  rowHeight?: number;
  /** 原表上已有的条件格式规则（模拟"标记已打上" + 用户规则 + 另一轮审阅规则）。 */
  conditions?: CondRule[];
  /** 副本的条件格式集合读不出来（宿主偶发）：用来钉"无法确认清理结果时要如实说"。 */
  copyConditionsUnreadable?: boolean;
  /** 直接替换副本上的集合，用来钉 `Count` 缺失 / 空句柄 / 缺公式 / 非表达式类型这几种读数形态。 */
  copyConditionsOverride?: unknown;
}) {
  const rows = options.values.length;
  const columns = options.values[0]?.length ?? 1;
  // 原表上的条件格式规则。默认一条"本次审阅"规则（公式由调用方给出时按给出的算）。
  const originalConditions: CondRule[] = options.conditions ??
    [{ formula: '=N("EB_REVIEW_deadbeef")=0', deleted: false }];
  // Copy() 会把原表规则原样带进副本。
  let copyConditions: CondRule[] = [];
  const cells: Cell[][] = options.values.map((row, r) =>
    row.map((value, c) => ({
      value,
      // 公式格：显示的"值"由调用方给定，`Formula` 里保存真实公式文本。
      formula: options.formulas?.[`${r + 1},${c + 1}`] ?? value,
      bold: options.bold === true
    }))
  );
  let rowHeight = options.rowHeight ?? 20;

  const cellRange = (r: number, c: number): WpsRange => {
    const cell = cells[r - 1][c - 1];
    return {
      get Value2() { return cell.value; },
      set Value2(next: string | null) { cell.value = next; cell.formula = next ?? null; },
      get Formula() { return cell.formula ?? ""; },
      set Formula(next: string) { cell.formula = next; cell.value = next.startsWith("=") ? "(公式)" : next; },
      MergeCells: false,
      RowHeight: rowHeight,
      get Font() {
        return {
          get Bold() { return cell.bold; },
          set Bold(next: boolean) { cell.bold = next; },
          Italic: false, Name: "宋体", Size: 11, Color: 0
        };
      },
      get Interior() { return { Color: undefined, Pattern: undefined }; },
      get NumberFormat() { return "General"; },
      get HorizontalAlignment() { return -4108; },
      get VerticalAlignment() { return -4108; },
      get WrapText() { return false; },
      get IndentLevel() { return 0; },
      get ShrinkToFit() { return false; },
      get Orientation() { return 0; },
      get Row() { return r; },
      get Column() { return c; }
    } as unknown as WpsRange;
  };

  const rect = (address: string) => {
    const match = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(address.replace(/\s/g, ""));
    if (!match) throw new Error(`替身无法解析区域「${address}」`);
    const toIndex = (letters: string) => letters.split("").reduce((sum, ch) => sum * 26 + (ch.charCodeAt(0) - 64), 0);
    return {
      startRow: Number(match[2]), startColumn: toIndex(match[1]),
      endRow: match[4] ? Number(match[4]) : Number(match[2]),
      endColumn: match[3] ? toIndex(match[3]) : toIndex(match[1])
    };
  };

  const makeRange = (address: string): WpsRange => {
    const box = rect(address);
    return {
      get Value2() {
        if (box.startRow === box.endRow && box.startColumn === box.endColumn) return cells[box.startRow - 1][box.startColumn - 1].value;
        return cells.slice(box.startRow - 1, box.endRow).map((row) => row.slice(box.startColumn - 1, box.endColumn).map((cell) => cell.value));
      },
      set Value2(_next: unknown) { throw new Error("替身只支持逐格写回"); },
      get Formula() {
        if (box.startRow === box.endRow && box.startColumn === box.endColumn) return cells[box.startRow - 1][box.startColumn - 1].formula;
        return cells.slice(box.startRow - 1, box.endRow).map((row) => row.slice(box.startColumn - 1, box.endColumn).map((cell) => cell.formula));
      },
      set Formula(_next: unknown) { throw new Error("替身只支持逐格写回"); },
      MergeCells: false,
      RowHeight: rowHeight,
      get FormatConditions() { return conditionsOf(originalConditions); },
      get Font() {
        return {
          get Bold() { return cells[box.startRow - 1][box.startColumn - 1].bold; },
          set Bold(next: boolean) { cells[box.startRow - 1][box.startColumn - 1].bold = next; },
          Italic: false, Name: "宋体", Size: 11, Color: 0
        };
      },
      get Interior() { return { Color: undefined, Pattern: undefined }; },
      get NumberFormat() { return "General"; },
      get HorizontalAlignment() { return -4108; },
      get VerticalAlignment() { return -4108; },
      get WrapText() { return false; },
      get IndentLevel() { return 0; },
      get ShrinkToFit() { return false; },
      get Orientation() { return 0; },
      Item: (r: number, c: number) => cellRange(box.startRow + r - 1, box.startColumn + c - 1)
    } as unknown as WpsRange;
  };

  const makeSheet = (name: string, data: Cell[][]): WpsWorksheet => ({
    Name: name,
    ProtectContents: false,
    Range: (address: string) => makeRange(address),
    get UsedRange() { return makeRange(`A1:${String.fromCharCode(64 + columns)}${rows}`); },
    Copy() {
      // 真实宿主：复制**当前**（= 执行后）状态。重建必须把执行前快照写回去。
      copyCells = data.map((row) => row.map((cell) => ({ ...cell })));
      copyRowHeight = rowHeight;
      // 条件格式规则也随 Copy() 一起带过去 —— 这正是"执行前副本也高亮"的来源。
      copyConditions = originalConditions.map((rule) => ({ ...rule }));
      active = copyWorkbook;
    }
  } as unknown as WpsWorksheet);

  const originalSheet = makeSheet("数据", cells);
  let copyCells: Cell[][] = [];
  let copyRowHeight = rowHeight;
  let active: WpsWorkbook;

  const copySheet = {
    Name: "数据",
    ProtectContents: false,
    Range: (address: string) => {
      const box = rect(address);
      const single = box.startRow === box.endRow && box.startColumn === box.endColumn;
      const get = (r: number, c: number) => copyCells[r - 1][c - 1];
      return {
        get Value2() {
          if (single) return get(box.startRow, box.startColumn).value;
          return copyCells.slice(box.startRow - 1, box.endRow).map((row) => row.slice(box.startColumn - 1, box.endColumn).map((cell) => cell.value));
        },
        set Value2(next: unknown) {
          // 单格区域本身就是单元格：宿主允许直接赋值；多格走逐格（下面 Item）。
          if (!single) throw new Error("替身只支持逐格写回");
          const cell = get(box.startRow, box.startColumn);
          cell.value = next as string | null;
          cell.formula = next as string | null;
        },
        get Formula() {
          if (single) return get(box.startRow, box.startColumn).formula;
          return copyCells.slice(box.startRow - 1, box.endRow).map((row) => row.slice(box.startColumn - 1, box.endColumn).map((cell) => cell.formula));
        },
        set Formula(next: unknown) {
          if (!single) throw new Error("替身只支持逐格写回");
          const cell = get(box.startRow, box.startColumn);
          cell.formula = String(next);
          cell.value = String(next).startsWith("=") ? "(公式)" : String(next);
        },
        MergeCells: false,
        get RowHeight() { return copyRowHeight; },
        set RowHeight(next: number) { copyRowHeight = next; },
        get FormatConditions() {
          if (options.copyConditionsOverride !== undefined) return options.copyConditionsOverride;
          return options.copyConditionsUnreadable ? undefined : conditionsOf(copyConditions);
        },
        get Font() {
          const cell = get(box.startRow, box.startColumn);
          return {
            get Bold() { return cell.bold; },
            set Bold(next: boolean) { cell.bold = next; },
            Italic: false, Name: "宋体", Size: 11, Color: 0
          };
        },
        get Interior() { return { Color: undefined, Pattern: undefined }; },
        get NumberFormat() { return "General"; },
        HorizontalAlignment: -4108,
        VerticalAlignment: -4108,
        WrapText: false,
        IndentLevel: 0,
        ShrinkToFit: false,
        Orientation: 0,
        Item: (r: number, c: number) => {
          const cell = get(box.startRow + r - 1, box.startColumn + c - 1);
          return {
            get Value2() { return cell.value; },
            set Value2(next: string | null) { cell.value = next; cell.formula = next ?? null; },
            get Formula() { return cell.formula ?? ""; },
            set Formula(next: string) { cell.formula = next; cell.value = next.startsWith("=") ? "(公式)" : next; },
            MergeCells: false,
            get RowHeight() { return copyRowHeight; },
            set RowHeight(next: number) { copyRowHeight = next; },
            get Font() {
              return {
                get Bold() { return cell.bold; },
                set Bold(next: boolean) { cell.bold = next; },
                Italic: false, Name: "宋体", Size: 11, Color: 0
              };
            },
            get Interior() { return { Color: undefined, Pattern: undefined }; },
            get NumberFormat() { return "General"; },
            HorizontalAlignment: -4108,
            VerticalAlignment: -4108,
            WrapText: false,
            IndentLevel: 0,
            ShrinkToFit: false,
            Orientation: 0
          } as unknown as WpsRange;
        }
      } as unknown as WpsRange;
    }
  };

  const original: WpsWorkbook = {
    Name: "测试.xlsx", FullName: "C:\\Data\\测试.xlsx", Path: "C:\\Data", ReadOnly: false,
    Worksheets: { Count: 1, Item: () => originalSheet },
    ActiveSheet: originalSheet,
    Activate() { active = original; }
  } as unknown as WpsWorkbook;
  const copyWorkbook: WpsWorkbook = {
    Name: "Book1", FullName: "", Path: "", ReadOnly: false,
    Worksheets: { Count: 1, Item: () => copySheet },
    ActiveSheet: copySheet as unknown as WpsWorksheet,
    Windows: { Count: 1, Item: () => ({ Caption: "Book1" }) },
    Activate() { active = copyWorkbook; },
    Close: () => true
  } as unknown as WpsWorkbook;
  active = original;

  const application = {
    Version: "12.0",
    get ActiveWorkbook() { return active; },
    get ActiveSheet() { return active.ActiveSheet!; },
    Windows: { Count: 2, Item: () => ({ Caption: "Book1" }), CompareSideBySideWith: () => true, SyncScrollingSideBySide: false },
    PluginStorage: { getItem: () => undefined, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as unknown as WpsApplication;

  return {
    application,
    copyCells: () => copyCells,
    copyRowHeight: () => copyRowHeight,
    copyConditions: () => copyConditions,
    originalCells: () => cells
  };
}

const sheetName = "数据";
function plan(title: string, actions: AnalysisPlan["actions"]): AnalysisPlan {
  return { id: "plan-before", title, summary: title, assumptions: [], warnings: [], actions } as AnalysisPlan;
}

describe("执行前副本按快照重建", () => {
  it("值是执行前的：副本按执行前快照重建，而不是沿用执行后的当前表", () => {
    // 执行前：A1:B1 是「旧值A/B」。
    const host = makeHost({ values: [["旧值A", "旧值B"]] });
    const prepared = prepareWpsPlan(host.application, [1], plan("写值", [
      { type: "writeValues", sheet: sheetName, range: "A1:B1", values: [["新值A", "新值B"]] }
    ]), []);
    const before = captureWpsBeforeResult(host.application, prepared);
    expect(before.values[0].values).toEqual([["旧值A", "旧值B"]]);

    // 模拟执行已经发生：当前表变成执行后的样子。
    host.originalCells()[0][0].value = "新值A";
    host.originalCells()[0][1].value = "新值B";

    stageWpsBeforeComparison(host.application, before);
    // 副本必须是**执行前**的两格；若沿用复制到的当前状态，这里会是新值 → 变红。
    expect(host.copyCells()[0].map((cell) => cell.value)).toEqual(["旧值A", "旧值B"]);
    // 原表一格不动（重建只写副本）。
    expect(host.originalCells()[0].map((cell) => cell.value)).toEqual(["新值A", "新值B"]);
  });

  it("公式使用执行前显示值，不恢复可重算的公式", () => {
    const host = makeHost({ values: [["常量的显示值", "公式的显示值"]], formulas: { "1,2": "=B9*2" } });
    const prepared = prepareWpsPlan(host.application, [1], plan("写公式", [
      { type: "writeFormulas", sheet: sheetName, range: "A1:B1", formulas: [["=\"常量\"", "=B9*2"]] }
    ]), []);
    const before = captureWpsBeforeResult(host.application, prepared);
    stageWpsBeforeComparison(host.application, before);
    const copy = host.copyCells()[0];
    // 常量和公式都使用执行前的显示值，不在副本中重新计算。
    expect(copy[0].formula).toBe("常量的显示值");
    expect(copy[1].formula).toBe("公式的显示值");
    expect(copy[1].value).toBe("公式的显示值");
  });

  it("清空动作的执行前副本恢复旧值，不再次清空", () => {
    const host = makeHost({ values: [["执行前内容"]] });
    const prepared = prepareWpsPlan(host.application, [1], plan("清空", [
      { type: "clearRange", sheet: sheetName, range: "A1", applyTo: "contents" }
    ]), []);
    const before = captureWpsBeforeResult(host.application, prepared);
    host.originalCells()[0][0].value = "";
    stageWpsBeforeComparison(host.application, before);
    expect(host.copyCells()[0][0].value).toBe("执行前内容");
    expect(host.originalCells()[0][0].value).toBe("");
  });

  it("格式动作：加粗与行高按执行前恢复，而不是沿用执行后的样子", () => {
    const host = makeHost({ values: [["甲"]], bold: false, rowHeight: 20 });
    const prepared = prepareWpsPlan(host.application, [1], plan("加粗", [
      { type: "setFont", sheet: sheetName, range: "A1", bold: true }
    ]), []);
    const before = captureWpsBeforeResult(host.application, prepared);
    expect(before.formats[0]?.captured).toBe(true);

    // 模拟执行后：加粗 + 行高变了（副本会连这些一起复制过去）。
    host.originalCells()[0][0].bold = true;
    stageWpsBeforeComparison(host.application, before);
    expect(host.copyCells()[0][0].bold).toBe(false);
    expect(host.copyRowHeight()).toBe(20);
  });

  it("列宽动作在计划层就被拒绝（未量化验证），不会走到「看起来对」的执行前副本", () => {
    const host = makeHost({ values: [["甲"]] });
    /*
     * 列宽是列级属性、且未纳入格式快照；`prepareWpsPlan` 目前直接拒绝该动作。
     * 这条用例锁的是**分层**：不能既允许执行、又在双视图里拿执行后的宽度冒充执行前。
     * `stageWpsBeforeComparison` 里另有一道防御（本次动作含列宽则拒绝打开对比）。
     */
    expect(() => prepareWpsPlan(host.application, [1], plan("列宽", [
      { type: "resizeRange", sheet: sheetName, range: "A1", columnWidth: 30 }
    ]), [])).toThrow(/columnWidth/);
    expect(host.originalCells()[0][0].value).toBe("甲");
  });

  it("执行前格式采集不到时，只在打开双视图这一步明确拒绝，不影响执行结论", () => {
    const host = makeHost({ values: [["甲"]] });
    const prepared = prepareWpsPlan(host.application, [1], plan("加粗", [
      { type: "setFont", sheet: sheetName, range: "A1", bold: true }
    ]), []);
    const before = captureWpsBeforeResult(host.application, prepared);
    // 模拟"格式没采到"：把采集结果抹掉（等价于宿主不提供逐格格式）。
    const degraded = { ...before, formats: [undefined] };
    expect(() => stageWpsBeforeComparison(host.application, degraded)).toThrow(/执行前格式未完整采集/);
  });
});
describe("执行前副本不带入本次审阅标记", () => {
  const THIS_RUN = '=N("EB_REVIEW_thisrun_abc123")=0';
  const OTHER_RUN = '=N("EB_REVIEW_otherrun_xyz789")=0';
  const USER_RULE = "=A1>0";

  it("Copy() 带过来的本次标记被清掉；另一轮审阅与用户规则保留", () => {
    const host = makeHost({
      values: [["旧值A"]],
      // 原表上有：本次审阅规则 + 另一轮审阅规则（同前缀、不同 nonce）+ 用户规则。
      conditions: [
        { formula: THIS_RUN, deleted: false },
        { formula: OTHER_RUN, deleted: false },
        { formula: USER_RULE, deleted: false }
      ]
    });
    const prepared = prepareWpsPlan(host.application, [1], plan("写值", [
      { type: "writeValues", sheet: sheetName, range: "A1", values: [["新值A"]] }
    ]), []);
    const before = captureWpsBeforeResult(host.application, prepared);
    host.originalCells()[0][0].value = "新值A";
    // 只给本次完整公式：另一轮审阅规则即使同 EB_REVIEW_ 前缀也不删。
    const staged = stageWpsBeforeComparison(host.application, before, THIS_RUN);
    // 清了几条必须**报告出来**：面板据此决定要不要提示"副本可能仍然同色"。
    expect(staged.reviewMarks).toEqual({ cleared: 1, unconfirmed: 0 });
    expect(host.copyConditions().filter(rule => !rule.deleted).map(rule => rule.formula))
      .toEqual([OTHER_RUN, USER_RULE]);
  });

  it("副本集合读不出来：报告 unconfirmed，不假装已经清干净", () => {
    const host = makeHost({
      values: [["旧值A"]],
      conditions: [{ formula: THIS_RUN, deleted: false }],
      copyConditionsUnreadable: true
    });
    const prepared = prepareWpsPlan(host.application, [1], plan("写值", [
      { type: "writeValues", sheet: sheetName, range: "A1", values: [["新值A"]] }
    ]), []);
    const before = captureWpsBeforeResult(host.application, prepared);
    const staged = stageWpsBeforeComparison(host.application, before, THIS_RUN);
    expect(staged.reviewMarks).toEqual({ cleared: 0, unconfirmed: 1 });
    // 读不出来就一个都不删（宁可副本高亮，也不按区域批量删）。
    expect(host.copyConditions().filter(rule => !rule.deleted)).toHaveLength(1);
  });

  /** 跑一遍真实 staging，只换掉副本上的条件格式读数形态。 */
  function stageWithCopyConditions(conditions: unknown) {
    const host = makeHost({
      values: [["旧值A"]],
      conditions: [{ formula: THIS_RUN, deleted: false }],
      copyConditionsOverride: conditions
    });
    const prepared = prepareWpsPlan(host.application, [1], plan("写值", [
      { type: "writeValues", sheet: sheetName, range: "A1", values: [["新值A"]] }
    ]), []);
    const before = captureWpsBeforeResult(host.application, prepared);
    return stageWpsBeforeComparison(host.application, before, THIS_RUN);
  }

  it("副本读数不完整（缺条数 / 空句柄 / 缺公式 / 未知类型）：一律 unconfirmed", () => {
    const deleted: string[] = [];
    const ours = {
      Type: 2,
      get Formula1() { return THIS_RUN; },
      Delete: () => { deleted.push("ours"); }
    };
    // 缺 `Count`：旧实现当成 0 条 ⇒ 副本还黄着，却报"已清理"。
    expect(stageWithCopyConditions({ Item: () => ours }).reviewMarks).toEqual({ cleared: 0, unconfirmed: 1 });
    // 表达式规则却读不出公式。
    expect(stageWithCopyConditions({ Count: 1, Item: () => ({ Type: 2, Formula1: null }) }).reviewMarks)
      .toEqual({ cleared: 0, unconfirmed: 1 });
    // 类型不认识、公式也读不出。
    expect(stageWithCopyConditions({ Count: 1, Item: () => ({ Type: 99, Formula1: null }) }).reviewMarks)
      .toEqual({ cleared: 0, unconfirmed: 1 });
    // 以上三种都是"读数不可信"：一次都不许删（宁可副本多一圈高亮）。
    expect(deleted).toEqual([]);
    // 空句柄：读得出的本次规则照常删，但整个区域仍报"没核实干净"。
    expect(stageWithCopyConditions({ Count: 2, Item: (index: number) => (index === 1 ? null : ours) }).reviewMarks)
      .toEqual({ cleared: 1, unconfirmed: 1 });
  });

  it("副本上的色阶规则不算\"无法核实\"：本次标记照常清掉，不产生假警报", () => {
    const staged = stageWithCopyConditions({
      Count: 2,
      Item: (index: number) => (index === 1
        ? { Type: 3 } // xlColorScale：结构上就没有 Formula1
        : { Type: 2, Formula1: THIS_RUN, Delete: () => {} })
    });
    expect(staged.reviewMarks).toEqual({ cleared: 1, unconfirmed: 0 });
  });

  it("没给本次公式就不清理：副本多一圈高亮，也好过误删别人的规则", () => {
    const host = makeHost({
      values: [["旧值A"]],
      conditions: [
        { formula: THIS_RUN, deleted: false },
        { formula: USER_RULE, deleted: false }
      ]
    });
    const prepared = prepareWpsPlan(host.application, [1], plan("写值", [
      { type: "writeValues", sheet: sheetName, range: "A1", values: [["新值A"]] }
    ]), []);
    const before = captureWpsBeforeResult(host.application, prepared);
    stageWpsBeforeComparison(host.application, before);
    // 什么都没清 —— 宁可副本也高亮，也不按前缀误删。
    expect(host.copyConditions().filter(rule => !rule.deleted)).toHaveLength(2);
  });
});
