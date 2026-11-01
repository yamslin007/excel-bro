/**
 * 空白策略的语义边界。
 *
 * 钉的是四件以前**表达不出来**的事：
 * 1. 「保留目标原值」—— 旧协议只能写个标签，于是"保留"被实现成"写空值"；
 * 2. 「清空」与「写空串」不是一回事 —— 前者是真空单元格，后者是常量空字符串；
 * 3. 没读到空白时**不该**因为没说空白怎么办就拦住预览；
 * 4. 读到空白且未定策略时是**待澄清**，不是错误，更不是默认标签。
 *
 * 取样值刻意与业务无关（`test_no_business_hardcoding.py` 只扫生产代码，
 * 但保持同样纪律能让夹具长期安全）。
 */

import { describe, expect, it } from "vitest";

import {
  BlankPolicyError,
  blankCellOutcome,
  blankCellPositions,
  blankPolicyFromLegacy,
  isBlankValue,
  parseBlankPolicy
} from "./blankPolicy";
import {
  BlankPolicyNeededError,
  DEFAULT_MAX_MAPPING_ACTIONS,
  BlankSplitLimitError,
  buildColumnMappingActions
} from "./columnMapping";
import type {
  ColumnMappingAction,
  ColumnMappingInput,
  ColumnMappingPlanAction
} from "./columnMapping";

/**
 * 取第一个「写值」动作的值矩阵。
 *
 * `ColumnMappingPlanAction` 是 `writeValues | clearRange` 的**联合**：`clearRange`
 * 只清内容、结构里根本没有 `values`。所以"断言矩阵"这件事必须先按 `type` 收窄 ——
 * 直接写 `actions[0].values` 会在类型层报错，而那正是它该报的：清空动作本来就没有值矩阵，
 * 硬取就是拿一个不存在的字段做断言。
 *
 * 找不到写值动作时抛错而不是返回 `undefined`：调用方都在断言"这一格应该是某个值"，
 * 拿不到动作就说明前置条件（`actions` 里至少有一段写值）已经不成立，测试应当当场炸掉。
 */
function writeMatrix(actions: readonly ColumnMappingPlanAction[]): string[][] {
  const action = actions.find(
    (candidate): candidate is ColumnMappingAction =>
      candidate.type === "writeValues"
  );
  if (!action) throw new Error("预期至少有一个 writeValues 动作，实际一个都没有");
  return action.values;
}

function build(
  sourceValues: (string | number | boolean | null)[][],
  extra: Partial<ColumnMappingInput> = {}
) {
  return buildColumnMappingActions({
    sourceSheet: "表一",
    targetSheet: "表一",
    sourceRange: `A1:A${sourceValues.length}`,
    targetRange: `C1:C${sourceValues.length}`,
    sourceValues,
    mapping: { 甲: "一", 乙: "二" },
    ...extra
  });
}

describe("旧协议兼容", () => {
  it("非空 blankLabel 等价于 kind=value，语义一点没动", () => {
    expect(blankPolicyFromLegacy("未分类")).toEqual({
      kind: "value",
      value: "未分类"
    });
  });

  it("空串/缺失 = 未定（旧协议里它们本来就被拒绝，不存在语义改变）", () => {
    expect(blankPolicyFromLegacy("")).toEqual({ kind: "undecided" });
    expect(blankPolicyFromLegacy(undefined)).toEqual({ kind: "undecided" });
    expect(blankPolicyFromLegacy(null)).toEqual({ kind: "undecided" });
  });

  it("旧规则在客户端仍能算出与以前逐格一致的值", () => {
    const actions = build([["甲"], [null], ["乙"]], { blankLabel: "未分类" });
    expect(actions).toHaveLength(1);
    expect(writeMatrix(actions)).toEqual([["一"], ["未分类"], ["二"]]);
  });
});

describe("没读到空白就不该拦", () => {
  it("源区域一格空白都没有、也没给策略：照常产出计划", () => {
    // 这是"还没读源数据就先问空白怎么办"的正面反例：旧实现在入口处
    // 直接以「空白标签为空」拒绝，逼得调用方在读到值之前先去问用户。
    const actions = build([["甲"], ["乙"], ["甲"]]);
    expect(actions).toHaveLength(1);
    expect(writeMatrix(actions)).toEqual([["一"], ["二"], ["一"]]);
  });

  it("显式空串映射项覆盖到的格子不算需要再决定的空白", () => {
    // mapping 里写了 "" 就是在说"空串按这个处理"，不该再去问用户。
    const actions = build([[""], ["甲"]], {
      mapping: { "": "空串", 甲: "一" }
    });
    expect(writeMatrix(actions)).toEqual([["空串"], ["一"]]);
  });
});

describe("读到空白且策略未定 = 待澄清", () => {
  it("抛的是 BlankPolicyNeededError，且带上真实读到的数量", () => {
    let thrown: unknown;
    try {
      build([["甲"], [null], ["乙"]]);
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(BlankPolicyNeededError);
    expect((thrown as BlankPolicyNeededError).blankCount).toBe(1);
  });

  it("计数只算真空白：数字 0 与 false 不是空白", () => {
    expect(isBlankValue(0)).toBe(false);
    expect(isBlankValue(false)).toBe(false);
    expect(isBlankValue(null)).toBe(true);
    expect(isBlankValue("   ")).toBe(true);
    expect(blankCellPositions([[0], [false], [null], [""]])).toEqual([
      { row: 2, column: 0 },
      { row: 3, column: 0 }
    ]);
  });
});

describe("四种意图互不冒充", () => {
  it("keep：被跳过的行**根本不出现在任何 range 里**", () => {
    const actions = build([["甲"], [null], ["乙"]], {
      blankPolicy: { kind: "keep" }
    });
    expect(actions.map((action) => action.range)).toEqual(["C1:C1", "C3:C3"]);
    // 反向锚点：若实现把"跳过"降级成"写个空值"，这里会变成一整块 C1:C3。
    expect(actions.some((action) => action.range === "C1:C3")).toBe(false);
  });

  it("clear：走 clearRange:contents，不是写空串", () => {
    const actions = build([["甲"], [null], ["乙"]], {
      blankPolicy: { kind: "clear" }
    });
    expect(actions[1]).toEqual({
      type: "clearRange",
      sheet: "表一",
      range: "C2:C2",
      applyTo: "contents"
    });
  });

  it("value 写指定值：整块区域照写", () => {
    const actions = build([["甲"], [null], ["乙"]], {
      blankPolicy: { kind: "value", value: "未分类" }
    });
    expect(actions).toHaveLength(1);
    expect(writeMatrix(actions)).toEqual([["一"], ["未分类"], ["二"]]);
  });

  it("value 可以是空串：那是显式写空字符串，与 clear、keep 都不同", () => {
    const actions = build([["甲"], [null], ["乙"]], {
      blankPolicy: { kind: "value", value: "" }
    });
    expect(actions).toHaveLength(1);
    expect(writeMatrix(actions)).toEqual([["一"], [""], ["二"]]);
  });

  it("连续多个空白也要合成一段，不逐格拆", () => {
    const actions = build([["甲"], [null], [null], ["乙"]], {
      blankPolicy: { kind: "clear" }
    });
    expect(actions.map((action) => action.range)).toEqual([
      "C1:C1",
      "C2:C3",
      "C4:C4"
    ]);
  });
});

describe("策略解析", () => {
  it("认不出来的 kind 抛错，不静默折成未定", () => {
    expect(() => parseBlankPolicy({ kind: "maybe" })).toThrow(BlankPolicyError);
    expect(() => parseBlankPolicy({ kind: "value" })).toThrow(BlankPolicyError);
  });

  it("四种形态都能解析，value 分支允许空串", () => {
    expect(parseBlankPolicy({ kind: "keep" })).toEqual({ kind: "keep" });
    expect(parseBlankPolicy({ kind: "clear" })).toEqual({ kind: "clear" });
    expect(parseBlankPolicy({ kind: "undecided" })).toEqual({
      kind: "undecided"
    });
    expect(parseBlankPolicy({ kind: "value", value: "" })).toEqual({
      kind: "value",
      value: ""
    });
  });

  it("未定的策略不允许落地：调用方必须先按真实数据拦掉", () => {
    expect(() => blankCellOutcome({ kind: "undecided" })).toThrow(
      BlankPolicyError
    );
  });
});

describe("拆分必须有上限", () => {
  it("零散空白超过上限时明说，不降级成写空值", () => {
    const values: (string | null)[][] = [];
    for (let index = 0; index < 40; index += 1) {
      values.push([index % 2 === 0 ? "甲" : null]);
    }
    let thrown: unknown;
    try {
      build(values, { blankPolicy: { kind: "keep" }, maxSplitActions: 4 });
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(BlankSplitLimitError);
    expect((thrown as BlankSplitLimitError).maxActions).toBe(4);
  });

  it("默认上限是有限值，不是不限", () => {
    expect(DEFAULT_MAX_MAPPING_ACTIONS).toBeGreaterThan(0);
  });
});
