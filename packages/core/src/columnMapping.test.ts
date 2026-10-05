/**
 * 共享夹具驱动的行为锁。
 *
 * 用例来自 `config/column_mapping_cases.json`，Python 侧
 * `server/tests/test_column_mapping_shared_cases.py` 加载同一份文件、
 * 断言同一批期望值。**改这里的期望值必须改夹具，改夹具会让两侧同时生效** ——
 * 单边修改实现会立刻有一侧变红。
 *
 * 本文件不重复 Python 侧已有的单元用例，它只负责"两份实现是否一致"。
 */

import { describe, expect, it } from "vitest";

import cases from "../../../config/column_mapping_cases.json";
import {
  buildColumnMappingActions,
  BlankPolicyNeededError,
  ColumnMappingError,
  NeedsClarificationError,
  OverlappingRangeError,
  ShapeMismatchError,
  UnmappedCategoryError,
  type ColumnMappingInput,
  type ColumnMappingPlanAction
} from "./columnMapping";
import type { CellValue } from "./contracts";

/** 只用于 `instanceof` 判定，构造签名各异，故收窄为 `never[]`。 */
type ErrorConstructor = new (...args: never[]) => ColumnMappingError;

/**
 * 错误种类 → 断言方式。与 Python 侧 `_EXPECTATIONS` 逐项对应。
 *
 * `categoriesUnknown` 与 `needsClarification` 是**两种不同的归因**,
 * 刻意分成两个 key:前者补救动作单一（补类别映射）,后者有两个可能
 * （补分隔符 / 补映射）,不能互相冒充。把它们合成一个错误种类会让
 * 调用方把"要用户选"误当成"补映射表就行"。
 */
const EXPECTATIONS: Record<string, { ctor: ErrorConstructor; fragment: string }> = {
  unmapped: { ctor: UnmappedCategoryError, fragment: "不在映射表内" },
  categoriesUnknown: { ctor: UnmappedCategoryError, fragment: "不在映射表内" },
  needsClarification: {
    ctor: NeedsClarificationError,
    fragment: "无法归类"
  },
  shape: { ctor: ShapeMismatchError, fragment: "形状" },
  overlap: { ctor: OverlappingRangeError, fragment: "重叠" },
  parse: { ctor: ColumnMappingError, fragment: "区域" },
  // 「读到真实空白但策略未定」是**待澄清**，不是报错：实现不替用户挑默认值。
  blankPolicy: { ctor: BlankPolicyNeededError, fragment: "空白" },
  sheetName: { ctor: ColumnMappingError, fragment: "工作表名" }
};

interface SharedCase {
  name: string;
  input: ColumnMappingInput;
  /** 不需要拆分时：单个 `writeValues` 动作的目标矩阵。 */
  expected?: string[][];
  /**
   * 需要跳过（保留原值）或清空时的**完整动作序列**。
   *
   * 只用 `expected` 断言不出"被跳过的位置根本没出现在 range 里" ——
   * 那正是"保留原值被实现成写空值"会悄悄通过的地方。
   */
  expectedActions?: Partial<ColumnMappingPlanAction>[];
  error?: string;
  unmappedValues?: CellValue[];
  /** 已确定归因时,未覆盖的类别（可含拆开的项,不只是整格）。 */
  unmatchedItems?: CellValue[];
  /** 归因不确定时,处理不了的整格原文。 */
  unresolvedCells?: CellValue[];
}

const sharedCases = cases.cases as unknown as SharedCase[];

describe("列映射共享夹具", () => {
  it("夹具本身非空，且每条用例只声明一种期望", () => {
    expect(sharedCases.length).toBeGreaterThan(0);
    for (const item of sharedCases) {
      const declared =
        (item.expected === undefined ? 0 : 1) +
        (item.expectedActions === undefined ? 0 : 1) +
        (item.error === undefined ? 0 : 1);
      expect(
        declared,
        `用例「${item.name}」必须且只能声明 expected / expectedActions / error 之一`
      ).toBe(1);
      if (item.error !== undefined) {
        expect(Object.keys(EXPECTATIONS)).toContain(item.error);
      }
    }
  });

  for (const item of sharedCases) {
    it(item.name, () => {
      if (item.expected !== undefined) {
        const actions = buildColumnMappingActions(item.input);
        expect(actions).toHaveLength(1);
        expect(actions[0].type).toBe("writeValues");
        expect(actions[0].sheet).toBe(item.input.targetSheet);
        expect(actions[0].range).toBe(item.input.targetRange);
        expect(actions[0].values).toEqual(item.expected);
        return;
      }

      if (item.expectedActions !== undefined) {
        const actions = buildColumnMappingActions(item.input);
        expect(actions).toHaveLength(item.expectedActions.length);
        item.expectedActions.forEach((expectedAction, index) => {
          const actual = actions[index] as Record<string, unknown>;
          expect(actual.sheet).toBe(item.input.targetSheet);
          for (const [key, value] of Object.entries(expectedAction)) {
            expect(actual[key], `第 ${index + 1} 个动作的 ${key}`).toEqual(value);
          }
        });
        return;
      }

      const expectation = EXPECTATIONS[item.error!];
      let thrown: unknown;
      try {
        buildColumnMappingActions(item.input);
      } catch (error) {
        thrown = error;
      }
      expect(thrown, "应当抛错但没有抛").toBeInstanceOf(expectation.ctor);
      expect((thrown as Error).message).toContain(expectation.fragment);
      if (item.unmappedValues !== undefined) {
        expect((thrown as UnmappedCategoryError).values).toEqual(
          item.unmappedValues
        );
      }
      if (item.unmatchedItems !== undefined) {
        expect((thrown as UnmappedCategoryError).values).toEqual(
          item.unmatchedItems
        );
      }
      if (item.unresolvedCells !== undefined) {
        expect((thrown as NeedsClarificationError).cells).toEqual(
          item.unresolvedCells
        );
      }
    });
  }
});
