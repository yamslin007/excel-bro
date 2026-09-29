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
  ColumnMappingError,
  OverlappingRangeError,
  ShapeMismatchError,
  UnmappedCategoryError,
  type ColumnMappingInput
} from "./columnMapping";
import type { CellValue } from "./contracts";

/** 只用于 `instanceof` 判定，构造签名各异，故收窄为 `never[]`。 */
type ErrorConstructor = new (...args: never[]) => ColumnMappingError;

/** 错误种类 → 断言方式。与 Python 侧 `_EXPECTATIONS` 逐项对应。 */
const EXPECTATIONS: Record<string, { ctor: ErrorConstructor; fragment: string }> = {
  unmapped: { ctor: UnmappedCategoryError, fragment: "不在映射表内" },
  shape: { ctor: ShapeMismatchError, fragment: "形状" },
  overlap: { ctor: OverlappingRangeError, fragment: "重叠" },
  parse: { ctor: ColumnMappingError, fragment: "区域" },
  blankLabel: { ctor: ColumnMappingError, fragment: "空白标签" },
  sheetName: { ctor: ColumnMappingError, fragment: "工作表名" }
};

interface SharedCase {
  name: string;
  input: ColumnMappingInput;
  expected?: string[][];
  error?: string;
  unmappedValues?: CellValue[];
}

const sharedCases = cases.cases as unknown as SharedCase[];

describe("列映射共享夹具", () => {
  it("夹具本身非空，且每条用例只声明期望值或错误之一", () => {
    expect(sharedCases.length).toBeGreaterThan(0);
    for (const item of sharedCases) {
      expect(
        item.expected === undefined ? 1 : 0,
        `用例「${item.name}」既没有期望值也没有期望错误，或两者都给了`
      ).toBe(item.error === undefined ? 0 : 1);
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
    });
  }
});
