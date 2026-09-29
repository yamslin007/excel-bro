/**
 * 切片逻辑的跨语言行为锁（TypeScript 侧，阶段 4 §24）。
 *
 * 用例来自 `config/mapping_slice_cases.json`，Python 侧
 * `server/tests/test_mapping_slice_shared_cases.py` 加载同一份文件，
 * 对 `_snapshot_region` 断言同一批期望值。**单边修改任一实现都会让一侧变红。**
 *
 * 为什么需要这一份：`column_mapping_cases.json` 锁的是纯函数
 * `buildColumnMappingActions`。两端各自的「从已读数据里切出源区域」不在它的
 * 覆盖范围内，此前只有各自语言的单元测试，一致性靠人工对读源码 —— 而对读
 * 得出的"两侧策略一致"结论是**错的**：本端不拦右侧列溢出、服务端不拦左侧列
 * 早于原点。本文件的存在就是为了不再靠人眼。
 *
 * 错误只按**种类**共享，消息文案两侧各自声明：服务端说「快照/已捕获」，
 * 客户端说「已读入」，同一件事在两侧的措辞本就不同，强行统一没有收益。
 */

import { describe, expect, it } from "vitest";

import cases from "../../../config/mapping_slice_cases.json";
import type { CellValue } from "../../excel-addin/src/contracts";
import { sliceMappingSource, type MappingSourceSheet } from "./wpsColumnMappingPlan";

/** 错误种类 → 本端消息片段。与 Python 侧 `_EXPECTATIONS` 逐项对应。 */
const EXPECTATIONS: Record<string, string> = {
  rowBeforeOrigin: "行，早于",
  rowBeyondCaptured: "行，超出已读取范围",
  columnBeforeOrigin: "列，早于",
  columnBeyondCaptured: "列，超出已读取范围",
  parse: "无法解析源区域",
  reversed: "无法解析源区域"
};

interface SharedCase {
  name: string;
  sheet: {
    startRow: number;
    startColumn: number;
    columnCount: number;
    values: CellValue[][];
    truncated: boolean;
  };
  range: string;
  expected?: CellValue[][];
  error?: string;
}

const sharedCases = (cases as { cases: SharedCase[] }).cases;

function toSheet(item: SharedCase): MappingSourceSheet {
  return {
    name: "共享用例表",
    startRow: item.sheet.startRow,
    startColumn: item.sheet.startColumn,
    columnCount: item.sheet.columnCount,
    values: item.sheet.values
  };
}

describe("切片共享夹具", () => {
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
        expect(sliceMappingSource(toSheet(item), item.range)).toEqual(item.expected);
        return;
      }
      expect(() => sliceMappingSource(toSheet(item), item.range)).toThrow(
        EXPECTATIONS[item.error!]
      );
    });
  }
});
