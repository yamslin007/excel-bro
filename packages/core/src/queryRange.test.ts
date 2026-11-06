import { describe, expect, it } from "vitest";
import {
  intersectQueryRect,
  parseQueryRange,
  queryRangeIssue,
  queryRectAddress,
  sliceForQueryRange
} from "./queryRange";

/**
 * `query_table` 的显式只读区域。
 *
 * 现场（2026-11-01）：用户说「只读取来源!A1:B5」，本端只能整表读或改用会写入的
 * `computeValues`。这里钉住写法判定与裁剪口径 —— 服务端 `models.py` 的
 * `validate_query_range` 按同一套规则镜像（用例表也保持一致）。
 */
describe("区域写法判定", () => {
  it("解析单个矩形", () => {
    expect(parseQueryRange("A1:B5")).toEqual({
      sheet: null,
      startRow: 1,
      startColumn: 1,
      rowCount: 5,
      columnCount: 2
    });
    expect(parseQueryRange("$B$2:$D$4")).toEqual({
      sheet: null,
      startRow: 2,
      startColumn: 2,
      rowCount: 3,
      columnCount: 3
    });
    // 单格也是合法矩形。
    expect(parseQueryRange("C3")).toMatchObject({
      startRow: 3,
      startColumn: 3,
      rowCount: 1,
      columnCount: 1
    });
  });

  it("带工作表名（可加单引号）", () => {
    expect(parseQueryRange("来源!A1:B5")?.sheet).toBe("来源");
    expect(parseQueryRange("'我的 表'!B2")?.sheet).toBe("我的 表");
  });

  it("拒绝会破坏「区域外不参与计算」的写法", () => {
    for (const value of [
      "",
      "   ",
      "A:A",
      "1:5",
      "A1:",
      "A1:B5:C7",
      "名称区域",
      "A0",
      "A1:B0",
      "A1:B5!",
      "a!b!A1",
      "A1:B5:B9"
    ]) {
      expect(parseQueryRange(value), value).toBeNull();
      expect(queryRangeIssue(value), value).not.toBeNull();
    }
  });

  it("越界与超长也要拒", () => {
    expect(parseQueryRange("A1048577")).toBeNull();
    // 16385 列 = XFE。
    expect(parseQueryRange("A1:XFE1")).toBeNull();
    // 超过 40 个字符（这里靠工作表名撑长）。
    expect(parseQueryRange(`${"x".repeat(40)}!A1`)).toBeNull();
    // 首尾空白只是被去掉，不算错。
    expect(parseQueryRange("  A1:B5  ")).toMatchObject({ rowCount: 5 });
  });

  it("工作表名不能带分隔符或通配符", () => {
    for (const value of ["a/b!A1", "a?b!A1", "a[b]!A1", "a:b!A1"]) {
      expect(parseQueryRange(value), value).toBeNull();
    }
  });

  it("合法写法的说明为空", () => {
    expect(queryRangeIssue("来源!A1:B5")).toBeNull();
  });
});

describe("区域交集与裁剪", () => {
  const used = { startRow: 1, startColumn: 1, rowCount: 6, columnCount: 2 };

  it("相交给出交集，不相交给 null", () => {
    expect(
      intersectQueryRect(
        { startRow: 1, startColumn: 1, rowCount: 5, columnCount: 2 },
        used
      )
    ).toEqual({ startRow: 1, startColumn: 1, rowCount: 5, columnCount: 2 });
    expect(
      intersectQueryRect(
        { startRow: 10, startColumn: 1, rowCount: 2, columnCount: 2 },
        used
      )
    ).toBeNull();
  });

  it("裁剪结果正好是区域形状，原点之外的格补 null", () => {
    const values = [
      ["甲", 1],
      ["乙", 2],
      ["丙", 3]
    ];
    // 矩阵从 B2 开始：区域起点比原点早，前面几格只能是空。
    // 注意列也会整体错位：矩阵的第 0 列是 B 列，所以区域的 A 列全空。
    expect(sliceForQueryRange(values, { row: 2, column: 2 }, {
      startRow: 1,
      startColumn: 1,
      rowCount: 3,
      columnCount: 2
    })).toEqual([
      [null, null],
      [null, "甲"],
      [null, "乙"]
    ]);
    // 起点对齐时原样返回同样形状。
    expect(sliceForQueryRange(values, { row: 1, column: 1 }, {
      startRow: 2,
      startColumn: 1,
      rowCount: 2,
      columnCount: 2
    })).toEqual([
      ["乙", 2],
      ["丙", 3]
    ]);
  });

  it("地址回写可用于错误信息", () => {
    expect(
      queryRectAddress({
        startRow: 1,
        startColumn: 1,
        rowCount: 5,
        columnCount: 2
      })
    ).toBe("A1:B5");
    expect(
      queryRectAddress({
        startRow: 3,
        startColumn: 3,
        rowCount: 1,
        columnCount: 1
      })
    ).toBe("C3");
  });
});
