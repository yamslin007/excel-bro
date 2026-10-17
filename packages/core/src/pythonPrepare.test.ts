/**
 * `prepare_python_data` 区域切片与值归一的**宿主无关**契约。
 *
 * 这里锁的是两端共用的那份实现；两端各自的宿主适配（Excel `getRange` 绝对坐标、
 * WPS 已用区域相对矩阵）由各自应用目录下的 `*PrepareShared.test.ts` 用
 * `config/python_prepare_cases.json` 跑同一批夹具。
 */

import { describe, expect, it } from "vitest";

import cases from "../../../config/python_prepare_cases.json";
import {
  assertPythonPrepareRectWithin,
  normalizeCellValue,
  parsePythonPrepareRect,
  PYTHON_PREPARE_REFUSAL_MESSAGE,
  PythonPrepareRegionRefused,
  slicePythonPrepareRegion
} from "./pythonPrepare";

const fixture = cases as {
  sheet: { startRow: number; startColumn: number; matrix: (string | number)[][] };
  slice: {
    name: string;
    range: string;
    expected?: (string | number)[][];
    refuse?: boolean;
  }[];
  values: {
    name: string;
    value: string | number | boolean;
    text: string;
    numberFormat: string;
    expected: string | number | boolean;
  }[];
  ragged: {
    name: string;
    matrix: (string | number)[][];
    rowCount: number;
    columnCount: number;
    range: string;
    expected?: (string | number)[][];
    refuse?: boolean;
  }[];
};

describe("区域解析", () => {
  it("A1 与 A1:B2 都能解析，且是 0 基绝对坐标", () => {
    expect(parsePythonPrepareRect("A1")).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 });
    expect(parsePythonPrepareRect("B3:E6")).toEqual({ r1: 2, c1: 1, r2: 5, c2: 4 });
  });

  it("反向区域、含 $、空串一律拒绝", () => {
    for (const bad of ["B2:A1", "E6:B3", "$B$3", "", "  ", "ZZZZ1"]) {
      expect(parsePythonPrepareRect(bad), bad).toBeNull();
    }
  });
});

describe("共享切片夹具", () => {
  it("夹具非空，每条用例只声明期望值或拒绝之一", () => {
    expect(fixture.slice.length).toBeGreaterThan(0);
    for (const item of fixture.slice) {
      expect(item.expected === undefined ? 1 : 0, item.name).toBe(
        item.refuse === undefined ? 0 : 1
      );
    }
  });

  for (const item of fixture.slice) {
    it(item.name, () => {
      const rect = parsePythonPrepareRect(item.range);
      expect(rect).not.toBeNull();
      if (item.refuse === true) {
        expect(() =>
          slicePythonPrepareRegion(fixture.sheet.matrix, fixture.sheet, rect!)
        ).toThrow(PYTHON_PREPARE_REFUSAL_MESSAGE);
        return;
      }
      expect(
        slicePythonPrepareRegion(fixture.sheet.matrix, fixture.sheet, rect!)
      ).toEqual(item.expected);
    });
  }
});

/**
 * 不完整矩阵（缺行/短行）。
 *
 * 只有"从已用区域读整块再切"的宿主会碰到（WPS）；Excel 的 `getRange` 恒返回
 * 完整矩形，所以它对应的保证是"越界一律拒绝"。两边都由**同一份**实现给出
 * 同一种拒绝：读了但没读到的格子绝不被补成空白。
 */
describe("缺失 / 不完整的已读矩阵", () => {
  for (const item of fixture.ragged) {
    it(item.name, () => {
      const rect = parsePythonPrepareRect(item.range)!;
      if (item.refuse === true) {
        expect(() =>
          slicePythonPrepareRegion(item.matrix, { startRow: 3, startColumn: 2 }, rect)
        ).toThrow(PythonPrepareRegionRefused);
        return;
      }
      expect(
        slicePythonPrepareRegion(item.matrix, { startRow: 3, startColumn: 2 }, rect)
      ).toEqual(item.expected);
    });
  }

  it("越界判定与切片判定给的是同一个拒绝（不是两套口径）", () => {
    const rect = parsePythonPrepareRect("B3:E7")!;
    expect(() =>
      assertPythonPrepareRectWithin(rect, { startRow: 3, startColumn: 2 }, 4, 4)
    ).toThrow(PythonPrepareRegionRefused);
  });
});

describe("值归一共享夹具（前导零 / 数字格式 / 日期）", () => {
  it("夹具覆盖了三类容易分叉的值", () => {
    const joined = fixture.values.map((item) => item.numberFormat).join("|");
    expect(fixture.values.some((item) => /^0+$/.test(item.numberFormat))).toBe(true);
    expect(fixture.values.some((item) => item.numberFormat === "@")).toBe(true);
    expect(joined).toContain("yyyy");
  });

  for (const item of fixture.values) {
    it(item.name, () => {
      expect(normalizeCellValue(item.value, item.text, item.numberFormat)).toEqual(
        item.expected
      );
    });
  }
});
