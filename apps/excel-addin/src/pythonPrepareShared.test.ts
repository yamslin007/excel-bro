/**
 * `prepare_python_data` 的**双端共同契约**（Excel 侧）。
 *
 * 夹具来自 `config/python_prepare_cases.json`，WPS 侧
 * `apps/wps-addin/src/wpsPythonPrepareShared.test.ts` 跑**同一份**用例。
 *
 * 为什么不能只靠"本端 26 passed"：Excel 用 `getRange(绝对地址)` 直接读，
 * 天然没有"已用区域不从 A1 起"的映射问题；WPS 读的是已用区域的**相对**矩阵。
 * 两端算法不同却必须给出**同一个答案**，这件事只有共用夹具才能证明 ——
 * 曾经 WPS 侧整体错位 `(startRow-1, startColumn-1)` 且**不报错**，
 * 单端测试全绿。
 */

import { beforeEach, describe, expect, it } from "vitest";

import { createOfficeTestHost, installOfficeTestHost } from "./officeTestHost";
import { executePreparePythonData } from "./dataTools";
import shared from "../../../config/python_prepare_cases.json";

const fixture = shared as {
  sheet: {
    name: string;
    startRow: number;
    startColumn: number;
    matrix: (string | number)[][];
    texts: string[][];
    numberFormats: string[][];
  };
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
};

const SHEET = fixture.sheet.name;

/** 把夹具矩阵按 `startRow/startColumn` 铺到**绝对**位置（左侧/上方留空）。 */
function padded(): (string | number)[][] {
  const { startRow, startColumn, matrix } = fixture.sheet;
  const top = Array.from({ length: startRow - 1 }, () =>
    Array.from({ length: startColumn - 1 + matrix[0].length }, () => "")
  );
  return [
    ...top,
    ...matrix.map((row) => [...Array.from({ length: startColumn - 1 }, () => ""), ...row])
  ];
}

function columnName(index: number): string {
  let value = index + 1;
  let text = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    text = String.fromCharCode(65 + remainder) + text;
    value = Math.floor((value - remainder) / 26);
  }
  return text;
}

function installSheetHost() {
  const host = createOfficeTestHost({
    url: "https://localhost/",
    sheets: [{ name: SHEET, values: padded() }]
  });
  /*
   * 只给**有内容的格**挂显示层信息：替身里空串是"没有这个格"，给它挂 text
   * 会把已用区域撑大，反而把"真实空白"这条用例的前提改掉。
   */
  const { startRow, startColumn, matrix, texts, numberFormats } = fixture.sheet;
  matrix.forEach((row, rowIndex) =>
    row.forEach((value, columnIndex) => {
      if (value === "" || value === null) return;
      const address = `${columnName(startColumn - 1 + columnIndex)}${startRow + rowIndex}`;
      host.setCellDisplay(SHEET, address, {
        text: texts[rowIndex]?.[columnIndex] ?? "",
        format: numberFormats[rowIndex]?.[columnIndex] ?? "General"
      });
    })
  );
  installOfficeTestHost(host);
  return host;
}

describe("prepare_python_data 双端共同契约（Excel 侧）", () => {
  beforeEach(() => {
    installSheetHost();
  });

  for (const item of fixture.slice) {
    it(item.name, async () => {
      const run = executePreparePythonData(
        [{ sheet: SHEET, range: item.range }],
        [SHEET]
      );
      if (item.refuse === true) {
        await expect(run).rejects.toThrow("数据登记区域超出已读取范围");
        return;
      }
      const blocks = await run;
      expect(blocks).toHaveLength(1);
      expect(blocks[0].sourceIndex).toBe(0);
      expect(blocks[0].values).toEqual(item.expected);
    });
  }

  it("真实空白与人为补齐必须可区分：区域内的空格是空串，区域外直接拒绝", async () => {
    const inside = await executePreparePythonData(
      [{ sheet: SHEET, range: "E6:E6" }],
      [SHEET]
    );
    expect(inside[0].values).toEqual([[""]]);
    await expect(
      executePreparePythonData([{ sheet: SHEET, range: "E7:E7" }], [SHEET])
    ).rejects.toThrow("数据登记区域超出已读取范围");
  });

  it("授权外的表一律拒绝，且不做任何读取", async () => {
    await expect(
      executePreparePythonData([{ sheet: "没授权的表", range: "A1:A1" }], [SHEET])
    ).rejects.toThrow("数据登记请求超出了用户选择的工作表范围");
  });

  it("取消：已 abort 的信号不得返回任何数据块", async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(
      executePreparePythonData([{ sheet: SHEET, range: "B3:E6" }], [SHEET], {
        signal: controller.signal
      })
    ).rejects.toThrow("本地数据登记已取消");
  });
});

describe("值归一共同契约（Excel 侧）", () => {
  for (const item of fixture.values) {
    it(item.name, async () => {
      const host = createOfficeTestHost({
        url: "https://localhost/",
        sheets: [{ name: "值契表", values: [[item.value as never]] }]
      });
      if (item.value !== "" && item.value !== null) {
        host.setCellDisplay("值契表", "A1", {
          text: item.text,
          format: item.numberFormat
        });
      }
      installOfficeTestHost(host);
      const blocks = await executePreparePythonData(
        [{ sheet: "值契表", range: "A1:A1" }],
        ["值契表"]
      );
      expect(blocks[0].values).toEqual([[item.expected]]);
    });
  }
});
