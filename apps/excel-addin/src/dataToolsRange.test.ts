import { afterEach, describe, expect, it, vi } from "vitest";
import {
  DataToolExecutionError,
  executeQueryTableData,
  executeQueryTableTool
} from "./dataTools";
import type { CellValue, DataToolRequest } from "./contracts";
import { createOfficeTestHost } from "./officeTestHost";

/**
 * `query_table` 的显式只读区域（2026-11-01 现场续）。
 *
 * 原需求：「只读取来源!A1:B5，告诉我 B 列五个数的合计，不修改任何工作表。」
 * 整表查询不能代替它：`B6=100` 必须**不参与**。所以这一组同时断言三件事 ——
 * 结果、区域外不参与、零写入。
 */
afterEach(() => {
  vi.unstubAllGlobals();
});

/** A1:B5 是五个标签+数字，B6 故意放 100 当"区域外诱饵"。 */
const SOURCE: CellValue[][] = [
  ["甲", 1],
  ["乙", 2],
  ["丙", 3],
  ["丁", 4],
  ["戊", 5],
  [null, 100]
];

function sumRequest(
  overrides: Partial<DataToolRequest["arguments"]> = {}
): DataToolRequest {
  return {
    id: "range-sum",
    tool: "query_table",
    arguments: {
      mode: "aggregate",
      scope: "selected",
      metrics: [{ operation: "sum", field: "列2", outputName: "合计" }],
      limit: 10,
      ...overrides
    }
  };
}

describe("Excel 执行器：显式区域只读", () => {
  it("查 A1:B5 时 B6=100 不参与，合计 15，且零写入", async () => {
    const host = createOfficeTestHost({
      url: "file:///C:/Data/book.xlsx",
      sheets: [{ name: "来源", values: SOURCE }]
    });
    vi.stubGlobal("Excel", host.Excel);

    const result = await executeQueryTableTool(
      sumRequest({ range: "来源!A1:B5", hasHeaders: false }),
      ["来源"],
      "来源"
    );

    expect(result.rows).toEqual([[15]]);
    expect(result.scannedRows).toBe(5);
    // 只读查询一次写入都没有：宿主替身按 store 记账，任何 values/formulas 赋值都会被记下。
    expect(host.writes()).toEqual([]);
  });

  it("同一张表整表查就是 115 —— 说明区域确实排掉了 B6", async () => {
    const host = createOfficeTestHost({
      url: "file:///C:/Data/book.xlsx",
      sheets: [{ name: "来源", values: SOURCE }]
    });
    vi.stubGlobal("Excel", host.Excel);

    const result = await executeQueryTableTool(
      sumRequest({ hasHeaders: false }),
      ["来源"],
      "来源"
    );

    // 没有区域时 100 本来就该算进来（1+2+3+4+5+100）；有 15 的对照才证明区域生效了。
    expect(result.rows).toEqual([[115]]);
    expect(host.writes()).toEqual([]);
  });

  it("有表头的老行为在区域内照样成立：表头行不进指标，区域外也不算", async () => {
    const host = createOfficeTestHost({
      url: "file:///C:/Data/book.xlsx",
      sheets: [
        {
          name: "明细",
          values: [["数量"], [1], [2], [3], [100]]
        }
      ]
    });
    vi.stubGlobal("Excel", host.Excel);

    const result = await executeQueryTableTool(
      sumRequest({
        range: "明细!A1:A4",
        metrics: [{ operation: "sum", field: "数量", outputName: "合计" }]
      }),
      ["明细"],
      "明细"
    );

    // 表头（数量）不算，A5=100 在区域外也不算。
    expect(result.rows).toEqual([[6]]);
    expect(result.scannedRows).toBe(3);
    expect(host.writes()).toEqual([]);
  });

  it("区域指到未授权工作表：按越权拒绝，不去读它", async () => {
    const host = createOfficeTestHost({
      url: "file:///C:/Data/book.xlsx",
      sheets: [
        { name: "来源", values: SOURCE },
        { name: "私密", values: [[1], [2]] }
      ]
    });
    vi.stubGlobal("Excel", host.Excel);

    await expect(
      executeQueryTableTool(
        sumRequest({ range: "私密!A1:A2", hasHeaders: false }),
        ["来源"],
        "来源"
      )
    ).rejects.toMatchObject({ code: "SCOPE_VIOLATION" });
    expect(host.writes()).toEqual([]);
  });

  it("区域写法不合法就报错，绝不回退成整表读", async () => {
    const host = createOfficeTestHost({
      url: "file:///C:/Data/book.xlsx",
      sheets: [{ name: "来源", values: SOURCE }]
    });
    vi.stubGlobal("Excel", host.Excel);

    // 整列写法在这里必须失败，而不是变成"那就算整表"。
    await expect(
      executeQueryTableTool(
        sumRequest({ range: "来源!B:B", hasHeaders: false }),
        ["来源"],
        "来源"
      )
    ).rejects.toBeInstanceOf(DataToolExecutionError);
    expect(host.writes()).toEqual([]);
  });

  it("scope=active 与区域指定的表矛盾时按越权处理", async () => {
    const host = createOfficeTestHost({
      url: "file:///C:/Data/book.xlsx",
      sheets: [
        { name: "来源", values: SOURCE },
        { name: "其他", values: [[1]] }
      ]
    });
    vi.stubGlobal("Excel", host.Excel);

    await expect(
      executeQueryTableTool(
        sumRequest({
          scope: "active",
          range: "其他!A1:A1",
          hasHeaders: false
        }),
        ["来源", "其他"],
        "来源"
      )
    ).rejects.toMatchObject({ code: "SCOPE_VIOLATION" });
  });
});

describe("数据矩阵路径（WPS 桥与 computeValues 共用）：显式区域", () => {
  it("有原点时按区域裁剪，区域外不参与", () => {
    const result = executeQueryTableData(sumRequest({ range: "A1:B5" }), [
      { name: "来源", values: SOURCE, origin: { row: 1, column: 1 } }
    ]);

    expect(result.rows).toEqual([[15]]);
    expect(result.scannedRows).toBe(5);
  });

  it("区域与矩阵原点错开时仍按绝对行列取", () => {
    // 矩阵从 B3 开始；区域 B3:C4 应该正好取到它左上角的 2×2。
    const result = executeQueryTableData(
      {
        id: "offset",
        tool: "query_table",
        arguments: {
          mode: "rows",
          scope: "selected",
          range: "B3:C4",
          hasHeaders: false,
          limit: 10
        }
      },
      [
        {
          name: "来源",
          values: [
            ["甲", 1],
            ["乙", 2]
          ],
          origin: { row: 3, column: 2 }
        }
      ]
    );

    expect(result.rows).toEqual([
      ["来源", "甲", 1],
      ["来源", "乙", 2]
    ]);
  });

  it("给了区域却拿不到工作表原点：报错，不静默整表读", () => {
    expect(() =>
      executeQueryTableData(sumRequest({ range: "A1:B5" }), [
        { name: "来源", values: SOURCE }
      ])
    ).toThrow(/工作表原点/);
  });

  it("区域指定的表不在这次查询的矩阵里：按越权拒绝", () => {
    expect(() =>
      executeQueryTableData(sumRequest({ range: "别的!A1:B5" }), [
        { name: "来源", values: SOURCE, origin: { row: 1, column: 1 } }
      ])
    ).toThrow(/不在本次查询范围/);
  });
});

/**
 * 「字段名是 列N 就判无表头」会不会误伤真实表头恰叫「列1/列2」的表？
 * 会 —— 所以那条旁证必须让事实与显式声明先走。
 */
describe("合成列名 vs 真实表头", () => {
  const realSyntheticHeader: CellValue[][] = [
    ["列1", "列2"],
    ["甲", 1],
    ["乙", 2]
  ];

  it("真实表头就叫 列1/列2 时按表头算（不被命名规则覆盖）", () => {
    const result = executeQueryTableData(
      {
        id: "real-synthetic-header",
        tool: "query_table",
        arguments: {
          mode: "aggregate",
          scope: "selected",
          metrics: [{ operation: "sum", field: "列2", outputName: "合计" }],
          limit: 10
        }
      },
      [
        {
          name: "来源",
          values: realSyntheticHeader,
          origin: { row: 1, column: 1 }
        }
      ]
    );

    // 表头行被识别出来：合计是 1+2=3，而不是把表头文字算进去。
    expect(result.rows).toEqual([[3]]);
    expect(result.scannedRows).toBe(2);
  });

  it("显式 hasHeaders=true 优先：即使区域里没有一行写着 列N 也以首行为表头", () => {
    const result = executeQueryTableData(
      {
        id: "explicit-true",
        tool: "query_table",
        arguments: {
          mode: "aggregate",
          scope: "selected",
          hasHeaders: true,
          metrics: [{ operation: "countRows", outputName: "笔数" }],
          limit: 10
        }
      },
      [
        {
          name: "来源",
          values: [
            ["列1", "列2"],
            ["甲", 1],
            ["乙", 2]
          ],
          origin: { row: 1, column: 1 }
        }
      ]
    );

    expect(result.rows).toEqual([[2]]);
    expect(result.scannedRows).toBe(2);
  });

  it("无表头数据区仍然判无表头（那条旁证没有被废掉）", () => {
    const result = executeQueryTableData(
      sumRequest({}),
      [
        {
          name: "来源",
          values: [
            ["甲", 1],
            ["乙", 2]
          ],
          origin: { row: 1, column: 1 }
        }
      ]
    );

    expect(result.rows).toEqual([[3]]);
    expect(result.scannedRows).toBe(2);
  });
});
