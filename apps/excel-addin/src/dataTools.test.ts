import { describe, expect, it } from "vitest";
import {
  createQueryTableAccumulator,
  DataToolExecutionError,
  executeQueryTableData,
  readSheetIntoAccumulator,
  resolveFilteredRowIndices
} from "./dataTools";
import type { CellValue, DataToolRequest } from "./contracts";
import capabilities from "../../../config/capabilities.json";

describe("executeQueryTableData", () => {
  it("reports the filtered record total across sheets before limiting groups", () => {
    const result = executeQueryTableData({
      id: "union-total", tool: "query_table", arguments: {
        mode: "aggregate", combine: { mode: "union" }, groupBy: ["类别"],
        metrics: [{ operation: "countRows", outputName: "数量" }], limit: 1,
        filters: [{ field: "类别", operator: "notEquals", value: "排除" }]
      }
    }, [
      { name: "一", values: [["类别"], ["甲"], ["乙"], ["排除"]] },
      { name: "二", values: [["类别"], ["甲"], ["丙"]] }
    ]);
    expect(result.rows).toHaveLength(1);
    expect(result.calculation).toContain("筛选后总记录数：4");
  });
  it("detects a title row and aggregates complete selected-sheet data", () => {
    const request: DataToolRequest = {
      id: "aggregate-movies",
      tool: "query_table",
      arguments: {
        mode: "aggregate",
        scope: "selected",
        groupBy: ["影片名称"],
        metrics: [
          {
            operation: "sum",
            field: "计数项:影厅",
            outputName: "影厅数",
            ratioOutputName: "占比"
          }
        ],
        sortBy: "影厅数",
        sortDirection: "desc",
        limit: 10
      }
    };

    const result = executeQueryTableData(request, [
      {
        name: "表1",
        values: [
          ["影院报表", null, null],
          [null, null, null],
          ["影院名称", "影片名称", "计数项:影厅"],
          ["影院A", "功夫女足", 11],
          ["影院B", "恐怖游轮", 2]
        ]
      },
      {
        name: "表2",
        values: [
          ["影院名称", "影片名称", "计数项:影厅"],
          ["影院C", "功夫女足", 13],
          ["影院D", "恐怖游轮", 1]
        ]
      }
    ]);

    expect(result.headers).toEqual(["影片名称", "影厅数", "占比"]);
    expect(result.rows).toEqual([
      ["功夫女足", 24, 24 / 27],
      ["恐怖游轮", 3, 3 / 27]
    ]);
    expect(result.scannedRows).toBe(4);
    expect(result.complete).toBe(true);
  });

  it("filters and returns the highest existing value", () => {
    const request: DataToolRequest = {
      id: "highest-score",
      tool: "query_table",
      arguments: {
        mode: "rows",
        fields: ["姓名", "分数"],
        filters: [
          {
            field: "学科",
            operator: "equals",
            value: "数学"
          }
        ],
        sortBy: "分数",
        sortDirection: "desc",
        limit: 1
      }
    };

    const result = executeQueryTableData(request, [
      {
        name: "成绩",
        values: [
          ["姓名", "学科", "分数"],
          ["小林", "数学", 88],
          ["阿狸", "数学", 93],
          ["马晓", "语文", 96]
        ]
      }
    ]);

    expect(result.rows).toEqual([["阿狸", 93]]);
    expect(result.scannedRows).toBe(3);
  });

  it("returns a structured retryable error for unknown fields", () => {
    const request: DataToolRequest = {
      id: "unknown-field",
      tool: "query_table",
      arguments: {
        mode: "rows",
        fields: ["不存在字段"]
      }
    };

    expect(() =>
      executeQueryTableData(request, [
        {
          name: "数据",
          values: [
            ["分类", "数值"],
            ["甲", 10]
          ]
        }
      ])
    ).toThrowError(DataToolExecutionError);

    try {
      executeQueryTableData(request, [
        {
          name: "数据",
          values: [
            ["分类", "数值"],
            ["甲", 10]
          ]
        }
      ]);
    } catch (error) {
      expect(error).toMatchObject({
        code: "FIELD_NOT_FOUND",
        retryable: true,
        availableFields: expect.arrayContaining(["分类", "数值"])
      });
    }
  });

  it(
    "calculates min and max across the configured 250,000-row limit",
    () => {
      const request: DataToolRequest = {
        id: "large-extremes",
        tool: "query_table",
        arguments: {
          mode: "aggregate",
          metrics: [
            {
              operation: "min",
              field: "数值",
              outputName: "最小值"
            },
            {
              operation: "max",
              field: "数值",
              outputName: "最大值"
            }
          ]
        }
      };
      const values = [
        ["数值"],
        ...Array.from({ length: 250_000 }, (_, index) => [index + 1])
      ];

      const result = executeQueryTableData(request, [
        { name: "大数据", values }
      ]);

      expect(result.rows).toEqual([[1, 250_000]]);
      expect(result.scannedRows).toBe(250_000);
    },
    15_000
  );

  it("preserves identifiers and handles dates, percentages and formatted numbers", () => {
    const aggregateRequest: DataToolRequest = {
      id: "typed-values",
      tool: "query_table",
      arguments: {
        mode: "aggregate",
        metrics: [
          {
            operation: "sum",
            field: "完成率",
            outputName: "完成率合计"
          },
          {
            operation: "sum",
            field: "金额",
            outputName: "金额合计"
          }
        ]
      }
    };
    const rowsRequest: DataToolRequest = {
      id: "dated-row",
      tool: "query_table",
      arguments: {
        mode: "rows",
        fields: ["日期", "编码"],
        filters: [
          {
            field: "日期",
            operator: "greaterThanOrEqual",
            value: "2026-07-02"
          }
        ]
      }
    };
    const sheets = [
      {
        name: "数据",
        values: [
          ["日期", "完成率", "编码", "金额"],
          ["2026-07-01", "25%", "001", "1,200.50"],
          ["2026-07-02", 0.5, "002", 300]
        ]
      }
    ];

    expect(executeQueryTableData(aggregateRequest, sheets).rows).toEqual([
      [0.75, 1500.5]
    ]);
    expect(executeQueryTableData(rowsRequest, sheets).rows).toEqual([
      ["2026-07-02", "002"]
    ]);
  });

  it("marks multi-sheet results incomplete when one selected sheet lacks fields", () => {
    const request: DataToolRequest = {
      id: "partial-multi-sheet",
      tool: "query_table",
      arguments: {
        mode: "aggregate",
        groupBy: ["分类"],
        metrics: [
          {
            operation: "sum",
            field: "数值",
            outputName: "合计"
          }
        ]
      }
    };

    const result = executeQueryTableData(request, [
      {
        name: "有效表",
        values: [
          ["分类", "数值"],
          ["甲", 10]
        ]
      },
      {
        name: "缺字段表",
        values: [
          ["分类", "备注"],
          ["乙", "无数值列"]
        ]
      }
    ]);

    expect(result.rows).toEqual([["甲", 10]]);
    expect(result.complete).toBe(false);
    expect(result.calculation).toContain("已处理记录数（结果不完整）：1");
    expect(result.calculation).not.toContain("总记录数");
    expect(result.warnings).toEqual([
      "「缺字段表」未找到字段：数值"
    ]);
  });

  it("produces exact aggregate results while consuming row chunks", () => {
    const request: DataToolRequest = {
      id: "chunked-average",
      tool: "query_table",
      arguments: {
        mode: "aggregate",
        groupBy: ["分类"],
        metrics: [
          {
            operation: "average",
            field: "数值",
            outputName: "平均值"
          },
          {
            operation: "countDistinct",
            field: "编号",
            outputName: "编号数"
          }
        ],
        sortBy: "分类",
        sortDirection: "asc"
      }
    };
    const accumulator = createQueryTableAccumulator(request);
    accumulator.addSheet({
      name: "数据",
      values: [
        ["分类", "编号", "数值"],
        ["甲", "001", 10],
        ["乙", "002", 20]
      ]
    });
    accumulator.addSheet({
      name: "数据",
      values: [
        ["分类", "编号", "数值"],
        ["甲", "001", 30],
        ["乙", "003", 40]
      ]
    });

    expect(accumulator.finish()).toMatchObject({
      rows: [
        ["甲", 20, 1],
        ["乙", 30, 2]
      ],
      scannedRows: 4,
      sourceSheets: ["数据"]
    });
  });

  it("rejects folder-only combine modes instead of silently ignoring them", () => {
    const request = {
      id: "join-current-workbook",
      tool: "query_table",
      arguments: {
        mode: "rows",
        combine: {
          mode: "join",
          leftSourceSheetId: "left",
          rightSourceSheetId: "right",
          leftKey: "编号",
          rightKey: "编号",
          joinHow: "inner"
        }
      }
    } as DataToolRequest;

    expect(() => executeQueryTableData(request, [])).toThrow(
      "请切换到文件夹模式"
    );
  });
});

describe("resolveFilteredRowIndices", () => {
  const values = [
    ["人员", "得分"],
    ["阿里", 44],
    ["企鹅", -3],
    ["阿里", 50],
    ["企鹅", -8]
  ];

  it("matches equals on a header-named field", () => {
    const indices = resolveFilteredRowIndices(values, 0, [
      { field: "人员", operator: "equals", value: "阿里" }
    ]);
    expect(indices).toEqual([1, 3]);
  });

  it("ANDs multiple filters", () => {
    const indices = resolveFilteredRowIndices(values, 0, [
      { field: "人员", operator: "equals", value: "阿里" },
      { field: "得分", operator: "greaterThan", value: 45 }
    ]);
    expect(indices).toEqual([3]);
  });

  it("matches contains case-insensitively", () => {
    const indices = resolveFilteredRowIndices(
      [["名称"], ["北京烤鸭"], ["上海小笼"]],
      0,
      [{ field: "名称", operator: "contains", value: "烤" }]
    );
    expect(indices).toEqual([1]);
  });

  it("returns empty when the field is not found", () => {
    const indices = resolveFilteredRowIndices(values, 0, [
      { field: "不存在", operator: "equals", value: "阿里" }
    ]);
    expect(indices).toEqual([]);
  });

  it("returns empty when there is no header row", () => {
    const indices = resolveFilteredRowIndices(values, -1, [
      { field: "人员", operator: "equals", value: "阿里" }
    ]);
    expect(indices).toEqual([]);
  });
});

describe("readSheetIntoAccumulator", () => {
  const CHUNK_ROWS = capabilities.queryTable.chunkRows;

  /**
   * 造一张跨分块的表：表头 + N 行数据，前一半属于「甲」组、后一半属于「乙」组。
   * 分界点故意落在第二块里，这样若后续块丢了表头（第一行数据被当成表头吞掉），
   * 分组计数就会少算。
   */
  function makeSheet(dataRows: number): CellValue[][] {
    const rows: CellValue[][] = [["人员", "得分"]];
    for (let index = 0; index < dataRows; index += 1) {
      rows.push([index < dataRows / 2 ? "甲" : "乙", 1]);
    }
    return rows;
  }

  /** 只实现 readSheetIntoAccumulator 用到的那点 Office.js 表面。 */
  function fakeContext(sheet: CellValue[][]) {
    const pendingLoads: { start: number; count: number; box: FakeRange }[] = [];
    interface FakeRange {
      values: CellValue[][];
      text: string[][];
      numberFormat: string[][];
      load(properties: string): void;
    }
    const worksheet = {
      getRangeByIndexes(
        rowIndex: number,
        _columnIndex: number,
        rowCount: number
      ) {
        const box: FakeRange = {
          values: [],
          text: [],
          numberFormat: [],
          load() {
            pendingLoads.push({ start: rowIndex, count: rowCount, box });
          }
        };
        return box;
      }
    };
    let syncCount = 0;
    const context = {
      sync: async () => {
        syncCount += 1;
        for (const { start, count, box } of pendingLoads.splice(0)) {
          box.values = sheet.slice(start, start + count);
          box.text = box.values.map((row) => row.map(() => ""));
          box.numberFormat = box.values.map((row) => row.map(() => "General"));
        }
      }
    };
    return {
      context: context as unknown as Excel.RequestContext,
      worksheet: worksheet as unknown as Excel.Worksheet,
      syncCount: () => syncCount
    };
  }

  function countingRequest(): DataToolRequest {
    return {
      id: "cross-chunk",
      tool: "query_table",
      arguments: {
        mode: "aggregate",
        scope: "selected",
        groupBy: ["人员"],
        metrics: [{ operation: "countRows", outputName: "笔数" }],
        limit: 10
      }
    };
  }

  it("re-injects the header so groups stay correct across chunks", async () => {
    const dataRows = CHUNK_ROWS * 2 + 37;
    const sheet = makeSheet(dataRows);
    const fake = fakeContext(sheet);
    const request = countingRequest();
    const accumulator = createQueryTableAccumulator(request);

    const rowsRead = await readSheetIntoAccumulator(
      fake.context,
      {
        name: "数据",
        worksheet: fake.worksheet,
        rowIndex: 0,
        columnIndex: 0,
        rowCount: sheet.length,
        columnCount: 2
      },
      accumulator,
      { queryArguments: request.arguments }
    );

    expect(rowsRead).toBe(sheet.length);
    // 三块（含表头共 10037 行 → 5000 + 5000 + 37），每块一次 sync。
    expect(fake.syncCount()).toBe(3);

    const result = accumulator.finish();
    const total = result.rows.reduce(
      (sum, row) => sum + Number(row[row.length - 1] ?? 0),
      0
    );
    // 一行都不能丢：跨块时若表头没拼回去，后续块的首行会被吞掉。
    expect(total).toBe(dataRows);
    expect(result.rows).toHaveLength(2);
    expect(result.scannedRows).toBe(dataRows);
  });

  it("reports an empty sheet without reading anything", async () => {
    const fake = fakeContext([]);
    const request = countingRequest();
    const accumulator = createQueryTableAccumulator(request);

    const rowsRead = await readSheetIntoAccumulator(
      fake.context,
      {
        name: "空表",
        worksheet: fake.worksheet,
        rowIndex: 0,
        columnIndex: 0,
        rowCount: 0,
        columnCount: 0
      },
      accumulator,
      { queryArguments: request.arguments }
    );

    expect(rowsRead).toBe(0);
    expect(fake.syncCount()).toBe(0);
  });

  it("aborts between chunks when the signal is already aborted", async () => {
    const sheet = makeSheet(CHUNK_ROWS + 10);
    const fake = fakeContext(sheet);
    const request = countingRequest();
    const accumulator = createQueryTableAccumulator(request);
    const controller = new AbortController();
    controller.abort();

    await expect(
      readSheetIntoAccumulator(
        fake.context,
        {
          name: "数据",
          worksheet: fake.worksheet,
          rowIndex: 0,
          columnIndex: 0,
          rowCount: sheet.length,
          columnCount: 2
        },
        accumulator,
        { queryArguments: request.arguments, signal: controller.signal }
      )
    ).rejects.toThrow(DataToolExecutionError);
    expect(fake.syncCount()).toBe(0);
  });

  it("synthesizes a header row when the source has none", async () => {
    // 无表头数据：第一行就是真实数据，不能被当成表头吞掉。
    const sheet: CellValue[][] = [
      ["甲", 1],
      ["甲", 1],
      ["乙", 1]
    ];
    const fake = fakeContext(sheet);
    const request: DataToolRequest = {
      id: "no-header",
      tool: "query_table",
      arguments: {
        mode: "aggregate",
        scope: "selected",
        groupBy: ["列1"],
        metrics: [{ operation: "countRows", outputName: "笔数" }],
        limit: 10
      }
    };
    const accumulator = createQueryTableAccumulator(request);

    await readSheetIntoAccumulator(
      fake.context,
      {
        name: "无表头",
        worksheet: fake.worksheet,
        rowIndex: 0,
        columnIndex: 0,
        rowCount: sheet.length,
        columnCount: 2
      },
      accumulator,
      { queryArguments: request.arguments, headerMode: "synthesize" }
    );

    const result = accumulator.finish();
    const total = result.rows.reduce(
      (sum, row) => sum + Number(row[row.length - 1] ?? 0),
      0
    );
    expect(total).toBe(3);
    expect(result.rows).toHaveLength(2);
  });
});
