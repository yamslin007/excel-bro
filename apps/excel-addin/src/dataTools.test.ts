import { describe, expect, it } from "vitest";
import {
  createQueryTableAccumulator,
  DataToolExecutionError,
  executeQueryTableData
} from "./dataTools";
import type { DataToolRequest } from "./contracts";

describe("executeQueryTableData", () => {
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
