import { describe, expect, it } from "vitest";
import {
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
});
