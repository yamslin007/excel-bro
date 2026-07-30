import { describe, expect, it, vi } from "vitest";
import { diagnosticEvents } from "./diagnostics";
import {
  executeSavedQueryTool,
  SavedQueryToolFallbackError
} from "./deterministicTools";
import { createQueryTool } from "./storage";

const workbook = {
  name: "demo.xlsx",
  capturedAt: "2026-07-29T00:00:00Z",
  activeWorksheet: "数据",
  worksheets: [
    {
      name: "数据",
      usedRange: "A1:B2",
      rowCount: 2,
      columnCount: 2,
      headers: ["分类", "金额"],
      dataRows: [],
      truncated: false
    }
  ]
};

describe("deterministic saved query tools", () => {
  it("runs locally with zero model calls", async () => {
    const tool = createQueryTool(
      "分类统计",
      "按分类统计金额",
      {
        id: "template",
        tool: "query_table",
        arguments: {
          mode: "aggregate",
          groupBy: ["分类"],
          metrics: [{ operation: "sum", field: "金额", outputName: "合计" }]
        }
      },
      "workbook",
      ["数据"],
      [],
      ["分类", "合计"]
    );
    const runner = vi.fn().mockResolvedValue({
      requestId: "run",
      tool: "query_table",
      title: "结果",
      headers: ["分类", "合计"],
      rows: [["甲", 10]],
      sourceSheets: ["数据"],
      scannedRows: 1,
      complete: true,
      calculation: "sum",
      warnings: []
    });

    await expect(
      executeSavedQueryTool(tool, workbook, {
        workbook: runner,
        folder: vi.fn()
      })
    ).resolves.toMatchObject({ rows: [["甲", 10]] });
    expect(runner).toHaveBeenCalledOnce();
    expect(diagnosticEvents().at(-1)).toMatchObject({
      phase: "saved_tool",
      modelCalls: 0
    });
  });

  it("stops locally when fields have changed", async () => {
    const tool = createQueryTool(
      "旧字段",
      "字段不兼容时降级",
      {
        id: "template",
        tool: "query_table",
        arguments: { mode: "rows", fields: ["不存在"] }
      },
      "workbook",
      ["数据"]
    );

    await expect(
      executeSavedQueryTool(tool, workbook, {
        workbook: vi.fn(),
        folder: vi.fn()
      })
    ).rejects.toBeInstanceOf(SavedQueryToolFallbackError);
  });
});
