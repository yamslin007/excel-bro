import { describe, expect, it } from "vitest";
import type { ExcelAction } from "./contracts";
import {
  batchSplitAggregateOutputs,
  buildSplitAggregateOutputs,
  safeWorksheetBaseName
} from "./splitAggregate";

describe("buildSplitAggregateOutputs", () => {
  it("finds a header below a title row and calculates group ratios", () => {
    const action: Extract<ExcelAction, { type: "splitGroupAggregate" }> = {
      type: "splitGroupAggregate",
      sheet: "sheet-dc1",
      splitBy: "影片名称",
      groupBy: ["影院名称", "影院编码"],
      metrics: [
        {
          operation: "countRows",
          outputName: "计数项:影厅",
          ratioOutputName: "占比"
        }
      ],
      includeBlankSplitValues: false,
      existingSheetPolicy: "rename",
      maxOutputSheets: 200
    };
    const outputs = buildSplitAggregateOutputs(
      [
        ["hunan(制表日期:2026-07-25)", null, null, null],
        [null, null, null, null],
        ["影院名称", "影院编码", "影厅", "影片名称"],
        ["怀化横店电影城", "43122301", "1号厅", "恐怖游轮"],
        ["怀化横店电影城", "43122301", "2号厅", "恐怖游轮"],
        ["怀化横店电影城", "43122301", "3号厅", "功夫女足"],
        ["广州影院", "44010001", "1号厅", "恐怖游轮"]
      ],
      action
    );

    expect(outputs.map((output) => output.splitValue)).toEqual([
      "恐怖游轮",
      "功夫女足"
    ]);
    expect(outputs[0].headers).toEqual([
      "影院名称",
      "影院编码",
      "影片名称",
      "计数项:影厅",
      "占比"
    ]);
    expect(outputs[0].rows).toEqual([
      ["怀化横店电影城", "43122301", "恐怖游轮", 2, 2 / 3],
      ["广州影院", "44010001", "恐怖游轮", 1, 1]
    ]);
    expect(outputs[0].ratioColumnIndexes).toEqual([4]);
  });

  it("rejects output explosions", () => {
    const action: Extract<ExcelAction, { type: "splitGroupAggregate" }> = {
      type: "splitGroupAggregate",
      sheet: "源数据",
      splitBy: "类别",
      groupBy: ["门店"],
      metrics: [{ operation: "countRows", outputName: "数量" }],
      includeBlankSplitValues: false,
      existingSheetPolicy: "rename",
      maxOutputSheets: 1
    };
    expect(() =>
      buildSplitAggregateOutputs(
        [
          ["门店", "类别"],
          ["A", "甲"],
          ["A", "乙"]
        ],
        action
      )
    ).toThrow("超过安全上限");
  });

  it("sanitizes ASCII and full-width characters rejected by Excel", () => {
    expect(safeWorksheetBaseName("三国第一部：争洛阳")).toBe(
      "三国第一部 争洛阳"
    );
    expect(safeWorksheetBaseName("'报表［终版］'")).toBe("报表 终版");
    expect(safeWorksheetBaseName("History")).toBe("History 结果");
  });

  it("batches worksheet outputs to reduce Office.js sync round trips", () => {
    const outputs = Array.from({ length: 26 }, (_, index) => ({
      splitValue: `影片 ${index + 1}`,
      headers: ["影院名称", "影片名称", "数量"],
      rows: [["影院", `影片 ${index + 1}`, 1]],
      ratioColumnIndexes: []
    }));

    expect(
      batchSplitAggregateOutputs(outputs, 25).map((batch) => batch.length)
    ).toEqual([25, 1]);
    expect(() => batchSplitAggregateOutputs(outputs, 0)).toThrow(
      "批次大小必须是正整数"
    );
  });
});
