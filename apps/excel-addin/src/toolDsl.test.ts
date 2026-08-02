import { describe, expect, it } from "vitest";
import type { AnalysisPlan } from "./contracts";
import { renderToolDsl } from "./toolDsl";

describe("renderToolDsl", () => {
  it("renders a split aggregate tool as a deterministic controlled plan", () => {
    const plan: AnalysisPlan = {
      id: "split-plan",
      title: "按影片拆分并统计",
      summary: "按影院统计影片记录数和占比",
      sourceFingerprintSheets: ["明细"],
      assumptions: [],
      warnings: ["空影片名称不会生成工作表"],
      actions: [
        {
          type: "splitGroupAggregate",
          sheet: "明细",
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
        }
      ]
    };

    const first = renderToolDsl(plan);
    const second = renderToolDsl(plan);

    expect(second).toBe(first);
    expect(first).toContain('TOOL "按影片拆分并统计"');
    expect(first).toContain('SOURCE_SCOPE ["明细"]');
    expect(first).toContain("STEP 1 SPLIT_GROUP_AGGREGATE");
    expect(first).toContain('SPLIT_BY = "影片名称"');
    expect(first).toContain("MAX_OUTPUT_SHEETS = 200");
    expect(first).toContain("WRITE_POLICY = PREVIEW_THEN_CONFIRM");
    expect(first).toContain("ARBITRARY_CODE = DISABLED");
  });

  it("escapes text and summarizes bulk payloads", () => {
    const plan: AnalysisPlan = {
      id: "write-plan",
      title: '写入 "结果"',
      summary: "第一行\n第二行",
      assumptions: [],
      warnings: [],
      actions: [
        {
          type: "writeValues",
          sheet: "结果",
          range: "A1:B2",
          values: [
            [1, 2],
            [3, 4]
          ]
        }
      ]
    };

    const dsl = renderToolDsl(plan);

    expect(dsl).toContain('TOOL "写入 \\"结果\\""');
    expect(dsl).toContain('SUMMARY "第一行\\n第二行"');
    expect(dsl).toContain("VALUES = <MATRIX ROWS=2 COLUMNS=2>");
    expect(dsl).not.toContain("[1, 2]");
  });
});
