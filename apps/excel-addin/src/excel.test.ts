import { describe, expect, it } from "vitest";
import {
  dataEpochsChanged,
  incompleteActionResults,
  preflightPlanActions,
  snapshotDataEpochs,
  sourceFingerprintForSnapshot,
  sortedRangePasses,
  verificationGaps,
  valuesEqual
} from "./excel";
import type { AnalysisPlan } from "./contracts";

describe("workbook source fingerprint", () => {
  const snapshot = {
    name: "example.xlsx",
    worksheets: [
      {
        name: "数据",
        usedRange: "数据!A1:B2",
        rowCount: 2,
        columnCount: 2,
        headers: ["编码", "金额"],
        dataRows: [["001", 10]],
        truncated: false
      }
    ]
  };

  it("is stable for the same selected source data", () => {
    expect(sourceFingerprintForSnapshot(snapshot, ["数据"])).toBe(
      sourceFingerprintForSnapshot(snapshot, ["数据"])
    );
  });

  it("changes when previewed source data changes", () => {
    const changed = {
      ...snapshot,
      worksheets: [
        {
          ...snapshot.worksheets[0],
          dataRows: [["001", 11]]
        }
      ]
    };
    expect(sourceFingerprintForSnapshot(changed, ["数据"])).not.toBe(
      sourceFingerprintForSnapshot(snapshot, ["数据"])
    );
  });
});

describe("data epoch safety gate", () => {
  it("reports unchanged when snapshot matches current epochs", () => {
    const snapshot = snapshotDataEpochs(["数据", "预售"]);
    expect(dataEpochsChanged(snapshot)).toBe(false);
  });

  it("detects a bumped sheet epoch as changed", () => {
    const snapshot = snapshotDataEpochs(["数据"]);
    const stale = { ...snapshot, 数据: (snapshot["数据"] ?? 0) - 1 };
    expect(dataEpochsChanged(stale)).toBe(true);
  });

  it("detects a bumped global epoch as changed", () => {
    const snapshot = snapshotDataEpochs(["数据"]);
    const stale = { ...snapshot, __global__: (snapshot.__global__ ?? 0) - 1 };
    expect(dataEpochsChanged(stale)).toBe(true);
  });

  it("treats an unknown sheet with a nonzero epoch as changed", () => {
    const snapshot = snapshotDataEpochs([]);
    const stale = { ...snapshot, 未知表: 5 };
    expect(dataEpochsChanged(stale)).toBe(true);
  });
});

describe("Excel verification value equivalence", () => {
  it("accepts harmless numeric text coercion", () => {
    expect(valuesEqual(43122301, "43122301")).toBe(true);
    expect(valuesEqual("43122301", 43122301)).toBe(true);
  });

  it("rejects numeric coercion that loses identifier information", () => {
    expect(valuesEqual(1105882025, "001105882025")).toBe(false);
  });

  it("accepts negligible floating point differences", () => {
    expect(valuesEqual(0.1 + 0.2, 0.3)).toBe(true);
  });
});

describe("Excel plan preflight", () => {
  const plan: AnalysisPlan = {
    id: "preflight-plan",
    title: "预检",
    summary: "执行前发现跨步骤问题",
    assumptions: [],
    warnings: [],
    actions: [
      {
        type: "writeValues",
        sheet: "结果",
        range: "A1",
        values: [[1]]
      },
      {
        type: "copyRange",
        sheet: "复制结果",
        sourceSheet: "不存在",
        sourceRange: "A1",
        targetRange: "A1",
        copyType: "values",
        skipBlanks: false,
        transpose: false
      },
      {
        type: "createWorksheet",
        sheet: "后续步骤"
      }
    ]
  };

  it("finds a later missing source before any action runs", () => {
    expect(preflightPlanActions(plan, ["数据"])).toEqual([
      {
        index: 1,
        message: "未找到复制源工作表「不存在」"
      }
    ]);
  });

  it("reports unsupported Excel API requirements before execution", () => {
    const notePlan: AnalysisPlan = {
      ...plan,
      actions: [
        {
          type: "addNote",
          sheet: "数据",
          cell: "A1",
          text: "说明"
        }
      ]
    };

    expect(
      preflightPlanActions(
        notePlan,
        ["数据"],
        (version) => version !== "1.18"
      )
    ).toEqual([
      {
        index: 0,
        message:
          "当前 Excel 不支持此操作（需要 ExcelApi 1.18 或更高版本）"
      }
    ]);
  });

  it("requires APIs that expose chart and PivotTable data sources", () => {
    const objectPlan: AnalysisPlan = {
      ...plan,
      actions: [
        {
          type: "createChart",
          sheet: "数据",
          sourceRange: "A1:B3",
          chartType: "ColumnClustered",
          targetRange: "D2"
        },
        {
          type: "createPivotTable",
          sheet: "透视",
          sourceSheet: "数据",
          sourceRange: "A1:B3",
          name: "SummaryPivot",
          destinationCell: "A1",
          rowFields: [],
          columnFields: [],
          valueFields: []
        }
      ]
    };

    expect(
      preflightPlanActions(
        objectPlan,
        ["数据"],
        (version) => version !== "1.12" && version !== "1.15"
      )
    ).toEqual([
      {
        index: 0,
        message:
          "当前 Excel 不支持此操作（需要 ExcelApi 1.12 或更高版本）"
      },
      {
        index: 1,
        message:
          "当前 Excel 不支持此操作（需要 ExcelApi 1.15 或更高版本）"
      }
    ]);
  });

  it("finds matrix size mismatches and object name conflicts", () => {
    const conflictPlan: AnalysisPlan = {
      ...plan,
      actions: [
        {
          type: "writeValues",
          sheet: "数据",
          range: "A1:B2",
          values: [[1, 2]]
        },
        {
          type: "createTable",
          sheet: "数据",
          range: "A1:B2",
          name: "SalesData",
          hasHeaders: true
        },
        {
          type: "addNamedRange",
          sheet: "数据",
          name: "salesdata",
          range: "A1:B2"
        },
        {
          type: "createPivotTable",
          sheet: "透视",
          sourceSheet: "数据",
          sourceRange: "A1:B2",
          name: "SummaryPivot",
          destinationCell: "A1:B2",
          rowFields: [],
          columnFields: [],
          valueFields: []
        }
      ]
    };

    expect(
      preflightPlanActions(
        conflictPlan,
        ["数据"],
        () => true,
        {
          tableNames: ["SalesData"],
          pivotTableNames: ["SummaryPivot"]
        }
      )
    ).toEqual([
      {
        index: 0,
        message: "目标区域 A1:B2 是 2×2，但写入矩阵是 1×2"
      },
      {
        index: 1,
        message: "表格或命名区域名称「SalesData」已存在"
      },
      {
        index: 2,
        message: "表格或命名区域名称「salesdata」已存在"
      },
      {
        index: 3,
        message: "数据透视表目标位置必须是单个单元格，当前为 A1:B2"
      },
      {
        index: 3,
        message: "数据透视表名称「SummaryPivot」已存在"
      }
    ]);
  });

  it("detects duplicate image names within the same worksheet", () => {
    const imagePlan: AnalysisPlan = {
      ...plan,
      actions: [
        {
          type: "addImage",
          sheet: "数据",
          base64: "image",
          targetRange: "A1:B2",
          name: "Preview"
        },
        {
          type: "addImage",
          sheet: "数据",
          base64: "image",
          targetRange: "D1:E2",
          name: "preview"
        }
      ]
    };

    expect(
      preflightPlanActions(imagePlan, ["数据"], () => true, {
        shapeNamesBySheet: { 数据: ["ExistingShape"] }
      })
    ).toEqual([
      {
        index: 1,
        message: "图片或形状名称「preview」已存在"
      }
    ]);
  });

  it("rejects unknown range names but accepts existing named ranges", () => {
    const rangePlan: AnalysisPlan = {
      ...plan,
      actions: [
        {
          type: "clearRange",
          sheet: "数据",
          range: "not-a-range",
          applyTo: "contents"
        },
        {
          type: "setFill",
          sheet: "数据",
          range: "KnownArea",
          color: "#FFFFFF"
        }
      ]
    };

    expect(
      preflightPlanActions(rangePlan, ["数据"], () => true, {
        namedRangeNames: ["KnownArea"]
      })
    ).toEqual([
      {
        index: 0,
        message:
          "区域「not-a-range」不是有效的 A1 地址或现有命名区域"
      }
    ]);
  });

  it("builds succeeded, failed and not-run action results", () => {
    const results = incompleteActionResults(
      plan,
      new Map([[1, "源工作表不存在"]]),
      [
        {
          index: 0,
          type: "writeValues",
          sheet: "结果",
          status: "succeeded"
        }
      ]
    );

    expect(results).toEqual([
      {
        index: 0,
        type: "writeValues",
        sheet: "结果",
        status: "succeeded"
      },
      {
        index: 1,
        type: "copyRange",
        sheet: "复制结果",
        status: "failed",
        message: "源工作表不存在"
      },
      {
        index: 2,
        type: "createWorksheet",
        sheet: "后续步骤",
        status: "not_run"
      }
    ]);
  });
});

describe("Excel verification coverage", () => {
  it("checks multi-key sorting from the actual range values", () => {
    const values = [
      ["组", "得分"],
      ["A", 2],
      ["A", 1],
      ["B", 3],
      ["B", null]
    ];

    expect(
      sortedRangePasses(
        values,
        [
          { column: 0, ascending: true },
          { column: 1, ascending: false }
        ],
        true
      )
    ).toBe(true);
    expect(
      sortedRangePasses(
        [["组", "得分"], ["B", 3], ["A", 2]],
        [{ column: 0, ascending: true }],
        true
      )
    ).toBe(false);
  });

  it("recognizes sorting, filtering, clearing filters and tables as checked", () => {
    const plan: AnalysisPlan = {
      id: "structured-verification",
      title: "对象验收",
      summary: "验证排序、筛选和表格",
      assumptions: [],
      warnings: [],
      actions: [
        {
          type: "sortRange",
          sheet: "数据",
          range: "A1:B4",
          keys: [{ column: 1, ascending: false }],
          hasHeaders: true
        },
        {
          type: "filterRange",
          sheet: "数据",
          range: "A1:B4",
          column: 0,
          values: ["A"]
        },
        {
          type: "createTable",
          sheet: "数据",
          range: "A1:B4",
          name: "ResultTable",
          hasHeaders: true
        }
      ],
      acceptanceCriteria: [
        {
          type: "rangeSorted",
          sheet: "数据",
          range: "A1:B4",
          keys: [{ column: 1, ascending: false }],
          hasHeaders: true
        },
        {
          type: "filterApplied",
          sheet: "数据",
          range: "A1:B4",
          column: 0,
          values: ["A"]
        },
        {
          type: "tableExists",
          sheet: "数据",
          range: "A1:B4",
          name: "ResultTable",
          hasHeaders: true
        }
      ]
    };

    expect(verificationGaps(plan)).toEqual([]);
    expect(
      verificationGaps({
        ...plan,
        actions: [{ type: "clearFilter", sheet: "数据" }],
        acceptanceCriteria: [{ type: "filterCleared", sheet: "数据" }]
      })
    ).toEqual([]);
  });

  it("does not treat worksheet existence as verification of formatting", () => {
    const plan: AnalysisPlan = {
      id: "verification-gap",
      title: "格式验证",
      summary: "写值并设置格式",
      assumptions: [],
      warnings: [],
      actions: [
        {
          type: "writeValues",
          sheet: "结果",
          range: "A1",
          values: [[1]]
        },
        {
          type: "setNumberFormat",
          sheet: "结果",
          range: "A1",
          formatCode: "0.00"
        }
      ],
      acceptanceCriteria: [
        {
          type: "worksheetExists",
          sheet: "结果"
        },
        {
          type: "rangeEquals",
          sheet: "结果",
          range: "A1",
          expected: [[1]]
        }
      ]
    };

    expect(verificationGaps(plan)).toEqual([
      {
        index: 1,
        type: "setNumberFormat",
        sheet: "结果",
        message:
          "第 2 步 setNumberFormat 已执行，但当前验收协议不能独立验证该操作的具体效果"
      }
    ]);
  });

  it("recognizes independently checked formats, validation and freeze panes", () => {
    const plan: AnalysisPlan = {
      id: "format-verification",
      title: "格式验收",
      summary: "验证格式、数据验证与冻结窗格",
      assumptions: [],
      warnings: [],
      actions: [
        {
          type: "setFill",
          sheet: "结果",
          range: "A1:B2",
          color: "#DFF3E4"
        },
        {
          type: "setNumberFormat",
          sheet: "结果",
          range: "B2",
          formatCode: "0.00"
        },
        {
          type: "setDataValidation",
          sheet: "结果",
          range: "B2:B10",
          validationType: "wholeNumber",
          values: [],
          formula1: 0,
          formula2: 100,
          operator: "between",
          allowBlank: false
        },
        {
          type: "freezePanes",
          sheet: "结果",
          rows: 1,
          columns: 2
        }
      ],
      acceptanceCriteria: [
        {
          type: "rangeFormatMatches",
          sheet: "结果",
          range: "A1:B2",
          fillColor: "#DFF3E4"
        },
        {
          type: "rangeFormatMatches",
          sheet: "结果",
          range: "B2",
          numberFormat: "0.00"
        },
        {
          type: "dataValidationMatches",
          sheet: "结果",
          range: "B2:B10",
          validationType: "wholeNumber",
          values: [],
          formula1: 0,
          formula2: 100,
          operator: "between",
          allowBlank: false
        },
        {
          type: "freezePanesMatches",
          sheet: "结果",
          rows: 1,
          columns: 2
        }
      ]
    };

    expect(verificationGaps(plan)).toEqual([]);
  });

  it("recognizes chart and PivotTable object checks", () => {
    const plan: AnalysisPlan = {
      id: "object-verification",
      title: "对象验收",
      summary: "核对图表与数据透视表",
      assumptions: [],
      warnings: [],
      actions: [
        {
          type: "createChart",
          sheet: "数据",
          sourceRange: "A1:B3",
          chartType: "ColumnClustered",
          title: "得分",
          targetRange: "D2"
        },
        {
          type: "createPivotTable",
          sheet: "透视",
          sourceSheet: "数据",
          sourceRange: "A1:B3",
          name: "SummaryPivot",
          destinationCell: "A1",
          rowFields: ["人员"],
          columnFields: [],
          valueFields: [{ field: "得分", aggregation: "sum" }]
        }
      ],
      acceptanceCriteria: [
        {
          type: "chartExists",
          sheet: "数据",
          name: "Chart 1",
          sourceRange: "A1:B3",
          chartType: "ColumnClustered",
          title: "得分",
          targetRange: "D2"
        },
        {
          type: "pivotTableExists",
          sheet: "透视",
          sourceSheet: "数据",
          sourceRange: "A1:B3",
          name: "SummaryPivot",
          destinationCell: "A1",
          rowFields: ["人员"],
          columnFields: [],
          valueFields: [{ field: "得分", aggregation: "sum" }]
        }
      ]
    };

    expect(verificationGaps(plan)).toEqual([]);
  });

  it("recognizes deterministic value and worksheet checks", () => {
    const plan: AnalysisPlan = {
      id: "verification-covered",
      title: "确定性验证",
      summary: "创建并写入",
      assumptions: [],
      warnings: [],
      actions: [
        { type: "createWorksheet", sheet: "结果" },
        {
          type: "writeValues",
          sheet: "结果",
          range: "A1:B1",
          values: [[1, 2]]
        }
      ],
      acceptanceCriteria: [
        { type: "worksheetExists", sheet: "结果" },
        {
          type: "rangeEquals",
          sheet: "结果",
          range: "A1:B1",
          expected: [[1, 2]]
        }
      ]
    };

    expect(verificationGaps(plan)).toEqual([]);
  });
});
