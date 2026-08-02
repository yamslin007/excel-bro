import { describe, expect, it, vi } from "vitest";
import type {
  AnalysisPlan,
  DataToolRequest,
  WorkbookSnapshot
} from "./contracts";
import {
  analyzeQueryToolCompatibility,
  analyzeToolEligibility,
  createQueryTool,
  createTool,
  deleteQueryTool,
  deleteTool,
  instantiateTool,
  saveQueryTool,
  saveTool
} from "./storage";

const splitPlan: AnalysisPlan = {
  id: "split-plan",
  title: "按影片拆分",
  summary: "按影院统计影片记录数和占比",
  assumptions: [],
  warnings: [],
  actions: [
    {
      type: "splitGroupAggregate",
      sheet: "旧数据",
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
  ],
  acceptanceCriteria: [{ type: "worksheetExists", sheet: "旧数据" }]
};

const workbook: WorkbookSnapshot = {
  name: "测试.xlsx",
  capturedAt: "2026-07-26T00:00:00.000Z",
  activeWorksheet: "本月数据",
  worksheets: [
    {
      name: "本月数据",
      usedRange: "A1:E20",
      rowCount: 20,
      columnCount: 5,
      headers: ["门店", "门店编码", "剧目", "场次", "占比"],
      dataRows: [],
      truncated: false
    },
    {
      name: "已有结果",
      usedRange: "A1",
      rowCount: 1,
      columnCount: 1,
      headers: ["完成"],
      dataRows: [],
      truncated: false
    }
  ]
};

describe("saved tools", () => {
  it("turns selected source sheets into reusable worksheet parameters", () => {
    const tool = createTool(
      splitPlan,
      "影院影片拆分",
      "生成影片子表",
      ["旧数据"]
    );

    expect(tool.version).toBe(2);
    expect(tool.parameters[0]).toMatchObject({
      id: "worksheet_1",
      label: "来源工作表",
      type: "worksheet",
      defaultValue: "旧数据",
      required: true
    });
    expect(
      tool.parameters
        .filter((parameter) => parameter.type === "field")
        .map((parameter) => parameter.defaultValue)
    ).toEqual(["影片名称", "影院名称", "影院编码"]);

    const instantiated = instantiateTool(tool, {
      worksheet_1: "本月数据",
      field_1: "剧目",
      field_2: "门店",
      field_3: "门店编码"
    });
    expect(instantiated.actions[0].sheet).toBe("本月数据");
    expect(instantiated.actions[0]).toMatchObject({
      type: "splitGroupAggregate",
      splitBy: "剧目",
      groupBy: ["门店", "门店编码"]
    });
    expect(instantiated.acceptanceCriteria?.[0].sheet).toBe("本月数据");
    expect(instantiated.title).toBe("影院影片拆分");
  });

  it("does not parameterize fixed output sheets that were not selected", () => {
    const plan: AnalysisPlan = {
      ...splitPlan,
      actions: [
        {
          type: "writeValues",
          sheet: "固定结果",
          range: "A1",
          values: [["完成"]]
        }
      ]
    };
    expect(() =>
      createTool(plan, "固定输出", "写入结果", ["旧数据"])
    ).toThrow("固定内容");
    const tool = createTool(
      plan,
      "固定输出",
      "写入结果",
      ["旧数据"],
      { fixedContent: true }
    );

    expect(tool.parameters).toEqual([]);
    expect(instantiateTool(tool, {}).actions[0].sheet).toBe("固定结果");
  });

  it("blocks embedded images and flags destructive actions", () => {
    const plan: AnalysisPlan = {
      ...splitPlan,
      actions: [
        {
          type: "addImage",
          sheet: "结果",
          base64: "AAAA",
          targetRange: "A1"
        },
        {
          type: "clearRange",
          sheet: "旧数据",
          range: "A1:B5",
          applyTo: "contents"
        }
      ]
    };

    const eligibility = analyzeToolEligibility(plan);

    expect(eligibility.blocked).toBe(true);
    expect(eligibility.requiredApprovals).toContain("destructive");
    expect(eligibility.issues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining(["EMBEDDED_IMAGE", "DESTRUCTIVE_ACTION"])
    );
  });

  it("parameterizes created output worksheets and rejects name collisions", () => {
    const plan: AnalysisPlan = {
      ...splitPlan,
      actions: [
        { type: "createWorksheet", sheet: "汇总结果" },
        {
          type: "writeValues",
          sheet: "汇总结果",
          range: "A1",
          values: [["完成"]]
        }
      ],
      acceptanceCriteria: [
        { type: "worksheetExists", sheet: "汇总结果" }
      ]
    };
    const tool = createTool(
      plan,
      "生成汇总",
      "创建新的汇总表",
      [],
      { fixedContent: true }
    );
    const output = tool.parameters.find(
      (parameter) => parameter.type === "outputWorksheet"
    );

    expect(output).toMatchObject({
      defaultValue: "汇总结果",
      label: "输出工作表名称"
    });
    const instantiated = instantiateTool(
      tool,
      { [output!.id]: "本次汇总" },
      workbook
    );
    expect(instantiated.actions.map((action) => action.sheet)).toEqual([
      "本次汇总",
      "本次汇总"
    ]);
    expect(instantiated.acceptanceCriteria?.[0].sheet).toBe("本次汇总");
    expect(() =>
      instantiateTool(
        tool,
        { [output!.id]: "已有结果" },
        workbook
      )
    ).toThrow("已存在");
    expect(() =>
      instantiateTool(tool, { [output!.id]: "非法/名称" })
    ).toThrow("无效");
  });

  it("parameterizes reusable source data ranges", () => {
    const plan: AnalysisPlan = {
      ...splitPlan,
      actions: [
        {
          type: "sortRange",
          sheet: "旧数据",
          range: "A1:E812",
          keys: [{ column: 4, ascending: false }],
          hasHeaders: true
        }
      ]
    };
    const tool = createTool(plan, "排序", "按占比排序", ["旧数据"]);
    const range = tool.parameters.find(
      (parameter) => parameter.type === "range"
    );

    expect(range).toMatchObject({
      defaultValue: "A1:E812",
      sourceParameterId: "worksheet_1"
    });
    const instantiated = instantiateTool(
      tool,
      {
        worksheet_1: "本月数据",
        [range!.id]: "A1:E20"
      },
      workbook
    );
    expect(instantiated.actions[0]).toMatchObject({
      type: "sortRange",
      sheet: "本月数据",
      range: "A1:E20"
    });
  });

  it("rejects a saved folder query when the stable source id changes", () => {
    const request: DataToolRequest = {
      id: "saved-folder-query",
      tool: "query_table",
      arguments: { mode: "rows", fields: ["门店"] }
    };
    const tool = createQueryTool(
      "门店查询",
      "读取门店",
      request,
      "folder",
      ["本月数据"],
      ["old-sheet-id"],
      ["门店"]
    );
    const current = structuredClone(workbook);
    current.worksheets[0].sourceSheetId = "new-sheet-id";

    expect(analyzeQueryToolCompatibility(tool, current)).toMatchObject({
      runnable: false,
      requiresModel: true,
      reasons: expect.arrayContaining([
        "文件夹来源 ID 已变化，请重新确认数据来源"
      ])
    });
  });

  it("deletes workflow and query tools from persistent storage", () => {
    const values = new Map<string, string>();
    vi.stubGlobal("localStorage", {
      getItem: (key: string) => values.get(key) ?? null,
      setItem: (key: string, value: string) => values.set(key, value),
      removeItem: (key: string) => values.delete(key),
      clear: () => values.clear(),
      key: (index: number) => [...values.keys()][index] ?? null,
      get length() {
        return values.size;
      }
    });
    try {
      const workflow = createTool(
        splitPlan,
        "影院影片拆分",
        "生成影片子表",
        ["旧数据"]
      );
      const query = createQueryTool(
        "门店查询",
        "读取门店",
        {
          id: "saved-query",
          tool: "query_table",
          arguments: { mode: "rows", fields: ["门店"] }
        },
        "workbook",
        ["本月数据"]
      );

      saveTool(workflow);
      saveQueryTool(query);

      expect(deleteTool(workflow.id)).toEqual([]);
      expect(deleteQueryTool(query.id)).toEqual([]);
      expect(JSON.parse(values.get("excel-bro.tools.v2") ?? "[]")).toEqual([]);
      expect(
        JSON.parse(values.get("excel-bro.query-tools.v1") ?? "[]")
      ).toEqual([]);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
