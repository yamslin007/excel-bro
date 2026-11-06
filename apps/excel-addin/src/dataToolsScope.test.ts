import { afterEach, describe, expect, it, vi } from "vitest";
import {
  DataToolExecutionError,
  executeQueryTableTool
} from "./dataTools";
import type { CellValue, DataToolRequest } from "./contracts";
import { createOfficeTestHost } from "./officeTestHost";

/**
 * 「本次查哪几张表」——从模型式参数到执行结果的整条链路。
 *
 * 现场（2026-11-01 真机）：用户只说「只读取来源!A1:B5」，结果卡却列出四张来源表、
 * 三张各报一条「缺少列2」，还把不完整说成「扫描已达上限」。账本里 `sourceSheets`
 * 是 `["来源","结果1","结果2","结果3"]`，而 `query_table` 的参数当时不落账。
 *
 * 这里的夹具刻意让**四张表都有可求和的第二列**：一旦范围没收窄，合计数会变成
 * 315（15 + 100×3）而不是 15 —— 数字本身就是判别点，不需要读源码断言。
 */
afterEach(() => {
  vi.unstubAllGlobals();
});

/** 四张表都已授权，且每张都有第二列数字：只要多读一张，合计就变。 */
const SHEETS: Array<{ name: string; values: CellValue[][] }> = [
  {
    name: "来源",
    values: [
      ["甲", 1],
      ["乙", 2],
      ["丙", 3],
      ["丁", 4],
      ["戊", 5]
    ]
  },
  { name: "结果1", values: [["x", 100]] },
  { name: "结果2", values: [["y", 100]] },
  { name: "结果3", values: [["z", 100]] }
];
const ALL_NAMES = SHEETS.map((sheet) => sheet.name);

function host() {
  const created = createOfficeTestHost({
    url: "file:///C:/Data/公式来源证据验收.xlsx",
    activeSheet: "来源",
    sheets: SHEETS
  });
  vi.stubGlobal("Excel", created.Excel);
  return created;
}

/** 模型按工具说明产出的参数形态（无表头数字区域 + 绑定了工作表）。 */
function sumRequest(
  overrides: Partial<DataToolRequest["arguments"]> = {}
): DataToolRequest {
  return {
    id: "scope-sum",
    tool: "query_table",
    arguments: {
      mode: "aggregate",
      scope: "selected",
      hasHeaders: false,
      metrics: [{ operation: "sum", field: "列2", outputName: "合计" }],
      limit: 10,
      ...overrides
    }
  };
}

describe("本次查询范围：只查被绑定的那张表", () => {
  it("range 带表名：合计 15，来源只有「来源」，完整且无截断、零写入", async () => {
    const office = host();
    const result = await executeQueryTableTool(
      sumRequest({ range: "来源!A1:B5" }),
      ALL_NAMES,
      "来源"
    );

    expect(result.rows).toEqual([[15]]);
    // 这条是缺陷的核心：结果卡上「来源：」后面只能有被要求的那张表。
    expect(result.sourceSheets).toEqual(["来源"]);
    // 没被要求的表不该被扫，也就不会冒出"缺字段"这种告警。
    expect(result.warnings).toEqual([]);
    expect(result.incomplete).toEqual([]);
    expect(result.complete).toBe(true);
    expect(result.truncated).toBe(false);
    expect(result.truncation).toBeNull();
    expect(result.scannedRows).toBe(5);
    expect(office.writes()).toEqual([]);
  });

  it("整片授权区都查才会是 315 —— 说明 15 不是「凑巧只读到一张表」", async () => {
    host();
    const result = await executeQueryTableTool(
      sumRequest({ sheets: ALL_NAMES }),
      ALL_NAMES,
      "来源"
    );
    // 15 + 100 + 100 + 100：四张表的第二列全部计入。与上一条构成对照。
    expect(result.rows).toEqual([[315]]);
    expect(result.sourceSheets).toEqual(ALL_NAMES);
  });

  it("sheets 绑定 + 区域不带表名：同一个矩形只落到那张表上", async () => {
    host();
    const result = await executeQueryTableTool(
      sumRequest({ sheets: ["来源"], range: "A1:B5" }),
      ALL_NAMES,
      "来源"
    );
    expect(result.rows).toEqual([[15]]);
    expect(result.sourceSheets).toEqual(["来源"]);
    expect(result.complete).toBe(true);
  });

  it("多张可查表却没绑定：拒绝并列出可选表名，不扫任何一张表", async () => {
    const office = host();
    const failure = await executeQueryTableTool(
      sumRequest(),
      ALL_NAMES,
      "来源"
    ).catch((reason: unknown) => reason);

    expect(failure).toBeInstanceOf(DataToolExecutionError);
    const error = failure as DataToolExecutionError;
    // 参数问题（模型改一个字段就能过），不是越权。
    expect(error.code).toBe("UNSUPPORTED_TOOL");
    expect(error.retryable).toBe(true);
    for (const name of ALL_NAMES) expect(error.message).toContain(`「${name}」`);
    expect(error.message).toContain("授权范围只是可访问的上限");
    expect(office.writes()).toEqual([]);
  });

  it("区域指的表与 sheets 不一致：拒绝，不替调用方挑一个", async () => {
    host();
    const failure = await executeQueryTableTool(
      sumRequest({ sheets: ["结果1"], range: "来源!A1:B5" }),
      ALL_NAMES,
      "来源"
    ).catch((reason: unknown) => reason);

    expect((failure as DataToolExecutionError).code).toBe("SCOPE_VIOLATION");
    expect((failure as DataToolExecutionError).message).toContain("不一致");
  });

  it("跨表确实要读时：列全了就都读，缺字段的表如实记账", async () => {
    host();
    const result = await executeQueryTableTool(
      sumRequest({ sheets: ALL_NAMES }),
      ALL_NAMES,
      "来源"
    );
    // 四张表都有「列2」，所以这里是 315 而不是"缺字段"：
    // 缺字段那条路径由下一条用例单独钉住（那张表真的没有该列）。
    expect(result.warnings).toEqual([]);
    expect(result.complete).toBe(true);
  });
});

describe("不完整与截断来自真实原因", () => {
  it("某张表缺请求的字段：complete=false 且原因写明是哪张表缺哪一列", async () => {
    const office = createOfficeTestHost({
      url: "file:///C:/Data/缺列.xlsx",
      activeSheet: "来源",
      sheets: [
        { name: "来源", values: [["甲", 1], ["乙", 2]] },
        // 这一张只有一列：请求的「列2」在这里不存在。
        { name: "结果1", values: [["target-1"]] }
      ]
    });
    vi.stubGlobal("Excel", office.Excel);

    const result = await executeQueryTableTool(
      sumRequest({ sheets: ["来源", "结果1"] }),
      ["来源", "结果1"],
      "来源"
    );

    expect(result.complete).toBe(false);
    expect(result.incomplete).toEqual([
      { kind: "missing_field", sheet: "结果1", fields: ["列2"] }
    ]);
    // 告警文案与结构化原因是同一次判断的两种投影，不是各算一遍。
    expect(result.warnings).toEqual(["「结果1」未找到字段：列2"]);
    // 数据完整性的问题不该被说成截断。
    expect(result.truncated).toBe(false);
  });

  it("展示截断只由 limit 产生，且不改变 complete", async () => {
    host();
    const result = await executeQueryTableTool(
      sumRequest({ range: "来源!A1:B5", limit: 2 }),
      ALL_NAMES,
      "来源"
    );

    // aggregate 只有一个分组，limit=2 切不到东西。
    expect(result.truncated).toBe(false);

    const rows = await executeQueryTableTool(
      {
        id: "scope-rows",
        tool: "query_table",
        arguments: {
          mode: "rows",
          scope: "selected",
          hasHeaders: false,
          range: "来源!A1:B5",
          limit: 2
        }
      },
      ALL_NAMES,
      "来源"
    );
    expect(rows.rows).toHaveLength(2);
    expect(rows.complete).toBe(true);
    expect(rows.truncated).toBe(true);
    expect(rows.truncation).toEqual({ shownRows: 2, totalRows: 5, limit: 2 });
    expect(rows.incomplete).toEqual([]);
  });

  it("超限仍然整条请求报错：不产出「部分结果」，也就不会有卡片说扫描达上限", async () => {
    /*
     * 用**列数**上限触发同一个限额闸（行数上限要铺 25 万行才能碰到，
     * 在替身宿主里造那张矩阵只会把测试拖死，判别点却完全一样：超限在执行前
     * 抛 `TOOL_LIMIT_EXCEEDED`，一条结果都不产出）。
     */
    const wide: CellValue[][] = [
      Array.from({ length: 101 }, (_, index) => `列${index + 1}`),
      Array.from({ length: 101 }, (_, index) => index)
    ];
    const office = createOfficeTestHost({
      url: "file:///C:/Data/超宽.xlsx",
      activeSheet: "来源",
      sheets: [{ name: "来源", values: wide }]
    });
    vi.stubGlobal("Excel", office.Excel);

    const failure = await executeQueryTableTool(
      // 不给区域：限额检查看的是"要读的那块已用区域"（101 列）。
      sumRequest({ sheets: ["来源"] }),
      ["来源"],
      "来源"
    ).catch((reason: unknown) => reason);

    expect(failure).toBeInstanceOf(DataToolExecutionError);
    expect((failure as DataToolExecutionError).code).toBe("TOOL_LIMIT_EXCEEDED");
    expect((failure as DataToolExecutionError).message).toContain("超过本地工具的");
    expect(office.writes()).toEqual([]);
  });
});
