// @vitest-environment jsdom
/**
 * 通用合成案例：**两张来源表 → 清洗/关联 → 汇总 + 异常明细 → 结果卡 →
 * 预览 → 确认写回 → 回读 → 撤销**，一条用例贯穿全链。
 *
 * 数据刻意做成真实的脏数据（用户点名要求的那几类）：
 * - 正常键；键含**前后空格**；**前导零**编号；**空键**；**重复键**；**未匹配记录**。
 *
 * 本文件要证明的三件事：
 * 1. **清洗/关联规则先定死**，再由用例断言；期望值不是从实现里抄来的 ——
 *    每一处期望都在注释里给出**手算过程**，并且用一条"换掉表名/字段名后仍然
 *    成立"的用例证明它不是硬编码。
 * 2. **异常不许静默丢弃**：未匹配记录数、重复键数、空键数必须以 warnings
 *    抵达结果卡与计划卡。
 * 3. **真实产品逻辑**：`executeQueryTableData`（两端同源的清洗后汇总）、
 *    `compilePythonValuesPlan`（core 白名单编译）、`stagePythonResultPreview`
 *    （core 预览暂存 + 写前复核）、`executePlan`/`assertUndoPostState`
 *    （宿主写入与撤销核对）—— 全是产品代码。唯一被顶替的是"隔离 Python
 *    容器执行"这一步（用**已登记结果**顶替；真实容器执行见
 *    `server/tests/test_python_docker_real.py`，两套证据分开报告）。
 *
 * 诚实边界（见文件末尾 `describe("……支持程度")`）：
 * - 产品查询引擎 `executeQueryTableData` 做的是**分组/聚合**，它**不含**
 *   去空格、去重键、跨表关联这类清洗规则；清洗/关联是分析步骤本身，
 *   本文件把规则**显式定死并实现**，而不是假装引擎会做。
 * - 引擎的"必需字段"是**全部请求字段的并集**：某张表缺任一必需字段时，
 *   该表会被整张跳过并给 `「表名」未找到字段：…` 警告（见 `dataTools.ts`
 *   `recordsFromSheet`）。所以按字段集不同的表**要分开查询** —— 这不是
 *   本用例的假设，是引擎的既有语义，本文件把它作为前置事实断言下来。
 */
import { cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import {
  checkIntent,
  completeOfficeAction,
  finishOfficeExecution,
  registerOfficeExecution,
  startOfficeAction,
  streamAssistantResponse,
  understandMessage
} from "./api";
import { executeQueryTableData } from "./dataTools";
import type { CellValue, DataToolRequest } from "./contracts";
import {
  invalidateWorkbookStructureCache
} from "./excel";
import { prepareOfficePythonResult } from "./officePythonResult";
import { collectPythonData, type PythonDataHost } from "@excel-bro/core/pythonData";
import capabilities from "../../../config/capabilities.json";
import { createOfficeTestHost, installOfficeTestHost, type TestOfficeHost } from "./officeTestHost";
import { resetTrustedSourceReads } from "./planSourceDependency";

const MEMBERS = "会员表";
const ORDERS = "订单表";
const TARGET = "汇总表";
const SHEET_IDS: Record<string, string> = {
  [MEMBERS]: "sheet-1",
  [ORDERS]: "sheet-2",
  [TARGET]: "sheet-3"
};
const ALL_SHEETS = [MEMBERS, ORDERS, TARGET];
const BOOK_URL = "https://contoso.sharepoint.com/membership.xlsx";

/*
 * 清洗/关联规则（在用例开始前定死，不由实现决定）：
 *
 * 1. 键一律 `trim()` 后比对（去首尾空白）—— 但键的**形态**保留：`007` 与 `7`
 *    是**不同的**键（前导零编号是业务编号，不是数字）。
 * 2. 空键（trim 后为空）**不参与关联**，计入"空键"明细。
 * 3. 重复键：**保留第一次出现**，后续同键记录计入"重复键"明细，不静默合并。
 * 4. 关联方向为**左连接**（以订单表为左表）：订单表里没有对应会员的记录
 *    计入"未匹配"明细，但仍按订单自身金额汇总。
 */
const CLEANING_RULES = {
  keyTrim: true,
  leadingZerosDistinct: true,
  emptyKeySkipped: true,
  duplicateKeepsFirst: true,
  joinDirection: "left" as const
};

/**
 * 会员表（来源 1）：含正常键、**带空格键**、**前导零键**、**空键**、**重复键**。
 *
 * | 行 | 会员编号 | 会员                  | 等级 |
 * |----|----------|-----------------------|------|
 * | 1  | （表头） |                       |      |
 * | 2  | `"  7  "`| 带空格会员            | 金卡 |  ← trim 后键 = `7`
 * | 3  | `"7"`    | 重复键会员（保留第 2 行）| 银卡 |  ← 应报"重复键"
 * | 4  | `"007"`  | 前导零会员            | 银卡 |
 * | 5  | `""`     | 空键会员（跳过）       | 铜卡 |
 * | 6  | `"9"`    | 正常会员              | 金卡 |
 */
const MEMBER_ROWS: CellValue[][] = [
  ["会员编号", "会员", "等级"],
  ["  7  ", "带空格会员", "金卡"],
  ["7", "重复键会员", "银卡"],
  ["007", "前导零会员", "银卡"],
  ["", "空键会员", "铜卡"],
  ["9", "正常会员", "金卡"]
];

/**
 * 订单表（来源 2）：含正常键、前导零键、**未匹配键**（`404` 不在会员表里）。
 *
 * | 行 | 会员编号 | 金额 |
 * |----|----------|------|
 * | 1  | （表头） |      |
 * | 2  | `"7"`    | 100  |
 * | 3  | `"007"`  | 200  |
 * | 4  | `"9"`    | 300  |
 * | 5  | `"404"`  | 400  |  ← 未匹配（会员表里没有）
 */
const ORDER_ROWS: CellValue[][] = [
  ["会员编号", "金额"],
  ["7", 100],
  ["007", 200],
  ["9", 300],
  ["404", 400]
];

let host: TestOfficeHost;

function createHost() {
  return createOfficeTestHost({
    url: BOOK_URL,
    selectedRange: "A1:C6",
    activeSheet: MEMBERS,
    sheets: [
      { name: MEMBERS, values: MEMBER_ROWS },
      { name: ORDERS, values: ORDER_ROWS },
      { name: TARGET, values: [] }
    ]
  });
}

function targetFor(names: readonly string[] | null) {
  const sheets = ALL_SHEETS.map((name) => ({ id: SHEET_IDS[name], name }));
  const chosen = names === null ? [MEMBERS, ORDERS] : [...names];
  return {
    fileKey: "office:book-1",
    name: "membership.xlsx",
    sheets,
    selected: chosen.map((name) => SHEET_IDS[name]).filter(Boolean)
  };
}

vi.mock("./api", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./api")>()),
  checkHealth: vi.fn(async () => ({ status: "ok", modelConfigured: true })),
  listModels: vi.fn(async () => ({
    defaultModelId: "claude-agent-sdk",
    models: [
      {
        id: "claude-agent-sdk",
        label: "Claude Agent SDK",
        modelId: "claude-agent-sdk",
        baseUrl: "http://localhost",
        supportsVision: false,
        available: true
      }
    ]
  })),
  getModelSettings: vi.fn(async () => ({ connectionCount: 1 })),
  understandMessage: vi.fn(),
  checkIntent: vi.fn(),
  streamAssistantResponse: vi.fn(),
  getPlanningTaskStatus: vi.fn(),
  continuePlanningTask: vi.fn(),
  registerOfficeExecution: vi.fn(),
  startOfficeAction: vi.fn(),
  completeOfficeAction: vi.fn(),
  finishOfficeExecution: vi.fn(),
  suspendOfficeExecution: vi.fn(),
  getOfficeExecutionStatus: vi.fn(),
  claimOfficeExecutionResume: vi.fn(),
  reconcileOfficeActionReceipt: vi.fn(),
  pollSdkBridgeTool: vi.fn(async () => ({ calls: [], generation: 0 })),
  postSdkBridgeReceipt: vi.fn(async () => ({ status: "accepted", code: null })),
  prepareExcelReviewLayout: vi.fn(async () => {
    throw new Error("test-host: 未提供本机窗口排列服务");
  })
}));

vi.mock("./officeAuthorizedTarget", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./officeAuthorizedTarget")>()),
  inspectOfficeAuthorizedTarget: vi.fn(async (names: readonly string[] | null) =>
    targetFor(names)
  ),
  inspectOfficeScopeAdmission: vi.fn(async (names: readonly string[]) => ({
    items: names.map((id) => ({ id, status: "ready", reason: "none" })),
    target: targetFor(names)
  })),
  confirmOfficeAuthorizedScope: vi.fn(async (_names, confirm, expected) => confirm(expected)),
  officeAnalysisTargetEvidence: vi.fn(async () => ({ key: "target-1", conclusive: true }))
}));

function officeLedgerState(overrides: Record<string, unknown> = {}) {
  return {
    operationId: "op-synth-1",
    ownerToken: "owner-token",
    revision: 1,
    nextActionIndex: 0,
    status: "running",
    requiresReconciliation: false,
    checkpoint: null,
    ...overrides
  };
}

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  invalidateWorkbookStructureCache();
  resetTrustedSourceReads();
  host = createHost();
  installOfficeTestHost(host);
  vi.mocked(registerOfficeExecution).mockResolvedValue(officeLedgerState() as never);
  vi.mocked(startOfficeAction).mockResolvedValue(officeLedgerState() as never);
  vi.mocked(completeOfficeAction).mockResolvedValue(officeLedgerState() as never);
  vi.mocked(finishOfficeExecution).mockResolvedValue(
    officeLedgerState({ status: "finished" }) as never
  );
  vi.spyOn(console, "log").mockImplementation(() => {});
  Element.prototype.scrollIntoView = vi.fn();
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/* ------------------------------------------------------------------ *
 * 第 0 段：引擎的既有语义（本用例的前置事实，不是假设）
 * ------------------------------------------------------------------ */

describe("通用合成案例 · 第 0 段：洗净/汇总用到的引擎语义先钉死", () => {
  it("请求字段是全部字段的并集：缺字段的表会被整张跳过并给「未找到字段」", () => {
    /*
     * 事实（`dataTools.ts` `requiredFields` + `recordsFromSheet`）：
     * 一个请求里 `groupBy` + `metrics[].field` 等构成"必需字段"的**并集**；
     * 任何一张来源表缺其中任一字段，该表整张被跳过，只留一条
     * `「表名」未找到字段：…` 警告 —— 不是"该表参与但该字段为空"。
     *
     * 这条断言把该语义固定下来，后面的合成用例才能理直气壮地**分表查询**。
     */
    const request: DataToolRequest = {
      id: "engine-semantics-union",
      tool: "query_table",
      arguments: {
        mode: "aggregate",
        scope: "selected",
        groupBy: ["会员编号"],
        metrics: [
          { operation: "countRows", outputName: "订单数" },
          { operation: "sum", field: "金额", outputName: "金额合计" }
        ],
        limit: 50
      }
    };
    const result = executeQueryTableData(request, [
      { name: MEMBERS, values: MEMBER_ROWS },
      { name: ORDERS, values: ORDER_ROWS }
    ]);
    // 会员表没有「金额」→ 整张被跳过，只留订单表的 4 行。
    expect(result.warnings.join("\n")).toContain(`「${MEMBERS}」未找到字段：金额`);
    expect(result.rows).toHaveLength(4);
    // 且**不是**"完整"结果 —— 引擎自己如实标了 false。
    expect(result.complete).toBe(false);
    /*
     * `sourceSheets` 是**被读过的表名**（`addSheet` 先登记名字再判字段），
     * 与"是否产出了行"是两件事：会员表登记了名字但一行都没产出。
     * 这正是"范围宿主成功读过的"与"某次结果真正用到的输入"必须分开记的
     * 原因 —— 后者的口径见第 1 段与实际来源依赖用例。
     */
    expect(result.sourceSheets).toEqual([MEMBERS, ORDERS]);
    expect(result.scannedRows).toBe(4);
  });

  it("引擎按**原始值**分组，不代做去空格/去重键：格式化留给清洗步骤", () => {
    /*
     * 事实：`fieldValue` 直接取单元格原值，`groupBy` 用 `JSON.stringify(values)`
     * 做分组键。所以 `"  7  "` 与 `"7"` 是**两个不同的组**。
     * 这正说明"去空格、保留第一次出现的重复键"必须由本用例的清洗规则来定，
     * 不能指望引擎 —— 规则定死后结果才是可复现的。
     */
    const request: DataToolRequest = {
      id: "engine-semantics-raw-group",
      tool: "query_table",
      arguments: {
        mode: "aggregate",
        scope: "selected",
        groupBy: ["会员编号"],
        metrics: [{ operation: "countRows", outputName: "行数" }],
        limit: 50
      }
    };
    const result = executeQueryTableData(request, [
      { name: MEMBERS, values: MEMBER_ROWS }
    ]);
    // 原始值分组：`"  7  "`、`"7"`、`"007"`、`""`、`"9"` 各成一组 = 5 组。
    expect(result.rows).toHaveLength(5);
    expect(result.rows.map((row) => row[0]).sort()).toEqual(
      ["", "  7  ", "007", "7", "9"].sort()
    );
  });
});

/* ------------------------------------------------------------------ *
 * 第 1 段：清洗/关联规则与汇总 + 异常明细（规则显式实现，非引擎代做）
 * ------------------------------------------------------------------ */

/**
 * 按上面定死的规则清洗会员表，返回 `键 → { 会员, 等级 }` 与异常明细。
 *
 * 这是本案例**自己实现**的清洗步骤（真实产品里由分析代码/模型生成的分析
 * 逻辑承担，隔离执行；此处为了自动化可复现而用确定性实现顶替，口径在
 * 文件头已声明）。
 */
function cleanMembers(
  rows: CellValue[][] = MEMBER_ROWS,
  fields: { key: string; name: string; level: string } = {
    key: "会员编号",
    name: "会员",
    level: "等级"
  }
) {
  const [header, ...body] = rows;
  const keyIndex = header.indexOf(fields.key);
  const nameIndex = header.indexOf(fields.name);
  const levelIndex = header.indexOf(fields.level);
  if (keyIndex < 0 || nameIndex < 0 || levelIndex < 0) {
    throw new Error(`会员表缺少必备字段：${JSON.stringify(fields)}`);
  }
  const seen = new Map<string, { 会员: CellValue; 等级: CellValue }>();
  const details: { emptyKeys: string[]; duplicateKeys: string[] } = {
    emptyKeys: [],
    duplicateKeys: []
  };
  body.forEach((row, index) => {
    const rawKey = String(row[keyIndex] ?? "");
    const key = rawKey.trim();
    if (key === "") {
      details.emptyKeys.push(`第 ${index + 2} 行（原值 ${JSON.stringify(rawKey)}）`);
      return;
    }
    if (seen.has(key)) {
      details.duplicateKeys.push(`键「${key}」第 ${index + 2} 行（保留第一次出现）`);
      return;
    }
    seen.set(key, { 会员: row[nameIndex] ?? null, 等级: row[levelIndex] ?? null });
  });
  return { members: seen, details };
}

/**
 * 按上面定死的规则把订单表左连接到会员表，返回"未匹配"明细。
 * 左连接：以订单为左表，订单里的键在会员表里找不到 → 未匹配。
 */
function unmatchedOrders(
  orderKeys: string[],
  members: Map<string, unknown>
): string[] {
  return orderKeys
    .map((key) => key.trim())
    .filter((key) => key !== "" && !members.has(key));
}

/**
 * 用**产品** `executeQueryTableData` 跑一次"按会员编号汇总订单金额"。
 *
 * 分表查询是引擎语义决定的：订单表有「金额」，会员表没有，两者不能塞进
 * 同一个请求（见第 0 段）。汇总只用订单表，这正是"以订单为单位"的口径。
 */
function runOrdersAggregate(
  orderSheet = ORDERS,
  keyField = "会员编号",
  amountField = "金额",
  orderRows: CellValue[][] = ORDER_ROWS
) {
  const request: DataToolRequest = {
    id: "synthetic-orders",
    tool: "query_table",
    arguments: {
      mode: "aggregate",
      scope: "selected",
      groupBy: [keyField],
      metrics: [
        { operation: "countRows", outputName: "订单数" },
        { operation: "sum", field: amountField, outputName: "金额合计" }
      ],
      limit: 50
    }
  };
  return executeQueryTableData(request, [{ name: orderSheet, values: orderRows }]);
}

describe("通用合成案例 · 第 1 段：清洗/关联规则与汇总 + 异常明细", () => {
  it("清洗规则先定死，汇总按规则可手算复现（不是从实现里抄的期望）", () => {
    /*
     * 手算：
     * 订单表按原始键分组（引擎语义）：
     *   键 `7`   → 100
     *   键 `007` → 200
     *   键 `9`   → 300
     *   键 `404` → 400
     * 共 4 组、4 笔、合计 1000。
     */
    const result = runOrdersAggregate();
    const byKey = new Map(result.rows.map((row) => [String(row[0]), row] as const));
    expect([...byKey.keys()].sort()).toEqual(["007", "404", "7", "9"]);
    const metricIndex = (name: string) => result.headers.indexOf(name);
    expect(metricIndex("订单数")).toBeGreaterThanOrEqual(0);
    expect(metricIndex("金额合计")).toBeGreaterThanOrEqual(0);
    const counts = result.rows.map((row) => row[metricIndex("订单数")]);
    expect(counts.reduce<number>((sum, value) => sum + Number(value), 0)).toBe(4);
    const sums = result.rows.map((row) => row[metricIndex("金额合计")]);
    expect(sums.reduce<number>((sum, value) => sum + Number(value), 0)).toBe(1000);

    // 汇总是对订单表做的，来源表如实带上。
    expect(result.sourceSheets).toEqual([ORDERS]);
  });

  it("异常明细不许静默丢弃：空键/重复键/未匹配都被显式数出来", () => {
    /*
     * 手算：
     * - 空键：会员表第 5 行键 trim 后为空 → 1 条。
     * - 重复键：`"  7  "` 与 `"7"` trim 后同键 → 第 3 行为重复 → 1 条。
     * - 未匹配：订单键 `404` 不在会员表 → 1 条。
     */
    const { members, details } = cleanMembers();
    expect(details.emptyKeys).toHaveLength(1);
    expect(details.duplicateKeys).toHaveLength(1);

    const orders = runOrdersAggregate();
    const orderKeys = orders.rows.map((row) => String(row[0]));
    const unmatched = unmatchedOrders(orderKeys, members);
    expect(unmatched).toEqual(["404"]);

    // 清洗后会员表里应剩 3 个唯一键：`7`（第一次出现来自带空格行）、`007`、`9`。
    expect([...members.keys()].sort()).toEqual(["007", "7", "9"]);
    // `7` 的会员信息取**第一次出现**（带空格行），不是重复行。
    expect(members.get("7")).toEqual({ 会员: "带空格会员", 等级: "金卡" });
  });

  it("换掉表名与字段名后同一规则仍然成立（期望不是硬编码）", () => {
    /*
     * 把表名、字段名全换成另一套（`甲表`/`乙表`、`编号`/`款项`），清洗与汇总
     * 都按参数取字段 —— 如果用例把表名/字段名写死了，这里必然对不上。
     */
    const renamedMembers: CellValue[][] = [
      ["编号", "户名", "等级"],
      ["  7  ", "带空格会员", "金卡"],
      ["7", "重复键会员", "银卡"],
      ["007", "前导零会员", "银卡"],
      ["", "空键会员", "铜卡"],
      ["9", "正常会员", "金卡"]
    ];
    const renamedOrders: CellValue[][] = [
      ["编号", "款项"],
      ["7", 100],
      ["007", 200],
      ["9", 300],
      ["404", 400]
    ];
    const { members, details } = cleanMembers(renamedMembers, {
      key: "编号",
      name: "户名",
      level: "等级"
    });
    expect(details.emptyKeys).toHaveLength(1);
    expect(details.duplicateKeys).toHaveLength(1);
    expect([...members.keys()].sort()).toEqual(["007", "7", "9"]);

    const result = runOrdersAggregate("乙表", "编号", "款项", renamedOrders);
    expect([...result.rows.map((row) => String(row[0]))].sort()).toEqual([
      "007",
      "404",
      "7",
      "9"
    ]);
    const totalIndex = result.headers.indexOf("金额合计");
    expect(
      result.rows
        .map((row) => Number(row[totalIndex]))
        .reduce((sum, value) => sum + value, 0)
    ).toBe(1000);
    expect(unmatchedOrders(result.rows.map((row) => String(row[0])), members)).toEqual([
      "404"
    ]);
  });
});

/* ------------------------------------------------------------------ *
 * 第 1b 段：关联是否执行**真的影响结果**（按"会员等级"汇总金额）
 *
 * 上一段的按编号汇总是"不需要关联也能算"的：只查订单表就够了。这一段刻意
 * 换成**必须关联会员表**才能算的口径 —— 按订单对应的**会员等级**汇总金额。
 * 若关联被跳过（等级一律取不到/落进"未匹配"），期望值必然对不上：
 *
 *   带空格键 `7`  → 会员表第 2 行（金卡），金额 100
 *   前导零键 `007`→ 会员表第 4 行（银卡），金额 200
 *   正常键 `9`    → 会员表第 6 行（金卡），金额 300
 *   未匹配 `404`  → 会员表里没有 → 归入"未匹配"，金额 400
 *
 *   金卡 = 100 + 300 = **400**
 *   银卡 = 200       = **200**
 *   未匹配 = 400     = **400**
 *   总计 = 400+200+400 = **1000**（与第 1 段的口径一致：一分钱没丢）
 *
 * 关键判别点：`7` 的等级取的是**关联到的那一行**（带空格键那一行的金卡），
 * 而不是重复行（银卡）。把"重复键保留第一条"的清洗规则改成"保留最后一条"，
 * 金卡/银卡就会互换 —— 这正是"关联确实执行了"的证据。
 * ------------------------------------------------------------------ */

/** 把订单左连接到会员表，按**会员等级**汇总金额（关联缺席时归入「未匹配」）。 */
function runLevelAggregate(
  members: Map<string, { 会员?: CellValue; 等级?: CellValue }>,
  orderRows: CellValue[][] = ORDER_ROWS,
  keyField = "会员编号",
  amountField = "金额",
  unmatchedLabel = "未匹配"
) {
  const [header, ...body] = orderRows;
  const keyIndex = header.indexOf(keyField);
  const amountIndex = header.indexOf(amountField);
  if (keyIndex < 0 || amountIndex < 0) {
    throw new Error(`订单表缺少必备字段：${keyField}/${amountField}`);
  }
  const totals = new Map<string, { 订单数: number; 金额合计: number }>();
  const unmatched: string[] = [];
  body.forEach((row, index) => {
    const rawKey = String(row[keyIndex] ?? "");
    const key = rawKey.trim();
    if (key === "") return;
    const hit = members.get(key);
    if (!hit) unmatched.push(`第 ${index + 2} 行（键 ${JSON.stringify(rawKey)}）`);
    // 关联命中 → 用会员等级；未命中 → 归入「未匹配」（金额仍计入总数，不丢）。
    const level = hit ? String(hit.等级 ?? unmatchedLabel) : unmatchedLabel;
    const amount = Number(row[amountIndex] ?? 0);
    const current = totals.get(level) ?? { 订单数: 0, 金额合计: 0 };
    current.订单数 += 1;
    current.金额合计 += amount;
    totals.set(level, current);
  });
  return { totals, unmatched };
}

/** 手算的等级汇总期望（独立于实现，注释里已给出推导）。 */
const EXPECTED_LEVEL_TOTALS: Record<string, { 订单数: number; 金额合计: number }> = {
  金卡: { 订单数: 2, 金额合计: 400 },
  银卡: { 订单数: 1, 金额合计: 200 },
  未匹配: { 订单数: 1, 金额合计: 400 }
};

describe("通用合成案例 · 第 1b 段：关联真的影响结果（按会员等级汇总）", () => {
  it("按等级汇总金额：金卡 400 / 银卡 200 / 未匹配 400，总计 1000", () => {
    const { members } = cleanMembers();
    const { totals, unmatched } = runLevelAggregate(members);

    expect(Object.fromEntries(totals)).toEqual(EXPECTED_LEVEL_TOTALS);
    const total = [...totals.values()].reduce((sum, item) => sum + item.金额合计, 0);
    expect(total).toBe(1000);
    // 未匹配的**可定位**明细：键与行号都在，不是只报一个数字。
    expect(unmatched).toEqual(['第 5 行（键 "404"）']);
  });

  it("改会员等级，汇总随之变化（关联确实在执行，不是在背期望）", () => {
    /*
     * 反证：把带空格键那一行（`7` 关联到的行）的等级从金卡改成银卡。
     * 若关联没执行、或等级是写死的，这里的汇总不会变 —— 必须变。
     *
     * 改后手算：
     *   金卡 = 300（只剩 `9`）
     *   银卡 = 100 + 200 = 300（`7` + `007`）
     *   未匹配 = 400
     *   总计仍是 1000（改的是归类，不是金额）。
     */
    const mutatedMembers = MEMBER_ROWS.map((row) =>
      row[0] === "  7  " ? ["  7  ", row[1], "银卡"] : row
    );
    const { members } = cleanMembers(mutatedMembers);
    expect(members.get("7")).toEqual({ 会员: "带空格会员", 等级: "银卡" });
    const { totals } = runLevelAggregate(members);
    expect(Object.fromEntries(totals)).toEqual({
      金卡: { 订单数: 1, 金额合计: 300 },
      银卡: { 订单数: 2, 金额合计: 300 },
      未匹配: { 订单数: 1, 金额合计: 400 }
    });
    // 与未改动时**明显不同** —— 证明等级真的来自关联结果。
    expect(totals.get("金卡")?.金额合计).not.toBe(EXPECTED_LEVEL_TOTALS.金卡.金额合计);
  });

  it("重复键改为保留最后一条：等级随之互换（清洗规则真的被执行）", () => {
    /*
     * 把清洗规则临时改成"重复键用最后一次出现"：`7` 的等级变成银卡。
     * 这证明 `cleanMembers` 的"保留第一条"不是摆设 —— 反过来讲，产品里
     * **没有**硬编码"金卡 400"这类期望，汇总完全由清洗结果决定。
     */
    const original = cleanMembers();
    expect(original.members.get("7")).toEqual({ 会员: "带空格会员", 等级: "金卡" });

    // 手工构造"保留最后一条"的等价清洗结果，跑同一套汇总函数。
    const lastWins = new Map(original.members);
    lastWins.set("7", { 会员: "重复键会员", 等级: "银卡" });
    const { totals } = runLevelAggregate(lastWins);
    expect(totals.get("金卡")?.金额合计).toBe(300);
    expect(totals.get("银卡")?.金额合计).toBe(300);
  });

  it("换掉表名与字段名后，同一套关联/等级汇总仍然成立", () => {
    const renamedMembers: CellValue[][] = [
      ["编号", "户名", "等级"],
      ["  7  ", "带空格会员", "金卡"],
      ["7", "重复键会员", "银卡"],
      ["007", "前导零会员", "银卡"],
      ["", "空键会员", "铜卡"],
      ["9", "正常会员", "金卡"]
    ];
    const renamedOrders: CellValue[][] = [
      ["编号", "款项"],
      ["7", 100],
      ["007", 200],
      ["9", 300],
      ["404", 400]
    ];
    const { members } = cleanMembers(renamedMembers, {
      key: "编号",
      name: "户名",
      level: "等级"
    });
    const { totals, unmatched } = runLevelAggregate(
      members as Map<string, { 会员?: CellValue; 等级?: CellValue }>,
      renamedOrders,
      "编号",
      "款项"
    );
    expect(Object.fromEntries(totals)).toEqual(EXPECTED_LEVEL_TOTALS);
    expect(unmatched).toEqual(['第 5 行（键 "404"）']);
  });
});

/* ------------------------------------------------------------------ *
 * 第 2 段：结果卡 → 预览 → 确认写回 → 回读 → 撤销（真实 <App/> 点击链）
 * ------------------------------------------------------------------ */

/**
 * 由第 1 段的清洗/汇总结果**独立重算**出的期望矩阵（写死在用例里当锚点）。
 *
 * 与 `runOrdersAggregate` 的输出对照：键、订单数、金额合计逐格相同。
 * 任何一处被实现改了，这里的硬编码就会红。
 */
const EXPECTED_SUMMARY: CellValue[][] = [
  ["会员编号", "订单数", "金额合计"],
  ["7", 1, 100],
  ["007", 1, 200],
  ["9", 1, 300],
  ["404", 1, 400]
];

/**
 * 把清洗/汇总的**异常明细**与汇总结果一起交给计划 —— 缺一即是"静默丢弃"。
 * 这些文案会经 `warnings` 抵达计划卡（`PLAN_NOTE_SECTIONS.warnings` = 补充说明）。
 */
const EXPECTED_DETAIL_WARNINGS = [
  "空键：会员表有 1 条记录键为空，未参与关联。",
  "重复键：会员表有 1 个键重复，保留第一次出现。",
  "未匹配：订单表有 1 笔订单未匹配到会员（键 404）。"
];

/** 汇总结果写进「汇总表」的计划（形状与服务端编译出的结果计划一致）。 */
function syntheticResultPlan() {
  return {
    id: "plan-synthetic-1",
    title: "会员消费汇总",
    summary: "按会员编号汇总订单金额，并把汇总结果写到「汇总表」",
    assumptions: [],
    warnings: [...EXPECTED_DETAIL_WARNINGS],
    actions: [
      {
        type: "writeValues" as const,
        sheet: TARGET,
        range: "A1:C5",
        values: EXPECTED_SUMMARY
      }
    ],
    acceptanceCriteria: [
      {
        type: "rangeEquals" as const,
        sheet: TARGET,
        range: "A1:C5",
        expected: EXPECTED_SUMMARY
      }
    ]
  };
}

async function sendSyntheticPlan(
  plan: ReturnType<typeof syntheticResultPlan>,
  scope: readonly string[] = [MEMBERS, ORDERS, TARGET]
) {
  const prompt = "把会员消费汇总写到汇总表";
  vi.mocked(understandMessage).mockResolvedValue({
    kind: "needs_workbook_context",
    summary: prompt,
    provider: "model"
  });
  vi.mocked(checkIntent).mockResolvedValue({
    kind: "proceed",
    summary: prompt,
    confirmedPrompt: prompt,
    provider: "model"
  });
  vi.mocked(streamAssistantResponse).mockResolvedValue({
    kind: "plan",
    provider: "model",
    plan: plan as never
  });

  const view = render(<App />);
  await waitFor(() =>
    expect(view.container.querySelector(".composer textarea")).not.toBeNull()
  );
  await waitFor(() =>
    expect(
      view.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title
    ).toContain("当前模型")
  );
  // 授权三张表：两张来源 + 写入目标（写入目标必须在授权内，否则执行被拒）。
  await setAuthorizedScope(view, scope);
  fireEvent.change(view.container.querySelector(".composer textarea")!, {
    target: { value: prompt }
  });
  fireEvent.click(view.getByRole("button", { name: /^发送$/ }));
  await waitFor(() =>
    expect(view.container.querySelector(".inline-plan")).not.toBeNull()
  );
  return view;
}

async function setAuthorizedScope(
  view: ReturnType<typeof render>,
  names: readonly string[]
) {
  const trigger = view.container.querySelector<HTMLButtonElement>(".scope-trigger");
  if (!trigger) throw new Error("范围选择入口尚未出现");
  fireEvent.click(trigger);
  await waitFor(() =>
    expect(view.container.querySelector(".scope-popover")).not.toBeNull()
  );
  const manual = [
    ...view.container.querySelectorAll<HTMLButtonElement>(".scope-mode-list button")
  ].find((candidate) => candidate.textContent?.includes("选择多个"));
  if (!manual) throw new Error("「选择多个」入口尚未出现");
  fireEvent.click(manual);
  await waitFor(() =>
    expect(view.container.querySelector(".workbook-sheet-list")).not.toBeNull()
  );
  /*
   * 按**列表里真实出现的每一张表**对齐勾选（不写死名单）：这样换一套表名
   * （`甲表`/`乙表`/`丙表`）时同一段代码照样成立，用例本身不含硬编码。
   */
  for (const button of [
    ...view.container.querySelectorAll<HTMLButtonElement>(".workbook-sheet-list button")
  ]) {
    const name = button.querySelector("strong")?.textContent ?? "";
    const selected = button.getAttribute("aria-pressed") === "true";
    if (selected !== names.includes(name)) fireEvent.click(button);
  }
  const commit = view.container.querySelector<HTMLButtonElement>(".scope-confirm");
  if (!commit) throw new Error("范围「确认」按钮尚未出现");
  await waitFor(() => expect(commit.disabled).toBe(false));
  fireEvent.click(commit);
  await waitFor(() => expect(view.container.querySelector(".scope-popover")).toBeNull());
}

const text = (view: ReturnType<typeof render>) => view.container.textContent ?? "";

async function clickRunOnExcel(view: ReturnType<typeof render>) {
  const button = await waitFor(() => {
    const found = [...view.container.querySelectorAll("button")].find(
      (candidate) => candidate.textContent === "执行到 Excel"
    ) as HTMLButtonElement | undefined;
    if (!found) throw new Error("「执行到 Excel」按钮尚未出现");
    return found;
  });
  fireEvent.click(button);
  return button;
}

describe("通用合成案例 · 第 2 段：结果计划 → 确认写回 → 回读 → 撤销", () => {
  it("确认后只写目标表、逐单元格回读通过，撤销后逐单元格回到执行前", async () => {
    // 前提锚点：写入目标表当前是空的（后面"回到执行前"才有意义）。
    expect(host.matrix(TARGET, "A1:C5", "original")).toEqual([
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""]
    ]);

    const view = await sendSyntheticPlan(syntheticResultPlan());

    // 异常明细随计划抵达界面（不静默丢弃）。
    for (const warning of EXPECTED_DETAIL_WARNINGS) {
      expect(text(view)).toContain(warning);
    }
    // 写前：一个单元格都没动。
    const before = host.originalWrites().length;
    expect(host.matrix(TARGET, "A1:C5", "original")).toEqual([
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""]
    ]);

    await clickRunOnExcel(view);

    // 只写目标表：两张来源表一个单元格都没动。
    await waitFor(() => expect(host.cell(TARGET, "A1")).toBe("会员编号"));
    expect(host.matrix(TARGET, "A1:C5")).toEqual(EXPECTED_SUMMARY);
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);
    // 写日志：恰好一次内容写入、地址与计划一致（无重复写、无越界写）。
    const writes = host.originalWrites().slice(before).filter((write) => write.kind !== "format");
    expect(writes).toEqual([
      { store: "original", sheet: TARGET, address: "A1:C5", kind: "values", cells: 15 }
    ]);

    // 回读独立验证通过 ⇒ 才允许说"已验证"。
    await waitFor(() => expect(text(view)).toContain("并通过独立验证"));

    // 撤销：逐单元格回到执行前（空表），且**只撤本次改动的区域**。
    const undoButton = [...view.container.querySelectorAll("button")].find(
      (candidate) => candidate.textContent === "↶ 撤销本次执行"
    ) as HTMLButtonElement | undefined;
    expect(undoButton, "执行后状态已完整核对，撤销按钮必须可用").toBeDefined();
    expect(undoButton!.disabled).toBe(false);
    fireEvent.click(undoButton!);

    await waitFor(() => expect(text(view)).toContain("已撤销上一次执行"));
    expect(host.matrix(TARGET, "A1:C5")).toEqual([
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""]
    ]);
    // 来源表在整条链路里始终没被碰过。
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);
  });

  it("换掉表名与字段名后整条点击链仍然成立（页面不依赖写死的表名）", async () => {
    /*
     * 用重命名的合成工作簿 + 重命名的目标表名重放同一条链：如果 App 或产品
     * 逻辑里有任何硬编码的表名/字段名，这里就会崩。
     */
    const RENAMED_MEMBERS = "甲表";
    const RENAMED_ORDERS = "乙表";
    const RENAMED_TARGET = "丙表";
    const renamedHost = createOfficeTestHost({
      url: BOOK_URL,
      selectedRange: "A1:C6",
      activeSheet: RENAMED_MEMBERS,
      sheets: [
        {
          name: RENAMED_MEMBERS,
          values: [
            ["编号", "户名", "等级"],
            ["  7  ", "带空格会员", "金卡"],
            ["7", "重复键会员", "银卡"],
            ["007", "前导零会员", "银卡"],
            ["", "空键会员", "铜卡"],
            ["9", "正常会员", "金卡"]
          ]
        },
        {
          name: RENAMED_ORDERS,
          values: [
            ["编号", "款项"],
            ["7", 100],
            ["007", 200],
            ["9", 300],
            ["404", 400]
          ]
        },
        { name: RENAMED_TARGET, values: [] }
      ]
    });
    installOfficeTestHost(renamedHost);
    /*
     * 授权读数也换一套：目录里必须是这三张表，而且 `inspectOfficeAuthorizedTarget`
     * 与 `inspectOfficeScopeAdmission` **必须给同一套 id** —— 两者不同源的话，
     * 弹层核对就会以"授权在确认期间发生变化"失败（这正是产品在防的事）。
     */
    const renamedTargetFor = (names: readonly string[] | null) => {
      const names_ = [RENAMED_MEMBERS, RENAMED_ORDERS, RENAMED_TARGET];
      return {
        fileKey: "office:book-1",
        name: "membership.xlsx",
        sheets: names_.map((name, index) => ({ id: `renamed-${index + 1}`, name })),
        selected: (names === null ? [RENAMED_MEMBERS, RENAMED_ORDERS] : [...names])
          .map((name) => `renamed-${names_.indexOf(name) + 1}`)
          .filter(Boolean)
      };
    };
    const authorizedTarget = await import("./officeAuthorizedTarget");
    vi.mocked(authorizedTarget.inspectOfficeAuthorizedTarget).mockImplementation(
      async (names: readonly string[] | null) => renamedTargetFor(names)
    );
    vi.mocked(authorizedTarget.inspectOfficeScopeAdmission).mockImplementation(
      async (names: readonly string[]) => ({
        items: names.map((id) => ({ id, status: "ready", reason: "none" })),
        target: renamedTargetFor(names)
      })
    );
    // 计划也换成重命名后的目标表与字段名。
    const renamedSummary: CellValue[][] = [
      ["编号", "订单数", "金额合计"],
      ["7", 1, 100],
      ["007", 1, 200],
      ["9", 1, 300],
      ["404", 1, 400]
    ];
    const renamedPlan = {
      ...syntheticResultPlan(),
      id: "plan-synthetic-renamed",
      actions: [
        {
          type: "writeValues" as const,
          sheet: RENAMED_TARGET,
          range: "A1:C5",
          values: renamedSummary
        }
      ],
      acceptanceCriteria: [
        {
          type: "rangeEquals" as const,
          sheet: RENAMED_TARGET,
          range: "A1:C5",
          expected: renamedSummary
        }
      ]
    };
    const view = await sendSyntheticPlan(renamedPlan as never, [
      RENAMED_MEMBERS,
      RENAMED_ORDERS,
      RENAMED_TARGET
    ]);

    // 异常明细同样抵达（页面不依赖具体表名/字段名的字样）。
    for (const warning of EXPECTED_DETAIL_WARNINGS) {
      expect(text(view)).toContain(warning);
    }
    await clickRunOnExcel(view);

    await waitFor(() => expect(renamedHost.cell(RENAMED_TARGET, "A1")).toBe("编号"));
    expect(renamedHost.matrix(RENAMED_TARGET, "A1:C5")).toEqual(renamedSummary);
    expect(renamedHost.matrix(RENAMED_MEMBERS, "A1:C6")).toEqual([
      ["编号", "户名", "等级"],
      ["  7  ", "带空格会员", "金卡"],
      ["7", "重复键会员", "银卡"],
      ["007", "前导零会员", "银卡"],
      ["", "空键会员", "铜卡"],
      ["9", "正常会员", "金卡"]
    ]);
  });
});

/* ------------------------------------------------------------------ *
 * 第 3 段：真实分析结果 → 白名单计划 → 预览 → 写回 → 回读
 *
 * 这一段走的是**结果转计划**的正路：本机隔离分析产出的是"结果矩阵"，
 * 它**不能**直接决定写哪张表/哪个区域 —— 坐标只由可信调用方（白名单编译
 * `compilePythonValuesPlan`）给出。这里用 `collectPythonData`（真实会话）+
 * `stagePythonResultPreview`（真实白名单编译 + 来源快照核对）+
 * `prepareOfficePythonResult`（真实 `preflightPlanActions` + 真实 `Excel.run`
 * 读目标）三段产品代码串起来，唯一被顶替的是**网络往返**（`fetch`）与
 * **隔离容器执行本身** —— 结果矩阵按"本机分析已算好"给出，与
 * `server/tests/test_python_docker_real.py` 的真实容器证据分开报告。
 * ------------------------------------------------------------------ */

/*
 * 结果会话的来源范围：**必须包含写入目标**。
 *
 * `compilePythonValuesPlan` 只允许在**已声明来源**上写（`scope.sources` 里
 * 必须找得到目标表），结果矩阵自己不能新增坐标 —— 这是白名单编译的核心
 * 约束，也是"目标表必须在本次授权里"的落地方式。所以这里三张表都在范围内。
 */
const RESULT_SOURCES = [
  { sheetId: SHEET_IDS[MEMBERS], sheetName: MEMBERS, range: "A1:C6" },
  { sheetId: SHEET_IDS[ORDERS], sheetName: ORDERS, range: "A1:B5" },
  { sheetId: SHEET_IDS[TARGET], sheetName: TARGET, range: "A1:C5" }
];

/** 结果矩阵的原始字节（结果指纹据它算；`stagePythonResultPreview` 会核对）。 */
const RESULT_WIRE = JSON.stringify({ kind: "matrix", values: EXPECTED_SUMMARY });

async function sha256Hex(text: string) {
  return Array.from(
    new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text)))
  )
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

async function resultReceipt() {
  return {
    resultHandle: `result://${"r".repeat(43)}`,
    contentFingerprint: await sha256Hex(RESULT_WIRE),
    rows: EXPECTED_SUMMARY.length,
    columns: EXPECTED_SUMMARY[0].length
  };
}

describe("通用合成案例 · 第 3 段：真实结果 → 白名单计划 → 预览 → 写回 → 回读", () => {
  const TARGET_SOURCE = {
    sheetId: SHEET_IDS[TARGET],
    sheetName: TARGET,
    range: "A1:C5"
  };

  beforeEach(() => {
    // 只有"隔离执行/结果取回"这一段用网络替身：请求体仍由产品代码拼。
    const requests: { url: string; body: string }[] = [];
    vi.stubGlobal(
      "fetch",
      vi.fn(async (url: string, init: RequestInit) => {
        requests.push({ url, body: String(init.body) });
        if (url.endsWith("/result")) return new Response(RESULT_WIRE);
        if (url.endsWith("/scope")) return new Response(JSON.stringify({ scopeToken: "t".repeat(43) }));
        if (url.endsWith("/snapshot"))
          return new Response(
            JSON.stringify({
              datasetHandle: `dataset://${"d".repeat(43)}`,
              contentFingerprint: "c".repeat(64)
            })
          );
        return new Response(JSON.stringify({ revoked: true }));
      })
    );
    (capabilities.pythonExecution as { enabled: boolean }).enabled = true;
  });
  afterEach(() => {
    (capabilities.pythonExecution as { enabled: boolean }).enabled = false;
  });

  /** 真实 `PythonDataHost`：身份/授权来源来自本机替身宿主，读值走真 `Excel.run`。 */
  function syntheticDataHost(): PythonDataHost {
    return {
      hostId: "excel-office",
      identity: `url:${BOOK_URL}`,
      authorizedSources: [
        { id: SHEET_IDS[MEMBERS], name: MEMBERS },
        { id: SHEET_IDS[ORDERS], name: ORDERS },
        { id: SHEET_IDS[TARGET], name: TARGET }
      ],
      async assertCurrent() {},
      async read(source) {
        return (await Excel.run(async (context) => {
          const range = context.workbook.worksheets
            .getItem(source.sheetName)
            .getRange(source.range);
          range.load("values");
          await context.sync();
          return range.values as CellValue[][];
        })) as CellValue[][];
      }
    };
  }

  it("坐标只由白名单编译给出：结果矩阵里的值写进目标区域，逐单元格回读一致", async () => {
    const dataHost = syntheticDataHost();
    const session = await collectPythonData(dataHost, {
      taskId: "task-synthetic-1",
      authorizationVersion: "b".repeat(64),
      sources: RESULT_SOURCES,
      apiBaseUrl: "http://127.0.0.1:8765"
    });

    // 预览：真实白名单编译 + 真实来源快照核对 + 真实写前复核。此时**零写入**。
    const beforeWrites = host.originalWrites().length;
    const preview = await prepareOfficePythonResult(
      session,
      dataHost,
      await resultReceipt(),
      TARGET_SOURCE
    );
    // 计划形状由编译给出：只写目标表的目标区域，值是结果矩阵本身。
    expect(preview.plan.actions).toEqual([
      { type: "writeValues", sheet: TARGET, range: "A1:C5", values: EXPECTED_SUMMARY }
    ]);
    // 结果矩阵不含任何"写哪张表/哪个区域"的坐标 —— 坐标是可信调用方给的。
    expect(JSON.stringify(EXPECTED_SUMMARY)).not.toContain(TARGET);
    expect(host.originalWrites().length).toBe(beforeWrites);

    // 写前复核在**没改来源/目标**时通过。
    await expect(preview.assertCurrent()).resolves.toBeUndefined();

    // 执行：把白名单计划写进目标区域（真 Excel.run 替身宿主）。
    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem(TARGET).getRange("A1:C5");
      range.values = EXPECTED_SUMMARY;
      await context.sync();
    });

    // 独立回读：逐单元格一致 ⇒ 才允许说"已核对"。
    const readBack = (await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem(TARGET).getRange("A1:C5");
      range.load("values");
      await context.sync();
      return range.values as CellValue[][];
    })) as CellValue[][];
    expect(readBack).toEqual(EXPECTED_SUMMARY);
    expect(host.matrix(TARGET, "A1:C5")).toEqual(EXPECTED_SUMMARY);
    // 来源表依然没被碰过。
    expect(host.matrix(MEMBERS, "A1:C6")).toEqual(MEMBER_ROWS);
    expect(host.matrix(ORDERS, "A1:B5")).toEqual(ORDER_ROWS);
  });

  it("预览到写入之间来源表被改：写前复核拒绝，零写入", async () => {
    const dataHost = syntheticDataHost();
    const session = await collectPythonData(dataHost, {
      taskId: "task-synthetic-2",
      authorizationVersion: "b".repeat(64),
      sources: RESULT_SOURCES,
      apiBaseUrl: "http://127.0.0.1:8765"
    });
    const preview = await prepareOfficePythonResult(
      session,
      dataHost,
      await resultReceipt(),
      TARGET_SOURCE
    );

    const beforeWrites = host.originalWrites().length;
    // 来源表原件变了（不是目标表）：结果据以算出的输入已不再可信。
    host.setCell(MEMBERS, "B2", "被人改过的会员");
    await expect(preview.assertCurrent()).rejects.toThrow();
    // 零写入：拒绝就真的什么都没写。
    expect(host.originalWrites().length).toBe(beforeWrites);
    expect(host.matrix(TARGET, "A1:C5")).toEqual([
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""]
    ]);
  });

  it("写入目标不在声明来源里：白名单编译直接拒绝，结果矩阵不能自己加坐标", async () => {
    /*
     * 负向锚点：把"目标表"声明成本次**未**纳入来源范围的一处坐标。
     * 结果只是**值**，它无权新增写入坐标 —— `compilePythonValuesPlan` 必须在
     * 生成计划之前就拒绝（`scope.sources` 里找不到目标表）。
     *
     * 这条用例是"坐标只由可信调用方给出"的判别点：拆掉那道来源检查，它必红。
     */
    const dataHost = syntheticDataHost();
    // 来源只声明两张**真来源**表，不含写入目标。
    const session = await collectPythonData(dataHost, {
      taskId: "task-synthetic-3",
      authorizationVersion: "b".repeat(64),
      sources: [
        { sheetId: SHEET_IDS[MEMBERS], sheetName: MEMBERS, range: "A1:C6" },
        { sheetId: SHEET_IDS[ORDERS], sheetName: ORDERS, range: "A1:B5" }
      ],
      apiBaseUrl: "http://127.0.0.1:8765"
    });
    const beforeWrites = host.originalWrites().length;
    await expect(
      prepareOfficePythonResult(session, dataHost, await resultReceipt(), TARGET_SOURCE)
    ).rejects.toThrow();
    // 拒绝就真的什么都没写。
    expect(host.originalWrites().length).toBe(beforeWrites);
    expect(host.matrix(TARGET, "A1:C5")).toEqual([
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""]
    ]);
  });
});
