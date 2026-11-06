// @vitest-environment jsdom
/**
 * 模型 → 工具 → 宿主的真实链路（Excel 侧）。
 *
 * 现场（2026-11-01 真机）：用户只说「只读取来源!A1:B5」，结果卡列出四张来源表。
 * 之前的测试都只调执行器、手工传正确参数，**跳过了桥这一跳** —— 而模型给的参数正是
 * 从这里进来的。这个文件让参数沿生产路径走：桥的 `execute` → 真实
 * `executeQueryTableTool` → 真实宿主替身（只替 Office.js），断言结果卡看到的事实。
 */

import { afterEach, describe, expect, it, vi } from "vitest";

import type { CellValue } from "./contracts";
import { createOfficeTestHost } from "./officeTestHost";
import { resetTrustedSourceReads, trustedToolReads } from "./planSourceDependency";
import {
  createSdkBridgeBinding,
  createSdkBridgeHostPorts,
  type SdkBridgeLiveContext
} from "./sdkBridgeHost";

const OP = "op_" + "3".repeat(32);
const FP = "e".repeat(64);

/** 四张表都已授权，且每张都有可求和的第二列：多读一张，合计就不是 15。 */
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
const NAMES = SHEETS.map((sheet) => sheet.name);

const LIVE: SdkBridgeLiveContext = {
  workbookIdentity: "file:///C:/Data/公式来源证据验收.xlsx",
  sourceFingerprint: null,
  authorizedSheetNames: NAMES,
  activeWorksheet: "来源"
};

function binding() {
  return createSdkBridgeBinding({
    taskId: "turn-scope-excel-0001",
    hostId: "excel-office",
    workbookName: "公式来源证据验收.xlsx",
    workbookIdentity: LIVE.workbookIdentity,
    sourceFingerprint: null,
    authorizedSheetNames: NAMES
  });
}

function query(arguments_: Record<string, unknown>) {
  return {
    operationId: OP,
    tool: "query_table" as const,
    arguments: arguments_,
    requestFingerprint: FP,
    expiresAt: 0
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
  resetTrustedSourceReads();
});

describe("桥上的 query_table：目标表必须绑定", () => {
  it("区域带表名：结果只覆盖「来源」，完整无截断、零写入", async () => {
    resetTrustedSourceReads();
    const office = createOfficeTestHost({
      url: LIVE.workbookIdentity,
      activeSheet: "来源",
      sheets: SHEETS
    });
    vi.stubGlobal("Excel", office.Excel);
    const bridge = binding();
    const hostPorts = createSdkBridgeHostPorts({ liveContext: () => LIVE }, bridge);

    const outcome = await hostPorts.execute(
      bridge,
      query({
        mode: "aggregate",
        scope: "selected",
        range: "来源!A1:B5",
        hasHeaders: false,
        metrics: [{ operation: "sum", field: "列2", outputName: "合计" }]
      })
    );

    expect(outcome.kind).toBe("result");
    const result = (outcome as { result: Record<string, unknown> }).result;
    expect(result.rows).toEqual([[15]]);
    expect(result.sourceSheets).toEqual(["来源"]);
    expect(result.complete).toBe(true);
    expect(result.incomplete).toEqual([]);
    expect(result.truncated).toBe(false);
    // 记账与结果同源：读过的表就是参与计算的那一张。
    expect(trustedToolReads(bridge.taskId)?.sheets).toEqual(["来源"]);
    expect(office.writes()).toEqual([]);
  });

  it("不给表名又有多张可查表：回错误码，不把四张表都读一遍", async () => {
    resetTrustedSourceReads();
    const office = createOfficeTestHost({
      url: LIVE.workbookIdentity,
      activeSheet: "来源",
      sheets: SHEETS
    });
    vi.stubGlobal("Excel", office.Excel);
    const bridge = binding();
    const hostPorts = createSdkBridgeHostPorts({ liveContext: () => LIVE }, bridge);

    const outcome = await hostPorts.execute(
      bridge,
      query({
        mode: "aggregate",
        scope: "selected",
        hasHeaders: false,
        metrics: [{ operation: "sum", field: "列2", outputName: "合计" }]
      })
    );

    // 参数能改对 ⇒ invalid_arguments（而不是"越权"），模型据此重试一次即可。
    expect(outcome).toEqual({ kind: "error", errorCode: "invalid_arguments" });
    expect(trustedToolReads(bridge.taskId)).toBeNull();
    expect(office.writes()).toEqual([]);
  });

  it("跨表查询显式列出四张表：都读，合计 315（对照组，证明上一条不是空转）", async () => {
    resetTrustedSourceReads();
    const office = createOfficeTestHost({
      url: LIVE.workbookIdentity,
      activeSheet: "来源",
      sheets: SHEETS
    });
    vi.stubGlobal("Excel", office.Excel);
    const bridge = binding();
    const hostPorts = createSdkBridgeHostPorts({ liveContext: () => LIVE }, bridge);

    const outcome = await hostPorts.execute(
      bridge,
      query({
        mode: "aggregate",
        scope: "selected",
        sheets: NAMES,
        hasHeaders: false,
        metrics: [{ operation: "sum", field: "列2", outputName: "合计" }]
      })
    );

    const result = (outcome as { result: Record<string, unknown> }).result;
    expect(result.rows).toEqual([[315]]);
    expect(result.sourceSheets).toEqual(NAMES);
    expect(result.complete).toBe(true);
  });
});
