import { describe, expect, it } from "vitest";
import {
  EXECUTION_RECOVERY_PACKAGE_VERSION,
  assertExecutionRecoveryPackage,
  executionRecoveryPackageVerdict,
  executionRecoveryPlanFingerprint,
  parseExecutionRecoveryPackage,
  type ExecutionRecoveryPackage
} from "./executionRecoveryPackage";
import cases from "../../../config/execution_recovery_package_cases.json";

it.each(cases.sourceEvidenceCases)("source evidence version $version empty=$empty valid=$valid", item=>{
  const candidate=structuredClone(cases.valid[0]);
  candidate.sourceEvidence.version=item.version;
  if(item.empty) candidate.sourceEvidence.sources=[];
  if(item.valid) expect(()=>assertExecutionRecoveryPackage(candidate)).not.toThrow();
  else expect(()=>assertExecutionRecoveryPackage(candidate)).toThrow();
});

/**
 * 公式动作的准入样例与服务端共用（`config/execution_recovery_package_cases.json`）。
 *
 * 这里刻意走**整包**校验而不是只验单个动作：真机上出错的是"包能读、动作写错"，
 * 只有在完整包里跑一遍，planFingerprint 重算与来源/目标不重叠这两条才不会被跳过。
 */
it.each((cases as { formulaActionCases: { name: string; valid: boolean; action: unknown }[] }).formulaActionCases)(
  "formula action $name valid=$valid",
  (item) => {
    const pkg = structuredClone(cases.valid[0]) as unknown as ExecutionRecoveryPackage;
    pkg.actions = [structuredClone(item.action) as ExecutionRecoveryPackage["actions"][number]];
    pkg.planFingerprint = executionRecoveryPlanFingerprint(pkg.actions);
    if (item.valid) expect(() => assertExecutionRecoveryPackage(pkg), item.name).not.toThrow();
    else expect(() => assertExecutionRecoveryPackage(pkg), item.name).toThrow();
  }
);

/**
 * 恢复包协议的两个判别点：
 * ① **白名单**：多一个字段、矩阵形状不对、指纹口径不符，一律拒绝（不是"尽量兼容"）；
 * ② **四类结论分开**：过期 / 不兼容 / 对不上 / 可用，不能用一个 false 抹平 ——
 *    否则界面上只会说"材料不可用"，用户无从知道是该早点回来、换工作簿还是重新规划。
 *
 * 跨语言那半（同一个 fixture 必须被 `server/app/models.py` 接受）在
 * `server/tests/test_execution_recovery_package.py` 里，读同一份 JSON。
 */
function sample(): ExecutionRecoveryPackage {
  const actions = [{ actionIndex: 0, type: "writeValues" as const, sheet: "目标", range: "A1:A1",
    rowCount: 1, columnCount: 1, cellCount: 1, previewFingerprint: "v2:1:aaaaaaaa", expectedFingerprint: "v2:1:bbbbbbbb" },
    { actionIndex: 1, type: "writeValues" as const, sheet: "目标", range: "A2:A2",
    rowCount: 1, columnCount: 1, cellCount: 1, previewFingerprint: "v2:1:cccccccc", expectedFingerprint: "v2:1:dddddddd" }];
  return {
    schemaVersion: EXECUTION_RECOVERY_PACKAGE_VERSION,
    host: "wps",
    operationId: "wps-execution-0001",
    taskId: "task-1",
    workbookIdentity: "wps-path:c:\\data\\test.xlsx",
    authorizedSheetNames: ["来源", "目标"],
    planId: "plan-resume",
    planTitle: "更新",
    planFingerprint: executionRecoveryPlanFingerprint(actions),
    hostVersion: "12.0",
    createdAt: "2026-10-20T00:00:00.000Z",
    actions: actions.map((action, index) => ({ ...action,
      values: [[index + 1]], beforeValues: [[`old${index + 1}`]], beforeFormulas: [[`old${index + 1}`]] })),
    sourceEvidence: { version: 1, hostVersion: "12.0", sources: [{ sheetName: "来源", range: "A1:A1",
      rowCount: 1, columnCount: 1, previewFingerprint: "v2:1:eeeeeeee", valueFingerprint: "v2:1:ffffffff" }] }
  };
}

describe("execution recovery package", () => {
  it("accepts the shared fixture and the locally built package", () => {
    expect(() => assertExecutionRecoveryPackage(sample())).not.toThrow();
    for (const item of (cases as { valid: unknown[] }).valid)
      expect(() => assertExecutionRecoveryPackage(item)).not.toThrow();
  });
  it("rejects every invalid case in the shared fixture with its own reason", () => {
    for (const item of (cases as { invalid: { name: string; value: unknown }[] }).invalid)
      expect(() => assertExecutionRecoveryPackage(item.value), item.name).toThrow();
  });
  it("rejects unknown fields, unsupported versions and stale plan fingerprints", () => {
    expect(() => assertExecutionRecoveryPackage({ ...sample(), plan: { actions: [] } })).toThrow(/未知字段/);
    expect(() => assertExecutionRecoveryPackage({ ...sample(), schemaVersion: 2 })).toThrow(/格式版本/);
    const relabelled = sample();
    relabelled.actions[1].previewFingerprint = "v2:1:99999999";
    // 写前指纹被改而计划指纹没跟着变：说明材料被改过，不能当成同一版计划。
    expect(() => assertExecutionRecoveryPackage(relabelled)).toThrow(/计划指纹/);
    const dropped = sample();
    dropped.actions = dropped.actions.slice(0, 1);
    expect(() => assertExecutionRecoveryPackage(dropped)).toThrow(/计划指纹/);
  });
  it("keeps source and target sheets disjoint", () => {
    const overlapping = sample();
    overlapping.actions[0].sheet = "来源";
    overlapping.planFingerprint = executionRecoveryPlanFingerprint(overlapping.actions);
    expect(() => assertExecutionRecoveryPackage(overlapping)).toThrow(/来源与写入目标重叠/);
  });
  it("parses only well-formed JSON text", () => {
    expect(parseExecutionRecoveryPackage(JSON.stringify(sample())).operationId).toBe("wps-execution-0001");
    expect(() => parseExecutionRecoveryPackage("{")).toThrow(/不是有效 JSON/);
  });
  it("separates expired, incompatible and mismatched material", () => {
    const base = { package: sample(), operationId: "wps-execution-0001", taskId: "task-1",
      workbookIdentity: "wps-path:c:\\data\\test.xlsx", authorizedSheetNames: ["来源", "目标"],
      hostVersion: "12.0", expiresAt: "2026-10-27T00:00:00.000Z",
      now: Date.parse("2026-10-26T00:00:00.000Z") };
    expect(executionRecoveryPackageVerdict(base)).toEqual({ kind: "ok" });
    expect(executionRecoveryPackageVerdict({ ...base, now: Date.parse("2026-10-28T00:00:00.000Z") }).kind)
      .toBe("expired");
    expect(executionRecoveryPackageVerdict({ ...base, hostVersion: "12.1" }).kind).toBe("incompatible");
    expect(executionRecoveryPackageVerdict({ ...base, workbookIdentity: "wps-path:c:\\data\\other.xlsx" }).kind)
      .toBe("mismatch");
    expect(executionRecoveryPackageVerdict({ ...base, authorizedSheetNames: ["来源"] }).kind).toBe("mismatch");
    expect(executionRecoveryPackageVerdict({ ...base, expiresAt: "not-a-date" }).kind).toBe("incompatible");
  });
  it("treats the exact expiry instant as already expired, like the server does", () => {
    // 服务端 `execution_recovery.load` 判 `deadline <= now`；这里必须是同一个边界，
    // 否则"恰好到期"在两端会给出相反结论（一端还能读、一端已经拒绝）。
    const base = { package: sample(), operationId: "wps-execution-0001", taskId: "task-1",
      workbookIdentity: "wps-path:c:\\data\\test.xlsx", authorizedSheetNames: ["来源", "目标"],
      hostVersion: "12.0", expiresAt: "2026-10-27T00:00:00.000Z" };
    const deadline = Date.parse(base.expiresAt);
    expect(executionRecoveryPackageVerdict({ ...base, now: deadline - 1 }).kind).toBe("ok");
    expect(executionRecoveryPackageVerdict({ ...base, now: deadline }).kind).toBe("expired");
    expect(executionRecoveryPackageVerdict({ ...base, now: deadline + 1 }).kind).toBe("expired");
  });
  it("recognises key order changes from SQLite (sorted keys) as the same plan", () => {
    const sorted = (value: unknown): unknown => Array.isArray(value) ? value.map(sorted) :
      value && typeof value === "object"
        ? Object.fromEntries(Object.entries(value).sort(([a], [b]) => a.localeCompare(b)).map(([k, v]) => [k, sorted(v)]))
        : value;
    const shuffled = JSON.parse(JSON.stringify(sorted(sample())));
    expect(() => assertExecutionRecoveryPackage(shuffled)).not.toThrow();
  });
  it("agrees on the plan fingerprint regardless of action order being preserved", () => {
    const first = sample();
    const reversed = { ...first, actions: [...first.actions].reverse().map((action, index) => ({ ...action, actionIndex: index })) };
    reversed.planFingerprint = executionRecoveryPlanFingerprint(reversed.actions);
    expect(reversed.planFingerprint).not.toBe(first.planFingerprint);
  });
});
