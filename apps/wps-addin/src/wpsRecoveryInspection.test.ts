import { describe, expect, it } from "vitest";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import type { WpsApplication } from "./wpsHost";
import { matrixFingerprintV2 } from "./wpsPlanExecutor";
import { inspectWpsRecoveryState } from "./wpsRecoveryInspection";

const fp = (value: string | null) => matrixFingerprintV2([[value]], [[value]]);

function fixture(values: Record<string, unknown> = { A1: "after", B2: "before" }) {
  const reads: string[] = [];
  const sheet = (name: string) => ({ Name: name, Range(address: string) {
    const value = values[address];
    return { Rows: { Count: 1 }, Columns: { Count: 1 },
      get Value2() { reads.push(`${name}!${address}`); return value; },
      get Formula() { return value; }
    };
  }});
  const sheets = [sheet("Allowed"), sheet("Other")];
  const application = { ActiveWorkbook: { Name: "test.xlsx", FullName: "C:\\data\\test.xlsx",
    Worksheets: { Count: 2, Item: (index: number) => sheets[index - 1] } } } as WpsApplication;
  const state: WpsExecutionState = { operationId: "execution-test", taskId: "task-test", status: "failed",
    workbookIdentity: "wps-path:c:\\data\\test.xlsx", revision: 2, authorizedSheetNames: ["Allowed"],
    ownerActive: false, canResume: true, requiresReconciliation: false,
    nextActionIndex: 1, totalActions: 2, inFlightActionIndex: null,
    descriptor: { actions: ["A1", "B2"].map((range, actionIndex) => ({
      actionIndex, type: "writeValues", sheet: "Allowed", range,
      rowCount: 1, columnCount: 1, cellCount: 1, previewFingerprint: fp("before"), expectedFingerprint: fp("after")
    })) }
  };
  return { application, state, reads };
}

describe("CAP-03 按登记坐标重读现场，非 PluginStorage 自证", () => {
  it("只读取已授权的明确目标，已完成和剩余动作分别对齐写后和写前", () => {
    const { application, state, reads } = fixture();
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("not_started");
    expect(reads).toEqual(["Allowed!A1", "Allowed!B2"]);
  });
  it("全部动作写后相符才证明零剩余完成", () => {
    const { application, state } = fixture({ A1: "after", B2: "after" });
    state.nextActionIndex = 2;
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("completed");
  });
  it.each(["A1", "B2"])("人工修改 %s：无论已完成或未执行都判冲突", (address) => {
    const { application, state } = fixture({ A1: "after", B2: "before", [address]: "human" });
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("conflict");
  });
  it.each(["before", "after"])("在途动作现场 %s：只核对不写入", (value) => {
    const { application, state } = fixture({ A1: "after", B2: value });
    state.inFlightActionIndex = 1;
    expect(inspectWpsRecoveryState(application, state, [1])).toBe(value === "after" ? "completed" : "not_started");
  });
  it("授权缺失时 Value2 零读取，不能沿用登记授权扩大今天范围", () => {
    const { application, state, reads } = fixture();
    expect(inspectWpsRecoveryState(application, state, [2])).toBe("unknown");
    expect(reads).toEqual([]);
  });
  it("后续动作越界：先验证全部范围，不能先读合法动作再拒绝", () => {
    const { application, state, reads } = fixture();
    state.descriptor.actions[1].sheet = "Other";
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("unknown");
    expect(reads).toEqual([]);
  });
  it.each([undefined, {}, []])("未采集/畸形数据不能归一化冒充空格：%j", (raw) => {
    const { application, state } = fixture({ A1: "after", B2: raw });
    state.descriptor.actions[1].previewFingerprint = fp(null);
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("unknown");
  });
  it("值未采集与真正空白 null 明确区分", () => {
    const { application, state } = fixture({ A1: "after", B2: null });
    state.descriptor.actions[1].previewFingerprint = fp(null);
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("not_started");
  });
  it("跨关闭只核对公式文本，不把计算值相同当作公式相同", () => {
    const { application, state } = fixture({ A1: "=A2+1", B2: "before" });
    state.descriptor.actions[0].type = "writeFormulas";
    state.descriptor.actions[0].expectedFingerprint = fp("=A2+2");
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("conflict");
  });
});
