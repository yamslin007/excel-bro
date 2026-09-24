import { describe, expect, it } from "vitest";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import { probeWpsWriteCapabilities, wpsWriteFeatureFlagEnabled } from "./wpsWriteCapabilities";

function application(overrides: {
  fullName?: string;
  readOnly?: boolean;
  protectContents?: boolean;
  cell?: WpsRange;
} = {}): WpsApplication {
  const sheet: WpsWorksheet = {
    Name: "明细",
    ProtectContents: overrides.protectContents ?? false,
    Range: () => overrides.cell ?? { Value2: null, Formula: "", MergeCells: false }
  };
  return {
    Version: "12.1",
    PluginStorage: { getItem: () => null, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined,
    ActiveWorkbook: {
      Name: "测试.xlsx",
      FullName: overrides.fullName ?? "C:\\Data\\测试.xlsx",
      ReadOnly: overrides.readOnly ?? false,
      Worksheets: { Count: 1, Item: () => sheet }
    }
  };
}

describe("WPS W3 read-only capability probe", () => {
  it("checks an authorized local workbook without writing", () => {
    const cell: WpsRange = { Value2: "原值", Formula: "原值", MergeCells: false };
    const result = probeWpsWriteCapabilities(application({ cell }), [1]);
    expect(result.readyForHostValidation).toBe(true);
    expect(result.writeEnabled).toBe(false);
    expect(result.checks.every((check) => check.passed)).toBe(true);
    expect(cell.Value2).toBe("原值");
  });

  it("fails closed for an empty authorization, an unknown read-only state, and unsupported paths", () => {
    const app = application({ fullName: "https://cloud.example.com/测试.xlsx" });
    app.ActiveWorkbook!.ReadOnly = undefined;
    const result = probeWpsWriteCapabilities(app, []);
    expect(result.readyForHostValidation).toBe(false);
    expect(result.checks.filter((check) => !check.passed).map((check) => check.id)).toContain("authorization");
    expect(result.checks.filter((check) => !check.passed).map((check) => check.id)).toContain("localXlsx");
    expect(result.checks.filter((check) => !check.passed).map((check) => check.id)).toContain("readOnly");
  });

  it("rejects protected sheets and unknown merge behavior", () => {
    const result = probeWpsWriteCapabilities(application({
      protectContents: true,
      cell: { Value2: 1, Formula: 1 }
    }), [1]);
    expect(result.readyForHostValidation).toBe(false);
    expect(result.checks.find((check) => check.id === "protection")?.passed).toBe(false);
    expect(result.checks.find((check) => check.id === "merge")?.passed).toBe(false);
  });

  it("does not claim formula recovery when the host omits Formula", () => {
    const result = probeWpsWriteCapabilities(application({
      cell: { Value2: "内容", MergeCells: false }
    }), [1]);
    expect(result.readyForHostValidation).toBe(false);
    expect(result.checks.find((check) => check.id === "formula")?.passed).toBe(false);
  });

  it("keeps the WPS write feature flag disabled until real host validation", () => {
    expect(wpsWriteFeatureFlagEnabled()).toBe(false);
  });
});
