// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { targetDiagnostic } from "@excel-bro/core/targetDiagnostics";
import {
  recordWpsTargetDiagnostic,
  resetWpsTargetDiagnostics,
  wpsTargetDiagnosticEvents,
  wpsTargetDiagnosticReport
} from "./wpsTargetDiagnostics";

beforeEach(() => { resetWpsTargetDiagnostics(); });
afterEach(() => { resetWpsTargetDiagnostics(); vi.restoreAllMocks(); });

function sample(overrides: Partial<Parameters<typeof targetDiagnostic>[0]> = {}) {
  return targetDiagnostic({
    outcome: "blocked", gate: "analysis", branch: "weak_evidence",
    conclusive: false, documentBound: false, hasConsent: false, explicitRequired: false,
    stableId: "none", pathIdentity: true, round: 1, seq: 1, ...overrides
  });
}

describe("WPS 范围判定诊断的记录端", () => {
  it("只按既有字段记录，报告里只有原因码与布尔值", () => {
    recordWpsTargetDiagnostic(sample());
    const report = wpsTargetDiagnosticReport();
    expect(report).toContain("weak_evidence");
    expect(report).toContain("documentBound=false");
    expect(report).toContain("stableId=none");
    // 诊断的字段集合被类型钉死：报告里不可能冒出路径/表名/证据键。
    expect(report).not.toMatch(/[A-Za-z]:\\|Sheet|PRIVATE/);
  });
  it("环形缓冲只保留最近若干条，不会一直长", () => {
    for (let index = 0; index < 12; index++) {
      recordWpsTargetDiagnostic(sample({ seq: index + 1 }));
    }
    const events = wpsTargetDiagnosticEvents();
    expect(events).toHaveLength(8);
    expect(events.at(-1)?.seq).toBe(12);
  });
  it("没有判定时报告明说这一轮还没判定，不是空白", () => {
    expect(wpsTargetDiagnosticReport()).toContain("还没有");
  });
});
