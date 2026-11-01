import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  formatWpsRecoveryDiagnosticEvent,
  recordWpsRecoveryDiagnostic,
  resetWpsRecoveryDiagnostics,
  safeDiagnosticErrorName,
  wpsRecoveryDiagnosticEvents,
  wpsRecoveryDiagnosticReport
} from "./wpsRecoveryDiagnostics";
import { reportDevProbeRecovery } from "./devProbe";
import { isDev } from "./isDev";

/*
 * 诊断是**观察层**，不参与任何判定。这组用例钉三件事：
 * ① 只在 DEV 记录（非 DEV 一条都不留）；
 * ② 记录里只有结构化标量，异常也只留白名单 `name`，不把 message 带出来；
 * ③ 遥测出口默认关着，只有显式开 `VITE_ENABLE_RECOVERY_DIAG` 才发。
 */
vi.mock("./isDev", () => ({ isDev: vi.fn(() => true) }));
vi.mock("./devProbe", () => ({ reportDevProbeRecovery: vi.fn() }));

const base = {
  stage: "discovery" as const,
  verdict: "insufficient" as const
};

beforeEach(() => {
  resetWpsRecoveryDiagnostics();
  vi.mocked(isDev).mockReturnValue(true);
  vi.mocked(reportDevProbeRecovery).mockClear();
});

afterEach(() => {
  vi.unstubAllEnvs();
  resetWpsRecoveryDiagnostics();
});

describe("恢复诊断：只在 DEV 记录", () => {
  it("DEV 下按顺序记录并自增序号", () => {
    recordWpsRecoveryDiagnostic({ ...base, outcome: "unknown", code: "value_read_undefined", phase: "values" });
    recordWpsRecoveryDiagnostic({ ...base, stage: "recheck", outcome: "completed", code: "value_read_undefined", phase: "values" });
    const events = wpsRecoveryDiagnosticEvents();
    expect(events.map((event) => event.seq)).toEqual([1, 2]);
    expect(events.map((event) => event.stage)).toEqual(["discovery", "recheck"]);
  });

  it("非 DEV 一条都不留（也不打日志、不发请求）", () => {
    vi.mocked(isDev).mockReturnValue(false);
    const debug = vi.spyOn(console, "debug").mockImplementation(() => undefined);
    recordWpsRecoveryDiagnostic({ ...base, outcome: "unknown", code: "authorize_threw", phase: "authorize" });
    expect(wpsRecoveryDiagnosticEvents()).toEqual([]);
    expect(debug).not.toHaveBeenCalled();
    expect(reportDevProbeRecovery).not.toHaveBeenCalled();
    debug.mockRestore();
  });

  it("缓冲区只保留最近若干条，不无限增长", () => {
    for (let index = 0; index < 45; index += 1) {
      recordWpsRecoveryDiagnostic({ ...base, outcome: "unknown", code: "descriptor_empty", phase: "descriptor" });
    }
    const events = wpsRecoveryDiagnosticEvents();
    expect(events).toHaveLength(40);
    // 丢的是最早的，序号连续 —— 否则"两次对照"会错位。
    expect(events[0].seq).toBe(6);
    expect(events.at(-1)!.seq).toBe(45);
  });
});

describe("恢复诊断：只收结构化标量，不收自由文本", () => {
  it("异常只留白名单 name，message 里的路径不会跟着进来", () => {
    const error = new Error("C:\\机密目录\\客户名单.xlsx 单元格 A1 读取失败");
    expect(safeDiagnosticErrorName(error)).toBe("Error");
    // 名字本身不像标识符（带盘符/斜杠/空格/中文）→ 一律丢弃。
    expect(safeDiagnosticErrorName({ name: "C:\\机密目录\\客户名单.xlsx" })).toBeUndefined();
    expect(safeDiagnosticErrorName({ name: "对象不存在" })).toBeUndefined();
    expect(safeDiagnosticErrorName("字符串异常")).toBeUndefined();
    expect(safeDiagnosticErrorName(null)).toBeUndefined();
  });

  // 防"以后有人往记录里塞一个自由文本字段"：键名必须都在白名单里，值只能是标量。
  it("记录字段是封闭集合，值都是标量", () => {
    reportWorstCase();
    const [event] = wpsRecoveryDiagnosticEvents();
    const allowed = new Set([
      "seq", "at", "stage", "verdict", "outcome", "code", "phase",
      "operationIdTail", "taskIdTail", "revision", "nextActionIndex", "totalActions",
      "inFlightActionIndex", "selectedSheetCount", "authorizedSheetCount",
      "actionIndex", "errorName"
    ]);
    for (const [key, value] of Object.entries(event)) {
      expect(allowed.has(key), `诊断里出现了未登记字段：${key}`).toBe(true);
      expect(["string", "number", "undefined"].includes(typeof value), `字段 ${key} 不是标量`).toBe(true);
    }
    const text = formatWpsRecoveryDiagnosticEvent(event);
    for (const secret of ["机密目录", "客户名单", "xlsx", "\\", "ownerToken"]) {
      expect(text).not.toContain(secret);
    }
  });
});

describe("恢复诊断：报告可直接复制", () => {
  it("头部声明口径，正文一行一条，能上下对照两次核验", () => {
    recordWpsRecoveryDiagnostic({ ...base, outcome: "unknown", code: "value_read_undefined", phase: "values",
      operationIdTail: "926ef1", taskIdTail: "b40493", at: Date.UTC(2026, 9, 28, 23, 32, 7, 600) });
    recordWpsRecoveryDiagnostic({ ...base, stage: "recheck", outcome: "not_started", code: "value_read_undefined",
      phase: "values", operationIdTail: "926ef1", taskIdTail: "b40493", at: Date.UTC(2026, 9, 28, 23, 39, 6, 877) });
    const report = wpsRecoveryDiagnosticReport();
    const lines = report.split("\n");
    expect(lines[0]).toContain("UTC");
    expect(lines[0]).toContain("共 2 条");
    expect(lines[1]).toContain("23:32:07.600");
    expect(lines[1]).toContain("discovery");
    expect(lines[2]).toContain("23:39:06.877");
    expect(lines[2]).toContain("recheck");
    // 两次的 code 一样、outcome 不一样：这正是"同一次核对前后"的样子。
    expect(lines[1]).toContain("code=value_read_undefined");
    expect(lines[2]).toContain("outcome=not_started");
  });

  it("没有记录时给一句说明，不返回空串", () => {
    expect(wpsRecoveryDiagnosticReport()).toContain("暂无记录");
  });
});

describe("恢复诊断：遥测出口默认关闭", () => {
  it("不设开关时一条都不发", () => {
    recordWpsRecoveryDiagnostic({ ...base, outcome: "unknown", code: "host_range_unreadable", phase: "read" });
    expect(reportDevProbeRecovery).not.toHaveBeenCalled();
  });

  it("显式开启后才发，且载荷里只有原因码与计数", () => {
    vi.stubEnv("VITE_ENABLE_RECOVERY_DIAG", "true");
    const error = new Error("C:\\机密目录\\客户名单.xlsx 打不开");
    recordWpsRecoveryDiagnostic({
      ...base,
      outcome: "unknown",
      code: "inspection_threw",
      phase: "read",
      operationIdTail: "926ef1",
      taskIdTail: "b40493",
      revision: 3,
      selectedSheetCount: 2,
      authorizedSheetCount: 2,
      errorName: safeDiagnosticErrorName(error)
    });
    expect(reportDevProbeRecovery).toHaveBeenCalledTimes(1);
    const payload = vi.mocked(reportDevProbeRecovery).mock.calls[0][0];
    expect(payload).toMatchObject({
      stage: "discovery", verdict: "insufficient", outcome: "unknown",
      code: "inspection_threw", phase: "read", operationIdTail: "926ef1", errorName: "Error"
    });
    const serialized = JSON.stringify(payload);
    for (const secret of ["机密目录", "客户名单", "xlsx", "C:", "\\", "ErrorMessage"]) {
      expect(serialized).not.toContain(secret);
    }
  });
});

/** 造一条字段最全的记录（值全部来自允许集合，没有业务数据）。 */
function reportWorstCase(): void {
  recordWpsRecoveryDiagnostic({
    stage: "pre-action",
    verdict: "blocked",
    outcome: "conflict",
    code: "completed_action_mismatch",
    phase: "values",
    operationIdTail: "926ef1",
    taskIdTail: "b40493",
    revision: 3,
    nextActionIndex: 1,
    totalActions: 2,
    inFlightActionIndex: 1,
    selectedSheetCount: 2,
    authorizedSheetCount: 2,
    actionIndex: 0,
    errorName: "Error"
  });
}
