import { describe, expect, it } from "vitest";
import {
  PLANNING_CONTEXT_CODES,
  PLANNING_CONTEXT_LIMITS,
  canRequestPlanningContinuation,
  findForbiddenKey,
  planningContextMessage,
  validatePlanningContext
} from "./planningContext";
import fixtures from "./planningContext.fixtures.json";

interface FixtureCase {
  name: string;
  record?: unknown;
  patch?: Record<string, unknown>;
  remove?: string[];
  expected: { ok: boolean; code?: string };
}

const LONG_PROMPT = "长".repeat(PLANNING_CONTEXT_LIMITS.maxPromptChars + 1);

export function materializeFixture(entry: FixtureCase): unknown {
  if ("record" in entry) {
    return entry.record === "__LONG_PROMPT__" ? LONG_PROMPT : entry.record;
  }
  const base = JSON.parse(JSON.stringify(fixtures.base)) as Record<string, unknown>;
  for (const key of entry.remove ?? []) delete base[key];
  for (const [key, value] of Object.entries(entry.patch ?? {})) {
    base[key] = value === "__LONG_PROMPT__" ? LONG_PROMPT : value;
  }
  return base;
}

describe("planning_context 契约（与 server/app/planning_context.py 共用同一份夹具）", () => {
  const cases = fixtures.cases as FixtureCase[];

  it("夹具本身覆盖了每条拒绝码", () => {
    const codes = new Set(
      cases.filter((item) => item.expected.code).map((item) => item.expected.code)
    );
    for (const code of Object.values(PLANNING_CONTEXT_CODES)) {
      expect(codes.has(code), `夹具缺少拒绝码 ${code}`).toBe(true);
    }
  });

  it.each(cases.map((item) => [item.name, item] as const))(
    "%s",
    (_name, entry) => {
      const verdict = validatePlanningContext(
        materializeFixture(entry),
        fixtures.expect
      );
      expect(verdict.ok).toBe(entry.expected.ok);
      if (!entry.expected.ok && !verdict.ok) {
        expect(verdict.code).toBe(entry.expected.code);
      }
    }
  );

  it("合法检查点被投影成白名单结构，未知键不会穿透", () => {
    const verdict = validatePlanningContext(fixtures.base, fixtures.expect);
    expect(verdict.ok).toBe(true);
    if (!verdict.ok) return;
    expect(Object.keys(verdict.record).sort()).toEqual(
      [
        "authorization",
        "completedTools",
        "confirmations",
        "executorId",
        "failedTool",
        "kind",
        "planningGoal",
        "prompt",
        "resultRefs",
        "resumability",
        "savedAt",
        "schemaVersion",
        "sourceFingerprint",
        "sourceFingerprintSheets",
        "stage",
        "target",
        "taskId",
        "taskRevision",
        "turnId",
        "unknownTool"
      ].sort()
    );
  });

  it("值矩阵键递归可见：嵌套在数组对象里也能抓到", () => {
    expect(
      findForbiddenKey({ list: [{ deep: { values: [[1]] } }] })
    ).toBe("values");
    expect(findForbiddenKey({ fine: [{ tool: "query_table" }] })).toBeNull();
  });

  it("每个拒绝码都有独立的用户可见说明（不许糊成一句）", () => {
    const messages = new Set<string>();
    for (const code of Object.values(PLANNING_CONTEXT_CODES)) {
      const text = planningContextMessage(code);
      expect(text.length).toBeGreaterThan(8);
      messages.add(text);
    }
    expect(messages.size).toBe(Object.values(PLANNING_CONTEXT_CODES).length);
  });

  it("客户端预检与服务端同判：缺现场不发请求", () => {
    const verdict = canRequestPlanningContinuation(null, fixtures.expect);
    expect(verdict.ok).toBe(false);
    expect(verdict.code).toBe(PLANNING_CONTEXT_CODES.missing);
    expect(verdict.message).toContain("重新发起");
  });
});
