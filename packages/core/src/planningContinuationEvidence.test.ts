/**
 * 「继续分析」现场证据的判据（共享层）。
 *
 * 这一层被两端**同一份**调用，所以每条断言都直接对应一个生产后果：
 * 判宽了就等于"读不到宿主也能继续"，判窄了会把纯结构任务堵死。
 */

import { describe, expect, it, vi } from "vitest";

import {
  PLANNING_EVIDENCE_FAILURE_MESSAGES,
  buildPlanningContinuationEvidence
} from "./planningContinuationEvidence";

const BASE = {
  workbookIdentity: "url:https://example.invalid/book.xlsx",
  authorizedSheetNames: ["数据", "明细"],
  sourceRequired: false,
  readSourceFingerprint: () => "sha256:abc"
};

describe("buildPlanningContinuationEvidence", () => {
  it("身份与授权都在、原任务不依赖来源：给 not_collected，且**不去读**数据", async () => {
    const readSourceFingerprint = vi.fn(() => "sha256:abc");
    const result = await buildPlanningContinuationEvidence({
      ...BASE,
      readSourceFingerprint
    });
    expect(result).toEqual({
      ok: true,
      evidence: {
        workbookIdentity: "url:https://example.invalid/book.xlsx",
        authorizedSheetNames: ["数据", "明细"],
        sourceVerification: "not_collected"
      }
    });
    expect(readSourceFingerprint).not.toHaveBeenCalled();
  });

  it("原任务依赖来源：真的重读一次，并把指纹带上", async () => {
    const result = await buildPlanningContinuationEvidence({
      ...BASE,
      sourceRequired: true
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.evidence.sourceVerification).toBe("verified");
    expect(result.evidence.sourceFingerprint).toBe("sha256:abc");
  });

  it("读不到工作簿身份：停下，**不**拿文件名/记录值补一个", async () => {
    for (const identity of [null, undefined, "", "   "]) {
      const result = await buildPlanningContinuationEvidence({
        ...BASE,
        workbookIdentity: identity
      });
      expect(result).toEqual({ ok: false, failure: "identity_unavailable" });
    }
  });

  it("授权成员为空（或全是空白）：停下，而不是发一个空集合去碰运气", async () => {
    for (const names of [[], ["", "  "]] as const) {
      const result = await buildPlanningContinuationEvidence({
        ...BASE,
        authorizedSheetNames: [...names]
      });
      expect(result).toEqual({ ok: false, failure: "scope_unavailable" });
    }
  });

  it("重读抛错：算**读取失败**，绝不退化成'尚未采集'", async () => {
    const result = await buildPlanningContinuationEvidence({
      ...BASE,
      sourceRequired: true,
      readSourceFingerprint: () => {
        throw new Error("宿主读取失败");
      }
    });
    expect(result).toEqual({ ok: false, failure: "source_unreadable" });
  });

  it("重读返回空：同样算读取失败（读到空气不等于数据没变）", async () => {
    const result = await buildPlanningContinuationEvidence({
      ...BASE,
      sourceRequired: true,
      readSourceFingerprint: () => "   "
    });
    expect(result).toEqual({ ok: false, failure: "source_unreadable" });
  });

  it("授权成员去重并去空白，顺序不参与身份", async () => {
    const result = await buildPlanningContinuationEvidence({
      ...BASE,
      authorizedSheetNames: [" 数据 ", "明细", "数据"]
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.evidence.authorizedSheetNames).toEqual(["数据", "明细"]);
  });

  it("三种失败都有独立文案（不许两种失败共用一句话）", () => {
    const messages = Object.values(PLANNING_EVIDENCE_FAILURE_MESSAGES);
    expect(messages).toHaveLength(3);
    expect(new Set(messages).size).toBe(3);
    for (const message of messages) expect(message.trim().length).toBeGreaterThan(8);
  });
});
