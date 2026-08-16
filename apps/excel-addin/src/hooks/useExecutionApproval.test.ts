// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import type { AnalysisPlan } from "../contracts";
import { analyzeToolEligibility } from "../storage";
import { useExecutionApproval } from "./useExecutionApproval";

vi.mock("../storage", () => ({
  analyzeToolEligibility: vi.fn()
}));

const analyzeToolEligibilityMock = vi.mocked(analyzeToolEligibility);

const plan = { id: "plan-1" } as unknown as AnalysisPlan;

describe("useExecutionApproval", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    analyzeToolEligibilityMock.mockReturnValue({
      issues: [],
      requiredApprovals: [],
      blocked: false
    });
  });

  it("should initialize with no candidate and empty verified set", () => {
    const { result } = renderHook(() => useExecutionApproval());

    expect(result.current.saveCandidate).toBeNull();
    expect(result.current.approveFixedContent).toBe(false);
    expect(result.current.approveDestructive).toBe(false);
    expect(result.current.verifiedPlanIds.size).toBe(0);
    expect(result.current.saveEligibility).toBeNull();
  });

  it("should begin save candidate and reset approvals", () => {
    const { result } = renderHook(() => useExecutionApproval());

    act(() => {
      result.current.setApproveFixedContent(true);
      result.current.setApproveDestructive(true);
      result.current.beginSaveCandidate(plan);
    });

    expect(result.current.saveCandidate).toBe(plan);
    expect(result.current.approveFixedContent).toBe(false);
    expect(result.current.approveDestructive).toBe(false);
  });

  it("should close save candidate and reset approvals", () => {
    const { result } = renderHook(() => useExecutionApproval());

    act(() => result.current.beginSaveCandidate(plan));
    act(() => result.current.closeSaveCandidate());

    expect(result.current.saveCandidate).toBeNull();
    expect(result.current.approveFixedContent).toBe(false);
    expect(result.current.approveDestructive).toBe(false);
  });

  it("should mark a plan as verified", () => {
    const { result } = renderHook(() => useExecutionApproval());

    act(() => result.current.markPlanVerified("plan-a"));
    act(() => result.current.markPlanVerified("plan-b"));

    expect(result.current.verifiedPlanIds.has("plan-a")).toBe(true);
    expect(result.current.verifiedPlanIds.has("plan-b")).toBe(true);
    expect(result.current.verifiedPlanIds.size).toBe(2);
  });

  it("should keep verified set stable for duplicate plan ids", () => {
    const { result } = renderHook(() => useExecutionApproval());

    act(() => result.current.markPlanVerified("plan-a"));
    act(() => result.current.markPlanVerified("plan-a"));

    expect(result.current.verifiedPlanIds.size).toBe(1);
  });

  it("should derive save eligibility from the candidate", () => {
    analyzeToolEligibilityMock.mockReturnValue({
      issues: [
        {
          code: "FIXED_CONTENT",
          severity: "approval",
          approval: "fixedContent",
          actionIndexes: [0],
          message: "包含固定内容"
        }
      ],
      requiredApprovals: ["fixedContent"],
      blocked: false
    });
    const { result } = renderHook(() => useExecutionApproval());

    act(() => result.current.beginSaveCandidate(plan));

    expect(analyzeToolEligibilityMock).toHaveBeenCalledWith(plan);
    expect(result.current.saveEligibility?.requiredApprovals).toEqual([
      "fixedContent"
    ]);
  });

  it("should expose approval setters", () => {
    const { result } = renderHook(() => useExecutionApproval());

    act(() => {
      result.current.setApproveFixedContent(true);
      result.current.setApproveDestructive(true);
    });

    expect(result.current.approveFixedContent).toBe(true);
    expect(result.current.approveDestructive).toBe(true);
  });
});
