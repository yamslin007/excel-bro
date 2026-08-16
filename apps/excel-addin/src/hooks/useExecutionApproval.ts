import { useCallback, useMemo, useState } from "react";
import type { AnalysisPlan } from "../contracts";
import {
  analyzeToolEligibility,
  type ToolEligibility
} from "../storage";

/**
 * 执行验收与工具固化批准 Hook
 *
 * 职责：
 * - 记录哪些计划已经执行并通过验证，允许用户“保存为工具”
 * - 管理“保存为工具”的候选计划和风险批准状态
 * - 派生固化资格检查结果，供确认对话框使用
 */
export function useExecutionApproval() {
  const [verifiedPlanIds, setVerifiedPlanIds] = useState<Set<string>>(
    () => new Set()
  );
  const [saveCandidate, setSaveCandidate] = useState<AnalysisPlan | null>(null);
  const [approveFixedContent, setApproveFixedContent] = useState(false);
  const [approveDestructive, setApproveDestructive] = useState(false);

  const saveEligibility = useMemo<ToolEligibility | null>(
    () => (saveCandidate ? analyzeToolEligibility(saveCandidate) : null),
    [saveCandidate]
  );

  const markPlanVerified = useCallback((planId: string) => {
    setVerifiedPlanIds((current) => {
      if (current.has(planId)) {
        return current;
      }
      return new Set(current).add(planId);
    });
  }, []);

  const beginSaveCandidate = useCallback((plan: AnalysisPlan) => {
    setSaveCandidate(plan);
    setApproveFixedContent(false);
    setApproveDestructive(false);
  }, []);

  const closeSaveCandidate = useCallback(() => {
    setSaveCandidate(null);
    setApproveFixedContent(false);
    setApproveDestructive(false);
  }, []);

  return {
    saveCandidate,
    setSaveCandidate,
    approveFixedContent,
    setApproveFixedContent,
    approveDestructive,
    setApproveDestructive,
    verifiedPlanIds,
    markPlanVerified,
    saveEligibility,
    beginSaveCandidate,
    closeSaveCandidate
  };
}
