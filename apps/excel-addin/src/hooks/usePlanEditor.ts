import { useCallback, useRef, useState } from "react";
import type { AnalysisPlan } from "../contracts";
import { assertEditablePlan } from "../contracts";
import { updateToolPlanTemplate, type SavedTool } from "../storage";
import {
  isToolSchemaFingerprint,
  toolSchemaFingerprintForSnapshot
} from "../excel";
import type { WorkbookSnapshot } from "../contracts";

export type PlanEditorPhase = "closed" | "editing" | "error";

export interface PlanEditorState {
  phase: PlanEditorPhase;
  /** JSON 文本，编辑中 */
  draft: string;
  /** 校验错误文案（phase=error 时有值） */
  errorMessage: string;
}

/** 编辑对象：消息里的计划，或固化工具的计划模板。 */
export type PlanEditorTarget =
  | { kind: "message"; messageId: string }
  | { kind: "savedTool"; toolId: string };

interface UsePlanEditorOptions {
  /**
   * 保存成功后回调。target 说明改的是消息里的计划还是固化工具；改工具时
   * updatedTools 是写回后的完整列表，调用方直接用它刷新状态即可。
   */
  onSave: (
    plan: AnalysisPlan,
    target: PlanEditorTarget,
    updatedTools?: SavedTool[]
  ) => void;
}

interface UsePlanEditorResult {
  state: PlanEditorState;
  open: (plan: AnalysisPlan, target: PlanEditorTarget) => void;
  close: () => void;
  setDraft: (text: string) => void;
  save: (workbook?: Pick<WorkbookSnapshot, "worksheets">) => Promise<void>;
}

const CLOSED: PlanEditorState = {
  phase: "closed",
  draft: "",
  errorMessage: ""
};

export function usePlanEditor(options: UsePlanEditorOptions): UsePlanEditorResult {
  const [state, setState] = useState<PlanEditorState>(CLOSED);
  const abortRef = useRef<AbortController | null>(null);
  const targetRef = useRef<PlanEditorTarget | null>(null);

  const open = useCallback((plan: AnalysisPlan, target: PlanEditorTarget) => {
    targetRef.current = target;
    setState({
      phase: "editing",
      draft: JSON.stringify(plan, null, 2),
      errorMessage: ""
    });
  }, []);

  const close = useCallback(() => {
    abortRef.current?.abort();
    targetRef.current = null;
    setState(CLOSED);
  }, []);

  const setDraft = useCallback((text: string) => {
    setState((previous) => ({
      ...previous,
      draft: text,
      errorMessage: "",
      phase: "editing"
    }));
  }, []);

  const save = useCallback(
    async (workbook?: Pick<WorkbookSnapshot, "worksheets">) => {
      // 本地 JSON 解析 + 白名单结构校验（快速保存，不请求服务端）。
      let parsed: unknown;
      try {
        parsed = JSON.parse(state.draft);
      } catch {
        setState((previous) => ({
          ...previous,
          phase: "error",
          errorMessage: "JSON 格式有误，请检查括号和引号是否匹配"
        }));
        return;
      }
      try {
        assertEditablePlan(parsed);
      } catch (reason) {
        setState((previous) => ({
          ...previous,
          phase: "error",
          errorMessage:
            reason instanceof Error ? reason.message : "计划结构不合法"
        }));
        return;
      }

      let normalized = parsed as AnalysisPlan;

      const target = targetRef.current;
      if (!target) return;

      // 固化工具：把规范化后的计划写回 storage，参数与审批项同步重算。
      let updatedTools: SavedTool[] | undefined;
      if (target.kind === "savedTool") {
        try {
          updatedTools = updateToolPlanTemplate(target.toolId, normalized);
        } catch (reason) {
          setState((previous) => ({
            ...previous,
            phase: "error",
            errorMessage:
              reason instanceof Error ? reason.message : "写回工具失败"
          }));
          return;
        }
      }

      // 指纹必须跟着计划一起重算：动作改了，schema 指纹依赖的「引用了哪些字段」
      // 也就变了；沿用旧指纹会在执行时误判成「来源已变化」。只有原本就是 schema
      // 指纹的计划（固化工具）才能这样重算 —— 值指纹要读实时单元格，不在这里做。
      if (
        workbook &&
        normalized.sourceFingerprint &&
        isToolSchemaFingerprint(normalized.sourceFingerprint)
      ) {
        normalized = {
          ...normalized,
          sourceFingerprint: toolSchemaFingerprintForSnapshot(
            normalized,
            workbook
          )
        };
      }

      options.onSave(normalized, target, updatedTools);
      targetRef.current = null;
      setState(CLOSED);
    },
    [state.draft, options]
  );

  return { state, open, close, setDraft, save };
}
