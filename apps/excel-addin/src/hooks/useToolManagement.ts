import { useCallback, useEffect, useRef, useState } from "react";
import type { WorkbookSnapshot } from "../contracts";
import {
  deleteQueryTool as removeSavedQueryTool,
  deleteTool as removeSavedTool,
  loadQueryTools,
  loadTools,
  saveQueryTool as persistQueryTool,
  saveTool as persistTool,
  type SavedQueryTool,
  type SavedTool,
  type ToolParameter
} from "../storage";
import { renderToolDsl } from "../toolDsl";
import { copyTextToClipboard } from "../utils";

export type ToolDrawerView = "library" | "detail" | "run";
export type ToolDetailMode = "standard" | "expert";
export type PendingToolDeletion = {
  kind: "workflow" | "query";
  id: string;
  name: string;
};

interface UseToolManagementOptions {
  workbook: WorkbookSnapshot | null;
  onToolDslCopyError?: () => void;
}

/**
 * 我的工具管理 Hook
 *
 * 职责：
 * - 管理工作流工具和本地查询工具的持久化状态
 * - 管理工具抽屉的选择、视图、删除确认和参数状态
 * - 提供工具保存、删除、参数准备和 DSL 复制等操作
 */
export function useToolManagement({
  workbook,
  onToolDslCopyError
}: UseToolManagementOptions) {
  const [tools, setTools] = useState<SavedTool[]>(loadTools);
  const [queryTools, setQueryTools] = useState<SavedQueryTool[]>(loadQueryTools);
  const [selectedToolId, setSelectedToolId] = useState<string | null>(null);
  const [selectedQueryToolId, setSelectedQueryToolId] = useState<string | null>(null);
  const [toolDrawerView, setToolDrawerView] = useState<ToolDrawerView>("library");
  const [toolDetailMode, setToolDetailMode] = useState<ToolDetailMode>("standard");
  const [pendingToolDeletion, setPendingToolDeletion] =
    useState<PendingToolDeletion | null>(null);
  const [copiedToolDslId, setCopiedToolDslId] = useState<string | null>(null);
  const [toolParameterValues, setToolParameterValues] = useState<Record<string, string>>({});
  const toolDslCopyTimerRef = useRef<number | null>(null);

  useEffect(
    () => () => {
      if (toolDslCopyTimerRef.current !== null) {
        window.clearTimeout(toolDslCopyTimerRef.current);
      }
    },
    []
  );

  const saveTool = useCallback((tool: SavedTool) => {
    const next = persistTool(tool);
    setTools(next);
    return next;
  }, []);

  const saveQueryTool = useCallback((tool: SavedQueryTool) => {
    const next = persistQueryTool(tool);
    setQueryTools(next);
    return next;
  }, []);

  const deleteTool = useCallback((toolId: string) => {
    const next = removeSavedTool(toolId);
    setTools(next);
    setSelectedToolId((current) => (current === toolId ? null : current));
    return next;
  }, []);

  const deleteQueryTool = useCallback((toolId: string) => {
    const next = removeSavedQueryTool(toolId);
    setQueryTools(next);
    setSelectedQueryToolId((current) =>
      current === toolId ? null : current
    );
    return next;
  }, []);

  const requestToolDeletion = useCallback(
    (
      kind: PendingToolDeletion["kind"],
      tool: Pick<SavedTool | SavedQueryTool, "id" | "name">
    ) => {
      setPendingToolDeletion({ kind, id: tool.id, name: tool.name });
    },
    []
  );

  const confirmToolDeletion = useCallback(() => {
    if (!pendingToolDeletion) return;
    if (pendingToolDeletion.kind === "workflow") {
      deleteTool(pendingToolDeletion.id);
    } else {
      deleteQueryTool(pendingToolDeletion.id);
    }
    setPendingToolDeletion(null);
    setToolDrawerView("library");
  }, [deleteQueryTool, deleteTool, pendingToolDeletion]);

  const resetToolDrawer = useCallback(() => {
    setToolDrawerView("library");
    setSelectedToolId(null);
    setSelectedQueryToolId(null);
    setToolDetailMode("standard");
    setPendingToolDeletion(null);
  }, []);

  const openWorkflowToolDetail = useCallback((tool: SavedTool) => {
    setSelectedToolId(tool.id);
    setSelectedQueryToolId(null);
    setToolDetailMode("standard");
    setToolDrawerView("detail");
  }, []);

  const openQueryToolDetail = useCallback((tool: SavedQueryTool) => {
    setSelectedQueryToolId(tool.id);
    setSelectedToolId(null);
    setToolDetailMode("standard");
    setToolDrawerView("detail");
  }, []);

  const selectTool = useCallback(
    (tool: SavedTool, snapshot: WorkbookSnapshot | null = workbook) => {
      setSelectedToolId(tool.id);
      const available = new Set(
        (snapshot?.worksheets ?? []).map((sheet) => sheet.name)
      );
      const fallback =
        snapshot?.activeWorksheet ?? snapshot?.worksheets[0]?.name ?? "";
      const suggestOutputName = (requested: string) => {
        if (
          ![...available].some(
            (name) =>
              name.toLocaleLowerCase() === requested.toLocaleLowerCase()
          )
        ) {
          return requested;
        }
        for (let index = 2; index < 1000; index += 1) {
          const suffix = ` (${index})`;
          const candidate = `${requested.slice(0, 31 - suffix.length)}${suffix}`;
          if (
            ![...available].some(
              (name) =>
                name.toLocaleLowerCase() === candidate.toLocaleLowerCase()
            )
          ) {
            return candidate;
          }
        }
        return requested.slice(0, 27) + " 副本";
      };
      const values: Record<string, string> = {};
      for (const parameter of tool.parameters) {
        if (parameter.type === "worksheet") {
          values[parameter.id] = available.has(parameter.defaultValue)
            ? parameter.defaultValue
            : fallback;
        } else if (parameter.type === "outputWorksheet") {
          values[parameter.id] = suggestOutputName(parameter.defaultValue);
        }
      }
      for (const parameter of tool.parameters) {
        if (parameter.type === "range") {
          const source = parameter.sourceParameterId
            ? values[parameter.sourceParameterId]
            : "";
          values[parameter.id] =
            snapshot?.worksheets.find((sheet) => sheet.name === source)
              ?.usedRange ?? parameter.defaultValue;
        } else if (parameter.type === "field") {
          const source = values[parameter.sourceParameterId];
          const headers = (
            snapshot?.worksheets.find((sheet) => sheet.name === source)
              ?.headers ?? []
          )
            .map((header) => String(header ?? "").trim())
            .filter(Boolean);
          values[parameter.id] =
            headers.find(
              (header) =>
                header.toLocaleLowerCase() ===
                parameter.defaultValue.toLocaleLowerCase()
            ) ?? "";
        }
      }
      setToolParameterValues(values);
    },
    [workbook]
  );

  const fieldOptions = useCallback(
    (
      tool: SavedTool,
      parameter: Extract<ToolParameter, { type: "field" }>
    ): string[] => {
      const sourceParameter = tool.parameters.find(
        (candidate) => candidate.id === parameter.sourceParameterId
      );
      const sourceSheet =
        toolParameterValues[parameter.sourceParameterId] ??
        sourceParameter?.defaultValue;
      const sheet = workbook?.worksheets.find(
        (candidate) => candidate.name === sourceSheet
      );
      const detected = [
        ...new Set(
          (sheet?.headers ?? [])
            .map((header) => String(header ?? "").trim())
            .filter(Boolean)
        )
      ];
      return detected.length > 0 ? detected : [parameter.defaultValue];
    },
    [toolParameterValues, workbook]
  );

  const updateToolParameter = useCallback(
    (tool: SavedTool, parameter: ToolParameter, value: string) => {
      setToolParameterValues((current) => {
        const next = { ...current, [parameter.id]: value };
        if (parameter.type !== "worksheet") return next;
        for (const candidate of tool.parameters) {
          if (
            candidate.type === "range" &&
            candidate.sourceParameterId === parameter.id
          ) {
            next[candidate.id] =
              workbook?.worksheets.find((sheet) => sheet.name === value)
                ?.usedRange ?? candidate.defaultValue;
            continue;
          }
          if (
            candidate.type !== "field" ||
            candidate.sourceParameterId !== parameter.id
          ) {
            continue;
          }
          const headers = (
            workbook?.worksheets.find((sheet) => sheet.name === value)
              ?.headers ?? []
          )
            .map((header) => String(header ?? "").trim())
            .filter(Boolean);
          next[candidate.id] =
            headers.find(
              (header) =>
                header.toLocaleLowerCase() ===
                candidate.defaultValue.toLocaleLowerCase()
            ) ?? "";
        }
        return next;
      });
    },
    [workbook]
  );

  const copyToolDsl = useCallback(
    async (tool: SavedTool) => {
      const text = renderToolDsl(tool.planTemplate);
      try {
        const copied = await copyTextToClipboard(text);
        if (!copied) throw new Error("copy failed");

        setCopiedToolDslId(tool.id);
        if (toolDslCopyTimerRef.current !== null) {
          window.clearTimeout(toolDslCopyTimerRef.current);
        }
        toolDslCopyTimerRef.current = window.setTimeout(() => {
          setCopiedToolDslId((current) =>
            current === tool.id ? null : current
          );
          toolDslCopyTimerRef.current = null;
        }, 1600);
      } catch {
        onToolDslCopyError?.();
      }
    },
    [onToolDslCopyError]
  );

  return {
    // 状态
    tools,
    queryTools,
    selectedToolId,
    selectedQueryToolId,
    toolDrawerView,
    toolDetailMode,
    pendingToolDeletion,
    copiedToolDslId,
    toolParameterValues,

    // 操作
    saveTool,
    saveQueryTool,
    deleteTool,
    deleteQueryTool,
    requestToolDeletion,
    confirmToolDeletion,
    resetToolDrawer,
    openWorkflowToolDetail,
    openQueryToolDetail,
    selectTool,
    fieldOptions,
    updateToolParameter,
    copyToolDsl,

    // 状态设置（给特殊场景使用）
    setTools,
    setQueryTools,
    setSelectedToolId,
    setSelectedQueryToolId,
    setToolDrawerView,
    setToolDetailMode,
    setPendingToolDeletion,
    setCopiedToolDslId,
    setToolParameterValues
  };
}
