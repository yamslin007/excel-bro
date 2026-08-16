// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import type {
  AnalysisPlan,
  WorkbookSnapshot
} from "../contracts";
import {
  deleteQueryTool,
  deleteTool,
  loadQueryTools,
  loadTools,
  saveQueryTool,
  saveTool,
  type SavedQueryTool,
  type SavedTool,
  type ToolParameter
} from "../storage";
import { renderToolDsl } from "../toolDsl";
import { useToolManagement } from "./useToolManagement";

vi.mock("../storage", () => ({
  loadTools: vi.fn(),
  loadQueryTools: vi.fn(),
  saveTool: vi.fn(),
  deleteTool: vi.fn(),
  saveQueryTool: vi.fn(),
  deleteQueryTool: vi.fn()
}));

vi.mock("../toolDsl", () => ({
  renderToolDsl: vi.fn()
}));

const loadToolsMock = vi.mocked(loadTools);
const loadQueryToolsMock = vi.mocked(loadQueryTools);
const saveToolMock = vi.mocked(saveTool);
const deleteToolMock = vi.mocked(deleteTool);
const saveQueryToolMock = vi.mocked(saveQueryTool);
const deleteQueryToolMock = vi.mocked(deleteQueryTool);
const renderToolDslMock = vi.mocked(renderToolDsl);

const workbook: WorkbookSnapshot = {
  name: "book.xlsx",
  capturedAt: "2026-08-16T00:00:00.000Z",
  activeWorksheet: "Sheet1",
  worksheets: [
    {
      name: "Sheet1",
      usedRange: "A1:B2",
      rowCount: 2,
      columnCount: 2,
      headers: ["Name", "Amount"],
      dataRows: [],
      truncated: false
    },
    {
      name: "Sheet2",
      usedRange: "A1:C3",
      rowCount: 3,
      columnCount: 3,
      headers: ["Name", "Amount", "Qty"],
      dataRows: [],
      truncated: false
    }
  ]
};

const worksheetParameter: ToolParameter = {
  id: "source-sheet",
  label: "来源表",
  defaultValue: "Sheet1",
  required: true,
  type: "worksheet",
  bindings: [{ actionIndex: 0, property: "sheet" }]
};

const fieldParameter: ToolParameter = {
  id: "field-name",
  label: "名称字段",
  defaultValue: "Name",
  required: true,
  type: "field",
  sourceParameterId: "source-sheet",
  bindings: [{ actionIndex: 0, property: "groupBy", itemIndex: 0 }]
};

const rangeParameter: ToolParameter = {
  id: "source-range",
  label: "来源范围",
  defaultValue: "A1:B2",
  required: true,
  type: "range",
  sourceParameterId: "source-sheet",
  bindings: [{ actionIndex: 0, property: "range" }]
};

const savedTool: SavedTool = {
  id: "tool-1",
  version: 2,
  name: "Tool",
  description: "desc",
  createdAt: "2026-08-16T00:00:00.000Z",
  updatedAt: "2026-08-16T00:00:00.000Z",
  verifiedAt: "2026-08-16T00:00:00.000Z",
  planTemplate: {} as AnalysisPlan,
  parameters: [worksheetParameter, fieldParameter, rangeParameter],
  approvals: []
};

const savedQueryTool = {
  id: "query-1",
  version: 1,
  name: "Query",
  description: "query",
  sourceMode: "workbook",
  request: { tool: "query_table", arguments: {} },
  sourceSheetNames: ["Sheet1"],
  sourceSheetIds: [],
  expectedHeaders: [],
  createdAt: "2026-08-16T00:00:00.000Z",
  updatedAt: "2026-08-16T00:00:00.000Z"
} as unknown as SavedQueryTool;

describe("useToolManagement", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.clearAllMocks();
    loadToolsMock.mockReturnValue([savedTool]);
    loadQueryToolsMock.mockReturnValue([savedQueryTool]);
    saveToolMock.mockImplementation((tool) => [tool]);
    deleteToolMock.mockReturnValue([]);
    saveQueryToolMock.mockImplementation((tool) => [tool]);
    deleteQueryToolMock.mockReturnValue([]);
    renderToolDslMock.mockReturnValue("DSL");
    Object.defineProperty(navigator, "clipboard", {
      configurable: true,
      value: {
        writeText: vi.fn().mockResolvedValue(undefined)
      }
    });
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("should initialize tools and query tools from storage", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    expect(result.current.tools).toEqual([savedTool]);
    expect(result.current.queryTools).toEqual([savedQueryTool]);
    expect(result.current.toolDrawerView).toBe("library");
  });

  it("should save a workflow tool", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    act(() => result.current.saveTool(savedTool));

    expect(saveToolMock).toHaveBeenCalledWith(savedTool);
    expect(result.current.tools).toEqual([savedTool]);
  });

  it("should delete a workflow tool and clear selection", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    act(() => result.current.setSelectedToolId("tool-1"));
    act(() => result.current.deleteTool("tool-1"));

    expect(deleteToolMock).toHaveBeenCalledWith("tool-1");
    expect(result.current.tools).toEqual([]);
    expect(result.current.selectedToolId).toBeNull();
  });

  it("should save and delete query tools", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    act(() => result.current.saveQueryTool(savedQueryTool));
    expect(result.current.queryTools).toEqual([savedQueryTool]);

    act(() => result.current.deleteQueryTool("query-1"));
    expect(deleteQueryToolMock).toHaveBeenCalledWith("query-1");
    expect(result.current.queryTools).toEqual([]);
  });

  it("should request and confirm workflow tool deletion", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    act(() =>
      result.current.requestToolDeletion("workflow", {
        id: "tool-1",
        name: "Tool"
      })
    );
    expect(result.current.pendingToolDeletion).toMatchObject({
      kind: "workflow",
      id: "tool-1",
      name: "Tool"
    });

    act(() => result.current.confirmToolDeletion());
    expect(deleteToolMock).toHaveBeenCalledWith("tool-1");
    expect(result.current.pendingToolDeletion).toBeNull();
    expect(result.current.toolDrawerView).toBe("library");
  });

  it("should open workflow tool detail", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    act(() => result.current.openWorkflowToolDetail(savedTool));

    expect(result.current.selectedToolId).toBe("tool-1");
    expect(result.current.selectedQueryToolId).toBeNull();
    expect(result.current.toolDetailMode).toBe("standard");
    expect(result.current.toolDrawerView).toBe("detail");
  });

  it("should open query tool detail", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    act(() => result.current.openQueryToolDetail(savedQueryTool));

    expect(result.current.selectedQueryToolId).toBe("query-1");
    expect(result.current.selectedToolId).toBeNull();
    expect(result.current.toolDrawerView).toBe("detail");
  });

  it("should reset tool drawer state", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    act(() => {
      result.current.openWorkflowToolDetail(savedTool);
      result.current.requestToolDeletion("workflow", {
        id: "tool-1",
        name: "Tool"
      });
    });
    act(() => result.current.resetToolDrawer());

    expect(result.current.toolDrawerView).toBe("library");
    expect(result.current.selectedToolId).toBeNull();
    expect(result.current.selectedQueryToolId).toBeNull();
    expect(result.current.pendingToolDeletion).toBeNull();
  });

  it("should prepare parameters when selecting a tool", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    act(() => result.current.selectTool(savedTool, workbook));

    expect(result.current.toolParameterValues["source-sheet"]).toBe("Sheet1");
    expect(result.current.toolParameterValues["source-range"]).toBe("A1:B2");
    expect(result.current.toolParameterValues["field-name"]).toBe("Name");
  });

  it("should infer field options from the selected worksheet", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    act(() => result.current.selectTool(savedTool, workbook));
    const options = result.current.fieldOptions(
      savedTool,
      fieldParameter as Extract<ToolParameter, { type: "field" }>
    );

    expect(options).toEqual(["Name", "Amount"]);
  });

  it("should update worksheet and dependent parameters", () => {
    const { result } = renderHook(() =>
      useToolManagement({ workbook })
    );

    act(() => result.current.selectTool(savedTool, workbook));
    act(() =>
      result.current.updateToolParameter(
        savedTool,
        worksheetParameter,
        "Sheet2"
      )
    );

    expect(result.current.toolParameterValues["source-sheet"]).toBe("Sheet2");
    expect(result.current.toolParameterValues["source-range"]).toBe("A1:C3");
    expect(result.current.toolParameterValues["field-name"]).toBe("Name");
  });

  it("should copy tool DSL and clear feedback after timeout", async () => {
    const onToolDslCopyError = vi.fn();
    const { result } = renderHook(() =>
      useToolManagement({ workbook, onToolDslCopyError })
    );

    await act(async () => {
      await result.current.copyToolDsl(savedTool);
    });

    expect(renderToolDslMock).toHaveBeenCalledWith(savedTool.planTemplate);
    expect(result.current.copiedToolDslId).toBe("tool-1");

    act(() => {
      vi.advanceTimersByTime(1600);
    });

    expect(result.current.copiedToolDslId).toBeNull();
    expect(onToolDslCopyError).not.toHaveBeenCalled();
  });

  it("should report DSL copy failure", async () => {
    renderToolDslMock.mockReturnValue("DSL");
    const onToolDslCopyError = vi.fn();
    Object.defineProperty(navigator, "clipboard", {
      configurable: true,
      value: {
        writeText: vi.fn().mockRejectedValue(new Error("denied"))
      }
    });
    const { result } = renderHook(() =>
      useToolManagement({ workbook, onToolDslCopyError })
    );

    await act(async () => {
      await result.current.copyToolDsl(savedTool);
    });

    expect(onToolDslCopyError).toHaveBeenCalled();
  });
});
