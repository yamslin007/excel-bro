// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { useAuthorizedTarget } from "./hooks/useAuthorizedTarget";
import { useScopeSelection } from "./hooks/useScopeSelection";
import { confirmOfficeAuthorizedScope, inspectOfficeAuthorizedTarget, officeAnalysisTargetEvidence } from "./officeAuthorizedTarget";

afterEach(() => { cleanup(); vi.unstubAllGlobals(); });

function setup() {
  const document = { url: "file:///C:/Data/Original.xlsx" };
  const readCells = vi.fn(() => { throw new Error("confirmation must not read cells"); });
  let sheets = [
    { id: "one", name: "Sheet1", load: vi.fn(), getRange: readCells },
    { id: "two", name: "Sheet2", load: vi.fn(), getRange: readCells }
  ];
  vi.stubGlobal("Office", { context: { document } });
  const sync = vi.fn(async () => {});
  vi.stubGlobal("Excel", { run: async (job: (context: unknown) => unknown) => job({
    workbook: { worksheets: { get items() { return sheets; }, load: vi.fn(), getActiveWorksheet: () => sheets[0] } }, sync
  }) });
  const analyze = vi.fn();
  let guard!: ReturnType<typeof useAuthorizedTarget>;
  function Picker() {
    const scope = useScopeSelection();
    guard = useAuthorizedTarget({ enabled: true,
      inspect: () => inspectOfficeAuthorizedTarget(scope.committedSheetNames.length ? scope.committedSheetNames : null),
      analysisEvidence: officeAnalysisTargetEvidence,
      adopt: (target) => scope.adoptCurrentWorkbookScope({} as never,
        target.sheets.filter((s) => target.selected.includes(s.id)).map((s) => s.name)) });
    return <>
      <button onClick={() => { scope.chooseManualScope(null); scope.toggleSheet("Sheet2"); }}>勾选 Sheet2</button>
      <button onClick={() => scope.chooseAutomaticScope(null)}>跟随当前表</button>
      <button onClick={() => void (async () => {
        const mode = scope.workbookScopeMode;
        if (await confirmOfficeAuthorizedScope(mode === "auto" ? null : scope.selectedSheetNames, guard.confirmScope)) {
          scope.setWorkbookScopeMode(mode);
        }
      })()}>确认范围</button>
      <button onClick={() => void guard.checkForAnalysis().then((ok) => { if (ok) analyze(); })}>追问</button>
      <span>{scope.committedSheetNames.join(",")}</span>
      <span data-testid="scope-mode">{scope.workbookScopeMode}</span>
      {guard.notice && <div role="dialog">需要确认</div>}
    </>;
  }
  const ui = render(<Picker />);
  return { ui, analyze, readCells, document, guard: () => guard, removeSelected: () => { sheets = sheets.slice(0, 1); } };
}

describe("Office range selection and follow-up authorization", () => {
  it("commits picked sheet ids then silently rechecks multiple follow-ups, without reading cells before analysis", async () => {
    const h = setup(); await waitFor(() => expect(h.guard().binding).not.toBeNull());
    fireEvent.click(h.ui.getByText("勾选 Sheet2"));
    fireEvent.click(h.ui.getByText("确认范围"));
    await waitFor(() => expect(h.guard().binding?.selected).toEqual(["two"]));
    expect(h.readCells).not.toHaveBeenCalled(); expect(h.analyze).not.toHaveBeenCalled();
    for (let count = 1; count <= 2; count++) {
      fireEvent.click(h.ui.getByText("追问"));
      await waitFor(() => expect(h.analyze).toHaveBeenCalledTimes(count));
      expect(h.ui.queryByRole("dialog")).toBeNull();
    }
    h.document.url = "file:///D:/Other/Original.xlsx";
    fireEvent.click(h.ui.getByText("追问"));
    await waitFor(() => expect(h.ui.getByRole("dialog")).toBeTruthy());
    expect(h.analyze).toHaveBeenCalledTimes(2);
  });
  it("does not commit a selected sheet that disappeared while the picker was open", async () => {
    const h = setup(); await waitFor(() => expect(h.guard().binding).not.toBeNull());
    fireEvent.click(h.ui.getByText("勾选 Sheet2")); h.removeSelected();
    await act(async () => { fireEvent.click(h.ui.getByText("确认范围")); });
    expect(h.guard().binding?.selected).toEqual(["one"]);
    expect(h.analyze).not.toHaveBeenCalled(); expect(h.readCells).not.toHaveBeenCalled();
  });
  it("replaced Office document context requires confirmation even at the same URL", async () => {
    const h = setup(); await waitFor(() => expect(h.guard().binding).not.toBeNull());
    fireEvent.click(h.ui.getByText("勾选 Sheet2")); fireEvent.click(h.ui.getByText("确认范围"));
    await waitFor(() => expect(h.guard().binding?.selected).toEqual(["two"]));
    vi.stubGlobal("Office", { context: { document: { ...h.document } } });
    fireEvent.click(h.ui.getByText("追问"));
    await waitFor(() => expect(h.ui.getByRole("dialog")).toBeTruthy());
    expect(h.analyze).not.toHaveBeenCalled(); expect(h.readCells).not.toHaveBeenCalled();
  });
  it("confirming automatic scope retains the mode chosen in the range picker", async () => {
    const h = setup(); await waitFor(() => expect(h.guard().binding).not.toBeNull());
    fireEvent.click(h.ui.getByText("勾选 Sheet2")); fireEvent.click(h.ui.getByText("跟随当前表"));
    fireEvent.click(h.ui.getByText("确认范围"));
    await waitFor(() => expect(h.guard().binding?.selected).toEqual(["one"]));
    expect(h.ui.getByTestId("scope-mode").textContent).toBe("auto");
    expect(h.readCells).not.toHaveBeenCalled(); expect(h.analyze).not.toHaveBeenCalled();
  });
  it("keeps the real App range-confirmation entry wired to the verified adapter", () => {
    const source = readFileSync(resolve(process.cwd(), "src/App.tsx"), "utf8");
    const picker = source.slice(source.indexOf("async function confirmSheetSelection()"), source.indexOf("function buildIntentScope("));
    expect(picker).toContain("await confirmOfficeAuthorizedScope(");
    expect(picker).toContain("targetGuard.confirmScope");
    expect(picker).toContain("setWorkbookScopeMode(workbookScopeMode)");
  });
});
