import { afterEach, describe, expect, it, vi } from "vitest";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { createOfficePythonDataHost } from "./officePythonData";
import { inspectOfficeAuthorizedTarget, officeAnalysisTargetEvidence } from "./officeAuthorizedTarget";

vi.mock("./officeAuthorizedTarget", () => ({ inspectOfficeAuthorizedTarget: vi.fn(), officeAnalysisTargetEvidence: vi.fn() }));
afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });

function fixture() {
  const bound: AuthorizedTarget = { fileKey: "url:file.xlsx", fileRef: {}, name: "file.xlsx",
    sheets: [{ id: "id-1", name: "Sheet1" }], selected: ["id-1"] };
  let current: AuthorizedTarget | null = bound;
  vi.mocked(inspectOfficeAuthorizedTarget).mockImplementation(async () => current!);
  vi.mocked(officeAnalysisTargetEvidence).mockResolvedValue({ key: "verified", conclusive: true });
  const write = vi.fn(), load = vi.fn(), getRange = vi.fn();
  const range = { load, get values() { return [[7]]; }, set values(_value) { write(); } };
  getRange.mockReturnValue(range);
  const context = { workbook: { worksheets: { getItem: vi.fn(() => ({ getRange })) } }, sync: vi.fn(async () => {}) };
  vi.stubGlobal("Excel", { run: vi.fn(async (work) => work(context)) });
  return { bound, current: () => current, change: (value: AuthorizedTarget | null) => { current = value; }, write, getRange, getItem: context.workbook.worksheets.getItem };
}

describe("Office Python readonly data adapter", () => {
  it("reads only authorized source through the real primitive, zero writes", async () => {
    const f = fixture(); const host = createOfficePythonDataHost(f.bound, f.current);
    expect(await host.read({ sheetId: "id-1", sheetName: "Sheet1", range: "A1" })).toEqual([[7]]);
    expect(f.getRange).toHaveBeenCalledWith("A1"); expect(f.write).not.toHaveBeenCalled();
    expect(f.getItem).toHaveBeenCalledWith("id-1");
  });
  it.each(["file", "scope", "unavailable"])("rejects %s before cell reads", async (change) => {
    const f = fixture(); const host = createOfficePythonDataHost(f.bound, f.current);
    f.change(change === "unavailable" ? null : { ...f.bound, ...(change === "file" ? { fileKey: "other" } : { selected: [] }) });
    await expect(host.read({ sheetId: "id-1", sheetName: "Sheet1", range: "A1" })).rejects.toThrow();
    expect(f.getRange).not.toHaveBeenCalled(); expect(f.write).not.toHaveBeenCalled();
  });
  it("rejects non-conclusive evidence and an in-place authorization change", async () => {
    const f = fixture(); const host = createOfficePythonDataHost(f.bound, f.current);
    vi.mocked(officeAnalysisTargetEvidence).mockResolvedValue(null);
    await expect(host.assertCurrent()).rejects.toThrow();
    f.bound.selected.length = 0;
    await expect(host.assertCurrent()).rejects.toThrow();
  });
});
