import { afterEach, describe, expect, it, vi } from "vitest";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import type { WpsApplication } from "./wpsHost";
import { createWpsPythonDataHost } from "./wpsPythonData";
import { getWpsApplication } from "./wpsHost";
import { inspectWpsAuthorizedTarget, wpsAnalysisTargetEvidence } from "./wpsAuthorizedTarget";
import { captureWpsObjectIdentity, verifyWpsObjectIdentity } from "./wpsObjectIdentity";

vi.mock("./wpsHost", () => ({ getWpsApplication: vi.fn() }));
vi.mock("./wpsAuthorizedTarget", () => ({ inspectWpsAuthorizedTarget: vi.fn(), wpsAnalysisTargetEvidence: vi.fn() }));
vi.mock("./wpsObjectIdentity", () => ({ captureWpsObjectIdentity: vi.fn(), verifyWpsObjectIdentity: vi.fn() }));
afterEach(() => { vi.clearAllMocks(); });

function fixture() {
  const write = vi.fn(), read = vi.fn(() => 7);
  const range = { Rows: { Count: 1 }, Columns: { Count: 1 }, get Value2() { return read(); }, set Value2(_value) { write(); } };
  const sheet = { Name: "Sheet1", Range: vi.fn(() => range) };
  const workbook = { Name: "file.xlsx", FullName: "C:\\Data\\file.xlsx", Path: "C:\\Data", Worksheets: { Count: 1, Item: () => sheet } };
  const app = { ActiveWorkbook: workbook } as unknown as WpsApplication;
  vi.mocked(getWpsApplication).mockReturnValue(app);
  const bound: AuthorizedTarget = { fileKey: workbook.FullName, fileRef: workbook, name: workbook.Name,
    sheets: [{ id: "1", name: "Sheet1" }], selected: ["1"] };
  let current: AuthorizedTarget | null = bound;
  vi.mocked(inspectWpsAuthorizedTarget).mockImplementation(async () => current!);
  vi.mocked(wpsAnalysisTargetEvidence).mockResolvedValue({ key: "verified", conclusive: true });
  vi.mocked(captureWpsObjectIdentity).mockReturnValue({ workbookRef: workbook, workbookName: workbook.Name,
    workbookFullName: workbook.FullName, workbookPath: workbook.Path, sheets: [{ index: 1, name: sheet.Name, ref: sheet }] });
  vi.mocked(verifyWpsObjectIdentity).mockReturnValue("same");
  return { bound, current: () => current, change: (value: AuthorizedTarget | null) => { current = value; }, write, read };
}

describe("WPS Python readonly data adapter", () => {
  it("reads authorized source with real authorization and primitive, zero writes", async () => {
    const f = fixture(); const host = createWpsPythonDataHost(f.bound, f.current);
    expect(await host.read({ sheetId: "1", sheetName: "Sheet1", range: "A1" })).toEqual([[7]]);
    expect(f.read).toHaveBeenCalledOnce(); expect(f.write).not.toHaveBeenCalled();
  });
  it.each(["file", "scope", "unavailable"])("rejects %s before cells", async (change) => {
    const f = fixture(); const host = createWpsPythonDataHost(f.bound, f.current);
    f.change(change === "unavailable" ? null : { ...f.bound, ...(change === "file" ? { fileKey: "other" } : { selected: [] }) });
    await expect(host.read({ sheetId: "1", sheetName: "Sheet1", range: "A1" })).rejects.toThrow();
    expect(f.read).not.toHaveBeenCalled(); expect(f.write).not.toHaveBeenCalled();
  });
  it("unknown/new wrapper is not authorization", async () => {
    const f = fixture(); const host = createWpsPythonDataHost(f.bound, f.current);
    vi.mocked(verifyWpsObjectIdentity).mockReturnValue("unknown");
    await expect(host.read({ sheetId: "1", sheetName: "Sheet1", range: "A1" })).rejects.toThrow();
    expect(f.read).not.toHaveBeenCalled();
  });
});
