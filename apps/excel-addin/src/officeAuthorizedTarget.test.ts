import { afterEach, describe, expect, it, vi } from "vitest";
import { inspectOfficeAuthorizedTarget, inspectOfficeScopeAdmission, officeAnalysisTargetEvidence } from "./officeAuthorizedTarget";
afterEach(() => vi.unstubAllGlobals());
function host() {
  const document = { url: "file:///C:/Data/Original.xlsx" };
  const sheet = { id: "stable-1", name: "Sheet1", load: vi.fn(), getRange: vi.fn(() => { throw new Error("must not read cells"); }) };
  const other = { id: "stable-2", name: "Other" };
  const collection = { items: [sheet, other], load: vi.fn(), getActiveWorksheet: () => sheet };
  const sync = vi.fn(async () => {});
  vi.stubGlobal("Office", { context: { document } });
  vi.stubGlobal("Excel", { run: async (job: (context: unknown) => unknown) => job({ workbook: { worksheets: collection }, sync }) });
  return { document, sheet, sync };
}
describe("Office target metadata adapter", () => {
  it("a picker catalog is not reused for a different document with the same filename", async () => {
    const h = host();
    const visible = await inspectOfficeAuthorizedTarget(["Sheet1"]);
    vi.stubGlobal("Office", { context: { document: { url: h.document.url } } });
    const report = await inspectOfficeScopeAdmission(["Sheet1"], visible);
    expect(report.items).toEqual([{ id: "Sheet1", status: "unavailable", reason: "target_changed" }]);
    expect(h.sheet.getRange).not.toHaveBeenCalled();
  });
  it("a missing selected sheet gets a precise unavailable result, not a reduced authorization", async () => {
    const h = host();
    const report = await inspectOfficeScopeAdmission(["Sheet1", "Deleted"]);
    expect(report.items).toEqual([
      { id: "Sheet1", status: "read_only", reason: "write_unverified" },
      { id: "Deleted", status: "unavailable", reason: "missing" }
    ]);
    expect(report.target).toBeUndefined();
    expect(h.sheet.getRange).not.toHaveBeenCalled();
  });
  it.each([[true, false, "read_only"], [false, true, "protected"], [false, false, "none"]])(
    "selection checks read-only/protection metadata without reading cells: %s/%s", async (readOnly, protectedSheet, reason) => {
      const h = host();
      const protection = { protected: protectedSheet, load: vi.fn() };
      const workbook = { readOnly, load: vi.fn(), worksheets: {
        items: [h.sheet], load: vi.fn(), getActiveWorksheet: () => h.sheet,
        getItem: vi.fn((id: string) => { expect(id).toBe(h.sheet.id); return { protection }; })
      } };
      vi.stubGlobal("Office", { context: { document: h.document, requirements: { isSetSupported: () => true } } });
      vi.stubGlobal("Excel", { run: async (job: (context: unknown) => unknown) => job({ workbook, sync: h.sync }) });
      const report = await inspectOfficeScopeAdmission(["Sheet1"]);
      expect(report.items).toEqual([{ id: "Sheet1", reason, status: reason === "none" ? "ready" : "read_only" }]);
      expect(h.sheet.getRange).not.toHaveBeenCalled();
      expect(workbook.worksheets.getItem).toHaveBeenCalledWith("stable-1");
    });
  it("unknown write capability is not falsely advertised as writable", async () => {
    const h = host();
    const report = await inspectOfficeScopeAdmission(["Sheet1"]);
    expect(report.items[0]).toMatchObject({ status: "read_only", reason: "write_unverified" });
    expect(h.sheet.getRange).not.toHaveBeenCalled();
  });
  it("analysis evidence is document-scoped, needs no source cells, and cannot authorize a same-URL replacement document", async () => {
    const h = host(); const target = await inspectOfficeAuthorizedTarget(["Sheet1"]);
    expect(await officeAnalysisTargetEvidence(target)).toMatchObject({ conclusive: true });
    expect(h.sheet.getRange).not.toHaveBeenCalled();
    vi.stubGlobal("Office", { context: { document: { url: h.document.url } } });
    expect(await officeAnalysisTargetEvidence(target)).toBeNull();
  });
  it("an unsaved stable Office document permits analysis but a changed saved URL does not reuse the binding", async () => {
    const h = host(); h.document.url = "";
    const target = await inspectOfficeAuthorizedTarget(["Sheet1"]);
    expect(await officeAnalysisTargetEvidence(target)).toMatchObject({ conclusive: true });
    h.document.url = "file:///D:/Elsewhere/Original.xlsx";
    expect(await officeAnalysisTargetEvidence(target)).toBeNull();
    expect(h.sheet.getRange).not.toHaveBeenCalled();
  });
  it("reads stable sheet identities, no raw values/usedRange, and never expands manual selection", async () => {
    const h = host(); const target = await inspectOfficeAuthorizedTarget(["Sheet1"]);
    expect(target.selected).toEqual(["stable-1"]); expect(target.sheets).toHaveLength(2);
    expect(h.sheet.getRange).not.toHaveBeenCalled();
    expect(target.name).toBe("Original.xlsx");
  });
  it("rejects a document switch in the middle of metadata sync", async () => {
    const h = host(); h.sync.mockImplementation(async () => { h.document.url = "file:///C:/Data/Other.xlsx"; });
    await expect(inspectOfficeAuthorizedTarget(null)).rejects.toThrow("检查期间");
  });
});
