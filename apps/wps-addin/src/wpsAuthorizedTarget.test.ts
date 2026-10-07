import { describe, expect, it, vi, afterEach } from "vitest";
import { isWpsReviewCopyTarget, wpsAnalysisTargetEvidence } from "./wpsAuthorizedTarget";
import { getWpsApplication } from "./wpsHost";
import type { WpsNativeComparison } from "./wpsNativeComparison";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
vi.mock("./wpsHost", async (importOriginal) => ({
  ...await importOriginal<typeof import("./wpsHost")>(), getWpsApplication: vi.fn()
}));
afterEach(() => vi.clearAllMocks());
describe("analysis evidence without source cells", () => {
  function host(churn = false) {
    const values = vi.fn();
    const sheet = { Name: "Sheet1", Index: 1, CodeName: "sheet-id", Range: values,
      UsedRange: { Row: 1, Column: 1, Rows: { Count: 6 }, Columns: { Count: 5 }, get Value2() { values(); return []; } } };
    const book = { Name: "Book.xlsx", FullName: "E:\\Private\\Book.xlsx", Path: "E:\\Private",
      Worksheets: { Count: 1, Item: () => churn ? { ...sheet } : sheet } };
    vi.mocked(getWpsApplication).mockReturnValue({ get ActiveWorkbook() {
      return churn ? { ...book } : book;
    } } as never);
    const target: AuthorizedTarget = { name: book.Name, fileKey: book.FullName, fileRef: book,
      sheets: [{ id: "1", name: "Sheet1" }], selected: ["1"] };
    return { values, target };
  }
  it("stable objects give conclusive evidence, new wrappers need explicit upfront consent; neither reads values", async () => {
    const stable = host();
    expect(await wpsAnalysisTargetEvidence(stable.target)).toEqual(expect.objectContaining({ conclusive: true }));
    expect(stable.values).not.toHaveBeenCalled();
    const wrapped = host(true);
    expect(await wpsAnalysisTargetEvidence(wrapped.target)).toEqual(expect.objectContaining({ conclusive: false }));
    expect(wrapped.values).not.toHaveBeenCalled();
  });
  it("evidence from another target is rejected before reading headers/rows", async () => {
    const h = host();
    expect(await wpsAnalysisTargetEvidence({ ...h.target, fileKey: "E:\\Other\\Book.xlsx" })).toBeNull();
    expect(await wpsAnalysisTargetEvidence({ ...h.target, sheets: [{ id: "1", name: "Replacement" }] })).toBeNull();
    expect(h.values).not.toHaveBeenCalled();
  });
});
describe("owned review copy metadata role, not a new original", () => {
  const target: AuthorizedTarget = { fileKey: null, name: "Book1", sheets: [], selected: [],
    fileRef: { Name: "Book1", FullName: "Book1", Path: "" } };
  const review = { copy: target.fileRef, copyIdentity: { name: "Book1", fullName: "Book1", path: null } } as WpsNativeComparison;
  it("recognizes the unsaved owned copy when FullName is a name, not a disk identity", () => {
    expect(isWpsReviewCopyTarget(target, review)).toBe(true);
    expect(isWpsReviewCopyTarget({ ...target, fileRef: { ...target.fileRef } }, review)).toBe(true);
  });
  it("does not exempt unrelated files or an absent review", () => {
    expect(isWpsReviewCopyTarget({ ...target, name: "Other" }, review)).toBe(false);
    expect(isWpsReviewCopyTarget({ ...target, fileRef: { Name: "Book1", FullName: "E:\\Other\\Book1.xlsx", Path: "E:\\Other" } }, review)).toBe(false);
    expect(isWpsReviewCopyTarget(target, null)).toBe(false);
  });
});
