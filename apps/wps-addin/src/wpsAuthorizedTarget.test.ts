import { describe, expect, it, vi, afterEach } from "vitest";
import { isWpsReviewCopyTarget, wpsAnalysisTargetEvidence } from "./wpsAuthorizedTarget";
import { bindWpsReviewTarget, releaseWpsReviewTarget, wpsReadWorkbook } from "./wpsResultReviewTarget";
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
describe("owned review copy is not a new original", () => {
  /*
   * 绑定的是**本次创建时**记下的身份；判据不用名字前缀/活动窗口/相似度。
   * 副本被 WPS 激活时读取仍按原表，无关工作簿仍然是无关工作簿。
   */
  const sheet = (name: string) => ({ Name: name, Index: 1, CodeName: `id-${name}`, Range: () => ({}),
    UsedRange: { Row: 1, Column: 1, Rows: { Count: 0 }, Columns: { Count: 0 }, Value2: [] } });
  const book = (name: string, fullName: string, path: string) => ({
    Name: name, FullName: fullName, Path: path,
    Worksheets: { Count: 1, Item: () => sheet("Sheet1") }
  });
  const original = book("ss.xlsx", "E:////Work////ss.xlsx", "E:////Work");
  const copy = book("工作簿1", "工作簿1", "");
  const unrelated = book("other.xlsx", "E:////Work////other.xlsx", "E:////Work");

  function bind(active: unknown, books: unknown[] = [original, copy]) {
    bindWpsReviewTarget({
      original: original as never,
      originalIdentity: { name: "ss.xlsx", fullName: "E:////Work////ss.xlsx", path: "E:////Work" },
      copyIdentity: { name: "工作簿1", fullName: "工作簿1", path: null },
      writeSheet: { index: 1, name: "Sheet1" }
    });
    vi.mocked(getWpsApplication).mockReturnValue({
      get ActiveWorkbook() { return active; },
      Workbooks: { Count: books.length, Item: (index: number) => books[index - 1] }
    } as never);
  }

  afterEach(() => releaseWpsReviewTarget());

  it("副本获得焦点时按原表读取（不是新目标）", () => {
    bind(copy);
    expect(isWpsReviewCopyTarget({ fileKey: null, name: "工作簿1", sheets: [], selected: [], fileRef: copy as never })).toBe(true);
    expect(wpsReadWorkbook(getWpsApplication())).toBe(original);
  });

  it("无关工作簿不被当成副本；原表关闭后不再解析（守卫会拦）", () => {
    bind(unrelated);
    expect(wpsReadWorkbook(getWpsApplication())).toBe(unrelated);
    expect(isWpsReviewCopyTarget({ fileKey: null, name: "other.xlsx", sheets: [], selected: [], fileRef: unrelated as never })).toBe(false);

    // 只剩副本（原表被关）⇒ 证不出"原表仍在"，不解析，交给守卫拦截。
    bind(copy, [copy]);
    expect(wpsReadWorkbook(getWpsApplication())).toBe(copy);
  });

  it("副本被另存成磁盘文件后不再是本次副本", () => {
    const saved = book("工作簿1", "E:////Work////工作簿1.xlsx", "E:////Work");
    bind(saved);
    expect(wpsReadWorkbook(getWpsApplication())).toBe(saved);
  });
});
