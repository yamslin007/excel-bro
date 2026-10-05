import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

describe("网页前后对比已退役，原生审阅与证据保留", () => {
  it("Excel 不再打开或构建网页对照窗口", () => {
    const source = readFileSync("src/App.tsx", "utf8");
    expect(source.includes("openReviewWindow")).toBe(false);
    expect(source.includes("reviewDialogRef")).toBe(false);
    expect(source.includes("./reviewState")).toBe(false);
    expect(readFileSync("vite.config.ts", "utf8")).not.toContain('"review.html"');
    for (const path of ["review.html", "src/ReviewWorkspace.tsx", "src/review.tsx", "src/reviewState.ts", "src/review.css", "src/components/ChangeReviewTable.tsx"]) {
      expect(existsSync(path), path).toBe(false);
    }
  });

  it("WPS 不再切换网页双网格，仍走独立副本原生审阅", () => {
    const source = readFileSync("../wps-addin/src/WpsAgentPanel.tsx", "utf8");
    expect(source.includes("ReviewWorkspace")).toBe(false);
    expect(source.includes("activeReviewExecutionId")).toBe(false);
    expect(source.includes("openChangeReview")).toBe(false);
    expect(readFileSync("../wps-addin/src/main.tsx", "utf8")).not.toContain("review.css");
    for (const entry of ["stageWpsNativeComparison", "verifyWpsNativeComparison", "discardWpsNativeComparison", "compareCommittedWpsResultToCopy"]) {
      expect(source).toContain(entry);
    }
  });

  it("两端继续保留执行回读和变更摘要，不以删除 UI 代替核对", () => {
    for (const path of ["src/App.tsx", "../wps-addin/src/WpsAgentPanel.tsx"]) {
      const source = readFileSync(path, "utf8");
      expect(source.includes("changeReviewsRef.current.set")).toBe(true);
      expect(source.includes("ChangeReviewSummaryCard")).toBe(true);
    }
    expect(readFileSync("../wps-addin/src/WpsAgentPanel.tsx", "utf8")).toContain("ChangeReviewUndoNote");
  });
});
