// @vitest-environment jsdom
import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { ScopeSelectionGuidance } from "./ScopeSelectionGuidance";

afterEach(() => cleanup());

describe("ScopeSelectionGuidance", () => {
  it("uses the same worksheet icon to point to the scope control above the composer", () => {
    render(
      <ScopeSelectionGuidance
        keeperName="sheet-dc1"
        missingSheetNames={[
          "恐怖游轮",
          "痴迷",
          "功夫女足",
          "群星闪耀时",
          "八仙！",
          "想你了",
          "四渡"
        ]}
      />
    );

    expect(screen.getByText("工作表范围")).toBeTruthy();
    expect(screen.getByText("▦").getAttribute("aria-hidden")).toBe("true");
    expect(screen.getByText(/请点击输入框上方的/)).toBeTruthy();
    expect(screen.getByText(/请同时勾选需要保留的/).textContent).toContain(
      "sheet-dc1"
    );
    expect(screen.getByText(/等 7 张工作表/)).toBeTruthy();
    expect(screen.queryByText(/点击下方/)).toBeNull();
  });
});
