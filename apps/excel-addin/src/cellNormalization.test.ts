import { describe, expect, it } from "vitest";
import { displayCellValue, normalizeCellValue } from "./cellNormalization";

describe("cell normalization", () => {
  it("converts Excel date serials to stable ISO dates", () => {
    expect(normalizeCellValue(46_572, "2027/7/4", "yyyy/m/d")).toBe(
      "2027-07-04"
    );
  });

  it("preserves formatted identifiers with leading zeroes", () => {
    expect(normalizeCellValue(1, "001", "000")).toBe("001");
    expect(normalizeCellValue("001", "001", "@")).toBe("001");
  });

  it("keeps calculation values separate from display text", () => {
    expect(normalizeCellValue(0.25, "25%", "0%")).toBe(0.25);
    expect(displayCellValue(0.25, "25%")).toBe("25%");
  });
});
