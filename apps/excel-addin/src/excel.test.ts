import { describe, expect, it } from "vitest";
import { valuesEqual } from "./excel";

describe("Excel verification value equivalence", () => {
  it("accepts harmless numeric text coercion", () => {
    expect(valuesEqual(43122301, "43122301")).toBe(true);
    expect(valuesEqual("43122301", 43122301)).toBe(true);
  });

  it("rejects numeric coercion that loses identifier information", () => {
    expect(valuesEqual(1105882025, "001105882025")).toBe(false);
  });

  it("accepts negligible floating point differences", () => {
    expect(valuesEqual(0.1 + 0.2, 0.3)).toBe(true);
  });
});
