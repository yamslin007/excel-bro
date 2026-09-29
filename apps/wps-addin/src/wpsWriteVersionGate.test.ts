import { describe, expect, it } from "vitest";
import { wpsWriteVersionStatus } from "./wpsWriteVersionGate";

describe("WPS write version gate", () => {
  it("allows the observed JSAPI 12.0 only in the development add-in", () => {
    expect(wpsWriteVersionStatus("12.0", true)).toBe("development");
    expect(wpsWriteVersionStatus("12.0", false)).toBe("blocked");
  });

  it("does not mistake the installation directory's 12.1 product version for a verified JSAPI value", () => {
    expect(wpsWriteVersionStatus("12.1.0.28488", true)).toBe("blocked");
    expect(wpsWriteVersionStatus("12.1.0.28488", false)).toBe("blocked");
  });

  it("rejects missing and unrelated versions in both modes", () => {
    for (const version of [null, "", "11.0", "12.01", "13.0"]) {
      expect(wpsWriteVersionStatus(version, true)).toBe("blocked");
      expect(wpsWriteVersionStatus(version, false)).toBe("blocked");
    }
  });
});
