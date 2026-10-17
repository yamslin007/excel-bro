import { describe, expect, it, vi } from "vitest";
import { readWpsRangeValues } from "./wpsRangeValues";

describe("WPS complete value capture", () => {
  it.each([null, "", 0, false, "ABC"])("captures explicit single-cell value %s", (Value2) => {
    expect(readWpsRangeValues({ Value2 }, 1, 1)).toEqual({ values: [[Value2]], valuesCaptured: true });
  });
  it("reads Value2 only once for a complete matrix", () => {
    const getter = vi.fn(() => [[null], [2]]);
    const result = readWpsRangeValues({ get Value2() { return getter(); } }, 2, 1);
    expect(result).toEqual({ values: [[null], [2]], valuesCaptured: true });
    expect(getter).toHaveBeenCalledTimes(1);
  });
  it.each([undefined, null, [[null]], [[undefined], [2]], [null], new Array(2)])(
    "does not treat missing or incomplete multi-cell values as captured: %s", (Value2) => {
      expect(readWpsRangeValues({ Value2 }, 2, 1).valuesCaptured).toBe(false);
    });
  it("can resolve a multi-cell scalar response through actual cell reads", () => {
    const Item = vi.fn((row: number, column?: number) => ({ Value2: row === 1 ? null : 2 }));
    expect(readWpsRangeValues({ Value2: null, Item }, 2, 1))
      .toEqual({ values: [[null], [2]], valuesCaptured: true });
    expect(Item.mock.calls).toEqual([[1, 1], [2, 1]]);
  });
  it("propagates host read errors without inventing empty cells", () => {
    expect(() => readWpsRangeValues({ get Value2() { throw new Error("read failed"); } }, 1, 1))
      .toThrow("read failed");
  });
});
