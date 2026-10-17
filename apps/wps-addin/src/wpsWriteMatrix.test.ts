import { describe, expect, it, vi } from "vitest";
import { writeWpsMatrix } from "./wpsWriteMatrix";
import type { WpsRange } from "./wpsHost";

function host(formats: string[][], failWrite = false) {
  const original = formats.map((row) => [...row]);
  let stored: unknown;
  const range: WpsRange = {
    Item: (r, c = 1) => ({
      get NumberFormat() { return formats[r - 1][c - 1]; },
      set NumberFormat(value) { formats[r - 1][c - 1] = String(value); }
    })
  };
  const setter = (values: unknown) => {
    if (failWrite) throw new Error("write rejected");
    stored = (values as unknown[][]).map((row, r) => row.map((value, c) =>
      typeof value === "string" && /^\d+$/.test(value) && formats[r][c] !== "@" ? Number(value) : value));
  };
  Object.defineProperty(range, "Value2", { set: setter });
  Object.defineProperty(range, "Formula", { set: setter });
  return { range, original, formats, stored: () => stored };
}

describe("WPS text-preserving matrix writes", () => {
  it("keeps numeric text, leading zeros and long IDs while retaining typed numbers and formats", () => {
    const target = host([["General", "0.00", "General", "0%", "@"]]);
    const values = [["7", "007", "12345678901234567890", 400, true]];
    writeWpsMatrix(target.range, values);
    expect(target.stored()).toEqual(values);
    expect(target.formats).toEqual(target.original);
  });
  it("restores formats even when writing throws", () => {
    const target = host([["General", "0.00"]], true);
    expect(() => writeWpsMatrix(target.range, [["007", 400]])).toThrow("write rejected");
    expect(target.formats).toEqual(target.original);
  });
  it("protects literal text in an undo matrix without putting formulas in text format", () => {
    const target = host([["General", "General", "0.00"]]);
    const values = [["007", "=1+1", 400]];
    writeWpsMatrix(target.range, values, "Formula");
    expect(target.stored()).toEqual(values);
    expect(target.formats).toEqual(target.original);
  });
  it("uses scalar assignments for a single cell and restores its format", () => {
    let value: unknown;
    const range: WpsRange = { NumberFormat: "General" };
    Object.defineProperty(range, "Value2", { set: (incoming) => {
      value = range.NumberFormat === "@" ? incoming : Number(incoming);
    } });
    writeWpsMatrix(range, [["007"]]);
    expect(value).toBe("007");
    expect(range.NumberFormat).toBe("General");
  });
  it("does not start writing if setting text format fails, and restores earlier changes", () => {
    const formats = ["0.00", "General"];
    const write = vi.fn();
    const range: WpsRange = { Item: (_r, c = 1) => ({
      get NumberFormat() { return formats[c - 1]; },
      set NumberFormat(value) {
        if (c === 2 && value === "@") throw new Error("format rejected");
        formats[c - 1] = String(value);
      }
    }) };
    Object.defineProperty(range, "Value2", { set: write });
    expect(() => writeWpsMatrix(range, [["007", "9"]])).toThrow("format rejected");
    expect(write).not.toHaveBeenCalled();
    expect(formats).toEqual(["0.00", "General"]);
  });
  it("reports a silent format-restore failure but still restores the other cells", () => {
    const formats = ["0.00", "General"];
    const write = vi.fn();
    const range: WpsRange = { Item: (_r, c = 1) => ({
      get NumberFormat() { return formats[c - 1]; },
      set NumberFormat(value) {
        if (c === 2 && value === "General") return;
        formats[c - 1] = String(value);
      }
    }) };
    Object.defineProperty(range, "Value2", { set: write });
    expect(() => writeWpsMatrix(range, [["007", "9"]])).toThrow("不能确认写入成功");
    expect(write).toHaveBeenCalledTimes(1);
    expect(formats).toEqual(["0.00", "@"]);
  });
  it("does not change formula-cell or numeric-cell formats even during assignment", () => {
    const formats = ["General", "0.00", "0%"];
    const observed: string[][] = [];
    const range: WpsRange = { Item: (_r, c = 1) => ({
      get NumberFormat() { return formats[c - 1]; },
      set NumberFormat(value) { formats[c - 1] = String(value); }
    }) };
    Object.defineProperty(range, "Formula", { set: () => observed.push([...formats]) });
    writeWpsMatrix(range, [["007", "=1+1", 400]], "Formula");
    expect(observed).toEqual([["@", "0.00", "0%"]]);
    expect(formats).toEqual(["General", "0.00", "0%"]);
  });
});
