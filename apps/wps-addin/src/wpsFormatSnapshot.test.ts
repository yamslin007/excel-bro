import { describe, expect, it } from "vitest";
import {
  captureWpsFormatSnapshot,
  restoreWpsFormatSnapshot,
  sameWpsFormatSnapshot,
  type WpsFormatSnapshot
} from "./wpsFormatSnapshot";
import type { WpsRange } from "./wpsHost";

/**
 * 带**逐格**格式状态的替身区域：`Item(r, c)` 返回各自独立的格式对象，
 * 模拟"混合格式"（区域内不同格有不同格式），这是单格探测覆盖不到的场景。
 */
function formatRange(rows: number, columns: number): WpsRange & { state: () => any } {
  const cells: Array<Array<{
    bold: boolean; italic: boolean; name: string; size: number; color: number;
    horizontalAlignment: number; wrapText: boolean; rowHeight: number;
  }>> = [];
  for (let r = 0; r < rows; r += 1) {
    const rowCells: any[] = [];
    for (let c = 0; c < columns; c += 1) {
      rowCells.push({
        bold: false, italic: false, name: "宋体", size: 11, color: 0,
        horizontalAlignment: 1, wrapText: false, rowHeight: 13.5
      });
    }
    cells.push(rowCells);
  }

  const cellCache = new Map<string, WpsRange>();
  const makeCell = (r: number, c: number): WpsRange => {
    const key = `${r},${c}`;
    const cached = cellCache.get(key);
    if (cached) return cached;
    const state = cells[r][c];
    const cell: WpsRange = {
      Font: {
        get Bold() { return state.bold; }, set Bold(v: unknown) { state.bold = v as boolean; },
        get Italic() { return state.italic; }, set Italic(v: unknown) { state.italic = v as boolean; },
        get Name() { return state.name; }, set Name(v: unknown) { state.name = v as string; },
        get Size() { return state.size; }, set Size(v: unknown) { state.size = v as number; },
        get Color() { return state.color; }, set Color(v: unknown) { state.color = v as number; }
      },
      get HorizontalAlignment() { return state.horizontalAlignment; },
      set HorizontalAlignment(v: unknown) { state.horizontalAlignment = v as number; },
      get WrapText() { return state.wrapText; }, set WrapText(v: unknown) { state.wrapText = v as boolean; },
      get RowHeight() { return state.rowHeight; }, set RowHeight(v: unknown) { state.rowHeight = v as number; }
    };
    cellCache.set(key, cell);
    return cell;
  };

  const range: WpsRange & { state: () => any } = {
    Rows: { Count: rows },
    Columns: { Count: columns },
    Item: (r: number, c: number) => makeCell(r - 1, c - 1),
    // 单格区域（Range("A1")）在真实 WPS 里**本身就是单元格**，`Font`/对齐/换行/
    // 行高直接挂在 range 上（`captureWpsFormatSnapshot` 单格路径读 `range` 本身，
    // 不走 `Item`）。替身把这些 getter/setter 委托到第一个格，模拟真实单格 range。
    Font: {
      get Bold() { return cells[0][0].bold; }, set Bold(v: unknown) { cells[0][0].bold = v as boolean; },
      get Italic() { return cells[0][0].italic; }, set Italic(v: unknown) { cells[0][0].italic = v as boolean; },
      get Name() { return cells[0][0].name; }, set Name(v: unknown) { cells[0][0].name = v as string; },
      get Size() { return cells[0][0].size; }, set Size(v: unknown) { cells[0][0].size = v as number; },
      get Color() { return cells[0][0].color; }, set Color(v: unknown) { cells[0][0].color = v as number; }
    },
    get HorizontalAlignment() { return cells[0][0].horizontalAlignment; },
    set HorizontalAlignment(v: unknown) { cells[0][0].horizontalAlignment = v as number; },
    get WrapText() { return cells[0][0].wrapText; }, set WrapText(v: unknown) { cells[0][0].wrapText = v as boolean; },
    get RowHeight() { return cells[0][0].rowHeight; }, set RowHeight(v: unknown) { cells[0][0].rowHeight = v as number; },
    state: () => cells
  };

  return range;
}

describe("WPS 格式快照（逐格/逐行）", () => {
  it("逐格采集：混合格式（有的格加粗有的不加粗）各格独立保留", () => {
    const range = formatRange(2, 2);
    // 第 1 行第 1 格加粗、第 2 行第 2 格斜体，制造混合格式。
    (range.Item!(1, 1).Font!).Bold = true;
    (range.Item!(2, 2).Font!).Italic = true;
    const snapshot = captureWpsFormatSnapshot(range, 2, 2);
    expect(snapshot.captured).toBe(true);
    expect(snapshot.cells[0][0].bold).toBe(true);
    expect(snapshot.cells[0][1].bold).toBe(false);
    expect(snapshot.cells[1][1].italic).toBe(true);
  });

  it("恢复后逐格回到写前（含混合格式）", () => {
    const range = formatRange(2, 2);
    (range.Item!(1, 1).Font!).Bold = true;
    const snapshot = captureWpsFormatSnapshot(range, 2, 2);
    // 改动格式。
    (range.Item!(1, 1).Font!).Bold = false;
    (range.Item!(1, 2).Font!).Name = "黑体";
    restoreWpsFormatSnapshot(range, snapshot);
    const after = captureWpsFormatSnapshot(range, 2, 2);
    expect(sameWpsFormatSnapshot(after, snapshot)).toBe(true);
    expect(after.cells[0][0].bold).toBe(true);
    expect(after.cells[0][1].name).toBe("宋体");
  });

  it("任一格任一目标属性读不到 → captured=false（不可用于撤销）", () => {
    const range = formatRange(1, 2);
    // 第 2 格不暴露 Font（混合情况下读不到字体）。
    const cell = range.Item!(1, 2);
    Object.defineProperty(cell, "Font", { get: () => undefined });
    const snapshot = captureWpsFormatSnapshot(range, 1, 2);
    expect(snapshot.captured).toBe(false);
  });

  it("逐行行高采集与恢复", () => {
    const range = formatRange(2, 1);
    (range.Item!(2, 1)).RowHeight = 25;
    const snapshot = captureWpsFormatSnapshot(range, 2, 1);
    expect(snapshot.rowHeights).toEqual([13.5, 25]);
    (range.Item!(1, 1)).RowHeight = 40;
    (range.Item!(2, 1)).RowHeight = 40;
    restoreWpsFormatSnapshot(range, snapshot);
    expect(range.Item!(1, 1).RowHeight).toBe(13.5);
    expect(range.Item!(2, 1).RowHeight).toBe(25);
  });

  it("sameWpsFormatSnapshot 逐格逐行严格比对", () => {
    const range = formatRange(1, 1);
    const a = captureWpsFormatSnapshot(range, 1, 1);
    const b = captureWpsFormatSnapshot(range, 1, 1);
    expect(sameWpsFormatSnapshot(a, b)).toBe(true);
    (range.Item!(1, 1).Font!).Bold = true;
    const c = captureWpsFormatSnapshot(range, 1, 1);
    expect(sameWpsFormatSnapshot(a, c)).toBe(false);
  });
});
