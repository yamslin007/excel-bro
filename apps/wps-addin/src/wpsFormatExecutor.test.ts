import { describe, expect, it } from "vitest";
import type { AnalysisPlan, CellValue } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import { captureWpsFormatSnapshot } from "./wpsFormatSnapshot";
import { wpsRecoveryDescriptor } from "./wpsExecutionLedger";
import {
  asFormatFields,
  composedFormatSnapshot,
  executeWpsPlan,
  formatEvidenceFingerprint,
  formatExpectedFingerprint,
  matrixFingerprintV2,
  prepareWpsPlan,
  undoWpsExecution,
  type WpsPreparedPlan
} from "./wpsPlanExecutor";

/**
 * 带**逐格**格式状态的替身宿主（Task #53 格式动作闭环）。
 *
 * 关键：格式必须逐格独立（`Item(r,c)` 返回各自的 Font/对齐/换行/行高），
 * 以覆盖"混合格式"——单格探测（Sheet1!A1）证明不了混合格式恢复。
 */
function formatHost(rows: number, columns: number) {
  const grid: Array<Array<{
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
    grid.push(rowCells);
  }

  const cellCache = new Map<string, WpsRange>();
  const makeCell = (r: number, c: number): WpsRange => {
    const key = `${r},${c}`;
    const cached = cellCache.get(key);
    if (cached) return cached;
    const state = grid[r][c];
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
      get RowHeight() { return state.rowHeight; }, set RowHeight(v: unknown) { state.rowHeight = v as number; },
      Rows: { Count: 1 },
      Columns: { Count: 1 }
    };
    cellCache.set(key, cell);
    return cell;
  };

  const range: WpsRange = {
    MergeCells: false,
    Rows: { Count: rows },
    Columns: { Count: columns },
    Item: (r: number, c: number) => makeCell(r - 1, c - 1),
    // 单格区域（Range("A1")）在真实 WPS 里**本身就是单元格**，`Font`/对齐/换行/
    // 行高直接挂在 range 上（`captureWpsFormatSnapshot` 单格路径读 `range` 本身，
    // 不走 `Item`）。替身把这些 getter/setter 委托到第一个格，模拟真实单格 range。
    Font: {
      get Bold() { return grid[0][0].bold; }, set Bold(v: unknown) { grid[0][0].bold = v as boolean; },
      get Italic() { return grid[0][0].italic; }, set Italic(v: unknown) { grid[0][0].italic = v as boolean; },
      get Name() { return grid[0][0].name; }, set Name(v: unknown) { grid[0][0].name = v as string; },
      get Size() { return grid[0][0].size; }, set Size(v: unknown) { grid[0][0].size = v as number; },
      get Color() { return grid[0][0].color; }, set Color(v: unknown) { grid[0][0].color = v as number; }
    },
    get HorizontalAlignment() { return grid[0][0].horizontalAlignment; },
    set HorizontalAlignment(v: unknown) { grid[0][0].horizontalAlignment = v as number; },
    get WrapText() { return grid[0][0].wrapText; }, set WrapText(v: unknown) { grid[0][0].wrapText = v as boolean; },
    get RowHeight() { return grid[0][0].rowHeight; }, set RowHeight(v: unknown) { grid[0][0].rowHeight = v as number; },
    // 格式动作的 `prepareFormatAction` 现在要读值矩阵作为"值不变"基线
    // （副本审阅 `previewFingerprint` 用真实值指纹，而非空矩阵指纹）。
    // 替身给全 null 矩阵即可 —— 格式测试不关心值内容，只关心 Formula 可读。
    Value2: Array.from({ length: rows }, () => Array.from({ length: columns }, () => null)),
    Formula: Array.from({ length: rows }, () => Array.from({ length: columns }, () => null))
  };
  const sheet: WpsWorksheet = {
    Name: "明细",
    ProtectContents: false,
    Range: () => range
  };
  const workbook = {
    Name: "格式测试.xlsx",
    FullName: "C:\\Data\\格式测试.xlsx",
    Path: "C:\\Data",
    ReadOnly: false,
    Worksheets: { Count: 1, Item: () => sheet },
    ActiveSheet: sheet
  };
  const storage = new Map<string, unknown>();
  const application: WpsApplication = {
    Version: "12.0",
    ActiveWorkbook: workbook,
    ActiveSheet: sheet,
    PluginStorage: {
      getItem: (key) => storage.get(key),
      setItem: (key, value) => { storage.set(key, value); }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  };
  return {
    application,
    grid,
    cell: (r: number, c: number) => grid[r][c],
    storedRecords: () => JSON.parse(String(storage.get("excel-bro.wps.undo.v1") ?? "[]"))
  };
}

function plan(actions: AnalysisPlan["actions"]): AnalysisPlan {
  return {
    id: "plan-fmt-1",
    title: "设置格式",
    summary: "设置字体与对齐",
    assumptions: [],
    warnings: [],
    actions,
    acceptanceCriteria: []
  };
}

describe("WPS 格式动作闭环（预览→执行→回读→撤销）", () => {
  it("setFont(bold) 完整闭环：执行后回读加粗，撤销后回到不加粗", async () => {
    const target = formatHost(2, 2);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "setFont", sheet: "明细", range: "A1:B2", bold: true }
    ]));
    // 预览后未执行：仍是不加粗。
    expect(target.cell(0, 0).bold).toBe(false);
    const record = await executeWpsPlan(target.application, prepared);
    // 执行后：所有格加粗。
    expect(target.cell(0, 0).bold).toBe(true);
    expect(target.cell(1, 1).bold).toBe(true);
    // 撤销后：回到不加粗。
    await undoWpsExecution(target.application, record.executionId);
    expect(target.cell(0, 0).bold).toBe(false);
    expect(target.cell(1, 1).bold).toBe(false);
  });

  it("setFont(name/size/color) 完整闭环：逐格改到目标值，撤销恢复", async () => {
    const target = formatHost(2, 1);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "setFont", sheet: "明细", range: "A1:A2", name: "黑体", size: 14, color: "#FF0000" }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    expect(target.cell(0, 0).name).toBe("黑体");
    expect(target.cell(0, 0).size).toBe(14);
    expect(target.cell(0, 0).color).toBe(0x0000ff);
    expect(target.cell(1, 0).name).toBe("黑体");
    await undoWpsExecution(target.application, record.executionId);
    expect(target.cell(0, 0).name).toBe("宋体");
    expect(target.cell(0, 0).size).toBe(11);
    expect(target.cell(0, 0).color).toBe(0);
  });

  it("setAlignment(horizontal/wrapText) 完整闭环", async () => {
    const target = formatHost(2, 1);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "setAlignment", sheet: "明细", range: "A1:A2", horizontal: "center", wrapText: true }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    expect(target.cell(0, 0).horizontalAlignment).toBe(-4108); // xlCenter
    expect(target.cell(0, 0).wrapText).toBe(true);
    await undoWpsExecution(target.application, record.executionId);
    expect(target.cell(0, 0).horizontalAlignment).toBe(1); // xlGeneral
    expect(target.cell(0, 0).wrapText).toBe(false);
  });

  it("resizeRange(rowHeight) 完整闭环：逐行改行高，撤销恢复", async () => {
    const target = formatHost(2, 1);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "resizeRange", sheet: "明细", range: "A1:A2", rowHeight: 24 }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    expect(target.cell(0, 0).rowHeight).toBe(24);
    expect(target.cell(1, 0).rowHeight).toBe(24);
    await undoWpsExecution(target.application, record.executionId);
    expect(target.cell(0, 0).rowHeight).toBe(13.5);
    expect(target.cell(1, 0).rowHeight).toBe(13.5);
  });

  it("混合格式：只改加粗，斜体保持各格原状，撤销后混合格式完整恢复", async () => {
    const target = formatHost(2, 1);
    // 预置混合格式：第 1 格斜体、第 2 格不加斜体。
    target.cell(0, 0).italic = true;
    target.cell(1, 0).italic = false;
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "setFont", sheet: "明细", range: "A1:A2", bold: true }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    // 执行后：两格都加粗，但斜体保持各格原状。
    expect(target.cell(0, 0).bold).toBe(true);
    expect(target.cell(1, 0).bold).toBe(true);
    expect(target.cell(0, 0).italic).toBe(true);
    expect(target.cell(1, 0).italic).toBe(false);
    await undoWpsExecution(target.application, record.executionId);
    // 撤销后：混合斜体完整恢复。
    expect(target.cell(0, 0).italic).toBe(true);
    expect(target.cell(1, 0).italic).toBe(false);
    expect(target.cell(0, 0).bold).toBe(false);
  });

  it("setAlignment 垂直属性读取不到时零写入拒绝", () => {
    const target = formatHost(1, 1);
    expect(() => prepareWpsPlan(target.application, [1], plan([
      { type: "setAlignment", sheet: "明细", range: "A1", horizontal: "center", vertical: "center" }
    ]))) .toThrow(/verticalAlignment/);
    // 拒绝后不动工作簿。
    expect(target.cell(0, 0).horizontalAlignment).toBe(1);
  });

  it("resizeRange 带 columnWidth 明确拒绝（未验证属性不静默忽略）", () => {
    const target = formatHost(1, 1);
    expect(() => prepareWpsPlan(target.application, [1], plan([
      { type: "resizeRange", sheet: "明细", range: "A1", rowHeight: 24, columnWidth: 12 }
    ]))).toThrow(/columnWidth/i);
    expect(target.cell(0, 0).rowHeight).toBe(13.5);
  });

  it("缺少填充或边框快照时拒绝；未实现的自适应仍拒绝", () => {
    const target = formatHost(1, 1);
    expect(() => prepareWpsPlan(target.application, [1], plan([
      { type: "setFill", sheet: "明细", range: "A1", color: "#DDEEFF" }
    ]))).toThrow("无法完整采集");
    expect(() => prepareWpsPlan(target.application, [1], plan([
      { type: "setBorders", sheet: "明细", range: "A1", sides: ["top"] }
    ]))).toThrow("borderTop");
    expect(() => prepareWpsPlan(target.application, [1], plan([
      { type: "autofit", sheet: "明细", range: "A1" }
    ]))).toThrow("当前不执行");
  });

  it("格式无法逐格采集（宿主未返回字体）→ prepareWpsPlan 拒绝，不产生不可撤销修改", () => {
    const target = formatHost(1, 1);
    // 让单格 range 不暴露 Font，模拟宿主读不到字体（单格路径读 range 本身）。
    const range = target.application.ActiveWorkbook!.Worksheets.Item(1).Range!("A1");
    Object.defineProperty(range, "Font", { get: () => undefined });
    expect(() => prepareWpsPlan(target.application, [1], plan([
      { type: "setFont", sheet: "明细", range: "A1", bold: true }
    ]))).toThrow(/无法完整逐格采集/);
  });

  it("执行后被人改过再撤销 → 拒绝覆盖（格式冲突守卫）", async () => {
    const target = formatHost(1, 1);
    const prepared = prepareWpsPlan(target.application, [1], plan([
      { type: "setFont", sheet: "明细", range: "A1", bold: true }
    ]));
    const record = await executeWpsPlan(target.application, prepared);
    expect(target.cell(0, 0).bold).toBe(true);
    // 执行后用户手动把加粗又改掉（改成斜体+取消加粗）。
    target.cell(0, 0).bold = false;
    target.cell(0, 0).italic = true;
    await expect(undoWpsExecution(target.application, record.executionId)).rejects.toThrow("又被修改");
    // 保留用户修改，不覆盖。
    expect(target.cell(0, 0).bold).toBe(false);
    expect(target.cell(0, 0).italic).toBe(true);
  });
});

/* ------------------------------------------------------------------------- *
 * 格式证据指纹与**合成写前现场**（Task #62）
 * ------------------------------------------------------------------------- *
 *
 * 真实 WPS 验收失败的直接原因是"格式动作根本没有可用的现场证据"：值矩阵指纹对格式
 * 改动没有区分力，所以格式动作必须走 `f1:` 格式证据指纹。下面这一组钉住三件事：
 * 1. 三种现场（写前 / 写后 / 外部修改）在格式指纹上是三个不同的值，而值矩阵指纹
 *    在同样三种现场上**完全一样**（这就是它没有区分力的证据）；
 * 2. 登记时算出的 `expectedFingerprint` 与真实回读现场的指纹**同一个值**（两边共用
 *    一套归一化，不是"回读通过但指纹对不上"）；
 * 3. 同区域按序叠加时，账本登记的写前指纹必须**逐字等于**执行器真正发出去的那一份
 *    —— 否则服务端会以"证据与登记的写前状态不一致"拒掉写前检查点，恢复链断在登记上。
 */

function rangeOf(prepared: WpsPreparedPlan, address: string, host: { application: WpsApplication }) {
  return host.application.ActiveWorkbook!.Worksheets.Item(1).Range!(address);
}

describe("格式证据指纹与合成写前现场（Task #62）", () => {
  it("写前 / 写后 / 外部修改：格式指纹三者互不相同，值矩阵指纹三者完全一样", async () => {
    const host = formatHost(3, 1);
    const prepared = prepareWpsPlan(host.application, [1], plan([
      { type: "setFont", sheet: "明细", range: "A1:A3", size: 20 }
    ]));
    const action = prepared.actions[0];
    const formatAction = asFormatFields(action);
    const range = rangeOf(prepared, "A1:A3", host);
    const read = () => {
      const snapshot = captureWpsFormatSnapshot(range, 3, 1);
      expect(snapshot.captured).toBe(true);
      return {
        format: formatEvidenceFingerprint(formatAction, snapshot),
        // 值矩阵指纹：格式动作不改值，三种现场上都一样。
        value: matrixFingerprintV2(
          snapshot.cells.map((row) => row.map(() => null)),
          snapshot.cells.map((row) => row.map(() => null))
        )
      };
    };

    const before = read();
    await executeWpsPlan(host.application, prepared);
    const after = read();
    host.cell(0, 0).size = 99; // 外部修改：值没动，格式动了
    const external = read();

    // 登记的目标指纹 == 真实回读现场的指纹（两边同一套归一化）。
    expect(after.format).toBe(formatExpectedFingerprint(formatAction, 3, 1));
    // 写前 / 写后 / 外部修改 三者互不相同 —— 恢复判定才可能区分它们。
    expect(new Set([before.format, after.format, external.format]).size).toBe(3);
    for (const fingerprint of [before.format, after.format, external.format]) {
      expect(fingerprint).toMatch(/^f1:/);
    }
    // 反证：值矩阵指纹在这三种现场上一模一样，对格式改动没有任何区分力。
    expect(before.value).toBe(after.value);
    expect(after.value).toBe(external.value);
  });

  it("同区域同属性按序叠加（14 → 18）：登记的写前/写后指纹逐字等于执行器发出的检查点", async () => {
    const host = formatHost(3, 1);
    const prepared = prepareWpsPlan(host.application, [1], plan([
      { type: "setFont", sheet: "明细", range: "A1:A3", size: 14 },
      { type: "setFont", sheet: "明细", range: "A1:A3", size: 18 }
    ]));
    const descriptor = wpsRecoveryDescriptor(prepared);
    const checkpoints: Array<{ phase: string; actionIndex: number; before: string; after: string | null }> = [];
    await executeWpsPlan(host.application, prepared, undefined, undefined, async (phase, checkpoint) => {
      checkpoints.push({
        phase,
        actionIndex: checkpoint.actionIndex,
        before: checkpoint.beforeFingerprint,
        after: checkpoint.afterFingerprint
      });
    });

    for (const [index, action] of descriptor.actions.entries()) {
      expect(action.type).toBe("setFont");
      expect(action.formatFields).toEqual(["size"]);
      expect(action.previewFingerprint).toMatch(/^f1:/);
      expect(action.expectedFingerprint).toMatch(/^f1:/);
      // 写前指纹是**合成**后的现场：第二个动作执行前，区域已经是 14 了。
      const before = checkpoints.find((item) => item.phase === "before" && item.actionIndex === index);
      const after = checkpoints.find((item) => item.phase === "after" && item.actionIndex === index);
      expect(before?.before).toBe(action.previewFingerprint);
      expect(after?.after).toBe(action.expectedFingerprint);
      expect(action.previewFingerprint).not.toBe(action.expectedFingerprint);
    }
    // 两个动作的写前现场本来就不同（第二个动作是在第一个动作之后）—— 不是同一个值。
    expect(descriptor.actions[0].previewFingerprint).not.toBe(
      descriptor.actions[1].previewFingerprint
    );
    expect(host.cell(0, 0).size).toBe(18);
  });

  it("合成写前现场只叠加前序动作**声明过**的属性与**相交**的范围", () => {
    const host = formatHost(3, 1);
    const prepared = prepareWpsPlan(host.application, [1], plan([
      { type: "setFont", sheet: "明细", range: "A1:A3", size: 14 },
      { type: "setAlignment", sheet: "明细", range: "A1:A3", horizontal: "center" }
    ]));
    // 前序动作声明的是 size，后一个动作声明的是 horizontalAlignment —— 不相交，
    // 所以合成结果与预览快照完全一致（不该混进 size 的改动）。
    expect(composedFormatSnapshot(prepared.actions, 1)).toEqual(prepared.actions[1].formatSnapshot);
  });

  it("值动作仍走值矩阵指纹、不带格式属性；格式动作缺快照时拒绝登记", () => {
    const host = formatHost(1, 1);
    const valuePlan: AnalysisPlan = {
      ...plan([{ type: "clearRange", sheet: "明细", range: "A1", applyTo: "contents" }])
    };
    const valueDescriptor = wpsRecoveryDescriptor(
      prepareWpsPlan(host.application, [1], valuePlan)
    );
    expect(valueDescriptor.actions[0].previewFingerprint).not.toMatch(/^f1:/);
    expect(valueDescriptor.actions[0].formatFields).toBeUndefined();
    expect(valueDescriptor.actions[0].expectedFingerprint).not.toMatch(/^f1:/);

    // 格式动作没有预览快照 ⇒ 拒绝登记，而不是拿空快照凑一个"永远达不到"的目标。
    const formatPrepared = prepareWpsPlan(host.application, [1], plan([
      { type: "setFont", sheet: "明细", range: "A1", bold: true }
    ]));
    delete formatPrepared.actions[0].formatSnapshot;
    expect(() => wpsRecoveryDescriptor(formatPrepared)).toThrow(/缺少预览时采集的格式快照/);
  });
});
