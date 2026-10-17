import { describe, expect, it } from "vitest";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange } from "./wpsHost";
import { matrixFingerprintV2, wpsFormatEvidenceFingerprint } from "./wpsPlanExecutor";
import type { WpsFormatSnapshot } from "./wpsFormatSnapshot";
import { inspectWpsRecoveryState } from "./wpsRecoveryInspection";

const fp = (value: string | null) => matrixFingerprintV2([[value]], [[value]]);

function fixture(values: Record<string, unknown> = { A1: "after", B2: "before" }) {
  const reads: string[] = [];
  const sheet = (name: string) => ({ Name: name, Range(address: string) {
    const value = values[address];
    return { Rows: { Count: 1 }, Columns: { Count: 1 },
      get Value2() { reads.push(`${name}!${address}`); return value; },
      get Formula() { return value; }
    };
  }});
  const sheets = [sheet("Allowed"), sheet("Other")];
  const application = { ActiveWorkbook: { Name: "test.xlsx", FullName: "C:\\data\\test.xlsx",
    Worksheets: { Count: 2, Item: (index: number) => sheets[index - 1] } } } as WpsApplication;
  const state: WpsExecutionState = { operationId: "execution-test", taskId: "task-test", status: "failed",
    workbookIdentity: "wps-path:c:\\data\\test.xlsx", revision: 2, authorizedSheetNames: ["Allowed"],
    ownerActive: false, canResume: true, requiresReconciliation: false,
    nextActionIndex: 1, totalActions: 2, inFlightActionIndex: null,
    descriptor: { actions: ["A1", "B2"].map((range, actionIndex) => ({
      actionIndex, type: "writeValues", sheet: "Allowed", range,
      rowCount: 1, columnCount: 1, cellCount: 1, previewFingerprint: fp("before"), expectedFingerprint: fp("after")
    })) }
  };
  return { application, state, reads };
}

describe("CAP-03 按登记坐标重读现场，非 PluginStorage 自证", () => {
  it("只读取已授权的明确目标，已完成和剩余动作分别对齐写后和写前", () => {
    const { application, state, reads } = fixture();
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("not_started");
    expect(reads).toEqual(["Allowed!A1", "Allowed!B2"]);
  });
  it("全部动作写后相符才证明零剩余完成", () => {
    const { application, state } = fixture({ A1: "after", B2: "after" });
    state.nextActionIndex = 2;
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("completed");
  });
  it.each(["A1", "B2"])("人工修改 %s：无论已完成或未执行都判冲突", (address) => {
    const { application, state } = fixture({ A1: "after", B2: "before", [address]: "human" });
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("conflict");
  });
  it.each(["before", "after"])("在途动作现场 %s：只核对不写入", (value) => {
    const { application, state } = fixture({ A1: "after", B2: value });
    state.inFlightActionIndex = 1;
    expect(inspectWpsRecoveryState(application, state, [1])).toBe(value === "after" ? "completed" : "not_started");
  });
  it("授权缺失时 Value2 零读取，不能沿用登记授权扩大今天范围", () => {
    const { application, state, reads } = fixture();
    expect(inspectWpsRecoveryState(application, state, [2])).toBe("unknown");
    expect(reads).toEqual([]);
  });
  it("后续动作越界：先验证全部范围，不能先读合法动作再拒绝", () => {
    const { application, state, reads } = fixture();
    state.descriptor.actions[1].sheet = "Other";
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("unknown");
    expect(reads).toEqual([]);
  });
  it.each([undefined, {}, []])("未采集/畸形数据不能归一化冒充空格：%j", (raw) => {
    const { application, state } = fixture({ A1: "after", B2: raw });
    state.descriptor.actions[1].previewFingerprint = fp(null);
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("unknown");
  });
  it("值未采集与真正空白 null 明确区分", () => {
    const { application, state } = fixture({ A1: "after", B2: null });
    state.descriptor.actions[1].previewFingerprint = fp(null);
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("not_started");
  });
  it("跨关闭只核对公式文本，不把计算值相同当作公式相同", () => {
    const { application, state } = fixture({ A1: "=A2+1", B2: "before" });
    state.descriptor.actions[0].type = "writeFormulas";
    state.descriptor.actions[0].expectedFingerprint = fp("=A2+2");
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("conflict");
  });
});

/* ------------------------------------------------------------------------- *
 * 格式动作的现场核对（Task #62）
 * ------------------------------------------------------------------------- *
 *
 * 格式动作**不改值**，值矩阵指纹在写前写后完全一样 —— 用它判"在途动作写没写成"
 * 永远判不出结果，只会把已写完的现场报成冲突，于是"核对完成并补记"这条路走不通。
 * 下面这一组钉住"按格式证据指纹核对、且一个 setter 都不碰"。
 */

type FormatCell = {
  bold: boolean;
  italic: boolean;
  name: string;
  size: number;
  color: number;
  horizontalAlignment: number;
  wrapText: boolean;
};

function baseFormat(overrides: Partial<FormatCell> = {}): FormatCell {
  return {
    bold: false,
    italic: false,
    name: "宋体",
    size: 11,
    color: 0,
    horizontalAlignment: 1,
    wrapText: false,
    ...overrides
  };
}

/** 快照只用于算指纹：声明属性之外的字段填什么都不影响 `setFont` 的 bold/size。 */
function snapshotOf(row: Array<Partial<FormatCell>>): WpsFormatSnapshot {
  return {
    // 单行两列：外层是"行"。
    cells: [row.map((cell) => ({ ...baseFormat(), ...cell }))],
    rowHeights: [],
    captured: true
  };
}

const FONT_FIELDS = ["bold", "size"] as const;

/**
 * 有状态格式宿主：`A1:B1` 两格各自持有字体状态。所有写入口都计一次 `writes` ——
 * 核对阶段必须是**纯读**，一个 setter 都不许被碰到。
 */
function formatFixture(cells: [Partial<FormatCell>, Partial<FormatCell>]) {
  const state: [FormatCell, FormatCell] = [baseFormat(cells[0]), baseFormat(cells[1])];
  let writes = 0;
  const guarded = <T extends object>(target: T): T =>
    new Proxy(target, {
      set: () => {
        writes += 1;
        return true;
      }
    });

  const cellAt = (index: number): WpsRange => {
    const cell = state[index];
    const font = guarded({
      get Bold() { return cell.bold; },
      set Bold(_value: unknown) { writes += 1; },
      get Italic() { return cell.italic; },
      set Italic(_value: unknown) { writes += 1; },
      get Name() { return cell.name; },
      set Name(_value: unknown) { writes += 1; },
      get Size() { return cell.size; },
      set Size(_value: unknown) { writes += 1; },
      get Color() { return cell.color; },
      set Color(_value: unknown) { writes += 1; }
    });
    return guarded({
      Font: font,
      get HorizontalAlignment() { return cell.horizontalAlignment; },
      set HorizontalAlignment(_value: unknown) { writes += 1; },
      get WrapText() { return cell.wrapText; },
      set WrapText(_value: unknown) { writes += 1; },
      get RowHeight() { return 20; },
      set RowHeight(_value: unknown) { writes += 1; }
    } as unknown as WpsRange);
  };

  const range = guarded({
    Rows: { Count: 1 },
    Columns: { Count: 2 },
    Font: cellAt(0).Font,
    HorizontalAlignment: 1,
    WrapText: false,
    RowHeight: 20,
    // 1 行 2 列：`Item(1, c)` 就是第 c 格（逐格定位，绝不回退整块）。
    Item: (_row: number, column: number) => cellAt(column - 1)
  } as unknown as WpsRange);
  const sheet = { Name: "Allowed", Range: () => range };
  const application = {
    ActiveWorkbook: {
      Name: "test.xlsx",
      FullName: "C:\\data\\test.xlsx",
      Worksheets: { Count: 1, Item: () => sheet }
    }
  } as unknown as WpsApplication;
  return { application, writes: () => writes, range };
}

function formatState(
  before: Array<Partial<FormatCell>>,
  after: Array<Partial<FormatCell>>,
  overrides: Partial<WpsExecutionState["descriptor"]["actions"][number]> = {}
): WpsExecutionState {
  const fingerprint = (row: Array<Partial<FormatCell>>) =>
    wpsFormatEvidenceFingerprint("setFont", [...FONT_FIELDS], snapshotOf(row));
  return {
    operationId: "execution-format",
    taskId: "task-format",
    status: "failed",
    workbookIdentity: "wps-path:c:\\data\\test.xlsx",
    revision: 3,
    authorizedSheetNames: ["Allowed"],
    ownerActive: false,
    canResume: true,
    requiresReconciliation: true,
    nextActionIndex: 0,
    totalActions: 1,
    inFlightActionIndex: 0,
    descriptor: {
      actions: [
        {
          actionIndex: 0,
          type: "setFont",
          sheet: "Allowed",
          range: "A1:B1",
          rowCount: 1,
          columnCount: 2,
          cellCount: 2,
          previewFingerprint: fingerprint(before),
          expectedFingerprint: fingerprint(after),
          formatFields: [...FONT_FIELDS],
          ...overrides
        }
      ]
    }
  };
}

const BEFORE: Array<Partial<FormatCell>> = [{ bold: false, size: 11 }, { bold: false, size: 11 }];
const AFTER: Array<Partial<FormatCell>> = [{ bold: true, size: 14 }, { bold: true, size: 14 }];

describe("格式动作的现场核对：按格式证据指纹判断写前 / 写后 / 都不是", () => {
  it("在途动作已是写后 → 已完成，且全程零写入", () => {
    const { application, writes } = formatFixture([{ bold: true, size: 14 }, { bold: true, size: 14 }]);
    expect(inspectWpsRecoveryState(application, formatState(BEFORE, AFTER), [1])).toBe("completed");
    expect(writes()).toBe(0);
  });

  it("在途动作仍是写前 → 未开始", () => {
    const { application, writes } = formatFixture([{ bold: false, size: 11 }, { bold: false, size: 11 }]);
    expect(inspectWpsRecoveryState(application, formatState(BEFORE, AFTER), [1])).toBe("not_started");
    expect(writes()).toBe(0);
  });

  it("只写了一半（一格写后、一格写前）→ 冲突，不倒向任何一侧", () => {
    const { application } = formatFixture([{ bold: true, size: 14 }, { bold: false, size: 11 }]);
    expect(inspectWpsRecoveryState(application, formatState(BEFORE, AFTER), [1])).toBe("conflict");
  });

  it("写后被外部改过（值没变、格式变了）→ 冲突；值矩阵指纹在这里没有区分力", () => {
    const { application } = formatFixture([{ bold: true, size: 18 }, { bold: true, size: 14 }]);
    expect(inspectWpsRecoveryState(application, formatState(BEFORE, AFTER), [1])).toBe("conflict");
  });

  it("属性名不可识别 → 未知，不用剩余属性凑一个哈希", () => {
    const { application } = formatFixture([{ bold: true, size: 14 }, { bold: true, size: 14 }]);
    const state = formatState(BEFORE, AFTER, { formatFields: ["bold", "underline"] });
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("unknown");
  });

  it("登记没声明属性 → 未知，不按值矩阵指纹回退核对", () => {
    const { application } = formatFixture([{ bold: true, size: 14 }, { bold: true, size: 14 }]);
    const state = formatState(BEFORE, AFTER, { formatFields: [] });
    expect(inspectWpsRecoveryState(application, state, [1])).toBe("unknown");
  });

  it("多格区域缺逐格定位（无 Item）→ 未知，不回退整块格式值", () => {
    const { application, range } = formatFixture([{ bold: true, size: 14 }, { bold: true, size: 14 }]);
    delete (range as { Item?: unknown }).Item;
    expect(inspectWpsRecoveryState(application, formatState(BEFORE, AFTER), [1])).toBe("unknown");
  });
});
