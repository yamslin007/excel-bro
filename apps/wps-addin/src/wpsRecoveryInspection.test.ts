import { describe, expect, it } from "vitest";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";
import type { WpsApplication, WpsRange } from "./wpsHost";
import { matrixFingerprintV2, wpsFormatEvidenceFingerprint } from "./wpsPlanExecutor";
import type { WpsFormatSnapshot } from "./wpsFormatSnapshot";
import { inspectWpsRecoveryState, inspectWpsRecoveryStateDetail } from "./wpsRecoveryInspection";

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

/* ------------------------------------------------------------------------- *
 * 原因码：把"证据不足"拆开
 * ------------------------------------------------------------------------- *
 *
 * 起因是 2026-10-28 夜间的真机现象：完整退出 WPS 后重开，第一次显示「证据不足」，
 * 点一次「重新核对（只读）」就恢复。`unknown` 有十几条产生路径，界面只有四个字，
 * 现场分不出是授权没解析、身份读不出、还是宿主把值返回成 undefined。
 *
 * 这一组的判据是**每个 return 点各有各的码**，且结论仍是 `unknown` ——
 * 加码不放松任何判据：能给出码的照样阻断，只是原因不再被吞掉。
 */
describe("现场核对的原因码：每条失败路径各自可辨", () => {
  it("已选表为空 → authorize_no_selection（重开后范围未就绪）", () => {
    const { application, state, reads } = fixture();
    const detail = inspectWpsRecoveryStateDetail(application, state, []);
    expect(detail.outcome).toBe("unknown");
    expect(detail.code).toBe("authorize_no_selection");
    expect(detail.phase).toBe("authorize");
    expect(detail.selectedSheetCount).toBe(0);
    expect(reads).toEqual([]);
  });

  it("授权解析抛错 → authorize_threw，且已经能报出异常 name", () => {
    const { application, state } = fixture();
    const detail = inspectWpsRecoveryStateDetail(application, state, [9]);
    expect(detail.outcome).toBe("unknown");
    expect(detail.code).toBe("authorize_threw");
    // 名字是 `WpsAuthorizationError`（不是 message）：够判"授权本身拒绝"，
    // 又不带出 message 里可能出现的表名。
    expect(detail.errorName).toBe("WpsAuthorizationError");
  });

  it("账本登记的表不在当前授权里 → authorized_sheet_not_allowed", () => {
    const { application, state } = fixture();
    expect(inspectWpsRecoveryStateDetail(application, state, [2]).code).toBe("authorized_sheet_not_allowed");
  });

  it("动作目标表不在当前授权里 → action_sheet_not_allowed", () => {
    const { application, state, reads } = fixture();
    state.descriptor.actions[1].sheet = "Other";
    const detail = inspectWpsRecoveryStateDetail(application, state, [1]);
    expect(detail.code).toBe("action_sheet_not_allowed");
    expect(detail.actionIndex).toBe(1);
    // 校验阶段就停住，一个值都没读。
    expect(reads).toEqual([]);
  });

  it("动作目标表不在登记授权里 → action_sheet_not_authorized", () => {
    const { application, state } = fixture();
    state.authorizedSheetNames = ["Other"];
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("authorized_sheet_not_allowed");
    state.authorizedSheetNames = ["Allowed"];
    state.descriptor.actions[0].sheet = "Allowed";
    state.descriptor.actions[1].sheet = "Other";
    expect(inspectWpsRecoveryStateDetail(application, state, [1, 2]).code).toBe("action_sheet_not_authorized");
  });

  it("动作序号错位 → action_index_mismatch", () => {
    const { application, state } = fixture();
    state.descriptor.actions[1].actionIndex = 5;
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("action_index_mismatch");
  });

  it("描述的动作数与总数不符 → descriptor_action_count", () => {
    const { application, state } = fixture();
    state.totalActions = 3;
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("descriptor_action_count");
  });

  it("描述里没有动作 → descriptor_empty", () => {
    const { application, state } = fixture();
    state.totalActions = 0;
    state.descriptor.actions = [];
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("descriptor_empty");
  });

  it("旧描述缺写后指纹 → action_fingerprint_missing", () => {
    const { application, state } = fixture();
    state.descriptor.actions[1].expectedFingerprint = "";
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("action_fingerprint_missing");
  });

  it("动作总数超写入上限 → cell_budget_exceeded", () => {
    const { application, state } = fixture();
    // 上限是 `capabilities.wpsWrite.maxCells`（1000）。形状必须先自洽（否则先报形状不符），
    // 再让总格数越界 —— 这正是"描述合法但规模不被允许"的那一支。
    state.totalActions = 1;
    state.nextActionIndex = 0;
    state.descriptor.actions = [{
      actionIndex: 0, type: "writeValues", sheet: "Allowed", range: "A1:B1000",
      rowCount: 1000, columnCount: 2, cellCount: 2000,
      previewFingerprint: fp(null), expectedFingerprint: fp("after")
    }];
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("cell_budget_exceeded");
  });

  it("身份与账本记录不同 → workbook_identity_mismatch", () => {
    const { application, state } = fixture();
    state.workbookIdentity = "wps-path:c://other//test.xlsx";
    const detail = inspectWpsRecoveryStateDetail(application, state, [1]);
    expect(detail.code).toBe("workbook_identity_mismatch");
    expect(detail.phase).toBe("identity");
  });

  it("身份读不出来（后缀不合规）→ workbook_identity_unreadable", () => {
    const { application, state } = fixture();
    (application.ActiveWorkbook as unknown as { FullName: string }).FullName = "C://data//test.xls";
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("workbook_identity_unreadable");
  });

  it("宿主交不出区域对象 → host_range_unreadable", () => {
    const { application, state } = fixture();
    flickAfterAuthorize(application, { itemReturns: () => ({ Name: "Allowed", Range: () => undefined }) });
    const detail = inspectWpsRecoveryStateDetail(application, state, [1]);
    expect(detail.code).toBe("host_range_unreadable");
    expect(detail.phase).toBe("read");
  });

  it("宿主交出的区域形状不符 → host_range_shape_mismatch", () => {
    const { application, state } = fixture();
    flickAfterAuthorize(application, {
      itemReturns: () => ({
        Name: "Allowed",
        Range: () => ({ Rows: { Count: 3 }, Columns: { Count: 1 }, Value2: "after", Formula: "after" })
      })
    });
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("host_range_shape_mismatch");
  });

  it("授权成功之后工作簿消失 → workbook_unreadable", () => {
    const { application, state } = fixture();
    flickAfterAuthorize(application, { bookVanishes: true });
    const detail = inspectWpsRecoveryStateDetail(application, state, [1]);
    expect(detail.outcome).toBe("unknown");
    expect(detail.code).toBe("workbook_unreadable");
    expect(detail.phase).toBe("read");
  });

  it("Value2/Formula 读回 undefined（宿主未就绪）→ value_read_undefined", () => {
    const { application, state } = fixture({ A1: "after", B2: undefined });
    state.descriptor.actions[1].previewFingerprint = fp(null);
    const detail = inspectWpsRecoveryStateDetail(application, state, [1]);
    expect(detail.code).toBe("value_read_undefined");
    expect(detail.actionIndex).toBe(1);
  });

  it("读回形状不是登记的矩阵 → value_shape_unexpected", () => {
    const { application, state } = fixture({ A1: "after", B2: {} });
    state.descriptor.actions[1].previewFingerprint = fp(null);
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("value_shape_unexpected");
  });

  it("现场既不是写后也不是写前 → conflict/in_flight_unexpected", () => {
    const { application, state } = fixture({ A1: "after", B2: "human" });
    state.inFlightActionIndex = 1;
    const detail = inspectWpsRecoveryStateDetail(application, state, [1]);
    expect(detail.outcome).toBe("conflict");
    expect(detail.code).toBe("in_flight_unexpected");
  });

  it("已完成动作被改过 → conflict/completed_action_mismatch", () => {
    const { application, state } = fixture({ A1: "human", B2: "before" });
    const detail = inspectWpsRecoveryStateDetail(application, state, [1]);
    expect(detail.outcome).toBe("conflict");
    expect(detail.code).toBe("completed_action_mismatch");
    expect(detail.actionIndex).toBe(0);
  });

  it("格式动作没声明属性 → format_fields_undeclared", () => {
    const { application } = formatFixture([{ bold: true, size: 14 }, { bold: true, size: 14 }]);
    const state = formatState(BEFORE, AFTER, { formatFields: [] });
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("format_fields_undeclared");
  });

  it("格式动作属性名认不出 → format_fields_unrecognized", () => {
    const { application } = formatFixture([{ bold: true, size: 14 }, { bold: true, size: 14 }]);
    const state = formatState(BEFORE, AFTER, { formatFields: ["bold", "underline"] });
    expect(inspectWpsRecoveryStateDetail(application, state, [1]).code).toBe("format_fields_unrecognized");
  });

  it("格式快照不完整 → format_snapshot_incomplete", () => {
    const { application, range } = formatFixture([{ bold: true, size: 14 }, { bold: true, size: 14 }]);
    delete (range as { Item?: unknown }).Item;
    const detail = inspectWpsRecoveryStateDetail(application, formatState(BEFORE, AFTER), [1]);
    expect(detail.code).toBe("format_snapshot_incomplete");
    expect(detail.phase).toBe("format");
  });

  it("核验函数自身抛错 → inspection_threw，仍带阶段与异常 name", () => {
    const { application, state } = fixture();
    flickAfterAuthorize(application, { itemThrows: true });
    const detail = inspectWpsRecoveryStateDetail(application, state, [1]);
    expect(detail.outcome).toBe("unknown");
    expect(detail.code).toBe("inspection_threw");
    expect(detail.phase).toBe("read");
    expect(detail.errorName).toBe("TypeError");
  });

  it("正常的两种结论不带原因码（加码不等于给正常路径加噪音）", () => {
    const { application, state } = fixture();
    const detail = inspectWpsRecoveryStateDetail(application, state, [1]);
    expect(detail.outcome).toBe("not_started");
    expect(detail.code).toBeUndefined();
    expect(detail.selectedSheetCount).toBe(1);
    expect(detail.authorizedSheetCount).toBe(1);
  });
});

/**
 * 让"授权之后"的宿主调用开始出问题。
 *
 * 直接替换 `Worksheets.Item` 会把失败点挪到授权阶段 —— `authorizeWpsSheets` 自己也要
 * `Item()` 读表名。那就测不到读取阶段的分支。这里用一个确定的时点：授权在**读完所有表名
 * 之后**才会读 `workbook.FullName`（用来拼身份），而读取值的循环不会再读它。所以让
 * `FullName` 的 getter 在返回后把开关翻掉，正好落在"授权成功、还没开始读值"之间。
 */
function flickAfterAuthorize(
  application: WpsApplication,
  options: { itemThrows?: boolean; itemReturns?: () => unknown; bookVanishes?: boolean } = {}
): void {
  const book = application.ActiveWorkbook!;
  const sheets = book.Worksheets;
  let authorized = false;
  if (options.itemThrows || options.itemReturns) {
    const originalItem = sheets.Item;
    (sheets as unknown as { Item: (index: number) => unknown }).Item = (index: number) => {
      if (!authorized) return originalItem(index);
      if (options.itemThrows) throw new TypeError("对象不存在");
      return options.itemReturns!();
    };
  }
  if (options.bookVanishes) {
    Object.defineProperty(application, "ActiveWorkbook", {
      configurable: true,
      get() { return authorized ? undefined : book; }
    });
  }
  const fullName = book.FullName;
  Object.defineProperty(book, "FullName", {
    configurable: true,
    get() { authorized = true; return fullName; }
  });
}
