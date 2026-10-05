/**
 * `wpsObjectIdentity` 的单元测试 —— 直接打判据本身，不经界面。
 *
 * 面板级用例（`wpsClarificationCollision.test.tsx`）证明的是"接上了、会失效"，
 * 这里证明的是**判据的内部逻辑**：哪一层能证明不同、哪一层只能判未知，
 * 以及"未知"必须按不安全处理。
 */
import { describe, expect, it } from "vitest";
import {
  captureWpsObjectIdentity,
  verifyWpsObjectIdentity,
  wpsIdentityIsConclusive
} from "./wpsObjectIdentity";
import type { WpsApplication } from "./wpsHost";

function host(options: {
  workbookName?: string;
  path?: string | null;
  unsaved?: boolean;
  sheets?: Array<{ name: string; codeName?: string | null; rows?: number; columns?: number }>;
}) {
  const workbookName = options.workbookName ?? "测试.xlsx";
  const path = options.path === undefined ? "C:\\Data" : options.path;
  const sheetSpecs = options.sheets ?? [{ name: "订单" }];

  const makeSheet = (spec: { name: string; codeName?: string | null; rows?: number; columns?: number }, position: number) => {
    const usedRange = {
      Rows: { Count: spec.rows ?? 3 },
      Columns: { Count: spec.columns ?? 2 },
      Row: 1,
      Column: 1,
      MergeCells: false,
      Value2: [["类别", "归属"]]
    } as Record<string, unknown>;
    usedRange.Item = () => usedRange;
    const sheet: Record<string, unknown> = {
      Name: spec.name,
      Index: position + 1,
      ProtectContents: false,
      UsedRange: usedRange,
      Range: () => ({ Value2: [["类别", "归属"]] })
    };
    // `codeName: null` 用来模拟"宿主不提供稳定 ID"。
    if (spec.codeName !== null) sheet.CodeName = spec.codeName ?? `${spec.name}-code`;
    return sheet;
  };

  let sheets = sheetSpecs.map(makeSheet);
  /**
   * `Count` 必须是**取值器**：否则"删掉一张表"之后数量还是旧值，
   * 越界判据就测不到 —— 那样测试会绿在错误的地方。
   */
  const makeWorkbook = (Name: string) => ({
    Name,
    Worksheets: {
      get Count() {
        return sheets.length;
      },
      Item: (index: number) => sheets[Number(index) - 1]
    },
    get ActiveSheet() {
      return sheets[0];
    }
  });
  let workbook: Record<string, unknown> = makeWorkbook(workbookName);
  if (!options.unsaved) {
    workbook.FullName = `${path ?? "C:\\Data"}\\${workbookName}`;
    workbook.Path = path;
  }

  const application: Record<string, unknown> = {
    Version: "12.0",
    get ActiveWorkbook() {
      return workbook;
    },
    get ActiveSheet() {
      return sheets[0];
    },
    PluginStorage: { getItem: () => undefined, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  };

  return {
    application: application as never,
    /** 打开**另一本**工作簿：全新对象（真宿主语义）。 */
    replaceWorkbook: (next: { workbookName?: string; path?: string | null; unsaved?: boolean }) => {
      const name = next.workbookName ?? workbookName;
      workbook = makeWorkbook(name);
      if (!next.unsaved) {
        const nextPath = next.path === undefined ? path : next.path;
        workbook.FullName = `${nextPath ?? "C:\\Data"}\\${name}`;
        workbook.Path = nextPath;
      }
    },
    /**
     * 只在**同一个工作簿对象**上改路径 —— 用来单独验证第 2 层判据
     * （引用没变、只有路径变了，也必须判不同）。
     */
    mutatePath: (nextPath: string) => {
      workbook.FullName = `${nextPath}\\${String(workbook.Name)}`;
      workbook.Path = nextPath;
    },
    /** 删除后以相同索引、同名、同表头、同尺寸重建。 */
    rebuildSheet: (position: number) => {
      const spec = sheetSpecs[position] ?? { name: "订单" };
      sheets = sheets.map((sheet, index) => (index === position ? makeSheet(spec, position) : sheet));
    },
    /** 只换结构尺寸（同索引同名同表头）。 */
    resizeSheet: (position: number, rows: number, columns: number) => {
      const spec = { ...(sheetSpecs[position] ?? { name: "订单" }), rows, columns };
      sheetSpecs[position] = spec;
      sheets = sheets.map((sheet, index) => (index === position ? makeSheet(spec, position) : sheet));
    },
    /** 删掉一张表（同名对象不动，只有集合变小）。 */
    dropSheet: (position: number) => {
      sheetSpecs.splice(position, 1);
      sheets = sheets.filter((_, index) => index !== position);
    },
    /** 模拟"工作簿被关掉"：把它置空。 */
    detachWorkbook: () => {
      workbook = undefined as never;
    }
  };
}

describe("宿主对象身份：捕获", () => {
  it("捕获得到工作簿引用、路径与每张表的稳定 ID 和结构", () => {
    const h = host({ sheets: [{ name: "订单" }, { name: "明细" }] });
    const identity = captureWpsObjectIdentity(h.application, [1, 2]);
    expect(identity.workbookName).toBe("测试.xlsx");
    expect(identity.workbookPath).toBe("C:\\Data");
    expect(identity.sheets.map((sheet) => sheet.name)).toEqual(["订单", "明细"]);
    expect(identity.sheets[0].id).toBe("订单-code");
    expect(identity.sheets[0].shape).toEqual([1, 1, 3, 2]);
  });

  it("宿主不提供稳定 ID 时留 undefined，不凭空造一个", () => {
    const h = host({ sheets: [{ name: "订单", codeName: null }] });
    const identity = captureWpsObjectIdentity(h.application, [1]);
    expect(identity.sheets[0].id).toBeUndefined();
    expect(identity.sheets[0].shape).toEqual([1, 1, 3, 2]);
  });

  it("系统表不进身份记录", () => {
    const h = host({ sheets: [{ name: "#EB_RULES" }, { name: "订单" }] });
    const identity = captureWpsObjectIdentity(h.application, [1, 2]);
    expect(identity.sheets.map((sheet) => sheet.name)).toEqual(["订单"]);
  });

  it("没有打开工作簿时抛错，不静默返回空身份", () => {
    const h = host({});
    h.detachWorkbook();
    expect(() => captureWpsObjectIdentity(h.application, [1])).toThrow();
  });

  it("序号越界时抛错", () => {
    const h = host({ sheets: [{ name: "订单" }] });
    expect(() => captureWpsObjectIdentity(h.application, [9])).toThrow();
  });
});

describe("宿主对象身份：复核", () => {
  it("什么都没变时判为相同", () => {
    const h = host({});
    const identity = captureWpsObjectIdentity(h.application, [1]);
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("same");
  });

  it("换到**另一个目录的同名工作簿**判为不同（指纹看不见的那一类）", () => {
    const h = host({ workbookName: "销售.xlsx", path: "D:\\甲公司" });
    const identity = captureWpsObjectIdentity(h.application, [1]);
    h.replaceWorkbook({ workbookName: "销售.xlsx", path: "E:\\乙公司" });
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("different");
  });

  it("第 2 层：工作簿对象没换、只有路径变了，也必须判不同", () => {
    const h = host({ workbookName: "销售.xlsx", path: "D:\\甲公司" });
    const identity = captureWpsObjectIdentity(h.application, [1]);
    // 同一份工作簿对象，只把路径改掉 —— 模拟宿主复用对象但换了文件。
    h.mutatePath("E:\\乙公司");
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("different");
  });

  it("同索引同名同表头地**重建**工作表：判 unknown（证明不了），不是 different", () => {
    const h = host({ sheets: [{ name: "订单" }] });
    const identity = captureWpsObjectIdentity(h.application, [1]);
    h.rebuildSheet(0);
    // 名称、表头、尺寸全同，只有对象引用变了 —— 但我们分辨不了
    // "真的重建了一张表" 与 "宿主对同一张表返回了新包装对象"。
    // 把这种情况说成 "different" 正是那个假话的来源：用户只改了一格又撤回，
    // 界面却告诉他"数据范围已经变化"。
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("unknown");
  });

  it("宿主不提供稳定 ID 时，重建同样判 unknown（不靠 ID，也不靠引用反推变化）", () => {
    const h = host({ sheets: [{ name: "订单", codeName: null }] });
    const identity = captureWpsObjectIdentity(h.application, [1]);
    expect(identity.sheets[0].id).toBeUndefined();
    h.rebuildSheet(0);
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("unknown");
  });

  it("引用没变但稳定 ID 被换掉：判 different（有正面变化证据）", () => {
    const h = host({ sheets: [{ name: "订单" }] });
    const identity = captureWpsObjectIdentity(h.application, [1]);
    // 同一个对象上把 CodeName 改掉。
    (h.application as unknown as { ActiveWorkbook: { Worksheets: { Item: (i: number) => Record<string, unknown> } } })
      .ActiveWorkbook.Worksheets.Item(1).CodeName = "别的-code";
    // 引用相同却没有"引用变了"的疑点，ID 还变了 —— 这是一条真正的变化证据。
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("different");
  });

  it("只改尺寸（同索引同名同表头）判 unknown，不判 different", () => {
    const h = host({ sheets: [{ name: "订单", codeName: null }] });
    const identity = captureWpsObjectIdentity(h.application, [1]);
    h.resizeSheet(0, 5, 2);
    // 尺寸本来就不参与"放行"，也不足以反推"换了对象"。
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("unknown");
  });

  it("工作表被删除（索引越界）判为不同", () => {
    const h = host({ sheets: [{ name: "订单" }, { name: "明细" }] });
    const identity = captureWpsObjectIdentity(h.application, [2]);
    // 删掉第一张后只剩一张：原来的索引 2 越界（工作簿对象本身没换）。
    h.dropSheet(0);
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("different");
  });

  it("工作簿被关掉时判为不同（保守，不是 unknown）", () => {
    const h = host({});
    const identity = captureWpsObjectIdentity(h.application, [1]);
    h.detachWorkbook();
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("different");
  });

  it("未保存工作簿：没有路径，但对象引用未变时仍判为相同", () => {
    const h = host({ unsaved: true });
    const identity = captureWpsObjectIdentity(h.application, [1]);
    expect(identity.workbookPath).toBeNull();
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("same");
  });

  it("未保存工作簿换成另一本：没有路径可依，只能判 unknown", () => {
    const h = host({ unsaved: true });
    const identity = captureWpsObjectIdentity(h.application, [1]);
    h.replaceWorkbook({ unsaved: true, workbookName: "另一本.xlsx" });
    // 未保存工作簿没有路径/完整名这条正面证据，换本与换包装分辨不了 ——
    // 这正是"未保存工作簿跨重开必然失效"的代价，由调用方按 unknown 处理。
    expect(verifyWpsObjectIdentity(h.application, identity)).toBe("unknown");
  });

  it("宿主读取抛错时判为未知（调用方按不安全处理）", () => {
    const h = host({});
    const identity = captureWpsObjectIdentity(h.application, [1]);
    const broken = {
      get ActiveWorkbook() {
        throw new Error("宿主拒绝访问");
      }
    } as unknown as WpsApplication;
    expect(verifyWpsObjectIdentity(broken, identity)).toBe("unknown");
  });
});

describe("会话身份是否达到放行强度", () => {
  it("没有身份、或一张表都没有时不算完整", () => {
    expect(wpsIdentityIsConclusive(null)).toBe(false);
  });

  it("捕获到至少一张表时算完整", () => {
    const h = host({});
    expect(wpsIdentityIsConclusive(captureWpsObjectIdentity(h.application, [1]))).toBe(true);
  });
});
