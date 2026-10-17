/**
 * 定向测试：澄清卡**授权身份指纹的碰撞边界与重生路径**。
 *
 * 上一条测试证明了「身份判据」与「表头判据」各自独立必要。这一条问更难的问题：
 * **这套指纹在什么时候会认不出来换了对象？**
 *
 * 逐项复核（用户点名）：
 *  1. 两个**不同目录的同名工作簿** —— 已知碰撞：指纹只带 `workbook.Name`，不带路径。
 *     这是**刻意**的隐私取舍（路径含用户名，指纹会落进对话历史）。这里如实钉住
 *     该行为，并在文档里列为**能力限制**，不声称已覆盖。
 *  2. 同名同表头的**替换工作表**：把索引 1 指向的表整张换掉（表名/表头/尺寸/数据全同）。
 *     —— 这是本指纹**结构性无法探测**的一类，必须靠"重开后待答澄清一律失效"兜底。
 *  3. 工作表**重排**、**改名**：指纹可探测（已在必要性测试覆盖，这里复测重生路径）。
 *  4. 面板**重开**后恢复澄清：内存里的建卡时身份已丢失，无法证明还是原授权对象 ——
 *     待答澄清必须**直接失效**，不能凭 localStorage 里存的指纹"自证"。
 *
 * 隐私边界同时被钉住：`localStorage` 里只允许出现工作簿**名**与工作表名，
 * 不得出现完整路径、也不得出现任何数据行。
 */
import { act, type ReactElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { listModels, getModelSettings, validatePlan } from "../../excel-addin/src/api";
import WpsAgentPanel from "./WpsAgentPanel";
import { runWpsReadonlyTurn } from "./wpsReadonlyAgent";
import { getWpsApplication } from "./wpsHost";

// These cases must still prove the stronger clarification guard independently.
vi.mock("../../excel-addin/src/hooks/useAuthorizedTarget", () => ({
  useAuthorizedTarget: () => ({ check: async () => true, checkForAnalysis: async () => true,
    clearAnalysisConsent: () => {}, hasAnalysisConsent: () => false,
    requestConfirmation: async () => false, notice: null, checking: false, binding: null })
}));

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  return { ...actual, understandMessage: vi.fn(async () => ({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" })), listModels: vi.fn(), getModelSettings: vi.fn(), validatePlan: vi.fn() };
});
vi.mock("./wpsReadonlyAgent", () => ({ runWpsReadonlyTurn: vi.fn() }));
vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const WPS_CHAT_STORAGE_KEY = "excel-bro.wps.chat.v1";

const listModelsMock = vi.mocked(listModels);
const getModelSettingsMock = vi.mocked(getModelSettings);
const validatePlanMock = vi.mocked(validatePlan);
const runTurnMock = vi.mocked(runWpsReadonlyTurn);
const getWpsApplicationMock = vi.mocked(getWpsApplication);
void validatePlanMock;

const firstConnection = {
  id: "connection:first",
  label: "第一个连接",
  provider: "model" as const,
  available: true,
  baseUrl: "https://example.test/v1",
  model: "test-model",
  hasApiKey: true,
  supportsVision: false
};

let mounted: Array<{ unmount: () => void }> = [];

async function act_(fn: () => void | Promise<void>) {
  await act(async () => {
    await fn();
  });
}

async function mount(element: ReactElement) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  await act_(() => {
    root.render(element);
  });
  mounted.push({ unmount: () => root.unmount() });
  return { container: host };
}

async function until(check: () => boolean, message: string, timeout = 2000) {
  const started = Date.now();
  while (Date.now() - started < timeout) {
    if (check()) return;
    await act_(() => new Promise((resolve) => setTimeout(resolve, 10)));
  }
  throw new Error(`等待超时：${message}`);
}

function textOf(container: HTMLElement) {
  return container.textContent ?? "";
}

async function click(target: HTMLElement | null) {
  if (!target) throw new Error("要点击的元素不存在");
  await act_(() => {
    target.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
}

async function type(target: HTMLTextAreaElement | null, value: string) {
  if (!target) throw new Error("要输入的元素不存在");
  const setter = Object.getOwnPropertyDescriptor(
    HTMLTextAreaElement.prototype,
    "value"
  )?.set;
  await act_(() => {
    setter?.call(target, value);
    target.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

async function submit(target: HTMLTextAreaElement | null) {
  if (!target) throw new Error("要提交的元素不存在");
  await act_(() => {
    target.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
  });
}

function panel() {
  return (
    <WpsAgentPanel
      selectedSheetIndexes={[1]}
      scopeDock={<div className="eb-scope-dock">当前表 Sheet1</div>}
    />
  );
}

/**
 * 表头、行列尺寸、数据值在三张表上完全一致的替身。
 *
 * 与必要性测试同一套语义（真宿主的 `sheet.Range(first, last)`，只接受两个参数），
 * 但额外暴露 `path`：用来构造「两个不同目录的同名工作簿」这一碰撞场景。
 */
function uniformHost(options: {
  workbookName?: string;
  path?: string;
  sheetNames: string[];
}) {
  const workbookName = options.workbookName ?? "测试.xlsx";
  const path = options.path ?? "C:\\Data";

  let headerRow: string[][] = [["类别", "归属"]];

  const makeRange = () => {
    const range: Record<string, unknown> = {
      Rows: { Count: 3 },
      Columns: { Count: 2 },
      MergeCells: false
    };
    Object.defineProperties(range, {
      Value2: {
        get: () => [
          ["类别", "归属"],
          ["办公用品", "华东"],
          ["家具", "华南"]
        ]
      },
      Item: { value: () => makeRange() }
    });
    return range;
  };

  const rangeFrom = () => {
    const range: Record<string, unknown> = {
      Rows: { Count: headerRow.length },
      Columns: { Count: headerRow[0]?.length ?? 0 },
      MergeCells: false
    };
    Object.defineProperties(range, {
      Value2: { get: () => headerRow.map((row) => [...row]) }
    });
    return range;
  };

  const storage = new Map<string, unknown>();

  /**
   * 构造一本工作簿的替身。
   *
   * **关键语义**：每次"打开另一本"都返回**全新的对象**，就像真宿主那样 ——
   * 同名同表头的另一本工作簿，在对象层面与原来那本毫无关系。指纹看不出差别，
   * 对象身份看得出。
   */
  const makeWorkbook = (options: {
    path: string;
    sheets: Array<Record<string, unknown>>;
    firstIndexOneBased?: boolean;
  }) => ({
    Name: workbookName,
    FullName: `${options.path}\\${workbookName}`,
    get Path() {
      return options.path;
    },
    ReadOnly: false,
    Worksheets: {
      Count: options.sheets.length,
      Item: (index: number) => options.sheets[Number(index) - 1]
    },
    ActiveSheet: options.sheets[0]
  });

  const makeSheet = (name: string, position: number) => ({
    Name: name,
    ProtectContents: false,
    /** 稳定 ID —— 真宿主由内核分配；这里每张表一个，重建时换新值。 */
    CodeName: `${name}__code__${position}`,
    get Index() {
      return position + 1;
    },
    get UsedRange() {
      return makeRange();
    },
    Range: () => rangeFrom()
  });

  let currentSheets = options.sheetNames.map(makeSheet);
  let workbook = makeWorkbook({ path, sheets: currentSheets });

  const application = {
    Version: "12.0",
    get ActiveWorkbook() {
      return workbook;
    },
    get ActiveSheet() {
      return currentSheets[0];
    },
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, item: unknown) => {
        storage.set(key, item);
      }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never;

  return {
    application,
    get sheets() {
      return currentSheets;
    },
    /** 打开**另一个目录**下的同名工作簿：名称、表名、表头全同，只有路径与本子对象变。 */
    openSameNameElsewhere: (otherPath: string) => {
      currentSheets = options.sheetNames.map(makeSheet);
      workbook = makeWorkbook({ path: otherPath, sheets: currentSheets });
    },
    /**
     * 删除后以**相同索引、相同名称、相同表头**重建一张表。
     *
     * 这是本模块真正要堵的缺口：指纹与表头判据都看不出来，只有对象身份能。
     */
    rebuildSameSheet: (position: number) => {
      const name = String(currentSheets[position].Name);
      currentSheets = currentSheets.map((sheet, index) =>
        index === position ? makeSheet(name, position) : sheet
      );
      workbook = makeWorkbook({ path, sheets: currentSheets });
    },
    rename: (position: number, name: string) => {
      currentSheets = currentSheets.map((sheet, index) =>
        index === position ? { ...sheet, Name: name } : sheet
      );
      workbook = makeWorkbook({ path, sheets: currentSheets });
    },
    reorder: (next: string[]) => {
      const byName = new Map(currentSheets.map((sheet) => [String(sheet.Name), sheet]));
      const ordered = next.flatMap((name) => {
        const sheet = byName.get(name);
        return sheet ? [sheet] : [];
      });
      if (ordered.length === currentSheets.length) currentSheets = ordered;
      workbook = makeWorkbook({ path, sheets: currentSheets });
    },
    setWorkbookName: (name: string) => {
      (workbook as { Name: string }).Name = name;
    },
    setHeader: (next: string[][]) => {
      headerRow = next;
    }
  };
}

function clarificationCard() {
  runTurnMock.mockResolvedValue({
    kind: "clarification",
    clarification: {
      response: {
        kind: "clarification",
        provider: "model",
        clarification: {
          id: `clarify-${Math.random().toString(16).slice(2)}`,
          summary: "需要明确目标",
          question: "要查询还是修改？",
          reason: "两种操作不同",
          scopeLabel: "当前表",
          options: [
            { id: "read", label: "查询", description: "只读", resolution: "查询数据" }
          ]
        }
      }
    },
    headerRows: [["类别", "归属"]]
  } as never);
}

async function pendingClarification() {
  const { container } = await mount(panel());
  await until(() => container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "应加载模型");
  const composer = container.querySelector<HTMLTextAreaElement>(".composer textarea");
  await type(composer, "帮我处理表格");
  await submit(composer);
  await until(() => Boolean(container.querySelector(".clarification-card")), "应显示澄清卡");
  await until(() => runTurnMock.mock.calls.length === 1, "应完成这一轮");
  return container;
}

function clarificationOptions(container: HTMLElement): HTMLButtonElement[] {
  return Array.from(
    container.querySelectorAll<HTMLButtonElement>(".clarification-options button")
  );
}

function storedHistoryTexts(): string[] {
  const raw = localStorage.getItem(WPS_CHAT_STORAGE_KEY);
  return raw == null ? [] : [raw];
}

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  listModelsMock.mockResolvedValue({
    defaultModelId: firstConnection.id,
    models: [firstConnection]
  });
  getModelSettingsMock.mockResolvedValue({} as never);
  getWpsApplicationMock.mockReturnValue({} as never);
});

afterEach(() => {
  mounted.forEach((entry) => entry.unmount());
  mounted = [];
  document.body.innerHTML = "";
});

describe("授权身份的碰撞边界与重生路径", () => {
  it("两个不同目录的同名工作簿：指纹相同，但对象身份必须识破并要求重发需求", async () => {
    clarificationCard();
    const host = uniformHost({ workbookName: "销售.xlsx", path: "D:\\甲公司", sheetNames: ["订单"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    // 换到**另一个目录**下的同名工作簿：工作表名、表头、尺寸、数据全同，只有路径变。
    host.openSameNameElsewhere("E:\\乙公司");
    expect(
      (host.application as { ActiveWorkbook: { FullName: string } }).ActiveWorkbook.FullName
    ).toBe("E:\\乙公司\\销售.xlsx");

    await click(clarificationOptions(container)[0]);
    /**
     * 指纹与表头判据**都**看不出差别（名字一样、表头一样），
     * 只有内存里的对象身份能 —— 这正是此前记录为"能力限制"的那一类，现已关闭。
     */
    await until(
      () => container.querySelector(".clarification-card.invalidated") !== null,
      "换到同名不同路径的工作簿后必须失效"
    );
    expect(runTurnMock).toHaveBeenCalledTimes(1);
    expect(textOf(container)).toContain("数据范围已经变化");

    // 路径绝不能为了"补救身份"而被写进对话历史。
    const stored = storedHistoryTexts().join("");
    expect(stored).not.toContain("E:\\乙公司");
    expect(stored).not.toContain("D:\\甲公司");
  });

  it("同索引、同名、同表头地删除后重建 —— 证明不了身份，不得谎称数据已变化", async () => {
    clarificationCard();
    const host = uniformHost({ sheetNames: ["订单", "明细"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    // 删掉索引 1 的「订单」，再在同一个位置新建一张同样叫「订单」、表头一致的表。
    host.rebuildSameSheet(0);

    await click(clarificationOptions(container)[0]);
    await until(
      () => container.querySelector(".clarification-card.unconfirmed") !== null,
      "同索引同名同表头的重建表必须挡住作答"
    );
    // 挡住是必须的；但**不能**断言"数据范围已经变化" —— 我们分辨不了
    // "真的重建了一张表"与"宿主换了包装对象"，那是两件事，只有前者才是变化。
    expect(textOf(container)).not.toContain("数据范围已经变化");
    expect(textOf(container)).toContain("无法确认");
    // 用户的原需求与他选的回答都要留着，供他重新确认范围后继续。
    expect(
      container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value ?? ""
    ).toContain("补充确认");
    expect(runTurnMock).toHaveBeenCalledTimes(1);
  });

  it("宿主对同一份工作簿/工作表返回新包装对象时，不得谎称数据范围已变化", async () => {
    clarificationCard();
    // 路径、工作簿名、表名、表头**全都不变**，只是宿主重新交了一份包装对象
    // —— 这就是"用户改了数据格又撤回"之后最可能的真实情形。
    const host = uniformHost({ sheetNames: ["订单", "明细"], path: "C:\\数据" });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    host.openSameNameElsewhere("C:\\数据");

    await click(clarificationOptions(container)[0]);
    await until(
      () => container.querySelector(".clarification-card.unconfirmed") !== null,
      "证明不了身份时必须挡住作答"
    );
    // 关键断言：范围其实没变，所以**不许**说"数据范围已经变化"。
    expect(textOf(container)).not.toContain("数据范围已经变化");
    expect(container.querySelector(".clarification-card.invalidated")).toBeNull();
    // 用户的回答被保留，重新确认范围后可以直接继续。
    expect(
      container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value ?? ""
    ).toContain("补充确认");
    expect(runTurnMock).toHaveBeenCalledTimes(1);
  });

  it("指纹兜底：同名同表头的替换工作表若伴随重排/改名，指纹也必须变", async () => {
    clarificationCard();
    const host = uniformHost({ sheetNames: ["订单", "订单（备份）"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    // 把索引 1 换成"另一张内容完全相同的表"，并让备份顶上原位置。
    host.rename(0, "订单（旧）");
    host.rename(1, "订单");
    host.reorder(["订单", "订单（旧）"]);

    await click(clarificationOptions(container)[0]);
    await until(
      () => container.querySelector(".clarification-card.invalidated") !== null,
      "名字集合变化后待答澄清必须失效"
    );
    expect(runTurnMock).toHaveBeenCalledTimes(1);
  });

  it("无法证明还是原授权对象时，保守要求重新选择范围并重发需求", async () => {
    clarificationCard();
    const host = uniformHost({ sheetNames: ["订单", "明细"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    host.rename(0, "订单★");

    await click(clarificationOptions(container)[0]);
    await until(
      () => container.querySelector(".clarification-card.invalidated") !== null,
      "改名后必须失效"
    );
    // 失效后不只是"红一下"：要给出可执行的下一步，并把原需求原样放回输入框。
    expect(textOf(container)).toContain("数据范围已经变化");
    expect(textOf(container)).toContain("请检查范围后重新发送需求");
    expect(
      container.querySelector<HTMLTextAreaElement>(".composer textarea")?.value
    ).toContain("帮我处理表格");
    // 澄清卡被收起，不会让用户对着已经失效的选项继续点。
    expect(container.querySelector(".clarification-options")).toBeNull();
  });

  it("重生：面板重开后，落盘的待答澄清收敛为“无法确认”（内存身份无从确认）", async () => {
    clarificationCard();
    const host = uniformHost({ sheetNames: ["订单", "明细"] });
    getWpsApplicationMock.mockReturnValue(host.application);

    const container = await pendingClarification();
    expect(container.querySelector(".clarification-card")).not.toBeNull();
    expect(container.querySelector(".clarification-options")).not.toBeNull();

    // 重开任务窗格：卸载后原样挂载（宿主存储保留）。
    mounted.forEach((entry) => entry.unmount());
    mounted = [];
    document.body.innerHTML = "";

    const reopened = await mount(panel());
    await until(() => reopened.container.querySelector<HTMLButtonElement>(".model-picker-trigger")?.title.includes("当前模型：第一个连接") ?? false, "重开应加载模型");
    await until(
      () => Boolean(reopened.container.querySelector(".clarification-card")),
      "重开后应恢复出那张卡（历史仍可见）"
    );
    // 身份只能在内存里确认，重开后即无从确认 —— 必须收敛成"无法确认"，
    // 而不是断言"数据范围已经变化"：重开本身没有给出任何变化证据。
    expect(reopened.container.querySelector(".clarification-card.unconfirmed")).not.toBeNull();
    expect(reopened.container.querySelector(".clarification-card.invalidated")).toBeNull();
    expect(reopened.container.querySelector(".clarification-options")).toBeNull();
  });

  it("未保存工作簿：同会话内对象引用仍可证明是同一本，正常作答", async () => {
    clarificationCard();
    // 未保存工作簿：没有 `FullName`/`Path`，只剩对象引用这一条判据。
    const host = uniformHost({ sheetNames: ["订单", "明细"] });
    const workbook = (host.application as { ActiveWorkbook: Record<string, unknown> })
      .ActiveWorkbook;
    delete workbook.FullName;
    delete workbook.Path;
    getWpsApplicationMock.mockReturnValue(host.application);

    runTurnMock
      .mockResolvedValueOnce({
        kind: "clarification",
        clarification: {
          response: {
            kind: "clarification",
            provider: "model",
            clarification: {
              id: "clarify-unsaved",
              summary: "需要明确目标",
              question: "要查询还是修改？",
              reason: "两种操作不同",
              scopeLabel: "当前表",
              options: [
                { id: "read", label: "查询", description: "只读", resolution: "查询数据" }
              ]
            }
          }
        },
        headerRows: [["类别", "归属"]]
      } as never)
      .mockResolvedValueOnce({
        kind: "answer",
        message: "已按补充说明处理。",
        resultContext: null,
        provider: "model"
      } as never);

    const container = await pendingClarification();
    await click(clarificationOptions(container)[0]);
    // 对象引用未变 ⇒ 能证明是同一本，正常继续（不能因为"没有路径"就一律误杀）。
    await until(() => runTurnMock.mock.calls.length === 2, "未保存工作簿同会话内应可继续");
    expect(textOf(container)).not.toContain("数据范围已经变化");
  });

  it("隐私：落盘的对话历史不含完整路径，也不含数据行", async () => {
    clarificationCard();
    const host = uniformHost({
      workbookName: "销售.xlsx",
      path: "C:\\Users\\someone\\Documents",
      sheetNames: ["订单", "明细"]
    });
    getWpsApplicationMock.mockReturnValue(host.application);

    await pendingClarification();

    const stored = storedHistoryTexts().join("");
    expect(stored.length).toBeGreaterThan(0);
    // 允许：工作簿名与工作表名（授权身份的可见部分）。
    expect(stored).toContain("销售.xlsx");
    expect(stored).toContain("订单");
    // 禁止：完整路径（含用户名）、稳定 ID 与任何数据行内容。
    expect(stored).not.toContain("C:\\Users\\someone");
    expect(stored).not.toContain("Documents");
    expect(stored).not.toContain("__code__");
    expect(stored).not.toContain("办公用品");
    expect(stored).not.toContain("华东");
    expect(stored).not.toContain("家具");
    expect(stored).not.toContain("华南");
  });
});
