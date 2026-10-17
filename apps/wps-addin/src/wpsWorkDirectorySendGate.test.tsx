/**
 * 定向测试：**工作目录范围下的发送闸**（真实 `WpsAgentPanel`，不换替身）。
 *
 * 为什么必须有这道闸：工作目录的数据在**服务端**（openpyxl），跟当前打开的
 * 工作簿毫无关系。本轮只打通了"勾选与授权"，服务端查询/执行链路还没接到 WPS 面板。
 * 这时候若放行，下游只能拿"当前打开的工作簿"当数据源 —— 那不是用户选的东西，
 * 是用一个他没授权的对象冒充他要的数据。
 *
 * 于是这条用例钉四件事，**各自读不同的数据源**：
 *  1. 零出站：`/api/turn`、`/api/intent/message` 一个请求都没有（读 fetch 记录）；
 *  2. 零工作簿读取：发送这一下宿主的值读取计数**没有增加**、写入为 0（读宿主替身计数器）；
 *  3. 诚实提示在场，并带上是哪个工作目录、选了几张表（读界面文本）；
 *  4. 输入框里的字原样留着（读 textarea.value）。
 *
 * 负向断言配**正向锚点**：同一套替身换成工作簿来源时必须真的发出去 ——
 * 否则"零出站"可能只是因为整条链路压根没接上（那样这条判据就是恒真的）。
 *
 * 判据与文案的唯一来源是共享层的 `workDirectorySendRefusal`（`@excel-bro/core/folderScope`），
 * 端到端这里只证"接线在"；决定本身由 `packages/core/src/folderScope.test.ts` 逐条钉住。
 */
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getModelSettings, listModels } from "../../excel-addin/src/api";
import WpsAgentPanel from "./WpsAgentPanel";
import { getWpsApplication } from "./wpsHost";

vi.mock("../../excel-addin/src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../excel-addin/src/api")>();
  return {
    ...actual,
    listModels: vi.fn(),
    getModelSettings: vi.fn()
  };
});
vi.mock("./wpsHost", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./wpsHost")>();
  return { ...actual, getWpsApplication: vi.fn(() => ({})) };
});

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
if (!Element.prototype.scrollIntoView) {
  Element.prototype.scrollIntoView = () => undefined;
}

const getWpsApplicationMock = vi.mocked(getWpsApplication);
const listModelsMock = vi.mocked(listModels);
const getModelSettingsMock = vi.mocked(getModelSettings);

const model = {
  id: "connection:first",
  label: "第一个连接",
  provider: "model" as const,
  available: true,
  supportsVision: false
};

type Captured = { url: string; body: string };
let captured: Captured[] = [];
let counts = { valueReads: 0, writes: 0 };

function jsonResponse(payload: unknown): Response {
  return {
    ok: true,
    status: 200,
    json: async () => payload,
    text: async () => JSON.stringify(payload)
  } as unknown as Response;
}

/**
 * 一个会**数数**的 WPS 宿主：读一次 `Value2` 就记一笔，写一次也记一笔。
 *
 * 这样"这一下有没有碰当前工作簿"是**量出来的**，而不是从"没发出请求"推断的。
 */
function countingHost() {
  const cellRange = () => {
    const range: Record<string, unknown> = {
      Rows: { Count: 3 },
      Columns: { Count: 2 },
      MergeCells: false,
      Row: 1,
      Column: 1
    };
    Object.defineProperties(range, {
      Value2: {
        get: () => {
          counts.valueReads += 1;
          return [["类别", "归属"], ["甲", "华东"], ["乙", "华南"]];
        },
        set: () => {
          counts.writes += 1;
        }
      },
      Item: { value: () => range }
    });
    return range;
  };
  const sheets = ["订单", "明细"].map((Name, position) => ({
    Name,
    Index: position + 1,
    CodeName: `SheetCode${position + 1}`,
    ProtectContents: false,
    get UsedRange() {
      return cellRange();
    },
    Range: () => cellRange()
  }));
  const workbook = {
    Name: "测试.xlsx",
    FullName: "C:\\数据\\测试.xlsx",
    Path: "C:\\数据",
    ReadOnly: false,
    Worksheets: { Count: sheets.length, Item: (index: number) => sheets[Number(index) - 1] },
    ActiveSheet: sheets[0]
  };
  const storage = new Map<string, unknown>();
  return {
    Version: "12.0",
    ActiveWorkbook: workbook,
    ActiveSheet: sheets[0],
    PluginStorage: {
      getItem: (key: string) => storage.get(key),
      setItem: (key: string, item: unknown) => {
        storage.set(key, item);
      }
    },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined
  } as never;
}

let mounted: Array<{ unmount: () => void }> = [];

async function settle(): Promise<void> {
  await act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  });
}

async function until(check: () => boolean, message: string): Promise<void> {
  for (let index = 0; index < 80; index += 1) {
    if (check()) return;
    await settle();
  }
  throw new Error(message);
}

async function mount(element: React.ReactElement) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(element);
  });
  mounted.push({
    unmount: () => {
      act(() => root.unmount());
      container.remove();
    }
  });
  return { container };
}

async function type(element: HTMLTextAreaElement, value: string): Promise<void> {
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")?.set;
  await act(async () => {
    setter?.call(element, value);
    element.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

async function submit(element: HTMLTextAreaElement): Promise<void> {
  const form = element.closest("form");
  if (!form) throw new Error("没有找到输入区表单");
  await act(async () => {
    form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
  });
  await settle();
}

function textOf(container: HTMLElement): string {
  return container.textContent ?? "";
}

/** 等模型目录就绪：没加载完的话，发送会先被"没有可用模型"那条闸拦下。 */
async function mountReadyPanel(props: {
  scopeSource?: "workbook" | "folder";
  folderScope?: { folderName: string; sheetCount: number } | null;
}) {
  const { container } = await mount(
    <WpsAgentPanel selectedSheetIndexes={[1]} scopeDock={<div className="eb-scope-dock" />} {...props} />
  );
  await until(
    () =>
      container
        .querySelector<HTMLButtonElement>(".model-picker-trigger")
        ?.title.includes("当前模型：第一个连接") ?? false,
    "应加载模型"
  );
  return { container };
}

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
  captured = [];
  counts = { valueReads: 0, writes: 0 };
  listModelsMock.mockResolvedValue({ defaultModelId: model.id, models: [model] });
  getModelSettingsMock.mockResolvedValue({ connectionCount: 1 } as never);
  getWpsApplicationMock.mockReturnValue(countingHost());
  // 记录每一个真实出站请求。放行的路径会落在这里；被拦下的路径不该出现任何一条。
  vi.stubGlobal(
    "fetch",
    vi.fn(async (input: RequestInfo | URL) => {
      captured.push({ url: typeof input === "string" ? input : String(input), body: "" });
      return jsonResponse({ kind: "needs_workbook_context", summary: "需要字段结构", provider: "model" });
    })
  );
});

afterEach(() => {
  vi.unstubAllGlobals();
  act(() => {
    mounted.forEach((entry) => entry.unmount());
  });
  mounted = [];
  document.body.innerHTML = "";
});

const TURN_URLS = /\/api\/(turn|intent)\b/;

describe("WPS 工作目录：发送侧不偷用当前工作簿", () => {
  it("工作目录范围下：零出站、零工作簿读取，给出诚实提示并把输入留回框里", async () => {
    const { container } = await mountReadyPanel({
      scopeSource: "folder",
      folderScope: { folderName: "数据", sheetCount: 3 }
    });
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea")!;
    expect(textarea).not.toBeNull();
    const readsBeforeTurn = counts.valueReads;

    const prompt = "把目录里的销售表按月份汇总";
    await type(textarea, prompt);
    await submit(textarea);

    // ① 零出站：消息没发出去。
    expect(captured.filter((item) => TURN_URLS.test(item.url))).toEqual([]);
    // ② 零工作簿读取/写入：发送这一下根本没碰当前打开的活页本。
    expect(counts.valueReads).toBe(readsBeforeTurn);
    expect(counts.writes).toBe(0);
    // ③ 诚实提示：说清是哪个目录、几张表，以及"什么都没读"。
    expect(textOf(container)).toContain("工作目录「数据」（已选 3 张表）");
    expect(textOf(container)).toContain("没有读取任何文件");
    expect(textOf(container)).toContain("当前打开的工作簿没有被读取或修改");
    // ④ 输入框里的字原样保留，切个范围就能直接发。
    expect(textarea.value).toBe(prompt);
  });

  it("工作簿来源仍要能正常发出去（正向锚点：证明「零出站」不是整条链路没接上）", async () => {
    const { container } = await mountReadyPanel({ scopeSource: "workbook" });
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea")!;
    await type(textarea, "把类别列翻译成归属");
    await submit(textarea);
    // 真的发出去了：至少一条业务请求，并且没有出现那道拒绝文案。
    expect(captured.length).toBeGreaterThan(0);
    expect(textOf(container)).not.toContain("工作目录的读取通道本轮还没接到这个面板");
  });

  it("没传来源时按工作簿处理，不受这道闸影响", async () => {
    const { container } = await mountReadyPanel({});
    const textarea = container.querySelector<HTMLTextAreaElement>(".composer textarea")!;
    await type(textarea, "把类别列翻译成归属");
    await submit(textarea);
    expect(textOf(container)).not.toContain("工作目录的读取通道本轮还没接到这个面板");
    expect(captured.length).toBeGreaterThan(0);
  });
});
