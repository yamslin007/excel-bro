// @vitest-environment jsdom
/**
 * WPS 侧桥宿主适配器 + 共享状态机的**完整链路**。
 *
 * 只换两处边界：
 * 1. **网络**：`fetch` 是替身，但断言真实发出的**请求体**（poll / receipt）。
 * 2. **JSAPI**：宿主替身提供 `UsedRange` / `Range().Value2`，模拟真宿主的
 *    "只接受两个参数"签名限制；**读取与计算都是生产代码**
 *    （`authorizeWpsSheets` + `readAuthorizedWpsSheets` + `executeQueryTableData`）。
 */

import { afterEach, describe, expect, it, vi } from "vitest";
import {
  applyBridgeProgressEvent,
  createSdkBridgeState,
  resetTrustedSourceReads,
  runBridgeUntilIdle,
  trustedToolReads,
  verifyBridgeAuthorization,
  type SdkBridgeBinding
} from "@excel-bro/core";

import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import {
  createWpsBridgeLiveContext,
  createWpsSdkBridgeBinding,
  createWpsSdkBridgeHostPorts,
  wpsBridgeErrorCode
} from "./wpsSdkBridgeHost";
import {
  DataToolExecutionError,
  type QueryToolProgress
} from "../../excel-addin/src/dataTools";
import {
  WpsAuthorizationError,
  WpsReadLimitError,
  wpsWorkbookBridgeIdentity
} from "./wpsWorkbookReader";

const OP = "op_" + "3".repeat(32);
const FP = "c".repeat(64);

/**
 * 造一张表。``origin`` 让 UsedRange **不从 A1 起**（默认 A1）。
 *
 * 真宿主常见这种情况：数据从 B3 才开始。此时 `UsedRange.Item(r,c)` 仍是
 * **相对**已用区域的坐标，但 `WpsSheetReadResult.startRow/startColumn` 是
 * **绝对**行列 —— 按 A1 区域切片必须做这层映射，否则会整体错位。
 */
function makeSheet(
  name: string,
  values: unknown[][],
  origin: { row?: number; column?: number } = {}
): WpsWorksheet {
  const rowCount = values.length;
  const columnCount = Math.max(0, ...values.map((row) => row.length));
  const usedRange: WpsRange = {
    Row: origin.row ?? 1,
    Column: origin.column ?? 1,
    Rows: { Count: rowCount },
    Columns: { Count: columnCount },
    Item: (rowIndex: number, columnIndex = 1) => ({ Row: rowIndex, Column: columnIndex })
  };
  return {
    Name: name,
    UsedRange: usedRange,
    // 真宿主的 Worksheet.Range 只接受两个参数；多传会当场抛。
    Range: (first: WpsRange, last?: WpsRange, ...extra: unknown[]) => {
      if (extra.length > 0) throw new Error("Too many parameters.");
      const startRow = Number(first.Row ?? 1);
      const startColumn = Number(first.Column ?? 1);
      const endRow = Number(last?.Row ?? startRow);
      const endColumn = Number(last?.Column ?? startColumn);
      return {
        Value2: values
          .slice(startRow - 1, endRow)
          .map((row) => row.slice(startColumn - 1, endColumn))
      };
    }
  };
}

function makeApplication(
  sheets: WpsWorksheet[],
  overrides: { name?: string; fullName?: string | null; path?: string } = {}
): WpsApplication {
  const workbook = {
    Name: overrides.name ?? "影院.xlsx",
    // 用 `in` 而不是 `??`：`{fullName: null}` 必须真的是 null（未保存），
    // 被 `??` 换成默认值就测不出"退回 name: 命名空间"。
    FullName: "fullName" in overrides ? overrides.fullName : "C:\\Data\\影院.xlsx",
    Path: "path" in overrides ? overrides.path : "C:\\Data",
    Worksheets: {
      Count: sheets.length,
      Item: (index: number | string) => sheets[Number(index) - 1]
    },
    ActiveSheet: sheets[0]
  };
  return {
    PluginStorage: { getItem: () => null, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined,
    ActiveWorkbook: workbook,
    ActiveSheet: sheets[0]
  } as WpsApplication;
}

/** 两张授权表：同一批列，需要跨表合并后再分组汇总。 */
function cinemaApplication(
  overrides: { name?: string; fullName?: string | null; path?: string } = {}
) {
  return makeApplication(
    [
      makeSheet("Sheet1", [
        ["影院名称", "影片名称", "计数项:影厅"],
        ["影院A", "功夫女足", 11],
        ["影院B", "恐怖游轮", 2]
      ]),
      makeSheet("Sheet2", [
        ["影院名称", "影片名称", "计数项:影厅"],
        ["影院C", "功夫女足", 13],
        ["影院D", "恐怖游轮", 1]
      ])
    ],
    overrides
  );
}

const QUERY_ARGUMENTS = {
  mode: "aggregate" as const,
  scope: "selected" as const,
  groupBy: ["影片名称"],
  metrics: [{ operation: "sum" as const, field: "计数项:影厅", outputName: "影厅数" }],
  sortBy: "影厅数",
  sortDirection: "desc" as const
};

interface Captured {
  url: string;
  body: Record<string, unknown>;
}

function stubNetwork(options: { poll?: unknown; receipt?: unknown } = {}) {
  const captured: Captured[] = [];
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      const target = String(url);
      captured.push({ url: target, body: JSON.parse(String(init?.body ?? "{}")) });
      const payload = target.includes("/receipt")
        ? options.receipt ?? { status: "accepted", code: null }
        : options.poll ?? {
            calls: [
              {
                operationId: OP,
                tool: "query_table",
                arguments: QUERY_ARGUMENTS,
                requestFingerprint: FP,
                expiresAt: 0
              }
            ],
            generation: 1
          };
      return new Response(JSON.stringify(payload), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    })
  );
  return captured;
}

function bindingOf(application: WpsApplication, indexes: number[] = [1, 2]): SdkBridgeBinding {
  const live = createWpsBridgeLiveContext(application, indexes);
  return createWpsSdkBridgeBinding({
    taskId: "turn-wps-0001",
    workbookName: live.workbookName,
    workbookIdentity: live.workbookIdentity,
    sourceFingerprint: live.sourceFingerprint,
    authorizedSheetNames: live.authorizedSheetNames
  });
}

/** 把状态机推到"已有指针、需要轮询"的状态。 */
function pollingState(binding: SdkBridgeBinding) {
  return applyBridgeProgressEvent(createSdkBridgeState(binding), {
    stepId: "bridge.query_table",
    bridgeTool: { operationId: OP, tool: "query_table" }
  });
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("WPS 桥绑定与现场读取", () => {
  it("绑定固定 wps-et 宿主，身份归一成 wps-path:", () => {
    const binding = bindingOf(cinemaApplication());
    expect(binding.hostId).toBe("wps-et");
    expect(binding.workbookIdentity).toBe("wps-path:c:\\data\\影院.xlsx");
  });

  it("未保存的工作簿用会话内对象身份，绝不退回会撞车的 name:", () => {
    // 未保存 = 既没有 FullName 也没有 Path，这才真的拿不到磁盘身份。
    const application = cinemaApplication({ fullName: null, path: "" });
    const first = bindingOf(application).workbookIdentity;
    // 同一个工作簿对象：重复读取必须给出**同一个**令牌（轮询/回执才对得上）。
    expect(bindingOf(application).workbookIdentity).toBe(first);
    expect(first.startsWith("wps-session:")).toBe(true);
    // 绝不退回 `name:` —— 两份都叫「影院.xlsx」的未保存工作簿不是同一个工作簿。
    expect(first).not.toBe("name:影院.xlsx");

    // 另一个同名未保存工作簿：必须是**另一个**身份，否则两份会互相冒充。
    const other = bindingOf(cinemaApplication({ fullName: null, path: "" })).workbookIdentity;
    expect(other).not.toBe(first);
    expect(other.startsWith("wps-session:")).toBe(true);
  });

  it("身份无法核实时明确阻断（连工作簿对象都没有 → 抛错，不退回名字）", () => {
    const application = cinemaApplication();
    (application as { ActiveWorkbook: unknown }).ActiveWorkbook = undefined;
    expect(() => wpsWorkbookBridgeIdentity(application)).toThrow(WpsAuthorizationError);
    // 走完整现场读取同样 fail-closed，而不是给出一个会撞车的身份。
    expect(() => createWpsBridgeLiveContext(application, [1])).toThrow(WpsAuthorizationError);
  });

  it("现场读取带回授权表名、序号与活动表", () => {
    const live = createWpsBridgeLiveContext(cinemaApplication(), [1, 2]);
    expect(live.workbookName).toBe("影院.xlsx");
    expect(live.authorizedSheetNames).toEqual(["Sheet1", "Sheet2"]);
    expect(live.selectedSheetIndexes).toEqual([1, 2]);
    expect(live.activeWorksheet).toBe("Sheet1");
    // 结果绑定到**真实源版本**：指纹非空（取自与读取同一批授权表的内容）。
    expect(live.sourceFingerprint).not.toBeNull();
  });

  it("源内容指纹绑定真实源版本：重复读稳定、改一个已授权值就变", () => {
    const application = cinemaApplication();
    const baseline = createWpsBridgeLiveContext(application, [1, 2]).sourceFingerprint;
    expect(baseline).not.toBeNull();
    // 同一份源重复读必须稳定 —— 否则每次执行前复核都会误判成 source_changed。
    expect(createWpsBridgeLiveContext(application, [1, 2]).sourceFingerprint).toBe(baseline);

    // 换一份源（工作簿身份相同、只改一个已授权单元格的值）：指纹必须变。
    const changedApplication = makeApplication([
      makeSheet("Sheet1", [
        ["影院名称", "影片名称", "计数项:影厅"],
        ["影院A", "功夫女足", 99],
        ["影院B", "恐怖游轮", 2]
      ]),
      makeSheet("Sheet2", [
        ["影院名称", "影片名称", "计数项:影厅"],
        ["影院C", "功夫女足", 13],
        ["影院D", "恐怖游轮", 1]
      ])
    ]);
    const changed = createWpsBridgeLiveContext(changedApplication, [1, 2]).sourceFingerprint;
    expect(changed).not.toBeNull();
    expect(changed).not.toBe(baseline);
  });

  it("宿主现场读不到时抛错（调用方会 fail-closed）", () => {
    const app = makeApplication([]);
    expect(() => createWpsBridgeLiveContext(app, [1])).toThrow();
  });
});

describe("wpsBridgeErrorCode", () => {
  it("WPS 读取层的两类错误各自映射到固定错误码", () => {
    expect(wpsBridgeErrorCode(new WpsAuthorizationError("敏感细节"))).toBe(
      "unauthorized_scope"
    );
    expect(wpsBridgeErrorCode(new WpsReadLimitError("敏感细节"))).toBe("result_too_large");
  });

  it("确定性查询错误沿用与 Excel 同一份映射", () => {
    expect(wpsBridgeErrorCode(new DataToolExecutionError("SCOPE_VIOLATION", "", false))).toBe(
      "unauthorized_scope"
    );
    expect(
      wpsBridgeErrorCode(new DataToolExecutionError("TOOL_LIMIT_EXCEEDED", "", false))
    ).toBe("result_too_large");
    expect(wpsBridgeErrorCode(new DataToolExecutionError("FIELD_NOT_FOUND", "", false))).toBe(
      "invalid_arguments"
    );
  });

  it("未知异常一律 internal —— 绝不把异常文本当作错误码", () => {
    expect(wpsBridgeErrorCode(new Error("内部细节"))).toBe("internal");
    expect(wpsBridgeErrorCode("随便什么")).toBe("internal");
  });
});

describe("共享复核判据（WPS 现场）", () => {
  it("授权未变放行；收窄则拒（判据为两端共用）", () => {
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const call = {
      operationId: OP,
      tool: "query_table" as const,
      arguments: QUERY_ARGUMENTS,
      requestFingerprint: FP,
      expiresAt: 0
    };
    expect(
      verifyBridgeAuthorization(binding, call, createWpsBridgeLiveContext(application, [1, 2]))
    ).toEqual({ ok: true });
    expect(
      verifyBridgeAuthorization(binding, call, createWpsBridgeLiveContext(application, [1]))
    ).toEqual({ ok: false, errorCode: "unauthorized_scope" });
  });
});

describe("端到端：轮询 → 复核 → 真读真算 → 回执", () => {
  it("真实请求体 + 真实跨表分组求和", async () => {
    const captured = stubNetwork();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const progress: QueryToolProgress[] = [];
    const ports = createWpsSdkBridgeHostPorts(
      {
        liveContext: () => createWpsBridgeLiveContext(application, [1, 2]),
        onProgress: (item) => progress.push(item)
      },
      binding,
      { application: () => application }
    );

    const result = await runBridgeUntilIdle(pollingState(binding), ports);

    const poll = captured.find((item) => item.url.includes("/poll"))?.body;
    expect(poll).toMatchObject({
      taskId: "turn-wps-0001",
      hostId: "wps-et",
      workbookName: "影院.xlsx",
      workbookIdentity: "wps-path:c:\\data\\影院.xlsx"
    });

    const receipt = captured.find((item) => item.url.includes("/receipt"))?.body;
    expect(receipt).toMatchObject({
      taskId: "turn-wps-0001",
      hostId: "wps-et",
      operationId: OP,
      requestFingerprint: FP
    });
    // 值是生产计算函数从真实宿主矩阵算出来的，不是预置的。
    const payload = (
      receipt?.receipt as { result?: { headers: unknown[]; rows: unknown[][]; calculation: string } }
    ).result;
    expect(payload?.headers).toEqual(["影片名称", "影厅数"]);
    // 指标之和是 27，但源记录只有 4 条；总数不能从显示指标相加。
    expect(payload?.calculation).toContain("筛选后总记录数：4");
    expect(payload?.rows).toEqual([
      ["功夫女足", 24],
      ["恐怖游轮", 3]
    ]);

    expect(result.executed).toBe(1);
    expect(progress.length).toBeGreaterThan(0);
    // 两端进度形状统一成 `QueryToolProgress`（scannedRows/sheet）。
    expect(progress.at(-1)).toMatchObject({
      sheet: "Sheet2",
      scannedRows: 3,
      totalRows: 3
    });
  });

  it("执行前授权收窄：一个单元格都不读，回执只带固定错误码", async () => {
    const captured = stubNetwork();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    let liveIndexes = [1, 2];
    const ports = createWpsSdkBridgeHostPorts(
      {
        // 绑定发生在授权完整时；执行前用户撤销了 Sheet2。
        liveContext: () => createWpsBridgeLiveContext(application, liveIndexes)
      },
      binding,
      { application: () => application }
    );
    liveIndexes = [1];

    const result = await runBridgeUntilIdle(pollingState(binding), ports);

    expect(result.executed).toBe(0);
    expect(captured.find((item) => item.url.includes("/receipt"))?.body.receipt).toEqual({
      errorCode: "unauthorized_scope"
    });
  });

  it("换簿（同名不同路径）：拒读，错误码 source_changed", async () => {
    const captured = stubNetwork();
    const original = cinemaApplication();
    const binding = bindingOf(original);
    const other = cinemaApplication({ fullName: "D:\\Other\\影院.xlsx" });
    const ports = createWpsSdkBridgeHostPorts(
      {
        // 复核读的是"当前"宿主 —— 已经是另一个同名工作簿。
        liveContext: () => createWpsBridgeLiveContext(other, [1, 2])
      },
      binding,
      { application: () => other }
    );

    const result = await runBridgeUntilIdle(pollingState(binding), ports);

    expect(result.executed).toBe(0);
    expect(captured.find((item) => item.url.includes("/receipt"))?.body.receipt).toEqual({
      errorCode: "source_changed"
    });
  });

  it("宿主读取失败时回执只带固定错误码，异常文本不出站", async () => {
    const captured = stubNetwork();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const ports = createWpsSdkBridgeHostPorts(
      {
        liveContext: () => createWpsBridgeLiveContext(application, [1, 2])
      },
      binding,
      {
        application: () => application,
        readSheets: async () => {
          throw new WpsAuthorizationError("内部路径 D:\\Secret\\x.xlsx");
        }
      }
    );

    const result = await runBridgeUntilIdle(pollingState(binding), ports);

    expect(result.executed).toBe(1); // 走了 execute 分支，但产出是错误回执
    const body = captured.find((item) => item.url.includes("/receipt"))?.body;
    expect(body?.receipt).toEqual({ errorCode: "unauthorized_scope" });
    expect(JSON.stringify(body)).not.toContain("Secret");
  });

  /*
   * "身份无法核实"必须**明确阻断**，而且要阻断成一张**回执**。
   *
   * 只抛异常是不够的：抛出去会被上层的 drain 吞掉（"桥失败不能拖垮整轮规划"），
   * 服务端一个回执都收不到，只能干等到 `waitToolTimeout` —— 无人回执的等待。
   */
  it("现场读不到（身份核实不了）时也回一张固定错误码回执，而不是静默无回执", async () => {
    const captured = stubNetwork();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const ports = createWpsSdkBridgeHostPorts(
      {
        liveContext: () => {
          throw new WpsAuthorizationError(
            "无法读取当前工作簿身份，不能建立本机数据桥。"
          );
        }
      },
      binding,
      { application: () => application }
    );

    const result = await runBridgeUntilIdle(pollingState(binding), ports);

    expect(result.executed).toBe(0);
    const receipt = captured.find((item) => item.url.includes("/receipt"))?.body;
    expect(receipt?.receipt).toEqual({ errorCode: "unauthorized_scope" });
  });

  /* 执行阶段才读不到现场（直接调用 execute 的路径）也不能让异常冒出去。 */
  it("execute 阶段现场读不到时同样只回固定错误码", async () => {
    const captured = stubNetwork();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    let readable = true;
    const ports = createWpsSdkBridgeHostPorts(
      {
        liveContext: () => {
          if (!readable) throw new WpsAuthorizationError("身份核实不了");
          return createWpsBridgeLiveContext(application, [1, 2]);
        }
      },
      binding,
      { application: () => application }
    );

    readable = false;
    const outcome = await ports.execute(binding, {
      operationId: OP,
      tool: "query_table",
      arguments: QUERY_ARGUMENTS,
      requestFingerprint: FP,
      expiresAt: 0
    });
    expect(outcome).toEqual({ kind: "error", errorCode: "unauthorized_scope" });
    expect(captured).toHaveLength(0);
  });

  it("取消后不再轮询，也不读宿主", async () => {
    const captured = stubNetwork();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const ports = createWpsSdkBridgeHostPorts(
      { liveContext: () => createWpsBridgeLiveContext(application, [1, 2]) },
      binding,
      { application: () => application }
    );

    const cancelled = applyBridgeProgressEvent(pollingState(binding), {
      stepId: "bridge.query_table",
      status: "cancelled"
    });
    const result = await runBridgeUntilIdle(cancelled, ports);

    expect(result.polled).toBe(false);
    expect(result.executed).toBe(0);
    expect(captured).toHaveLength(0);
  });

  it("prepare_python_data 按来源区域切片原始值并回 datasets 回执", async () => {
    const captured = stubNetwork();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const ports = createWpsSdkBridgeHostPorts(
      { liveContext: () => createWpsBridgeLiveContext(application, [1, 2]) },
      binding,
      { application: () => application }
    );

    const outcome = await ports.execute(binding, {
      operationId: OP,
      tool: "prepare_python_data",
      arguments: {
        sources: [
          { sheet: "Sheet1", range: "A1:C3" },
          { sheet: "Sheet2", range: "B1:B3" }
        ]
      },
      requestFingerprint: FP,
      expiresAt: 0
    });
    expect(outcome.kind).toBe("result");
    const payload = (outcome as { kind: "result"; result: { datasets: unknown[] } })
      .result;
    expect(payload.datasets).toHaveLength(2);
    const datasets = payload.datasets as { sourceIndex: number; values: unknown[][] }[];
    expect(datasets[0].sourceIndex).toBe(0);
    expect(datasets[0].values).toEqual([
      ["影院名称", "影片名称", "计数项:影厅"],
      ["影院A", "功夫女足", 11],
      ["影院B", "恐怖游轮", 2]
    ]);
    // Sheet2 的 B1:B3 切片只有 1 列。
    expect(datasets[1].sourceIndex).toBe(1);
    expect(datasets[1].values).toEqual([
      ["影片名称"],
      ["功夫女足"],
      ["恐怖游轮"]
    ]);

    await ports.receipt(binding, {
      operationId: OP,
      tool: "prepare_python_data",
      arguments: { sources: [] },
      requestFingerprint: FP,
      expiresAt: 0
    }, outcome);
    const receipt = captured.find((item) => item.url.includes("/receipt"))?.body;
    expect(receipt?.receipt).toHaveProperty("datasets");
    expect((receipt?.receipt as { result?: unknown }).result).toBeUndefined();
  });

  /*
   * 「宿主成功读过的范围」——WPS 侧同样要记账，且与 Excel 侧**逐字同源**
   * （`@excel-bro/core/planSourceDependency`）。漏掉这一条，WPS 轮次的
   * "实际来源依赖"恒为空，同一份计划在 Excel 可恢复、在 WPS 不可恢复。
   */
  it("prepare_python_data 成功读取后按实际来源记账", async () => {
    stubNetwork();
    resetTrustedSourceReads();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const ports = createWpsSdkBridgeHostPorts(
      { liveContext: () => createWpsBridgeLiveContext(application, [1, 2]) },
      binding,
      { application: () => application }
    );

    const outcome = await ports.execute(binding, {
      operationId: OP,
      tool: "prepare_python_data",
      arguments: {
        sources: [
          { sheet: "Sheet1", range: "A1:C3" },
          { sheet: "Sheet2", range: "B1:B3" }
        ]
      },
      requestFingerprint: FP,
      expiresAt: 0
    });
    expect(outcome.kind).toBe("result");
    expect(trustedToolReads(binding.taskId)).toEqual({
      origin: "tool-execution-record",
      sheets: ["Sheet1", "Sheet2"]
    });
  });

  /* 只把输入的一部分交给执行器时，记账只有**实际读过**的那张。 */
  it("prepare_python_data 只记账实际读取的来源", async () => {
    stubNetwork();
    resetTrustedSourceReads();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const ports = createWpsSdkBridgeHostPorts(
      { liveContext: () => createWpsBridgeLiveContext(application, [1, 2]) },
      binding,
      { application: () => application }
    );

    await ports.execute(binding, {
      operationId: OP,
      tool: "prepare_python_data",
      arguments: { sources: [{ sheet: "Sheet2", range: "B1:B3" }] },
      requestFingerprint: FP,
      expiresAt: 0
    });
    expect(trustedToolReads(binding.taskId)?.sheets).toEqual(["Sheet2"]);
  });

  /* 越权来源被拒时**不留**记账：没读到就不算"读过"。 */
  it("prepare_python_data 越权被拒时不留记账", async () => {
    stubNetwork();
    resetTrustedSourceReads();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const ports = createWpsSdkBridgeHostPorts(
      { liveContext: () => createWpsBridgeLiveContext(application, [1, 2]) },
      binding,
      { application: () => application }
    );

    const outcome = await ports.execute(binding, {
      operationId: OP,
      tool: "prepare_python_data",
      arguments: { sources: [{ sheet: "Sheet9", range: "A1:B2" }] },
      requestFingerprint: FP,
      expiresAt: 0
    });
    expect(outcome).toEqual({ kind: "error", errorCode: "unauthorized_scope" });
    expect(trustedToolReads(binding.taskId)).toBeNull();
  });

  /* query_table 也要记账：WPS 侧此前**完全没有**这一层。 */
  it("query_table 按实际取用的工作表记账", async () => {
    stubNetwork();
    resetTrustedSourceReads();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const ports = createWpsSdkBridgeHostPorts(
      { liveContext: () => createWpsBridgeLiveContext(application, [1, 2]) },
      binding,
      { application: () => application }
    );

    await ports.execute(binding, {
      operationId: OP,
      tool: "query_table",
      arguments: QUERY_ARGUMENTS,
      requestFingerprint: FP,
      expiresAt: 0
    });
    const record = trustedToolReads(binding.taskId);
    expect(record?.origin).toBe("tool-execution-record");
    expect(record?.sheets.length).toBeGreaterThan(0);
  });


  describe("prepare_python_data 非 A1 起点的已用区域", () => {
    /** 已用区域从 B3 起：矩阵 4×4 对应绝对 B3:E6。 */
    const OFFSETS = [
      ["B3", "C3", "D3", "E3"],
      ["B4", "C4", "D4", "E4"],
      ["B5", "C5", "D5", "E5"],
      ["B6", "C6", "D6", "E6"]
    ];

    function offsetApplication() {
      return makeApplication([makeSheet("Sheet1", OFFSETS, { row: 3, column: 2 })]);
    }

    function offsetPorts() {
      const application = offsetApplication();
      const binding = createWpsSdkBridgeBinding({
        taskId: "turn-wps-offset",
        workbookName: "影院.xlsx",
        workbookIdentity: "wps-path:c:\\data\\影院.xlsx",
        sourceFingerprint: null,
        authorizedSheetNames: ["Sheet1"]
      });
      return {
        binding,
        ports: createWpsSdkBridgeHostPorts(
          { liveContext: () => createWpsBridgeLiveContext(application, [1]) },
          binding,
          { application: () => application }
        )
      };
    }

    async function slice(range: string) {
      const { binding, ports } = offsetPorts();
      const outcome = await ports.execute(binding, {
        operationId: OP,
        tool: "prepare_python_data" as const,
        arguments: { sources: [{ sheet: "Sheet1", range }] },
        requestFingerprint: FP,
        expiresAt: 0
      });
      expect(outcome.kind).toBe("result");
      const datasets = (
        outcome as { kind: "result"; result: { datasets: { values: unknown[][] }[] } }
      ).result.datasets;
      return datasets[0].values;
    }

    it("整体已用区域：B3:E6 返回整块矩阵", async () => {
      expect(await slice("B3:E6")).toEqual(OFFSETS);
    });

    it("子区域按绝对坐标取：C4:D5 只取中间 2×2", async () => {
      expect(await slice("C4:D5")).toEqual([
        ["C4", "D4"],
        ["C5", "D5"]
      ]);
    });

    it("边界行/列：B3:B3 只取左上角一格", async () => {
      expect(await slice("B3:B3")).toEqual([["B3"]]);
    });

    it("右下边界：E6:E6 只取右下角一格", async () => {
      expect(await slice("E6:E6")).toEqual([["E6"]]);
    });

    /* 越界必须拒绝 —— 绝不把"没读到的部分"补成空白冒充读取成功。 */
    it("区域起点早于已用区域（A1）时拒绝，不补空白", async () => {
      const { binding, ports } = offsetPorts();
      const outcome = await ports.execute(binding, {
        operationId: OP,
        tool: "prepare_python_data" as const,
        arguments: { sources: [{ sheet: "Sheet1", range: "A1:C4" }] },
        requestFingerprint: FP,
        expiresAt: 0
      });
      expect(outcome).toEqual({ kind: "error", errorCode: "unauthorized_scope" });
    });

    it("区域超出已用区域右下角时拒绝，不补空白", async () => {
      const { binding, ports } = offsetPorts();
      const outcome = await ports.execute(binding, {
        operationId: OP,
        tool: "prepare_python_data" as const,
        arguments: { sources: [{ sheet: "Sheet1", range: "D5:G9" }] },
        requestFingerprint: FP,
        expiresAt: 0
      });
      expect(outcome).toEqual({ kind: "error", errorCode: "unauthorized_scope" });
    });

    /* 完全落在已用区域之外（A1:A2）同样拒绝。 */
    it("区域整体落在已用区域之外时拒绝", async () => {
      const { binding, ports } = offsetPorts();
      const outcome = await ports.execute(binding, {
        operationId: OP,
        tool: "prepare_python_data" as const,
        arguments: { sources: [{ sheet: "Sheet1", range: "A1:A2" }] },
        requestFingerprint: FP,
        expiresAt: 0
      });
      expect(outcome).toEqual({ kind: "error", errorCode: "unauthorized_scope" });
    });
  });

  it("prepare_python_data 来源越权时回固定错误码", async () => {
    const captured = stubNetwork();
    const application = cinemaApplication();
    const binding = bindingOf(application);
    const ports = createWpsSdkBridgeHostPorts(
      { liveContext: () => createWpsBridgeLiveContext(application, [1, 2]) },
      binding,
      { application: () => application }
    );

    const outcome = await ports.execute(binding, {
      operationId: OP,
      tool: "prepare_python_data",
      arguments: { sources: [{ sheet: "Sheet9", range: "A1:B2" }] },
      requestFingerprint: FP,
      expiresAt: 0
    });
    expect(outcome).toEqual({ kind: "error", errorCode: "unauthorized_scope" });
  });
});
