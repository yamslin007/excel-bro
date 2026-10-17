/**
 * WPS（JSAPI）侧的桥宿主适配器。
 *
 * 与 Excel 侧 `sdkBridgeHost.ts` 是**同一套语义**（复核判据在共享核心里，
 * 轮询/回执/取值都走同一份 `api.ts`），差别只在"怎么读宿主"：
 *
 * - Excel 走 Office.js `Excel.run` + `usedRange`；
 * - WPS 走 JSAPI `authorizeWpsSheets` + `readAuthorizedWpsSheets`。
 *
 * 计算那一步两端**完全相同**：都是既有的 `executeQueryTableData`。
 *
 * 边界与 Excel 侧一致：只读；授权复核不过就一个宿主单元格都不读；
 * 失败只回固定错误码，绝不把异常文本透给模型。
 */

import {
  BRIDGE_UNVERIFIABLE_ERROR_CODE,
  recordTrustedToolReads,
  verifyBridgeAuthorizationOrRefuse,
  type BridgeCallOutcome,
  type BridgeHostId,
  type BridgeHostPorts,
  type SdkBridgeBinding
} from "@excel-bro/core";

import {
  fetchSdkBridgeResult,
  pollSdkBridgeTool,
  postSdkBridgeReceipt
} from "../../excel-addin/src/api";
import type {
  DataToolRequest,
  QueryTableArguments,
  SdkBridgeDatasetReceipt,
  SdkBridgeResultResponse,
  SdkBridgeToolReceipt
} from "../../excel-addin/src/contracts";
import {
  DataToolExecutionError,
  executeQueryTableData,
  type PreparePythonDataSource,
  type QueryToolProgress
} from "../../excel-addin/src/dataTools";
import {
  bridgeErrorCode,
  createSdkBridgeBinding,
  type BridgeResultLookup,
  type SdkBridgeBindingOptions
} from "../../excel-addin/src/sdkBridgeHost";
import {
  normalizeCellValue,
  PythonPrepareRegionRefused,
  slicePythonPrepareRegion,
  type PythonPrepareRect
} from "@excel-bro/core/pythonPrepare";

import { getWpsApplication, type WpsApplication } from "./wpsHost";
import { sheetsForRequest } from "./wpsReadonlyAgent";
import {
  authorizeWpsSheets,
  listWpsWorksheets,
  readAuthorizedWpsSheets,
  WpsAuthorizationError,
  WpsReadLimitError,
  wpsAuthorizedSourceFingerprint,
  wpsWorkbookBridgeIdentity,
  type WpsCellValue,
  type WpsSheetReadResult
} from "./wpsWorkbookReader";

/** 执行前重读的 WPS 宿主现场。 */
export interface WpsSdkBridgeLiveContext {
  /** 工作簿显示名（建绑定用）。 */
  workbookName: string;
  workbookIdentity: string;
  sourceFingerprint: string | null;
  /** 复核用：当前授权的工作表**名字**。 */
  authorizedSheetNames: string[];
  /** 读取用：当前授权的工作表**序号**（JSAPI 按序号定位）。 */
  selectedSheetIndexes: number[];
  activeWorksheet: string;
}

export type WpsSdkBridgeBindingOptions = Omit<SdkBridgeBindingOptions, "hostId">;

/**
 * 读一次**当前**宿主现场。
 *
 * 绑定快照与执行前复核都用这一份 —— 两个读取器分开写，迟早会对不上
 * （一族"收窄测不出来"的缺陷就是这么来的）。
 *
 * 授权读不到就**抛**：调用方（hook 的 begin）会退回空范围，复核一律拒绝。
 * 这是 fail-closed 的方向 —— 绝不拿"读不到"当成"没变化"。
 */
export function createWpsBridgeLiveContext(
  application: WpsApplication,
  selectedSheetIndexes: readonly number[]
): WpsSdkBridgeLiveContext {
  const authorization = authorizeWpsSheets(application, selectedSheetIndexes);
  const metadata = listWpsWorksheets(application);
  return {
    workbookName: metadata.name,
    workbookIdentity: wpsWorkbookBridgeIdentity(application),
    // 结果与**真实源版本**绑定：指纹取自与读取同一批授权工作表的内容。
    // 绑定与执行前复核各自重读一次，等待期间源变了就会被判成 `source_changed`。
    sourceFingerprint: wpsAuthorizedSourceFingerprint(application, selectedSheetIndexes),
    authorizedSheetNames: authorization.sheets.map((sheet) => sheet.name),
    selectedSheetIndexes: authorization.sheets.map((sheet) => sheet.index),
    activeWorksheet: metadata.sheets.find((sheet) => sheet.isActive)?.name ?? ""
  };
}

/** 建绑定：hostId 固定为 `wps-et`，调用方不必也不该自己填。 */
export function createWpsSdkBridgeBinding(
  options: WpsSdkBridgeBindingOptions
): SdkBridgeBinding {
  return createSdkBridgeBinding({ ...options, hostId: "wps-et" satisfies BridgeHostId });
}

/**
 * WPS 侧错误 → 服务端固定错误码。
 *
 * 先认 WPS 读取层自己的两类错误，其余沿用与 Excel 同一份映射
 * （`DataToolExecutionError` / AbortError）。**任何**异常文本都不外发。
 */
export function wpsBridgeErrorCode(reason: unknown): string {
  if (reason instanceof WpsAuthorizationError) return "unauthorized_scope";
  if (reason instanceof WpsReadLimitError) return "result_too_large";
  return bridgeErrorCode(reason);
}

export interface WpsSdkBridgeExecutionOptions {
  /**
   * 每次执行前调用一次，**重新读**当前授权与工作簿身份（不是启动时快照）。
   * 允许异步：读取宿主本身就要 await。
   */
  liveContext: () =>
    | WpsSdkBridgeLiveContext
    | Promise<WpsSdkBridgeLiveContext>;
  signal?: AbortSignal;
  onProgress?: (progress: QueryToolProgress) => void;
}

export interface WpsSdkBridgeHostDependencies {
  poll?: typeof pollSdkBridgeTool;
  receipt?: typeof postSdkBridgeReceipt;
  fetchResult?: typeof fetchSdkBridgeResult;
  readSheets?: typeof readAuthorizedWpsSheets;
  authorize?: typeof authorizeWpsSheets;
  /** 只替换宿主入口（测试用）；默认 `getWpsApplication()`。 */
  application?: () => WpsApplication;
}

/**
 * 组装共享状态机需要的四个端口。
 *
 * ``execute`` 的顺序刻意与本地分析一致：
 * 复核（在状态机里）→ 授权（重读宿主元数据）→ 真读矩阵 → 既有确定性计算。
 * 中间任何一步失败都只回固定错误码。
 */
export function createWpsSdkBridgeHostPorts(
  options: WpsSdkBridgeExecutionOptions,
  binding: SdkBridgeBinding,
  dependencies: WpsSdkBridgeHostDependencies = {}
): BridgeHostPorts {
  const poll = dependencies.poll ?? pollSdkBridgeTool;
  const postReceipt = dependencies.receipt ?? postSdkBridgeReceipt;
  const authorize = dependencies.authorize ?? authorizeWpsSheets;
  const readSheets = dependencies.readSheets ?? readAuthorizedWpsSheets;
  const application = dependencies.application ?? (() => getWpsApplication());

  return {
    poll: (current) => poll(current, options.signal),

    /*
     * 执行前的授权复核：判据在共享核心里，与 Excel 端**同一份**。
     * 工作簿身份 / 来源指纹 / 授权范围（含**收窄**）任一变化都必须拒绝。
     * 现场在这一刻重读 —— 拿启动快照复核等于没复核。
     *
     * 读不到现场（未保存工作簿身份核实不了、授权序号失效…）走共享的
     * fail-closed 拒绝：**必须回一张固定错误码的回执**，不能让异常冒出去被
     * 上层吞掉 —— 那样服务端一个回执都收不到，只能干等到超时。
     */
    verify: (_current, call) =>
      verifyBridgeAuthorizationOrRefuse(binding, call, () => options.liveContext()),

    execute: async (_current, call): Promise<BridgeCallOutcome> => {
      let live: WpsSdkBridgeLiveContext;
      try {
        /*
         * 现场读取单独接住：读不到 = **无法核实**（未保存工作簿身份核实不了、
         * 已授权序号失效…），回与复核同一个固定错误码；混成 `internal`
         * 用户拿不到可执行的下一步。`verify` 正常会先拦住，这里是防御性出口。
         */
        live = await options.liveContext();
      } catch {
        return { kind: "error", errorCode: BRIDGE_UNVERIFIABLE_ERROR_CODE };
      }
      try {
        if (call.tool === "prepare_python_data") {
          const sources = asWpsPreparePythonSources(call.arguments);
          const authorization = authorize(application(), live.selectedSheetIndexes);
          const reads = await readSheets(application(), authorization, {
            signal: options.signal,
            /*
             * 数据登记必须取回 `Text`/`NumberFormat`：值矩阵里前导零编号退化成
             * 数字、日期退化成序列值，不还原就和 Excel 端（走 `normalizeCellValue`）
             * 不是一个语义 —— 同一张表两端算出不同结果是不可接受的。
             */
            includeDisplay: true
          });
          const blocks: SdkBridgeDatasetReceipt[] = sliceWpsPythonSources(
            sources,
            reads,
            live.authorizedSheetNames
          );
          /*
           * 可信来源记账（与 Excel 侧逐字同源）：数据登记也是"确定性地读过表"，
           * 读了谁由**实际读取成功的来源**回报，不是模型自报的参数集合。失败路径
           * 会在下面 `catch` 里变成固定错误码，此处只记成功 —— "没读到" ≠ "读过"。
           */
          recordTrustedToolReads(
            binding.taskId,
            blocks.map((block) => sources[block.sourceIndex]?.sheet ?? "").filter(Boolean)
          );
          return { kind: "result", result: { datasets: blocks } };
        }
        // 授权现场**此刻**重读：不能用绑定时的快照去读现在的表。
        const authorization = authorize(application(), live.selectedSheetIndexes);
        const reads = await readSheets(application(), authorization, {
          signal: options.signal,
          // 两端进度形状统一成 `QueryToolProgress`，面板就不用按宿主分叉。
          onProgress: options.onProgress
            ? ({ sheet, rowsRead, totalRows }) =>
                options.onProgress?.({
                  sheet: sheet.name,
                  scannedRows: rowsRead,
                  totalRows
                })
            : undefined
        });
        const request: DataToolRequest = {
          // 用桥上的操作 ID 作为请求身份，便于把本地结果与那次调用对上。
          id: call.operationId,
          tool: "query_table",
          arguments: asQueryTableArguments(call.arguments)
        };
        const sheets = sheetsForRequest(request, reads, live.activeWorksheet);
        /*
         * 可信来源记账（与 Excel 侧逐字同源）：读了哪几张表由**确定性执行器**
         * 实际取用的那些表决定（`sheetsForRequest` 已按 `scope` 过滤），不是模型
         * 自报的授权范围。断点恢复的"实际来源依赖"只从这条记录来。
         */
        recordTrustedToolReads(
          binding.taskId,
          sheets.map((sheet) => sheet.name)
        );
        return {
          kind: "result",
          result: executeQueryTableData(request, sheets)
        };
      } catch (reason) {
        return { kind: "error", errorCode: wpsBridgeErrorCode(reason) };
      }
    },

    receipt: async (current, call, outcome) => {
      const receipt: SdkBridgeToolReceipt =
        outcome.kind === "result"
          ? call.tool === "prepare_python_data"
            ? { datasets: (outcome.result as { datasets: SdkBridgeDatasetReceipt[] }).datasets }
            : { result: outcome.result as SdkBridgeToolReceipt["result"] }
          : { errorCode: outcome.errorCode };
      return postReceipt(
        {
          ...current,
          operationId: call.operationId,
          requestFingerprint: call.requestFingerprint,
          receipt
        },
        options.signal
      );
    }
  };
}

/**
 * 桥上的参数由服务端用 ``QueryTableArguments`` 校验过，这里只做形状闸：
 * 非对象直接按参数错误处理，真正的字段校验交给既有的确定性查询实现。
 */
function asQueryTableArguments(value: unknown): QueryTableArguments {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    throw new DataToolExecutionError("UNSUPPORTED_TOOL", "本地查询参数无效", false);
  }
  return value as QueryTableArguments;
}

/** 桥上的 ``prepare_python_data`` 参数形状闸（范围核对在 ``sliceWpsPythonSources``）。 */
function asWpsPreparePythonSources(value: unknown): PreparePythonDataSource[] {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    throw new DataToolExecutionError("UNSUPPORTED_TOOL", "数据登记参数无效", false);
  }
  const sources = (value as { sources?: unknown }).sources;
  if (!Array.isArray(sources)) {
    throw new DataToolExecutionError("UNSUPPORTED_TOOL", "数据登记来源无效", false);
  }
  return sources as PreparePythonDataSource[];
}

/** 解析 ``A1:B2`` 或 ``A1`` 为 0 基的 [r1,c1,r2,c2]。非法返回 null。 */
function parseA1Range(range: string): [number, number, number, number] | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/.exec(
    range.trim().toUpperCase()
  );
  if (!match) return null;
  const colToIndex = (col: string) => {
    let value = 0;
    for (const ch of col) value = value * 26 + (ch.charCodeAt(0) - 64);
    return value - 1;
  };
  const c1 = colToIndex(match[1]);
  const r1 = Number(match[2]) - 1;
  const c2 = match[3] ? colToIndex(match[3]) : c1;
  const r2 = match[4] ? Number(match[4]) - 1 : r1;
  if (r2 < r1 || c2 < c1) return null;
  return [r1, c1, r2, c2];
}

/**
 * 从已读的整表值里按来源声明的 A1 区域切片。
 *
 * 与 Excel 端 ``executePreparePythonData`` 同语义：每个来源的 sheet 必须落在
 * 授权工作表内，range 必须能解析为合法区域，且切片不能越界。越界/超限一律
 * 抛 ``SCOPE_VIOLATION``（由 ``wpsBridgeErrorCode`` 映射为固定错误码），不把
 * 细节透给模型。
 */
/**
 * 导出只为**双端共同契约测试**能直接驱动本宿主的切片路径
 * （`config/python_prepare_cases.json` 的那一批用例）；生产调用点只有
 * `execute` 里的 `prepare_python_data` 分支。
 */
export function sliceWpsPythonSources(
  sources: PreparePythonDataSource[],
  reads: WpsSheetReadResult[],
  authorizedSheetNames: string[]
): SdkBridgeDatasetReceipt[] {
  const allowed = new Set(authorizedSheetNames);
  const blocks: SdkBridgeDatasetReceipt[] = [];
  sources.forEach((source, sourceIndex) => {
    if (
      source === null ||
      typeof source !== "object" ||
      typeof source.sheet !== "string" ||
      typeof source.range !== "string"
    ) {
      throw new DataToolExecutionError("SCOPE_VIOLATION", "数据登记来源无效", false);
    }
    if (!allowed.has(source.sheet)) {
      throw new DataToolExecutionError(
        "SCOPE_VIOLATION",
        "数据登记请求超出了用户选择的工作表范围",
        false
      );
    }
    const bounds = parseA1Range(source.range);
    if (!bounds) {
      throw new DataToolExecutionError("SCOPE_VIOLATION", "数据登记区域无效", false);
    }
    const read = reads.find((item) => item.sheet.name === source.sheet);
    if (!read) {
      throw new DataToolExecutionError("SCOPE_VIOLATION", "找不到已授权工作表", false);
    }
    /*
     * 切片与"不补空白"的判定与 Excel 端**同一份实现**
     * （`@excel-bro/core/pythonPrepare`）。A1 是绝对坐标、`read.values` 是
     * `UsedRange` 的相对矩阵，直接拿绝对下标当矩阵下标会静默错位
     * `(startRow-1, startColumn-1)` 且**不报错**。
     */
    const rect = { r1: bounds[0], c1: bounds[1], r2: bounds[2], c2: bounds[3] };
    let sliced: WpsCellValue[][];
    try {
      sliced = slicePythonPrepareRegion(read.values, read, rect) as WpsCellValue[][];
    } catch (error) {
      if (error instanceof PythonPrepareRegionRefused) {
        throw new DataToolExecutionError("SCOPE_VIOLATION", error.message, false);
      }
      throw error;
    }
    blocks.push({ sourceIndex, values: normalizeWpsPreparedValues(sliced, read, rect) });
  });
  return blocks;
}

/**
 * 把切出来的原始值按**显示文本 + 数字格式**还原成用户看到的值。
 *
 * 与 Excel 端 `executePreparePythonData` 对 `range.values/text/numberFormat`
 * 的处理**逐字同一份**归一函数，只有"取哪三份矩阵"这一步按宿主不同：
 * Excel 一次读精确区域，WPS 从已用区域矩阵里按同一绝对坐标取回对应格。
 */
function normalizeWpsPreparedValues(
  values: WpsCellValue[][],
  read: WpsSheetReadResult,
  rect: PythonPrepareRect
): WpsCellValue[][] {
  if (!read.texts || !read.numberFormats) return values;
  const startRow = Number(read.startRow ?? 1);
  const startColumn = Number(read.startColumn ?? 1);
  return values.map((row, rowIndex) =>
    row.map((value, columnIndex) => {
      const matrixRow = rect.r1 + rowIndex + 1 - startRow;
      const matrixColumn = rect.c1 + columnIndex + 1 - startColumn;
      return normalizeCellValue(
        value,
        read.texts?.[matrixRow]?.[matrixColumn] ?? "",
        read.numberFormats?.[matrixRow]?.[matrixColumn] ?? "General"
      ) as WpsCellValue;
    })
  );
}

/**
 * 按引用取回结果值（WPS 结果卡的渲染路径）。
 *
 * 此时会话通常已关闭：服务端校验的是**授权归属**而不是活跃会话，所以要回带
 * 产生该结果那一轮的绑定（任务 + 工作簿身份），不能用"当前"绑定顶替。
 */
export async function loadWpsBridgeResult(
  binding: BridgeResultLookup,
  resultRef: string,
  signal?: AbortSignal
): Promise<SdkBridgeResultResponse | null> {
  try {
    return await fetchSdkBridgeResult({ ...binding, resultRef }, signal);
  } catch {
    return null;
  }
}
