/**
 * L3 端口层：**只有接口定义,没有实现**（方案 §5.2）。
 *
 * 环境依赖在这里收口 —— 存储、宿主、网络、时间各自抽成接口,
 * 由两个 app 注入具体实现。core 的 L4 编排只认这些接口。
 *
 * 本阶段只定义阶段 4（参数化列映射落成 `writeValues`）真正需要的成员。
 * 不预先铺 38 个动作的原语:没有使用点的接口无法被测试约束,
 * 只会在阶段 5 逐动作推进时被改写。
 */

import type { AuthorizedSheetRef, CellValue, WorkbookLocator } from "./contracts";

/** 宿主标识。服务端不感知该值（方案 §2.3 第 2 条:恢复按通道分开,不靠服务端分支）。 */
export type HostId = "excel-office" | "wps-et" | "folder";

/**
 * 同步键值存储。Excel 注入 `localStorage`,WPS 注入 `PluginStorage`。
 *
 * 保持同步是已确认的决策：`WpsPluginStorage`（`wpsHost.ts:43-47`）本就是同步的,
 * 不需要水合层。但适配器有两件事要做:
 * - `getItem` 在 WPS 侧返回 `unknown`,须做字符串强制并把非字符串返回折成 `null`
 * - `removeItem` 在 WPS 侧是**可选成员**,适配器需在缺失时退化为写空串或自行兜底
 */
export interface KeyValueStore {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

/**
 * 动作在"宿主 × 版本"上的可用状态（方案 §4.2）。
 *
 * `unverified` 与 `unsupported` 都**不进计划**,区别只在给用户的说明:
 * 前者是未验收,后者是宿主 API 不具备。两者都必须带原因,
 * 因为方案 §4.3 要求排除发生在规划前,用户不能看到一个外观相同的空入口。
 */
export type ActionCapability =
  | { state: "supported" }
  | { state: "unverified"; reason: string }
  | { state: "unsupported"; reason: string };

/** 能力查询器。`actionType` 用 string 而非 `ExcelAction["type"]`：后者仍在 app 侧的 contracts。 */
export interface CapabilityResolver {
  capability(actionType: string): ActionCapability;
}

/**
 * 工作簿快照的读取入口**不在本接口里**。
 *
 * 原先这里有 `readSheets(sheets, budget): Promise<unknown>`。返回 `unknown` 的
 * 成员约束不了任何东西 —— 两端实现返回什么都合法,类型检查形同虚设,
 * 而真正的快照类型 `WorkbookSnapshot` 仍在 app 侧的 `contracts.ts`。
 * 等它随 contracts 拆分迁入 L0 后再定义这个成员。
 *
 * 阶段 4 的参数化列映射只需要读已授权的源区域,`HostPrimitiveOps.readValues`
 * 已覆盖;两端现有的 `captureWorkbook` / `listWpsWorksheets` 不经本接口。
 */

/**
 * 写入原语。**只做一件事:把值落到宿主上,不做验收判定。**
 *
 * 回读比对属于 L2 语义层,不能由宿主实现各自解释 —— 那正是
 * `normalizeCell` 与 `canonicalCell` 分叉的来源。原语只负责
 * "写"和"原样读回",判定统一在 core。
 */
export interface HostPrimitiveOps {
  setValues(sheet: AuthorizedSheetRef, address: string, values: CellValue[][]): Promise<void>;
  readValues(sheet: AuthorizedSheetRef, address: string): Promise<CellValue[][]>;
  readFormulas(sheet: AuthorizedSheetRef, address: string): Promise<CellValue[][]>;
  clearContents(sheet: AuthorizedSheetRef, address: string): Promise<void>;
}

/**
 * 宿主适配器。
 *
 * 不含撤销成员：三条通道的撤销机制完全不同（方案 §2.2 —— 服务端账本 /
 * `PluginStorage` / 文件级副本）,粒度也不同（区域 / 区域 / 整文件原子替换）。
 * 把它们塞进一个接口就是在两套语义之间架桥,而项目定论是让桥不必存在。
 * 撤销按通道各自定义,阶段 4 再决定共享的是哪一层。
 */
export interface HostAdapter extends CapabilityResolver {
  readonly id: HostId;
  readonly storage: KeyValueStore;
  isAvailable(): boolean;
  /** 完整版本串。能力矩阵按版本前缀匹配（方案 §4.2）,故不要在此处截断。 */
  hostVersion(): string | null;
  ready(): Promise<void>;
  /** 无法稳定识别时返回 `undefined`,不返回占位身份。 */
  locator(): WorkbookLocator | undefined;
  /**
   * 批量执行原语。
   *
   * **不承诺整体原子性。** Office 有 `Excel.run` 的事务语义,WPS 没有等价物,
   * 顺序执行到一半失败就是半写状态。跨宿主一致的是"半写可恢复",
   * 靠检查点与写前快照兜底,不是靠假装有事务。
   */
  runBatch<T>(work: (ops: HostPrimitiveOps) => Promise<T>): Promise<T>;
}

/**
 * 服务端客户端。请求/响应**类型**属 L0,传输实现留 app（方案 §5.3）。
 *
 * 本阶段只定义最小形态,`api.ts` 的 1146 行不在阶段 3 范围内。
 */
export interface ApiClient {
  post<TRequest, TResponse>(
    path: string,
    body: TRequest,
    init?: { signal?: AbortSignal }
  ): Promise<TResponse>;
  get<TResponse>(path: string, init?: { signal?: AbortSignal }): Promise<TResponse>;
}

/** 时间源。抽成接口才能让验收判定与超时逻辑可测。 */
export interface Clock {
  now(): number;
  /** ISO 8601。落盘记录用它,不要各处自行 `new Date().toISOString()`。 */
  nowIso(): string;
}
