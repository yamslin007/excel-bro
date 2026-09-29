/**
 * Office 侧的 `HostAdapter` 实现（方案阶段 4 前置）。
 *
 * 与 WPS 适配器对称：只做包装,不新增写入语义。现有 `executeAction`
 * （`excel.ts` 约 1050 行）保持原样,阶段 5 才逐动作改走原语。
 */

import type { AuthorizedSheetRef, CellValue, WorkbookLocator } from "@excel-bro/core/contracts";
import type {
  ActionCapability,
  HostAdapter,
  HostPrimitiveOps,
  KeyValueStore
} from "@excel-bro/core/ports";
import { currentOfficeWorkbookLocator, isRunningInExcel } from "./excel";
import { OFFICE_EXECUTED_ACTION_TYPES } from "./officeActions";

/**
 * `localStorage` → `KeyValueStore`。
 *
 * 取不到时返回一个恒空的实现,而不是抛错：`localStorage` 在
 * 部分 Office 宿主容器与隐私模式下会不可用,而调用方大多是读配置,
 * 读不到走默认值即可。**写入失败不静默** —— 让异常抛出,
 * 否则用户会以为设置保存成功了。
 */
export function createOfficeKeyValueStore(
  scope: { localStorage?: Storage } = globalThis as { localStorage?: Storage }
): KeyValueStore {
  const store = scope.localStorage;
  if (!store) {
    return {
      getItem: () => null,
      setItem: () => {
        throw new Error("当前环境不提供 localStorage，无法保存设置。");
      },
      removeItem: () => {}
    };
  }
  return {
    getItem: (key) => store.getItem(key),
    setItem: (key, value) => store.setItem(key, value),
    removeItem: (key) => store.removeItem(key)
  };
}

/**
 * Office 通道的动作能力。
 *
 * 现阶段全部报 `supported`：`executeAction` 覆盖 `ExcelAction` 全部成员,
 * 且这些动作本就是在 Office 上开发并长期使用的。
 *
 * 这**不是**"Office 无需能力矩阵"的结论 —— 方案 §4.3 要求可续跑/可规划是
 * 宿主 × 版本 × 动作的交集。Office 侧目前没有按版本收窄的证据,
 * 硬编码一份未经验收的版本表只会制造假精确。真需要收窄时（例如某动作在
 * 特定 Office 版本上行为不同）,在此处按 `Office.context.diagnostics.version` 加判断。
 */
export function resolveOfficeCapability(actionType: string): ActionCapability {
  if (!actionType.trim()) {
    return { state: "unsupported", reason: "动作类型为空。" };
  }
  return { state: "supported" };
}

/**
 * Office 侧的能力**自报**集合（方案 §4.3 第 1 条）。
 *
 * 与 `wpsReportedActionTypes` 对称：这是"本端认为可用的动作"，**不是授权凭据**。
 * 服务端按 `capabilities.json` 的 `hostCapabilities["excel-office"]` 算上界后与它
 * 求交，客户端只能收窄不能扩大。
 *
 * 名单从 `ExcelAction` 的真实执行入口派生（`excel.ts` 的 `executeAction` 分支），
 * 而不是手工维护第二份清单 —— 两份名单漂移时，规划前交集与执行预检会各说各话。
 *
 * 传 `undefined` 表示"本端不声明任何动作"，用于拿不到宿主环境（如纯对话页）时；
 * 此时服务端交集为空，不会有任何写入类动作进计划。
 */
export function officeReportedActionTypes(
  supported: readonly string[] | undefined
): string[] {
  if (!supported) return [];
  // 去重但不改顺序：调用方若按重要性排序，交集结果也保持同样的顺序。
  return [...new Set(supported.map((entry) => entry.trim()).filter(Boolean))];
}

/** 规划请求里与宿主能力收窄相关的三个字段。 */
export interface OfficeHostRequestFields {
  hostId: "excel-office" | null;
  hostVersion: string | null;
  selfReportedActions: string[];
}

/**
 * 组装 Office 通道的规划请求字段（方案 §4.3）。
 *
 * 抽成纯函数是为了可测：真实的 `Excel`/`Office` 全局对象在单测里不存在，
 * 而这段逻辑恰恰是"报错方向必须是少做而不是多做"的关键，必须能被断言。
 *
 * @param runningInExcel 是否真的跑在 Excel 宿主内
 * @param hostVersion 宿主自报版本，拿不到传 `null`
 */
export function officeHostRequestFields(
  runningInExcel: boolean,
  hostVersion: string | null
): OfficeHostRequestFields {
  if (!runningInExcel) {
    // 文件夹模式与浏览器预览走 openpyxl 通道，动作覆盖面与 Office 不同，
    // 不能借 Office 的名义放宽边界。
    return { hostId: null, hostVersion: null, selfReportedActions: [] };
  }
  const version = hostVersion?.trim() || null;
  return {
    hostId: "excel-office",
    // 拿不到版本时仍报 excel-office，但版本为 null：服务端按前缀匹配，
    // null 不命中 `16.0` 这一族，上界为空 —— 这是"宁可少做"的失败方向。
    hostVersion: version,
    selfReportedActions: officeReportedActionTypes(OFFICE_EXECUTED_ACTION_TYPES)
  };
}

/** 从全局 `Office` 取宿主版本；不可用时返回 `null`。 */
export function currentOfficeHostVersion(): string | null {
  if (typeof Office === "undefined") return null;
  const version = (Office as unknown as {
    context?: { diagnostics?: { version?: unknown } };
  }).context?.diagnostics?.version;
  return typeof version === "string" && version.trim() ? version.trim() : null;
}

export function createOfficeHostAdapter(): HostAdapter {
  return {
    id: "excel-office",
    storage: createOfficeKeyValueStore(),
    isAvailable: isRunningInExcel,
    hostVersion() {
      if (typeof Office === "undefined") return null;
      const version = (Office as any).context?.diagnostics?.version;
      return typeof version === "string" && version.trim() ? version.trim() : null;
    },
    ready() {
      if (typeof Office === "undefined" || typeof Office.onReady !== "function") {
        return Promise.resolve();
      }
      return Office.onReady().then(() => undefined);
    },
    locator(): WorkbookLocator | undefined {
      return currentOfficeWorkbookLocator();
    },
    capability: resolveOfficeCapability,
    runBatch(work) {
      return Excel.run(async (context) => work(createOfficePrimitiveOps(context)));
    }
  };
}

/**
 * 写入原语。
 *
 * `Excel.run` 的 `context.sync()` 是事务边界,所以这里每个原语内部自己 sync ——
 * 与 WPS 逐条落地的行为对齐。core 的验收判定因此不必知道两端谁有事务：
 * 一致的保证是"半写可恢复",不是"要么全成要么全不动"。
 */
export function createOfficePrimitiveOps(
  context: Excel.RequestContext
): HostPrimitiveOps {
  const resolve = (sheet: AuthorizedSheetRef, address: string) =>
    context.workbook.worksheets.getItem(sheet.name).getRange(address);

  return {
    async setValues(sheet, address, values) {
      resolve(sheet, address).values = values as unknown[][];
      await context.sync();
    },
    async readValues(sheet, address) {
      const range = resolve(sheet, address);
      range.load("values");
      await context.sync();
      return range.values as CellValue[][];
    },
    async readFormulas(sheet, address) {
      const range = resolve(sheet, address);
      range.load("formulas");
      await context.sync();
      return range.formulas as CellValue[][];
    },
    async clearContents(sheet, address) {
      resolve(sheet, address).clear(Excel.ClearApplyTo.contents);
      await context.sync();
    }
  };
}
