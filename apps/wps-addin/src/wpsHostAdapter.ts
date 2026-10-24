/**
 * WPS 侧的 `HostAdapter` 实现（方案阶段 4 前置）。
 *
 * 这一层只做**包装**：能力判定、存储强制转换、身份判别式都委托给已有实现或 core,
 * 不在这里新增写入语义。执行预检（`prepareWpsPlan`）保持原样,按方案 §4.3
 * 它是第三道防线,不因为有了 `capability()` 就删。
 */

import capabilities from "../../../config/capabilities.json";
import type { AuthorizedSheetRef, CellValue, WorkbookLocator } from "@excel-bro/core/contracts";
import type {
  ActionCapability,
  HostAdapter,
  HostPrimitiveOps,
  KeyValueStore
} from "@excel-bro/core/ports";
import {
  getWpsApplication,
  WpsHostUnavailableError,
  type WpsApplication,
  type WpsPluginStorage
} from "./wpsHost";
import {
  assertWpsAuthorization,
  normalizeWpsMatrix,
  type WpsReadAuthorization
} from "./wpsWorkbookReader";
import { wpsWorkbookLocator } from "./wpsPlanExecutor";
import { wpsWriteVersionStatus } from "./wpsWriteVersionGate";

/**
 * `PluginStorage` → `KeyValueStore`。
 *
 * 两处必须处理,不是防御性代码而是接口差异:
 * - `getItem` 返回 `unknown`（`wpsHost.ts:44`）。非字符串一律折成 `null`,
 *   **不要 `String(value)`** —— 那会把 `undefined` 变成 `"undefined"` 这种
 *   能通过 JSON.parse 失败检查、却毫无意义的值存下去。
 * - `removeItem` 是**可选成员**（`wpsHost.ts:46`）。宿主没提供时退化为写空串,
 *   并保持 `getItem` 对空串返回 `null`,使"已删除"与"存过空串"在读侧一致。
 */
export function createWpsKeyValueStore(storage: WpsPluginStorage): KeyValueStore {
  return {
    getItem(key) {
      const raw = storage.getItem(key);
      if (typeof raw !== "string" || raw === "") return null;
      return raw;
    },
    setItem(key, value) {
      storage.setItem(key, value);
    },
    removeItem(key) {
      if (typeof storage.removeItem === "function") {
        storage.removeItem(key);
        return;
      }
      storage.setItem(key, "");
    }
  };
}

/**
 * 动作能力三态（方案 §4.2）。
 *
 * 一处诚实说明：配置只记录"哪些动作已验收",**不记录宿主 API 是否具备**。
 * 所以白名单外的动作只能报 `unverified`,不能报 `unsupported` ——
 * 后者是"宿主 API 不具备"的断言,拿不出证据就不该下。
 * 两者行为相同（都不进计划）,区别只在给用户的说明,谎报无收益。
 */
export function resolveWpsCapability(
  actionType: string,
  hostVersion: string | null
): ActionCapability {
  const allowed = [...capabilities.wpsWrite.allowedActions, ...capabilities.worksheetCheckpoint.allowedActions].some(
    (entry) => entry === actionType || entry.startsWith(`${actionType}:`)
  );
  if (!capabilities.wpsWrite.enabled) {
    return { state: "unverified", reason: "WPS 写入能力当前整体关闭。" };
  }
  if (!allowed) {
    return {
      state: "unverified",
      reason: `EB 的 WPS 执行器尚未开放「${actionType}」的执行、回读与恢复闭环，规划阶段已排除；不代表 WPS 软件没有此能力。`
    };
  }
  if (wpsWriteVersionStatus(hostVersion) === "blocked") {
    return {
      state: "unverified",
      reason: `当前 WPS JSAPI 版本 ${hostVersion || "未知"} 尚未开放写入。`
    };
  }
  return { state: "supported" };
}

/**
 * WPS 侧的能力**自报**集合（方案 §4.3 第 1 条）。
 *
 * 语义要说清楚：这是"本端认为可用的动作"，**不是授权凭据**。
 * 服务端按可信配置算上界后与它求交，客户端只能收窄。
 * 因此这里多报不产生任何收益（多出来的会被上界砍掉），少报只会让自己
 * 的功能更少 —— 这也正是它不需要额外可信机制的原因。
 *
 * 自报内容严格从 `wpsWrite.allowedActions` 与版本闸派生，
 * 不在这里手工维护第二份名单：两份名单漂移时，规划前交集与执行预检
 * 会各说各话。
 */
export function wpsReportedActionTypes(hostVersion: string | null): string[] {
  if (!capabilities.wpsWrite.enabled) return [];
  if (wpsWriteVersionStatus(hostVersion) === "blocked") return [];
  return [...capabilities.wpsWrite.allowedActions, ...capabilities.worksheetCheckpoint.allowedActions];
}

export function createWpsHostAdapter(
  application: WpsApplication = getWpsApplication(),
  authorization?: WpsReadAuthorization
): HostAdapter {
  const storage = createWpsKeyValueStore(application.PluginStorage);

  const hostVersion = (): string | null => {
    const version = String(application.Version ?? "").trim();
    return version || null;
  };

  return {
    id: "wps-et",
    storage,
    isAvailable() {
      return (
        Boolean(application.PluginStorage) &&
        typeof application.CreateTaskPane === "function"
      );
    },
    hostVersion,
    async ready() {
      // WPS 的 JSAPI 在页面脚本执行时已可用，没有 Office.onReady 的等价物。
    },
    locator(): WorkbookLocator | undefined {
      if (!authorization) return undefined;
      try {
        return wpsWorkbookLocator(authorization);
      } catch {
        // 未保存/非 .xlsx。返回 undefined 而不是占位身份：
        // 身份不稳定时后续一切恢复承诺都不成立，应让调用方看见"没有身份"。
        return undefined;
      }
    },
    capability(actionType) {
      return resolveWpsCapability(actionType, hostVersion());
    },
    async runBatch(work) {
      if (!authorization) {
        throw new WpsHostUnavailableError("尚未完成工作表授权，拒绝执行任何写入原语。");
      }
      return work(createWpsPrimitiveOps(application, authorization));
    }
  };
}

/**
 * 写入原语。**只负责落值与原样读回,不做验收判定** ——
 * 判定归 core,否则又会长出第二套折叠语义（`normalizeCell` / `canonicalCell` 的教训）。
 */
export function createWpsPrimitiveOps(
  application: WpsApplication,
  authorization: WpsReadAuthorization
): HostPrimitiveOps {
  const resolve = (sheet: AuthorizedSheetRef, address: string) => {
    const authorized = authorization.sheets.find((entry) => entry.name === sheet.name);
    if (!authorized) {
      throw new Error(`计划引用了未授权工作表「${sheet.name}」，已拒绝。`);
    }
    const worksheet = assertWpsAuthorization(application, authorization, authorized);
    const range = worksheet.Range?.(address);
    if (!range) throw new Error(`无法解析区域「${address}」。`);
    return range;
  };

  // 宿主对单个单元格返回标量、对单列/单行返回一维数组，形状要靠 Rows/Columns 还原。
  const shape = (range: ReturnType<typeof resolve>, address: string) => {
    const rows = Number(range.Rows?.Count ?? 0);
    const columns = Number(range.Columns?.Count ?? 0);
    if (rows < 1 || columns < 1) throw new Error(`无法确定区域「${address}」的形状。`);
    return { rows, columns };
  };

  return {
    async setValues(sheet, address, values) {
      resolve(sheet, address).Value2 = values;
    },
    async readValues(sheet, address) {
      const range = resolve(sheet, address);
      const { rows, columns } = shape(range, address);
      return normalizeWpsMatrix(range.Value2, rows, columns);
    },
    async readFormulas(sheet, address) {
      const range = resolve(sheet, address);
      const { rows, columns } = shape(range, address);
      return normalizeWpsMatrix(range.Formula, rows, columns);
    },
    async clearContents(sheet, address) {
      const range = resolve(sheet, address);
      const { rows, columns } = shape(range, address);
      range.Value2 = Array.from({ length: rows }, () =>
        Array.from({ length: columns }, () => null as CellValue)
      );
    }
  };
}
