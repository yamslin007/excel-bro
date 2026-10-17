import type { WpsRecoveryDescriptor } from "../../excel-addin/src/contracts";
import {
  asFormatFields,
  composedFormatSnapshot,
  declaredFormatFields,
  formatEvidenceFingerprint,
  formatExpectedFingerprint,
  isFormatActionType,
  matrixFingerprintV2,
  type WpsPreparedPlan
} from "./wpsPlanExecutor";
import {
  registerWpsExecution,
  getWpsExecutionStatus,
  startWpsAction,
  completeWpsAction,
  suspendWpsExecution,
  claimWpsExecutionResume,
  reconcileWpsActionReceipt,
  finishWpsExecution,
} from "../../excel-addin/src/api";
import type { WpsExecutionState } from "../../excel-addin/src/contracts";

/**
 * WPS 写入耐久恢复的持久登记模块。
 *
 * 权威现场是**服务端 SQLite 账本**（`office_executions` 表，WPS 用 `wps-path:` 身份
 * 命名空间）。本地只保存一份**发现索引** —— 记录 operationId、workbookIdentity 与
 * taskId，让面板重载后能重新查询服务端真实状态；**不**保存完整计划、授权、写前/写后
 * 矩阵或任何单元格值。索引丢失只会失去"自动发现"的能力，不会伪造或丢失写入现场。
 */

const PENDING_WPS_EXECUTION_STORAGE_KEY = "excel-bro.pending-wps-execution.v1";

/** 本地发现索引。只用于重载后定位服务端记录，不是权威状态。 */
export interface PendingWpsExecution {
  schemaVersion: 1;
  operationId: string;
  conversationId: string;
  activityMessageId: string;
  workbookIdentity: string;
  taskId: string;
  createdAt: string;
}

interface StorageLike {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

function browserStorage(): StorageLike | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

/**
 * WPS 跨关闭工作簿身份：已保存工作簿的小写磁盘完整路径，加 `wps-path:` 前缀。
 *
 * 不能单独用工作簿名（同名异路径会撞）、工作表索引（重排/删除会移位）或 JSAPI 对象
 * 引用（跨关闭失效）作为身份。完整路径是已保存 `.xlsx` 的稳定磁盘身份；前缀把它与
 * Office 的 `url:` 命名空间隔开，服务端对已存在记录按**完整字符串**比较，两者互不冒充。
 */
export function wpsWorkbookIdentity(fullName: string | null | undefined): string | null {
  const normalized = String(fullName ?? "").trim().toLowerCase();
  if (!/^[a-z]:\\[^\r\n]+\.xlsx$/.test(normalized)) return null;
  return `wps-path:${normalized}`;
}

export function savePendingWpsExecution(
  value: PendingWpsExecution,
  storage: StorageLike | null = browserStorage()
): boolean {
  if (!storage) return false;
  try {
    storage.setItem(PENDING_WPS_EXECUTION_STORAGE_KEY, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
}

function isPendingWpsExecution(value: unknown): value is PendingWpsExecution {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<PendingWpsExecution>;
  return (
    candidate.schemaVersion === 1 &&
    typeof candidate.operationId === "string" &&
    candidate.operationId.length >= 8 &&
    typeof candidate.conversationId === "string" &&
    Boolean(candidate.conversationId) &&
    typeof candidate.activityMessageId === "string" &&
    Boolean(candidate.activityMessageId) &&
    typeof candidate.workbookIdentity === "string" &&
    candidate.workbookIdentity.startsWith("wps-path:") &&
    typeof candidate.taskId === "string" &&
    Boolean(candidate.taskId) &&
    typeof candidate.createdAt === "string"
  );
}

export function loadPendingWpsExecution(
  storage: StorageLike | null = browserStorage()
): PendingWpsExecution | null {
  if (!storage) return null;
  try {
    const parsed: unknown = JSON.parse(
      storage.getItem(PENDING_WPS_EXECUTION_STORAGE_KEY) ?? "null"
    );
    if (isPendingWpsExecution(parsed)) return parsed;
    storage.removeItem(PENDING_WPS_EXECUTION_STORAGE_KEY);
    return null;
  } catch {
    try {
      storage.removeItem(PENDING_WPS_EXECUTION_STORAGE_KEY);
    } catch {
      // 本地索引只负责重载发现；服务端记录仍是权威来源。
    }
    return null;
  }
}

export function clearPendingWpsExecution(
  operationId: string,
  storage: StorageLike | null = browserStorage()
): void {
  if (!storage) return;
  try {
    const current = loadPendingWpsExecution(storage);
    if (!current || current.operationId === operationId) {
      storage.removeItem(PENDING_WPS_EXECUTION_STORAGE_KEY);
    }
  } catch {
    // 下次加载仍会按 operationId 查询服务端，不影响写入幂等状态。
  }
}

/**
 * 把已确认的写入计划投影成**恢复描述**。
 *
 * 这是出站闸的收口点：调用方只能传 `WpsPreparedPlan`，而这里只取动作类型、目标
 * 工作表与区域、目标形状，以及本机计算的写前/写后逻辑指纹。指纹不是加密。
 * `WpsPreparedAction.values`（目标值矩阵）与 `sourceSnapshots`（来源值）**一律不取**，
 * 所以它们没有任何路径能进入登记请求。
 *
 * `applyTo` / `formulaR1C1` 取自 `AnalysisPlan` 的原始动作：服务端要靠它们拒绝
 * 不受支持的变体（`clearRange` 非 contents、`writeFormulas` 用 R1C1）。
 *
 * **格式动作的指纹是格式证据，不是值矩阵指纹**：值动作的 `previewFingerprint`
 * 是"写前值"、`expectedFingerprint` 是"写后目标值"；格式动作不改值（值矩阵指纹在
 * 写前/写后完全一样，没有区分力），所以走 `formatEvidenceFingerprint`（写前现场，
 * 由预览时采集的逐格/逐行快照算出）与 `formatExpectedFingerprint`（写后目标，
 * 只由声明的属性与目标值加形状算出）。两者都**不含任何单元格内容**，也**不传格式
 * 快照本身** —— 快照含逐格属性值，属于本机现场。
 *
 * 同时投影出 `formatFields`（本次声明要改的属性名）：服务端按它做**属性级**校验，
 * 未验证属性（`underline` / `vertical` / `columnWidth`）在这里根本不会出现，
 * 因为 `declaredFormatFields` 只会产出已验证的属性。
 */
export function wpsRecoveryDescriptor(prepared: WpsPreparedPlan): WpsRecoveryDescriptor {
  return {
    actions: prepared.actions.map((action, index) => {
      const source = prepared.plan.actions[action.actionIndex];
      const shared = {
        actionIndex: action.actionIndex,
        type: action.type,
        sheet: action.sheetName,
        range: action.range,
        rowCount: action.rowCount,
        columnCount: action.columnCount,
        cellCount: action.cellCount,
        ...(source && "applyTo" in source ? { applyTo: source.applyTo } : {}),
        ...(source && "formulaR1C1" in source ? { formulaR1C1: Boolean(source.formulaR1C1) } : {})
      };
      if (isFormatActionType(action.type)) {
        // 格式动作必须有预览快照：没有它就无从证明"写前现场"，宁可拒绝登记。
        const expectedBefore = composedFormatSnapshot(prepared.actions, index);
        if (!expectedBefore) {
          throw new Error("格式动作缺少预览时采集的格式快照，无法生成恢复证据，已拒绝登记。");
        }
        const formatAction = asFormatFields(action);
        const formatFields = declaredFormatFields(formatAction);
        if (!formatFields.length) {
          throw new Error("格式动作没有声明任何可恢复属性，无法生成恢复证据，已拒绝登记。");
        }
        return {
          ...shared,
          /*
           * 写前证据取**合成后**的现场（前序同区域格式动作的声明值已叠加）：执行器在
           * 真正写这一笔之前看到的现场就是这个。若直接用预览原始快照，登记出的
           * `previewFingerprint` 会**永远达不到**，服务端随后以"证据与登记的写前状态
           * 不一致"拒掉写前检查点 —— 恢复链整条断在登记上。
           */
          previewFingerprint: formatEvidenceFingerprint(formatAction, expectedBefore),
          expectedFingerprint: formatExpectedFingerprint(
            formatAction,
            action.rowCount,
            action.columnCount
          ),
          formatFields: [...formatFields]
        };
      }
      return {
        ...shared,
        previewFingerprint: action.previewFingerprint,
        expectedFingerprint: matrixFingerprintV2(action.values, action.values)
      };
    })
  };
}

/**
 * 首写前登记：成功返回带 ownerToken 的状态；失败抛错，调用方据此中止写入。
 *
 * 入参是**已确认的 `WpsPreparedPlan`** 而不是 `AnalysisPlan`：描述投影在这里
 * 一次性完成，注册请求里不出现任何单元格矩阵。
 */
export async function registerWpsWriteIntent(input: {
  operationId: string;
  prepared: WpsPreparedPlan;
  authorizedSheetNames: string[];
  workbookIdentity: string;
  taskId: string;
}): Promise<WpsExecutionState> {
  return registerWpsExecution({
    operationId: input.operationId,
    descriptor: wpsRecoveryDescriptor(input.prepared),
    authorizedSheetNames: input.authorizedSheetNames,
    workbookIdentity: input.workbookIdentity,
    taskId: input.taskId
  });
}

export {
  getWpsExecutionStatus,
  startWpsAction,
  completeWpsAction,
  suspendWpsExecution,
  claimWpsExecutionResume,
  reconcileWpsActionReceipt,
  finishWpsExecution,
};
