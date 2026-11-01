/**
 * 测试用的**本机后端恢复包存储**替身（不是宿主临时文件，也不是内存对象直传）。
 *
 * 为什么替身必须长这样（2026-10-27 真实 WPS 验收的教训）：恢复包存在本机后端的应用
 * 数据目录（`server/app/execution_recovery.py`），**完整退出 WPS 只会丢掉宿主自己的
 * `PluginStorage`，不会丢后端那一份**。所以"WPS 重启后还能续跑"的正确替身是：
 *
 * - 换掉**宿主**（新的 `WpsApplication` 对象 + 空的 `PluginStorage`）；
 * - 保留**后端存储**（它是另一个进程，WPS 重启与它无关）；
 * - 材料只能以**文本**跨边界：写入时序列化、读回时重新解析。把内存对象直接递给调用方
 *   会抹掉"文本损坏/字段漂移/逐字节不一致"这一整类判据。
 *
 * 语义对齐服务端（`server/app/execution_recovery.py`）：
 * - 保存前必须过共享协议校验（`assertExecutionRecoveryPackage`）；
 * - 同 `operationId` + 同正文 = 幂等，返回既有元数据，**不刷新** `expiresAt`；
 * - 同 `operationId` + 不同正文 = 冲突；
 * - 过期后读不出正文（线上是 410，这里是抛错，让调用方走 `unavailable`）。
 */

import { createHash } from "node:crypto";
import { assertExecutionRecoveryPackage, parseExecutionRecoveryPackage } from "@excel-bro/core/executionRecoveryPackage";
import type { ExecutionRecoveryPackageMaterial } from "../../excel-addin/src/contracts";
import capabilities from "../../../config/capabilities.json";

type Stored = ExecutionRecoveryPackageMaterial;

function digestOf(text: string): string {
  return createHash("sha256").update(text, "utf8").digest("hex");
}

function sizeOf(text: string): number {
  return typeof TextEncoder === "function" ? new TextEncoder().encode(text).length : text.length;
}

/**
 * 与服务端异常族同形的最小失败对象：线上是 `ApiRequestError`（带 `code`），
 * 替身只带 `recoveryState` 标记。`wpsRecoveryPackage` 两条都认，所以替身不必
 * 反过来 import 被 mock 的 api 模块（那会绕出一个循环依赖）。
 */
function failure(state: "expired" | "corrupt" | "conflict" | "quota", message: string): Error {
  return Object.assign(new Error(message), {recoveryState: state});
}

export type RecoveryBackend = {
  /** `PUT /api/execution-recovery/{id}`。 */
  save(operationId: string, packageJson: string): Stored;
  /** `GET /api/execution-recovery/{id}`：没有就是 `null`（线上 404）。 */
  load(operationId: string): Stored | null;
  /** 模拟"保留期到了"：把到期时间改到过去，正文仍在盘上。 */
  forceExpire(operationId: string): void;
  /** 模拟磁盘上的字节被外部改坏（哈希对不上）。 */
  corrupt(operationId: string, replacement: string): void;
  /** 终态回收（`DELETE`）：不存在也算成功，与服务端幂等语义一致。 */
  remove(operationId: string): void;
  /** 后端进程重启：内存索引清空，模拟"这一份材料从来没有存过"。 */
  clear(): void;
  /** 磁盘上还剩哪些执行的材料。 */
  ids(): string[];
  /** 落盘调用次数：证明"写前落盘"真的发生过、且只发生一次。 */
  saveCalls: number;
};

export function createRecoveryBackend(
  options: { retentionDays?: number; now?: () => number } = {}
): RecoveryBackend {
  const files = new Map<string, Stored>();
  const now = options.now ?? (() => Date.now());
  const retentionDays = options.retentionDays ?? capabilities.taskRecovery.unfinishedRecoveryPackageRetentionDays;
  const backend: RecoveryBackend = {
    saveCalls: 0,
    save(operationId, packageJson) {
      backend.saveCalls += 1;
      // 先校验再落盘：形状不对的材料绝不允许出现在"以后能被读回来"的地方。
      assertExecutionRecoveryPackage(parseExecutionRecoveryPackage(packageJson));
      const existing = files.get(operationId);
      if (existing) {
        if (existing.packageJson !== packageJson) throw failure("conflict", "该执行已保存过不同的恢复材料，拒绝覆盖。");
        // 幂等：同正文不重写，也不顺带把保留期往后推。
        // 已经过期也一样——过期之后重复保存**不能**重新生效（线上是 410）。
        if (Date.parse(existing.expiresAt) <= now()) throw failure("expired", "恢复材料已超过保留期，不能重新保存或续期。");
        return { ...existing };
      }
      const savedAt = new Date(now()).toISOString();
      const stored: Stored = {
        packageJson,
        digest: digestOf(packageJson),
        size: sizeOf(packageJson),
        createdAt: savedAt,
        savedAt,
        expiresAt: new Date(now() + retentionDays * 86_400_000).toISOString()
      };
      files.set(operationId, stored);
      return { ...stored };
    },
    load(operationId) {
      const stored = files.get(operationId);
      if (!stored) return null;
      // 与服务端一样"按状态分开报"：过期 ≠ 损坏。调用方据此给用户不同的说明，
      // 而不是把两种情况都说成"找不到材料"。
      if (digestOf(stored.packageJson) !== stored.digest)
        throw failure("corrupt", "恢复材料在磁盘上已损坏。");
      if (Date.parse(stored.expiresAt) <= now())
        throw failure("expired", "恢复材料已超过保留期。");
      return { ...stored };
    },
    forceExpire(operationId) {
      const stored = files.get(operationId);
      if (stored) files.set(operationId, { ...stored, expiresAt: new Date(now() - 1000).toISOString() });
    },
    corrupt(operationId, replacement) {
      const stored = files.get(operationId);
      // 只改正文不改哈希：这正是"磁盘被外部改过"的样子。
      if (stored) files.set(operationId, { ...stored, packageJson: replacement });
    },
    remove(operationId) {
      files.delete(operationId);
    },
    clear() {
      files.clear();
    },
    ids() {
      return [...files.keys()];
    }
  };
  return backend;
}

/**
 * 测试文件共享的"那台后端"。
 *
 * 模块级而不是每个 `fixture()` 一个新的：WPS 重启不该清空它（它是另一个进程）。
 * 需要"后端也重启"的用例显式调用 `resetBackendRecovery()`。
 */
let backend = createRecoveryBackend();

export function backendRecovery(): RecoveryBackend {
  return backend;
}

export function resetBackendRecovery(options?: { retentionDays?: number; now?: () => number }): RecoveryBackend {
  backend = createRecoveryBackend(options);
  return backend;
}

/** 供 `vi.mock("../../excel-addin/src/api")` 直接挂上去的两个替身。 */
export function backendSave(operationId: string, packageJson: string): Stored {
  return backend.save(operationId, packageJson);
}

export function backendLoad(operationId: string): Stored | null {
  return backend.load(operationId);
}

export function backendRemove(operationId: string): void {
  backend.remove(operationId);
}
