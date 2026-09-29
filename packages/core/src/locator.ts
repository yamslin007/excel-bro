/**
 * `WorkbookLocator` 与出网字符串之间的唯一折叠点（方案 §2.3 第 1 条、§2.4）。
 *
 * 共享层内部用判别式区分三种身份;`/api/office/...` 的 `workbookIdentity`
 * 仍是 `url:<...>` 字符串（`models.py:2020-2025` 强制该前缀,已落盘的
 * `OfficeExecutionState` 账本记录也按字符串解析）。折叠只在本文件发生,
 * 别处不得自行拼接或解析该前缀。
 */

import type { WorkbookLocator } from "./contracts";

/** 出网身份的固定前缀。服务端校验依赖它,不可更改。 */
const OFFICE_IDENTITY_PREFIX = "url:";

/**
 * Office 文档 URL → 出网身份字符串。
 *
 * 归一规则逐字复制 `excel.ts:3766`,不是重写：反斜杠转正斜杠 + 转小写。
 * 已落盘账本里的旧记录是按这条规则生成的,改动任何一步都会让在途任务
 * 的续跑与撤销入口对不上。
 */
export function officeIdentityFromUrl(url: string): string | null {
  const trimmed = url.trim();
  if (!trimmed) return null;
  return `${OFFICE_IDENTITY_PREFIX}${trimmed.replace(/\\/g, "/").toLocaleLowerCase()}`;
}

/**
 * `WorkbookLocator` → 出网身份字符串。
 *
 * 非 Office 通道**抛错而非返回伪身份**：WPS 恢复按方案 §2.3 第 2 条不走服务端,
 * 文件夹通道有自己的 manifest。若某处把 wps-path 塞进 `/api/office/...`,
 * 那是调用方分流错了,应当在此处炸开,而不是拼一个能通过 `url:` 校验的假串
 * 让服务端把两条通道的现场混在一个账本里。
 */
export function officeIdentityFromLocator(locator: WorkbookLocator): string | null {
  if (locator.kind !== "office-url") {
    throw new Error(
      `仅 Office 通道可折叠为 workbookIdentity，收到 kind=${locator.kind}`
    );
  }
  return officeIdentityFromUrl(locator.value);
}

/**
 * 出网身份字符串 → `WorkbookLocator`。读旧账本记录时用。
 *
 * 不认识的输入返回 `null` 而不是猜：缺前缀说明这不是 Office 身份,
 * 当成 office-url 处理会让后续比对静默通过。
 */
export function locatorFromOfficeIdentity(identity: string): WorkbookLocator | null {
  if (!identity.startsWith(OFFICE_IDENTITY_PREFIX)) return null;
  const value = identity.slice(OFFICE_IDENTITY_PREFIX.length);
  if (!value) return null;
  return { kind: "office-url", value };
}

/**
 * 两个身份是否指向同一工作簿。
 *
 * `kind` 不同一律视为不等 —— 三条通道的身份语义不同,磁盘路径与文档 URL
 * 即便字符串相同也不是同一个现场。同 kind 内按各自规则比：Office 已在
 * 折叠时归一,WPS 路径大小写不敏感（Windows 文件系统）,文件夹相对路径
 * 由服务端分配,原样比对。
 */
export function sameLocator(a: WorkbookLocator, b: WorkbookLocator): boolean {
  if (a.kind !== b.kind) return false;
  if (a.kind === "office-url") {
    return officeIdentityFromUrl(a.value) === officeIdentityFromUrl(b.value);
  }
  if (a.kind === "wps-path") {
    return (
      a.value.trim().replace(/\\/g, "/").toLocaleLowerCase() ===
      b.value.trim().replace(/\\/g, "/").toLocaleLowerCase()
    );
  }
  return a.value === b.value;
}
