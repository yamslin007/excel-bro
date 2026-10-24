import type { WpsApplication, WpsWorkbook } from './wpsHost';

/**
 * **本次结果审阅**与它所属原表之间的绑定。
 *
 * 解决的真实问题：点「双视图对比」后 WPS 会把副本激活成 `ActiveWorkbook`。
 * 而确认/放弃、数据范围核对都从 `ActiveWorkbook` 取目标身份，于是副本被当成
 * "另一个工作簿"，界面报"工作簿或授权发生变化"，范围面板也开始读副本。
 *
 * 设计边界（不可放宽）：
 *  - 只认**本次审阅**创建的那一份副本，判据是创建时记下的完整身份
 *    （name + fullName + path）+ 未保存 + 单表 + 与登记的原表并存且唯一；
 *    不用"工作簿1"、`EB_` 前缀、活动窗口或名字相似度。
 *  - 只影响**读取**（识别当前目标是谁、列出工作表）；写入/关闭各自有明确目标，
 *    不经过这里，也不做"全局把 ActiveWorkbook 换成原表" —— 无关工作簿仍是无关工作簿。
 *  - 绑定只在这一个任务窗格内有效，不授予任何新权限。
 */

export type WpsReviewCopyIdentity = { name: string; fullName: string; path: string | null };
export type WpsReviewTargetBinding = {
  /** 执行时那一本原表（宿主对象；可能被宿主重建包装，故同时按身份查找）。 */
  original: WpsWorkbook;
  originalIdentity: WpsReviewCopyIdentity;
  /** 本次副本创建时的身份。 */
  copyIdentity: WpsReviewCopyIdentity;
  /** 写入目标表（副本只复制这一张）。 */
  writeSheet: { index: number; name: string };
};

let binding: WpsReviewTargetBinding | null = null;

const sameIdentity = (left: WpsReviewCopyIdentity, right: WpsReviewCopyIdentity) =>
  left.name === right.name && left.fullName === right.fullName && left.path === right.path;

export function identityOf(book: WpsWorkbook | undefined | null): WpsReviewCopyIdentity | null {
  if (!book) return null;
  try {
    return {
      name: String(book.Name ?? ""),
      // 与 `wpsWorkbookReader.workbookFullName` 同一口径，但在本模块内联，避免循环依赖。
      fullName: String(book.FullName ?? "") ||
        (book.Path ? `${String(book.Path).replace(/[\\/]$/, "")}\\${String(book.Name)}` : ""),
      path: String(book.Path ?? "").trim() || null
    };
  } catch { return null; }
}

export function bindWpsReviewTarget(value: WpsReviewTargetBinding): void { binding = value; }
export function releaseWpsReviewTarget(): void { binding = null; }
export function currentWpsReviewTarget(): WpsReviewTargetBinding | null { return binding; }

/** 这本工作簿是否就是**本次审阅**的副本（不是"像副本"）。 */
export function isWpsReviewCopyWorkbook(book: WpsWorkbook | undefined | null): boolean {
  if (!binding || !book) return false;
  const identity = identityOf(book);
  if (!identity || !sameIdentity(identity, binding.copyIdentity)) return false;
  // 已另存成磁盘文件 ⇒ 那是用户自己的文件，不再当成我们的副本。
  if (identity.path) return false;
  try {
    if (Number(book.Worksheets?.Count ?? 0) !== 1) return false;
    if (String(book.Worksheets.Item(1).Name ?? "") !== binding.writeSheet.name) return false;
  } catch { return false; }
  return true;
}

/**
 * 如果 `ActiveWorkbook` 是本次副本，返回**仍在场的原表**；否则返回 null。
 *
 * "仍在场"要能被证明：有 `Workbooks` 集合时要求原表与副本各恰好一份；
 * 没有集合时退回到"绑定的原表对象仍可读、且那张表还在"。证不出来就返回 null ——
 * 上层据此保持原判定（会拦住操作），而不是猜一个原表继续。
 */
export function resolveWpsReviewOriginal(app: WpsApplication): WpsWorkbook | null {
  const active = app.ActiveWorkbook;
  if (!binding || !isWpsReviewCopyWorkbook(active)) return null;
  try {
    if (app.Workbooks && Number(app.Workbooks.Count ?? 0) > 0) {
      let originals = 0;
      let copies = 0;
      for (let index = 1; index <= Number(app.Workbooks.Count); index += 1) {
        const identity = identityOf(app.Workbooks.Item(index));
        if (!identity) continue;
        if (sameIdentity(identity, binding.originalIdentity)) originals += 1;
        if (sameIdentity(identity, binding.copyIdentity)) copies += 1;
      }
      if (originals !== 1 || copies !== 1) return null;
    }
    const original = binding.original;
    if (!original) return null;
    if (!sameIdentity(identityOf(original) ?? binding.originalIdentity, binding.originalIdentity)) return null;
    if (String(original.Worksheets.Item(binding.writeSheet.index)?.Name ?? "") !== binding.writeSheet.name) return null;
    return original;
  } catch { return null; }
}

/**
 * **读取用**的工作簿：当前活动簿是本次副本时解析为原表，否则就是活动簿本身。
 *
 * 只给"识别目标 / 列工作表 / 取身份"这类读取用；写入与关闭不经过这里。
 */
export function wpsReadWorkbook(app: WpsApplication): WpsWorkbook | undefined {
  return resolveWpsReviewOriginal(app) ?? app.ActiveWorkbook;
}

/** 当前活动簿是否就是本次副本（供"这是副本，不是新目标"的判断使用）。 */
export function wpsActiveIsReviewCopy(app: WpsApplication): boolean {
  return isWpsReviewCopyWorkbook(app.ActiveWorkbook);
}