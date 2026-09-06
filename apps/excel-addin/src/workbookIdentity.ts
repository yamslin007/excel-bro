export function workbookNameFromDocumentUrl(documentUrl?: string | null): string {
  const raw = documentUrl?.trim();
  if (!raw) return "未保存的工作簿";

  const withoutQuery = raw.split(/[?#]/, 1)[0].replace(/\\/g, "/");
  const encodedName = withoutQuery.slice(withoutQuery.lastIndexOf("/") + 1);
  if (!encodedName) return "未保存的工作簿";

  try {
    return decodeURIComponent(encodedName);
  } catch {
    return encodedName;
  }
}

export type WorkbookLocation =
  | { kind: "local"; folderPath: string }
  | { kind: "unsaved" }
  | { kind: "remote" };

/**
 * 从 Office.context.document.url 推导工作簿所在的本地文件夹。
 *
 * 备份目录固定在数据文件夹下（server 的 _backup_root），所以只要拿到本地目录就能
 * 定位备份。云端（OneDrive/SharePoint/Excel Online）路径不是本地文件系统路径，
 * 后端 openpyxl 也访问不到，直接判为 remote。
 */
export function workbookFolderFromDocumentUrl(
  documentUrl?: string | null
): WorkbookLocation {
  const raw = documentUrl?.trim();
  if (!raw) return { kind: "unsaved" };

  // file:/// 协议前缀不影响本地路径判断，先去掉。
  const withoutProtocol = raw.replace(/^file:\/\/\//, "");
  const isDriveLetterPath = /^[A-Za-z]:[\\/]/.test(withoutProtocol);
  const isUncPath = /^\\\\[^\\]/.test(withoutProtocol);
  if (!isDriveLetterPath && !isUncPath) return { kind: "remote" };

  const separator = withoutProtocol.includes("\\") ? "\\" : "/";
  const lastSeparator = Math.max(
    withoutProtocol.lastIndexOf("\\"),
    withoutProtocol.lastIndexOf("/")
  );
  if (lastSeparator < 0) return { kind: "unsaved" };

  const folderPath = withoutProtocol.slice(0, lastSeparator);
  // 盘符根目录（C:\ 之类）截断后只剩 "C:"，补回分隔符才是合法路径。
  const normalized = /^[A-Za-z]:$/.test(folderPath)
    ? `${folderPath}${separator}`
    : folderPath;
  return normalized ? { kind: "local", folderPath: normalized } : { kind: "unsaved" };
}

function normalizedDate(year: string, month: string, day: string): string {
  return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
}

export function extractWorkbookDataPeriod(
  workbookName: string
): string | null {
  const matches = [
    ...workbookName.matchAll(
      /(\d{4})\s*(?:-|\/|\.|年)\s*(\d{1,2})\s*(?:-|\/|\.|月)\s*(\d{1,2})\s*日?/g
    )
  ].map((match) => normalizedDate(match[1], match[2], match[3]));

  if (matches.length === 0) return null;
  if (matches.length === 1 || matches[0] === matches[1]) return matches[0];
  return `${matches[0]} 至 ${matches[1]}`;
}
