function stripUrlSuffix(value: string): string {
  const suffixIndex = value.search(/[?#]/);
  return suffixIndex >= 0 ? value.slice(0, suffixIndex) : value;
}

function decodePath(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function collapseSlashes(value: string): string {
  const isUnc = value.startsWith("//");
  const collapsed = value.replace(/\/{2,}/g, "/");
  return isUnc ? `/${collapsed}` : collapsed;
}

/** 归一化为可比较的 Windows 绝对路径；无法确定时返回 null。 */
export function normalizeWorkbookPath(raw?: string | null): string | null {
  const trimmed = raw?.trim();
  if (!trimmed) return null;

  let value = stripUrlSuffix(trimmed.replace(/\\/g, "/"));
  if (/^file:\/\//i.test(value)) {
    const remainder = value.replace(/^file:\/\//i, "");
    value = /^\/+[a-z]:\//i.test(remainder)
      ? remainder.replace(/^\/+/, "")
      : `//${remainder.replace(/^\/+/, "")}`;
  } else if (
    /^[a-z][a-z0-9+.-]*:/i.test(value) &&
    !/^[a-z]:\//i.test(value)
  ) {
    return null;
  }

  value = collapseSlashes(decodePath(value).replace(/\\/g, "/")).replace(
    /\/+$/,
    ""
  );
  if (!/^[a-z]:\//i.test(value) && !/^\/\/[^/]+\/[^/]+/.test(value)) {
    return null;
  }
  return value.toLocaleLowerCase();
}

/** 文件夹侧的 folderPath + relativePath 拼成同一形态。 */
export function folderFileFullPath(
  folderPath: string,
  relativePath: string
): string | null {
  const folder = normalizeWorkbookPath(folderPath);
  const relative = relativePath.trim().replace(/\\/g, "/");
  if (
    !folder ||
    !relative ||
    /^[a-z]:\//i.test(relative) ||
    /^\/\//.test(relative) ||
    /^[a-z][a-z0-9+.-]*:/i.test(relative)
  ) {
    return null;
  }
  return normalizeWorkbookPath(
    `${folder.replace(/\/+$/, "")}/${relative.replace(/^\/+/, "")}`
  );
}