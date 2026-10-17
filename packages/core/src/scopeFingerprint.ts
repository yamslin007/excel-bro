/** Local conversation binding, not a model/API payload or a write authorization. */
export interface ConversationScope {
  workbookName: string;
  sourceMode: "workbook" | "folder";
  selectionMode: "auto" | "manual" | "folder";
  activeWorksheet: string;
  sheets: ReadonlyArray<{ name: string }>;
}

/** Server supplies an absolute root. Preserve POSIX case and Windows UNC identity. */
export function folderRootIdentity(root?: string | null): string | null {
  if (!root?.trim()) return null;
  const slashPath = root.trim().replace(/\\/g, "/");
  if (/^[a-z]:\/+$/i.test(slashPath)) return `${slashPath.slice(0, 2).toLowerCase()}/`;
  const path = slashPath.replace(/\/+$/, "") || (slashPath === "/" ? "/" : "");
  if (/^[a-z]:\//i.test(path) || /^\/\/[^/]+\/[^/]+/.test(path)) {
    return path.toLowerCase();
  }
  return path.startsWith("/") && !path.startsWith("//") ? path || "/" : null;
}

export function intentScopeFingerprint(scope: ConversationScope, folderRoot?: string | null): string {
  const rootIdentity = scope.sourceMode === "folder" ? folderRootIdentity(folderRoot) : null;
  if (scope.sourceMode === "folder" && rootIdentity === null) {
    throw new Error("无法核实工作目录身份，请重新选择目录并确认工作表。");
  }
  return JSON.stringify({
    workbookName: scope.workbookName,
    sourceMode: scope.sourceMode,
    selectionMode: scope.selectionMode,
    sheets: scope.sheets.map((sheet) => sheet.name),
    ...(scope.selectionMode === "auto" ? { activeWorksheet: scope.activeWorksheet } : {}),
    // Never use the directory's basename as a file identity. Old folder bindings
    // lacking this field fail comparison and require a fresh request.
    ...(scope.sourceMode === "folder" ? { folderRoot: rootIdentity } : {})
  });
}
