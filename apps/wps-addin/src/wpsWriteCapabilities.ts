import capabilities from "../../../config/capabilities.json";
import type { WpsApplication, WpsWorkbook, WpsWorksheet } from "./wpsHost";
import { wpsWriteVersionStatus, type WpsWriteVersionStatus } from "./wpsWriteVersionGate";

export type WpsWriteCapabilityCheck = {
  id: string;
  label: string;
  passed: boolean;
  detail: string;
};

export type WpsWriteCapabilityProbe = {
  hostVersion: string | null;
  versionStatus: WpsWriteVersionStatus;
  workbookName: string | null;
  selectedSheetCount: number;
  checks: WpsWriteCapabilityCheck[];
  readyForHostValidation: boolean;
  writeEnabled: boolean;
};

function isLocalXlsx(fullName: string): boolean {
  return /^[A-Za-z]:\\[^\r\n]+\.xlsx$/i.test(fullName);
}

function selectedSheets(
  workbook: WpsWorkbook,
  indexes: readonly number[]
): WpsWorksheet[] | null {
  const unique = [...new Set(indexes)];
  if (unique.length === 0) return null;
  const count = Number(workbook.Worksheets?.Count ?? 0);
  if (unique.some((index) => !Number.isInteger(index) || index < 1 || index > count)) {
    return null;
  }
  try {
    const sheets = unique.map((index) => workbook.Worksheets.Item(index));
    return sheets.every((sheet) => Boolean(sheet?.Name)) ? sheets : null;
  } catch {
    return null;
  }
}

/**
 * W3 preflight. It never sets a WPS property or sends workbook data to the server.
 * Passing the probe only makes preview/confirmation available; execution repeats all checks.
 */
export function probeWpsWriteCapabilities(
  application: WpsApplication,
  selectedSheetIndexes: readonly number[]
): WpsWriteCapabilityProbe {
  const workbook = application.ActiveWorkbook;
  const selected = workbook ? selectedSheets(workbook, selectedSheetIndexes) : null;
  const checks: WpsWriteCapabilityCheck[] = [];
  const add = (id: string, label: string, passed: boolean, detail: string) => {
    checks.push({ id, label, passed, detail });
  };

  const fullName = String(workbook?.FullName ?? "").trim();
  const hostVersion = String(application.Version ?? "").trim();
  const versionStatus = wpsWriteVersionStatus(hostVersion);
  const versionAllowed = versionStatus !== "blocked";
  add("workbook", "当前工作簿", Boolean(workbook), workbook ? "已连接" : "没有打开的工作簿");
  add(
    "hostVersion",
    "WPS 宿主版本",
    versionAllowed,
    versionStatus === "verified"
      ? `已验证 JSAPI 版本族 ${hostVersion}`
      : versionStatus === "development"
        ? `开发调试模式临时开放 JSAPI 版本 ${hostVersion}，仍需实机写入与撤销验收`
        : `当前 JSAPI 版本 ${hostVersion || "未知"} 尚未开放写入`
  );
  add(
    "localXlsx",
    "已保存的本地 .xlsx",
    isLocalXlsx(fullName),
    isLocalXlsx(fullName) ? "文件类型与路径符合首版范围" : "需要本地磁盘上的已保存 .xlsx"
  );
  add(
    "readOnly",
    "工作簿可编辑状态",
    workbook?.ReadOnly === false,
    workbook?.ReadOnly === false ? "宿主明确报告可编辑" : "宿主未明确报告可编辑"
  );
  add(
    "authorization",
    "工作表授权",
    selected !== null,
    selected ? `已选择 ${selected.length} 张有效工作表` : "请重新选择至少一张有效工作表"
  );

  let rangeApiAvailable = Boolean(selected?.length);
  let formulaReadable = Boolean(selected?.length);
  let mergeReadable = Boolean(selected?.length);
  let protectionReadable = Boolean(selected?.length);
  for (const sheet of selected ?? []) {
    if (sheet.ProtectContents !== false) protectionReadable = false;
    try {
      const cell = sheet.Range?.("A1");
      if (!cell || typeof cell !== "object") {
        rangeApiAvailable = false;
        formulaReadable = false;
        mergeReadable = false;
        continue;
      }
      // Read only. Blank cells may return null/undefined for Value2 or Formula.
      void cell.Value2;
      if (cell.Formula === undefined) formulaReadable = false;
      if (typeof cell.MergeCells !== "boolean") mergeReadable = false;
    } catch {
      rangeApiAvailable = false;
      formulaReadable = false;
      mergeReadable = false;
    }
  }
  add("range", "区域读取接口", rangeApiAvailable, rangeApiAvailable ? "A1 可读取" : "Range 或 Value2 不可读取");
  add("formula", "公式快照接口", formulaReadable, formulaReadable ? "Formula 读取未报错" : "Formula 无法读取");
  add("merge", "合并单元格检查", mergeReadable, mergeReadable ? "MergeCells 返回布尔值" : "MergeCells 无法确定");
  add("protection", "工作表保护状态", protectionReadable, protectionReadable ? "已选表均明确未保护" : "存在保护表或状态无法确定");

  return {
    hostVersion: hostVersion || null,
    versionStatus,
    workbookName: workbook?.Name ? String(workbook.Name) : null,
    selectedSheetCount: selected?.length ?? 0,
    checks,
    readyForHostValidation: checks.every((check) => check.passed),
    writeEnabled: capabilities.wpsWrite.enabled && checks.every((check) => check.passed)
  };
}

export function wpsWriteFeatureFlagEnabled(): boolean {
  return capabilities.wpsWrite.enabled;
}
