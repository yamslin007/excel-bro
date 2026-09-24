export type WpsCollection<T> = {
  Count: number;
  Item(index: number | string): T;
};

export type WpsRange = {
  Address?: string | ((...args: unknown[]) => unknown);
  Text?: unknown;
  Value2?: unknown;
  Formula?: unknown;
  MergeCells?: unknown;
  Row?: number;
  Column?: number;
  Rows?: { Count: number };
  Columns?: { Count: number };
  Item?(rowIndex: number, columnIndex?: number): WpsRange;
};

export type WpsWorksheet = {
  Name: string;
  Index?: number;
  ProtectContents?: boolean;
  UsedRange?: WpsRange;
  Range?(cell1: WpsRange | string, cell2?: WpsRange | string): WpsRange;
};

export type WpsWorkbook = {
  Name: string;
  FullName?: string;
  Path?: string;
  ReadOnly?: boolean;
  Worksheets: WpsCollection<WpsWorksheet>;
  ActiveSheet?: WpsWorksheet;
};

export type WpsTaskPane = {
  ID: string | number;
  Visible: boolean;
  DockPosition?: number;
  Width?: number;
};

export type WpsPluginStorage = {
  getItem(key: string): unknown;
  setItem(key: string, value: unknown): void;
  removeItem?(key: string): void;
};

export type WpsApplication = {
  Version?: string;
  ActiveWorkbook?: WpsWorkbook;
  ActiveSheet?: WpsWorksheet;
  Selection?: WpsRange;
  PluginStorage: WpsPluginStorage;
  CreateTaskPane(url: string, title?: string): WpsTaskPane;
  GetTaskPane(id: string | number): WpsTaskPane | undefined;
  ribbonUI?: object;
};

export type WpsHostProbe = {
  host: "wps-et";
  workbookName: string;
  workbookPath: string | null;
  activeSheetName: string;
  sheetNames: string[];
  selectionAddress: string | null;
  selectionPreview: string | null;
};

declare global {
  interface Window {
    Application?: WpsApplication;
  }
}

export class WpsHostUnavailableError extends Error {
  constructor(message = "当前页面没有连接到 WPS 表格宿主。") {
    super(message);
    this.name = "WpsHostUnavailableError";
  }
}

export function getWpsApplication(scope: Pick<Window, "Application"> = window): WpsApplication {
  const application = scope.Application;
  if (!application?.PluginStorage || typeof application.CreateTaskPane !== "function") {
    throw new WpsHostUnavailableError();
  }
  return application;
}

function normalizeSelectionPreview(value: unknown): string | null {
  if (value == null || value === "") return null;
  if (Array.isArray(value)) {
    return value
      .slice(0, 3)
      .map((row) => Array.isArray(row) ? row.slice(0, 3).join(" · ") : String(row))
      .join(" / ");
  }
  return String(value);
}

function readRangeAddress(range: WpsRange | undefined): string | null {
  const address = range?.Address;
  if (typeof address === "function") {
    const value = address.call(range);
    return value == null || value === "" ? null : String(value);
  }
  return address == null || address === "" ? null : String(address);
}

export function readWpsHostProbe(application: WpsApplication): WpsHostProbe {
  const workbook = application.ActiveWorkbook;
  if (!workbook) {
    throw new WpsHostUnavailableError("WPS 表格中当前没有打开的工作簿。");
  }

  const sheetNames: string[] = [];
  for (let index = 1; index <= Number(workbook.Worksheets?.Count ?? 0); index += 1) {
    const sheet = workbook.Worksheets.Item(index);
    if (sheet?.Name) sheetNames.push(String(sheet.Name));
  }

  const activeSheet = application.ActiveSheet ?? workbook.ActiveSheet;
  const selection = application.Selection;
  const selectionText = selection?.Text ?? selection?.Value2;

  return {
    host: "wps-et",
    workbookName: String(workbook.Name || "未命名工作簿"),
    workbookPath: workbook.FullName ? String(workbook.FullName) : workbook.Path ? String(workbook.Path) : null,
    activeSheetName: activeSheet?.Name ? String(activeSheet.Name) : "未识别",
    sheetNames,
    selectionAddress: readRangeAddress(selection),
    selectionPreview: normalizeSelectionPreview(selectionText)
  };
}
