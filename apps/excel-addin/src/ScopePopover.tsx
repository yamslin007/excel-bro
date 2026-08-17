// 范围浮层：选择本次对话使用的数据范围（跟随工作表 / 多表 / 文件夹）。
// 纯展示组件，JSX 从 App.tsx 逐字搬移；状态与回调由 props 传入。
import type { Dispatch, SetStateAction } from "react";
import type {
  FolderCatalog,
  WorkbookSnapshot,
  WorksheetSnapshot
} from "./contracts";
import { folderSheetKey } from "./utils";
import type { Status } from "./types/chat";
import type { SourceMode, WorkbookScopeMode } from "./types/workbook";

interface ScopePopoverProps {
  contextOpen: boolean;
  setContextOpen: Dispatch<SetStateAction<boolean>>;
  workbook: WorkbookSnapshot | null;
  sourceMode: SourceMode;
  workbookScopeMode: WorkbookScopeMode;
  workbookDataPeriod: string | null;
  chooseAutomaticScope: () => void;
  chooseManualScope: () => void;
  chooseFolderScope: () => void;
  sheetSearch: string;
  setSheetSearch: Dispatch<SetStateAction<string>>;
  selectedSheetNames: string[];
  setSelectedSheetNames: Dispatch<SetStateAction<string[]>>;
  setSelectionConfirmed: Dispatch<SetStateAction<boolean>>;
  filteredWorksheets: WorksheetSnapshot[];
  toggleSheet: (sheetName: string) => void;
  busy: boolean;
  browseFolder: () => Promise<void>;
  folderCatalog: FolderCatalog | null;
  folderSheetKeys: string[];
  toggleFolderSheet: (fileId: string, sheetName: string) => void;
  scan: (options?: { announce?: boolean }) => Promise<void>;
  status: Status;
  confirmSheetSelection: () => Promise<void>;
}

export default function ScopePopover({
  contextOpen,
  setContextOpen,
  workbook,
  sourceMode,
  workbookScopeMode,
  workbookDataPeriod,
  chooseAutomaticScope,
  chooseManualScope,
  chooseFolderScope,
  sheetSearch,
  setSheetSearch,
  selectedSheetNames,
  setSelectedSheetNames,
  setSelectionConfirmed,
  filteredWorksheets,
  toggleSheet,
  busy,
  browseFolder,
  folderCatalog,
  folderSheetKeys,
  toggleFolderSheet,
  scan,
  status,
  confirmSheetSelection
}: ScopePopoverProps) {
  return (
    <>
      {contextOpen && workbook && (
        <section className="scope-popover" aria-label="选择数据范围">
          <div className="scope-popover-header">
            <div>
              <strong>这次对话使用哪些数据？</strong>
              <span>默认跟随你当前正在查看的工作表</span>
            </div>
            <button
              onClick={() => setContextOpen(false)}
              aria-label="关闭数据范围选择"
            >
              ×
            </button>
          </div>

          {sourceMode === "workbook" && (
            <div className="workbook-identity">
              <i aria-hidden="true">XLSX</i>
              <span>
                <small>当前文件</small>
                <strong title={workbook.name}>{workbook.name}</strong>
                <em>
                  {workbookDataPeriod
                    ? `报表日期 ${workbookDataPeriod}`
                    : "文件名中未识别到报表日期"}
                  {" · "}
                  当前工作表 {workbook.activeWorksheet}
                </em>
              </span>
            </div>
          )}

          <div className="scope-mode-list">
            <button
              className={
                sourceMode === "workbook" && workbookScopeMode === "auto"
                  ? "selected"
                  : ""
              }
              onClick={chooseAutomaticScope}
            >
              <i>◎</i>
              <span>
                <strong>跟随当前工作表</strong>
                <small>发送时自动使用正在查看的工作表</small>
              </span>
            </button>
            <button
              className={
                sourceMode === "workbook" && workbookScopeMode === "manual"
                  ? "selected"
                  : ""
              }
              onClick={chooseManualScope}
            >
              <i>☷</i>
              <span>
                <strong>选择多个工作表</strong>
                <small>用于跨表查询、比较或汇总</small>
              </span>
            </button>
            <button
              className={sourceMode === "folder" ? "selected" : ""}
              onClick={chooseFolderScope}
            >
              <i>⌑</i>
              <span>
                <strong>选择文件夹</strong>
                <small>批量处理多个本地工作簿</small>
              </span>
            </button>
          </div>

          {sourceMode === "workbook" && workbookScopeMode === "manual" && (
            <div className="workbook-sheet-picker">
              <div className="sheet-search-row">
                <input
                  value={sheetSearch}
                  onChange={(event) => setSheetSearch(event.target.value)}
                  placeholder="搜索工作表…"
                  aria-label="搜索工作表"
                />
                <span>
                  已选 {selectedSheetNames.length}/{workbook.worksheets.length}
                </span>
              </div>
              <div className="sheet-picker-toolbar">
                <button
                  onClick={() => {
                    setSelectedSheetNames((current) => [
                      ...new Set([
                        ...current,
                        ...filteredWorksheets.map((sheet) => sheet.name)
                      ])
                    ]);
                    setSelectionConfirmed(false);
                  }}
                >
                  全选搜索结果
                </button>
                <button
                  onClick={() => {
                    setSelectedSheetNames([]);
                    setSelectionConfirmed(false);
                  }}
                >
                  清空
                </button>
              </div>
              <div className="workbook-sheet-list">
                {filteredWorksheets.map((sheet) => {
                  const selected = selectedSheetNames.includes(sheet.name);
                  return (
                    <button
                      className={selected ? "selected" : ""}
                      key={sheet.name}
                      onClick={() => toggleSheet(sheet.name)}
                      aria-pressed={selected}
                    >
                      <i aria-hidden="true">{selected ? "✓" : ""}</i>
                      <span>
                        <strong>{sheet.name}</strong>
                        <small>
                          {sheet.name === workbook.activeWorksheet
                            ? "当前工作表 · "
                            : ""}
                          {sheet.rowCount} 行 · {sheet.columnCount} 列
                        </small>
                      </span>
                    </button>
                  );
                })}
                {filteredWorksheets.length === 0 && (
                  <p>没有匹配的工作表</p>
                )}
              </div>
            </div>
          )}

          {sourceMode === "folder" && (
            <div className="folder-scope-picker">
              <button
                className="browse-folder-button"
                disabled={busy}
                onClick={() => void browseFolder()}
              >
                {folderCatalog
                  ? `更换文件夹 · ${folderCatalog.folderName}`
                  : "选择文件夹"}
              </button>

              {folderCatalog && (
                <div className="folder-file-list">
                  {folderCatalog.files.map((file) => (
                    <div className="folder-file" key={file.id}>
                      <strong>{file.relativePath}</strong>
                      {file.error ? (
                        <small className="file-error">{file.error}</small>
                      ) : (
                        <div className="sheet-picker-options">
                          {file.worksheets.map((sheet) => {
                            const selected = folderSheetKeys.includes(
                              folderSheetKey(file.id, sheet.name)
                            );
                            return (
                              <button
                                key={sheet.name}
                                className={selected ? "selected" : ""}
                                onClick={() =>
                                  toggleFolderSheet(file.id, sheet.name)
                                }
                                aria-pressed={selected}
                              >
                                <span>{sheet.name}</span>
                                <small>
                                  {sheet.rowCount} 行 · {sheet.columnCount} 列
                                </small>
                              </button>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          <div className="scope-popover-actions">
            {sourceMode === "workbook" ? (
              <button
                disabled={busy}
                onClick={() => void scan({ announce: true })}
              >
                {status === "scanning" ? "读取中…" : "刷新工作簿"}
              </button>
            ) : (
              <span>已选 {folderSheetKeys.length} 个工作表</span>
            )}
            <button
              className="scope-confirm"
              disabled={
                sourceMode === "workbook"
                  ? workbookScopeMode === "manual" &&
                    selectedSheetNames.length === 0
                  : folderSheetKeys.length === 0
              }
              onClick={() => void confirmSheetSelection()}
            >
              {status === "scanning" ? "读取中…" : "完成"}
            </button>
          </div>
        </section>
      )}
    </>
  );
}
