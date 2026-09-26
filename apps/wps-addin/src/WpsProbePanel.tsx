import { useCallback, useEffect, useState } from "react";
import { reportDevProbe } from "./devProbe";
import { initializeRestoredTaskPaneWidth } from "./ribbon";
import WpsAgentPanel from "./WpsAgentPanel";
import { getWpsApplication, readWpsHostProbe, type WpsHostProbe } from "./wpsHost";
import {
  listWpsWorksheets,
  type WpsWorkbookMetadata
} from "./wpsWorkbookReader";

type ProbeState =
  | { status: "loading" }
  | { status: "ready"; value: WpsHostProbe }
  | { status: "error"; message: string };

function runProbe(): ProbeState {
  try {
    return { status: "ready", value: readWpsHostProbe(getWpsApplication()) };
  } catch (error) {
    return {
      status: "error",
      message: error instanceof Error ? error.message : "无法读取 WPS 表格宿主。"
    };
  }
}

export default function WpsProbePanel() {
  const [probe, setProbe] = useState<ProbeState>({ status: "loading" });
  const [metadata, setMetadata] = useState<WpsWorkbookMetadata | null>(null);
  const [selectedIndexes, setSelectedIndexes] = useState<number[]>([]);
  const [scopeOpen, setScopeOpen] = useState(false);
  const [scopeMode, setScopeMode] = useState<"auto" | "manual">("auto");

  useEffect(() => {
    try {
      initializeRestoredTaskPaneWidth();
    } catch {
      // WPS may have restored the webview before its task pane handle is available.
    }
    const nextProbe = runProbe();
    setProbe(nextProbe);
    if (nextProbe.status === "ready") {
      const nextMetadata = listWpsWorksheets(getWpsApplication());
      setMetadata(nextMetadata);
      setSelectedIndexes(nextMetadata.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index));
    }
    reportDevProbe("taskpane", {
      connected: nextProbe.status === "ready",
      hasWorkbook: nextProbe.status === "ready" && Boolean(nextProbe.value.workbookName),
      sheetCount: nextProbe.status === "ready" ? nextProbe.value.sheetNames.length : 0
    });
  }, []);

  const refresh = () => {
    const nextProbe = runProbe();
    setProbe(nextProbe);
    if (nextProbe.status !== "ready") {
      setMetadata(null);
      setSelectedIndexes([]);
      return;
    }
    const nextMetadata = listWpsWorksheets(getWpsApplication());
    setMetadata(nextMetadata);
    setSelectedIndexes((current) => {
      if (scopeMode === "auto") {
        return nextMetadata.sheets
          .filter((sheet) => sheet.isActive)
          .map((sheet) => sheet.index);
      }
      const available = new Set(nextMetadata.sheets.map((sheet) => sheet.index));
      return current.filter((index) => available.has(index));
    });
  };

  const chooseAutoScope = () => {
    setScopeMode("auto");
    setSelectedIndexes(
      metadata?.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index) ?? []
    );
  };

  const chooseManualScope = () => {
    setScopeMode("manual");
    if (selectedIndexes.length > 0) return;
    setSelectedIndexes(
      metadata?.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index) ?? []
    );
  };

  const toggleSheet = (index: number) => {
    setSelectedIndexes((current) =>
      current.includes(index) ? current.filter((value) => value !== index) : [...current, index]
    );
  };

  const resolveSelectedSheetIndexes = useCallback((): readonly number[] => {
    try {
      const nextProbe = runProbe();
      if (nextProbe.status !== "ready") {
        setProbe(nextProbe);
        setMetadata(null);
        setSelectedIndexes([]);
        return [];
      }
      const nextMetadata = listWpsWorksheets(getWpsApplication());
      const available = new Set(nextMetadata.sheets.map((sheet) => sheet.index));
      const nextIndexes = scopeMode === "auto"
        ? nextMetadata.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index)
        : selectedIndexes.filter((index) => available.has(index));
      setProbe(nextProbe);
      setMetadata(nextMetadata);
      setSelectedIndexes(nextIndexes);
      return nextIndexes;
    } catch {
      setSelectedIndexes([]);
      return [];
    }
  }, [scopeMode, selectedIndexes]);

  if (probe.status === "loading") {
    return (
      <main className="connection-shell">
        <span className="probe-logo">EB</span>
        <div><strong>Excel Bro</strong><p>正在连接 WPS 表格…</p></div>
      </main>
    );
  }

  if (probe.status === "error") {
    return (
      <main className="connection-shell error">
        <span className="probe-logo">EB</span>
        <div><strong>尚未连接 WPS 表格</strong><p>{probe.message}</p></div>
        <button type="button" onClick={refresh}>重新检测</button>
      </main>
    );
  }

  const selectedSheets = metadata?.sheets.filter((sheet) => selectedIndexes.includes(sheet.index)) ?? [];
  const scopeLabel = selectedSheets.length === 0
    ? "尚未选择工作表"
    : scopeMode === "auto" && selectedSheets.length === 1
      ? `当前表 ${selectedSheets[0].name}`
      : selectedSheets.length === 1
      ? `${selectedSheets[0].isActive ? "当前表" : "已选"} ${selectedSheets[0].name}`
      : `已固定 ${selectedSheets.length} 个工作表`;

  const scopeDock = (
    <section className={`wps-scope-dock ${scopeOpen ? "open" : ""}`}>
      {scopeOpen && (
        <section className="scope-popover wps-scope-popover" aria-label="选择数据范围">
          <div className="scope-popover-scroll">
          <div className="scope-popover-header">
            <div>
              <strong>这次对话使用哪些数据？</strong>
              <span>默认跟随你当前正在查看的工作表</span>
            </div>
            <button type="button" aria-label="关闭数据范围" onClick={() => setScopeOpen(false)}>×</button>
          </div>
          <div className="wps-current-file-card">
            <span aria-hidden="true">XLSX</span>
            <div>
              <small>当前文件</small>
              <strong>{probe.value.workbookName}</strong>
              <em>当前工作表 {probe.value.activeSheetName}</em>
            </div>
          </div>
          <div className="scope-controls">
            <div className="scope-mode-list">
              <button
                type="button"
                className={scopeMode === "auto" ? "selected" : ""}
                onClick={chooseAutoScope}
              >
                <i aria-hidden="true">⊙</i>
                <span><strong>跟随当前</strong><small>发送时自动使用正在查看的工作表</small></span>
              </button>
              <button
                type="button"
                className={scopeMode === "manual" ? "selected" : ""}
                onClick={chooseManualScope}
              >
                <i aria-hidden="true">☷</i>
                <span><strong>选择多个</strong><small>用于跨表查询、比较或汇总</small></span>
              </button>
              <button type="button" disabled title="WPS 当前仅支持打开的工作簿">
                <i aria-hidden="true">▫</i>
                <span><strong>工作目录</strong><small>WPS 当前仅支持打开的工作簿</small></span>
              </button>
            </div>

            {scopeMode === "manual" && (
              <div className="sheet-picker-options wps-sheet-picker-options">
                {metadata?.sheets.map((sheet) => (
                  <button
                    type="button"
                    className={selectedIndexes.includes(sheet.index) ? "selected" : ""}
                    key={`${sheet.index}:${sheet.name}`}
                    aria-pressed={selectedIndexes.includes(sheet.index)}
                    onClick={() => toggleSheet(sheet.index)}
                  >
                    <i aria-hidden="true">{selectedIndexes.includes(sheet.index) ? "✓" : ""}</i>
                    <span>{sheet.name}</span>
                    {sheet.isActive && <small>当前表</small>}
                  </button>
                ))}
                {selectedIndexes.length === 0 && (
                  <p className="authorization-empty">尚未选择工作表，不会读取或修改任何数据。</p>
                )}
              </div>
            )}
          </div>
          </div>
          <div className="wps-scope-actions">
            <button type="button" className="scope-refresh" onClick={refresh}>↻ 刷新</button>
            <button
              type="button"
              className="scope-confirm"
              disabled={selectedIndexes.length === 0}
              onClick={() => setScopeOpen(false)}
            >
              完成
            </button>
          </div>
        </section>
      )}
      <div className="scope-entry-row">
        <button
          className={`scope-trigger ${scopeOpen ? "open" : ""}`}
          type="button"
          aria-expanded={scopeOpen}
          onClick={() => setScopeOpen((current) => !current)}
        >
          <i aria-hidden="true">▦</i>
          <span>
            <small title={probe.value.workbookName}>{probe.value.workbookName}</small>
            <span className="scope-sep" aria-hidden="true">|</span>
            <strong>{scopeLabel}</strong>
          </span>
          <b>{scopeOpen ? "⌄" : "⌃"}</b>
        </button>
      </div>
    </section>
  );

  return (
    <main className="probe-shell">
      <WpsAgentPanel
        selectedSheetIndexes={selectedIndexes}
        resolveSelectedSheetIndexes={resolveSelectedSheetIndexes}
        scopeDock={scopeDock}
      />
    </main>
  );
}
