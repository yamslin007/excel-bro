import { useEffect, useRef, useState } from "react";
import { reportDevProbe } from "./devProbe";
import { initializeRestoredTaskPaneWidth } from "./ribbon";
import WpsAgentPanel from "./WpsAgentPanel";
import { getWpsApplication, readWpsHostProbe, type WpsHostProbe } from "./wpsHost";
import {
  authorizeWpsSheets,
  listWpsWorksheets,
  readAuthorizedWpsSheets,
  type WpsSheetReadResult,
  type WpsWorkbookMetadata
} from "./wpsWorkbookReader";
import {
  probeWpsWriteCapabilities,
  type WpsWriteCapabilityProbe
} from "./wpsWriteCapabilities";

type ProbeState =
  | { status: "loading" }
  | { status: "ready"; value: WpsHostProbe }
  | { status: "error"; message: string };

type ReadState =
  | { status: "idle" }
  | { status: "reading"; message: string }
  | { status: "ready"; results: WpsSheetReadResult[] }
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
  const [readState, setReadState] = useState<ReadState>({ status: "idle" });
  const [scopeOpen, setScopeOpen] = useState(false);
  const [writeProbe, setWriteProbe] = useState<WpsWriteCapabilityProbe | null>(null);
  const readAbortRef = useRef<AbortController | null>(null);

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
    readAbortRef.current?.abort();
    const nextProbe = runProbe();
    setProbe(nextProbe);
    setReadState({ status: "idle" });
    setWriteProbe(null);
    if (nextProbe.status !== "ready") {
      setMetadata(null);
      setSelectedIndexes([]);
      return;
    }
    const nextMetadata = listWpsWorksheets(getWpsApplication());
    setMetadata(nextMetadata);
    setSelectedIndexes(nextMetadata.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index));
  };

  const toggleSheet = (index: number) => {
    setSelectedIndexes((current) =>
      current.includes(index) ? current.filter((value) => value !== index) : [...current, index]
    );
    setReadState({ status: "idle" });
    setWriteProbe(null);
  };

  const readAuthorizedData = async () => {
    const controller = new AbortController();
    readAbortRef.current?.abort();
    readAbortRef.current = controller;
    try {
      const application = getWpsApplication();
      const authorization = authorizeWpsSheets(application, selectedIndexes);
      setReadState({ status: "reading", message: "正在核对授权范围…" });
      const results = await readAuthorizedWpsSheets(application, authorization, {
        signal: controller.signal,
        onProgress: ({ sheet, rowsRead, totalRows }) => {
          setReadState({
            status: "reading",
            message: `正在读取「${sheet.name}」：${rowsRead} / ${totalRows} 行`
          });
        }
      });
      setReadState({ status: "ready", results });
    } catch (error) {
      setReadState({
        status: "error",
        message: error instanceof Error ? error.message : "WPS 工作表读取失败。"
      });
    } finally {
      if (readAbortRef.current === controller) readAbortRef.current = null;
    }
  };

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
    : selectedSheets.length === 1
      ? `${selectedSheets[0].isActive ? "当前表" : "已选"} ${selectedSheets[0].name}`
      : `已选 ${selectedSheets.length} 张表`;

  const scopeDock = (
    <section className={`eb-scope-dock ${scopeOpen ? "open" : ""}`}>
      {scopeOpen && (
        <div className="eb-scope-panel">
          <div className="scope-panel-heading">
            <div>
              <strong>授权读取范围</strong>
              <p>只读取明确勾选的工作表；取消全部选择不会自动扩大范围。</p>
            </div>
            <button type="button" onClick={refresh}>刷新结构</button>
          </div>
          <div className="sheet-options">
            {metadata?.sheets.map((sheet) => (
              <label className="sheet-option" key={`${sheet.index}:${sheet.name}`}>
                <input
                  type="checkbox"
                  checked={selectedIndexes.includes(sheet.index)}
                  onChange={() => toggleSheet(sheet.index)}
                />
                <span>{sheet.name}</span>
                {sheet.isActive && <small>当前表</small>}
              </label>
            ))}
          </div>
          {selectedIndexes.length === 0 && (
            <p className="authorization-empty">尚未选择工作表，不会读取任何数据。</p>
          )}
          <div className="scope-actions">
            <button
              className="scope-verify-button"
              type="button"
              disabled={selectedIndexes.length === 0 || readState.status === "reading"}
              onClick={() => void readAuthorizedData()}
            >
              {readState.status === "reading" ? "读取中…" : "验证读取"}
            </button>
            {readState.status === "reading" && (
              <button className="scope-stop-button" type="button" onClick={() => readAbortRef.current?.abort()}>
                停止
              </button>
            )}
          </div>
          {readState.status === "reading" && <p className="read-status-text">{readState.message}</p>}
          {readState.status === "error" && <p className="read-error">{readState.message}</p>}
          {readState.status === "ready" && (
            <div className="read-results">
              {readState.results.map((result) => (
                <details key={`${result.sheet.index}:${result.sheet.name}`}>
                  <summary>
                    {result.sheet.name} · {result.rowCount} 行 × {result.columnCount} 列
                  </summary>
                  <div className="preview-table-wrap">
                    <table>
                      <tbody>
                        {result.values.slice(0, 5).map((row, rowIndex) => (
                          <tr key={rowIndex}>
                            {row.slice(0, 5).map((cell, columnIndex) => (
                              <td key={columnIndex}>{cell == null ? "" : String(cell)}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </details>
              ))}
            </div>
          )}
          <details className="scope-diagnostics">
            <summary>连接信息</summary>
            <dl>
              <div><dt>工作簿</dt><dd>{probe.value.workbookName}</dd></div>
              <div><dt>当前表</dt><dd>{probe.value.activeSheetName}</dd></div>
              <div><dt>当前选区</dt><dd>{probe.value.selectionAddress ?? "未识别"}</dd></div>
            </dl>
          </details>
          {import.meta.env.DEV && (
            <details className="scope-diagnostics">
              <summary>W3 写入能力检查（只读）</summary>
              <p>仅检查当前工作簿和接口，不会修改单元格。</p>
              <button
                type="button"
                onClick={() => setWriteProbe(
                  probeWpsWriteCapabilities(getWpsApplication(), selectedIndexes)
                )}
              >
                检查当前工作簿
              </button>
              {writeProbe && (
                <div>
                  <p>宿主版本：{writeProbe.hostVersion ?? "未提供"}</p>
                  <p>{writeProbe.readyForHostValidation
                    ? "只读前置检查通过；仍需真实写入与恢复验收。"
                    : "前置检查未通过，写入保持关闭。"}</p>
                  <ul>
                    {writeProbe.checks.map((check) => (
                      <li key={check.id}>
                        {check.passed ? "✓" : "!"} {check.label}：{check.detail}
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </details>
          )}
        </div>
      )}
      <button
        className="eb-scope-summary"
        type="button"
        aria-expanded={scopeOpen}
        onClick={() => setScopeOpen((current) => !current)}
      >
        <span className="scope-icon" aria-hidden="true">▦</span>
        <span className="scope-workbook">{probe.value.workbookName}</span>
        <span className="scope-divider" aria-hidden="true" />
        <strong>{scopeLabel}</strong>
        <span className="scope-chevron" aria-hidden="true">{scopeOpen ? "⌄" : "⌃"}</span>
      </button>
    </section>
  );

  return (
    <main className="probe-shell">
      <WpsAgentPanel selectedSheetIndexes={selectedIndexes} scopeDock={scopeDock} />
    </main>
  );
}
