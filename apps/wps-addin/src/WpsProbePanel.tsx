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
  /**
   * **已提交**的授权范围 —— 这次对话真正的数据边界。
   *
   * 它只在用户点「确认范围」时改变。勾选、取消勾选、切换"跟随当前/选择多个"
   * 都只改下面的草稿：否则用户每取消一次勾选，就会把正在等回答的任务一起打掉，
   * 而他本来只是想看一眼另一张表。
   */
  const [selectedIndexes, setSelectedIndexes] = useState<number[]>([]);
  const [scopeMode, setScopeMode] = useState<"auto" | "manual">("auto");
  /** 勾选中的**草稿**。打开面板时从已提交范围复制；取消编辑即丢弃。 */
  const [draftIndexes, setDraftIndexes] = useState<number[]>([]);
  const [draftMode, setDraftMode] = useState<"auto" | "manual">("auto");
  const [scopeOpen, setScopeOpen] = useState(false);
  const [sheetSearch, setSheetSearch] = useState("");

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
    /*
     * 刷新只对齐**已提交**范围，不碰草稿：用户正在弹层里勾的表不该被一次刷新
     * 悄悄改掉，那会让他"确认"出一份自己没勾过的范围。
     */
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

  /**
   * 打开范围面板：把**已提交**范围复制成草稿。
   *
   * 每次打开都从已提交范围起稿 —— 上次看了一半就关掉的草稿不会被当成授权。
   */
  const openScope = () => {
    setDraftIndexes(selectedIndexes);
    setDraftMode(scopeMode);
    setSheetSearch("");
    setScopeOpen(true);
  };

  /** 关闭而不提交：草稿丢弃，已提交授权**一字不改**。 */
  const discardScopeEdit = () => {
    setDraftIndexes(selectedIndexes);
    setDraftMode(scopeMode);
    setScopeOpen(false);
  };

  const chooseAutoScope = () => {
    setDraftMode("auto");
    setDraftIndexes(
      metadata?.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index) ?? []
    );
  };

  const chooseManualScope = () => {
    setDraftMode("manual");
    if (draftIndexes.length > 0) return;
    setDraftIndexes(
      metadata?.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index) ?? []
    );
  };

  const scopeDraftDirty =
    draftMode !== scopeMode ||
    [...draftIndexes].sort((a, b) => a - b).join(",") !==
      [...selectedIndexes].sort((a, b) => a - b).join(",");

  const toggleSheet = (index: number) => {
    setDraftIndexes((current) =>
      current.includes(index) ? current.filter((value) => value !== index) : [...current, index]
    );
  };

  /**
   * 提交草稿为**新一版授权**。
   *
   * 关键：点击时**重新读取真实宿主元数据**再解析范围 —— 不能拿弹层刚打开时
   * 的那份快照提交。弹层可能开了很久，期间用户可能已经删表、关本子或换了文件。
   */
  const commitScope = () => {
    const nextProbe = runProbe();
    if (nextProbe.status !== "ready") {
      setProbe(nextProbe);
      setMetadata(null);
      setSelectedIndexes([]);
      setDraftIndexes([]);
      setScopeOpen(false);
      return;
    }
    const nextMetadata = listWpsWorksheets(getWpsApplication());
    const available = new Set(nextMetadata.sheets.map((sheet) => sheet.index));
    const activeIndexes = nextMetadata.sheets
      .filter((sheet) => sheet.isActive)
      .map((sheet) => sheet.index);
    const nextIndexes =
      draftMode === "auto"
        ? activeIndexes
        : draftIndexes.filter((index) => available.has(index));
    setProbe(nextProbe);
    setMetadata(nextMetadata);
    setScopeMode(draftMode);
    setSelectedIndexes(nextIndexes);
    setDraftIndexes(nextIndexes);
    setScopeOpen(false);
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
  const filteredSheets = metadata?.sheets.filter((sheet) =>
    sheet.name.toLocaleLowerCase().includes(sheetSearch.trim().toLocaleLowerCase())
  ) ?? [];
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
            <button type="button" aria-label="关闭数据范围" onClick={discardScopeEdit}>×</button>
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
                className={draftMode === "auto" ? "selected" : ""}
                onClick={chooseAutoScope}
              >
                <i aria-hidden="true">⊙</i>
                <span><strong>跟随当前</strong><small>发送时自动使用正在查看的工作表</small></span>
              </button>
              <button
                type="button"
                className={draftMode === "manual" ? "selected" : ""}
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

            {draftMode === "manual" && (
              <div className="wps-manual-sheet-picker">
                <div className="wps-sheet-search-row">
                  <input
                    aria-label="搜索工作表"
                    placeholder="搜索工作表…"
                    value={sheetSearch}
                    onChange={(event) => setSheetSearch(event.target.value)}
                  />
                  <span>已选 {draftIndexes.length}/{metadata?.sheets.length ?? 0}</span>
                </div>
                <div className="sheet-picker-toolbar wps-sheet-picker-toolbar">
                  <button type="button" disabled={filteredSheets.length === 0} onClick={() => {
                    const visible = filteredSheets.map((sheet) => sheet.index);
                    setDraftIndexes((current) => [...new Set([...current, ...visible])]);
                  }}>全选搜索结果</button>
                  <button type="button" disabled={draftIndexes.length === 0} onClick={() => setDraftIndexes([])}>清空</button>
                </div>
                <div className="sheet-picker-options wps-sheet-picker-options">
                {filteredSheets.map((sheet) => (
                  <button
                    type="button"
                    className={draftIndexes.includes(sheet.index) ? "selected" : ""}
                    key={`${sheet.index}:${sheet.name}`}
                    aria-pressed={draftIndexes.includes(sheet.index)}
                    onClick={() => toggleSheet(sheet.index)}
                  >
                    <i aria-hidden="true">{draftIndexes.includes(sheet.index) ? "✓" : ""}</i>
                    <span>{sheet.name}</span>
                    {sheet.isActive && <small>当前表</small>}
                  </button>
                ))}
                {filteredSheets.length === 0 && (
                  <p className="authorization-empty">没有匹配的工作表。</p>
                )}
                {draftIndexes.length === 0 && (
                  <p className="authorization-empty">尚未选择工作表，不会读取或修改任何数据。</p>
                )}
                </div>
              </div>
            )}
          </div>
          </div>
          <div className="wps-scope-actions">
            <button type="button" className="scope-refresh" onClick={refresh}>↻ 刷新</button>
            <button
              type="button"
              className="scope-confirm"
              disabled={draftMode === "manual" && draftIndexes.length === 0}
              onClick={commitScope}
            >
              确认范围
            </button>
            {scopeDraftDirty && (
              <button type="button" className="scope-discard" onClick={discardScopeEdit}>
                放弃修改
              </button>
            )}
          </div>
          {scopeDraftDirty && (
            <p className="wps-scope-draft-hint" role="status">
              勾选还只是草稿：点「确认范围」才会把它作为这次对话的数据边界提交；
              放弃修改不会改变已确认的范围。
            </p>
          )}
        </section>
      )}
      <div className="scope-entry-row">
        <button
          className={`scope-trigger ${scopeOpen ? "open" : ""}`}
          type="button"
          aria-expanded={scopeOpen}
          onClick={() => (scopeOpen ? discardScopeEdit() : openScope())}
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
