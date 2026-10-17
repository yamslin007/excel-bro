import { useCallback, useEffect, useRef, useState } from "react";
import type { AuthorizedTarget } from "@excel-bro/core/authorizedTarget";
import { folderRootIdentity } from "@excel-bro/core/scopeFingerprint";
import {
  addedFolderSheetKeys,
  canPreserveFolderAuthorization,
  folderScopeDraftDirty,
  folderSelectionsFromKeys,
  folderSheetKey,
  isEBSystemSheet,
  selectableFolderSheetKeys
} from "@excel-bro/core/folderScope";
import type { FolderCatalog } from "../../excel-addin/src/contracts";
import {
  checkFolderAdmission,
  createFolderSnapshot,
  scanFolderByPath,
  selectFolder
} from "../../excel-addin/src/api";
import { inspectWpsScopeAdmission } from "./wpsAuthorizedTarget";
import { useScopeAdmission } from "../../excel-addin/src/hooks/useScopeAdmission";
import { ScopeAdmissionSummary } from "../../excel-addin/src/components/ScopeAdmissionStatus";
import { admissionItem } from "@excel-bro/core/scopeAdmission";
import { reportDevProbe } from "./devProbe";
import { brandLogoUrl } from "../../excel-addin/src/brandLogo";
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
    return { status: "ready", value: readWpsHostProbe(getWpsApplication(), { includeSelectionPreview: false }) };
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
   * 它只在用户点「确认」时改变。勾选、取消勾选、切换"跟随当前/选择多个"
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
  const [scopeConfirming, setScopeConfirming] = useState(false);
  const [scopeError, setScopeError] = useState("");
  const [scopeCatalogAvailable, setScopeCatalogAvailable] = useState(true);
  /** 每次点「刷新」自增：作为图标的 `key`，换 key 即重挂载 → 旋转动画重放一次。 */
  const [refreshSpin, setRefreshSpin] = useState(0);

  /**
   * 数据来源。`workbook` = 当前打开的工作簿（原有能力）；`folder` = 工作目录。
   *
   * 工作目录的数据读写全在服务端（openpyxl），**不经 WPS 宿主**，所以 WPS 不需要
   * 任何"打开磁盘文件"的宿主能力。这也是这张卡能在 WPS 上放开的原因。
   */
  const [scopeSource, setScopeSource] = useState<"workbook" | "folder">("workbook");
  /** 来源也是草稿：点卡片只改草稿，点「确认」才成为这次的授权来源。 */
  const [draftSource, setDraftSource] = useState<"workbook" | "folder">("workbook");
  const [folderCatalog, setFolderCatalog] = useState<FolderCatalog | null>(null);
  /**
   * **已提交**的工作目录授权：根身份 + 勾选键。
   *
   * 根身份来自 `folderRootIdentity`（盘符小写、UNC 保留、无绝对根返回 null）。
   * 与 `folderCatalog` 分开存：目录被"更换"成另一个根时，这份授权必须一起作废，
   * 而不是跟着新目录继续生效。
   */
  const [folderScope, setFolderScope] = useState<{ root: string; keys: string[] } | null>(null);
  /** 勾选**草稿**（文件 id + 表名合成的键）。取消即回退到 `folderScope.keys`。 */
  const [draftFolderKeys, setDraftFolderKeys] = useState<string[]>([]);
  const [folderBusy, setFolderBusy] = useState(false);
  const [folderError, setFolderError] = useState("");
  const [folderPickerView, setFolderPickerView] = useState<"files" | "sheets">("files");
  const [activeFolderFileId, setActiveFolderFileId] = useState<string | null>(null);
  const [folderSheetSearch, setFolderSheetSearch] = useState("");
  const [showSelectedFolderSheets, setShowSelectedFolderSheets] = useState(false);
  const [recentlyAddedFolderSheetKeys, setRecentlyAddedFolderSheetKeys] = useState<string[]>([]);
  const [folderRefreshNotice, setFolderRefreshNotice] = useState("");
  const folderRequestLock = useRef(false);
  const scopeConfirmation = useRef<((target: AuthorizedTarget) => Promise<boolean>) | null>(null);
  const scopeConfirmLock = useRef(false);
  const registerScopeConfirmation = useCallback((handler: ((target: AuthorizedTarget) => Promise<boolean>) | null) => {
    scopeConfirmation.current = handler;
  }, []);
  /**
   * 承认检查。**同一个引擎**同时服务两种来源：
   *  - workbook：`inspectWpsScopeAdmission` 向 WPS 宿主逐表核实；
   *  - folder：`checkFolderAdmission` 由服务端在本地文件系统上逐表核实。
   *
   * 检查永远**不构成授权** —— 它只回答"现在能不能分析"，提交仍要用户点「确认」。
   * 两种来源的 id 口径不同（工作表序号 vs 文件夹键），所以按来源分别给。
   */
  const folderAdmissionIds = draftFolderKeys;
  const workbookAdmissionIds = draftIndexes.map(String);
  const admissionIds = draftSource === "folder" ? folderAdmissionIds : workbookAdmissionIds;
  const admission = useScopeAdmission({
    enabled: scopeOpen && (draftSource === "folder" ? folderCatalog !== null : metadata !== null),
    scopeKey: draftSource === "folder"
      ? JSON.stringify(["folder", folderCatalog?.sessionId, folderCatalog?.folderPath, folderBusy])
      : JSON.stringify([metadata?.fullName, metadata?.name, metadata?.sheets.map((s) => [s.index, s.name]), draftMode, refreshSpin, scopeCatalogAvailable]),
    ids: admissionIds,
    inspect: async () => {
      if (draftSource === "folder") {
        const catalog = folderCatalog;
        if (!catalog) return { items: [] };
        const selections = folderSelectionsFromKeys(catalog, draftFolderKeys);
        if (selections.length === 0) return { items: [] };
        return checkFolderAdmission(catalog.sessionId, selections);
      }
      return !scopeCatalogAvailable
        ? { items: draftIndexes.map((index) => admissionItem(String(index), "identity_unverified")) }
        : inspectWpsScopeAdmission(draftIndexes, {
        name: metadata!.name,
        fileKey: metadata!.fullName && /^(?:[a-z]:[\\/]|\\\\)/i.test(metadata!.fullName) ? metadata!.fullName : null,
        // Unsaved WPS identity is established by the live host adapter, not manufactured here.
        fileRef: getWpsApplication().ActiveWorkbook,
        sheets: metadata!.sheets.map((sheet) => ({ id: String(sheet.index), name: sheet.name })),
        selected: draftIndexes.map(String)
      });
    }
  });

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
    if (scopeConfirmLock.current) return;
    // 先记一次点击反馈（放在最前面，早退分支同样要转）。
    setRefreshSpin((value) => value + 1);
    setScopeError("");
    const nextProbe = runProbe();
    if (nextProbe.status !== "ready") {
      setScopeCatalogAvailable(false);
      if (scopeOpen) setScopeError("暂时无法读取目标，请刷新后重试；未变更授权。");
      else setProbe(nextProbe);
      return;
    }
    try {
      const nextMetadata = listWpsWorksheets(getWpsApplication());
      setScopeCatalogAvailable(true);
      setProbe(nextProbe);
      setMetadata(nextMetadata);
      // Only the automatic draft follows the active sheet. Refresh is not
      // consent: committed indexes and manual picks, including missing ones,
      // stay intact until explicit confirmation.
      if (scopeOpen && draftMode === "auto") {
        setDraftIndexes(nextMetadata.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index));
      }
    } catch {
      setScopeCatalogAvailable(false);
      setScopeError("暂时无法读取目标，请刷新后重试；未变更授权。");
    }
  };

  /** 扫描结果里可勾选（排除 EB 系统遗留表）的表数，用于各类计数文案。 */
  const selectableSheetCount = (catalog: FolderCatalog): number =>
    selectableFolderSheetKeys(catalog).length;

  /**
   * 选择/更换工作目录。
   *
   * 目录框是**服务端**用系统原生对话框弹出来的（`POST /api/folders/select`），
   * 所以 WPS 宿主不需要任何"选文件夹"能力。用户在目录框里点取消 ⇒ 后端返回 null
   * ⇒ 这里**什么都不改**（不是"选了个空目录"）。
   *
   * 换根 = 换目标：旧勾选与旧授权一起作废，绝不跨目录沿用（目录 A 里那句
   * "已选 3 张表"对目录 B 毫无意义）。
   */
  const chooseWorkDirectory = async () => {
    if (folderRequestLock.current || folderBusy) return;
    folderRequestLock.current = true;
    setFolderBusy(true);
    setFolderError("");
    setFolderRefreshNotice("");
    setRecentlyAddedFolderSheetKeys([]);
    try {
      const catalog = await selectFolder();
      if (!catalog) return;
      const nextRoot = folderRootIdentity(catalog.folderPath);
      const sameRoot = nextRoot !== null && nextRoot === folderRootIdentity(folderCatalog?.folderPath);
      setFolderCatalog(catalog);
      setActiveFolderFileId((catalog.files.find((file) => !file.error) ?? catalog.files[0])?.id ?? null);
      setFolderPickerView("files");
      setFolderSheetSearch("");
      setShowSelectedFolderSheets(false);
      if (!sameRoot) {
        setDraftFolderKeys([]);
        setFolderScope(null);
      }
      setFolderRefreshNotice(`已读取 ${catalog.files.length} 个文件、${selectableSheetCount(catalog)} 张工作表`);
    } catch (reason) {
      setFolderError(reason instanceof Error ? reason.message : "读取工作目录失败");
    } finally {
      folderRequestLock.current = false;
      setFolderBusy(false);
    }
  };

  /**
   * 刷新工作目录。
   *
   * 刷新**不是同意**：只有「同根 + 同会话 + 至少勾一张 + 每张都还在」才保留原确认，
   * 判据在共享层 `canPreserveFolderAuthorization`。新增的表不会自动勾上 —— 自动勾
   * 等于替用户扩权。被删/改名的表会从勾选里消失（那个勾选已经不是事实了）。
   */
  const refreshWorkDirectory = async (options: { quiet?: boolean; keys?: readonly string[] } = {}) => {
    if (!folderCatalog || folderRequestLock.current || folderBusy) return;
    // 要核实"还算不算同一件授权"的那一组键。默认是当前草稿；打开面板时的静默刷新
    // 显式传**已提交**的键 —— 那一刻草稿刚被重置，闭包里的 `draftFolderKeys` 还是
    // 上一轮渲染的值，读它会核错对象。
    const keys = options.keys ?? draftFolderKeys;
    folderRequestLock.current = true;
    setFolderBusy(true);
    setFolderError("");
    setFolderRefreshNotice("");
    try {
      const refreshed = await scanFolderByPath(folderCatalog.folderPath);
      const addedKeys = addedFolderSheetKeys(folderCatalog, refreshed);
      const preserveAuthorization = canPreserveFolderAuthorization(
        folderCatalog,
        refreshed,
        keys
      );
      setFolderCatalog(refreshed);
      setDraftFolderKeys((current) =>
        current.filter((key) => folderSelectionsFromKeys(refreshed, [key]).length > 0)
      );
      if (!preserveAuthorization) setFolderScope(null);
      setRecentlyAddedFolderSheetKeys(addedKeys);
      if (addedKeys.length > 0) {
        const separator = addedKeys[0].indexOf("\u0000");
        if (separator > 0) {
          setActiveFolderFileId(addedKeys[0].slice(0, separator));
          setFolderPickerView("sheets");
        }
        setFolderRefreshNotice(`发现 ${addedKeys.length} 张新增工作表`);
        return;
      }
      if (options.quiet) return;
      setFolderRefreshNotice(
        `已刷新，共 ${refreshed.files.length} 个文件、${selectableSheetCount(refreshed)} 张工作表`
      );
    } catch (reason) {
      setFolderError(reason instanceof Error ? reason.message : "刷新工作目录失败");
    } finally {
      folderRequestLock.current = false;
      setFolderBusy(false);
    }
  };

  const toggleFolderSheet = (fileId: string, sheetName: string) => {
    const key = folderSheetKey(fileId, sheetName);
    setDraftFolderKeys((current) =>
      current.includes(key) ? current.filter((item) => item !== key) : [...current, key]
    );
  };

  /** 全选/清空都只动**本文件**，也只动草稿；已提交授权要等「确认」。 */
  const setFolderFileSelection = (fileId: string, selected: boolean) => {
    const file = folderCatalog?.files.find((candidate) => candidate.id === fileId);
    if (!file) return;
    const keys = file.worksheets
      .filter((sheet) => !isEBSystemSheet(sheet.name))
      .map((sheet) => folderSheetKey(fileId, sheet.name));
    setDraftFolderKeys((current) => selected
      ? [...new Set([...current, ...keys])]
      : current.filter((key) => !keys.includes(key)));
  };

  /**
   * 清空**整个工作目录**的草稿勾选（不是"本文件"）。
   *
   * 与「选择多个」的「清空」同一件事：只清草稿，已提交的 `folderScope` 一字不动，
   * 要换成更小的范围仍需按「确认」。
   */
  const clearFolderSelection = () => {
    if (draftFolderKeys.length === 0) return;
    setDraftFolderKeys([]);
    setRecentlyAddedFolderSheetKeys([]);
    if (showSelectedFolderSheets) setShowSelectedFolderSheets(false);
  };

  /**
   * 打开范围面板：把**已提交**范围复制成草稿。
   *
   * 每次打开都从已提交范围起稿 —— 上次看了一半就关掉的草稿不会被当成授权。
   */
  const openScope = () => {
    if (scopeConfirmLock.current) return;
    setScopeError("");
    try {
      const current = listWpsWorksheets(getWpsApplication());
      setScopeCatalogAvailable(true);
      const changedFile = metadata?.fullName !== current.fullName || metadata?.name !== current.name;
      setMetadata(current);
      const nextProbe = runProbe();
      if (nextProbe.status === "ready") setProbe(nextProbe);
      setDraftIndexes(changedFile || scopeMode === "auto"
        ? current.sheets.filter((s) => s.isActive).map((s) => s.index) : selectedIndexes);
    } catch {
      setScopeCatalogAvailable(false);
      setDraftIndexes(selectedIndexes);
      setScopeError("暂时无法核实目标，请刷新后重试。");
    }
    setDraftMode(scopeMode);
    // 来源与工作目录勾选也一起从**已提交**状态起稿。
    setDraftSource(scopeSource);
    setDraftFolderKeys(folderScope?.keys ?? []);
    setFolderError("");
    setSheetSearch("");
    setFolderSheetSearch("");
    setScopeOpen(true);
    // 打开时静默刷新一次工作目录（对齐 Excel：目录里的文件可能已经改名/新增）。
    // 刷新不改授权，所以这次刷新不作任何提示；用户点「刷新」才给反馈。
    if (scopeSource === "folder" && folderCatalog) {
      void refreshWorkDirectory({ quiet: true, keys: folderScope?.keys ?? [] });
    }
  };

  /** 关闭而不提交：草稿丢弃，已提交授权**一字不改**。 */
  const discardScopeEdit = () => {
    if (scopeConfirmLock.current) return;
    setDraftIndexes(selectedIndexes);
    setDraftMode(scopeMode);
    setDraftSource(scopeSource);
    setDraftFolderKeys(folderScope?.keys ?? []);
    setFolderError("");
    setScopeOpen(false);
  };

  const chooseAutoScope = () => {
    setDraftSource("workbook");
    setDraftMode("auto");
    setDraftIndexes(
      metadata?.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index) ?? []
    );
  };

  const chooseManualScope = () => {
    setDraftSource("workbook");
    setDraftMode("manual");
    if (draftIndexes.length > 0) return;
    setDraftIndexes(
      metadata?.sheets.filter((sheet) => sheet.isActive).map((sheet) => sheet.index) ?? []
    );
  };

  /**
   * 切到「工作目录」。
   *
   * 只是把来源选成 folder 并展开目录选择区 —— **不产生任何授权**：
   * 候选目录登记完还要逐张勾选，「确认」才提交。已有目录则静默刷新一次。
   */
  const chooseFolderSource = () => {
    setDraftSource("folder");
    setFolderPickerView("files");
    setFolderSheetSearch("");
    if (folderCatalog && !folderBusy) void refreshWorkDirectory({ quiet: true });
  };

  const committedFolderKeys = folderScope?.keys ?? [];
  const scopeDraftDirty = draftSource !== scopeSource ||
    (draftSource === "folder"
      ? folderScopeDraftDirty(draftFolderKeys, committedFolderKeys)
      : draftMode !== scopeMode ||
        [...draftIndexes].sort((a, b) => a - b).join(",") !==
          [...selectedIndexes].sort((a, b) => a - b).join(","));

  const toggleSheet = (index: number) => {
    setDraftIndexes((current) =>
      current.includes(index) ? current.filter((value) => value !== index) : [...current, index]
    );
  };

  /**
   * 提交草稿为**新一版授权**。
   *
   * 两条路都先重核一遍检查结果：**看得见的绿勾不是永久许可**。
   *  - workbook：点击时重新读取真实宿主元数据再解析范围 —— 不能拿弹层刚打开时的
   *    那份快照提交。弹层可能开了很久，期间用户可能已经删表、关本子或换了文件。
   *  - folder：点击时重新用**服务端**按当前勾选算一次候选范围。数据读写都在服务端，
   *    这里拿到的是"这次授权到底覆盖了哪些表"的回执，不是内存里的旧勾选。
   */
  const commitScope = async () => {
    if (scopeConfirmLock.current || !admission.canCommit) return;
    scopeConfirmLock.current = true;
    setScopeConfirming(true);
    setScopeError("");
    try {
      const checked = await admission.verify();
      if (!checked) return;
      if (draftSource === "folder") {
        const catalog = folderCatalog;
        if (!catalog || draftFolderKeys.length === 0) return;
        const selections = folderSelectionsFromKeys(catalog, draftFolderKeys);
        if (selections.length === 0) {
          setScopeError("所选工作表已不存在，请刷新后重新选择；未变更授权。");
          return;
        }
        const root = folderRootIdentity(catalog.folderPath);
        if (!root) {
          setScopeError("无法核实工作目录身份，请重新选择目录；未变更授权。");
          return;
        }
        const snapshot = await createFolderSnapshot(catalog.sessionId, selections);
        if (snapshot.worksheets.length !== draftFolderKeys.length) {
          // 服务端按这些选择算出来的表数对不上 ⇒ 期间有表被删/改名。不提交半截范围。
          setScopeError("工作目录内容在确认期间发生变化，请刷新后重新选择；未变更授权。");
          return;
        }
        setFolderScope({ root, keys: [...draftFolderKeys] });
        setScopeSource("folder");
        setScopeOpen(false);
        return;
      }
      if (!checked.target) return;
      const nextProbe = runProbe();
      if (nextProbe.status !== "ready") {
        setScopeError("暂时无法读取目标身份，请刷新后重试；未变更授权。");
        return;
      }
      const nextMetadata = listWpsWorksheets(getWpsApplication());
      const nextIndexes = [...draftIndexes];
      if (nextIndexes.some((index) => !nextMetadata.sheets.some((s) => s.index === index)) ||
          (draftMode === "auto" && !nextMetadata.sheets.some((s) => s.isActive && nextIndexes.includes(s.index)))) {
        setScopeError("工作表范围在确认期间发生变化，请刷新后重新选择；未变更授权。");
        return;
      }
      const target = checked.target;
      if (!scopeConfirmation.current || !(await scopeConfirmation.current(target))) {
        setScopeError("暂时无法核实文件与工作表，请刷新后重试；未变更授权。");
        return;
      }
      setProbe(nextProbe);
      setMetadata(nextMetadata);
      setScopeMode(draftMode);
      setScopeSource("workbook");
      setSelectedIndexes(nextIndexes);
      setDraftIndexes(nextIndexes);
      setScopeOpen(false);
    } catch (reason) {
      setScopeError(
        reason instanceof Error
          ? reason.message
          : "暂时无法读取目标身份，请刷新后重试；未变更授权。"
      );
    } finally {
      scopeConfirmLock.current = false;
      setScopeConfirming(false);
    }
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
        <img className="probe-logo" src={brandLogoUrl} alt="" />
        <div><strong>Excel Bro</strong><p>正在连接 WPS 表格…</p></div>
      </main>
    );
  }

  if (probe.status === "error") {
    return (
      <main className="connection-shell error">
        <img className="probe-logo" src={brandLogoUrl} alt="" />
        <div><strong>尚未连接 WPS 表格</strong><p>{probe.message}</p></div>
        <button type="button" onClick={refresh}>重新检测</button>
      </main>
    );
  }

  const selectedSheets = metadata?.sheets.filter((sheet) => selectedIndexes.includes(sheet.index)) ?? [];
  // Both labels must describe the same metadata snapshot, not independent
  // probe/committed-selection snapshots. This display does not grant access.
  const currentSheetName = metadata?.sheets.find((sheet) => sheet.isActive)?.name;
  const currentFileName = metadata?.name ?? probe.value.workbookName;
  const filteredSheets = metadata?.sheets.filter((sheet) =>
    sheet.name.toLocaleLowerCase().includes(sheetSearch.trim().toLocaleLowerCase())
  ) ?? [];
  const workbookScopeLabel = scopeMode === "auto"
    ? currentSheetName ? `当前表 ${currentSheetName}` : "当前工作表待核实"
    : selectedSheets.length === 0
    ? "尚未选择工作表"
      : selectedSheets.length === 1
      ? `${selectedSheets[0].isActive ? "当前表" : "已选"} ${selectedSheets[0].name}`
      : `已固定 ${selectedSheets.length} 个工作表`;
  /**
   * 触发器上那句范围说明。
   *
   * 工作目录模式下**只**说工作目录 —— 这里要是顺带写上一句当前工作簿名，
   * 等于在告诉用户"这次用的是这个工作簿"，而事实恰恰相反。
   */
  const scopeLabel = scopeSource === "folder"
    ? folderCatalog
      ? `工作目录 ${folderCatalog.folderName} · 已选 ${committedFolderKeys.length} 张表`
      : "设置工作目录"
    : workbookScopeLabel;

  /** 勾选草稿按当前扫描结果算出的「文件 → 表」投影，计数与提交都用它。 */
  const draftFolderSelections = folderCatalog
    ? folderSelectionsFromKeys(folderCatalog, draftFolderKeys)
    : [];
  const draftFolderFileCount = draftFolderSelections.length;
  const totalSelectableFolderSheets = folderCatalog ? selectableSheetCount(folderCatalog) : 0;
  const folderSelectionCounts = new Map<string, number>(
    draftFolderSelections.map((selection) => [selection.fileId, selection.sheets.length])
  );
  const activeFolderFile =
    folderCatalog?.files.find((file) => file.id === activeFolderFileId) ?? null;
  const activeFolderSheets = activeFolderFile
    ? activeFolderFile.worksheets.filter((sheet) => !isEBSystemSheet(sheet.name))
    : [];
  const filteredFolderSheets = activeFolderSheets.filter((sheet) =>
    sheet.name.toLocaleLowerCase().includes(folderSheetSearch.trim().toLocaleLowerCase())
  );
  const visibleFolderFiles = folderCatalog
    ? (showSelectedFolderSheets
        ? folderCatalog.files.filter((file) => (folderSelectionCounts.get(file.id) ?? 0) > 0)
        : folderCatalog.files)
    : [];

  const scopeDock = (
    <section className={`wps-scope-dock ${scopeOpen ? "open" : ""}`}>
      {scopeOpen && (
        <section className="scope-popover wps-scope-popover" aria-label="选择数据范围">
          <div className="scope-popover-scroll">
          <div className="scope-popover-header">
            <div>
              <strong>这次对话使用哪些数据？</strong>
              <span>
                {draftSource === "folder"
                  ? "批量处理本地工作簿；数据由本机服务端读取"
                  : "默认跟随你当前正在查看的工作表"}
              </span>
            </div>
            <button type="button" aria-label="关闭数据范围" onClick={discardScopeEdit}>×</button>
          </div>
          <div className="wps-current-file-card">
            <span aria-hidden="true">XLSX</span>
            <div>
              <small>当前文件</small>
              <strong>{currentFileName}</strong>
              <em>{currentSheetName ? `当前工作表 ${currentSheetName}` : "当前工作表待核实"}</em>
            </div>
          </div>
          <div className="scope-controls">
            <fieldset className="scope-selection-fields" disabled={scopeConfirming}>
            <div className="scope-mode-list">
              <button
                type="button"
                className={draftSource === "workbook" && draftMode === "auto" ? "selected" : ""}
                onClick={chooseAutoScope}
              >
                <i aria-hidden="true">⊙</i>
                <span><strong>跟随当前</strong><small>发送时自动使用正在查看的工作表</small></span>
              </button>
              <button
                type="button"
                className={draftSource === "workbook" && draftMode === "manual" ? "selected" : ""}
                onClick={chooseManualScope}
              >
                <i aria-hidden="true">☷</i>
                <span><strong>选择多个</strong><small>用于跨表查询、比较或汇总</small></span>
              </button>
              {/* 面板常挂载、由 `.open` 类驱动收放：这样**收起也有动画**
                  （条件渲染会在切走时直接卸载，收缩无从播放）。收起态靠
                  CSS 的 0fr + visibility:hidden，不进 Tab 序。 */}
              <div className={`scope-picker-reveal${draftSource === "workbook" && draftMode === "manual" ? " open" : ""}`}>
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
                      {/*
                        行内只有「勾选框 + 工作表名（+ 当前表）」。授权状态文案
                        （「可分析 · 修改仍需预览确认」等）**不进这一行**：它会把单行
                        撑成两行、行高跟着跳，也不是这一行要表达的东西。整次选择的
                        授权状态只有一个去处 —— 列表下方的 ScopeAdmissionSummary。
                        勾选框复用共享的 `.sheet-checkbox`（与 Excel 同一份声明），
                        免得两端各写一套、各自漂移。
                      */}
                      <i className="sheet-checkbox" aria-hidden="true">
                        {draftIndexes.includes(sheet.index) ? "✓" : ""}
                      </i>
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
              </div>
              <button
                type="button"
                className={draftSource === "folder" ? "selected" : ""}
                onClick={chooseFolderSource}
              >
                <i aria-hidden="true">⌑</i>
                <span><strong>工作目录</strong><small>批量处理多个本地工作簿</small></span>
              </button>
            </div>
            </fieldset>
            {draftSource === "workbook" && (
              <ScopeAdmissionSummary items={admission.report?.items ?? []} count={draftIndexes.length}
                checking={admission.checking} />
            )}
          </div>

          {/*
            工作目录选择区。结构、类名与文案全部对齐 Excel 的 `.folder-*`
            （WPS 已经加载了 Excel 的样式表，所以同样的类名就是同一套视觉）。
            数据读写都在服务端，这里只负责把"允许处理哪些表"勾清楚。
          */}
          {draftSource === "folder" && (
            <div className="folder-scope-picker">
              {!folderCatalog ? (
                <section className="work-directory-empty" aria-label="设置工作目录">
                  <i aria-hidden="true">⌂</i>
                  <div>
                    <strong>设置工作目录</strong>
                    <span>
                      选择存放待处理 Excel 文件的文件夹。这里仅确定候选范围，稍后仍需勾选允许处理的工作表。
                    </span>
                    <small>本机服务端会弹出系统目录选择框；取消不会改变当前范围。</small>
                  </div>
                  <button
                    type="button"
                    className="browse-folder-button"
                    disabled={folderBusy}
                    onClick={() => void chooseWorkDirectory()}
                  >
                    {folderBusy ? "正在读取…" : "选择工作目录"}
                  </button>
                </section>
              ) : (
                <div className="folder-toolbar">
                  <button
                    type="button"
                    className="browse-folder-button"
                    disabled={folderBusy}
                    onClick={() => void chooseWorkDirectory()}
                  >
                    {`更换工作目录 · ${folderCatalog.folderName}`}
                  </button>
                  <button
                    type="button"
                    className="folder-refresh-button"
                    disabled={folderBusy}
                    onClick={() => void refreshWorkDirectory()}
                    title="重新扫描文件和工作表"
                  >
                    ↻ {folderBusy ? "刷新中" : "刷新"}
                  </button>
                </div>
              )}

              {folderCatalog && (
                <>
                  <p className="folder-save-hint">
                    在其他工作簿新增或改名后，请先保存文件再刷新。
                  </p>
                  {folderRefreshNotice && (
                    <div
                      className={`folder-refresh-notice${
                        recentlyAddedFolderSheetKeys.length > 0 ? " has-new" : ""
                      }`}
                      role="status"
                    >
                      {folderRefreshNotice}
                    </div>
                  )}
                  <div className="folder-selection-filter">
                    <span>
                      {folderCatalog.files.length} 个文件 · {totalSelectableFolderSheets} 张表
                    </span>
                    <div className="folder-selection-filter-actions">
                      <button
                        type="button"
                        disabled={draftFolderKeys.length === 0}
                        onClick={clearFolderSelection}
                      >
                        清空
                      </button>
                      <button
                        type="button"
                        className={showSelectedFolderSheets ? "selected" : ""}
                        disabled={!showSelectedFolderSheets && draftFolderKeys.length === 0}
                        onClick={() => setShowSelectedFolderSheets((current) => !current)}
                        aria-pressed={showSelectedFolderSheets}
                      >
                        {showSelectedFolderSheets ? "显示全部" : "只看已选"}
                      </button>
                    </div>
                  </div>
                  <div className={`folder-master-detail folder-picker-view-${folderPickerView}`}>
                    <div className="folder-file-list" role="list" aria-label="工作簿列表">
                      {visibleFolderFiles.map((file) => {
                        const selectableCount = file.worksheets.filter(
                          (sheet) => !isEBSystemSheet(sheet.name)
                        ).length;
                        const selectedCount = folderSelectionCounts.get(file.id) ?? 0;
                        const active = activeFolderFile?.id === file.id;
                        return (
                          <button
                            type="button"
                            className={`folder-file-option${active ? " active" : ""}${file.error ? " error" : ""}`}
                            key={file.id}
                            onClick={() => {
                              setActiveFolderFileId(file.id);
                              setFolderPickerView("sheets");
                              setFolderSheetSearch("");
                            }}
                            aria-pressed={active}
                            role="listitem"
                          >
                            <i aria-hidden="true">X</i>
                            <span>
                              <strong title={file.relativePath}>{file.relativePath}</strong>
                              <small>{file.error ? "文件暂时无法读取" : `${selectableCount} 张工作表`}</small>
                            </span>
                            <b>{selectedCount}/{selectableCount}</b>
                          </button>
                        );
                      })}
                      {visibleFolderFiles.length === 0 && (
                        <p className="folder-empty">
                          {showSelectedFolderSheets ? "还没有勾选工作表" : "文件夹中没有可用工作簿"}
                        </p>
                      )}
                    </div>

                    {activeFolderFile && (
                      <section
                        className="folder-sheet-detail"
                        aria-label={`选择 ${activeFolderFile.relativePath} 中的工作表`}
                      >
                        <div className="folder-sheet-header">
                          <button
                            type="button"
                            className="folder-back-button"
                            onClick={() => setFolderPickerView("files")}
                          >
                            ← 文件
                          </button>
                          <span>
                            <strong title={activeFolderFile.relativePath}>
                              {activeFolderFile.relativePath}
                            </strong>
                            <small>
                              已选 {folderSelectionCounts.get(activeFolderFile.id) ?? 0}/
                              {activeFolderSheets.length}
                            </small>
                          </span>
                        </div>

                        {activeFolderFile.error ? (
                          <p className="file-error">{activeFolderFile.error}</p>
                        ) : (
                          <>
                            <div className="sheet-search-row folder-sheet-search">
                              <input
                                value={folderSheetSearch}
                                onChange={(event) => setFolderSheetSearch(event.target.value)}
                                placeholder="搜索当前文件的工作表..."
                                aria-label="搜索当前文件的工作表"
                              />
                              <span>{filteredFolderSheets.length} 项</span>
                            </div>
                            <div className="sheet-picker-toolbar">
                              <button
                                type="button"
                                onClick={() => setFolderFileSelection(activeFolderFile.id, true)}
                              >
                                全选本文件
                              </button>
                              <button
                                type="button"
                                onClick={() => setFolderFileSelection(activeFolderFile.id, false)}
                              >
                                清空本文件
                              </button>
                            </div>
                            <div className="folder-sheet-list">
                              {filteredFolderSheets.map((sheet) => {
                                const key = folderSheetKey(activeFolderFile.id, sheet.name);
                                const selected = draftFolderKeys.includes(key);
                                return (
                                  <button
                                    type="button"
                                    key={sheet.name}
                                    className={selected ? "selected" : ""}
                                    onClick={() => toggleFolderSheet(activeFolderFile.id, sheet.name)}
                                    aria-pressed={selected}
                                  >
                                    <i className="sheet-checkbox" aria-hidden="true">
                                      {selected ? "✓" : ""}
                                    </i>
                                    <span>
                                      <strong>{sheet.name}</strong>
                                      <small>
                                        {sheet.rowCount} 行 · {sheet.columnCount} 列
                                      </small>
                                    </span>
                                    {recentlyAddedFolderSheetKeys.includes(key) && <em>新增</em>}
                                  </button>
                                );
                              })}
                              {filteredFolderSheets.length === 0 && (
                                <p className="folder-empty">
                                  {showSelectedFolderSheets
                                    ? "当前文件没有已选工作表"
                                    : "没有匹配的工作表"}
                                </p>
                              )}
                            </div>
                          </>
                        )}
                      </section>
                    )}
                  </div>
                </>
              )}
            </div>
          )}
          </div>
          {/*
            工作目录侧的固定页脚 —— 结构、类名与内部顺序都与 Excel 的
            `.folder-selection-footer` **逐字对齐**：计数在左，按钮在右
            （共享样式表里 `.folder-selection-footer .scope-confirm { margin-left: auto }`
            把确认贴到右边），承认汇总落在页脚**下方**。

            页脚的"层"和 Excel 一样：它是 `.scope-popover-scroll` 的**兄弟**，不在滚动容器里。
            放进滚动容器它就只是"内容末尾的一行"，内容一滚就跟着滚走 —— 这条判据在
            `wpsWorkDirectoryScope.test.tsx`（有目录夹具的那份）里钉。
          */}
          {draftSource === "folder" && folderCatalog && (
            <div className="folder-selection-footer">
              <div className="folder-selection-summary">
                <strong>
                  已选 {draftFolderFileCount} 个文件 · {draftFolderKeys.length} 张表
                </strong>
                <details className="folder-selection-help">
                  <summary aria-label="查看工作表选择保留说明" title="查看说明">ⓘ</summary>
                  <span role="note">切换文件不会丢失已勾选的工作表</span>
                </details>
              </div>
              <button
                type="button"
                className="scope-confirm"
                disabled={scopeConfirming || !admission.canCommit || draftFolderKeys.length === 0}
                onClick={() => void commitScope()}
              >
                {scopeConfirming ? "核对中..." : "确认"}
              </button>
              {/* 「取消」只在草稿与已提交授权不一致时出现 —— 两端（Excel / WPS）文件夹页脚
                  到此**逐字同形**：计数 + ⓘ → 确认 →（可选）取消。出口语义与工作簿动作条那颗
                  是同一条规则（`discardScopeEdit`）：草稿回到已提交授权，已提交一字不改。 */}
              {scopeDraftDirty && (
                <button type="button" className="scope-discard" onClick={discardScopeEdit}>
                  取消
                </button>
              )}
            </div>
          )}
          {draftSource === "folder" && (
            <ScopeAdmissionSummary items={admission.report?.items ?? []}
              count={draftFolderKeys.length} checking={admission.checking} />
          )}
          {/* 工作簿模式的动作条（与 Excel 的 `.scope-action-buttons` 同一层、同一份样式声明）；
              工作目录模式不用它 —— 那一边的固定页脚就是上面的 `.folder-selection-footer`，
              刷新在目录工具条上。 */}
          {draftSource === "workbook" && (
            <div className="wps-scope-actions">
              <button type="button" className="scope-refresh" disabled={scopeConfirming} onClick={refresh}>
                <i
                  className={`scope-refresh-glyph${refreshSpin ? " spinning" : ""}`}
                  key={refreshSpin}
                  aria-hidden="true"
                >↻</i>
                {" "}刷新
              </button>
              <button
                type="button"
                className="scope-confirm"
                disabled={scopeConfirming || !admission.canCommit}
                onClick={() => void commitScope()}
              >
                {scopeConfirming ? "核对中..." : "确认"}
              </button>
              {scopeDraftDirty && (
                <button type="button" className="scope-discard" onClick={discardScopeEdit}>
                  取消
                </button>
              )}
            </div>
          )}
          {folderError && <p role="alert" className="authorized-target-error">{folderError}</p>}
          {scopeError && <p role="alert" className="authorized-target-error">{scopeError}</p>}
        </section>
      )}
      <div className="scope-entry-row">
        <button
          className={`scope-trigger ${scopeOpen ? "open" : ""}`}
          type="button"
          aria-expanded={scopeOpen}
          onClick={() => (scopeOpen ? discardScopeEdit() : openScope())}
        >
          <i aria-hidden="true">
            <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
              <path d="M19.2 12.1l-6.9 6.9a4.6 4.6 0 0 1-6.5-6.5l6.9-6.9a3 3 0 0 1 4.3 4.3l-6.9 6.9a1.4 1.4 0 0 1-2-2l6.2-6.2" />
            </svg>
          </i>
          <span>
            {/* 工作目录模式下**不显示当前工作簿名**：这一行是"这次用哪些数据"，
                写上工作簿名就等于在说"这次用它"，而工作目录模式恰恰不用它。 */}
            {scopeSource === "workbook" && (
              <>
                <small title={currentFileName}>{currentFileName}</small>
                <span className="scope-sep" aria-hidden="true">|</span>
              </>
            )}
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
        scopeMode={scopeMode}
        /*
          工作目录模式**不把工作簿范围传下去**：传了等于让下游以为
          "范围就是这些工作表"，那正是"偷用当前工作簿"。下游据此明确拒绝发送。
        */
        scopeSource={scopeSource}
        folderScope={
          scopeSource === "folder" && folderScope
            ? {
                folderName: folderCatalog?.folderName ?? "工作目录",
                sheetCount: folderScope.keys.length
              }
            : null
        }
        onOpenScope={openScope}
        scopePickerOpen={scopeOpen}
        onScopeConfirmationReady={registerScopeConfirmation}
        onAuthorizeTarget={(indexes) => {
          const nextMetadata = listWpsWorksheets(getWpsApplication());
          setMetadata(nextMetadata);
          setProbe(runProbe());
          setScopeMode("manual");
          setScopeSource("workbook");
          setSelectedIndexes(indexes);
          setDraftIndexes(indexes);
        }}
        scopeDock={scopeDock}
      />
    </main>
  );
}
