import { getWpsApplication, type WpsTaskPane } from "./wpsHost";
import { readReviewPaneLease, TASK_PANE_STORAGE_KEY } from './wpsReviewPaneLease';

const DEFAULT_TASK_PANE_WIDTH = 1080;
const PREVIOUS_TASK_PANE_WIDTH = 720;
const initializedPaneIds = new Set<string>();
type RibbonTabSwitcher = { ActivateTabMso?: (tabId: string) => void };
let ribbonUiFromOnLoad: RibbonTabSwitcher | undefined;

function taskPaneUrl(): string {
  const url = new URL(window.location.href);
  url.hash = "/taskpane";
  return url.toString();
}

export function initializeVisiblePaneWidth(pane: WpsTaskPane, reused: boolean): void {
  const paneId = String(pane.ID);
  if (initializedPaneIds.has(paneId)) return;

  try {
    // Ribbon and pane live in separate WebViews. Do not migrate a temporary review width.
    if (readReviewPaneLease(getWpsApplication())?.paneId === paneId) return;
    const widthBefore = Number(pane.Width) || 0;
    // Migrate panes still near the former narrow default; preserve wider manual sizes.
    if (!reused || widthBefore <= PREVIOUS_TASK_PANE_WIDTH) {
      pane.Width = DEFAULT_TASK_PANE_WIDTH;
    }
    initializedPaneIds.add(paneId);
  } catch (error) {
    console.warn("WPS task pane width could not be set:", error);
  }
}

export function initializeRestoredTaskPaneWidth(): void {
  const application = getWpsApplication();
  const paneId = application.PluginStorage.getItem(TASK_PANE_STORAGE_KEY);
  if (paneId == null || paneId === "") return;
  const pane = application.GetTaskPane(paneId as string | number);
  if (pane) initializeVisiblePaneWidth(pane, true);
}

function createTaskPane(): WpsTaskPane {
  const application = getWpsApplication();
  const pane = application.CreateTaskPane(taskPaneUrl(), "Excel Bro");
  application.PluginStorage.setItem(TASK_PANE_STORAGE_KEY, pane.ID);
  if ("DockPosition" in pane) pane.DockPosition = 2;
  pane.Visible = true;
  initializeVisiblePaneWidth(pane, false);
  return pane;
}

export function showTaskPane(): WpsTaskPane {
  const application = getWpsApplication();
  const paneId = application.PluginStorage.getItem(TASK_PANE_STORAGE_KEY);
  if (paneId != null && paneId !== "") {
    try {
      const pane = application.GetTaskPane(paneId as string | number);
      if (pane) {
        pane.Visible = true;
        initializeVisiblePaneWidth(pane, true);
        return pane;
      }
    } catch {
      application.PluginStorage.removeItem?.(TASK_PANE_STORAGE_KEY);
    }
  }
  return createTaskPane();
}

export const ribbon = {
  OnAddinLoad(ribbonUI: unknown): boolean {
    getWpsApplication();
    ribbonUiFromOnLoad = ribbonUI && typeof ribbonUI === "object"
      ? ribbonUI as RibbonTabSwitcher
      : undefined;
    return true;
  },

  OnAction(control: { Id?: string }): boolean {
    if (control?.Id === "excelBroShowTaskPane") {
      showTaskPane();
      const application = getWpsApplication();
      const hostRibbonUI = application.ribbonUI as RibbonTabSwitcher | undefined;
      const ribbonUI = typeof hostRibbonUI?.ActivateTabMso === "function"
        ? hostRibbonUI
        : ribbonUiFromOnLoad;
      try {
        ribbonUI?.ActivateTabMso?.("TabHome");
      } catch {
        // The task pane is already open; a host-specific tab-switch failure is cosmetic.
      }
    }
    return true;
  },

  GetImage(): string {
    // 走 PNG 而不是 SVG：ribbon.xml 用的是 Office 2006 customUI 的 `getImage`
    // 回调，该回调按约定返回**栅格图 URL**（PNG 是这条路径上支持面最广的格式）。
    //
    // 文件名里的 `-v2` 是**必须的**，不是笔误：WPS 和 Office 都**按 URL 缓存**
    // 功能区图标。2026-10-12 实机踩到 —— 图片文件已换成新 logo、标签也已是新的，
    // 功能区仍显示退役的「EB」字标，就是因为 URL 没变、宿主直接吃缓存。
    // **以后换图必须同时把这里和 `manifest.xml` 的版本号一起 +1。**
    // 无版本号的老文件名仍在 public/assets 下作别名（由 scripts/build_logo_assets.py
    // 同一次运行产出），避免老 URL 404 出破图标。
    return "assets/excel-bro-icon-v2.png";
  }
};

declare global {
  interface Window {
    ribbon: typeof ribbon;
  }
}
