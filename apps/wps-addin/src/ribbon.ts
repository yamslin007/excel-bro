import { getWpsApplication, type WpsTaskPane } from "./wpsHost";

const TASK_PANE_STORAGE_KEY = "excel_bro_taskpane_id";
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
    return "assets/excel-bro-icon.svg";
  }
};

declare global {
  interface Window {
    ribbon: typeof ribbon;
  }
}
