import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { initializeRestoredTaskPaneWidth, ribbon, showTaskPane } from "./ribbon";
import type { WpsApplication, WpsTaskPane } from "./wpsHost";

const ribbonXml = readFileSync("public/ribbon.xml", "utf8");

describe("WPS ribbon placement", () => {
  it("uses a separate Excel Bro tab without modifying Home", () => {
    expect(ribbonXml).toMatch(/<tab id="excelBroWpsTab" label="Excel Bro">/);
    expect(ribbonXml).not.toContain('idMso="TabHome"');
    expect(ribbonXml).toContain('id="excelBroShowTaskPane"');
  });
});

afterEach(() => {
  if (window.Application) ribbon.OnAddinLoad(null);
  delete window.Application;
});

function attachApplication(
  storedId: unknown,
  getTaskPane: WpsApplication["GetTaskPane"]
): {
  createTaskPane: ReturnType<typeof vi.fn>;
  setItem: ReturnType<typeof vi.fn>;
} {
  const createTaskPane = vi.fn((): WpsTaskPane => ({ ID: 17, Visible: false }));
  const setItem = vi.fn();
  window.Application = {
    PluginStorage: {
      getItem: () => storedId,
      setItem
    },
    CreateTaskPane: createTaskPane,
    GetTaskPane: getTaskPane
  };
  return { createTaskPane, setItem };
}

describe("WPS task pane lifecycle", () => {
  it("reuses the existing pane", () => {
    const existing = { ID: 9, Visible: false, Width: 840 };
    const getTaskPane = vi.fn(() => existing);
    const { createTaskPane } = attachApplication(9, getTaskPane);

    expect(showTaskPane()).toBe(existing);
    expect(existing.Visible).toBe(true);
    expect(existing.Width).toBe(840);
    expect(createTaskPane).not.toHaveBeenCalled();
  });

  it("creates and stores a pane when no live pane exists", () => {
    const getTaskPane = vi.fn(() => undefined);
    const { createTaskPane, setItem } = attachApplication("stale", getTaskPane);

    const pane = showTaskPane();
    expect(createTaskPane).toHaveBeenCalledTimes(1);
    expect(setItem).toHaveBeenCalledWith("excel_bro_taskpane_id", 17);
    expect(pane.Visible).toBe(true);
    expect(pane.Width).toBe(1080);
  });

  it("sets the new pane width after making it visible", () => {
    let width = 0;
    let visibleWhenWidthSet = false;
    const pane: WpsTaskPane = {
      ID: 44,
      Visible: false,
      get Width() { return width; },
      set Width(value) {
        visibleWhenWidthSet = this.Visible;
        width = value ?? 0;
      }
    };
    window.Application = {
      PluginStorage: { getItem: () => null, setItem: vi.fn() },
      CreateTaskPane: () => pane,
      GetTaskPane: () => undefined
    };

    showTaskPane();
    expect(visibleWhenWidthSet).toBe(true);
    expect(width).toBe(1080);
  });

  it("widens a reused pane with the former default only once after showing it", () => {
    const existing: WpsTaskPane = { ID: 42, Visible: false, Width: 720 };
    attachApplication(42, () => existing);

    showTaskPane();
    expect(existing.Visible).toBe(true);
    expect(existing.Width).toBe(1080);

    existing.Width = 500;
    existing.Visible = false;
    showTaskPane();
    expect(existing.Width).toBe(500);
  });

  it("widens a restored pane even when WPS opens it without a Ribbon click", () => {
    const existing: WpsTaskPane = { ID: 43, Visible: true, Width: 520 };
    attachApplication(43, () => existing);

    initializeRestoredTaskPaneWidth();
    expect(existing.Width).toBe(1080);
  });

  it("opens the pane before returning to the built-in Home tab", () => {
    const getTaskPane = vi.fn(() => undefined);
    const { createTaskPane } = attachApplication(null, getTaskPane);
    const activateHome = vi.fn(() => {
      expect(createTaskPane).toHaveBeenCalledTimes(1);
    });
    window.Application!.ribbonUI = { ActivateTabMso: activateHome };

    expect(ribbon.OnAction({ Id: "excelBroShowTaskPane" })).toBe(true);
    expect(activateHome).toHaveBeenCalledWith("TabHome");
  });

  it("keeps the pane open when WPS cannot switch tabs", () => {
    const getTaskPane = vi.fn(() => undefined);
    const { createTaskPane } = attachApplication(null, getTaskPane);
    window.Application!.ribbonUI = {
      ActivateTabMso: () => { throw new Error("tab unavailable"); }
    };

    expect(() => ribbon.OnAction({ Id: "excelBroShowTaskPane" })).not.toThrow();
    expect(createTaskPane).toHaveBeenCalledTimes(1);
  });

  it("uses the ribbon object supplied by WPS on load when the app property is absent", () => {
    const getTaskPane = vi.fn(() => undefined);
    attachApplication(null, getTaskPane);
    const activateHome = vi.fn();
    expect(ribbon.OnAddinLoad({ ActivateTabMso: activateHome })).toBe(true);

    ribbon.OnAction({ Id: "excelBroShowTaskPane" });
    expect(activateHome).toHaveBeenCalledWith("TabHome");
  });
});
