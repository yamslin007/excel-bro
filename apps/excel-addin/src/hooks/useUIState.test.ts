// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, renderHook, waitFor } from "@testing-library/react";
import {
  PET_VISIBILITY_STORAGE_KEY,
  normalizePetVisibility
} from "../conversation";
import {
  normalizePreferredTaskPaneWidth,
  TASK_PANE_WIDTH_STORAGE_KEY,
  useUIState,
  WORK_DIRECTORY_GUIDE_STORAGE_KEY
} from "./useUIState";

function stubTaskPaneApi(supported = true) {
  const setWidth = vi.fn();
  vi.stubGlobal("Office", {
    onReady: (callback: () => void) => {
      callback();
      return Promise.resolve();
    },
    context: {
      requirements: {
        isSetSupported: vi.fn(() => supported)
      }
    },
    extensionLifeCycle: {
      taskpane: { setWidth }
    }
  });
  return setWidth;
}

describe("useUIState", () => {
  beforeEach(() => {
    localStorage.clear();
    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      writable: true,
      value: 1200
    });
  });

  afterEach(() => {
    cleanup();
    vi.unstubAllGlobals();
  });

  it("should normalize only widened task pane widths", () => {
    expect(normalizePreferredTaskPaneWidth(null)).toBeNull();
    expect(normalizePreferredTaskPaneWidth("320")).toBeNull();
    expect(normalizePreferredTaskPaneWidth("719.6")).toBe(720);
    expect(normalizePreferredTaskPaneWidth("invalid")).toBeNull();
  });

  it("should restore the saved task pane width when TaskPaneApi is supported", async () => {
    localStorage.setItem(TASK_PANE_WIDTH_STORAGE_KEY, "720");
    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      writable: true,
      value: 320
    });
    const setWidth = stubTaskPaneApi();

    const { result } = renderHook(() => useUIState());

    await waitFor(() => expect(setWidth).toHaveBeenCalledWith(720));
    expect(result.current.widenStepDone).toBe(true);
  });

  it("waits for verified host measurement and never restores the wide chat preference on focus", () => {
    localStorage.setItem(TASK_PANE_WIDTH_STORAGE_KEY,"720");
    const setWidth = stubTaskPaneApi();
    const {result, rerender} = renderHook(({review}) => useUIState(review), {initialProps:{review:false}});
    setWidth.mockClear();
    rerender({review:true});
    expect(setWidth).not.toHaveBeenCalled();
    expect(result.current.comparisonPaneNeedsResize).toBe(true); // request isn't proof
    window.innerWidth = 360;
    act(() => window.dispatchEvent(new Event('resize')));
    expect(result.current.comparisonPaneNeedsResize).toBe(true); // iframe width cannot prove host ratio
    expect(localStorage.getItem(TASK_PANE_WIDTH_STORAGE_KEY)).toBe('720');
    setWidth.mockClear();
    act(() => {
      window.dispatchEvent(new Event('focus'));
      document.dispatchEvent(new Event('visibilitychange'));
    });
    expect(setWidth).not.toHaveBeenCalled();
    rerender({review:false});
    expect(setWidth).toHaveBeenCalledWith(720); // normal preference remains available
  });

  it.each(['unsupported','silent','throw'] as const)("keeps manual width guidance when the host is %s", reason => {
    localStorage.setItem(TASK_PANE_WIDTH_STORAGE_KEY,'720');
    const setWidth = stubTaskPaneApi(reason !== 'unsupported');
    if(reason === 'throw') setWidth.mockImplementation(() => {throw new Error('unsupported runtime');});
    const {result} = renderHook(() => useUIState(true));
    expect(result.current.comparisonPaneNeedsResize).toBe(true);
    expect(localStorage.getItem(TASK_PANE_WIDTH_STORAGE_KEY)).toBe('720');
    expect(setWidth).not.toHaveBeenCalled(); // no scope, no speculative width
    window.innerWidth = 350;
    act(() => window.dispatchEvent(new Event('resize')));
    expect(result.current.comparisonPaneNeedsResize).toBe(true);
    expect(localStorage.getItem(TASK_PANE_WIDTH_STORAGE_KEY)).toBe('720');
  });

  it("does not widen a narrow review pane or fight subsequent manual dragging", () => {
    window.innerWidth = 320;
    const setWidth = stubTaskPaneApi();
    const {result} = renderHook(() => useUIState(true));
    expect(setWidth).not.toHaveBeenCalled();
    expect(result.current.comparisonPaneNeedsResize).toBe(true);
    window.innerWidth = 430;
    act(() => window.dispatchEvent(new Event('resize')));
    act(() => window.dispatchEvent(new Event('focus')));
    expect(setWidth).not.toHaveBeenCalled();
    expect(result.current.comparisonPaneNeedsResize).toBe(true);
  });

  it("should remember the actual width after the user widens the pane", () => {
    stubTaskPaneApi();
    renderHook(() => useUIState());

    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      writable: true,
      value: 684
    });
    act(() => window.dispatchEvent(new Event("resize")));

    expect(localStorage.getItem(TASK_PANE_WIDTH_STORAGE_KEY)).toBe("684");
  });

  it("should reapply the saved width when the task pane becomes active again", () => {
    localStorage.setItem(TASK_PANE_WIDTH_STORAGE_KEY, "720");
    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      writable: true,
      value: 320
    });
    const setWidth = stubTaskPaneApi();
    renderHook(() => useUIState());
    setWidth.mockClear();

    act(() => window.dispatchEvent(new Event("focus")));

    expect(setWidth).toHaveBeenCalledWith(720);
  });

  it("should keep manual resizing as the fallback without TaskPaneApi", () => {
    localStorage.setItem(TASK_PANE_WIDTH_STORAGE_KEY, "720");
    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      writable: true,
      value: 320
    });
    const setWidth = stubTaskPaneApi(false);

    const { result } = renderHook(() => useUIState());

    expect(setWidth).not.toHaveBeenCalled();
    expect(result.current.widenStepDone).toBe(false);
  });

  it("should initialize drawers, menus and guides as closed", () => {
    const { result } = renderHook(() => useUIState());

    expect(result.current.toolsOpen).toBe(false);
    expect(result.current.historyOpen).toBe(false);
    expect(result.current.settingsOpen).toBe(false);
    expect(result.current.modelMenuOpen).toBe(false);
    expect(result.current.moreMenuOpen).toBe(false);
    expect(result.current.isRuleManagerOpen).toBe(false);
    expect(result.current.focusOpening).toBe(false);
    expect(result.current.widenStepDone).toBe(false);
    expect(result.current.workDirectoryGuideDismissed).toBe(false);
  });

  it("should initialize the work directory guide from localStorage", () => {
    localStorage.setItem(WORK_DIRECTORY_GUIDE_STORAGE_KEY, "dismissed");

    const { result } = renderHook(() => useUIState());

    expect(result.current.workDirectoryGuideDismissed).toBe(true);
  });

  it("should dismiss and persist the work directory guide", () => {
    const { result } = renderHook(() => useUIState());

    act(() => result.current.dismissWorkDirectoryGuide());

    expect(result.current.workDirectoryGuideDismissed).toBe(true);
    expect(localStorage.getItem(WORK_DIRECTORY_GUIDE_STORAGE_KEY)).toBe(
      "dismissed"
    );
  });

  it("should initialize pet visibility from localStorage", () => {
    localStorage.setItem(PET_VISIBILITY_STORAGE_KEY, "hidden");
    const { result } = renderHook(() => useUIState());

    expect(result.current.petVisible).toBe(false);
  });

  it("should default pet visibility to visible when storage is empty", () => {
    const { result } = renderHook(() => useUIState());

    expect(result.current.petVisible).toBe(
      normalizePetVisibility(localStorage.getItem(PET_VISIBILITY_STORAGE_KEY))
    );
    expect(result.current.petVisible).toBe(true);
  });

  it("should initialize pane width based on window.innerWidth", () => {
    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      writable: true,
      value: 320
    });

    const { result } = renderHook(() => useUIState());

    expect(result.current.isNarrowPane).toBe(true);
  });

  it("should close all drawers", () => {
    const { result } = renderHook(() => useUIState());

    act(() => {
      result.current.setToolsOpen(true);
      result.current.setHistoryOpen(true);
      result.current.setSettingsOpen(true);
    });

    act(() => result.current.closeAllDrawers());

    expect(result.current.toolsOpen).toBe(false);
    expect(result.current.historyOpen).toBe(false);
    expect(result.current.settingsOpen).toBe(false);
  });

  it("should close all menus", () => {
    const { result } = renderHook(() => useUIState());

    act(() => {
      result.current.setModelMenuOpen(true);
      result.current.setMoreMenuOpen(true);
      result.current.setIsRuleManagerOpen(true);
    });

    act(() => result.current.closeAllMenus());

    expect(result.current.modelMenuOpen).toBe(false);
    expect(result.current.moreMenuOpen).toBe(false);
    expect(result.current.isRuleManagerOpen).toBe(false);
  });

  it("should toggle pet visibility and persist it", () => {
    const { result } = renderHook(() => useUIState());

    act(() => result.current.togglePetVisibility());
    expect(result.current.petVisible).toBe(false);
    expect(localStorage.getItem(PET_VISIBILITY_STORAGE_KEY)).toBe("hidden");

    act(() => result.current.togglePetVisibility());
    expect(result.current.petVisible).toBe(true);
    expect(localStorage.getItem(PET_VISIBILITY_STORAGE_KEY)).toBe("visible");
  });

  it("should update narrow pane state on resize", () => {
    const { result } = renderHook(() => useUIState());

    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      writable: true,
      value: 320
    });

    act(() => {
      window.dispatchEvent(new Event("resize"));
    });

    expect(result.current.isNarrowPane).toBe(true);
    expect(result.current.widenStepDone).toBe(false);
  });

  it("should mark widen step done when resizing to wide pane", () => {
    const { result } = renderHook(() => useUIState());

    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      writable: true,
      value: 1400
    });

    act(() => {
      window.dispatchEvent(new Event("resize"));
    });

    expect(result.current.isNarrowPane).toBe(false);
    expect(result.current.widenStepDone).toBe(true);
  });

  it("should expose internal setters for UI state", () => {
    const { result } = renderHook(() => useUIState());

    act(() => {
      result.current.setFocusOpening(true);
      result.current.setComposerHeight(180);
    });

    expect(result.current.focusOpening).toBe(true);
    expect(result.current.composerHeight).toBe(180);
  });
});
