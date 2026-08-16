// @vitest-environment jsdom
import { describe, it, expect, beforeEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import {
  PET_VISIBILITY_STORAGE_KEY,
  normalizePetVisibility
} from "../conversation";
import { useUIState } from "./useUIState";

describe("useUIState", () => {
  beforeEach(() => {
    localStorage.clear();
    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      writable: true,
      value: 1200
    });
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
