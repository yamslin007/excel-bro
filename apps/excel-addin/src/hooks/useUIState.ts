import { useCallback, useEffect, useState } from "react";
import {
  PET_VISIBILITY_STORAGE_KEY,
  normalizePetVisibility
} from "../conversation";

const PANE_WIDEN_THRESHOLD = 380;

/**
 * UI 状态管理 Hook
 *
 * 职责：
 * - 管理抽屉状态（工具、历史、设置）
 * - 管理菜单状态（模型菜单、更多菜单、规则管理器）
 * - 管理 UI 辅助状态（宠物可见性、窗格宽度、引导步骤、输入框高度）
 */
export function useUIState() {
  // 抽屉状态
  const [toolsOpen, setToolsOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);

  // 菜单状态
  const [modelMenuOpen, setModelMenuOpen] = useState(false);
  const [moreMenuOpen, setMoreMenuOpen] = useState(false);
  const [isRuleManagerOpen, setIsRuleManagerOpen] = useState(false);

  // UI 辅助状态
  const [petVisible, setPetVisible] = useState(() =>
    normalizePetVisibility(localStorage.getItem(PET_VISIBILITY_STORAGE_KEY))
  );
  const [focusOpening, setFocusOpening] = useState(false);
  const [isNarrowPane, setIsNarrowPane] = useState(
    () =>
      typeof window !== "undefined" &&
      window.innerWidth < PANE_WIDEN_THRESHOLD
  );
  const [widenStepDone, setWidenStepDone] = useState(false);
  const [composerHeight, setComposerHeight] = useState<number | null>(null);

  // 宠物可见性持久化
  useEffect(() => {
    try {
      localStorage.setItem(
        PET_VISIBILITY_STORAGE_KEY,
        petVisible ? "visible" : "hidden"
      );
    } catch {
      // The preference remains active for this session if storage is blocked.
    }
  }, [petVisible]);

  // 窗格宽度监听
  useEffect(() => {
    const handleResize = () => {
      const wide = window.innerWidth >= PANE_WIDEN_THRESHOLD;
      setIsNarrowPane(!wide);
      if (wide && !widenStepDone) {
        setWidenStepDone(true);
      }
    };

    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, [widenStepDone]);

  // 关闭所有抽屉
  const closeAllDrawers = useCallback(() => {
    setToolsOpen(false);
    setHistoryOpen(false);
    setSettingsOpen(false);
  }, []);

  // 关闭所有菜单
  const closeAllMenus = useCallback(() => {
    setModelMenuOpen(false);
    setMoreMenuOpen(false);
    setIsRuleManagerOpen(false);
  }, []);

  // 切换宠物可见性
  const togglePetVisibility = useCallback(() => {
    setPetVisible((current) => !current);
  }, []);

  return {
    // 抽屉状态
    toolsOpen,
    historyOpen,
    settingsOpen,
    setToolsOpen,
    setHistoryOpen,
    setSettingsOpen,
    closeAllDrawers,

    // 菜单状态
    modelMenuOpen,
    moreMenuOpen,
    isRuleManagerOpen,
    setModelMenuOpen,
    setMoreMenuOpen,
    setIsRuleManagerOpen,
    closeAllMenus,

    // UI 辅助状态
    petVisible,
    focusOpening,
    isNarrowPane,
    widenStepDone,
    composerHeight,
    togglePetVisibility,
    setFocusOpening,
    setIsNarrowPane,
    setWidenStepDone,
    setComposerHeight,
    setPetVisible
  };
}
