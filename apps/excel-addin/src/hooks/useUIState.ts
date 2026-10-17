import { useCallback, useEffect, useRef, useState } from "react";
import {
  PET_VISIBILITY_STORAGE_KEY,
  normalizePetVisibility
} from "../conversation";
import { PANE_WIDEN_THRESHOLD } from "../utils";
import { useReviewPaneRatio, type ReviewPaneScope } from './useReviewPaneRatio';

const MAX_PREFERRED_TASK_PANE_WIDTH = 4096;
export const TASK_PANE_WIDTH_STORAGE_KEY =
  "excelBro.preferredTaskPaneWidth.v1";
export const WORK_DIRECTORY_GUIDE_STORAGE_KEY =
  "excelBro.workDirectoryGuideDismissed.v1";

export function normalizePreferredTaskPaneWidth(
  value: string | null
): number | null {
  if (value === null || value.trim() === "") return null;
  const width = Number(value);
  if (
    !Number.isFinite(width) ||
    width < PANE_WIDEN_THRESHOLD ||
    width > MAX_PREFERRED_TASK_PANE_WIDTH
  ) {
    return null;
  }
  return Math.round(width);
}

function readPreferredTaskPaneWidth(): number | null {
  try {
    return normalizePreferredTaskPaneWidth(
      localStorage.getItem(TASK_PANE_WIDTH_STORAGE_KEY)
    );
  } catch {
    return null;
  }
}

function storePreferredTaskPaneWidth(width: number | null): void {
  try {
    if (width === null) {
      localStorage.removeItem(TASK_PANE_WIDTH_STORAGE_KEY);
      return;
    }
    localStorage.setItem(TASK_PANE_WIDTH_STORAGE_KEY, String(width));
  } catch {
    // 存储不可用时只影响跨会话恢复，当前窗格仍可正常使用。
  }
}

/**
 * UI 状态管理 Hook
 *
 * 职责：
 * - 管理抽屉状态（工具、历史、设置）
 * - 管理菜单状态（模型菜单、更多菜单、规则管理器）
 * - 管理 UI 辅助状态（宠物可见性、窗格宽度、引导步骤、输入框高度）
 */
export function useUIState(comparisonActive = false, reviewScope?: ReviewPaneScope, reviewRearm = 0) {
  const ratioNeedsResize = useReviewPaneRatio(reviewScope, reviewRearm);
  const hasReviewScope = Boolean(reviewScope);
  const [initialPreferredTaskPaneWidth] = useState(
    readPreferredTaskPaneWidth
  );
  const preferredTaskPaneWidthRef = useRef(initialPreferredTaskPaneWidth);

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
  const [widenStepDone, setWidenStepDone] = useState(
    initialPreferredTaskPaneWidth !== null
  );
  const [workDirectoryGuideDismissed, setWorkDirectoryGuideDismissed] =
    useState(
      () =>
        localStorage.getItem(WORK_DIRECTORY_GUIDE_STORAGE_KEY) === "dismissed"
    );
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

  // 监听用户拖动，并在 TaskPaneApi 1.1 可用时恢复上次加宽后的实际宽度。
  useEffect(() => {
    let disposed = false;
    let canManageTaskPaneWidth = false;

    const restorePreferredWidth = () => {
      const preferredWidth = preferredTaskPaneWidthRef.current;
      if (
        comparisonActive ||
        !canManageTaskPaneWidth ||
        preferredWidth === null ||
        Math.abs(window.innerWidth - preferredWidth) < 2
      ) {
        return;
      }
      try {
        Office.extensionLifeCycle.taskpane.setWidth(preferredWidth);
      } catch {
        // 运行时能力声明与实际对象不一致时，保留手动拖动的降级路径。
        canManageTaskPaneWidth = false;
        if (window.innerWidth < PANE_WIDEN_THRESHOLD) {
          setWidenStepDone(false);
        }
      }
    };

    const handleResize = () => {
      const wide = window.innerWidth >= PANE_WIDEN_THRESHOLD;
      setIsNarrowPane(!wide);
      if (wide) setWidenStepDone(true);
      // Review resizing is temporary. Never overwrite the normal chat preference
      // with a programmatic resize or reapply it on focus during comparison.
      if (comparisonActive) return;
      if (!canManageTaskPaneWidth) return;

      const preferredWidth = normalizePreferredTaskPaneWidth(
        String(window.innerWidth)
      );
      preferredTaskPaneWidthRef.current = preferredWidth;
      storePreferredTaskPaneWidth(preferredWidth);
    };

    const handleFocus = () => restorePreferredWidth();
    const handleVisibilityChange = () => {
      if (document.visibilityState === "visible") restorePreferredWidth();
    };

    window.addEventListener("resize", handleResize);
    if (typeof Office !== "undefined") {
      Office.onReady(() => {
        if (disposed) return;
        try {
          canManageTaskPaneWidth =
            Office.context.requirements.isSetSupported("TaskPaneApi", "1.1");
        } catch {
          canManageTaskPaneWidth = false;
        }

        if (canManageTaskPaneWidth) {
          // Wait for a verified review scope: no preliminary fixed-width jump.
          if (!comparisonActive) restorePreferredWidth();
          if (canManageTaskPaneWidth) {
            window.addEventListener("focus", handleFocus);
            document.addEventListener(
              "visibilitychange",
              handleVisibilityChange
            );
          }
        } else if (window.innerWidth < PANE_WIDEN_THRESHOLD) {
          // 已保存的偏好在旧版 Excel 中无法自动恢复，继续展示手动拉宽引导。
          setWidenStepDone(false);
        }
      });
    }

    return () => {
      disposed = true;
      window.removeEventListener("resize", handleResize);
      window.removeEventListener("focus", handleFocus);
      document.removeEventListener("visibilitychange", handleVisibilityChange);
    };
  }, [comparisonActive]);

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

  const dismissWorkDirectoryGuide = useCallback(() => {
    setWorkDirectoryGuideDismissed(true);
    try {
      localStorage.setItem(WORK_DIRECTORY_GUIDE_STORAGE_KEY, "dismissed");
    } catch {
      // 存储不可用时，本次会话仍保持已跳过状态。
    }
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
    workDirectoryGuideDismissed,
    composerHeight,
    comparisonPaneNeedsResize: hasReviewScope ? ratioNeedsResize : comparisonActive,
    togglePetVisibility,
    setFocusOpening,
    setIsNarrowPane,
    setWidenStepDone,
    dismissWorkDirectoryGuide,
    setComposerHeight,
    setPetVisible
  };
}
