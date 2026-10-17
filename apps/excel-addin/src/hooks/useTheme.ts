import { useCallback, useEffect, useRef, useState } from "react";
import {
  clearBackgroundImage,
  loadBackgroundImage,
  loadThemeSettings,
  saveBackgroundImage,
  saveThemeSettings,
  type ThemePreset,
  type ThemeSettings
} from "../utils/imageStorage";
import {
  convertToJPEG,
  validateImageFile
} from "../utils/imageProcessing";

interface PresetPalette {
  background: string;
  backgroundSubtle: string;
  surface: string;
  surfaceHover: string;
  textPrimary: string;
  textSecondary: string;
  textTertiary: string;
  brandBg: string;
  brandText: string;
  brandHover: string;
  /**
   * 「有新进度」浮标的**前景色**（两个字形共用一个值：运行中的三颗圆点、
   * 跑完之后的锚图标）。
   *
   * 为什么要单独一个令牌，而不是在 CSS 里用 `color-mix()` 从 `brandText` 现算：
   *
   * 1. **宿主引擎的 `color-mix()` 支持不可假设**。Excel 侧是 WebView2（新版
   *    Chromium，支持），WPS 侧是 QtWebEngine，版本随安装包走。一旦不支持，
   *    那条声明被整条丢弃；而**降级行救不了**——`background: var(--x)` 与
   *    `background: color-mix(...)` 是同一个属性，后者在前者之后、赢在层叠，
   *    只是"算出无效"而非"语法无效"，于是 `background` 落到初始值
   *    `transparent`，图形直接**看不见**（不是"浅一点的主题色"）。
   *    写成令牌就没有任何运行时色彩函数参与，任何引擎都拿得到确定值。
   * 2. 值可以**被守卫断言**。`styles.test.ts` 能直接读它、算对比度、并校验它
   *    确实是本主题 `brandText` 的同色相压深，而不是"某个近白色"。
   *
   * 取值 = 本主题 `brandText` 向黑混 20%（只降明度、不动色相）。20% 是**必要
   * 且充分**的下限：压在环形渐变最暗端（`brandHover`）上，warm-orange 需要它
   * 才能从 2.58 抬到 3.53，vivid-green 从 2.93 抬到 4.00，双双过非文本 UI 的
   * **3:1**（WCAG 1.4.11）。压深 10% 在静态下够、悬停端点只有 2.89，不够。
   *
   * 字段名与令牌一样刻意**不跟着字形走**（旧名 `progressDot`）：浮标现在有两个
   * 字形，名字跟着字形改就会在下一次换字形时变成谎话。
   */
  progressGlyph: string;
  border: string;
  borderStrong: string;
  shadowSm: string;
  shadowMd: string;
  pattern: string;
}

const PRESETS: Record<ThemePreset, PresetPalette> = {
  default: {
    background: "#ffffff",
    backgroundSubtle: "#ffffff",
    surface: "#fafaf9",
    surfaceHover: "#f1f0ee",
    textPrimary: "#37352f",
    textSecondary: "#787774",
    textTertiary: "#9b9a97",
    brandBg: "#edf3ec",
    brandText: "#2f6b47",
    brandHover: "#e3ebe2",
    progressGlyph: "#265639",
    border: "#e9e9e7",
    borderStrong: "#d3d2cf",
    shadowSm: "0 1px 2px rgba(0, 0, 0, 0.04)",
    shadowMd: "0 2px 4px rgba(0, 0, 0, 0.06)",
    pattern: "none"
  },
  "warm-orange": {
    background: "#fffcf7",
    backgroundSubtle: "#fff5e8",
    surface: "#ffefd9",
    surfaceHover: "#ffe6c7",
    textPrimary: "#5c3d2e",
    textSecondary: "#8b6f5c",
    textTertiary: "#b39a87",
    brandBg: "#ffe4c4",
    brandText: "#d97739",
    brandHover: "#ffd9b0",
    progressGlyph: "#ae5f2e",
    border: "#f5d9b8",
    borderStrong: "#e8c89f",
    shadowSm: "0 1px 2px rgba(217, 119, 57, 0.08)",
    shadowMd: "0 2px 4px rgba(217, 119, 57, 0.12)",
    pattern: `url("data:image/svg+xml,%3Csvg width='60' height='60' xmlns='http://www.w3.org/2000/svg'%3E%3Ccircle cx='10' cy='10' r='3' fill='%23d97739' opacity='0.04'/%3E%3Cpath d='M10 5 L10 6 M10 14 L10 15 M5 10 L6 10 M14 10 L15 10 M7 7 L7.7 7.7 M12.3 12.3 L13 13 M13 7 L12.3 7.7 M7.7 12.3 L7 13' stroke='%23d97739' opacity='0.04' stroke-width='0.5'/%3E%3Ccircle cx='50' cy='50' r='3' fill='%23d97739' opacity='0.03'/%3E%3C/svg%3E")`
  },
  "calm-blue": {
    background: "#f9fbff",
    backgroundSubtle: "#f0f6ff",
    surface: "#e3f0ff",
    surfaceHover: "#d4e8ff",
    textPrimary: "#2c3e5c",
    textSecondary: "#5a6b8a",
    textTertiary: "#8a9bb5",
    brandBg: "#d4e8ff",
    brandText: "#3b7dd6",
    brandHover: "#c2deff",
    progressGlyph: "#2f64ab",
    border: "#c7dff7",
    borderStrong: "#aecef0",
    shadowSm: "0 1px 2px rgba(59, 125, 214, 0.08)",
    shadowMd: "0 2px 4px rgba(59, 125, 214, 0.12)",
    pattern: `url("data:image/svg+xml,%3Csvg width='80' height='60' xmlns='http://www.w3.org/2000/svg'%3E%3Cellipse cx='15' cy='20' rx='8' ry='5' fill='%233b7dd6' opacity='0.04'/%3E%3Cellipse cx='20' cy='18' rx='6' ry='4' fill='%233b7dd6' opacity='0.04'/%3E%3Cellipse cx='10' cy='21' rx='5' ry='3' fill='%233b7dd6' opacity='0.04'/%3E%3Cellipse cx='55' cy='45' rx='7' ry='4' fill='%233b7dd6' opacity='0.03'/%3E%3C/svg%3E")`
  },
  "vivid-green": {
    background: "#f7fffb",
    backgroundSubtle: "#edfff4",
    surface: "#dfffea",
    surfaceHover: "#ceffd9",
    textPrimary: "#1e4d3b",
    textSecondary: "#4a7562",
    textTertiary: "#7a9f8e",
    brandBg: "#c8f5d9",
    brandText: "#2d9b63",
    brandHover: "#b5f0ca",
    progressGlyph: "#247c4f",
    border: "#b8e8ca",
    borderStrong: "#9ddab3",
    shadowSm: "0 1px 2px rgba(45, 155, 99, 0.08)",
    shadowMd: "0 2px 4px rgba(45, 155, 99, 0.12)",
    pattern: `url("data:image/svg+xml,%3Csvg width='60' height='60' xmlns='http://www.w3.org/2000/svg'%3E%3Cpath d='M12 18 Q15 12, 18 18 Q15 15, 12 18' fill='%232d9b63' opacity='0.04'/%3E%3Cpath d='M8 10 Q10 6, 12 10 Q10 8, 8 10' fill='%232d9b63' opacity='0.04'/%3E%3Cpath d='M48 48 Q50 44, 52 48 Q50 46, 48 48' fill='%232d9b63' opacity='0.03'/%3E%3C/svg%3E")`
  }
};

const DEFAULT_SETTINGS: ThemeSettings = {
  preset: "default",
  opacity: 100,
  autoMask: false,
  hasCustomBackground: false
};

function rootStyle(): CSSStyleDeclaration {
  return document.documentElement.style;
}

/** 把主题预设的调色板应用到当前 documentElement（独立窗口复用）。 */
export function applyThemePalette(preset: ThemePreset): void {
  const palette = PRESETS[preset];
  const style = rootStyle();
  style.setProperty("--color-background", palette.background);
  style.setProperty("--color-background-subtle", palette.backgroundSubtle);
  style.setProperty("--color-surface", palette.surface);
  style.setProperty("--color-surface-hover", palette.surfaceHover);
  style.setProperty("--color-text-primary", palette.textPrimary);
  style.setProperty("--color-text-secondary", palette.textSecondary);
  style.setProperty("--color-text-tertiary", palette.textTertiary);
  style.setProperty("--color-brand-bg", palette.brandBg);
  style.setProperty("--color-brand-text", palette.brandText);
  style.setProperty("--color-brand-hover", palette.brandHover);
  style.setProperty("--color-progress-glyph", palette.progressGlyph);
  style.setProperty("--color-border", palette.border);
  style.setProperty("--color-border-strong", palette.borderStrong);
  style.setProperty("--shadow-sm", palette.shadowSm);
  style.setProperty("--shadow-md", palette.shadowMd);
  style.setProperty("--theme-pattern", palette.pattern);
}

function applyBackground(url: string | null, opacity: number): void {
  const style = rootStyle();
  const container = document.querySelector<HTMLElement>(".chat-shell");
  if (!url) {
    style.setProperty("--custom-background-image", "none");
    style.setProperty("--custom-background-opacity", "0");
    container?.removeAttribute("data-has-background");
    return;
  }
  style.setProperty("--custom-background-image", `url("${url}")`);
  style.setProperty("--custom-background-opacity", String(opacity / 100));
  container?.setAttribute("data-has-background", "true");
}

function applyMask(enabled: boolean): void {
  const container = document.querySelector<HTMLElement>(".chat-shell");
  if (!container) return;
  if (enabled && container.hasAttribute("data-has-background")) {
    container.setAttribute("data-auto-mask", "true");
  } else {
    container.removeAttribute("data-auto-mask");
  }
}

export function useTheme() {
  const [settings, setSettings] = useState<ThemeSettings>(DEFAULT_SETTINGS);
  const [backgroundUrl, setBackgroundUrl] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState("");
  const objectUrlRef = useRef<string | null>(null);

  const persist = useCallback(async (next: ThemeSettings) => {
    setSettings(next);
    try {
      await saveThemeSettings(next);
    } catch {
      // IndexedDB 不可用时仍允许本次会话使用主题。
    }
  }, []);

  useEffect(() => {
    let active = true;
    async function initialize() {
      try {
        const stored = await loadThemeSettings();
        const next = stored ?? DEFAULT_SETTINGS;
        applyThemePalette(next.preset);
        applyMask(next.autoMask);
        setSettings(next);

        if (next.hasCustomBackground) {
          const blob = await loadBackgroundImage();
          if (active && blob) {
            const url = URL.createObjectURL(blob);
            objectUrlRef.current = url;
            setBackgroundUrl(url);
            applyBackground(url, next.opacity);
          }
        }
      } catch {
        // 降级为默认主题。
      } finally {
        if (active) setReady(true);
      }
    }
    void initialize();
    return () => {
      active = false;
      if (objectUrlRef.current) {
        URL.revokeObjectURL(objectUrlRef.current);
        objectUrlRef.current = null;
      }
    };
  }, []);

  const applyPreset = useCallback(
    async (preset: ThemePreset) => {
      applyThemePalette(preset);
      await persist({ ...settings, preset });
    },
    [persist, settings]
  );

  const uploadBackground = useCallback(
    async (file: File) => {
      setError("");
      try {
        console.log("[Theme] 上传开始:", file.name, file.size, file.type);
        validateImageFile(file);
        setUploading(true);
        const blob = await convertToJPEG(file, 0.9);
        await saveBackgroundImage(blob);
        if (objectUrlRef.current) {
          URL.revokeObjectURL(objectUrlRef.current);
        }
        const url = URL.createObjectURL(blob);
        objectUrlRef.current = url;
        setBackgroundUrl(url);
        applyBackground(url, settings.opacity);
        applyMask(settings.autoMask);
        await persist({
          ...settings,
          hasCustomBackground: true
        });
      } catch (reason) {
        setError(reason instanceof Error ? reason.message : "图片上传失败");
      } finally {
        setUploading(false);
      }
    },
    [persist, settings]
  );

  const removeBackground = useCallback(async () => {
    if (objectUrlRef.current) {
      URL.revokeObjectURL(objectUrlRef.current);
      objectUrlRef.current = null;
    }
    setBackgroundUrl(null);
    applyBackground(null, 0);
    applyMask(false);
    try {
      await clearBackgroundImage();
    } catch {
      // 即使删除失败，本次会话仍应恢复默认。
    }
    await persist({
      ...settings,
      hasCustomBackground: false
    });
  }, [persist, settings]);

  const setOpacity = useCallback(
    async (opacity: number) => {
      const next = { ...settings, opacity };
      setSettings(next);
      if (backgroundUrl) {
        applyBackground(backgroundUrl, opacity);
      }
      await persist(next);
    },
    [backgroundUrl, persist, settings]
  );

  const setAutoMask = useCallback(
    async (autoMask: boolean) => {
      const next = { ...settings, autoMask };
      setSettings(next);
      applyMask(autoMask);
      await persist(next);
    },
    [persist, settings]
  );

  const resetTheme = useCallback(async () => {
    applyThemePalette("default");
    await removeBackground();
    await persist({ ...DEFAULT_SETTINGS });
  }, [persist, removeBackground]);

  return {
    ready,
    settings,
    backgroundUrl,
    uploading,
    error,
    applyPreset,
    uploadBackground,
    removeBackground,
    setOpacity,
    setAutoMask,
    resetTheme
  };
}
