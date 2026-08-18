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
  border: string;
  borderStrong: string;
  shadowSm: string;
  shadowMd: string;
}

const PRESETS: Record<ThemePreset, PresetPalette> = {
  default: {
    background: "#ffffff",
    backgroundSubtle: "#f7f6f3",
    surface: "#fafaf9",
    surfaceHover: "#f1f0ee",
    textPrimary: "#37352f",
    textSecondary: "#787774",
    textTertiary: "#9b9a97",
    brandBg: "#edf3ec",
    brandText: "#2f6b47",
    brandHover: "#e3ebe2",
    border: "#e9e9e7",
    borderStrong: "#d3d2cf",
    shadowSm: "0 1px 2px rgba(0, 0, 0, 0.04)",
    shadowMd: "0 2px 4px rgba(0, 0, 0, 0.06)"
  },
  "warm-orange": {
    background: "#faf9f7",
    backgroundSubtle: "#f5f3f0",
    surface: "#f0ede9",
    surfaceHover: "#ebe7e2",
    textPrimary: "#3d3730",
    textSecondary: "#6b6358",
    textTertiary: "#9a9186",
    brandBg: "#f0ebe5",
    brandText: "#7d6854",
    brandHover: "#e8e2db",
    border: "#e5e0db",
    borderStrong: "#d6cfc8",
    shadowSm: "0 1px 2px rgba(61, 55, 48, 0.04)",
    shadowMd: "0 2px 4px rgba(61, 55, 48, 0.06)"
  },
  "calm-blue": {
    background: "#f8f9fb",
    backgroundSubtle: "#f1f3f6",
    surface: "#ebeef2",
    surfaceHover: "#e3e7ec",
    textPrimary: "#2f3d4f",
    textSecondary: "#5a6678",
    textTertiary: "#8a95a5",
    brandBg: "#e8ecf1",
    brandText: "#5a6b7d",
    brandHover: "#dfe4ea",
    border: "#e3e8ed",
    borderStrong: "#d1d9e2",
    shadowSm: "0 1px 2px rgba(47, 61, 79, 0.04)",
    shadowMd: "0 2px 4px rgba(47, 61, 79, 0.06)"
  },
  "vivid-green": {
    background: "#f8faf9",
    backgroundSubtle: "#f2f5f3",
    surface: "#ecf0ed",
    surfaceHover: "#e4e9e6",
    textPrimary: "#2f3d35",
    textSecondary: "#5a6a60",
    textTertiary: "#8a9890",
    brandBg: "#e9efec",
    brandText: "#5a7065",
    brandHover: "#dfe6e2",
    border: "#e4ebe7",
    borderStrong: "#d3ddd7",
    shadowSm: "0 1px 2px rgba(47, 61, 53, 0.04)",
    shadowMd: "0 2px 4px rgba(47, 61, 53, 0.06)"
  }
};

const DEFAULT_SETTINGS: ThemeSettings = {
  preset: "default",
  opacity: 40,
  autoMask: true,
  hasCustomBackground: false
};

function rootStyle(): CSSStyleDeclaration {
  return document.documentElement.style;
}

function applyPalette(preset: ThemePreset): void {
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
  style.setProperty("--color-border", palette.border);
  style.setProperty("--color-border-strong", palette.borderStrong);
  style.setProperty("--shadow-sm", palette.shadowSm);
  style.setProperty("--shadow-md", palette.shadowMd);
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
        applyPalette(next.preset);
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
      applyPalette(preset);
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
    applyPalette("default");
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
