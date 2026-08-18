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
  textPrimary: string;
  brandBg: string;
  brandText: string;
  brandHover: string;
}

const PRESETS: Record<ThemePreset, PresetPalette> = {
  default: {
    background: "#ffffff",
    backgroundSubtle: "#f7f6f3",
    textPrimary: "#37352f",
    brandBg: "#edf3ec",
    brandText: "#2f6b47",
    brandHover: "#e3ebe2"
  },
  "warm-orange": {
    background: "#fffbf5",
    backgroundSubtle: "#fff4e6",
    textPrimary: "#4a3520",
    brandBg: "#ffe8cc",
    brandText: "#c87a3a",
    brandHover: "#ffdbb0"
  },
  "calm-blue": {
    background: "#f8fbff",
    backgroundSubtle: "#eff6ff",
    textPrimary: "#1e3a5f",
    brandBg: "#dbeafe",
    brandText: "#3b82f6",
    brandHover: "#c7defe"
  },
  "vivid-green": {
    background: "#f7fff7",
    backgroundSubtle: "#ecfdf5",
    textPrimary: "#1a4d2e",
    brandBg: "#d1fae5",
    brandText: "#059669",
    brandHover: "#a7f3d0"
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
  style.setProperty("--color-text-primary", palette.textPrimary);
  style.setProperty("--color-brand-bg", palette.brandBg);
  style.setProperty("--color-brand-text", palette.brandText);
  style.setProperty("--color-brand-hover", palette.brandHover);
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
