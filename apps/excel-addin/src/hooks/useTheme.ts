import { useCallback, useEffect, useRef, useState } from "react";
import {
  loadBackgroundImage,
  loadThemeSettings,
  saveThemeSnapshot,
  type ThemePreset
} from "../utils/imageStorage";
import {
  convertToJPEG,
  validateImageFile
} from "../utils/imageProcessing";

import { APPEARANCE_DEFAULTS, normalizeAppearance, type AppearanceSettings } from "../utils/appearanceSettings";

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

function rgba(hex: string, alpha: number): string {
  const n = parseInt(hex.slice(1), 16);
  return `rgba(${n >> 16}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`;
}

function lighten(hex: string): string {
  const n = parseInt(hex.slice(1), 16);
  return "#" + [n >> 16, (n >> 8) & 255, n & 255]
    .map((channel) => Math.round(channel * 0.55 + 255 * 0.45).toString(16).padStart(2, "0")).join("");
}

/** Keep accent hue while ensuring small links remain readable on the paper. */
function readingLinkColor(hex: string, background: string, dark: boolean): string {
  const channels = (color: string) => [1, 3, 5].map(index => parseInt(color.slice(index, index + 2), 16));
  const light = (rgb: number[]) => rgb.map(value => {
    const c = value / 255;
    return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  }).reduce((sum, value, index) => sum + value * [0.2126, 0.7152, 0.0722][index], 0);
  let rgb = channels(hex);
  const base = light(channels(background));
  for (let i = 0; i < 32; i++) {
    const ink = light(rgb);
    if ((Math.max(ink, base) + 0.05) / (Math.min(ink, base) + 0.05) >= 5) break;
    rgb = rgb.map(value => Math.round(value * 0.9 + (dark ? 255 : 0) * 0.1));
  }
  return '#' + rgb.map(value => value.toString(16).padStart(2, '0')).join('');
}

export function applyAppearance(settings: AppearanceSettings, url: string | null, systemDark = false): void {
  applyThemePalette(settings.preset);
  const style = rootStyle();
  const dark = settings.colorMode === "dark" || (settings.colorMode === "system" && systemDark);
  const palette = PRESETS[settings.preset];
  const surface = dark ? "#202824" : palette.background;
  const selectedAccent = settings.accentColor ?? palette.brandText;
  const accent = dark ? lighten(selectedAccent) : selectedAccent;
  document.documentElement.dataset.ebAppearance = dark ? "dark" : "light";
  if (dark) {
    for (const [name, value] of Object.entries({
      "background": "#161c19", "background-subtle": "#1b231f", "surface": "#242e28", "surface-hover": "#303d34",
      "text-primary": "#edf3ef", "text-secondary": "#b8c7bd", "text-tertiary": "#9aac9f",
      "border": "#3b4a40", "border-strong": "#56685b", "brand-bg": "#263f30", "brand-hover": "#34523f"
    })) style.setProperty(`--color-${name}`, value);
  }
  style.setProperty("--color-brand-text", accent);
  style.setProperty("--appearance-accent", accent);
  // Controls use a contrasting foreground even for a very light custom accent.
  const n = parseInt(accent.slice(1), 16);
  const luminance = (0.2126 * (n >> 16) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255)) / 255;
  style.setProperty("--appearance-accent-ink", luminance > 0.62 ? "#18231b" : "#ffffff");
  style.setProperty("--appearance-card-bg", rgba(surface, settings.cardOpacity / 100));
  // Only local text surfaces get a readability floor; empty space stays transparent.
  style.setProperty("--appearance-reading-bg", rgba(surface, Math.max(0.90, settings.cardOpacity / 100)));
  style.setProperty("--appearance-dock-bg", rgba(surface, 0.72));
  style.setProperty("--appearance-reading-secondary", dark ? "#b8c7bd" : "#535b57");
  style.setProperty("--appearance-reading-link", readingLinkColor(accent, surface, dark));
  style.setProperty("--appearance-input-bg", rgba(surface, Math.max(0.94, settings.cardOpacity / 100)));
  // Keep the selected theme hue; darken light accents until white text is readable.
  const userSurface = readingLinkColor(selectedAccent, "#ffffff", false);
  style.setProperty("--appearance-user-bg", userSurface);
  style.setProperty("--appearance-user-ink", "#ffffff");
  style.setProperty("--appearance-user-border", userSurface);
  style.setProperty("--appearance-card-blur", `${settings.cardBlur}px`);
  style.setProperty("--appearance-radius", `${settings.cardRadius}px`);
  style.setProperty("--appearance-fit", settings.fit);
  style.setProperty("--appearance-zoom", String(settings.zoom / 100));
  style.setProperty("--appearance-position", `${settings.positionX}% ${settings.positionY}%`);
  style.setProperty("--appearance-filter", `brightness(${settings.brightness / 100}) blur(${settings.blur}px)`);
  style.setProperty("--appearance-image-bleed", `${settings.blur}px`);
  if (dark) style.setProperty("--color-progress-glyph", accent);
  style.setProperty("--appearance-mask", settings.autoMask ? "0.38" : "0");
  style.setProperty("--custom-background-image", url ? `url("${url}")` : "none");
  style.setProperty("--custom-background-opacity", url ? String(settings.opacity / 100) : "0");
  document.querySelectorAll<HTMLElement>(".chat-shell").forEach((shell) => {
    if (url) shell.setAttribute("data-has-background", "true");
    else shell.removeAttribute("data-has-background");
    if (url && settings.autoMask) shell.setAttribute("data-auto-mask", "true");
    else shell.removeAttribute("data-auto-mask");
  });
}

type Snapshot = { settings: AppearanceSettings; image: Blob | null };

export function useTheme() {
  const [snapshot, setSnapshot] = useState<Snapshot>({ settings: APPEARANCE_DEFAULTS, image: null });
  const [backgroundUrl, setBackgroundUrl] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [editing, setEditing] = useState(false);
  const [error, setError] = useState("");
  const saved = useRef(snapshot);
  const current = useRef(snapshot);
  const editingRef = useRef(false);
  const savingRef = useRef(false);
  const uploadVersion = useRef(0);
  const uploadInFlight = useRef(false);
  const active = useRef(true);
  const [systemDark, setSystemDark] = useState(false);
  const replace = useCallback((next: Snapshot) => { current.current = next; setSnapshot(next); }, []);

  useEffect(() => {
    const query = window.matchMedia?.("(prefers-color-scheme: dark)");
    if (!query) return;
    const update = () => setSystemDark(query.matches);
    update();
    if (query.addEventListener) query.addEventListener("change", update);
    else query.addListener?.(update);
    return () => { if (query.removeEventListener) query.removeEventListener("change", update); else query.removeListener?.(update); };
  }, []);

  useEffect(() => {
    active.current = true;
    let cancelled = false;
    void (async () => {
      try {
        const stored = normalizeAppearance(await loadThemeSettings());
        const image = stored.hasCustomBackground ? await loadBackgroundImage() : null;
        if (cancelled) return;
        const next = { settings: { ...stored, hasCustomBackground: Boolean(image) }, image };
        saved.current = next;
        replace(next);
        if (stored.hasCustomBackground && !image) setError("已保存的壁纸无法读取，请重新选择图片。");
      } catch {
        if (!cancelled) setError("无法读取已保存的外观；可以预览，应用时将重新尝试保存。");
      } finally { if (!cancelled) setReady(true); }
    })();
    return () => { cancelled = true; active.current = false; uploadVersion.current++; };
  }, [replace]);

  useEffect(() => {
    const url = snapshot.image ? URL.createObjectURL(snapshot.image) : null;
    setBackgroundUrl(url);
    return () => { if (url) URL.revokeObjectURL(url); };
  }, [snapshot.image]);

  useEffect(() => { applyAppearance(snapshot.settings, backgroundUrl, systemDark); }, [snapshot.settings, backgroundUrl, systemDark]);
  useEffect(() => {
    document.documentElement.toggleAttribute("data-appearance-editing", editing);
    return () => document.documentElement.removeAttribute("data-appearance-editing");
  }, [editing]);

  const beginEdit = useCallback(() => {
    if (editingRef.current || savingRef.current) return;
    editingRef.current = true;
    setEditing(true);
    replace(saved.current);
  }, [replace]);
  const cancelEdit = useCallback(() => {
    if (savingRef.current || !editingRef.current) return;
    uploadVersion.current++;
    uploadInFlight.current = false;
    editingRef.current = false;
    setEditing(false);
    setUploading(false);
    setError("");
    replace(saved.current);
  }, [replace]);
  const updateSettings = useCallback((patch: Partial<AppearanceSettings>) => {
    if (!editingRef.current || savingRef.current) return;
    replace({ ...current.current, settings: normalizeAppearance({ ...current.current.settings, ...patch }) });
  }, [replace]);
  const uploadBackground = useCallback(async (file: File) => {
    if (!editingRef.current || savingRef.current) return;
    const version = ++uploadVersion.current;
    uploadInFlight.current = true;
    setError("");
    try {
      validateImageFile(file);
      setUploading(true);
      const image = await convertToJPEG(file, 0.9);
      if (!active.current || version !== uploadVersion.current || !editingRef.current) return;
      replace({ settings: { ...current.current.settings, hasCustomBackground: true }, image });
    } catch (reason) {
      if (active.current && version === uploadVersion.current) setError(reason instanceof Error ? reason.message : "图片处理失败");
    } finally { if (active.current && version === uploadVersion.current) { uploadInFlight.current = false; setUploading(false); } }
  }, [replace]);
  const removeBackground = useCallback(() => {
    if (!editingRef.current || savingRef.current) return;
    uploadVersion.current++;
    uploadInFlight.current = false;
    setUploading(false);
    replace({ settings: { ...current.current.settings, hasCustomBackground: false }, image: null });
  }, [replace]);
  const resetTheme = useCallback(() => {
    if (!editingRef.current || savingRef.current) return;
    uploadVersion.current++;
    uploadInFlight.current = false;
    setUploading(false);
    replace({ settings: { ...APPEARANCE_DEFAULTS }, image: null });
  }, [replace]);
  const applyChanges = useCallback(async () => {
    if (!editingRef.current || savingRef.current || uploadInFlight.current) return false;
    savingRef.current = true;
    setSaving(true);
    setError("");
    const next = current.current;
    try {
      await saveThemeSnapshot(next.settings, next.image);
      saved.current = next;
      editingRef.current = false;
      if (active.current) setEditing(false);
      return true;
    } catch {
      if (active.current) setError("外观未保存，请检查可用存储空间后重试；当前仍是临时预览。");
      return false;
    } finally { savingRef.current = false; if (active.current) setSaving(false); }
  }, []);

  return { ready, settings: snapshot.settings, backgroundUrl, uploading, saving, editing, error,
    beginEdit, cancelEdit, updateSettings, uploadBackground, removeBackground, resetTheme, applyChanges };
}
