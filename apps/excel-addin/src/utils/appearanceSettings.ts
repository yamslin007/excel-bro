import type { ThemeSettings } from "./imageStorage";

export interface AppearanceSettings extends ThemeSettings {
  fit: "cover" | "contain";
  zoom: number;
  positionX: number;
  positionY: number;
  brightness: number;
  blur: number;
  cardOpacity: number;
  cardBlur: number;
  cardRadius: number;
  colorMode: "light" | "dark" | "system";
  accentColor: string | null;
}

export const APPEARANCE_DEFAULTS: AppearanceSettings = {
  preset: "default", opacity: 100, autoMask: false, hasCustomBackground: false,
  fit: "cover", zoom: 100, positionX: 50, positionY: 50, brightness: 100, blur: 0,
  cardOpacity: 90, cardBlur: 10, cardRadius: 16, colorMode: "light", accentColor: null
};

export const APPEARANCE_PRESETS = [
  { id: "office", label: "清晰办公", values: { cardOpacity: 97, cardBlur: 0, blur: 6, brightness: 90, autoMask: true } },
  { id: "glass", label: "轻透玻璃", values: { cardOpacity: 88, cardBlur: 12, blur: 4, brightness: 95, autoMask: false } },
  { id: "wallpaper", label: "突出壁纸", values: { cardOpacity: 76, cardBlur: 8, blur: 0, brightness: 100, autoMask: false } }
] as const;

/** Old records keep their image, palette and opacity; unknown fields never reach CSS. */
export function normalizeAppearance(value: Partial<AppearanceSettings> | null | undefined): AppearanceSettings {
  const raw = value ?? {};
  const next = { ...APPEARANCE_DEFAULTS };
  const number = (key: keyof AppearanceSettings, min: number, max: number) => {
    const value = raw[key];
    return typeof value === "number" && Number.isFinite(value)
      ? Math.min(max, Math.max(min, Math.round(value))) : APPEARANCE_DEFAULTS[key] as number;
  };
  if (["default", "warm-orange", "calm-blue", "vivid-green"].includes(raw.preset ?? "")) next.preset = raw.preset!;
  next.opacity = number("opacity", 0, 100);
  next.autoMask = raw.autoMask === true;
  next.hasCustomBackground = raw.hasCustomBackground === true;
  next.fit = raw.fit === "contain" ? "contain" : "cover";
  next.zoom = number("zoom", 100, 200);
  next.positionX = number("positionX", 0, 100);
  next.positionY = number("positionY", 0, 100);
  next.brightness = number("brightness", 40, 120);
  next.blur = number("blur", 0, 24);
  next.cardOpacity = number("cardOpacity", 55, 100);
  next.cardBlur = number("cardBlur", 0, 24);
  next.cardRadius = number("cardRadius", 0, 28);
  next.colorMode = raw.colorMode === "dark" || raw.colorMode === "system" ? raw.colorMode : "light";
  next.accentColor = typeof raw.accentColor === "string" && /^#[\da-f]{6}$/i.test(raw.accentColor) ? raw.accentColor : null;
  return next;
}
