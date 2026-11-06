// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, renderHook, screen, waitFor } from "@testing-library/react";
import { useState } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { useTheme } from "./useTheme";
import ThemePanel from "../components/ThemePanel";
import { APPEARANCE_DEFAULTS, normalizeAppearance } from "../utils/appearanceSettings";
import * as storage from "../utils/imageStorage";
import { convertToJPEG } from "../utils/imageProcessing";

vi.mock("../utils/imageStorage", () => ({ loadThemeSettings: vi.fn(), loadBackgroundImage: vi.fn(), saveThemeSnapshot: vi.fn() }));
vi.mock("../utils/imageProcessing", () => ({ convertToJPEG: vi.fn(), validateImageFile: vi.fn() }));
const originalImage = new Blob(["original"], { type: "image/jpeg" });

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(storage.loadThemeSettings).mockResolvedValue({ preset: "calm-blue", opacity: 73, autoMask: true, hasCustomBackground: true });
  vi.mocked(storage.loadBackgroundImage).mockResolvedValue(originalImage);
  vi.mocked(storage.saveThemeSnapshot).mockResolvedValue();
  vi.mocked(convertToJPEG).mockResolvedValue(new Blob(["new"], { type: "image/jpeg" }));
  let index = 0;
  Object.defineProperty(URL, "createObjectURL", { configurable: true, value: vi.fn(() => `blob:theme-${++index}`) });
  Object.defineProperty(URL, "revokeObjectURL", { configurable: true, value: vi.fn() });
});
afterEach(() => { cleanup(); document.documentElement.removeAttribute("style"); delete document.documentElement.dataset.ebAppearance; });

async function mountTheme() {
  const hook = renderHook(useTheme);
  await waitFor(() => expect(hook.result.current.ready).toBe(true));
  return hook;
}

describe("transactional appearance editing", () => {
  it("migrates old records, previews immediately and cancels without writing", async () => {
    const { result } = await mountTheme();
    expect(result.current.settings).toMatchObject({ preset: "calm-blue", opacity: 73, autoMask: true, zoom: 100 });
    act(() => { result.current.beginEdit(); result.current.updateSettings({ brightness: 65, cardOpacity: 80 }); });
    expect(document.documentElement.style.getPropertyValue("--appearance-filter")).toContain("0.65");
    expect(storage.saveThemeSnapshot).not.toHaveBeenCalled();
    act(() => result.current.cancelEdit());
    expect(result.current.settings.brightness).toBe(100);
    expect(result.current.settings.opacity).toBe(73);
    expect(storage.saveThemeSnapshot).not.toHaveBeenCalled();
  });
  it("cancel restores a removed or replaced image and revokes transient URLs", async () => {
    const { result } = await mountTheme();
    act(() => result.current.beginEdit());
    await act(async () => { await result.current.uploadBackground(new File(["new"], "new.png", { type: "image/png" })); });
    expect(URL.revokeObjectURL).toHaveBeenCalled();
    act(() => { result.current.removeBackground(); result.current.cancelEdit(); });
    expect(result.current.settings.hasCustomBackground).toBe(true);
    expect(result.current.backgroundUrl).not.toBeNull();
    expect(storage.saveThemeSnapshot).not.toHaveBeenCalled();
  });
  it("apply atomically saves image and settings; later cancel restores the applied state", async () => {
    const { result } = await mountTheme();
    act(() => { result.current.beginEdit(); result.current.updateSettings({ cardOpacity: 81 }); });
    await act(async () => { expect(await result.current.applyChanges()).toBe(true); });
    expect(storage.saveThemeSnapshot).toHaveBeenCalledWith(expect.objectContaining({ cardOpacity: 81 }), originalImage);
    act(() => { result.current.beginEdit(); result.current.resetTheme(); result.current.cancelEdit(); });
    expect(result.current.settings.cardOpacity).toBe(81);
    expect(result.current.settings.hasCustomBackground).toBe(true);
    expect(storage.saveThemeSnapshot).toHaveBeenCalledTimes(1);
  });
  it("save failure stays in preview and allows cancellation without claiming persistence", async () => {
    vi.mocked(storage.saveThemeSnapshot).mockRejectedValueOnce(new Error("quota"));
    const { result } = await mountTheme();
    act(() => { result.current.beginEdit(); result.current.removeBackground(); });
    await act(async () => { expect(await result.current.applyChanges()).toBe(false); });
    expect(result.current.editing).toBe(true);
    expect(result.current.error).toContain("未保存");
    act(() => result.current.cancelEdit());
    expect(result.current.settings.hasCustomBackground).toBe(true);
  });
  it("an upload finishing after cancel/reopen cannot overwrite the new draft", async () => {
    let finish!: (value: Blob) => void;
    vi.mocked(convertToJPEG).mockReturnValueOnce(new Promise((resolve) => { finish = resolve; }));
    const { result } = await mountTheme();
    act(() => result.current.beginEdit());
    let upload!: Promise<void>;
    act(() => { upload = result.current.uploadBackground(new File(["late"], "late.png", { type: "image/png" })); });
    await act(async () => { expect(await result.current.applyChanges()).toBe(false); });
    expect(storage.saveThemeSnapshot).not.toHaveBeenCalled();
    act(() => { result.current.cancelEdit(); result.current.beginEdit(); result.current.updateSettings({ zoom: 135 }); });
    await act(async () => { finish(new Blob(["late"])); await upload; });
    expect(result.current.settings.zoom).toBe(135);
    await act(async () => { await result.current.applyChanges(); });
    expect(storage.saveThemeSnapshot).toHaveBeenCalledWith(expect.objectContaining({ zoom: 135 }), originalImage);
  });
  it("validates corrupt settings and does not erase saved data on read failure", async () => {
    expect(normalizeAppearance({ zoom: Infinity, cardOpacity: -50, accentColor: 'url(bad)' })).toMatchObject({ zoom: 100, cardOpacity: 55, accentColor: null });
    vi.mocked(storage.loadThemeSettings).mockRejectedValueOnce(new Error("unavailable"));
    const { result } = await mountTheme();
    expect(result.current.error).toContain("无法读取");
    expect(storage.saveThemeSnapshot).not.toHaveBeenCalled();
  });
  it("blocks changes and duplicate applies during an in-flight atomic commit", async () => {
    let finish!: () => void;
    vi.mocked(storage.saveThemeSnapshot).mockImplementationOnce(() => new Promise((resolve) => { finish = resolve; }));
    const { result } = await mountTheme();
    act(() => { result.current.beginEdit(); result.current.updateSettings({ zoom: 145 }); });
    let saving!: Promise<boolean>;
    act(() => { saving = result.current.applyChanges(); result.current.updateSettings({ zoom: 180 }); result.current.cancelEdit(); });
    await act(async () => { expect(await result.current.applyChanges()).toBe(false); finish(); await saving; });
    expect(result.current.settings.zoom).toBe(145);
    expect(storage.saveThemeSnapshot).toHaveBeenCalledTimes(1);
  });
});

it.each(["excel", "wps"])("%s opens the shared preview, switches controls and Escape restores the prior appearance", async (host) => {
  function Harness() {
    const api = useTheme(); const [open, setOpen] = useState(false);
    return <div className={`chat-shell ${host === "wps" ? "wps-chat-shell" : ""}`}><button onClick={() => setOpen(true)}>壁纸</button>
      <ThemePanel open={open} onClose={() => setOpen(false)} {...api} /></div>;
  }
  render(<Harness />);
  fireEvent.click(screen.getByRole("button", { name: "壁纸" }));
  await waitFor(() => expect(screen.getByRole("button", { name: "应用" }).hasAttribute("disabled")).toBe(false));
  fireEvent.change(screen.getByLabelText("亮度"), { target: { value: "60" } });
  expect(document.documentElement.style.getPropertyValue("--appearance-filter")).toContain("0.6");
  fireEvent.click(screen.getByRole("tab", { name: "内容卡片" }));
  expect(screen.getByLabelText("圆角")).toBeTruthy();
  fireEvent.click(screen.getByRole("tab", { name: "配色" }));
  fireEvent.click(screen.getByRole("button", { name: "深色" }));
  expect(document.documentElement.dataset.ebAppearance).toBe("dark");
  fireEvent.keyDown(document, { key: "Escape" });
  expect(screen.queryByRole("dialog")).toBeNull();
  expect(document.documentElement.dataset.ebAppearance).toBe("light");
  expect(document.documentElement.style.getPropertyValue("--appearance-filter")).toContain("brightness(1)");
  expect(storage.saveThemeSnapshot).not.toHaveBeenCalled();
});
