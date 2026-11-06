import { useEffect, useRef, useState, type PointerEvent } from "react";
import { createPortal } from "react-dom";
import type { useTheme } from "../hooks/useTheme";
import { APPEARANCE_DEFAULTS, APPEARANCE_PRESETS, type AppearanceSettings } from "../utils/appearanceSettings";
import { brandLogoUrl } from "../brandLogo";
import { ConversationMessage } from "./ConversationMessage";

type ThemeApi = ReturnType<typeof useTheme>;
interface ThemePanelProps extends ThemeApi { open: boolean; onClose: () => void; }
const COLORS = [
  { id: "default", label: "默认 Notion", color: "#2f6b47" },
  { id: "warm-orange", label: "温暖米色", color: "#d97739" },
  { id: "calm-blue", label: "石板灰蓝", color: "#3b7dd6" },
  { id: "vivid-green", label: "鼠尾草绿", color: "#2d9b63" }
] as const;
const TABS = ["壁纸", "内容卡片", "配色"] as const;
type NumericKey = "zoom" | "positionX" | "positionY" | "brightness" | "blur" | "opacity" | "cardOpacity" | "cardBlur" | "cardRadius";

export function AppearancePreview({ hasImage }: { hasImage: boolean }) {
  return <section className="appearance-preview" data-image={hasImage || undefined} aria-label="外观实时预览">
    <div className="appearance-preview-wallpaper" /><div className="appearance-preview-mask" />
    <div className="appearance-preview-label"><i />实时预览 <span>尚未应用</span></div>
    <div className="appearance-preview-messages">
      <div className="appearance-sample-user">帮我检查这三张表的公式</div>
      <ConversationMessage role="assistant" author>
        <div className="appearance-sample-card appearance-sample-status"><span>已完成 · 3 秒</span><span>查看过程 ›</span></div>
        <div className="appearance-sample-card"><strong>三张工作表的公式一致。</strong><p>引用与预期相符，本次检查未修改工作表。</p></div>
      </ConversationMessage>
      <small className="appearance-preview-note">示例内容，仅用于预览外观</small>
    </div>
    <div className="appearance-sample-composer">
      <div className="appearance-sample-scope"><span>＋ 新对话</span><span>↶</span><span>◷</span><span>工作簿.xlsx · 已选 3 张表</span></div>
      <div className="appearance-sample-input"><span>描述你想查询、分析或修改的内容…</span>
        <div><img src={brandLogoUrl} alt="" /><span>模型⌄</span><b>▲</b></div>
      </div>
    </div>
  </section>;
}

export default function ThemePanel(props: ThemePanelProps) {
  const { open, onClose, ready, settings, backgroundUrl, uploading, saving, editing, error,
    beginEdit, cancelEdit, updateSettings, uploadBackground, removeBackground, resetTheme, applyChanges } = props;
  const [tab, setTab] = useState<typeof TABS[number]>("壁纸");
  const fileInput = useRef<HTMLInputElement>(null);
  const dialog = useRef<HTMLDivElement>(null);
  const closeRef = useRef(() => {});
  closeRef.current = () => { if (!saving) { cancelEdit(); onClose(); } };
  useEffect(() => {
    if (!open || !ready) return;
    beginEdit();
    return cancelEdit;
  }, [open, ready, beginEdit, cancelEdit]);
  useEffect(() => {
    if (!open) return;
    setTab("壁纸");
    const previous = document.activeElement as HTMLElement | null;
    dialog.current?.focus();
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); closeRef.current(); }
      if (event.key !== "Tab") return;
      const elements = Array.from(dialog.current?.querySelectorAll<HTMLElement>('button:not(:disabled), input:not(:disabled):not([type="file"]), [tabindex="0"]') ?? [])
        .filter((element) => element.getAttribute("tabindex") !== "-1");
      if (!elements.length) { event.preventDefault(); return; }
      const first = elements[0], last = elements[elements.length - 1];
      if (event.shiftKey && (document.activeElement === first || document.activeElement === dialog.current)) { event.preventDefault(); last.focus(); }
      else if (!event.shiftKey && (document.activeElement === last || document.activeElement === dialog.current)) { event.preventDefault(); first.focus(); }
    };
    document.addEventListener("keydown", keydown, true);
    return () => { document.removeEventListener("keydown", keydown, true); previous?.focus(); };
  }, [open]);
  if (!open) return null;

  const range = (key: NumericKey, label: string, min: number, max: number, unit = "%", disabled = false) => <div className="appearance-range" key={key}>
    <label htmlFor={`appearance-${key}`}>{label}</label>
    <input id={`appearance-${key}`} type="range" min={min} max={max} value={settings[key]} disabled={disabled}
      onChange={(event) => updateSettings({ [key]: Number(event.target.value) })} />
    <div className="appearance-number"><input type="number" aria-label={`${label}数值`} min={min} max={max} value={settings[key]} disabled={disabled}
      onChange={(event) => { if (event.target.value !== "") updateSettings({ [key]: Number(event.target.value) }); }} /><span>{unit}</span></div>
    <button type="button" className="appearance-reset-one" disabled={disabled} aria-label={`重置${label}`} title={`重置${label}`}
      onClick={() => updateSettings({ [key]: APPEARANCE_DEFAULTS[key] })}>↺</button>
  </div>;
  const position = (event: PointerEvent<HTMLDivElement>) => {
    if (!backgroundUrl || saving || !event.currentTarget.hasPointerCapture?.(event.pointerId)) return;
    const bounds = event.currentTarget.getBoundingClientRect();
    if (bounds.width && bounds.height) updateSettings({ positionX: (event.clientX - bounds.left) / bounds.width * 100,
      positionY: (event.clientY - bounds.top) / bounds.height * 100 });
  };
  const disabled = !ready || !editing || saving;
  return createPortal(<div className="appearance-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) closeRef.current(); }}>
    <div ref={dialog} className="appearance-dialog" role="dialog" aria-modal="true" aria-labelledby="appearance-title" tabIndex={-1}>
      <AppearancePreview hasImage={Boolean(backgroundUrl)} />
      <section className="appearance-controls">
        <header className="appearance-header"><div><h2 id="appearance-title">壁纸与外观</h2><p>调节后即时预览，应用后保存</p></div>
          <button type="button" disabled={saving} onClick={() => closeRef.current()} aria-label="关闭壁纸设置">×</button></header>
        <fieldset className="appearance-options" disabled={disabled}>
          <div className="appearance-presets" aria-label="外观预设">{APPEARANCE_PRESETS.map((preset) => {
            const selected = Object.entries(preset.values).every(([key, value]) => settings[key as keyof AppearanceSettings] === value);
            return <button type="button" key={preset.id} aria-pressed={selected} onClick={() => updateSettings(preset.values)}>
              <span className={`appearance-preset-art ${preset.id}`}><i /><i /></span><b>{preset.label}</b>{selected && <em>✓</em>}
            </button>;
          })}</div>
          <div className="appearance-tabs" role="tablist" aria-label="外观分类">{TABS.map((name, index) => <button type="button" key={name}
            role="tab" id={`appearance-tab-${index}`} aria-controls="appearance-tab-content" aria-selected={tab === name} tabIndex={tab === name ? 0 : -1}
            onClick={() => setTab(name)} onKeyDown={(event) => { if (event.key === "ArrowRight" || event.key === "ArrowLeft") {
              event.preventDefault(); const next = (index + (event.key === "ArrowRight" ? 1 : 2)) % 3;
              setTab(TABS[next]); document.getElementById(`appearance-tab-${next}`)?.focus();
            } }}>{name}</button>)}</div>
          <div className="appearance-scroll" role="tabpanel" id="appearance-tab-content" aria-labelledby={`appearance-tab-${TABS.indexOf(tab)}`}>
            {tab === "壁纸" && <>
              <div className="appearance-upload"><div className="appearance-thumbnail" style={{ backgroundImage: backgroundUrl ? `url("${backgroundUrl}")` : undefined }}>
                {!backgroundUrl && <span>选择一张喜欢的壁纸</span>}</div><div><button type="button" disabled={uploading} onClick={() => fileInput.current?.click()}>{uploading ? "处理中…" : backgroundUrl ? "更换图片" : "上传图片"}</button>
                {backgroundUrl && <button type="button" onClick={removeBackground}>删除背景</button>}</div></div>
              <p className="appearance-help">JPG / PNG / WebP，小于 10 MB · 图片保存在本机</p>
              <input ref={fileInput} type="file" hidden accept="image/jpeg,image/png,image/webp" aria-label="选择壁纸图片" onChange={(event) => {
                const file = event.target.files?.[0]; if (file) void uploadBackground(file); event.target.value = "";
              }} />
              <div className="appearance-segment" aria-label="壁纸显示方式">{(["cover", "contain"] as const).map((fit) => <button type="button" key={fit} aria-pressed={settings.fit === fit} onClick={() => updateSettings({ fit })}>{fit === "cover" ? "填充" : "适应"}</button>)}</div>
              <div className="appearance-position" aria-label="拖动壁纸位置" style={{ backgroundImage: backgroundUrl ? `url("${backgroundUrl}")` : undefined }}
                onPointerDown={(event) => { if (backgroundUrl) { event.currentTarget.setPointerCapture(event.pointerId); position(event); } }}
                onPointerMove={position} onPointerUp={(event) => { if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId); }}>
                <i style={{ left: `${settings.positionX}%`, top: `${settings.positionY}%` }} />
                <span>拖动调整位置</span>
              </div>
              {range("zoom", "缩放", 100, 200)}
              {range("positionX", "水平位置", 0, 100)}
              {range("positionY", "垂直位置", 0, 100)}
              {range("brightness", "亮度", 40, 120)}
              {range("blur", "壁纸模糊", 0, 24, "px")}
              {range("opacity", "壁纸浓度", 0, 100)}
              <label className="appearance-checkbox"><input type="checkbox" checked={settings.autoMask} onChange={(event) => updateSettings({ autoMask: event.target.checked })} />柔和遮罩，减少背景干扰</label>
            </>}
            {tab === "内容卡片" && <><h3>内容卡片</h3><p className="appearance-help">统一调节消息、过程和结果卡片。输入框保留较实的底色。</p>
              {range("cardRadius", "圆角", 0, 28, "px")}</>}
            {tab === "配色" && <><h3>显示模式</h3><div className="appearance-segment">{(["light", "dark", "system"] as const).map((mode, index) => <button type="button" key={mode} aria-pressed={settings.colorMode === mode} onClick={() => updateSettings({ colorMode: mode })}>{["浅色", "深色", "跟随系统"][index]}</button>)}</div>
              <h3>主题配色</h3><div className="appearance-palettes">{COLORS.map((color) => <button type="button" key={color.id} aria-pressed={settings.preset === color.id && !settings.accentColor} onClick={() => updateSettings({ preset: color.id, accentColor: null })}><i style={{ background: color.color }} />{color.label}</button>)}</div>
              <label className="appearance-custom-color">自定义主题色<input type="color" aria-label="自定义主题色" value={settings.accentColor ?? COLORS.find((item) => item.id === settings.preset)!.color} onChange={(event) => updateSettings({ accentColor: event.target.value })} /></label>
              <button type="button" className="appearance-link" onClick={() => updateSettings({ accentColor: null })}>恢复主题配色</button>
            </>}
            <section className="appearance-readability"><h3>内容可读性</h3>
              {range("cardOpacity", "界面底色浓度", 55, 100)}
              {range("cardBlur", "工具区磨砂", 0, 24, "px")}
              <p className="appearance-help">调节消息卡片与工具区的通透感。文字背后保留必要底色，消息间隙与空白区域完整展示壁纸。</p>
            </section>
          </div>
        </fieldset>
        <footer className="appearance-footer">
          {!ready && <p role="status">正在读取外观设置…</p>}
          {error && <p role="alert" className="appearance-error">{error}</p>}
          <div><button type="button" disabled={saving} onClick={() => closeRef.current()}>取消</button>
            <button type="button" disabled={disabled} onClick={resetTheme}>恢复默认</button>
            <button type="button" className="appearance-apply" disabled={disabled || uploading} onClick={() => void applyChanges().then((ok) => { if (ok) onClose(); })}>{saving ? "保存中…" : "应用"}</button></div>
          <small>调节仅临时预览，应用后保存</small>
        </footer>
      </section>
    </div>
  </div>, document.body);
}
