import { useRef } from "react";
import type { ThemePreset } from "../utils/imageStorage";
import type { useTheme } from "../hooks/useTheme";

type ThemeApi = ReturnType<typeof useTheme>;

interface ThemePanelProps extends ThemeApi {
  open: boolean;
  onClose: () => void;
}

const PRESET_OPTIONS: Array<{
  id: ThemePreset;
  label: string;
  description: string;
}> = [
  { id: "default", label: "默认 Notion", description: "暖灰、浅绿、克制" },
  { id: "warm-orange", label: "温暖橙", description: "柔和橙色背景" },
  { id: "calm-blue", label: "宁静蓝", description: "清爽蓝色背景" },
  { id: "vivid-green", label: "活力绿", description: "品牌绿色增强" }
];

export default function ThemePanel({
  open,
  onClose,
  settings,
  uploading,
  error,
  applyPreset,
  uploadBackground,
  removeBackground,
  setOpacity,
  setAutoMask,
  resetTheme
}: ThemePanelProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);

  return (
    <>
      {open && <button className="theme-panel-backdrop" onClick={onClose} aria-label="关闭主题面板" />}
      <section className={`theme-panel${open ? " open" : ""}`} aria-hidden={!open}>
        <header className="theme-panel-header">
          <div>
            <span className="theme-panel-eyebrow">隐藏彩蛋</span>
            <strong>主题定制</strong>
          </div>
          <button type="button" onClick={onClose} aria-label="关闭主题面板">×</button>
        </header>

        <div className="theme-panel-body">
          <section className="theme-section">
            <h3>预设主题</h3>
            <div className="preset-grid">
              {PRESET_OPTIONS.map((option) => (
                <button
                  key={option.id}
                  type="button"
                  className={`preset-button${settings.preset === option.id ? " selected" : ""}`}
                  onClick={() => void applyPreset(option.id)}
                >
                  <span className={`preset-swatch preset-${option.id}`} />
                  <span>
                    <strong>{option.label}</strong>
                    <small>{option.description}</small>
                  </span>
                </button>
              ))}
            </div>
          </section>

          <section className="theme-section">
            <h3>自定义背景</h3>
            <div className="theme-upload-row">
              <button
                type="button"
                className="secondary-button"
                disabled={uploading}
                onClick={() => fileInputRef.current?.click()}
              >
                {uploading ? "处理中…" : "上传图片"}
              </button>
              <span>支持 JPG / PNG / WebP，小于 10MB</span>
              <input
                ref={fileInputRef}
                className="theme-file-input"
                type="file"
                accept="image/jpeg,image/png,image/webp"
                onChange={(event) => {
                  const file = event.target.files?.[0];
                  if (file) void uploadBackground(file);
                  event.target.value = "";
                }}
              />
            </div>

            <label className="opacity-control">
              <span>背景透明度</span>
              <input
                type="range"
                min="0"
                max="100"
                value={settings.opacity}
                onChange={(event) => void setOpacity(Number(event.target.value))}
              />
              <b>{settings.opacity}%</b>
            </label>

            <label className="theme-mask-toggle">
              <input
                type="checkbox"
                checked={settings.autoMask}
                onChange={(event) => void setAutoMask(event.target.checked)}
              />
              <span>自动遮罩，确保文字可读</span>
            </label>
          </section>

          {error && <p className="theme-error">{error}</p>}

          <section className="theme-preview" aria-label="主题预览">
            <div className="theme-preview-bubble">这条消息会按新主题显示。</div>
            <button type="button" className="theme-preview-button">按钮示例</button>
          </section>
        </div>

        <footer className="theme-actions">
          <button type="button" className="secondary-button" onClick={() => void resetTheme()}>
            重置
          </button>
          <button type="button" className="secondary-button" onClick={() => void removeBackground()}>
            删除背景
          </button>
          <button type="button" className="primary-button" onClick={onClose}>
            应用主题
          </button>
        </footer>
      </section>
    </>
  );
}
