import { useEffect, useMemo, useRef, useState } from "react";

export interface JsonEditorValidation {
  ok: boolean;
  line?: number;
  column?: number;
  message?: string;
}

interface JsonCodeEditorProps {
  value: string;
  onChange: (text: string) => void;
  disabled?: boolean;
  ariaLabel?: string;
  /** 来自父级的错误提示（如保存时的结构校验失败） */
  externalError?: string | null;
  /** 标记内容相对初始值是否已修改 */
  dirty?: boolean;
}

const SKIN_STORAGE_KEY = "EB_PLAN_EDITOR_CODE_SKIN";

const TOKEN_PATTERN =
  /("(?:[^"\\]|\\.)*")(\s*:)?|\b(true|false)\b|\bnull\b|-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/g;

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * 轻量 JSON 语法高亮：只做 token 上色，不做解析。
 * 使用 textarea + 底层 <pre> 叠加渲染，不引入 Monaco / CodeMirror 类依赖。
 */
export function highlightJsonText(json: string): string {
  let out = "";
  let last = 0;
  TOKEN_PATTERN.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = TOKEN_PATTERN.exec(json)) !== null) {
    out += escapeHtml(json.slice(last, match.index));
    if (match[1]) {
      out += match[2]
        ? `<span class="tk-key">${escapeHtml(match[1])}</span>${escapeHtml(match[2])}`
        : `<span class="tk-str">${escapeHtml(match[1])}</span>`;
    } else if (match[3]) {
      out += `<span class="tk-bool">${escapeHtml(match[3])}</span>`;
    } else if (match[0] === "null") {
      out += `<span class="tk-null">null</span>`;
    } else {
      out += `<span class="tk-num">${escapeHtml(match[0])}</span>`;
    }
    last = match.index + match[0].length;
  }
  out += escapeHtml(json.slice(last));
  return out;
}

export function validateJsonText(text: string): JsonEditorValidation {
  try {
    JSON.parse(text);
    return { ok: true };
  } catch (reason) {
    const message =
      reason instanceof Error ? reason.message : "JSON 格式有误";
    const lineMatch = message.match(/line (\d+) column (\d+)/i);
    if (lineMatch) {
      return {
        ok: false,
        line: Number(lineMatch[1]),
        column: Number(lineMatch[2]),
        message
      };
    }
    const positionMatch = message.match(/position (\d+)/i);
    if (positionMatch) {
      const before = text.slice(0, Number(positionMatch[1])).split("\n");
      return {
        ok: false,
        line: before.length,
        column: before[before.length - 1].length + 1,
        message
      };
    }
    return { ok: false, line: 1, column: 1, message };
  }
}

function fallbackCopy(textarea: HTMLTextAreaElement | null, value: string) {
  if (textarea) {
    textarea.focus();
    textarea.select();
    try {
      document.execCommand("copy");
    } catch {
      // 剪贴板不可用时静默失败
    }
  } else {
    void navigator.clipboard?.writeText(value).catch(() => undefined);
  }
}

/**
 * 计划 JSON 代码编辑器：语法高亮 + 行号 + 本地实时校验 + 格式化/复制。
 *
 * 故意不引入 Monaco / CodeMirror：这里编辑的是几十行 JSON，轻量叠加层足够，
 * 且不会把包体积翻倍。保存前由父级做本地 JSON 语法与白名单结构检查。
 */
export function JsonCodeEditor({
  value,
  onChange,
  disabled,
  ariaLabel,
  externalError,
  dirty
}: JsonCodeEditorProps) {
  const [skin, setSkin] = useState<"light" | "dark">(() => {
    try {
      return localStorage.getItem(SKIN_STORAGE_KEY) === "dark"
        ? "dark"
        : "light";
    } catch {
      return "light";
    }
  });
  const [position, setPosition] = useState({ line: 1, column: 1 });
  const [bannerDismissed, setBannerDismissed] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const highlightRef = useRef<HTMLPreElement | null>(null);
  const gutterRef = useRef<HTMLPreElement | null>(null);

  useEffect(() => {
    try {
      localStorage.setItem(SKIN_STORAGE_KEY, skin);
    } catch {
      // 存储不可用时仅影响皮肤偏好持久化
    }
  }, [skin]);

  const validation = useMemo(() => validateJsonText(value), [value]);
  const highlighted = useMemo(() => highlightJsonText(value), [value]);
  const gutterNumbers = useMemo(() => {
    let count = 1;
    for (let i = 0; i < value.length; i += 1) {
      if (value[i] === "\n") count += 1;
    }
    return Array.from({ length: count }, (_, i) => String(i + 1)).join("\n");
  }, [value]);

  const bannerText = !validation.ok
    ? `第 ${validation.line} 行第 ${validation.column} 列附近有语法错误`
    : (externalError ?? "");

  useEffect(() => {
    setBannerDismissed(false);
  }, [bannerText]);

  const syncScroll = () => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    if (highlightRef.current) {
      highlightRef.current.scrollTop = textarea.scrollTop;
      highlightRef.current.scrollLeft = textarea.scrollLeft;
    }
    if (gutterRef.current) {
      gutterRef.current.style.transform = `translateY(${-textarea.scrollTop}px)`;
    }
  };

  const updatePosition = () => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    const before = textarea.value.slice(0, textarea.selectionStart);
    const parts = before.split("\n");
    setPosition({
      line: parts.length,
      column: parts[parts.length - 1].length + 1
    });
  };

  const handleFormat = () => {
    if (!validation.ok) return;
    try {
      onChange(JSON.stringify(JSON.parse(value), null, 2));
    } catch {
      // 本地校验已拦截，正常不会走到这里
    }
  };

  const handleCopy = () => {
    if (navigator.clipboard?.writeText) {
      navigator.clipboard.writeText(value).catch(() => {
        fallbackCopy(textareaRef.current, value);
      });
    } else {
      fallbackCopy(textareaRef.current, value);
    }
  };

  const showBanner = bannerText !== "" && !bannerDismissed;

  return (
    <div className={`json-editor skin-${skin}`}>
      <div className="json-editor-toolbar">
        <span
          className={`json-editor-status ${validation.ok ? "ok" : "err"}`}
        >
          <i className="json-editor-dot" aria-hidden="true" />
          <span className="json-editor-status-text">
            {validation.ok
              ? "JSON 有效"
              : `格式错误 · 第 ${validation.line} 行`}
          </span>
        </span>
        <span className="json-editor-spacer" />
        <button
          type="button"
          className={`json-editor-tool${skin === "light" ? " active" : ""}`}
          onClick={() => setSkin("light")}
          disabled={disabled || skin === "light"}
          title="浅色代码皮肤"
        >
          浅色
        </button>
        <button
          type="button"
          className={`json-editor-tool${skin === "dark" ? " active" : ""}`}
          onClick={() => setSkin("dark")}
          disabled={disabled || skin === "dark"}
          title="深色代码皮肤"
        >
          深色
        </button>
        <button
          type="button"
          className="json-editor-tool"
          onClick={handleCopy}
          disabled={disabled}
        >
          复制
        </button>
        <button
          type="button"
          className="json-editor-tool"
          onClick={handleFormat}
          disabled={disabled}
        >
          格式化
        </button>
      </div>

      <div className="json-editor-wrap">
        {showBanner && (
          <div className="json-editor-error" role="alert">
            <span>{bannerText}</span>
            <button
              type="button"
              className="json-editor-error-close"
              onClick={() => setBannerDismissed(true)}
              aria-label="关闭错误提示"
            >
              ×
            </button>
          </div>
        )}
        <div className="json-editor-scroll">
          <div className="json-editor-gutter" aria-hidden="true">
            <pre ref={gutterRef}>{gutterNumbers}</pre>
          </div>
          <div className="json-editor-overlay">
            <pre
              ref={highlightRef}
              className="json-editor-hl"
              aria-hidden="true"
              dangerouslySetInnerHTML={{ __html: highlighted }}
            />
            <textarea
              ref={textareaRef}
              value={value}
              onChange={(event) => onChange(event.target.value)}
              onScroll={syncScroll}
              onKeyUp={updatePosition}
              onClick={updatePosition}
              onSelect={updatePosition}
              spellCheck={false}
              disabled={disabled}
              aria-label={ariaLabel}
            />
          </div>
        </div>
      </div>

      <footer className="json-editor-statusbar">
        <span>
          行 {position.line} · 列 {position.column}
        </span>
        <span className="json-editor-spacer" />
        {dirty !== undefined && <span>{dirty ? "已修改" : "未修改"}</span>}
      </footer>
    </div>
  );
}
