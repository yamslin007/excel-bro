// 通用小工具（从 App.tsx 抽出）：剪贴板、LRU 缓存、组合输入尺寸常量等。
import type { ModelConnectionDraft } from "./hooks/useModelManagement";
import { DataToolExecutionError } from "./dataTools";

// 复制文本到剪贴板：优先 navigator.clipboard，退化到 execCommand。
export async function copyTextToClipboard(text: string): Promise<boolean> {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    const copied = document.execCommand("copy");
    textarea.remove();
    return copied;
  } catch {
    return false;
  }
}

export const RESULT_CACHE_LIMIT = 24;
export const PROMPT_KEY_CACHE_LIMIT = 48;
// 模块加载时生成一次的会话 id；换会话（刷新页面）即失效缓存。
export const CACHE_SESSION_ID =
  globalThis.crypto?.randomUUID?.() ?? `session-${Date.now()}`;

// Map 迭代顺序即插入顺序，超限时删最早的键 => 朴素 LRU。
export function lruSet<K, V>(
  cache: Map<K, V>,
  key: K,
  value: V,
  limit: number
): void {
  if (cache.has(key)) cache.delete(key);
  cache.set(key, value);
  while (cache.size > limit) {
    const oldest = cache.keys().next().value as K | undefined;
    if (oldest === undefined) break;
    cache.delete(oldest);
  }
}

export function lruGet<K, V>(cache: Map<K, V>, key: K): V | undefined {
  const value = cache.get(key);
  if (value !== undefined) {
    cache.delete(key);
    cache.set(key, value);
  }
  return value;
}

export const COMPOSER_MAX_HEIGHT = 156;
export const COMPOSER_MIN_HEIGHT = 44;
// 首次引导第一步「拉宽窗格」的阈值：窗格宽度达到该值即视为已拉宽，
// 之后承接模型引导。与内联窄窗格判断共用，避免两处口径漂移。
export const PANE_WIDEN_THRESHOLD = 380;

export function emptyModelConnectionDraft(): ModelConnectionDraft {
  return {
    id: null,
    label: "",
    baseUrl: "",
    modelId: "",
    apiKey: "",
    clearApiKey: false,
    supportsVision: false
  };
}

// isAbortError：区分「用户主动打断」和真正的错误。fetch 被 abort 时抛 AbortError，
// 本地工具取消抛 DataToolExecutionError(code=CANCELLED)。这两种都不该弹错误提示。
export function isAbortError(reason: unknown): boolean {
  if (reason instanceof DOMException && reason.name === "AbortError") {
    return true;
  }
  if (
    reason instanceof DataToolExecutionError &&
    reason.code === "CANCELLED"
  ) {
    return true;
  }
  return false;
}

// 文件夹模式下的工作表标识键：文件 id + 表名合成，供 folderSheetKeys 记录勾选。
export function folderSheetKey(fileId: string, sheetName: string): string {
  return `${fileId}\u0000${sheetName}`;
}

// /function 试算结果是否算"验算不通过"：Excel 错误值（#REF! 等）或试算异常。
export function formulaTrialFailed(result: string): boolean {
  const trimmed = result.trim();
  return (
    trimmed.startsWith("#") ||
    trimmed.startsWith("（试算失败") ||
    trimmed.startsWith("(试算失败")
  );
}
