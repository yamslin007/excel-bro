/**
 * WPS 专注窗口的**宿主适配层**：只负责打开 / 关闭 / 去重，不传内容。
 *
 * 与开发期能力探针（`wpsWideViewPoc.ts`）不同：这里打开的是正式页面
 * `focus.html?session=<id>`，窗口会话标识进 URL，握手由
 * `wpsFocusChannel.ts` + 共享协议完成。**`requested` 只表示宿主接受了打开请求**，
 * 不表示窗口已显示、也不表示通信成功 —— "已就绪"只能由子窗口的 `hello` 确认。
 *
 * 会话隔离方式：每次打开生成一个随机 `sessionId`，页面、通道与（可能写下的）打开时
 * 快照都按它隔离；没有全局快照、没有固定频道名。
 */

import type { WpsApplication, WpsWebDialog } from "./wpsHost";

export const WPS_FOCUS_PAGE = "focus.html";
export const WPS_FOCUS_TITLE = "Excel Bro · 专注窗口";
/** 逻辑像素（WPS 文档口径），宿主自行按 DPI 缩放。 */
export const WPS_FOCUS_WINDOW_SIZE = { width: 1180, height: 820 } as const;

export type WpsFocusOpenOutcome =
  | {
      status: "requested";
      openedWith: "CreateWebDialog" | "ShowDialog";
      /** `ShowDialog` 只返回 boolean —— 没有句柄，不能据此宣称能关闭窗口。 */
      handle: WpsWebDialog | null;
      url: string;
    }
  | { status: "missing"; reason: string }
  | { status: "error"; reason: string };

export function focusWindowUrl(sessionId: string, base?: string): string {
  const url = new URL(WPS_FOCUS_PAGE, base ?? window.location.href);
  url.searchParams.set("session", sessionId);
  return url.toString();
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error ?? "");
}

/**
 * 尝试真正打开一次专注窗口。
 *
 * 顺序：`CreateWebDialog`（可能返回句柄，父窗格随后可主动关闭）→ `ShowDialog`
 * （只返回 boolean）。两个 API 都失败才报 `error`；一个都没有报 `missing`。
 * 任何一条路径都不会在这里虚构成功 —— 返回值只描述"请求是否被接受"。
 */
export function openWpsFocusWindow(
  application: WpsApplication,
  options: { sessionId: string; url?: string }
): WpsFocusOpenOutcome {
  const url = options.url ?? focusWindowUrl(options.sessionId);
  const errors: string[] = [];

  if (typeof application.CreateWebDialog === "function") {
    try {
      const result = application.CreateWebDialog(
        url,
        WPS_FOCUS_TITLE,
        WPS_FOCUS_WINDOW_SIZE.width,
        WPS_FOCUS_WINDOW_SIZE.height,
        false
      );
      if (result && typeof result === "object") {
        const handle = result as WpsWebDialog;
        try {
          if ("Visible" in handle) handle.Visible = true;
          else if (typeof handle.Show === "function") handle.Show();
        } catch (error) {
          errors.push(`CreateWebDialog 显示失败：${errorText(error)}`);
        }
        return { status: "requested", openedWith: "CreateWebDialog", handle, url };
      }
      errors.push("CreateWebDialog 没有返回窗口句柄，无法确认窗口已创建。");
    } catch (error) {
      errors.push(`CreateWebDialog 抛错：${errorText(error)}`);
    }
  }

  if (typeof application.ShowDialog === "function") {
    try {
      const accepted = application.ShowDialog(
        url,
        WPS_FOCUS_TITLE,
        WPS_FOCUS_WINDOW_SIZE.width,
        WPS_FOCUS_WINDOW_SIZE.height,
        false
      );
      // 官方签名返回 boolean；true 只代表宿主接受了请求，不代表窗口已显示。
      if (accepted) {
        return { status: "requested", openedWith: "ShowDialog", handle: null, url };
      }
      errors.push("ShowDialog 返回 false，窗口未创建。");
    } catch (error) {
      errors.push(`ShowDialog 抛错：${errorText(error)}`);
    }
  }

  if (!errors.length) {
    return {
      status: "missing",
      reason: "宿主没有暴露 ShowDialog / CreateWebDialog，当前 WPS 版本无法打开独立窗口。"
    };
  }
  return { status: "error", reason: errors.join("；") };
}

/**
 * 关闭一个**有句柄**的专注窗口。`ShowDialog` 打开的窗口没有句柄，返回 `false`：
 * 那时关闭按钮由子页面自己执行（`window.close()`），父窗格只能通过 `closing`
 * 消息告知对方，不能假装自己关掉了窗口。
 */
export function closeWpsFocusWindow(handle: WpsWebDialog | null): boolean {
  if (handle && typeof handle.Close === "function") {
    try {
      handle.Close();
      return true;
    } catch {
      return false;
    }
  }
  return false;
}
