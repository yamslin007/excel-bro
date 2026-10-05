/**
 * 宽幅审阅视图 —— WPS 独立宽窗口能力的**实机探针**（只读、不含任何工作表数据）。
 *
 * 目的：确认 `Application.ShowDialog` / `CreateWebDialog` 在当前 WPS 实机上
 * 是否真的存在、能否打开到可用宽度、能否关闭。**这里不读工作簿、不发单元格值**，
 * 只报告"宿主有没有这个能力"——所以它可以在任何工作簿里安全运行。
 *
 * 结论只作三态输出（`requested` / `missing` / `error`），不做任何降级决策；
 * 已存在的窗格内全幅只是明确标注的过渡视图，不能冒充目标三栏宽视图。
 */

import type { WpsApplication, WpsWebDialog, WpsShowDialog } from "./wpsHost";

export type WideViewDialogProbeStatus = "requested" | "missing" | "error";

export type WideViewDialogProbe = {
  status: WideViewDialogProbeStatus;
  /** 宿主自报版本，失败时用来定位"哪个版本没有这个 API"。 */
  hostVersion: string | null;
  /** 实际探测到存在的 API 名（可能多个，便于判断该用哪个）。 */
  presentApis: string[];
  /** 调用后被返回了句柄的 API 名；空表示"有方法但没拿到句柄"。 */
  openedWith: string | null;
  /** 句柄上实际可用的成员名，用来判断能不能关闭。 */
  dialogMembers: string[];
  /** 请求的宽高（逻辑像素；WPS 自行按 DPI 缩放）。 */
  requestedSize: { width: number; height: number };
  /** status=error 时的原始错误文本；这是失败记录的一部分，必须保留。 */
  error: string | null;
  /** 探针不自动关窗，留给用户观察真实宽度与关闭行为。 */
  closed: boolean | null;
};

export const WIDE_VIEW_POC_URL = "poc-dialog.html";
export const WIDE_VIEW_POC_TITLE = "Excel Bro 变更审阅（能力探针）";

/**
 * 目标宽度（CSS 像素）。取 1280 是因为两张并排网格 + 行号列在
 * 1366 宽笔记本上仍能完整显示；低于 960 两张网格就必然要横向滚动。
 */
export const WIDE_VIEW_TARGET_WIDTH = 1280;
export const WIDE_VIEW_TARGET_HEIGHT = 860;

function dialogMembers(dialog: WpsWebDialog | undefined): string[] {
  if (!dialog || typeof dialog !== "object") return [];
  const members = new Set(Object.keys(dialog as object));
  for (const name of ["Close", "Show", "Visible", "ID"] as const) {
    if (name in dialog) members.add(name);
  }
  return [...members].sort();
}

/**
 * 探测并**真实尝试**打开一次对话框。返回值只描述事实，不含判断。
 *
 * 注意：`requested` 只说明宿主接受了打开请求，**不等于真机已显示到目标宽度**。
 * 调用方必须在 `error`/`missing` 时如实报错，**不允许**在这里
 * 回退到 localStorage 快照或窗格内浮层 —— 那会把"没能力"粉饰成"能打开"。
 */
export function probeWideViewDialog(application: WpsApplication, probeNonce?: string): WideViewDialogProbe {
  const hostVersion = String(application.Version ?? "").trim() || null;
  const presentApis: string[] = [];
  for (const name of ["ShowDialog", "CreateWebDialog", "GetWebDialog"] as const) {
    if (typeof application[name] === "function") presentApis.push(name);
  }

  const requestedSize = {
    width: WIDE_VIEW_TARGET_WIDTH,
    height: WIDE_VIEW_TARGET_HEIGHT
  };

  if (!presentApis.includes("ShowDialog") && !presentApis.includes("CreateWebDialog")) {
    return {
      status: "missing",
      hostVersion,
      presentApis,
      openedWith: null,
      dialogMembers: [],
      requestedSize,
      error: "宿主没有暴露 ShowDialog / CreateWebDialog。",
      closed: null
    };
  }

  const probeUrl = new URL(WIDE_VIEW_POC_URL, window.location.href);
  if (probeNonce) probeUrl.searchParams.set("probeNonce", probeNonce);
  const url = probeUrl.toString();
  const attempts: Array<[string, WpsShowDialog]> = [];
  if (typeof application.ShowDialog === "function") {
    attempts.push(["ShowDialog", application.ShowDialog.bind(application)]);
  }
  if (typeof application.CreateWebDialog === "function") {
    attempts.push(["CreateWebDialog", application.CreateWebDialog.bind(application)]);
  }

  let lastError: string | null = null;
  for (const [name, call] of attempts) {
    try {
      const result = call(
        url,
        WIDE_VIEW_POC_TITLE,
        requestedSize.width,
        requestedSize.height,
        false
      );
      if (!result) {
        lastError = `${name} 返回 false/空值，未能确认窗口创建。`;
        continue;
      }
      const dialog = typeof result === "object" ? result : undefined;
      // ShowDialog 的当前官方签名返回 boolean；不能把 true 当成窗口句柄。
      // CreateWebDialog 创建后未必显示，只有句柄时才尝试 Visible/Show。
      if (name === "CreateWebDialog" && !dialog) {
        lastError = "CreateWebDialog 未返回可显示的窗口句柄。";
        continue;
      }
      if (name === "CreateWebDialog" && dialog) {
        if ("Visible" in dialog) dialog.Visible = true;
        else if (typeof dialog.Show === "function") dialog.Show();
        else {
          lastError = "CreateWebDialog 返回句柄，但没有 Visible/Show，无法证明窗口已显示。";
          continue;
        }
      }
      return {
        status: "requested",
        hostVersion,
        presentApis,
        openedWith: name,
        dialogMembers: dialogMembers(dialog),
        requestedSize,
        error: null,
        closed: null
      };
    } catch (error) {
      lastError = `${name} 抛错：${error instanceof Error ? error.message : String(error)}`;
    }
  }

  return {
    status: "error",
    hostVersion,
    presentApis,
    openedWith: null,
    dialogMembers: [],
    requestedSize,
    error: lastError ?? "所有候选 API 都未能打开对话框。",
    closed: null
  };
}
