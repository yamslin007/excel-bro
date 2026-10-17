export function reportDevProbe(
  stage: "bootstrap" | "taskpane",
  details: Record<string, boolean | number>
): void {
  if (!import.meta.env.DEV) return;
  void fetch("/__wps_probe", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ stage, ...details })
  }).catch(() => undefined);
}

/**
 * 上报**结构化逐项**探测报告（用于 WPS 格式能力探测）。
 *
 * 与 `reportDevProbe`（只收 boolean/number 汇总）不同，这里把整份逐项报告作为
 * JSON 字符串打进 `__wps_probe` 的 body，供 dev server console 直接 `[wps-bridge]`
 * 打印出完整 JSON —— 用户可复制。仅 DEV 下生效。
 */
export function reportDevProbeDetail(
  stage: "taskpane",
  details: Record<string, boolean | number | string>
): void {
  if (!import.meta.env.DEV) return;
  void fetch("/__wps_probe", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ stage, ...details })
  }).catch(() => undefined);
}
