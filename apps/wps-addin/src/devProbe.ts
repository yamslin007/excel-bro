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
