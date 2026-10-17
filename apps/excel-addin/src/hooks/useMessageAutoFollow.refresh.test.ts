// @vitest-environment node
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import react from "@vitejs/plugin-react";
import { createServer } from "vite";
import { describe, expect, it } from "vitest";

describe("message auto-follow development refresh", () => {
  it.each(["excel-addin", "wps-addin"])("gives the shared hook a refresh signature in %s", async (app) => {
    const directory = new URL("./", import.meta.url);
    const hookPath = fileURLToPath(new URL("useMessageAutoFollowRuntime.tsx", directory)).replaceAll("\\", "/");
    const appRoot = fileURLToPath(new URL(`../../../${app}/`, import.meta.url));
    const server = await createServer({
      configFile: false,
      root: appRoot,
      // Never re-optimize dependencies into the running add-in's dev cache.
      cacheDir: `${appRoot}/node_modules/.vite-refresh-tests`,
      plugins: [react()],
      server: { middlewareMode: true },
      optimizeDeps: { noDiscovery: true, include: [] }
    });
    try {
      const transformed = await server.transformRequest(`/@fs/${hookPath}`);
      expect(transformed?.code).toContain("$RefreshSig$");
      expect(transformed?.code).toMatch(/_s\w*\(useMessageAutoFollow,/);
      const entryPath = fileURLToPath(new URL("useMessageAutoFollow.ts", directory)).replaceAll("\\", "/");
      const entry = await server.transformRequest(`/@fs/${entryPath}`);
      expect(entry?.code).toContain("useMessageAutoFollowRuntime.tsx");

      // A previously unsigned custom hook must invalidate the component's
      // saved hook order, rather than reuse it and crash on the added useRef.
      const refresh = createRequire(import.meta.url)("react-refresh/runtime");
      const key = transformed!.code.match(/_s\w*\(useMessageAutoFollow,\s*("(?:[^"\\]|\\.)*")/)?.[1];
      expect(key).toBeTruthy();
      const legacyHook = () => {};
      const nextHook = () => {};
      refresh.createSignatureFunctionForTransform()(nextHook, JSON.parse(key!));
      const PreviousPanel = () => null;
      const UpdatedPanel = () => null;
      refresh.createSignatureFunctionForTransform()(PreviousPanel, "same-panel", false, () => [legacyHook]);
      refresh.collectCustomHooksForSignature(PreviousPanel);
      refresh.register(PreviousPanel, `auto-follow-refresh-${app}`);
      refresh.createSignatureFunctionForTransform()(UpdatedPanel, "same-panel", false, () => [nextHook]);
      refresh.register(UpdatedPanel, `auto-follow-refresh-${app}`);
      const update = refresh.performReactRefresh();
      expect(update.staleFamilies.has(refresh.getFamilyByType(UpdatedPanel))).toBe(true);
      expect(update.updatedFamilies.has(refresh.getFamilyByType(UpdatedPanel))).toBe(false);
    } finally {
      await server.close();
    }
  });
});
