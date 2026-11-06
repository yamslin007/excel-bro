import fs from "node:fs";
import path from "node:path";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

/**
 * 开发端口只有一个来源：`config/capabilities.json` 的 `devServers`。
 *
 * 后端来源白名单从同一份配置派生，启动器（`scripts/start_wps_dev.mjs`）按声明顺序
 * 实测挑出真正能绑定的那个，再经 `EB_WPS_DEV_PORT` 传进来。这里**拒绝**列表外的端口：
 * 那会让任务窗格与后端白名单错位，只剩 `Failed to fetch`。
 * 也不用端口扫描（wpsjs + portfinder 就是这样静默漂到 3987 的）。
 */
const CAPABILITIES_PATH = path.resolve(__dirname, "../../config/capabilities.json");
const DEV_SERVERS = JSON.parse(fs.readFileSync(CAPABILITIES_PATH, "utf8"))?.devServers ?? {};
const WPS_PORT_CANDIDATES: number[] = DEV_SERVERS.wpsPortCandidates ?? [];

function resolveDevPort(): number {
  if (!Array.isArray(WPS_PORT_CANDIDATES) || WPS_PORT_CANDIDATES.length === 0) {
    throw new Error(
      "config/capabilities.json 缺少 devServers.wpsPortCandidates，无法确定 WPS 开发端口"
    );
  }
  const raw = process.env.EB_WPS_DEV_PORT;
  if (!raw) {
    // 直接跑 `vite` 时退到首个候选：仍是声明过的端口，不做扫描、不漂移。
    console.info(
      `[wps-dev] 未指定 EB_WPS_DEV_PORT，使用候选列表首个端口 ${WPS_PORT_CANDIDATES[0]}；` +
        "推荐用 npm run dev:wps（自动挑可用端口并同步登记与后端白名单）。"
    );
    return WPS_PORT_CANDIDATES[0];
  }
  const port = Number(raw);
  if (!Number.isInteger(port) || !WPS_PORT_CANDIDATES.includes(port)) {
    throw new Error(
      `EB_WPS_DEV_PORT=${raw} 不在 devServers.wpsPortCandidates（${WPS_PORT_CANDIDATES.join("、")}）内；` +
        "后端只允许该列表内的来源，用列表外端口会直接被拒。"
    );
  }
  return port;
}

export default defineConfig({
  base: "./",
  resolve: {
    alias: {
      "@excel-bro/core": path.resolve(__dirname, "../../packages/core/src")
    },
    // 两端共用 packages/core 的 React hooks，双份 React 会导致 invalid hook call
    dedupe: ["react", "react-dom"]
  },
  plugins: [
    react(),
    {
      name: "excel-bro-wps-request-log",
      configureServer(server) {
        server.middlewares.use((request, _response, next) => {
          const path = request.url?.split("?", 1)[0];
          if (path === "/__wps_probe" && request.method === "POST") {
            let body = "";
            request.setEncoding("utf8");
            request.on("data", (chunk) => { body += chunk; });
            request.on("end", () => {
              console.info(`[wps-bridge] ${body}`);
              _response.statusCode = 204;
              _response.end();
            });
            return;
          }
          if (path && !path.startsWith("/@vite") && !path.startsWith("/node_modules/")) {
            const userAgent = request.headers["user-agent"]?.includes("QtWebEngine")
              ? "QtWebEngine"
              : request.headers["user-agent"]?.includes("Chrome")
                ? "Chromium"
                : "other";
            // 带 Origin 的请求能直接说明"任务窗格是从哪个端口加载的"，是端口对齐的现场证据；
            // Origin 只是回环来源，不含工作簿名/表名/单元格值。
            const origin = request.headers.origin;
            console.info(
              `[wps-request] ${request.method ?? "GET"} ${path} ${userAgent}` +
                `${origin ? ` from=${origin}` : ""}`
            );
          }
          next();
        });
      }
    }
  ],
  server: {
    host: "127.0.0.1",
    port: resolveDevPort(),
    // 端口不许漂移：绑定失败就报错退出，而不是悄悄换一个后端不认识的端口。
    strictPort: true,
    proxy: {
      "/api": {
        // 默认连本机 8765；原生验收可用 `EB_TEST_API_TARGET` 指到独立测试服务端
        // （例如 http://127.0.0.1:8877），让真实 WPS 窗格连接测试实例而不是碰用户 8765。
        target: process.env.EB_TEST_API_TARGET ?? "http://127.0.0.1:8765",
        changeOrigin: false,
        ws: true
      }
    }
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      // 多页构建：任务窗格 + 专注窗口 + 宽视图能力探针。
      // `focus.html` 是专注窗口的独立入口（`ShowDialog("./focus.html?session=…")` 打开），
      // `poc-dialog.html` 是只读能力探针，两者都必须进生产产物，否则部署后打开的是空页面。
      input: {
        taskpane: path.resolve(__dirname, "index.html"),
        focus: path.resolve(__dirname, "focus.html"),
        "poc-dialog": path.resolve(__dirname, "poc-dialog.html")
      }
    }
  },
  test: {
    environment: "jsdom",
    // 未替身的网络出口直接让用例失败（见 `@excel-bro/core/testNetworkGate`）。
    setupFiles: ["./src/testSetup.ts"]
  }
});
