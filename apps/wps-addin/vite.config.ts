import path from "node:path";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

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
            console.info(`[wps-request] ${request.method ?? "GET"} ${path} ${userAgent}`);
          }
          next();
        });
      }
    }
  ],
  server: {
    host: "127.0.0.1",
    port: 3889,
    strictPort: true,
    proxy: {
      "/api": {
        target: "http://127.0.0.1:8765",
        changeOrigin: false,
        ws: true
      }
    }
  },
  build: {
    outDir: "dist"
  },
  test: {
    environment: "jsdom"
  }
});
