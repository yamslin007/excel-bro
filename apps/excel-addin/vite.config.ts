import fs from "node:fs";
import path from "node:path";
import { defineConfig, type Connect, type ViteDevServer } from "vite";
import react from "@vitejs/plugin-react";
import devCerts from "office-addin-dev-certs";

export default defineConfig(async ({ command }) => {
  const isVitest = process.env.VITEST === "true";
  const https = command === "serve" && !isVitest
    ? await devCerts.getHttpsServerOptions()
    : undefined;

  return {
    resolve: {
      alias: {
        "@excel-bro/core": path.resolve(__dirname, "../../packages/core/src")
      },
      // 两端共用 packages/core 的 React hooks，双份 React 会导致 invalid hook call
      dedupe: ["react", "react-dom"]
    },
    plugins: [
      react(),
      // 临时调试：记录 Excel 对函数元数据等资源的请求
      {
        name: "eb-request-log",
        configureServer(server: ViteDevServer) {
          server.middlewares.use(
            (
              req: Connect.IncomingMessage,
              res: import("node:http").ServerResponse,
              next: Connect.NextFunction
            ) => {
            if (req.url && !req.url.includes("request-log")) {
              fs.appendFileSync(
                path.resolve(__dirname, "request-log.txt"),
                `${new Date().toISOString()} ${req.headers["user-agent"] ?? "?"} ${req.url}\n`
              );
            }
            // Excel 的 Custom Functions Script URL 必须返回 JavaScript，不能让 Vite 将 /functions.js 回退到 index.html。
            if (req.url?.split("?", 1)[0] === "/functions.js") {
              const bundlePath = path.resolve(__dirname, "dist/functions.js");
              if (fs.existsSync(bundlePath)) {
                res.statusCode = 200;
                res.setHeader("Content-Type", "application/javascript; charset=utf-8");
                res.end(fs.readFileSync(bundlePath));
                return;
              }
            }
            next();
          });
        }
      }
    ],
    server: {
      host: "127.0.0.1",
      port: 3000,
      strictPort: true,
      https,
      cors: true
    },
    build: {
      outDir: "dist",
      rollupOptions: {
        input: {
          taskpane: path.resolve(__dirname, "index.html"),
          focus: path.resolve(__dirname, "focus.html"),
          commands: path.resolve(__dirname, "commands.html"),
          functions: path.resolve(__dirname, "functions.html"),
          planEditor: path.resolve(__dirname, "plan-editor.html")
        }
      }
    },
    test: {
      // 未替身的网络出口直接让用例失败（见 `@excel-bro/core/testNetworkGate`）。
      setupFiles: ["./src/testSetup.ts"]
    }
  };
});
