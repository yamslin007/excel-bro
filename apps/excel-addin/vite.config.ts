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
    plugins: [
      react(),
      // 临时调试：记录 Excel 对函数元数据等资源的请求
      {
        name: "eb-request-log",
        configureServer(server: ViteDevServer) {
          server.middlewares.use(
            (
              req: Connect.IncomingMessage,
              _res: import("node:http").ServerResponse,
              next: Connect.NextFunction
            ) => {
            if (req.url && !req.url.includes("request-log")) {
              fs.appendFileSync(
                path.resolve(__dirname, "request-log.txt"),
                `${new Date().toISOString()} ${req.headers["user-agent"] ?? "?"} ${req.url}\n`
              );
            }
            next();
          });
        }
      }
    ],
    server: {
      host: "localhost",
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
          functions: path.resolve(__dirname, "functions.html")
        }
      }
    }
  };
});
