import path from "node:path";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import devCerts from "office-addin-dev-certs";

export default defineConfig(async ({ command }) => {
  const isVitest = process.env.VITEST === "true";
  const https = command === "serve" && !isVitest
    ? await devCerts.getHttpsServerOptions()
    : undefined;

  return {
    plugins: [react()],
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
          commands: path.resolve(__dirname, "commands.html")
        }
      }
    }
  };
});
