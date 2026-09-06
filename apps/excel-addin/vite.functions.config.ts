import path from "node:path";
import { defineConfig } from "vite";

/**
 * Excel Custom Functions 的 Script URL 由宿主直接加载，必须是可执行的
 * JavaScript bundle，而不是 functions.html 中的 module script。
 */
export default defineConfig({
  build: {
    outDir: "dist",
    emptyOutDir: false,
    lib: {
      entry: path.resolve(__dirname, "src/functions/functions.ts"),
      formats: ["iife"],
      fileName: () => "functions.js",
      name: "ExcelBroFunctions"
    },
    rollupOptions: {
      output: {
        inlineDynamicImports: true,
        extend: false
      }
    }
  }
});