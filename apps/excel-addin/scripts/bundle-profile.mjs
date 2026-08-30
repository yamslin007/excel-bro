import path from "node:path";
import { build } from "vite";
import { visualizer } from "rollup-plugin-visualizer";

const projectRoot = process.cwd();
const addinRoot = path.resolve(projectRoot, "apps/excel-addin");
const reportRoot = path.resolve(projectRoot, "docs/performance/bundle-profile");

await build({
  root: addinRoot,
  configFile: path.resolve(addinRoot, "vite.config.ts"),
  plugins: [
    visualizer({
      filename: path.join(reportRoot, "bundle-visualizer.html"),
      title: "Excel Bro bundle profile",
      template: "treemap",
      gzipSize: true,
      brotliSize: true
    }),
    visualizer({
      filename: path.join(reportRoot, "bundle-stats.json"),
      template: "raw-data",
      gzipSize: true,
      brotliSize: true
    })
  ],
  build: {
    outDir: path.resolve(addinRoot, "dist-profile"),
    emptyOutDir: true
  }
});
