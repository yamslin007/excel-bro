import React from "react";
import ReactDOM from "react-dom/client";
import { AppRoutes } from "./AppRoutes";
import { reportDevProbe } from "./devProbe";
import { ribbon } from "./ribbon";
// WPS 与 Excel 任务窗格共用同一套视觉基线。宿主差异只留在数据读取和
// 写入适配器中，不再通过另一套产品样式制造两个界面。
import "../../excel-addin/src/styles.css";
import "./styles.css";

window.ribbon = ribbon;
reportDevProbe("bootstrap", {
  hasApplication: Boolean(window.Application),
  hasCreateTaskPane: typeof window.Application?.CreateTaskPane === "function",
  hasPluginStorage: Boolean(window.Application?.PluginStorage),
  ribbonRegistered: Boolean(window.ribbon)
});

const rootElement = document.getElementById("root");
if (rootElement) {
  ReactDOM.createRoot(rootElement).render(
    <React.StrictMode>
      <AppRoutes />
    </React.StrictMode>
  );
}
