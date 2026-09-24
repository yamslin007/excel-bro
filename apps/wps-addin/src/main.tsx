import React from "react";
import ReactDOM from "react-dom/client";
import WpsProbePanel from "./WpsProbePanel";
import { reportDevProbe } from "./devProbe";
import { ribbon } from "./ribbon";
import "./styles.css";

window.ribbon = ribbon;
reportDevProbe("bootstrap", {
  hasApplication: Boolean(window.Application),
  hasCreateTaskPane: typeof window.Application?.CreateTaskPane === "function",
  hasPluginStorage: Boolean(window.Application?.PluginStorage),
  ribbonRegistered: Boolean(window.ribbon)
});

function Root() {
  if (window.location.hash === "#/taskpane") return <WpsProbePanel />;
  return (
    <main className="bootstrap-shell">
      <span className="probe-logo">EB</span>
      <p>Excel Bro WPS 加载项已初始化。</p>
    </main>
  );
}

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>
);
