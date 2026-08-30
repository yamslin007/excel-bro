import React from "react";
import ReactDOM from "react-dom/client";
import PlanEditorDialog from "./components/PlanEditorDialog";
import "./styles.css";

const root = document.getElementById("plan-editor-root");
if (root) {
  ReactDOM.createRoot(root).render(
    <React.StrictMode>
      <PlanEditorDialog />
    </React.StrictMode>
  );
}
