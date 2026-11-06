import React from "react";
import ReactDOM from "react-dom/client";
import WpsFocusWorkspace from "./WpsFocusWorkspace";
import "../../excel-addin/src/styles.css";
import "../../excel-addin/src/focus.css";
import "./styles.css";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <WpsFocusWorkspace />
  </React.StrictMode>
);
