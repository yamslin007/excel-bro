import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./styles.css";
// 共享运行时：自定义函数与任务窗格同页加载，必须在这里注册 EB()
import "./functions/functions";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

