import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@/index.css";
import App from "@/App";
import { DialogsHost } from "@/lib/dialogs";
import { ProgressHost } from "@/lib/progress";
import { installGlobalErrorReporting } from "@/lib/errorReporting";

const params = new URLSearchParams(window.location.search);
const view = params.get("view") || ""; // tool1..tool12 / settings

// 兜住"漏了 .catch 的桥调用"：这类失败以前是静默的（用户以为点了没反应）。
// 统一写进宿主日志 + 顶部提示条，不再无声无息。
installGlobalErrorReporting();

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App view={view} />
    <DialogsHost />
    <ProgressHost />
  </StrictMode>
);
