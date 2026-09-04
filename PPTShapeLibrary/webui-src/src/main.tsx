import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@/index.css";
import App from "@/App";

const params = new URLSearchParams(window.location.search);
const view = params.get("view") || ""; // tool1..tool12 / settings

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App view={view} />
  </StrictMode>
);
