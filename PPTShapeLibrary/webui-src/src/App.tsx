import { useEffect, useState } from "react";
import LibraryView from "@/components/library/LibraryView";
import ToolView from "@/components/library/ToolView";
import { bridge } from "@/lib/bridge";

function applyTheme(theme: string) {
  document.documentElement.classList.toggle("dark", theme === "dark");
}

export default function App({ view }: { view: string }) {
  const [theme, setTheme] = useState<string>("light");

  // 读取设置主题（与旧逻辑一致）
  useEffect(() => {
    bridge
      .init()
      .then((d) => {
        setTheme(d.theme || "light");
        applyTheme(d.theme || "light");
      })
      .catch(() => {});
  }, []);

  useEffect(() => {
    bridge.on("settingsChanged", (s: any) => {
      if (s?.theme) {
        setTheme(s.theme);
        applyTheme(s.theme);
      }
    });
  }, []);

  if (view && view !== "library") {
    return <ToolView view={view} theme={theme} setTheme={setTheme} />;
  }
  return <LibraryView theme={theme} setTheme={setTheme} />;
}
