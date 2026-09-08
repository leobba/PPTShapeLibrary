import LibraryView from "@/components/library/LibraryView";
import ToolView from "@/components/library/ToolView";

// 深色模式已移除（第二十一批）：界面固定浅色

export default function App({ view }: { view: string }) {
  if (view && view !== "library") {
    return <ToolView view={view} />;
  }
  return <LibraryView />;
}
