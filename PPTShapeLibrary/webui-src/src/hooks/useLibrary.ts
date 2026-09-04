import { useCallback, useEffect, useRef, useState } from "react";
import { bridge, InitData, ListItem } from "@/lib/bridge";

export type SortMode = "最新" | "最旧" | "名称" | "最近使用" | "最常用" | "自定义";

export interface LibraryState {
  init: InitData | null;
  group: string; // "全部" 或组名
  category: string;
  kind: string; // 全部/Shape/Style
  filter: string;
  tag: string;
  sort: SortMode;
  items: ListItem[];
  filteredTotal: number;
  total: number;
  loading: boolean;
  selected: Set<string>;
}

export function useLibrary() {
  const [init, setInit] = useState<InitData | null>(null);
  const [group, setGroup] = useState("全部");
  const [category, setCategory] = useState("全部");
  const [kind, setKind] = useState("全部");
  const [filter, setFilter] = useState("");
  const [tag, setTag] = useState("");
  const [sort, setSort] = useState<SortMode>("最新");
  const [items, setItems] = useState<ListItem[]>([]);
  const [filteredTotal, setFilteredTotal] = useState(0);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(false);
  const [selected, setSelected] = useState<Set<string>>(new Set());

  const page = useRef(0);
  const debounceRef = useRef<number | undefined>(undefined);

  const reload = useCallback(
    (resetPage = true) => {
      if (resetPage) page.current = 0;
      const offset = resetPage ? 0 : page.current;
      setLoading(true);
      bridge
        .list(filter, category, kind, tag, group, sort, offset, 200)
        .then((d) => {
          setLoading(false);
          setFilteredTotal(d.filteredTotal);
          setTotal(d.total);
          if (resetPage) setItems(d.items || []);
          else setItems((prev) => [...prev, ...(d.items || [])]);
          page.current = offset + (d.items ? d.items.length : 0);
        })
        .catch((e) => {
          setLoading(false);
          console.error("list error", e);
        });
    },
    [filter, category, kind, tag, group, sort]
  );

  // init 加载
  useEffect(() => {
    bridge.init().then((d) => setInit(d));
  }, []);

  // 防抖搜索
  useEffect(() => {
    if (!init) return;
    if (debounceRef.current) window.clearTimeout(debounceRef.current);
    debounceRef.current = window.setTimeout(() => reload(true), 150);
    return () => {
      if (debounceRef.current) window.clearTimeout(debounceRef.current);
    };
  }, [filter, category, kind, tag, group, sort, init, reload]);

  // 下行 refresh（C# 数据变化）
  useEffect(() => {
    const fn = () => bridge.init().then(setInit).catch(() => {});
    bridge.on("refresh", fn);
    bridge.on("settingsChanged", (s) => setInit((prev) => (prev ? { ...prev, ...s } : prev)));
    return () => {
      // bridge.on 无退订 API：保留监听（面板生命周期=页面生命周期，可接受）
    };
  }, []);
  // 注意：首次 init 到达后由上方"防抖 effect（依赖 init）"触发首次列表加载，无需单独 effect，避免重复请求

  const toggleSelect = useCallback((id: string) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }, []);

  const clearSelection = useCallback(() => setSelected(new Set()), []);

  return {
    init,
    items,
    filteredTotal,
    total,
    loading,
    selected,
    group,
    category,
    kind,
    filter,
    tag,
    sort,
    setGroup: (g: string) => {
      setGroup(g);
      setCategory("全部");
    },
    setCategory,
    setKind,
    setFilter,
    setTag,
    setSort,
    toggleSelect,
    clearSelection,
    reload,
    loadMore: () => {
      if (!loading && page.current < filteredTotal) reload(false);
    },
    refreshMeta: () => bridge.init().then(setInit).catch(() => {}),
  };
}
