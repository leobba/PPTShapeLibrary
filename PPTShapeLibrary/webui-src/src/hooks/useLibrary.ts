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
  const seqRef = useRef(0); // 请求序号：丢弃过期响应，防快速切组/筛选时旧结果覆盖新结果
  const pageSizeRef = useRef(200); // 每页条数（取自设置 pageSize，默认 200）

  const reload = useCallback(
    (resetPage = true) => {
      if (resetPage) {
        page.current = 0;
        // 筛选/切组/搜索会改变可见集：清空旧选中，防止"看不到却可批量删除"的误删
        setSelected(new Set());
      }
      const offset = resetPage ? 0 : page.current;
      const seq = ++seqRef.current;
      setLoading(true);
      bridge
        .list(filter, category, kind, tag, group, sort, offset, pageSizeRef.current)
        .then((d) => {
          if (seq !== seqRef.current) return; // 过期响应丢弃
          setLoading(false);
          setFilteredTotal(d.filteredTotal);
          setTotal(d.total);
          if (resetPage) setItems(d.items || []);
          else setItems((prev) => [...prev, ...(d.items || [])]);
          page.current = offset + (d.items ? d.items.length : 0);
        })
        .catch((e) => {
          if (seq !== seqRef.current) return;
          setLoading(false);
          console.error("list error", e);
        });
    },
    [filter, category, kind, tag, group, sort]
  );

  // init 加载（失败自动重试最多 3 次，避免大库冷启动偶发超时后永久骨架）
  useEffect(() => {
    let tries = 0;
    let alive = true;
    const load = () => {
      bridge
        .init()
        .then((d) => {
          if (alive) {
            setInit(d);
            const ps = (d as any)?.pageSize;
            if (typeof ps === "number" && ps > 0 && ps <= 500) pageSizeRef.current = ps;
          }
        })
        .catch(() => {
          tries++;
          if (alive && tries < 3) window.setTimeout(load, 1200);
        });
    };
    load();
    return () => {
      alive = false;
    };
  }, []);

  // 首次 init 到位后加载一次列表（仅一次；后续 refresh/refreshMeta 只更新元数据，不再整表跳回第一页）
  const didInitialLoad = useRef(false);
  useEffect(() => {
    if (!init) return;
    if (!didInitialLoad.current) {
      didInitialLoad.current = true;
      reload(true);
    }
  }, [init, reload]);

  // 组/分类有效性：当前选中的组或分类已被删除时回退（避免请求旧名返回空、筛选态悬空）
  useEffect(() => {
    if (!init) return;
    if (group !== "全部" && !(init.groups || []).includes(group)) {
      setGroup("全部");
    }
  }, [init, group]); // eslint-disable-line react-hooks/exhaustive-deps

  // 防抖搜索（筛选/排序变化 150ms 后重载；不依赖 init——init 仅首次驱动加载）
  useEffect(() => {
    if (!didInitialLoad.current) return;
    if (debounceRef.current) window.clearTimeout(debounceRef.current);
    debounceRef.current = window.setTimeout(() => reload(true), 150);
    return () => {
      if (debounceRef.current) window.clearTimeout(debounceRef.current);
    };
  }, [filter, category, kind, tag, group, sort, reload]);

  // 下行 refresh（C# 数据变化）：仅刷新元数据(侧栏/计数)，不整表重载
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
