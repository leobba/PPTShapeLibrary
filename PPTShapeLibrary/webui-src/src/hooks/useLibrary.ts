import { useCallback, useEffect, useRef, useState } from "react";
import { bridge, InitData, ListItem } from "@/lib/bridge";

export type SortMode = "最新" | "最旧" | "名称" | "最近使用" | "最常用" | "自定义";

export interface LibraryState {
  init: InitData | null;
  group: string; // "全部" 或组名
  category: string;
  kind: string; // 全部/Shape/Style
  filter: string;
  sort: SortMode;
  unusedOnly: boolean;
  colorFilter: string;
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
  const [sort, setSort] = useState<SortMode>("最新");
  const [unusedOnly, setUnusedOnly] = useState(false);   // 仅未使用
  const [colorFilter, setColorFilter] = useState("");     // 主色近似筛选 "RRGGBB"
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
        .list(filter, category, kind, group, sort, offset, pageSizeRef.current, unusedOnly, colorFilter)
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
    [filter, category, kind, group, sort, unusedOnly, colorFilter]
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
      // 页面重载/窗口迁移重建后，恢复上次浏览的 组/分类/类型/搜索/排序（避免跳回"全部"）
      const lg = ((init as any)?.lastGroup as string | undefined) || "";
      const lc = ((init as any)?.lastCategory as string | undefined) || "";
      const lk = ((init as any)?.lastKind as string | undefined) || "";
      const lf = ((init as any)?.lastFilter as string | undefined) || "";
      const ls = ((init as any)?.lastSort as string | undefined) || "";
      let changed = false;
      if (lg && lg !== group) { setGroup(lg); changed = true; }
      if (lc && lc !== "全部" && lc !== category) { setCategory(lc); changed = true; }
      if (lk && lk !== "全部" && lk !== kind) { setKind(lk); changed = true; }
      if (lf && lf !== filter) { setFilter(lf); changed = true; }
      if (ls && ls !== sort) { setSort(ls as SortMode); changed = true; }
      if (changed) {
        // 状态变化后由"防抖搜索 effect"(依赖这些字段) 自动 reload
        return;
      }
      reload(true);
    }
  }, [init, reload]); // eslint-disable-line react-hooks/exhaustive-deps

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
  }, [filter, category, kind, group, sort, unusedOnly, colorFilter, reload]);

  // 下行 refresh（C# 数据变化：收藏/删除/移动等）：刷新元数据 + 重载当前列表。
  // 关键：监听只注册一次并用 reloadRef 取"最新"的 reload——否则每次筛选变化都会累积一个
  // 监听器，旧监听器闭包里的陈旧筛选(如初次 group="全部")会把列表打回全部（历史 bug）。
  const refreshInFlight = useRef(false);
  const reloadRef = useRef(reload);
  useEffect(() => {
    reloadRef.current = reload;
  }, [reload]);
  useEffect(() => {
    const fn = () => {
      if (refreshInFlight.current) return;   // 防抖合并连续下行
      refreshInFlight.current = true;
      bridge
        .init()
        .then((d) => {
          setInit(d);
          // 用最新筛选重载当前视图
          if (didInitialLoad.current) reloadRef.current(true);
        })
        .catch(() => {})
        .finally(() => {
          window.setTimeout(() => {
            refreshInFlight.current = false;
          }, 400);
        });
    };
    const offRefresh = bridge.on("refresh", fn);
    const offSettings = bridge.on("settingsChanged", (s) => setInit((prev) => (prev ? { ...prev, ...s } : prev)));
    return () => {
      offRefresh();
      offSettings();
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
    sort,
    unusedOnly,
    colorFilter,
    setUnusedOnly,
    setColorFilter,
    setGroup: (g: string) => {
      setGroup(g);
      setCategory("全部");
    },
    setCategory,
    setKind,
    setFilter,
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
