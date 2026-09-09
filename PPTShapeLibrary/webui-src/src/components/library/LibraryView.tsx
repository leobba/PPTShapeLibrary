import { useEffect, useMemo, useRef, useState } from "react";
import { bridge, ListItem } from "@/lib/bridge";
import { useLibrary } from "@/hooks/useLibrary";
import { dialogs } from "@/lib/dialogs";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { Skeleton } from "@/components/ui/skeleton";
import { ScrollArea } from "@/components/ui/scroll-area";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuLabel,
  ContextMenuSeparator,
  ContextMenuTrigger,
} from "@/components/ui/context-menu";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Folder,
  Search,
  MoreVertical,
  Plus,
  Upload,
  Trash2,
  RotateCcw,
  RotateCw,
  Star,
  Download,
  RefreshCw,
  X,
  ScanSearch,
  ArrowLeft,
  ArrowRight,
  ArrowUp,
  ArrowDown,
  EyeOff,
  Eye,
  LayoutTemplate,
  Type,
  Palette,
  ChartPie,
  FolderPlus,
  Package,
  LayoutGrid,
  PanelRight,
  PanelLeft,
  ChevronDown,
} from "lucide-react";

const GROUP_ICONS: Record<string, React.ReactNode> = {
  我的: <Star className="size-4" />,
  模板: <LayoutTemplate className="size-4" />,
  素材: <Folder className="size-4" />,
  文字: <Type className="size-4" />,
  背景: <Palette className="size-4" />,
  图例: <ChartPie className="size-4" />,
};

function tooltip(e: React.MouseEvent) {
  /* noop */
}

export default function LibraryView() {
  const lib = useLibrary();
  const { init, items, selected, loadMore } = lib;
  const selCount = selected.size;
  const [preview, setPreview] = useState<ListItem | null>(null);
  const [hoverImg, setHoverImg] = useState<{ url: string; kind: string } | null>(null);
  const [showTrash, setShowTrash] = useState(false);
  const [showGroups, setShowGroups] = useState(false);
  const [showCats, setShowCats] = useState(false);   // 管理二级分类（按当前一级组）
  const [showDupes, setShowDupes] = useState(false);
  const [trashConfirmIds, setTrashConfirmIds] = useState<string[] | null>(null);
  const [toastMsg, setToastMsg] = useState<string | null>(null);
  const toastTimer = useRef<number | undefined>(undefined);
  const showToast = (msg: string) => {
    setToastMsg(msg);
    if (toastTimer.current) window.clearTimeout(toastTimer.current);
    toastTimer.current = window.setTimeout(() => setToastMsg(null), 3000);
  };
  // 插入后调整条（旋转/透明度/微调；作用于刚插入并已选中的形状）
  const [showAdjust, setShowAdjust] = useState(false);
  const adjustTimer = useRef<number | undefined>(undefined);
  const showAdjustBar = () => {
    setShowAdjust(true);
    if (adjustTimer.current) window.clearTimeout(adjustTimer.current);
    adjustTimer.current = window.setTimeout(() => setShowAdjust(false), 10000);
  };
  const adjust = (action: string) => {
    bridge
      .adjust(action)
      .then(() => {})
      .catch((e: Error) => showToast(e.message || "调整失败（请先选中形状）"));
  };
  // 插入（防连点：一次只允许一个插入在途；超时给出友好提示而非裸 timeout）
  const insertingRef = useRef(false);
  const insertItem = (id: string) => {
    if (insertingRef.current) {
      showToast("正在插入，请稍候…");
      return;
    }
    insertingRef.current = true;
    bridge
      .insert(id)
      .then(() => {
        insertingRef.current = false;
        showToast("✓ 已插入到当前页");
        showAdjustBar();
        lib.reload(true);
      })
      .catch((e: Error) => {
        insertingRef.current = false;
        const msg = /timeout/i.test(e.message)
          ? "插入处理超时：若已出现在页面请勿重复点击；稍后可重试"
          : e.message || "插入失败";
        showToast(msg);
      });
  };
  const applyStyleWithFeedback = (id: string, all?: boolean) => {
    const p = all ? bridge.applyStyleSameTypeAllSlides(id) : bridge.applyStyle(id);
    p.then(() => {
      showToast("样式已应用");
      lib.refreshMeta();
      lib.reload(true); // TouchUsed 更新了使用统计/最近使用排序，重载列表
    }).catch((e: Error) => showToast(e.message || "应用失败（请先选中形状）"));
  };
  // 面板宽度自适应：容器宽度 → 卡片列最小宽（窄面板小卡，宽面板大卡）
  const [gridMin, setGridMin] = useState(96);
  // 宽面板固定每行两个素材（>=760px 时 2 列），窄面板回退自适应列
  const [gridCols, setGridCols] = useState<string | null>(null);
  const gridRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    const el = gridRef.current;
    if (!el) return;
    if (typeof ResizeObserver === "undefined") return;
    // 卡片密度系数：compact 小卡 / standard 标准 / large 大图
    const k =
      (init as any)?.cardSize === "compact" ? 0.85 : (init as any)?.cardSize === "large" ? 1.3 : 1;
    const ro = new ResizeObserver(() => {
      const w = el.clientWidth;
      if (w >= 690) setGridCols("repeat(2, 1fr)");
      else setGridCols(null);
      const base = w >= 620 ? 148 : w >= 420 ? 120 : w >= 300 ? 98 : 84;
      setGridMin(Math.round(base * k));
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, [init]);
  // 网格模板：宽面板固定 2 列，否则 auto-fill 自适应
  const gridTemplate = gridCols || `repeat(auto-fill, minmax(${gridMin}px, 1fr))`;

  const groups = useMemo(() => ["全部", ...(init?.groups || [])], [init]);

  // 列表内容变化（插入/删除/重载后卡片可能已不存在）时清除 hover 大图，防遮罩残留
  useEffect(() => {
    setHoverImg(null);
  }, [items]);
  const cats = useMemo(() => {
    if (lib.group === "全部") {
      const set = new Set<string>();
      Object.values(init?.groupCategories || {}).forEach((cs) => cs.forEach((c) => set.add(c)));
      return ["全部", ...Array.from(set).sort()];
    }
    return ["全部", ...(init?.groupCategories?.[lib.group] || [])];
  }, [init, lib.group]);

  return (
    <div className="flex h-screen flex-col bg-background text-foreground">
      {/* 顶栏（无标题块：直接搜索 + 管理） */}
      <header className="flex items-center gap-2 border-b bg-card/70 px-3 py-2 backdrop-blur">
        <div className="ml-0 flex flex-1 items-center gap-1 rounded-md bg-muted px-2">
          <Search className="size-3.5 text-muted-foreground" />
          <Input
            className="h-7 border-0 bg-transparent focus-visible:ring-0"
            placeholder="搜索 名称 / 分类 / 标签"
            value={lib.filter}
            onChange={(e) => lib.setFilter(e.target.value)}
            onFocus={(e) => e.target.select()}
            onKeyDown={(e) => {
              if (e.key === "Escape") lib.setFilter("");
            }}
          />
          {lib.filter && (
            <button
              className="flex size-4 shrink-0 items-center justify-center rounded-full text-muted-foreground hover:bg-muted hover:text-foreground"
              title="清空搜索 (Esc)"
              onClick={() => lib.setFilter("")}
            >
              <X className="size-3" />
            </button>
          )}
        </div>
        <Select value={lib.sort} onValueChange={(v) => lib.setSort(v as any)}>
          <SelectTrigger className="h-7 w-28 text-xs">
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            {["最新", "最旧", "名称", "最近使用", "最常用", "自定义"].map((s) => (
              <SelectItem key={s} value={s}>{s}</SelectItem>
            ))}
          </SelectContent>
        </Select>

        {/* 库管理 */}
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button variant="ghost" size="icon" className="h-7 w-7">
              <MoreVertical className="size-4" />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end">
            <DropdownMenuItem onSelect={() => setShowGroups(true)}>
              <Folder className="size-4" /> 管理分组
            </DropdownMenuItem>
            <DropdownMenuItem
              onSelect={() => {
                if (lib.group === "全部") {
                  dialogs.alert("请先在左侧一级栏选中要管理分类的分组（如：我的 / 模板 / 素材…）。");
                  return;
                }
                setShowCats(true);
              }}
            >
              <LayoutGrid className="size-4" /> 管理分类…
            </DropdownMenuItem>
            <DropdownMenuItem onSelect={() => setShowDupes(true)}>
              <ScanSearch className="size-4" /> 去重清理
            </DropdownMenuItem>
            <DropdownMenuItem
              onSelect={async () => {
                const ok = await dialogs.confirm(
                  "重建全部形状条目的预览缩略图？\n（修复旧收藏中预览与内容错位的问题，按新方式从收藏文件重新生成）",
                  { title: "重建预览图", okText: "重建", danger: false }
                );
                if (!ok) return;
                try {
                  const d = await bridge.rebuildPreviews();
                  dialogs.alert("已重建 " + (d.rebuilt || 0) + " 条预览缩略图。");
                  lib.refreshMeta();
                  lib.reload(true);
                } catch (e: any) {
                  dialogs.alert("重建失败：" + (e?.message || e));
                }
              }}
            >
              <RefreshCw className="size-4" /> 重建预览图…
            </DropdownMenuItem>
            <DropdownMenuItem onSelect={() => setShowTrash(true)}>
              <Trash2 className="size-4" /> 回收站
            </DropdownMenuItem>
            <DropdownMenuItem
              onSelect={() => {
                bridge.importPictures().then((d) => {
                  lib.refreshMeta();
                  dialogs.alert("已导入 " + (d.imported || 0) + " 张图片到 素材 › 图片");
                });
              }}
            >
              <Upload className="size-4" /> 导入图片…
            </DropdownMenuItem>
            <DropdownMenuItem
              onSelect={async () => {
                // 直接弹文件选择器选 .ppack.zip（点对点分享一键导入，可多选）
                try {
                  const d = await bridge.importPackFile();
                  if (d.imported > 0) {
                    lib.refreshMeta();
                    lib.reload(true);
                    let msg = "已导入 " + (d.imported || 0) + " 个素材";
                    if (d.conflicts > 0) msg += "（" + d.conflicts + " 个 id 冲突已自动换新）";
                    if (d.packName) msg += "\n素材包：" + d.packName;
                    if (d.author) msg += "\n作者：" + d.author;
                    if (d.license) msg += "  · 许可：" + d.license;
                    if (d.sourceUrl) msg += "\n来源：" + d.sourceUrl;
                    dialogs.alert(msg);
                  }
                  // imported=0：取消选择则静默；选了但没导入成功则提示
                  if (d.imported === 0 && d.hadSelection) {
                    dialogs.alert("所选文件未能导入：请确认是有效的素材包（.ppack.zip）。");
                  }
                } catch (e: any) {
                  dialogs.alert("导入失败：" + (e?.message || e));
                }
              }}
            >
              <Download className="size-4" /> 导入素材包…
            </DropdownMenuItem>
            <DropdownMenuSeparator />
            <DropdownMenuLabel className="text-[11px]">窗格停靠</DropdownMenuLabel>
            <DropdownMenuItem
              onSelect={() => bridge.setPaneDock("left").then(() => lib.refreshMeta())}
              className={(((init as any)?.paneDock as string) || "left") === "left" ? "bg-accent text-accent-foreground" : ""}
            >
              <PanelLeft className="size-4" /> 停靠左侧
            </DropdownMenuItem>
            <DropdownMenuItem
              onSelect={() => bridge.setPaneDock("right").then(() => lib.refreshMeta())}
              className={(((init as any)?.paneDock as string) || "left") === "right" ? "bg-accent text-accent-foreground" : ""}
            >
              <PanelRight className="size-4" /> 停靠右侧
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      </header>

      <div className="flex min-h-0 flex-1">
        {/* 左侧分组栏（第一层）：我的/模板/自定义组 下挂 全部/形状/样式；素材等内置组直接一级 */}
        <nav className="flex w-[78px] shrink-0 flex-col items-stretch gap-1 overflow-y-auto border-r bg-card/40 px-1.5 py-2">
          {groups.map((g) => {
            const count = init?.groupCounts?.[g];
            // 有类型子项：我的/模板 及用户自定义组（预设内置组素材/文字/背景/图例 与虚拟"全部" 无类型细分，仅组）
            const presetPlain = ["素材", "文字", "背景", "图例"];
            const typed = g !== "全部" && !presetPlain.includes(g);
            const inThisGroup = lib.group === g;
            return (
              <div key={g} className="flex flex-col items-stretch gap-0.5">
                <button
                  title={g}
                  onClick={() => {
                    lib.setGroup(g);
                    lib.setKind("全部"); // 点组头=该组全部类型（无类型组恒 全部）
                  }}
                  className={
                    "flex items-center gap-1 rounded-lg px-1.5 py-1.5 text-[11px] transition-colors " +
                    (inThisGroup && lib.kind === "全部"
                      ? "bg-primary/15 font-medium text-primary"
                      : "text-muted-foreground hover:bg-muted hover:text-foreground")
                  }
                >
                  <span className="relative shrink-0">
                    {GROUP_ICONS[g] || <Folder className="size-3.5" />}
                    {!!count && (
                      <span className="absolute -right-2 -top-1.5 rounded-full bg-muted px-0.5 text-[8px] leading-3">
                        {count > 99 ? "99+" : count}
                      </span>
                    )}
                  </span>
                  <span className="max-w-full truncate">{g}</span>
                </button>
                {typed && (
                  <div className="flex flex-col gap-0.5 pl-4">
                    {(
                      [
                        ["全部", "全部"],
                        ["形状", "Shape"],
                        ["样式", "Style"],
                      ] as const
                    ).map(([label, k]) => (
                      <button
                        key={k}
                        onClick={() => {
                          lib.setGroup(g);
                          lib.setKind(k);
                        }}
                        className={
                          "self-start rounded px-1.5 py-0.5 text-[10px] leading-none transition-colors " +
                          (inThisGroup && lib.kind === k
                            ? "bg-primary text-primary-foreground"
                            : "text-muted-foreground hover:bg-muted hover:text-foreground")
                        }
                      >
                        {label}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
          <Separator className="my-1" />
          <button
            className="flex flex-col items-center gap-1 rounded-lg border border-dashed px-0.5 py-2 text-[10px] text-muted-foreground hover:border-primary hover:text-primary"
            onClick={async () => {
              const name = await dialogs.prompt("新建一级分组名称：", {
                placeholder: "如：课程作业",
                okText: "创建",
              });
              if (name?.trim()) bridge.createGroup(name.trim()).then(() => lib.refreshMeta());
            }}
          >
            <Plus className="size-3.5" /> 新建
          </button>
        </nav>

        {/* 主区 */}
        <div className="flex min-w-0 flex-1 flex-col">
          {/* 二级分类行：全部 | 当前分类下拉(列出全部分类) | 新建分类 */}
          <div className="flex items-center gap-1.5 border-b bg-card/40 px-2 py-1.5">
            <button
              onClick={() => lib.setCategory("全部")}
              className={
                "shrink-0 rounded-full px-2.5 py-1 text-xs transition-colors " +
                (lib.category === "全部"
                  ? "bg-primary text-primary-foreground"
                  : "text-muted-foreground hover:bg-muted hover:text-foreground")
              }
            >
              全部
            </button>
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <button
                  className={
                    "flex shrink-0 items-center gap-1 rounded-full border px-2.5 py-1 text-xs transition-colors " +
                    (lib.category !== "全部"
                      ? "border-primary bg-primary/10 text-primary"
                      : "border-border text-muted-foreground hover:bg-muted hover:text-foreground")
                  }
                  title="选择二级分类"
                >
                  {lib.category !== "全部" ? lib.category : "选择分类"}
                  <ChevronDown className="size-3 opacity-60" />
                </button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="start" className="max-h-72 overflow-y-auto">
                {(cats as string[])
                  .filter((c) => c !== "全部")
                  .map((c) => (
                    <DropdownMenuItem key={c} onSelect={() => lib.setCategory(c)}>
                      <span className={"mr-1 size-2 rounded-full " + (lib.category === c ? "bg-primary" : "bg-muted")} />
                      {c}
                    </DropdownMenuItem>
                  ))}
                {(cats as string[]).filter((c) => c !== "全部").length === 0 && (
                  <p className="px-2 py-1 text-[11px] text-muted-foreground">该组暂无二级分类</p>
                )}
              </DropdownMenuContent>
            </DropdownMenu>
            {lib.group !== "全部" && (
              <button
                className="flex shrink-0 items-center gap-0.5 rounded-full border border-dashed px-2 py-1 text-xs text-muted-foreground hover:border-primary hover:text-primary"
                title="新建二级分类"
                onClick={async () => {
                  const name = await dialogs.prompt("新建二级分类（当前组：「" + lib.group + "」）", {
                    placeholder: "分类名，如：箭头",
                    okText: "创建",
                  });
                  if (name?.trim()) {
                    bridge
                      .createGroupCategory(lib.group, name.trim())
                      .then(() => {
                        lib.refreshMeta();
                        lib.setCategory(name.trim()); // 选中新分类
                      })
                      .catch((e: any) => dialogs.alert("创建失败：" + (e?.message || e)));
                  }
                }}
              >
                <Plus className="size-3" /> 新建分类
              </button>
            )}
          </div>

          {/* 状态行 */}
          <div className="flex items-center justify-between px-3 py-1 text-[11px] text-muted-foreground">
            <span className="flex items-center gap-1">
              {lib.group !== "全部" && <span className="font-medium text-foreground">{lib.group}</span>}
              {lib.group !== "全部" && lib.category !== "全部" && <span>›</span>}
              {lib.category !== "全部" && <span className="text-foreground">{lib.category}</span>}
              {lib.loading ? (
                <span className="italic">加载中…</span>
              ) : lib.category !== "全部" || lib.kind !== "全部" || lib.filter  ? (
                <span>
                  <span className="font-semibold text-foreground">{items.length}</span>
                  {" / "}
                  <span className="font-semibold text-foreground">{lib.filteredTotal}</span>
                  {" 条匹配"}
                </span>
              ) : (
                <span>
                  共 <span className="font-semibold text-foreground">{lib.filteredTotal}</span> 项
                </span>
              )}
            </span>
            {/* 撤销/重做 */}
            <span className="flex items-center gap-1">
              <Button variant="ghost" size="icon" className="h-6 w-6" onClick={() => bridge.undo()}>
                <RotateCcw className="size-3.5" />
              </Button>
              <Button variant="ghost" size="icon" className="h-6 w-6" onClick={() => bridge.redo()}>
                <RotateCw className="size-3.5" />
              </Button>
            </span>
          </div>

          {/* 最近使用横条（仅"全部"且无任何筛选时显示全局最近使用 Top6） */}
          {init?.recentTop &&
            (init.recentTop as any[]).length > 0 &&
            lib.group === "全部" &&
            lib.category === "全部" &&
            lib.kind === "全部" &&
            !lib.filter && (
              <div className="flex items-center gap-2 overflow-x-auto border-b bg-card/30 px-3 py-2">
                <span className="shrink-0 text-[11px] font-medium text-muted-foreground">最近使用</span>
                <div className="flex shrink-0 gap-2">
                  {(init.recentTop as any[]).map((r: any) => (
                    <button
                      key={r.id}
                      className="group w-16 shrink-0 text-left"
                      title={r.name}
                      onClick={() => {
                        if (r.kind === "Style") applyStyleWithFeedback(r.id);
                        else insertItem(r.id);
                      }}
                    >
                      <div className="aspect-square w-full overflow-hidden rounded-md border bg-muted/40">
                        <img
                          src={r.previewUrl}
                          alt={r.name}
                          loading="lazy"
                          className="h-full w-full object-cover transition-transform group-hover:scale-105"
                        />
                      </div>
                      <p className="truncate text-[9px] text-muted-foreground">{r.name}</p>
                    </button>
                  ))}
                </div>
              </div>
            )}

          {/* 卡片网格（原生滚动：移出区清除 hover，滚底加载更多，宽度自适应列宽） */}
          <div
            ref={gridRef}
            className="min-h-0 flex-1 overflow-y-auto"
            onMouseLeave={() => setHoverImg(null)}
            onScroll={(e) => {
              setHoverImg(null);
              const el = e.currentTarget;
              if (el.scrollTop + el.clientHeight >= el.scrollHeight - 240) {
                loadMore();
              }
            }}
          >
              {!init && (
                <div className="grid gap-2 p-2.5" style={{ gridTemplateColumns: gridTemplate }}>
                  {Array.from({ length: 8 }).map((_, i) => (
                    <div key={i} className="flex flex-col gap-1.5">
                      <Skeleton className="aspect-[1/.92] w-full rounded-lg" />
                      <Skeleton className="h-3 w-3/4" />
                    </div>
                  ))}
                </div>
              )}
              {init && !lib.loading && items.length === 0 ? (
                <div className="flex h-full flex-col items-center justify-center gap-2 p-8 text-center">
                  <Folder className="size-9 text-muted-foreground/40" />
                  <p className="text-sm font-medium text-muted-foreground">
                    {lib.group !== "全部" || lib.category !== "全部" || lib.kind !== "全部" || lib.filter 
                      ? "没有符合筛选条件的素材"
                      : "素材库还是空的"}
                  </p>
                  <p className="max-w-xs text-xs text-muted-foreground/70">
                    {lib.group !== "全部" || lib.category !== "全部" || lib.kind !== "全部" || lib.filter 
                      ? "试试清空搜索词、切换一级分组 / 二级分类 / 类型 / 标签"
                      : "在演示文稿中选中形状，点击「素材库」选项卡里的「收藏形状」/「收藏整页」/「收藏样式」等收录素材"}
                  </p>
                  {lib.group === "全部" &&
                    lib.category === "全部" &&
                    lib.kind === "全部" &&
                    !lib.filter && (
                      <div className="mt-2 flex gap-2">
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={async () => {
                            try {
                              const d = await bridge.importPackFile();
                              if (d.imported > 0) {
                                lib.refreshMeta();
                                lib.reload(true);
                                dialogs.alert("已导入 " + (d.imported || 0) + " 个素材。");
                              }
                            } catch (e: any) {
                              dialogs.alert("导入失败：" + (e?.message || e));
                            }
                          }}
                        >
                          <Download className="mr-1 size-3.5" /> 导入素材包
                        </Button>
                        <Button
                          variant="ghost"
                          size="sm"
                          onClick={() => {
                            bridge.importPictures().then((d) => {
                              lib.refreshMeta();
                              lib.reload(true);
                              if (d.imported > 0) dialogs.alert("已导入 " + d.imported + " 张图片到 素材 › 图片");
                            });
                          }}
                        >
                          <Upload className="mr-1 size-3.5" /> 导入图片
                        </Button>
                      </div>
                    )}
                </div>
              ) : (
                <div
                  className="grid gap-2 p-2.5"
                  style={{ gridTemplateColumns: gridTemplate }}
                >
                  {items.map((item) => {
                    // 最近使用横条显示时，网格内不重复展示同 id 卡片
                    const inRecent =
                      lib.group === "全部" &&
                      lib.category === "全部" &&
                      lib.kind === "全部" &&
                      !lib.filter &&
                      (init?.recentTop as any[])?.some((r: any) => r.id === item.id);
                    if (inRecent) return null;
                    return (
                    <ItemCard
                      key={item.id}
                      item={item}
                      selected={selected.has(item.id)}
                      dragEnabled={lib.sort === "自定义"}
                      clickToInsert={(init as any)?.clickToInsert !== false}
                      onDropCard={(fromId, toId) => {
                        // 自定义排序：把 from 移到 to 的位置并持久化全量顺序
                        const ids = items.map((x) => x.id);
                        const fi = ids.indexOf(fromId);
                        const ti = ids.indexOf(toId);
                        if (fi < 0 || ti < 0) return;
                        ids.splice(fi, 1);
                        ids.splice(ti, 0, fromId);
                        bridge
                          .reorder(ids)
                          .then(() => {
                            showToast("已保存排序");
                            lib.reload(true);
                          })
                          .catch((e: Error) => showToast(e.message || "排序保存失败"));
                      }}
                      onToggle={() => lib.toggleSelect(item.id)}
                      onPreview={() => setPreview(item)}
                      onHover={setHoverImg}
                      onInsert={() => insertItem(item.id)}
                      onApplyStyle={() => applyStyleWithFeedback(item.id)}
                      onChanged={() => {
                        lib.refreshMeta(); // 刷新侧栏组计数/元数据
                        lib.reload(true); // 重载当前列表（删除/重命名/移动后即时生效）
                      }}
                    />
                    );
                  })}
                  {lib.loading &&
                    items.length === 0 &&
                    Array.from({ length: 8 }).map((_, i) => (
                      <div key={i} className="flex flex-col gap-1.5">
                        <Skeleton className="aspect-[1/.92] w-full rounded-lg" />
                        <Skeleton className="h-3 w-3/4" />
                      </div>
                    ))}
                </div>
              )}
          </div>

          {/* hover 大图预览（图片查看器式，纯展示不拦截事件） */}
          {hoverImg && <HoverLens data={hoverImg} />}
        </div>
      </div>

      {/* 插入后调整条（作用于当前选中形状：旋转/透明度/微调；10 秒自动隐藏） */}
      {showAdjust && selCount === 0 && (
        <div className="fixed bottom-4 left-1/2 z-40 flex -translate-x-1/2 items-center gap-1 rounded-xl border bg-popover px-2 py-1.5 shadow-xl">
          <span className="mr-1 px-1 text-[10px] text-muted-foreground">微调</span>
          {[
            [<RotateCcw key="rl" className="size-3.5" />, "rotateL", "左转 15°"],
            [<RotateCw key="rr" className="size-3.5" />, "rotateR", "右转 15°"],
            [<EyeOff key="od" className="size-3.5" />, "opacityDown", "透明度 -"],
            [<Eye key="ou" className="size-3.5" />, "opacityUp", "透明度 +"],
            [<ArrowLeft key="nl" className="size-3.5" />, "nudgeL", "左移"],
            [<ArrowRight key="nr" className="size-3.5" />, "nudgeR", "右移"],
            [<ArrowUp key="nu" className="size-3.5" />, "nudgeU", "上移"],
            [<ArrowDown key="nd" className="size-3.5" />, "nudgeD", "下移"],
          ].map(([icon, a, tip]) => (
            <Button
              key={a as string}
              variant="ghost"
              size="icon-sm"
              className="h-7 w-7"
              title={tip as string}
              onClick={() => adjust(a as string)}
            >
              {icon as React.ReactNode}
            </Button>
          ))}
          <Button
            variant="ghost"
            size="sm"
            className="h-7 px-2 text-xs text-muted-foreground"
            title="撤销上一步"
            onClick={() => bridge.undo().catch(() => {})}
          >
            <RotateCcw className="mr-1 size-3" />
            撤销
          </Button>
          <Button
            variant="ghost"
            size="sm"
            className="h-7 px-2 text-xs text-muted-foreground"
            title="重做"
            onClick={() => bridge.redo().catch(() => {})}
          >
            <RotateCw className="mr-1 size-3" />
            重做
          </Button>
        </div>
      )}

      {/* 批量操作浮条 */}
      {selCount > 0 && (
        <div className="fixed bottom-4 left-1/2 z-40 flex max-w-[calc(100vw-16px)] -translate-x-1/2 items-center gap-1 overflow-x-auto rounded-xl border bg-popover px-2.5 py-2 shadow-xl">
          <span className="shrink-0 px-1 text-xs font-medium">已选 {selCount} 项</span>
          <Button
            variant="ghost"
            size="sm"
            className="h-7 shrink-0 px-2 text-xs"
            title="取消选择"
            onClick={() => lib.clearSelection()}
          >
            <X className="mr-1 size-3.5" />
            取消
          </Button>
          <Button
            variant="outline"
            size="sm"
            className="h-7 shrink-0 px-2 text-xs"
            onClick={async () => {
              const c = await dialogs.prompt("输入二级分类（已有或新分类）：", {
                initial: lib.category !== "全部" ? lib.category : "",
                placeholder: "如：箭头",
                okText: "移动",
              });
              if (c?.trim()) {
                const ids = Array.from(selected);
                bridge
                  .batchMove(ids, c.trim())
                  .then(() => {
                    lib.clearSelection();
                    lib.reload(true);
                    showToast("已移动 " + ids.length + " 项到「" + c.trim() + "」");
                  })
                  .catch((e: Error) => showToast(e.message || "移动失败"));
              }
            }}
          >
            <Folder className="mr-1 size-3.5" />
            移动分类
          </Button>
          <Button
            variant="outline"
            size="sm"
            className="h-7 shrink-0 px-2 text-xs"
            onClick={async () => {
              const g = await dialogs.prompt("输入一级分组（已有或新分组）：", {
                initial: lib.group !== "全部" ? lib.group : "",
                placeholder: "如：课程作业",
                okText: "移动",
              });
              if (g?.trim()) {
                const ids = Array.from(selected);
                bridge
                  .createGroup(g.trim())
                  .then(() => bridge.batchMoveGroup(ids, g.trim()))
                  .then(() => {
                    lib.clearSelection();
                    lib.reload(true);
                    lib.refreshMeta();
                    showToast("已移动 " + ids.length + " 项到「" + g.trim() + "」");
                  })
                  .catch((e: Error) => showToast(e.message || "移动失败"));
              }
            }}
          >
            <FolderPlus className="mr-1 size-3.5" />
            移动分组
          </Button>
          <Button
            variant="outline"
            size="sm"
            className="h-7 shrink-0 px-2 text-xs"
            onClick={() => {
              const ids = Array.from(selected);
              bridge
                .exportPack(ids, "")
                .then((d) => {
                  lib.clearSelection();
                  showToast("已导出素材包" + (d?.path ? "：" + (d.path as string).split(/[\\/]/).pop() : ""));
                })
                .catch((e: Error) => showToast(e.message || "导出失败"));
            }}
          >
            <Package className="mr-1 size-3.5" />
            导出素材包
          </Button>
          <Button
            variant="destructive"
            size="sm"
            className="h-7 shrink-0 px-2 text-xs"
            onClick={() => setTrashConfirmIds(Array.from(selected))}
          >
            <Trash2 className="mr-1 size-3.5" />
            移入回收站
          </Button>
        </div>
      )}

      {/* 大图 Dialog（lightbox） */}
      <Dialog open={!!preview} onOpenChange={(o) => !o && setPreview(null)}>
        {preview && (
          <DialogContent className="max-w-lg">
            <DialogHeader>
              <DialogTitle className="pr-8">
                <span className="block truncate" title={preview.name}>{preview.name}</span>
              </DialogTitle>
              <DialogDescription>
                {preview.group || "我的"} › {preview.category} ·{" "}
                {preview.kind === "Style" ? "样式" : "形状"}
              </DialogDescription>
            </DialogHeader>
            <div className="checker flex items-center justify-center rounded-lg p-2">
              <img
                src={preview.previewUrl}
                alt={preview.name}
                className="max-h-[55vh] w-auto rounded object-contain"
              />
            </div>
            <div className="flex justify-end gap-2">
              {preview.kind !== "Style" ? (
                <Button
                  onClick={() => {
                    setPreview(null);
                    insertItem(preview.id);
                  }}
                >
                  插入到当前页
                </Button>
              ) : (
                <Button
                  onClick={() => {
                    setPreview(null);
                    applyStyleWithFeedback(preview.id);
                  }}
                >
                  应用到选中形状
                </Button>
              )}
              <Button variant="outline" onClick={() => setPreview(null)}>
                关闭
              </Button>
            </div>
          </DialogContent>
        )}
      </Dialog>

      {showTrash && (
        <TrashDialog
          onClose={() => {
            setShowTrash(false);
            lib.refreshMeta();
            lib.reload(true); // 回收站恢复/删除后刷新列表
          }}
        />
      )}
      {showGroups && (
        <GroupDialog
          groups={init?.groups || []}
          onClose={() => {
            setShowGroups(false);
            lib.refreshMeta();
            lib.reload(true); // 分组改名/删除后刷新
          }}
        />
      )}
      {showDupes && (
        <DuplicatesDialog
          onClose={() => {
            setShowDupes(false);
            lib.refreshMeta();
            lib.reload(true);
          }}
        />
      )}
      {showCats && (
        <CategoriesDialog
          group={lib.group}
          onClose={() => {
            setShowCats(false);
            lib.refreshMeta();
            lib.reload(true);
          }}
        />
      )}
      <AlertDialog
        open={trashConfirmIds !== null}
        onOpenChange={(o) => !o && setTrashConfirmIds(null)}
      >
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>移入回收站？</AlertDialogTitle>
            <AlertDialogDescription>
              将选中的 {trashConfirmIds?.length || 0} 项移入回收站。可稍后在「回收站」恢复。
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>取消</AlertDialogCancel>
            <AlertDialogAction
              onClick={() => {
                if (trashConfirmIds)
                  bridge.del(trashConfirmIds).then(() => {
                    lib.clearSelection();
                    lib.reload(true);
                  });
              }}
            >
              移入回收站
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
      {toastMsg && (
        <div className="pointer-events-none fixed bottom-6 left-1/2 z-[200] -translate-x-1/2 rounded-lg bg-foreground px-3.5 py-2 text-sm text-background shadow-lg">
          {toastMsg}
        </div>
      )}
    </div>
  );
}

/** 卡片（含右键菜单） */
function ItemCard({
  item,
  selected,
  onToggle,
  onPreview,
  onHover,
  onInsert,
  onApplyStyle,
  onChanged,
  dragEnabled,
  onDropCard,
  clickToInsert,
}: {
  item: ListItem;
  selected: boolean;
  onToggle: () => void;
  onPreview: () => void;
  onHover: (v: { url: string; kind: string } | null) => void;
  onInsert: () => void;
  onApplyStyle: () => void;
  onChanged: () => void;
  dragEnabled?: boolean;
  onDropCard?: (fromId: string, toId: string) => void;
  /** 设置：true=单击即插；false=单击预览、双击插入 */
  clickToInsert?: boolean;
}) {
  const isStyle = item.kind === "Style";
  const clickTimerRef = useRef<number | undefined>(undefined);
  const [dragState, setDragState] = useState<"dragging" | "over" | null>(null);
  const handleClick = (e: React.MouseEvent) => {
    if (e.ctrlKey) {
      onToggle();
      return;
    }
    if (isStyle) {
      onApplyStyle();
      return;
    }
    if (clickToInsert !== false) {
      // 单击即插入
      onInsert();
      return;
    }
    // 单击预览 / 双击插入：延迟判定双击
    if (clickTimerRef.current) {
      window.clearTimeout(clickTimerRef.current);
      clickTimerRef.current = undefined;
      onInsert(); // 第二次点击构成双击 → 插入
      return;
    }
    clickTimerRef.current = window.setTimeout(() => {
      clickTimerRef.current = undefined;
      onPreview(); // 无双击 → 预览
    }, 260);
  };
  return (
    <ContextMenu>
      <ContextMenuTrigger>
        <div
          className={
            "group relative cursor-pointer overflow-hidden rounded-lg border bg-card shadow-sm transition-all hover:-translate-y-0.5 hover:shadow-md " +
            (selected ? "border-primary ring-2 ring-primary/60" : "border-border") +
            (dragState === "dragging" ? " opacity-40" : "") +
            (dragState === "over" ? " border-primary ring-2 ring-primary/50" : "")
          }
          draggable={!!dragEnabled}
          onDragStart={(e) => {
            setDragState("dragging");
            e.dataTransfer.setData("text/plain", item.id);
            e.dataTransfer.effectAllowed = "move";
          }}
          onDragEnd={() => setDragState(null)}
          onDragOver={(e) => {
            if (dragEnabled) {
              e.preventDefault();
              e.dataTransfer.dropEffect = "move";
              if (dragState !== "over") setDragState("over");
            }
          }}
          onDragLeave={() => {
            if (dragState === "over") setDragState(null);
          }}
          onDrop={(e) => {
            setDragState(null);
            if (dragEnabled && onDropCard) {
              e.preventDefault();
              const fromId = e.dataTransfer.getData("text/plain");
              if (fromId && fromId !== item.id) onDropCard(fromId, item.id);
            }
          }}
          onClick={handleClick}
          onMouseEnter={() => onHover({ url: item.previewUrl, kind: item.kind })}
          onMouseLeave={() => onHover(null)}
          title={item.name}
        >
          <div className="aspect-[1/.92] w-full overflow-hidden bg-muted/30">
            <img
              src={item.previewUrl}
              alt={item.name}
              loading="lazy"
              className="h-full w-full object-cover transition-transform group-hover:scale-105"
            />
          </div>
          <div className="absolute left-1 top-1">
            {item.group === "我的" && <Star className="size-3.5 fill-amber-400 text-amber-400" />}
          </div>
          <div className="absolute right-1 top-1">
            <Badge variant={isStyle ? "secondary" : "default"} className="h-4 px-1 text-[9px]">
              {isStyle ? "样式" : "形状"}
            </Badge>
          </div>
          <p className="line-clamp-2 min-h-[2.2em] whitespace-normal break-words px-1.5 pb-1 pt-0.5 text-[11px] leading-snug">
            {item.name}
          </p>
        </div>
      </ContextMenuTrigger>
      <ContextMenuContent>
        <ContextMenuLabel className="text-xs text-muted-foreground">
          {(item.group || "我的") + " › " + item.category}
        </ContextMenuLabel>
        <ContextMenuSeparator />
        {!isStyle ? (
          <ContextMenuItem onSelect={onInsert}>插入到当前页</ContextMenuItem>
        ) : (
          <>
            <ContextMenuItem onSelect={() => bridge.applyStyle(item.id)}>应用样式到选中形状</ContextMenuItem>
            <ContextMenuItem onSelect={() => bridge.applyStyleSameType(item.id)}>应用到本页同类</ContextMenuItem>
          </>
        )}
        <ContextMenuItem onSelect={onPreview}>预览大图…</ContextMenuItem>
        <ContextMenuSeparator />
        <ContextMenuItem
          onSelect={async () => {
            const g = await dialogs.prompt("输入一级分组（不存在将新建）：", {
              initial: item.group || "我的",
              okText: "移动",
            });
            if (g?.trim())
              bridge
                .createGroup(g.trim())
                .then(() => bridge.setGroup(item.id, g.trim()))
                .then(onChanged);
          }}
        >
          移动到分组…
        </ContextMenuItem>
        <ContextMenuItem
          onSelect={async () => {
            const c = await dialogs.prompt("分类名称（输入新分类或已有分类）：", {
              initial: item.category || "",
              okText: "移动",
            });
            if (c?.trim()) bridge.setCategory(item.id, c.trim()).then(onChanged);
          }}
        >
          移动到分类…
        </ContextMenuItem>
        <ContextMenuItem
          onSelect={async () => {
            const n = await dialogs.prompt("重命名收藏", {
              initial: item.name,
              okText: "重命名",
            });
            if (n?.trim()) bridge.rename(item.id, n.trim()).then(onChanged);
          }}
        >
          重命名
        </ContextMenuItem>
        <ContextMenuSeparator />
        <ContextMenuItem onSelect={() => bridge.openFolder(item.id)}>打开收藏文件夹</ContextMenuItem>
        <ContextMenuItem
          onSelect={() => {
            bridge
              .exportPack([item.id], "")
              .then(onChanged)
              .catch(() => {});
          }}
        >
          导出到素材包…
        </ContextMenuItem>
        <ContextMenuSeparator />
        <ContextMenuItem
          className="text-destructive focus:text-destructive"
          onSelect={async () => {
            const ok = await dialogs.confirm("将「" + item.name + "」移入回收站？", {
              title: "移入回收站",
              okText: "移入回收站",
              danger: true,
            });
            if (ok) bridge.del([item.id]).then(onChanged);
          }}
        >
          移入回收站
        </ContextMenuItem>
      </ContextMenuContent>
    </ContextMenu>
  );
}

/** hover 放大预览（图片查看器式，纯展示不拦截事件；提示文案随素材类型） */
function HoverLens({ data }: { data: { url: string; kind: string } }) {
  const isStyle = data.kind === "Style";
  return (
    <div className="pointer-events-none fixed inset-0 z-50 flex flex-col items-center justify-center bg-black/45 p-6 backdrop-blur-[2px]">
      <div className="flex max-h-[74%] max-w-[min(560px,92%)] items-center justify-center overflow-hidden rounded-xl bg-muted/80 p-2 shadow-2xl">
        <img src={data.url} alt="" className="max-h-[70vh] w-auto rounded object-contain" />
      </div>
      <div className="mt-3 flex items-center gap-1.5 rounded-full bg-black/50 px-3 py-1 text-xs text-white">
        {isStyle ? "单击 应用到选中形状 · 右键更多操作" : "单击 插入到当前页 · 右键更多操作"}
      </div>
    </div>
  );
}

/** 回收站 Dialog */
function TrashDialog({ onClose }: { onClose: () => void }) {
  const [list, setList] = useState<any[] | null>(null);
  useEffect(() => {
    bridge.trashList().then((d) => setList(d.items || []));
  }, []);
  const refresh = () => bridge.trashList().then((d) => setList(d.items || []));
  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-w-md">
        <DialogHeader>
          <DialogTitle>回收站</DialogTitle>
          <DialogDescription>删除的素材先移到这里，可恢复或彻底删除。</DialogDescription>
        </DialogHeader>
        <div className="flex max-h-[50vh] flex-col gap-1 overflow-y-auto">
          {!list ? (
            <Skeleton className="h-10 w-full" />
          ) : list.length === 0 ? (
            <p className="py-6 text-center text-xs text-muted-foreground">回收站是空的。</p>
          ) : (
            list.map((m) => (
              <div key={m.id} className="flex items-center justify-between rounded-md border px-2 py-1.5 text-xs">
                <span className="truncate">{m.name}</span>
                <span className="flex shrink-0 gap-1">
                  <Button
                    variant="outline"
                    size="sm"
                    className="h-6 px-2 text-[11px]"
                    onClick={() => bridge.trashRestore(m.id).then(refresh)}
                  >
                    恢复
                  </Button>
                  <Button
                    variant="destructive"
                    size="sm"
                    className="h-6 px-2 text-[11px]"
                    onClick={async () => {
                      const ok = await dialogs.confirm("彻底删除「" + m.name + "」？此操作不可恢复。", {
                        title: "彻底删除",
                        okText: "彻底删除",
                        danger: true,
                      });
                      if (ok) bridge.trashPurgeOne(m.id).then(refresh);
                    }}
                  >
                    删除
                  </Button>
                </span>
              </div>
            ))
          )}
        </div>
        {list && list.length > 0 && (
          <div className="flex justify-end">
            <Button
              variant="outline"
              size="sm"
              onClick={async () => {
                const ok = await dialogs.confirm("清空回收站？此操作不可恢复。", {
                  title: "清空回收站",
                  okText: "清空",
                  danger: true,
                });
                if (ok) bridge.trashPurge().then(refresh);
              }}
            >
              清空回收站
            </Button>
          </div>
        )}
      </DialogContent>
    </Dialog>
  );
}

/** 管理分组（一级）：列出全部组，可新建自定义组；预设组只读 */
function GroupDialog({ groups, onClose }: { groups: string[]; onClose: () => void }) {
  const presets = ["我的", "模板", "素材", "文字", "背景", "图例"];
  const custom = groups.filter((g) => !presets.includes(g));
  const [list, setList] = useState<string[]>(custom);
  const refresh = () => {
    bridge.init().then((d) => {
      setList((d.groups || []).filter((g: string) => !presets.includes(g)));
      onClose();
    });
  };
  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-w-sm">
        <DialogHeader>
          <DialogTitle>管理一级分组</DialogTitle>
          <DialogDescription>预设组不可改；自定义组可重命名/删除（条目归「我的」）。</DialogDescription>
        </DialogHeader>
        <div className="flex max-h-[45vh] flex-col gap-1.5 overflow-y-auto">
          {list.length === 0 && (
            <p className="py-4 text-center text-xs text-muted-foreground">暂无自定义分组。</p>
          )}
          {list.map((g) => (
            <div key={g} className="flex items-center justify-between rounded-md border px-2 py-1.5 text-xs">
              <span className="truncate">{g}</span>
              <span className="flex shrink-0 gap-1">
                <Button
                  variant="outline"
                  size="sm"
                  className="h-6 px-2 text-[11px]"
                  onClick={async () => {
                    const n = await dialogs.prompt("重命名分组为：", { initial: g, okText: "重命名" });
                    if (n?.trim() && n.trim() !== g) bridge.renameGroup(g, n.trim()).then(refresh);
                  }}
                >
                  重命名
                </Button>
                <Button
                  variant="destructive"
                  size="sm"
                  className="h-6 px-2 text-[11px]"
                  onClick={async () => {
                    const ok = await dialogs.confirm("删除分组「" + g + "」？其下条目移到「我的」。", {
                      title: "删除分组",
                      okText: "删除",
                      danger: true,
                    });
                    if (ok) bridge.deleteGroup(g).then(refresh);
                  }}
                >
                  删除
                </Button>
              </span>
            </div>
          ))}
        </div>
        <div className="flex justify-end gap-2">
          <Button
            variant="outline"
            size="sm"
            onClick={async () => {
              const n = await dialogs.prompt("新建一级分组名称：", { placeholder: "如：课程作业", okText: "创建" });
              if (n?.trim()) bridge.createGroup(n.trim()).then(refresh);
            }}
          >
            新建分组
          </Button>
          <Button size="sm" onClick={onClose}>
            完成
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}


/** 管理二级分类（针对当前一级组）：列出组内真实分类∪自定义分类，可新建/重命名/删除 */
function CategoriesDialog({ group, onClose }: { group: string; onClose: () => void }) {
  const [cats, setCats] = useState<string[] | null>(null);
  const load = () => {
    bridge.init().then((d) => {
      const per = (d.groupCategories as Record<string, string[]>) || {};
      setCats((per[group] as string[]) || []);
    });
  };
  useEffect(load, []); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-w-sm">
        <DialogHeader>
          <DialogTitle>管理分类 · {group}</DialogTitle>
          <DialogDescription>为该一级组新增/重命名/删除二级分类。删除会把该分类下素材归入「未分类」。</DialogDescription>
        </DialogHeader>
        <div className="flex max-h-[45vh] flex-col gap-1.5 overflow-y-auto">
          {!cats ? (
            <Skeleton className="h-8 w-full" />
          ) : cats.length === 0 ? (
            <p className="py-4 text-center text-xs text-muted-foreground">
              该组还没有分类。点下方「新建分类」添加。
            </p>
          ) : (
            cats.map((name) => (
              <div key={name} className="flex items-center justify-between rounded-md border px-2 py-1.5 text-xs">
                <span className="truncate">{name}</span>
                <span className="flex shrink-0 gap-1">
                  <Button
                    variant="outline"
                    size="sm"
                    className="h-6 px-2 text-[11px]"
                    onClick={async () => {
                      const n = await dialogs.prompt("重命名分类「" + name + "」为：", {
                        initial: name,
                        okText: "重命名",
                      });
                      if (n?.trim() && n.trim() !== name)
                        bridge.renameGroupCategory(group, name, n.trim()).then(load);
                    }}
                  >
                    重命名
                  </Button>
                  <Button
                    variant="destructive"
                    size="sm"
                    className="h-6 px-2 text-[11px]"
                    onClick={async () => {
                      const ok = await dialogs.confirm(
                        "删除分类「" + name + "」？其下素材将归入「未分类」（可稍后移回）。",
                        { title: "删除分类", okText: "删除", danger: true }
                      );
                      if (ok) bridge.deleteGroupCategory(group, name).then(load);
                    }}
                  >
                    删除
                  </Button>
                </span>
              </div>
            ))
          )}
        </div>
        <div className="flex justify-end gap-2">
          <Button
            variant="outline"
            size="sm"
            onClick={async () => {
              const n = await dialogs.prompt("新建分类（「" + group + "」组）：", {
                placeholder: "分类名，如：箭头",
                okText: "创建",
              });
              if (n?.trim()) bridge.createGroupCategory(group, n.trim()).then(load);
            }}
          >
            <Plus className="mr-1 size-3.5" /> 新建分类
          </Button>
          <Button size="sm" onClick={onClose}>
            完成
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}

/** 去重清理：扫描完全重复（同类型/数量/尺寸）的条目，可勾选批量移入回收站 */
function DuplicatesDialog({ onClose }: { onClose: () => void }) {
  const [list, setList] = useState<any[] | null>(null);
  const [checked, setChecked] = useState<Set<string>>(new Set());
  const scan = () => {
    setList(null);
    setChecked(new Set());
    bridge
      .scanDuplicates()
      .then((d) => {
        const arr = (d.duplicates as any[]) || [];
        setList(arr);
        // 默认全选可删项
        setChecked(new Set(arr.map((x) => x.id)));
      })
      .catch((e: Error) => dialogs.alert("扫描失败：" + e.message));
  };
  useEffect(() => {
    scan();
  }, []); // eslint-disable-line
  const toggle = (id: string) => {
    setChecked((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };
  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-w-md">
        <DialogHeader>
          <DialogTitle>去重清理</DialogTitle>
          <DialogDescription>
            按 同类型/同形状数量/同尺寸 判定"完全重复"，保留最早收藏，其余可移入回收站（可恢复）。
          </DialogDescription>
        </DialogHeader>
        <div className="flex max-h-[50vh] flex-col gap-1 overflow-y-auto">
          {list === null ? (
            <div className="flex items-center justify-center py-8 text-xs text-muted-foreground">扫描中…</div>
          ) : list.length === 0 ? (
            <p className="py-6 text-center text-xs text-muted-foreground">未发现完全重复的素材。</p>
          ) : (
            list.map((m) => (
              <label
                key={m.id}
                className="flex cursor-pointer items-center gap-2 rounded-md border px-2 py-1.5 text-xs"
              >
                <Checkbox checked={checked.has(m.id)} onCheckedChange={() => toggle(m.id)} />
                <img src={m.previewUrl} alt="" className="h-8 w-12 rounded object-cover" />
                <span className="min-w-0 flex-1 truncate">{m.name}</span>
                <span className="shrink-0 text-muted-foreground">{m.kind === "Style" ? "样式" : "形状"}</span>
              </label>
            ))
          )}
        </div>
        {list !== null && list.length > 0 && (
          <div className="flex justify-between gap-2">
            <Button variant="ghost" size="sm" onClick={scan}>
              重新扫描
            </Button>
            <Button
              size="sm"
              disabled={checked.size === 0}
              onClick={async () => {
                const ids = Array.from(checked);
                const ok = await dialogs.confirm(
                  "将 " + ids.length + " 条重复项移入回收站？（保留最早收藏）",
                  { title: "移入回收站", okText: "移入回收站", danger: true }
                );
                if (ok) {
                  bridge
                    .del(ids)
                    .then(() => {
                      onClose();
                    })
                    .catch((e: Error) => dialogs.alert("删除失败：" + e.message));
                }
              }}
            >
              移入回收站 ({checked.size})
            </Button>
          </div>
        )}
      </DialogContent>
    </Dialog>
  );
}
