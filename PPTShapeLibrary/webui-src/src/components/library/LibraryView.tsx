import { useEffect, useMemo, useState } from "react";
import { bridge, ListItem } from "@/lib/bridge";
import { useLibrary } from "@/hooks/useLibrary";
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
  Folder,
  Search,
  Moon,
  Sun,
  MoreVertical,
  Plus,
  Upload,
  Trash2,
  RotateCcw,
  RotateCw,
  Star,
} from "lucide-react";

const GROUP_ICONS: Record<string, React.ReactNode> = {
  我的: <Star className="size-4" />,
  模板: <span className="text-base">📐</span>,
  素材: <Folder className="size-4" />,
  文字: <span className="text-base">🔤</span>,
  背景: <span className="text-base">🎨</span>,
  图例: <span className="text-base">📊</span>,
};

function tooltip(e: React.MouseEvent) {
  /* noop */
}

export default function LibraryView({
  theme,
  setTheme,
}: {
  theme: string;
  setTheme: (t: string) => void;
}) {
  const lib = useLibrary();
  const { init, items, selected, loadMore } = lib;
  const selCount = selected.size;
  const [preview, setPreview] = useState<ListItem | null>(null);
  const [hoverImg, setHoverImg] = useState<string | null>(null);
  const [showTrash, setShowTrash] = useState(false);

  const groups = useMemo(() => ["全部", ...(init?.groups || [])], [init]);
  const cats = useMemo(() => {
    if (lib.group === "全部") {
      const set = new Set<string>();
      Object.values(init?.groupCategories || {}).forEach((cs) => cs.forEach((c) => set.add(c)));
      return ["全部", ...Array.from(set).sort()];
    }
    return ["全部", ...(init?.groupCategories?.[lib.group] || [])];
  }, [init, lib.group]);

  const toggleTheme = () => {
    const next = theme === "dark" ? "light" : "dark";
    setTheme(next);
    document.documentElement.classList.toggle("dark", next === "dark");
    bridge.saveSettings({ theme: next }).catch(() => {});
  };

  return (
    <div className="flex h-screen flex-col bg-background text-foreground">
      {/* 顶栏 */}
      <header className="flex items-center gap-2 border-b bg-card/70 px-3 py-2 backdrop-blur">
        <span className="flex items-center gap-1.5 text-sm font-semibold">
          <Folder className="size-4 text-primary" /> 素材库
        </span>
        <div className="ml-2 flex flex-1 items-center gap-1 rounded-md bg-muted px-2">
          <Search className="size-3.5 text-muted-foreground" />
          <Input
            className="h-7 border-0 bg-transparent focus-visible:ring-0"
            placeholder="搜索 名称 / 分类 / 标签"
            value={lib.filter}
            onChange={(e) => lib.setFilter(e.target.value)}
          />
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
            <DropdownMenuItem onSelect={() => setShowTrash(true)}>
              <Trash2 className="size-4" /> 回收站
            </DropdownMenuItem>
            <DropdownMenuItem
              onSelect={() => {
                bridge.importPictures().then((d) => {
                  lib.refreshMeta();
                  alert("已导入 " + (d.imported || 0) + " 张图片到 素材 › 图片");
                });
              }}
            >
              <Upload className="size-4" /> 导入图片…
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
        <Button variant="ghost" size="icon" className="h-7 w-7" onClick={toggleTheme}>
          {theme === "dark" ? <Sun className="size-4" /> : <Moon className="size-4" />}
        </Button>
      </header>

      <div className="flex min-h-0 flex-1">
        {/* 左侧一级分组栏 */}
        <nav className="flex w-16 shrink-0 flex-col items-stretch gap-0.5 overflow-y-auto border-r bg-card/40 px-1.5 py-2">
          {groups.map((g) => {
            const count = init?.groupCounts?.[g];
            const active = lib.group === g;
            return (
              <button
                key={g}
                title={g}
                onClick={() => lib.setGroup(g)}
                className={
                  "flex flex-col items-center gap-1 rounded-lg px-0.5 py-2 text-[11px] transition-colors " +
                  (active
                    ? "bg-primary/10 text-primary"
                    : "text-muted-foreground hover:bg-muted hover:text-foreground")
                }
              >
                <span className="relative">
                  {GROUP_ICONS[g] || <Folder className="size-4" />}
                  {!!count && (
                    <span className="absolute -right-2.5 -top-1.5 rounded-full bg-muted px-1 text-[9px] leading-3">
                      {count > 99 ? "99+" : count}
                    </span>
                  )}
                </span>
                <span className="max-w-full truncate">{g}</span>
              </button>
            );
          })}
          <Separator className="my-1" />
          <button
            className="flex flex-col items-center gap-1 rounded-lg border border-dashed px-0.5 py-2 text-[10px] text-muted-foreground hover:border-primary hover:text-primary"
            onClick={() => {
              const name = prompt("新建一级分组名称：");
              if (name?.trim()) bridge.createGroup(name.trim()).then(() => lib.refreshMeta());
            }}
          >
            <Plus className="size-3.5" /> 新建
          </button>
        </nav>

        {/* 主区 */}
        <div className="flex min-w-0 flex-1 flex-col">
          {/* 顶部二级分类胶囊 */}
          <div className="flex items-center gap-1 overflow-x-auto border-b bg-card/40 px-2 py-1.5">
            {cats.map((c) => (
              <button
                key={c}
                onClick={() => lib.setCategory(c)}
                className={
                  "shrink-0 rounded-full px-2.5 py-1 text-xs transition-colors " +
                  (lib.category === c
                    ? "bg-primary text-primary-foreground"
                    : "text-muted-foreground hover:bg-muted")
                }
              >
                {c}
              </button>
            ))}
          </div>

          {/* 状态行 */}
          <div className="flex items-center justify-between px-3 py-1 text-[11px] text-muted-foreground">
            <span>
              {lib.group !== "全部" ? lib.group + " · " : ""}
              {lib.category !== "全部" || lib.kind !== "全部" || lib.filter || lib.tag
                ? `筛选 ${items.length} / ${lib.filteredTotal}`
                : `共 ${lib.total} 项`}
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

          {/* 卡片网格（移出网格区清除 hover 预览） */}
          <div className="min-h-0 flex-1" onMouseLeave={() => setHoverImg(null)}>
            <ScrollArea className="h-full">
              <div className="grid grid-cols-[repeat(auto-fill,minmax(96px,1fr))] gap-2 p-2.5">
                {items.map((item) => (
                  <ItemCard
                    key={item.id}
                    item={item}
                    selected={selected.has(item.id)}
                    onToggle={() => lib.toggleSelect(item.id)}
                    onPreview={() => setPreview(item)}
                    onHover={setHoverImg}
                    onInsert={() => bridge.insert(item.id).then(() => lib.reload(true))}
                    onChanged={() => lib.refreshMeta()}
                  />
                ))}
                {lib.loading &&
                  items.length === 0 &&
                  Array.from({ length: 8 }).map((_, i) => (
                    <div key={i} className="flex flex-col gap-1.5">
                      <Skeleton className="aspect-[1/.92] w-full rounded-lg" />
                      <Skeleton className="h-3 w-3/4" />
                    </div>
                  ))}
              </div>
            </ScrollArea>
          </div>

          {/* hover 大图预览（图片查看器式，纯展示不拦截事件） */}
          {hoverImg && <HoverLens src={hoverImg} />}
        </div>
      </div>

      {/* 批量操作浮条 */}
      {selCount > 0 && (
        <div className="fixed bottom-4 left-1/2 z-40 flex -translate-x-1/2 items-center gap-2 rounded-xl border bg-popover px-3 py-2 shadow-xl">
          <span className="text-xs font-medium">已选 {selCount} 项</span>
          <Button
            variant="outline"
            size="sm"
            className="h-7 text-xs"
            onClick={() => lib.clearSelection()}
          >
            取消
          </Button>
          <Button
            variant="destructive"
            size="sm"
            className="h-7 text-xs"
            onClick={() => {
              const ids = Array.from(selected);
              if (confirm("将选中的 " + ids.length + " 项移入回收站？"))
                bridge.del(ids).then(() => {
                  lib.clearSelection();
                  lib.reload(true);
                });
            }}
          >
            移入回收站
          </Button>
        </div>
      )}

      {/* 大图 Dialog（lightbox） */}
      <Dialog open={!!preview} onOpenChange={(o) => !o && setPreview(null)}>
        {preview && (
          <DialogContent className="max-w-lg">
            <DialogHeader>
              <DialogTitle>{preview.name}</DialogTitle>
              <DialogDescription>
                {preview.group || "我的"} › {preview.category} ·{" "}
                {preview.kind === "Style" ? "样式" : "形状"}
              </DialogDescription>
            </DialogHeader>
            <div className="flex items-center justify-center rounded-lg bg-muted/40 p-2">
              <img
                src={preview.previewUrl}
                alt={preview.name}
                className="max-h-[55vh] w-auto rounded object-contain"
              />
            </div>
            <div className="flex justify-end gap-2">
              {preview.kind !== "Style" && (
                <Button
                  onClick={() => {
                    bridge.insert(preview.id).then(() => {
                      setPreview(null);
                      lib.reload(true);
                    });
                  }}
                >
                  插入到当前页
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
          }}
        />
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
  onChanged,
}: {
  item: ListItem;
  selected: boolean;
  onToggle: () => void;
  onPreview: () => void;
  onHover: (url: string | null) => void;
  onInsert: () => void;
  onChanged: () => void;
}) {
  const isStyle = item.kind === "Style";
  return (
    <ContextMenu>
      <ContextMenuTrigger>
        <div
          className={
            "group relative cursor-pointer overflow-hidden rounded-lg border bg-card shadow-sm transition-all hover:-translate-y-0.5 hover:shadow-md " +
            (selected ? "border-primary ring-2 ring-primary/60" : "border-border")
          }
          onClick={(e) => {
            if (e.ctrlKey) {
              onToggle();
              return;
            }
            onInsert();
          }}
          onDoubleClick={() => {}}
          onMouseEnter={() => onHover(item.previewUrl)}
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
          <p className="truncate px-1.5 pb-1 pt-0.5 text-[11px]">{item.name}</p>
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
          onSelect={() => {
            const g = prompt("输入一级分组（不存在将新建）：", item.group || "我的");
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
          onSelect={() => {
            const n = prompt("新名称：", item.name);
            if (n?.trim()) bridge.rename(item.id, n.trim()).then(onChanged);
          }}
        >
          重命名
        </ContextMenuItem>
        <ContextMenuSeparator />
        <ContextMenuItem
          className="text-destructive focus:text-destructive"
          onSelect={() => {
            if (confirm("将「" + item.name + "」移入回收站？")) bridge.del([item.id]).then(onChanged);
          }}
        >
          移入回收站
        </ContextMenuItem>
      </ContextMenuContent>
    </ContextMenu>
  );
}

/** hover 放大预览 */
function HoverLens({ src }: { src: string }) {
  return (
    <div className="pointer-events-none fixed inset-0 z-50 flex flex-col items-center justify-center bg-black/45 p-6 backdrop-blur-[2px]">
      <div className="flex max-h-[74%] max-w-[min(560px,92%)] items-center justify-center overflow-hidden rounded-xl bg-muted/80 p-2 shadow-2xl">
        <img src={src} alt="" className="max-h-[70vh] w-auto rounded object-contain" />
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
                    onClick={() => {
                      if (confirm("彻底删除「" + m.name + "」？不可恢复。")) bridge.trashPurgeOne(m.id).then(refresh);
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
              onClick={() => {
                if (confirm("清空回收站？不可恢复。")) bridge.trashPurge().then(refresh);
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
