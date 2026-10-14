import { useCallback, useEffect, useRef, useState } from "react";
import { bridge } from "@/lib/bridge";
import { dialogs } from "@/lib/dialogs";
import { X, ChevronDownIcon, RefreshCw } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { Separator } from "@/components/ui/separator";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import {
  Command,
  CommandEmpty,
  CommandGroup,
  CommandInput,
  CommandItem,
  CommandList,
} from "@/components/ui/command";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle, AlertDialogTrigger } from "@/components/ui/alert-dialog";

const META: Record<string, { title: string; desc: string }> = {
  tool1: { title: "文字样式", desc: "批量修改字体/字号/颜色/对齐" },
  tool2: { title: "一键换字体", desc: "取消原字体，直接改为新中/英文字体" },
  tool3: { title: "查找 / 替换", desc: "跨页文本替换" },
  tool4: { title: "文字清理", desc: "去多余空白 / 统一标点" },
  tool5: { title: "导出图片", desc: "整稿导出 PNG/JPG" },
  tool6: { title: "清理瘦身", desc: "删空白页/隐藏对象等" },
  tool7: { title: "统一尺寸", desc: "批量统一形状尺寸" },
  tool8: { title: "图片网格", desc: "多图排网格插入" },
  tool9: { title: "页脚页码", desc: "为全部页加页脚与页码" },
  tool10: { title: "导出 PDF / 大纲", desc: "整稿 PDF 或文本大纲" },
  tool11: { title: "图表美化", desc: "预设调色板美化图表" },
  tool12: { title: "成套模板", desc: "插入封面/目录/内容/结束成套页" },
  settings: { title: "设置", desc: "偏好设置" },
  about: { title: "关于 / 更新", desc: "版本信息、检查更新与更新提醒" },
};

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-1.5">
      <Label className="text-xs text-muted-foreground">{label}</Label>
      {children}
    </div>
  );
}

/** HTML 颜色 #RRGGBB → PowerPoint RGB int(BGR 字节序) */
/**
 * HTML 颜色 #RRGGBB → PowerPoint 颜色 int。
 * PowerPoint 的 .Color.RGB 为 0xBBGGRR 字节序（低字节=R，见 C# HexToRgb: r|g<<8|b<<16）。
 * #RRGGBB 解析后 r=RR,g=GG,b=BB，返回 (BB<<16)|(GG<<8)|RR。
 */
function toBgr(hex: string): number {
  const h = hex.replace("#", "");
  const r = parseInt(h.slice(0, 2), 16) || 0;
  const g = parseInt(h.slice(2, 4), 16) || 0;
  const b = parseInt(h.slice(4, 6), 16) || 0;
  return (b << 16) | (g << 8) | r;
}

/** 拉取字体数据：docFonts=文稿字体, systemFonts=系统全部字体 */
function useFontData() {
  const [docFonts, setDocFonts] = useState<string[]>([]);
  const [systemFonts, setSystemFonts] = useState<string[]>([]);
  const [noDoc, setNoDoc] = useState(false);
  const refresh = () => {
    bridge
      .toolGetFonts()
      .then((d) => {
        setDocFonts((d.fonts as string[]) || []);
        setSystemFonts((d.systemFonts as string[]) || []);
        setNoDoc(!d.hasPresentation);
      })
      .catch(() => {
        setDocFonts([]);
        setSystemFonts([]);
        setNoDoc(true);
      });
  };
  useEffect(() => {
    refresh();
  }, []);
  return { docFonts, systemFonts, noDoc, refresh };
}

/** 字体名匹配：大小写不敏感的子串匹配（对中文/英文混排的字体名都直观） */
function fontFilter(value: string, search: string) {
  const s = search.trim().toLowerCase();
  if (!s) return 1;
  return value.toLowerCase().includes(s) ? 1 : 0;
}

/** 单条字体选项：用字体本身渲染名字，选中项显示勾选 */
function FontPickerItem({
  name,
  selected,
  onSelect,
}: {
  name: string;
  selected: boolean;
  onSelect: (v: string) => void;
}) {
  return (
    <CommandItem
      value={name}
      data-checked={selected ? "true" : undefined}
      onSelect={() => onSelect(name)}
      className="py-2"
    >
      <span className="truncate" style={{ fontFamily: name }}>{name}</span>
    </CommandItem>
  );
}

/**
 * 字体下拉选择器：Popover + Command（可搜索、全量列出、按"文稿/系统"分组）。
 * 之前是自绘 div 弹层：宽度只有触发按钮那么宽、列表 max-h-56、且被 slice(0,80) 截断，
 * 系统 197 个字体只能看到 80 个；现在改为 portal 弹层（不被面板裁剪）+ 可滚动全量列表。
 */
function FontPicker({
  value,
  onChange,
  docFonts,
  systemFonts,
  placeholder,
}: {
  value: string;
  onChange: (v: string) => void;
  docFonts: string[];
  systemFonts: string[];
  placeholder?: string;
}) {
  const [open, setOpen] = useState(false);
  const docSet = new Set(docFonts || []);
  const allFonts = Array.from(new Set([...(docFonts || []), ...(systemFonts || [])])).sort((a, b) =>
    a.localeCompare(b, "zh"),
  );
  const docList = allFonts.filter((f) => docSet.has(f));
  const sysList = allFonts.filter((f) => !docSet.has(f));
  const loading = allFonts.length === 0;

  const select = (f: string) => {
    onChange(f);
    setOpen(false);
  };

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <Button type="button" variant="outline" className="h-8 w-full justify-between px-3 font-normal">
          <span className="truncate" style={value ? { fontFamily: value } : undefined}>
            {value || placeholder || "选择字体…"}
          </span>
          <ChevronDownIcon data-icon="inline-end" className="opacity-60" />
        </Button>
      </PopoverTrigger>
      {/* 宽高都跟着 Radix 的"可用空间"收缩，靠近窗口边缘也不会被裁剪 */}
      <PopoverContent
        align="start"
        sideOffset={4}
        collisionPadding={8}
        className="w-[min(92vw,var(--radix-popover-content-available-width,22rem),22rem)] gap-0 p-0"
      >
        <Command filter={fontFilter}>
          <CommandInput
            autoFocus
            placeholder={loading ? "字体加载中…" : `搜索字体（共 ${allFonts.length} 个）`}
          />
          <CommandList className="max-h-[min(65vh,var(--radix-popover-content-available-height,22rem),22rem)]">
            <CommandEmpty>{loading ? "字体列表加载中…" : "无匹配字体"}</CommandEmpty>
            {docList.length > 0 && (
              <CommandGroup heading={`文稿字体（${docList.length}）`}>
                {docList.map((f) => (
                  <FontPickerItem key={`doc-${f}`} name={f} selected={f === value} onSelect={select} />
                ))}
              </CommandGroup>
            )}
            <CommandGroup heading={`系统字体（${sysList.length}）`}>
              {sysList.map((f) => (
                <FontPickerItem key={`sys-${f}`} name={f} selected={f === value} onSelect={select} />
              ))}
            </CommandGroup>
          </CommandList>
          <div className="flex items-center justify-between border-t px-3 py-1.5 text-[11px] text-muted-foreground">
            <span>共 {allFonts.length} 个字体</span>
            <span className="truncate pl-2" style={value ? { fontFamily: value } : undefined}>
              {value || "未选择"}
            </span>
          </div>
        </Command>
      </PopoverContent>
    </Popover>
  );
}

/** 小工具壳：统一 Card + 结果行 */
function Shell({ title, desc, children }: { title: string; desc: string; children: React.ReactNode }) {
  return (
    <Card>
      <CardHeader className="pb-2">
        <CardTitle className="text-sm">{title}</CardTitle>
        <p className="text-xs text-muted-foreground">{desc}</p>
      </CardHeader>
      <CardContent
        className="flex flex-col gap-3"
        onKeyDown={(e) => {
          // 表单内按 Enter 触发主操作（替换/清理/导出等），提升键盘效率
          if (e.key !== "Enter" || e.shiftKey) return;
          const t = e.target as HTMLElement;
          if (t && (t.tagName === "TEXTAREA" || t.closest("[role='dialog']") || t.closest("[role='listbox']") || t.closest("[role='combobox']"))) return;
          const card = e.currentTarget;
          // 找第一个可用的主按钮（primary 优先，其次非 outline/ghost）
          const btns = Array.from(card.querySelectorAll("button:not([disabled])")) as HTMLButtonElement[];
          const primary =
            btns.find((b) => b.className.includes("primary")) ||
            btns.find((b) => !b.className.includes("outline") && !b.className.includes("ghost") && b.textContent);
          if (primary) {
            e.preventDefault();
            primary.click();
          }
        }}
      >
        {children}
      </CardContent>
    </Card>
  );
}

/** 结果提示 */
function Result({ out }: { out: string }) {
  if (!out) return null;
  return (
    <>
      <Separator />
      <p className="whitespace-pre-wrap text-xs text-muted-foreground">{out}</p>
    </>
  );
}

export default function ToolView({ view }: { view: string }) {
  const meta = META[view] || { title: "工具", desc: "" };
  useEffect(() => {
    const head = document.getElementById("popup-drag");
    if (!head) return;
    const onDown = (e: MouseEvent) => {
      // 交互控件(关闭钮/可点元素)上按下不拖动窗口，避免系统拖拽吞掉 click 导致关不掉
      const t = e.target as HTMLElement;
      if (t && (t.closest("button") || t.closest("input") || t.closest("select") || t.closest("[role='combobox']"))) return;
      try {
        (window as any).chrome?.webview?.postMessage(JSON.stringify({ type: "popupDrag" }));
      } catch {}
    };
    head.addEventListener("mousedown", onDown);
    return () => head.removeEventListener("mousedown", onDown);
  }, []);

  // Esc 兜底关闭（宿主无边框无系统按钮）
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        try {
          (window as any).chrome?.webview?.postMessage(JSON.stringify({ type: "popupClose" }));
        } catch {}
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  const close = () => {
    try {
      (window as any).chrome?.webview?.postMessage(JSON.stringify({ type: "popupClose" }));
    } catch {}
  };

  return (
    <div className="flex h-screen flex-col bg-background text-foreground">
      <div
        id="popup-drag"
        className="flex cursor-move select-none items-center justify-between border-b bg-card px-4 py-2.5"
      >
        <div className="flex items-center gap-2">
          <h1 className="text-sm font-semibold">{meta.title}</h1>
          <span className="text-xs text-muted-foreground">{meta.desc}</span>
        </div>
        <Button
          variant="ghost"
          size="icon"
          className="h-7 w-7 rounded-md text-muted-foreground hover:bg-muted hover:text-foreground"
          aria-label="关闭"
          onClick={close}
        >
          <X className="size-4" />
        </Button>
      </div>
      <div className="flex-1 overflow-y-auto p-4">
        <ToolBody view={view} />
      </div>
    </div>
  );
}

function ToolBody({ view }: { view: string }) {
  const metaObj = META[view] || { title: "工具", desc: "" };
  if (view === "settings") return <SettingsForm />;
  if (view === "about") return <AboutForm />;

  switch (view) {
    case "tool1": return <Tool1 />;
    case "tool2": return <Tool2 />;
    case "tool3": return <Tool3 />;
    case "tool4": return <Tool4 />;
    case "tool5": return <Tool5 />;
    case "tool6": return <Tool6 />;
    case "tool7": return <Tool7 />;
    case "tool8": return <Tool8 />;
    case "tool9": return <Tool9 />;
    case "tool10": return <Tool10 />;
    case "tool11": return <Tool11 />;
    case "tool12": return <Tool12 />;
    default:
      return (
        <Shell title={metaObj.title} desc={metaObj.desc}>
          <p className="text-xs text-muted-foreground">工具迁移中（view={view}）。</p>
        </Shell>
      );
  }
}

function useRun() {
  const [out, setOut] = useState("");
  const [running, setRunning] = useState(false);
  async function run(action: () => Promise<any>) {
    if (running) return; // 防重复点击
    setRunning(true);
    setOut("处理中…请稍候");
    try {
      const d = await action();
      setOut(d?.message || "完成");
    } catch (e: any) {
      setOut("失败：" + (e?.message || e));
    } finally {
      setRunning(false);
    }
  }
  return { out, run, running };
}

/** 工具预设记忆：一次取回全部 prefs；set(key,value) 自动持久化。 */
function usePrefs() {
  const [prefs, setPrefs] = useState<Record<string, string>>({});
  useEffect(() => {
    bridge
      .getToolPrefs()
      .then((d: any) => setPrefs(d?.prefs || {}))
      .catch(() => {});
  }, []);
  const setPref = (key: string, value: string) => {
    setPrefs((prev) => ({ ...prev, [key]: value }));
    bridge.setToolPref(key, value).catch(() => {});
  };
  const getPref = (key: string, fallback = "") => prefs[key] ?? fallback;
  return { prefs, getPref, setPref };
}

/** tool1 文字样式 */
function Tool1() {
  const { out, run } = useRun();
  const { docFonts, systemFonts, noDoc, refresh } = useFontData();
  const [scope, setScope] = useState("all");
  const [font, setFont] = useState("");
  const [size, setSize] = useState("");
  const [bold, setBold] = useState(false);
  const [italic, setItalic] = useState(false);
  const [align, setAlign] = useState("keep");
  const [color, setColor] = useState("");
  const selColor = color ? color : "";
  const apply = () =>
    run(() =>
      bridge.applyTextStyle({
        scope,
        fontName: font.trim(),
        fontSize: parseFloat(size) || undefined,
        bold: bold || undefined,
        italic: italic || undefined,
        align: align === "keep" ? undefined : align,
        // HTML #RRGGBB → PowerPoint .Color.RGB 为 BGR 序
        color: selColor ? toBgr(selColor) : undefined,
      })
    );
  return (
    <Shell title="文字样式" desc="批量改字体/字号/粗细/对齐/颜色（留空不改）">
      {noDoc && <p className="rounded-md bg-muted px-2 py-1.5 text-xs text-muted-foreground">请先打开演示文稿。</p>}
      <Field label="作用范围">
        <Select value={scope} onValueChange={setScope}>
          <SelectTrigger><SelectValue /></SelectTrigger>
          <SelectContent>
            <SelectItem value="selection">选中形状</SelectItem>
            <SelectItem value="page">当前页</SelectItem>
            <SelectItem value="all">全部页面</SelectItem>
          </SelectContent>
        </Select>
      </Field>
      <Field label="字体（留空不改）">
        <FontPicker value={font} onChange={setFont} docFonts={docFonts} systemFonts={systemFonts} placeholder="选择字体…" />
      </Field>
      <div className="grid grid-cols-2 gap-2">
        <Field label="字号">
          <Input value={size} onChange={(e) => setSize(e.target.value)} placeholder="如 28" inputMode="decimal" />
        </Field>
        <Field label="对齐">
          <Select value={align} onValueChange={setAlign}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="keep">保持原样</SelectItem>
              <SelectItem value="left">左对齐</SelectItem>
              <SelectItem value="center">居中</SelectItem>
              <SelectItem value="right">右对齐</SelectItem>
            </SelectContent>
          </Select>
        </Field>
      </div>
      <div className="grid grid-cols-2 gap-2">
        <div className="flex items-center gap-2">
          <Checkbox id="t1b" checked={bold} onCheckedChange={(v) => setBold(!!v)} />
          <Label htmlFor="t1b" className="text-sm">加粗</Label>
        </div>
        <div className="flex items-center gap-2">
          <Checkbox id="t1i" checked={italic} onCheckedChange={(v) => setItalic(!!v)} />
          <Label htmlFor="t1i" className="text-sm">斜体</Label>
        </div>
      </div>
      <Field label="字体颜色（留空不改）">
        <div className="flex items-center gap-2">
          <input
            type="color"
            className="h-8 w-10 cursor-pointer rounded border bg-transparent p-0.5"
            value={selColor || "#000000"}
            onChange={(e) => setColor(e.target.value)}
          />
          <span className="text-xs text-muted-foreground">{selColor || "未设置（保持原色）"}</span>
        </div>
      </Field>
      <Button onClick={apply} disabled={noDoc} title={noDoc ? "请先打开演示文稿" : undefined}>
        应用
      </Button>
      {noDoc && <p className="text-xs text-muted-foreground">打开演示文稿后即可应用。</p>}
      <Result out={out} />
    </Shell>
  );
}

/** tool2 一键换字体 */
function Tool2() {
  const { out, run } = useRun();
  const { docFonts, systemFonts, noDoc } = useFontData();
  const [cnFont, setCnFont] = useState("");
  const [enFont, setEnFont] = useState("");
  const [all, setAll] = useState(true);
  const cn = cnFont.trim();
  const en = enFont.trim();
  const ready = !!cn || !!en;
  return (
    <Shell title="一键换字体" desc="取消原有字体设置，直接改为新的中文字体 / 英文字体">
      {noDoc && <p className="rounded-md bg-muted px-2 py-1.5 text-xs text-muted-foreground">未检测到打开的演示文稿——字体列表已含系统全部字体，可直接选用。</p>}
      <Field label="中文字体（覆盖原中文字体）">
        <FontPicker value={cnFont} onChange={setCnFont} docFonts={docFonts} systemFonts={systemFonts} placeholder="选择/搜索中文字体…" />
      </Field>
      <Field label="英文字体（覆盖原英文字体）">
        <FontPicker value={enFont} onChange={setEnFont} docFonts={docFonts} systemFonts={systemFonts} placeholder="选择/搜索英文字体…" />
      </Field>
      <div className="flex items-center gap-2">
        <Checkbox id="t2all" checked={all} onCheckedChange={(v) => setAll(!!v)} />
        <Label htmlFor="t2all" className="text-sm">整份演示文稿（取消则当前页）</Label>
      </div>
      <Button
        disabled={!ready}
        onClick={() => run(() => bridge.setFonts({ chineseFont: cn, latinFont: en, allSlides: all }))}
      >
        直接改为新字体
      </Button>
      {!ready && <p className="text-xs text-muted-foreground">请选择新字体：中文字体、英文字体至少选一个。</p>}
      {ready && (!cn || !en) && (
        <p className="text-xs text-muted-foreground">未填写的一项保持原样；两项都填可同时覆盖中英文（含表格与组合形状内文字）。</p>
      )}
      <Result out={out} />
    </Shell>
  );
}

/** tool3 查找 / 替换 */
function Tool3() {
  const { out, run } = useRun();
  const [find, setFind] = useState("");
  const [rep, setRep] = useState("");
  const [all, setAll] = useState(true);
  const [match, setMatch] = useState(false);
  return (
    <Shell title="查找 / 替换" desc="跨页替换文本（保留格式）">
      <Field label="查找内容"><Input value={find} onChange={(e) => setFind(e.target.value)} /></Field>
      <Field label="替换为"><Input value={rep} onChange={(e) => setRep(e.target.value)} /></Field>
      <div className="flex items-center gap-2">
        <Checkbox id="t3all" checked={all} onCheckedChange={(v) => setAll(!!v)} />
        <Label htmlFor="t3all" className="text-sm">整份演示文稿（取消则当前页）</Label>
      </div>
      <div className="flex items-center gap-2">
        <Checkbox id="t3case" checked={match} onCheckedChange={(v) => setMatch(!!v)} />
        <Label htmlFor="t3case" className="text-sm">区分大小写</Label>
      </div>
      <Button
        disabled={!find}
        onClick={() => run(() => bridge.findReplace({ find, replace: rep, matchCase: match, allSlides: all }))}
      >
        替换全部
      </Button>
      {!find && <p className="text-xs text-muted-foreground">请先输入查找内容。</p>}
      {find && !rep && <p className="text-xs text-muted-foreground">提示：替换为空即删除查找到的文本。</p>}
      <Result out={out} />
    </Shell>
  );
}

/** tool4 文字清理 */
function Tool4() {
  const { out, run } = useRun();
  const [ws, setWs] = useState(true);
  const [punct, setPunct] = useState(false);
  return (
    <Shell title="文字清理" desc="整份演示文稿文字规整">
      <div className="flex items-center gap-2">
        <Checkbox id="t4ws" checked={ws} onCheckedChange={(v) => setWs(!!v)} />
        <Label htmlFor="t4ws" className="text-sm">去多余空白（全角空格/连续空格/行首尾）</Label>
      </div>
      <div className="flex items-center gap-2">
        <Checkbox id="t4p" checked={punct} onCheckedChange={(v) => setPunct(!!v)} />
        <Label htmlFor="t4p" className="text-sm">中文间西文标点转全角</Label>
      </div>
      <Button variant="secondary" onClick={() => run(() => bridge.cleanText({ trimWhitespace: ws, unifyPunctuation: punct }))}>开始清理</Button>
      <Result out={out} />
    </Shell>
  );
}

/** tool5 导出图片 */
function Tool5() {
  const { out, run } = useRun();
  const { prefs, setPref } = usePrefs();
  const [dir, setDir] = useState("");
  const [fmt, setFmt] = useState("PNG");
  const [w, setW] = useState("0");
  // 预设回填（首帧 prefs 到达后）
  const inited = useRef(false);
  useEffect(() => {
    if (inited.current || !prefs || Object.keys(prefs).length === 0) return;
    inited.current = true;
    if (prefs.tool5_dir) setDir(prefs.tool5_dir);
    if (prefs.tool5_fmt) setFmt(prefs.tool5_fmt);
    if (prefs.tool5_w) setW(prefs.tool5_w);
  }, [prefs]);
  return (
    <Shell title="导出图片" desc="把整份演示文稿逐页导出为图片">
      <Field label="导出目录">
        <div className="flex gap-1.5">
          <Input value={dir} onChange={(e) => setDir(e.target.value)} placeholder="选择目录…" />
          <Button variant="outline" onClick={() => bridge.requestFolder("选择导出目录").then((d) => { if (d?.path) { setDir(d.path); setPref("tool5_dir", d.path); } })}>浏览…</Button>
        </div>
      </Field>
      <div className="flex items-center gap-3">
        <Field label="格式">
          <Select value={fmt} onValueChange={(v) => { setFmt(v); setPref("tool5_fmt", v); }}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent><SelectItem value="PNG">PNG</SelectItem><SelectItem value="JPG">JPG</SelectItem></SelectContent>
          </Select>
        </Field>
        <Field label="宽度">
          <Select value={w} onValueChange={(v) => { setW(v); setPref("tool5_w", v); }}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent><SelectItem value="0">原始</SelectItem><SelectItem value="1280">1280</SelectItem><SelectItem value="1920">1920</SelectItem></SelectContent>
          </Select>
        </Field>
      </div>
      <Button
        disabled={!dir.trim()}
        onClick={() => run(() => bridge.formatExportImages({ dir: dir.trim(), format: fmt, width: parseInt(w) || 0 }))}
      >
        开始导出
      </Button>
      {!dir.trim() && <p className="text-xs text-muted-foreground">请先选择导出目录。</p>}
      <Result out={out} />
    </Shell>
  );
}

/** tool6 清理瘦身 */
function Tool6() {
  const { out, run } = useRun();
  const [delEmpty, setDelEmpty] = useState(true);
  const [delHidden, setDelHidden] = useState(false);
  const [delOff, setDelOff] = useState(false);
  const [clearNotes, setClearNotes] = useState(true);
  const [removeAct, setRemoveAct] = useState(false);
  const opts = { delEmpty, delHidden, delOffSlide: delOff, clearNotes, removeActions: removeAct };
  return (
    <Shell title="清理瘦身" desc="先扫描统计，再执行清理（不可恢复，建议先保存）">
      {[
        ["删除空白幻灯片", delEmpty, setDelEmpty],
        ["删除隐藏对象", delHidden, setDelHidden],
        ["删除页外对象（完全超出画布）", delOff, setDelOff],
        ["清空备注文字", clearNotes, setClearNotes],
        ["移除形状动作/超链接", removeAct, setRemoveAct],
      ].map(([text, val, set]) => (
        <div key={text as string} className="flex items-center gap-2">
          <Checkbox id={"t6" + text} checked={val as boolean} onCheckedChange={(v) => (set as any)(!!v)} />
          <Label htmlFor={"t6" + text} className="text-sm">{text as string}</Label>
        </div>
      ))}
      <div className="flex gap-2">
        <Button variant="outline" onClick={() => run(() => bridge.formatScanClean(opts))}>扫描统计</Button>
        <AlertDialog>
          <AlertDialogTrigger asChild>
            <Button variant="destructive">执行清理</Button>
          </AlertDialogTrigger>
          <AlertDialogContent>
            <AlertDialogHeader>
              <AlertDialogTitle>确认执行清理？</AlertDialogTitle>
              <AlertDialogDescription>此操作不可恢复，建议先保存文档。</AlertDialogDescription>
            </AlertDialogHeader>
            <AlertDialogFooter>
              <AlertDialogCancel>取消</AlertDialogCancel>
              <AlertDialogAction onClick={() => run(() => bridge.formatRunClean(opts))}>执行</AlertDialogAction>
            </AlertDialogFooter>
          </AlertDialogContent>
        </AlertDialog>
      </div>
      <Result out={out} />
    </Shell>
  );
}

/** tool7 统一尺寸 */
function Tool7() {
  const { out, run } = useRun();
  const { prefs, setPref } = usePrefs();
  const [scope, setScope] = useState("selection");
  const [w, setW] = useState("");
  const [h, setH] = useState("");
  const inited = useRef(false);
  useEffect(() => {
    if (inited.current || !prefs || Object.keys(prefs).length === 0) return;
    inited.current = true;
    if (prefs.tool7_scope) setScope(prefs.tool7_scope);
    if (prefs.tool7_w) setW(prefs.tool7_w);
    if (prefs.tool7_h) setH(prefs.tool7_h);
  }, [prefs]);
  return (
    <Shell title="统一尺寸" desc="按宽/高统一形状（可只填一项，保持比例）">
      <Field label="范围">
        <Select value={scope} onValueChange={(v) => { setScope(v); setPref("tool7_scope", v); }}>
          <SelectTrigger><SelectValue /></SelectTrigger>
          <SelectContent>
            <SelectItem value="selection">选中形状</SelectItem>
            <SelectItem value="page">当前页</SelectItem>
            <SelectItem value="all">全部页面</SelectItem>
          </SelectContent>
        </Select>
      </Field>
      <div className="flex items-center gap-3">
        <Field label="宽度(磅)"><Input value={w} onChange={(e) => setW(e.target.value)} onBlur={() => setPref("tool7_w", w)} inputMode="decimal" /></Field>
        <Field label="高度(磅)"><Input value={h} onChange={(e) => setH(e.target.value)} onBlur={() => setPref("tool7_h", h)} inputMode="decimal" /></Field>
      </div>
      <Button
        disabled={!parseFloat(w) && !parseFloat(h)}
        onClick={() => {
          setPref("tool7_w", w);
          setPref("tool7_h", h);
          run(() => bridge.formatResize({ scope, width: parseFloat(w) || 0, height: parseFloat(h) || 0, forceImageRatio: false }));
        }}
      >
        应用尺寸
      </Button>
      {!parseFloat(w) && !parseFloat(h) && <p className="text-xs text-muted-foreground">请至少填写宽度或高度。</p>}
      <Result out={out} />
    </Shell>
  );
}

/** tool8 图片网格 */
function Tool8() {
  const { out, run } = useRun();
  const [files, setFiles] = useState<string[]>([]);
  const [cols, setCols] = useState("3");
  const [toNew, setToNew] = useState(true);
  return (
    <Shell title="图片网格插入" desc="多图按网格插入">
      <Field label="图片">
        <div className="flex gap-1.5">
          <Input value={files.length ? files.length + " 张已选" : ""} readOnly placeholder="未选择图片" />
          <Button variant="outline" onClick={() =>
            bridge
              .requestFiles("选择图片", "图片文件 (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp")
              .then((d) => setFiles(d.files || []))
              .catch((e: Error) => dialogs.alert("选择图片失败：" + (e?.message || e)).catch(() => {}))
          }>选择…</Button>
        </div>
      </Field>
      <Field label="每行张数"><Input value={cols} onChange={(e) => setCols(e.target.value)} inputMode="numeric" /></Field>
      <div className="flex items-center gap-2">
        <Checkbox id="t8n" checked={toNew} onCheckedChange={(v) => setToNew(!!v)} />
        <Label htmlFor="t8n" className="text-sm">排满自动翻新页（取消则插入当前页）</Label>
      </div>
      <Button
        disabled={files.length === 0}
        onClick={() => run(() => bridge.formatImageGrid({ files, cols: parseInt(cols) || 3, toNewPage: toNew }))}
      >
        生成网格
      </Button>
      {files.length === 0 && <p className="text-xs text-muted-foreground">请先选择至少一张图片。</p>}
      <Result out={out} />
    </Shell>
  );
}

/** tool9 页脚页码 */
function Tool9() {
  const { out, run } = useRun();
  const [text, setText] = useState("");
  const [showDate, setShowDate] = useState(true);
  const [showPage, setShowPage] = useState(true);
  const [firstOnly, setFirstOnly] = useState(false);
  return (
    <Shell title="页脚页码" desc="为页面添加页脚文本/日期/页码（可撤销）">
      <Field label="页脚文本（姓名/单位，可留空）">
        <Input value={text} onChange={(e) => setText(e.target.value)} placeholder="张三 · 课题组" />
      </Field>
      <div className="flex items-center gap-2">
        <Checkbox id="t9d" checked={showDate} onCheckedChange={(v) => setShowDate(!!v)} />
        <Label htmlFor="t9d" className="text-sm">显示日期</Label>
      </div>
      <div className="flex items-center gap-2">
        <Checkbox id="t9p" checked={showPage} onCheckedChange={(v) => setShowPage(!!v)} />
        <Label htmlFor="t9p" className="text-sm">显示页码</Label>
      </div>
      <div className="flex items-center gap-2">
        <Checkbox id="t9f" checked={firstOnly} onCheckedChange={(v) => setFirstOnly(!!v)} />
        <Label htmlFor="t9f" className="text-sm">页码仅首页显示（文本/日期仍加到每页）</Label>
      </div>
      <Button
        disabled={!showDate && !showPage && !text.trim()}
        onClick={() => run(() => bridge.academicFooter({ footerText: text, showDate, showPageNumber: showPage, firstOnly }))}
      >
        添加
      </Button>
      {!showDate && !showPage && !text.trim() && (
        <p className="text-xs text-muted-foreground">请至少勾选日期/页码或填写页脚文本。</p>
      )}
      <Result out={out} />
    </Shell>
  );
}

/** tool10 导出 PDF / 大纲 */
function Tool10() {
  const { run } = useRun();
  const [mode, setMode] = useState("pdf");
  const [dir, setDir] = useState("");
  const [name, setName] = useState("");
  const [out, setOut] = useState("");
  const [outText, setOutText] = useState("");
  const { prefs, setPref } = usePrefs();
  const prefInited = useRef(false);
  useEffect(() => {
    if (prefInited.current || !prefs || Object.keys(prefs).length === 0) return;
    prefInited.current = true;
    if (prefs.tool10_dir) setDir(prefs.tool10_dir);
    if (prefs.tool10_name) setName(prefs.tool10_name);
  }, [prefs]);
  async function go() {
    try {
      setPref("tool10_dir", dir.trim());
      if (!outline) setPref("tool10_name", name.trim());
      const d = await (mode === "pdf"
        ? bridge.academicExportPdf({ dir: dir.trim(), name: name.trim(), style: 0, perPage: 1 })
        : bridge.academicExportOutline({ dir: dir.trim(), name: name.trim() }));
      setOut(d?.message || "完成");
      setOutText(mode === "outline" && !dir.trim() ? d?.text || "" : "");
    } catch (e: any) {
      setOut("失败：" + (e?.message || e));
      setOutText("");
    }
  }
  const outline = mode === "outline";
  return (
    <Shell title="导出 PDF / 大纲" desc={outline ? "整稿文字生成大纲：填目录存文件，或留空直接预览/复制" : "整稿转 PDF"}>
      <Field label="类型">
        <Select value={mode} onValueChange={(v) => { setMode(v); setOutText(""); }}>
          <SelectTrigger><SelectValue /></SelectTrigger>
          <SelectContent><SelectItem value="pdf">PDF</SelectItem><SelectItem value="outline">大纲文本</SelectItem></SelectContent>
        </Select>
      </Field>
      <Field label={outline ? "保存目录（可选，留空则直接预览）" : "保存目录"}>
        <div className="flex gap-1.5">
          <Input value={dir} onChange={(e) => setDir(e.target.value)} placeholder="选择目录…" />
          <Button variant="outline" onClick={() => bridge.requestFolder("选择保存目录").then((d) => { if (d?.path) { setDir(d.path); setPref("tool10_dir", d.path); } })}>浏览…</Button>
        </div>
      </Field>
      {!outline && <Field label="文件名（可留空，自动命名）"><Input value={name} onChange={(e) => setName(e.target.value)} /></Field>}
      <Button
        disabled={!outline && !dir.trim()}
        onClick={go}
      >
        {outline ? (dir.trim() ? "生成大纲文件" : "生成大纲") : "导出 PDF"}
      </Button>
      {!outline && !dir.trim() && <p className="text-xs text-muted-foreground">请先选择保存目录。</p>}
      {outline && !dir.trim() && <p className="text-xs text-muted-foreground">留空目录将直接显示大纲文本，可一键复制。</p>}
      <Result out={out} />
      {outText && (
        <div className="flex flex-col gap-1.5">
          <div className="max-h-56 overflow-y-auto rounded-md border bg-muted/40 p-2">
            <pre className="whitespace-pre-wrap text-xs">{outText}</pre>
          </div>
          <Button
            variant="outline"
            size="sm"
            className="self-end"
            onClick={() => {
              navigator.clipboard?.writeText(outText).then(
                () => setOut("已复制大纲文本"),
                () => setOut("复制失败，请手动选择复制")
              );
            }}
          >
            复制大纲
          </Button>
        </div>
      )}
    </Shell>
  );
}

/** tool11 图表美化 */
function Tool11() {
  const { out, run } = useRun();
  const [palette, setPalette] = useState("0");
  return (
    <Shell title="图表美化" desc="选中图表后套用预设配色">
      <Field label="配色方案">
        <Select value={palette} onValueChange={setPalette}>
          <SelectTrigger><SelectValue /></SelectTrigger>
          <SelectContent>
            <SelectItem value="0">学术蓝</SelectItem>
            <SelectItem value="1">清新绿</SelectItem>
            <SelectItem value="2">暖橙</SelectItem>
            <SelectItem value="3">莫兰迪</SelectItem>
          </SelectContent>
        </Select>
      </Field>
      <Button onClick={() => run(() => bridge.academicBeautifyChart({ palette: parseInt(palette), fontName: "微软雅黑" }))}>美化选中图表</Button>
      <Result out={out} />
    </Shell>
  );
}

/** tool12 成套模板 */
function Tool12() {
  const { out, run } = useRun();
  const [type, setType] = useState("毕业答辩");
  return (
    <Shell title="成套模板" desc="在文稿末尾插入封面/目录/内容/结束成套页">
      <Field label="模板类型">
        <Select value={type} onValueChange={setType}>
          <SelectTrigger><SelectValue /></SelectTrigger>
          <SelectContent>
            {["毕业答辩", "开题报告", "课程汇报", "组会汇报"].map((t) => (
              <SelectItem key={t} value={t}>{t}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      </Field>
      <Button onClick={() => run(() => bridge.insertTemplatePack({ type, contentSlides: 3 }))}>插入成套模板</Button>
      <Result out={out} />
    </Shell>
  );
}

/** 关于 / 更新：版本信息 + 检查更新 + 更新内容 + 提醒间隔。
 *  从素材库面板的「⋯」菜单里搬过来的 —— 更新是插件级信息，与设置并列更符合预期。 */
function AboutForm() {
  const [info, setInfo] = useState<any>(null);
  const [busy, setBusy] = useState(false);
  const [minutes, setMinutes] = useState<number>(10);

  const check = useCallback(async () => {
    setBusy(true);
    try {
      setInfo(await bridge.checkUpdate());
    } catch (e: any) {
      setInfo({ ok: false, error: e?.message || String(e) });
    } finally {
      setBusy(false);
    }
  }, []);

  useEffect(() => {
    void check();
  }, [check]);

  const notes = ((info?.notes as string) || "").trim();
  const hasNotes = notes.length > 0 && !notes.startsWith("（");

  return (
    <Shell title="关于 / 更新" desc="版本信息、检查更新与更新提醒">
      <div className="space-y-5">
        <div className="rounded-lg border bg-muted/30 px-4 py-3">
          <div className="flex items-baseline gap-2">
            <span className="text-xs text-muted-foreground">当前版本</span>
            <span className="text-lg font-semibold">{info?.currentVersion || "…"}</span>
            {info?.ok && info?.hasUpdate && (
              <span className="rounded-full bg-blue-100 px-2 py-0.5 text-[11px] font-medium text-blue-700">
                有新版本 {info.latestVersion}
              </span>
            )}
            {info?.ok && !info?.hasUpdate && (
              <span className="rounded-full bg-emerald-100 px-2 py-0.5 text-[11px] font-medium text-emerald-700">
                已是最新
              </span>
            )}
          </div>
          <div className="mt-3 flex items-center gap-2">
            <Button variant="outline" onClick={() => void check()} disabled={busy}>
              <RefreshCw className="mr-1 size-4" />
              {busy ? "检查中…" : "检查更新"}
            </Button>
            {info?.ok && info?.hasUpdate && (
              <Button
                onClick={async () => {
                  try {
                    const r = (await bridge.applyUpdate()) as any;
                    if (!r?.ok) dialogs.alert("无法开始更新：" + (r?.message || "未知原因"));
                  } catch (e: any) {
                    dialogs.alert("无法开始更新：" + (e?.message || e));
                  }
                }}
              >
                立即更新到 {info.latestVersion}
              </Button>
            )}
            <Button
              variant="ghost"
              onClick={() => {
                void bridge.openRelease(info?.releaseUrl || "").catch(() => {});
              }}
            >
              打开发布页
            </Button>
          </div>
        </div>

        <div>
          <div className="mb-1.5 text-xs text-muted-foreground">更新内容</div>
          {hasNotes ? (
            <pre className="max-h-56 overflow-auto whitespace-pre-wrap break-words rounded-lg border bg-card p-3 font-sans text-[12px] leading-relaxed">
              {notes}
            </pre>
          ) : (
            <div className="rounded-lg border bg-card px-3 py-2 text-[12px] text-muted-foreground">
              {info?.ok ? "当前版本没有更新说明。" : busy ? "正在获取…" : "暂时获取不到更新内容（可能网络不通）。"}
            </div>
          )}
        </div>

        <div>
          <div className="mb-1.5 text-xs text-muted-foreground">更新提醒间隔</div>
          <div className="flex flex-wrap gap-1.5">
            {[0, 10, 30, 60].map((m) => (
              <Button
                key={m}
                variant={minutes === m ? "default" : "outline"}
                size="sm"
                onClick={() => {
                  setMinutes(m);
                  void bridge.setUpdateReminder(m).catch(() => {});
                }}
              >
                {m === 0 ? "关闭提醒" : m + " 分钟"}
              </Button>
            ))}
          </div>
          <p className="mt-1.5 text-[11px] text-muted-foreground">
            有新版时到点弹窗提醒（默认 10 分钟）。打开素材库面板时也会在顶部显示一条提示。
          </p>
        </div>
      </div>
    </Shell>
  );
}

/** settings（深色主题已移除，固定浅色） */
function SettingsForm() {
  const [s, setS] = useState<Record<string, any> | null>(null);
  useEffect(() => {
    bridge.getSettings().then((d) => setS({ ...d })).catch(() => setS({}));
  }, []);
  const set = (k: string, v: any) => setS((prev) => (prev ? { ...prev, [k]: v } : prev));
  const [saved, setSaved] = useState(false);

  if (!s) return null;

  const save = () => {
    bridge.saveSettings(s).then(() => {
      setSaved(true);
      setTimeout(() => setSaved(false), 1500);
    });
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-sm">偏好设置</CardTitle>
        <p className="text-xs text-muted-foreground">改动后点击保存生效</p>
      </CardHeader>
      <CardContent className="flex flex-col gap-3">
        <Row label="插入位置">
          <Select value={s.insertPosition === "original" ? "original" : "center"} onValueChange={(v) => set("insertPosition", v)}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="center">页面居中</SelectItem>
              <SelectItem value="original">原始位置</SelectItem>
            </SelectContent>
          </Select>
        </Row>
        <Row label="卡片大小">
          <Select value={s.cardSize || "compact"} onValueChange={(v) => set("cardSize", v)}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="compact">紧凑</SelectItem>
              <SelectItem value="standard">标准</SelectItem>
              <SelectItem value="large">大图</SelectItem>
            </SelectContent>
          </Select>
        </Row>
        <Row label="每行列数">
          <Select value={s.gridColumns || "2"} onValueChange={(v) => set("gridColumns", v)}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="2">2 列</SelectItem>
              <SelectItem value="3">3 列</SelectItem>
              <SelectItem value="4">4 列</SelectItem>
              <SelectItem value="auto">自适应</SelectItem>
            </SelectContent>
          </Select>
        </Row>
        <Row label="强调色">
          <Select value={s.accentColor || "blue"} onValueChange={(v) => set("accentColor", v)}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="blue">蓝</SelectItem>
              <SelectItem value="indigo">靛蓝</SelectItem>
              <SelectItem value="violet">紫</SelectItem>
              <SelectItem value="emerald">绿</SelectItem>
              <SelectItem value="rose">玫红</SelectItem>
              <SelectItem value="amber">琥珀</SelectItem>
            </SelectContent>
          </Select>
        </Row>
        <Row label="分页数量">
          <Input className="h-8 w-24" inputMode="numeric" value={s.pageSize ?? 200} onChange={(e) => set("pageSize", parseInt(e.target.value) || 200)} />
        </Row>
        <Row label="面板宽度">
          <Input className="h-8 w-24" inputMode="numeric" value={s.paneWidth ?? 300} onChange={(e) => set("paneWidth", parseInt(e.target.value) || 300)} />
        </Row>
        <CheckRow label="按形状类型自动命名" checked={!!s.autoName} onChecked={(v) => set("autoName", !!v)} />
        <CheckRow label="单击卡片直接插入（取消则单击预览、双击插入）" checked={s.clickToInsert !== false} onChecked={(v) => set("clickToInsert", !!v)} />
        <CheckRow label="应用样式时把主色映射到当前文档主题色" checked={!!s.themeAdaptStyle} onChecked={(v) => set("themeAdaptStyle", !!v)} />
        <CheckRow
          label="收藏形状时把文字替换为占位提示"
          checked={!!s.collectPlaceholderText}
          onChecked={(v) => set("collectPlaceholderText", !!v)}
          hint="开启后收藏的形状内文字会变成「双击编辑文字」，适合把常用版式存成模板"
        />
        <CheckRow
          label="显示「类型」行（全部/形状/样式）"
          checked={s.showTypeRow !== false}
          onChecked={(v) => set("showTypeRow", !!v)}
          hint="关闭后左侧选组即看该组全部；样式卡仍可通过右键/双击「应用到形状」使用"
        />
        <CheckRow
          label="插入整页后清理未使用的版式"
          checked={!!s.cleanLayoutsOnInsertPage}
          onChecked={(v) => set("cleanLayoutsOnInsertPage", !!v)}
          hint="整页快照会带入源文件版式；开启后自动删除目标文稿中未被使用的自定义版式（每个母版至少保留 1 个）"
        />
        <Row label="素材作者署名（导出素材包时写入）">
          <Input
            className="h-8 w-56"
            placeholder="你的昵称 / 团队名"
            value={s.authorName || ""}
            onChange={(e) => set("authorName", e.target.value)}
          />
        </Row>
        <Button className="self-end" onClick={save}>{saved ? "已保存 ✓" : "保存设置"}</Button>
      </CardContent>
    </Card>
  );
}

function Row({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="flex items-center justify-between gap-3">
      <Label className="shrink-0 text-xs text-muted-foreground">{label}</Label>
      <div className="flex min-w-0 flex-1 justify-end">{children}</div>
    </div>
  );
}

function CheckRow({
  label,
  checked,
  onChecked,
  hint,
}: {
  label: string;
  checked: boolean;
  onChecked: (v: boolean) => void;
  hint?: string;
}) {
  return (
    <div className="flex flex-col gap-0.5">
      <div className="flex items-center gap-2">
        <Checkbox id={label} checked={checked} onCheckedChange={(v) => onChecked(!!v)} />
        <Label htmlFor={label} className="text-sm">{label}</Label>
      </div>
      {hint && <p className="pl-6 text-[11px] text-muted-foreground">{hint}</p>}
    </div>
  );
}
