import { useEffect, useState } from "react";
import { bridge } from "@/lib/bridge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { Separator } from "@/components/ui/separator";
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
  tool2: { title: "一键换字体", desc: "文档内字体整体替换" },
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
};

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-1.5">
      <Label className="text-xs text-muted-foreground">{label}</Label>
      {children}
    </div>
  );
}

/** 拉取当前演示文稿用到的字体列表（无文稿/失败时为空数组并给出提示文案） */
function useDocFonts() {
  const [fonts, setFonts] = useState<string[]>([]);
  const [noDoc, setNoDoc] = useState(false);
  const refresh = () => {
    bridge
      .toolGetFonts()
      .then((d) => {
        setFonts((d.fonts as string[]) || []);
        setNoDoc(!d.hasPresentation);
      })
      .catch(() => {
        setFonts([]);
        setNoDoc(true);
      });
  };
  useEffect(() => {
    refresh();
  }, []);
  return { fonts, noDoc, refresh };
}

/** 字体下拉选择器：文档字体 + 可输入自定义(combobox 简化: 下拉含文档字体与常用字体) */
const COMMON_FONTS = ["微软雅黑", "微软雅黑 Light", "宋体", "黑体", "楷体", "仿宋", "Arial", "Arial Black", "Calibri", "Times New Roman", "Consolas", "Segoe UI"];

/** HTML 颜色 #RRGGBB → PowerPoint RGB int(BGR 字节序) */
function toBgr(hex: string): number {
  const n = parseInt(hex.replace("#", ""), 16) || 0;
  const r = n & 0xff;
  const g = (n >> 8) & 0xff;
  const b = (n >> 16) & 0xff;
  return (b << 16) | (g << 8) | r;
}

function FontPicker({
  value,
  onChange,
  docFonts,
  placeholder,
  allowCustom = true,
}: {
  value: string;
  onChange: (v: string) => void;
  docFonts: string[];
  placeholder?: string;
  allowCustom?: boolean;
}) {
  const all = Array.from(new Set([...docFonts, ...COMMON_FONTS]));
  // 当前值若在选项中则高亮，否则视为"自定义"
  const selValue = all.includes(value) ? value : "__none__";
  return (
    <div className="flex gap-1.5">
      <Select value={selValue} onValueChange={onChange}>
        <SelectTrigger className="min-w-0 flex-1">
          <SelectValue placeholder={placeholder || "选择字体…"} />
        </SelectTrigger>
        <SelectContent className="max-h-56">
          {docFonts.length > 0 &&
            docFonts.map((f) => (
              <SelectItem key={"d" + f} value={f}>{f}</SelectItem>
            ))}
          {docFonts.length > 0 && <div className="px-2 py-1 text-[10px] text-muted-foreground">—— 常用 ——</div>}
          {COMMON_FONTS.filter((f) => !docFonts.includes(f)).map((f) => (
            <SelectItem key={"c" + f} value={f}>{f}</SelectItem>
          ))}
          <SelectItem value="__none__" className="hidden" aria-hidden>
            {value || "（自定义）"}
          </SelectItem>
        </SelectContent>
      </Select>
      {allowCustom && (
        <Input
          className="h-8 w-32"
          placeholder="自定义…"
          value={all.includes(value) ? "" : value}
          onChange={(e) => onChange(e.target.value)}
        />
      )}
    </div>
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
      <CardContent className="flex flex-col gap-3">{children}</CardContent>
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

export default function ToolView({
  view,
  theme,
  setTheme,
}: {
  view: string;
  theme: string;
  setTheme: (t: string) => void;
}) {
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
          ✕
        </Button>
      </div>
      <div className="flex-1 overflow-y-auto p-4">
        <ToolBody view={view} theme={theme} setTheme={setTheme} />
      </div>
    </div>
  );
}

function ToolBody({
  view,
  theme,
  setTheme,
}: {
  view: string;
  theme: string;
  setTheme: (t: string) => void;
}) {
  const metaObj = META[view] || { title: "工具", desc: "" };
  if (view === "settings") return <SettingsForm theme={theme} setTheme={setTheme} />;

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
  async function run(action: () => Promise<any>) {
    try {
      const d = await action();
      setOut(d?.message || "完成");
    } catch (e: any) {
      setOut("失败：" + (e?.message || e));
    }
  }
  return { out, run };
}

/** tool1 文字样式 */
/** tool1 文字样式 */
function Tool1() {
  const { out, run } = useRun();
  const { fonts, noDoc, refresh } = useDocFonts();
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
        <FontPicker value={font} onChange={setFont} docFonts={fonts} allowCustom placeholder="选择字体…" />
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
      <Button onClick={apply}>应用</Button>
      <Result out={out} />
    </Shell>
  );
}

/** tool2 一键换字体 */
function Tool2() {
  const { out, run } = useRun();
  const { fonts, noDoc, refresh } = useDocFonts();
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("微软雅黑");
  const [all, setAll] = useState(true);
  return (
    <Shell title="一键换字体" desc="把文档中的原字体替换为新字体（原字体从下方文档字体中选择）">
      {noDoc && <p className="rounded-md bg-muted px-2 py-1.5 text-xs text-muted-foreground">请先打开演示文稿，以便列出其中用到的字体。</p>}
      <Field label="原字体（要替换掉的）">
        <FontPicker value={from} onChange={setFrom} docFonts={fonts} allowCustom placeholder="从文档字体中选择…" />
      </Field>
      <Field label="新字体（替换为）">
        <FontPicker value={to} onChange={setTo} docFonts={fonts} allowCustom placeholder="选择新字体…" />
      </Field>
      <div className="flex items-center gap-2">
        <Checkbox id="t2all" checked={all} onCheckedChange={(v) => setAll(!!v)} />
        <Label htmlFor="t2all" className="text-sm">整份演示文稿（取消则当前页）</Label>
      </div>
      <Button
        disabled={!from.trim()}
        onClick={() => run(() => bridge.replaceFont({ fromFont: from.trim(), toFont: to.trim() || from.trim(), allSlides: all }))}
      >
        开始替换
      </Button>
      {!from.trim() && <p className="text-xs text-muted-foreground">请先选择要替换的原字体。</p>}
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
      <Button onClick={() => run(() => bridge.findReplace({ find, replace: rep, matchCase: match, allSlides: all }))}>替换全部</Button>
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
  const [dir, setDir] = useState("");
  const [fmt, setFmt] = useState("PNG");
  const [w, setW] = useState("0");
  return (
    <Shell title="导出图片" desc="把整份演示文稿逐页导出为图片">
      <Field label="导出目录">
        <div className="flex gap-1.5">
          <Input value={dir} onChange={(e) => setDir(e.target.value)} placeholder="选择目录…" />
          <Button variant="outline" onClick={() => bridge.requestFolder("选择导出目录").then((d) => d?.path && setDir(d.path))}>浏览…</Button>
        </div>
      </Field>
      <div className="flex items-center gap-3">
        <Field label="格式">
          <Select value={fmt} onValueChange={setFmt}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent><SelectItem value="PNG">PNG</SelectItem><SelectItem value="JPG">JPG</SelectItem></SelectContent>
          </Select>
        </Field>
        <Field label="宽度">
          <Select value={w} onValueChange={setW}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent><SelectItem value="0">原始</SelectItem><SelectItem value="1280">1280</SelectItem><SelectItem value="1920">1920</SelectItem></SelectContent>
          </Select>
        </Field>
      </div>
      <Button onClick={() => run(() => bridge.formatExportImages({ dir, format: fmt, width: parseInt(w) }))}>开始导出</Button>
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
  const [scope, setScope] = useState("selection");
  const [w, setW] = useState("");
  const [h, setH] = useState("");
  return (
    <Shell title="统一尺寸" desc="按宽/高统一形状（可只填一项，保持比例）">
      <Field label="范围">
        <Select value={scope} onValueChange={setScope}>
          <SelectTrigger><SelectValue /></SelectTrigger>
          <SelectContent>
            <SelectItem value="selection">选中形状</SelectItem>
            <SelectItem value="page">当前页</SelectItem>
            <SelectItem value="all">全部页面</SelectItem>
          </SelectContent>
        </Select>
      </Field>
      <div className="flex items-center gap-3">
        <Field label="宽度(磅)"><Input value={w} onChange={(e) => setW(e.target.value)} /></Field>
        <Field label="高度(磅)"><Input value={h} onChange={(e) => setH(e.target.value)} /></Field>
      </div>
      <Button onClick={() => run(() => bridge.formatResize({ scope, width: parseFloat(w) || 0, height: parseFloat(h) || 0, forceImageRatio: false }))}>应用尺寸</Button>
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
          <Button variant="outline" onClick={() => bridge.requestFiles("选择图片", "图片文件 (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp").then((d) => setFiles(d.files || []))}>选择…</Button>
        </div>
      </Field>
      <Field label="每行张数"><Input value={cols} onChange={(e) => setCols(e.target.value)} inputMode="numeric" /></Field>
      <div className="flex items-center gap-2">
        <Checkbox id="t8n" checked={toNew} onCheckedChange={(v) => setToNew(!!v)} />
        <Label htmlFor="t8n" className="text-sm">排满自动翻新页（取消则插入当前页）</Label>
      </div>
      <Button onClick={() => run(() => bridge.formatImageGrid({ files, cols: parseInt(cols) || 3, toNewPage: toNew }))}>生成网格</Button>
      <Result out={out} />
    </Shell>
  );
}

/** tool9 页脚页码 */
function Tool9() {
  const { out, run } = useRun();
  const [text, setText] = useState("");
  return (
    <Shell title="页脚页码" desc="为全部页添加页脚与页码（可撤销）">
      <Field label="页脚文本（姓名/单位，可留空）">
        <Input value={text} onChange={(e) => setText(e.target.value)} placeholder="张三 · 课题组" />
      </Field>
      <Button onClick={() => run(() => bridge.academicFooter({ footerText: text, showDate: true, showPageNumber: true, firstOnly: false }))}>添加到全部页面</Button>
      <Result out={out} />
    </Shell>
  );
}

/** tool10 导出 PDF / 大纲 */
function Tool10() {
  const { out, run } = useRun();
  const [mode, setMode] = useState("pdf");
  const [dir, setDir] = useState("");
  const [name, setName] = useState("");
  return (
    <Shell title="导出 PDF / 大纲" desc="整稿转 PDF 或导出文字大纲">
      <Field label="类型">
        <Select value={mode} onValueChange={setMode}>
          <SelectTrigger><SelectValue /></SelectTrigger>
          <SelectContent><SelectItem value="pdf">PDF</SelectItem><SelectItem value="outline">大纲文本 (.txt)</SelectItem></SelectContent>
        </Select>
      </Field>
      <Field label="保存目录">
        <div className="flex gap-1.5">
          <Input value={dir} onChange={(e) => setDir(e.target.value)} placeholder="选择目录…" />
          <Button variant="outline" onClick={() => bridge.requestFolder("选择保存目录").then((d) => d?.path && setDir(d.path))}>浏览…</Button>
        </div>
      </Field>
      <Field label="文件名（可留空）"><Input value={name} onChange={(e) => setName(e.target.value)} /></Field>
      <Button
        onClick={() =>
          run(() =>
            mode === "pdf"
              ? bridge.academicExportPdf({ dir, name, style: 0, perPage: 1 })
              : bridge.academicExportOutline({ dir, name })
          )
        }
      >
        导出
      </Button>
      <Result out={out} />
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

/** settings */
function SettingsForm({
  theme,
  setTheme,
}: {
  theme: string;
  setTheme: (t: string) => void;
}) {
  const [s, setS] = useState<Record<string, any> | null>(null);
  useEffect(() => {
    bridge.getSettings().then((d) => setS({ ...d })).catch(() => setS({ theme }));
  }, []);
  const set = (k: string, v: any) => setS((prev) => (prev ? { ...prev, [k]: v } : prev));
  const [saved, setSaved] = useState(false);

  if (!s) return null;

  const save = () => {
    bridge.saveSettings(s).then((d) => {
      setTheme(d.theme || s.theme);
      document.documentElement.classList.toggle("dark", (d.theme || s.theme) === "dark");
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
        <Row label="主题">
          <Select value={s.theme === "dark" ? "dark" : "light"} onValueChange={(v) => set("theme", v)}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="light">浅色</SelectItem>
              <SelectItem value="dark">深色</SelectItem>
            </SelectContent>
          </Select>
        </Row>
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
        <Row label="分页数量">
          <Input className="h-8 w-24" inputMode="numeric" value={s.pageSize ?? 200} onChange={(e) => set("pageSize", parseInt(e.target.value) || 200)} />
        </Row>
        <Row label="面板宽度">
          <Input className="h-8 w-24" inputMode="numeric" value={s.paneWidth ?? 300} onChange={(e) => set("paneWidth", parseInt(e.target.value) || 300)} />
        </Row>
        <CheckRow label="按形状类型自动命名" checked={!!s.autoName} onChecked={(v) => set("autoName", !!v)} />
        <CheckRow label="单击卡片直接插入（取消则单击预览、双击插入）" checked={s.clickToInsert !== false} onChecked={(v) => set("clickToInsert", !!v)} />
        <CheckRow label="应用样式时把主色映射到当前文档主题色" checked={!!s.themeAdaptStyle} onChecked={(v) => set("themeAdaptStyle", !!v)} />
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

function CheckRow({ label, checked, onChecked }: { label: string; checked: boolean; onChecked: (v: boolean) => void }) {
  return (
    <div className="flex items-center gap-2">
      <Checkbox id={label} checked={checked} onCheckedChange={(v) => onChecked(!!v)} />
      <Label htmlFor={label} className="text-sm">{label}</Label>
    </div>
  );
}
