import { useEffect, useState } from "react";
import { bridge } from "@/lib/bridge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Label } from "@/components/ui/label";
import { Separator } from "@/components/ui/separator";

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

/** 轻量字段：统一 shadcn 风格表单 */
function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-1.5">
      <Label className="text-xs text-muted-foreground">{label}</Label>
      {children}
    </div>
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
  // 拖动弹窗：复用宿主 popupDrag 协议
  useEffect(() => {
    const head = document.getElementById("popup-drag");
    if (!head) return;
    const onDown = () => {
      try {
        (window as any).chrome?.webview?.postMessage(JSON.stringify({ type: "popupDrag" }));
      } catch {}
    };
    head.addEventListener("mousedown", onDown);
    return () => head.removeEventListener("mousedown", onDown);
  }, []);

  return (
    <div className="flex h-screen flex-col bg-background text-foreground">
      {/* 标题条（无边框宿主用） */}
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
          size="icon-sm"
          className="h-6 w-6"
          aria-label="关闭"
          onClick={() => {
            try {
              (window as any).chrome?.webview?.postMessage(JSON.stringify({ type: "popupClose" }));
            } catch {}
          }}
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
  // 通用状态
  const [a, setA] = useState("");
  const [b, setB] = useState("");
  const [sel, setSel] = useState("0");
  const [out, setOut] = useState("");

  async function run(action: () => Promise<any>) {
    try {
      const d = await action();
      setOut(d?.message || "完成");
    } catch (e: any) {
      setOut("失败：" + (e?.message || e));
    }
  }

  const metaObj = META[view] || { title: "工具", desc: "" };

  if (view === "settings") {
    return <SettingsForm theme={theme} setTheme={setTheme} />;
  }

  // 各工具：仅放最核心操作的表单（渐进迁移中；消息直达 C#）
  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-sm">{metaObj.title}</CardTitle>
      </CardHeader>
      <CardContent className="flex flex-col gap-3">
        {view === "tool1" && (
          <>
            <Field label="作用范围">
              <Select value={sel} onValueChange={setSel}>
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="0">选中形状</SelectItem>
                  <SelectItem value="1">当前页</SelectItem>
                  <SelectItem value="2">全部页面</SelectItem>
                </SelectContent>
              </Select>
            </Field>
            <Field label="字体名（留空不改）"><Input value={a} onChange={(e) => setA(e.target.value)} /></Field>
            <Field label="字号（留空不改）"><Input value={b} onChange={(e) => setB(e.target.value)} /></Field>
            <Button
              onClick={() =>
                run(() =>
                  bridge.applyTextStyle({
                    scope: sel === "0" ? "selection" : sel === "1" ? "page" : "all",
                    fontName: a.trim(),
                    fontSize: parseFloat(b) || 0,
                    color: 0,
                  })
                )
              }
            >
              应用
            </Button>
          </>
        )}

        {view === "tool9" && (
          <>
            <Field label="页脚文本（姓名/单位，可留空）">
              <Input value={a} onChange={(e) => setA(e.target.value)} placeholder="张三 · 课题组" />
            </Field>
            <Button
              onClick={() =>
                run(() =>
                  bridge.academicFooter({
                    footerText: a,
                    showDate: true,
                    showPageNumber: true,
                    firstOnly: false,
                  })
                )
              }
            >
              添加到全部页面
            </Button>
          </>
        )}

        {view === "tool11" && (
          <>
            <Field label="配色方案">
              <Select value={sel} onValueChange={setSel}>
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="0">学术蓝</SelectItem>
                  <SelectItem value="1">清新绿</SelectItem>
                  <SelectItem value="2">暖橙</SelectItem>
                  <SelectItem value="3">莫兰迪</SelectItem>
                </SelectContent>
              </Select>
            </Field>
            <Button onClick={() => run(() => bridge.academicBeautifyChart({ palette: parseInt(sel), fontName: "微软雅黑" }))}>
              美化选中图表
            </Button>
          </>
        )}

        {view === "tool12" && (
          <>
            <Field label="模板类型">
              <Select value={sel} onValueChange={setSel}>
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="毕业答辩">毕业答辩</SelectItem>
                  <SelectItem value="开题报告">开题报告</SelectItem>
                  <SelectItem value="课程汇报">课程汇报</SelectItem>
                  <SelectItem value="组会汇报">组会汇报</SelectItem>
                </SelectContent>
              </Select>
            </Field>
            <Button
              onClick={() =>
                run(() =>
                  bridge.insertTemplatePack({ type: sel, contentSlides: 3 })
                )
              }
            >
              插入成套模板
            </Button>
          </>
        )}

        {!["tool1", "tool9", "tool11", "tool12"].includes(view) && (
          <p className="text-xs text-muted-foreground">
            该工具的 React 表单迁移进行中；功能仍可通过 Ribbon/旧入口使用。
            （view={view}）
          </p>
        )}

        {out && (
          <>
            <Separator />
            <p className="whitespace-pre-wrap text-xs text-muted-foreground">{out}</p>
          </>
        )}
      </CardContent>
    </Card>
  );
}

function SettingsForm({
  theme,
  setTheme,
}: {
  theme: string;
  setTheme: (t: string) => void;
}) {
  const [dark, setDark] = useState(theme === "dark");
  return (
    <Card>
      <CardHeader><CardTitle className="text-sm">偏好设置</CardTitle></CardHeader>
      <CardContent className="flex flex-col gap-3">
        <div className="flex items-center justify-between">
          <span className="text-sm">深色主题</span>
          <input
            type="checkbox"
            className="h-4 w-4 accent-[var(--primary)]"
            checked={dark}
            onChange={(e) => {
              const v = e.target.checked;
              setDark(v);
              const next = v ? "dark" : "light";
              setTheme(next);
              document.documentElement.classList.toggle("dark", v);
              bridge.saveSettings({ theme: next }).catch(() => {});
            }}
          />
        </div>
      </CardContent>
    </Card>
  );
}
