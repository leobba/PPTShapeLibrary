import { useCallback, useEffect, useRef, useState } from "react";
import { bridge } from "@/lib/bridge";
import { dialogs } from "@/lib/dialogs";

type UpdateInfo = {
  ok: boolean;
  hasUpdate: boolean;
  currentVersion: string;
  latestVersion: string;
  releaseUrl: string;
  downloadUrl: string;
  sha256: string;
  notes: string;
  error: string;
};

/**
 * 更新提示条：挂到素材库窗格顶部。
 *
 * 设计（与安装器保持同一套逻辑，但放在插件内 —— 用户在 PowerPoint 里就能看到）：
 *  · 打开面板时静默查一次（24h 缓存由 C# 侧管，不会反复联网）
 *  · 有新版本 → 顶部常驻一条，展示**更新内容**，两个动作：去下载 / 知道了
 *  · 没有 → 什么都不显示，不打扰
 *  · 查不到（断网/限流）→ 静默失败，不打扰
 */
export function UpdateBanner() {
  const [info, setInfo] = useState<UpdateInfo | null>(null);
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [showNotes, setShowNotes] = useState(false);   // 更新内容默认收起，点按钮才展开
  const tried = useRef(false);

  const check = useCallback(async (silent: boolean) => {
    setBusy(true);
    try {
      const r = await bridge.checkUpdate();
      setInfo(r as UpdateInfo);
      setOpen(true);
      // 有新版本才自动展开；没有就收起（不打扰）
      if (!(r as UpdateInfo)?.hasUpdate) setOpen(false);
      void silent;
    } catch {
      /* 检查失败就静默，不打扰用户 */
    } finally {
      setBusy(false);
    }
  }, []);

  useEffect(() => {
    if (tried.current) return;
    tried.current = true;
    void check(true);

    // 定时提醒：C# 侧到点会主动推 updateAvailable，这里收到就**弹窗提醒**
    // （不是下次打开面板才显示 —— 用户要的是"每隔一段时间提醒"）
    // 更新检查现在由 C# 侧在**插件启动时**就做（后台线程，不卡界面），
    // 结果通过 updateAvailable 事件推过来；这里只负责显示提示条。
    // 面板打开时再查一次是为了覆盖"启动时网络不通、后来通了"的情况（有 24h 缓存，开销很小）。
    const off = bridge.on("updateAvailable", (p: any) => {
      setInfo({
        ok: true,
        hasUpdate: true,
        currentVersion: p?.currentVersion || "",
        latestVersion: p?.latestVersion || "",
        releaseUrl: p?.releaseUrl || "",
        downloadUrl: p?.downloadUrl || "",
        sha256: "",
        notes: p?.notes || "",
        error: "",
      });
      setOpen(true);
      setShowNotes(false);
      // 不再弹网页里的对话框：C# 侧已经有弹窗提醒了，重复弹很烦。
      // 这里只把顶部提示条显示出来。
    });
    return off;
  }, [check]);

  if (!info || !info.ok || !info.hasUpdate || !open) return null;

  const notes = (info.notes || "").trim();
  const hasNotes = notes.length > 0 && !notes.startsWith("（");

  return (
    <div
      className="mb-2 rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-[12px] text-blue-900"
      role="status"
      aria-live="polite"
    >
      <div className="flex items-center justify-between gap-4">
        <div className="min-w-0 truncate font-medium">
          发现新版本 {info.latestVersion}
          <span className="ml-1 font-normal text-blue-700/80">（当前 {info.currentVersion}）</span>
        </div>
        <div className="flex shrink-0 items-center gap-1.5">
          <button
            className="rounded-md bg-blue-600 px-2.5 py-1 text-[11px] font-medium text-white hover:bg-blue-700 disabled:opacity-50"
            disabled={busy}
            onClick={async () => {
              // 打开**本地更新弹窗**（C# WinForms）：里面显示下载进度/速度，
              // 下载完会检查未保存的文稿并给"保存全部"按钮，确认后才卸载+安装。
              setBusy(true);
              try {
                const r = (await bridge.applyUpdate()) as any;
                setBusy(false);
                if (!r?.ok) {
                  await dialogs.alert("无法开始更新：" + (r?.message || "未知原因") + "\r\n\r\n也可点「更新内容」查看说明，或打开发布页手动下载。");
                }
              } catch (e: any) {
                setBusy(false);
                await dialogs.alert("无法开始更新：" + (e?.message || e));
              }
            }}
          >
            {busy ? "打开中…" : "立即更新"}
          </button>
          {hasNotes && (
            <button
              className="rounded-md border border-blue-300 px-2.5 py-1 text-[11px] text-blue-700 hover:bg-blue-100"
              onClick={() => setShowNotes((v) => !v)}
            >
              {showNotes ? "收起内容" : "更新内容"}
            </button>
          )}
          <button
            className="rounded-md px-2 py-1 text-[11px] text-blue-700 hover:bg-blue-100"
            onClick={() => setOpen(false)}
          >
            稍后
          </button>
        </div>
      </div>
      {showNotes && hasNotes && (
        <pre className="mt-2 max-h-56 overflow-auto whitespace-pre-wrap break-words rounded-md border border-blue-200 bg-white/70 p-3 font-sans text-[11px] leading-relaxed text-blue-900/90">
          {notes}
        </pre>
      )}
    </div>
  );
}
