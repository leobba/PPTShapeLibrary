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
    const off = bridge.on("updateAvailable", (p: any) => {
      setInfo((cur) => ({
        ok: true,
        hasUpdate: true,
        currentVersion: p?.currentVersion || cur?.currentVersion || "",
        latestVersion: p?.latestVersion || "",
        releaseUrl: p?.releaseUrl || "",
        downloadUrl: p?.downloadUrl || "",
        sha256: "",
        notes: p?.notes || "",
        error: "",
      }));
      setOpen(true);
      const nv = p?.latestVersion || "新版本";
      const nt = (p?.notes || "").trim();
      dialogs.alert(
        "发现新版本 " + nv + "（当前 " + (p?.currentVersion || "?") + "）\r\n\r\n" + (nt ? nt + "\r\n\r\n" : "") +
        "点下方「去下载」即可获取新版。"
      );
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
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <div className="font-medium">
            发现新版本 {info.latestVersion}（当前 {info.currentVersion}）
          </div>
          {hasNotes && (
            <pre className="mt-1 max-h-24 overflow-auto whitespace-pre-wrap break-words font-sans text-[11px] leading-relaxed text-blue-800/90">
              {notes}
            </pre>
          )}
        </div>
        <div className="flex shrink-0 items-center gap-1.5">
          <button
            className="rounded-md bg-blue-600 px-2.5 py-1 text-[11px] font-medium text-white hover:bg-blue-700 disabled:opacity-50"
            disabled={busy}
            onClick={async () => {
              // 一键更新：下载+校验+解压 → 确认关闭 PowerPoint → 安装器在退出后自动装新版
              const yes = await dialogs.confirm(
                "将下载新版 " + info.latestVersion + " 并自动安装。\r\n\r\n" +
                "安装需要**关闭 PowerPoint**（新版文件要替换正在使用的旧版），" +
                "确认后会立即关闭，安装器会在后台完成升级，完成后重新打开 PowerPoint 即可。\r\n\r\n" +
                "（素材库数据不受影响）\r\n\r\n是否现在更新？",
                { title: "立即更新", okText: "更新并关闭 PowerPoint" }
              );
              if (!yes) return;
              setBusy(true);
              try {
                const r = (await bridge.applyUpdate()) as any;
                if (!r?.ok) {
                  setBusy(false);
                  await dialogs.alert("更新失败：" + (r?.message || "未知原因") + "\r\n\r\n也可点「查看更新」到发布页手动下载。");
                  return;
                }
                await bridge.quitForUpdate(r.installer, r.dir);
                // PowerPoint 即将关闭，这里不用再改 UI 状态
              } catch (e: any) {
                setBusy(false);
                await dialogs.alert("更新失败：" + (e?.message || e));
              }
            }}
          >
            {busy ? "准备中…" : "立即更新"}
          </button>
          <button
            className="rounded-md border border-blue-300 px-2.5 py-1 text-[11px] text-blue-700 hover:bg-blue-100"
            onClick={() => {
              // 打开发布页（一定存在，还能看到更新说明和历史版本），不会出现 404
              void bridge.openRelease(info.releaseUrl).catch(() => {});
            }}
          >
            查看更新
          </button>
          <button
            className="rounded-md px-2 py-1 text-[11px] text-blue-700 hover:bg-blue-100"
            onClick={() => setOpen(false)}
          >
            知道了
          </button>
        </div>
      </div>
    </div>
  );
}
