// 全局错误兜底：把"未捕获异常 / 未处理的 Promise 拒绝"变成可见信息，而不是静默失败。
//
// 背景：桥调用里有一批写操作漏了 .catch（例如 bridge.del(...).then(refresh)），
// 宿主返回 ok=false 时 Promise 会 reject → 用户只看到"点了没反应"。
// 逐个补 catch 容易漏，这里做一层兜底：任何未处理拒绝都写进宿主日志并在顶部提示。
import { bridge } from "@/lib/bridge";

let banner: HTMLDivElement | null = null;
let hideTimer: number | undefined;

function showBanner(text: string) {
  try {
    if (!banner) {
      banner = document.createElement("div");
      banner.id = "global-error-banner";
      banner.setAttribute("role", "alert");
      banner.setAttribute("aria-live", "assertive");
      banner.style.cssText = [
        "position:fixed", "left:50%", "top:8px", "transform:translateX(-50%)",
        "z-index:2147483647", "max-width:min(92vw,520px)",
        "background:#b3261e", "color:#fff", "border-radius:8px",
        "padding:8px 12px", "font-size:12px", "line-height:1.5",
        "box-shadow:0 6px 20px rgba(0,0,0,.28)", "pointer-events:none",
        "word-break:break-word", "text-align:center",
      ].join(";");
      document.body.appendChild(banner);
    }
    banner.textContent = text;
    if (hideTimer) window.clearTimeout(hideTimer);
    hideTimer = window.setTimeout(() => {
      if (banner && banner.parentNode) banner.parentNode.removeChild(banner);
      banner = null;
    }, 6000);
  } catch {
    /* 连提示都失败就算了，至少日志里有 */
  }
}

function describe(reason: unknown): string {
  if (reason instanceof Error) return reason.message || String(reason);
  if (typeof reason === "string") return reason;
  try {
    return JSON.stringify(reason);
  } catch {
    return String(reason);
  }
}

/** 只在宿主（WebView2）里生效；浏览器里跑 npm run dev 时也不会抛。 */
export function installGlobalErrorReporting() {
  window.addEventListener("unhandledrejection", (e) => {
    const msg = describe((e as PromiseRejectionEvent).reason) || "未知错误";
    // 超时类错误已有专门的 UI 提示（插入/重建预览），不重复弹
    if (!/timeout/i.test(msg)) showBanner("操作未完成：" + msg);
    try {
      bridge.log("error", "unhandledrejection: " + msg);
    } catch {
      /* noop */
    }
  });

  window.addEventListener("error", (e) => {
    const msg = e.message || "脚本错误";
    showBanner("界面脚本错误：" + msg);
    try {
      bridge.log("error", "window.onerror: " + msg);
    } catch {
      /* noop */
    }
  });
}
