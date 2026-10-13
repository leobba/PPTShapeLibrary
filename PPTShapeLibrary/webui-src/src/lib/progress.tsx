// 长任务进度浮层：整稿收藏 / 重建缩略图 / 整稿导出图片这类操作过去会静默卡住界面，
// 现在宿主逐项推 progress 消息，这里显示进度条 + 取消按钮。
//
// 说明：宿主在长任务期间会泵消息（Application.DoEvents），所以"取消"这条上行消息
// 能被及时处理；请求结束（成功或失败）时由 bridge.onResult 精确收起浮层。
import { useEffect, useState } from "react";
import { bridge } from "@/lib/bridge";
import { Button } from "@/components/ui/button";

type Progress = { reqId: string; title: string; current: number; total: number; cancelUrl?: string };

export function ProgressHost() {
  const [state, setState] = useState<Progress | null>(null);
  const [cancelling, setCancelling] = useState(false);

  useEffect(() => {
    const offProgress = bridge.on("progress", (p) => {
      if (!p || !p.reqId) return;
      setState({
        reqId: p.reqId,
        title: p.title || "处理中",
        current: Number(p.current) || 0,
        total: Number(p.total) || 0,
        cancelUrl: p.cancelUrl || undefined,
      });
      setCancelling(false);
    });
    const offResult = bridge.onResult((reqId) => {
      setState((cur) => (cur && cur.reqId === reqId ? null : cur));
      setCancelling(false);
    });
    return () => {
      offProgress();
      offResult();
    };
  }, []);

  if (!state) return null;

  const pct = state.total > 0 ? Math.min(100, Math.round((state.current / state.total) * 100)) : 0;
  const indeterminate = state.total <= 0;

  const requestCancel = () => {
    setCancelling(true);
    // 主通道：宿主在长任务期间开了一个只监听回环的临时端口（普通桥消息此时送不进去，
    // 实测 Application.DoEvents() 也无法让 WebView2 消息在宿主 UI 线程被占用时送达）。
    if (state.cancelUrl) {
      fetch(state.cancelUrl).catch(() => {});
    }
    // 备用通道：宿主不忙时普通桥消息立刻生效
    bridge.cancel().catch(() => {});
    // 兜底：万一都没生效，别让按钮永远卡在"正在取消"
    window.setTimeout(() => setCancelling(false), 4000);
  };

  return (
    <div
      className="fixed inset-0 z-[300] flex items-center justify-center bg-black/25 backdrop-blur-[1px]"
      role="dialog"
      aria-modal="true"
      aria-label={state.title}
    >
      <div className="w-[min(360px,88vw)] rounded-xl border bg-popover p-4 shadow-2xl">
        <div className="mb-1 flex items-baseline justify-between">
          <span className="text-sm font-medium">{state.title}</span>
          <span className="text-xs tabular-nums text-muted-foreground">
            {indeterminate ? "准备中…" : `${state.current} / ${state.total}`}
          </span>
        </div>
        <div
          className="h-1.5 w-full overflow-hidden rounded-full bg-muted"
          role="progressbar"
          aria-valuemin={0}
          aria-valuemax={state.total || 100}
          aria-valuenow={indeterminate ? undefined : state.current}
          aria-valuetext={indeterminate ? "准备中" : `${state.current} / ${state.total}`}
        >
          <div
            className={
              "h-full rounded-full bg-primary transition-[width] duration-200 " +
              (indeterminate ? "w-1/3 animate-pulse" : "")
            }
            style={indeterminate ? undefined : { width: pct + "%" }}
          />
        </div>
        <div className="mt-3 flex items-center justify-end gap-2">
          <span className="mr-auto text-[11px] text-muted-foreground">
            {cancelling ? "正在取消…" : "可随时取消，已完成的部分会保留"}
          </span>
          <Button
            variant="outline"
            size="sm"
            disabled={cancelling}
            onClick={requestCancel}
          >
            取消
          </Button>
        </div>
      </div>
    </div>
  );
}
