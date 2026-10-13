import { useEffect, useRef, useState } from "react";
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
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";

// ---------------- 全局原生对话框替代（shadcn） ----------------
// 用法：import { dialogs, DialogsHost } from "@/lib/dialogs"
//   <DialogsHost /> 放应用根部一次；任意处 await dialogs.confirm("...") / dialogs.prompt("...") / dialogs.alert("...")

let confirmResolve: ((v: boolean) => void) | null = null;
let promptResolve: ((v: string | null) => void) | null = null;
let alertResolve: (() => void) | null = null;

type Cfg = { title?: string; okText?: string; danger?: boolean };
type Pcfg = { description?: string; initial?: string; placeholder?: string; okText?: string };
type UiState = {
  confirm: { message: string; cfg?: Cfg } | null;
  prompt: { message: string; cfg?: Pcfg } | null;
  alert: { message: string } | null;
};

function showConfirm(message: string, cfg?: Cfg) {
  return new Promise<boolean>((resolve) => {
    // 覆盖前先把上一个未决的同类型请求按"取消"结掉：全局单例 resolver 若被覆盖，
    // 前一个 await 会永远不 settle（调用方的后续逻辑静默不执行）。
    if (confirmResolve) {
      const prev = confirmResolve;
      confirmResolve = null;
      prev(false);
    }
    confirmResolve = resolve;
    _setUi?.((s) => ({ ...s, confirm: { message, cfg } }));
  });
}
function showPrompt(message: string, cfg?: Pcfg) {
  return new Promise<string | null>((resolve) => {
    if (promptResolve) {
      const prev = promptResolve;
      promptResolve = null;
      prev(null);
    }
    promptResolve = resolve;
    _setUi?.((s) => ({ ...s, prompt: { message, cfg } }));
  });
}
function showAlert(message: string) {
  return new Promise<void>((resolve) => {
    if (alertResolve) {
      const prev = alertResolve;
      alertResolve = null;
      prev();
    }
    alertResolve = resolve;
    _setUi?.((s) => ({ ...s, alert: { message } }));
  });
}

export const dialogs = { confirm: showConfirm, prompt: showPrompt, alert: showAlert };

let _setUi: React.Dispatch<React.SetStateAction<UiState>> | null = null;

/** 放一次在应用根部；内部渲染 Confirm/Prompt/Alert 三个 shadcn Dialog */
export function DialogsHost() {
  const [ui, setUi] = useState<UiState>({ confirm: null, prompt: null, alert: null });
  const inputRef = useRef<HTMLInputElement | null>(null);
  useEffect(() => {
    _setUi = setUi;
    return () => {
      _setUi = null;
    };
  }, []);

  const closeConfirm = (v: boolean) => {
    setUi((s) => ({ ...s, confirm: null }));
    const r = confirmResolve;
    confirmResolve = null;
    r?.(v);
  };
  const closePrompt = (v: string | null) => {
    setUi((s) => ({ ...s, prompt: null }));
    const r = promptResolve;
    promptResolve = null;
    r?.(v);
  };
  const closeAlert = () => {
    setUi((s) => ({ ...s, alert: null }));
    const r = alertResolve;
    alertResolve = null;
    r?.();
  };

  return (
    <>
      <AlertDialog open={!!ui.confirm} onOpenChange={(o) => !o && closeConfirm(false)}>
        {ui.confirm && (
          <AlertDialogContent>
            <AlertDialogHeader>
              <AlertDialogTitle>{ui.confirm.cfg?.title || "确认"}</AlertDialogTitle>
              {ui.confirm.message && <AlertDialogDescription>{ui.confirm.message}</AlertDialogDescription>}
            </AlertDialogHeader>
            <AlertDialogFooter>
              <AlertDialogCancel onClick={() => closeConfirm(false)}>取消</AlertDialogCancel>
              <AlertDialogAction
                className={
                  ui.confirm.cfg?.danger ? "bg-destructive text-white hover:bg-destructive/90" : undefined
                }
                onClick={() => closeConfirm(true)}
              >
                {ui.confirm.cfg?.okText || "确定"}
              </AlertDialogAction>
            </AlertDialogFooter>
          </AlertDialogContent>
        )}
      </AlertDialog>

      <Dialog open={!!ui.prompt} onOpenChange={(o) => !o && closePrompt(null)}>
        {ui.prompt && (
          <DialogContent className="max-w-sm">
            <DialogHeader>
              <DialogTitle>{ui.prompt.message}</DialogTitle>
              {ui.prompt.cfg?.description && (
                <DialogDescription>{ui.prompt.cfg.description}</DialogDescription>
              )}
            </DialogHeader>
            <Input
              ref={(el) => {
                inputRef.current = el;
              }}
              autoFocus
              defaultValue={ui.prompt.cfg?.initial}
              placeholder={ui.prompt.cfg?.placeholder}
              onKeyDown={(e) => {
                if (e.key === "Enter") closePrompt(inputRef.current?.value ?? "");
              }}
            />
            <DialogFooter className="gap-2">
              <Button variant="outline" onClick={() => closePrompt(null)}>
                取消
              </Button>
              <Button onClick={() => closePrompt(inputRef.current?.value ?? "")}>
                {ui.prompt.cfg?.okText || "确定"}
              </Button>
            </DialogFooter>
          </DialogContent>
        )}
      </Dialog>

      <Dialog open={!!ui.alert} onOpenChange={(o) => !o && closeAlert()}>
        {ui.alert && (
          <DialogContent className="max-w-sm">
            <DialogHeader>
              <DialogTitle>提示</DialogTitle>
            </DialogHeader>
            <p className="text-sm text-muted-foreground">{ui.alert.message}</p>
            <DialogFooter>
              <Button onClick={closeAlert}>知道了</Button>
            </DialogFooter>
          </DialogContent>
        )}
      </Dialog>
    </>
  );
}
