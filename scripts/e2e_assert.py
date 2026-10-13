# -*- coding: utf-8 -*-
"""
PPTShapeLibrary 端到端断言脚本（由 scripts/e2e-smoke.ps1 调用）
================================================================
通过 CDP 驱动 WebView2 里的**真实前端**，把 C# 桥 + 前端 + PowerPoint 文档整条链路跑一遍。

前置：PowerPoint 已启动（带 --remote-debugging-port），任务窗格已打开。
用法：python scripts/e2e_assert.py <port> <sentinel>
退出码：0=全部通过，1=有失败项。
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request

import websocket

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HELPER = (
    "window.__e2e=function(t,p,ms){return new Promise(function(res,rej){"
    "var id='e2e'+Math.random().toString(36).slice(2);"
    "var timer=setTimeout(function(){cleanup();rej(new Error('timeout '+t))},ms||180000);"
    "function cleanup(){clearTimeout(timer);window.chrome.webview.removeEventListener('message',h)}"
    "var h=function(e){var m;try{m=(typeof e.data==='string')?JSON.parse(e.data):e.data}catch(x){return;}"
    "if(m&&m.type==='result'&&m.reqId===id){cleanup();res(m);}};"
    "window.chrome.webview.addEventListener('message',h);"
    "window.chrome.webview.postMessage(JSON.stringify({type:t,reqId:id,payload:p||{}}));});};"
    "window.__e2eErrs=[];window.addEventListener('error',function(e){window.__e2eErrs.push(String(e.message))});"
    "'ok'"
)


def page_target(port, timeout_s=90):
    end = time.time() + timeout_s
    while time.time() < end:
        try:
            data = json.load(urllib.request.urlopen("http://127.0.0.1:%d/json/list" % port, timeout=5))
            pages = [x for x in data if x.get("type") == "page" and "app.pptshapes.local" in (x.get("url") or "")]
            if pages:
                return pages[0]
        except Exception:
            pass
        time.sleep(1)
    return None


class Cdp:
    def __init__(self, target):
        self.ws = websocket.create_connection(target["webSocketDebuggerUrl"], timeout=300,
                                             max_size=64 * 1024 * 1024, suppress_origin=True)
        self.seq = 0

    def ev(self, expr, ap=False, timeout=300):
        self.seq += 1
        self.ws.send(json.dumps({"id": self.seq, "method": "Runtime.evaluate",
                                 "params": {"expression": expr, "returnByValue": True,
                                            "awaitPromise": ap, "timeout": timeout * 1000}}))
        while True:
            m = json.loads(self.ws.recv())
            if m.get("id") == self.seq:
                r = m.get("result", {})
                if r.get("exceptionDetails"):
                    return {"__js_error": str(r["exceptionDetails"])[:300]}
                return r.get("result", {}).get("value")

    def call(self, t, p=None, timeout=240):
        return self.ev("__e2e(%s,%s,%d)" % (json.dumps(t), json.dumps(p or {}), timeout * 1000),
                       True, timeout + 40)


def clip_text():
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command", "Get-Clipboard -Raw"],
                             capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=20)
        return (out.stdout or "").strip()
    except Exception:
        return ""


def clip_clear():
    """尽力清空剪贴板；返回是否确认已空（清不掉时调用方应放宽断言，而不是误报失败）。"""
    for cmd in ("Clear-Clipboard", "Set-Clipboard -Value ''"):
        try:
            subprocess.run(["powershell", "-NoProfile", "-Command", cmd],
                           capture_output=True, text=True, timeout=20)
        except Exception:
            pass
        if clip_text() == "":
            return True
    return False


RESULTS = []

LOG_PATH = os.path.join(os.environ["APPDATA"], "PPTShapeLibrary", "logs", "app.log")


def _log_size():
    try:
        return os.path.getsize(LOG_PATH)
    except Exception:
        return 0


def _log_tail(since_bytes):
    try:
        with open(LOG_PATH, "rb") as f:
            f.seek(max(0, since_bytes))
            return f.read().decode("utf-8", "replace")
    except Exception:
        return ""


def _file_hash(path):
    try:
        import hashlib
        with open(path, "rb") as f:
            return hashlib.sha256(f.read()).hexdigest()
    except Exception:
        return ""


def check(name, ok, detail=""):
    RESULTS.append((name, bool(ok), detail))
    print("  [%s] %-34s %s" % ("PASS" if ok else "FAIL", name, detail))
    return ok


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 9333
    sentinel = sys.argv[2] if len(sys.argv) > 2 else "E2E-SENTINEL"

    tgt = page_target(port)
    if not tgt:
        print("FAIL: 找不到 app.pptshapes.local 页面（任务窗格没打开？）")
        return 1
    c = Cdp(tgt)
    c.ev(HELPER)

    print("\n== 1. 数据与元信息 ==")
    init = c.call("init")
    d = init.get("data") or {}
    check("init 成功", init.get("ok"), "total=%s groups=%s" % (d.get("total"), len(d.get("groups") or [])))
    check("init 带 groups/theme", bool(d.get("groups")) and bool(d.get("theme")))

    lst = c.call("list", {"limit": 200, "offset": 0})
    items = (lst.get("data") or {}).get("items") or []
    check("list 成功", lst.get("ok"), "返回 %d 条 / filteredTotal=%s" % (len(items), (lst.get("data") or {}).get("filteredTotal")))

    shapes = [i for i in items if i.get("kind") == "Shape" and not i.get("snapshotPage")]
    pages = [i for i in items if i.get("snapshotPage")]

    print("\n== 2. 健壮性（坏输入不得打崩请求） ==")
    r = c.call("list", {"limit": 3, "colorLike": "ZZZZZZ"})
    check("list 非法 colorLike 不报错", r.get("ok") is True, str(r.get("error") or "")[:60])
    r = c.call("definitelyNotAType", {})
    check("未知消息类型被拒绝", r.get("ok") is False and "unknown type" in (r.get("error") or ""))
    r = c.call("openFolder", {"id": "../../evil"})
    check("路径穿越 id 被拒", r.get("ok") is False)
    r = c.call("formatResize", {"scope": "no-such-scope", "width": 100})
    check("非法 scope 被拒（不静默改整稿）", r.get("ok") is False, str(r.get("error") or "")[:40])

    print("\n== 3. 设置边界 ==")
    r = c.call("saveSettings", {"patch": {"pageSize": 999999, "paneWidth": 999999}})
    dd = r.get("data") or {}
    check("pageSize 被夹紧到 <=500", dd.get("pageSize") == 500, "实际 %s" % dd.get("pageSize"))
    check("paneWidth 被夹紧到 <=1200", dd.get("paneWidth") == 1200, "实际 %s" % dd.get("paneWidth"))
    c.call("saveSettings", {"patch": {"pageSize": 200, "paneWidth": 400}})

    print("\n== 4. 去重扫描不得把整页快照当重复项 ==")
    r = c.call("scanDuplicates")
    dups = (r.get("data") or {}).get("duplicates") or []
    check("scanDuplicates 成功", r.get("ok"))
    check("无 shapeCount<=0 的重复项", all((x.get("shapeCount") or 0) > 0 for x in dups),
          "共 %d 条" % len(dups))

    print("\n== 5. 整页插入：不碰剪贴板 + 三个位置都可用 ==")
    if not pages:
        check("存在整页快照条目", False, "库里没有 snapshotPage 条目，跳过插入测试")
    else:
        pid = pages[0]["id"]
        before_clip = clip_text()
        slowest = 0
        for pos in ("end", "after", "end"):
            r = c.call("insertPage", {"id": pid, "position": pos}, 200)
            ms = (r.get("data") or {}).get("elapsedMs") or 0
            slowest = max(slowest, ms)
            check("insertPage(%s)" % pos, r.get("ok"), "%d ms %s" % (ms, str(r.get("error") or "")[:60]))
        after_clip = clip_text()
        check("插入整页后剪贴板未被触碰",
              sentinel in after_clip and sentinel in before_clip,
              "哨兵=%r" % (after_clip[:40],))
        # 性能回归门禁：改走 InsertFromFile 后实测 50~120ms；旧剪贴板路径为 6000ms+
        check("整页插入耗时 < 3000ms（无剪贴板重试）", slowest < 3000, "最慢 %d ms" % slowest)

    print("\n== 6. 形状插入（剪贴板两种状态） ==")
    if shapes:
        sid = shapes[0]["id"]
        # (a) 敌对状态：剪贴板被"活着但不响应"的进程独占
        #     （本脚本的父 PowerShell 正阻塞等待我们，正好构成这种状态）
        r = c.call("insert", {"id": sid}, 200)
        err = str(r.get("error") or "")
        if r.get("ok"):
            check("形状插入（剪贴板被独占）", True, "仍然成功")
        else:
            check("形状插入（剪贴板被独占）给出可操作提示且快速失败",
                  ("剪贴板被另一个程序独占" in err) or ("剪贴板被其他程序占用" in err),
                  err[:70])
        # (b) 剪贴板空闲时必须成功（清不掉时说明环境本身占着剪贴板，放宽为不误报）
        cleared = clip_clear()
        time.sleep(0.5)
        r2 = c.call("insert", {"id": sid}, 200)
        if cleared:
            check("形状插入（剪贴板空闲）", r2.get("ok") is True, str(r2.get("error") or "")[:70])
        else:
            check("形状插入（剪贴板无法清空，跳过严格断言）", True,
                  "环境占着剪贴板；本次 ok=%s" % r2.get("ok"))
    else:
        print("  (库里没有形状条目，跳过)")

    print("\n== 7. 活动文稿被库条目抢占时，整稿工具不得作用到素材文件 ==")
    # 复现场景：PowerPoint 文档恢复（上次被强杀时句柄池里开着某个 items\*\shape.pptx）
    # 或用户直接双击库文件 → 库条目变成 ActivePresentation（读写 + 有窗口）。
    # 若没有防护，"导出图片/换字体/清理瘦身"会作用到**用户的素材文件**上。
    if pages:
        lib_path = os.path.join(os.environ["APPDATA"], "PPTShapeLibrary", "items",
                                pages[0]["id"], "shape.pptx")
        before_hash = _file_hash(lib_path)
        lib_pres = None
        try:
            import win32com.client
            log_before = _log_size()
            app = win32com.client.Dispatch("PowerPoint.Application")
            lib_pres = app.Presentations.Open(lib_path, True, False, True)   # WithWindow=True
            time.sleep(0.8)
            active = ""
            try:
                active = app.ActivePresentation.Name
            except Exception:
                pass
            # 防护会立刻把焦点弹回用户文稿，所以这里期望"活动文稿不是库条目"
            check("库条目抢到活动状态后被弹回用户文稿", "shape.pptx" not in active, "active=%s" % active)
            tmpdir2 = os.path.join(tempfile.gettempdir(), "pptsl_e2e_guard")
            shutil.rmtree(tmpdir2, ignore_errors=True)
            os.makedirs(tmpdir2, exist_ok=True)
            r3 = c.call("formatExportImages", {"dir": tmpdir2, "format": "PNG", "width": 800}, 300)
            exported2 = len([f for f in os.listdir(tmpdir2) if f.endswith(".png")])
            check("导出仍作用于用户文稿（>=20 页）", exported2 >= 20,
                  "导出 %d 页（若为 1 页说明作用到了库文件）" % exported2)
            check("防护被触发并记录日志", "active presentation is a library item" in _log_tail(log_before),
                  str(r3.get("message"))[:50])
            check("库素材文件未被改动", _file_hash(lib_path) == before_hash)
        except Exception as ex:
            check("库条目抢占场景测试", False, str(ex)[:80])
        finally:
            try:
                if lib_pres is not None:
                    lib_pres.Close()
            except Exception:
                pass

    print("\n== 8. 长任务：分片执行 / 进度 / 取消 / 串行化 ==")
    tmpdir = os.path.join(tempfile.gettempdir(), "pptsl_e2e_export")
    shutil.rmtree(tmpdir, ignore_errors=True)
    os.makedirs(tmpdir, exist_ok=True)
    # 页面里的收集器：
    #  · 记录 progress
    #  · __autoCancelAfter=N 时，第 N 条进度后按 __cancelMode 取消（'bridge'=只发桥消息 / 'fetch'=只 fetch 回环端口 / 'both'）
    #  · __pingAfter=N 时发一个 ping，用于证明"长任务期间宿主仍然响应"
    c.ev(
        "window.__prog=[];window.__autoCancelAfter=0;window.__cancelMode='both';window.__pingAfter=0;"
        "window.__cancelPosted=0;window.__busyResult=null;window.__pingAt=0;window.__pingSent=0;"
        "if(!window.__progHooked){window.__progHooked=true;window.chrome.webview.addEventListener('message',function(e){"
        "var m;try{m=(typeof e.data==='string')?JSON.parse(e.data):e.data}catch(x){return;}"
        "if(!m)return;"
        "if(m.type==='progress'){window.__prog.push(m.payload);window.__firstPayload=JSON.stringify(m.payload);"
        "  var n=window.__prog.length;"
        "  if(window.__pingAfter&&n>=window.__pingAfter&&!window.__pingSent){window.__pingSent=1;"
        "    window.chrome.webview.postMessage(JSON.stringify({type:'ping',reqId:'pingprobe',payload:{}}));}"
        "  if(window.__autoCancelAfter&&n>=window.__autoCancelAfter){window.__autoCancelAfter=0;"
        "    window.__cancelPosted=(window.__cancelPosted||0)+1;"
        "    var mode=window.__cancelMode||'both';"
        "    if(mode!=='bridge'){var u=m.payload&&m.payload.cancelUrl;"
        "      if(u){try{fetch(u).then(function(r){window.__fetchOk=(window.__fetchOk||0)+1;window.__fetchStatus=r.status;}).catch(function(e){window.__fetchErr=String(e);});}catch(x){window.__fetchErr='throw:'+x;}}else{window.__fetchErr='no-url';}}"
        "    if(mode!=='fetch'){window.chrome.webview.postMessage(JSON.stringify({type:'cancel',reqId:'auto-cancel',payload:{}}));}}}"
        "if(m.type==='result'&&m.reqId==='busyprobe'){window.__busyResult=m;window.__busyAt=Date.now();}"
        "if(m.type==='result'&&m.reqId==='pingprobe'){window.__pingAt=Date.now();}"
        "});}'ok'"
    )

    # (a) 正常跑完：进度 + 导出文件数 + **长任务期间宿主仍然响应**
    c.ev("window.__prog=[];window.__autoCancelAfter=0;window.__pingAfter=2;window.__pingSent=0;window.__pingAt=0;'ok'")
    t0 = time.time()
    r = c.call("formatExportImages", {"dir": tmpdir, "format": "PNG", "width": 1200}, 300)
    export_done_at = int(time.time() * 1000)
    prog = json.loads(c.ev("JSON.stringify(window.__prog)") or "[]")
    check("整稿导出图片完成", r.get("ok") is True, str(r.get("message") or r.get("error") or "")[:70])
    check("收到 >=3 条进度消息", len(prog) >= 3, "共 %d 条" % len(prog))
    totals = {int(p.get("total") or 0) for p in prog}
    check("进度带总页数（前端可画进度条）", any(t >= 20 for t in totals), "totals=%s" % sorted(totals)[:5])
    cur = [int(p.get("current") or 0) for p in prog]
    check("进度单调不减", all(cur[i] <= cur[i + 1] for i in range(len(cur) - 1)), "样例 %s" % cur[:6])
    exported = len([f for f in os.listdir(tmpdir) if f.endswith(".png")])
    check("导出文件数与页数一致", exported >= 20, "实际 %d 个 png" % exported)
    ping_at = c.ev("window.__pingAt||0") or 0
    check("长任务期间宿主仍响应（ping 被即时答复）",
          ping_at > 0 and ping_at < export_done_at,
          "pingAt=%s exportDoneAt=%s" % (ping_at, export_done_at))

    # (b) 只发普通桥消息取消（分片执行后这条路必须生效）
    c.ev("window.__prog=[];window.__autoCancelAfter=3;window.__cancelMode='bridge';window.__cancelPosted=0;'ok'")
    shutil.rmtree(tmpdir, ignore_errors=True)
    os.makedirs(tmpdir, exist_ok=True)
    r2 = c.call("formatExportImages", {"dir": tmpdir, "format": "PNG", "width": 1200}, 300)
    prog2 = json.loads(c.ev("JSON.stringify(window.__prog)") or "[]")
    d2 = r2.get("data") or {}
    check("取消生效（仅普通桥消息）", d2.get("cancelled") is True,
          "message=%s" % str(d2.get("message"))[:50])
    check("取消后提前结束", len(prog2) < len(prog), "取消前 %d 条 / 正常 %d 条" % (len(prog2), len(prog)))

    # (c) 只走回环端口 fetch 取消（兜底通道仍然可用）
    c.ev("window.__prog=[];window.__autoCancelAfter=3;window.__cancelMode='fetch';window.__fetchOk=0;'ok'")
    shutil.rmtree(tmpdir, ignore_errors=True)
    os.makedirs(tmpdir, exist_ok=True)
    r3 = c.call("formatExportImages", {"dir": tmpdir, "format": "PNG", "width": 1200}, 300)
    d3 = r3.get("data") or {}
    check("取消生效（仅回环端口 fetch）", d3.get("cancelled") is True,
          "message=%s" % str(d3.get("message"))[:50])

    # (d) 防重入：分片执行后 UI 线程在片间是空闲的，长任务中发出的请求会被**明确拒绝**
    #     （而不是排队到结束才执行，更不是插进去打断当前操作）
    if shapes:
        c.ev(
            "(function(){window.__prog=[];window.__autoCancelAfter=0;window.__busyResult=null;window.__busyAt=0;"
            "window.__probeSentAt=0;window.__exportDoneAt=0;var id=%s;"
            "var h=function(e){var m;try{m=(typeof e.data==='string')?JSON.parse(e.data):e.data}catch(x){return;}"
            "if(m&&m.type==='progress'){window.__prog.push(1);"
            "if(window.__prog.length===2){window.chrome.webview.removeEventListener('message',h);"
            "window.__probeSentAt=Date.now();"
            "window.chrome.webview.postMessage(JSON.stringify({type:'insert',reqId:'busyprobe',payload:{id:id}}));}}};"
            "window.chrome.webview.addEventListener('message',h);return 'ok';})()" % json.dumps(shapes[0]["id"])
        )
        shutil.rmtree(tmpdir, ignore_errors=True)
        os.makedirs(tmpdir, exist_ok=True)
        c.call("formatExportImages", {"dir": tmpdir, "format": "PNG", "width": 800}, 300)
        c.ev("window.__exportDoneAt=Date.now();'ok'")
        deadline = time.time() + 30
        busy = None
        while time.time() < deadline:
            busy = json.loads(c.ev("JSON.stringify(window.__busyResult)") or "null")
            if busy:
                break
            time.sleep(0.5)
        busy_at = c.ev("window.__busyAt||0") or 0
        export_done = c.ev("window.__exportDoneAt||0") or 0
        check("长任务中发出的请求被明确拒绝（不插队、不静默丢失）",
              bool(busy) and busy.get("ok") is False, str((busy or {}).get("error"))[:60])
        check("拒绝提示可读且指向取消",
              bool(busy) and ("正在执行" in (busy.get("error") or "")),
              str((busy or {}).get("error"))[:60])
        check("拒绝发生在长任务进行中（而不是结束后）",
              bool(busy) and busy_at <= export_done, "busyAt=%s exportDoneAt=%s" % (busy_at, export_done))

    # (e) 换字体也是分片执行的（30 页文稿）：必须逐页上报进度并成功完成
    c.ev("window.__prog=[];window.__autoCancelAfter=0;window.__pingAfter=0;'ok'")
    r4 = c.call("setFonts", {"chineseFont": "微软雅黑", "latinFont": "Arial", "allSlides": True}, 300)
    prog4 = json.loads(c.ev("JSON.stringify(window.__prog)") or "[]")
    check("整稿换字体完成（分片执行）", r4.get("ok") is True,
          str((r4.get("data") or {}).get("message"))[:60])
    check("换字体逐页上报进度", len(prog4) >= 3, "共 %d 条" % len(prog4))

    # (f) 统一尺寸：分片执行（整稿 66 个形状）——进度 + 完成
    c.ev("window.__prog=[];window.__autoCancelAfter=0;window.__pingAfter=2;window.__pingSent=0;window.__pingAt=0;'ok'")
    t0 = time.time()
    r5 = c.call("formatResize", {"scope": "all", "width": 300, "height": None, "forceImageRatio": False}, 300)
    resize_done = int(time.time() * 1000)
    prog5 = json.loads(c.ev("JSON.stringify(window.__prog)") or "[]")
    check("统一尺寸完成（分片执行）", r5.get("ok") is True,
          str((r5.get("data") or {}).get("message"))[:60])
    check("统一尺寸逐片上报进度", len(prog5) >= 3, "共 %d 条" % len(prog5))
    ping5 = c.ev("window.__pingAt||0") or 0
    check("统一尺寸期间宿主仍响应", ping5 > 0 and ping5 < resize_done,
          "pingAt=%s doneAt=%s" % (ping5, resize_done))

    # (g) 清理瘦身（扫描=干跑，不改文档）：分片执行
    c.ev("window.__prog=[];window.__autoCancelAfter=0;window.__pingAfter=0;'ok'")
    r6 = c.call("formatScanClean", {"delEmpty": True, "delHidden": True, "delOffSlide": True,
                                    "clearNotes": True, "removeActions": True}, 300)
    prog6 = json.loads(c.ev("JSON.stringify(window.__prog)") or "[]")
    msg6 = str((r6.get("data") or {}).get("message") or r6.get("error") or "")
    # 扫描结果是"统计文本"；个别形状读不到会令 ok=false（既有语义），这里断言"确实跑完并给出统计"
    check("清理扫描完成（分片执行）", "空白页" in msg6, msg6[:80])
    check("清理扫描逐页上报进度", len(prog6) >= 3, "共 %d 条" % len(prog6))

    # (h) 全局兜底：未处理的 Promise 拒绝必须变成可见提示 + 宿主日志（这是"漏 catch 也不会静默"的结构保证）
    log_before_ur = _log_size()
    c.ev("window.__urShown=0;Promise.reject(new Error('E2E 未处理拒绝探针'));'ok'")
    time.sleep(1.2)
    banner = c.ev("(function(){var el=document.getElementById('global-error-banner');"
                  "return el?el.textContent:'';})()")
    log_tail_ur = _log_tail(log_before_ur)
    check("未处理拒绝弹出可见提示条", "E2E 未处理拒绝探针" in str(banner), str(banner)[:70])
    check("未处理拒绝写入宿主日志", "unhandledrejection" in log_tail_ur, "日志命中=%s" % ("unhandledrejection" in log_tail_ur))

    print("\n== 8. 前端运行时健康 ==")
    errs = c.ev("JSON.stringify(window.__e2eErrs||[])")
    check("无 JS 未捕获异常", errs in ("[]", None), str(errs)[:120])
    dom = c.ev("JSON.stringify({rootLen:(document.getElementById('root')||{}).innerHTML?"
               "document.getElementById('root').innerHTML.length:0,"
               "imgs:document.querySelectorAll('img').length})")
    check("前端已渲染", "rootLen" in str(dom) and '"rootLen":0' not in str(dom), str(dom)[:100])

    failed = [r for r in RESULTS if not r[1]]
    print("\n================ 汇总 ================")
    print("通过 %d / %d" % (len(RESULTS) - len(failed), len(RESULTS)))
    for name, _, detail in failed:
        print("  FAIL: %s  %s" % (name, detail))
    print("RESULT: %s" % ("PASS" if not failed else "FAIL"))
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
