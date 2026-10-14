using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using Microsoft.Web.WebView2.Core;

namespace PPTShapeLibrary
{
    /// <summary>
    /// C#↔JS 消息路由器：协议唯一实现处。只做"协议翻译"，不含业务逻辑。
    /// </summary>
    public class BridgeController
    {
        /// <summary>设置被保存时触发（任意宿主）；ThisAddIn 订阅后即时应用 paneWidth 等需要宿主动作的项。</summary>
        public static event Action<Settings> SettingsSaved;

        /// <summary>库数据被任意宿主修改时触发（删除/改名/移动/回收站等写操作后）。各宿主订阅后刷新自己的列表，
        /// 解决"在工具弹窗里改数据、任务窗格列表停留在旧缓存不刷新"的跨宿主不同步。</summary>
        public static event Action RefreshRequested;

        private readonly CoreWebView2 _core;
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private readonly string _hostKey = Guid.NewGuid().ToString("N");   // 每个宿主(面板/弹窗)独立浏览上下文
        private volatile bool _closed = false;                             // 宿主已释放（见 MarkClosed）

        public BridgeController(CoreWebView2 core)
        {
            _core = core;
        }

        public void Attach()
        {
            // 首屏数据在页面 ready 后由前端发 init 拉取
        }

        /// <summary>WebView 消息仅接受本地受信任 UI 页面，避免外部导航页面调用宿主桥。</summary>
        public static bool IsTrustedMessageSource(string source)
        {
            Uri uri;
            return Uri.TryCreate(source, UriKind.Absolute, out uri)
                && string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                && string.Equals(uri.Host, "app.pptshapes.local", StringComparison.OrdinalIgnoreCase)
                && uri.IsDefaultPort;
        }

        // ---------------- 长任务（进度 / 取消 / 分片执行） ----------------

        private static LongOperation _current;   // 同一时刻只允许一个长任务

        /// <summary>把回调投递回 UI 线程的消息队列（片与片之间回到消息循环）。
        /// 用 SynchronizationContext：WebView2/WinForms 控件创建时已在本线程安装，
        /// 没有时退化为一次性 WinForms 计时器（同样依赖消息循环）。</summary>
        private static void PostToUi(Action action)
        {
            PostToUiStatic(action);
        }

        /// <summary>供其它类（如 UpdateDialog）投递到 UI 队列用。</summary>
        internal static void PostToUiStatic(Action action)
        {
            var ctx = System.Threading.SynchronizationContext.Current;
            if (ctx != null)
            {
                ctx.Post(_ => { try { action(); } catch (Exception ex) { Log.Error(ex, "PostToUi"); } }, null);
                return;
            }
            var timer = new System.Windows.Forms.Timer { Interval = 1 };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                try { action(); } catch (Exception ex) { Log.Error(ex, "PostToUi(timer)"); }
            };
            timer.Start();
        }

        /// <summary>
        /// 分片执行长任务：每片在 UI 线程上跑一小步，**片间回到消息循环**。
        /// 效果：PowerPoint 不再被整段占死（能重绘/响应点击），桥消息（含 cancel、ping）
        /// 也能正常送达——这比"泵 DoEvents"可靠（实测 DoEvents 送不进 WebView2 消息）。
        ///
        /// 协议不变：本方法**不立即**下发 result，全部完成后异步发出，前端仍按 reqId 等待。
        /// 调用方在 switch 里直接 break，不要自己再发 result。
        /// </summary>
        private void RunChunked(string reqId, string title, int total, Action<int> slice,
                                Func<bool, TextToolService.ToolResult> finish)
        {
            var op = new LongOperation(reqId, title, PushProgress);
            _current = op;
            op.OpenCancelChannel();   // 兼容前端"fetch 取消"通道（分片后普通消息也生效，双保险）
            op.Report(0, total);
            Log.Write("[bridge] chunked op begin: " + title + " total=" + total + " req=" + reqId);

            int i = 0;
            Action step = null;
            step = () =>
            {
                if (op.Cancelled || i >= total)
                {
                    FinishChunked(op, reqId, finish, op.Cancelled);
                    return;
                }
                try { slice(i); }
                catch (Exception ex) { Log.Error(ex, "chunk " + title + " #" + i); }
                i++;
                op.Report(i, total);
                if (i < total && !op.Cancelled) PostToUi(step);
                else FinishChunked(op, reqId, finish, op.Cancelled);
            };
            step();
        }

        private void FinishChunked(LongOperation op, string reqId,
                                   Func<bool, TextToolService.ToolResult> finish, bool cancelled)
        {
            op.CloseCancelChannel();
            if (ReferenceEquals(_current, op)) _current = null;
            TextToolService.ToolResult r = null;
            try { r = finish(cancelled); }
            catch (Exception ex) { Log.Error(ex, "chunked finish"); }
            Log.Write("[bridge] chunked op end: " + op.Title + " cancelled=" + cancelled);
            if (r == null)
            {
                Result(reqId, false, null, "操作失败。");
                return;
            }
            Result(reqId, r.Ok,
                   r.Ok ? new Dictionary<string, object>
                       {
                           { "changed", r.Changed },
                           { "message", r.Message },
                           { "cancelled", cancelled }
                       } : null,
                   r.Ok ? null : r.Message);
        }

        private LongOperation BeginLongOperation(string reqId, string title)
        {
            var op = new LongOperation(reqId, title, PushProgress);
            _current = op;
            op.RegisterComFilter();
            op.OpenCancelChannel();   // 取消通道（只监听回环，任务结束即关）
            op.Report(0, 0);
            Log.Write("[bridge] long op begin: " + title + " req=" + reqId);
            return op;
        }

        private void EndLongOperation(LongOperation op)
        {
            if (op == null) return;
            op.RevokeComFilter();
            op.CloseCancelChannel();
            if (ReferenceEquals(_current, op)) _current = null;
            Log.Write("[bridge] long op end: " + op.Title + " cancelled=" + op.Cancelled);
        }

        private void PushProgress(string reqId, int current, int total)
        {
            var cur = _current;
            var payload = new Dictionary<string, object>
            {
                { "reqId", reqId },
                { "current", current },
                { "total", total },
                { "title", cur != null ? cur.Title : "" }
            };
            // 取消地址随进度下发：前端据此 fetch 回环端口，绕开被占住的 UI 消息循环
            if (cur != null && !string.IsNullOrEmpty(cur.CancelUrl)) payload["cancelUrl"] = cur.CancelUrl;
            Post(new Dictionary<string, object> { { "type", "progress" }, { "payload", payload } });
        }

        // ---------------- 下行推送 ----------------

        public void PushRefresh()
        {
            Post(new Dictionary<string, object> { { "type", "refresh" }, { "payload", new object[0] } });
        }

        /// <summary>写操作后广播刷新：刷新自身宿主并通知其他宿主（跨弹窗/窗格数据同步）。</summary>
        public void BroadcastRefresh()
        {
            PushRefresh();
            try { RefreshRequested?.Invoke(); } catch { }
        }

        public void PushSettingsChanged(Settings s)
        {
            Post(new Dictionary<string, object> { { "type", "settingsChanged" }, { "payload", SettingsDict(s) } });
        }

        // ---------------- 上行处理 ----------------

        public void HandleWebMessage(string raw)
        {
            string type = "?";
            string reqId = "";
            try
            {
                var msg = _json.Deserialize<Dictionary<string, object>>(raw);
                if (msg == null) return;
                type = ToStr(msg, "type");
                reqId = ToStr(msg, "reqId");
                var payload = msg.ContainsKey("payload") ? msg["payload"] as Dictionary<string, object> : new Dictionary<string, object>();
                if (payload == null) payload = new Dictionary<string, object>();

                // 长任务执行期间只放行"取消/探活/日志"：DoEvents 会把队列里的其他请求送进来，
                // 若不拦住就会重入（在整稿操作中途又发起一次插入/删除）。
                var running = _current;
                if (running != null && type != "cancel" && type != "ping" && type != "log")
                {
                    Log.Write("[bridge] rejected while busy: " + type);
                    Result(reqId, false, null,
                        "正在执行「" + running.Title + "」，请等待完成或点击「取消」。");
                    return;
                }

                switch (type)
                {
                    case "ping":
                        Ok(reqId, new Dictionary<string, object> { { "pong", true } });
                        break;
                    case "cancel":
                        {
                            // 长任务里的 DoEvents 会让这条消息有机会被处理，从而把 Cancelled 置位
                            var cur = _current;
                            if (cur != null)
                            {
                                cur.Cancelled = true;
                                Log.Write("[bridge] cancel requested for " + cur.Title + " req=" + reqId);
                                Ok(reqId, new Dictionary<string, object> { { "cancelling", true } });
                            }
                            else
                            {
                                Ok(reqId, new Dictionary<string, object> { { "cancelling", false } });
                            }
                            break;
                        }
                    case "init":
                        Ok(reqId, BuildInit());
                        break;
                    // 检查更新：给前端用（插件内直接提示升级，不必让用户跑去开安装器）。
                    // 通道与安装器一致：① 固定地址 update.json → ② GitHub API 兜底 → ③ 加速站回退。
                    case "checkUpdate":
                        {
                            var u = UpdateService.Check();
                            var d = new Dictionary<string, object>
                            {
                                { "ok", u.Ok },
                                { "hasUpdate", u.HasUpdate },
                                { "currentVersion", u.CurrentVersion },
                                { "latestVersion", u.LatestVersion },
                                { "releaseUrl", u.ReleaseUrl },
                                { "downloadUrl", u.DownloadUrl },
                                { "sha256", u.Sha256 },
                                { "notes", u.Notes },
                                { "error", u.Error },
                            };
                            Ok(reqId, d);
                        }
                        break;
                    case "openRelease":
                        {
                            string url = ToStr(payload, "url");
                            // 一律以**发布页**为准：清单里的直连地址可能失效（旧版本/资产被删）→ 404；
                            // 而发布页永远存在，还能看到更新说明与历史版本。
                            // （早期版本会先 HEAD 探测直连包，但实测在复杂网络下容易误判成不可用，故放弃探测。）
                            if (string.IsNullOrEmpty(url) || url.Contains("/releases/download/"))
                            {
                                url = UpdateService.ReleasesPageUrl;
                            }
                            try
                            {
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true });
                                Ok(reqId, new Dictionary<string, object> { { "opened", true } });
                            }
                            catch (Exception ex)
                            {
                                Result(reqId, false, null, ex.Message);
                            }
                            break;
                        }
                    // 定时提醒：前端可设提醒间隔（分钟，0=关闭）。C# 侧起定时器，
                    // 到点检查一次，有新版就**主动推** updateAvailable 给前端，不等用户下次手点。
                    // 一键更新：打开**本地更新弹窗**（下载进度/速度 → 检查未保存文稿 → 保存按钮 → 卸载+安装）
                    // 检查要联网 → 丢后台线程，绝不卡住 PowerPoint 界面；有结果再回 UI 线程开窗。
                    case "applyUpdate":
                        {
                            Ok(reqId, new Dictionary<string, object>
                            {
                                { "ok", true },
                                { "message", "正在检查更新…" },
                            });
                            UiDispatcher.RunBackground(() => UpdateReminder.CheckCore(prompt: true), null);
                            break;
                        }
                    // 关掉 PowerPoint 并让安装器完成更新（用户在弹窗里点"立即更新"后走这里）
                    case "quitForUpdate":
                        {
                            string installer = ToStr(payload, "installer");
                            string dir = ToStr(payload, "dir");
                            if (string.IsNullOrEmpty(installer) || !File.Exists(installer))
                            {
                                Result(reqId, false, null, "安装器不存在，无法自动更新");
                                break;
                            }
                            try
                            {
                                // 先起安装器（它会等 PowerPoint 退出），再关 PowerPoint
                                var psi = new System.Diagnostics.ProcessStartInfo(installer, "--pending-update \"" + dir + "\"");
                                psi.UseShellExecute = true;
                                psi.WorkingDirectory = Path.GetDirectoryName(installer);
                                System.Diagnostics.Process.Start(psi);
                                Log.Write("[update] 已启动安装器：" + installer + " --pending-update " + dir);
                                Ok(reqId, new Dictionary<string, object> { { "started", true } });
                                // 关 PowerPoint（延后一点，让安装器进程先起来）
                                ScheduleQuitPowerPoint();
                            }
                            catch (Exception ex)
                            {
                                Result(reqId, false, null, ex.Message);
                            }
                            break;
                        }
                    case "setUpdateReminder":
                        {
                            int minutes = 0;
                            int.TryParse(ToStr(payload, "minutes"), out minutes);
                            UpdateReminder.SetInterval(minutes);
                            Ok(reqId, new Dictionary<string, object> { { "minutes", UpdateReminder.Minutes } });
                            break;
                        }
                    case "list":
                        Ok(reqId, DoList(payload));
                        break;
                    case "getSettings":
                        Ok(reqId, SettingsDict(ShapeLibraryService.LoadSettings()));
                        break;
                    case "saveSettings":
                        Ok(reqId, DoSaveSettings(payload));
                        break;
                    case "insert":
                        {
                            string id = ToStr(payload, "id");
                            var sw = System.Diagnostics.Stopwatch.StartNew();
                            bool ok = ShapeExporter.InsertShape(id, false, out string insErr);   // 静默，前端 Toast；InsertShape 内部已 TouchUsed，此处不再重复计
                            sw.Stop();
                            if (sw.ElapsedMilliseconds > 1000)
                                Log.Write("insert slow: " + sw.ElapsedMilliseconds + "ms id=" + id + " ok=" + ok);
                            Result(reqId, ok, ok ? new Dictionary<string, object> { { "kind", KindOf(id) }, { "elapsedMs", sw.ElapsedMilliseconds } } : null,
                                   ok ? null : insErr);
                            break;
                        }
                    case "insertPage":
                        {
                            string id = ToStr(payload, "id");
                            string position = ToStr(payload, "position");
                            var swp = System.Diagnostics.Stopwatch.StartNew();
                            bool ok = ShapeExporter.InsertPageSnapshot(id, position, out string pgErr);
                            swp.Stop();
                            if (swp.ElapsedMilliseconds > 1000)
                                Log.Write("insertPage slow: " + swp.ElapsedMilliseconds + "ms id=" + id + " ok=" + ok);
                            Result(reqId, ok, ok ? new Dictionary<string, object> { { "kind", "Page" }, { "elapsedMs", swp.ElapsedMilliseconds } } : null,
                                   ok ? null : pgErr);
                            break;
                        }
                    case "applyStyle":
                        {
                            string id = ToStr(payload, "id");
                            bool ok = ShapeExporter.ApplyStyle(id, false, out string styErr);   // 静默
                            if (ok) ShapeLibraryService.TouchUsed(id);
                            Result(reqId, ok, null, ok ? null : styErr);
                            break;
                        }
                    case "applyStyleSameType":
                        {
                            string id = ToStr(payload, "id");
                            bool ok = ShapeExporter.ApplyStyleToSameType(id, false, out string batErr);   // 静默
                            if (ok) ShapeLibraryService.TouchUsed(id);
                            Result(reqId, ok, null, ok ? null : batErr);
                            break;
                        }
                    case "applyStyleSameTypeAllSlides":
                        {
                            string id = ToStr(payload, "id");
                            bool ok = ShapeExporter.ApplyStyleToSameTypeAllSlides(id, false, out string allErr);   // 静默
                            if (ok) ShapeLibraryService.TouchUsed(id);
                            Result(reqId, ok, null, ok ? null : allErr);
                            break;
                        }
                    case "collectClipboard":
                        {
                            // Ribbon 直达能力；Web 侧仅转发（供“更多工具”入口复用）
                            string nid = ShapeExporter.CollectClipboard();
                            if (nid != null) BroadcastRefresh();
                            Ok(reqId, nid != null
                                ? new Dictionary<string, object> { { "id", nid } }
                                : new Dictionary<string, object>());
                            break;
                        }
                    case "collectAllSlides":
                        {
                            // 分片执行：按页切片（每页会产生一个库条目），片间回到消息循环
                            string caErr;
                            var job = ShapeExporter.CollectAllSlidesJob.Create(out caErr);
                            if (job == null)
                            {
                                Result(reqId, false, null, caErr);
                                break;
                            }
                            RunChunked(reqId, "收藏整稿", job.Total,
                                       idx => job.Step(idx),
                                       cancelled =>
                                       {
                                           var res = job.Finish(cancelled);
                                           if (job.Collected > 0) BroadcastRefresh();
                                           return res;
                                       });
                            break;
                        }
                    case "undo":
                        ShapeExporter.UndoLastAction();
                        Ok(reqId, new Dictionary<string, object>());
                        break;
                    case "redo":
                        ShapeExporter.RedoLastAction();
                        Ok(reqId, new Dictionary<string, object>());
                        break;
                    case "adjust":
                        ShapeExporter.AdjustSelected(ToStr(payload, "action"));
                        Ok(reqId, new Dictionary<string, object>());
                        break;
                    case "delete":
                        {
                            int n = 0;
                            int failed = 0;
                            foreach (var id in GetIds(payload))
                            {
                                if (ShapeLibraryService.DeleteItem(id)) n++;
                                else failed++;
                            }
                            // 部分成功也必须广播：否则其他宿主（工具弹窗等）仍显示已删除的条目
                            if (n > 0) BroadcastRefresh();
                            if (failed > 0)
                                Result(reqId, false,
                                    new Dictionary<string, object> { { "deleted", n }, { "failed", failed } },
                                    "已移入回收站 " + n + " 项；有 " + failed + " 项失败（可能被占用或文件锁定），已保留在原位。");
                            else
                                Ok(reqId, new Dictionary<string, object> { { "deleted", n } });
                            break;
                        }
                    case "rename":
                        {
                            string id = ToStr(payload, "id");
                            ShapeLibraryService.RenameItem(id, ToStr(payload, "name"));
                            Ok(reqId, new Dictionary<string, object>());
                            break;
                        }
                    case "setCategory":
                        {
                            string id = ToStr(payload, "id");
                            ShapeLibraryService.SetCategory(id, ToStr(payload, "category"));
                            Ok(reqId, new Dictionary<string, object>());
                            break;
                        }
                    case "batchMove":
                        {
                            int n = 0;
                            string cat = ToStr(payload, "category");
                            foreach (var id in GetIds(payload))
                            {
                                ShapeLibraryService.SetCategory(id, cat);
                                n++;
                            }
                            Ok(reqId, new Dictionary<string, object> { { "moved", n } });
                            break;
                        }
                    // ---------------- 一级分组管理 ----------------
                    case "setGroup":
                        {
                            string id = ToStr(payload, "id");
                            ShapeLibraryService.SetGroup(id, ToStr(payload, "group"));
                            Ok(reqId, new Dictionary<string, object>());
                            break;
                        }
                    case "batchMoveGroup":
                        {
                            int n = 0;
                            string g = ToStr(payload, "group");
                            foreach (var id in GetIds(payload))
                            {
                                ShapeLibraryService.SetGroup(id, g);
                                n++;
                            }
                            Ok(reqId, new Dictionary<string, object> { { "moved", n } });
                            break;
                        }
                    case "createGroup":
                        {
                            bool ok = ShapeLibraryService.CreateUserGroup(ToStr(payload, "name"));
                            Ok(reqId, new Dictionary<string, object> { { "ok", ok } });
                            break;
                        }
                    case "renameGroup":
                        {
                            ShapeLibraryService.RenameGroup(ToStr(payload, "oldName"), ToStr(payload, "newName"));
                            Ok(reqId, new Dictionary<string, object>());
                            break;
                        }
                    case "deleteGroup":
                        {
                            int n = ShapeLibraryService.DeleteGroup(ToStr(payload, "name"));
                            Ok(reqId, new Dictionary<string, object> { { "moved", n } });
                            break;
                        }
                    case "reorder":
                        {
                            var ordered = payload.ContainsKey("orderedIds") ? payload["orderedIds"] as System.Collections.ArrayList : null;
                            var ids = new List<string>();
                            if (ordered != null)
                            {
                                for (int i = 0; i < ordered.Count; i++) ids.Add(ordered[i].ToString());
                            }
                            // 由服务层按"全局顺序回填 + 全量重新编号"落盘，避免只给子集写序号导致全局顺序错乱
                            int renumbered = ShapeLibraryService.ReorderItems(ids);
                            Ok(reqId, new Dictionary<string, object> { { "renumbered", renumbered } });
                            break;
                        }
                    case "openFolder":
                        {
                            string id = ToStr(payload, "id");
                            if (!ShapeLibraryService.IsSafeId(id))
                            {
                                Result(reqId, false, null, "素材条目标识无效");
                                break;
                            }
                            string dir = Path.Combine(ShapeLibraryService.ItemsDir, id);
                            if (!Directory.Exists(dir))
                            {
                                Result(reqId, false, null, "素材条目不存在");
                                break;
                            }
                            System.Diagnostics.Process.Start("explorer.exe", dir);
                            Ok(reqId, new Dictionary<string, object>());
                            break;
                        }
                    // ---------------- 回收站 ----------------
                    case "trashList":
                        {
                            var items = ShapeLibraryService.ListTrash();
                            var arr = new List<Dictionary<string, object>>();
                            foreach (var m in items)
                            {
                                arr.Add(new Dictionary<string, object>
                                {
                                    { "id", m.Id },
                                    { "name", m.Name },
                                    { "kind", m.Kind },
                                    { "category", m.GetCategory() },
                                    { "createdAt", m.CreatedAt }
                                });
                            }
                            Ok(reqId, new Dictionary<string, object> { { "items", arr } });
                            break;
                        }
                    case "trashRestore":
                        {
                            bool ok = ShapeLibraryService.RestoreTrashItem(ToStr(payload, "id"));
                            Result(reqId, ok, ok ? new Dictionary<string, object>() : null,
                                   ok ? null : "恢复失败（条目不存在或已被清除）");
                            BroadcastRefresh();
                            break;
                        }
                    case "trashPurge":
                        {
                            int n = ShapeLibraryService.PurgeTrash();
                            Ok(reqId, new Dictionary<string, object> { { "purged", n } });
                            break;
                        }
                    case "trashPurgeOne":
                        {
                            ShapeLibraryService.HardDelete(ToStr(payload, "id"));
                            Ok(reqId, new Dictionary<string, object>());
                            break;
                        }
                    case "rebuildPreviews":
                        {
                            // 分片执行：逐条重建（每条要打开一次库文稿），片间回到消息循环
                            string rpErr;
                            var job = ShapeExporter.RebuildPreviewsJob.Create(out rpErr);
                            if (job == null)
                            {
                                Result(reqId, false, null, rpErr);
                                break;
                            }
                            RunChunked(reqId, "重建缩略图", job.Total,
                                       idx => job.Step(idx),
                                       cancelled =>
                                       {
                                           var res = job.Finish(cancelled);
                                           BroadcastRefresh();
                                           return res;
                                       });
                            break;
                        }
                    case "setPaneDock":
                        {
                            Globals.ThisAddIn.SetPaneDock(ToStr(payload, "dock"));
                            Ok(reqId, new Dictionary<string, object>());
                            BroadcastRefresh();
                            break;
                        }
                    case "getToolPrefs":
                        {
                            var prefs = ShapeLibraryService.GetToolPrefs();
                            Ok(reqId, new Dictionary<string, object> { { "prefs", prefs } });
                            break;
                        }
                    case "setToolPref":
                        {
                            ShapeLibraryService.SetToolPref(ToStr(payload, "key"), ToStr(payload, "value"));
                            Ok(reqId, new Dictionary<string, object>());
                            break;
                        }
                    // ---------------- PPT 工具（Web 工具页） ----------------
                    case "toolGetFonts":
                        {
                            // 系统字体：GDI+ 已安装字体 + 用户级字体（见 FontCatalog）
                            var sysFonts = FontCatalog.GetFamilies();
                            Ok(reqId, new Dictionary<string, object>
                            {
                                { "fonts", TextToolService.GetDocumentFonts() },
                                { "systemFonts", sysFonts },
                                { "hasPresentation", TextToolService.HasPresentation() }
                            });
                            break;
                        }
                    case "applyTextStyle":
                        {
                            var r = TextToolService.ApplyTextStyle(
                                ToStr(payload, "scope"),
                                ToStr(payload, "fontName"),
                                ToFloatNullable(payload, "fontSize"),
                                ToBoolNullable(payload, "bold"),
                                ToBoolNullable(payload, "italic"),
                                ToIntNullable(payload, "color"),
                                ToStr(payload, "align"));
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    case "setFonts":
                        {
                            // 分片执行：按页切片，片间回到消息循环（整稿换字体不再卡死 PowerPoint）
                            bool allSlides = ToBoolNullable(payload, "allSlides") ?? true;   // 默认整稿
                            string sfErr;
                            var job = TextToolService.SetFontsJob.Create(
                                ToStr(payload, "chineseFont"), ToStr(payload, "latinFont"), allSlides, out sfErr);
                            if (job == null)
                            {
                                Result(reqId, false, null, sfErr);
                                break;
                            }
                            RunChunked(reqId, "一键换字体", job.Total,
                                       idx => job.Step(idx),
                                       cancelled => job.Finish(cancelled));
                            break;
                        }
                    case "findReplace":
                        {
                            var r = TextToolService.FindReplace(
                                ToStr(payload, "find"),
                                ToStr(payload, "replace"),
                                true == ToBoolNullable(payload, "matchCase"),
                                true == ToBoolNullable(payload, "allSlides"));
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    case "cleanText":
                        {
                            var r = TextToolService.CleanText(
                                true == ToBoolNullable(payload, "trimWhitespace"),
                                true == ToBoolNullable(payload, "unifyPunctuation"));
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    // ---------------- 版式批处理（FormatToolService） ----------------
                    case "requestFolder":
                        {
                            using (var dlg = new System.Windows.Forms.FolderBrowserDialog
                            {
                                Description = string.IsNullOrEmpty(ToStr(payload, "title")) ? "选择文件夹" : ToStr(payload, "title")
                            })
                            {
                                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                {
                                    Ok(reqId, new Dictionary<string, object> { { "path", dlg.SelectedPath } });
                                }
                                else
                                {
                                    Ok(reqId, new Dictionary<string, object> { { "path", "" } });
                                }
                            }
                            break;
                        }
                    case "requestFiles":
                        {
                            using (var dlg = new System.Windows.Forms.OpenFileDialog
                            {
                                Title = string.IsNullOrEmpty(ToStr(payload, "title")) ? "选择文件" : ToStr(payload, "title"),
                                Multiselect = true,
                                Filter = string.IsNullOrEmpty(ToStr(payload, "filter"))
                                    ? "所有文件 (*.*)|*.*" : ToStr(payload, "filter")
                            })
                            {
                                var files = new System.Collections.ArrayList();
                                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                {
                                    foreach (var f in dlg.FileNames) files.Add(f);
                                }
                                Ok(reqId, new Dictionary<string, object> { { "files", files } });
                            }
                            break;
                        }
                    case "importPackFile":
                        {
                            // 直接弹文件选择器选 .ppack.zip → 导入素材库（点对点分享一键导入）
                            using (var dlg = new System.Windows.Forms.OpenFileDialog
                            {
                                Title = "选择要导入的素材包",
                                Multiselect = true,
                                Filter = "素材包 (*.ppack.zip)|*.ppack.zip|所有文件 (*.*)|*.*"
                            })
                            {
                                int total = 0, conflicts = 0;
                                bool hadSelection = false;
                                string lastAuthor = "", lastLicense = "", lastSource = "", lastName = "";
                                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                {
                                    hadSelection = true;
                                    foreach (var f in dlg.FileNames)
                                    {
                                        try
                                        {
                                            int n = PackService.ImportPack(f, out int cf, out string err, out var meta);
                                            if (err == null)
                                            {
                                                total += n;
                                                conflicts += cf;
                                                if (meta != null)
                                                {
                                                    if (!string.IsNullOrEmpty(meta.Author)) lastAuthor = meta.Author;
                                                    if (!string.IsNullOrEmpty(meta.License)) lastLicense = meta.License;
                                                    if (!string.IsNullOrEmpty(meta.SourceUrl)) lastSource = meta.SourceUrl;
                                                    if (!string.IsNullOrEmpty(meta.Name)) lastName = meta.Name;
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                var res = new Dictionary<string, object> { { "imported", total }, { "conflicts", conflicts }, { "hadSelection", hadSelection } };
                                if (!string.IsNullOrEmpty(lastName)) res["packName"] = lastName;
                                if (!string.IsNullOrEmpty(lastAuthor)) res["author"] = lastAuthor;
                                if (!string.IsNullOrEmpty(lastLicense)) res["license"] = lastLicense;
                                if (!string.IsNullOrEmpty(lastSource)) res["sourceUrl"] = lastSource;
                                Ok(reqId, res);
                            }
                            BroadcastRefresh();
                            break;
                        }
                    case "importPictures":
                        {
                            using (var dlg = new System.Windows.Forms.OpenFileDialog
                            {
                                Title = "选择要导入的图片",
                                Multiselect = true,
                                Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp"
                            })
                            {
                                int n = 0;
                                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                {
                                    foreach (var f in dlg.FileNames)
                                    {
                                        try
                                        {
                                            if (ShapeExporter.ImportPictureToLibrary(f) != null) n++;
                                        }
                                        catch { }
                                    }
                                }
                                Ok(reqId, new Dictionary<string, object> { { "imported", n } });
                            }
                            BroadcastRefresh();
                            break;
                        }
                    case "formatExportImages":
                        {
                            // 分片执行：逐页导出，片间回到消息循环（PowerPoint 保持响应、取消即时生效）
                            string exErr;
                            var job = FormatToolService.ExportImagesJob.Create(
                                ToStr(payload, "dir"), ToStr(payload, "format"), ToInt(payload, "width"), out exErr);
                            if (job == null)
                            {
                                Result(reqId, false, null, exErr);
                                break;
                            }
                            RunChunked(reqId, "导出图片", job.Total,
                                       idx => job.Step(idx),
                                       cancelled => job.Finish(cancelled));
                            break;
                        }
                    case "formatScanClean":
                    case "formatRunClean":
                        {
                            // 分片执行：按页扫描/清理，片间回到消息循环（大文档不再整段卡死）
                            bool dryRun = type == "formatScanClean";
                            var opts = CleanOptsFrom(payload);
                            string cErr;
                            var job = FormatToolService.CleanJob.Create(opts, dryRun, out cErr);
                            if (job == null)
                            {
                                Result(reqId, false, null, cErr);
                                break;
                            }
                            RunChunked(reqId, dryRun ? "扫描清理项" : "清理瘦身", job.Total,
                                       idx => job.Step(idx),
                                       cancelled => job.Finish(cancelled));
                            break;
                        }
                    case "formatResize":
                        {
                            // 分片执行：逐个形状改尺寸，片间回到消息循环
                            string rsErr;
                            var job = FormatToolService.ResizeShapesJob.Create(
                                ToStr(payload, "scope"),
                                ToFloatNullable(payload, "width"),
                                ToFloatNullable(payload, "height"),
                                true == ToBoolNullable(payload, "forceImageRatio"),
                                out rsErr);
                            if (job == null)
                            {
                                Result(reqId, false, null, rsErr);
                                break;
                            }
                            RunChunked(reqId, "统一尺寸", job.Total,
                                       idx => job.Step(idx),
                                       cancelled => job.Finish(cancelled));
                            break;
                        }
                    case "formatImageGrid":
                        {
                            var r = FormatToolService.ImageGrid(
                                GetStringArray(payload, "files"),
                                Math.Max(1, Math.Min(8, ToInt(payload, "cols"))),
                                true == ToBoolNullable(payload, "toNewPage"));
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    // ---------------- 学术/学习工具（AcademicToolService） ----------------
                    case "academicFooter":
                        {
                            var r = AcademicToolService.AddFooterPageNumbers(
                                ToStr(payload, "footerText"),
                                true == ToBoolNullable(payload, "showDate"),
                                true == ToBoolNullable(payload, "showPageNumber"),
                                true == ToBoolNullable(payload, "firstOnly"));
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    case "academicExportPdf":
                        {
                            var r = AcademicToolService.ExportPdf(
                                ToStr(payload, "dir"),
                                ToStr(payload, "name"),
                                ToInt(payload, "style"),
                                Math.Max(1, Math.Min(9, ToInt(payload, "perPage"))));
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    case "academicExportOutline":
                        {
                            var r = AcademicToolService.ExportOutlineText(
                                ToStr(payload, "dir"),
                                ToStr(payload, "name"));
                            // Ok 时 dir 为空则 Message 是大纲文本（供前端复制）
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed }, { "text", r.Message } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    case "academicBeautifyChart":
                        {
                            var r = AcademicToolService.BeautifyChart(
                                ToInt(payload, "palette"),
                                ToStr(payload, "fontName"));
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    case "insertTemplatePack":
                        {
                            var r = AcademicToolService.InsertTemplatePack(
                                ToStr(payload, "type"),
                                ToInt(payload, "contentSlides"));
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    case "scanDuplicates":
                        {
                            // 全库扫描"完全重复"（同类型/同数量/同尺寸；保留最早，返回可删的重复项）
                            var dups = ShapeLibraryService.FindExactDuplicates();
                            var arr = new List<Dictionary<string, object>>();
                            foreach (var d in dups)
                            {
                                arr.Add(new Dictionary<string, object>
                                {
                                    { "id", d.Id }, { "name", d.Name },
                                    { "kind", d.Kind }, { "shapeCount", d.ShapeCount },
                                    { "previewUrl", "https://data.pptshapes.local/" + d.Id + "/preview.png?v=" + PreviewStamp(d.Id) }
                                });
                            }
                            Ok(reqId, new Dictionary<string, object> { { "duplicates", arr } });
                            break;
                        }
                    case "createGroupCategory":
                        {
                            ShapeLibraryService.CreateGroupCategory(ToStr(payload, "group"), ToStr(payload, "category"));
                            Ok(reqId, new Dictionary<string, object>());
                            BroadcastRefresh();
                            break;
                        }
                    case "renameGroupCategory":
                        {
                            int n = ShapeLibraryService.RenameGroupCategory(
                                ToStr(payload, "group"), ToStr(payload, "oldCategory"), ToStr(payload, "newCategory"));
                            Ok(reqId, new Dictionary<string, object> { { "touched", n } });
                            BroadcastRefresh();
                            break;
                        }
                    case "deleteGroupCategory":
                        {
                            int n = ShapeLibraryService.DeleteGroupCategory(
                                ToStr(payload, "group"), ToStr(payload, "category"));
                            Ok(reqId, new Dictionary<string, object> { { "touched", n } });
                            BroadcastRefresh();
                            break;
                        }
                    case "exportPack":
                        {
                            // saveAs 为纯文件名（不含扩展名），实际落盘到 %APPDATA%\PPTShapeLibrary\packs\
                            string name = SanitizeFileName(ToStr(payload, "saveAs"));
                            if (name.Length == 0) name = "素材包_" + DateTime.Now.ToString("yyyyMMdd-HHmmss");
                            string dir = System.IO.Path.Combine(ShapeLibraryService.RootDir, "packs");
                            try { System.IO.Directory.CreateDirectory(dir); } catch { }
                            string saveAs = System.IO.Path.Combine(dir, name + ".ppack.zip");
                            // 已存在则追加时间戳，避免静默覆盖旧素材包
                            if (System.IO.File.Exists(saveAs))
                            {
                                saveAs = System.IO.Path.Combine(dir,
                                    name + "_" + DateTime.Now.ToString("HHmmss") + ".ppack.zip");
                            }
                            bool ok = PackService.ExportPack(GetIds(payload), saveAs, out string pkErr);
                            Result(reqId, ok,
                                   ok ? new Dictionary<string, object> { { "path", saveAs } } : null,
                                   ok ? null : pkErr);
                            break;
                        }
                    case "importPack":
                        {
                            // openFrom 为纯文件名；从 packs 目录读取（也兼容用户手工拷贝到此目录）
                            string name = SanitizeFileName(ToStr(payload, "openFrom"));
                            string dir = System.IO.Path.Combine(ShapeLibraryService.RootDir, "packs");
                            string full = name.EndsWith(".ppack.zip", StringComparison.OrdinalIgnoreCase)
                                ? System.IO.Path.Combine(dir, name)
                                : System.IO.Path.Combine(dir, name + ".ppack.zip");
                            int n = PackService.ImportPack(full, out int conflicts, out string impErr, out var packMeta);
                            bool ok = impErr == null;
                            if (ok)
                            {
                                var data = new Dictionary<string, object> { { "imported", n }, { "conflicts", conflicts } };
                                if (packMeta != null)
                                {
                                    if (!string.IsNullOrEmpty(packMeta.Name)) data["packName"] = packMeta.Name;
                                    if (!string.IsNullOrEmpty(packMeta.Author)) data["author"] = packMeta.Author;
                                    if (!string.IsNullOrEmpty(packMeta.License)) data["license"] = packMeta.License;
                                    if (!string.IsNullOrEmpty(packMeta.SourceUrl)) data["sourceUrl"] = packMeta.SourceUrl;
                                    if (!string.IsNullOrEmpty(packMeta.Description)) data["description"] = packMeta.Description;
                                }
                                Result(reqId, true, data, null);
                            }
                            else
                            {
                                Result(reqId, false, null, impErr);
                            }
                            BroadcastRefresh();
                            break;
                        }
                    case "log":
                        {
                            string level = ToStr(payload, "level");
                            string message = ToStr(payload, "message");
                            if (level == "error") Log.Error(new Exception(message), "[webview]");
                            else Log.Write("[webview] " + message);
                            Ok(reqId, new Dictionary<string, object>());
                            break;
                        }
                    default:
                        Result(reqId, false, null, "unknown type: " + type);
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Bridge:" + type);
                Result(reqId, false, null, ex.Message);
            }
        }

        // ---------------- 处理器实现 ----------------

        private Dictionary<string, object> BuildInit()
        {
            var d = new Dictionary<string, object>();
            var s = ShapeLibraryService.LoadSettings();
            foreach (var kv in SettingsDict(s)) d[kv.Key] = kv.Value;
            var all = ShapeLibraryService.LoadAllItems();
            d["categories"] = ShapeLibraryService.GetAllCategories();
            d["total"] = all.Count;
            // 浏览上下文：页面重载/窗口迁移重建后前端据此恢复上次浏览的 组/分类（避免跳回"全部"）
            if (ShapeLibraryService.BrowseWasSet(_hostKey))
            {
                d["lastGroup"] = ShapeLibraryService.BrowseGroupForRestore(_hostKey);
                d["lastCategory"] = ShapeLibraryService.BrowseCategoryForRestore(_hostKey) ?? "全部";
                d["lastKind"] = ShapeLibraryService.BrowseKindForRestore(_hostKey);
                d["lastFilter"] = ShapeLibraryService.BrowseFilterForRestore(_hostKey);
                d["lastSort"] = ShapeLibraryService.BrowseSortForRestore(_hostKey);
            }

            // 分类计数（前端胶囊徽标）
            var counts = new Dictionary<string, int>();
            foreach (var m in all)
            {
                string c = m.GetCategory();
                if (!counts.ContainsKey(c)) counts[c] = 0;
                counts[c]++;
            }
            d["categoryCounts"] = counts;

            // A6：全库最近使用 Top6（供前端"最近使用"横条；跨排序/筛选的全局数据）
            d["recentTop"] = all
                .Where(m => !string.IsNullOrWhiteSpace(m.LastUsedAt))
                .OrderByDescending(m => m.LastUsedAt)
                .Take(6)
                .Select(m => new Dictionary<string, object>
                {
                    { "id", m.Id },
                    { "name", m.Name },
                    { "kind", m.Kind },
                    { "category", m.GetCategory() },
                    { "lastUsedAt", m.LastUsedAt ?? "" },
                    { "useCount", m.UseCount },
                    { "previewUrl", "https://data.pptshapes.local/" + m.Id + "/preview.png?v=" + PreviewStamp(m.Id) },
                    { "snapshotPage", m.SnapshotPage }
                })
                .ToList();

            // 一级分组元数据：顺序组列表、计数、每组 icon 与“该组二级分类列表”
            var groups = ShapeLibraryService.GetGroups();
            d["groups"] = groups;
            d["presetGroups"] = ShapeLibraryService.PresetGroups.ToList();
            d["groupCounts"] = ShapeLibraryService.GetGroupCounts();
            var groupCats = new Dictionary<string, object>();
            foreach (var g in groups)
            {
                groupCats[g] = ShapeLibraryService.GetAllCategoriesOfGroup(g);
            }
            d["groupCategories"] = groupCats;
            return d;
        }

        private Dictionary<string, object> DoList(Dictionary<string, object> payload)
        {
            string filter = ToStr(payload, "filter");
            string category = ToStr(payload, "category");
            string kind = ToStr(payload, "kind");
            string group = ToStr(payload, "group");
            string sort = ToStr(payload, "sort");
            // 登记浏览上下文（按宿主隔离）：收藏落点 + 本宿主重载后视图恢复
            ShapeLibraryService.SetBrowseContext(_hostKey, group, category, kind, filter, sort);
            Log.Write("[list] host=" + _hostKey.Substring(0, 8) + " group=[" + group + "] cat=[" + category + "] kind=[" + kind + "] filter=[" + filter + "]");
            int offset = ToInt(payload, "offset");
            int limit = ToInt(payload, "limit");
            if (limit <= 0) limit = 200;

            var items = ShapeLibraryService.LoadAllItems().AsEnumerable();

            if (!string.IsNullOrWhiteSpace(filter))
            {
                items = items.Where(m =>
                    m.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.GetCategory().IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            if (!string.IsNullOrEmpty(category) && category != "全部")
            {
                items = items.Where(m => m.GetCategory() == category);
            }
            if (!string.IsNullOrEmpty(kind) && kind != "全部")
            {
                string k = kind;
                items = items.Where(m => (m.Kind ?? "").Equals(k, StringComparison.OrdinalIgnoreCase));
            }
            if (!string.IsNullOrEmpty(group) && group != "全部")
            {
                string g = group;
                items = items.Where(m => m.GetGroup() == g);
            }
            // 仅未使用（UseCount==0）
            if (ToBool(payload, "unusedOnly"))
            {
                items = items.Where(m => m.UseCount <= 0);
            }
            // 按主色近似筛选（RGB 曼哈顿距离阈值）
            string colorLike = ToStr(payload, "colorLike");
            int tr, tg, tb;
            if (!string.IsNullOrWhiteSpace(colorLike) && colorLike.Length >= 6 && TryParseHexColor(colorLike, out tr, out tg, out tb))
            {
                items = items.Where(m =>
                {
                    string hex = m.ColorHex;
                    if (string.IsNullOrWhiteSpace(hex) || hex.Length < 6) return false;
                    try
                    {
                        int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                        int g2 = Convert.ToInt32(hex.Substring(2, 2), 16);
                        int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                        return Math.Abs(r - tr) + Math.Abs(g2 - tg) + Math.Abs(b - tb) <= 90;
                    }
                    catch { return false; }
                });
            }

            int filteredTotal = items.Count();
            switch (sort)
            {
                case "最旧": items = items.OrderBy(m => m.CreatedAt); break;
                case "名称": items = items.OrderBy(m => m.Name, StringComparer.OrdinalIgnoreCase); break;
                case "最近使用": items = items.OrderByDescending(m => m.LastUsedAt ?? ""); break;
                case "最常用": items = items.OrderByDescending(m => m.UseCount).ThenByDescending(m => m.LastUsedAt ?? ""); break;
                case "自定义": items = items.OrderBy(m => m.SortOrder).ThenByDescending(m => m.CreatedAt); break;
                default: items = items.OrderByDescending(m => m.CreatedAt); break;
            }

            var page = items.Skip(offset).Take(limit).ToList();

            var arr = new List<Dictionary<string, object>>();
            foreach (var m in page)
            {
                arr.Add(new Dictionary<string, object>
                {
                    { "id", m.Id },
                    { "name", m.Name },
                    { "kind", m.Kind },
                    { "category", m.GetCategory() },
                    { "group", m.GetGroup() },
                    { "shapeCount", m.ShapeCount },
                    { "width", m.Width },
                    { "height", m.Height },
                    { "useCount", m.UseCount },
                    { "lastUsedAt", m.LastUsedAt ?? "" },
                    { "sortOrder", m.SortOrder },
                    { "previewUrl", "https://data.pptshapes.local/" + m.Id + "/preview.png?v=" + PreviewStamp(m.Id) },
                    { "snapshotPage", m.SnapshotPage },
                    { "colorHex", m.ColorHex ?? "" }
                });
            }

            return new Dictionary<string, object>
            {
                { "items", arr },
                { "filteredTotal", filteredTotal },
                { "total", ShapeLibraryService.LoadAllItems().Count },
                { "offset", offset },
                { "limit", limit }
            };
        }

        private Dictionary<string, object> DoSaveSettings(Dictionary<string, object> payload)
        {
            var patch = payload.ContainsKey("patch") ? payload["patch"] as Dictionary<string, object> : null;
            var s = ShapeLibraryService.LoadSettings();
            if (patch != null)
            {
                if (patch.ContainsKey("theme")) s.Theme = patch["theme"].ToString();
                if (patch.ContainsKey("insertPosition")) s.InsertPosition = patch["insertPosition"].ToString();
                if (patch.ContainsKey("cardSize")) s.CardSize = patch["cardSize"].ToString();
                if (patch.ContainsKey("autoName"))
                {
                    bool b; if (bool.TryParse(patch["autoName"].ToString(), out b)) s.AutoName = b;
                }
                if (patch.ContainsKey("clickToInsert"))
                {
                    bool b; if (bool.TryParse(patch["clickToInsert"].ToString(), out b)) s.ClickToInsert = b;
                }
                if (patch.ContainsKey("themeAdaptStyle"))
                {
                    bool b; if (bool.TryParse(patch["themeAdaptStyle"].ToString(), out b)) s.ThemeAdaptStyle = b;
                }
                if (patch.ContainsKey("collectPlaceholderText"))
                {
                    bool b; if (bool.TryParse(patch["collectPlaceholderText"].ToString(), out b)) s.CollectPlaceholderText = b;
                }
                if (patch.ContainsKey("gridColumns"))
                {
                    string v = patch["gridColumns"]?.ToString() ?? "";
                    if (v == "2" || v == "3" || v == "4" || v == "auto") s.GridColumns = v;
                }
                if (patch.ContainsKey("showTypeRow"))
                {
                    bool b; if (bool.TryParse(patch["showTypeRow"].ToString(), out b)) s.ShowTypeRow = b;
                }
                if (patch.ContainsKey("cleanLayoutsOnInsertPage"))
                {
                    bool b; if (bool.TryParse(patch["cleanLayoutsOnInsertPage"].ToString(), out b)) s.CleanLayoutsOnInsertPage = b;
                }
                if (patch.ContainsKey("accentColor"))
                {
                    string v = patch["accentColor"]?.ToString() ?? "";
                    string[] allowed = { "blue", "indigo", "violet", "emerald", "rose", "amber" };
                    if (Array.IndexOf(allowed, v) >= 0) s.AccentColor = v;
                }
                if (patch.ContainsKey("paneDock"))
                {
                    string v = patch["paneDock"]?.ToString() ?? "";
                    if (v == "right" || v == "left" || v == "top" || v == "bottom") s.PaneDock = v;
                }
                if (patch.ContainsKey("paneHeight"))
                {
                    int n; if (int.TryParse(patch["paneHeight"].ToString(), out n) && n >= 100 && n <= 800) s.PaneHeight = n;
                }
                if (patch.ContainsKey("pageSize"))
                {
                    int n; if (int.TryParse(patch["pageSize"].ToString(), out n)) s.PageSize = Math.Max(1, Math.Min(500, n));
                }
                if (patch.ContainsKey("paneWidth"))
                {
                    // 必须夹紧：该值会被 ThisAddIn.OnSettingsSaved 立即应用到任务窗格宽度并持久化，
                    // 未夹紧时一个坏值（如 999999）会让窗格宽度异常且重启后依然生效。
                    int n; if (int.TryParse(patch["paneWidth"].ToString(), out n)) s.PaneWidth = Math.Max(200, Math.Min(1200, n));
                }
                if (patch.ContainsKey("authorName"))
                {
                    s.AuthorName = (patch["authorName"]?.ToString() ?? "").Trim();
                    if (s.AuthorName.Length > 60) s.AuthorName = s.AuthorName.Substring(0, 60);
                }
                ShapeLibraryService.SaveSettings(s);
            }
            var result = SettingsDict(ShapeLibraryService.LoadSettings());
            var saved = ShapeLibraryService.LoadSettings();
            PushSettingsChanged(saved);
            try { SettingsSaved?.Invoke(saved); } catch { }
            return result;
        }

        // ---------------- 工具 ----------------

        private string KindOf(string id)
        {
            var meta = ShapeLibraryService.LoadMeta(id);
            return meta == null ? "" : meta.Kind;
        }

        /// <summary>
        /// 缩略图缓存失效戳：取 meta.json 文件修改时间的 Unix 秒。
        /// 重收藏/改名/改分类等元数据变化时该时间自然更新，浏览器缓存自动失效；
        /// 未变化的收藏保持原 URL，命中 HTTP 缓存（方案 4.3）。
        /// </summary>
        private static long PreviewStamp(string id)
        {
            try
            {
                string path = ShapeLibraryService.MetaFile(id);
                if (System.IO.File.Exists(path))
                {
                    long stamp = (long)(System.IO.File.GetLastWriteTimeUtc(path)
                        - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
                    if (stamp > 0) return stamp;
                }
            }
            catch { }
            return 0;
        }

        private static List<string> GetIds(Dictionary<string, object> payload)
        {
            var list = new List<string>();
            if (payload.ContainsKey("ids") && payload["ids"] is System.Collections.ArrayList arr)
            {
                foreach (var x in arr) list.Add(x.ToString());
            }
            return list;
        }

        private static List<string> GetStringArray(Dictionary<string, object> payload, string key)
        {
            var list = new List<string>();
            if (payload.ContainsKey(key) && payload[key] is System.Collections.ArrayList arr)
            {
                foreach (var x in arr) list.Add(x.ToString());
            }
            return list;
        }

        private static FormatToolService.CleanOptions CleanOptsFrom(Dictionary<string, object> payload)
        {
            return new FormatToolService.CleanOptions
            {
                DelEmptySlides = true == ToBoolNullable(payload, "delEmpty"),
                DelHiddenShapes = true == ToBoolNullable(payload, "delHidden"),
                DelOffSlideShapes = true == ToBoolNullable(payload, "delOffSlide"),
                ClearNotes = true == ToBoolNullable(payload, "clearNotes"),
                RemoveActions = true == ToBoolNullable(payload, "removeActions")
            };
        }

        private static string ToStr(Dictionary<string, object> d, string key)
        {
            return d.ContainsKey(key) && d[key] != null ? d[key].ToString() : "";
        }

        /// <summary>解析 "RRGGBB"（允许带 # 前缀）。非法输入返回 false，绝不抛异常——
        /// 此前直接 Convert.ToInt32(sub,16) 会让一个坏颜色值把整次 list 请求打成错误。</summary>
        private static bool TryParseHexColor(string s, out int r, out int g, out int b)
        {
            r = g = b = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim().TrimStart('#');
            if (s.Length < 6) return false;
            return int.TryParse(s.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out r)
                && int.TryParse(s.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out g)
                && int.TryParse(s.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out b);
        }

        /// <summary>文件名净化：替换路径分隔符/冒号/控制字符与 .. 段，防止路径穿越（覆盖 Windows 与便携语义）。</summary>
        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "";
            string s = name.Trim();
            var invalid = new System.Collections.Generic.HashSet<char>(System.IO.Path.GetInvalidFileNameChars());
            // 显式补：/ : 反斜杠与常见穿越字符（GetInvalidFileNameChars 在个别平台不含全量）
            foreach (char c in new[] { '/', '\\', ':', '\0', '\n', '\r', '\t' }) invalid.Add(c);
            var chars = new List<char>();
            foreach (char c in s)
            {
                chars.Add(invalid.Contains(c) ? '_' : c);
            }
            string result = new string(chars.ToArray()).Trim();
            // 消除纯 ".." 目录段（split 后重组），杜绝 "..\..\x" 残留穿越
            var segs = result.Split(new[] { '/', '\\' });
            var keep = segs.Where(seg => seg != ".." && seg != ".").ToArray();
            result = string.Join("_", keep).Trim();
            if (result == "." || result == "..") return "";
            // 按完整字符截断（避免切在代理对中间抛异常）
            return result.Length > 80 ? result.Substring(0, 80) : result;
        }

        private static int ToInt(Dictionary<string, object> d, string key)
        {
            if (d.ContainsKey(key))
            {
                try { return Convert.ToInt32(Convert.ToDouble(d[key])); } catch { }
            }
            return 0;
        }

        private static bool ToBool(Dictionary<string, object> d, string key)
        {
            if (d.ContainsKey(key) && d[key] != null)
            {
                try
                {
                    if (d[key] is bool b) return b;
                    return bool.TryParse(d[key].ToString(), out bool r) && r;
                }
                catch { }
            }
            return false;
        }

        private static float? ToFloatNullable(Dictionary<string, object> d, string key)
        {
            if (d.ContainsKey(key) && d[key] != null)
            {
                try
                {
                    float v = Convert.ToSingle(Convert.ToDouble(d[key]));
                    return v > 0 ? (float?)v : null;
                }
                catch { }
            }
            return null;
        }

        private static bool? ToBoolNullable(Dictionary<string, object> d, string key)
        {
            if (d.ContainsKey(key) && d[key] != null)
            {
                bool b;
                if (bool.TryParse(d[key].ToString(), out b)) return b;
                if (d[key].ToString() == "1") return true;
            }
            return null;
        }

        private static int? ToIntNullable(Dictionary<string, object> d, string key)
        {
            if (d.ContainsKey(key) && d[key] != null)
            {
                try { return Convert.ToInt32(Convert.ToDouble(d[key])); } catch { }
            }
            return null;
        }

        private Dictionary<string, object> SettingsDict(Settings s)
        {
            return new Dictionary<string, object>
            {
                { "theme", s.Theme },
                { "insertPosition", s.InsertPosition },
                { "paneWidth", s.PaneWidth },
                { "autoName", s.AutoName },
                { "pageSize", s.PageSize },
                { "cardSize", s.CardSize },
                { "clickToInsert", s.ClickToInsert },
                { "themeAdaptStyle", s.ThemeAdaptStyle },
                { "authorName", s.AuthorName ?? "" },
                { "collectPlaceholderText", s.CollectPlaceholderText },
                { "paneDock", s.PaneDock ?? "left" },
                { "paneHeight", s.PaneHeight },
                { "gridColumns", s.GridColumns ?? "2" },
                { "showTypeRow", s.ShowTypeRow },
                { "cleanLayoutsOnInsertPage", s.CleanLayoutsOnInsertPage },
                { "accentColor", s.AccentColor ?? "blue" }
            };
        }

        /// <summary>延后一点关闭 PowerPoint，让安装器进程先起来（它会等 PowerPoint 退出后再装）。
        ///
        /// ⚠️ 实测：在 WebView2 消息回调里调 `Application.Quit()` **有时会被 COM 吞掉**
        /// （日志显示调用了，但 PowerPoint 不关）。所以这里**重试多次**；
        /// 若始终关不掉（比如文稿有未保存改动、或 Office 处于某种模态状态），
        /// 就明确提示用户手动关闭 —— 安装器会一直在等，手动关掉后自动完成升级。</summary>
        private void ScheduleQuitPowerPoint()
        {
            var t = new System.Windows.Forms.Timer();
            t.Interval = 2000;
            int tries = 0;
            t.Tick += (s, e) =>
            {
                tries++;
                try
                {
                    var app = Globals.ThisAddIn != null ? Globals.ThisAddIn.Application : null;
                    if (app != null)
                    {
                        if (tries == 1) Log.Write("[update] 正在关闭 PowerPoint 以完成更新…（最多重试 6 次）");
                        app.Quit();
                    }
                }
                catch (Exception ex)
                {
                    Log.Write("[update] 关闭 PowerPoint 失败：" + ex.Message);
                }
                if (tries >= 6)
                {
                    t.Stop();
                    t.Dispose();
                    Log.Write("[update] 自动关闭未生效 → 请手动关闭 PowerPoint，安装器会自动完成升级");
                }
            };
            t.Start();
        }

        /// <summary>主动告诉前端"有新版本"，前端据此弹窗/显示提示条。
        /// （定时提醒由 UpdateReminder 负责，它在插件启动时就跑，不依赖面板是否打开。）</summary>
        private void PushUpdateAvailable(object u)
        {
            var ur = u as UpdateResult;
            if (ur == null) return;
            Post(new Dictionary<string, object>
            {
                { "type", "updateAvailable" },
                {
                    "payload", new Dictionary<string, object>
                    {
                        { "currentVersion", ur.CurrentVersion },
                        { "latestVersion", ur.LatestVersion },
                        { "notes", ur.Notes },
                        { "releaseUrl", ur.ReleaseUrl },
                        { "downloadUrl", ur.DownloadUrl },
                    }
                }
            });
        }

        // ---------------- 下行 ----------------

        private void Ok(string reqId, Dictionary<string, object> data)
        {
            Result(reqId, true, data, null);
        }

        private void Result(string reqId, bool ok, Dictionary<string, object> data, string error)
        {
            var m = new Dictionary<string, object> { { "type", "result" }, { "reqId", reqId }, { "ok", ok } };
            if (data != null) m["data"] = data;
            if (error != null) m["error"] = error;
            Post(m);
        }

        private void Post(Dictionary<string, object> msg)
        {
            if (_closed) return;   // 宿主已释放（窗格迁移/退出）：再发消息只会刷一堆 disposed 异常日志
            try
            {
                _core.PostWebMessageAsJson(_json.Serialize(msg));
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Bridge.Post");
            }
        }

        /// <summary>宿主释放时调用：之后所有下行消息直接丢弃。</summary>
        public void MarkClosed()
        {
            _closed = true;
        }
    }
}
