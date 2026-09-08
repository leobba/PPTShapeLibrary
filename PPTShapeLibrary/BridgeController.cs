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

        public BridgeController(CoreWebView2 core)
        {
            _core = core;
        }

        public void Attach()
        {
            // 首屏数据在页面 ready 后由前端发 init 拉取
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

        public void PushSwitchView(string view)
        {
            Post(new Dictionary<string, object> { { "type", "switchView" }, { "payload", new Dictionary<string, object> { { "view", view } } } });
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

                switch (type)
                {
                    case "ping":
                        Ok(reqId, new Dictionary<string, object> { { "pong", true } });
                        break;
                    case "init":
                        Ok(reqId, BuildInit());
                        break;
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
                            Result(reqId, ok, ok ? new Dictionary<string, object> { { "kind", KindOf(id) } } : null,
                                   ok ? null : insErr);
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
                            int n = ShapeExporter.CollectAllSlides();
                            if (n > 0) BroadcastRefresh();
                            Result(reqId, n > 0,
                                   n > 0 ? new Dictionary<string, object> { { "collected", n } } : null,
                                   n > 0 ? null : "没有可收藏的页面（跳过空白/占位页）");
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
                            if (failed > 0)
                                Result(reqId, false, null,
                                    "有 " + failed + " 项移入回收站失败（可能被占用或文件锁定），已保留在原位。");
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
                    case "setTags":
                        {
                            string id = ToStr(payload, "id");
                            ShapeLibraryService.SetTags(id, ToStr(payload, "tags"));
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
                            if (ordered != null)
                            {
                                for (int i = 0; i < ordered.Count; i++)
                                {
                                    ShapeLibraryService.SetSortOrder(ordered[i].ToString(), i);
                                }
                            }
                            Ok(reqId, new Dictionary<string, object>());
                            break;
                        }
                    case "openFolder":
                        {
                            string dir = Path.Combine(ShapeLibraryService.ItemsDir, ToStr(payload, "id"));
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
                            // 修复历史错位缩略图：重建全部形状条目预览（逐条打开库文稿导出）
                            int n = ShapeExporter.RebuildAllPreviews();
                            Ok(reqId, new Dictionary<string, object> { { "rebuilt", n } });
                            BroadcastRefresh();
                            break;
                        }
                    // ---------------- PPT 工具（Web 工具页） ----------------
                    case "toolGetFonts":
                        {
                            var sysFonts = new List<string>();
                            try
                            {
                                using (var fc = new System.Drawing.Text.InstalledFontCollection())
                                {
                                    foreach (var fam in fc.Families)
                                    {
                                        string n = fam.Name;
                                        // 仅排除本插件的图标字体家族（MaterialIcons/Material Symbols），
                                        // 保留 Segoe UI Symbol 等真实可用符号字体
                                        if (n.StartsWith("Material", StringComparison.OrdinalIgnoreCase)) continue;
                                        if (n.Contains("PPTShapeLibrary")) continue;
                                        if (!sysFonts.Contains(n)) sysFonts.Add(n);
                                    }
                                }
                            }
                            catch { }
                            sysFonts.Sort(StringComparer.OrdinalIgnoreCase);
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
                    case "replaceFont":
                        {
                            bool allSlides = ToBoolNullable(payload, "allSlides") ?? true;   // 默认整稿
                            var r = TextToolService.ReplaceFont(
                                ToStr(payload, "fromFont"),
                                ToStr(payload, "toFont"),
                                allSlides);
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
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
                            var r = FormatToolService.ExportImages(
                                ToStr(payload, "dir"),
                                ToStr(payload, "format"),
                                ToInt(payload, "width"));
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    case "formatScanClean":
                        {
                            var opts = CleanOptsFrom(payload);
                            var r = FormatToolService.ScanClean(opts);
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed }, { "message", r.Message } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    case "formatRunClean":
                        {
                            var opts = CleanOptsFrom(payload);
                            var r = FormatToolService.RunClean(opts);
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed }, { "message", r.Message } } : null,
                                   r.Ok ? null : r.Message);
                            break;
                        }
                    case "formatResize":
                        {
                            var r = FormatToolService.ResizeShapes(
                                ToStr(payload, "scope"),
                                ToFloatNullable(payload, "width"),
                                ToFloatNullable(payload, "height"),
                                true == ToBoolNullable(payload, "forceImageRatio"));
                            Result(reqId, r.Ok, r.Ok ? new Dictionary<string, object> { { "changed", r.Changed } } : null,
                                   r.Ok ? null : r.Message);
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
                    case "renameTag":
                        {
                            int n = ShapeLibraryService.RenameTagAll(ToStr(payload, "oldTag"), ToStr(payload, "newTag"));
                            Ok(reqId, new Dictionary<string, object> { { "touched", n } });
                            break;
                        }
                    case "deleteTag":
                        {
                            int n = ShapeLibraryService.DeleteTagAll(ToStr(payload, "tag"));
                            Ok(reqId, new Dictionary<string, object> { { "touched", n } });
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
            d["tags"] = ShapeLibraryService.GetAllTags();
            d["tagCounts"] = ShapeLibraryService.GetTagCounts();
            d["total"] = all.Count;

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
                    { "tags", m.GetTags() },
                    { "lastUsedAt", m.LastUsedAt ?? "" },
                    { "useCount", m.UseCount },
                    { "previewUrl", "https://data.pptshapes.local/" + m.Id + "/preview.png?v=" + PreviewStamp(m.Id) }
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
            string tag = ToStr(payload, "tag");
            string group = ToStr(payload, "group");
            string sort = ToStr(payload, "sort");
            int offset = ToInt(payload, "offset");
            int limit = ToInt(payload, "limit");
            if (limit <= 0) limit = 200;

            var items = ShapeLibraryService.LoadAllItems().AsEnumerable();

            if (!string.IsNullOrWhiteSpace(filter))
            {
                items = items.Where(m =>
                    m.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.GetCategory().IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    string.Join(" ", m.GetTags()).IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
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
            if (!string.IsNullOrEmpty(tag))
            {
                string t = tag;
                items = items.Where(m => m.GetTags().Any(x => x.Equals(t, StringComparison.OrdinalIgnoreCase)));
            }
            if (!string.IsNullOrEmpty(group) && group != "全部")
            {
                string g = group;
                items = items.Where(m => m.GetGroup() == g);
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
                var tags = m.GetTags();
                arr.Add(new Dictionary<string, object>
                {
                    { "id", m.Id },
                    { "name", m.Name },
                    { "kind", m.Kind },
                    { "category", m.GetCategory() },
                    { "group", m.GetGroup() },
                    { "tags", tags },
                    { "shapeCount", m.ShapeCount },
                    { "width", m.Width },
                    { "height", m.Height },
                    { "useCount", m.UseCount },
                    { "lastUsedAt", m.LastUsedAt ?? "" },
                    { "sortOrder", m.SortOrder },
                    { "previewUrl", "https://data.pptshapes.local/" + m.Id + "/preview.png?v=" + PreviewStamp(m.Id) }
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
                if (patch.ContainsKey("pageSize"))
                {
                    int n; if (int.TryParse(patch["pageSize"].ToString(), out n) && n > 0) s.PageSize = n;
                }
                if (patch.ContainsKey("paneWidth"))
                {
                    int n; if (int.TryParse(patch["paneWidth"].ToString(), out n)) s.PaneWidth = n;
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
                { "authorName", s.AuthorName ?? "" }
            };
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
            try
            {
                _core.PostWebMessageAsJson(_json.Serialize(msg));
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Bridge.Post");
            }
        }
    }
}
