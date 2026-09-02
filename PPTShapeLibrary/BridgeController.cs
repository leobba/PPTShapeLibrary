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
                            bool ok = ShapeExporter.InsertShape(id, false, out string insErr);   // 静默，前端 Toast
                            if (ok) ShapeLibraryService.TouchUsed(id);
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
                    case "undo":
                        ShapeExporter.UndoLastAction();
                        Ok(reqId, new Dictionary<string, object>());
                        break;
                    case "adjust":
                        ShapeExporter.AdjustSelected(ToStr(payload, "action"));
                        Ok(reqId, new Dictionary<string, object>());
                        break;
                    case "delete":
                        {
                            int n = 0;
                            foreach (var id in GetIds(payload))
                            {
                                ShapeLibraryService.DeleteItem(id);
                                n++;
                            }
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
                            int n = PackService.ImportPack(full, out string impErr);
                            bool ok = impErr == null;
                            Result(reqId, ok, ok ? new Dictionary<string, object> { { "imported", n } } : null,
                                   ok ? null : impErr);
                            PushRefresh();
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
            return d;
        }

        private Dictionary<string, object> DoList(Dictionary<string, object> payload)
        {
            string filter = ToStr(payload, "filter");
            string category = ToStr(payload, "category");
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
                ShapeLibraryService.SaveSettings(s);
            }
            var result = SettingsDict(ShapeLibraryService.LoadSettings());
            PushSettingsChanged(ShapeLibraryService.LoadSettings());
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

        private static string ToStr(Dictionary<string, object> d, string key)
        {
            return d.ContainsKey(key) && d[key] != null ? d[key].ToString() : "";
        }

        /// <summary>文件名净化：去掉路径分隔符与非法字符，防止路径穿越。</summary>
        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "";
            string s = name.Trim();
            var invalid = System.IO.Path.GetInvalidFileNameChars();
            var chars = new List<char>();
            foreach (char c in s)
            {
                chars.Add(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
            }
            string result = new string(chars.ToArray()).Trim();
            if (result == "." || result == "..") return "";
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
                { "themeAdaptStyle", s.ThemeAdaptStyle }
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
