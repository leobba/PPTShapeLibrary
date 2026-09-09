using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTShapeLibrary
{
    public enum LibraryItemKind
    {
        Shape,
        Style
    }

    [DataContract]
    public class LibraryMeta
    {
        [DataMember(Name = "id")]
        public string Id { get; set; }
        [DataMember(Name = "name")]
        public string Name { get; set; }
        [DataMember(Name = "kind")]
        public string Kind { get; set; }
        [DataMember(Name = "createdAt")]
        public string CreatedAt { get; set; }
        [DataMember(Name = "shapeCount")]
        public int ShapeCount { get; set; }
        [DataMember(Name = "sourceSlide")]
        public string SourceSlide { get; set; }
        [DataMember(Name = "category")]
        public string Category { get; set; }
        [DataMember(Name = "group", EmitDefaultValue = false)]
        public string Group { get; set; }          // 一级分组（我的/模板/素材/文字/背景/图例/自定义…），缺省=我的
        [DataMember(Name = "lastUsedAt")]
        public string LastUsedAt { get; set; }
        [DataMember(Name = "sortOrder")]
        public int SortOrder { get; set; }
        [DataMember(Name = "useCount")]
        public int UseCount { get; set; }
        [DataMember(Name = "width")]
        public float Width { get; set; }
        [DataMember(Name = "height")]
        public float Height { get; set; }
        [DataMember(Name = "sourceAutoType")]
        public int SourceAutoType { get; set; }

        public string GetCategory()
        {
            return string.IsNullOrWhiteSpace(Category) ? "未分类" : Category;
        }

        public const string DefaultGroup = "我的";

        public string GetGroup()
        {
            return string.IsNullOrWhiteSpace(Group) ? DefaultGroup : Group;
        }
    }

    /// <summary>
    /// 本地文件夹存储：%APPDATA%\PPTShapeLibrary\items\{id}\
    ///   meta.json      元数据
    ///   shape.pptx      Kind=Shape 时保存的形状（单页演示文稿）
    ///   style.json      Kind=Style 时保存的样式属性
    ///   preview.png    缩略图
    /// </summary>
    public class ShapeLibraryService
    {
        public static string RootDir
        {
            get
            {
                // 注意：目录名保持 "PPTShapeLibrary"（历史数据/备份路径兼容），勿改为中文
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PPTShapeLibrary");
                Directory.CreateDirectory(dir);
                return dir;
            }
        }

        public static string ItemsDir
        {
            get
            {
                string dir = Path.Combine(RootDir, "items");
                Directory.CreateDirectory(dir);
                return dir;
            }
        }

        /// <summary>回收站目录（软删条目移动至此，可恢复/彻底清除）。</summary>
        public static string TrashDir
        {
            get
            {
                string dir = Path.Combine(RootDir, "trash");
                Directory.CreateDirectory(dir);
                return dir;
            }
        }

        public static string CreateItemDir(string id)
        {
            string dir = Path.Combine(ItemsDir, id);
            Directory.CreateDirectory(dir);
            return dir;
        }

        public static string NewId()
        {
            return DateTime.Now.ToString("yyyyMMdd-HHmmss-") + Guid.NewGuid().ToString("N").Substring(0, 6);
        }

        private static readonly System.Text.RegularExpressions.Regex IdSafeRegex =
            new System.Text.RegularExpressions.Regex("^[0-9A-Za-z_\\-]+$");

        /// <summary>安全 id 校验：id 只允许字母数字下划线连字符（NewId 生成格式），杜绝路径穿越（..、盘符、分隔符）。</summary>
        public static bool IsSafeId(string id)
        {
            return !string.IsNullOrEmpty(id) && id.Length <= 80 && IdSafeRegex.IsMatch(id);
        }

        public static string ShapeFile(string id)
        {
            return Path.Combine(ItemsDir, id, "shape.pptx");
        }

        public static string StyleFile(string id)
        {
            return Path.Combine(ItemsDir, id, "style.json");
        }

        public static string PreviewFile(string id)
        {
            return Path.Combine(ItemsDir, id, "preview.png");
        }

        public static string MetaFile(string id)
        {
            return Path.Combine(ItemsDir, id, "meta.json");
        }

        public static void SaveMeta(LibraryMeta meta)
        {
            Directory.CreateDirectory(Path.Combine(ItemsDir, meta.Id));
            var ser = new DataContractJsonSerializer(typeof(LibraryMeta));
            using (var fs = File.Create(MetaFile(meta.Id)))
            {
                ser.WriteObject(fs, meta);
            }
            InvalidateCache();
        }

        public static LibraryMeta LoadMeta(string id)
        {
            string path = MetaFile(id);
            if (!File.Exists(path)) return null;
            try
            {
                var ser = new DataContractJsonSerializer(typeof(LibraryMeta));
                using (var fs = File.OpenRead(path))
                {
                    return (LibraryMeta)ser.ReadObject(fs);
                }
            }
            catch
            {
                return null;
            }
        }

        // ---------------- 进程内缓存（B 组优化） ----------------
        // 收藏列表按目录扫描 + meta.json 反序列化生成，高频调用（list/搜索/分类计数/标签
        // 列表）代价高。此处以"2 秒 TTL + 任意写操作置脏"缓存 LoadAllItems 结果。
        // 注意：仅缓存"已排序列表的快照引用"；meta 对象只读共享，任何写路径都先 LoadMeta
        // 读文件、SaveMeta 落盘后置脏，绝不原地修改缓存中的 meta（防止脏读）。

        private static readonly object CacheLock = new object();
        private static List<LibraryMeta> _allCache;
        private static DateTime _allCacheTime;
        private static bool _allDirty = true;

        private static void InvalidateCache()
        {
            lock (CacheLock)
            {
                _allDirty = true;
            }
        }

        /// <summary>供外部（素材包导入等直接写 items 目录的路径）显式失效缓存。</summary>
        public static void InvalidateCachePublic()
        {
            InvalidateCache();
        }

        private const int CacheTtlMs = 2000;

        /// <summary>全部收藏（CreatedAt 倒序）。带 2s TTL 进程内缓存，写操作自动失效。</summary>
        public static List<LibraryMeta> LoadAllItems()
        {
            lock (CacheLock)
            {
                if (!_allDirty && _allCache != null
                    && (DateTime.UtcNow - _allCacheTime).TotalMilliseconds < CacheTtlMs)
                {
                    return new List<LibraryMeta>(_allCache); // 浅拷贝防调用方误改缓存本体
                }

                var list = new List<LibraryMeta>();
                if (Directory.Exists(ItemsDir))
                {
                    foreach (var dir in Directory.GetDirectories(ItemsDir))
                    {
                        string id = Path.GetFileName(dir);
                        var meta = LoadMeta(id);
                        if (meta != null)
                        {
                            list.Add(meta);
                        }
                    }
                }
                _allCache = list.OrderByDescending(m => m.CreatedAt).ToList();
                _allCacheTime = DateTime.UtcNow;
                _allDirty = false;
                return new List<LibraryMeta>(_allCache);
            }
        }

        /// <summary>软删除：把条目移入回收站（可恢复）。失败时保留原位并返回 false（不静默删除）。</summary>
        public static bool DeleteItem(string id)
        {
            if (!IsSafeId(id)) { Log.Error(new Exception("DeleteItem: unsafe id=" + (id ?? "")), "Security"); return false; }
            string dir = Path.Combine(ItemsDir, id);
            if (!Directory.Exists(dir)) return true;
            string trash = Path.Combine(TrashDir, id);
            try
            {
                if (Directory.Exists(trash)) Directory.Delete(trash, true);
                Directory.Move(dir, trash);
                InvalidateCache();
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "DeleteItem(soft) id=" + id);
                return false;   // 保留原位，调用方提示，绝不静默永久删除
            }
        }

        /// <summary>彻底删除（仅回收站内清除）。注意：不触碰 items 下可能同名的新条目，
        /// 避免"软删后从备份恢复的同 id 条目被回收站彻底删除连带误删"。</summary>
        public static void HardDelete(string id)
        {
            if (!IsSafeId(id)) { Log.Error(new Exception("HardDelete: unsafe id=" + (id ?? "")), "Security"); return; }
            string dir = Path.Combine(TrashDir, id);
            if (Directory.Exists(dir))
            {
                try { Directory.Delete(dir, true); } catch { }
            }
            InvalidateCache();
        }

        /// <summary>从回收站恢复条目到素材库。</summary>
        public static bool RestoreTrashItem(string id)
        {
            if (!IsSafeId(id)) { Log.Error(new Exception("RestoreTrashItem: unsafe id=" + (id ?? "")), "Security"); return false; }
            string trash = Path.Combine(TrashDir, id);
            if (!Directory.Exists(trash)) return false;
            string dest = Path.Combine(ItemsDir, id);
            if (Directory.Exists(dest))
            {
                // 目标冲突：改名恢复（新 id 安全）
                dest = Path.Combine(ItemsDir, NewId());
            }
            try
            {
                Directory.Move(trash, dest);
                InvalidateCache();
                // P1-6: 改名恢复后同步 meta.json 内 id 字段，避免目录名与 meta id 不一致
                if (Path.GetFileName(dest) != id)
                {
                    try { RewriteMetaIdText(Path.Combine(dest, "meta.json"), Path.GetFileName(dest)); } catch { }
                }
                return true;
            }
            catch { return false; }
        }

        /// <summary>回收站列表（meta 可读的条目）。</summary>
        public static List<LibraryMeta> ListTrash()
        {
            var list = new List<LibraryMeta>();
            if (!Directory.Exists(TrashDir)) return list;
            foreach (var dir in Directory.GetDirectories(TrashDir))
            {
                string id = Path.GetFileName(dir);
                var meta = LoadMetaFromDir(Path.Combine(dir, "meta.json"));
                if (meta != null)
                {
                    meta.Id = id;
                    list.Add(meta);
                }
            }
            return list.OrderByDescending(m => m.CreatedAt).ToList();
        }

        /// <summary>把 meta.json 文本里的首个 id 字段改写为新 id（目录改名/冲突恢复后同步，防 id 不一致）。</summary>
        private static void RewriteMetaIdText(string metaFile, string newId)
        {
            try
            {
                if (!System.IO.File.Exists(metaFile)) return;
                string text = System.IO.File.ReadAllText(metaFile, System.Text.Encoding.UTF8);
                var rx = new System.Text.RegularExpressions.Regex("\"id\"\\s*:\\s*\"[^\"]*\"");
                var m = rx.Match(text);
                if (m.Success)
                {
                    text = text.Substring(0, m.Index)
                        + "\"id\":\"" + newId + "\""
                        + text.Substring(m.Index + m.Length);
                    System.IO.File.WriteAllText(metaFile, text, System.Text.Encoding.UTF8);
                }
            }
            catch { }
        }

        private static LibraryMeta LoadMetaFromDir(string metaPath)
        {
            if (!File.Exists(metaPath)) return null;
            try
            {
                var ser = new DataContractJsonSerializer(typeof(LibraryMeta));
                using (var fs = File.OpenRead(metaPath))
                {
                    return (LibraryMeta)ser.ReadObject(fs);
                }
            }
            catch { return null; }
        }

        /// <summary>清空回收站。返回清除条数。</summary>
        public static int PurgeTrash()
        {
            int n = 0;
            if (!Directory.Exists(TrashDir)) return 0;
            foreach (var dir in Directory.GetDirectories(TrashDir))
            {
                try { Directory.Delete(dir, true); n++; } catch { }
            }
            InvalidateCache();
            return n;
        }

        public static void RenameItem(string id, string newName)
        {
            var meta = LoadMeta(id);
            if (meta == null) return;
            // 空名/纯空白保持原名，超长截断，避免列表出现空白条目或异常名
            if (string.IsNullOrWhiteSpace(newName)) return;
            string name = newName.Trim();
            if (name.Length > 120) name = name.Substring(0, 120);
            meta.Name = name;
            SaveMeta(meta);
            InvalidateCache();
        }

        public static void SaveStyleJson(string id, string json)
        {
            File.WriteAllText(StyleFile(id), json, Encoding.UTF8);
        }

        /// <summary>孤儿目录清理：把 items 下没有 meta.json（收藏流程中断残留）的目录移入回收站，返回清理数。
        /// 不硬删，用户仍可在回收站查看/彻底删除。</summary>
        public static int CollectOrphanDirs()
        {
            int n = 0;
            try
            {
                if (!Directory.Exists(ItemsDir)) return 0;
                foreach (var dir in Directory.GetDirectories(ItemsDir))
                {
                    string id = Path.GetFileName(dir);
                    if (!IsSafeId(id)) continue;
                    if (File.Exists(Path.Combine(dir, "meta.json"))) continue;
                    string trash = Path.Combine(TrashDir, id);
                    try
                    {
                        if (Directory.Exists(trash)) Directory.Delete(trash, true);
                        Directory.Move(dir, trash);
                        n++;
                    }
                    catch { }
                }
                if (n > 0) InvalidateCache();
            }
            catch { }
            return n;
        }

        public static string LoadStyleJson(string id)
        {
            string path = StyleFile(id);
            return File.Exists(path) ? File.ReadAllText(path, Encoding.UTF8) : null;
        }

        public static void SavePreviewImage(string id, byte[] png)
        {
            File.WriteAllBytes(PreviewFile(id), png);
        }

        public static byte[] LoadPreviewImage(string id)
        {
            string path = PreviewFile(id);
            return File.Exists(path) ? File.ReadAllBytes(path) : null;
        }

        // ---------------- 分类 ----------------

        public static List<string> GetAllCategories()
        {
            var list = new List<string>();
            foreach (var meta in LoadAllItems())
            {
                string cat = meta.GetCategory();
                if (!list.Contains(cat)) list.Add(cat);
            }
            return list.OrderBy(c => c).ToList();
        }

        public static void SetCategory(string id, string category)
        {
            var meta = LoadMeta(id);
            if (meta == null) return;
            meta.Category = string.IsNullOrWhiteSpace(category) ? "未分类" : category.Trim();
            SaveMeta(meta);
            InvalidateCache();
        }

        /// <summary>记录最近使用时间与使用次数。</summary>
        public static void TouchUsed(string id)
        {
            var meta = LoadMeta(id);
            if (meta == null) return;
            meta.LastUsedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            meta.UseCount++;
            SaveMeta(meta);
            InvalidateCache();
        }

        public static void SetSortOrder(string id, int order)
        {
            var meta = LoadMeta(id);
            if (meta == null) return;
            meta.SortOrder = order;
            SaveMeta(meta);
            InvalidateCache();
        }

        // ---------------- 一级分组（Group） ----------------

        /// <summary>预设组（固定顺序与图标名；「全部」为前端虚拟）。图标名供前端 msym 用。</summary>
        public static readonly string[] PresetGroups = { "我的", "模板", "素材", "文字", "背景", "图例" };

        /// <summary>用户自建组文件（保留顺序）。</summary>
        private static string UserGroupsFile
        {
            get { return Path.Combine(RootDir, "usergroups.json"); }
        }

        private static List<string> _userGroups;

        public static List<string> GetUserGroups()
        {
            if (_userGroups == null)
            {
                _userGroups = new List<string>();
                try
                {
                    if (File.Exists(UserGroupsFile))
                    {
                        string json = File.ReadAllText(UserGroupsFile);
                        var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                        var arr = ser.Deserialize<List<string>>(json);
                        if (arr != null) _userGroups = arr.Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
                    }
                }
                catch { }
            }
            return new List<string>(_userGroups);
        }

        private static void SaveUserGroups()
        {
            try
            {
                var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                File.WriteAllText(UserGroupsFile, ser.Serialize(_userGroups ?? new List<string>()));
            }
            catch { }
        }

        /// <summary>全部一级组（含「全部」占位由前端处理，这里只给实体组），图标名映射前端。</summary>
        public static List<string> GetGroups()
        {
            var list = new List<string>();
            foreach (var g in PresetGroups) list.Add(g);
            foreach (var g in GetUserGroups())
            {
                if (!list.Contains(g)) list.Add(g);
            }
            return list;
        }

        /// <summary>组 → 条目计数（供前端右侧栏角标）。</summary>
        public static Dictionary<string, int> GetGroupCounts()
        {
            var counts = new Dictionary<string, int>();
            foreach (var m in LoadAllItems())
            {
                string g = m.GetGroup();
                counts.TryGetValue(g, out int n);
                counts[g] = n + 1;
            }
            return counts;
        }

        public static void SetGroup(string id, string group)
        {
            var meta = LoadMeta(id);
            if (meta == null) return;
            meta.Group = string.IsNullOrWhiteSpace(group) ? LibraryMeta.DefaultGroup : group.Trim();
            SaveMeta(meta);
            InvalidateCache();
        }

        /// <summary>新建自定义组。返回 false 表示重名或非法。</summary>
        public static bool CreateUserGroup(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            name = name.Trim();
            if (PresetGroups.Contains(name)) return false;
            var ug = GetUserGroups();
            if (ug.Contains(name)) return false;
            ug.Add(name);
            _userGroups = ug;
            SaveUserGroups();
            return true;
        }

        /// <summary>重命名组：groups.json 更名，并把该组条目的 Group 一并更新。</summary>
        public static void RenameGroup(string oldName, string newName)
        {
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName)) return;
            newName = newName.Trim();
            if (PresetGroups.Contains(oldName)) return;          // 预设组不允许改名
            var ug = GetUserGroups();
            int i = ug.IndexOf(oldName);
            if (i < 0) return;
            ug[i] = newName;
            _userGroups = ug;
            SaveUserGroups();
            foreach (var m in LoadAllItems())
            {
                if (m.GetGroup() == oldName)
                {
                    m.Group = newName;
                    SaveMeta(m);
                }
            }
            InvalidateCache();
        }

        /// <summary>删除用户组：其条目移到「我的」。返回受影响条目数。</summary>
        public static int DeleteGroup(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return 0;
            if (PresetGroups.Contains(name)) return 0;
            var ug = GetUserGroups();
            if (!ug.Remove(name)) return 0;
            _userGroups = ug;
            SaveUserGroups();
            int n = 0;
            foreach (var m in LoadAllItems())
            {
                if (m.GetGroup() == name)
                {
                    m.Group = LibraryMeta.DefaultGroup;
                    SaveMeta(m);
                    n++;
                }
            }
            InvalidateCache();
            return n;
        }

        /// <summary>某组内出现的全部二级分类（按条目 category 汇总 + 组预设空分类也列出空组）。</summary>
        public static List<string> GetCategoriesOfGroup(string group)
        {
            var list = new List<string>();
            foreach (var m in LoadAllItems())
            {
                if (m.GetGroup() != group) continue;
                string c = m.GetCategory();
                if (!list.Contains(c)) list.Add(c);
            }
            return list.OrderBy(c => c).ToList();
        }

        // ---------------- 用户二级分类清单（每组可声明的分类，含空分类；持久化 usercategories.json） ----------------

        private static string UserCatsFile
        {
            get { return Path.Combine(RootDir, "usercategories.json"); }
        }
        private static Dictionary<string, List<string>> _userCats;

        private static Dictionary<string, List<string>> LoadUserCats()
        {
            if (_userCats != null) return _userCats;
            _userCats = new Dictionary<string, List<string>>();
            try
            {
                if (File.Exists(UserCatsFile))
                {
                    var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                    var raw = ser.Deserialize<Dictionary<string, object>>(File.ReadAllText(UserCatsFile));
                    if (raw != null)
                    {
                        foreach (var kv in raw)
                        {
                            var arr = kv.Value as System.Collections.ArrayList;
                            if (arr != null)
                                _userCats[kv.Key] = arr.Cast<object>()
                                    .Select(o => o?.ToString() ?? "").Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
                        }
                    }
                }
            }
            catch { }
            return _userCats;
        }

        private static void SaveUserCats()
        {
            try
            {
                var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                File.WriteAllText(UserCatsFile, ser.Serialize(LoadUserCats()));
            }
            catch { }
        }

        /// <summary>用户声明过的二级分类（含空分类），可自行管理。</summary>
        public static List<string> GetUserCategoriesOfGroup(string group)
        {
            var d = LoadUserCats();
            return d.ContainsKey(group ?? "") ? new List<string>(d[group ?? ""]) : new List<string>();
        }

        /// <summary>组可展示的全部二级分类 = 组内条目真实分类 ∪ 用户声明分类（去重排序）。</summary>
        public static List<string> GetAllCategoriesOfGroup(string group)
        {
            var set = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in GetCategoriesOfGroup(group)) if (c != "未分类") set.Add(c);
            foreach (var c in GetUserCategoriesOfGroup(group)) set.Add(c);
            return set.ToList();
        }

        /// <summary>新建二级分类（写清单；同名已存在则不重复）。</summary>
        public static void CreateGroupCategory(string group, string cat)
        {
            if (string.IsNullOrWhiteSpace(group) || string.IsNullOrWhiteSpace(cat)) return;
            var d = LoadUserCats();
            string g = group.Trim();
            if (!d.ContainsKey(g)) d[g] = new List<string>();
            string c = cat.Trim();
            if (!d[g].Any(x => x.Equals(c, StringComparison.OrdinalIgnoreCase))) d[g].Add(c);
            SaveUserCats();
        }

        /// <summary>重命名二级分类：改清单 + 该组条目批量改 category。返回受影响条目数。</summary>
        public static int RenameGroupCategory(string group, string oldCat, string newCat)
        {
            if (string.IsNullOrWhiteSpace(oldCat) || string.IsNullOrWhiteSpace(newCat)) return 0;
            if (oldCat == newCat) return 0;
            int touched = 0;
            // 条目批量改
            foreach (var m in LoadAllItems())
            {
                if (m.GetGroup() != group) continue;
                if (m.GetCategory().Equals(oldCat, StringComparison.OrdinalIgnoreCase))
                {
                    SetCategory(m.Id, newCat.Trim());
                    touched++;
                }
            }
            // 清单同步
            var d = LoadUserCats();
            if (d.ContainsKey(group ?? ""))
            {
                var lst = d[group ?? ""];
                int i = lst.FindIndex(x => x.Equals(oldCat, StringComparison.OrdinalIgnoreCase));
                if (i >= 0) { lst[i] = newCat.Trim(); SaveUserCats(); }
            }
            if (touched > 0) InvalidateCache();
            return touched;
        }

        /// <summary>删除二级分类：清单移除 + 该组条目 category 归「未分类」。返回受影响条目数。</summary>
        public static int DeleteGroupCategory(string group, string cat)
        {
            int touched = 0;
            foreach (var m in LoadAllItems())
            {
                if (m.GetGroup() != group) continue;
                if (m.GetCategory().Equals(cat, StringComparison.OrdinalIgnoreCase))
                {
                    SetCategory(m.Id, "未分类");
                    touched++;
                }
            }
            var d = LoadUserCats();
            if (d.ContainsKey(group ?? ""))
            {
                d[group ?? ""].RemoveAll(x => x.Equals(cat, StringComparison.OrdinalIgnoreCase));
                if (d[group ?? ""].Count == 0) d.Remove(group ?? "");
                SaveUserCats();
            }
            if (touched > 0) InvalidateCache();
            return touched;
        }

        // ---------------- 设置 ----------------

        public static string SettingsFile
        {
            get { return Path.Combine(RootDir, "settings.json"); }
        }

        public static Settings LoadSettings()
        {
            try
            {
                string path = SettingsFile;
                if (File.Exists(path))
                {
                    var ser = new DataContractJsonSerializer(typeof(Settings));
                    using (var fs = File.OpenRead(path))
                    {
                        return (Settings)ser.ReadObject(fs);
                    }
                }
            }
            catch { }
            return new Settings();
        }

        public static void SaveSettings(Settings settings)
        {
            try
            {
                var ser = new DataContractJsonSerializer(typeof(Settings));
                using (var fs = File.Create(SettingsFile))
                {
                    ser.WriteObject(fs, settings);
                }
            }
            catch { }
        }

        // ---------------- 去重检测 / 清理 ----------------

        /// <summary>计算 ShapeRange 的总宽高。</summary>
        private static void RangeSize(PowerPoint.ShapeRange range, out float w, out float h)
        {
            w = 0; h = 0;
            try
            {
                foreach (PowerPoint.Shape s in range)
                {
                    w += (float)s.Width;
                    h += (float)s.Height;
                }
            }
            catch { }
        }

        /// <summary>检测是否已有形状、尺寸完全一致的收藏（提示可能重复）。</summary>
        public static string FindDuplicate(PowerPoint.ShapeRange range)
        {
            try
            {
                float w, h;
                RangeSize(range, out w, out h);
                int count = range.Count;
                foreach (var meta in LoadAllItems())
                {
                    if (meta.Kind != "Shape") continue;
                    if (meta.ShapeCount != count) continue;
                    if (meta.Width > 0 && Math.Abs(meta.Width - w) < 2 && Math.Abs(meta.Height - h) < 2)
                    {
                        return meta.Name;
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>找出完全重复的收藏（同类型、同数量、尺寸一致），返回重复项列表（保留最早的）。</summary>
        public static List<LibraryMeta> FindExactDuplicates()
        {
            var dups = new List<LibraryMeta>();
            try
            {
                var all = LoadAllItems();
                var seen = new Dictionary<string, LibraryMeta>();
                foreach (var meta in all.OrderBy(m => m.CreatedAt))
                {
                    if (meta.Kind != "Shape") continue;
                    string key = meta.Kind + "|" + meta.ShapeCount + "|" + meta.Width.ToString("F1") + "|" + meta.Height.ToString("F1");
                    if (seen.ContainsKey(key))
                    {
                        dups.Add(meta);
                    }
                    else
                    {
                        seen[key] = meta;
                    }
                }
            }
            catch { }
            return dups;
        }

        // ---------------- 备份 / 恢复 ----------------

        /// <summary>把整个收藏库打包为 .zip 备份文件。</summary>
        public static bool BackupTo(string zipPath)
        {
            try
            {
                if (File.Exists(zipPath)) File.Delete(zipPath);
                System.IO.Compression.ZipFile.CreateFromDirectory(ItemsDir, zipPath,
                    System.IO.Compression.CompressionLevel.Optimal, false);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>从 .zip 备份恢复收藏库（覆盖现有 items 内容）。两阶段：先安全解压到临时目录并校验，
        /// 成功后再替换，避免 zip 损坏/恶意路径导致现有收藏丢失或路径穿越。</summary>
        public static bool RestoreFrom(string zipPath)
        {
            string tmp = null;
            try
            {
                tmp = Path.Combine(Path.GetTempPath(), "PPTShapeLibrary_restore_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tmp);
                string itemsRoot = Path.GetFullPath(tmp) + Path.DirectorySeparatorChar;

                using (var fs = File.OpenRead(zipPath))
                using (var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read))
                {
                    foreach (var entry in zip.Entries)
                    {
                        if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue; // 目录项
                        string rel = entry.FullName.Replace('\\', '/');
                        // 安全校验：拒绝绝对路径与 .. 穿越
                        string full = Path.GetFullPath(Path.Combine(tmp, rel.Replace('/', Path.DirectorySeparatorChar)));
                        if (!full.StartsWith(itemsRoot, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("备份包含非法路径：" + entry.FullName);
                        Directory.CreateDirectory(Path.GetDirectoryName(full));
                        using (var es = entry.Open())
                        using (var outFs = File.Create(full))
                        {
                            es.CopyTo(outFs);
                        }
                    }
                }

                // 校验成功：清空现有 items 并替换
                string items = ItemsDir;
                if (Directory.Exists(items))
                {
                    foreach (var dir in Directory.GetDirectories(items))
                    {
                        try { Directory.Delete(dir, true); } catch { }
                    }
                }
                Directory.CreateDirectory(items);
                foreach (var sub in Directory.GetDirectories(tmp))
                {
                    string dest = Path.Combine(items, Path.GetFileName(sub));
                    if (Directory.Exists(dest)) try { Directory.Delete(dest, true); } catch { }
                    Directory.Move(sub, dest);
                }
                InvalidateCache();
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                try { if (tmp != null && Directory.Exists(tmp)) Directory.Delete(tmp, true); } catch { }
            }
        }
    }
}
