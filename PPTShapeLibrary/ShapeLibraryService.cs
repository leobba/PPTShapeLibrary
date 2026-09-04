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
        [DataMember(Name = "tags")]
        public string Tags { get; set; }
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

        public string[] GetTags()
        {
            if (string.IsNullOrWhiteSpace(Tags)) return new string[0];
            return Tags.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries);
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

        /// <summary>软删除：把条目移入回收站（可恢复）。</summary>
        public static void DeleteItem(string id)
        {
            string dir = Path.Combine(ItemsDir, id);
            if (!Directory.Exists(dir)) return;
            string trash = Path.Combine(TrashDir, id);
            try
            {
                if (Directory.Exists(trash)) Directory.Delete(trash, true);
                Directory.Move(dir, trash);
            }
            catch
            {
                // 移动失败（占用等）则直接删除，避免残留
                try { Directory.Delete(dir, true); } catch { }
            }
            InvalidateCache();
        }

        /// <summary>彻底删除（回收站内清除）。</summary>
        public static void HardDelete(string id)
        {
            string dir = Path.Combine(TrashDir, id);
            if (Directory.Exists(dir))
            {
                try { Directory.Delete(dir, true); } catch { }
            }
            // 也兜底 items 下残留
            string item = Path.Combine(ItemsDir, id);
            if (Directory.Exists(item))
            {
                try { Directory.Delete(item, true); } catch { }
            }
            InvalidateCache();
        }

        /// <summary>从回收站恢复条目到素材库。</summary>
        public static bool RestoreTrashItem(string id)
        {
            string trash = Path.Combine(TrashDir, id);
            if (!Directory.Exists(trash)) return false;
            string dest = Path.Combine(ItemsDir, id);
            if (Directory.Exists(dest))
            {
                // 目标冲突：改名恢复
                dest = Path.Combine(ItemsDir, NewId());
            }
            try
            {
                Directory.Move(trash, dest);
                InvalidateCache();
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
            meta.Name = newName;
            SaveMeta(meta);
            InvalidateCache();
        }

        public static void SaveStyleJson(string id, string json)
        {
            File.WriteAllText(StyleFile(id), json, Encoding.UTF8);
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

        public static void SetTags(string id, string tags)
        {
            var meta = LoadMeta(id);
            if (meta == null) return;
            meta.Tags = tags == null ? "" : tags.Trim();
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

        public static List<string> GetAllTags()
        {
            var list = new List<string>();
            foreach (var meta in LoadAllItems())
            {
                foreach (var t in meta.GetTags())
                {
                    if (!list.Contains(t)) list.Add(t);
                }
            }
            return list.OrderBy(t => t).ToList();
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

        /// <summary>标签使用计数（供标签管理器显示）。</summary>
        public static Dictionary<string, int> GetTagCounts()
        {
            var counts = new Dictionary<string, int>();
            foreach (var meta in LoadAllItems())
            {
                foreach (var t in meta.GetTags())
                {
                    counts.TryGetValue(t, out int n);
                    counts[t] = n + 1;
                }
            }
            return counts;
        }

        /// <summary>全局重命名标签：把所有收藏中的 oldTag 替换为 newTag（已存在则合并）。</summary>
        public static int RenameTagAll(string oldTag, string newTag)
        {
            if (string.IsNullOrWhiteSpace(oldTag) || string.IsNullOrWhiteSpace(newTag)) return 0;
            if (oldTag == newTag) return 0;
            int touched = 0;
            foreach (var meta in LoadAllItems())
            {
                var tags = meta.GetTags().ToList();
                if (tags.Contains(oldTag))
                {
                    tags.Remove(oldTag);
                    if (!tags.Contains(newTag)) tags.Add(newTag);
                    SetTags(meta.Id, string.Join(",", tags));
                    touched++;
                }
            }
            if (touched > 0) InvalidateCache();
            return touched;
        }

        /// <summary>全局删除标签：从所有收藏中移除。返回受影响条目数。</summary>
        public static int DeleteTagAll(string tag)
        {
            if (string.IsNullOrWhiteSpace(tag)) return 0;
            int touched = 0;
            foreach (var meta in LoadAllItems())
            {
                var tags = meta.GetTags().ToList();
                if (tags.Contains(tag))
                {
                    tags.Remove(tag);
                    SetTags(meta.Id, string.Join(",", tags));
                    touched++;
                }
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

        /// <summary>从 .zip 备份恢复收藏库（覆盖现有 items 内容）。</summary>
        public static bool RestoreFrom(string zipPath)
        {
            try
            {
                string items = ItemsDir;
                if (Directory.Exists(items))
                {
                    foreach (var dir in Directory.GetDirectories(items))
                    {
                        try { Directory.Delete(dir, true); } catch { }
                    }
                }
                Directory.CreateDirectory(items);
                System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, items);
                InvalidateCache();
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
