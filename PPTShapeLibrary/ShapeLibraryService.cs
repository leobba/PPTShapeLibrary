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
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "素材库");
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

        public static void DeleteItem(string id)
        {
            string dir = Path.Combine(ItemsDir, id);
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, true);
            }
            InvalidateCache();
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
