using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
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
        [DataMember(Name = "snapshotPage", EmitDefaultValue = false)]
        public bool SnapshotPage { get; set; }       // true=整页快照（含背景/版式的整页复制，以页为单位插入）
        [DataMember(Name = "colorHex", EmitDefaultValue = false)]
        public string ColorHex { get; set; }         // 主色 "RRGGBB"（收藏时取填充/渐变首色，供颜色筛选）

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
        private const string LibraryMutationMutexName = @"Local\PPTShapeLibrary.LibraryMutation";

        /// <summary>跨 PowerPoint 进程串行化素材库的复合写操作。</summary>
        public static T WithLibraryMutationLock<T>(Func<T> operation)
        {
            if (operation == null) throw new ArgumentNullException("operation");
            using (var mutex = new Mutex(false, LibraryMutationMutexName))
            {
                bool acquired = false;
                try
                {
                    try { acquired = mutex.WaitOne(TimeSpan.FromSeconds(30)); }
                    catch (AbandonedMutexException) { acquired = true; }
                    if (!acquired) throw new TimeoutException("素材库正被另一个 PowerPoint 实例占用，请稍后重试。");
                    return operation();
                }
                finally
                {
                    if (acquired) mutex.ReleaseMutex();
                }
            }
        }

        /// <summary>库结构即将变更（删除/恢复/覆盖）时触发：用于关闭"库文稿句柄池"，
        /// 避免 PowerPoint 仍持有 shape.pptx 导致文件/目录移动删除失败。</summary>
        public static event Action LibraryWillMutate;

        internal static void RaiseWillMutate()
        {
            var h = LibraryWillMutate;
            if (h == null) return;
            try { h(); } catch { }
        }
        // ---------------- 浏览上下文（按宿主隔离：每个面板/弹窗各自一份；收藏落点用最近活跃宿主） ----------------

        /// <summary>单个宿主的浏览状态。</summary>
        public class BrowseState
        {
            public string Group = "全部";
            public string Category;        // null=未选分类
            public string Kind = "全部";
            public string Filter = "";
            public string Sort = "最新";
            public bool Set;               // 是否浏览过（未浏览过则 init 不强制恢复）
        }

        private static readonly object BrowseSync = new object();
        private static readonly Dictionary<string, BrowseState> _browseByHost = new Dictionary<string, BrowseState>();
        private static string _activeHostKey;   // 最近与用户交互的宿主（收藏落点取它的上下文）

        /// <summary>前端浏览到某 组/分类/类型/搜索/排序 时登记（list 请求携带；按宿主隔离避免多面板互相覆盖）。</summary>
        public static void SetBrowseContext(string hostKey, string group, string category, string kind, string filter, string sort)
        {
            if (string.IsNullOrEmpty(hostKey)) hostKey = "default";
            lock (BrowseSync)
            {
                BrowseState st;
                if (!_browseByHost.TryGetValue(hostKey, out st))
                {
                    st = new BrowseState();
                    _browseByHost[hostKey] = st;
                }
                st.Set = true;
                st.Group = string.IsNullOrWhiteSpace(group) ? "全部" : group.Trim();
                st.Category = (string.IsNullOrWhiteSpace(category) || category == "全部") ? null : category.Trim();
                st.Kind = string.IsNullOrWhiteSpace(kind) ? "全部" : kind.Trim();
                st.Filter = filter ?? "";
                st.Sort = string.IsNullOrWhiteSpace(sort) ? "最新" : sort.Trim();
                _activeHostKey = hostKey;
            }
        }

        private static BrowseState ActiveState()
        {
            lock (BrowseSync)
            {
                BrowseState st;
                if (_activeHostKey != null && _browseByHost.TryGetValue(_activeHostKey, out st)) return st;
                return null;
            }
        }

        private static BrowseState StateOf(string hostKey)
        {
            lock (BrowseSync)
            {
                BrowseState st;
                if (hostKey != null && _browseByHost.TryGetValue(hostKey, out st)) return st;
                return null;
            }
        }

        /// <summary>收藏应写入的目标组（浏览"全部"/未浏览时回落「我的」）。</summary>
        public static string BrowseGroupForCollect
        {
            get
            {
                var st = ActiveState();
                if (st == null || !st.Set || string.IsNullOrWhiteSpace(st.Group) || st.Group == "全部") return "我的";
                return st.Group;
            }
        }

        /// <summary>收藏应写入的目标分类（当前浏览分类；未明确则 null 回落后由收藏方填默认）。</summary>
        public static string BrowseCategoryForCollect
        {
            get { var st = ActiveState(); return st != null ? st.Category : null; }
        }

        /// <summary>指定宿主页面重载后恢复：是否已浏览过。</summary>
        public static bool BrowseWasSet(string hostKey)
        {
            var st = StateOf(hostKey);
            return st != null && st.Set;
        }

        /// <summary>指定宿主重载后应恢复到的组（原样，可能"全部"）。</summary>
        public static string BrowseGroupForRestore(string hostKey)
        {
            var st = StateOf(hostKey);
            return (st == null || string.IsNullOrWhiteSpace(st.Group)) ? "全部" : st.Group;
        }

        public static string BrowseCategoryForRestore(string hostKey)
        {
            var st = StateOf(hostKey);
            return st != null ? st.Category : null;
        }

        public static string BrowseKindForRestore(string hostKey)
        {
            var st = StateOf(hostKey);
            return (st == null || string.IsNullOrWhiteSpace(st.Kind)) ? "全部" : st.Kind;
        }

        public static string BrowseFilterForRestore(string hostKey)
        {
            var st = StateOf(hostKey);
            return st != null ? (st.Filter ?? "") : "";
        }

        public static string BrowseSortForRestore(string hostKey)
        {
            var st = StateOf(hostKey);
            return (st == null || string.IsNullOrWhiteSpace(st.Sort)) ? "最新" : st.Sort;
        }

        /// <summary>重置浏览上下文（新会话兜底）。</summary>
        public static void ResetBrowseContext()
        {
            lock (BrowseSync)
            {
                _browseByHost.Clear();
                _activeHostKey = null;
            }
        }

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
            if (!IsSafeId(id)) throw new System.ArgumentException("素材条目标识无效。", "id");
            return WithLibraryMutationLock(() =>
            {
                string dir = Path.Combine(ItemsDir, id);
                Directory.CreateDirectory(dir);
                return dir;
            });
        }

        public static string NewId()
        {
            return WithLibraryMutationLock(() =>
            {
                string id;
                do
                {
                    id = DateTime.Now.ToString("yyyyMMdd-HHmmss-") + Guid.NewGuid().ToString("N").Substring(0, 12);
                }
                while (Directory.Exists(Path.Combine(ItemsDir, id)) || Directory.Exists(Path.Combine(TrashDir, id)));
                return id;
            });
        }

        private static readonly System.Text.RegularExpressions.Regex IdSafeRegex =
            new System.Text.RegularExpressions.Regex("^[0-9A-Za-z_\\-]+$");

        /// <summary>安全 id 校验：id 只允许字母数字下划线连字符（NewId 生成格式），杜绝路径穿越（..、盘符、分隔符）。</summary>
        public static bool IsSafeId(string id)
        {
            if (string.IsNullOrEmpty(id) || id.Length > 80 || !IdSafeRegex.IsMatch(id)) return false;
            // Windows reserves these path components for device names, even without extensions.
            string upper = id.ToUpperInvariant();
            if (upper == "CON" || upper == "PRN" || upper == "AUX" || upper == "NUL") return false;
            if (upper.Length == 4 && (upper.StartsWith("COM", StringComparison.Ordinal) || upper.StartsWith("LPT", StringComparison.Ordinal))
                && upper[3] >= '1' && upper[3] <= '9') return false;
            return true;
        }

        public static string ShapeFile(string id)
        {
            return ItemFile(id, "shape.pptx");
        }

        public static string StyleFile(string id)
        {
            return ItemFile(id, "style.json");
        }

        public static string PreviewFile(string id)
        {
            return ItemFile(id, "preview.png");
        }

        public static string MetaFile(string id)
        {
            return ItemFile(id, "meta.json");
        }

        private static string ItemFile(string id, string fileName)
        {
            if (!IsSafeId(id)) throw new System.ArgumentException("素材条目标识无效。", "id");
            return Path.Combine(ItemsDir, id, fileName);
        }

        private static void WriteFileAtomically(string path, Action<string> writeTemp)
        {
            string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                writeTemp(temp);
                if (File.Exists(path)) File.Replace(temp, path, null);
                else File.Move(temp, path);
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { }
            }
        }

        public static void SaveMeta(LibraryMeta meta)
        {
            if (meta == null || !IsSafeId(meta.Id)) throw new System.ArgumentException("素材条目标识无效。", "meta");
            WithLibraryMutationLock(() =>
            {
                string dir = Path.Combine(ItemsDir, meta.Id);
                Directory.CreateDirectory(dir);
                string path = MetaFile(meta.Id);
                string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    var ser = new DataContractJsonSerializer(typeof(LibraryMeta));
                    using (var fs = File.Create(temp)) ser.WriteObject(fs, meta);
                    if (File.Exists(path)) File.Replace(temp, path, null);
                    else File.Move(temp, path);
                    InvalidateCache();
                }
                finally
                {
                    try { if (File.Exists(temp)) File.Delete(temp); } catch { }
                }
                return true;
            });
        }

        public static LibraryMeta LoadMeta(string id)
        {
            if (!IsSafeId(id)) return null;
            string path = MetaFile(id);
            if (!File.Exists(path)) return null;
            try
            {
                var ser = new DataContractJsonSerializer(typeof(LibraryMeta));
                using (var fs = File.OpenRead(path))
                {
                    var meta = (LibraryMeta)ser.ReadObject(fs);
                    // The folder name is the trusted identity. Never allow imported metadata
                    // to redirect later file operations to a different library path.
                    if (meta != null) meta.Id = id;
                    return meta;
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
                        if (!IsSafeId(id)) continue;
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
            return WithLibraryMutationLock(() => DeleteItemCore(id));
        }

        private static bool DeleteItemCore(string id)
        {
            if (!IsSafeId(id)) { Log.Error(new Exception("DeleteItem: unsafe id=" + (id ?? "")), "Security"); return false; }
            RaiseWillMutate();   // 关闭库文稿句柄，避免文件被占用导致移动失败
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
            WithLibraryMutationLock(() =>
            {
                HardDeleteCore(id);
                return true;
            });
        }

        private static void HardDeleteCore(string id)
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
            return WithLibraryMutationLock(() => RestoreTrashItemCore(id));
        }

        private static bool RestoreTrashItemCore(string id)
        {
            if (!IsSafeId(id)) { Log.Error(new Exception("RestoreTrashItem: unsafe id=" + (id ?? "")), "Security"); return false; }
            RaiseWillMutate();
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
            return WithLibraryMutationLock(() => PurgeTrashCore());
        }

        private static int PurgeTrashCore()
        {
            RaiseWillMutate();
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
            WithLibraryMutationLock(() =>
            {
                var meta = LoadMeta(id);
                if (meta == null || string.IsNullOrWhiteSpace(newName)) return true;
                string name = newName.Trim();
                if (name.Length > 120) name = name.Substring(0, 120);
                meta.Name = name;
                SaveMeta(meta);
                InvalidateCache();
                return true;
            });
        }

        public static void SaveStyleJson(string id, string json)
        {
            WithLibraryMutationLock(() =>
            {
                string path = StyleFile(id);
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                WriteFileAtomically(path, temp => File.WriteAllText(temp, json ?? "", Encoding.UTF8));
                return true;
            });
        }

        /// <summary>孤儿目录清理：把 items 下没有 meta.json（收藏流程中断残留）的目录移入回收站，返回清理数。
        /// 不硬删，用户仍可在回收站查看/彻底删除。</summary>
        public static int CollectOrphanDirs()
        {
            return WithLibraryMutationLock(() => CollectOrphanDirsCore());
        }

        private static int CollectOrphanDirsCore()
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
            WithLibraryMutationLock(() =>
            {
                string path = PreviewFile(id);
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                WriteFileAtomically(path, temp => File.WriteAllBytes(temp, png ?? new byte[0]));
                return true;
            });
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
            WithLibraryMutationLock(() =>
            {
                var meta = LoadMeta(id);
                if (meta == null) return true;
                meta.Category = string.IsNullOrWhiteSpace(category) ? "未分类" : category.Trim();
                SaveMeta(meta);
                InvalidateCache();
                return true;
            });
        }

        /// <summary>记录最近使用时间与使用次数。</summary>
        public static void TouchUsed(string id)
        {
            WithLibraryMutationLock(() =>
            {
                var meta = LoadMeta(id);
                if (meta == null) return true;
                meta.LastUsedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                meta.UseCount++;
                SaveMeta(meta);
                InvalidateCache();
                return true;
            });
        }

        public static void SetSortOrder(string id, int order)
        {
            WithLibraryMutationLock(() =>
            {
                var meta = LoadMeta(id);
                if (meta == null) return true;
                meta.SortOrder = order;
                SaveMeta(meta);
                InvalidateCache();
                return true;
            });
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
        private static DateTime _userGroupsWriteTimeUtc = DateTime.MinValue;

        public static List<string> GetUserGroups()
        {
            DateTime writeTime = File.Exists(UserGroupsFile) ? File.GetLastWriteTimeUtc(UserGroupsFile) : DateTime.MinValue;
            if (_userGroups == null || writeTime != _userGroupsWriteTimeUtc)
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
                _userGroupsWriteTimeUtc = File.Exists(UserGroupsFile) ? File.GetLastWriteTimeUtc(UserGroupsFile) : DateTime.MinValue;
            }
            return new List<string>(_userGroups);
        }

        private static void SaveUserGroups()
        {
            WithLibraryMutationLock(() =>
            {
                try
                {
                    var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                    WriteFileAtomically(UserGroupsFile, temp => File.WriteAllText(temp, ser.Serialize(_userGroups ?? new List<string>())));
                    _userGroupsWriteTimeUtc = File.GetLastWriteTimeUtc(UserGroupsFile);
                }
                catch { }
                return true;
            });
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
            WithLibraryMutationLock(() =>
            {
                var meta = LoadMeta(id);
                if (meta == null) return true;
                meta.Group = string.IsNullOrWhiteSpace(group) ? LibraryMeta.DefaultGroup : group.Trim();
                SaveMeta(meta);
                InvalidateCache();
                return true;
            });
        }

        /// <summary>新建自定义组。返回 false 表示重名或非法。</summary>
        public static bool CreateUserGroup(string name)
        {
            return WithLibraryMutationLock(() => CreateUserGroupCore(name));
        }

        private static bool CreateUserGroupCore(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            name = name.Trim();
            if (PresetGroups.Contains(name)) return false;
            _userGroups = null; // reload while holding the cross-process lock
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
            WithLibraryMutationLock(() =>
            {
                RenameGroupCore(oldName, newName);
                return true;
            });
        }

        private static void RenameGroupCore(string oldName, string newName)
        {
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName)) return;
            newName = newName.Trim();
            if (PresetGroups.Contains(oldName)) return;          // 预设组不允许改名
            _userGroups = null;
            var ug = GetUserGroups();
            int i = ug.IndexOf(oldName);
            if (i < 0) return;
            ug[i] = newName;
            _userGroups = ug;
            SaveUserGroups();
            InvalidateCache();
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
            return WithLibraryMutationLock(() => DeleteGroupCore(name));
        }

        private static int DeleteGroupCore(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return 0;
            if (PresetGroups.Contains(name)) return 0;
            _userGroups = null;
            var ug = GetUserGroups();
            if (!ug.Remove(name)) return 0;
            _userGroups = ug;
            SaveUserGroups();
            int n = 0;
            InvalidateCache();
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
        private static DateTime _userCatsWriteTimeUtc = DateTime.MinValue;

        private static Dictionary<string, List<string>> LoadUserCats()
        {
            DateTime writeTime = File.Exists(UserCatsFile) ? File.GetLastWriteTimeUtc(UserCatsFile) : DateTime.MinValue;
            if (_userCats != null && writeTime == _userCatsWriteTimeUtc) return _userCats;
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
            _userCatsWriteTimeUtc = File.Exists(UserCatsFile) ? File.GetLastWriteTimeUtc(UserCatsFile) : DateTime.MinValue;
            return _userCats;
        }

        private static void SaveUserCats()
        {
            WithLibraryMutationLock(() =>
            {
                try
                {
                    var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                    WriteFileAtomically(UserCatsFile, temp => File.WriteAllText(temp, ser.Serialize(LoadUserCats())));
                    _userCatsWriteTimeUtc = File.GetLastWriteTimeUtc(UserCatsFile);
                }
                catch { }
                return true;
            });
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
            WithLibraryMutationLock(() =>
            {
                CreateGroupCategoryCore(group, cat);
                return true;
            });
        }

        private static void CreateGroupCategoryCore(string group, string cat)
        {
            if (string.IsNullOrWhiteSpace(group) || string.IsNullOrWhiteSpace(cat)) return;
            _userCats = null;
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
            return WithLibraryMutationLock(() => RenameGroupCategoryCore(group, oldCat, newCat));
        }

        private static int RenameGroupCategoryCore(string group, string oldCat, string newCat)
        {
            if (string.IsNullOrWhiteSpace(oldCat) || string.IsNullOrWhiteSpace(newCat)) return 0;
            if (oldCat == newCat) return 0;
            int touched = 0;
            InvalidateCache();
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
            _userCats = null;
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
            return WithLibraryMutationLock(() => DeleteGroupCategoryCore(group, cat));
        }

        private static int DeleteGroupCategoryCore(string group, string cat)
        {
            int touched = 0;
            InvalidateCache();
            foreach (var m in LoadAllItems())
            {
                if (m.GetGroup() != group) continue;
                if (m.GetCategory().Equals(cat, StringComparison.OrdinalIgnoreCase))
                {
                    SetCategory(m.Id, "未分类");
                    touched++;
                }
            }
            _userCats = null;
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
                        var s = (Settings)ser.ReadObject(fs);
                        // 布局默认迁移：旧默认(右侧300 / 左侧830) → 新默认(左侧400)，仅当用户从未自定义过(宽度仍为默认值)
                        bool oldDefaultW = s.PaneWidth <= 300 || s.PaneWidth == 830;
                        if (oldDefaultW && (s.PaneDock ?? "right") == "right")
                        {
                            s.PaneDock = "left";
                        }
                        if (oldDefaultW)
                        {
                            s.PaneWidth = 400;
                            SaveSettings(s);
                        }
                        return s;
                    }
                }
            }
            catch { }
            return new Settings();
        }

        public static void SaveSettings(Settings settings)
        {
            WithLibraryMutationLock(() =>
            {
                try
                {
                    var ser = new DataContractJsonSerializer(typeof(Settings));
                    WriteFileAtomically(SettingsFile, temp =>
                    {
                        using (var fs = File.Create(temp)) ser.WriteObject(fs, settings);
                    });
                }
                catch { }
                return true;
            });
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
            return WithLibraryMutationLock(() => BackupToCore(zipPath));
        }

        private static bool BackupToCore(string zipPath)
        {
            string tempZip = null;
            try
            {
                if (string.IsNullOrWhiteSpace(zipPath)) return false;
                string target = Path.GetFullPath(zipPath);
                string itemsRoot = Path.GetFullPath(ItemsDir).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    + Path.DirectorySeparatorChar;
                if (target.StartsWith(itemsRoot, StringComparison.OrdinalIgnoreCase)) return false;
                string targetDir = Path.GetDirectoryName(target);
                if (string.IsNullOrEmpty(targetDir)) return false;
                tempZip = Path.Combine(targetDir, ".PPTShapeLibrary_backup_" + Guid.NewGuid().ToString("N") + ".tmp");
                using (var fs = new FileStream(tempZip, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                using (var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create))
                {
                    foreach (string dir in Directory.GetDirectories(ItemsDir))
                    {
                        string id = Path.GetFileName(dir);
                        if (!IsSafeId(id)) continue;
                        var itemMeta = LoadMeta(id);
                        if (itemMeta == null) continue; // exclude incomplete/orphan entries
                        if (itemMeta.Kind == "Shape" && !File.Exists(Path.Combine(dir, "shape.pptx"))) continue;
                        if (itemMeta.Kind == "Style" && !File.Exists(Path.Combine(dir, "style.json"))) continue;
                        if (itemMeta.Kind != "Shape" && itemMeta.Kind != "Style") continue;

                        foreach (string file in Directory.GetFiles(dir))
                        {
                            string name = Path.GetFileName(file);
                            if (!string.Equals(name, "meta.json", StringComparison.OrdinalIgnoreCase)
                                && !string.Equals(name, "shape.pptx", StringComparison.OrdinalIgnoreCase)
                                && !string.Equals(name, "style.json", StringComparison.OrdinalIgnoreCase)
                                && !string.Equals(name, "preview.png", StringComparison.OrdinalIgnoreCase)) continue;
                            var entry = zip.CreateEntry(id + "/" + name, System.IO.Compression.CompressionLevel.Optimal);
                            using (var entryStream = entry.Open())
                            using (var input = File.OpenRead(file)) input.CopyTo(entryStream);
                        }
                    }
                }
                if (File.Exists(target)) File.Replace(tempZip, target, null);
                else File.Move(tempZip, target);
                tempZip = null;
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                try { if (tempZip != null && File.Exists(tempZip)) File.Delete(tempZip); } catch { }
            }
        }

        /// <summary>自动备份：当日尚未备份时把整库打包到 backups/backup-yyyyMMdd.zip，并只保留最近 7 份。
        /// 失败静默（记日志由调用方处理），不阻塞启动。</summary>
        public static void AutoBackupIfNeeded(int keep = 7)
        {
            try
            {
                string dir = Path.Combine(RootDir, "backups");
                Directory.CreateDirectory(dir);
                string today = DateTime.Now.ToString("yyyyMMdd");
                string target = Path.Combine(dir, "backup-" + today + ".zip");
                if (!File.Exists(target))
                {
                    if (!BackupTo(target)) return;
                }
                // 轮转：仅保留最近 keep 份
                var files = new DirectoryInfo(dir).GetFiles("backup-*.zip")
                    .OrderByDescending(f => f.Name).ToList();
                for (int i = keep; i < files.Count; i++)
                {
                    try { files[i].Delete(); } catch { }
                }
            }
            catch { }
        }

        /// <summary>从 .zip 备份恢复收藏库（覆盖现有 items 内容）。两阶段：先安全解压到临时目录并校验，
        /// 成功后再替换，避免 zip 损坏/恶意路径导致现有收藏丢失或路径穿越。</summary>
        public static bool RestoreFrom(string zipPath)
        {
            RaiseWillMutate();   // 覆盖 items 前先关闭库文稿句柄
            return WithLibraryMutationLock(() => RestoreFromCore(zipPath));
        }

        private static bool RestoreFromCore(string zipPath)
        {
            string tmp = null;
            string backup = null;
            try
            {
                // Keep staging beside items so the final directory renames stay on one volume.
                tmp = Path.Combine(RootDir, ".restore_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tmp);
                string itemsRoot = Path.GetFullPath(tmp) + Path.DirectorySeparatorChar;

                using (var fs = File.OpenRead(zipPath))
                using (var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read))
                {
                    if (zip.Entries.Count > 20004)
                        throw new InvalidOperationException("备份包含过多文件。");
                    long expandedBytes = 0;
                    long extractedBytes = 0;
                    var itemIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var entry in zip.Entries)
                    {
                        string rel = entry.FullName.Replace('\\', '/');
                        bool isDirectory = rel.EndsWith("/", StringComparison.Ordinal);
                        if (isDirectory) rel = rel.Substring(0, rel.Length - 1);
                        string[] parts = rel.Split('/');
                        // Backups contain only {id}/{file}; accepting arbitrary ZIP layouts can
                        // replace the library with an empty or unusable tree (for example a pack).
                        if (parts.Length != (isDirectory ? 1 : 2) || !IsSafeId(parts[0]))
                            throw new InvalidOperationException("备份包含非法条目：" + entry.FullName);
                        itemIds.Add(parts[0]);
                        if (entry.Length > 256L * 1024 * 1024)
                            throw new InvalidOperationException("备份中的文件过大：" + entry.FullName);
                        expandedBytes += entry.Length;
                        if (expandedBytes > 512L * 1024 * 1024)
                            throw new InvalidOperationException("备份解压后超过允许大小。");
                        if (isDirectory)
                        {
                            Directory.CreateDirectory(Path.Combine(tmp, parts[0]));
                            continue;
                        }
                        string fileName = parts[1];
                        if (!string.Equals(fileName, "meta.json", StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(fileName, "shape.pptx", StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(fileName, "style.json", StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(fileName, "preview.png", StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("备份包含不支持的文件：" + entry.FullName);
                        if (!seenPaths.Add(rel))
                            throw new InvalidOperationException("备份包含重复文件：" + entry.FullName);
                        if (string.Equals(fileName, "meta.json", StringComparison.OrdinalIgnoreCase)
                            && entry.Length > 4L * 1024 * 1024)
                            throw new InvalidOperationException("备份元数据过大：" + entry.FullName);

                        string full = Path.GetFullPath(Path.Combine(tmp, rel.Replace('/', Path.DirectorySeparatorChar)));
                        if (!full.StartsWith(itemsRoot, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("备份包含非法路径：" + entry.FullName);
                        Directory.CreateDirectory(Path.GetDirectoryName(full));
                        using (var es = entry.Open())
                        using (var outFs = File.Create(full))
                        {
                            var buffer = new byte[81920];
                            int read;
                            long entryBytes = 0;
                            while ((read = es.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                entryBytes += read;
                                extractedBytes += read;
                                if (entryBytes > 256L * 1024 * 1024 || extractedBytes > 512L * 1024 * 1024
                                    || (string.Equals(fileName, "meta.json", StringComparison.OrdinalIgnoreCase)
                                        && entryBytes > 4L * 1024 * 1024))
                                    throw new InvalidOperationException("备份实际解压大小超过允许限制：" + entry.FullName);
                                outFs.Write(buffer, 0, read);
                            }
                        }
                    }

                    foreach (string id in itemIds)
                    {
                        string itemDir = Path.Combine(tmp, id);
                        string metaPath = Path.Combine(itemDir, "meta.json");
                        var itemMeta = LoadMetaFromDir(metaPath);
                        if (itemMeta == null)
                            throw new InvalidOperationException("备份条目缺少有效元数据：" + id);
                        if (itemMeta.Kind == "Shape" && !File.Exists(Path.Combine(itemDir, "shape.pptx")))
                            throw new InvalidOperationException("备份形状条目缺少 shape.pptx：" + id);
                        if (itemMeta.Kind == "Style" && !File.Exists(Path.Combine(itemDir, "style.json")))
                            throw new InvalidOperationException("备份样式条目缺少 style.json：" + id);
                        if (itemMeta.Kind != "Shape" && itemMeta.Kind != "Style")
                            throw new InvalidOperationException("备份条目类型无效：" + id);
                        if (!string.Equals(itemMeta.Id, id, StringComparison.Ordinal))
                            RewriteMetaIdText(metaPath, id);
                    }
                }

                // Swap whole directories instead of deleting existing items one by one. If the
                // staged directory cannot be installed, restore the original directory intact.
                string items = ItemsDir;
                bool movedOriginal = false;
                backup = Path.Combine(RootDir, ".items_before_restore_" + Guid.NewGuid().ToString("N"));
                try
                {
                    if (Directory.Exists(items))
                    {
                        Directory.Move(items, backup);
                        movedOriginal = true;
                    }
                    Directory.Move(tmp, items);
                    tmp = null;
                }
                catch
                {
                    if (movedOriginal && !Directory.Exists(items) && Directory.Exists(backup))
                    {
                        try { Directory.Move(backup, items); }
                        catch (Exception rollbackError)
                        {
                            Log.Error(rollbackError, "RestoreFrom rollback");
                            Log.Write("RestoreFrom rollback failed; original library retained at " + backup);
                        }
                    }
                    throw;
                }

                InvalidateCache();
                if (movedOriginal && Directory.Exists(backup))
                {
                    try { Directory.Delete(backup, true); }
                    catch (Exception cleanupError) { Log.Error(cleanupError, "RestoreFrom old library cleanup"); }
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "RestoreFrom");
                return false;
            }
            finally
            {
                try { if (tmp != null && Directory.Exists(tmp)) Directory.Delete(tmp, true); } catch { }
            }
        }

        // ---------------- 工具预设记忆（键值对，持久化 toolprefs.json） ----------------

        private static string ToolPrefsFile
        {
            get { return Path.Combine(RootDir, "toolprefs.json"); }
        }
        private static Dictionary<string, string> _toolPrefs;
        private static DateTime _toolPrefsWriteTimeUtc = DateTime.MinValue;

        private static Dictionary<string, string> LoadToolPrefs()
        {
            DateTime writeTime = File.Exists(ToolPrefsFile) ? File.GetLastWriteTimeUtc(ToolPrefsFile) : DateTime.MinValue;
            if (_toolPrefs != null && writeTime == _toolPrefsWriteTimeUtc) return _toolPrefs;
            _toolPrefs = new Dictionary<string, string>();
            try
            {
                if (File.Exists(ToolPrefsFile))
                {
                    var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                    var raw = ser.Deserialize<Dictionary<string, object>>(File.ReadAllText(ToolPrefsFile));
                    if (raw != null)
                    {
                        foreach (var kv in raw)
                        {
                            if (kv.Value != null) _toolPrefs[kv.Key] = kv.Value.ToString();
                        }
                    }
                }
            }
            catch { }
            _toolPrefsWriteTimeUtc = File.Exists(ToolPrefsFile) ? File.GetLastWriteTimeUtc(ToolPrefsFile) : DateTime.MinValue;
            return _toolPrefs;
        }

        /// <summary>读取全部工具预设（前端一次取回回填）。</summary>
        public static Dictionary<string, string> GetToolPrefs()
        {
            return new Dictionary<string, string>(LoadToolPrefs());
        }

        /// <summary>写入/更新一个工具预设键。</summary>
        public static void SetToolPref(string key, string value)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            WithLibraryMutationLock(() =>
            {
                try
                {
                    _toolPrefs = null; // reread under the cross-process lock to preserve other processes' updates
                    var d = LoadToolPrefs();
                    if (value == null) d.Remove(key);
                    else d[key] = value;
                    var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                    WriteFileAtomically(ToolPrefsFile, temp => File.WriteAllText(temp, ser.Serialize(d)));
                    _toolPrefsWriteTimeUtc = File.GetLastWriteTimeUtc(ToolPrefsFile);
                }
                catch { }
                return true;
            });
        }
    }
}
