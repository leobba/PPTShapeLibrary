using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 素材包导出/导入（C 组新增功能）。
    /// 素材包 = 一个 .ppack.zip 文件，内部结构：
    ///   manifest.json   描述信息（name/exportedAt/版本）+ 条目列表
    ///   items/{id}/     —— 与本地收藏库 items 目录逐字节一致的条目目录（含 meta.json/shape.pptx/style.json/preview.png）
    /// 导入时如遇 id 冲突（本地已有同名 id 目录），为防覆盖本地数据，冲突条目重新生成 id 导入。
    /// </summary>
    public static class PackService
    {
        [DataContract]
        public class PackEntry
        {
            [DataMember(Name = "id")] public string Id { get; set; }
            [DataMember(Name = "name")] public string Name { get; set; }
            [DataMember(Name = "kind")] public string Kind { get; set; }
        }

        [DataContract]
        public class PackManifest
        {
            [DataMember(Name = "format")] public string Format { get; set; }
            /// <summary>规范版本：v1 初始 / v1.1 增加 author/license/sourceUrl/description（向后兼容，旧包缺失不影响导入）。</summary>
            [DataMember(Name = "specVersion")] public string SpecVersion { get; set; }
            [DataMember(Name = "name")] public string Name { get; set; }
            [DataMember(Name = "author")] public string Author { get; set; }
            [DataMember(Name = "license")] public string License { get; set; }
            [DataMember(Name = "sourceUrl")] public string SourceUrl { get; set; }
            [DataMember(Name = "description")] public string Description { get; set; }
            [DataMember(Name = "exportedAt")] public string ExportedAt { get; set; }
            [DataMember(Name = "items")] public List<PackEntry> Items { get; set; }
        }

        private const string PackFormat = "pptshape-pack-v1";
        private const string PackSpecVersion = "1.1";
        private const long MaxPackExpandedBytes = 512L * 1024 * 1024;
        private const long MaxPackEntryBytes = 256L * 1024 * 1024;
        private const int MaxPackEntries = 20004;

        /// <summary>导入后回传的包摘要（来源展示）。</summary>
        public class PackMetaOut
        {
            public string Name;
            public string Author;
            public string License;
            public string SourceUrl;
            public string Description;
            public string SpecVersion;
            public string ExportedAt;
        }
        private static readonly DataContractJsonSerializer ManifestSer = new DataContractJsonSerializer(typeof(PackManifest));

        /// <summary>把指定 id 集合打包为素材包 zip。返回是否成功，失败原因写 error。</summary>
        public static bool ExportPack(IEnumerable<string> ids, string zipPath, out string error)
        {
            string localError = null;
            bool ok = ShapeLibraryService.WithLibraryMutationLock(() => ExportPackCore(ids, zipPath, out localError));
            error = localError;
            return ok;
        }

        private static bool ExportPackCore(IEnumerable<string> ids, string zipPath, out string error)
        {
            error = null;
            var idList = (ids ?? new string[0]).Distinct().ToList();
            if (idList.Count == 0)
            {
                error = "未选择任何收藏";
                return false;
            }
            try
            {
                if (File.Exists(zipPath)) File.Delete(zipPath);
                using (var fs = File.Create(zipPath))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    var manifest = new PackManifest
                    {
                        Format = PackFormat,
                        SpecVersion = PackSpecVersion,
                        Name = Path.GetFileNameWithoutExtension(zipPath),
                        Author = ShapeLibraryService.LoadSettings().AuthorName,
                        License = "CC BY 4.0",   // 默认许可：可商用可再分发，署名即可（用户导出即授权）
                        ExportedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        Items = new List<PackEntry>()
                    };

                    foreach (string id in idList)
                    {
                        if (!ShapeLibraryService.IsSafeId(id)) continue;
                        string srcDir = Path.Combine(ShapeLibraryService.ItemsDir, id);
                        if (!Directory.Exists(srcDir)) continue;
                        var meta = ShapeLibraryService.LoadMeta(id);
                        if (meta == null) continue;

                        // 压缩 items/{id} 目录全部文件
                        foreach (string file in Directory.GetFiles(srcDir))
                        {
                            string entryName = "items/" + id + "/" + Path.GetFileName(file);
                            var entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
                            using (var es = entry.Open())
                            using (var inFs = File.OpenRead(file))
                            {
                                inFs.CopyTo(es);
                            }
                        }
                        manifest.Items.Add(new PackEntry { Id = id, Name = meta.Name, Kind = meta.Kind });
                    }

                    if (manifest.Items.Count == 0)
                    {
                        error = "选中的收藏均不存在或不可读";
                        return false;
                    }

                    // manifest.json 放 zip 根
                    var me = zip.CreateEntry("manifest.json");
                    using (var ms = new MemoryStream())
                    {
                        ManifestSer.WriteObject(ms, manifest);
                        ms.Position = 0;
                        using (var es = me.Open())
                        {
                            ms.CopyTo(es);
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                error = "导出素材包失败：" + ex.Message;
                return false;
            }
        }

        /// <summary>导入素材包 zip 到本地库。冲突 id 自动重写为新 id。返回导入条目数，conflicts 返回冲突重写数。
        /// meta 输出包的 manifest 摘要（作者/许可等，供 UI 展示来源）。</summary>
        public static int ImportPack(string zipPath, out int conflicts, out string error,
            out PackMetaOut meta)
        {
            int localConflicts = 0;
            string localError = null;
            PackMetaOut localMeta = null;
            int imported = ShapeLibraryService.WithLibraryMutationLock(
                () => ImportPackCore(zipPath, out localConflicts, out localError, out localMeta));
            conflicts = localConflicts;
            error = localError;
            meta = localMeta;
            return imported;
        }

        private static int ImportPackCore(string zipPath, out int conflicts, out string error,
            out PackMetaOut meta)
        {
            error = null;
            conflicts = 0;
            meta = new PackMetaOut();
            string stagingRoot = null;
            try
            {
                if (!File.Exists(zipPath))
                {
                    error = "素材包文件不存在";
                    return 0;
                }
                // Include orphaned/corrupt item folders as occupied too; LoadAllItems only
                // returns readable metadata and would otherwise allow an import to overwrite them.
                var existingIds = new HashSet<string>(ShapeLibraryService.LoadAllItems().Select(m => m.Id), StringComparer.OrdinalIgnoreCase);
                if (Directory.Exists(ShapeLibraryService.ItemsDir))
                {
                    foreach (string dir in Directory.GetDirectories(ShapeLibraryService.ItemsDir))
                    {
                        string folderId = Path.GetFileName(dir);
                        if (ShapeLibraryService.IsSafeId(folderId)) existingIds.Add(folderId);
                    }
                }
                stagingRoot = Path.Combine(ShapeLibraryService.RootDir, ".pack_import_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(stagingRoot);
                var idMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);   // srcId -> destId

                using (var fs = File.OpenRead(zipPath))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    if (zip.Entries.Count > MaxPackEntries)
                        throw new InvalidDataException("素材包包含过多文件。");
                    long expandedBytes = 0;
                    foreach (var entry in zip.Entries)
                    {
                        if (entry.Length > MaxPackEntryBytes)
                            throw new InvalidDataException("素材包中的文件过大：" + entry.FullName);
                        expandedBytes += entry.Length;
                        if (expandedBytes > MaxPackExpandedBytes)
                            throw new InvalidDataException("素材包解压后超过允许大小。");
                    }

                    // 读 manifest（可选——即便缺失也按 items/ 目录结构导入）
                    var manifestEntry = zip.GetEntry("manifest.json");
                    PackManifest manifest = null;
                    if (manifestEntry != null)
                    {
                        if (manifestEntry.Length > 4L * 1024 * 1024)
                            throw new InvalidDataException("素材包清单过大。");
                        try
                        {
                            using (var ms = new MemoryStream())
                            {
                                using (var manifestStream = manifestEntry.Open())
                                {
                                    var buffer = new byte[81920];
                                    int read;
                                    while ((read = manifestStream.Read(buffer, 0, buffer.Length)) > 0)
                                    {
                                        if (ms.Length + read > 4L * 1024 * 1024)
                                            throw new InvalidDataException("素材包清单实际解压后过大。");
                                        ms.Write(buffer, 0, read);
                                    }
                                }
                                ms.Position = 0;
                                manifest = (PackManifest)ManifestSer.ReadObject(ms);
                            }
                        }
                        catch { manifest = null; }
                        if (manifest == null || !string.Equals(manifest.Format, PackFormat, StringComparison.Ordinal))
                            throw new InvalidDataException("素材包清单无效或格式不受支持。");
                    }
                    if (manifest != null)
                    {
                        if (manifest.Items == null)
                            throw new InvalidDataException("素材包清单缺少条目列表。");
                        meta.Name = manifest.Name;
                        meta.Author = manifest.Author;
                        meta.License = manifest.License;
                        meta.SourceUrl = manifest.SourceUrl;
                        meta.Description = manifest.Description;
                        meta.SpecVersion = manifest.SpecVersion;
                        meta.ExportedAt = manifest.ExportedAt;
                    }
                    var validIds = manifest != null && manifest.Items != null
                        ? new HashSet<string>(manifest.Items.Where(i => i != null).Select(i => i.Id), StringComparer.OrdinalIgnoreCase)
                        : null;
                    var seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    long extractedBytes = 0;

                    // 第一遍：遍历 zip 内 items/{id}/{file} 条目，确定 id 映射并落盘文件
                    foreach (var entry in zip.Entries)
                    {
                        string path = entry.FullName.Replace('\\', '/');
                        if (!path.StartsWith("items/", StringComparison.OrdinalIgnoreCase)) continue;
                        string[] parts = path.Split('/');
                        // 必须恰为三段 items/{id}/{file}；多余段（含 .. 穿越/嵌套目录）一律拒绝
                        if (parts.Length != 3) throw new InvalidDataException("素材包包含非法条目：" + entry.FullName);
                        if (!seenPaths.Add(path)) throw new InvalidDataException("素材包包含重复文件：" + entry.FullName);
                        string srcId = parts[1];
                        string fileName = parts[2];
                        // srcId 白名单 + fileName 不得含分隔符或 .. 或空
                        if (!ShapeLibraryService.IsSafeId(srcId)) throw new InvalidDataException("素材包包含非法 id：" + entry.FullName);
                        if (string.IsNullOrWhiteSpace(fileName) || fileName == "." || fileName == ".."
                            || fileName.IndexOf('/') >= 0 || fileName.IndexOf('\\') >= 0 || fileName.IndexOf(':') >= 0)
                            throw new InvalidDataException("素材包包含非法文件名：" + entry.FullName);
                        if (!string.Equals(fileName, "meta.json", StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(fileName, "shape.pptx", StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(fileName, "style.json", StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(fileName, "preview.png", StringComparison.OrdinalIgnoreCase))
                            throw new InvalidDataException("素材包包含不支持的文件：" + entry.FullName);
                        if (validIds != null && !validIds.Contains(srcId)) continue; // 不在 manifest 的孤儿文件跳过

                        if (!idMap.ContainsKey(srcId))
                        {
                            string destId = srcId;
                            if (existingIds.Contains(destId))
                            {
                                do { destId = ShapeLibraryService.NewId(); }
                                while (existingIds.Contains(destId));
                                existingIds.Add(destId);
                                conflicts++;
                            }
                            else
                            {
                                existingIds.Add(destId);
                            }
                            idMap[srcId] = destId;
                        }

                        // 落盘前双重校验：最终完整路径必须位于 ItemsDir 下（防任何遗漏的穿越向量）
                        string itemsRoot = Path.GetFullPath(stagingRoot) + Path.DirectorySeparatorChar;
                        string destFile = Path.GetFullPath(Path.Combine(stagingRoot, idMap[srcId], fileName));
                        if (!destFile.StartsWith(itemsRoot, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidDataException("素材包条目越出库目录：" + entry.FullName);
                        Directory.CreateDirectory(Path.GetDirectoryName(destFile));
                        using (var es = entry.Open())
                        using (var outFs = File.Create(destFile))
                        {
                            var buffer = new byte[81920];
                            int read;
                            long entryBytes = 0;
                            while ((read = es.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                entryBytes += read;
                                extractedBytes += read;
                                if (entryBytes > MaxPackEntryBytes || extractedBytes > MaxPackExpandedBytes)
                                    throw new InvalidDataException("素材包实际解压大小超过允许限制：" + entry.FullName);
                                outFs.Write(buffer, 0, read);
                            }
                        }
                    }

                    // 第二遍：对冲突重写（srcId != destId）的条目，把 meta.json 里的 "id" 字段文本改写为 destId
                    foreach (var kv in idMap)
                    {
                        if (kv.Key == kv.Value) continue;
                        string metaPath = Path.Combine(stagingRoot, kv.Value, "meta.json");
                        if (File.Exists(metaPath)) RewriteMetaIdText(metaPath, kv.Value);
                    }
                }

                // Publish the package only after every entry has been extracted successfully.
                // If any directory move fails, move already-published entries back to staging.
                var committed = new List<KeyValuePair<string, string>>(); // destination -> staged source
                try
                {
                    foreach (var kv in idMap)
                    {
                        string stagedDir = Path.Combine(stagingRoot, kv.Value);
                        if (!Directory.Exists(stagedDir)) continue;
                        string destDir = Path.Combine(ShapeLibraryService.ItemsDir, kv.Value);
                        if (Directory.Exists(destDir))
                            throw new IOException("导入期间素材条目发生冲突：" + kv.Value);
                        Directory.Move(stagedDir, destDir);
                        committed.Add(new KeyValuePair<string, string>(destDir, stagedDir));
                    }
                }
                catch
                {
                    bool rollbackComplete = true;
                    for (int i = committed.Count - 1; i >= 0; i--)
                    {
                        try { Directory.Move(committed[i].Key, committed[i].Value); }
                        catch (Exception rollbackError)
                        {
                            rollbackComplete = false;
                            Log.Error(rollbackError, "ImportPack rollback");
                        }
                    }
                    if (!rollbackComplete)
                    {
                        Log.Write("ImportPack rollback incomplete; recovery data retained at " + stagingRoot);
                        stagingRoot = null; // preserve staged files for recovery
                    }
                    if (committed.Count > 0) ShapeLibraryService.InvalidateCachePublic();
                    throw;
                }

                ShapeLibraryService.InvalidateCachePublic();

                // 统计真正新入库的条目（导出前不存在、导入后 meta.json 可读）
                int imported = 0;
                foreach (var kv in idMap)
                {
                    string metaPath = Path.Combine(ShapeLibraryService.ItemsDir, kv.Value, "meta.json");
                    if (File.Exists(metaPath)) imported++;
                }
                return imported;
            }
            catch (Exception ex)
            {
                error = "导入素材包失败：" + ex.Message;
                return 0;
            }
            finally
            {
                if (stagingRoot != null && Directory.Exists(stagingRoot))
                {
                    try { Directory.Delete(stagingRoot, true); }
                    catch (Exception cleanupError) { Log.Error(cleanupError, "ImportPack staging cleanup"); }
                }
            }
        }

        /// <summary>直接把 meta.json 中的 "id":"..." 字段改写为新值（避免依赖反序列化-再落盘的时序问题）。</summary>
        private static void RewriteMetaIdText(string metaFile, string newId)
        {
            try
            {
                string text = File.ReadAllText(metaFile, Encoding.UTF8);
                // DataContractJsonSerializer 输出形如 {"id":"xxxx",...}；替换首个 id 字段（声明顺序 id 排第一）
                var rx = new System.Text.RegularExpressions.Regex("\"id\"\\s*:\\s*\"[^\"]*\"");
                var m = rx.Match(text);
                if (m.Success)
                {
                    text = text.Substring(0, m.Index)
                        + "\"id\":\"" + newId + "\""
                        + text.Substring(m.Index + m.Length);
                    File.WriteAllText(metaFile, text, Encoding.UTF8);
                }
            }
            catch { }
        }
    }
}
