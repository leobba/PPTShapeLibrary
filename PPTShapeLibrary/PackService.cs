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
            [DataMember(Name = "name")] public string Name { get; set; }
            [DataMember(Name = "exportedAt")] public string ExportedAt { get; set; }
            [DataMember(Name = "items")] public List<PackEntry> Items { get; set; }
        }

        private const string PackFormat = "pptshape-pack-v1";
        private static readonly DataContractJsonSerializer ManifestSer = new DataContractJsonSerializer(typeof(PackManifest));

        /// <summary>把指定 id 集合打包为素材包 zip。返回是否成功，失败原因写 error。</summary>
        public static bool ExportPack(IEnumerable<string> ids, string zipPath, out string error)
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
                        Name = Path.GetFileNameWithoutExtension(zipPath),
                        ExportedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        Items = new List<PackEntry>()
                    };

                    foreach (string id in idList)
                    {
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

        /// <summary>导入素材包 zip 到本地库。冲突 id 自动重写为新 id。返回导入条目数。</summary>
        public static int ImportPack(string zipPath, out string error)
        {
            error = null;
            try
            {
                if (!File.Exists(zipPath))
                {
                    error = "素材包文件不存在";
                    return 0;
                }
                var existingIds = new HashSet<string>(ShapeLibraryService.LoadAllItems().Select(m => m.Id));
                var idMap = new Dictionary<string, string>();   // srcId -> destId

                using (var fs = File.OpenRead(zipPath))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    // 读 manifest（可选——即便缺失也按 items/ 目录结构导入）
                    var manifestEntry = zip.GetEntry("manifest.json");
                    PackManifest manifest = null;
                    if (manifestEntry != null)
                    {
                        try
                        {
                            using (var ms = new MemoryStream())
                            {
                                manifestEntry.Open().CopyTo(ms);
                                ms.Position = 0;
                                manifest = (PackManifest)ManifestSer.ReadObject(ms);
                            }
                        }
                        catch { manifest = null; }
                    }
                    var validIds = manifest != null && manifest.Items != null
                        ? new HashSet<string>(manifest.Items.Select(i => i.Id))
                        : null;

                    // 第一遍：遍历 zip 内 items/{id}/ 条目，确定 id 映射并落盘文件
                    foreach (var entry in zip.Entries)
                    {
                        string path = entry.FullName.Replace('\\', '/');
                        if (!path.StartsWith("items/", StringComparison.OrdinalIgnoreCase)) continue;
                        string[] parts = path.Split('/');
                        if (parts.Length < 3) continue;              // items/{id}/{file}
                        string srcId = parts[1];
                        string fileName = parts[2];
                        if (string.IsNullOrWhiteSpace(srcId) || string.IsNullOrWhiteSpace(fileName)) continue;
                        if (validIds != null && !validIds.Contains(srcId)) continue; // 不在 manifest 的孤儿文件跳过

                        if (!idMap.ContainsKey(srcId))
                        {
                            string destId = srcId;
                            if (existingIds.Contains(destId))
                            {
                                do { destId = ShapeLibraryService.NewId(); }
                                while (existingIds.Contains(destId));
                                existingIds.Add(destId);
                            }
                            else
                            {
                                existingIds.Add(destId);
                            }
                            idMap[srcId] = destId;
                        }

                        string destFile = Path.Combine(ShapeLibraryService.ItemsDir, idMap[srcId], fileName);
                        Directory.CreateDirectory(Path.GetDirectoryName(destFile));
                        using (var es = entry.Open())
                        using (var outFs = File.Create(destFile))
                        {
                            es.CopyTo(outFs);
                        }
                    }

                    // 第二遍：对冲突重写（srcId != destId）的条目，把 meta.json 里的 "id" 字段文本改写为 destId
                    foreach (var kv in idMap)
                    {
                        if (kv.Key == kv.Value) continue;
                        string metaPath = Path.Combine(ShapeLibraryService.ItemsDir, kv.Value, "meta.json");
                        if (File.Exists(metaPath)) RewriteMetaIdText(metaPath, kv.Value);
                    }
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
