using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace PPTShapeLibraryInstaller
{
    /// <summary>一次更新检查的结果。</summary>
    internal class UpdateInfo
    {
        public bool Ok;                 // 检查本身是否成功（网络/解析失败时为 false）
        public bool FromCache;          // 结果来自本地缓存（24h 内不重复联网）
        public bool HasUpdate;          // 是否有比当前更新的版本
        public string CurrentVersion;   // 当前安装器/产品版本
        public string LatestVersion;    // 最新版本（规范化成 4 段）
        public string TagName;          // 原始 tag，如 v1.0.2.0
        public string ReleaseUrl;       // 发布页
        public string Notes;            // 发布说明（截断）
        public string AssetName;        // 分发包资产名
        public string AssetUrl;         // 分发包下载地址
        public string Sha256Url;        // 可选：同名 .sha256 资产
        public string Error;            // Ok=false 时的原因
    }

    /// <summary>
    /// GitHub Releases 更新检查 / 下载 / 应用。
    ///
    /// 设计要点（对应 skill「桌面应用 → Releases + 自建检查」）：
    ///  · 仓库必须**公开**，检查走 api.github.com，无需 token；
    ///  · 24 小时缓存，避免未认证 API 的 60 次/小时/IP 限流；
    ///  · 下载优先直连，失败按顺序回退 GitHub 加速站前缀（加速站只代理 github.com 域名）；
    ///  · 有 .sha256 资产就校验，校验不过拒绝安装；
    ///  · 应用更新 = 解压分发包 → 启动其中的安装器（由它负责卸载旧版 + 装新版）。
    /// </summary>
    internal static class UpdateChecker
    {
        public const string Owner = "leobba";
        public const string Repo = "PPTShapeLibrary";
        public static readonly string LatestApiUrl =
            "https://api.github.com/repos/" + Owner + "/" + Repo + "/releases/latest";
        /// <summary>更新清单的**固定地址**：GitHub 的 releases/latest/download/&lt;asset&gt; 会重定向到
        /// 最新 Release 的同名资产，**不经过 API、不受匿名限流影响**（60 次/小时/IP 很容易被共享出口用尽，
        /// 实测本机就撞上过 403）。发版时每个 Release 都上传同名 update.json 覆盖即可。</summary>
        public static readonly string UpdateJsonUrl =
            "https://github.com/" + Owner + "/" + Repo + "/releases/latest/download/update.json";
        public static readonly string ReleasesPageUrl =
            "https://github.com/" + Owner + "/" + Repo + "/releases";

        /// <summary>GitHub 加速站前缀，按顺序尝试；空串表示直连优先。
        /// 只对 github.com / objects.githubusercontent.com 的资源有效。</summary>
        public static readonly string[] MirrorPrefixes =
        {
            "",                             // 直连
            "https://ghproxy.net/",
            "https://gh-proxy.com/",
            "https://ghfast.top/",
        };

        private static readonly TimeSpan CacheTtl = TimeSpan.FromHours(24);

        /// <summary>
        /// "当前版本"取两处更准确的来源（传统 csproj 不会自动生成 AssemblyVersion，
        /// 所以不去读程序集版本）：
        ///   ① 已安装版本：HKCU/HKLM 卸载项里的 DisplayVersion；
        ///   ② 分发包自带版本：&lt;exe 目录&gt;\App\Application Files\PPTShapeLibrary_x_y_z_w\ 的目录名。
        /// 都没找到返回 0.0.0.0（此时"更新检查"退化为"看最新版本"）。
        /// </summary>
        public static string CurrentVersion()
        {
            string installed = InstalledVersion();
            if (!string.IsNullOrEmpty(installed)) return NormalizeVersion(installed);
            string bundled = BundledVersion();
            if (!string.IsNullOrEmpty(bundled)) return NormalizeVersion(bundled);
            return "0.0.0.0";
        }

        private static string InstalledVersion()
        {
            try
            {
                var roots = new[]
                {
                    Microsoft.Win32.Registry.CurrentUser,
                    Microsoft.Win32.Registry.LocalMachine,
                    Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine,
                        Microsoft.Win32.RegistryView.Registry32)
                };
                foreach (var hive in roots)
                {
                    using (var root = hive.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
                    {
                        if (root == null) continue;
                        foreach (var sub in root.GetSubKeyNames())
                        {
                            using (var k = root.OpenSubKey(sub))
                            {
                                if (k == null) continue;
                                var name = k.GetValue("DisplayName") as string;
                                if (string.IsNullOrEmpty(name)) continue;
                                if (name.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) < 0) continue;
                                var ver = k.GetValue("DisplayVersion") as string;
                                if (!string.IsNullOrEmpty(ver)) return ver;
                            }
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        private static string BundledVersion()
        {
            try
            {
                string appFiles = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "App", "Application Files");
                if (!Directory.Exists(appFiles)) return null;
                foreach (var d in Directory.GetDirectories(appFiles, "PPTShapeLibrary_*"))
                {
                    // 目录名形如 PPTShapeLibrary_1_0_1_0：版本段本身用下划线分隔，
                    // 所以要取"前缀之后全部"再换成点，不能取最后一个下划线之后（那会得到 "0"）。
                    string leaf = Path.GetFileName(d);
                    const string prefix = "PPTShapeLibrary_";
                    if (leaf.Length <= prefix.Length) continue;
                    if (!leaf.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                    string ver = leaf.Substring(prefix.Length).Replace('_', '.');
                    if (!string.IsNullOrEmpty(ver)) return ver;
                }
            }
            catch { }
            return null;
        }

        /// <summary>把 "v1.0.2" / "1.0.2.0" / "1.0.2-beta" 规范化成可比较的 4 段数字。</summary>
        public static string NormalizeVersion(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "0.0.0.0";
            string s = raw.Trim();
            if (s.StartsWith("v", StringComparison.OrdinalIgnoreCase)) s = s.Substring(1);
            int dash = s.IndexOf('-');
            if (dash > 0) s = s.Substring(0, dash);       // 预发布后缀不参与比较
            var parts = s.Split('.');
            var nums = new List<string>();
            for (int i = 0; i < 4; i++)
            {
                int n = 0;
                if (i < parts.Length) int.TryParse(parts[i], out n);
                nums.Add(n.ToString());
            }
            return string.Join(".", nums.ToArray());
        }

        private static int CompareVersions(string a, string b)
        {
            var pa = NormalizeVersion(a).Split('.');
            var pb = NormalizeVersion(b).Split('.');
            for (int i = 0; i < 4; i++)
            {
                int x = int.Parse(pa[i]), y = int.Parse(pb[i]);
                if (x != y) return x < y ? -1 : 1;
            }
            return 0;
        }

        // ---------------- 缓存 ----------------

        private static string DefaultCachePath
        {
            get
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PPTShapeLibrary");
                return Path.Combine(dir, "update-check.json");
            }
        }

        private class CacheRecord
        {
            public string CheckedAt { get; set; }
            public string TagName { get; set; }
            public string LatestVersion { get; set; }
            public string ReleaseUrl { get; set; }
            public string Notes { get; set; }
            public string AssetName { get; set; }
            public string AssetUrl { get; set; }
            public string Sha256Url { get; set; }
        }

        private static CacheRecord LoadCache(string path)
        {
            try
            {
                if (!File.Exists(path)) return null;
                var json = File.ReadAllText(path, Encoding.UTF8);
                var rec = new JavaScriptSerializer().Deserialize<CacheRecord>(json);
                if (rec == null || string.IsNullOrEmpty(rec.CheckedAt)) return null;
                return rec;
            }
            catch { return null; }
        }

        private static void SaveCache(string path, CacheRecord rec)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, new JavaScriptSerializer().Serialize(rec), Encoding.UTF8);
            }
            catch { }
        }

        // ---------------- 检查 ----------------

        /// <summary>检查更新。force=false 时 24h 内直接用缓存。
        /// 通道顺序：① Release 固定地址的 update.json（不走 API）→ ② GitHub API（兜底）。</summary>
        public static UpdateInfo Check(bool force)
        {
            string cachePath = DefaultCachePath;
            var info = new UpdateInfo { CurrentVersion = CurrentVersion() };

            var cached = LoadCache(cachePath);
            if (!force && cached != null)
            {
                DateTime at;
                if (DateTime.TryParse(cached.CheckedAt, out at) && DateTime.Now - at < CacheTtl)
                {
                    Fill(info, cached);
                    info.Ok = true;
                    info.FromCache = true;
                    return info;
                }
            }

            // ① update.json（首选：固定地址、无 API 限流）
            string viaErr;
            var rec = TryUpdateJson(out viaErr);
            if (rec != null)
            {
                SaveCache(cachePath, rec);
                Fill(info, rec);
                info.Ok = true;
                return info;
            }

            // ② API 兜底
            var apiRec = TryApi(out viaErr);
            if (apiRec != null)
            {
                SaveCache(cachePath, apiRec);
                Fill(info, apiRec);
                info.Ok = true;
                return info;
            }

            info.Error = "检查更新失败：" + viaErr;
            return info;
        }

        /// <summary>读 Release 固定地址的 update.json（直连失败自动走加速站）。</summary>
        private static CacheRecord TryUpdateJson(out string error)
        {
            error = null;
            Exception last = null;
            foreach (var prefix in MirrorPrefixes)
            {
                string url = prefix.Length == 0 ? UpdateJsonUrl : prefix + UpdateJsonUrl;
                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    var req = (HttpWebRequest)WebRequest.Create(url);
                    req.UserAgent = "PPTShapeLibrary-installer-updater";
                    req.Accept = "application/json";
                    req.Timeout = 15000;
                    req.ReadWriteTimeout = 15000;
                    req.AllowAutoRedirect = true;
                    string body;
                    using (var resp = (HttpWebResponse)req.GetResponse())
                    using (var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                    {
                        body = sr.ReadToEnd();
                    }

                    var d = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(body);
                    if (d == null) throw new Exception("update.json 内容为空");
                    var rec = new CacheRecord
                    {
                        CheckedAt = DateTime.Now.ToString("o"),
                        TagName = Str(d, "tag"),
                        LatestVersion = NormalizeVersion(Str(d, "version").Length > 0 ? Str(d, "version") : Str(d, "tag")),
                        ReleaseUrl = Str(d, "releaseUrl"),
                        Notes = Truncate(Str(d, "notes"), 800),
                        AssetName = Str(d, "assetName"),
                        AssetUrl = Str(d, "url"),
                        Sha256Url = Str(d, "sha256Url"),
                    };
                    if (string.IsNullOrEmpty(rec.AssetUrl))
                    {
                        var sha = Str(d, "sha256");
                        if (sha.Length > 0) rec.Notes = (rec.Notes + "\n\n(sha256: " + sha + ")").Trim();
                    }
                    return rec;
                }
                catch (WebException wex)
                {
                    var hr = wex.Response as HttpWebResponse;
                    if (hr != null && hr.StatusCode == HttpStatusCode.NotFound)
                    {
                        // 直连 404 = 仓库还没有 Release（正常）；加速站 404 视为该镜像不可用，继续试下一个
                        if (prefix.Length == 0)
                        {
                            return new CacheRecord
                            {
                                CheckedAt = DateTime.Now.ToString("o"),
                                TagName = "",
                                LatestVersion = CurrentVersion(),
                                Notes = "（仓库还没有发布任何 Release）",
                                ReleaseUrl = ReleasesPageUrl,
                            };
                        }
                    }
                    last = wex;
                }
                catch (Exception ex) { last = ex; }
            }
            error = "update.json 通道不可用：" + (last != null ? last.Message : "未知");
            return null;
        }

        /// <summary>GitHub API 兜底（可能被匿名限流）。</summary>
        private static CacheRecord TryApi(out string error)
        {
            error = null;
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                var req = (HttpWebRequest)WebRequest.Create(LatestApiUrl);
                req.UserAgent = "PPTShapeLibrary-installer-updater";
                req.Accept = "application/vnd.github+json";
                req.Timeout = 15000;
                req.ReadWriteTimeout = 15000;
                string body;
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                {
                    body = sr.ReadToEnd();
                }

                var root = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(body);
                if (root == null) throw new Exception("响应为空");

                var rec = new CacheRecord
                {
                    CheckedAt = DateTime.Now.ToString("o"),
                    TagName = Str(root, "tag_name"),
                    ReleaseUrl = Str(root, "html_url"),
                    Notes = Truncate(Str(root, "body"), 800),
                };
                rec.LatestVersion = NormalizeVersion(rec.TagName);

                object assetsObj;
                if (root.TryGetValue("assets", out assetsObj))
                {
                    var assets = assetsObj as System.Collections.ArrayList;
                    if (assets != null)
                    {
                        string fallbackUrl = null, fallbackName = null;
                        foreach (var a in assets)
                        {
                            var d = a as Dictionary<string, object>;
                            if (d == null) continue;
                            string name = Str(d, "name");
                            string url = Str(d, "browser_download_url");
                            if (string.IsNullOrEmpty(url)) continue;
                            if (name.EndsWith(".sha256", StringComparison.OrdinalIgnoreCase))
                            {
                                if (name.IndexOf("dist", StringComparison.OrdinalIgnoreCase) >= 0) rec.Sha256Url = url;
                                continue;
                            }
                            if (!name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase)) continue;
                            if (name.IndexOf("dist", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                rec.AssetName = name; rec.AssetUrl = url; break;
                            }
                            if (fallbackUrl == null) { fallbackUrl = url; fallbackName = name; }
                        }
                        if (string.IsNullOrEmpty(rec.AssetUrl) && fallbackUrl != null)
                        {
                            rec.AssetUrl = fallbackUrl; rec.AssetName = fallbackName;
                        }
                    }
                }
                return rec;
            }
            catch (WebException wex)
            {
                var hr = wex.Response as HttpWebResponse;
                if (hr != null && hr.StatusCode == HttpStatusCode.NotFound)
                {
                    return new CacheRecord
                    {
                        CheckedAt = DateTime.Now.ToString("o"),
                        TagName = "",
                        LatestVersion = CurrentVersion(),
                        Notes = "（仓库还没有发布任何 Release）",
                        ReleaseUrl = ReleasesPageUrl,
                    };
                }
                if (hr != null && hr.StatusCode == HttpStatusCode.Forbidden)
                {
                    error = "GitHub API 匿名限流（每个出口 IP 每小时 60 次，当前已用尽）。稍后再试即可，" +
                            "或到发布页手动查看：" + ReleasesPageUrl;
                    return null;
                }
                error = "GitHub API 失败：" + wex.Message;
                return null;
            }
            catch (Exception ex)
            {
                error = "GitHub API 失败：" + ex.Message;
                return null;
            }
        }

        private static void Fill(UpdateInfo info, CacheRecord rec)
        {
            info.TagName = rec.TagName;
            info.LatestVersion = string.IsNullOrEmpty(rec.LatestVersion) ? NormalizeVersion(rec.TagName) : rec.LatestVersion;
            info.ReleaseUrl = rec.ReleaseUrl;
            info.Notes = rec.Notes;
            info.AssetName = rec.AssetName;
            info.AssetUrl = rec.AssetUrl;
            info.Sha256Url = rec.Sha256Url;
            info.HasUpdate = !string.IsNullOrEmpty(info.LatestVersion)
                             && CompareVersions(info.LatestVersion, info.CurrentVersion) > 0;
        }

        private static string Str(Dictionary<string, object> d, string key)
        {
            object v;
            return d.TryGetValue(key, out v) && v != null ? v.ToString() : "";
        }

        private static string Truncate(string s, int n)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Replace("\r\n", "\n").Trim();
            return s.Length <= n ? s : s.Substring(0, n) + "…";
        }

        // ---------------- 下载 / 校验 / 应用 ----------------

        /// <summary>下载（直连优先，失败回退加速站）。progress 可为 null。</summary>
        public static void Download(string url, string destPath, Action<long, long> progress)
        {
            Exception last = null;
            foreach (var prefix in MirrorPrefixes)
            {
                string u = prefix.Length == 0 ? url : prefix + url;
                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    var req = (HttpWebRequest)WebRequest.Create(u);
                    req.UserAgent = "PPTShapeLibrary-installer-updater";
                    req.Timeout = 30000;
                    using (var resp = (HttpWebResponse)req.GetResponse())
                    using (var src = resp.GetResponseStream())
                    using (var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                    {
                        long total = resp.ContentLength;
                        long got = 0;
                        var buf = new byte[81920];
                        int n;
                        while ((n = src.Read(buf, 0, buf.Length)) > 0)
                        {
                            fs.Write(buf, 0, n);
                            got += n;
                            if (progress != null) progress(got, total);
                        }
                    }
                    return;   // 成功
                }
                catch (Exception ex)
                {
                    last = ex;
                    try { if (File.Exists(destPath)) File.Delete(destPath); } catch { }
                }
            }
            throw new Exception("下载失败（直连与加速站都试过了）：" + (last != null ? last.Message : "未知"));
        }

        public static string Sha256(string path)
        {
            using (var sha = SHA256.Create())
            using (var fs = File.OpenRead(path))
            {
                var hash = sha.ComputeHash(fs);
                var sb = new StringBuilder();
                foreach (var b in hash) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        /// <summary>下载 .sha256 资产并解析出期望哈希（失败返回 null，不阻断安装）。</summary>
        public static string FetchExpectedHash(string sha256Url)
        {
            if (string.IsNullOrEmpty(sha256Url)) return null;
            try
            {
                var tmp = Path.Combine(Path.GetTempPath(), "pptsl_expected.sha256");
                Download(sha256Url, tmp, null);
                string txt = File.ReadAllText(tmp).Trim();
                try { File.Delete(tmp); } catch { }
                // 兼容 "hash" 与 "hash  filename" 两种写法
                var first = txt.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                return first.Length > 0 ? first[0].ToLowerInvariant() : null;
            }
            catch { return null; }
        }

        /// <summary>解压分发包并启动其中的安装器。返回启动的进程。</summary>
        public static Process ApplyUpdate(string zipPath)
        {
            string dir = Path.Combine(Path.GetTempPath(), "pptsl_update_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(dir);
            System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, dir);

            // 分发包里应当有「安装 PPTShapeLibrary.exe」
            string exe = null;
            foreach (var f in Directory.GetFiles(dir, "*.exe", SearchOption.AllDirectories))
            {
                if (Path.GetFileName(f).IndexOf("安装", StringComparison.Ordinal) >= 0 ||
                    Path.GetFileName(f).IndexOf("install", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    exe = f; break;
                }
            }
            if (exe == null)
            {
                var any = Directory.GetFiles(dir, "*.exe", SearchOption.AllDirectories);
                if (any.Length > 0) exe = any[0];
            }
            if (exe == null) throw new Exception("分发包里没有找到安装器（.exe）。");

            return Process.Start(new ProcessStartInfo(exe) { UseShellExecute = true, WorkingDirectory = dir });
        }
    }
}
