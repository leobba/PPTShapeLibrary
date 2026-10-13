using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.Win32;

namespace PPTShapeLibrary
{
    /// <summary>一键更新的准备结果。</summary>
    public class UpdateApply
    {
        public bool Ok;
        public string Message;
        public string Dir;        // 新版解压目录
        public string Installer;  // 解压出来的安装器路径
    }

    /// <summary>下载结果类型。</summary>
    public enum DownloadStatus
    {
        Ok,
        NotFound,   // HTTP 404：资产没托管
        Failed      // 网络/连接问题
    }

    /// <summary>一次更新检查的结果（发给前端）。</summary>
    public class UpdateResult
    {
        public bool Ok;
        public bool HasUpdate;
        public string CurrentVersion;
        public string LatestVersion;
        public string ReleaseUrl;    // 发布页
        public string DownloadUrl;   // 分发包直链（前端可直接给用户点）
        public string Sha256;
        public string Notes;
        public string Error;
    }

    /// <summary>
    /// 插件内的更新检查（用户在 PowerPoint 里就能看到"有新版本"和更新内容）。
    ///
    /// 与安装器（installer/UpdateChecker.cs）保持**同一套协议与通道**，但独立实现：
    ///  · 清单：① 固定地址 `releases/latest/download/update.json`（不走 API、不受匿名限流）
    ///          → ② GitHub API（releases/latest）兜底
    ///  · 下载：直连 GitHub → 失败回退加速站（ghproxy.net / gh-proxy.com / ghfast.top）
    ///  · 版本：读插件程序集版本（ClickOnce 安装后就是真实版本号）
    ///  · 缓存：24 小时（%APPDATA%\PPTShapeLibrary\update-check.json），避免反复联网
    /// </summary>
    public static class UpdateService
    {
        public const string Owner = "leobba";
        public const string Repo = "PPTShapeLibrary";
        public static readonly string ReleasesPageUrl =
            "https://github.com/" + Owner + "/" + Repo + "/releases";
        public static readonly string UpdateJsonUrl =
            "https://github.com/" + Owner + "/" + Repo + "/releases/latest/download/update.json";
        public static readonly string LatestApiUrl =
            "https://api.github.com/repos/" + Owner + "/" + Repo + "/releases/latest";

        /// <summary>GitHub 加速站前缀，按顺序尝试；空串 = 直连优先。
        /// 只对 github.com / objects.githubusercontent.com 有效（加速站不代理微软域名）。</summary>
        public static readonly string[] MirrorPrefixes =
        {
            "",
            "https://ghproxy.net/",
            "https://gh-proxy.com/",
            "https://ghfast.top/",
        };

        private static readonly TimeSpan CacheTtl = TimeSpan.FromHours(24);

        /// <summary>
        /// 当前版本。**不能**读程序集版本：AssemblyInfo.cs 的版本停在 1.0.1.0 不再随发版更新，
        /// 而 ClickOnce 安装后的真实版本体现在分发包装目录名 `App\Application Files\PPTShapeLibrary_x_y_z_w`。
        /// 取不到时才退回程序集版本。
        /// </summary>
        public static string CurrentVersion()
        {
            // ① 已安装版本：控制面板卸载项里的 DisplayVersion（ClickOnce 安装后就是真实版本）
            var installed = InstalledVersion();
            if (!string.IsNullOrEmpty(installed)) return NormalizeVersion(installed);

            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                // ② 分发包布局：<exe 目录>\App\Application Files\PPTShapeLibrary_x_y_z_w
                string appFiles = Path.Combine(baseDir, "App", "Application Files");
                if (Directory.Exists(appFiles))
                {
                    foreach (var d in Directory.GetDirectories(appFiles, "PPTShapeLibrary_*"))
                    {
                        var v = VersionFromDirName(Path.GetFileName(d));
                        if (v != null) return v;
                    }
                }
            }
            catch { }

            // ③ 兜底：程序集版本（AssemblyInfo 可能不随发版更新，仅作最后手段）
            try
            {
                var v = Assembly.GetExecutingAssembly().GetName().Version;
                return v == null ? "0.0.0.0" : v.ToString();
            }
            catch { return "0.0.0.0"; }
        }

        /// <summary>从控制面板卸载项读已安装版本（找不到返回 null）。</summary>
        private static string InstalledVersion()
        {
            // ① 开发模式（|vstolocal）：dev-register.ps1 会把真实版本写进加载项注册表的 ProductVersion。
            //    这一路优先，因为开发模式下没有 ARP 项、程序集版本也不随发版变。
            var regVer = RegistryProductVersion();
            var arpVer = ArpVersion();
            // 两个来源都可能不准：注册表 ProductVersion 是 dev-register 写的（开发模式），
            // ARP DisplayVersion 来自早期 vsto 清单（可能陈旧，实测见过 1.0.1.0）。
            // 取**较大的**那个，避免被陈旧值拖成"永远有新版本"或显示错误的当前版本。
            var best = regVer;
            if (!string.IsNullOrEmpty(arpVer) &&
                (string.IsNullOrEmpty(best) || CompareVersions(arpVer, best) > 0))
            {
                best = arpVer;
            }
            return string.IsNullOrEmpty(best) ? null : best;
        }

        /// <summary>开发模式：dev-register.ps1 写入加载项注册表的 ProductVersion。</summary>
        private static string RegistryProductVersion()
        {
            try
            {
                foreach (var keyPath in new[]
                {
                    @"Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary",
                    @"Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary"
                })
                {
                    using (var k = Registry.CurrentUser.OpenSubKey(keyPath))
                    {
                        if (k == null) continue;
                        var v = k.GetValue("ProductVersion") as string;
                        if (!string.IsNullOrEmpty(v)) return v;
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>正式安装：控制面板卸载项里的 DisplayVersion（HKLM 两视图 + HKCU）。</summary>
        private static string ArpVersion()
        {
            try
            {
                foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                {
                    using (var root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view)
                            .OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
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
                using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default)
                        .OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
                {
                    if (root != null)
                    {
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

        private static string VersionFromDirName(string dirName)
        {
            const string prefix = "PPTShapeLibrary_";
            if (string.IsNullOrEmpty(dirName) || !dirName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return null;
            if (dirName.Length <= prefix.Length) return null;
            var ver = dirName.Substring(prefix.Length).Replace('_', '.');
            return NormalizeVersion(ver);
        }

        private static string CachePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "PPTShapeLibrary", "update-check.json");
            }
        }

        private static string NormalizeVersion(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "0.0.0.0";
            string s = raw.Trim();
            if (s.StartsWith("v", StringComparison.OrdinalIgnoreCase)) s = s.Substring(1);
            int dash = s.IndexOf('-');
            if (dash > 0) s = s.Substring(0, dash);
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

        private static string Truncate(string s, int n)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Replace("\r\n", "\n").Trim();
            return s.Length <= n ? s : s.Substring(0, n) + "…";
        }

        private static string Str(Dictionary<string, object> d, string key)
        {
            object v;
            return d.TryGetValue(key, out v) && v != null ? v.ToString() : "";
        }

        /// <summary>
        /// 手工解析 update.json 的字段。
        ///
        /// 为什么不用 JavaScriptSerializer / Newtonsoft：实测 `JavaScriptSerializer` 在普通
        /// WinForms/VSTO 进程里反序列化会抛
        ///   Could not load type 'System.Web.UI.WebResourceAttribute' from assembly 'System.Web…'
        /// —— 这是 .NET Framework 的已知坑（它假定跑在 Web 宿主里）。而 update.json
        /// 的字段结构是我们自己定的、很稳定，正则提取足够且零依赖。
        /// </summary>
        private static CacheRecord ParseManifest(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;
            var rec = new CacheRecord
            {
                LatestVersion = NormalizeVersion(
                    Pick(json, "version").Length > 0 ? Pick(json, "version") : Pick(json, "tag")),
                ReleaseUrl = Pick(json, "releaseUrl"),
                DownloadUrl = Pick(json, "url"),
                Sha256 = Pick(json, "sha256"),
                Notes = Truncate(Pick(json, "notes"), 800),
            };
            if (rec.LatestVersion == "0.0.0.0" && rec.DownloadUrl.Length == 0) return null;
            return rec;
        }

        /// <summary>从 JSON 文本里取一个字符串字段的值（不回复杂制）。</summary>
        private static string Pick(string json, string key)
        {
            if (string.IsNullOrEmpty(json)) return "";
            var m = System.Text.RegularExpressions.Regex.Match(
                json, "\"" + System.Text.RegularExpressions.Regex.Escape(key) + "\"\\s*:\\s*\"(.*?)\"",
                System.Text.RegularExpressions.RegexOptions.Singleline);
            return m.Success ? UnescapeJson(m.Groups[1].Value) : "";
        }

        private static string UnescapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\r", "\r").Replace("\\n", "\n").Replace("\\t", "\t")
                    .Replace("\\\"", "\"").Replace("\\\\", "\\");
        }

        private class CacheRecord
        {
            public string CheckedAt { get; set; }
            public string LatestVersion { get; set; }
            public string ReleaseUrl { get; set; }
            public string DownloadUrl { get; set; }
            public string Sha256 { get; set; }
            public string Notes { get; set; }
        }

        private static CacheRecord LoadCache()
        {
            try
            {
                if (!File.Exists(CachePath)) return null;
                var rec = new JavaScriptSerializer().Deserialize<CacheRecord>(
                    File.ReadAllText(CachePath, Encoding.UTF8));
                if (rec == null || string.IsNullOrEmpty(rec.CheckedAt)) return null;
                return rec;
            }
            catch { return null; }
        }

        private static void SaveCache(CacheRecord rec)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(CachePath));
                File.WriteAllText(CachePath, new JavaScriptSerializer().Serialize(rec), Encoding.UTF8);
            }
            catch { }
        }

        private static void Fill(UpdateResult info, CacheRecord rec)
        {
            info.LatestVersion = rec.LatestVersion;
            info.ReleaseUrl = rec.ReleaseUrl;
            info.DownloadUrl = rec.DownloadUrl;
            info.Sha256 = rec.Sha256;
            info.Notes = rec.Notes;
            info.HasUpdate = !string.IsNullOrEmpty(info.LatestVersion)
                             && CompareVersions(info.LatestVersion, info.CurrentVersion) > 0;
        }

        /// <summary>检查更新（24h 内直接返回缓存，不联网）。</summary>
        public static UpdateResult Check()
        {
            var info = new UpdateResult { CurrentVersion = CurrentVersion() };

            var cached = LoadCache();
            if (cached != null)
            {
                DateTime at;
                if (DateTime.TryParse(cached.CheckedAt, out at) && DateTime.Now - at < CacheTtl)
                {
                    Fill(info, cached);
                    info.Ok = true;
                    return info;
                }
            }

            string err;
            var rec = TryUpdateJson(out err);
            if (rec == null) rec = TryApi(out err);
            if (rec == null)
            {
                info.Error = "检查更新失败：" + err;
                return info;
            }

            rec.CheckedAt = DateTime.Now.ToString("o");
            SaveCache(rec);
            Fill(info, rec);
            info.Ok = true;
            return info;
        }

        /// <summary>① 固定地址 update.json（首选，不走 API）。</summary>
        private static CacheRecord TryUpdateJson(out string error)
        {
            error = null;
            Exception last = null;
            foreach (var prefix in MirrorPrefixes)
            {
                string url = prefix.Length == 0 ? UpdateJsonUrl : prefix + UpdateJsonUrl;
                try
                {
                    string body = GetString(url);
                    var rec = ParseManifest(body);
                    if (rec == null) throw new Exception("update.json 内容为空或格式不对");
                    if (string.IsNullOrEmpty(rec.ReleaseUrl)) rec.ReleaseUrl = ReleasesPageUrl;
                    return rec;
                }
                catch (WebException wex)
                {
                    var hr = wex.Response as HttpWebResponse;
                    if (hr != null && hr.StatusCode == HttpStatusCode.NotFound && prefix.Length == 0)
                    {
                        // 直连 404 = 还没有 Release（正常）
                        return new CacheRecord
                        {
                            CheckedAt = DateTime.Now.ToString("o"),
                            LatestVersion = CurrentVersion(),
                            ReleaseUrl = ReleasesPageUrl,
                            Notes = "（仓库还没有发布任何 Release）",
                        };
                    }
                    last = wex;
                }
                catch (Exception ex) { last = ex; }
            }
            error = "update.json 通道不可用：" + (last != null ? last.Message : "未知");
            return null;
        }

        /// <summary>② GitHub API 兜底（可能被匿名限流 403）。</summary>
        private static CacheRecord TryApi(out string error)
        {
            error = null;
            try
            {
                string body = GetString(LatestApiUrl);
                var root = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(body);
                if (root == null) throw new Exception("响应为空");

                var rec = new CacheRecord
                {
                    LatestVersion = NormalizeVersion(Str(root, "tag_name")),
                    ReleaseUrl = Str(root, "html_url"),
                    Notes = Truncate(Str(root, "body"), 800),
                };
                object assetsObj;
                if (root.TryGetValue("assets", out assetsObj))
                {
                    var assets = assetsObj as System.Collections.ArrayList;
                    if (assets != null)
                    {
                        string fallback = null;
                        foreach (var a in assets)
                        {
                            var d = a as Dictionary<string, object>;
                            if (d == null) continue;
                            string name = Str(d, "name");
                            string url = Str(d, "browser_download_url");
                            if (string.IsNullOrEmpty(url)) continue;
                            if (name.EndsWith(".sha256", StringComparison.OrdinalIgnoreCase)) continue;
                            if (!name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase)) continue;
                            if (name.IndexOf("dist", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                rec.DownloadUrl = url; break;
                            }
                            if (fallback == null) fallback = url;
                        }
                        if (string.IsNullOrEmpty(rec.DownloadUrl)) rec.DownloadUrl = fallback;
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
                        LatestVersion = CurrentVersion(),
                        ReleaseUrl = ReleasesPageUrl,
                        Notes = "（仓库还没有发布任何 Release）",
                    };
                }
                if (hr != null && hr.StatusCode == HttpStatusCode.Forbidden)
                {
                    error = "GitHub API 匿名限流（每 IP 每小时 60 次已用尽），稍后再试；" +
                            "也可到发布页手动查看：" + ReleasesPageUrl;
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

        /// <summary>
        /// 一键更新：下载新版本 → 校验 SHA256 → 解压 → 启动其中的安装器（带 --pending-update）
        /// → 由安装器等 PowerPoint 退出后完成"卸载旧版 + 装新版"。
        ///
        /// 返回：Ok / Message / Dir（解压目录，供安装器使用）。
        /// 注意：本方法**不**关闭 PowerPoint，关闭由调用方（用户在弹窗里确认后）执行。
        /// </summary>
        public static UpdateApply ApplyUpdate(Action<long, long> progress)
        {
            var res = new UpdateApply();
            var info = Check();
            if (!info.Ok || !info.HasUpdate)
            {
                res.Message = info.Ok ? "当前已是最新版本，无需更新。" : ("检查更新失败：" + info.Error);
                return res;
            }
            if (string.IsNullOrEmpty(info.DownloadUrl))
            {
                res.Message = "该版本没有提供下载地址，请到发布页手动下载。";
                return res;
            }

            string zip = null;
            try
            {
                zip = Path.Combine(Path.GetTempPath(), info.LatestVersion + "-dist.zip");
                Log.Write("[update] 下载 " + info.DownloadUrl);
                if (Download(info.DownloadUrl, zip, progress) != DownloadStatus.Ok)
                    throw new Exception("下载失败（直连与加速站都试过了）");

                if (!string.IsNullOrEmpty(info.Sha256))
                {
                    var actual = Sha256(zip);
                    if (!string.Equals(actual, info.Sha256, StringComparison.OrdinalIgnoreCase))
                        throw new Exception("校验失败，已放弃安装（下载的文件与发布的不一致）");
                    Log.Write("[update] SHA256 校验通过");
                }

                // 解压到**持久目录**，不要用 %TEMP%：
                // ClickOnce 会把"安装源 URL"记下来（后续升级/修复/卸载都要回到这个路径），
                // 放 temp 里迟早被清理，安装源失效会导致以后无法升级或卸载。
                var root = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PPTShapeLibrary", "updates");
                var dir = Path.Combine(root, info.LatestVersion);
                try
                {
                    if (Directory.Exists(dir)) Directory.Delete(dir, true);
                }
                catch { }
                Directory.CreateDirectory(dir);
                System.IO.Compression.ZipFile.ExtractToDirectory(zip, dir);
                Log.Write("[update] 已解压到 " + dir);

                var installer = Directory.GetFiles(dir, "安装 PPTShapeLibrary.exe", SearchOption.AllDirectories);
                if (installer.Length == 0) throw new Exception("分发包里没有找到安装器");

                try { File.Delete(zip); } catch { }

                res.Ok = true;
                res.Dir = dir;
                res.Installer = installer[0];
                res.Message = "新版已准备好：" + info.LatestVersion;
                return res;
            }
            catch (Exception ex)
            {
                try { if (zip != null && File.Exists(zip)) File.Delete(zip); } catch { }
                res.Message = ex.Message;
                Log.Write("[update] 更新失败：" + ex.Message);
                return res;
            }
        }

        /// <summary>下载（直连优先，失败回退加速站）。</summary>
        public static DownloadStatus Download(string url, string destPath, Action<long, long> progress)
        {
            Exception last = null;
            foreach (var prefix in MirrorPrefixes)
            {
                string u = prefix.Length == 0 ? url : prefix + url;
                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    var req = (HttpWebRequest)WebRequest.Create(u);
                    req.UserAgent = "PPTShapeLibrary-plugin-updater";
                    req.Accept = "application/octet-stream";
                    req.Timeout = 30000;
                    req.ReadWriteTimeout = 30000;
                    req.Proxy = null;
                    using (var resp = (HttpWebResponse)req.GetResponse())
                    using (var src = resp.GetResponseStream())
                    using (var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                    {
                        long total = resp.ContentLength, got = 0;
                        var buf = new byte[81920];
                        int n;
                        while ((n = src.Read(buf, 0, buf.Length)) > 0)
                        {
                            fs.Write(buf, 0, n);
                            got += n;
                            if (progress != null) progress(got, total);
                        }
                    }
                    return DownloadStatus.Ok;
                }
                catch (WebException wex)
                {
                    var hr = wex.Response as HttpWebResponse;
                    if (hr != null && hr.StatusCode == HttpStatusCode.NotFound) return DownloadStatus.NotFound;
                    last = wex;
                    try { if (File.Exists(destPath)) File.Delete(destPath); } catch { }
                }
                catch (Exception ex)
                {
                    last = ex;
                    try { if (File.Exists(destPath)) File.Delete(destPath); } catch { }
                }
            }
            Log.Write("[update] 下载失败：" + (last != null ? last.Message : "未知"));
            return DownloadStatus.Failed;
        }

        public static string Sha256(string path)
        {
            using (var sha = System.Security.Cryptography.SHA256.Create())
            using (var fs = File.OpenRead(path))
            {
                var sb = new StringBuilder();
                foreach (var b in sha.ComputeHash(fs)) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        /// <summary>判断一个 HTTP(S) 地址是否可用（只要响应码是 2xx 就算可用）。</summary>
        private static string GetString(string url)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.UserAgent = "PPTShapeLibrary-plugin-updater";
            // 注意：这里**不能**用 application/vnd.github+json —— 那是给 API 用的媒体类型，
            // 对 update.json 这种普通文件下载会让 GitHub 返回 406/415。
            // 安装器那条链路实测能通用的是 application/json（见 installer/UpdateChecker.cs）。
            req.Accept = "application/json";
            req.Timeout = 15000;
            req.ReadWriteTimeout = 15000;
            req.Proxy = null;
            using (var resp = (HttpWebResponse)req.GetResponse())
            using (var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
            {
                return sr.ReadToEnd();
            }
        }
    }
}
