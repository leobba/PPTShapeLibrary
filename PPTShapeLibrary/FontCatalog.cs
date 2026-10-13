using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using Microsoft.Win32;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 系统可用字体目录（供前端字体下拉使用），目标是与 PowerPoint 原生字体下拉尽量一致。
    ///
    /// 来源：
    ///   1) GDI+ InstalledFontCollection —— 已安装字体族（含本地化中文名），实测本机 197 族且与
    ///      逐个解析字体文件的结果完全一致，系统字体基本已全覆盖；
    ///   2) 用户级安装的字体（"为当前用户安装"）—— HKCU 字体注册表 + %LOCALAPPDATA%\Microsoft\Windows\Fonts。
    ///      这类字体不写 HKLM、也可能不出现在 GDI+ 列表里，必须按字体文件解析真实族名，
    ///      否则用户在 PowerPoint 里能看到、插件下拉里却没有。
    ///
    /// 刻意不遍历 HKLM 的全部字体文件：那边约 200 个文件逐个 AddFontFile 需 ~2 秒，而 GDI+ 已覆盖。
    /// 结果缓存 30 秒，新装字体无需重启 PowerPoint 即可出现。
    /// </summary>
    internal static class FontCatalog
    {
        private static readonly object Gate = new object();
        private static readonly TimeSpan Ttl = TimeSpan.FromSeconds(30);
        private const int MaxUserFontFiles = 100;   // 兜底上限，避免异常环境拖慢首帧

        private static List<string> _cache;
        private static DateTime _cachedAtUtc = DateTime.MinValue;

        /// <summary>返回去重排序后的字体族名列表（含文稿可用字体）。</summary>
        public static List<string> GetFamilies()
        {
            lock (Gate)
            {
                if (_cache != null && (DateTime.UtcNow - _cachedAtUtc) < Ttl) return _cache;

                var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                CollectInstalledFamilies(set);
                CollectUserFontFamilies(set);

                var list = set.Where(IsUsableName).ToList();
                list.Sort(StringComparer.OrdinalIgnoreCase);

                _cache = list;
                _cachedAtUtc = DateTime.UtcNow;
                return list;
            }
        }

        /// <summary>清空缓存（字体安装后可立即刷新）。</summary>
        public static void Invalidate()
        {
            lock (Gate) { _cache = null; _cachedAtUtc = DateTime.MinValue; }
        }

        private static void CollectInstalledFamilies(HashSet<string> set)
        {
            try
            {
                using (var fc = new InstalledFontCollection())
                {
                    foreach (var fam in fc.Families)
                        if (fam != null && !string.IsNullOrEmpty(fam.Name)) set.Add(fam.Name);
                }
            }
            catch { }
        }

        /// <summary>收集用户级字体文件并解析其真实族名。</summary>
        private static void CollectUserFontFamilies(HashSet<string> set)
        {
            var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // ① HKCU 字体注册表（值 = 字体文件名或绝对路径）
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"))
                {
                    if (key != null)
                    {
                        foreach (var name in key.GetValueNames())
                        {
                            var v = key.GetValue(name) as string;
                            if (string.IsNullOrWhiteSpace(v)) continue;
                            string path = v.IndexOf(':') > 0
                                ? v
                                : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", v);
                            if (File.Exists(path)) files.Add(path);
                        }
                    }
                }
            }
            catch { }

            // ② 用户字体目录（部分安装方式只落文件不写注册表）
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    @"Microsoft\Windows\Fonts");
                if (Directory.Exists(dir))
                    foreach (var f in Directory.GetFiles(dir))
                        if (IsFontFile(f)) files.Add(f);
            }
            catch { }

            int n = 0;
            foreach (var f in files)
            {
                if (n++ >= MaxUserFontFiles) break;
                try
                {
                    using (var pfc = new PrivateFontCollection())
                    {
                        pfc.AddFontFile(f);
                        foreach (var fam in pfc.Families)
                            if (fam != null && !string.IsNullOrEmpty(fam.Name)) set.Add(fam.Name);
                    }
                }
                catch { }   // 个别字体文件损坏/受保护：跳过
            }
        }

        private static bool IsFontFile(string path)
        {
            string ext = Path.GetExtension(path);
            return ".ttf".Equals(ext, StringComparison.OrdinalIgnoreCase)
                || ".otf".Equals(ext, StringComparison.OrdinalIgnoreCase)
                || ".ttc".Equals(ext, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>过滤不可用的名字：主题字体引用(+mn-*)、竖排变体(@*)、本插件图标字体。</summary>
        private static bool IsUsableName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            if (name[0] == '+' || name[0] == '@') return false;
            if (name.StartsWith("Material", StringComparison.OrdinalIgnoreCase)) return false;
            if (name.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            return true;
        }
    }
}
