using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;

namespace PPTShapeLibrary
{
    /// <summary>
    /// Ribbon 图标工厂：优先使用 Material Icons 静态字体（Regular，笔画清晰、GDI+ 友好），
    /// 图标名未收录时回退 Material Symbols Outlined variable TTF。
    /// 字体与码点表位于程序目录 RibbonFonts\ 下；Web 端用另一份 woff2（见 WebUI/fonts）。
    /// </summary>
    public static class RibbonIconFactory
    {
        private static PrivateFontCollection _fonts;
        private static FontFamily _familyMain;     // MaterialIcons-Regular（首选，清晰）
        private static FontFamily _familyAlt;      // MaterialSymbolsOutlined（回退）
        private static Dictionary<string, int> _codepoints;      // MaterialIcons 码点
        private static Dictionary<string, int> _codepointsAlt;   // MaterialSymbols 码点
        private static readonly object Sync = new object();

        // ---------------- 加载 ----------------

        private static void EnsureLoaded()
        {
            if (_fonts != null) return;
            lock (Sync)
            {
                if (_fonts != null) return;
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                try
                {
                    _fonts = new PrivateFontCollection();
                    _codepoints = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    _codepointsAlt = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    string mainTtf = Path.Combine(baseDir, "RibbonFonts", "MaterialIcons-Regular.ttf");
                    if (File.Exists(mainTtf))
                    {
                        _fonts.AddFontFile(mainTtf);
                        LoadCodepoints(Path.Combine(baseDir, "RibbonFonts", "MaterialIcons-Regular.codepoints"), _codepoints);
                    }
                    string altTtf = Path.Combine(baseDir, "RibbonFonts", "MaterialSymbolsOutlined.ttf");
                    if (File.Exists(altTtf))
                    {
                        _fonts.AddFontFile(altTtf);
                        LoadCodepoints(Path.Combine(baseDir, "RibbonFonts", "MaterialSymbolsOutlined.codepoints"), _codepointsAlt);
                    }
                    // 按字体族名匹配主/备字体（不依赖 Families 索引顺序，防止主 TTF 缺失时
                    // 把备字体误当主字体使用导致字形与码点表错位）
                    foreach (var fam in _fonts.Families)
                    {
                        string nm = fam.Name;
                        if (nm.IndexOf("Material Icons", StringComparison.OrdinalIgnoreCase) >= 0 && _familyMain == null)
                            _familyMain = fam;
                        else if (nm.IndexOf("Material Symbols", StringComparison.OrdinalIgnoreCase) >= 0 && _familyAlt == null)
                            _familyAlt = fam;
                    }
                    if (_familyMain == null && _familyAlt != null) _familyMain = _familyAlt;
                    if (_familyMain == null && _fonts.Families.Length > 0)
                        _familyMain = _fonts.Families[0]; // 兜底：有任何字体都比无字强
                    if (_familyMain == null)
                        Log.Error(new Exception("Ribbon fonts missing: neither MaterialIcons nor MaterialSymbols ttf found in " + baseDir + "\\RibbonFonts"), "RibbonIconFactory");
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "RibbonIconFactory load");
                }
            }
        }

        private static void LoadCodepoints(string cpFile, Dictionary<string, int> target)
        {
            if (!File.Exists(cpFile)) return;
            foreach (string line in File.ReadAllLines(cpFile))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                string[] parts = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2)
                {
                    int v;
                    if (int.TryParse(parts[1], System.Globalization.NumberStyles.HexNumber, null, out v))
                    {
                        if (!target.ContainsKey(parts[0])) target[parts[0]] = v;
                    }
                }
            }
        }

        /// <summary>图标名 → (字体族, 码点)。优先主字体，缺则回退。返回 false 表示均缺失。</summary>
        private static bool ResolveGlyph(string iconName, out FontFamily family, out int codepoint)
        {
            family = null; codepoint = 0;
            if (_codepoints != null && _codepoints.ContainsKey(iconName) && _familyMain != null)
            {
                family = _familyMain; codepoint = _codepoints[iconName]; return true;
            }
            if (_codepointsAlt != null && _codepointsAlt.ContainsKey(iconName) && _familyAlt != null)
            {
                family = _familyAlt; codepoint = _codepointsAlt[iconName]; return true;
            }
            return false;
        }

        // ---------------- 绘制 ----------------

        /// <summary>
        /// 高辨识彩色字形：纯彩色字形（无底色块），内部 4x 超采样后缩回，
        /// 小尺寸下依然清晰锐利；颜色按功能语义分配。
        /// </summary>
        public static Image GetColorGlyph(string iconName, int size, Color color)
        {
            try
            {
                EnsureLoaded();
                FontFamily fam;
                int cp;
                if (!ResolveGlyph(iconName, out fam, out cp)) return null;
                int ss = size * 4;                       // 超采样
                var bmp = new Bitmap(ss, ss, PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.AntiAlias;
                    using (var font = new Font(fam, ss * 0.92f, FontStyle.Regular, GraphicsUnit.Pixel))
                    using (var brush = new SolidBrush(color))
                    {
                        string ch = char.ConvertFromUtf32(cp);
                        var fmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                        RectangleF rc = new RectangleF(0, -ss * 0.02f, ss, ss);
                        g.DrawString(ch, font, brush, rc, fmt);
                    }
                }
                // 缩回目标尺寸
                var small = new Bitmap(size, size, PixelFormat.Format32bppArgb);
                using (var g2 = Graphics.FromImage(small))
                {
                    g2.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g2.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g2.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    g2.DrawImage(bmp, 0, 0, size, size);
                }
                bmp.Dispose();
                return small;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "RibbonIconFactory glyph " + iconName);
                return null;
            }
        }

        /// <summary>每个动作的独立高辨识色（按功能语义，颜色鲜艳且彼此可区分）。</summary>
        private static Color ActionColor(string action)
        {
            switch (action)
            {
                case "CollectShapes":    return Color.FromArgb(0x1F, 0x77, 0xD4); // 蓝
                case "CollectPage":      return Color.FromArgb(0x00, 0x8A, 0x8A); // 青
                case "CollectStyle":     return Color.FromArgb(0x7A, 0x3F, 0xB8); // 紫
                case "CollectClipboard": return Color.FromArgb(0x00, 0x70, 0xC0); // 天蓝
                case "CollectAllSlides": return Color.FromArgb(0x2E, 0x9B, 0x4F); // 绿
                case "OpenPane":         return Color.FromArgb(0xD4, 0x8A, 0x0E); // 琥珀
                case "Backup":           return Color.FromArgb(0xB0, 0x6A, 0x00); // 棕橙
                case "Restore":          return Color.FromArgb(0x2E, 0x9B, 0x4F); // 绿
                case "ExportPptx":       return Color.FromArgb(0xD6, 0x3A, 0x2A); // 红橙
                case "Tool1":            return Color.FromArgb(0x1F, 0x77, 0xD4); // 蓝
                case "Tool2":            return Color.FromArgb(0x00, 0x8A, 0x8A); // 青
                case "Tool3":            return Color.FromArgb(0xD4, 0x7F, 0x00); // 橙
                case "Tool4":            return Color.FromArgb(0x7A, 0x3F, 0xB8); // 紫
                case "Tool5":            return Color.FromArgb(0xD6, 0x3A, 0x2A); // 红
                case "Tool6":            return Color.FromArgb(0x2E, 0x9B, 0x4F); // 绿
                case "Tool7":            return Color.FromArgb(0x0B, 0x7A, 0x9B); // 深青
                case "Tool8":            return Color.FromArgb(0x9B, 0x59, 0xC8); // 紫红
                case "Tool9":            return Color.FromArgb(0x3A, 0x6E, 0xA5); // 靛蓝：页脚
                case "Tool10":           return Color.FromArgb(0xC6, 0x3D, 0x2F); // 红：PDF
                case "Tool11":           return Color.FromArgb(0x4C, 0x8B, 0x5C); // 绿：图表
                case "Tool12":           return Color.FromArgb(0xE0, 0x8A, 0x2E); // 橙金：成套模板
                case "Settings":         return Color.FromArgb(0x60, 0x6B, 0x7A); // 灰
                case "Advanced":         return Color.FromArgb(0x55, 0x6B, 0x8A); // 靛
                default:                 return Color.FromArgb(0x1B, 0x5E, 0xAB);
            }
        }

        /// <summary>Ribbon 动作按钮高辨识彩色图标。action 对应 ActionColor 色相。</summary>
        public static Image GetActionGlyph(string iconName, string action, int size = 32)
        {
            return GetColorGlyph(iconName, size, ActionColor(action));
        }

        // ---------------- 图标名常量 ----------------

        public const string Tool1Icon = "text_fields";
        public const string Tool2Icon = "font_download";
        public const string Tool3Icon = "manage_search";
        public const string Tool4Icon = "cleaning_services";
        public const string Tool5Icon = "image";
        public const string Tool6Icon = "delete_sweep";
        public const string Tool7Icon = "aspect_ratio";
        public const string Tool8Icon = "grid_view";
        public const string Tool9Icon = "text_snippet";      // 页脚页码
        public const string Tool10Icon = "picture_as_pdf";   // 导出PDF
        public const string Tool11Icon = "bar_chart";        // 图表美化
        public const string Tool12Icon = "view_day";         // 成套模板
        public const string SettingsIcon = "settings";

        public const string CollectShapesIcon = "content_copy";
        public const string CollectPageIcon = "select_all";
        public const string CollectStyleIcon = "format_paint";
        public const string CollectClipboardIcon = "content_paste";
        public const string CollectAllSlidesIcon = "collections";
        public const string OpenPaneIcon = "photo_library";
        public const string BackupIcon = "save";
        public const string RestoreIcon = "unarchive";
        public const string ExportPptxIcon = "slideshow";
        public const string AdvancedIcon = "construction";
    }
}
