using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 从 Material Symbols Outlined variable TTF 按图标名（ligature 名）取码点，
    /// 用 GDI+ 绘成 32×32 深色透明 PNG 供 Ribbon 按钮 Image 使用。
    /// 字体文件与码点表位于程序目录 RibbonFonts\ 下。
    /// </summary>
    public static class RibbonIconFactory
    {
        private static PrivateFontCollection _fonts;
        private static FontFamily _family;
        private static Dictionary<string, int> _codepoints;
        private static readonly object Sync = new object();

        private static void EnsureLoaded()
        {
            if (_fonts != null) return;
            lock (Sync)
            {
                if (_fonts != null) return;
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string ttf = Path.Combine(baseDir, "RibbonFonts", "MaterialSymbolsOutlined.ttf");
                string cpFile = Path.Combine(baseDir, "RibbonFonts", "MaterialSymbolsOutlined.codepoints");
                try
                {
                    _fonts = new PrivateFontCollection();
                    if (File.Exists(ttf))
                    {
                        _fonts.AddFontFile(ttf);
                        if (_fonts.Families.Length > 0) _family = _fonts.Families[0];
                    }
                    _codepoints = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    if (File.Exists(cpFile))
                    {
                        foreach (string line in File.ReadAllLines(cpFile))
                        {
                            if (string.IsNullOrWhiteSpace(line)) continue;
                            string[] parts = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                            if (parts.Length >= 2)
                            {
                                int v;
                                if (int.TryParse(parts[1], System.Globalization.NumberStyles.HexNumber, null, out v))
                                    _codepoints[parts[0]] = v;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "RibbonIconFactory load");
                }
            }
        }

        /// <summary>取图标 PNG。color 默认深灰蓝(Ribbon 常规前景)。失败返回 null。</summary>
        public static Image GetIcon(string iconName, int size = 32, Color? color = null)
        {
            try
            {
                EnsureLoaded();
                if (_family == null || _codepoints == null || !_codepoints.ContainsKey(iconName)) return null;
                Color c = color ?? Color.FromArgb(0x1B, 0x1F, 0x23);
                var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.AntiAlias;
                    using (var font = new Font(_family, size * 0.95f, FontStyle.Regular, GraphicsUnit.Pixel))
                    using (var brush = new SolidBrush(c))
                    {
                        string ch = char.ConvertFromUtf32(_codepoints[iconName]);
                        var fmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                        RectangleF rc = new RectangleF(-1, -1, size + 2, size + 2);
                        g.DrawString(ch, font, brush, rc, fmt);
                    }
                }
                return bmp;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "RibbonIconFactory draw " + iconName);
                return null;
            }
        }

        /// <summary>工具组 8 按钮 + 设置图标名。</summary>
        public const string Tool1Icon = "text_fields";
        public const string Tool2Icon = "font_download";
        public const string Tool3Icon = "manage_search";
        public const string Tool4Icon = "cleaning_services";
        public const string Tool5Icon = "image";
        public const string Tool6Icon = "delete_sweep";
        public const string Tool7Icon = "aspect_ratio";
        public const string Tool8Icon = "grid_view";
        public const string SettingsIcon = "settings";
    }
}
