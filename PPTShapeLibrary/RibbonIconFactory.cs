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

        /// <summary>
        /// 彩色图标：圆角渐变色底 + 白色字形（类似 Office 彩色大图标观感）。
        /// 配色按功能语义由调用方传入（浅/深 Ribbon 均清晰）。
        /// </summary>
        public static Image GetColorIcon(string iconName, int size, Color c1, Color c2, float cornerRadius)
        {
            try
            {
                EnsureLoaded();
                if (_family == null || _codepoints == null || !_codepoints.ContainsKey(iconName)) return null;
                int pad = (int)Math.Max(1, size * 0.06f);
                int glyphSize = (int)(size * 0.60f);
                var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.AntiAlias;
                    // 渐变圆角底（上→下）
                    RectangleF bg = new RectangleF(pad, pad, size - 2 * pad, size - 2 * pad);
                    using (var path = RoundedRect(bg, cornerRadius))
                    using (var lgb = new LinearGradientBrush(bg, c1, c2, 90f))
                    {
                        g.FillPath(lgb, path);
                    }
                    // 白字形居中
                    using (var font = new Font(_family, glyphSize, FontStyle.Regular, GraphicsUnit.Pixel))
                    using (var brush = new SolidBrush(Color.White))
                    {
                        string ch = char.ConvertFromUtf32(_codepoints[iconName]);
                        var fmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                        g.DrawString(ch, font, brush, new RectangleF(0, 0, size, size), fmt);
                    }
                }
                return bmp;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "RibbonIconFactory color draw " + iconName);
                return null;
            }
        }

        private static GraphicsPath RoundedRect(RectangleF r, float radius)
        {
            var p = new GraphicsPath();
            float d = radius * 2;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
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

        // 收藏/素材库操作图标名
        public const string CollectShapesIcon = "content_copy";      // 收藏形状
        public const string CollectPageIcon = "select_all";         // 收藏整页
        public const string CollectStyleIcon = "format_paint";      // 收藏样式
        public const string CollectClipboardIcon = "content_paste"; // 收藏剪贴板
        public const string CollectAllSlidesIcon = "collections";   // 收藏整稿
        public const string OpenPaneIcon = "photo_library";         // 打开素材库
        public const string BackupIcon = "save";                    // 导出备份
        public const string RestoreIcon = "unarchive";              // 导入备份
        public const string ExportPptxIcon = "slideshow";           // 导出PPTX
        public const string AdvancedIcon = "construction";          // 高级批处理

        /// <summary>按功能名返回其语义色（渐变起止）。</summary>
        private static void GetActionColors(string action, out Color c1, out Color c2)
        {
            switch (action)
            {
                // 收藏系列 —— 蓝
                case "CollectShapes":
                case "CollectPage":
                case "CollectStyle":
                case "CollectClipboard":
                case "CollectAllSlides":
                    c1 = Color.FromArgb(0x2E, 0x86, 0xDE); c2 = Color.FromArgb(0x1B, 0x5E, 0xAB); break;
                // 素材库/插入 —— 青绿
                case "OpenPane":
                case "ExportPptx":
                    c1 = Color.FromArgb(0x2B, 0xB5, 0x8A); c2 = Color.FromArgb(0x17, 0x86, 0x63); break;
                // 备份/恢复 —— 琥珀/绿
                case "Backup":
                case "Restore":
                    c1 = Color.FromArgb(0xE8, 0xA3, 0x3C); c2 = Color.FromArgb(0xC7, 0x79, 0x1B); break;
                // 工具 1-4 —— 紫/粉
                case "Tool1":
                case "Tool2":
                case "Tool3":
                case "Tool4":
                    c1 = Color.FromArgb(0x9C, 0x6B, 0xDF); c2 = Color.FromArgb(0x7A, 0x47, 0xC1); break;
                // 工具 5-8 —— 粉/青
                case "Tool5":
                case "Tool6":
                case "Tool7":
                case "Tool8":
                    c1 = Color.FromArgb(0xF0, 0x6A, 0x8E); c2 = Color.FromArgb(0xD1, 0x4E, 0x74); break;
                // 高级/设置 —— 靛灰
                default:
                    c1 = Color.FromArgb(0x8A, 0x93, 0xA6); c2 = Color.FromArgb(0x62, 0x6B, 0x7C); break;
            }
        }

        /// <summary>Ribbon 动作按钮彩色图标（32px 徽章）。action 用上面的常量名匹配配色。</summary>
        public static Image GetActionIcon(string iconName, string action, int size = 32)
        {
            Color c1, c2;
            GetActionColors(action, out c1, out c2);
            return GetColorIcon(iconName, size, c1, c2, size * 0.24f);
        }
    }
}
