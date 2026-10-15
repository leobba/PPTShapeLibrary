# gen-icon-preview.ps1 —— 生成 exe 图标的**预览对比板**（先看效果，满意了再接到构建里）
#
# 用 Add-Type 编译一段 C# 来绘制（PowerShell 直接操作 PrivateFontCollection 会报 null，
# 走 .NET 编译最稳；且与 RibbonIconFactory 同一套 System.Drawing 技术栈）。
#
# 候选（都是"库/素材"语义）：
#   photo_library (e413) 现「打开素材库」按钮同款，语义最一致
#   collections   (e3b6) 收藏集
#   grid_view     (e9b0) 素材墙
#   format_shapes (e25e) 形状（核心）
#   category      (e574) 分类
#   palette       (e40a) 样式/调色
#
# 用法：
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test\gen-icon-preview.ps1
param([string]$Root = "E:\code\PPTShapeLibrary")

$ErrorActionPreference = "Stop"
$ttf = Join-Path $Root "PPTShapeLibrary\bin\Release\RibbonFonts\MaterialIcons-Regular.ttf"
if (-not (Test-Path $ttf)) { throw "icon font not found: $ttf" }

$outDir = Join-Path $Root "build\icon-preview"
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$src = @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Globalization;
using System.IO;

public static class IconPreview
{
    static FontFamily _family;
    static Color _brand = Color.FromArgb(0x1B, 0x5E, 0xAB);

    public static string Run(string ttf, string outPng)
    {
        var coll = new PrivateFontCollection();
        coll.AddFontFile(ttf);
        _family = coll.Families[0];

        var cands = new[] {
            new { Name = "photo_library", Code = "e413", Desc = "shape library (same as button)" },
            new { Name = "collections",   Code = "e3b6", Desc = "collections" },
            new { Name = "grid_view",     Code = "e9b0", Desc = "material wall" },
            new { Name = "format_shapes", Code = "e25e", Desc = "shapes (core)" },
            new { Name = "category",      Code = "e574", Desc = "categories" },
            new { Name = "palette",       Code = "e40a", Desc = "style / palette" },
        };
        var sizes = new[] { 256, 64, 48, 32, 24, 16 };
        int pad = 14, labelH = 34, rowH = 256 + pad;
        int colW = 0; foreach (var s in sizes) colW += s + pad;
        int w = 60 + colW + 200;
        int h = 50 + cands.Length * (rowH + labelH + 10) * 2 + 120;

        using (var bmp = new Bitmap(w, h))
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.FromArgb(0xF5, 0xF7, 0xFA));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

            var tf = new Font("Microsoft YaHei UI", 18, FontStyle.Bold);
            var nf = new Font("Microsoft YaHei UI", 12, FontStyle.Bold);
            var df = new Font("Microsoft YaHei UI", 10);
            var sf = new Font("Microsoft YaHei UI", 9);
            var dark = new SolidBrush(Color.FromArgb(0x1B, 0x3A, 0x57));
            var gray = new SolidBrush(Color.Gray);
            var blue = new SolidBrush(_brand);

            int y = 20;
            g.DrawString("PPTShapeLibrary - exe icon preview (pick one, then wire into build)", tf, dark, 20, y);
            y += 46;

            for (int st = 0; st < 2; st++)
            {
                bool plate = (st == 1);
                string title = plate
                    ? "STYLE B: rounded blue plate + white glyph (app-like)"
                    : "STYLE A: plain colored glyph (matches existing Ribbon buttons)";
                g.DrawString(title, nf, blue, 20, y);
                y += 24;

                foreach (var c in cands)
                {
                    int x = 70;
                    foreach (var s in sizes)
                    {
                        using (var img = MakeGlyph(c.Code, s, plate ? Color.White : _brand, plate))
                        {
                            int iy = y + (int)Math.Round((256 - s) / 2.0);
                            g.DrawImage(img, x, iy, s, s);
                        }
                        g.DrawString(s.ToString(), sf, gray, x, y + 258);
                        x += s + pad;
                    }
                    g.DrawString(c.Name, nf, dark, x + 6, y + 100);
                    g.DrawString(c.Desc, df, gray, x + 6, y + 124);
                    y += rowH + labelH + 10;
                }
                y += 20;
            }

            tf.Dispose(); nf.Dispose(); df.Dispose(); sf.Dispose();
            dark.Dispose(); gray.Dispose(); blue.Dispose();

            bmp.Save(outPng, ImageFormat.Png);
        }
        coll.Dispose();
        return w + "x" + h;
    }

    static Bitmap MakeGlyph(string cpHex, int size, Color color, bool onPlate)
    {
        int ss = size * 4;                       // supersample like RibbonIconFactory
        var bmp = new Bitmap(ss, ss);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            g.Clear(Color.Transparent);

            if (onPlate)
            {
                using (var br = new SolidBrush(_brand))
                {
                    float r = ss * 0.18f;
                    var path = new GraphicsPath();
                    float d = r * 2;
                    path.AddArc(0, 0, d, d, 180, 90);
                    path.AddArc(ss - d, 0, d, d, 270, 90);
                    path.AddArc(ss - d, ss - d, d, d, 0, 90);
                    path.AddArc(0, ss - d, d, d, 90, 90);
                    path.CloseFigure();
                    g.FillPath(br, path);
                    path.Dispose();
                }
            }

            int cp = int.Parse(cpHex, NumberStyles.HexNumber);
            string ch = char.ConvertFromUtf32(cp);
            float fontSize = (float)Math.Round(ss * (onPlate ? 0.60 : 0.92));
            using (var font = new Font(_family, fontSize, FontStyle.Regular, GraphicsUnit.Pixel))
            using (var tb = new SolidBrush(color))
            {
                var sfmt = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                var rect = new RectangleF(0, 0, ss, ss);
                g.DrawString(ch, font, tb, rect, sfmt);
                sfmt.Dispose();
            }
        }

        var res = new Bitmap(size, size);
        using (var rg = Graphics.FromImage(res))
        {
            rg.SmoothingMode = SmoothingMode.AntiAlias;
            rg.InterpolationMode = InterpolationMode.HighQualityBicubic;
            rg.Clear(Color.Transparent);
            rg.DrawImage(bmp, 0, 0, size, size);
        }
        bmp.Dispose();
        return res;
    }
}
'@

Add-Type -TypeDefinition $src -ReferencedAssemblies @('System.Drawing', 'System.Windows.Forms', 'System.Runtime.InteropServices', 'System.IO.Compression.FileSystem', 'System.Globalization', 'System.Core')
$png = Join-Path $outDir "icon-preview.png"
$dim = [IconPreview]::Run($ttf, $png)
Write-Host "saved: $png ($dim)"
