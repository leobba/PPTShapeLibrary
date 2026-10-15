# gen-icon-candidates-v4.ps1 —— 第四代图标：**用 PPT 素材**
#
# "用 PPT 素材"落实为两件事：
#   1. **PPT 原生形状**：圆角矩形 / 椭圆 / 三角形 / 五角星（就是 PPT「插入 → 形状」里的那些），
#      不再用 Material Icons 的通用字形；
#   2. **PPT 官方橙红**：从 POWERPNT.EXE 提取的真实图标采样得到（#D24726 一带），
#      让图标一眼关联到 PowerPoint。
#
# 6 种组合：
#   shapes-grid   4 个原生形状 2x2 排布（"形状素材库"）
#   slide-shapes  白色幻灯片 + 内含 3 个原生形状
#   slide-row     白色幻灯片 + 一排 4 个小形状
#   slide-big     白色幻灯片 + 一个大形状（"素材"为主体）
#   shapes-scatter 5 个形状错落排布（纯"素材库"）
#   real-materials 用你库里**真实素材的缩略图**拼 2x2（独特，但有隐私考虑）
#
# 用法：
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test\gen-icon-candidates-v4.ps1
param([string]$Root = "E:\code\PPTShapeLibrary")

$ErrorActionPreference = "Stop"

$src = @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Collections.Generic;

public static class IconGen4
{
    // ---------------- 基础 ----------------

    static GraphicsPath Rounded(RectangleF r, float radius)
    {
        var p = new GraphicsPath();
        if (radius <= 0.5f) { p.AddRectangle(r); return p; }
        float d = radius * 2;
        if (d > r.Width) d = r.Width;
        if (d > r.Height) d = r.Height;
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    static void FillRound(Graphics g, RectangleF r, float radius, Color c)
    {
        using (var path = Rounded(r, radius))
        using (var b = new SolidBrush(c))
            g.FillPath(b, path);
    }

    static void FillCircle(Graphics g, RectangleF r, Color c)
    {
        using (var b = new SolidBrush(c))
            g.FillEllipse(b, r);
    }

    static void FillTriangle(Graphics g, RectangleF r, Color c)
    {
        using (var b = new SolidBrush(c))
            g.FillPolygon(b, new[]
            {
                new PointF(r.X + r.Width / 2f, r.Y),
                new PointF(r.Right, r.Bottom),
                new PointF(r.X, r.Bottom)
            });
    }

    static void FillStar(Graphics g, RectangleF r, Color c)
    {
        float cx = r.X + r.Width / 2f, cy = r.Y + r.Height / 2f;
        float R = Math.Min(r.Width, r.Height) / 2f;
        float ri = R * 0.42f;
        var pts = new PointF[10];
        for (int i = 0; i < 10; i++)
        {
            double ang = -Math.PI / 2 + i * Math.PI / 5;
            float rad = (i % 2 == 0) ? R : ri;
            pts[i] = new PointF((float)(cx + rad * Math.Cos(ang)), (float)(cy + rad * Math.Sin(ang)));
        }
        using (var b = new SolidBrush(c))
            g.FillPolygon(b, pts);
    }

    // PPT 原生形状的"形状序号"：0 圆角矩形 1 椭圆 2 三角形 3 五角星
    static void FillNative(Graphics g, int kind, RectangleF r, Color c)
    {
        switch (kind)
        {
            case 0: FillRound(g, r, Math.Min(r.Width, r.Height) * 0.22f, c); break;
            case 1: FillCircle(g, r, c); break;
            case 2: FillTriangle(g, r, c); break;
            default: FillStar(g, r, c); break;
        }
    }

    static RectangleF SlideRect(RectangleF box, float scale)
    {
        float w = box.Width * scale;
        float h = w * 9f / 16f;
        if (h > box.Height * scale) { h = box.Height * scale; w = h * 16f / 9f; }
        return new RectangleF(box.X + (box.Width - w) / 2f, box.Y + (box.Height - h) / 2f, w, h);
    }

    /// <summary>在一行里均分 n 个正方形单元。</summary>
    static RectangleF[] RowCells(RectangleF box, int n, float gapRatio)
    {
        float gap = box.Width * gapRatio;
        float cw = (box.Width - gap * (n - 1)) / n;
        float s = Math.Min(cw, box.Height);
        float totalW = s * n + gap * (n - 1);
        float x = box.X + (box.Width - totalW) / 2f;
        float y = box.Y + (box.Height - s) / 2f;
        var res = new RectangleF[n];
        for (int i = 0; i < n; i++) res[i] = new RectangleF(x + i * (s + gap), y, s, s);
        return res;
    }

    // ---------------- 组合 ----------------

    static void Comp_ShapesGrid(Graphics g, RectangleF box, Color plate)
    {
        float gap = box.Width * 0.14f;
        float cw = (box.Width - gap) / 2f;
        float totalH = cw * 2 + gap;
        float y0 = box.Y + (box.Height - totalH) / 2f;
        int[] kinds = { 0, 1, 2, 3 };
        int k = 0;
        for (int i = 0; i < 2; i++)
            for (int j = 0; j < 2; j++)
            {
                var cell = new RectangleF(box.X + j * (cw + gap), y0 + i * (cw + gap), cw, cw);
                float inset = cw * 0.10f;
                FillNative(g, kinds[k++], new RectangleF(cell.X + inset, cell.Y + inset, cell.Width - inset * 2, cell.Height - inset * 2), Color.White);
            }
    }

    static void Comp_SlideShapes(Graphics g, RectangleF box, Color plate)
    {
        var s = SlideRect(box, 0.96f);
        FillRound(g, s, s.Height * 0.15f, Color.White);
        var inner = new RectangleF(s.X + s.Width * 0.16f, s.Y + s.Height * 0.22f, s.Width * 0.68f, s.Height * 0.56f);
        var cells = RowCells(inner, 3, 0.14f);
        int[] kinds = { 0, 1, 2 };
        for (int i = 0; i < 3; i++) FillNative(g, kinds[i], cells[i], plate);
    }

    static void Comp_SlideRow(Graphics g, RectangleF box, Color plate)
    {
        var s = SlideRect(box, 0.96f);
        FillRound(g, s, s.Height * 0.15f, Color.White);
        var inner = new RectangleF(s.X + s.Width * 0.12f, s.Y + s.Height * 0.28f, s.Width * 0.76f, s.Height * 0.44f);
        var cells = RowCells(inner, 4, 0.16f);
        int[] kinds = { 3, 2, 1, 0 };
        for (int i = 0; i < 4; i++) FillNative(g, kinds[i], cells[i], plate);
    }

    static void Comp_SlideBig(Graphics g, RectangleF box, Color plate)
    {
        var s = SlideRect(box, 0.96f);
        FillRound(g, s, s.Height * 0.15f, Color.White);
        float sz = s.Height * 0.62f;
        var c = new RectangleF(s.X + (s.Width - sz) / 2f, s.Y + (s.Height - sz) / 2f, sz, sz);
        FillNative(g, 3, c, plate);      // 大五角星 = "精华素材"
    }

    static void Comp_ShapesScatter(Graphics g, RectangleF box, Color plate)
    {
        // 5 个形状错落：大星居中偏上，四角小形状
        float big = box.Width * 0.42f;
        FillNative(g, 3, new RectangleF(box.X + box.Width * 0.30f, box.Y + box.Height * 0.10f, big, big), Color.White);
        float sm = box.Width * 0.26f;
        FillNative(g, 0, new RectangleF(box.X + box.Width * 0.04f, box.Y + box.Height * 0.52f, sm, sm), Color.White);
        FillNative(g, 1, new RectangleF(box.X + box.Width * 0.68f, box.Y + box.Height * 0.50f, sm, sm), Color.White);
        FillNative(g, 2, new RectangleF(box.X + box.Width * 0.30f, box.Y + box.Height * 0.68f, sm * 0.9f, sm * 0.9f), Color.White);
        float xs = box.Width * 0.17f;
        FillNative(g, 1, new RectangleF(box.X + box.Width * 0.66f, box.Y + box.Height * 0.16f, xs, xs), Color.White);
    }

    public static Bitmap Make(string comp, Color plate, int size, List<string> realPreviews)
    {
        int ss = size * 4;
        var big = new Bitmap(ss, ss);
        using (var g = Graphics.FromImage(big))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.Clear(Color.Transparent);

            bool lightPlate = (comp == "real-materials");
            Color plateCol = lightPlate ? Color.FromArgb(0xF7, 0xF9, 0xFC) : plate;
            FillRound(g, new RectangleF(0, 0, ss, ss), ss * 0.215f, plateCol);

            float pad = ss * 0.195f;
            var box = new RectangleF(pad, pad, ss - pad * 2f, ss - pad * 2f);

            if (comp == "real-materials")
            {
                // 2x2 真实素材缩略图 + 橙色细框
                FillRound(g, new RectangleF(ss * 0.055f, ss * 0.055f, ss * 0.89f, ss * 0.89f), ss * 0.16f, plate);
                var inner = new RectangleF(ss * 0.085f, ss * 0.085f, ss * 0.83f, ss * 0.83f);
                float gap = inner.Width * 0.06f;
                float cw = (inner.Width - gap) / 2f;
                int idx = 0;
                for (int i = 0; i < 2; i++)
                    for (int j = 0; j < 2; j++)
                    {
                        var cell = new RectangleF(inner.X + j * (cw + gap), inner.Y + i * (cw + gap), cw, cw);
                        FillRound(g, cell, cw * 0.12f, Color.White);
                        if (realPreviews != null && idx < realPreviews.Count)
                        {
                            try
                            {
                                using (var im = Image.FromFile(realPreviews[idx]))
                                {
                                    float ip = cw * 0.10f;
                                    var target = new RectangleF(cell.X + ip, cell.Y + ip, cell.Width - ip * 2, cell.Height - ip * 2);
                                    // 按宽高比居中缩放
                                    float sc = Math.Min(target.Width / im.Width, target.Height / im.Height);
                                    float dw = im.Width * sc, dh = im.Height * sc;
                                    var dr = new RectangleF(target.X + (target.Width - dw) / 2f, target.Y + (target.Height - dh) / 2f, dw, dh);
                                    g.DrawImage(im, dr);
                                }
                            }
                            catch { }
                        }
                        idx++;
                    }
            }
            else if (comp == "shapes-grid") Comp_ShapesGrid(g, box, plate);
            else if (comp == "slide-shapes") Comp_SlideShapes(g, box, plate);
            else if (comp == "slide-row") Comp_SlideRow(g, box, plate);
            else if (comp == "slide-big") Comp_SlideBig(g, box, plate);
            else if (comp == "shapes-scatter") Comp_ShapesScatter(g, box, plate);
        }
        var res = new Bitmap(size, size);
        using (var rg = Graphics.FromImage(res))
        {
            rg.SmoothingMode = SmoothingMode.AntiAlias;
            rg.InterpolationMode = InterpolationMode.HighQualityBicubic;
            rg.Clear(Color.Transparent);
            rg.DrawImage(big, 0, 0, size, size);
        }
        big.Dispose();
        return res;
    }

    public static void SaveIco(string comp, Color plate, string path, int[] sizes, List<string> realPreviews)
    {
        var imgs = new Bitmap[sizes.Length];
        for (int i = 0; i < sizes.Length; i++) imgs[i] = Make(comp, plate, sizes[i], realPreviews);
        using (var ms = new MemoryStream())
        using (var bw = new BinaryWriter(ms))
        {
            bw.Write((ushort)0); bw.Write((ushort)1); bw.Write((ushort)imgs.Length);
            int offset = 6 + 16 * imgs.Length;
            var blobs = new byte[imgs.Length][];
            for (int i = 0; i < imgs.Length; i++)
                using (var m2 = new MemoryStream()) { imgs[i].Save(m2, ImageFormat.Png); blobs[i] = m2.ToArray(); }
            for (int i = 0; i < imgs.Length; i++)
            {
                int w = imgs[i].Width, h = imgs[i].Height;
                bw.Write((byte)(w >= 256 ? 0 : w));
                bw.Write((byte)(h >= 256 ? 0 : h));
                bw.Write((byte)0); bw.Write((byte)0);
                bw.Write((ushort)1); bw.Write((ushort)32);
                bw.Write((uint)blobs[i].Length);
                bw.Write((uint)offset);
                offset += blobs[i].Length;
            }
            foreach (var b in blobs) bw.Write(b);
            bw.Flush();
            File.WriteAllBytes(path, ms.ToArray());
        }
        foreach (var im in imgs) im.Dispose();
    }

    public static string Sheet(string outPng, string[] comps, string[] colorNames, Color[] colors,
                              List<string> realPreviews, int cell, int sizePx)
    {
        int cols = colors.Length, rows = comps.Length;
        int labelW = 250, headH = 110, rowLabelH = 24;
        int W = labelW + cols * (cell + 16) + 24;
        int H = headH + rows * (cell + rowLabelH + 14) + 40;
        using (var bmp = new Bitmap(W, H))
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.FromArgb(0xF5, 0xF7, 0xFA));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            var fT = new Font("Microsoft YaHei UI", 17, FontStyle.Bold);
            var fH = new Font("Microsoft YaHei UI", 10, FontStyle.Bold);
            var fR = new Font("Microsoft YaHei UI", 12, FontStyle.Bold);
            var fS = new Font("Microsoft YaHei UI", 9);
            var dark = new SolidBrush(Color.FromArgb(0x1B, 0x3A, 0x57));
            var gray = new SolidBrush(Color.FromArgb(0x70, 0x7A, 0x88));

            g.DrawString("PPTShapeLibrary v4  -  PPT native shapes + real PPT brand color", fT, dark, 20, 16);
            g.DrawString("shapes from PowerPoint's own shape gallery; orange sampled from POWERPNT.EXE (#D24726)", fS, gray, 20, 48);
            g.DrawString("each cell: 128 / 32 / 16 px", fS, gray, 20, 66);

            for (int c = 0; c < cols; c++)
            {
                int x = labelW + c * (cell + 16);
                g.DrawString(colorNames[c], fH, dark, x, 80);
                using (var sw = new SolidBrush(colors[c])) g.FillRectangle(sw, x, 98, 40, 8);
            }

            int y = headH;
            for (int r = 0; r < rows; r++)
            {
                g.DrawString(comps[r], fR, dark, 20, y + 50);
                for (int c = 0; c < cols; c++)
                {
                    int x = labelW + c * (cell + 16);
                    using (var img = Make(comps[r], colors[c], sizePx, realPreviews)) g.DrawImage(img, x, y + 10, sizePx, sizePx);
                    using (var i32 = Make(comps[r], colors[c], 32, realPreviews)) g.DrawImage(i32, x + sizePx + 8, y + 10, 32, 32);
                    using (var i16 = Make(comps[r], colors[c], 16, realPreviews)) g.DrawImage(i16, x + sizePx + 8, y + 48, 16, 16);
                }
                y += cell + rowLabelH + 14;
            }
            fT.Dispose(); fH.Dispose(); fR.Dispose(); fS.Dispose(); dark.Dispose(); gray.Dispose();
            bmp.Save(outPng, ImageFormat.Png);
            return W + "x" + H;
        }
    }
}
'@

Add-Type -TypeDefinition $src -ReferencedAssemblies @('System.Drawing')

$outDir = Join-Path $Root "build\icon-v4"
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$icoDir = Join-Path $outDir "ico"
New-Item -ItemType Directory -Path $icoDir -Force | Out-Null

# 挑 4 张"真实素材"缩略图：优先接近正方形、体积适中（避免空图）
$itemsDir = Join-Path $env:APPDATA "PPTShapeLibrary\items"
$real = New-Object System.Collections.Generic.List[string]
try {
    Add-Type -AssemblyName System.Drawing
    $cands = Get-ChildItem $itemsDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $p = Join-Path $_.FullName "preview.png"
        if (Test-Path $p) {
            $f = Get-Item $p
            if ($f.Length -gt 8KB) {
                try {
                    $im = [System.Drawing.Image]::FromFile($p)
                    $ar = $im.Width / [double]$im.Height
                    $im.Dispose()
                    [pscustomobject]@{ Path = $p; Size = $f.Length; Aspect = $ar; Score = [math]::Abs([math]::Log($ar)) }
                } catch { }
            }
        }
    } | Sort-Object Score, { -$_.Size } | Select-Object -First 4
    foreach ($c in $cands) { $real.Add($c.Path) }
    Write-Host ("real previews picked: {0}" -f $real.Count)
} catch { Write-Host "real preview pick failed: $($_.Exception.Message)" }

$comps = @("shapes-grid", "slide-shapes", "slide-row", "slide-big", "shapes-scatter", "real-materials")
$colorNames = @("PPT-orange", "brand-blue", "indigo", "graphite")
$colors = @(
    [System.Drawing.Color]::FromArgb(0xD2, 0x47, 0x26),   # PowerPoint 官方橙红（从 POWERPNT.EXE 采样）
    [System.Drawing.Color]::FromArgb(0x1B, 0x5E, 0xAB),
    [System.Drawing.Color]::FromArgb(0x5E, 0x35, 0xB1),
    [System.Drawing.Color]::FromArgb(0x37, 0x47, 0x4F)
)

$png = Join-Path $outDir "icon-v4-sheet.png"
$dim = [IconGen4]::Sheet($png, $comps, $colorNames, $colors, $real, 132, 128)
Write-Host "sheet: $png  ($dim)"

$sizes = @(256, 64, 48, 32, 24, 16)
$n = 0
foreach ($c in $comps) {
    for ($i = 0; $i -lt $colors.Count; $i++) {
        [IconGen4]::SaveIco($c, $colors[$i], (Join-Path $icoDir ("{0}-{1}.ico" -f $c, $colorNames[$i])), $sizes, $real)
        $n++
    }
}
Write-Host "icons: $n files in $icoDir"
