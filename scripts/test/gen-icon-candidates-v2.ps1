# gen-icon-candidates-v2.ps1 —— 第二代图标候选：**突出 PPT 主体 + 素材库**
#
# 设计思路（响应反馈："之前的不突出 PPT，像通用图标"）：
#   不再用 Material Icons 的通用字形（放大镜/网格/调色板），而是**自己画组合图形**：
#   以「幻灯片」(16:9 圆角矩形) 为**主体**，叠加「素材/库」的元素
#   （形状集合 / 幻灯片堆叠 / 素材架 / 添加徽标）。
#
# 6 种组合 × 6 种颜色 = 36 个候选，统一"圆角底板 + 白色图形"（B 风格）。
#
# 用法：
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test\gen-icon-candidates-v2.ps1
param([string]$Root = "E:\code\PPTShapeLibrary")

$ErrorActionPreference = "Stop"

$src = @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

public static class IconGen2
{
    // ---------------- 基础绘制 ----------------

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

    static void FillCircle(Graphics g, float cx, float cy, float radius, Color c)
    {
        using (var b = new SolidBrush(c))
            g.FillEllipse(b, cx - radius, cy - radius, radius * 2, radius * 2);
    }

    static void FillTriangle(Graphics g, PointF a, PointF b, PointF c, Color col)
    {
        using (var br = new SolidBrush(col))
            g.FillPolygon(br, new[] { a, b, c });
    }

    /// <summary>把 16:9 的"幻灯片"画进给定方框（居中），返回实际用的矩形。</summary>
    static RectangleF SlideRect(RectangleF box, float scale)
    {
        float w = box.Width * scale;
        float h = w * 9f / 16f;
        if (h > box.Height * scale) { h = box.Height * scale; w = h * 16f / 9f; }
        float x = box.X + (box.Width - w) / 2f;
        float y = box.Y + (box.Height - h) / 2f;
        return new RectangleF(x, y, w, h);
    }

    // ---------------- 6 种组合 ----------------
    // 约定：box 是"内容区"，plateColor 是底板色（用于镂空/对比），白 = 前景

    // A 幻灯片 + 内部 2x2 素材网格（"幻灯片里装着素材库"）
    static void Comp_SlideGrid(Graphics g, RectangleF box, Color plate)
    {
        var s = SlideRect(box, 0.92f);
        FillRound(g, s, s.Height * 0.14f, Color.White);
        // 内部 2x2 小方块（用底板色镂空）
        float pad = s.Height * 0.16f;
        float cw = (s.Width - pad * 3) / 2f;
        float chh = (s.Height - pad * 3) / 2f;
        float rr = Math.Min(cw, chh) * 0.22f;
        for (int i = 0; i < 2; i++)
            for (int j = 0; j < 2; j++)
            {
                var cell = new RectangleF(s.X + pad + j * (cw + pad), s.Y + pad + i * (chh + pad), cw, chh);
                FillRound(g, cell, rr, plate);
            }
    }

    // B 三张幻灯片堆叠（"一叠素材"）
    static void Comp_SlideStack(Graphics g, RectangleF box, Color plate)
    {
        float dx = box.Width * 0.075f;
        float dy = box.Height * 0.062f;
        // 后两张（半透明白 = 显得退后）
        var back2 = SlideRect(new RectangleF(box.X - dx * 1.6f, box.Y + dy * 1.6f, box.Width, box.Height), 0.80f);
        var back1 = SlideRect(new RectangleF(box.X - dx * 0.8f, box.Y + dy * 0.8f, box.Width, box.Height), 0.80f);
        FillRound(g, back2, back2.Height * 0.13f, Color.FromArgb(105, 255, 255, 255));
        FillRound(g, back1, back1.Height * 0.13f, Color.FromArgb(165, 255, 255, 255));
        // 最前一张实心白 + 底部一条底板色标题带
        var front = SlideRect(new RectangleF(box.X + dx * 0.4f, box.Y - dy * 0.4f, box.Width, box.Height), 0.80f);
        FillRound(g, front, front.Height * 0.13f, Color.White);
        float barH = front.Height * 0.17f;
        var bar = new RectangleF(front.X + front.Width * 0.12f, front.Bottom - barH * 1.5f, front.Width * 0.52f, barH);
        FillRound(g, bar, barH * 0.3f, plate);
    }

    // C 幻灯片 + 内部 形状三件套（圆/方/三角 = "形状素材库"）
    static void Comp_SlideShapes(Graphics g, RectangleF box, Color plate)
    {
        var s = SlideRect(box, 0.94f);
        FillRound(g, s, s.Height * 0.14f, Color.White);
        float cy = s.Y + s.Height * 0.54f;
        float r = s.Height * 0.19f;
        float gap = s.Width * 0.055f;
        float totalW = r * 2 * 3 + gap * 2;
        float x = s.X + (s.Width - totalW) / 2f;
        FillCircle(g, x + r, cy, r, plate);
        x += r * 2 + gap;
        FillRound(g, new RectangleF(x, cy - r, r * 2, r * 2), r * 0.28f, plate);
        x += r * 2 + gap;
        FillTriangle(g, new PointF(x + r, cy - r), new PointF(x + r * 2, cy + r), new PointF(x, cy + r), plate);
    }

    // D 2x2 幻灯片方阵（"一整库幻灯片"）
    static void Comp_DeckGrid(Graphics g, RectangleF box, Color plate)
    {
        float gapX = box.Width * 0.10f;
        float gapY = box.Height * 0.10f;
        float cw = (box.Width - gapX) / 2f;
        float chh = cw * 9f / 16f;
        float totalH = chh * 2 + gapY;
        float y0 = box.Y + (box.Height - totalH) / 2f;
        float x0 = box.X;
        var alphas = new[] { 255, 150, 150, 205 };
        int k = 0;
        for (int i = 0; i < 2; i++)
            for (int j = 0; j < 2; j++)
            {
                var cell = new RectangleF(x0 + j * (cw + 0), y0 + i * (chh + gapY), cw, chh);
                FillRound(g, cell, chh * 0.16f, Color.FromArgb(alphas[k], 255, 255, 255));
                // 每张小幻灯片里一道"内容条"
                var bar = new RectangleF(cell.X + cell.Width * 0.14f, cell.Y + cell.Height * 0.52f, cell.Width * 0.72f, cell.Height * 0.20f);
                using (var b = new SolidBrush(Color.FromArgb(150, plate)))
                    g.FillRectangle(b, bar);
                k++;
            }
    }

    // E 幻灯片 + 右下角"加入库"圆徽标（+）
    static void Comp_SlidePlus(Graphics g, RectangleF box, Color plate)
    {
        var s = SlideRect(new RectangleF(box.X, box.Y, box.Width * 0.86f, box.Height * 0.86f), 0.94f);
        FillRound(g, s, s.Height * 0.14f, Color.White);
        // 幻灯片内部两条内容线
        using (var b = new SolidBrush(Color.FromArgb(135, plate)))
        {
            g.FillRectangle(b, s.X + s.Width * 0.14f, s.Y + s.Height * 0.34f, s.Width * 0.56f, s.Height * 0.15f);
            g.FillRectangle(b, s.X + s.Width * 0.14f, s.Y + s.Height * 0.60f, s.Width * 0.38f, s.Height * 0.15f);
        }
        // 徽标：底板色圆 + 白色 +
        float br = box.Width * 0.20f;
        float bcx = box.Right - br * 0.92f;
        float bcy = box.Bottom - br * 0.92f;
        FillCircle(g, bcx, bcy, br, plate);
        float t = br * 0.22f;
        FillRound(g, new RectangleF(bcx - br * 0.52f, bcy - t, br * 1.04f, t * 2f), t * 0.6f, Color.White);
        FillRound(g, new RectangleF(bcx - t, bcy - br * 0.52f, t * 2f, br * 1.04f), t * 0.6f, Color.White);
    }

    // F 幻灯片 + 下方"素材架"（两条横条 = 书架/资源架）
    static void Comp_SlideShelf(Graphics g, RectangleF box, Color plate)
    {
        var s = SlideRect(new RectangleF(box.X, box.Y, box.Width, box.Height * 0.72f), 0.94f);
        FillRound(g, s, s.Height * 0.14f, Color.White);
        using (var b = new SolidBrush(Color.FromArgb(130, plate)))
            g.FillRectangle(b, s.X + s.Width * 0.14f, s.Y + s.Height * 0.34f, s.Width * 0.50f, s.Height * 0.16f);
        // 书架横条
        float barH = box.Height * 0.085f;
        float y = s.Bottom + box.Height * 0.10f;
        float w1 = box.Width * 0.86f;
        float w2 = box.Width * 0.58f;
        FillRound(g, new RectangleF(box.X + (box.Width - w1) / 2f, y, w1, barH), barH * 0.35f, Color.White);
        FillRound(g, new RectangleF(box.X + (box.Width - w2) / 2f, y + barH * 1.85f, w2, barH), barH * 0.35f, Color.White);
    }

    static void DrawComp(Graphics g, string comp, RectangleF box, Color plate)
    {
        switch (comp)
        {
            case "slide-grid":   Comp_SlideGrid(g, box, plate); break;
            case "slide-stack":  Comp_SlideStack(g, box, plate); break;
            case "slide-shapes": Comp_SlideShapes(g, box, plate); break;
            case "deck-grid":    Comp_DeckGrid(g, box, plate); break;
            case "slide-plus":   Comp_SlidePlus(g, box, plate); break;
            case "slide-shelf":  Comp_SlideShelf(g, box, plate); break;
        }
    }

    // ---------------- 生成单个图标 ----------------

    public static Bitmap Make(string comp, Color plate, int size)
    {
        int ss = size * 4;                       // 4x 超采样，小尺寸也锐利
        var big = new Bitmap(ss, ss);
        using (var g = Graphics.FromImage(big))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.Clear(Color.Transparent);
            // 圆角底板
            FillRound(g, new RectangleF(0, 0, ss, ss), ss * 0.215f, plate);
            // 内容区（四周留白）
            float pad = ss * 0.205f;
            var box = new RectangleF(pad, pad, ss - pad * 2f, ss - pad * 2f);
            DrawComp(g, comp, box, plate);
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

    // ---------------- 多尺寸 .ico ----------------

    public static void SaveIco(string comp, Color plate, string path, int[] sizes)
    {
        var imgs = new Bitmap[sizes.Length];
        for (int i = 0; i < sizes.Length; i++) imgs[i] = Make(comp, plate, sizes[i]);
        using (var ms = new MemoryStream())
        using (var bw = new BinaryWriter(ms))
        {
            bw.Write((ushort)0); bw.Write((ushort)1); bw.Write((ushort)imgs.Length);
            int offset = 6 + 16 * imgs.Length;
            var blobs = new byte[imgs.Length][];
            for (int i = 0; i < imgs.Length; i++)
            {
                using (var m2 = new MemoryStream()) { imgs[i].Save(m2, ImageFormat.Png); blobs[i] = m2.ToArray(); }
            }
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

    // ---------------- 对比图 ----------------

    public static string Sheet(string outPng, string[] comps, string[] colorNames, Color[] colors,
                              string[] compLabels, int cell, int sizePx)
    {
        int cols = colors.Length, rows = comps.Length;
        int labelW = 250, headH = 96, rowLabelH = 30;
        int W = labelW + cols * (cell + 16) + 24;
        int H = headH + rows * (cell + rowLabelH + 16) + 40;

        using (var bmp = new Bitmap(W, H))
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.FromArgb(0xF5, 0xF7, 0xFA));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

            var fTitle = new Font("Microsoft YaHei UI", 17, FontStyle.Bold);
            var fHead  = new Font("Microsoft YaHei UI", 10, FontStyle.Bold);
            var fRow   = new Font("Microsoft YaHei UI", 11, FontStyle.Bold);
            var fSmall = new Font("Microsoft YaHei UI", 9);
            var dark = new SolidBrush(Color.FromArgb(0x1B, 0x3A, 0x57));
            var gray = new SolidBrush(Color.FromArgb(0x70, 0x7A, 0x88));

            g.DrawString("PPTShapeLibrary icon candidates v2  -  slide motif x 6 colors (plate style)", fTitle, dark, 20, 16);
            g.DrawString("each cell: 128px + 32px + 16px", fSmall, gray, 20, 46);

            // 色列表头
            for (int c = 0; c < cols; c++)
            {
                int x = labelW + c * (cell + 16);
                g.DrawString(colorNames[c], fHead, dark, x, 66);
                using (var sw = new SolidBrush(colors[c]))
                    g.FillRectangle(sw, x, 84, 40, 8);
            }

            int y = headH;
            for (int r = 0; r < rows; r++)
            {
                g.DrawString(comps[r], fRow, dark, 20, y + 10);
                g.DrawString(compLabels[r], fSmall, gray, 20, y + 34);
                for (int c = 0; c < cols; c++)
                {
                    int x = labelW + c * (cell + 16);
                    using (var img = Make(comps[r], colors[c], sizePx))
                        g.DrawImage(img, x, y + 14, sizePx, sizePx);
                    // 32 / 16 小尺寸同排（验证小尺寸清晰度）
                    using (var i32 = Make(comps[r], colors[c], 32))
                        g.DrawImage(i32, x + sizePx + 8, y + 14, 32, 32);
                    using (var i16 = Make(comps[r], colors[c], 16))
                        g.DrawImage(i16, x + sizePx + 8, y + 52, 16, 16);
                }
                y += cell + rowLabelH + 16;
            }

            fTitle.Dispose(); fHead.Dispose(); fRow.Dispose(); fSmall.Dispose();
            dark.Dispose(); gray.Dispose();
            bmp.Save(outPng, ImageFormat.Png);
            return W + "x" + H;
        }
    }
}
'@

Add-Type -TypeDefinition $src -ReferencedAssemblies @('System.Drawing')

$outDir = Join-Path $Root "build\icon-v2"
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$icoDir = Join-Path $outDir "ico"
New-Item -ItemType Directory -Path $icoDir -Force | Out-Null

$comps = @("slide-grid", "slide-stack", "slide-shapes", "deck-grid", "slide-plus", "slide-shelf")
$compLabels = @(
    "slide with 2x2 material grid",
    "stack of 3 slides",
    "slide with circle/square/triangle",
    "2x2 deck grid",
    "slide + add-to-library badge",
    "slide above library shelves"
)
$colorNames = @("PPT-orange", "brand-blue", "indigo", "teal", "magenta", "graphite")
$colors = @(
    [System.Drawing.Color]::FromArgb(0xC4, 0x3E, 0x1C),
    [System.Drawing.Color]::FromArgb(0x1B, 0x5E, 0xAB),
    [System.Drawing.Color]::FromArgb(0x5E, 0x35, 0xB1),
    [System.Drawing.Color]::FromArgb(0x00, 0x79, 0x6B),
    [System.Drawing.Color]::FromArgb(0xC2, 0x18, 0x5B),
    [System.Drawing.Color]::FromArgb(0x37, 0x47, 0x4F)
)

$png = Join-Path $outDir "icon-v2-sheet.png"
$dim = [IconGen2]::Sheet($png, $comps, $colorNames, $colors, $compLabels, 132, 128)
Write-Host "sheet: $png  ($dim)"

$sizes = @(256, 64, 48, 32, 24, 16)
$n = 0
for ($r = 0; $r -lt $comps.Count; $r++) {
    for ($c = 0; $c -lt $colors.Count; $c++) {
        $name = "{0}-{1}.ico" -f $comps[$r], $colorNames[$c]
        [IconGen2]::SaveIco($comps[$r], $colors[$c], (Join-Path $icoDir $name), $sizes)
        $n++
    }
}
Write-Host "icons: $n files in $icoDir"
