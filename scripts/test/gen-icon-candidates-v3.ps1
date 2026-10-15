# gen-icon-candidates-v3.ps1 —— 第三代图标候选：**幻灯片主体 + 具体元素徽标**
#
# 设计（响应"更具体的元素"）：主图仍是 16:9「幻灯片」（一眼是 PPT），
# 右下角叠一个**可辨识的元素徽标**（圆底 + 白色图形），表达"素材库能做什么"。
#
# 8 种元素：
#   wand      魔术棒（一键美化）
#   puzzle    拼图块（组件化）
#   book      书（素材书/知识库）
#   star      星标（精选）
#   blocks    积木（素材块）
#   folder    文件夹（素材库）
#   bookmark  书签（收藏）
#   box       素材盒（收纳）
#
# 用法：
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test\gen-icon-candidates-v3.ps1
param([string]$Root = "E:\code\PPTShapeLibrary")

$ErrorActionPreference = "Stop"

$src = @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

public static class IconGen3
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

    static void FillCircle(Graphics g, float cx, float cy, float rad, Color c)
    {
        using (var b = new SolidBrush(c))
            g.FillEllipse(b, cx - rad, cy - rad, rad * 2, rad * 2);
    }

    /// <summary>五角星（10 顶点交替长短半径）。</summary>
    static void FillStar(Graphics g, float cx, float cy, float R, float r, Color c, float rotDeg)
    {
        var pts = new PointF[10];
        double rot = rotDeg * Math.PI / 180.0;
        for (int i = 0; i < 10; i++)
        {
            double ang = -Math.PI / 2 + rot + i * Math.PI / 5;
            float rad = (i % 2 == 0) ? R : r;
            pts[i] = new PointF((float)(cx + rad * Math.Cos(ang)), (float)(cy + rad * Math.Sin(ang)));
        }
        using (var b = new SolidBrush(c))
            g.FillPolygon(b, pts);
    }

    /// <summary>旋转的粗线段（圆头），用于画魔术棒。</summary>
    static void FillCapsule(Graphics g, PointF a, PointF b, float thick, Color c)
    {
        using (var pen = new Pen(c, thick))
        {
            pen.StartCap = LineCap.Round;
            pen.EndCap = LineCap.Round;
            g.DrawLine(pen, a, b);
        }
    }

    static RectangleF SlideRect(RectangleF box, float scale)
    {
        float w = box.Width * scale;
        float h = w * 9f / 16f;
        if (h > box.Height * scale) { h = box.Height * scale; w = h * 16f / 9f; }
        float x = box.X + (box.Width - w) / 2f;
        float y = box.Y + (box.Height - h) / 2f;
        return new RectangleF(x, y, w, h);
    }

    // ---------------- 元素徽标（画在圆形徽标内，白色）----------------

    static void Badge_Wand(Graphics g, float cx, float cy, float R)
    {
        // 斜置细棒 + 顶端小星
        float t = R * 0.20f;
        FillCapsule(g, new PointF(cx - R * 0.52f, cy + R * 0.55f), new PointF(cx + R * 0.22f, cy - R * 0.22f), t, Color.White);
        FillStar(g, cx + R * 0.44f, cy - R * 0.44f, R * 0.42f, R * 0.17f, Color.White, 0f);
    }

    static void Badge_Puzzle(Graphics g, float cx, float cy, float R, Color plate)
    {
        float s = R * 0.86f;                       // 主体方块边长
        var sq = new RectangleF(cx - s / 2f, cy - s / 2f, s, s);
        float rr = s * 0.14f;
        FillRound(g, sq, rr, Color.White);
        // 右侧凸起（卡扣）
        FillCircle(g, sq.Right, cy, s * 0.20f, Color.White);
        // 上方凹口（用底板色挖掉）
        FillCircle(g, cx, sq.Top, s * 0.20f, plate);
    }

    static void Badge_Book(Graphics g, float cx, float cy, float R, Color plate)
    {
        float w = R * 1.12f, h = R * 1.34f;
        var cover = new RectangleF(cx - w / 2f, cy - h / 2f, w, h);
        FillRound(g, cover, w * 0.14f, Color.White);
        // 书脊
        float spine = w * 0.20f;
        using (var b = new SolidBrush(plate))
            g.FillRectangle(b, cover.X + w * 0.16f, cover.Y + h * 0.10f, spine * 0.28f, h * 0.80f);
        // 两条"文字"线
        using (var b = new SolidBrush(Color.FromArgb(150, plate)))
        {
            g.FillRectangle(b, cover.X + w * 0.40f, cover.Y + h * 0.30f, w * 0.42f, h * 0.09f);
            g.FillRectangle(b, cover.X + w * 0.40f, cover.Y + h * 0.50f, w * 0.30f, h * 0.09f);
        }
    }

    static void Badge_Star(Graphics g, float cx, float cy, float R)
    {
        FillStar(g, cx, cy, R * 1.12f, R * 0.46f, Color.White, 0f);
    }

    static void Badge_Blocks(Graphics g, float cx, float cy, float R)
    {
        float s = R * 0.52f, gap = R * 0.10f;
        float rr = s * 0.20f;
        // 下排两块
        FillRound(g, new RectangleF(cx - s - gap / 2f, cy + gap / 2f, s, s), rr, Color.White);
        FillRound(g, new RectangleF(cx + gap / 2f, cy + gap / 2f, s, s), rr, Color.White);
        // 上排一块（居中）
        FillRound(g, new RectangleF(cx - s / 2f, cy - s - gap / 2f, s, s), rr, Color.White);
    }

    static void Badge_Folder(Graphics g, float cx, float cy, float R, Color plate)
    {
        float w = R * 1.34f, h = R * 0.96f;
        var back = new RectangleF(cx - w / 2f, cy - h / 2f, w, h);
        // 顶部标签
        var tab = new RectangleF(back.X, back.Y - h * 0.22f, w * 0.46f, h * 0.30f);
        FillRound(g, tab, h * 0.10f, Color.White);
        FillRound(g, back, h * 0.14f, Color.White);
        // 前层分割线
        using (var b = new SolidBrush(plate))
            g.FillRectangle(b, back.X + w * 0.10f, back.Y + h * 0.34f, w * 0.80f, h * 0.10f);
    }

    static void Badge_Bookmark(Graphics g, float cx, float cy, float R, Color plate)
    {
        float w = R * 0.82f, h = R * 1.36f;
        var body = new RectangleF(cx - w / 2f, cy - h / 2f, w, h);
        FillRound(g, body, w * 0.16f, Color.White);
        // 底部 V 形缺口（用底板色三角挖掉）
        using (var br = new SolidBrush(plate))
            g.FillPolygon(br, new[]
            {
                new PointF(body.X, body.Bottom),
                new PointF(body.Right, body.Bottom),
                new PointF(cx, body.Bottom - h * 0.34f)
            });
    }

    static void Badge_Box(Graphics g, float cx, float cy, float R, Color plate)
    {
        float w = R * 1.24f, h = R * 0.80f;
        // 盒身
        var body = new RectangleF(cx - w / 2f, cy - h * 0.10f, w, h);
        FillRound(g, body, h * 0.16f, Color.White);
        // 盒盖（稍宽，错开在上方）
        float lw = w * 1.10f, lh = h * 0.34f;
        var lid = new RectangleF(cx - lw / 2f, body.Y - lh * 1.05f, lw, lh);
        FillRound(g, lid, lh * 0.30f, Color.White);
        // 中间分隔
        using (var b = new SolidBrush(plate))
            g.FillRectangle(b, body.X + w * 0.08f, body.Y + h * 0.22f, w * 0.84f, h * 0.10f);
    }

    // ---------------- 组装：幻灯片 + 徽标 ----------------

    static void DrawBadge(Graphics g, string mot, float cx, float cy, float R, Color plate)
    {
        // 徽标底：底板色圆（把幻灯片"挖"出一个圆），外加一圈底板色的"环"做分隔
        FillCircle(g, cx, cy, R * 1.20f, plate);
        switch (mot)
        {
            case "wand":     Badge_Wand(g, cx, cy, R); break;
            case "puzzle":   Badge_Puzzle(g, cx, cy, R, plate); break;
            case "book":     Badge_Book(g, cx, cy, R, plate); break;
            case "star":     Badge_Star(g, cx, cy, R); break;
            case "blocks":   Badge_Blocks(g, cx, cy, R); break;
            case "folder":   Badge_Folder(g, cx, cy, R, plate); break;
            case "bookmark": Badge_Bookmark(g, cx, cy, R, plate); break;
            case "box":      Badge_Box(g, cx, cy, R, plate); break;
        }
    }

    public static Bitmap Make(string mot, Color plate, int size)
    {
        int ss = size * 4;
        var big = new Bitmap(ss, ss);
        using (var g = Graphics.FromImage(big))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.Clear(Color.Transparent);

            FillRound(g, new RectangleF(0, 0, ss, ss), ss * 0.215f, plate);

            float pad = ss * 0.195f;
            var box = new RectangleF(pad, pad, ss - pad * 2f, ss - pad * 2f);

            if (mot == "wand")
            {
                // 魔术棒直接斜跨幻灯片（不做徽标），更醒目
                var s = SlideRect(box, 0.84f);
                var s2 = new RectangleF(s.X, s.Y - box.Height * 0.03f, s.Width, s.Height);
                FillRound(g, s2, s2.Height * 0.14f, Color.White);
                using (var b = new SolidBrush(Color.FromArgb(135, plate)))
                    g.FillRectangle(b, s2.X + s2.Width * 0.13f, s2.Y + s2.Height * 0.34f, s2.Width * 0.52f, s2.Height * 0.15f);

                float th = box.Width * 0.115f;
                var a = new PointF(box.X + box.Width * 0.14f, box.Bottom - box.Height * 0.10f);
                var c = new PointF(box.Right - box.Width * 0.20f, box.Y + box.Height * 0.22f);
                // 先画一条底板色的粗线做"描边分隔"，再画白色棒身
                FillCapsule(g, a, c, th * 1.85f, plate);
                FillCapsule(g, a, c, th, Color.White);
                // 顶端大星 + 两颗小星
                FillStar(g, c.X + box.Width * 0.10f, c.Y - box.Height * 0.10f, box.Width * 0.24f, box.Width * 0.10f, Color.White, 0f);
                FillStar(g, box.X + box.Width * 0.18f, box.Y + box.Height * 0.16f, box.Width * 0.085f, box.Width * 0.035f, Color.White, 0f);
                FillStar(g, box.Right - box.Width * 0.10f, box.Bottom - box.Height * 0.16f, box.Width * 0.070f, box.Width * 0.029f, Color.White, 0f);
            }
            else
            {
                // 幻灯片偏左上，右下留徽标位
                var inner = new RectangleF(box.X, box.Y, box.Width * 0.90f, box.Height * 0.90f);
                var s = SlideRect(inner, 0.96f);
                FillRound(g, s, s.Height * 0.14f, Color.White);
                using (var b = new SolidBrush(Color.FromArgb(130, plate)))
                    g.DrawRectangle(new Pen(Color.FromArgb(120, plate), s.Height * 0.07f),
                        s.X + s.Width * 0.12f, s.Y + s.Height * 0.30f, s.Width * 0.40f, s.Height * 0.34f);

                float R = box.Width * 0.285f;
                DrawBadge(g, mot, box.Right - R * 0.92f, box.Bottom - R * 0.92f, R, plate);
            }
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

    public static void SaveIco(string mot, Color plate, string path, int[] sizes)
    {
        var imgs = new Bitmap[sizes.Length];
        for (int i = 0; i < sizes.Length; i++) imgs[i] = Make(mot, plate, sizes[i]);
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

    public static string Sheet(string outPng, string[] mots, string[] colorNames, Color[] colors, int cell, int sizePx)
    {
        int cols = colors.Length, rows = mots.Length;
        int labelW = 220, headH = 96, rowLabelH = 26;
        int W = labelW + cols * (cell + 16) + 24;
        int H = headH + rows * (cell + rowLabelH + 14) + 40;
        using (var bmp = new Bitmap(W, H))
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.FromArgb(0xF5, 0xF7, 0xFA));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            var fTitle = new Font("Microsoft YaHei UI", 17, FontStyle.Bold);
            var fHead = new Font("Microsoft YaHei UI", 10, FontStyle.Bold);
            var fRow = new Font("Microsoft YaHei UI", 12, FontStyle.Bold);
            var fSmall = new Font("Microsoft YaHei UI", 9);
            var dark = new SolidBrush(Color.FromArgb(0x1B, 0x3A, 0x57));
            var gray = new SolidBrush(Color.FromArgb(0x70, 0x7A, 0x88));

            g.DrawString("PPTShapeLibrary icon candidates v3  -  slide + element badge x 6 colors", fTitle, dark, 20, 16);
            g.DrawString("each cell: 128px / 32px / 16px", fSmall, gray, 20, 46);

            for (int c = 0; c < cols; c++)
            {
                int x = labelW + c * (cell + 16);
                g.DrawString(colorNames[c], fHead, dark, x, 66);
                using (var sw = new SolidBrush(colors[c])) g.FillRectangle(sw, x, 84, 40, 8);
            }

            int y = headH;
            for (int r = 0; r < rows; r++)
            {
                g.DrawString(mots[r], fRow, dark, 20, y + 40);
                for (int c = 0; c < cols; c++)
                {
                    int x = labelW + c * (cell + 16);
                    using (var img = Make(mots[r], colors[c], sizePx)) g.DrawImage(img, x, y + 10, sizePx, sizePx);
                    using (var i32 = Make(mots[r], colors[c], 32)) g.DrawImage(i32, x + sizePx + 8, y + 10, 32, 32);
                    using (var i16 = Make(mots[r], colors[c], 16)) g.DrawImage(i16, x + sizePx + 8, y + 48, 16, 16);
                }
                y += cell + rowLabelH + 14;
            }
            fTitle.Dispose(); fHead.Dispose(); fRow.Dispose(); fSmall.Dispose(); dark.Dispose(); gray.Dispose();
            bmp.Save(outPng, ImageFormat.Png);
            return W + "x" + H;
        }
    }
}
'@

Add-Type -TypeDefinition $src -ReferencedAssemblies @('System.Drawing')

$outDir = Join-Path $Root "build\icon-v3"
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$icoDir = Join-Path $outDir "ico"
New-Item -ItemType Directory -Path $icoDir -Force | Out-Null

$mots = @("wand", "puzzle", "book", "star", "blocks", "folder", "bookmark", "box")
$colorNames = @("PPT-orange", "brand-blue", "indigo", "teal", "magenta", "graphite")
$colors = @(
    [System.Drawing.Color]::FromArgb(0xC4, 0x3E, 0x1C),
    [System.Drawing.Color]::FromArgb(0x1B, 0x5E, 0xAB),
    [System.Drawing.Color]::FromArgb(0x5E, 0x35, 0xB1),
    [System.Drawing.Color]::FromArgb(0x00, 0x79, 0x6B),
    [System.Drawing.Color]::FromArgb(0xC2, 0x18, 0x5B),
    [System.Drawing.Color]::FromArgb(0x37, 0x47, 0x4F)
)

$png = Join-Path $outDir "icon-v3-sheet.png"
$dim = [IconGen3]::Sheet($png, $mots, $colorNames, $colors, 132, 128)
Write-Host "sheet: $png  ($dim)"

$sizes = @(256, 64, 48, 32, 24, 16)
$n = 0
foreach ($m in $mots) {
    for ($c = 0; $c -lt $colors.Count; $c++) {
        [IconGen3]::SaveIco($m, $colors[$c], (Join-Path $icoDir ("{0}-{1}.ico" -f $m, $colorNames[$c])), $sizes)
        $n++
    }
}
Write-Host "icons: $n files in $icoDir"
