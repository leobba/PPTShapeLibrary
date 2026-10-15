# make-app-icon.py —— 把用户自制的图标转成 Windows 应用图标（多尺寸 .ico）
#
# 输入：用户提供的 WebP（扩展名写成了 .jpeg，实际是 RIFF/WEBP，GDI+ 读不了，用 Pillow）
# 处理：
#   1. 按 alpha 取内容边界，**裁掉多余透明留白**
#      （原图 2048x2048，内容仅 1528x1472，四周留白 12.5%~14%；
#       若不裁，做出来的图标在小尺寸下会显得比别的图标"小一圈"）
#   2. 按 5% 边距重新居中成正方形（保留 alpha）
#   3. 逐尺寸 LANCZOS 缩放（16/20/24/32/40/48/64/96/128/256）
#   4. 合成多尺寸 .ico（Pillow 会为 ≤128 写 BMP 条目、256 写 PNG 条目 —— 兼容性最好）
#   5. 输出预览图，便于检查小尺寸清晰度
#
# 用法：
#   python scripts\make-app-icon.py
import os
import sys
from PIL import Image, ImageDraw, ImageFont

SRC = r"E:\图片\fiGIMUyNz.jpeg"
ROOT = r"E:\code\PPTShapeLibrary"
OUT_ICO = os.path.join(ROOT, "assets", "app.ico")
OUT_PREVIEW_DIR = os.path.join(ROOT, "build", "icon-user")
SRC_COPY = os.path.join(ROOT, "assets", "app-icon-source.webp")

MARGIN = 0.05                       # 重新居中后的边距（内容占 90%）
SIZES = [256, 128, 96, 64, 48, 40, 32, 24, 20, 16]

os.makedirs(os.path.dirname(OUT_ICO), exist_ok=True)
os.makedirs(OUT_PREVIEW_DIR, exist_ok=True)

# ---------------------------------------------------------------- 1) 读入
im = Image.open(SRC)
print("source : %s  %s %dx%d mode=%s" % (os.path.basename(SRC), im.format, im.width, im.height, im.mode))
im = im.convert("RGBA")

# 顺手把源图原样拷一份到仓库（可复现）
# 注意：**不要重新编码**（曾试过存 lossless WebP，体积从 244KB 涨到 755KB）；
# 原文件已经是有损 WebP，直接复制字节即可。
try:
    import shutil
    shutil.copyfile(SRC, SRC_COPY)
    print("source copied -> %s (%d KB)" % (SRC_COPY, os.path.getsize(SRC_COPY) // 1024))
except Exception as e:
    print("copy source failed:", e)

# ---------------------------------------------------------------- 2) 裁到内容边界 + 重新居中
bbox = im.getchannel("A").getbbox()
print("content bbox : %s  ->  %dx%d" % (bbox, bbox[2] - bbox[0], bbox[3] - bbox[1]))
content = im.crop(bbox)

cw, ch = content.size
side = int(round(max(cw, ch) / (1.0 - 2.0 * MARGIN)))
canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
canvas.paste(content, ((side - cw) // 2, (side - ch) // 2), content)
print("normalized   : %dx%d  (content %.1f%% of side, margin %.0f%%)" % (
    side, side, max(cw, ch) * 100.0 / side, MARGIN * 100))

# ---------------------------------------------------------------- 3) 逐尺寸缩放
frames = {}
for s in SIZES:
    frames[s] = canvas.resize((s, s), Image.LANCZOS)

# ---------------------------------------------------------------- 4) 写 .ico
# 用 256 那帧作为基准，Pillow 会按 sizes 逐档生成合适的条目格式
frames[256].save(OUT_ICO, format="ICO", sizes=[(s, s) for s in SIZES])
print("ico          : %s  (%d KB, %d sizes)" % (OUT_ICO, os.path.getsize(OUT_ICO) // 1024, len(SIZES)))

# 单独存 PNG（方便别处引用 / 网页）
for s in SIZES:
    frames[s].save(os.path.join(OUT_PREVIEW_DIR, "icon-%d.png" % s))

# ---------------------------------------------------------------- 5) 预览条
def load_font(size):
    for p in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\segoeui.ttf", r"C:\Windows\Fonts\arial.ttf"):
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()

f_big = load_font(22)
f_small = load_font(13)

pad = 18
label_h = 26
zoom = 3                                   # 小尺寸放大展示（看清细节）
row1_h = 256
row2_h = max(SIZES) * zoom
W = pad + sum(s + pad for s in SIZES) + pad
H = pad + 40 + row1_h + 20 + label_h + row2_h + label_h + pad + 10

strip = Image.new("RGBA", (W, H), (245, 247, 250, 255))
d = ImageDraw.Draw(strip)
d.text((pad, 10), "app.ico  -  actual sizes (top) / 3x zoom (bottom)", font=f_big, fill=(27, 58, 87, 255))

x = pad
base1 = 46
for s in SIZES:
    strip.paste(frames[s], (x, base1 + (row1_h - s) // 2), frames[s])
    d.text((x, base1 + row1_h + 4), "%d" % s, font=f_small, fill=(110, 122, 136, 255))
    x += s + pad

x = pad
base2 = base1 + row1_h + 20 + label_h
for s in SIZES:
    z = s * zoom
    big = canvas.resize((z, z), Image.LANCZOS)
    strip.paste(big, (x, base2 + (row2_h - z) // 2), big)
    d.rectangle([x - 1, base2 + (row2_h - z) // 2 - 1, x + z, base2 + (row2_h - z) // 2 + z],
                outline=(210, 216, 224, 255))
    x += z + pad

strip_path = os.path.join(OUT_PREVIEW_DIR, "preview-strip.png")
strip.save(strip_path)
print("preview      : %s  (%dx%d)" % (strip_path, W, H))
