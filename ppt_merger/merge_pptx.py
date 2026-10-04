#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT 批量合成工具（基于本机 Microsoft PowerPoint COM）
=====================================================

把指定目录下的多个 .pptx 按文件名顺序合并成一个 PPT，完整保留每页的
图表、SmartArt、配色、版式等（PowerPoint 原生 InsertFromFile 机制）。
默认还会把"源文件名"以角标形式放到该文件每一页的左上角，方便识别。

要求：
  - Windows 且已安装 Microsoft PowerPoint（2010+ 均可，实测 16.0）
  - pip install pywin32

用法示例：
  python merge_pptx.py "D:\\PPT模板\\图表" -o "D:\\PPT模板\\图表合集.pptx"
  python merge_pptx.py "D:\\PPT模板\\图表" -o out.pptx --no-label     # 不加文件名角标
"""

import argparse
import glob
import os
import re
import sys
import tempfile

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# PowerPoint 常量（避免依赖 win32com.client.constants 生成）
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_ALIGN_LEFT = 1
LABEL_LEFT = 10.0     # pt，距左边缘
LABEL_TOP = 6.0       # pt，距上边缘
LABEL_HEIGHT = 18.0   # pt
LABEL_FONT_SIZE = 10
LABEL_COLOR = 0x8C8C8C  # OLE_COLOR 灰 (BGR: 140,140,140)


def natural_key(path):
    """自然排序 key：让 file2 < file10"""
    name = os.path.basename(path).lower()
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", name)]


def collect_files(src_dir, pattern="*.pptx"):
    files = glob.glob(os.path.join(src_dir, pattern))
    files = [f for f in files if os.path.isfile(f)]
    return sorted(files, key=natural_key)


def add_filename_label(slide, label):
    """在 slide 左上角添加一个文件名小字文本框。"""
    try:
        slide_w = float(slide.Parent.PageSetup.SlideWidth)  # 单位 pt
    except Exception:
        slide_w = 960.0
    box_w = min(600.0, max(120.0, slide_w * 0.5))

    tb = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL,
                                 LABEL_LEFT, LABEL_TOP, box_w, LABEL_HEIGHT)
    tf = tb.TextFrame
    tf.WordWrap = False
    tf.MarginLeft = 0
    tf.MarginTop = 0
    tf.MarginRight = 0
    tf.MarginBottom = 0
    tr = tf.TextRange
    tr.Text = label
    tr.Font.Size = LABEL_FONT_SIZE
    tr.Font.Color.RGB = LABEL_COLOR
    try:
        tr.ParagraphFormat.Alignment = MSO_ALIGN_LEFT
    except Exception:
        pass
    return tb


def merge(files, out_path, label=True):
    """
    用 PowerPoint COM 把 files 全部页顺序合并到 out_path。
    返回 (源文件数, 总页数, 失败列表)。
    """
    import win32com.client  # 延迟导入，方便无 Office 时给出友好提示

    if not files:
        raise ValueError("没有找到可合并的 PPT 文件")

    for f in files:
        if not f.lower().endswith(".pptx"):
            raise ValueError(f"仅支持 .pptx：{f}")

    out_abs = os.path.abspath(out_path)
    out_key = os.path.normcase(out_abs)
    for f in files:
        same_path = os.path.normcase(os.path.abspath(f)) == out_key
        try:
            same_path = same_path or os.path.samefile(f, out_abs)
        except OSError:
            pass
        if same_path:
            raise ValueError(f"输出路径与输入文件相同，拒绝覆盖源文件：{f}")

    out_dir = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(out_dir, exist_ok=True)
    temp_out = None

    app = None
    pres = None
    ok_files = []
    failed = []

    def label_range(start_idx, end_idx, text):
        """给第 start_idx..end_idx 张幻灯片(1-based)加角标"""
        if not label:
            return
        for sidx in range(start_idx, end_idx + 1):
            try:
                add_filename_label(pres.Slides(sidx), text)
            except Exception as e:
                print(f"      [角标失败] 第{sidx}页: {e}")

    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        # 以第一个文件为基底（自带母版/主题），其余文件顺序插入
        base = files[0]
        label_text = os.path.splitext(os.path.basename(base))[0]
        print(f"[1/{len(files)}] 基底: {os.path.basename(base)}")
        pres = app.Presentations.Open(base, ReadOnly=True, WithWindow=False)
        ok_files.append(base)
        label_range(1, pres.Slides.Count, label_text)

        for i, f in enumerate(files[1:], start=2):
            before = pres.Slides.Count
            try:
                pres.Slides.InsertFromFile(f, pres.Slides.Count)
                added = pres.Slides.Count - before
                label_text = os.path.splitext(os.path.basename(f))[0]
                label_range(before + 1, pres.Slides.Count, label_text)
                ok_files.append(f)
                print(f"[{i}/{len(files)}] +{added}页  左上角标: {label_text}")
            except Exception as e:
                failed.append((f, str(e)))
                print(f"[{i}/{len(files)}] 跳过(失败): {os.path.basename(f)} -> {e}")

        # Save beside the destination, then replace it only after PowerPoint has
        # completed and closed the new file. A failed save leaves any old output intact.
        fd, temp_out = tempfile.mkstemp(prefix=".pptshape_merge_", suffix=".pptx", dir=out_dir)
        os.close(fd)
        os.remove(temp_out)
        pres.SaveAs(temp_out)
        total = pres.Slides.Count
        pres.Close()
        pres = None
        os.replace(temp_out, out_abs)
        temp_out = None
        print(f"[保存] 共 {total} 页 -> {out_abs}")
        return len(ok_files), total, failed
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if temp_out and os.path.exists(temp_out):
            try:
                os.remove(temp_out)
            except OSError:
                pass


def main():
    ap = argparse.ArgumentParser(
        description="把目录下的 .pptx 按文件名顺序合并成一个 PPT，并在每页左上角标注来源文件名"
                    "（需本机安装 PowerPoint）")
    ap.add_argument("src_dir", help="源目录（如 D:\\PPT模板\\图表）")
    ap.add_argument("-o", "--out", default="",
                    help="输出文件路径（默认: <源目录>\\..\\<源目录名>合集.pptx）")
    ap.add_argument("--pattern", default="*.pptx",
                    help="匹配模式，默认 *.pptx（不递归子目录）")
    ap.add_argument("--no-label", action="store_true",
                    help="不加文件名角标，仅纯合并")
    args = ap.parse_args()

    src_dir = os.path.abspath(args.src_dir)
    if not os.path.isdir(src_dir):
        print(f"[错误] 目录不存在: {src_dir}")
        return 2

    out_path = args.out or os.path.join(
        os.path.dirname(src_dir), os.path.basename(src_dir.rstrip("\\/")) + "合集.pptx")

    files = collect_files(src_dir, args.pattern)
    print(f"[扫描] {src_dir}  找到 {len(files)} 个文件")
    for f in files:
        print("   -", os.path.basename(f))
    if not files:
        print("[错误] 没有匹配的文件，退出。")
        return 1

    if os.path.exists(os.path.abspath(out_path)):
        print(f"[提示] 输出文件已存在，将覆盖: {out_path}")

    try:
        n_ok, total, failed = merge(files, out_path, label=not args.no_label)
    except (ValueError, OSError) as e:
        print(f"[错误] {e}")
        return 2
    print("=" * 56)
    print(f"合并完成：成功 {n_ok}/{len(files)} 个文件，共 {total} 页")
    if failed:
        print("以下文件合并失败（已跳过）：")
        for f, e in failed:
            print(f"   - {f}: {e}")
    print(f"输出: {os.path.abspath(out_path)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
