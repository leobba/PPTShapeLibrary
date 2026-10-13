# -*- coding: utf-8 -*-
"""
整页插入保真度回归（守卫"可以不用剪贴板"这个前提）
================================================================
背景：整页插入自 2026-10 起改为只用 `Slides.InsertFromFile`（不再碰剪贴板），
**前提是库文件的设计/母版/版式已唯一改名**（`MakeDesignUniqueForLibrary` → `Lib<id>-*`）。
一旦有人去掉这个改名，PowerPoint 会按名复用目标文稿的默认设计 → 版式被重绑 → 外观走样，
而"整页插入"看起来仍然"成功"（不会报错），只会静默变样。所以必须有脚本守住这个前提。

做法：
  1. 造一个"默认 Office 主题"的干净目标文稿（正是名字冲突场景）；
  2. 目标 A：剪贴板路径（Slides(1).Copy → Slides.Paste）；目标 B：InsertFromFile；
  3. 解包逐项比对：插入页 slideN.xml、slideN.xml.rels 的版式指向、主题 clrScheme、
     设计/母版/版式名、媒体文件；
  4. 两张插入页导出同尺寸 PNG（960×540）做像素差。

判据（全部满足才算 PASS）：
  · 版式绑定相同、主题 clrScheme 相同、设计/母版/版式名集合相同、媒体文件名集合相同
  · 像素平均差 < 0.5/255
  · slide XML 除随机生成的 <p14:creationId> 外无差异（该 id 每次插入都不同，无语义影响）

用法：python scripts/verify-insert-fidelity.py [<库条目shape.pptx> ...]
      不带参数时自动取库中最近 3 个整页快照。
退出码：0=PASS，1=FAIL。
"""
import difflib
import glob
import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
try:
    import win32com.client
    from PIL import Image, ImageChops
except ImportError as e:
    print("需要 pywin32 与 Pillow：pip install pywin32 pillow  (%s)" % e)
    sys.exit(2)

APP = None
PREEXISTING = 0


def open_app():
    global APP, PREEXISTING
    APP = win32com.client.Dispatch("PowerPoint.Application")
    try:
        PREEXISTING = APP.Presentations.Count
    except Exception:
        PREEXISTING = 0
    return APP


def quit_app():
    """PowerPoint 的 COM 服务器是单实例：用户自己开着 PowerPoint 时绝不能 Quit。"""
    if PREEXISTING == 0:
        try:
            APP.Quit()
        except Exception:
            pass
    else:
        print("[info] 检测到 PowerPoint 已有打开的文稿，跳过 Quit（不干扰用户会话）")


def new_target(width, height):
    pres = APP.Presentations.Add(0)          # WithWindow = msoFalse → 零页文稿
    pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight = width, height
    pres.Slides.Add(1, 12)                   # 12 = ppLayoutBlank
    tb = pres.Slides(1).Shapes.AddTextbox(1, 40, 40, 300, 60)
    tb.TextFrame.TextRange.Text = "目标文稿占位"
    return pres


def run_clipboard(lib_path, out_path, w, h):
    pres = new_target(w, h)
    try:
        lib = APP.Presentations.Open(lib_path, ReadOnly=True, Untitled=False, WithWindow=False)
        try:
            ok = False
            for attempt in range(4):
                lib.Slides(1).Copy()
                time.sleep(0.35 + 0.2 * attempt)
                r = pres.Slides.Paste(pres.Slides.Count + 1)
                if r is not None and r.Count > 0:
                    ok = True
                    break
            if not ok:
                raise RuntimeError("剪贴板路径粘贴失败（本机剪贴板被占用？）")
            pres.SaveAs(out_path)
        finally:
            lib.Close()
    finally:
        pres.Close()


def run_insertfromfile(lib_path, out_path, w, h):
    pres = new_target(w, h)
    try:
        pres.Slides.InsertFromFile(lib_path, pres.Slides.Count, 1, 1)
        pres.SaveAs(out_path)
    finally:
        pres.Close()


def parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def inserted_slide(ps, marker="目标文稿占位"):
    for n in sorted(ps):
        if re.fullmatch(r"ppt/slides/slide\d+\.xml", n) and marker.encode("utf-8") not in ps[n]:
            return n
    return None


def layout_of(ps, slide_name):
    num = re.search(r"slide(\d+)\.xml$", slide_name).group(1)
    rels = "ppt/slides/_rels/slide%s.xml.rels" % num
    if rels not in ps:
        return None
    txt = ps[rels].decode("utf-8", "replace")
    m = re.search(r'Type="[^"]*slideLayout"[^>]*Target="([^"]+)"', txt) or \
        re.search(r'Target="([^"]+)"[^>]*Type="[^"]*slideLayout"', txt)
    return m.group(1) if m else None


def theme_scheme(ps):
    for n in sorted(ps):
        if re.fullmatch(r"ppt/theme/theme\d+\.xml", n):
            m = re.search(r"<a:clrScheme.*?</a:clrScheme>", ps[n].decode("utf-8", "replace"), re.S)
            if m:
                return re.sub(r"\s+", " ", m.group(0))
    return None


def names(ps):
    out = []
    for n in sorted(ps):
        if re.fullmatch(r"ppt/slideMasters/slideMaster\d+\.xml", n) or \
           re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", n):
            m = re.search(r'name="([^"]*)"', ps[n].decode("utf-8", "replace"))
            if m:
                out.append(m.group(1))
    return sorted(out)


def norm(b):
    t = b.decode("utf-8", "replace")
    return re.sub(r"\s+", " ", re.sub(r">\s+<", "><", t)).strip()


def xml_equivalent(xa, xb):
    """除 <p14:creationId>（每次插入随机生成）外是否完全一致。"""
    ta, tb = norm(xa), norm(xb)
    if ta == tb:
        return True, "完全相同"
    sa = re.findall(r"<[^>]+>|[^<]+", ta)
    sb = re.findall(r"<[^>]+>|[^<]+", tb)
    sm = difflib.SequenceMatcher(a=sa, b=sb, autojunk=False)
    diffs = [( "".join(sa[i1:i2]), "".join(sb[j1:j2]) )
             for tag, i1, i2, j1, j2 in sm.get_opcodes() if tag != "equal"]
    only_creation = all("creationId" in a or "creationId" in b for a, b in diffs)
    return only_creation, "%d 处差异(仅 creationId=%s) 相似度=%.6f" % (len(diffs), only_creation, sm.ratio())


def export_png(pptx, png, slide_no, w=960, h=540):
    pres = APP.Presentations.Open(pptx, ReadOnly=True, Untitled=False, WithWindow=False)
    try:
        pres.Slides(slide_no).Export(png, "PNG", w, h)
    finally:
        pres.Close()


def pixel_mean(a, b):
    ia, ib = Image.open(a).convert("RGB"), Image.open(b).convert("RGB")
    if ia.size != ib.size:
        ib = ib.resize(ia.size)
    d = ImageChops.difference(ia, ib)
    hist = d.histogram()
    total = sum(hist) * 3
    return (sum(i * hist[i] for i in range(256)) * 3 / total) if total else 0.0, d.getbbox()


def check_one(lib_path, work):
    pres = APP.Presentations.Open(lib_path, ReadOnly=True, Untitled=False, WithWindow=False)
    try:
        w, h = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
    finally:
        pres.Close()

    a, b = os.path.join(work, "clip.pptx"), os.path.join(work, "file.pptx")
    run_clipboard(lib_path, a, w, h)
    run_insertfromfile(lib_path, b, w, h)
    pa, pb = parts(a), parts(b)
    na, nb = inserted_slide(pa), inserted_slide(pb)
    if not na or not nb:
        return False, "未能定位插入页"
    xml_ok, xml_note = xml_equivalent(pa[na], pb[nb])
    la, lb = layout_of(pa, na), layout_of(pb, nb)
    ta, tb = theme_scheme(pa), theme_scheme(pb)
    na_, nb_ = names(pa), names(pb)
    ma = sorted(os.path.basename(x) for x in pa if x.startswith("ppt/media/"))
    mb = sorted(os.path.basename(x) for x in pb if x.startswith("ppt/media/"))
    png_a, png_b = os.path.join(work, "a.png"), os.path.join(work, "b.png")
    export_png(a, png_a, int(re.search(r"\d+", na).group(0)))
    export_png(b, png_b, int(re.search(r"\d+", nb).group(0)))
    mean, bbox = pixel_mean(png_a, png_b)

    checks = [
        ("版式绑定相同", la == lb, "%s / %s" % (la, lb)),
        ("主题 clrScheme 相同", ta == tb and ta is not None, ""),
        ("设计/母版/版式名相同", na_ == nb_, "%s" % (na_[:2],)),
        ("媒体文件相同", ma == mb, "%d 个" % len(ma)),
        ("slide XML 仅差 creationId", xml_ok, xml_note),
        ("像素平均差 < 0.5", mean < 0.5, "%.4f bbox=%s" % (mean, bbox)),
    ]
    for name, ok, detail in checks:
        print("   [%s] %-26s %s" % ("PASS" if ok else "FAIL", name, detail))
    return all(c[1] for c in checks), ""


def default_libs():
    items = os.path.join(os.environ["APPDATA"], "PPTShapeLibrary", "items")
    found = []
    for meta in glob.glob(os.path.join(items, "*", "meta.json")):
        try:
            j = json.load(open(meta, encoding="utf-8"))
        except Exception:
            continue
        if j.get("snapshotPage"):
            found.append((j.get("createdAt", ""), os.path.join(os.path.dirname(meta), "shape.pptx")))
    found.sort(reverse=True)
    return [p for _, p in found[:3] if os.path.exists(p)]


def main():
    libs = sys.argv[1:] or default_libs()
    if not libs:
        print("找不到整页快照条目（%APPDATA%\\PPTShapeLibrary\\items 下没有 snapshotPage）")
        return 2
    open_app()
    results = []
    try:
        for lib in libs:
            print("\n=== %s ===" % os.path.basename(os.path.dirname(lib)))
            work = tempfile.mkdtemp(prefix="insertfid_")
            try:
                ok, note = check_one(lib, work)
            except Exception as ex:
                ok, note = False, str(ex)
                print("   [FAIL] 异常: %s" % ex)
            finally:
                shutil.rmtree(work, ignore_errors=True)
            results.append((os.path.basename(os.path.dirname(lib)), ok, note))
    finally:
        quit_app()

    print("\n===== 汇总 =====")
    for name, ok, note in results:
        print("  %-26s %s %s" % (name, "PASS" if ok else "FAIL", note))
    all_ok = all(r[1] for r in results)
    print("RESULT:", "PASS" if all_ok else "FAIL")
    if not all_ok:
        print("提示：若'版式绑定/设计名'不一致，先检查库文件的 ppt/slideMasters/slideMaster1.xml 的 name")
        print("      是否带 Lib<id>- 前缀（MakeDesignUniqueForLibrary 是否还生效）。")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
