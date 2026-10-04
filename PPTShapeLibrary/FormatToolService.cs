using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTShapeLibrary
{
    /// <summary>
    /// PPT 版式/批处理工具的共享静态服务层（Web 桥与 WinForms 窗体共用）。
    /// 无 UI 依赖；结果统一为 TextToolService.ToolResult（Ok/Message/Changed），由调用方呈现。
    /// </summary>
    public static class FormatToolService
    {
        private static PowerPoint.Application App
        {
            get { return Globals.ThisAddIn.Application; }
        }

        /// <summary>清理瘦身选项。</summary>
        public class CleanOptions
        {
            public bool DelEmptySlides = true;
            public bool DelHiddenShapes;
            public bool DelOffSlideShapes;
            public bool ClearNotes = true;
            public bool RemoveActions;
        }

        private class CleanStats
        {
            public int EmptySlides, HiddenShapes, OffSlideShapes, Notes, Actions;
            public int Failed;
        }

        private static string StatsText(CleanStats st)
        {
            return "空白页 " + st.EmptySlides + " · 隐藏对象 " + st.HiddenShapes +
                   " · 页外对象 " + st.OffSlideShapes + " · 备注 " + st.Notes +
                   " · 动作链接 " + st.Actions + " · 失败 " + st.Failed;
        }

        // ================= ① 整稿导出 PNG/JPG =================

        /// <summary>逐页导出图片到 dir，命名 slide01.png…。format="PNG"/"JPG"；width=0 为原始分辨率。</summary>
        public static TextToolService.ToolResult ExportImages(string dir, string format, int width)
        {
            try
            {
                if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
                    return new TextToolService.ToolResult { Ok = false, Message = "导出目录不存在。" };
                var pres = App.ActivePresentation;
                if (pres == null)
                    return new TextToolService.ToolResult { Ok = false, Message = "没有打开的演示文稿。" };

                bool isJpg = "JPG".Equals(format, StringComparison.OrdinalIgnoreCase);
                int w = width > 0 ? width : 0;
                int done = 0;
                int failed = 0;
                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    string file = Path.Combine(dir, "slide" + i.ToString("D2") + "." + (isJpg ? "jpg" : "png"));
                    try { pres.Slides[i].Export(file, isJpg ? "JPG" : "PNG", w, 0); done++; }
                    catch { failed++; }
                }
                if (failed > 0 && done == 0)
                    return new TextToolService.ToolResult { Ok = false, Message = "导出失败：" + failed + " 页均无法导出（请检查目录权限/格式支持）。", Changed = 0 };
                return new TextToolService.ToolResult
                {
                    Ok = done > 0,
                    Message = "已导出 " + done + " 页到：" + dir + (failed > 0 ? "（" + failed + " 页失败）" : ""),
                    Changed = done
                };
            }
            catch (Exception ex)
            {
                return new TextToolService.ToolResult { Ok = false, Message = ex.Message };
            }
        }

        // ================= ② 清理瘦身 =================

        /// <summary>扫描统计（不实际删除）。</summary>
        public static TextToolService.ToolResult ScanClean(CleanOptions opts)
        {
            try
            {
                var st = CleanCore(opts, true);
                return new TextToolService.ToolResult
                {
                    Ok = st.Failed == 0,
                    Message = StatsText(st),
                    Changed = st.EmptySlides + st.HiddenShapes + st.OffSlideShapes + st.Notes + st.Actions
                };
            }
            catch (Exception ex)
            {
                return new TextToolService.ToolResult { Ok = false, Message = ex.Message };
            }
        }

        /// <summary>执行清理（真实删除/清空）。</summary>
        public static TextToolService.ToolResult RunClean(CleanOptions opts)
        {
            try
            {
                var st = CleanCore(opts, false);
                return new TextToolService.ToolResult
                {
                    Ok = st.Failed == 0,
                    Message = (st.Failed == 0 ? "清理完成：" : "清理部分完成：") + StatsText(st),
                    Changed = st.EmptySlides + st.HiddenShapes + st.OffSlideShapes + st.Notes + st.Actions
                };
            }
            catch (Exception ex)
            {
                return new TextToolService.ToolResult { Ok = false, Message = ex.Message };
            }
        }

        private static CleanStats CleanCore(CleanOptions o, bool dryRun)
        {
            var st = new CleanStats();
            var pres = App.ActivePresentation;
            if (pres == null) return st;
            var slideW = (float)pres.PageSetup.SlideWidth;
            var slideH = (float)pres.PageSetup.SlideHeight;

            int remainingSlides = pres.Slides.Count;
            for (int slideIndex = pres.Slides.Count; slideIndex >= 1; slideIndex--)
            {
                var slide = pres.Slides[slideIndex];
                for (int i = slide.Shapes.Count; i >= 1; i--)
                {
                    var shp = slide.Shapes[i];
                    try
                    {
                        if (o.DelHiddenShapes && shp.Visible == MsoTriState.msoFalse)
                        {
                            if (dryRun) st.HiddenShapes++;
                            else
                            {
                                try { shp.Delete(); st.HiddenShapes++; }
                                catch { st.Failed++; }
                            }
                            continue;
                        }
                        if (o.DelOffSlideShapes)
                        {
                            float l = 0, t = 0, r = 0, b = 0;
                            try { l = (float)shp.Left; t = (float)shp.Top; r = l + (float)shp.Width; b = t + (float)shp.Height; }
                            catch { st.Failed++; continue; }
                            // 完全在画布外才删，避免误伤半出画布对象
                            if (r < -5 || b < -5 || l > slideW + 5 || t > slideH + 5)
                            {
                                if (dryRun) st.OffSlideShapes++;
                                else
                                {
                                    try { shp.Delete(); st.OffSlideShapes++; }
                                    catch { st.Failed++; }
                                }
                                continue;
                            }
                        }
                        if (o.RemoveActions)
                        {
                            try
                            {
                                if (shp.ActionSettings != null)
                                {
                                    bool any = false;
                                    bool changedAny = false;
                                    var activations = new[]
                                    {
                                        PowerPoint.PpMouseActivation.ppMouseClick,
                                        PowerPoint.PpMouseActivation.ppMouseOver
                                    };
                                    foreach (var actKey in activations)
                                    {
                                        try
                                        {
                                            var act = shp.ActionSettings[actKey];
                                            if (act.Action != PowerPoint.PpActionType.ppActionNone)
                                            {
                                                if (dryRun) any = true;
                                                else
                                                {
                                                    try { act.Action = PowerPoint.PpActionType.ppActionNone; changedAny = true; }
                                                    catch { st.Failed++; }
                                                }
                                            }
                                        }
                                        catch { st.Failed++; }
                                    }
                                    if (any || changedAny) st.Actions++;
                                }
                            }
                            catch { st.Failed++; }
                        }
                    }
                    catch { st.Failed++; }
                }
                bool slideWillBeDeleted = false;
                if (o.DelEmptySlides && remainingSlides > 1
                    && !HasVisibleContent(slide, o, slideW, slideH))
                {
                    if (dryRun)
                    {
                        st.EmptySlides++;
                        remainingSlides--;
                        slideWillBeDeleted = true;
                    }
                    else
                    {
                        try
                        {
                            slide.Delete();
                            st.EmptySlides++;
                            remainingSlides--;
                            slideWillBeDeleted = true;
                        }
                        catch { st.Failed++; }
                    }
                }
                if (o.ClearNotes && !slideWillBeDeleted)
                {
                    try
                    {
                        var notes = slide.NotesPage.Shapes.Placeholders[(int)PowerPoint.PpPlaceholderType.ppPlaceholderBody];
                        var tr = notes.TextFrame.TextRange;
                        if (tr != null && !string.IsNullOrWhiteSpace(tr.Text))
                        {
                            if (dryRun) st.Notes++;
                            else
                            {
                                try { tr.Text = ""; st.Notes++; }
                                catch { st.Failed++; }
                            }
                        }
                    }
                    catch { st.Failed++; }
                }
            }
            return st;
        }

        /// <summary>页是否有实质可见内容：占位符且无文字视为空；其它可见对象均算内容。</summary>
        private static bool HasVisibleContent(PowerPoint.Slide slide, CleanOptions options, float slideW, float slideH)
        {
            try
            {
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shp = slide.Shapes[i];
                    try
                    {
                        if (shp.Visible != MsoTriState.msoTrue)
                        {
                            // Preserve hidden content unless the user explicitly selected its removal.
                            if (!options.DelHiddenShapes) return true;
                            continue;
                        }
                        if (options.DelOffSlideShapes)
                        {
                            float l = (float)shp.Left, t = (float)shp.Top;
                            float r = l + (float)shp.Width, b = t + (float)shp.Height;
                            if (r < -5 || b < -5 || l > slideW + 5 || t > slideH + 5) continue;
                        }
                        bool isPlaceholder = false;
                        try { isPlaceholder = shp.PlaceholderFormat != null; } catch { }
                        if (isPlaceholder)
                        {
                            try
                            {
                                if (shp.HasTextFrame == MsoTriState.msoTrue && shp.TextFrame.HasText == MsoTriState.msoTrue
                                    && !string.IsNullOrWhiteSpace(shp.TextFrame.TextRange.Text)) return true;
                            }
                            catch { return true; }
                            continue;
                        }
                        return true;
                    }
                    catch { return true; }
                }
            }
            catch
            {
                // An unreadable slide is not safely known to be empty. Preserve it rather than
                // letting an automation/COM error turn into a destructive deletion.
                return true;
            }
            return false;
        }

        // ================= ③ 统一尺寸 =================

        /// <summary>scope=selection/page/all；宽高至少一项；图片且 forceImageRatio 时等比（取 Min 比例）。</summary>
        public static TextToolService.ToolResult ResizeShapes(string scope, float? width, float? height, bool forceImageRatio)
        {
            try
            {
                if (width == null && height == null)
                    return new TextToolService.ToolResult { Ok = false, Message = "请至少填写宽度或高度。" };
                var shapes = ShapesByScope(scope);
                int done = 0, failed = 0;
                foreach (var shp in shapes)
                {
                    try
                    {
                        float w = (float)shp.Width, h = (float)shp.Height;
                        if (w <= 0 || h <= 0) continue;
                        bool isPic = false;
                        try { isPic = shp.Type == MsoShapeType.msoPicture; } catch { }
                        if (width != null && height != null)
                        {
                            if (isPic && forceImageRatio)
                            {
                                float scale = Math.Min(width.Value / w, height.Value / h);
                                shp.Width = w * scale; shp.Height = h * scale;
                            }
                            else { shp.Width = width.Value; shp.Height = height.Value; }
                        }
                        else if (width != null)
                        {
                            float scale = width.Value / w;
                            shp.Width = width.Value; shp.Height = h * scale;
                        }
                        else
                        {
                            float scale = height.Value / h;
                            shp.Height = height.Value; shp.Width = w * scale;
                        }
                        done++;
                    }
                    catch { failed++; }
                }
                return new TextToolService.ToolResult
                {
                    Ok = failed == 0,
                    Message = "已统一 " + done + " 个形状尺寸" + (failed > 0 ? "，" + failed + " 个形状调整失败。" : "。"),
                    Changed = done
                };
            }
            catch (Exception ex)
            {
                return new TextToolService.ToolResult { Ok = false, Message = ex.Message };
            }
        }

        // ================= ④ 图片网格 =================

        /// <summary>多图按 cols 排成网格；toNewPage 时每页排满（cols×4）自动加新空白页；contain 等比居中不裁切。</summary>
        public static TextToolService.ToolResult ImageGrid(List<string> files, int cols, bool toNewPage)
        {
            try
            {
                if (files == null || files.Count == 0)
                    return new TextToolService.ToolResult { Ok = false, Message = "请先选择图片文件。" };
                var pres = App.ActivePresentation;
                if (pres == null)
                    return new TextToolService.ToolResult { Ok = false, Message = "没有打开的演示文稿。" };
                if (cols < 1) cols = 1;
                float sw = (float)pres.PageSetup.SlideWidth;
                float sh = (float)pres.PageSetup.SlideHeight;

                PowerPoint.Slide slide = null;
                if (!toNewPage)
                {
                    slide = pres.Slides[App.ActiveWindow.View.Slide.SlideIndex];
                }
                const float margin = 12f, gap = 8f;
                int perPage = cols * 4;
                int onPage = 0;
                int inserted = 0, failed = 0;
                foreach (string file in files)
                {
                    if (string.IsNullOrEmpty(file) || !File.Exists(file)) { failed++; continue; }
                    if (toNewPage && (slide == null || onPage >= perPage))
                    {
                        slide = pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                        onPage = 0;
                    }
                    int rowInPage = onPage / cols;
                    int colInPage = onPage % cols;
                    float cellW = (sw - margin * 2 - gap * (cols - 1)) / cols;
                    float cellH = (sh - margin * 2 - gap * (4 - 1)) / 4;
                    float x = margin + colInPage * (cellW + gap);
                    float y = margin + rowInPage * (cellH + gap);
                    try
                    {
                        var pic = slide.Shapes.AddPicture(file, MsoTriState.msoFalse, MsoTriState.msoTrue, x, y, cellW, cellH);
                        float pw = (float)pic.Width, ph = (float)pic.Height;
                        if (pw > 0 && ph > 0)
                        {
                            float scale = Math.Min(cellW / pw, cellH / ph);
                            float dw = pw * scale, dh = ph * scale;
                            pic.Width = dw; pic.Height = dh;
                            pic.Left = x + (cellW - dw) / 2f;
                            pic.Top = y + (cellH - dh) / 2f;
                        }
                        else
                        {
                            pic.Left = x; pic.Top = y; pic.Width = cellW; pic.Height = cellH;
                        }
                        inserted++; onPage++;
                    }
                    catch { failed++; }
                }
                return new TextToolService.ToolResult
                {
                    Ok = inserted > 0,
                    Message = "已插入 " + inserted + " 张图片" +
                              (toNewPage ? "（按需新增幻灯片）" : "到当前页") +
                              (failed > 0 ? "，" + failed + " 张失败。" : "。"),
                    Changed = inserted
                };
            }
            catch (Exception ex)
            {
                return new TextToolService.ToolResult { Ok = false, Message = ex.Message };
            }
        }

        // ================= 辅助 =================

        private static List<PowerPoint.Shape> ShapesByScope(string scope)
        {
            var result = new List<PowerPoint.Shape>();
            try
            {
                if (scope == "selection")
                {
                    var sel = App.ActiveWindow.Selection;
                    if (sel != null && sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange != null)
                        for (int i = 1; i <= sel.ShapeRange.Count; i++) result.Add(sel.ShapeRange[i]);
                }
                else if (scope == "page")
                {
                    var slide = App.ActivePresentation.Slides[App.ActiveWindow.View.Slide.SlideIndex];
                    for (int i = 1; i <= slide.Shapes.Count; i++) result.Add(slide.Shapes[i]);
                }
                else
                {
                    foreach (PowerPoint.Slide slide in App.ActivePresentation.Slides)
                        for (int i = 1; i <= slide.Shapes.Count; i++) result.Add(slide.Shapes[i]);
                }
            }
            catch { }
            return result;
        }
    }
}
