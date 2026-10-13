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

        public class CleanStats
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

        /// <summary>
        /// "整稿导出图片"的分片作业：把逐页导出拆成 Step(i)，由桥在 UI 线程上**一片一片**执行
        /// （片间回到消息循环）→ PowerPoint 不卡死、取消按钮也走普通桥消息即时生效。
        /// 同步调用方可以直接 while 循环跑完（行为与旧实现一致）。
        /// </summary>
        public class ExportImagesJob
        {
            public int Total { get; private set; }
            public int Done { get; private set; }
            public int Failed { get; private set; }

            private readonly string _dir;
            private readonly bool _isJpg;
            private readonly int _width;
            private PowerPoint.Presentation _pres;

            private ExportImagesJob(PowerPoint.Presentation pres, string dir, bool isJpg, int width)
            {
                _pres = pres;
                _dir = dir;
                _isJpg = isJpg;
                _width = width;
                Total = pres.Slides.Count;
            }

            /// <summary>校验参数并建立作业；不满足时返回 null 并给出原因。</summary>
            public static ExportImagesJob Create(string dir, string format, int width, out string error)
            {
                error = null;
                if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
                {
                    error = "导出目录不存在。";
                    return null;
                }
                var pres = ShapeExporter.ActiveUserPresentation();
                if (pres == null)
                {
                    error = "没有打开的演示文稿。";
                    return null;
                }
                bool isJpg = "JPG".Equals(format, StringComparison.OrdinalIgnoreCase);
                return new ExportImagesJob(pres, dir, isJpg, width > 0 ? width : 0);
            }

            /// <summary>导出第 i 页（0-based）。</summary>
            public void Step(int i)
            {
                if (_pres == null) return;
                int idx = i + 1;
                if (idx > _pres.Slides.Count) return;   // 处理期间用户删了页：跳过
                string file = Path.Combine(_dir, "slide" + idx.ToString("D2") + "." + (_isJpg ? "jpg" : "png"));
                try
                {
                    _pres.Slides[idx].Export(file, _isJpg ? "JPG" : "PNG", _width, 0);
                    Done++;
                }
                catch { Failed++; }
            }

            public TextToolService.ToolResult Finish(bool cancelled)
            {
                if (Failed > 0 && Done == 0)
                    return new TextToolService.ToolResult
                    {
                        Ok = false,
                        Message = "导出失败：" + Failed + " 页均无法导出（请检查目录权限/格式支持）。",
                        Changed = 0
                    };
                return new TextToolService.ToolResult
                {
                    Ok = Done > 0,
                    Message = (cancelled ? "已取消，" : "") + "已导出 " + Done + " / " + Total + " 页到：" + _dir +
                              (Failed > 0 ? "（" + Failed + " 页失败）" : ""),
                    Changed = Done
                };
            }
        }

        /// <summary>逐页导出图片（同步版：非桥调用方使用；桥走 ExportImagesJob 分片执行）。</summary>
        public static TextToolService.ToolResult ExportImages(string dir, string format, int width)
        {
            try
            {
                string err;
                var job = ExportImagesJob.Create(dir, format, width, out err);
                if (job == null) return new TextToolService.ToolResult { Ok = false, Message = err };
                for (int i = 0; i < job.Total; i++) job.Step(i);
                return job.Finish(false);
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
            string err;
            var job = CleanJob.Create(o, dryRun, out err);
            if (job == null) return st;
            for (int i = 0; i < job.Total; i++) job.Step(i);
            return job.Stats;
        }

        /// <summary>
        /// "清理瘦身"的分片作业：按**页**切片（从后往前，避免删页后索引错位）。
        /// 由桥在 UI 线程上一片一片执行 → 大文档清理不再卡死 PowerPoint、取消即时生效。
        /// 注意：分片执行期间用户可以继续操作文档，因此每片都重新确认目标页是否还存在。
        /// </summary>
        public class CleanJob
        {
            public CleanStats Stats { get; private set; }
            public int Total { get; private set; }

            private readonly PowerPoint.Presentation _pres;
            private readonly CleanOptions _o;
            private readonly bool _dryRun;
            private readonly float _slideW;
            private readonly float _slideH;
            private int _remainingSlides;

            private CleanJob(PowerPoint.Presentation pres, CleanOptions o, bool dryRun)
            {
                _pres = pres;
                _o = o;
                _dryRun = dryRun;
                Stats = new CleanStats();
                _slideW = (float)pres.PageSetup.SlideWidth;
                _slideH = (float)pres.PageSetup.SlideHeight;
                Total = pres.Slides.Count;
                _remainingSlides = Total;
            }

            public static CleanJob Create(CleanOptions o, bool dryRun, out string error)
            {
                error = null;
                var pres = ShapeExporter.ActiveUserPresentation();
                if (pres == null)
                {
                    error = "没有打开的演示文稿。";
                    return null;
                }
                return new CleanJob(pres, o ?? new CleanOptions(), dryRun);
            }

            /// <summary>处理第 i 片（0-based，对应从后往前数的第 i 页）。</summary>
            public void Step(int i)
            {
                int slideIndex = Total - i;                 // 从最后一页开始
                if (slideIndex < 1 || slideIndex > _pres.Slides.Count) return;   // 期间被删了：跳过
                PowerPoint.Slide slide;
                try { slide = _pres.Slides[slideIndex]; }
                catch { return; }

                for (int k = slide.Shapes.Count; k >= 1; k--)
                {
                    PowerPoint.Shape shp;
                    try { shp = slide.Shapes[k]; } catch { Stats.Failed++; continue; }
                    try
                    {
                        if (_o.DelHiddenShapes && shp.Visible == MsoTriState.msoFalse)
                        {
                            if (_dryRun) Stats.HiddenShapes++;
                            else { try { shp.Delete(); Stats.HiddenShapes++; } catch { Stats.Failed++; } }
                            continue;
                        }
                        if (_o.DelOffSlideShapes)
                        {
                            float l, t, r, b;
                            try { l = (float)shp.Left; t = (float)shp.Top; r = l + (float)shp.Width; b = t + (float)shp.Height; }
                            catch { Stats.Failed++; continue; }
                            if (r < -5 || b < -5 || l > _slideW + 5 || t > _slideH + 5)
                            {
                                if (_dryRun) Stats.OffSlideShapes++;
                                else { try { shp.Delete(); Stats.OffSlideShapes++; } catch { Stats.Failed++; } }
                                continue;
                            }
                        }
                        if (_o.RemoveActions)
                        {
                            try
                            {
                                if (shp.ActionSettings != null)
                                {
                                    bool any = false, changedAny = false;
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
                                                if (_dryRun) any = true;
                                                else { try { act.Action = PowerPoint.PpActionType.ppActionNone; changedAny = true; } catch { Stats.Failed++; } }
                                            }
                                        }
                                        catch { Stats.Failed++; }
                                    }
                                    if (any || changedAny) Stats.Actions++;
                                }
                            }
                            catch { Stats.Failed++; }
                        }
                    }
                    catch { Stats.Failed++; }
                }

                bool slideWillBeDeleted = false;
                if (_o.DelEmptySlides && _remainingSlides > 1 && !HasVisibleContent(slide, _o, _slideW, _slideH))
                {
                    if (_dryRun)
                    {
                        Stats.EmptySlides++;
                        _remainingSlides--;
                        slideWillBeDeleted = true;
                    }
                    else
                    {
                        try
                        {
                            slide.Delete();
                            Stats.EmptySlides++;
                            _remainingSlides--;
                            slideWillBeDeleted = true;
                        }
                        catch { Stats.Failed++; }
                    }
                }
                if (_o.ClearNotes && !slideWillBeDeleted)
                {
                    try
                    {
                        var notes = slide.NotesPage.Shapes.Placeholders[(int)PowerPoint.PpPlaceholderType.ppPlaceholderBody];
                        var tr = notes.TextFrame.TextRange;
                        if (tr != null && !string.IsNullOrWhiteSpace(tr.Text))
                        {
                            if (_dryRun) Stats.Notes++;
                            else { try { tr.Text = ""; Stats.Notes++; } catch { Stats.Failed++; } }
                        }
                    }
                    catch { Stats.Failed++; }
                }
            }

            public TextToolService.ToolResult Finish(bool cancelled)
            {
                bool scan = _dryRun;
                return new TextToolService.ToolResult
                {
                    Ok = Stats.Failed == 0,
                    Message = (cancelled ? "已取消（部分完成）：" : (scan ? "" : (Stats.Failed == 0 ? "清理完成：" : "清理部分完成："))) +
                              StatsText(Stats),
                    Changed = Stats.EmptySlides + Stats.HiddenShapes + Stats.OffSlideShapes + Stats.Notes + Stats.Actions
                };
            }
        }

        /// <summary>页是否有实质可见内容：占位符且无文字视为空；其它可见对象均算内容。
        /// 读不到内容时**保守判定为"有内容"**——宁可少删一页，也不要把用户的东西删掉。</summary>
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
        /// <summary>
        /// "统一尺寸"的分片作业：先把目标形状列表抓下来，再逐个处理。
        /// 由桥分片执行 → 大文档批量改尺寸不再卡死 PowerPoint。
        /// 分片期间形状可能被用户删掉：每片都 try/catch，失败只计数。
        /// </summary>
        public class ResizeShapesJob
        {
            public int Total { get; private set; }
            public int Done { get; private set; }
            public int Failed { get; private set; }

            private readonly List<PowerPoint.Shape> _shapes;
            private readonly float? _width;
            private readonly float? _height;
            private readonly bool _forceImageRatio;

            private ResizeShapesJob(List<PowerPoint.Shape> shapes, float? width, float? height, bool forceImageRatio)
            {
                _shapes = shapes;
                _width = width;
                _height = height;
                _forceImageRatio = forceImageRatio;
                Total = shapes.Count;
            }

            public static ResizeShapesJob Create(string scope, float? width, float? height, bool forceImageRatio,
                                                 out string error)
            {
                error = null;
                if (width == null && height == null)
                {
                    error = "请至少填写宽度或高度。";
                    return null;
                }
                try
                {
                    var shapes = ShapesByScope(scope);
                    if (shapes.Count == 0)
                    {
                        error = "所选范围内没有形状。";
                        return null;
                    }
                    return new ResizeShapesJob(shapes, width, height, forceImageRatio);
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return null;
                }
            }

            public void Step(int i)
            {
                if (i < 0 || i >= _shapes.Count) return;
                var shp = _shapes[i];
                try
                {
                    float w = (float)shp.Width, h = (float)shp.Height;
                    if (w <= 0 || h <= 0) return;
                    bool isPic = false;
                    try { isPic = shp.Type == MsoShapeType.msoPicture; } catch { }
                    if (_width != null && _height != null)
                    {
                        if (isPic && _forceImageRatio)
                        {
                            float scale = Math.Min(_width.Value / w, _height.Value / h);
                            shp.Width = w * scale; shp.Height = h * scale;
                        }
                        else { shp.Width = _width.Value; shp.Height = _height.Value; }
                    }
                    else if (_width != null)
                    {
                        float scale = _width.Value / w;
                        shp.Width = _width.Value; shp.Height = h * scale;
                    }
                    else
                    {
                        float scale = _height.Value / h;
                        shp.Height = _height.Value; shp.Width = w * scale;
                    }
                    Done++;
                }
                catch { Failed++; }
            }

            public TextToolService.ToolResult Finish(bool cancelled)
            {
                return new TextToolService.ToolResult
                {
                    Ok = Failed == 0,
                    Message = (cancelled ? "已取消，" : "") + "已统一 " + Done + " / " + Total + " 个形状尺寸" +
                              (Failed > 0 ? "，" + Failed + " 个形状调整失败。" : "。"),
                    Changed = Done
                };
            }
        }

        /// <summary>scope=selection/page/all；宽高至少一项；图片且 forceImageRatio 时等比（取 Min 比例）。
        /// 同步版：桥走 ResizeShapesJob 分片执行。</summary>
        public static TextToolService.ToolResult ResizeShapes(string scope, float? width, float? height, bool forceImageRatio)
        {
            try
            {
                string err;
                var job = ResizeShapesJob.Create(scope, width, height, forceImageRatio, out err);
                if (job == null) return new TextToolService.ToolResult { Ok = false, Message = err };
                for (int i = 0; i < job.Total; i++) job.Step(i);
                return job.Finish(false);
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
                var pres = ShapeExporter.ActiveUserPresentation();
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
            // 与 TextToolService.ShapesByScope 保持一致：范围非法/无文稿/无选中都抛明确异常，
            // 由上层转成 Ok=false 的提示。旧实现 catch 后返回空列表 → "已统一 0 个形状"的假成功，
            // 且 scope 写错时会静默落到"整稿"，把不该动的形状全改掉。
            if (string.IsNullOrEmpty(scope)) scope = "all";
            if (scope != "selection" && scope != "page" && scope != "all")
                throw new InvalidOperationException("范围参数无效：" + scope);

            var result = new List<PowerPoint.Shape>();
            try
            {
                var pres = ShapeExporter.ActiveUserPresentation();
                if (pres == null) throw new InvalidOperationException("没有打开的演示文稿。");
                if (scope == "selection")
                {
                    var win = App.ActiveWindow;
                    if (win == null) throw new InvalidOperationException("没有可用的活动窗口。");
                    var sel = win.Selection;
                    if (sel != null && sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange != null)
                        for (int i = 1; i <= sel.ShapeRange.Count; i++) result.Add(sel.ShapeRange[i]);
                    if (result.Count == 0) throw new InvalidOperationException("未选中形状。请先选中要处理的形状。");
                }
                else if (scope == "page")
                {
                    var win = App.ActiveWindow;
                    if (win == null || win.View == null || win.View.Slide == null)
                        throw new InvalidOperationException("没有可用的活动窗口/页面。");
                    var slide = pres.Slides[win.View.Slide.SlideIndex];
                    for (int i = 1; i <= slide.Shapes.Count; i++) result.Add(slide.Shapes[i]);
                }
                else
                {
                    foreach (PowerPoint.Slide slide in pres.Slides)
                        for (int i = 1; i <= slide.Shapes.Count; i++) result.Add(slide.Shapes[i]);
                }
            }
            catch (InvalidOperationException) { throw; }
            catch { throw new InvalidOperationException("没有打开的演示文稿或活动窗口。"); }
            return result;
        }
    }
}
