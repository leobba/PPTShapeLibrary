using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTShapeLibrary
{
    /// <summary>
    /// PPT 文字/文本工具的共享静态服务层（Web 桥与 WinForms 窗体共用）。
    /// 每个方法返回 {ok, message, changed} 类结果而非弹窗，由调用方决定呈现方式。
    /// </summary>
    public static class TextToolService
    {
        private static PowerPoint.Application App
        {
            get { return Globals.ThisAddIn.Application; }
        }

        public class ToolResult
        {
            public bool Ok;
            public string Message;
            public int Changed;
        }

        /// <summary>按范围枚举文本形状：selection(选中)/page(当前页)/all(全部页)。
        /// 无演示文稿/无活动窗口时抛 InvalidOperationException（上层转为明确错误），避免静默"已处理 0 个"假成功。</summary>
        private static List<PowerPoint.Shape> ShapesByScope(string scope)
        {
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

        private static bool IsTextShape(PowerPoint.Shape shp)
        {
            try { return shp.HasTextFrame == MsoTriState.msoTrue && shp.TextFrame.HasText == MsoTriState.msoTrue; }
            catch { return false; }
        }

        /// <summary>批量修改文字样式：scope=selection/page/all；仅非空字段生效。</summary>
        public static ToolResult ApplyTextStyle(string scope, string fontName, float? size, bool? bold, bool? italic, int? colorRgb, string align)
        {
            int changed = 0, failed = 0;
            try
            {
                foreach (var shp in ShapesByScope(scope))
                {
                    if (!IsTextShape(shp)) continue;
                    try
                    {
                        var range = shp.TextFrame.TextRange;
                        if (!string.IsNullOrEmpty(fontName)) range.Font.Name = fontName;
                        if (size != null) range.Font.Size = size.Value;
                        if (bold != null) range.Font.Bold = bold.Value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                        if (italic != null) range.Font.Italic = italic.Value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                        if (colorRgb != null) range.Font.Color.RGB = colorRgb.Value;
                        if (align == "left") range.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                        else if (align == "center") range.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                        else if (align == "right") range.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignRight;
                        changed++;
                    }
                    catch { failed++; }
                }
                return new ToolResult
                {
                    Ok = failed == 0,
                    Message = "已处理 " + changed + " 个文字形状" + (failed > 0 ? "，" + failed + " 个处理失败。" : "。"),
                    Changed = changed
                };
            }
            catch (Exception ex)
            {
                return new ToolResult { Ok = false, Message = ex.Message };
            }
        }

        /// <summary>
        /// "一键换字体"的分片作业：按**页**切片，由桥在 UI 线程上一片一片执行
        /// （片间回到消息循环）→ 整稿换字体不再卡死 PowerPoint，取消也即时生效。
        /// </summary>
        public class SetFontsJob
        {
            public int Total { get; private set; }
            public int Shapes;
            public int Chars;
            public int Failed;

            private readonly PowerPoint.Presentation _pres;
            private readonly string _cn;
            private readonly string _en;
            private readonly bool _allSlides;
            private readonly bool _currentPageOnly;

            private SetFontsJob(PowerPoint.Presentation pres, string cn, string en, bool allSlides, bool currentPageOnly)
            {
                _pres = pres;
                _cn = cn;
                _en = en;
                _allSlides = allSlides;
                _currentPageOnly = currentPageOnly;
                Total = currentPageOnly ? 1 : pres.Slides.Count;
            }

            public static SetFontsJob Create(string chineseFont, string latinFont, bool allSlides, out string error)
            {
                error = null;
                string cn = (chineseFont ?? "").Trim();
                string en = (latinFont ?? "").Trim();
                if (cn.Length == 0 && en.Length == 0)
                {
                    error = "请至少选择一种新字体（中文字体或英文字体）。";
                    return null;
                }
                var pres = ShapeExporter.ActiveUserPresentation();
                if (pres == null)
                {
                    error = "没有打开的演示文稿。";
                    return null;
                }
                bool currentOnly = !allSlides;
                if (currentOnly)
                {
                    var win = App.ActiveWindow;
                    if (win == null || win.View == null || win.View.Slide == null)
                    {
                        error = "没有可用的活动窗口/页面。";
                        return null;
                    }
                }
                return new SetFontsJob(pres, cn, en, allSlides, currentOnly);
            }

            /// <summary>处理第 i 片（整稿=第 i 页；当前页=只有第 0 片）。</summary>
            public void Step(int i)
            {
                try
                {
                    PowerPoint.Slide slide;
                    if (_currentPageOnly)
                    {
                        int idx = App.ActiveWindow.View.Slide.SlideIndex;
                        if (idx < 1 || idx > _pres.Slides.Count) return;
                        slide = _pres.Slides[idx];
                    }
                    else
                    {
                        if (i + 1 > _pres.Slides.Count) return;   // 处理期间用户删了页
                        slide = _pres.Slides[i + 1];
                    }
                    for (int s = 1; s <= slide.Shapes.Count; s++)
                    {
                        int sh = 0, ch = 0, fl = 0;
                        ApplyFonts(slide.Shapes[s], _cn, _en, 0, ref sh, ref ch, ref fl);
                        Shapes += sh; Chars += ch; Failed += fl;
                    }
                }
                catch { Failed++; }
            }

            public ToolResult Finish(bool cancelled)
            {
                string what = _cn.Length > 0 && _en.Length > 0
                    ? "中文字体「" + _cn + "」＋英文字体「" + _en + "」"
                    : (_cn.Length > 0 ? "中文字体「" + _cn + "」" : "英文字体「" + _en + "」");
                return new ToolResult
                {
                    Ok = Failed == 0,
                    Message = (cancelled ? "已取消，" : "") + "已把 " + Shapes + " 个文字形状（约 " + Chars + " 字）改为 " + what +
                              (_allSlides ? "，范围：整份演示文稿" : "，范围：当前页") +
                              (Failed > 0 ? "；" + Failed + " 个形状处理失败。" : "。"),
                    Changed = Shapes
                };
            }
        }

        /// <summary>
        /// 一键换字体（同步版）：不需要指定"原字体"，直接把文字统一改成新的中文字体 / 英文字体。
        /// 通过显式写入 Font.NameFarEast（中文）与 Font.NameAscii / NameOther（西文）覆盖
        /// 原有的逐字符字体设置与主题字体引用（即"取消原字体"）。
        /// 覆盖普通文本框、组合形状内部、以及表格单元格内文字。
        /// 桥调用走 SetFontsJob 分片执行，本方法供非桥调用方（如 Ribbon）使用。
        /// </summary>
        /// <param name="chineseFont">新的中文字体；为空表示中文保持不变</param>
        /// <param name="latinFont">新的英文字体；为空表示西文保持不变</param>
        public static ToolResult SetFonts(string chineseFont, string latinFont, bool allSlides)
        {
            try
            {
                string err;
                var job = SetFontsJob.Create(chineseFont, latinFont, allSlides, out err);
                if (job == null) return new ToolResult { Ok = false, Message = err };
                for (int i = 0; i < job.Total; i++) job.Step(i);
                return job.Finish(false);
            }
            catch (Exception ex)
            {
                return new ToolResult { Ok = false, Message = ex.Message };
            }
        }

        /// <summary>递归处理形状：组合形状展开、表格逐格、普通文本框整体设置字体。</summary>
        private static void ApplyFonts(PowerPoint.Shape shp, string cn, string en, int depth,
                                       ref int shapes, ref int chars, ref int failed)
        {
            if (shp == null || depth > 6) return;
            try
            {
                if (shp.Type == MsoShapeType.msoGroup && shp.GroupItems != null)
                {
                    for (int i = 1; i <= shp.GroupItems.Count; i++)
                        ApplyFonts(shp.GroupItems[i], cn, en, depth + 1, ref shapes, ref chars, ref failed);
                    return;
                }
                if (shp.HasTable == MsoTriState.msoTrue)
                {
                    var tbl = shp.Table;
                    if (tbl == null) return;
                    for (int r = 1; r <= tbl.Rows.Count; r++)
                    {
                        for (int c = 1; c <= tbl.Columns.Count; c++)
                        {
                            try
                            {
                                var cellRange = tbl.Cell(r, c).Shape.TextFrame.TextRange;
                                if (SetFontOnRange(cellRange, cn, en))
                                {
                                    shapes++;
                                    chars += cellRange.Length;
                                }
                            }
                            catch { }   // 合并单元格/空单元格等：跳过
                        }
                    }
                    return;
                }
                if (IsTextShape(shp))
                {
                    var range = shp.TextFrame.TextRange;
                    if (SetFontOnRange(range, cn, en))
                    {
                        shapes++;
                        chars += range.Length;
                    }
                }
            }
            catch { failed++; }
        }

        /// <summary>把一段文字的中文/西文字体显式改写为目标字体；返回是否有文字被处理。</summary>
        private static bool SetFontOnRange(PowerPoint.TextRange range, string cn, string en)
        {
            if (range == null || range.Length <= 0) return false;
            var font = range.Font;
            if (cn.Length > 0) font.NameFarEast = cn;   // 中文字体（东亚）
            if (en.Length > 0)
            {
                font.NameAscii = en;   // 英文/数字/半角标点
                font.NameOther = en;   // 其他脚本，避免残留原字体
            }
            return true;
        }

        /// <summary>跨页查找替换（保留格式）。返回替换总数与跳过形状数。</summary>
        public static ToolResult FindReplace(string find, string replace, bool matchCase, bool allSlides)
        {
            int total = 0, shapesTouched = 0, skipped = 0;
            try
            {
                if (string.IsNullOrEmpty(find))
                    return new ToolResult { Ok = false, Message = "查找内容不能为空。" };
                var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                string scope = allSlides ? "all" : "page";
                foreach (var shp in ShapesByScope(scope))
                {
                    if (!IsTextShape(shp)) continue;
                    try
                    {
                        var textRange = shp.TextFrame.TextRange;
                        string full = textRange.Text ?? "";
                        int fromIdx = 0, inShape = 0, safety = 0;
                        while (safety++ < 2000)
                        {
                            int idx = full.IndexOf(find, fromIdx, cmp);
                            if (idx < 0) break;
                            textRange.Characters(idx + 1, find.Length).Text = replace;
                            total++; inShape++;
                            full = textRange.Text ?? "";
                            fromIdx = idx + find.Length;
                        }
                        if (inShape > 0) shapesTouched++;
                    }
                    catch { skipped++; }
                }
                return new ToolResult
                {
                    Ok = true,
                    Message = "共替换 " + total + " 处，涉及 " + shapesTouched + " 个文本形状。" +
                              (skipped > 0 ? "（" + skipped + " 个形状因异常跳过）" : ""),
                    Changed = total
                };
            }
            catch (Exception ex)
            {
                return new ToolResult { Ok = false, Message = ex.Message };
            }
        }

        /// <summary>文字清理（整稿）：去多余空白 / CJK 间西文标点转全角。整段重设并回刷首字符格式。</summary>
        public static ToolResult CleanText(bool trimWhitespace, bool unifyPunctuation)
        {
            int changed = 0, failed = 0;
            try
            {
                if (!trimWhitespace && !unifyPunctuation)
                    return new ToolResult { Ok = false, Message = "请至少选择一项清理规则。" };
                foreach (var shp in ShapesByScope("all"))
                {
                    if (!IsTextShape(shp)) continue;
                    try
                    {
                        var textRange = shp.TextFrame.TextRange;
                        string t = textRange.Text ?? "";
                        string o = t;
                        if (trimWhitespace)
                        {
                            t = t.Replace('\u3000', ' ');
                            t = System.Text.RegularExpressions.Regex.Replace(t, "[ ]{2,}", " ");
                            t = System.Text.RegularExpressions.Regex.Replace(t, "(?m)^[ \t]+|[ \t]+$", "");
                        }
                        if (unifyPunctuation)
                        {
                            t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*,\\s*([\u4e00-\u9fff])", "$1，$2");
                            t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*\\.\\s*([\u4e00-\u9fff])", "$1。$2");
                            t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*\\?\\s*([\u4e00-\u9fff])", "$1？$2");
                            t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*!\\s*([\u4e00-\u9fff])", "$1！$2");
                        }
                        if (t != o)
                        {
                            try
                            {
                                var f = textRange.Characters(1, 1).Font;
                                string fontName = f.Name; float size = f.Size; int rgb = f.Color.RGB;
                                bool bold = f.Bold == MsoTriState.msoTrue;
                                textRange.Text = t;
                                if (textRange.Text.Length > 0)
                                {
                                    var nf = textRange.Font;
                                    if (!string.IsNullOrEmpty(fontName)) nf.Name = fontName;
                                    if (size > 0) nf.Size = size;
                                    nf.Color.RGB = rgb;
                                    nf.Bold = bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                                }
                            }
                            catch { textRange.Text = t; }
                            changed++;
                        }
                    }
                    catch { failed++; }
                }
                return new ToolResult
                {
                    Ok = failed == 0,
                    Message = "已清理 " + changed + " 个文本形状" + (failed > 0 ? "，" + failed + " 个处理失败。" : "。"),
                    Changed = changed
                };
            }
            catch (Exception ex)
            {
                return new ToolResult { Ok = false, Message = ex.Message };
            }
        }

        /// <summary>当前演示文稿用到的全部字体（供前端字体下拉的"文稿字体"分组）。
        /// 同时收集中文/西文字体名，并过滤主题字体引用（+mn-* 等）。</summary>
        public static List<string> GetDocumentFonts()
        {
            var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var docPres = ShapeExporter.ActiveUserPresentation();
                if (docPres == null) return fonts.OrderBy(x => x).ToList();
                foreach (PowerPoint.Slide slide in docPres.Slides)
                {
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        try
                        {
                            var shp = slide.Shapes[i];
                            if (!IsTextShape(shp)) continue;
                            var font = shp.TextFrame.TextRange.Font;
                            // 逐个取值各自兜异常：混排文本上某个属性可能取不到，但不该影响其它属性
                            TryAddFontName(fonts, () => font.Name);
                            TryAddFontName(fonts, () => font.NameFarEast);
                            TryAddFontName(fonts, () => font.NameAscii);
                            TryAddFontName(fonts, () => font.NameOther);
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return fonts.OrderBy(x => x).ToList();
        }

        private static void TryAddFontName(HashSet<string> set, Func<string> get)
        {
            try { AddFontName(set, get()); } catch { }
        }

        /// <summary>收集字体名：跳过空值与主题字体引用（+mn-ea/+mj-lt 等）。</summary>
        private static void AddFontName(HashSet<string> set, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            if (name[0] == '+') return;          // 主题字体引用，不是真实字体名
            set.Add(name.Trim());
        }

        /// <summary>当前是否打开演示文稿（无打开时前端应提示）。</summary>
        public static bool HasPresentation()
        {
            try { return ShapeExporter.ActiveUserPresentation() != null; }
            catch { return false; }
        }
    }
}
