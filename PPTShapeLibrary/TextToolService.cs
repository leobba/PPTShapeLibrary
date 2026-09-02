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

        /// <summary>按范围枚举文本形状：selection(选中)/page(当前页)/all(全部页)。</summary>
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

        private static bool IsTextShape(PowerPoint.Shape shp)
        {
            try { return shp.HasTextFrame == MsoTriState.msoTrue && shp.TextFrame.HasText == MsoTriState.msoTrue; }
            catch { return false; }
        }

        /// <summary>批量修改文字样式：scope=selection/page/all；仅非空字段生效。</summary>
        public static ToolResult ApplyTextStyle(string scope, string fontName, float? size, bool? bold, bool? italic, int? colorRgb, string align)
        {
            int changed = 0;
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
                    catch { }
                }
                return new ToolResult { Ok = true, Message = "已处理 " + changed + " 个文字形状。", Changed = changed };
            }
            catch (Exception ex)
            {
                return new ToolResult { Ok = false, Message = ex.Message };
            }
        }

        /// <summary>一键替换字体：遍历整稿（scope 参数保留扩展），将 fromFont 文本替换为 toFont。</summary>
        public static ToolResult ReplaceFont(string fromFont, string toFont, bool allSlides)
        {
            int changed = 0;
            try
            {
                string scope = allSlides ? "all" : "page";
                foreach (var shp in ShapesByScope(scope))
                {
                    if (!IsTextShape(shp)) continue;
                    try
                    {
                        var range = shp.TextFrame.TextRange;
                        if (range.Font.Name.Equals(fromFont, StringComparison.OrdinalIgnoreCase))
                        {
                            range.Font.Name = toFont;
                            changed++;
                        }
                    }
                    catch { }
                }
                return new ToolResult { Ok = true, Message = "已将 " + changed + " 处文字从「" + fromFont + "」替换为「" + toFont + "」。", Changed = changed };
            }
            catch (Exception ex)
            {
                return new ToolResult { Ok = false, Message = ex.Message };
            }
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
            int changed = 0;
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
                    catch { }
                }
                return new ToolResult { Ok = true, Message = "已清理 " + changed + " 个文本形状。", Changed = changed };
            }
            catch (Exception ex)
            {
                return new ToolResult { Ok = false, Message = ex.Message };
            }
        }

        /// <summary>当前演示文稿用到的全部字体（供前端"原字体"下拉）。</summary>
        public static List<string> GetDocumentFonts()
        {
            var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (PowerPoint.Slide slide in App.ActivePresentation.Slides)
                {
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        try
                        {
                            var shp = slide.Shapes[i];
                            if (IsTextShape(shp))
                            {
                                string fn = shp.TextFrame.TextRange.Font.Name;
                                if (!string.IsNullOrEmpty(fn)) fonts.Add(fn);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return fonts.OrderBy(x => x).ToList();
        }

        /// <summary>当前是否打开演示文稿（无打开时前端应提示）。</summary>
        public static bool HasPresentation()
        {
            try { return App.ActivePresentation != null; }
            catch { return false; }
        }
    }
}
