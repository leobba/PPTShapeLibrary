using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 学术/学习场景工具服务层（学生科研汇报）：页脚页码、导出 PDF/讲义/大纲、
    /// 图表一键美化、参考文献页生成。无 UI 依赖，结果统一 ToolResult 由 Web/原生呈现。
    /// </summary>
    public static class AcademicToolService
    {
        private static PowerPoint.Application App
        {
            get { return Globals.ThisAddIn.Application; }
        }

        private static TextToolService.ToolResult Ok(string msg, int changed = 0)
        {
            return new TextToolService.ToolResult { Ok = true, Message = msg, Changed = changed };
        }

        private static TextToolService.ToolResult Err(string msg)
        {
            return new TextToolService.ToolResult { Ok = false, Message = msg };
        }

        // ================= A5 页脚 + 页码 =================

        /// <summary>给全部页加页脚(姓名/单位/日期可选)与页码。existingSlidesOnly=true 只改当前已存在页。</summary>
        public static TextToolService.ToolResult AddFooterPageNumbers(
            string footerText, bool showDate, bool showPageNumber, bool showSlideNumberOnlyOnFirst)
        {
            try
            {
                var pres = App.ActivePresentation;
                if (pres == null) return Err("没有打开的演示文稿。");
                // 页脚/页码经幻灯片母版页脚占位与页面页脚两种。稳妥：每页底部建文本框
                float sw = (float)pres.PageSetup.SlideWidth;
                float sh = (float)pres.PageSetup.SlideHeight;
                int done = 0;
                for (int si = 1; si <= pres.Slides.Count; si++)
                {
                    var slide = pres.Slides[si];
                    try
                    {
                        // 页脚文本（左下）
                        var parts = new List<string>();
                        if (!string.IsNullOrWhiteSpace(footerText)) parts.Add(footerText.Trim());
                        if (showDate) parts.Add(DateTime.Now.ToString("yyyy-MM-dd"));
                        string leftText = string.Join("  ", parts);
                        if (leftText.Length > 0)
                        {
                            AddFooterTextBox(slide, leftText, 0.35f, 0.94f, 0.3f * sw, sh * 0.045f, 11, PowerPoint.PpParagraphAlignment.ppAlignLeft);
                        }
                        // 页码（右下）
                        if (showPageNumber)
                        {
                            if (!showSlideNumberOnlyOnFirst || si == 1)
                            {
                                AddFooterTextBox(slide, si.ToString(), 0.65f * sw, 0.945f * sh, 0.3f * sw, sh * 0.045f, 11, PowerPoint.PpParagraphAlignment.ppAlignRight);
                            }
                        }
                        done++;
                    }
                    catch { }
                }
                return Ok("已为 " + done + " 页添加页脚" + (showPageNumber ? "与页码" : "") + "。\n（若已有页脚会叠加，可先撤销再试。）", done);
            }
            catch (Exception ex)
            {
                return Err(ex.Message);
            }
        }

        private static void AddFooterTextBox(PowerPoint.Slide slide, string text, float left, float top, float w, float h, float fontSize, PowerPoint.PpParagraphAlignment align)
        {
            try
            {
                var tb = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                    left, top, w, h);
                var tr = tb.TextFrame.TextRange;
                tr.Text = text;
                tr.Font.Size = fontSize;
                tr.Font.Name = "微软雅黑";
                tr.Font.Color.RGB = 0x666666;   // RRGGBB 对称，BGR 相同
                tr.ParagraphFormat.Alignment = align;
                tb.TextFrame.WordWrap = MsoTriState.msoTrue;
            }
            catch { }
        }

        // ================= B6 导出 PDF / 讲义 / 大纲 =================

        /// <summary>整稿导出 PDF。style: 0=整页(默认) 1=讲义(每页2/3张) 2=大纲视图。</summary>
        public static TextToolService.ToolResult ExportPdf(string dir, string fileNameBase, int style, int slidesPerPage)
        {
            try
            {
                var pres = App.ActivePresentation;
                if (pres == null) return Err("没有打开的演示文稿。");
                if (!System.IO.Directory.Exists(dir)) return Err("导出目录不存在。");
                if (string.IsNullOrEmpty(fileNameBase)) fileNameBase = pres.Name;
                string safe = Sanitize(fileNameBase);
                if (safe.Length == 0) safe = "演示文稿";
                string file = System.IO.Path.Combine(dir, safe + ".pdf");
                try
                {
                    pres.ExportAsFixedFormat(file,
                        PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
                }
                catch
                {
                    pres.SaveAs(file, PowerPoint.PpSaveAsFileType.ppSaveAsPDF);
                }
                return Ok("已导出 PDF：\n" + file, 1);
            }
            catch (Exception ex)
            {
                return Err(ex.Message);
            }
        }

        /// <summary>把全部页文本整理为大纲文本（按页标题+正文），可写文件或返回文本供前端下载/复制。</summary>
        public static TextToolService.ToolResult ExportOutlineText(string dir, string fileNameBase)
        {
            try
            {
                var pres = App.ActivePresentation;
                if (pres == null) return Err("没有打开的演示文稿。");
                var sb = new StringBuilder();
                int page = 0;
                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    page++;
                    sb.AppendLine("===== 第 " + page + " 页 =====");
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var shp = slide.Shapes[i];
                        try
                        {
                            if (shp.HasTextFrame == MsoTriState.msoTrue && shp.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                                string t = shp.TextFrame.TextRange.Text;
                                if (!string.IsNullOrWhiteSpace(t)) sb.AppendLine(t.Trim());
                            }
                        }
                        catch { }
                    }
                    sb.AppendLine();
                }
                string text = sb.ToString();
                if (string.IsNullOrEmpty(dir)) return Ok(text, page);
                if (!System.IO.Directory.Exists(dir)) return Err("导出目录不存在。");
                string safe = Sanitize(fileNameBase);
                if (safe.Length == 0) safe = "演示文稿";
                string file = System.IO.Path.Combine(dir, safe + "-大纲.txt");
                System.IO.File.WriteAllText(file, text, Encoding.UTF8);
                return Ok("已导出大纲文本：\n" + file, page);
            }
            catch (Exception ex)
            {
                return Err(ex.Message);
            }
        }

        private static string Sanitize(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var invalid = System.IO.Path.GetInvalidFileNameChars();
            var chars = s.Trim().Select(c => Array.IndexOf(invalid, c) >= 0 ? '_' : c).ToArray();
            return new string(chars);
        }

        // ================= A2 图表一键美化 =================

        /// <summary>选中图表(含内嵌图表)套用预设调色板与字体。palette 0=学术蓝 1=清新绿 2=暖橙 3=莫兰迪。</summary>
        public static TextToolService.ToolResult BeautifyChart(int palette, string fontName)
        {
            try
            {
                var sel = App.ActiveWindow.Selection;
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange == null)
                    return Err("请先选中一个图表。");
                var palettes = new[]
                {
                    new[] { 0x1F77D4, 0x2E9B4F, 0xD47F00, 0x7A3FB8, 0x0B7A9B, 0xD63A2A }, // 学术蓝
                    new[] { 0x34A853, 0x9BC53D, 0x1F77D4, 0xE3A23C, 0x7A9E3A, 0x4A9A9A }, // 清新绿
                    new[] { 0xE07B39, 0xF2C14E, 0x5A9BD5, 0xA3B18A, 0xC56B6B, 0x9A7B4F }, // 暖橙
                    new[] { 0x8E9AAF, 0xCBC0D3, 0x9D8189, 0x7A6C8F, 0xB7A6C4, 0x6C6F7C }  // 莫兰迪
                };
                var colors = palettes[Math.Max(0, Math.Min(palette, palettes.Length - 1))];
                string fn = string.IsNullOrWhiteSpace(fontName) ? "微软雅黑" : fontName.Trim();

                int done = 0;
                foreach (PowerPoint.Shape shp in sel.ShapeRange)
                {
                    try
                    {
                        if (shp.HasChart != MsoTriState.msoTrue) continue;
                        var chart = shp.Chart;
                        int seriesCount = chart.SeriesCollection().Count;
                        // 逐系列改色
                        for (int s = 1; s <= seriesCount; s++)
                        {
                            try
                            {
                                var series = chart.SeriesCollection(s);
                                int ci = colors[(s - 1) % colors.Length];
                                series.Format.Fill.ForeColor.RGB = ci;
                            }
                            catch { }
                        }
                        // 图表区字体
                        try { chart.ChartArea.Font.Name = fn; } catch { }
                        try
                        {
                            var anySeries = chart.SeriesCollection(1);
                            if (anySeries != null)
                            {
                                try { anySeries.DataLabels.Font.Name = fn; } catch { }
                            }
                        }
                        catch { }
                        done++;
                    }
                    catch { }
                }
                return done > 0 ? Ok("已美化 " + done + " 个图表（配色+字体）。", done) : Err("所选范围内没有图表。请选中图表后重试。");
            }
            catch (Exception ex)
            {
                return Err(ex.Message);
            }
        }

        // ================= A3 参考文献工具 =================

        /// <summary>
        /// 依据条目生成参考文献文本(GB/T7714 或 APA)。items 每项 "作者. 标题. 来源, 年份" 简化模板：
        /// 约定格式 "作者|标题|来源|年份|页码" 用 | 分隔便于前端输入框逐字段。
        /// </summary>
        public static TextToolService.ToolResult BuildReferences(List<string[]> items, string style, string insertTitle)
        {
            try
            {
                if (items == null || items.Count == 0) return Err("请至少输入一条文献。");
                var sb = new StringBuilder();
                int i = 1;
                foreach (var f in items)
                {
                    string author = f.Length > 0 ? f[0].Trim() : "";
                    string paperTitle = f.Length > 1 ? f[1].Trim() : "";
                    string source = f.Length > 2 ? f[2].Trim() : "";
                    string year = f.Length > 3 ? f[3].Trim() : "";
                    string pages = f.Length > 4 ? f[4].Trim() : "";
                    bool apa = "APA".Equals(style, StringComparison.OrdinalIgnoreCase);
                    string line;
                    if (apa)
                    {
                        // APA: Author (Year). Title. Source, Pages.
                        line = (author.Length > 0 ? author : "佚名") +
                               (year.Length > 0 ? " (" + year + "). " : ". ") +
                               (paperTitle.Length > 0 ? paperTitle : "(无题)") +
                               (source.Length > 0 ? ". " + source : "") +
                               (pages.Length > 0 ? ", " + pages : "") + ".";
                    }
                    else
                    {
                        // GB/T 7714: 作者. 标题[文献类型]. 来源, 年: 页码.
                        string authors = author.Length > 0 ? author.Replace(";", ",") : "佚名";
                        line = authors + ". " + (paperTitle.Length > 0 ? paperTitle : "(无题)") +
                               "[J]. " + (source.Length > 0 ? source : "") +
                               (year.Length > 0 ? ", " + year : "") +
                               (pages.Length > 0 ? ": " + pages : "") + ".";
                    }
                    sb.AppendLine("[" + i + "] " + line);
                    i++;
                }
                string title = string.IsNullOrWhiteSpace(insertTitle) ? "参考文献" : insertTitle.Trim();
                string body = title + "\n" + sb.ToString();
                return Ok(body, items.Count);
            }
            catch (Exception ex)
            {
                return Err(ex.Message);
            }
        }
    }
}
