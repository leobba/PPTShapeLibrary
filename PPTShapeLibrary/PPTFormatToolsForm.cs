using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTShapeLibrary
{
    /// <summary>
    /// PPT 版式/批处理工具：整稿导出 PNG/JPG、批量清理瘦身、统一尺寸、批量插图片成网格。
    /// </summary>
    public class PPTFormatToolsForm : Form
    {
        public PPTFormatToolsForm()
        {
            Text = "PPT 版式 / 批处理工具";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(470, 400);
            var font = new Font("Microsoft YaHei", 9F);

            var tabs = new TabControl { Dock = DockStyle.Fill, Font = font };
            tabs.TabPages.Add(BuildExportTab(font));
            tabs.TabPages.Add(BuildCleanTab(font));
            tabs.TabPages.Add(BuildSizeTab(font));
            tabs.TabPages.Add(BuildGridTab(font));
            Controls.Add(tabs);

            var btnClose = new Button { Text = "关闭", DialogResult = DialogResult.Cancel, Anchor = AnchorStyles.Bottom | AnchorStyles.Right, Location = new Point(358, 366), Size = new Size(100, 28), Font = font };
            CancelButton = btnClose;
            Controls.Add(btnClose);
        }

        private static PowerPoint.Application App
        {
            get { return Globals.ThisAddIn.Application; }
        }

        private static void Warn(string msg)
        {
            MessageBox.Show(msg, "PPT 工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void Fail(string msg)
        {
            MessageBox.Show("操作失败：" + msg, "PPT 工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        // ================= ① 整稿导出 PNG/JPG =================
        private TabPage BuildExportTab(Font font)
        {
            var page = new TabPage("整稿导出图片") { Padding = new Padding(8) };
            var grp = new GroupBox { Text = "把整份演示文稿每页导出为图片", Dock = DockStyle.Fill, Font = font };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(10) };

            var lblFmt = new Label { Text = "图片格式：", AutoSize = true, Font = font };
            var cmbFmt = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = font };
            cmbFmt.Items.AddRange(new[] { "PNG", "JPG" });
            cmbFmt.SelectedIndex = 0;
            var lblFolder = new Label { Text = "导出目录：", AutoSize = true, Font = font };
            var txtFolder = new TextBox { Font = font, Text = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) };
            var btnPick = new Button { Text = "浏览…", Font = font };
            btnPick.Click += (s, e) =>
            {
                using (var dlg = new FolderBrowserDialog { Description = "选择图片导出目录" })
                {
                    if (dlg.ShowDialog() == DialogResult.OK) txtFolder.Text = dlg.SelectedPath;
                }
            };
            var lblScale = new Label { Text = "分辨率：", AutoSize = true, Font = font };
            var cmbScale = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = font };
            cmbScale.Items.AddRange(new[] { "原始大小", "宽 1280", "宽 1920" });
            cmbScale.SelectedIndex = 0;
            var btnExport = new Button
            {
                Text = "开始导出", Font = font, Height = 30,
                BackColor = Color.FromArgb(0x1F, 0x6F, 0xA2), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };

            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            AddRow(layout, 0, lblFmt, cmbFmt);
            AddRow(layout, 1, lblFolder, MakeStack(txtFolder, btnPick, 300));
            AddRow(layout, 2, lblScale, cmbScale);
            layout.Controls.Add(btnExport, 0, 3);
            layout.SetColumnSpan(btnExport, 2);

            btnExport.Click += (s, e) =>
            {
                try
                {
                    string dir = txtFolder.Text.Trim();
                    if (!System.IO.Directory.Exists(dir)) { Fail("导出目录不存在。"); return; }
                    var pres = App.ActivePresentation;
                    if (pres == null) { Fail("没有打开的演示文稿。"); return; }
                    string ext = cmbFmt.SelectedItem as string ?? "PNG";
                    bool isJpg = ext.Equals("JPG", StringComparison.OrdinalIgnoreCase);
                    int w = 0;
                    string scale = cmbScale.SelectedItem as string;
                    if (scale != null && scale.Contains("1280")) w = 1280;
                    if (scale != null && scale.Contains("1920")) w = 1920;

                    int done = 0;
                    Cursor = Cursors.WaitCursor;
                    try
                    {
                        for (int slideIdx = 1; slideIdx <= pres.Slides.Count; slideIdx++)
                        {
                            string file = System.IO.Path.Combine(dir, "slide" + slideIdx.ToString("D2") + "." + (isJpg ? "jpg" : "png"));
                            pres.Slides[slideIdx].Export(file, isJpg ? "JPG" : "PNG", w, 0);
                            done++;
                        }
                    }
                    finally { Cursor = Cursors.Default; }
                    Warn("已导出 " + done + " 页到：\n" + dir);
                }
                catch (Exception ex) { Fail(ex.Message); }
            };

            grp.Controls.Add(layout);
            page.Controls.Add(grp);
            return page;
        }

        // ================= ② 批量清理瘦身 =================
        private CheckBox chkDelEmpty, chkDelHidden, chkDelOffSlide, chkClearNotes, chkDelActions;
        private Label lblCleanPreview;

        private TabPage BuildCleanTab(Font font)
        {
            var page = new TabPage("清理瘦身") { Padding = new Padding(8) };
            var grp = new GroupBox { Text = "批量清理（先扫描预览，再执行）", Dock = DockStyle.Fill, Font = font };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(10) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 230));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            chkDelEmpty = new CheckBox { Text = "删除空白幻灯片", Font = font, Checked = true };
            chkDelHidden = new CheckBox { Text = "删除隐藏对象（含隐藏的标题占位符）", Font = font, Checked = false };
            chkDelOffSlide = new CheckBox { Text = "删除页面外对象（超出画布）", Font = font, Checked = false };
            chkClearNotes = new CheckBox { Text = "清空备注页文字", Font = font, Checked = true };
            chkDelActions = new CheckBox { Text = "移除形状超链接与动作（保留外观）", Font = font, Checked = false };

            var btnScan = new Button { Text = "扫描统计", Font = font, Height = 30 };
            var btnRun = new Button { Text = "执行清理", Font = font, Height = 30, BackColor = Color.FromArgb(0x7B, 0x1F, 0xA2), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            lblCleanPreview = new Label { Text = "（尚未扫描）", Font = font, AutoSize = true, ForeColor = Color.FromArgb(0x66, 0x70, 0x7E) };

            layout.Controls.Add(chkDelEmpty, 0, 0);
            layout.Controls.Add(chkDelHidden, 0, 1);
            layout.Controls.Add(chkDelOffSlide, 0, 2);
            layout.Controls.Add(chkClearNotes, 0, 3);
            layout.Controls.Add(chkDelActions, 0, 4);
            var scanRun = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight };
            scanRun.Controls.Add(btnScan);
            scanRun.Controls.Add(btnRun);
            layout.Controls.Add(scanRun, 0, 5);
            layout.Controls.Add(lblCleanPreview, 0, 6);
            layout.SetColumnSpan(lblCleanPreview, 2);

            btnScan.Click += (s, e) => ScanClean();
            btnRun.Click += (s, e) => RunClean();

            grp.Controls.Add(layout);
            page.Controls.Add(grp);
            return page;
        }

        private class CleanStats { public int EmptySlides, HiddenShapes, OffSlideShapes, Notes, Actions; }

        private bool HasVisibleContent(PowerPoint.Slide slide)
        {
            try
            {
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shp = slide.Shapes[i];
                    try
                    {
                        if (shp.Visible != MsoTriState.msoTrue) continue;
                        // 占位符且无文字视为空内容（对齐收藏逻辑的 IsPlaceholder 语义）
                        bool isPlaceholder = false;
                        try { isPlaceholder = shp.PlaceholderFormat != null; } catch { }
                        if (isPlaceholder)
                        {
                            try
                            {
                                if (shp.HasTextFrame == MsoTriState.msoTrue && shp.TextFrame.HasText == MsoTriState.msoTrue
                                    && !string.IsNullOrWhiteSpace(shp.TextFrame.TextRange.Text)) return true;
                            }
                            catch { }
                            continue;
                        }
                        return true;   // 非占位符可见对象（含纯色块/图片/空文本框装饰）都算内容
                    }
                    catch { return true; }
                }
            }
            catch { }
            return false;
        }

        private CleanStats ScanCleanCore(bool dryRun)
        {
            var st = new CleanStats();
            var pres = App.ActivePresentation;
            if (pres == null) return st;
            var slideW = (float)pres.PageSetup.SlideWidth;
            var slideH = (float)pres.PageSetup.SlideHeight;

            foreach (PowerPoint.Slide slide in pres.Slides)
            {
                // 空白页（无可见形状）
                if (chkDelEmpty.Checked && !HasVisibleContent(slide)) st.EmptySlides++;
                for (int i = slide.Shapes.Count; i >= 1; i--)
                {
                    var shp = slide.Shapes[i];
                    try
                    {
                        if (chkDelHidden.Checked && shp.Visible == MsoTriState.msoFalse) { st.HiddenShapes++; if (!dryRun) shp.Delete(); continue; }
                        if (chkDelOffSlide.Checked)
                        {
                            float l = 0, t = 0, r = 0, b = 0;
                            try { l = (float)shp.Left; t = (float)shp.Top; r = l + (float)shp.Width; b = t + (float)shp.Height; }
                            catch { continue; }
                            // 完全在画布外（或超出过多）才删，避免误伤半出画布对象
                            if (r < -5 || b < -5 || l > slideW + 5 || t > slideH + 5)
                            {
                                st.OffSlideShapes++;
                                if (!dryRun) shp.Delete();
                                continue;
                            }
                        }
                        if (chkDelActions.Checked)
                        {
                            try
                            {
                                if (shp.ActionSettings != null)
                                {
                                    bool any = false;
                                    // ActionSettings 按 PpMouseActivation（Click/Hover）索引
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
                                                act.Action = PowerPoint.PpActionType.ppActionNone;
                                                any = true;
                                            }
                                        }
                                        catch { }
                                    }
                                    if (any) { st.Actions++; }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
                if (chkClearNotes.Checked)
                {
                    try
                    {
                        var notes = slide.NotesPage.Shapes.Placeholders[(int)PowerPoint.PpPlaceholderType.ppPlaceholderBody];
                        var tr = notes.TextFrame.TextRange;
                        if (tr != null && !string.IsNullOrWhiteSpace(tr.Text)) { st.Notes++; if (!dryRun) tr.Text = ""; }
                    }
                    catch { }
                }
            }
            return st;
        }

        private void ScanClean()
        {
            try
            {
                var st = ScanCleanCore(true);
                lblCleanPreview.Text =
                    "空白页 " + st.EmptySlides +
                    " · 隐藏对象 " + st.HiddenShapes +
                    " · 页外对象 " + st.OffSlideShapes +
                    " · 备注 " + st.Notes +
                    " · 动作链接 " + st.Actions;
            }
            catch (Exception ex) { Fail(ex.Message); }
        }

        private void RunClean()
        {
            try
            {
                if (MessageBox.Show("执行后将不可恢复（建议先保存文档），确定继续？", "清理瘦身",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                CleanStats st;
                Cursor = Cursors.WaitCursor;
                try { st = ScanCleanCore(false); }
                finally { Cursor = Cursors.Default; }
                Warn("清理完成：\n空白页 " + st.EmptySlides + " · 隐藏对象 " + st.HiddenShapes +
                     " · 页外对象 " + st.OffSlideShapes + " · 备注 " + st.Notes + " · 动作链接 " + st.Actions);
            }
            catch (Exception ex) { Fail(ex.Message); }
        }

        // ================= ③ 统一尺寸 =================
        private TabPage BuildSizeTab(Font font)
        {
            var page = new TabPage("统一尺寸") { Padding = new Padding(8) };
            var grp = new GroupBox { Text = "把形状统一到指定宽/高（可只填一项，保持比例）", Dock = DockStyle.Fill, Font = font };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(10) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));

            var lblScope = new Label { Text = "作用范围：", AutoSize = true, Font = font };
            var cmbScope = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = font };
            cmbScope.Items.AddRange(new[] { "选中形状", "当前页全部形状", "全部页面形状" });
            cmbScope.SelectedIndex = 0;
            var lblW = new Label { Text = "宽度(磅)：", AutoSize = true, Font = font };
            var txtW = new TextBox { Font = font };
            var lblH = new Label { Text = "高度(磅)：", AutoSize = true, Font = font };
            var txtH = new TextBox { Font = font };
            var chkImages = new CheckBox { Text = "图片强制等比（以宽或高为准）", Font = font, Checked = false };
            var btnRun = new Button { Text = "应用尺寸", Font = font, Height = 30, BackColor = Color.FromArgb(0x1F, 0x6F, 0xA2), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            AddRow(layout, 0, lblScope, cmbScope);
            AddRow(layout, 1, lblW, txtW);
            AddRow(layout, 2, lblH, txtH);
            layout.Controls.Add(chkImages, 1, 3);
            layout.Controls.Add(btnRun, 0, 4);
            layout.SetColumnSpan(btnRun, 2);

            btnRun.Click += (s, e) =>
            {
                try
                {
                    float? nw = ParseFloat(txtW.Text);
                    float? nh = ParseFloat(txtH.Text);
                    if (nw == null && nh == null) { Fail("请至少填写宽度或高度。"); return; }
                    string scope = cmbScope.SelectedItem as string ?? "选中形状";
                    var shapes = new List<PowerPoint.Shape>();
                    if (scope == "选中形状")
                    {
                        var sel = App.ActiveWindow.Selection;
                        if (sel != null && sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange != null)
                            for (int i = 1; i <= sel.ShapeRange.Count; i++) shapes.Add(sel.ShapeRange[i]);
                    }
                    else if (scope == "当前页全部形状")
                    {
                        var slide = App.ActivePresentation.Slides[App.ActiveWindow.View.Slide.SlideIndex];
                        for (int i = 1; i <= slide.Shapes.Count; i++) shapes.Add(slide.Shapes[i]);
                    }
                    else
                    {
                        foreach (PowerPoint.Slide slide in App.ActivePresentation.Slides)
                            for (int i = 1; i <= slide.Shapes.Count; i++) shapes.Add(slide.Shapes[i]);
                    }
                    int done = 0;
                    foreach (var shp in shapes)
                    {
                        try
                        {
                            float w = (float)shp.Width, h = (float)shp.Height;
                            if (w <= 0 || h <= 0) continue;
                            bool isPic = false;
                            try { isPic = shp.Type == MsoShapeType.msoPicture; } catch { }
                            float scale = 1f;
                            if (nw != null && nh != null)
                            {
                                if (isPic && chkImages.Checked)
                                {
                                    scale = Math.Min(nw.Value / w, nh.Value / h);
                                    shp.Width = w * scale; shp.Height = h * scale;
                                }
                                else { shp.Width = nw.Value; shp.Height = nh.Value; }
                            }
                            else if (nw != null) { scale = nw.Value / w; shp.Width = nw.Value; shp.Height = h * scale; }
                            else { scale = nh.Value / h; shp.Height = nh.Value; shp.Width = w * scale; }
                            done++;
                        }
                        catch { }
                    }
                    Warn("已统一 " + done + " 个形状尺寸。");
                }
                catch (Exception ex) { Fail(ex.Message); }
            };

            grp.Controls.Add(layout);
            page.Controls.Add(grp);
            return page;
        }

        // ================= ④ 批量插图片成网格 =================
        private TabPage BuildGridTab(Font font)
        {
            var page = new TabPage("图片网格插入") { Padding = new Padding(8) };
            var grp = new GroupBox { Text = "把多张图片排成整齐网格插入", Dock = DockStyle.Fill, Font = font };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(10) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));

            var lblCols = new Label { Text = "每行张数：", AutoSize = true, Font = font };
            var numCols = new NumericUpDown { Minimum = 1, Maximum = 8, Value = 3, Font = font, Width = 60 };
            var lblFiles = new Label { Text = "图片文件：", AutoSize = true, Font = font };
            var txtFiles = new TextBox { Font = font, Multiline = true, Height = 90, ScrollBars = ScrollBars.Vertical, ReadOnly = true };
            var btnPick = new Button { Text = "选择图片…", Font = font };
            var lblNewPage = new Label { Text = "插入方式：", AutoSize = true, Font = font };
            var cmbInsert = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = font };
            cmbInsert.Items.AddRange(new[] { "插入到新幻灯片（每页排满自动翻页）", "插入到当前页（超画布不裁剪）" });
            cmbInsert.SelectedIndex = 0;
            var btnRun = new Button { Text = "生成网格", Font = font, Height = 30, BackColor = Color.FromArgb(0x1F, 0x6F, 0xA2), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 110));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            AddRow(layout, 0, lblCols, numCols);
            layout.Controls.Add(lblFiles, 0, 1);
            var fileStack = MakeStack(txtFiles, btnPick, 300);
            layout.Controls.Add(fileStack, 1, 1);
            layout.SetRowSpan(fileStack, 2);
            AddRow(layout, 3, lblNewPage, cmbInsert);
            layout.Controls.Add(btnRun, 0, 4);
            layout.SetColumnSpan(btnRun, 2);

            var fileList = new List<string>();
            btnPick.Click += (s, e) =>
            {
                using (var dlg = new OpenFileDialog
                {
                    Multiselect = true,
                    Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp"
                })
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        fileList.Clear();
                        fileList.AddRange(dlg.FileNames);
                        txtFiles.Text = string.Join("\r\n", dlg.FileNames);
                    }
                }
            };

            btnRun.Click += (s, e) =>
            {
                try
                {
                    if (fileList.Count == 0) { Fail("请先选择图片文件。"); return; }
                    int cols = (int)numCols.Value;
                    bool toNewPage = cmbInsert.SelectedIndex == 0;
                    var pres = App.ActivePresentation;
                    if (pres == null) { Fail("没有打开的演示文稿。"); return; }
                    float sw = (float)pres.PageSetup.SlideWidth;
                    float sh = (float)pres.PageSetup.SlideHeight;

                    PowerPoint.Slide slide = null;
                    if (!toNewPage)
                    {
                        slide = pres.Slides[App.ActiveWindow.View.Slide.SlideIndex];
                    }
                    const float margin = 12f, gap = 8f;
                    int perPage = cols * 4;   // 默认每页 4 行
                    int onPage = 0;
                    for (int k = 0; k < fileList.Count; k++)
                    {
                        string file = fileList[k];
                        if (!System.IO.File.Exists(file)) continue;
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
                            // 等比 contain：完整显示图片（不裁切、不溢出单元格），短边居中
                            float pw = (float)pic.Width, ph = (float)pic.Height;
                            if (pw > 0 && ph > 0)
                            {
                                float scale = Math.Min(cellW / pw, cellH / ph);
                                float dw = pw * scale, dh = ph * scale;
                                // 视觉居中再补 2% 内边距避免贴边
                                pic.Width = dw; pic.Height = dh;
                                pic.Left = x + (cellW - dw) / 2f;
                                pic.Top = y + (cellH - dh) / 2f;
                            }
                            else
                            {
                                pic.Left = x; pic.Top = y; pic.Width = cellW; pic.Height = cellH;
                            }
                            onPage++;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("插入失败：" + file + "\n" + ex.Message, "图片网格");
                        }
                    }
                    Warn("已插入 " + fileList.Count + " 张图片" + (toNewPage ? "（按需新增幻灯片）。" : "到当前页。"));
                }
                catch (Exception ex) { Fail(ex.Message); }
            };

            grp.Controls.Add(layout);
            page.Controls.Add(grp);
            return page;
        }

        // ================= 辅助 =================
        private static float? ParseFloat(string s)
        {
            float v;
            return float.TryParse(s.Trim(), out v) && v > 0 ? v : (float?)null;
        }

        private static Control MakeStack(Control main, Control second, int mainWidth)
        {
            var p = new TableLayoutPanel { ColumnCount = 2, AutoSize = true };
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 84));
            main.Width = mainWidth;
            p.Controls.Add(main, 0, 0);
            p.Controls.Add(second, 1, 0);
            return p;
        }

        private static void AddRow(TableLayoutPanel layout, int row, Control left, Control right)
        {
            left.Dock = DockStyle.Fill;
            right.Dock = DockStyle.Fill;
            layout.Controls.Add(left, 0, row);
            layout.Controls.Add(right, 1, row);
        }
    }
}
