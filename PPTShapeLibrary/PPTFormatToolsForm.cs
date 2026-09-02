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
                    string ext = cmbFmt.SelectedItem as string ?? "PNG";
                    int w = 0;
                    string scale = cmbScale.SelectedItem as string;
                    if (scale != null && scale.Contains("1280")) w = 1280;
                    if (scale != null && scale.Contains("1920")) w = 1920;

                    Cursor = Cursors.WaitCursor;
                    TextToolService.ToolResult r;
                    try { r = FormatToolService.ExportImages(dir, ext, w); }
                    finally { Cursor = Cursors.Default; }
                    if (r.Ok) Warn(r.Message); else Fail(r.Message);
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

        private void ScanClean()
        {
            try
            {
                var opts = ReadCleanOptions();
                var r = FormatToolService.ScanClean(opts);
                lblCleanPreview.Text = r.Ok ? r.Message : ("扫描失败：" + r.Message);
            }
            catch (Exception ex) { Fail(ex.Message); }
        }

        private void RunClean()
        {
            try
            {
                if (MessageBox.Show("执行后将不可恢复（建议先保存文档），确定继续？", "清理瘦身",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                var opts = ReadCleanOptions();
                Cursor = Cursors.WaitCursor;
                TextToolService.ToolResult r;
                try { r = FormatToolService.RunClean(opts); }
                finally { Cursor = Cursors.Default; }
                if (r.Ok) Warn(r.Message); else Fail(r.Message);
            }
            catch (Exception ex) { Fail(ex.Message); }
        }

        private FormatToolService.CleanOptions ReadCleanOptions()
        {
            return new FormatToolService.CleanOptions
            {
                DelEmptySlides = chkDelEmpty.Checked,
                DelHiddenShapes = chkDelHidden.Checked,
                DelOffSlideShapes = chkDelOffSlide.Checked,
                ClearNotes = chkClearNotes.Checked,
                RemoveActions = chkDelActions.Checked
            };
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
                    string scope = cmbScope.SelectedItem as string ?? "选中形状";
                    string scopeKey = scope == "选中形状" ? "selection" : scope == "当前页全部形状" ? "page" : "all";
                    Cursor = Cursors.WaitCursor;
                    TextToolService.ToolResult r;
                    try { r = FormatToolService.ResizeShapes(scopeKey, nw, nh, chkImages.Checked); }
                    finally { Cursor = Cursors.Default; }
                    if (r.Ok) Warn(r.Message); else Fail(r.Message);
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
                    Cursor = Cursors.WaitCursor;
                    TextToolService.ToolResult r;
                    try { r = FormatToolService.ImageGrid(fileList, cols, toNewPage); }
                    finally { Cursor = Cursors.Default; }
                    if (r.Ok) Warn(r.Message); else Fail(r.Message);
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
