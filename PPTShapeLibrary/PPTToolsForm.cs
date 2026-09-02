using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTShapeLibrary
{
    /// <summary>PPT 小工具：批量修改文字样式 + 一键更换字体。</summary>
    public class PPTToolsForm : Form
    {
        private static readonly string[] CommonFonts =
        {
            "微软雅黑", "宋体", "黑体", "仿宋", "楷体", "等线",
            "Arial", "Times New Roman", "Calibri", "Segoe UI", "Georgia", "Verdana"
        };

        // 文字样式区
        private readonly ComboBox cmbTextScope;
        private readonly ComboBox cmbFontName;
        private readonly TextBox txtFontSize;
        private readonly ComboBox cmbBold;
        private readonly ComboBox cmbItalic;
        private readonly TextBox txtColor;
        private readonly ComboBox cmbAlign;

        // 换字体区
        private readonly ComboBox cmbFromFont;
        private readonly ComboBox cmbToFont;
        private readonly ComboBox cmbFontScope;

        public PPTToolsForm()
        {
            Text = "PPT 小工具";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(440, 520);

            // ---------- 批量修改文字样式 ----------
            var grp1 = new GroupBox
            {
                Text = "批量修改文字样式",
                Location = new Point(16, 12),
                Size = new Size(408, 260),
                Font = new Font("Microsoft YaHei", 9F)
            };

            AddLabel(grp1, "作用范围：", 20, 30);
            cmbTextScope = AddCombo(grp1, new[] { "选中形状", "当前页", "全部页面" }, 130, 27);

            AddLabel(grp1, "字体（留空不改）：", 20, 62);
            cmbFontName = new ComboBox
            {
                Location = new Point(130, 59),
                Size = new Size(250, 25),
                DropDownStyle = ComboBoxStyle.DropDown,
                Font = new Font("Microsoft YaHei", 9F)
            };
            cmbFontName.Items.AddRange(CommonFonts);
            grp1.Controls.Add(cmbFontName);

            AddLabel(grp1, "字号（留空不改）：", 20, 94);
            txtFontSize = new TextBox { Location = new Point(130, 91), Size = new Size(80, 25), Font = new Font("Microsoft YaHei", 9F) };
            grp1.Controls.Add(txtFontSize);

            AddLabel(grp1, "加粗：", 20, 126);
            cmbBold = AddCombo(grp1, new[] { "不改", "加粗", "取消加粗" }, 130, 123);
            AddLabel(grp1, "斜体：", 240, 126);
            cmbItalic = AddCombo(grp1, new[] { "不改", "斜体", "取消斜体" }, 300, 123, 80);

            AddLabel(grp1, "颜色（RRGGBB，留空不改）：", 20, 158);
            txtColor = new TextBox { Location = new Point(230, 155), Size = new Size(150, 25), Font = new Font("Microsoft YaHei", 9F) };
            grp1.Controls.Add(txtColor);

            AddLabel(grp1, "对齐：", 20, 190);
            cmbAlign = AddCombo(grp1, new[] { "不改", "左对齐", "居中", "右对齐" }, 130, 187);

            var btnApply = new Button
            {
                Text = "应用文字样式",
                Location = new Point(130, 222),
                Size = new Size(150, 30),
                Font = new Font("Microsoft YaHei", 9F),
                BackColor = Color.FromArgb(0x1F, 0x6F, 0xA2),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnApply.Click += (s, e) => ApplyTextStyle();
            grp1.Controls.Add(btnApply);

            // ---------- 一键更换字体 ----------
            var grp2 = new GroupBox
            {
                Text = "一键更换字体",
                Location = new Point(16, 284),
                Size = new Size(408, 180),
                Font = new Font("Microsoft YaHei", 9F)
            };

            AddLabel(grp2, "原字体：", 20, 30);
            cmbFromFont = AddCombo(grp2, new string[0], 130, 27, 250);
            cmbFromFont.DropDownStyle = ComboBoxStyle.DropDown;

            AddLabel(grp2, "新字体：", 20, 62);
            cmbToFont = new ComboBox
            {
                Location = new Point(130, 59),
                Size = new Size(250, 25),
                DropDownStyle = ComboBoxStyle.DropDown,
                Font = new Font("Microsoft YaHei", 9F)
            };
            cmbToFont.Items.AddRange(CommonFonts);
            cmbToFont.Text = "微软雅黑";
            grp2.Controls.Add(cmbToFont);

            AddLabel(grp2, "作用范围：", 20, 94);
            cmbFontScope = AddCombo(grp2, new[] { "当前页", "全部页面" }, 130, 91);

            var btnReplace = new Button
            {
                Text = "开始替换",
                Location = new Point(130, 132),
                Size = new Size(150, 30),
                Font = new Font("Microsoft YaHei", 9F),
                BackColor = Color.FromArgb(0x7B, 0x1F, 0xA2),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnReplace.Click += (s, e) => ReplaceFonts();
            grp2.Controls.Add(btnReplace);

            var btnClose = new Button
            {
                Text = "关闭",
                DialogResult = DialogResult.Cancel,
                Location = new Point(320, 478),
                Size = new Size(100, 30),
                Font = new Font("Microsoft YaHei", 9F)
            };
            CancelButton = btnClose;

            Controls.Add(grp1);
            Controls.Add(grp2);
            Controls.Add(btnClose);

            LoadFontsIntoFromBox();
        }

        private static void AddLabel(Control parent, string text, int x, int y)
        {
            var lbl = new Label
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                Font = new Font("Microsoft YaHei", 9F)
            };
            parent.Controls.Add(lbl);
        }

        private static ComboBox AddCombo(Control parent, string[] items, int x, int y, int w = 150)
        {
            var cb = new ComboBox
            {
                Location = new Point(x, y),
                Size = new Size(w, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei", 9F)
            };
            cb.Items.AddRange(items);
            if (items.Length > 0) cb.SelectedIndex = 0;
            parent.Controls.Add(cb);
            return cb;
        }

        // ---------------- 逻辑 ----------------

        private static PowerPoint.Application App
        {
            get { return Globals.ThisAddIn.Application; }
        }

        private List<PowerPoint.Shape> GetShapesByScope(string scope)
        {
            var result = new List<PowerPoint.Shape>();
            try
            {
                if (scope == "选中形状")
                {
                    var sel = App.ActiveWindow.Selection;
                    if (sel != null && sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange != null)
                    {
                        for (int i = 1; i <= sel.ShapeRange.Count; i++)
                        {
                            result.Add(sel.ShapeRange[i]);
                        }
                    }
                }
                else if (scope == "当前页")
                {
                    var slide = App.ActivePresentation.Slides[App.ActiveWindow.View.Slide.SlideIndex];
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        result.Add(slide.Shapes[i]);
                    }
                }
                else // 全部页面
                {
                    foreach (PowerPoint.Slide slide in App.ActivePresentation.Slides)
                    {
                        for (int i = 1; i <= slide.Shapes.Count; i++)
                        {
                            result.Add(slide.Shapes[i]);
                        }
                    }
                }
            }
            catch { }
            return result;
        }

        private void ApplyTextStyle()
        {
            try
            {
                string scope = cmbTextScope.SelectedItem as string ?? "选中形状";
                string fontName = cmbFontName.Text.Trim();
                string sizeText = txtFontSize.Text.Trim();
                float? size = null;
                float tmp;
                if (float.TryParse(sizeText, out tmp) && tmp > 0) size = tmp;
                bool? bold = null, italic = null;
                string b = cmbBold.SelectedItem as string;
                string i = cmbItalic.SelectedItem as string;
                if (b == "加粗") bold = true;
                if (b == "取消加粗") bold = false;
                if (i == "斜体") italic = true;
                if (i == "取消斜体") italic = false;
                int? color = null;
                string c = txtColor.Text.Trim();
                if (c.Length == 6)
                {
                    try { color = ShapeStyleExtractor.HexToRgb(c); } catch { }
                }
                int? align = null;
                string a = cmbAlign.SelectedItem as string;
                if (a == "左对齐") align = 1;
                if (a == "居中") align = 2;
                if (a == "右对齐") align = 3;

                int changed = 0;
                foreach (var shp in GetShapesByScope(scope))
                {
                    try
                    {
                        if (shp.HasTextFrame != MsoTriState.msoTrue) continue;
                        if (shp.TextFrame.HasText != MsoTriState.msoTrue) continue;
                        var range = shp.TextFrame.TextRange;
                        if (!string.IsNullOrEmpty(fontName)) range.Font.Name = fontName;
                        if (size != null) range.Font.Size = size.Value;
                        if (bold != null) range.Font.Bold = bold.Value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                        if (italic != null) range.Font.Italic = italic.Value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                        if (color != null) range.Font.Color.RGB = color.Value;
                        if (align != null) range.ParagraphFormat.Alignment = (PowerPoint.PpParagraphAlignment)align.Value;
                        changed++;
                    }
                    catch { }
                }
                MessageBox.Show("已处理 " + changed + " 个文字形状。", "PPT 小工具",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("操作失败：" + ex.Message, "PPT 小工具");
            }
        }

        private void LoadFontsIntoFromBox()
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
                            if (shp.HasTextFrame == MsoTriState.msoTrue && shp.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                                string f = shp.TextFrame.TextRange.Font.Name;
                                if (!string.IsNullOrEmpty(f)) fonts.Add(f);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            cmbFromFont.Items.Clear();
            foreach (var f in fonts.OrderBy(x => x)) cmbFromFont.Items.Add(f);
            foreach (var f in CommonFonts)
            {
                if (!cmbFromFont.Items.Contains(f)) cmbFromFont.Items.Add(f);
            }
            if (cmbFromFont.Items.Count > 0) cmbFromFont.SelectedIndex = 0;
        }

        private void ReplaceFonts()
        {
            try
            {
                string fromFont = cmbFromFont.Text.Trim();
                string toFont = cmbToFont.Text.Trim();
                string scope = cmbFontScope.SelectedItem as string ?? "当前页";
                if (fromFont == "" || toFont == "")
                {
                    MessageBox.Show("请选择原字体和新字体。", "PPT 小工具");
                    return;
                }
                if (fromFont.Equals(toFont, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("原字体和新字体相同。", "PPT 小工具");
                    return;
                }

                string effScope = scope == "全部页面" ? "全部页面" : "当前页";
                int changed = 0;
                foreach (var shp in GetShapesByScope(effScope))
                {
                    try
                    {
                        if (shp.HasTextFrame != MsoTriState.msoTrue) continue;
                        if (shp.TextFrame.HasText != MsoTriState.msoTrue) continue;
                        var range = shp.TextFrame.TextRange;
                        if (range.Font.Name.Equals(fromFont, StringComparison.OrdinalIgnoreCase))
                        {
                            range.Font.Name = toFont;
                            changed++;
                        }
                    }
                    catch { }
                }
                MessageBox.Show("已将 " + changed + " 处文字从「" + fromFont + "」替换为「" + toFont + "」。", "PPT 小工具",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("替换失败：" + ex.Message, "PPT 小工具");
            }
        }
    }
}
