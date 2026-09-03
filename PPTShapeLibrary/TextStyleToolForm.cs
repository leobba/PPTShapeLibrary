using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace PPTShapeLibrary
{
    /// <summary>文字样式工具（独立弹窗）。批量改字体/字号/加粗/斜体/颜色/对齐。</summary>
    public class TextStyleToolForm : Form
    {
        private readonly ComboBox cmbScope;
        private readonly ComboBox cmbFont;
        private readonly TextBox txtSize;
        private readonly ComboBox cmbBold;
        private readonly ComboBox cmbItalic;
        private readonly TextBox txtColor;
        private readonly ComboBox cmbAlign;

        public TextStyleToolForm()
        {
            Text = "文字样式";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(400, 340);
            var font = new Font("Microsoft YaHei", 9F);

            var grp = new GroupBox { Text = "批量修改文字样式", Location = new Point(14, 10), Size = new Size(372, 240), Font = font };

            cmbScope = AddCombo(grp, new[] { "选中形状", "当前页", "全部页面" }, 120, 26, 230);
            AddLabel(grp, "作用范围：", 16, 30);

            cmbFont = new ComboBox { Location = new Point(120, 62), Size = new Size(230, 25), Font = font };
            cmbFont.Items.AddRange(TextToolService.GetDocumentFonts().ToArray());
            cmbFont.Text = "";
            grp.Controls.Add(cmbFont);
            AddLabel(grp, "字体(留空不改)：", 16, 66);

            txtSize = new TextBox { Location = new Point(120, 98), Size = new Size(80, 25), Font = font };
            grp.Controls.Add(txtSize);
            AddLabel(grp, "字号(留空不改)：", 16, 102);

            cmbBold = AddCombo(grp, new[] { "不改", "加粗", "取消加粗" }, 120, 134);
            cmbItalic = AddCombo(grp, new[] { "不改", "斜体", "取消斜体" }, 230, 134, 120);
            AddLabel(grp, "加粗：", 16, 138);
            AddLabel(grp, "斜体：", 232, 138);

            txtColor = new TextBox { Location = new Point(120, 170), Size = new Size(230, 25), Font = font };
            grp.Controls.Add(txtColor);
            AddLabel(grp, "颜色 RRGGBB(留空不改)：", 16, 174);

            cmbAlign = AddCombo(grp, new[] { "不改", "左对齐", "居中", "右对齐" }, 120, 206, 150);
            AddLabel(grp, "对齐：", 16, 210);

            var btn = new Button
            {
                Text = "应用",
                Location = new Point(150, 258), Size = new Size(100, 30),
                Font = font, BackColor = Color.FromArgb(0x1F, 0x6F, 0xA2), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btn.Click += (s, e) => Apply();
            Controls.Add(grp);
            Controls.Add(btn);

            var close = new Button { Text = "关闭", DialogResult = DialogResult.Cancel, Location = new Point(284, 300), Size = new Size(100, 28), Font = font };
            CancelButton = close;
            Controls.Add(close);
        }

        private void Apply()
        {
            try
            {
                string scope = cmbScope.SelectedIndex == 0 ? "selection" : cmbScope.SelectedIndex == 1 ? "page" : "all";
                string fn = cmbFont.Text.Trim();
                float? size = null;
                float tmp;
                if (float.TryParse(txtSize.Text.Trim(), out tmp) && tmp > 0) size = tmp;
                bool? bold = null, italic = null;
                if (cmbBold.SelectedIndex == 1) bold = true;
                if (cmbBold.SelectedIndex == 2) bold = false;
                if (cmbItalic.SelectedIndex == 1) italic = true;
                if (cmbItalic.SelectedIndex == 2) italic = false;
                int? rgb = null;
                string c = txtColor.Text.Trim();
                if (c.Length == 6) { try { rgb = ShapeStyleExtractor.HexToRgb(c); } catch { } }
                string align = cmbAlign.SelectedIndex == 1 ? "left" : cmbAlign.SelectedIndex == 2 ? "center" : cmbAlign.SelectedIndex == 3 ? "right" : "";

                var r = TextToolService.ApplyTextStyle(scope, fn, size, bold, italic, rgb, align);
                MessageBox.Show(r.Message, "文字样式", MessageBoxButtons.OK,
                    r.Ok ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("操作失败：" + ex.Message, "文字样式", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static void AddLabel(Control parent, string text, int x, int y)
        {
            parent.Controls.Add(new Label { Text = text, Location = new Point(x, y), AutoSize = true, Font = new Font("Microsoft YaHei", 9F) });
        }

        private static ComboBox AddCombo(Control parent, string[] items, int x, int y, int w = 100)
        {
            var cb = new ComboBox { Location = new Point(x, y), Size = new Size(w, 25), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft YaHei", 9F) };
            cb.Items.AddRange(items);
            cb.SelectedIndex = 0;
            parent.Controls.Add(cb);
            return cb;
        }
    }
}
