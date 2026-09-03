using System;
using System.Drawing;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>一键换字体工具（独立弹窗）。自动列出文档字体，原 → 新（整稿或当前页）。</summary>
    public class FontReplaceToolForm : Form
    {
        private readonly ComboBox cmbFrom;
        private readonly ComboBox cmbTo;
        private readonly ComboBox cmbScope;

        public FontReplaceToolForm()
        {
            Text = "一键换字体";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(400, 220);
            var font = new Font("Microsoft YaHei", 9F);

            var grp = new GroupBox { Text = "把文档中的字体整体替换为另一种", Location = new Point(14, 10), Size = new Size(372, 130), Font = font };

            cmbFrom = new ComboBox { Location = new Point(120, 26), Size = new Size(230, 25), Font = font, DropDownStyle = ComboBoxStyle.DropDown };
            cmbFrom.Items.AddRange(TextToolService.GetDocumentFonts().ToArray());
            if (cmbFrom.Items.Count > 0) cmbFrom.SelectedIndex = 0;
            grp.Controls.Add(cmbFrom);
            grp.Controls.Add(new Label { Text = "原字体：", Location = new Point(16, 30), AutoSize = true, Font = font });

            cmbTo = new ComboBox { Location = new Point(120, 60), Size = new Size(230, 25), Font = font, DropDownStyle = ComboBoxStyle.DropDown };
            cmbTo.Items.AddRange(new[] { "微软雅黑", "宋体", "黑体", "仿宋", "楷体", "等线", "Arial", "Times New Roman", "Calibri", "Segoe UI", "Georgia", "Verdana" });
            cmbTo.SelectedIndex = 0;
            grp.Controls.Add(cmbTo);
            grp.Controls.Add(new Label { Text = "新字体：", Location = new Point(16, 64), AutoSize = true, Font = font });

            cmbScope = new ComboBox { Location = new Point(120, 94), Size = new Size(150, 25), DropDownStyle = ComboBoxStyle.DropDownList, Font = font };
            cmbScope.Items.AddRange(new[] { "整份演示文稿", "当前页" });
            cmbScope.SelectedIndex = 0;
            grp.Controls.Add(cmbScope);
            grp.Controls.Add(new Label { Text = "范围：", Location = new Point(16, 98), AutoSize = true, Font = font });

            var btn = new Button
            {
                Text = "开始替换",
                Location = new Point(150, 150), Size = new Size(100, 30),
                Font = font, BackColor = Color.FromArgb(0x7B, 0x1F, 0xA2), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btn.Click += (s, e) => Replace();
            Controls.Add(grp);
            Controls.Add(btn);

            var close = new Button { Text = "关闭", DialogResult = DialogResult.Cancel, Location = new Point(284, 182), Size = new Size(100, 28), Font = font };
            CancelButton = close;
            Controls.Add(close);
        }

        private void Replace()
        {
            try
            {
                string from = cmbFrom.Text.Trim();
                string to = cmbTo.Text.Trim();
                if (from.Length == 0 || to.Length == 0)
                {
                    MessageBox.Show("请选择原字体和新字体。", "一键换字体");
                    return;
                }
                bool allSlides = cmbScope.SelectedIndex == 0;
                var r = TextToolService.ReplaceFont(from, to, allSlides);
                MessageBox.Show(r.Message, "一键换字体", MessageBoxButtons.OK,
                    r.Ok ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("替换失败：" + ex.Message, "一键换字体", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
