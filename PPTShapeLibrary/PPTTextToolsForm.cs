using System;
using System.Drawing;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// PPT 文本工具：跨页查找/替换（保留格式） + 文字清理（去多余空白/统一标点/合并断行）。
    /// 替换用 TextRange 逐处定位实现，保留原有字体/字号/颜色等格式。
    /// </summary>
    public class PPTTextToolsForm : Form
    {
        private readonly ComboBox cmbScope;
        private readonly TextBox txtFind;
        private readonly TextBox txtReplace;
        private readonly CheckBox chkCase;
        private readonly CheckBox chkCleanWs;
        private readonly CheckBox chkCleanPunct;
        private readonly Button btnReplace;
        private readonly Button btnClean;
        private GroupBox _grpReplace;
        private GroupBox _grpClean;

        public PPTTextToolsForm() : this("both") { }

        /// <summary>mode: "replace"(仅查找替换) / "clean"(仅文字清理) / "both"(默认全部)。</summary>
        public PPTTextToolsForm(string mode)
        {
            Text = mode == "replace" ? "查找 / 替换" : mode == "clean" ? "文字清理" : "文本工具";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(430, 330);
            var font = new Font("Microsoft YaHei", 9F);

            // ---- 查找替换 ----
            var grp1 = new GroupBox { Text = "查找 / 替换（保留格式）", Location = new Point(14, 10), Size = new Size(400, 160), Font = font };
            _grpReplace = grp1;            var lblScope = new Label { Text = "作用范围：", Location = new Point(16, 28), AutoSize = true, Font = font };
            cmbScope = new ComboBox
            {
                Location = new Point(130, 25), Size = new Size(240, 25), DropDownStyle = ComboBoxStyle.DropDownList, Font = font
            };
            cmbScope.Items.AddRange(new[] { "当前页", "全部页面" });
            cmbScope.SelectedIndex = 1;
            grp1.Controls.Add(lblScope); grp1.Controls.Add(cmbScope);

            var lblFind = new Label { Text = "查找内容：", Location = new Point(16, 62), AutoSize = true, Font = font };
            txtFind = new TextBox { Location = new Point(130, 58), Size = new Size(240, 25), Font = font };
            grp1.Controls.Add(lblFind); grp1.Controls.Add(txtFind);

            var lblRep = new Label { Text = "替换为：", Location = new Point(16, 96), AutoSize = true, Font = font };
            txtReplace = new TextBox { Location = new Point(130, 92), Size = new Size(240, 25), Font = font };
            chkCase = new CheckBox { Text = "区分大小写", Location = new Point(16, 126), AutoSize = true, Font = font };
            grp1.Controls.Add(lblRep); grp1.Controls.Add(txtReplace); grp1.Controls.Add(chkCase);

            btnReplace = new Button
            {
                Text = "替换全部", Location = new Point(280, 122), Size = new Size(100, 28),
                Font = font, BackColor = Color.FromArgb(0x1F, 0x6F, 0xA2), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnReplace.Click += (s, e) => DoReplace();
            grp1.Controls.Add(btnReplace);

            // ---- 清理 ----
            var grp2 = new GroupBox { Text = "文字清理（整份演示文稿）", Location = new Point(14, 178), Size = new Size(400, 110), Font = font };
            _grpClean = grp2;
            chkCleanWs = new CheckBox
            {
                Text = "去多余空白（连续空格/全角空格、行首尾空格）",
                Location = new Point(16, 24), AutoSize = true, Font = font, Checked = true
            };
            chkCleanPunct = new CheckBox
            {
                Text = "统一标点（英文逗号句号 → 中文全角）",
                Location = new Point(16, 52), AutoSize = true, Font = font, Checked = false
            };
            grp2.Controls.Add(chkCleanWs); grp2.Controls.Add(chkCleanPunct);

            btnClean = new Button
            {
                Text = "开始清理", Location = new Point(280, 68), Size = new Size(100, 28),
                Font = font, BackColor = Color.FromArgb(0x7B, 0x1F, 0xA2), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnClean.Click += (s, e) => DoClean();
            grp2.Controls.Add(btnClean);

            var btnClose = new Button { Text = "关闭", DialogResult = DialogResult.Cancel, Location = new Point(314, 296), Size = new Size(100, 28), Font = font };
            CancelButton = btnClose;

            Controls.Add(grp1); Controls.Add(grp2); Controls.Add(btnClose);

            // mode 聚焦：隐藏非目标区块，并把窗体调紧凑
            if (mode == "replace") { grp2.Visible = false; ClientSize = new Size(430, 210); btnClose.Location = new Point(314, 178); }
            else if (mode == "clean") { grp1.Visible = false; ClientSize = new Size(430, 210); btnClose.Location = new Point(314, 178); }
        }

        private void DoReplace()
        {
            try
            {
                string find = txtFind.Text;
                string rep = txtReplace.Text;
                string scopeSel = cmbScope.SelectedItem as string ?? "全部页面";
                bool allSlides = scopeSel == "全部页面";
                Cursor = Cursors.WaitCursor;
                TextToolService.ToolResult r;
                try { r = TextToolService.FindReplace(find, rep, chkCase.Checked, allSlides); }
                finally { Cursor = Cursors.Default; }
                MessageBox.Show(r.Message, "查找 / 替换", MessageBoxButtons.OK,
                    r.Ok ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("替换失败：" + ex.Message, "文本工具");
            }
        }

        private void DoClean()
        {
            try
            {
                if (!chkCleanWs.Checked && !chkCleanPunct.Checked)
                {
                    MessageBox.Show("请至少勾选一项清理规则。", "文本工具");
                    return;
                }
                Cursor = Cursors.WaitCursor;
                TextToolService.ToolResult r;
                try { r = TextToolService.CleanText(chkCleanWs.Checked, chkCleanPunct.Checked); }
                finally { Cursor = Cursors.Default; }
                MessageBox.Show(r.Message + "\n（文本会按首字符格式统一；多格式混排需保留时请谨慎使用或事后撤销。）",
                    "文字清理", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("清理失败：" + ex.Message, "文本工具");
            }
        }
    }
}
