using System;
using System.Drawing;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>插件设置对话框。</summary>
    public class SettingsForm : Form
    {
        private readonly ComboBox cmbTheme;
        private readonly ComboBox cmbInsert;
        private readonly ComboBox cmbCardSize;
        private readonly CheckBox chkAutoName;
        private readonly CheckBox chkClickToInsert;
        private readonly CheckBox chkThemeAdapt;
        private readonly TextBox txtPageSize;
        private readonly NumericUpDown numWidth;

        public Settings Result;

        public SettingsForm(Settings current)
        {
            Text = "PPT形状库设置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(360, 260);

            var lblTheme = new Label { Text = "主题：", Location = new Point(24, 24), AutoSize = true, Font = new Font("Microsoft YaHei", 9F) };
            cmbTheme = new ComboBox
            {
                Location = new Point(130, 21),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei", 9F)
            };
            cmbTheme.Items.Add("浅色");
            cmbTheme.Items.Add("深色");
            cmbTheme.SelectedIndex = current.Theme == "dark" ? 1 : 0;

            var lblInsert = new Label { Text = "插入位置：", Location = new Point(24, 62), AutoSize = true, Font = new Font("Microsoft YaHei", 9F) };
            cmbInsert = new ComboBox
            {
                Location = new Point(130, 59),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei", 9F)
            };
            cmbInsert.Items.Add("页面居中");
            cmbInsert.Items.Add("原始位置");
            cmbInsert.SelectedIndex = current.InsertPosition == "original" ? 1 : 0;

            var lblName = new Label { Text = "自动命名：", Location = new Point(24, 100), AutoSize = true, Font = new Font("Microsoft YaHei", 9F) };
            chkAutoName = new CheckBox
            {
                Location = new Point(130, 96),
                Size = new Size(200, 24),
                Text = "按形状类型自动命名",
                Font = new Font("Microsoft YaHei", 9F),
                Checked = current.AutoName
            };

            var lblPage = new Label { Text = "分页大小：", Location = new Point(24, 138), AutoSize = true, Font = new Font("Microsoft YaHei", 9F) };
            txtPageSize = new TextBox
            {
                Location = new Point(130, 135),
                Size = new Size(80, 25),
                Text = current.PageSize.ToString(),
                Font = new Font("Microsoft YaHei", 9F)
            };

            var lblWidth = new Label { Text = "面板宽度：", Location = new Point(24, 176), AutoSize = true, Font = new Font("Microsoft YaHei", 9F) };
            numWidth = new NumericUpDown
            {
                Location = new Point(130, 173),
                Size = new Size(80, 25),
                Minimum = 220,
                Maximum = 500,
                Value = current.PaneWidth > 0 ? current.PaneWidth : 300
            };

            // 第二列：卡片大小 / 交互 / 主题适配
            var lblCard = new Label { Text = "卡片大小：", Location = new Point(24, 214), AutoSize = true, Font = new Font("Microsoft YaHei", 9F) };
            cmbCardSize = new ComboBox
            {
                Location = new Point(130, 211),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei", 9F)
            };
            cmbCardSize.Items.Add("紧凑");
            cmbCardSize.Items.Add("标准");
            cmbCardSize.Items.Add("大图");
            cmbCardSize.SelectedIndex = current.CardSize == "standard" ? 1 : current.CardSize == "large" ? 2 : 0;

            chkClickToInsert = new CheckBox
            {
                Location = new Point(24, 252),
                Size = new Size(300, 24),
                Text = "单击卡片直接插入（取消则单击预览、双击插入）",
                Font = new Font("Microsoft YaHei", 9F),
                Checked = current.ClickToInsert
            };

            chkThemeAdapt = new CheckBox
            {
                Location = new Point(24, 284),
                Size = new Size(300, 24),
                Text = "应用样式时把主色映射到当前文档主题色",
                Font = new Font("Microsoft YaHei", 9F),
                Checked = current.ThemeAdaptStyle
            };

            var btnTagMgr = new Button
            {
                Text = "标签管理…",
                Location = new Point(130, 316),
                Size = new Size(120, 28),
                Font = new Font("Microsoft YaHei", 9F)
            };
            btnTagMgr.Click += (s, e) => OpenTagManager();

            var btnOk = new Button { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(150, 356), Size = new Size(90, 30) };
            var btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(250, 356), Size = new Size(90, 30) };
            btnOk.Click += (s, e) => BuildResult();

            Controls.Add(lblTheme); Controls.Add(cmbTheme);
            Controls.Add(lblInsert); Controls.Add(cmbInsert);
            Controls.Add(lblName); Controls.Add(chkAutoName);
            Controls.Add(lblPage); Controls.Add(txtPageSize);
            Controls.Add(lblWidth); Controls.Add(numWidth);
            Controls.Add(lblCard); Controls.Add(cmbCardSize);
            Controls.Add(chkClickToInsert);
            Controls.Add(chkThemeAdapt);
            Controls.Add(btnTagMgr);
            Controls.Add(btnOk); Controls.Add(btnCancel);
            AcceptButton = btnOk;
            CancelButton = btnCancel;
            ClientSize = new Size(380, 396);
        }

        private void OpenTagManager()
        {
            using (var frm = new TagManagerForm())
            {
                frm.ShowDialog(this);
            }
        }

        private void BuildResult()
        {
            var s = new Settings();
            s.Theme = cmbTheme.SelectedIndex == 1 ? "dark" : "light";
            s.InsertPosition = cmbInsert.SelectedIndex == 1 ? "original" : "center";
            s.AutoName = chkAutoName.Checked;
            s.ClickToInsert = chkClickToInsert.Checked;
            s.ThemeAdaptStyle = chkThemeAdapt.Checked;
            s.CardSize = cmbCardSize.SelectedIndex == 1 ? "standard" : cmbCardSize.SelectedIndex == 2 ? "large" : "compact";
            int ps;
            if (!int.TryParse(txtPageSize.Text, out ps) || ps <= 0) ps = 200;
            s.PageSize = ps;
            s.PaneWidth = (int)numWidth.Value;
            Result = s;
        }
    }
}
