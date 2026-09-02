using System;
using System.Drawing;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// PPT 小工具总览 Hub：汇总全部工具入口，取代 Ribbon 直达单个对话框。
    /// </summary>
    public class ToolsHubForm : Form
    {
        public ToolsHubForm()
        {
            Text = "PPT 小工具";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(500, 430);
            var font = new Font("Microsoft YaHei", 9F);

            var title = new Label
            {
                Text = "PPT 小工具",
                Font = new Font("Microsoft YaHei UI", 15F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0x1B, 0x1F, 0x23),
                Location = new Point(18, 14), AutoSize = true
            };
            var subtitle = new Label
            {
                Text = "选择一个工具开始（部分操作会改动文档，建议先保存）",
                Font = font,
                ForeColor = Color.FromArgb(0x66, 0x70, 0x7E),
                Location = new Point(20, 46), AutoSize = true
            };

            var flow = new FlowLayoutPanel
            {
                Location = new Point(14, 78),
                Size = new Size(472, 320),
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true
            };

            AddTool(flow, "文字样式 / 字体", "批量改字体字号颜色对齐；一键整稿换字体", () => Show<PPTToolsForm>());
            AddTool(flow, "文本查找替换 / 清理", "跨页查找替换（保留格式）；去多余空白、统一标点", () => Show<PPTTextToolsForm>());
            AddTool(flow, "版式 / 批处理", "整稿导出 PNG/JPG；清理瘦身；统一尺寸；图片网格插入", () => Show<PPTFormatToolsForm>());

            Controls.Add(title);
            Controls.Add(subtitle);
            Controls.Add(flow);

            var btnClose = new Button
            {
                Text = "关闭", DialogResult = DialogResult.Cancel,
                Location = new Point(388, 390), Size = new Size(100, 28), Font = font
            };
            CancelButton = btnClose;
            Controls.Add(btnClose);
        }

        private void AddTool(FlowLayoutPanel flow, string name, string desc, Action open)
        {
            var card = new Button
            {
                Width = 440,
                Height = 66,
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei", 9.5F),
                Padding = new Padding(12, 0, 0, 0),
                Cursor = Cursors.Hand,
                Tag = desc
            };
            card.FlatAppearance.BorderSize = 1;
            card.FlatAppearance.BorderColor = Color.FromArgb(0xE1, 0xE7, 0xEF);
            card.Text = name + "\r\n" + desc;
            card.ForeColor = Color.FromArgb(0x1B, 0x1F, 0x23);
            card.MouseEnter += (s, e) => { card.FlatAppearance.BorderColor = Color.FromArgb(0x00, 0x67, 0xC0); };
            card.Click += (s, e) => open();
            flow.Controls.Add(card);
        }

        private void Show<T>() where T : Form, new()
        {
            using (var frm = new T())
            {
                frm.ShowDialog(this);
            }
        }
    }
}
