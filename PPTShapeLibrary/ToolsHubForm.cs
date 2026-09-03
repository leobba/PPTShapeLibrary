using System;
using System.Drawing;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// PPT 高级工具 Hub：承载仍需原生窗口的批处理工具
    /// （文字/查找等常用工具已 Web 化到面板「工具」视图）。
    /// </summary>
    public class ToolsHubForm : Form
    {
        public ToolsHubForm()
        {
            Text = "高级工具";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(500, 430);
            var font = new Font("Microsoft YaHei", 9F);

            var title = new Label
            {
                Text = "高级工具",
                Font = new Font("Microsoft YaHei UI", 15F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0x1B, 0x1F, 0x23),
                Location = new Point(18, 14), AutoSize = true
            };
            var subtitle = new Label
            {
                Text = "批处理类工具（操作会改动文档，建议先保存）",
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

            AddTool(flow, "导出图片", "整稿逐页导出 PNG/JPG（Web 界面）", () => OpenWebTool("tool5"));
            AddTool(flow, "清理瘦身", "删空白页/隐藏/页外对象、清空备注、移除动作（Web 界面）", () => OpenWebTool("tool6"));
            AddTool(flow, "统一尺寸", "按宽/高/等比批量统一（Web 界面）", () => OpenWebTool("tool7"));
            AddTool(flow, "图片网格", "多图 contain 排成网格插入（Web 界面）", () => OpenWebTool("tool8"));
            var note = new Label
            {
                Text = "文字样式、换字体、查找替换、清理等同样以 Web 工具窗口提供（Ribbon「工具」组直达）。",
                Font = new Font("Microsoft YaHei", 8.5F),
                ForeColor = Color.FromArgb(0x66, 0x70, 0x7E),
                Width = 440,
                Height = 30,
                Margin = new Padding(6, 2, 0, 0)
            };
            flow.Controls.Add(note);

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

        private void OpenWebTool(string viewKey)
        {
            using (var popup = new ToolPopupHostForm(viewKey))
            {
                popup.ShowDialog(this);
            }
        }
    }
}
