using System.Drawing;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>简单的单行输入对话框（WinForms）。</summary>
    internal class InputDialog : Form
    {
        private readonly TextBox _textBox;

        public string Value
        {
            get { return _textBox.Text; }
        }

        public InputDialog(string title, string label, string initialValue = "")
        {
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(340, 104);

            var lbl = new Label
            {
                Text = label,
                Location = new Point(12, 12),
                AutoSize = true,
                Font = new Font("Microsoft YaHei", 9F)
            };
            _textBox = new TextBox
            {
                Location = new Point(12, 34),
                Size = new Size(316, 23),
                Text = initialValue
            };
            var ok = new Button { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(162, 68), Size = new Size(80, 28) };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(248, 68), Size = new Size(80, 28) };

            Controls.Add(lbl);
            Controls.Add(_textBox);
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
            _textBox.SelectAll();
        }
    }
}
