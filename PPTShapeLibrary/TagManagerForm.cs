using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>标签管理器：重命名 / 删除标签（C1）。</summary>
    public class TagManagerForm : Form
    {
        private readonly ListBox lstTags;
        private readonly Button btnRename;
        private readonly Button btnDelete;
        private readonly Button btnClose;

        public TagManagerForm()
        {
            Text = "标签管理";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(340, 380);

            var lbl = new Label
            {
                Text = "标签（点击选中后可重命名/删除）：",
                Location = new Point(16, 14),
                AutoSize = true,
                Font = new Font("Microsoft YaHei", 9F)
            };

            lstTags = new ListBox
            {
                Location = new Point(16, 44),
                Size = new Size(308, 240),
                Font = new Font("Microsoft YaHei", 9F)
            };

            btnRename = new Button { Text = "重命名", Location = new Point(16, 300), Size = new Size(90, 30) };
            btnDelete = new Button { Text = "删除", Location = new Point(116, 300), Size = new Size(90, 30) };
            btnClose = new Button { Text = "关闭", DialogResult = DialogResult.Cancel, Location = new Point(232, 300), Size = new Size(90, 30) };

            btnRename.Click += (s, e) => RenameTag();
            btnDelete.Click += (s, e) => DeleteTag();

            Controls.Add(lbl);
            Controls.Add(lstTags);
            Controls.Add(btnRename);
            Controls.Add(btnDelete);
            Controls.Add(btnClose);
            CancelButton = btnClose;

            Reload();
        }

        private void Reload()
        {
            lstTags.Items.Clear();
            foreach (var tag in ShapeLibraryService.GetAllTags())
            {
                int count = ShapeLibraryService.LoadAllItems().Count(m => m.GetTags().Contains(tag));
                lstTags.Items.Add(tag + "（" + count + " 项）");
                lstTags.Tag = tag;
            }
        }

        private string SelectedTag()
        {
            if (lstTags.SelectedIndex < 0) return null;
            string text = lstTags.SelectedItem.ToString();
            int idx = text.IndexOf('（');
            return idx > 0 ? text.Substring(0, idx) : text;
        }

        private void RenameTag()
        {
            string tag = SelectedTag();
            if (tag == null)
            {
                MessageBox.Show("请先选择一个标签。", "标签管理", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var input = new InputDialog("重命名标签", "新的标签名：", tag);
            if (input.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(input.Value)) return;
            string newTag = input.Value.Trim();
            foreach (var meta in ShapeLibraryService.LoadAllItems())
            {
                var tags = meta.GetTags().ToList();
                if (tags.Contains(tag))
                {
                    tags.Remove(tag);
                    if (!tags.Contains(newTag)) tags.Add(newTag);
                    ShapeLibraryService.SetTags(meta.Id, string.Join(",", tags));
                }
            }
            Reload();
        }

        private void DeleteTag()
        {
            string tag = SelectedTag();
            if (tag == null)
            {
                MessageBox.Show("请先选择一个标签。", "标签管理", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (MessageBox.Show(this, "从所有收藏中删除标签「" + tag + "」？", "标签管理",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            foreach (var meta in ShapeLibraryService.LoadAllItems())
            {
                var tags = meta.GetTags().ToList();
                if (tags.Contains(tag))
                {
                    tags.Remove(tag);
                    ShapeLibraryService.SetTags(meta.Id, string.Join(",", tags));
                }
            }
            Reload();
        }
    }
}
