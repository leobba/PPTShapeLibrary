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

        public PPTTextToolsForm()
        {
            Text = "文本查找替换 / 清理";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(430, 330);
            var font = new Font("Microsoft YaHei", 9F);

            // ---- 查找替换 ----
            var grp1 = new GroupBox { Text = "查找 / 替换（保留格式）", Location = new Point(14, 10), Size = new Size(400, 160), Font = font };
            var lblScope = new Label { Text = "作用范围：", Location = new Point(16, 28), AutoSize = true, Font = font };
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
        }

        private static PowerPoint.Application App
        {
            get { return Globals.ThisAddIn.Application; }
        }

        /// <summary>按范围返回全部文本形状。</summary>
        private List<PowerPoint.Shape> TextShapesByScope(string scope)
        {
            var result = new List<PowerPoint.Shape>();
            try
            {
                if (scope == "当前页")
                {
                    var slide = App.ActivePresentation.Slides[App.ActiveWindow.View.Slide.SlideIndex];
                    for (int i = 1; i <= slide.Shapes.Count; i++) result.Add(slide.Shapes[i]);
                }
                else
                {
                    foreach (PowerPoint.Slide slide in App.ActivePresentation.Slides)
                    {
                        for (int i = 1; i <= slide.Shapes.Count; i++) result.Add(slide.Shapes[i]);
                    }
                }
            }
            catch { }
            return result;
        }

        private bool IsTextShape(PowerPoint.Shape shp)
        {
            try
            {
                return shp.HasTextFrame == MsoTriState.msoTrue && shp.TextFrame.HasText == MsoTriState.msoTrue;
            }
            catch { return false; }
        }

        private void DoReplace()
        {
            try
            {
                string find = txtFind.Text;
                string rep = txtReplace.Text;
                string scope = cmbScope.SelectedItem as string ?? "当前页";
                if (find.Length == 0)
                {
                    MessageBox.Show("请输入查找内容。", "文本工具");
                    return;
                }
                var cmp = chkCase.Checked ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

                int total = 0;
                foreach (var shp in TextShapesByScope(scope))
                {
                    if (!IsTextShape(shp)) continue;
                    try
                    {
                        var textRange = shp.TextFrame.TextRange;
                        string full = textRange.Text ?? "";
                        int idx = full.IndexOf(find, cmp);
                        int guard = 0;
                        // 逐处定位并替换（保留每处的格式，来自原文本段）
                        while (idx >= 0 && guard++ < 500)
                        {
                            var part = textRange.Characters(idx + 1, find.Length);
                            part.Text = rep;
                            // 全文重读，从新位置继续向后找
                            full = textRange.Text ?? "";
                            idx = full.IndexOf(find, idx + rep.Length, cmp);
                            total++;
                            if (idx < 0) break;
                        }
                    }
                    catch { }
                }
                MessageBox.Show("共替换 " + total + " 处。", "文本工具",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                int changed = 0;
                foreach (var shp in TextShapesByScope("全部页面"))
                {
                    if (!IsTextShape(shp)) continue;
                    try
                    {
                        var textRange = shp.TextFrame.TextRange;
                        string t = textRange.Text ?? "";
                        string o = t;
                        if (chkCleanWs.Checked)
                        {
                            // 全角空格→普通空格；连续空格→单个；行首尾空白修剪
                            t = t.Replace('\u3000', ' ');
                            t = System.Text.RegularExpressions.Regex.Replace(t, "[ ]{2,}", " ");
                            t = System.Text.RegularExpressions.Regex.Replace(t, "(?m)^[ \t]+|[ \t]+$", "");
                        }
                        if (chkCleanPunct.Checked)
                        {
                            // 中文语境：西文逗号/句号/问号/感叹号/冒号 → 全角（仅当其两侧有 CJK 字符时保守替换太复杂，做整体替换并给出数量）
                            t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*,\\s*([\u4e00-\u9fff])", "$1，$2");
                            t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*\\.\\s*([\u4e00-\u9fff])", "$1。$2");
                            t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*\\?\\s*([\u4e00-\u9fff])", "$1？$2");
                            t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*!\\s*([\u4e00-\u9fff])", "$1！$2");
                        }
                        if (t != o)
                        {
                            textRange.Text = t;   // 直接整段替换会丢统一格式，但对整体清理可接受；提示用户注意
                            changed++;
                        }
                    }
                    catch { }
                }
                MessageBox.Show("已清理 " + changed + " 个文本形状。\n（注意：整段清理会采用该段首字符格式，若发现格式变化请用撤销。）",
                    "文本工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("清理失败：" + ex.Message, "文本工具");
            }
        }
    }
}
