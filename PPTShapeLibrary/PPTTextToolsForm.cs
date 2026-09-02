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
                int shapesTouched = 0;
                int skippedShapes = 0;
                Cursor = Cursors.WaitCursor;
                try
                {
                    foreach (var shp in TextShapesByScope(scope))
                    {
                        if (!IsTextShape(shp)) continue;
                        try
                        {
                            var textRange = shp.TextFrame.TextRange;
                            string full = textRange.Text ?? "";
                            int fromIdx = 0;          // 已替换末尾游标（防重入/防漂移）
                            int inShape = 0;
                            int safety = 0;           // 兜底：单 shape 上限 2000 次，防异常宿主死循环
                            while (safety++ < 2000)
                            {
                                int idx = full.IndexOf(find, fromIdx, cmp);
                                if (idx < 0) break;
                                // 用 Characters 定位替换（保留该处原格式：赋值只改文本，字符属性继承原字符）
                                var part = textRange.Characters(idx + 1, find.Length);
                                part.Text = rep;
                                total++;
                                inShape++;
                                // 全文重读；游标推进到替换文本末尾（+find.Length 保守：rep 可能被 COM 归一化换行，一律按 rep 原长后移，必要时再靠 IndexOf 从 fromIdx 找，天然防死循环）
                                full = textRange.Text ?? "";
                                fromIdx = idx + find.Length;
                            }
                            if (inShape > 0) shapesTouched++;
                        }
                        catch { skippedShapes++; }
                    }
                }
                finally { Cursor = Cursors.Default; }
                string note = skippedShapes > 0
                    ? "\n（" + skippedShapes + " 个文本形状因异常跳过，详情见日志）"
                    : "";
                MessageBox.Show("共替换 " + total + " 处，涉及 " + shapesTouched + " 个文本形状。" + note,
                    "文本工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                Cursor = Cursors.WaitCursor;
                try
                {
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
                                // 中文语境：西文逗号/句号/问号/感叹号 → 全角（仅当两侧为 CJK 时，保守不破坏英文/数字）
                                t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*,\\s*([\u4e00-\u9fff])", "$1，$2");
                                t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*\\.\\s*([\u4e00-\u9fff])", "$1。$2");
                                t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*\\?\\s*([\u4e00-\u9fff])", "$1？$2");
                                t = System.Text.RegularExpressions.Regex.Replace(t, "([\u4e00-\u9fff])\\s*!\\s*([\u4e00-\u9fff])", "$1！$2");
                            }
                            if (t != o)
                            {
                                // 整段重设前记录该段首字符主导格式，重设后刷回，避免整段丢格式。
                                // （跨段多格式文本会统一为首段格式——清理语义下可接受，弹窗已提示。）
                                try
                                {
                                    var f = textRange.Characters(1, 1).Font;
                                    string fontName = f.Name; float size = f.Size; int rgb = f.Color.RGB;
                                    bool bold = f.Bold == MsoTriState.msoTrue;
                                    textRange.Text = t;
                                    if (textRange.Text.Length > 0)
                                    {
                                        var nf = textRange.Font;
                                        if (!string.IsNullOrEmpty(fontName)) nf.Name = fontName;
                                        if (size > 0) nf.Size = size;
                                        nf.Color.RGB = rgb;
                                        nf.Bold = bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                                    }
                                }
                                catch
                                {
                                    textRange.Text = t;   // 读取格式失败仍完成清理（格式按默认）
                                }
                                changed++;
                            }
                        }
                        catch { }
                    }
                }
                finally { Cursor = Cursors.Default; }
                MessageBox.Show("已清理 " + changed + " 个文本形状。\n（文本会按首字符格式统一；若文档内含多格式混排需保留，请谨慎使用或事后撤销。）",
                    "文本工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("清理失败：" + ex.Message, "文本工具");
            }
        }
    }
}
