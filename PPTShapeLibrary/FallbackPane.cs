using System;
using System.Drawing;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// WebView2 不可用时的极薄降级面板（B 阶段取代旧 WPF 面板成为兜底路径）。
    /// 不承载任何收藏功能——仅提示安装 WebView2 Runtime 并给出官方下载入口，
    /// 保证插件在 Runtime 缺失/初始化失败环境下仍可启动、不弹未处理异常。
    /// 收藏数据不受影响（存储与 UI 无关，安装 Runtime 重启 PowerPoint 即恢复 Web 面板）。
    /// </summary>
    public class FallbackPane : UserControl, ILibraryPane
    {
        public FallbackPane(string reason)
        {
            var title = new Label
            {
                Text = "新界面需要 WebView2 Runtime",
                Font = new Font("Microsoft YaHei UI", 13F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0x17, 0x1B, 0x21),
                Dock = DockStyle.Top,
                Padding = new Padding(0, 24, 0, 8),
                TextAlign = ContentAlignment.TopCenter
            };

            var desc = new Label
            {
                Text = reason + "\r\n\r\n" +
                       "本插件的收藏库界面基于 WebView2（Microsoft Edge 内核）运行。\r\n" +
                       "安装后重启 PowerPoint 即可自动恢复新界面，您的收藏不受影响。",
                Font = new Font("Microsoft YaHei UI", 9F),
                ForeColor = Color.FromArgb(0x55, 0x5F, 0x6B),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.TopCenter,
                Padding = new Padding(24, 0, 24, 0)
            };

            var btn = new Button
            {
                Text = "打开 WebView2 Runtime 下载页",
                AutoSize = true,
                Dock = DockStyle.Bottom,
                Margin = new Padding(0, 12, 0, 28),
                Font = new Font("Microsoft YaHei UI", 9F),
                BackColor = Color.FromArgb(0x00, 0x67, 0xC0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Padding = new Padding(12, 4, 12, 4)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, e) =>
            {
                try
                {
                    System.Diagnostics.Process.Start(
                        "https://developer.microsoft.com/microsoft-edge/webview2/");
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "FallbackPane open download page");
                }
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F));   // 标题
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));   // 正文
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));   // 按钮
            layout.Controls.Add(title, 0, 0);
            layout.Controls.Add(desc, 0, 1);
            layout.Controls.Add(btn, 0, 2);

            Controls.Add(layout);
            BackColor = Color.White;
            Name = "FallbackPane";
        }

        public void RefreshLibrary()
        {
            // 降级面板无可刷新内容
        }

        public void ApplySettings(Settings s)
        {
            // 降级面板无可应用设置
        }
    }
}
