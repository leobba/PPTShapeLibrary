using System;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 插件内的"更新定时提醒"：到点检查一次，有新版本就弹窗提醒。
    ///
    /// 为什么单独成类、且在插件启动时就跑：
    ///  · 任务窗格是"首次打开才创建"（Startup 不建，避免无活动窗口时绑到幽灵窗口），
    ///    如果只由前端发消息来启动定时器，用户没打开过面板就永远不会被提醒；
    ///  · 所以这里在 ThisAddIn_Startup 里直接起，跟 WebView 是否加载无关。
    ///
    /// 频率可控：菜单「⋯ → 更新提醒间隔」可设 0/10/30/60 分钟（0=关闭）。
    /// 注意：UpdateService.Check() 本身有 24h 缓存，所以"每 10 分钟"只是提醒频率，不会每 10 分钟真去联网。
    /// </summary>
    public static class UpdateReminder
    {
        private static Timer _timer;
        private static int _minutes = 10;   // 默认 10 分钟
        private static DateTime _lastPrompt = DateTime.MinValue;

        public static int Minutes { get { return _minutes; } }

        /// <summary>插件启动时调用：按默认间隔（10 分钟）起定时器。</summary>
        public static void StartDefault()
        {
            SetInterval(_minutes);
        }

        /// <summary>设置提醒间隔（分钟；&lt;=0 表示关闭）。</summary>
        public static void SetInterval(int minutes)
        {
            if (_timer != null)
            {
                _timer.Stop();
                _timer.Dispose();
                _timer = null;
            }
            _minutes = minutes;
            if (minutes <= 0)
            {
                Log.Write("[update] 定时提醒已关闭");
                return;
            }
            _timer = new Timer();
            _timer.Interval = minutes * 60 * 1000;
            _timer.Tick += OnTick;
            _timer.Start();
            Log.Write("[update] 定时提醒已开启，每 " + minutes + " 分钟一次");
        }

        private static void OnTick(object sender, EventArgs e)
        {
            try
            {
                var u = UpdateService.Check();
                if (!u.Ok || !u.HasUpdate) return;

                // 同一版本不重复弹（用户点"知道了"后至少隔 1 小时再说）
                if ((DateTime.Now - _lastPrompt).TotalMinutes < 60) return;

                var notes = (u.Notes ?? "").Trim();
                if (notes.StartsWith("（")) notes = "";      // "（仓库还没有发布任何 Release）"这类不算更新内容

                var msg = "发现新版本 " + u.LatestVersion + "（当前 " + u.CurrentVersion + "）\r\n\r\n" +
                          (notes.Length > 0 ? notes + "\r\n\r\n" : "") +
                          "点「确定」前往发布页下载新版；或稍后到安装器里点「检查更新」。";

                Log.Write("[update] 提醒：新版本 " + u.LatestVersion);

                var r = MessageBox.Show(
                    msg,
                    "素材库 · 有新版本",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1);

                _lastPrompt = DateTime.Now;

                if (r == DialogResult.OK)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(UpdateService.ReleasesPageUrl) { UseShellExecute = true });
                    }
                    catch (Exception ex)
                    {
                        Log.Write("[update] 打开发布页失败：" + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                // 提醒失败绝不能影响正常使用
                Log.Write("[update] 定时检查失败：" + ex.Message);
            }
        }
    }
}
