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
        private static string _lastPromptedVersion = null;

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
            // 定时器跑在 UI 线程上，检查要联网 → 丢到后台，避免卡住 PowerPoint
            CheckAsync();
        }

        /// <summary>启动时调用：后台查一次，有新版就弹窗（无新版什么都不做）。</summary>
        public static void CheckOnStartupAsync()
        {
            // 延后 4 秒再查：启动瞬间 PowerPoint 自己也在忙，且测速（后台）还没出结果
            UiDispatcher.RunBackgroundDelayed(4000, () => CheckCore(prompt: true), null);
        }

        /// <summary>后台检查 + 回 UI 线程弹窗（定时器用）。</summary>
        private static void CheckAsync()
        {
            UiDispatcher.RunBackground(() => CheckCore(prompt: true), null);
        }

        /// <summary>
        /// 检查一次更新；<paramref name="prompt"/> = true 时，**有新版才弹窗**
        /// （无新版什么都不做 —— 不要把"已是最新"这种信息弹出来打扰人）。
        ///
        /// 网络请求在后台线程做；弹窗回到 UI 线程做。
        /// </summary>
        public static void CheckCore(bool prompt)
        {
            UpdateResult u;
            try { u = UpdateService.Check(); }
            catch (Exception ex) { Log.Write("[update] 检查更新失败：" + ex.Message); return; }

            if (!u.Ok || !u.HasUpdate) return;      // 没更新 → 静默，绝不弹窗
            Log.Write("[update] 发现新版本 " + u.LatestVersion + "（当前 " + u.CurrentVersion + "）");
            if (!prompt) return;

            UiDispatcher.Post(() => PromptAndAct(u));
        }

        private static void PromptAndAct(UpdateResult u)
        {
            try
            {
                // 同一版本不重复弹（用户处理过之后至少隔 1 小时再说）
                if (_lastPromptedVersion == u.LatestVersion &&
                    (DateTime.Now - _lastPrompt).TotalMinutes < 60)
                {
                    Log.Write("[update] 同版本 1 小时内已提醒过，跳过");
                    return;
                }

                var notes = (u.Notes ?? "").Trim();
                if (notes.StartsWith("（")) notes = "";      // 占位说明不算更新内容
                if (notes.Length > 280) notes = notes.Substring(0, 280) + "…";

                var msg = "发现新版本 " + u.LatestVersion + "（当前 " + u.CurrentVersion + "）\r\n\r\n" +
                          (notes.Length > 0 ? notes + "\r\n\r\n" : "") +
                          "点「立即更新」可直接下载安装（会显示下载进度/速度，并检查文稿是否已保存）。";

                var r = MessageBox.Show(
                    msg,
                    "素材库 · 有新版本",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1);

                _lastPrompt = DateTime.Now;
                _lastPromptedVersion = u.LatestVersion;

                if (r == DialogResult.OK)
                {
                    // 直接用刚拿到的结果开窗（里面不再联网检查，避免又等一轮）
                    try { UpdateDialog.ShowWith(u); }
                    catch (Exception ex) { Log.Write("[update] 打开更新弹窗失败：" + ex.Message); }
                }
            }
            catch (Exception ex)
            {
                // 提醒失败绝不能影响正常使用
                Log.Write("[update] 弹窗提醒失败：" + ex.Message);
            }
        }
    }
}
