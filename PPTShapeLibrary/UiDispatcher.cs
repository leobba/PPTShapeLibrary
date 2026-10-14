using System;
using System.Threading;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 把回调投递到**插件 UI 线程**（PowerPoint 主线程）执行。
    ///
    /// 为什么需要单独一个类：更新检查要联网（几秒），绝不能跑在 UI 线程上
    /// （会把 PowerPoint 启动/界面卡住）；但检查完要弹窗、又要回到 UI 线程。
    ///
    /// 实现：在 UI 线程上创建一个隐藏 Control 并强制建句柄，之后任意线程都能
    /// 用它的 BeginInvoke 把委托投回 UI 线程 —— 这是 WinForms 互操作里最可靠的
    /// 跨线程投递方式（不依赖 SynchronizationContext 是否已安装）。
    /// </summary>
    internal static class UiDispatcher
    {
        private static Control _invoker;

        /// <summary>在**插件启动时（UI 线程）**调用一次。</summary>
        public static void Init()
        {
            if (_invoker != null) return;
            try
            {
                var c = new Control();
                c.CreateControl();          // 在 UI 线程建句柄，BeginInvoke 才能工作
                _invoker = c;
            }
            catch (Exception ex)
            {
                Log.Write("[ui] UiDispatcher 初始化失败（将退化为直接执行）：" + ex.Message);
            }
        }

        public static bool IsAvailable
        {
            get { return _invoker != null && _invoker.IsHandleCreated; }
        }

        /// <summary>投递到 UI 线程；不可用时就在当前线程直接执行（至少不会丢）。</summary>
        public static void Post(Action action)
        {
            if (action == null) return;
            try
            {
                var c = _invoker;
                if (c != null && c.IsHandleCreated)
                {
                    c.BeginInvoke(action);
                    return;
                }
            }
            catch (Exception ex)
            {
                Log.Write("[ui] 投递失败，改为直接执行：" + ex.Message);
            }
            try { action(); } catch (Exception ex) { Log.Write("[ui] 直接执行异常：" + ex.Message); }
        }

        /// <summary>在后台线程跑一段可能有网络/耗时操作的代码，完成后回 UI 线程执行 done。</summary>
        public static void RunBackground(Action work, Action done)
        {
            var t = new Thread(() =>
            {
                try { work(); }
                catch (Exception ex) { Log.Write("[ui] 后台任务异常：" + ex.Message); }
                if (done != null) Post(done);
            });
            t.IsBackground = true;
            t.Start();
        }

        /// <summary>延后 delayMs 再在后台线程跑（用于启动后稍等一会儿再联网）。</summary>
        public static void RunBackgroundDelayed(int delayMs, Action work, Action done)
        {
            var t = new Thread(() =>
            {
                try { Thread.Sleep(delayMs); } catch { }
                try { work(); }
                catch (Exception ex) { Log.Write("[ui] 后台任务异常：" + ex.Message); }
                if (done != null) Post(done);
            });
            t.IsBackground = true;
            t.Start();
        }
    }
}
