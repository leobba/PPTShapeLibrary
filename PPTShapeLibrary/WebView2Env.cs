using System;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;

namespace PPTShapeLibrary
{
    /// <summary>
    /// WebView2 环境共享：任务窗格与工具弹窗共用同一 CoreWebView2Environment 与用户数据目录。
    /// 好处：
    /// 1. 同一浏览器进程（任务窗格常驻时其 WebView2 实例保持浏览器进程存活），
    ///    之后打开工具弹窗不再冷启动，明显减少"打开慢/卡顿"；
    /// 2. 环境对象只创建一次，减少重复初始化开销。
    /// 说明：弹窗与面板渲染的是同一 Web UI（app.pptshapes.local 同源），共享 profile 无冲突。
    /// </summary>
    public static class WebView2Env
    {
        private static CoreWebView2Environment _env;
        private static Task<CoreWebView2Environment> _creating;
        private static readonly object Sync = new object();

        public static string UserDataFolder
        {
            get
            {
                return System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "PPTShapeLibrary", "webview2");
            }
        }

        /// <summary>获取（并缓存）共享环境。首次创建后复用；并发调用只创建一次
        /// （此前任务窗格与工具弹窗同时初始化会各自 CreateAsync 一个环境，多余那个被丢弃但会多起浏览器进程）。</summary>
        public static Task<CoreWebView2Environment> GetAsync()
        {
            lock (Sync)
            {
                if (_env != null) return Task.FromResult(_env);
                if (_creating != null) return _creating;
                _creating = CreateOnceAsync();
                return _creating;
            }
        }

        private static async Task<CoreWebView2Environment> CreateOnceAsync()
        {
            try
            {
                System.IO.Directory.CreateDirectory(UserDataFolder);
                var env = await CoreWebView2Environment.CreateAsync(null, UserDataFolder, new CoreWebView2EnvironmentOptions());
                lock (Sync)
                {
                    if (_env == null) _env = env;
                    return _env;
                }
            }
            catch
            {
                lock (Sync) { _creating = null; }   // 失败允许下次重试
                throw;
            }
        }

        /// <summary>
        /// 带重试的控制器初始化。实测坑：任务窗格在"打开为新演示文稿"等场景下迁移到新文档窗口时，
        /// 旧宿主刚被 Dispose、新的 WebView2 立刻初始化，会偶发
        /// `COMException 0x80004004 (E_ABORT)`（浏览器进程正在重启），
        /// 结果是面板直接降级成"请安装 WebView2"提示页，只能重启 PowerPoint 才恢复。
        /// 这里退避重试若干次；重试期间浏览器进程通常已经就绪。
        /// </summary>
        public static async Task EnsureCoreWebView2WithRetryAsync(
            Microsoft.Web.WebView2.WinForms.WebView2 view, int attempts = 3)
        {
            var env = await GetAsync();
            Exception last = null;
            for (int i = 1; i <= attempts; i++)
            {
                if (view.IsDisposed)   // 期间窗格又被迁移重建：本控件的 CoreWebView2 已被释放，重试无意义
                {
                    Log.Write("WebView2 ensure aborted: control disposed (attempt " + i + "/" + attempts + ")");
                    throw last ?? new ObjectDisposedException("WebView2");
                }
                try
                {
                    await view.EnsureCoreWebView2Async(env);
                    if (i > 1) Log.Write("WebView2 ensure succeeded on attempt " + i + "/" + attempts);
                    return;
                }
                catch (Exception ex)
                {
                    last = ex;
                    Log.Write("WebView2 ensure attempt " + i + "/" + attempts + " failed: " + ex.Message);
                    if (i == attempts) break;
                    await Task.Delay(250 * i);
                }
            }
            throw last ?? new InvalidOperationException("WebView2 初始化失败");
        }
    }
}
