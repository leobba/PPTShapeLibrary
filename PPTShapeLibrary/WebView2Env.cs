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

        /// <summary>获取（并缓存）共享环境。首次创建后复用。</summary>
        public static async Task<CoreWebView2Environment> GetAsync()
        {
            if (_env != null) return _env;
            System.IO.Directory.CreateDirectory(UserDataFolder);
            var env = await CoreWebView2Environment.CreateAsync(null, UserDataFolder, new CoreWebView2EnvironmentOptions());
            lock (Sync)
            {
                if (_env == null) _env = env;
            }
            return _env;
        }
    }
}
