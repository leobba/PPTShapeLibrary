using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;

namespace PPTShapeLibrary
{
    /// <summary>
    /// WebView2 任务窗格宿主：初始化环境、虚拟域名映射、加载 Web UI、失败降级信号。
    /// </summary>
    public class WebPaneHost : UserControl, ILibraryPane
    {
        private Microsoft.Web.WebView2.WinForms.WebView2 _webView;
        private BridgeController _bridge;
        private bool _initialized = false;
        private bool _failed = false;
        private volatile bool _shutdown = false;   // 已被 PrepareShutdown 释放（迁移/退出），此后的失败不算失败

        /// <summary>初始化失败（Runtime 缺失/导航失败）时触发，ThisAddIn 据此降级 WPF。</summary>
        public event Action<Exception> InitializationFailed;

        public WebPaneHost()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this._webView = new Microsoft.Web.WebView2.WinForms.WebView2();
            this.SuspendLayout();
            this._webView.Dock = DockStyle.Fill;
            this.Controls.Add(this._webView);
            this.Name = "WebPaneHost";
            this.ResumeLayout(false);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            InitializeAsync();
        }

        private async void InitializeAsync()
        {
            try
            {
                // Runtime 探测
                string version = CoreWebView2Environment.GetAvailableBrowserVersionString();
                Log.Write("WebView2 runtime: " + version);

                // 共享环境 + 共享用户数据目录（保持浏览器进程常驻，加速后续工具弹窗）
                // 带重试：窗格迁移到新文档窗口时旧宿主刚 Dispose，控制器创建可能偶发 E_ABORT
                await WebView2Env.EnsureCoreWebView2WithRetryAsync(_webView);

                var core = _webView.CoreWebView2;
                core.NavigationStarting += (s, e) =>
                {
                    if (!BridgeController.IsTrustedMessageSource(e.Uri)) e.Cancel = true;
                };

                // 虚拟域名映射（约束 C4：禁 file://）
                string webUiDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WebUI");
                core.SetVirtualHostNameToFolderMapping(
                    "app.pptshapes.local", webUiDir,
                    CoreWebView2HostResourceAccessKind.Allow);
                core.SetVirtualHostNameToFolderMapping(
                    "data.pptshapes.local", ShapeLibraryService.ItemsDir,
                    CoreWebView2HostResourceAccessKind.Allow);

                // Debug 构建 F12 入口
#if DEBUG
                this.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.F12) core.OpenDevToolsWindow();
                };
#endif

                // 消息桥
                _bridge = new BridgeController(core);
                _bridge.Attach();

                // 跨宿主刷新：其他宿主(工具弹窗等)改了库数据时刷新本窗格列表
                BridgeController.RefreshRequested += OnCrossRefresh;

                core.NavigationCompleted += (s, e) =>
                {
                    if (!e.IsSuccess)
                    {
                        Log.Error(new Exception("Navigation failed: " + e.WebErrorStatus), "WebView2 nav");
                        Fail(new Exception("页面加载失败: " + e.WebErrorStatus));
                        return;
                    }
                    _initialized = true;
                    Log.Write("Pane ready (WebView2)");
                };

                core.WebMessageReceived += (s, e) =>
                {
                    try
                    {
                        if (!BridgeController.IsTrustedMessageSource(e.Source)) return;
                        _bridge?.HandleWebMessage(e.TryGetWebMessageAsString());
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "WebMessageReceived");
                    }
                };

                await _webView.CoreWebView2.ExecuteScriptAsync("document.title='ready'");
                // 时间戳参数防 WebView2 缓存旧页面资源
                var stamp = DateTime.Now.Ticks.ToString();
                _webView.Source = new Uri("https://app.pptshapes.local/index.html?v=" + stamp);
            }
            catch (Exception ex)
            {
                Fail(ex);
            }
        }

        private void Fail(Exception ex)
        {
            // 迁移到别的文档窗口时，旧宿主会被 PrepareShutdown 释放；此时初始化失败属预期，
            // 不能当成"WebView2 不可用"去降级成提示页（历史现象：面板突然变成安装引导页）。
            if (_failed || _shutdown) return;
            _failed = true;
            Log.Error(ex, "WebPaneHost init failed");
            InitializationFailed?.Invoke(ex);
        }

        public bool IsFailed
        {
            get { return _failed; }
        }

        /// <summary>WebView2 首次导航成功（前端可交互）后为 true。供迁移重建后等待显示。</summary>
        public bool IsReady
        {
            get { return _initialized; }
        }

        /// <summary>跨宿主刷新回调（其他宿主改数据后触发本窗格刷新）。</summary>
        private void OnCrossRefresh()
        {
            try
            {
                if (_initialized && _bridge != null) _bridge.PushRefresh();
            }
            catch (Exception ex) { Log.Error(ex, "WebPaneHost.OnCrossRefresh"); }
        }

        /// <summary>关闭前释放 WebView2 宿主（停止导航、释放底层进程与 userData 文件锁）。</summary>
        public void PrepareShutdown()
        {
            try
            {
                _shutdown = true;
                BridgeController.RefreshRequested -= OnCrossRefresh;
                if (_bridge != null) _bridge.MarkClosed();   // 之后不再向已释放的 CoreWebView2 发消息
                if (_webView != null)
                {
                    try { _webView.Dispose(); } catch { }
                    _webView = null;
                }
                _bridge = null;
                _initialized = false;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "WebPaneHost.PrepareShutdown");
            }
        }

        // ---------------- ILibraryPane ----------------

        public void RefreshLibrary()
        {
            try
            {
                if (_initialized && _bridge != null)
                {
                    _bridge.PushRefresh();
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "WebPaneHost.RefreshLibrary");
            }
        }

        public void ApplySettings(Settings s)
        {
            try
            {
                if (_initialized && _bridge != null)
                {
                    _bridge.PushSettingsChanged(s);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "WebPaneHost.ApplySettings");
            }
        }
    }
}
