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

        /// <summary>初始化失败（Runtime 缺失/导航失败）时触发，ThisAddIn 据此降级 WPF。</summary>
        public event Action<Exception> InitializationFailed;

        public WebPaneHost()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                _failed = true;
                Log.Error(ex, "WebPaneHost ctor");
                InitializationFailed?.Invoke(ex);
            }
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
                var env = await WebView2Env.GetAsync();
                await _webView.EnsureCoreWebView2Async(env);

                var core = _webView.CoreWebView2;

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
            if (_failed) return;
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
                BridgeController.RefreshRequested -= OnCrossRefresh;
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

        /// <summary>切换前端视图（library/tools/settings）——供 Ribbon 直达按钮使用。</summary>
        public void SwitchView(string view)
        {
            try
            {
                if (_initialized && _bridge != null)
                {
                    _bridge.PushSwitchView(view);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "WebPaneHost.SwitchView");
            }
        }
    }
}
