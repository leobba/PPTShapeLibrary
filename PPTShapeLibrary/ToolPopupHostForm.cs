using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 独立工具/设置弹窗宿主：内嵌 WebView2，加载与任务窗格同一 WebUI，
    /// 通过 URL 查询参数 ?view=toolX|settings 让前端进入"单工具模式"，
    /// 从而获得与素材库面板完全一致的 Web UI 观感（而非 WinForms 原生控件）。
    /// </summary>
    public class ToolPopupHostForm : Form
    {
        private readonly string _viewKey;
        private WebView2 _webView;
        private bool _failed = false;

        public ToolPopupHostForm(string viewKey)
        {
            _viewKey = viewKey ?? "tool1";
            Text = ViewTitle(_viewKey);
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new System.Drawing.Size(440, 560);
            MinimumSize = new System.Drawing.Size(360, 420);
            ShowInTaskbar = false;

            _webView = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_webView);
            Load += (s, e) => InitAsync();
        }

        private static string ViewTitle(string key)
        {
            switch (key)
            {
                case "settings": return "设置";
                case "tool1": return "文字样式";
                case "tool2": return "一键换字体";
                case "tool3": return "查找 / 替换";
                case "tool4": return "文字清理";
                case "tool5": return "导出图片";
                case "tool6": return "清理瘦身";
                case "tool7": return "统一尺寸";
                case "tool8": return "图片网格";
                default: return "工具";
            }
        }

        private async void InitAsync()
        {
            try
            {
                string version = CoreWebView2Environment.GetAvailableBrowserVersionString();
                Log.Write("Popup WebView2 runtime: " + version);

                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "PPTShapeLibrary", "webview2");
                Directory.CreateDirectory(userDataFolder);

                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder, new CoreWebView2EnvironmentOptions());
                await _webView.EnsureCoreWebView2Async(env);

                var core = _webView.CoreWebView2;
                string webUiDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WebUI");
                core.SetVirtualHostNameToFolderMapping("app.pptshapes.local", webUiDir,
                    CoreWebView2HostResourceAccessKind.Allow);
                core.SetVirtualHostNameToFolderMapping("data.pptshapes.local", ShapeLibraryService.ItemsDir,
                    CoreWebView2HostResourceAccessKind.Allow);

                // 弹窗与任务窗格共享消息桥：同一个桥控制器负责 (view 由 URL 参数决定)
                var bridge = new BridgeController(core);
                bridge.Attach();
                core.WebMessageReceived += (s, e) =>
                {
                    try { bridge.HandleWebMessage(e.TryGetWebMessageAsString()); }
                    catch (Exception ex) { Log.Error(ex, "Popup WebMessageReceived"); }
                };

                core.NavigationCompleted += (s, e) =>
                {
                    if (!e.IsSuccess)
                    {
                        Log.Error(new Exception("Popup nav failed: " + e.WebErrorStatus), "Popup WebView2");
                        Fail();
                    }
                };

                var stamp = DateTime.Now.Ticks.ToString();
                _webView.Source = new Uri("https://app.pptshapes.local/index.html?view=" + _viewKey + "&v=" + stamp);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Popup WebView2 init");
                Fail();
            }
        }

        private void Fail()
        {
            if (_failed) return;
            _failed = true;
            try
            {
                MessageBox.Show("工具窗口初始化失败：请确认本机已安装 WebView2 Runtime。\n" +
                                "https://developer.microsoft.com/microsoft-edge/webview2/", "素材库工具");
            }
            catch { }
        }
    }
}
