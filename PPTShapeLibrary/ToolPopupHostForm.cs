using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 独立工具/设置弹窗宿主：内嵌 WebView2，加载与任务窗格同一 WebUI，
    /// 通过 URL 查询参数 ?view=toolX|settings 让前端进入"单工具模式"，
    /// 从而获得与素材库面板完全一致的 Web UI 观感（而非 WinForms 原生控件）。
    /// 无系统标题栏，标题由 Web 页内 popup-head 呈现（带图标 + 关闭钮）。
    /// </summary>
    public class ToolPopupHostForm : Form
    {
        private readonly string _viewKey;
        private WebView2 _webView;
        private bool _failed = false;

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        /// <summary>网页 popup-head 按住拖动时调用（无边框窗口移动）。</summary>
        private void BeginDragMove()
        {
            try
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Popup drag");
            }
        }

        public ToolPopupHostForm(string viewKey)
        {
            _viewKey = viewKey ?? "tool1";
            // 无边框：标题由 Web 页内 popup-head 呈现（带图标+关闭钮），避免 Windows 标题与页内标题双标题
            Text = ViewTitle(_viewKey);
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new System.Drawing.Size(440, 560);
            MinimumSize = new System.Drawing.Size(360, 400);
            ShowInTaskbar = false;
            BackColor = System.Drawing.Color.White;

            _webView = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_webView);
            // Esc 直接关闭（无边框弹窗，页内 ✕ 与 Esc 双途径，避免"关不掉"）
            KeyPreview = true;
            KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    e.Handled = true;
                    Close();
                }
            };
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
                case "tool9": return "页脚页码";
                case "tool10": return "导出 PDF / 大纲";
                case "tool11": return "图表美化";
                case "tool12": return "成套模板";
                default: return "工具";
            }
        }

        private async void InitAsync()
        {
            try
            {
                string version = CoreWebView2Environment.GetAvailableBrowserVersionString();
                Log.Write("Popup WebView2 runtime: " + version);

                // 与任务窗格共享环境与用户数据目录：浏览器进程常驻，弹窗打开不冷启动（快）
                var env = await WebView2Env.GetAsync();
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
                    try
                    {
                        string raw = e.TryGetWebMessageAsString();
                        // 宿主级消息：精确解析 type 后关窗/拖动，避免 Contains 误匹配业务消息
                        if (!string.IsNullOrEmpty(raw))
                        {
                            string type = "";
                            try
                            {
                                var msg = new System.Web.Script.Serialization.JavaScriptSerializer()
                                    .DeserializeObject(raw) as System.Collections.Generic.Dictionary<string, object>;
                                if (msg != null && msg.ContainsKey("type")) type = msg["type"]?.ToString() ?? "";
                            }
                            catch { /* 非 JSON 按未知处理，交给桥 */ }
                            if (type == "popupClose") { Close(); return; }
                            if (type == "popupDrag") { BeginDragMove(); return; }
                        }
                        bridge.HandleWebMessage(raw);
                    }
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
