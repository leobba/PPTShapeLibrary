using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTShapeLibrary
{
    public partial class ThisAddIn
    {
        private ILibraryPane _paneControl;
        private UserControl _paneHost;
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                Log.Write("Startup begin");
                CreatePane();
                Log.Write("TaskPane created");
                _taskPane = this.CustomTaskPanes.Add(_paneHost, "素材库");
                _taskPane.Width = ShapeLibraryService.LoadSettings().PaneWidth > 0
                    ? ShapeLibraryService.LoadSettings().PaneWidth : 300;
                Log.Write("TaskPane added");

                // 内置素材延迟安装（避免启动期复杂 COM 操作）
                var timer = new System.Windows.Forms.Timer { Interval = 8000 };
                timer.Tick += (s2, e2) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    try { ShapeExporter.InstallBuiltinLibrary(); Log.Write("Builtin library installed"); }
                    catch (Exception ex) { Log.Error(ex, "InstallBuiltinLibrary"); }
                };
                timer.Start();
                Log.Write("Startup done");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ThisAddIn_Startup");
                MessageBox.Show("插件启动失败：" + ex.Message, "素材库");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        /// <summary>宿主选择决策树：WebView2 Runtime 检测 + 失败降级到极薄提示页。</summary>
        private void CreatePane()
        {
            // uiBackend="wpf" 已废弃：旧 WPF 面板随 B 阶段完成被移除。
            // 保留该值解析仅为向后兼容——等价于走降级提示页。
            var settings = ShapeLibraryService.LoadSettings();
            if (settings.UiBackend == "wpf")
            {
                Log.Write("uiBackend=wpf is deprecated, using fallback pane");
                UseFallbackPane("settings.json 中的 uiBackend=\"wpf\" 已不再使用。");
                return;
            }

            try
            {
                string ver = Microsoft.Web.WebView2.Core.CoreWebView2Environment.GetAvailableBrowserVersionString();
                Log.Write("WebView2 runtime available: " + ver);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "WebView2 runtime missing");
                UseFallbackPane("检测到本机缺少 WebView2 Runtime。");
                return;
            }

            // Web 面板（失败自动降级到提示页）
            var webHost = new WebPaneHost();
            webHost.InitializationFailed += ex =>
            {
                try
                {
                    Log.Error(ex, "WebPane init failed -> fallback pane");
                    if (_taskPane != null)
                    {
                        MessageBox.Show("新界面初始化失败，将显示降级提示页。", "素材库",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        SwitchToFallbackPane("新界面初始化失败：" + ex.Message);
                    }
                }
                catch { }
            };
            _paneControl = webHost;
            _paneHost = webHost;
        }

        /// <summary>极薄降级提示页（WebView2 缺失/初始化失败时使用）。</summary>
        private void UseFallbackPane(string reason)
        {
            var fb = new FallbackPane(reason);
            _paneControl = fb;
            _paneHost = fb;
        }

        private void SwitchToFallbackPane(string reason)
        {
            try
            {
                bool visible = _taskPane != null && _taskPane.Visible;
                int width = _taskPane != null ? _taskPane.Width : 300;
                this.CustomTaskPanes.Remove(_taskPane);

                UseFallbackPane(reason);
                _taskPane = this.CustomTaskPanes.Add(_paneHost, "素材库");
                _taskPane.Width = width;
                _taskPane.Visible = visible;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "SwitchToFallbackPane");
            }
        }

        /// <summary>注册 VSTO 原生 Ribbon（自定义「形状库」选项卡）。</summary>
        protected override Microsoft.Office.Tools.Ribbon.IRibbonExtension[] CreateRibbonObjects()
        {
            return new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { new Ribbon1() };
        }

        /// <summary>导出全部收藏为备份文件。</summary>
        public void DoBackup()
        {
            var dlg = new SaveFileDialog
            {
                Title = "导出收藏备份",
                Filter = "收藏备份 (*.zip)|*.zip",
                FileName = "素材库备份_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".zip"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            if (ShapeLibraryService.BackupTo(dlg.FileName))
            {
                MessageBox.Show("备份完成：" + dlg.FileName, "素材库");
            }
            else
            {
                MessageBox.Show("备份失败，请检查目标位置是否可写。", "素材库", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>从备份文件恢复收藏库。</summary>
        public void DoRestore()
        {
            var dlg = new OpenFileDialog
            {
                Title = "导入收藏备份",
                Filter = "收藏备份 (*.zip)|*.zip"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            if (MessageBox.Show("导入将覆盖当前所有收藏，确定继续？", "素材库",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            if (ShapeLibraryService.RestoreFrom(dlg.FileName))
            {
                MessageBox.Show("导入完成。", "素材库");
                RefreshLibrary();
            }
            else
            {
                MessageBox.Show("导入失败，备份文件可能已损坏。", "素材库", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>导出收藏库为 PPTX 分享文件。</summary>
        public void DoExportPptx()
        {
            var dlg = new SaveFileDialog
            {
                Title = "导出收藏为 PPTX",
                Filter = "PowerPoint 演示文稿 (*.pptx)|*.pptx",
                FileName = "素材库_" + DateTime.Now.ToString("yyyyMMdd-HHmm") + ".pptx"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            if (ShapeExporter.ExportLibraryToPptx(dlg.FileName))
            {
                MessageBox.Show("导出完成：" + dlg.FileName, "素材库");
            }
        }

        /// <summary>打开设置对话框并应用。</summary>
        public void DoSettings()
        {
            var current = ShapeLibraryService.LoadSettings();
            using (var frm = new SettingsForm(current))
            {
                if (frm.ShowDialog() == DialogResult.OK && frm.Result != null)
                {
                    ShapeLibraryService.SaveSettings(frm.Result);
                    if (_taskPane != null)
                    {
                        _taskPane.Width = frm.Result.PaneWidth;
                        _paneControl.ApplySettings(frm.Result);
                    }
                }
            }
        }

        /// <summary>显示/隐藏收藏任务窗格。</summary>
        public void ToggleTaskPane()
        {
            if (_taskPane == null) return;
            _taskPane.Visible = !_taskPane.Visible;
            if (_taskPane.Visible)
            {
                _paneControl.RefreshLibrary();
            }
        }

        public void ShowTaskPane()
        {
            if (_taskPane == null) return;
            _taskPane.Visible = true;
            _paneControl.RefreshLibrary();
        }

        public void RefreshLibrary()
        {
            if (_paneControl != null) _paneControl.RefreshLibrary();
        }

        #region VSTO generated code

        /// <summary>
        /// 设计器支持所需的方法——不要使用代码编辑器修改
        /// 此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
