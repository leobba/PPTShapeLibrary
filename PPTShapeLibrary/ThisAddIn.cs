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
        private System.Windows.Forms.Timer _refreshDebounce;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                Log.Write("Startup begin");
                HookAppEvents();
                BridgeController.SettingsSaved += OnSettingsSaved;
                CreatePane();
                Log.Write("TaskPane created");
                // 注意：不在 Startup 立即 Add 任务窗格——此时可能无活动文档窗口，
                // Add 会绑定"幽灵窗口"导致窗格 Visible=true 却永不显示（用户实测问题）。
                // 改为首次打开素材库（必有活动窗口）时再 Add（见 EnsureTaskPaneForActiveWindow）。
                Log.Write("TaskPane deferred to first open");

                // 内置素材延迟安装（避免启动期复杂 COM 操作）
                var timer = new System.Windows.Forms.Timer { Interval = 8000 };
                timer.Tick += (s2, e2) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    try { ShapeExporter.InstallBuiltinLibrary(); Log.Write("Builtin library installed"); }
                    catch (Exception ex) { Log.Error(ex, "InstallBuiltinLibrary"); }
                    try { ShapeExporter.InstallAcademicTemplates(); Log.Write("Academic templates installed"); }
                    catch (Exception ex) { Log.Error(ex, "InstallAcademicTemplates"); }
                    try { ShapeExporter.InstallLibraryContent(); Log.Write("Library content installed"); }
                    catch (Exception ex) { Log.Error(ex, "InstallLibraryContent"); }
                    try { ShapeExporter.GenerateSamplePictures(); Log.Write("Sample pictures installed"); }
                    catch (Exception ex) { Log.Error(ex, "GenerateSamplePictures"); }
                    try { ShapeExporter.EnsureGroupAssignments(); }
                    catch (Exception ex) { Log.Error(ex, "EnsureGroupAssignments"); }
                    try { int orphans = ShapeLibraryService.CollectOrphanDirs(); if (orphans > 0) Log.Write("Orphan dirs cleaned: " + orphans); }
                    catch (Exception ex) { Log.Error(ex, "CollectOrphanDirs"); }
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
            try
            {
                UnhookAppEvents();
                BridgeController.SettingsSaved -= OnSettingsSaved;
                if (_refreshDebounce != null) { _refreshDebounce.Stop(); _refreshDebounce.Dispose(); _refreshDebounce = null; }
                // 释放宿主（WebView2 底层进程与 userData 锁随之释放，避免残留实例占用 profile）
                var host = _paneControl as WebPaneHost;
                try
                {
                    if (host != null)
                    {
                        host.PrepareShutdown();
                        _paneControl = null;
                        _paneHost = null;
                    }
                }
                catch { }
            }
            catch { }
        }

        /// <summary>设置保存后即时应用需要宿主配合的项（如任务窗格宽度）。</summary>
        private void OnSettingsSaved(Settings s)
        {
            try
            {
                if (_taskPane != null && s.PaneWidth >= 200)
                {
                    _taskPane.Width = s.PaneWidth;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "OnSettingsSaved");
            }
        }

        /// <summary>订阅演示文稿生命周期事件：打开/关闭/新建/激活/停用任一演示文稿时防抖刷新素材库。幂等：重复调用不重复订阅。</summary>
        private bool _appEventsHooked = false;
        private void HookAppEvents()
        {
            try
            {
                if (_appEventsHooked) return;
                var app = this.Application;
                if (app == null) return;
                // 经 ApplicationEvents_Event 事件接口订阅
                var events = (PowerPoint.EApplication_Event)app;
                events.NewPresentation += OnNewPres;
                events.PresentationOpen += OnPresOpen;
                events.PresentationClose += OnPresClose;
                events.WindowActivate += OnWindowActivate;
                _appEventsHooked = true;
                Log.Write("App events hooked");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "HookAppEvents");
            }
        }

        private void UnhookAppEvents()
        {
            try
            {
                if (!_appEventsHooked) return;
                var app = this.Application;
                if (app == null) return;
                var events = (PowerPoint.EApplication_Event)app;
                events.NewPresentation -= OnNewPres;
                events.PresentationOpen -= OnPresOpen;
                events.PresentationClose -= OnPresClose;
                events.WindowActivate -= OnWindowActivate;
                _appEventsHooked = false;
                Log.Write("App events unhooked");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "UnhookAppEvents");
            }
        }

        private void OnNewPres(PowerPoint.Presentation pres) { ScheduleRefresh(); }
        private void OnPresOpen(PowerPoint.Presentation pres) { ScheduleRefresh(); }
        private void OnPresClose(PowerPoint.Presentation pres) { ScheduleRefresh(); }
        private void OnWindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow win)
        {
            ScheduleRefresh();
            try
            {
                string cap = "?";
                try { cap = win != null ? win.Caption : "null"; } catch { }
                bool sameWin = false;
                try { sameWin = _taskPane != null && ReferenceEquals(_taskPane.Window, win); } catch { }
                Log.Write("[pane] WindowActivate caption=[" + cap + "] sameWindow=" + sameWin + " paneVisible=" + (_taskPane != null && _taskPane.Visible));
                if (_taskPane == null) return;
                if (!_taskPane.Visible) return;
                if (sameWin)
                {
                    // 同窗口多文档：PowerPoint 在文档切换后可能不重绘任务窗格（Visible=true 但不显示）。
                    // 强制 false→true 触发 Office 重挂（轻量，不重建 WebView2）。
                    _taskPane.Visible = false;
                    _taskPane.Visible = true;
                    Log.Write("[pane] WindowActivate force-repaint done");
                }
                else
                {
                    EnsurePaneForWindow(win);
                }
            }
            catch (Exception ex) { Log.Error(ex, "OnWindowActivate"); }
        }

        private void ScheduleRefresh()
        {
            try
            {
                if (_refreshDebounce == null)
                {
                    _refreshDebounce = new System.Windows.Forms.Timer { Interval = 600 };
                    _refreshDebounce.Tick += (s, e) =>
                    {
                        _refreshDebounce.Stop();
                        Log.Write("Doc event -> refresh library");
                        if (_paneControl != null) _paneControl.RefreshLibrary();
                    };
                }
                _refreshDebounce.Stop();
                _refreshDebounce.Start();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ScheduleRefresh");
            }
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
                if (_taskPane != null) this.CustomTaskPanes.Remove(_taskPane);

                UseFallbackPane(reason);
                PowerPoint.DocumentWindow win = null;
                try { win = Application.ActiveWindow; } catch { }
                _taskPane = win != null
                    ? this.CustomTaskPanes.Add(_paneHost, "素材库", win)
                    : this.CustomTaskPanes.Add(_paneHost, "素材库");
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

        /// <summary>
        /// 确保素材库任务窗格出现在当前活动窗口。
        /// PowerPoint 独立窗口各自承载窗格：若窗格绑定窗口 ≠ 活动窗口，则把窗格迁移到活动窗口
        /// （移除旧窗格并为当前窗口重建，VSTO 的 CustomTaskPane.Window 只读，无法原地改绑）。
        /// </summary>
        private void EnsurePaneOnActiveWindow()
        {
            EnsurePaneForWindow(null);
        }

        /// <summary>首次打开时创建任务窗格（Startup 不建，避免无活动窗口时绑定幽灵窗口导致永不显示）。</summary>
        private void EnsureTaskPaneCreated()
        {
            try
            {
                if (_taskPane != null) return;
                if (_paneHost == null) return;
                PowerPoint.DocumentWindow win = null;
                try { win = Application.ActiveWindow; } catch { }
                if (win == null) return;   // 仍无活动窗口：等下次调用

                _taskPane = this.CustomTaskPanes.Add(_paneHost, "素材库", win);
                int w = ShapeLibraryService.LoadSettings().PaneWidth;
                _taskPane.Width = w > 0 ? w : 300;
                Log.Write("Task pane created lazily for active window");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "EnsureTaskPaneCreated");
            }
        }

        private void EnsurePaneForWindow(PowerPoint.DocumentWindow win)
        {
            try
            {
                EnsureTaskPaneCreated();   // 懒建兜底
                if (_taskPane == null) return;
                if (win == null)
                {
                    try { win = Application.ActiveWindow; } catch { }
                }
                if (win == null) return;
                object current = null;
                try { current = _taskPane.Window; } catch { }
                if (ReferenceEquals(current, win)) return;   // 已绑定当前窗口

                Log.Write("Migrating task pane to window (was bound to another)");
                bool wasVisible = _taskPane.Visible;
                int width = _taskPane.Width;
                this.CustomTaskPanes.Remove(_taskPane);
                // 释放旧宿主，避免 WebView2 残留
                var oldHost = _paneControl as WebPaneHost;
                try { oldHost?.PrepareShutdown(); } catch { }

                // 为当前窗口重建（WebView2 重新初始化）
                CreatePane();
                _taskPane = this.CustomTaskPanes.Add(_paneHost, "素材库", win);
                _taskPane.Width = width;
                // 关键：必须立即 Visible=true 触发内容控件 Load → WebPaneHost.InitializeAsync 启动；
                // 若设为 false，UserControl.Load 不触发，WebView2 永不初始化、窗格永不显示（此前 bug）。
                _taskPane.Visible = true;
                Log.Write("Task pane rebuilt & shown for new window (wasVisible=" + wasVisible + ")");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "EnsurePaneForWindow");
            }
        }

        /// <summary>诊断：记录窗格绑定窗口 vs 当前活动窗口，用于多窗口问题定位。</summary>
        private void LogPaneState(string who)
        {
            try
            {
                string active = "?";
                try { active = Application.ActiveWindow != null ? Application.ActiveWindow.Caption : "null"; } catch { }
                string bound = "?";
                try { bound = _taskPane != null && _taskPane.Window != null ? "bound-window" : "none"; } catch { }
                string vis = _taskPane != null ? _taskPane.Visible.ToString() : "no-pane";
                Log.Write("[pane] " + who + " active=[" + active + "] bound=" + bound + " visible=" + vis);
            }
            catch { }
        }

        /// <summary>显示/隐藏收藏任务窗格。</summary>
        public void ToggleTaskPane()
        {
            LogPaneState("Toggle enter");
            EnsureTaskPaneCreated();   // 首次/懒创建（Startup 未建）
            if (_taskPane == null) return;
            if (!_taskPane.Visible) EnsurePaneOnActiveWindow();   // 要显示时确保在当前窗口
            _taskPane.Visible = !_taskPane.Visible;
            if (_taskPane.Visible)
            {
                _paneControl.RefreshLibrary();
            }
            LogPaneState("Toggle exit");
        }

        public void ShowTaskPane()
        {
            LogPaneState("Show enter");
            EnsureTaskPaneCreated();
            if (_taskPane == null) return;
            EnsurePaneOnActiveWindow();
            _taskPane.Visible = true;
            _paneControl.RefreshLibrary();
            LogPaneState("Show exit");
        }

        public void RefreshLibrary()
        {
            if (_paneControl != null) _paneControl.RefreshLibrary();
        }

        /// <summary>把 Web 面板切到"工具"视图（Ribbon「小工具」→ 文字样式首项）。</summary>
        public void ShowToolsView()
        {
            ShowTaskPane();
            if (_paneControl is WebPaneHost web) web.SwitchView("tool1");
        }

        /// <summary>把 Web 面板切到指定工具视图（tool1..tool8 / settings / library）。</summary>
        public void ShowToolView(string view)
        {
            ShowTaskPane();
            if (_paneControl is WebPaneHost web) web.SwitchView(view);
        }

        /// <summary>把 Web 面板切到"设置"视图（Ribbon「设置」）。</summary>
        public void ShowSettingsView()
        {
            ShowTaskPane();
            if (_paneControl is WebPaneHost web) web.SwitchView("settings");
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
