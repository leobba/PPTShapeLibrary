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
        private int _boundWindowHwnd;   // 窗格绑定窗口的 HWND（判断"是否当前窗口"用句柄，ReferenceEquals 对 COM 不可靠）

        /// <summary>取文档窗口句柄（不同 PPT 文档窗口 HWND 不同，用于准确判断窗格是否该迁移）。</summary>
        private static int WindowHwnd(PowerPoint.DocumentWindow win)
        {
            try { return win != null ? win.HWND : 0; } catch { return 0; }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                Log.Write("Startup begin");
                HookAppEvents();
                BridgeController.SettingsSaved += OnSettingsSaved;
                ShapeLibraryService.LibraryWillMutate += () => ShapeExporter.SweepLibPresPool(true, 0);   // 库变更前释放库文稿占用
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
                    try { ShapeLibraryService.AutoBackupIfNeeded(); Log.Write("Auto backup checked"); }
                    catch (Exception ex) { Log.Error(ex, "AutoBackupIfNeeded"); }
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
                try { ShapeExporter.SweepLibPresPool(true, 0); } catch { }   // 释放库文稿句柄池（避免残留打开的文件）
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
                int hwnd = WindowHwnd(win);
                Log.Write("[pane] WindowActivate caption=[" + cap + "] hwnd=" + hwnd + " boundHwnd=" + _boundWindowHwnd + " paneVisible=" + (_taskPane != null && _taskPane.Visible));
                if (_taskPane == null) return;
                if (!_taskPane.Visible) return;
                if (hwnd != 0 && hwnd == _boundWindowHwnd)
                {
                    // 同一文档窗口：PowerPoint 切换演示文稿后可能不重绘任务窗格，强制 false→true 重挂
                    _taskPane.Visible = false;
                    _taskPane.Visible = true;
                    Log.Write("[pane] WindowActivate force-repaint done");
                }
                else
                {
                    EnsurePaneForWindow(win);   // 不同文档窗口（独立 PPT）：迁移窗格过去
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
                ApplyPaneLayout();
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

        /// <summary>窗格是否仍然有效。演示文稿窗口关闭/迁移后 PowerPoint 会删除窗格对象，
        /// 而 _taskPane 字段仍指向它，之后任何访问都抛 0x800A01A8
        /// （用户现象：点「素材库」按钮打不开、日志里反复报"该任务窗格已被删除"）。
        /// 这里探活，失效即清空字段，交给 EnsureTaskPaneCreated 重建。</summary>
        private bool PaneIsAlive()
        {
            if (_taskPane == null) return false;
            try
            {
                bool v = _taskPane.Visible;   // 已删除的窗格：此访问即抛
                GC.KeepAlive(v);
                return true;
            }
            catch
            {
                _taskPane = null;
                return false;
            }
        }

        /// <summary>首次打开时创建任务窗格（Startup 不建，避免无活动窗口时绑定幽灵窗口导致永不显示）。</summary>
        private void EnsureTaskPaneCreated()
        {
            try
            {
                if (_taskPane != null && !PaneIsAlive())
                {
                    Log.Write("Task pane stale (deleted by PowerPoint) -> recreate");
                }
                if (_taskPane != null) return;
                if (_paneHost == null) return;
                PowerPoint.DocumentWindow win = null;
                try { win = Application.ActiveWindow; } catch { }
                if (win == null) return;   // 仍无活动窗口：等下次调用

                _taskPane = this.CustomTaskPanes.Add(_paneHost, "素材库", win);
                _boundWindowHwnd = WindowHwnd(win);
                ApplyPaneLayout();
                Log.Write("Task pane created lazily, boundHwnd=" + _boundWindowHwnd);
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
                int hwnd = WindowHwnd(win);
                if (hwnd != 0 && hwnd == _boundWindowHwnd) return;   // 已绑定当前窗口（按句柄判断）

                Log.Write("Migrating task pane (boundHwnd=" + _boundWindowHwnd + " -> " + hwnd + ")");
                bool wasVisible = _taskPane.Visible;
                int width = _taskPane.Width;
                this.CustomTaskPanes.Remove(_taskPane);
                // 释放旧宿主，避免 WebView2 残留
                var oldHost = _paneControl as WebPaneHost;
                try { oldHost?.PrepareShutdown(); } catch { }

                // 为当前窗口重建（WebView2 重新初始化）
                CreatePane();
                _taskPane = this.CustomTaskPanes.Add(_paneHost, "素材库", win);
                _boundWindowHwnd = hwnd;
                ApplyPaneLayout();
                // 关键：必须立即 Visible=true 触发内容控件 Load → WebPaneHost.InitializeAsync 启动；
                // 若设为 false，UserControl.Load 不触发，WebView2 永不初始化、窗格永不显示（此前 bug）。
                _taskPane.Visible = true;
                Log.Write("Task pane rebuilt & shown for new window (wasVisible=" + wasVisible + ", hwnd=" + hwnd + ")");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "EnsurePaneForWindow");
            }
        }

        /// <summary>把窗格按设置应用停靠位置与宽度（仅左/右停靠；旧 top/bottom 设置归一到右侧）。</summary>
        private void ApplyPaneLayout()
        {
            try
            {
                if (_taskPane == null) return;
                var s = ShapeLibraryService.LoadSettings();
                string dock = s.PaneDock ?? "left";
                if (dock != "left") dock = "right";   // 停靠仅左右（用户 24 批要求；top/bottom 兼容为 right）
                _taskPane.DockPosition = dock == "left"
                    ? Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft
                    : Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                int w = s.PaneWidth > 0 ? s.PaneWidth : 400;
                try { _taskPane.Width = w; } catch { }
            }
            catch (Exception ex) { Log.Error(ex, "ApplyPaneLayout"); }
        }

        /// <summary>切换窗格停靠位置（仅 "right"/"left"）并保存记忆。</summary>
        public void SetPaneDock(string dock)
        {
            try
            {
                if (dock != "right" && dock != "left") return;
                var s = ShapeLibraryService.LoadSettings();
                s.PaneDock = dock;
                ShapeLibraryService.SaveSettings(s);
                ApplyPaneLayout();
                Log.Write("Pane dock set to " + dock);
            }
            catch (Exception ex) { Log.Error(ex, "SetPaneDock"); }
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
