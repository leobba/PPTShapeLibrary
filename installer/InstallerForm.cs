using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PPTShapeLibraryInstaller
{
    /// <summary>单项检测结果状态。</summary>
    internal enum CheckState
    {
        Ok,       // 满足
        Missing,  // 缺失（必需）
        Warn      // 缺失（可选/建议）
    }

    /// <summary>一项环境依赖。</summary>
    internal class ReqItem
    {
        public string Key;
        public string Title;
        public string Detail;
        public bool Required;          // 必需项：不满足不可安装
        public CheckState State;
        public string DownloadUrl;     // 官方下载页
        public string LocalPattern;    // 期望的离线安装包文件名通配（同目录"离线安装包"下）
        public string InstallArgs;     // 离线包静默安装参数
        public string LocalFile;       // 实际找到的离线包路径
        public Label StateLabel;
        public Label DetailLabel;
        public Button ActionButton;
    }

    internal class InstallerForm : Form
    {
        private readonly ReqItem[] _items;
        private Panel _listPanel;
        private TextBox _log;
        private Button _reinstall;
        private Button _install;
        private Button _uninstallOld;
        private Label _summary;
        private Label _existingLabel;
        private string _existingVersion;         // 已安装版本（用于冲突提示）
        private string _existingUninstallString; // 已安装项的卸载命令

        public InstallerForm()
        {
            _items = new ReqItem[]
            {
                new ReqItem
                {
                    Key = "ppt", Title = "Microsoft PowerPoint", Required = true,
                    Detail = "需要 PowerPoint 2016 / 2019 / 2021 / Microsoft 365（桌面版）",
                    DownloadUrl = "https://www.microsoft.com/microsoft-365",
                    LocalPattern = null, InstallArgs = null
                },
                new ReqItem
                {
                    Key = "net48", Title = ".NET Framework 4.8", Required = true,
                    Detail = "插件运行库（Windows 10 1809+ / Windows 11 通常已内置）",
                    DownloadUrl = "https://dotnet.microsoft.com/download/dotnet-framework/net48",
                    LocalPattern = "ndp48*.exe", InstallArgs = "/q /norestart"
                },
                new ReqItem
                {
                    Key = "vsto", Title = "Visual Studio 2010 Tools for Office Runtime (VSTO 4.0)", Required = true,
                    Detail = "Office 加载项运行时；缺失时 PowerPoint 里不会出现本插件",
                    DownloadUrl = "https://learn.microsoft.com/visualstudio/vsto/how-to-install-the-visual-studio-tools-for-office-runtime-redistributable",
                    LocalPattern = "vstor*.exe", InstallArgs = "/quiet /norestart"
                },
                new ReqItem
                {
                    Key = "webview2", Title = "Microsoft Edge WebView2 Runtime", Required = false,
                    Detail = "插件界面渲染组件；缺失时插件仍可安装，但面板会显示降级提示页",
                    DownloadUrl = "https://developer.microsoft.com/microsoft-edge/webview2/#download",
                    LocalPattern = "MicrosoftEdgeWebview2Setup.exe", InstallArgs = "/silent /install"
                },
            };

            BuildUi();
            RefreshChecks();
        }

        // ---------------- UI ----------------

        private void BuildUi()
        {
            Text = "PPTShapeLibrary 安装程序";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(700, 600);
            MinimumSize = new Size(660, 580);
            Font = new Font("Microsoft YaHei UI", 9F);

            var title = new Label
            {
                Text = "PPTShapeLibrary 安装程序",
                Font = new Font("Microsoft YaHei UI", 14F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(18, 14)
            };
            Controls.Add(title);

            var subtitle = new Label
            {
                Text = "安装前会检测运行环境；缺失的组件可直接下载安装，全部就绪后点击「开始安装」。",
                AutoSize = true,
                ForeColor = Color.FromArgb(90, 90, 90),
                Location = new Point(20, 46)
            };
            Controls.Add(subtitle);

            _listPanel = new Panel
            {
                Location = new Point(18, 74),
                Size = new Size(664, 300),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };
            Controls.Add(_listPanel);

            int y = 8;
            foreach (var it in _items)
            {
                var row = new Panel { Location = new Point(6, y), Size = new Size(636, 68), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };

                var state = new Label
                {
                    Text = "…",
                    Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold),
                    Location = new Point(6, 20),
                    Size = new Size(28, 26),
                    TextAlign = ContentAlignment.MiddleCenter
                };
                row.Controls.Add(state);
                it.StateLabel = state;

                var name = new Label
                {
                    Text = it.Title,
                    Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                    Location = new Point(40, 8),
                    Size = new Size(430, 20)
                };
                row.Controls.Add(name);

                var detail = new Label
                {
                    Text = it.Detail,
                    ForeColor = Color.FromArgb(100, 100, 100),
                    Location = new Point(40, 28),
                    Size = new Size(430, 34)
                };
                row.Controls.Add(detail);
                it.DetailLabel = detail;

                var btn = new Button
                {
                    Text = "下载安装",
                    Location = new Point(482, 14),
                    Size = new Size(110, 32),
                    FlatStyle = FlatStyle.System
                };
                var captured = it;
                btn.Click += (s, e) => OnAction(captured);
                row.Controls.Add(btn);
                it.ActionButton = btn;

                _listPanel.Controls.Add(row);
                y += 74;
            }

            var logLabel = new Label { Text = "检测 / 安装日志：", AutoSize = true, Location = new Point(18, 382) };
            Controls.Add(logLabel);

            _log = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Location = new Point(18, 404),
                Size = new Size(664, 104),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                BackColor = Color.FromArgb(250, 250, 250)
            };
            Controls.Add(_log);

            // 已安装旧版本提示（ClickOnce 不允许从新位置升级已安装项，必须先卸载）
            _existingLabel = new Label
            {
                Text = "",
                Location = new Point(18, 512),
                Size = new Size(664, 32),
                ForeColor = Color.FromArgb(120, 80, 0),
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            Controls.Add(_existingLabel);

            _summary = new Label
            {
                Text = "",
                Location = new Point(18, 548),
                Size = new Size(300, 20),
                Anchor = AnchorStyles.Left | AnchorStyles.Bottom
            };
            Controls.Add(_summary);

            _uninstallOld = new Button
            {
                Text = "卸载旧版本",
                Size = new Size(120, 30),
                Location = new Point(330, 546),
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
                Visible = false
            };
            _uninstallOld.Click += (s, e) => UninstallExisting();
            Controls.Add(_uninstallOld);

            _reinstall = new Button { Text = "重新检测", Size = new Size(90, 30), Location = new Point(458, 546), Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
            _reinstall.Click += (s, e) => RefreshChecks();
            Controls.Add(_reinstall);

            _install = new Button { Text = "开始安装", Size = new Size(126, 30), Location = new Point(556, 546), Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
            _install.Click += (s, e) => StartInstall();
            Controls.Add(_install);
        }

        // ---------------- 检测 ----------------

        private void RefreshChecks()
        {
            _log.Clear();
            Log("开始环境检测…");
            foreach (var it in _items)
            {
                switch (it.Key)
                {
                    case "ppt": CheckPowerPoint(it); break;
                    case "net48": CheckNet48(it); break;
                    case "vsto": CheckVsto(it); break;
                    case "webview2": CheckWebView2(it); break;
                }
                UpdateRow(it);
            }
            DetectExistingInstall();
            UpdateExistingUi();
            UpdateSummary();
        }

        // ---------------- 已安装版本检测 / 卸载（解决 ClickOnce "已安装另一版本" 冲突） ----------------

        private void DetectExistingInstall()
        {
            _existingVersion = null;
            _existingUninstallString = null;
            try
            {
                using (var root = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
                {
                    if (root != null)
                    {
                        foreach (var sub in root.GetSubKeyNames())
                        {
                            using (var k = root.OpenSubKey(sub))
                            {
                                if (k == null) continue;
                                string name = k.GetValue("DisplayName") as string;
                                if (string.IsNullOrEmpty(name)) continue;
                                if (name.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) < 0) continue;
                                _existingVersion = k.GetValue("DisplayVersion") as string;
                                _existingUninstallString = k.GetValue("UninstallString") as string;
                                break;
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void UpdateExistingUi()
        {
            if (!string.IsNullOrEmpty(_existingUninstallString))
            {
                _existingLabel.Text = "已检测到已安装版本 " + (_existingVersion ?? "") +
                    "。ClickOnce 不允许从新位置直接升级，若「开始安装」报『已安装另一版本』，请先点「卸载旧版本」。";
                _uninstallOld.Visible = true;
                Log("！ 检测到已安装版本 " + (_existingVersion ?? ""));
            }
            else
            {
                _existingLabel.Text = "未检测到已安装的旧版本。";
                _uninstallOld.Visible = false;
            }
        }

        /// <summary>卸载已安装版本；并清掉构建期遗留的 |vstolocal 开发式注册（否则会与正式版抢注册表）。</summary>
        private void UninstallExisting()
        {
            if (string.IsNullOrEmpty(_existingUninstallString))
            {
                MessageBox.Show("未检测到已安装版本。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var r = MessageBox.Show(
                "将卸载已安装的 PPTShapeLibrary（版本 " + (_existingVersion ?? "") + "）。\r\n\r\n" +
                "卸载命令：\r\n" + _existingUninstallString + "\r\n\r\n是否继续？（卸载期间请勿打开 PowerPoint）",
                "卸载旧版本", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (r != DialogResult.OK) return;

            try
            {
                Log("执行卸载：" + _existingUninstallString);
                RunCommandLine(_existingUninstallString);
                TryRemoveStaleVstoLocalKey();
                Log("卸载命令已执行完毕，并已清理开发式注册残留。");
                MessageBox.Show("卸载已完成。请点「重新检测」确认，然后再点「开始安装」。",
                    "卸载完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DetectExistingInstall();
                UpdateExistingUi();
                UpdateSummary();
            }
            catch (Exception ex)
            {
                Log("卸载失败：" + ex.Message);
                MessageBox.Show("卸载失败：" + ex.Message + "\r\n\r\n可改用：控制面板 → 程序和功能 → 卸载 PPTShapeLibrary",
                    "卸载失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void RunCommandLine(string cmd)
        {
            int exeEnd = cmd.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
            if (exeEnd > 0)
            {
                string exe = cmd.Substring(0, exeEnd + 4).Trim().Trim('"');
                string args = cmd.Substring(exeEnd + 4).Trim();
                var p = Process.Start(new ProcessStartInfo(exe, args) { UseShellExecute = false });
                if (p != null) p.WaitForExit();
            }
            else
            {
                var p = Process.Start(new ProcessStartInfo("cmd.exe", "/c " + cmd) { UseShellExecute = false });
                if (p != null) p.WaitForExit();
            }
        }

        /// <summary>仅当注册表项是构建期写入的 |vstolocal（开发式）时才删除，避免误删正式安装的注册。</summary>
        private static void TryRemoveStaleVstoLocalKey()
        {
            try
            {
                const string sub = @"Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary";
                string manifest = null;
                using (var k = Registry.CurrentUser.OpenSubKey(sub, false))
                {
                    if (k == null) return;
                    manifest = k.GetValue("Manifest") as string;
                }
                if (manifest == null || manifest.IndexOf("|vstolocal", StringComparison.OrdinalIgnoreCase) < 0) return;
                Registry.CurrentUser.DeleteSubKeyTree(sub, false);
            }
            catch { }
        }

        private void CheckPowerPoint(ReqItem it)
        {
            string path = ReadAppPath(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE");
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                it.State = CheckState.Ok;
                it.Detail = "已检测到：" + path;
                Log("✓ PowerPoint：" + path);
            }
            else
            {
                it.State = CheckState.Missing;
                it.Detail = "未检测到 PowerPoint（请先安装 Office/PowerPoint 桌面版）";
                Log("✗ 未检测到 PowerPoint 桌面版");
            }
        }

        private void CheckNet48(ReqItem it)
        {
            try
            {
                using (var k = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                        .OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
                {
                    if (k != null)
                    {
                        object rel = k.GetValue("Release");
                        int release = rel == null ? 0 : Convert.ToInt32(rel);
                        if (release >= 528040)   // 4.8
                        {
                            it.State = CheckState.Ok;
                            it.Detail = ".NET Framework 4.8 或更高（Release=" + release + "）";
                            Log("✓ .NET Framework：Release=" + release);
                            return;
                        }
                        it.State = CheckState.Missing;
                        it.Detail = "已装版本偏低（Release=" + release + "），需要 4.8 或更高";
                        Log("✗ .NET Framework 版本偏低：Release=" + release);
                        return;
                    }
                }
            }
            catch { }
            it.State = CheckState.Missing;
            it.Detail = "未检测到 .NET Framework 4.8";
            Log("✗ 未检测到 .NET Framework 4.8");
        }

        private void CheckVsto(ReqItem it)
        {
            string ver = null;
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                try
                {
                    using (var k = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view)
                            .OpenSubKey(@"SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"))
                    {
                        if (k != null)
                        {
                            object v = k.GetValue("Version");
                            if (v != null) { ver = v.ToString(); break; }
                        }
                    }
                }
                catch { }
            }
            if (!string.IsNullOrEmpty(ver))
            {
                it.State = CheckState.Ok;
                it.Detail = "已安装 VSTO Runtime（版本 " + ver + "）";
                Log("✓ VSTO Runtime：" + ver);
            }
            else
            {
                it.State = CheckState.Missing;
                it.Detail = "未安装 VSTO Runtime 4.0（缺失时插件无法被 PowerPoint 加载）";
                Log("✗ 未检测到 VSTO Runtime 4.0");
            }
        }

        private void CheckWebView2(ReqItem it)
        {
            const string guid = "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}";
            string ver = null;
            // 机器级（32/64 视图）+ 用户级
            foreach (var view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
            {
                foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
                {
                    try
                    {
                        using (var k = RegistryKey.OpenBaseKey(hive, view)
                                .OpenSubKey(@"SOFTWARE\Microsoft\EdgeUpdate\Clients\" + guid))
                        {
                            if (k != null)
                            {
                                object pv = k.GetValue("pv");
                                if (pv != null && !string.IsNullOrWhiteSpace(pv.ToString())) { ver = pv.ToString(); break; }
                            }
                        }
                    }
                    catch { }
                }
                if (!string.IsNullOrEmpty(ver)) break;
            }
            if (!string.IsNullOrEmpty(ver))
            {
                it.State = CheckState.Ok;
                it.Detail = "已安装 WebView2 Runtime（版本 " + ver + "）";
                Log("✓ WebView2 Runtime：" + ver);
            }
            else
            {
                it.State = CheckState.Warn;
                it.Detail = "未检测到 WebView2 Runtime（可继续安装，但插件面板会显示降级提示页）";
                Log("！ 未检测到 WebView2 Runtime");
            }
        }

        private static string ReadAppPath(string subKey)
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                try
                {
                    using (var k = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view).OpenSubKey(subKey))
                    {
                        if (k != null)
                        {
                            object v = k.GetValue(null);
                            if (v != null) return v.ToString().Trim('"');
                        }
                    }
                }
                catch { }
            }
            return null;
        }

        // ---------------- 行状态与操作 ----------------

        private void UpdateRow(ReqItem it)
        {
            if (it.State == CheckState.Ok)
            {
                it.StateLabel.Text = "✓";
                it.StateLabel.ForeColor = Color.FromArgb(28, 140, 60);
                it.ActionButton.Text = "已满足";
                it.ActionButton.Enabled = false;
            }
            else
            {
                it.StateLabel.Text = it.State == CheckState.Missing ? "✗" : "!";
                it.StateLabel.ForeColor = it.State == CheckState.Missing ? Color.FromArgb(200, 40, 40) : Color.FromArgb(214, 140, 0);

                // 本地离线包优先
                it.LocalFile = FindLocalPackage(it.LocalPattern);
                if (it.LocalFile != null)
                {
                    it.ActionButton.Text = "安装本地包";
                    it.ActionButton.Enabled = true;
                    it.Detail = it.Detail + "（已找到离线包：" + Path.GetFileName(it.LocalFile) + "）";
                }
                else
                {
                    it.ActionButton.Text = "下载安装";
                    it.ActionButton.Enabled = true;
                }
            }
            it.DetailLabel.Text = it.Detail;
        }

        private static string FindLocalPackage(string pattern)
        {
            if (string.IsNullOrEmpty(pattern)) return null;
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string[] dirs = { Path.Combine(baseDir, "离线安装包"), baseDir };
                foreach (var d in dirs)
                {
                    if (!Directory.Exists(d)) continue;
                    var files = Directory.GetFiles(d, pattern);
                    if (files.Length > 0) return files[0];
                }
            }
            catch { }
            return null;
        }

        private void OnAction(ReqItem it)
        {
            try
            {
                if (it.LocalFile != null && File.Exists(it.LocalFile))
                {
                    var r = MessageBox.Show(
                        "将静默安装本地离线包：\r\n" + it.LocalFile + "\r\n\r\n安装过程中可能弹出系统 UAC 授权提示，请允许。是否继续？",
                        "安装 " + it.Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (r != DialogResult.OK) return;
                    Log("正在安装本地包：" + Path.GetFileName(it.LocalFile) + " " + (it.InstallArgs ?? ""));
                    var psi = new ProcessStartInfo(it.LocalFile, it.InstallArgs ?? "")
                    {
                        UseShellExecute = true,
                        Verb = "runas"     // 需要管理员权限
                    };
                    var p = Process.Start(psi);
                    p.WaitForExit();
                    Log("本地包安装进程已退出（ExitCode=" + p.ExitCode + "），请点击「重新检测」确认结果。");
                    MessageBox.Show("安装程序已结束。请点击「重新检测」查看结果；部分组件需要重启计算机后生效。",
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Log("打开下载页面：" + it.DownloadUrl);
                Process.Start(new ProcessStartInfo(it.DownloadUrl) { UseShellExecute = true });
                MessageBox.Show(
                    "已打开官方页面，请下载并安装：" + it.Title + "\r\n\r\n" +
                    "安装完成后回到本窗口，点击「重新检测」。\r\n\r\n" +
                    "· 若页面上的下载链接已变化或无法下载（微软下载中心会调整），请改用离线方式：\r\n" +
                    "  把官方安装包放到本程序同目录的「离线安装包」文件夹，" + Environment.NewLine +
                    "  本窗口会自动识别并把按钮变为「安装本地包」：\r\n" +
                    "    ndp48-x86-x64-allos-enu.exe / vstor_redist.exe / MicrosoftEdgeWebview2Setup.exe",
                    "请安装 " + it.Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("操作失败：" + ex.Message);
                MessageBox.Show("操作失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateSummary()
        {
            bool blocked = false;
            bool webviewMissing = false;
            foreach (var it in _items)
            {
                if (it.Required && it.State != CheckState.Ok) blocked = true;
                if (it.Key == "webview2" && it.State != CheckState.Ok) webviewMissing = true;
            }
            if (blocked)
            {
                _summary.Text = "存在未满足的必需组件，请先安装后再点「开始安装」。";
                _summary.ForeColor = Color.FromArgb(200, 40, 40);
                _install.Enabled = false;
            }
            else if (webviewMissing)
            {
                _summary.Text = "必需组件已齐备；WebView2 缺失，界面将显示降级提示页。";
                _summary.ForeColor = Color.FromArgb(214, 140, 0);
                _install.Enabled = true;
            }
            else
            {
                _summary.Text = "环境检测通过，可以开始安装。";
                _summary.ForeColor = Color.FromArgb(28, 140, 60);
                _install.Enabled = true;
            }
        }

        // ---------------- 安装插件本体 ----------------

        private void StartInstall()
        {
            try
            {
                string setup = FindClickOnceSetup();
                if (setup == null)
                {
                    MessageBox.Show(
                        "未找到安装包（setup.exe）。\r\n请确认本程序与「ClickOnce」文件夹在同一目录下（完整分发目录应包含：安装程序、ClickOnce\\setup.exe、离线安装包\\）。",
                        "缺少安装包", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Log("✗ 未找到 ClickOnce\\setup.exe");
                    return;
                }

                // 已安装过其它位置的版本时，ClickOnce 会拒绝安装（"已安装另一版本"）→ 先卸载
                DetectExistingInstall();
                UpdateExistingUi();
                if (!string.IsNullOrEmpty(_existingUninstallString))
                {
                    var ask = MessageBox.Show(
                        "检测到已安装版本 " + (_existingVersion ?? "") + "。\r\n\r\n" +
                        "ClickOnce 不允许从新位置直接升级，若直接安装通常会报『已安装另一版本的自定义项』。\r\n" +
                        "是否先卸载该版本，再继续安装？",
                        "先卸载旧版本？", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (ask == DialogResult.Cancel) return;
                    if (ask == DialogResult.Yes)
                    {
                        UninstallExisting();
                        DetectExistingInstall();
                        UpdateExistingUi();
                        if (!string.IsNullOrEmpty(_existingUninstallString))
                        {
                            MessageBox.Show("旧版本似乎仍存在，已中止安装。请手动在「控制面板 → 程序和功能」卸载后再试。",
                                "已中止", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }

                Log("开始安装插件：" + setup);
                var p = Process.Start(new ProcessStartInfo(setup) { UseShellExecute = true, WorkingDirectory = Path.GetDirectoryName(setup) });
                p.WaitForExit();
                Log("安装程序已退出（ExitCode=" + p.ExitCode + "）。");
                DetectExistingInstall();
                UpdateExistingUi();
                MessageBox.Show(
                    "安装流程已结束。\r\n\r\n请打开 PowerPoint，在功能区应出现「素材库」选项卡。\r\n" +
                    "若未出现：文件 → 选项 → 加载项 → 管理「COM 加载项」→ 勾选 PPTShapeLibrary，然后重启 PowerPoint。",
                    "安装完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("安装失败：" + ex.Message);
                MessageBox.Show("安装失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string FindClickOnceSetup()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates =
            {
                Path.Combine(baseDir, "ClickOnce", "setup.exe"),
                Path.Combine(baseDir, "setup.exe"),
                Path.Combine(baseDir, "publish", "setup.exe"),
            };
            foreach (var c in candidates)
            {
                if (File.Exists(c)) return c;
            }
            return null;
        }

        private void Log(string msg)
        {
            _log.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + msg + Environment.NewLine);
        }
    }
}
