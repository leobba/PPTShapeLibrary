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
        private Button _repair;
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

            _repair = new Button
            {
                Text = "修复加载项注册",
                Size = new Size(120, 30),
                Location = new Point(330, 512),
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom
            };
            _repair.Click += (s, e) => RepairRegistration();
            Controls.Add(_repair);
            var repairTip = new ToolTip();
            repairTip.SetToolTip(_repair,
                "功能区没有「素材库」时点这里：重写加载项注册（当前用户）。\r\n" +
                "「所有用户」模式会额外写入 HKLM（需要管理员）。");

            _checkUpdate = new Button
            {
                Text = "检查更新",
                Size = new Size(100, 30),
                Location = new Point(222, 512),
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom
            };
            _checkUpdate.Click += (s, e) => CheckUpdateAsync(true);
            Controls.Add(_checkUpdate);

            _updateLabel = new Label
            {
                Text = "",
                Location = new Point(18, 516),
                Size = new Size(196, 22),
                Anchor = AnchorStyles.Left | AnchorStyles.Bottom,
                ForeColor = Color.FromArgb(90, 100, 110),
                AutoEllipsis = true
            };
            Controls.Add(_updateLabel);

            // 启动后自动检查一次（24h 缓存，不重复联网）
            Load += (s, e) => CheckUpdateAsync(false);
        }

        // ---------------- 更新检查（GitHub Releases） ----------------

        private Button _checkUpdate;
        private Label _updateLabel;
        private bool _updateChecking;

        /// <summary>manual=false 时走 24h 缓存；结果异步回填到界面。</summary>
        private void CheckUpdateAsync(bool manual)
        {
            if (_updateChecking) return;
            _updateChecking = true;
            if (manual) _checkUpdate.Enabled = false;
            _updateLabel.Text = "正在检查更新…";

            var ui = System.Threading.SynchronizationContext.Current;
            var t = new System.Threading.Thread(() =>
            {
                UpdateInfo info;
                try { info = UpdateChecker.Check(manual); }
                catch (Exception ex) { info = new UpdateInfo { Ok = false, Error = ex.Message }; }
                Action apply = () => ApplyUpdateInfo(info, manual);
                if (ui != null) ui.Post(_ => apply(), null);
                else { try { BeginInvoke(apply); } catch { } }
            });
            t.IsBackground = true;
            t.Start();
        }

        private void ApplyUpdateInfo(UpdateInfo info, bool manual)
        {
            _updateChecking = false;
            _checkUpdate.Enabled = true;
            string cur = UpdateChecker.CurrentVersion();

            if (!info.Ok)
            {
                _updateLabel.Text = "更新检查失败（点「检查更新」重试）";
                _updateLabel.ForeColor = Color.FromArgb(150, 120, 60);
                Log("更新检查失败：" + info.Error);
                if (manual)
                    MessageBox.Show(info.Error + "\r\n\r\n不影响安装；也可直接到发布页手动查看：\r\n" +
                                    UpdateChecker.ReleasesPageUrl,
                        "检查更新", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!info.HasUpdate)
            {
                _updateLabel.Text = "已是最新版本 " + cur;
                _updateLabel.ForeColor = Color.FromArgb(90, 100, 110);
                Log("更新检查：已是最新（当前 " + cur + "，最新 " + info.LatestVersion + "）" +
                    (info.FromCache ? "（缓存）" : ""));
                if (manual)
                    MessageBox.Show("当前版本 " + cur + " 已是最新。", "检查更新",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _updateLabel.Text = "发现新版本 " + info.LatestVersion + "（点「检查更新」查看）";
            _updateLabel.ForeColor = Color.FromArgb(0, 102, 204);
            Log("更新检查：发现新版本 " + info.LatestVersion + "（当前 " + cur + "）" +
                (info.FromCache ? "（缓存）" : ""));
            // 自动检查只做被动提示，不弹模态窗打断用户；手动点「检查更新」才弹下载确认
            if (manual) PromptUpdate(info);
        }

        private void PromptUpdate(UpdateInfo info)
        {
            string notes = string.IsNullOrWhiteSpace(info.Notes) ? "" : "\r\n\r\n更新说明：\r\n" + info.Notes;
            string assetLine = string.IsNullOrEmpty(info.AssetUrl)
                ? "\r\n\r\n（该版本没有上传分发包，只能到发布页手动下载）"
                : "";
            var r = MessageBox.Show(
                "发现新版本 " + info.LatestVersion + "（当前 " + info.CurrentVersion + "）。" + notes + assetLine +
                "\r\n\r\n是否现在下载并安装？（下载会先试直连，失败自动走 GitHub 加速站）",
                "发现新版本", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (r != DialogResult.Yes) return;

            if (string.IsNullOrEmpty(info.AssetUrl))
            {
                Process.Start(new ProcessStartInfo(info.ReleaseUrl ?? UpdateChecker.ReleasesPageUrl) { UseShellExecute = true });
                return;
            }

            string zip = Path.Combine(Path.GetTempPath(), info.AssetName ?? "pptshape-update.zip");
            try
            {
                Log("开始下载：" + info.AssetName);
                _updateLabel.Text = "正在下载…";
                UpdateChecker.Download(info.AssetUrl, zip, (got, total) =>
                {
                    try
                    {
                        string txt = total > 0
                            ? string.Format("正在下载… {0:0.0}/{1:0.0} MB", got / 1048576.0, total / 1048576.0)
                            : string.Format("正在下载… {0:0.0} MB", got / 1048576.0);
                        if (_updateLabel != null) _updateLabel.Text = txt;
                    }
                    catch { }
                });

                string expected = UpdateChecker.FetchExpectedHash(info.Sha256Url);
                string actual = UpdateChecker.Sha256(zip);
                Log("下载完成：" + actual);
                if (!string.IsNullOrEmpty(expected) && !string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("下载文件校验失败，已放弃安装。\r\n期望 " + expected + "\r\n实际 " + actual,
                        "校验失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    try { File.Delete(zip); } catch { }
                    _updateLabel.Text = "校验失败";
                    return;
                }

                _updateLabel.Text = "正在启动新版安装程序…";
                var p = UpdateChecker.ApplyUpdate(zip);
                Log("已启动新版安装程序：" + (p != null ? p.ProcessName : "?"));
                MessageBox.Show(
                    "新版安装程序已启动，请按它的提示完成安装。\r\n\r\n" +
                    "建议：先完全退出 PowerPoint 再安装。安装完成后本窗口可关闭。",
                    "开始安装新版本", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("更新失败：" + ex.Message);
                _updateLabel.Text = "更新失败";
                MessageBox.Show("更新失败：" + ex.Message + "\r\n\r\n可到发布页手动下载：\r\n" +
                                (info.ReleaseUrl ?? UpdateChecker.ReleasesPageUrl),
                    "更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ---------------- 加载项注册修复 ----------------

        /// <summary>加载项注册的 4 个位置：HKCU/HKLM × 版本无关/版本化。</summary>
        private static readonly string[][] AddinKeyPaths =
        {
            new[] { "HKCU", @"Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary" },
            new[] { "HKCU", @"Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary" },
            new[] { "HKLM", @"Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary" },
            new[] { "HKLM", @"Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary" }
        };

        private static bool IsAdmin()
        {
            try
            {
                var id = System.Security.Principal.WindowsIdentity.GetCurrent();
                return new System.Security.Principal.WindowsPrincipal(id)
                    .IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch { return false; }
        }

        /// <summary>写入加载项注册。manifestPath 为 .vsto 的绝对路径（写入时转成 file:///…|vstolocal）。</summary>
        private static int WriteAddinKeys(string manifestPath, bool includeUser, bool includeMachine)
        {
            string manifest = "file:///" + manifestPath.Replace('\\', '/') + "|vstolocal";
            int written = 0;
            foreach (var item in AddinKeyPaths)
            {
                bool machine = item[0] == "HKLM";
                if (machine && !includeMachine) continue;
                if (!machine && !includeUser) continue;
                try
                {
                    var root = machine ? Registry.LocalMachine : Registry.CurrentUser;
                    using (var key = root.CreateSubKey(item[1]))
                    {
                        if (key == null) continue;
                        key.SetValue("Description", "PPTShapeLibrary", RegistryValueKind.String);
                        key.SetValue("FriendlyName", "PPTShapeLibrary", RegistryValueKind.String);
                        key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                        key.SetValue("Manifest", manifest, RegistryValueKind.String);
                        written++;
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("写入失败 " + item[0] + "\\" + item[1] + "：" + ex.Message);
                }
            }
            return written;
        }

        /// <summary>提权子模式的入口：只写 HKLM 两处键，返回写入数量（0=失败）。</summary>
        public static int WriteMachineAddinKeys(string manifestPath)
        {
            return WriteAddinKeys(manifestPath, false, true);
        }

        /// <summary>「修复加载项注册」：这是"HKCU 版本无关键被清掉 → 功能区没有素材库"的用户级修复入口。</summary>
        private void RepairRegistration()
        {
            string manifest = FindDeploymentManifest();
            if (manifest == null)
            {
                MessageBox.Show(
                    "未找到 App\\PPTShapeLibrary.vsto，无法修复注册。\r\n" +
                    "请确认安装目录完整（App\\PPTShapeLibrary.vsto + App\\Application Files\\）。",
                    "缺少安装包", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (IsPowerPointRunning())
            {
                MessageBox.Show("请先完全退出 PowerPoint 再修复注册（否则可能被覆盖回去）。",
                    "请先关闭 PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int n = WriteAddinKeys(manifest, true, IsAdmin());
            Log("修复加载项注册：写入 " + n + " 处（当前用户" + (IsAdmin() ? " + 所有用户" : "") + "）");

            string msg = "已写入 " + n + " 处加载项注册（当前用户）。";
            if (!IsAdmin())
            {
                var ask = MessageBox.Show(
                    msg + "\r\n\r\n是否同时为**所有用户**写入注册（HKLM）？\r\n" +
                    "这会弹一次 UAC；写入后即使用户级注册被清掉，插件仍能加载。",
                    "是否写入所有用户？", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (ask == DialogResult.Yes)
                {
                    try
                    {
                        var psi = new ProcessStartInfo
                        {
                            FileName = Application.ExecutablePath,
                            Arguments = "--repair-machine \"" + manifest + "\"",
                            UseShellExecute = true,
                            Verb = "runas"
                        };
                        var p = Process.Start(psi);
                        if (p != null) p.WaitForExit(60000);
                        Log("已请求提权写入 HKLM（退出码 " + (p != null ? p.ExitCode.ToString() : "?") + "）");
                    }
                    catch (Exception ex)
                    {
                        Log("提权失败：" + ex.Message);
                        MessageBox.Show("提权被取消或失败：" + ex.Message, "未写入所有用户",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            MessageBox.Show(msg + "\r\n\r\n请重新打开 PowerPoint，功能区应出现「素材库」选项卡。",
                "修复完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                // 用户级（HKCU）与机器级（HKLM，含 32 位视图）都要扫：
                // 只扫 HKCU 时，"为所有用户安装"的 ClickOnce/ARP 记录看不见，
                // 「卸载旧版本」按钮就不会出现，而文档恰恰让用户用它来解「已安装另一版本」。
                var roots = new[]
                {
                    Registry.CurrentUser,
                    Registry.LocalMachine,
                    RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                };
                foreach (var hive in roots)
                {
                    if (_existingUninstallString != null) break;
                    using (var root = hive.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
                    {
                        if (root == null) continue;
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

        /// <summary>PowerPoint 是否在运行。安装/卸载期间它会锁住加载项 DLL 与清单，
        /// 导致 VSTOInstaller 半途失败、留下半注册状态。</summary>
        private static bool IsPowerPointRunning()
        {
            try { return Process.GetProcessesByName("POWERPNT").Length > 0; }
            catch { return false; }
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
            if (IsPowerPointRunning())
            {
                MessageBox.Show("PowerPoint 正在运行，卸载会因文件被占用而失败。请完全退出 PowerPoint 后重试。",
                    "请先关闭 PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Log("✗ PowerPoint 仍在运行，卸载中止");
                return;
            }

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

        private void RunCommandLine(string cmd)
        {
            if (string.IsNullOrWhiteSpace(cmd)) return;
            string exe, args;
            int exeEnd = cmd.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
            if (exeEnd > 0)
            {
                // 只在".exe"之后是字符串末尾或空白时才当成可执行文件边界，
                // 否则形如 "C:\Program Files\My.exe Tools\uninstall.exe" /S 的路径会被切错。
                int after = exeEnd + 4;
                if (after < cmd.Length && !char.IsWhiteSpace(cmd[after]))
                {
                    exeEnd = cmd.LastIndexOf(".exe", StringComparison.OrdinalIgnoreCase);
                    after = exeEnd + 4;
                }
                if (exeEnd <= 0) { exe = cmd.Trim().Trim('"'); args = ""; }
                else
                {
                    exe = cmd.Substring(0, after).Trim().Trim('"');
                    args = cmd.Substring(after).Trim();
                }
            }
            else
            {
                exe = "cmd.exe";
                args = "/c " + cmd;
            }
            var p = Process.Start(new ProcessStartInfo(exe, args) { UseShellExecute = false });
            if (p == null) return;
            // 必须有超时：卸载器卡住（或 dfshim 弹了错误框）时不能把安装器 UI 线程永久挂死
            if (!p.WaitForExit(5 * 60 * 1000))
            {
                Log("卸载命令 5 分钟未结束，已放弃等待（进程仍在后台运行）：" + exe);
                try { p.Kill(); } catch { }
            }
        }

        /// <summary>仅当注册表项是构建期写入的 |vstolocal（开发式）时才删除，避免误删正式安装的注册。
        /// 注意：加载项注册可能同时存在于"版本无关"与"16.0 版本化"两处
        /// （scripts\dev-register.ps1 两处都写），只清一处会留下残留继续与正式安装抢注册表，
        /// 再次触发"已安装另一版本的自定义项"。两处都要检查。</summary>
        private static void TryRemoveStaleVstoLocalKey()
        {
            string[] subs =
            {
                @"Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary",
                @"Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary"
            };
            foreach (string sub in subs)
            {
                try
                {
                    string manifest;
                    using (var k = Registry.CurrentUser.OpenSubKey(sub, false))
                    {
                        if (k == null) continue;
                        manifest = k.GetValue("Manifest") as string;
                    }
                    if (manifest == null || manifest.IndexOf("|vstolocal", StringComparison.OrdinalIgnoreCase) < 0) continue;
                    Registry.CurrentUser.DeleteSubKeyTree(sub, false);
                }
                catch { }
            }
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
                // PowerPoint 在运行会锁住加载项 DLL/清单，导致 VSTOInstaller 半途失败、留下半注册状态
                if (IsPowerPointRunning())
                {
                    var c = MessageBox.Show(
                        "检测到 PowerPoint 正在运行。\r\n\r\n" +
                        "安装/升级过程中 PowerPoint 会锁住插件文件，容易导致安装半途失败并留下半注册状态。\r\n" +
                        "请先完全退出 PowerPoint（含所有窗口），然后点「确定」继续；或点「取消」稍后再装。",
                        "请先关闭 PowerPoint", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (c != DialogResult.OK) { Log("！ 用户选择先关闭 PowerPoint 再安装"); return; }
                    if (IsPowerPointRunning())
                    {
                        MessageBox.Show("PowerPoint 仍在运行，安装已中止。请完全退出后重试。",
                            "仍检测到 PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        Log("✗ PowerPoint 仍在运行，安装中止");
                        return;
                    }
                }

                // 优先用部署清单(.vsto)安装：其 codebase 是相对路径(Application Files\...)，
                // 因此安装源就是本目录，拷到任何机器/路径都能装；
                // 而 setup.exe 里烤入了发布时的绝对 URL，不适用于分发。
                string manifest = FindDeploymentManifest();
                string setup = FindClickOnceSetup();
                if (manifest == null && setup == null)
                {
                    MessageBox.Show(
                        "未找到安装包。\r\n完整分发目录应包含：安装程序、App\\PPTShapeLibrary.vsto、App\\Application Files\\、离线安装包\\。",
                        "缺少安装包", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Log("✗ 未找到 App\\PPTShapeLibrary.vsto 或 setup.exe");
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

                bool installed = false;
                if (manifest != null)
                {
                    string url = new Uri(manifest).AbsoluteUri;   // file:///...（按本目录实际位置生成，可移植）
                    string vstoInstaller = FindVstoInstaller();
                    Log("安装源（部署清单）：" + url);
                    if (vstoInstaller != null)
                    {
                        Log("调用 VSTOInstaller /Install …");
                        var p = Process.Start(new ProcessStartInfo(vstoInstaller, "/Install \"" + url + "\"") { UseShellExecute = false });
                        if (p != null) p.WaitForExit();
                        int code = p == null ? -1 : p.ExitCode;
                        Log("VSTOInstaller 退出码：" + code);
                        // 以"是否真的出现安装记录"为准，而不是仅看退出码
                        DetectExistingInstall();
                        installed = !string.IsNullOrEmpty(_existingUninstallString);
                        if (!installed) Log("未检测到安装记录，改用系统关联方式重试…");
                    }
                    if (!installed)
                    {
                        // 回退：等同双击 .vsto，由系统关联的安装程序弹出向导
                        Log("改用系统关联方式安装 .vsto（会弹出安装向导，请按提示完成）…");
                        var p2 = Process.Start(new ProcessStartInfo(manifest) { UseShellExecute = true });
                        if (p2 != null) p2.WaitForExit();
                        DetectExistingInstall();
                        installed = !string.IsNullOrEmpty(_existingUninstallString);
                    }
                }
                else
                {
                    Log("开始安装插件（setup.exe 回退路径）：" + setup);
                    var p = Process.Start(new ProcessStartInfo(setup) { UseShellExecute = true, WorkingDirectory = Path.GetDirectoryName(setup) });
                    if (p != null) p.WaitForExit();
                    installed = true;
                }

                DetectExistingInstall();
                UpdateExistingUi();
                MessageBox.Show(
                    (installed ? "安装流程已结束。\r\n\r\n" : "安装未确认成功（可能被取消）。\r\n\r\n") +
                    "请打开 PowerPoint，在功能区应出现「素材库」选项卡。\r\n" +
                    "若未出现：文件 → 选项 → 加载项 → 管理「COM 加载项」→ 勾选 PPTShapeLibrary，然后重启 PowerPoint。",
                    "安装", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("安装失败：" + ex.Message);
                MessageBox.Show("安装失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>VSTO 安装器路径（用于 /Install、/Uninstall）。</summary>
        private static string FindVstoInstaller()
        {
            string[] candidates =
            {
                @"C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
                @"C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
            };
            foreach (var c in candidates)
            {
                if (File.Exists(c)) return c;
            }
            return null;
        }

        /// <summary>部署清单(.vsto)：优先 App\（可移植），兼容 ClickOnce\ 与根目录。</summary>
        private static string FindDeploymentManifest()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates =
            {
                Path.Combine(baseDir, "App", "PPTShapeLibrary.vsto"),
                Path.Combine(baseDir, "ClickOnce", "PPTShapeLibrary.vsto"),
                Path.Combine(baseDir, "PPTShapeLibrary.vsto"),
            };
            foreach (var c in candidates)
            {
                if (File.Exists(c)) return c;
            }
            return null;
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
