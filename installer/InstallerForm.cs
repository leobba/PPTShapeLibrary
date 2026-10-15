using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
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
        /// <summary>托管在**本仓库固定 tag**上的资产候选（按优先级）。微软官方源在国内常常打不开，
        /// 而 GitHub 资产可以走加速站；这里放的就是随仓库分发的官方安装包。</summary>
        public PrereqAsset[] Packages;
        public Label StateLabel;
        public Label DetailLabel;
        public Button ActionButton;
    }

    /// <summary>一个可下载的依赖包（固定 tag 的 Release 资产）。</summary>
    internal class PrereqAsset
    {
        public string FileName;   // 资产文件名
        public string Kind;       // exe / msi / zip（zip 会解压后运行其中的 exe 或 msi）
        public string Args;       // 静默安装参数（msi 会被替换成 msiexec /i … /qn）
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
                    LocalPattern = "vstor*.exe;vstor*.zip", InstallArgs = "/quiet /norestart",
                    // 官方源在国内基本打不开 → 优先从本仓库固定 tag 拉（可用加速站）。
                    // vstor_redist.exe 是官方自解压包；vstor_runtime.zip 是同一份内容的打包形式
                    // （install.exe + vstor40_x64.MSI/CAB，均为微软签名文件），官方包拉不到时用它。
                    Packages = new[]
                    {
                        new PrereqAsset { FileName = "vstor_redist.exe", Kind = "exe", Args = "/quiet /norestart" },
                        new PrereqAsset { FileName = "vstor_runtime.zip", Kind = "zip", Args = "/q" },
                    }
                },
                new ReqItem
                {
                    Key = "webview2", Title = "Microsoft Edge WebView2 Runtime", Required = false,
                    Detail = "插件界面渲染组件；缺失时插件仍可安装，但面板会显示降级提示页",
                    DownloadUrl = "https://developer.microsoft.com/microsoft-edge/webview2/#download",
                    LocalPattern = "MicrosoftEdgeWebview2Setup.exe", InstallArgs = "/silent /install",
                    Packages = new[]
                    {
                        new PrereqAsset { FileName = "MicrosoftEdgeWebview2Setup.exe", Kind = "exe", Args = "/silent /install" },
                    }
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
                    "未找到安装包，无法修复注册。\r\n\r\n" +
                    "· 单文件包安装：请先运行安装器完成安装（文件在 %LOCALAPPDATA%\\PPTShapeLibrary\\addin\\）\r\n" +
                    "· 目录分发形态：请确认本程序同目录有 App\\PPTShapeLibrary.vsto + App\\Application Files\\",
                    "缺少安装包", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Log("用于修复的清单：" + manifest);
            if (IsPowerPointRunning())
            {
                MessageBox.Show("请先完全退出 PowerPoint 再修复注册（否则可能被覆盖回去）。",
                    "请先关闭 PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 修复前先把会让 VSTO 拒绝加载/安装的残留清掉：
            // 光写加载项注册是不够的 —— 实测若 VSTA\Solutions 里还留着本产品记录，
            // 或 ARP 里有指向 bin\Debug 的陈旧卸载项，VSTO 会认为"已安装另一版本"而拒绝加载。
            int n = RepairAddinRegistrationCore(manifest, Log);

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

        /// <summary>
        /// 修复加载项注册的**核心逻辑**（GUI 与 CLI --repair 共用）：
        /// 清 VSTA 记录 + 清开发式 ARP + 清损坏的 ClickOnce 状态 → 写加载项注册。返回写入处数。
        /// </summary>
        internal static int RepairAddinRegistrationCore(string manifestPath, Action<string> say)
        {
            // 光写注册不够：实测只要 VSTA\Solutions 里还留着本产品记录，或 ARP 里有指向
            // bin\Debug 的陈旧卸载项，VSTO 就会认为"已安装另一版本"而拒绝加载/安装。
            ClearVstaRecords(say);
            ClearDevArpRecords(say);
            // 损坏的 ClickOnce 部署状态（如 UpdateSkipTime 元数据坏了）会让安装永远失败 -400
            int coc = ClearBrokenClickOnceState(say);
            if (coc > 0) say("已清理 " + coc + " 条损坏的 ClickOnce 记录");

            int n = WriteAddinKeys(manifestPath, true, IsAdmin());
            say("修复加载项注册：写入 " + n + " 处（当前用户" + (IsAdmin() ? " + 所有用户" : "") + "）");
            return n;
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

        /// <summary>从 ARP 读已安装版本的卸载命令（找不到返回 null）。</summary>
        private static string FindExistingUninstallString()
        {
            try
            {
                var roots = new[]
                {
                    Registry.CurrentUser,
                    Registry.LocalMachine,
                    RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                };
                foreach (var hive in roots)
                {
                    using (var root = hive.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
                    {
                        if (root == null) continue;
                        foreach (var sub in root.GetSubKeyNames())
                        {
                            using (var k = root.OpenSubKey(sub))
                            {
                                if (k == null) continue;
                                var name = k.GetValue("DisplayName") as string;
                                if (string.IsNullOrEmpty(name)) continue;
                                if (name.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) < 0) continue;
                                return k.GetValue("UninstallString") as string;
                            }
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 无人值守更新：`--pending-update &lt;解压目录&gt;` 走这里。
        ///
        /// 流程：等 PowerPoint 退出 → 卸载已装的旧版（ClickOnce 不允许就地升级）
        ///       → 用解压出来的部署清单装新版 → 结果写到 %TEMP%\pptsl_pending_update.log。
        /// 全程**不弹任何 UI**（用户在 PowerPoint 里点完"立即更新"就把 PowerPoint 关了，不该再被打断）。
        /// 返回 0 表示成功。
        /// </summary>
        public static int RunPendingUpdate(string dir)
        {
            var logPath = Path.Combine(Path.GetTempPath(), "pptsl_pending_update.log");
            Action<string> say = (m) =>
            {
                var line = DateTime.Now.ToString("HH:mm:ss") + "  " + m;
                try { File.AppendAllText(logPath, line + Environment.NewLine, Encoding.UTF8); } catch { }
            };
            try { File.Delete(logPath); } catch { }

            try
            {
                say("待安装目录：" + dir);
                var manifest = Path.Combine(dir, "App", "PPTShapeLibrary.vsto");
                if (!File.Exists(manifest))
                {
                    say("✗ 未找到 " + manifest);
                    return 2;
                }

                // ① 等 PowerPoint 退出（它锁着插件 DLL/清单，装到一半失败会留半注册状态）
                //    注意：只等**真有窗口**的 PowerPoint。实测见过"僵尸进程"——
                //    用户关掉窗口后进程没退干净（无主窗口、却有近百线程常驻），
                //    按进程名判断会白等 180 秒超时，用户完全不知道为什么。
                int zombie = CountZombiePowerPoint();
                if (zombie > 0)
                    say("提示：检测到 " + zombie + " 个无窗口的 PowerPoint 残留进程（不影响安装，继续）");

                say("等待 PowerPoint 退出…");
                int waited = 0;
                while (IsPowerPointRunning() && waited < 180)
                {
                    Thread.Sleep(2000);
                    waited += 2;
                }
                if (IsPowerPointRunning())
                {
                    say("✗ 等待超时（" + waited + " 秒），PowerPoint 仍在运行，已放弃");
                    // 明确告诉用户该怎么办（原来只有一行日志，用户在安装器界面上看不到原因）
                    try
                    {
                        MessageBox.Show(
                            "PowerPoint 还在运行，安装无法继续（它会锁住插件文件）。\r\n\r\n" +
                            "请**完全关闭 PowerPoint**（确认所有窗口都已关闭），然后重新点更新。\r\n\r\n" +
                            "如果任务管理器里已经没有 PowerPoint 了，请告诉我，我再排查。",
                            "请关闭 PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    catch { }
                    return 3;
                }
                say("PowerPoint 已退出（等待 " + waited + " 秒）");

                // ② 卸载旧版
                //    分两种：
                //     · 开发式（Addins 键带 |vstolocal，或 ARP 的卸载串指向 bin\Debug|bin\Release）
                //       —— 这种**不是** VSTOInstaller 正规装的，用 /Uninstall 会挂住（实测卡 120s 超时）
                //       → 直接删注册表键
                //     · 正式安装（ClickOnce）→ 走 ARP 里的卸载串
                if (HasDevRegistration() || ArpRecordIsDev())
                {
                    say("检测到开发式注册/陈旧卸载项：直接清理注册表键（VSTOInstaller 处理不了这种）");
                    ClearDevRegistration(say);
                }
                else
                {
                    var uninstall = FindExistingUninstallString();
                    if (!string.IsNullOrEmpty(uninstall))
                    {
                        // 先清一次损坏的 ClickOnce 状态：状态坏了的时候 /Uninstall 也会**卡死**
                        // （实测卡满 120 秒超时），而且卸载本来就是要重装的，不卸也行。
                        int pre = ClearBrokenClickOnceState(say);
                        if (pre > 0)
                        {
                            say("检测到 ClickOnce 状态损坏，已清理 " + pre + " 条 → 跳过卸载，直接重装");
                        }
                        else
                        {
                            say("卸载旧版：" + uninstall);
                            try { RunCommandLineStatic(uninstall, say, 120000); }
                            catch (Exception ex) { say("卸载异常（继续尝试安装）：" + ex.Message); }
                        }
                    }
                    else { say("未检测到旧版安装记录，直接安装"); }
                }

                // ★ 无论走哪条路，装之前都要清掉 VSTA\Solutions 里的本产品记录。
                //   它与"当前是否注册了加载项"无关：VSTO 用它判定"该产品已安装另一版本"，
                //   残留时 VSTOInstaller 一定报 AddInAlreadyInstalledException（实测多轮）。
                ClearVstaRecords(say);

                // ③ 安装新版
                var vi = FindVstoInstaller();
                var url = new Uri(manifest).AbsoluteUri;
                say("安装源：" + url);

                // 装之前先清一次损坏的 ClickOnce 状态：坏的时候任何位置都装不进去（-400），
                // 且用户完全无从下手。清理只针对本产品，正常机器上是空操作。
                int cleaned = ClearBrokenClickOnceState(say);
                if (cleaned > 0) say("已清理 " + cleaned + " 条损坏的 ClickOnce 记录，准备重装");

                int vstoCode = 0;
                if (vi != null)
                {
                    var p = Process.Start(new ProcessStartInfo(vi, "/Install \"" + url + "\"") { UseShellExecute = false });
                    if (p != null) p.WaitForExit();
                    vstoCode = p == null ? -1 : p.ExitCode;
                    say("VSTOInstaller 退出码：" + vstoCode);

                    // -400 = 部署状态坏了：再清一次（这次连缓存一起清）后重试一遍
                    if (vstoCode != 0)
                    {
                        say("安装失败（" + vstoCode + "），按『ClickOnce 状态损坏』再清一次后重试…");
                        int again = ClearBrokenClickOnceState(say);
                        say("  本轮又清理 " + again + " 条");
                        var p2 = Process.Start(new ProcessStartInfo(vi, "/Install \"" + url + "\"") { UseShellExecute = false });
                        if (p2 != null) p2.WaitForExit();
                        vstoCode = p2 == null ? -1 : p2.ExitCode;
                        say("重试后 VSTOInstaller 退出码：" + vstoCode);
                    }
                }
                else
                {
                    var p = Process.Start(new ProcessStartInfo(manifest) { UseShellExecute = true });
                    if (p != null) p.WaitForExit();
                    say("已用系统关联方式安装（无 VSTOInstaller）");
                }

                var now = FindExistingUninstallString();
                bool ok = !string.IsNullOrEmpty(now);
                say(ok ? "✓ 安装完成" : "✗ 未检测到安装记录（可能被取消或失败）");
                return ok ? 0 : 4;
            }
            catch (Exception ex)
            {
                say("✗ 异常：" + ex.Message);
                return 1;
            }
        }

        /// <summary>是否存在开发式注册（Manifest 带 |vstolocal）。</summary>
        private static bool HasDevRegistration()
        {
            try
            {
                foreach (var keyPath in new[]
                {
                    @"Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary",
                    @"Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary"
                })
                {
                    using (var k = Registry.CurrentUser.OpenSubKey(keyPath))
                    {
                        if (k == null) continue;
                        var mf = k.GetValue("Manifest") as string;
                        if (!string.IsNullOrEmpty(mf) && mf.IndexOf("vstolocal", StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>ARP 里是否存在**指向开发路径**的陈旧卸载项（bin\Debug / bin\Release / vstolocal）。
        /// 这种记录 VSTOInstaller /Uninstall 卸不掉（会挂住等超时），只能直接删键。</summary>
        private static bool ArpRecordIsDev()
        {
            try
            {
                foreach (var hive in new[] { Registry.CurrentUser, Registry.LocalMachine })
                {
                    using (var root = hive.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
                    {
                        if (root == null) continue;
                        foreach (var sub in root.GetSubKeyNames())
                        {
                            using (var k = root.OpenSubKey(sub))
                            {
                                if (k == null) continue;
                                var name = k.GetValue("DisplayName") as string;
                                if (string.IsNullOrEmpty(name)) continue;
                                if (name.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) < 0) continue;
                                var un = k.GetValue("UninstallString") as string;
                                if (string.IsNullOrEmpty(un)) continue;
                                // 注意：卸载串里是 **正斜杠**（file:///E:/…/bin/Debug/…），
                                // 直接找 "bin\Debug" 会漏判 —— 必须先把分隔符统一。
                                var norm = un.Replace('/', '\\');
                                if (norm.IndexOf("bin\\Debug", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    norm.IndexOf("bin\\Release", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    norm.IndexOf("vstolocal", StringComparison.OrdinalIgnoreCase) >= 0)
                                    return true;
                            }
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>清掉开发式注册：两处 Addins 键 + VSTO SolutionMetadata + AddinsData 缓存。</summary>
        private static void ClearDevRegistration(Action<string> say)
        {
            try
            {
                foreach (var keyPath in new[]
                {
                    @"Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary",
                    @"Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary"
                })
                {
                    try { Registry.CurrentUser.DeleteSubKeyTree(keyPath, false); say("已删除注册：" + keyPath); }
                    catch (Exception ex) { say("删除注册失败（忽略）：" + keyPath + " → " + ex.Message); }
                }
                try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Office\16.0\PowerPoint\AddinsData\PPTShapeLibrary", false); } catch { }
                try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Office\PowerPoint\AddinsData\PPTShapeLibrary", false); } catch { }

                // VSTO 的"已安装"记录（以**产品名**为唯一键）：不清掉的话后面装新版会报
                // AddInAlreadyInstalledException（"已安装另一版本的自定义项"）。
                using (var sm = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VSTO\SolutionMetadata"))
                {
                    if (sm != null)
                    {
                        foreach (var sub in sm.GetSubKeyNames())
                        {
                            using (var k = sm.OpenSubKey(sub))
                            {
                                if (k == null) continue;
                                var name = k.GetValue("addInName") as string;
                                if (!string.IsNullOrEmpty(name) &&
                                    name.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    try
                                    {
                                        Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\VSTO\SolutionMetadata\" + sub, false);
                                        say("已删除 VSTO 安装记录：" + sub);
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }

                // ⚠️ 还要删掉**陈旧的 ARP 卸载项**（指向开发路径的那些），详见 ClearDevArpRecords 注释。
                ClearDevArpRecords(say);

                // ★ 最关键的一条：VSTA\Solutions —— VSTO 用它记"这个产品装过哪个版本"，
                //   是 AddInAlreadyInstalledException（"已安装另一版本的自定义项"）的**真正判定来源**。
                ClearVstaRecords(say);
            }
            catch (Exception ex) { say("清理开发式注册异常（忽略）：" + ex.Message); }
        }

        /// <summary>删掉 ARP 里**指向开发路径**的卸载项（bin\Debug / bin\Release / |vstolocal）。
        /// 这种记录不是 VSTOInstaller 正规装的，用 /Uninstall 去卸会**挂死**（实测卡满 120s 超时）；
        /// 而且它同样会让 VSTO 认为"已安装另一版本"。正式安装（指向别人机器/cache 的）保持不动。</summary>
        private static void ClearDevArpRecords(Action<string> say)
        {
            foreach (var hive in new[] { Registry.CurrentUser, Registry.LocalMachine })
            {
                try
                {
                    using (var root = hive.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
                    {
                        if (root == null) continue;
                        foreach (var sub in root.GetSubKeyNames())
                        {
                            using (var k = root.OpenSubKey(sub))
                            {
                                if (k == null) continue;
                                var name = k.GetValue("DisplayName") as string;
                                if (string.IsNullOrEmpty(name)) continue;
                                if (name.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) < 0) continue;
                                var un = k.GetValue("UninstallString") as string;
                                // 卸载串里是**正斜杠**（file:///E:/…/bin/Debug/…）→ 必须先统一分隔符，
                                // 否则按 "bin\Debug" 匹配会漏判（实测踩过）。
                                var norm = (un ?? "").Replace('/', '\\');
                                bool isDev = norm.IndexOf("bin\\Debug", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                             norm.IndexOf("bin\\Release", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                             norm.IndexOf("vstolocal", StringComparison.OrdinalIgnoreCase) >= 0;
                                if (!isDev) continue;
                                try
                                {
                                    hive.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + sub, false);
                                    say("已删除陈旧的卸载项（开发版残留）：" + sub);
                                }
                                catch (Exception ex) { say("删除卸载项失败（忽略）：" + ex.Message); }
                            }
                        }
                    }
                }
                catch { }
            }
        }

        /// <summary>清掉 VSTA\Solutions 里本产品的记录（"已安装另一版本"的判定来源）。</summary>
        private static void ClearVstaRecords(Action<string> say)
        {
            try
            {
                using (var sol = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VSTA\Solutions"))
                {
                    if (sol == null) return;
                    foreach (var sub in sol.GetSubKeyNames())
                    {
                        using (var k = sol.OpenSubKey(sub))
                        {
                            if (k == null) continue;
                            var product = k.GetValue("ProductName") as string;
                            var url = k.GetValue("Url") as string;
                            var appId = k.GetValue("ApplicationID") as string;
                            bool mine = (!string.IsNullOrEmpty(product) &&
                                         product.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) >= 0)
                                        || (!string.IsNullOrEmpty(url) &&
                                            url.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) >= 0)
                                        || (!string.IsNullOrEmpty(appId) &&
                                            appId.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) >= 0);
                            if (!mine) continue;
                            try
                            {
                                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\VSTA\Solutions\" + sub, false);
                                say("已删除 VSTA 记录（『已安装另一版本』的判定来源）：" + sub);
                            }
                            catch (Exception ex) { say("删除 VSTA 记录失败：" + ex.Message); }
                        }
                    }
                }
            }
            catch { }
        }

        /// <summary>执行一条命令行（按首个 ".exe" 切分，避免路径里的空格被切错）；输出走 say。</summary>
        private static void RunCommandLineStatic(string cmd, Action<string> say, int timeoutMs = 5 * 60 * 1000)
        {
            int i = cmd.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
            string exe, args;
            if (i > 0)
            {
                exe = cmd.Substring(0, i + 4).Trim().Trim('"');
                args = cmd.Substring(i + 4).Trim();
            }
            else { exe = cmd; args = ""; }
            say("执行：" + exe + " " + args);
            var p = Process.Start(new ProcessStartInfo(exe, args) { UseShellExecute = false });
            if (p == null) { say("未能启动进程"); return; }
            if (!p.WaitForExit(timeoutMs))
            {
                say("超时未结束（" + (timeoutMs / 1000) + "s），放弃等待并结束该进程");
                try { p.Kill(); } catch { }
                return;
            }
            say("退出码：" + p.ExitCode);
        }

        /// <summary>
        /// 清理**本产品**损坏的 ClickOnce 部署状态。
        ///
        /// 为什么需要（实测踩过的坑）：
        ///   VSTOInstaller 报 -400，事件日志里是
        ///     System.Deployment.Application.DeploymentException: 存储元数据"UpdateSkipTime"无效
        ///        在 ClickOnceAddInDeploymentManager.GetManifests(...)
        ///        在 ClickOnceAddInDeploymentManager.InstallAddIn()
        ///   即 ClickOnce 在 HKCU\...\Deployment\SideBySide\2.0 下给本产品存的元数据坏了
        ///   （PackageMetadata 里的 UpdateSkipTime 成了 0001/01/01），此后**任何位置都装不进去**，
        ///   表现就是"更新一直失败 / 插件装不上"。
        ///
        /// 处理：把 SideBySide\2.0 下 **只属于本产品**（按公钥 token 或 ppts.. 前缀识别）的
        /// PackageMetadata / Installations / Components / Marks / StateManager 等条目删掉，
        /// 让 ClickOnce 重新生成。**不动**其它产品的记录。
        /// </summary>
        /// <summary>
        /// 清理历史遗留的安装状态（供本地安装模式调用）：
        /// 老的 ClickOnce 部署记录、VSTA\Solutions、指向开发路径的 ARP、损坏的 SideBySide 元数据。
        /// 迁到 |vstolocal 本地安装后这些都不再需要；留着会干扰（"已安装另一版本"、-400 等）。
        /// 返回清理条数。
        /// </summary>
        public static int CleanLegacyState(Action<string> say)
        {
            int n = 0;
            try { n += ClearBrokenClickOnceState(say); } catch { }
            try { ClearVstaRecords(say); n++; } catch { }
            try { ClearDevArpRecords(say); } catch { }
            // 老的 ClickOnce 卸载项（指向 VSTOInstaller 的）也一并清掉，避免控制面板出现两个条目
            try
            {
                foreach (var hive in new[] { Registry.CurrentUser, Registry.LocalMachine })
                {
                    using (var root = hive.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
                    {
                        if (root == null) continue;
                        foreach (var sub in root.GetSubKeyNames())
                        {
                            if (string.Equals(sub, "PPTShapeLibrary", StringComparison.OrdinalIgnoreCase)) continue;
                            using (var k = root.OpenSubKey(sub))
                            {
                                if (k == null) continue;
                                var name = k.GetValue("DisplayName") as string;
                                if (string.IsNullOrEmpty(name)) continue;
                                if (name.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) < 0) continue;
                                var un = k.GetValue("UninstallString") as string;
                                if (string.IsNullOrEmpty(un)) continue;
                                if (un.IndexOf("VSTOInstaller", StringComparison.OrdinalIgnoreCase) < 0) continue;
                                try
                                {
                                    hive.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + sub, false);
                                    say("已删除旧 ClickOnce 卸载项：" + sub);
                                    n++;
                                }
                                catch { }
                            }
                        }
                    }
                }
            }
            catch { }
            return n;
        }

        private static int ClearBrokenClickOnceState(Action<string> say)
        {
            int deleted = 0;
            try
            {
                const string basePath = @"Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0";
                string[] sections =
                {
                    "PackageMetadata", "Installations", "Components", "Marks",
                    "StateManager", "Visibility", "VisibilityRoots", "Categories", "Assemblies"
                };

                using (var root = Registry.CurrentUser.OpenSubKey(basePath, true))
                {
                    if (root == null) return 0;
                    foreach (var sec in sections)
                    {
                        using (var s = root.OpenSubKey(sec, true))
                        {
                            if (s == null) continue;
                            foreach (var name in CollectOurs(s, sec))
                            {
                                try { s.DeleteSubKeyTree(name, false); deleted++; } catch { }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { say("清理 ClickOnce 状态异常（忽略）：" + ex.Message); }

            // 顺带删掉本产品在本地的部署缓存目录（Apps\2.0 里的散列目录）
            try
            {
                string apps = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Apps", "2.0");
                if (Directory.Exists(apps))
                {
                    foreach (var d in Directory.GetDirectories(apps))
                    {
                        string leaf = Path.GetFileName(d);
                        string probe = Path.Combine(d, "manifests");
                        if (!Directory.Exists(probe)) continue;
                        bool ours = false;
                        foreach (var f in Directory.GetFiles(probe))
                        {
                            if (Path.GetFileName(f).IndexOf("ppt", StringComparison.OrdinalIgnoreCase) >= 0) { ours = true; break; }
                        }
                        if (ours)
                        {
                            try { Directory.Delete(d, true); deleted++; say("已删除本地部署缓存：" + leaf); } catch { }
                        }
                    }
                }
            }
            catch { }

            return deleted;
        }

        /// <summary>收集某个 section 下"属于本产品"的**顶层**子键名（递归找出后归到顶层）。</summary>
        private static List<string> CollectOurs(RegistryKey section, string sectionName)
        {
            var top = new List<string>();
            foreach (var name in section.GetSubKeyNames())
            {
                if (IsOurs(name)) { top.Add(name); continue; }
                using (var sub = section.OpenSubKey(name))
                {
                    if (sub == null) continue;
                    if (HasOursDeep(sub)) top.Add(name);
                }
            }
            return top;
        }

        private static bool HasOursDeep(RegistryKey k)
        {
            foreach (var n in k.GetSubKeyNames())
            {
                if (IsOurs(n)) return true;
                using (var sub = k.OpenSubKey(n))
                {
                    if (sub != null && HasOursDeep(sub)) return true;
                }
            }
            return false;
        }

        /// <summary>名字是否属于本产品：ClickOnce 用公钥 token 做后缀，另有 ppts.. 前缀。</summary>
        private static bool IsOurs(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            if (name.IndexOf("a8a7103479c94f3f", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (name.StartsWith("ppts..", StringComparison.OrdinalIgnoreCase)) return true;
            if (name.IndexOf("PPTShapeLibrary", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        /// <summary>PowerPoint 是否**真的**还在用。安装/卸载期间它会锁住加载项 DLL 与清单，
        /// 导致 VSTOInstaller 半途失败、留下半注册状态。
        ///
        /// ⚠️ 不能只看进程名：实测见过"僵尸 PowerPoint 进程"——用户关掉 PPT 窗口后进程没退干净，
        /// 无主窗口（MainWindowHandle==0）却有近百个线程常驻，能存活几小时。
        /// 只按进程名判断的后果：安装器一直等它退出（等满 180 秒超时），
        /// 用户明明"任务管理器里没有 PPT"却只能手动关掉安装器。
        ///
        /// 判定：**有主窗口**才算真在用；无窗口的孤儿进程不算（也不去杀它，避免误伤）。</summary>
        private static bool IsPowerPointRunning()
        {
            return CountRealPowerPoint() > 0;
        }

        /// <summary>返回"真正在用"的 PowerPoint 进程数（有主窗口的那些）。</summary>
        private static int CountRealPowerPoint()
        {
            try
            {
                int n = 0;
                foreach (var p in Process.GetProcesses())
                {
                    try
                    {
                        if (!string.Equals(p.ProcessName, "POWERPNT", StringComparison.OrdinalIgnoreCase)) continue;
                        if (p.MainWindowHandle != IntPtr.Zero) n++;
                    }
                    catch { }
                    finally { try { p.Dispose(); } catch { } }
                }
                return n;
            }
            catch { return 0; }
        }

        /// <summary>无主窗口的孤儿 PowerPoint 进程数（仅用于日志/提示）。</summary>
        private static int CountZombiePowerPoint()
        {
            try
            {
                int n = 0;
                foreach (var p in Process.GetProcesses())
                {
                    try
                    {
                        if (!string.Equals(p.ProcessName, "POWERPNT", StringComparison.OrdinalIgnoreCase)) continue;
                        if (p.MainWindowHandle == IntPtr.Zero) n++;
                    }
                    catch { }
                    finally { try { p.Dispose(); } catch { } }
                }
                return n;
            }
            catch { return 0; }
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
                // ⚠️ 必须带 --quiet：不带的话被调起的卸载器**自己会弹一个「已卸载」MessageBox**，
                // 而 RunCommandLine 要等它退出 → 那个框若被主窗口挡住，用户看到的就是
                // "点完卸载就没反应"（实测日志里出现过 74 秒的等待），误以为卡死。
                string cmd = _existingUninstallString ?? "";
                if (cmd.IndexOf("--quiet", StringComparison.OrdinalIgnoreCase) < 0) cmd = cmd.TrimEnd() + " --quiet";
                Log("执行卸载：" + cmd);
                RunCommandLine(cmd);
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
            string installPath = null;
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
                            if (v != null)
                            {
                                ver = v.ToString();
                                installPath = k.GetValue("InstallPath") as string;
                                break;
                            }
                        }
                    }
                }
                catch { }
            }

            if (string.IsNullOrEmpty(ver))
            {
                it.State = CheckState.Missing;
                it.Detail = "未安装 VSTO Runtime 4.0（缺失时插件无法被 PowerPoint 加载）";
                Log("✗ 未检测到 VSTO Runtime 4.0");
                return;
            }

            // 光看注册表不够：实测卸载 VSTO 后（未重启）
            //   · msiexec 退出码 0、控制面板卸载项清零、32 位视图键移除
            //   · **但 64 位视图键残留**，且运行时文件**也还在**（并没有被安排重启删除）
            // 所以"键 + 文件"都还在 → 只查这两项会**误报已安装**，用户点开始安装后才发现插件加载不了。
            // 以**控制面板卸载项（MSI 产品注册）**为准最可靠：装了就有，卸了就没。
            bool arpFound = IsVstoListedInPrograms();

            // 运行时文件校验作为辅助：卸载项在但文件不在 → 安装坏了
            string dir = installPath;
            if (string.IsNullOrEmpty(dir))
            {
                // 32 位视图的键没有 InstallPath，退回默认位置
                dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles),
                    "Microsoft Shared", "VSTO", "10.0");
            }
            bool loaderOk = false;
            try { loaderOk = File.Exists(Path.Combine(dir, "VSTOLoader.dll")); } catch { }

            if (!arpFound)
            {
                it.State = CheckState.Missing;
                it.Detail = "VSTO Runtime 已卸载（注册表项残留版本 " + ver +
                            "，但控制面板里已无该产品）。请重启计算机完成清理，或点下面的按钮重新安装 VSTO Runtime";
                Log("✗ VSTO Runtime 已卸载：控制面板卸载项不存在（注册表残留版本 " + ver + "）");
                return;
            }
            if (!loaderOk)
            {
                it.State = CheckState.Missing;
                it.Detail = "VSTO Runtime 安装不完整：控制面板有该产品，但运行时文件不在 " + dir +
                            "。请重新安装 VSTO Runtime（缺失时插件无法被 PowerPoint 加载）";
                Log("✗ VSTO Runtime 文件缺失（" + dir + "）");
                return;
            }

            it.State = CheckState.Ok;
            it.Detail = "已安装 VSTO Runtime（版本 " + ver + "）";
            Log("✓ VSTO Runtime：" + ver);
        }

        /// <summary>控制面板"程序和功能"里是否还列着 VSTO Runtime（MSI 产品注册是否还在）。</summary>
        private static bool IsVstoListedInPrograms()
        {
            try
            {
                foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                {
                    using (var root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view)
                            .OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall"))
                    {
                        if (root == null) continue;
                        foreach (var sub in root.GetSubKeyNames())
                        {
                            using (var k = root.OpenSubKey(sub))
                            {
                                if (k == null) continue;
                                var name = k.GetValue("DisplayName") as string;
                                if (!string.IsNullOrEmpty(name) &&
                                    name.IndexOf("Visual Studio 2010 Tools for Office", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            return false;
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

        /// <summary>在「离线安装包」目录里找离线包。pattern 可以是用 `;` 分隔的多个通配
        /// （如 `vstor*.exe;vstor*.zip` —— 有的包是 zip 形式的官方负载）。</summary>
        private static string FindLocalPackage(string pattern)
        {
            if (string.IsNullOrEmpty(pattern)) return null;
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string[] dirs = { Path.Combine(baseDir, "离线安装包"), baseDir };
                foreach (var p in pattern.Split(';'))
                {
                    var pat = p.Trim();
                    if (pat.Length == 0) continue;
                    foreach (var d in dirs)
                    {
                        if (!Directory.Exists(d)) continue;
                        var files = Directory.GetFiles(d, pat);
                        if (files.Length > 0) return files[0];
                    }
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
                    Log("正在安装本地包：" + Path.GetFileName(it.LocalFile));
                    // 参数必须按"包的形态"取：zip 里的 install.exe 要 /q，
                    // 而官方 vstor_redist.exe 要 /quiet /norestart——混用会得到退出码 4096（参数无效、什么都没做）。
                    var pk = MatchPackage(it, it.LocalFile);
                    string lkind = pk != null ? pk.Kind
                        : Path.GetExtension(it.LocalFile).TrimStart('.').ToLowerInvariant();
                    string largs = pk != null ? pk.Args : it.InstallArgs;
                    if (lkind == "zip") Log("（zip 形式：解压后按 " + largs + " 运行其中的安装程序）");
                    RunPrereqInstaller(it.LocalFile, lkind, largs);
                    Log("本地包安装已结束，正在重新检测…");
                    RefreshChecks();
                    Log(SummaryLine());
                    MessageBox.Show(SummaryLine() + "\r\n\r\n详情见下方日志。部分组件需要重启计算机后生效。",
                        "安装结束", MessageBoxButtons.OK,
                        BlockedByRequired() ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                    return;
                }

                // ① 仓库固定 tag 上的安装包（可走加速站）——微软官方源在国内基本打不开，这是首选
                if (it.Packages != null && it.Packages.Length > 0)
                {
                    var r = MessageBox.Show(
                        "将自动下载安装包并安装（直连优先，失败自动走 GitHub 加速站）：\r\n" +
                        string.Join("\r\n", Array.ConvertAll(it.Packages, x => "  · " + x.FileName)) +
                        "\r\n\r\n来源：本项目的 GitHub Release（固定 tag " + UpdateChecker.PrereqTag +
                        "）。过程中可能弹出 UAC 授权提示，请允许。是否继续？",
                        "下载安装 " + it.Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (r != DialogResult.OK) return;

                    PrereqAsset asset = null;
                    string localPkg = null;
                    foreach (var cand in it.Packages)
                    {
                        try
                        {
                            var dest = Path.Combine(Path.GetTempPath(), cand.FileName);
                            Log("下载：" + cand.FileName);
                            string why;
                            var st = UpdateChecker.TryDownload(UpdateChecker.PrereqBaseUrl + cand.FileName, dest,
                                (got, total) =>
                                {
                                    // 只在每 25% 打一次，避免刷屏
                                    if (total > 0 && got * 100 / total % 25 == 0)
                                        Log("  " + (got * 100 / total) + "%");
                                }, out why);
                            if (st == DownloadStatus.Ok)
                            {
                                localPkg = dest; asset = cand; break;
                            }
                            Log("  " + (st == DownloadStatus.NotFound
                                    ? "跳过（仓库里没有这个资产）"
                                    : "失败：" + why));
                        }
                        catch (Exception dex)
                        {
                            Log("  下载失败：" + dex.Message);
                        }
                    }
                    if (localPkg == null)
                    {
                        var fb = MessageBox.Show(
                            "从 GitHub（含加速站）下载失败。\r\n\r\n是否打开官方下载页手动下载？",
                            "下载失败", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (fb == DialogResult.Yes)
                            Process.Start(new ProcessStartInfo(it.DownloadUrl) { UseShellExecute = true });
                        return;
                    }

                    Log("下载完成，开始安装：" + asset.FileName);
                    RunPrereqInstaller(localPkg, asset.Kind, asset.Args);
                    Log("安装程序已结束，正在重新检测…");
                    RefreshChecks();
                    Log(SummaryLine());
                    MessageBox.Show(SummaryLine() + "\r\n\r\n详情见下方日志。部分组件需要重启计算机后生效。",
                        "安装结束", MessageBoxButtons.OK,
                        BlockedByRequired() ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                    return;
                }

                // ② 没有托管包 → 打开官方页（原来的行为），并提示离线方式
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

        /// <summary>把本地文件与 Packages 里的候选对上（用于取正确的 Kind/Args）。</summary>
        private static PrereqAsset MatchPackage(ReqItem it, string localFile)
        {
            if (it.Packages == null) return null;
            var name = Path.GetFileName(localFile);
            foreach (var c in it.Packages)
            {
                if (string.Equals(c.FileName, name, StringComparison.OrdinalIgnoreCase)) return c;
            }
            // 名字对不上时按扩展名兜个底
            var ext = Path.GetExtension(localFile).TrimStart('.').ToLowerInvariant();
            foreach (var c in it.Packages)
            {
                if (string.Equals(c.Kind, ext, StringComparison.OrdinalIgnoreCase)) return c;
            }
            return null;
        }

        private bool BlockedByRequired()
        {
            foreach (var x in _items) if (x.Required && x.State != CheckState.Ok) return true;
            return false;
        }

        private string SummaryLine()
        {
            var bad = new List<string>();
            foreach (var x in _items) if (x.Required && x.State != CheckState.Ok) bad.Add(x.Title);
            return bad.Count == 0
                ? "环境检测通过，可以开始安装。"
                : "仍缺失必需组件：" + string.Join("、", bad.ToArray()) + "。请按上面的提示处理后再试。";
        }

        /// <summary>运行依赖安装程序：exe 直接跑；msi 走 msiexec；zip 解压后跑里面的 exe/msi。</summary>
        private void RunPrereqInstaller(string path, string kind, string args)
        {
            string exe = path, exeArgs = args ?? "";

            if (string.Equals(kind, "zip", StringComparison.OrdinalIgnoreCase))
            {
                string dir = Path.Combine(Path.GetTempPath(), "pptsl_prereq_" + Guid.NewGuid().ToString("N").Substring(0, 8));
                Directory.CreateDirectory(dir);
                System.IO.Compression.ZipFile.ExtractToDirectory(path, dir);
                Log("已解压到：" + dir);
                // 优先装 exe（bootstrapper），其次 msi
                var exes = Directory.GetFiles(dir, "*.exe", SearchOption.AllDirectories);
                var msis = Directory.GetFiles(dir, "*.msi", SearchOption.AllDirectories);
                if (exes.Length > 0) { exe = exes[0]; exeArgs = args ?? ""; }
                else if (msis.Length > 0) { exe = msis[0]; }
                else throw new Exception("压缩包里没有找到 .exe 或 .msi");
            }

            ProcessStartInfo psi;
            if (exe.EndsWith(".msi", StringComparison.OrdinalIgnoreCase))
            {
                psi = new ProcessStartInfo("msiexec.exe", "/i \"" + exe + "\" /qn /norestart");
            }
            else
            {
                psi = new ProcessStartInfo(exe, exeArgs);
            }
            psi.UseShellExecute = true;
            psi.Verb = "runas";
            Log("执行：" + psi.FileName + " " + psi.Arguments);
            var p = Process.Start(psi);
            if (p == null)
            {
                Log("⚠ 未能启动安装程序（可能被 UAC 取消）");
                return;
            }
            p.WaitForExit();
            int code = p.ExitCode;
            Log("退出码：" + code);
            if (code != 0 && code != 3010 && code != 1641)
            {
                // 实测：给 install.exe 传错参数会得到 4096（=0x1000，"参数无效"），看起来"成功"其实什么都没做
                Log("⚠ 退出码非 0（0=成功 3010=成功但需重启）。若是 4096，通常说明安装参数与这个包不匹配。");
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

        /// <summary>
        /// 【单文件包】从内嵌载荷安装：解到临时目录 → 清理历史残留 → 装到 %LOCALAPPDATA% → 写注册。
        /// 全程在**后台线程**执行（解压约 8MB + 拷贝数百个文件，放 UI 线程会明显卡顿），
        /// 日志通过 SynchronizationContext 回填到界面。
        /// </summary>
        private void InstallFromEmbeddedPayload()
        {
            SetBusy(true);
            Log("检测到内嵌载荷（单文件包）→ 完全自包含安装，不需要 App\\ 或 离线安装包\\");

            var ui = System.Threading.SynchronizationContext.Current;
            var t = new System.Threading.Thread(() =>
            {
                bool ok = false;
                string err = null, ver = null;
                Action<string> say = m =>
                {
                    try
                    {
                        if (ui != null) ui.Post(_ => Log(m), null);
                        else { try { BeginInvoke(new Action(() => Log(m))); } catch { } }
                    }
                    catch { }
                };
                try
                {
                    // 与 CLI（--install-local）**共用**同一段安装逻辑，避免两条路径行为不一致
                    // （旧 bug：CLI 能用内嵌载荷，GUI 却只找同目录的 App\ / 离线安装包\）
                    ver = InstallFlow.InstallFromEmbedded(say);
                    ok = true;
                }
                catch (Exception ex)
                {
                    err = ex.Message;
                    say("✗ 安装失败：" + ex.Message);
                }
                finally
                {
                    // 临时载荷目录由 InstallFlow 自己清理；这里只兜底
                }

                Action done = () =>
                {
                    try { DetectExistingInstall(); UpdateExistingUi(); UpdateSummary(); } catch { }
                    SetBusy(false);
                    if (ok)
                        MessageBox.Show(
                            "安装完成。\r\n\r\n" +
                            "打开 PowerPoint，功能区应出现「素材库」选项卡。\r\n" +
                            "若未出现：请先完全退出 PowerPoint 再重新打开。\r\n\r\n" +
                            "插件目录：" + LocalInstall.VersionDir(ver),
                            "安装完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MessageBox.Show("安装失败：\r\n\r\n" + err, "安装失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                };
                try
                {
                    if (ui != null) ui.Post(_ => done(), null);
                    else { try { BeginInvoke(done); } catch { } }
                }
                catch { }
            });
            t.IsBackground = true;
            t.Start();
        }

        /// <summary>忙碌态：禁按钮 + 等待光标，避免用户重复点击、也让人知道程序在干活。</summary>
        private void SetBusy(bool busy)
        {
            foreach (var b in new[] { _install, _uninstallOld, _repair, _reinstall, _checkUpdate })
            {
                try { if (b != null) b.Enabled = !busy; } catch { }
            }
            try { UseWaitCursor = busy; } catch { }
            try { if (busy) _summary.Text = "正在安装…（请勿关闭本窗口）"; } catch { }
        }

        private void StartInstall()
        {
            try
            {
                // PowerPoint 在运行会锁住加载项 DLL/清单，导致 VSTOInstaller 半途失败、留下半注册状态。
                // 只把**有窗口**的 PowerPoint 算作"正在运行"：无窗口的孤儿进程不锁文件，
                // 没必要因此拦住用户（实测：任务管理器里看不到 PPT，却提示"请关闭 PowerPoint"）。
                int zombiePpt = CountZombiePowerPoint();
                if (zombiePpt > 0)
                    Log("提示：有 " + zombiePpt + " 个无窗口的 PowerPoint 残留进程（不拦截安装）");

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

                // 【单文件包】自带内嵌载荷 → 完全自包含安装，**不依赖同目录的 App\ 和 离线安装包\**。
                //
                // 曾经的 bug：单 exe 里明明带着载荷，但 GUI 的安装流程仍然只会去同目录找
                // App\PPTShapeLibrary.vsto / 离线安装包\，找不到就弹"完整分发目录应包含…"，
                // 于是"我只发了一个 exe"的用户装不上。内嵌载荷流程原先只在 CLI
                // （--install-local）里实现，GUI 没接上。
                if (LocalInstall.HasPayload)
                {
                    InstallFromEmbeddedPayload();
                    return;
                }

                // 下面都是**旧的分发目录形态**（App\ + Application Files\ + 离线安装包\）的回退路径。
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

                // 装之前无条件清掉会让 VSTO 拒绝安装的残留（VSTA 记录 / 开发版 ARP）。
                // 不清的话，即使 Addins 键和 ARP 都没了，VSTOInstaller 仍会报
                // "已安装另一版本的自定义项"（见 docs/PowerPoint-COM经验与坑.md §9.8）。
                ClearVstaRecords(Log);
                ClearDevArpRecords(Log);
                // 装之前也清一次损坏的 ClickOnce 状态（正常机器上是空操作）
                int cocPre = ClearBrokenClickOnceState(Log);
                if (cocPre > 0) Log("已清理 " + cocPre + " 条损坏的 ClickOnce 记录");

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
        internal static string FindDeploymentManifest()
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
            // 回退：**已安装**的清单。单文件包（本地安装）把文件放在 addin\<版本>\，
            // 自己所在目录下永远不会有 App\ —— 不回退的话「修复加载项注册」会误报"缺少安装包"。
            return LocalInstall.InstalledManifestPath();
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

        /// <summary>写日志：界面里显示，同时**落盘**（%TEMP%\pptsl_installer.log）——
        /// 界面上那个框会滚动、也不好整段复制，出问题直接把这个文件发出来最省事。</summary>
        private void Log(string msg)
        {
            var line = DateTime.Now.ToString("HH:mm:ss") + "  " + msg + Environment.NewLine;
            try { _log.AppendText(line); } catch { }
            try
            {
                var path = Path.Combine(Path.GetTempPath(), "pptsl_installer.log");
                // 超过约 1MB 就重开，避免无限增长
                if (File.Exists(path) && new FileInfo(path).Length > 1024 * 1024) File.Delete(path);
                File.AppendAllText(path, line, System.Text.Encoding.UTF8);
            }
            catch { }
        }
    }
}
