using System;
using System.IO;
using System.Windows.Forms;

namespace PPTShapeLibraryInstaller
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            // ---- 新增：本地安装模式（不再依赖 ClickOnce/VSTOInstaller）----
            //   --install-local [--dir <解压好的载荷目录>] [--version x.y.z.w] [--quiet]
            //   不给 --dir 时用安装器**内嵌的载荷**（单 exe 分发就是这条路）。
            //   行为：放文件到 %LOCALAPPDATA%\PPTShapeLibrary\addin\<版本>\ + 写 2 个加载项键 + 写卸载项。
            //   **不需要关闭 PowerPoint**（只改注册表指向，运行中的实例继续用旧目录，下次启动生效）
            if (args != null && Array.IndexOf(args, "--install-local") >= 0)
            {
                Environment.Exit(RunInstallLocal(args));
            }

            // --uninstall [--quiet]：删键 + 删目录 + 清残留（控制面板卸载项调它）
            if (args != null && Array.IndexOf(args, "--uninstall") >= 0)
            {
                Environment.Exit(RunUninstall(args));
            }

            // --payload-extract <目标目录>：只把内嵌载荷解出来（打包/排错用）
            if (args != null && args.Length >= 2 && args[0] == "--payload-extract")
            {
                try
                {
                    LocalInstall.ExtractPayload(args[1], s => Console.WriteLine(s));
                    Environment.Exit(0);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex.Message);
                    Environment.Exit(1);
                }
            }

            // 打印机模式：报告内嵌载荷信息（打包脚本自检用）
            if (args != null && Array.IndexOf(args, "--payload-info") >= 0)
            {
                Console.WriteLine("HasPayload=" + LocalInstall.HasPayload);
                Environment.Exit(LocalInstall.HasPayload ? 0 : 1);
            }

            // 提权子模式：由「修复加载项注册」以 runas 拉起，只写 HKLM 的加载项注册后立即退出。
            // 不弹 UI，避免用户看到第二个窗口；退出码 0=写入成功。
            if (args != null && args.Length >= 2 && args[0] == "--repair-machine")
            {
                int n = InstallerForm.WriteMachineAddinKeys(args[1]);
                Environment.Exit(n > 0 ? 0 : 1);
            }

            // 无人值守更新子模式：由**插件**拉起（用户在 PowerPoint 里点了「立即更新」并关闭 PowerPoint）。
            // 参数：--pending-update <新版解压目录>
            // 行为：等 PowerPoint 退出 → 卸载旧版 → 装新版 → 结果写 %TEMP%\pptsl_pending_update.log。
            // 全程不弹 UI；退出码 0=成功。
            if (args != null && args.Length >= 2 && args[0] == "--pending-update")
            {
                int rc = InstallerForm.RunPendingUpdate(args[1]);
                Environment.Exit(rc);
            }

            // 命令行模式：静默检查更新（便于脚本化验证 / 定时任务）
            //   --check-update            只查并打印结果
            //   --check-update --download 再下载并校验 SHA256（不安装）
            //   --check-update --apply    下载校验后解压并启动新版安装器
            if (args != null && Array.IndexOf(args, "--check-update") >= 0)
            {
                string outFile = ArgValue(args, "--out");
                return RunCheckUpdate(
                    Array.IndexOf(args, "--download") >= 0,
                    Array.IndexOf(args, "--apply") >= 0,
                    outFile);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new InstallerForm());
            return 0;
        }

        private static string ArgValue(string[] args, string name)
        {
            int i = Array.IndexOf(args, name);
            if (i >= 0 && i + 1 < args.Length && !args[i + 1].StartsWith("--")) return args[i + 1];
            return null;
        }

        /// <summary>版本号：优先 --version，其次取载荷里的 vsto/程序集版本，最后退回调用方给的默认。</summary>
        private static string ResolveVersion(string[] args, string payloadDir)
        {
            string v = ArgValue(args, "--version");
            if (!string.IsNullOrEmpty(v)) return v;

            // 从载荷目录里的 PPTShapeLibrary.dll 读文件版本（最可靠）
            try
            {
                if (!string.IsNullOrEmpty(payloadDir))
                {
                    string dll = Path.Combine(payloadDir, "PPTShapeLibrary.dll");
                    if (File.Exists(dll))
                    {
                        var fv = System.Diagnostics.FileVersionInfo.GetVersionInfo(dll).FileVersion;
                        if (!string.IsNullOrEmpty(fv)) return fv;
                    }
                }
            }
            catch { }

            // 从自己所在目录找（安装器与载荷同级时）
            try
            {
                string dll = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PPTShapeLibrary.dll");
                if (File.Exists(dll))
                {
                    var fv = System.Diagnostics.FileVersionInfo.GetVersionInfo(dll).FileVersion;
                    if (!string.IsNullOrEmpty(fv)) return fv;
                }
            }
            catch { }

            return "0.0.0.0";
        }

        private static int RunInstallLocal(string[] args)
        {
            string logPath = Path.Combine(Path.GetTempPath(), "pptsl_local_install.log");
            Action<string> say = m =>
            {
                string line = DateTime.Now.ToString("HH:mm:ss") + "  " + m;
                try { File.AppendAllText(logPath, line + Environment.NewLine, System.Text.Encoding.UTF8); } catch { }
                try { Console.WriteLine(line); } catch { }
            };
            try { File.Delete(logPath); } catch { }

            bool quiet = Array.IndexOf(args, "--quiet") >= 0;
            try
            {
                string dir = ArgValue(args, "--dir");
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    say("✗ 载荷目录不存在：" + dir);
                    if (!quiet) MessageBox.Show("载荷目录不存在：" + dir, "安装失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 2;
                }

                string stage = dir;
                bool tempStage = false;
                if (string.IsNullOrEmpty(stage))
                {
                    // 单 exe：先把内嵌载荷解到临时目录
                    stage = Path.Combine(Path.GetTempPath(), "pptsl_payload_" + Guid.NewGuid().ToString("N").Substring(0, 8));
                    say("解出自带载荷 → " + stage);
                    LocalInstall.ExtractPayload(stage, say);
                    tempStage = true;
                }

                string version = ResolveVersion(args, stage);
                say("版本：" + version);

                say("正在结束可能残留的 PowerPoint 安装记录（清理 ClickOnce/VSTA 残留）…");
                int cleaned = InstallerForm.CleanLegacyState(say);
                if (cleaned > 0) say("已清理 " + cleaned + " 条历史残留");

                int n = LocalInstall.Install(version, stage, say);
                say("✓ 安装成功：" + LocalInstall.VersionDir(version) + "（注册 " + n + " 处）");

                if (tempStage)
                {
                    try { Directory.Delete(stage, true); } catch { }
                }

                if (!quiet)
                {
                    MessageBox.Show(
                        "安装完成。\r\n\r\n" +
                        "打开 PowerPoint，功能区应出现「素材库」选项卡。\r\n" +
                        "若没有出现：请先完全退出 PowerPoint 再重新打开；仍没有就运行安装器的「修复加载项注册」。\r\n\r\n" +
                        "插件目录：" + LocalInstall.VersionDir(version),
                        "安装完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return 0;
            }
            catch (Exception ex)
            {
                say("✗ 安装失败：" + ex.Message);
                if (!quiet) MessageBox.Show("安装失败：\r\n\r\n" + ex.Message, "安装失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        private static int RunUninstall(string[] args)
        {
            string logPath = Path.Combine(Path.GetTempPath(), "pptsl_local_install.log");
            Action<string> say = m =>
            {
                string line = DateTime.Now.ToString("HH:mm:ss") + "  " + m;
                try { File.AppendAllText(logPath, line + Environment.NewLine, System.Text.Encoding.UTF8); } catch { }
                try { Console.WriteLine(line); } catch { }
            };

            bool quiet = Array.IndexOf(args, "--quiet") >= 0;
            try
            {
                LocalInstall.Uninstall(say);
                if (!quiet)
                    MessageBox.Show("已卸载 PPTShapeLibrary。\r\n\r\n（你的素材库数据仍在 %APPDATA%\\PPTShapeLibrary，未删除）",
                        "卸载完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            catch (Exception ex)
            {
                say("✗ 卸载失败：" + ex.Message);
                if (!quiet) MessageBox.Show("卸载失败：\r\n\r\n" + ex.Message, "卸载失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        /// <summary>
        /// 静默检查更新。退出码：0=已是最新；1=检查/下载失败；2=有新版（便于脚本判断）。
        ///
        /// 注意：本程序是 WinExe，从 PowerShell 捕获时 stdout 常常拿不到，
        /// 所以结果**同时**写一份到文件（默认 %TEMP%\pptsl_update_check.log，可用 --out 指定）。
        /// </summary>
        private static int RunCheckUpdate(bool download, bool apply, string outFile)
        {
            if (string.IsNullOrEmpty(outFile))
                outFile = Path.Combine(Path.GetTempPath(), "pptsl_update_check.log");
            var log = new System.Text.StringBuilder();
            Action<string> say = s => { Console.WriteLine(s); log.AppendLine(s); };
            Action<string> err = s => { Console.Error.WriteLine(s); log.AppendLine("错误: " + s); };

            int code;
            try
            {
                var info = UpdateChecker.Check(true);
                say("当前版本: " + info.CurrentVersion);
                if (!info.Ok)
                {
                    err(info.Error);
                    code = 1;
                }
                else
                {
                    say("最新版本: " + info.LatestVersion + "  (tag " + info.TagName + ")");
                    say("发布页:   " + info.ReleaseUrl);
                    if (!info.HasUpdate)
                    {
                        say("已是最新。");
                        code = 0;
                    }
                    else
                    {
                        say("资产:     " + info.AssetName);
                        say("下载地址: " + info.AssetUrl);
                        code = 2;
                        if ((download || apply) && !string.IsNullOrEmpty(info.AssetUrl))
                        {
                            var dest = Path.Combine(Path.GetTempPath(), info.AssetName ?? "pptshape-update.zip");
                            say("开始下载（直连优先，失败回退加速站）…");
                            long lastPct = -1;
                            UpdateChecker.Download(info.AssetUrl, dest, (got, total) =>
                            {
                                if (total <= 0) return;
                                long pct = got * 100 / total;
                                if (pct != lastPct && pct % 25 == 0) { say("  " + pct + "%"); lastPct = pct; }
                            });
                            var actual = UpdateChecker.Sha256(dest);
                            var expected = UpdateChecker.FetchExpectedHash(info.Sha256Url);
                            say("已下载:   " + dest + "  (" + new FileInfo(dest).Length + " 字节)");
                            say("sha256:   " + actual);
                            if (!string.IsNullOrEmpty(expected))
                            {
                                say("期望值:   " + expected);
                                if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
                                {
                                    err("校验失败，拒绝安装。");
                                    code = 1;
                                }
                                else say("校验通过。");
                            }
                            else say("（该 Release 没有 .sha256 资产，跳过校验）");

                            if (apply && code == 2)
                            {
                                var p = UpdateChecker.ApplyUpdate(dest);
                                say("已启动新版安装器: " + (p != null ? p.ProcessName : "?"));
                            }
                        }
                        else if (download || apply)
                        {
                            err("该 Release 没有上传分发包资产，无法自动下载。");
                            code = 1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                err(ex.Message);
                code = 1;
            }

            try { File.WriteAllText(outFile, log.ToString(), System.Text.Encoding.UTF8); } catch { }
            return code;
        }
    }
}
