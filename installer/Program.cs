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
