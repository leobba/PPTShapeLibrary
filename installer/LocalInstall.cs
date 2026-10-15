using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PPTShapeLibraryInstaller
{
    /// <summary>
    /// 「本地安装」模式：**不用 ClickOnce / VSTOInstaller**，直接把插件文件放好并写注册表。
    ///
    /// 为什么换掉 ClickOnce（本项目的实测结论）：
    ///   ClickOnce 一路踩到这些坑 —— AddInAlreadyInstalledException（已安装另一版本）、
    ///   卸载卡死 120 秒、部署元数据损坏导致 VSTOInstaller 永远 -400、
    ///   安装源目录换位置就装不上……而且每次升级都要"关 PowerPoint 等它退出"。
    ///
    /// 而 <c>Manifest = file:///…\PPTShapeLibrary.vsto|vstolocal</c> 这种注册方式
    /// （VS 按 F5 调试用的就是它）实测**不需要 VSTO 信任项、不需要 ClickOnce**：
    ///   · 安装 = 把文件放到 %LOCALAPPDATA%\PPTShapeLibrary\addin\&lt;版本&gt;\ + 写 2 个键
    ///   · 升级 = 放一个新版本目录 + 改写 2 个键 → **不需要关 PowerPoint**（下次启动生效）
    ///   · 卸载 = 删 2 个键 + 删目录
    /// </summary>
    internal static class LocalInstall
    {
        /// <summary>插件文件放这里（按版本分目录，便于升级时并存/回滚）。</summary>
        public static string AddinRoot
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PPTShapeLibrary", "addin");
            }
        }

        public static string VersionDir(string version)
        {
            return Path.Combine(AddinRoot, version);
        }

        /// <summary>安装器 exe 自己也被复制一份到插件目录，保证控制面板卸载永远可用
        /// （用户可能把下载下来的 exe 删了）。</summary>
        public static string InstalledExePath(string version)
        {
            return Path.Combine(VersionDir(version), "PPTShapeLibrary-卸载修复.exe");
        }

        // ---------------- 载荷（内嵌 zip）----------------

        private const string PayloadResource = "payload.zip";

        /// <summary>把内嵌的插件载荷解压到目标目录。返回解压出的文件数。</summary>
        public static int ExtractPayload(string targetDir, Action<string> say)
        {
            var asm = Assembly.GetExecutingAssembly();
            Stream src = null;
            foreach (var name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith(PayloadResource, StringComparison.OrdinalIgnoreCase))
                {
                    src = asm.GetManifestResourceStream(name);
                    break;
                }
            }
            if (src == null) throw new Exception("安装器内部没有找到插件载荷（payload.zip）—— 文件可能损坏");

            Directory.CreateDirectory(targetDir);
            int count = 0;
            using (src)
            using (var zip = new ZipArchive(src, ZipArchiveMode.Read))
            {
                foreach (var entry in zip.Entries)
                {
                    if (string.IsNullOrEmpty(entry.Name)) continue;   // 目录项
                    string rel = entry.FullName.Replace('/', Path.DirectorySeparatorChar);
                    string dest = Path.Combine(targetDir, rel);
                    Directory.CreateDirectory(Path.GetDirectoryName(dest));
                    entry.ExtractToFile(dest, true);
                    count++;
                }
            }
            say("已解压 " + count + " 个文件到 " + targetDir);
            return count;
        }

        public static bool HasPayload
        {
            get
            {
                var asm = Assembly.GetExecutingAssembly();
                foreach (var name in asm.GetManifestResourceNames())
                    if (name.EndsWith(PayloadResource, StringComparison.OrdinalIgnoreCase)) return true;
                return false;
            }
        }

        // ---------------- 注册 ----------------

        /// <summary>加载项注册的 4 个位置（HKCU/HKLM × 版本无关/版本化）。</summary>
        private static readonly string[][] AddinKeyPaths =
        {
            new[] { "HKCU", @"Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary" },
            new[] { "HKCU", @"Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary" },
            new[] { "HKLM", @"Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary" },
            new[] { "HKLM", @"Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary" }
        };

        private const string UninstallKey =
            @"Software\Microsoft\Windows\CurrentVersion\Uninstall\PPTShapeLibrary";

        /// <summary>
        /// 安装/升级：把载荷放到 addin\&lt;版本&gt;\、写加载项注册（含 ProductVersion，供插件内"检查更新"读）、
        /// 写控制面板卸载项。**不需要关闭 PowerPoint**：只改注册表指向，运行中的实例继续用旧目录。
        /// </summary>
        public static int Install(string version, string payloadSource, Action<string> say)
        {
            string dir = VersionDir(version);
            say("安装目录：" + dir);

            // 1) 放文件
            if (!string.IsNullOrEmpty(payloadSource) && Directory.Exists(payloadSource))
            {
                CopyTree(payloadSource, dir, say);
            }
            else
            {
                ExtractPayload(dir, say);
            }

            string vsto = Path.Combine(dir, "PPTShapeLibrary.vsto");
            if (!File.Exists(vsto)) throw new Exception("安装失败：载荷里没有 PPTShapeLibrary.vsto");

            // 2) 把安装器自身复制进去，卸载/修复不依赖用户手里的下载文件
            try
            {
                string self = Application.ExecutablePath;
                string dst = InstalledExePath(version);
                if (!string.Equals(self, dst, StringComparison.OrdinalIgnoreCase))
                {
                    File.Copy(self, dst, true);
                    say("已复制安装器副本：" + dst);
                }
            }
            catch (Exception ex) { say("复制安装器副本失败（不影响使用）：" + ex.Message); }

            // 3) 写注册
            int n = WriteAddinKeys(dir, version, say);
            WriteUninstallEntry(version, dir, say);

            // 4) 顺手清理**旧版本目录**（best-effort）。
            //    升级到新目录时旧目录的文件通常被运行中的 PowerPoint 锁着，删不掉；
            //    那就留着 —— 下次升级（或用户重启 PowerPoint 后再装）时自然能删掉。
            CleanupOldVersions(version, say);

            say("✓ 安装完成（写入 " + n + " 处加载项注册）");
            return n;
        }

        /// <summary>删掉除 keepVersion 之外的旧版本目录（删不掉就跳过，不影响使用）。</summary>
        private static void CleanupOldVersions(string keepVersion, Action<string> say)
        {
            try
            {
                if (!Directory.Exists(AddinRoot)) return;
                foreach (var d in Directory.GetDirectories(AddinRoot))
                {
                    string name = Path.GetFileName(d);
                    if (string.Equals(name, keepVersion, StringComparison.OrdinalIgnoreCase)) continue;
                    try
                    {
                        Directory.Delete(d, true);
                        say("已清理旧版本目录：" + name);
                    }
                    catch
                    {
                        say("旧版本目录暂被占用，留待以后清理：" + name);
                    }
                }
            }
            catch { }
        }

        private static void CopyTree(string src, string dst, Action<string> say)
        {
            Directory.CreateDirectory(dst);
            int count = 0;
            var locked = new List<string>();
            foreach (var f in Directory.GetFiles(src, "*", SearchOption.AllDirectories))
            {
                string rel = f.Substring(src.Length).TrimStart('\\', '/');
                string target = Path.Combine(dst, rel);
                Directory.CreateDirectory(Path.GetDirectoryName(target));
                try
                {
                    File.Copy(f, target, true);
                    count++;
                }
                catch (IOException)
                {
                    // 目标文件被占用（PowerPoint 正跑着**同一个版本目录**里的加载项时会发生）。
                    // 正常升级是装到**新版本目录**，不会碰到这个；这里只在"重装同版本"时出现。
                    locked.Add(rel);
                }
                catch (UnauthorizedAccessException) { locked.Add(rel); }
            }
            say("已复制 " + count + " 个文件" + (locked.Count > 0 ? "（" + locked.Count + " 个被占用）" : ""));

            if (locked.Count > 0)
            {
                // 让上层决定怎么办：多数文件已就位，但版本目录不完整
                throw new IOException(
                    "有 " + locked.Count + " 个文件被占用（PowerPoint 正在使用同版本的文件）：" +
                    string.Join("、", locked.GetRange(0, Math.Min(3, locked.Count)).ToArray()) +
                    (locked.Count > 3 ? " 等" : "") +
                    "。\r\n请关闭 PowerPoint 后重试（升级到**不同版本**时不需要关闭 PowerPoint）。");
            }
        }

        private static int WriteAddinKeys(string addinDir, string version, Action<string> say)
        {
            string vsto = Path.Combine(addinDir, "PPTShapeLibrary.vsto");
            string manifest = "file:///" + vsto.Replace('\\', '/') + "|vstolocal";
            int written = 0;
            foreach (var item in AddinKeyPaths)
            {
                bool machine = item[0] == "HKLM";
                if (machine && !IsAdmin()) continue;      // HKLM 需要管理员，写不上就跳过（HKCU 够用）
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
                        // 插件内"检查更新"读它作为"当前版本"（优先于 ARP）
                        key.SetValue("ProductVersion", version, RegistryValueKind.String);
                        written++;
                    }
                }
                catch (Exception ex)
                {
                    say("写注册失败（忽略）：" + item[0] + "\\" + item[1] + " → " + ex.Message);
                }
            }
            return written;
        }

        /// <summary>写控制面板「程序和功能」里的卸载项，卸载命令指向插件目录里的安装器副本。</summary>
        private static void WriteUninstallEntry(string version, string dir, Action<string> say)
        {
            try
            {
                string exe = InstalledExePath(version);
                if (!File.Exists(exe)) exe = Application.ExecutablePath;
                long sizeKb = 0;
                try
                {
                    foreach (var f in Directory.GetFiles(dir, "*", SearchOption.AllDirectories))
                        sizeKb += new FileInfo(f).Length / 1024;
                }
                catch { }

                using (var k = Registry.CurrentUser.CreateSubKey(UninstallKey))
                {
                    if (k == null) return;
                    k.SetValue("DisplayName", "PPTShapeLibrary（PowerPoint 素材库插件）", RegistryValueKind.String);
                    k.SetValue("DisplayVersion", version, RegistryValueKind.String);
                    k.SetValue("Publisher", "PPTShapeLibrary", RegistryValueKind.String);
                    k.SetValue("InstallLocation", dir, RegistryValueKind.String);
                    k.SetValue("UninstallString", "\"" + exe + "\" --uninstall", RegistryValueKind.String);
                    k.SetValue("QuietUninstallString", "\"" + exe + "\" --uninstall --quiet", RegistryValueKind.String);
                    k.SetValue("DisplayIcon", exe, RegistryValueKind.String);
                    k.SetValue("EstimatedSize", (int)Math.Min(int.MaxValue, sizeKb), RegistryValueKind.DWord);
                    k.SetValue("NoModify", 1, RegistryValueKind.DWord);
                    k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                }
                say("已写入卸载项（控制面板可卸载）");
            }
            catch (Exception ex) { say("写卸载项失败（忽略）：" + ex.Message); }
        }

        /// <summary>卸载：删注册 + 删所有版本目录 + 清 ClickOnce/VSTA 残留。</summary>
        public static void Uninstall(Action<string> say)
        {
            // 1) 删加载项注册
            foreach (var item in AddinKeyPaths)
            {
                try
                {
                    var root = item[0] == "HKLM" ? Registry.LocalMachine : Registry.CurrentUser;
                    root.DeleteSubKeyTree(item[1], false);
                    say("已删除注册：" + item[0] + "\\" + item[1]);
                }
                catch { }
            }
            // 2) 删 AddinsData 缓存
            foreach (var p in new[]
            {
                @"Software\Microsoft\Office\PowerPoint\AddinsData\PPTShapeLibrary",
                @"Software\Microsoft\Office\16.0\PowerPoint\AddinsData\PPTShapeLibrary"
            })
            {
                try { Registry.CurrentUser.DeleteSubKeyTree(p, false); } catch { }
            }
            // 3) 删卸载项
            try { Registry.CurrentUser.DeleteSubKeyTree(UninstallKey, false); say("已删除卸载项"); } catch { }

            // 4) 删文件目录
            //    注意：卸载器自己就住在 addin\<版本>\ 里（控制面板调用的就是那份副本），
            //    所以**直接删一定失败**（"对路径…的访问被拒绝"）。删不掉就交给一个
            //    分离的 cmd 进程：等我们退出后再删。
            try
            {
                if (Directory.Exists(AddinRoot))
                {
                    try
                    {
                        Directory.Delete(AddinRoot, true);
                        say("已删除插件目录：" + AddinRoot);
                    }
                    catch (Exception ex)
                    {
                        say("直接删除失败（" + ex.Message + "），改由后台进程稍后清理");
                        ScheduleDirectoryDelete(AddinRoot, say);
                    }
                }
            }
            catch { }

            say("✓ 已卸载（素材库数据在 %APPDATA%\\PPTShapeLibrary，未删除）");
        }

        /// <summary>
        /// 派一个**分离**的 cmd 进程，等本进程退出后再删目录。
        /// 卸载器就住在待删目录里，自己删自己一定被拒 —— 这是标准做法。
        /// 用 ping 当 sleep（`timeout` 需要控制台，在隐藏窗口里会立刻报错）。
        /// </summary>
        public static void ScheduleDirectoryDelete(string dir, Action<string> say)
        {
            try
            {
                string cmd = "/c ping -n 3 127.0.0.1 >nul & rmdir /s /q \"" + dir + "\"";
                var psi = new ProcessStartInfo("cmd.exe", cmd)
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                };
                Process.Start(psi);
                say("已安排后台清理：" + dir);
            }
            catch (Exception ex) { say("安排后台清理失败（可手动删除）：" + ex.Message); }
        }

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
    }
}
