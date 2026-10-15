using System;
using System.IO;
using System.Windows.Forms;

namespace PPTShapeLibraryInstaller
{
    /// <summary>
    /// 安装流程的**共享实现**：CLI（--install-local）与 GUI（安装按钮）都走这里。
    ///
    /// 【为什么要抽出来】曾经 GUI 的 StartInstall 只会在同目录找 App\ / 离线安装包\，
    /// 而"从内嵌载荷安装"只在 CLI 里实现过 —— 于是单文件包交给用户后，
    /// 双击运行会被要求提供 App\ 和 离线安装包\（明明 exe 里就带着载荷）。
    /// 两条路径共用同一段代码后，CLI 自动化测试即可覆盖 GUI 的行为。
    /// </summary>
    internal static class InstallFlow
    {
        /// <summary>
        /// 从**内嵌载荷**安装（单文件包）。返回安装的版本号。
        /// 步骤：解到临时目录 → 读版本 → 清 ClickOnce/VSTA 历史残留 → 装到 %LOCALAPPDATA% → 删临时目录。
        /// 异常向上抛（调用方决定怎么报错）。
        /// </summary>
        public static string InstallFromEmbedded(Action<string> say)
        {
            string stage = Path.Combine(Path.GetTempPath(),
                "pptsl_payload_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            try
            {
                say("解出自带载荷 → " + stage);
                int n = LocalInstall.ExtractPayload(stage, say);
                say("已解出 " + n + " 个文件");
                return InstallFromDir(stage, say);
            }
            finally
            {
                try { if (Directory.Exists(stage)) Directory.Delete(stage, true); } catch { }
            }
        }

        /// <summary>从**已解开的载荷目录**安装。返回安装的版本号。</summary>
        public static string InstallFromDir(string payloadDir, Action<string> say)
        {
            if (string.IsNullOrEmpty(payloadDir) || !Directory.Exists(payloadDir))
                throw new DirectoryNotFoundException("载荷目录不存在：" + payloadDir);

            string ver = Program.ResolveVersion(new string[0], payloadDir);
            say("版本：" + ver);

            say("清理历史残留（ClickOnce / VSTA）…");
            int cleaned = InstallerForm.CleanLegacyState(say);
            if (cleaned > 0) say("已清理 " + cleaned + " 条历史残留");

            int n = LocalInstall.Install(ver, payloadDir, say);
            say("✓ 安装成功：" + LocalInstall.VersionDir(ver) + "（注册 " + n + " 处）");
            return ver;
        }
    }
}
