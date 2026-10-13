using System;
using System.Windows.Forms;

namespace PPTShapeLibraryInstaller
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            // 提权子模式：由「修复加载项注册」以 runas 拉起，只写 HKLM 的加载项注册后立即退出。
            // 不弹 UI，避免用户看到第二个窗口；退出码 0=写入成功。
            if (args != null && args.Length >= 2 && args[0] == "--repair-machine")
            {
                int n = InstallerForm.WriteMachineAddinKeys(args[1]);
                Environment.Exit(n > 0 ? 0 : 1);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new InstallerForm());
        }
    }
}
