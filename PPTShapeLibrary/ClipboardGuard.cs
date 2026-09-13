using System;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 剪贴板保护：PowerPoint COM 移动形状必须 Copy→Paste，Copy 会占用并覆盖系统剪贴板，
    /// 覆盖用户正在使用的剪贴板内容体验很差。本类在操作前抓取剪贴板快照，
    /// 操作完成后恢复原内容。
    ///
    /// 说明：
    /// - 快照为延迟渲染 IDataObject（GetDataObject 不立即取全部数据，恢复时 SetDataObject(copy=true)
    ///   由原数据源回放），对文本/图片/文件及 Office 对象均适用。
    /// - 极端情况（来源程序已退出、剪贴板被系统清空）恢复可能失败——失败静默，不影响插入本身。
    /// </summary>
    public static class ClipboardGuard
    {
        /// <summary>
        /// 抓取当前剪贴板快照。
        /// 关键：逐个格式 <c>GetData(format, autoConvert:true)</c> **立即实体化**（延迟渲染的 Office 数据
        /// 也一并取成内存数据）再放进独立的 DataObject。这样即使随后我们自己的 Copy 覆盖了剪贴板
        /// （其原始 owner 的数据被释放），之后仍能把用户的剪贴板内容原样放回。
        /// 之前"遇到 PowerPoint 形状数据就跳过保护"的做法会让用户复制的内容被我们的操作冲掉。
        /// </summary>
        public static IDataObject Snapshot()
        {
            try
            {
                var src = Clipboard.GetDataObject();
                if (src == null) return null;
                string[] formats;
                try { formats = src.GetFormats(); } catch { return null; }
                if (formats == null || formats.Length == 0) return null;   // 空剪贴板：无需保护

                var copy = new DataObject();
                bool any = false;
                foreach (var f in formats)
                {
                    if (string.IsNullOrEmpty(f)) continue;
                    try
                    {
                        object data = src.GetData(f, true);   // true = 允许自动转换/强制渲染
                        if (data == null) continue;
                        copy.SetData(f, data);
                        any = true;
                    }
                    catch { /* 个别格式无法实体化：跳过该格式 */ }
                }
                return any ? (IDataObject)copy : null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>恢复剪贴板为快照内容（尽力而为，失败静默）。</summary>
        public static void Restore(IDataObject snapshot)
        {
            if (snapshot == null) return;
            for (int attempt = 1; attempt <= 2; attempt++)
            {
                try
                {
                    Clipboard.SetDataObject(snapshot, true);
                    return;
                }
                catch
                {
                    // 偶发占用（剪贴板被其他进程短暂锁定）：稍后重试一次，仍失败则静默放弃
                    try { System.Threading.Thread.Sleep(60); } catch { }
                }
            }
        }

        /// <summary>用剪贴板保护执行一次会覆盖剪贴板的操作。</summary>
        public static T WithProtect<T>(Func<T> operation)
        {
            var snapshot = Snapshot();
            try
            {
                return operation();
            }
            finally
            {
                Restore(snapshot);
            }
        }
    }
}
