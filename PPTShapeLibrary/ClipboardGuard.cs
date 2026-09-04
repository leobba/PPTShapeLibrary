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
        /// 抓取当前剪贴板快照。仅当剪贴板内容"可被可靠回放"时才保护：
        /// 纯 PowerPoint 形状数据属延迟渲染且源即本进程，恢复不可靠且用户多半在 PPT 内操作，跳过。
        /// </summary>
        public static IDataObject Snapshot()
        {
            try
            {
                bool hasText = Clipboard.ContainsText();
                bool hasImage = Clipboard.ContainsImage();
                bool hasFiles = Clipboard.ContainsFileDropList();
                bool hasPpt = Clipboard.ContainsData("PowerPoint Shapedata");
                if (!hasText && !hasImage && !hasFiles)
                {
                    return null;   // 只有 Office 专有/空剪贴板——无需（无法可靠）保护
                }
                // 文本/图片/文件为主：可回放，抓快照
                return Clipboard.GetDataObject();
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
            try
            {
                // 延迟到操作后一小段时间执行，避免与 PowerPoint 内部剪贴板处理竞争
                Clipboard.SetDataObject(snapshot, true);
            }
            catch
            {
                // 恢复失败静默——用户仍可手动 Ctrl+V 原内容（若源仍在）
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
