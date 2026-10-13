using System;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 剪贴板保护：PowerPoint COM 插入形状必须 Copy→Paste，Copy 会占用并覆盖系统剪贴板，
    /// 覆盖用户正在使用的剪贴板内容体验很差。本类在操作前抓取剪贴板快照，操作完成后恢复原内容。
    ///
    /// 说明：
    /// - 快照**只抓文本/图片/文件**这三类"可安全回放"的格式，并立即拷成**托管 DataObject**：
    ///   ① 回放 Office/OLE 专有格式会在封送阶段损坏堆（FatalExecutionEngineError 0xc0000005，
    ///      无法 try/catch），见 docs/PowerPoint-COM经验与坑.md 第 5 条；
    ///   ② 更重要的：`Clipboard.GetDataObject()` 返回的是**延迟渲染的 OLE 包装**，
    ///      持有它等于在整个 Copy→Paste 期间拽着剪贴板引用，会让本进程的
    ///      `Clipboard.Clear()` / PowerPoint 的 `Copy()` 拿不到剪贴板所有权
    ///      （实测现象：连续 6 次 `Clipboard is empty or contains data which may not be pasted here`，
    ///      且 `GetFormats()` 直接抛异常）。拷成托管对象后引用立即释放，不再自锁。
    /// - 极端情况（来源程序已退出、剪贴板被系统清空）恢复可能失败——失败静默，不影响插入本身。
    /// </summary>
    public static class ClipboardGuard
    {
        /// <summary>
        /// 抓取当前剪贴板快照（仅文本/图片/文件；返回托管副本，不持有 OLE 引用）。
        /// 无可保护内容（空剪贴板 / 只有 Office 专有格式）时返回 null。
        /// </summary>
        public static IDataObject Snapshot()
        {
            var data = new DataObject();
            bool any = false;
            try
            {
                if (Clipboard.ContainsText())
                {
                    string text = Clipboard.GetText();
                    if (!string.IsNullOrEmpty(text))
                    {
                        data.SetData(DataFormats.UnicodeText, text);
                        any = true;
                    }
                }
            }
            catch { }
            try
            {
                if (Clipboard.ContainsImage())
                {
                    var img = Clipboard.GetImage();
                    if (img != null)
                    {
                        data.SetData(DataFormats.Bitmap, img);
                        any = true;
                    }
                }
            }
            catch { }
            try
            {
                if (Clipboard.ContainsFileDropList())
                {
                    var files = Clipboard.GetFileDropList();
                    if (files != null && files.Count > 0)
                    {
                        var arr = new string[files.Count];
                        files.CopyTo(arr, 0);
                        data.SetData(DataFormats.FileDrop, arr);
                        any = true;
                    }
                }
            }
            catch { }
            return any ? data : null;
        }

        /// <summary>恢复剪贴板为快照内容（尽力而为，失败静默）。</summary>
        public static void Restore(IDataObject snapshot)
        {
            if (snapshot == null) return;
            try
            {
                // copy=true：让内容在插件退出后仍留在剪贴板
                Clipboard.SetDataObject(snapshot, true);
            }
            catch
            {
                // 恢复失败静默——用户仍可手动 Ctrl+V 原内容（若源仍在）
            }
        }

        /// <summary>剪贴板当前是否"可读"：读不到（抛异常）通常意味着被别的程序独占且未响应。
        /// 供 Copy→Paste 重试逻辑尽早失败并给出准确提示，而不是白等十几秒。</summary>
        public static bool IsReadable()
        {
            try
            {
                Clipboard.GetDataObject();
                return true;
            }
            catch
            {
                return false;
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
