using System.Drawing;

namespace PPTShapeLibrary
{
    /// <summary>
    /// Ribbon 图标工厂（精简版）。
    /// 第十八批起停用自绘 TTF 图标字体（原 MaterialIcons/MaterialSymbols TTF 约 10.6MB，
    /// 且 Ribbon 改为 Office 原生文字按钮，无需运行时字形绘制）。
    /// 本类仅保留图标名常量供代码语义引用；GetActionGlyph/GetColorGlyph 保持签名兼容但返回 null。
    /// </summary>
    public static class RibbonIconFactory
    {
        /// <summary>兼容入口：返回 null（无图标，Ribbon 按钮为纯文字；ShowImage=false）。</summary>
        public static Image GetColorGlyph(string iconName, int size, Color color)
        {
            return null;
        }

        /// <summary>兼容入口：返回 null。</summary>
        public static Image GetActionGlyph(string iconName, string action, int size = 32)
        {
            return null;
        }

        // ---------------- 图标名常量（语义引用；供未来 OfficeImageId 方案对照） ----------------

        public const string Tool1Icon = "text_fields";
        public const string Tool2Icon = "font_download";
        public const string Tool3Icon = "manage_search";
        public const string Tool4Icon = "cleaning_services";
        public const string Tool5Icon = "image";
        public const string Tool6Icon = "delete_sweep";
        public const string Tool7Icon = "aspect_ratio";
        public const string Tool8Icon = "grid_view";
        public const string Tool9Icon = "text_snippet";      // 页脚页码
        public const string Tool10Icon = "picture_as_pdf";   // 导出PDF
        public const string Tool11Icon = "bar_chart";        // 图表美化
        public const string Tool12Icon = "view_day";         // 成套模板
        public const string SettingsIcon = "settings";

        public const string CollectShapesIcon = "content_copy";
        public const string CollectPageIcon = "select_all";
        public const string CollectStyleIcon = "format_paint";
        public const string CollectClipboardIcon = "content_paste";
        public const string CollectAllSlidesIcon = "collections";
        public const string OpenPaneIcon = "photo_library";
        public const string BackupIcon = "save";
        public const string RestoreIcon = "unarchive";
        public const string ExportPptxIcon = "slideshow";
        public const string AdvancedIcon = "construction";
    }
}
