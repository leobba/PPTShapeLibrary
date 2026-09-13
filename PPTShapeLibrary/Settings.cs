using System.Runtime.Serialization;

namespace PPTShapeLibrary
{
    /// <summary>插件设置（JSON 持久化到 %APPDATA%\PPTShapeLibrary\settings.json）。</summary>
    [DataContract]
    public class Settings
    {
        [DataMember(Name = "theme")]
        public string Theme { get; set; }              // "light" / "dark"

        [DataMember(Name = "insertPosition")]
        public string InsertPosition { get; set; }     // "center" / "original"

        [DataMember(Name = "paneWidth")]
        public int PaneWidth { get; set; }             // 面板宽度

        [DataMember(Name = "autoName")]
        public bool AutoName { get; set; }             // 自动按形状类型命名

        [DataMember(Name = "pageSize")]
        public int PageSize { get; set; }              // 分页每页数量

        [DataMember(Name = "cardSize")]
        public string CardSize { get; set; }           // "compact"/"standard"/"large"

        [DataMember(Name = "clickToInsert")]
        public bool ClickToInsert { get; set; }        // true=单击插入；false=单击预览、双击插入

        [DataMember(Name = "themeAdaptStyle")]
        public bool ThemeAdaptStyle { get; set; }      // 应用样式时映射当前主题色

        [DataMember(Name = "uiBackend", EmitDefaultValue = false)]
        public string UiBackend { get; set; }          // "webview2"(默认) / "wpf"(手动降级)

        [DataMember(Name = "authorName", EmitDefaultValue = false)]
        public string AuthorName { get; set; }          // 素材作者署名（导出素材包时写入 manifest）

        [DataMember(Name = "collectPlaceholderText")]
        public bool CollectPlaceholderText { get; set; }  // 收藏形状时把文字替换为占位提示（模板化收藏）

        [DataMember(Name = "paneDock")]
        public string PaneDock { get; set; }              // 任务窗格停靠："right"/"left"/"top"/"bottom"

        [DataMember(Name = "paneHeight")]
        public int PaneHeight { get; set; }               // 停靠 上/下 时窗格高度

        [DataMember(Name = "gridColumns")]
        public string GridColumns { get; set; }           // 素材网格列数："2"/"3"/"4"/"auto"(自适应)

        [DataMember(Name = "showTypeRow")]
        public bool ShowTypeRow { get; set; }             // 是否显示"类型：全部/形状/样式"行

        public Settings()
        {
            Theme = "light";
            InsertPosition = "center";
            PaneWidth = 400;
            AutoName = true;
            PageSize = 200;
            CardSize = "compact";
            ClickToInsert = true;
            ThemeAdaptStyle = false;
            UiBackend = "webview2";
            AuthorName = "";
            CollectPlaceholderText = false;
            PaneDock = "left";
            PaneHeight = 180;
            GridColumns = "2";
            ShowTypeRow = true;
        }
    }
}
