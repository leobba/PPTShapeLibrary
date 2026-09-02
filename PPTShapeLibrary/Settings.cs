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

        public Settings()
        {
            Theme = "light";
            InsertPosition = "center";
            PaneWidth = 300;
            AutoName = true;
            PageSize = 200;
            CardSize = "compact";
            ClickToInsert = true;
            ThemeAdaptStyle = false;
            UiBackend = "webview2";
        }
    }
}
