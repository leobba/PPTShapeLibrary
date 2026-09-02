namespace PPTShapeLibrary
{
    /// <summary>任务窗格面板统一接口：ThisAddIn 对 WPF/Web 面板零感知切换。</summary>
    public interface ILibraryPane
    {
        void RefreshLibrary();
        void ApplySettings(Settings s);
    }
}
