using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 「形状库」自定义功能区选项卡（VSTO 原生 Ribbon 设计器机制）。
    /// </summary>
    [ComVisible(true)]
    public class Ribbon1 : RibbonBase
    {
        internal Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private RibbonTab tabLibrary;
        private RibbonGroup grpCollect;
        private RibbonGroup grpLibrary;
        private RibbonGroup grpTools;
        private RibbonButton btnCollectShapes;
        private RibbonButton btnCollectPage;
        private RibbonButton btnCollectStyle;
        private RibbonButton btnOpenPane;
        private RibbonButton btnBackup;
        private RibbonButton btnRestore;
        private RibbonButton btnExportPptx;
        private RibbonButton btnTools;
        private RibbonButton btnSettings;

        private void InitializeComponent()
        {
            var f = Globals.Factory.GetRibbonFactory();

            this.tabLibrary = f.CreateRibbonTab();
            this.grpCollect = f.CreateRibbonGroup();
            this.grpLibrary = f.CreateRibbonGroup();
            this.grpTools = f.CreateRibbonGroup();
            this.btnCollectShapes = f.CreateRibbonButton();
            this.btnCollectPage = f.CreateRibbonButton();
            this.btnCollectStyle = f.CreateRibbonButton();
            this.btnOpenPane = f.CreateRibbonButton();
            this.btnBackup = f.CreateRibbonButton();
            this.btnRestore = f.CreateRibbonButton();
            this.btnExportPptx = f.CreateRibbonButton();
            this.btnTools = f.CreateRibbonButton();
            this.btnSettings = f.CreateRibbonButton();

            // ---- 收藏组 ----
            this.btnCollectShapes.Label = "收藏形状";
            this.btnCollectShapes.OfficeImageId = "Copy";
            this.btnCollectShapes.ShowImage = true;
            this.btnCollectShapes.ScreenTip = "收藏选中形状";
            this.btnCollectShapes.SuperTip = "把当前选中的形状（含组合）整组存入本地收藏库，稍后可一键插入当前页。";
            this.btnCollectShapes.Click += OnCollectShapes;

            this.btnCollectPage.Label = "收藏整页";
            this.btnCollectPage.OfficeImageId = "SelectAll";
            this.btnCollectPage.ShowImage = true;
            this.btnCollectPage.ScreenTip = "收藏整页形状";
            this.btnCollectPage.SuperTip = "把当前页面的全部形状打包为一个收藏。";
            this.btnCollectPage.Click += OnCollectPage;

            this.btnCollectStyle.Label = "收藏样式";
            this.btnCollectStyle.OfficeImageId = "FormatPainter";
            this.btnCollectStyle.ShowImage = true;
            this.btnCollectStyle.ScreenTip = "收藏选中形状的样式";
            this.btnCollectStyle.SuperTip = "保存填充、线条、阴影、三维、文字等样式属性，之后可应用到任意形状。";
            this.btnCollectStyle.Click += OnCollectStyle;

            this.grpCollect.Label = "收藏";
            this.grpCollect.Items.Add(this.btnCollectShapes);
            this.grpCollect.Items.Add(this.btnCollectPage);
            this.grpCollect.Items.Add(this.btnCollectStyle);

            // ---- 素材库组 ----
            this.btnOpenPane.Label = "打开收藏库";
            this.btnOpenPane.OfficeImageId = "Open";
            this.btnOpenPane.ShowImage = true;
            this.btnOpenPane.ScreenTip = "显示/隐藏收藏面板";
            this.btnOpenPane.SuperTip = "在右侧打开收藏面板：单击卡片插入/应用，右键管理。";
            this.btnOpenPane.Click += OnOpenPane;

            this.btnBackup.Label = "导出备份";
            this.btnBackup.OfficeImageId = "SaveAs";
            this.btnBackup.ShowImage = true;
            this.btnBackup.ScreenTip = "导出收藏备份";
            this.btnBackup.SuperTip = "把全部收藏打包为 zip 备份文件（可跨电脑恢复）。";
            this.btnBackup.Click += OnBackup;

            this.btnRestore.Label = "导入备份";
            this.btnRestore.OfficeImageId = "FileOpen";
            this.btnRestore.ShowImage = true;
            this.btnRestore.ScreenTip = "导入收藏备份";
            this.btnRestore.SuperTip = "从 zip 备份文件恢复收藏库。";
            this.btnRestore.Click += OnRestore;

            this.btnExportPptx.Label = "导出PPTX";
            this.btnExportPptx.OfficeImageId = "FileSaveAs";
            this.btnExportPptx.ShowImage = true;
            this.btnExportPptx.ScreenTip = "导出为 PPTX";
            this.btnExportPptx.SuperTip = "把收藏库导出为 PowerPoint 文件（每页一个收藏），方便分享。";
            this.btnExportPptx.Click += OnExportPptx;

            this.grpLibrary.Label = "素材库";
            this.grpLibrary.Items.Add(this.btnOpenPane);
            this.grpLibrary.Items.Add(this.btnBackup);
            this.grpLibrary.Items.Add(this.btnRestore);
            this.grpLibrary.Items.Add(this.btnExportPptx);

            // ---- 工具组 ----
            this.btnTools.Label = "小工具";
            this.btnTools.OfficeImageId = "MacroSecurity";
            this.btnTools.ShowImage = true;
            this.btnTools.ScreenTip = "PPT 小工具";
            this.btnTools.SuperTip = "批量修改文字样式、一键更换字体等实用工具。";
            this.btnTools.Click += OnTools;

            this.btnSettings.Label = "设置";
            this.btnSettings.OfficeImageId = "ToolsOptions";
            this.btnSettings.ShowImage = true;
            this.btnSettings.ScreenTip = "插件设置";
            this.btnSettings.SuperTip = "主题（浅色/深色）、插入位置、卡片大小、面板宽度等。";
            this.btnSettings.Click += OnSettings;

            this.grpTools.Label = "工具";
            this.grpTools.Items.Add(this.btnTools);
            this.grpTools.Items.Add(this.btnSettings);

            // ---- 组装选项卡 ----
            this.tabLibrary.Groups.Add(this.grpCollect);
            this.tabLibrary.Groups.Add(this.grpLibrary);
            this.tabLibrary.Groups.Add(this.grpTools);
            this.tabLibrary.Label = "形状库";

            this.Tabs.Add(this.tabLibrary);
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.PerformLayout();
        }

        // ---------------- 事件 ----------------

        private void OnCollectShapes(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string id = ShapeExporter.CollectShapes(null);
                if (id != null) Globals.ThisAddIn.ShowTaskPane();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ribbon CollectShapes");
                System.Windows.Forms.MessageBox.Show("操作失败：" + ex.Message, "PPT形状库");
            }
        }

        private void OnCollectPage(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string id = ShapeExporter.CollectWholePage();
                if (id != null) Globals.ThisAddIn.ShowTaskPane();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ribbon CollectPage");
                System.Windows.Forms.MessageBox.Show("操作失败：" + ex.Message, "PPT形状库");
            }
        }

        private void OnCollectStyle(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string id = ShapeExporter.CollectStyle(null);
                if (id != null) Globals.ThisAddIn.ShowTaskPane();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ribbon CollectStyle");
                System.Windows.Forms.MessageBox.Show("操作失败：" + ex.Message, "PPT形状库");
            }
        }

        private void OnOpenPane(object sender, RibbonControlEventArgs e)
        {
            try { Globals.ThisAddIn.ToggleTaskPane(); }
            catch (Exception ex) { Log.Error(ex, "Ribbon OpenPane"); }
        }

        private void OnBackup(object sender, RibbonControlEventArgs e)
        {
            try { Globals.ThisAddIn.DoBackup(); }
            catch (Exception ex) { Log.Error(ex, "Ribbon Backup"); }
        }

        private void OnRestore(object sender, RibbonControlEventArgs e)
        {
            try { Globals.ThisAddIn.DoRestore(); }
            catch (Exception ex) { Log.Error(ex, "Ribbon Restore"); }
        }

        private void OnExportPptx(object sender, RibbonControlEventArgs e)
        {
            try { Globals.ThisAddIn.DoExportPptx(); }
            catch (Exception ex) { Log.Error(ex, "Ribbon ExportPptx"); }
        }

        private void OnTools(object sender, RibbonControlEventArgs e)
        {
            try
            {
                using (var frm = new PPTToolsForm())
                {
                    frm.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ribbon Tools");
            }
        }

        private void OnSettings(object sender, RibbonControlEventArgs e)
        {
            try { Globals.ThisAddIn.DoSettings(); }
            catch (Exception ex) { Log.Error(ex, "Ribbon Settings"); }
        }
    }
}
