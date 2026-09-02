# PPTShapeLibrary —— PowerPoint 形状收藏与 PPT 小工具插件（VSTO）

对标 Canva/创客贴素材库的 PowerPoint 形状收藏插件。

## 功能总览

### 收藏与插入
- **收藏选中形状**：整组收藏，插入时自动居中、超大自动缩放、插入后自动选中
- **收藏整页**：一键把当前页所有形状打包收藏
- **收藏样式**：完整还原渐变/圆角/旋转/阴影/三维/文字/发光等属性
- **点击插入/应用**：单击卡片插入（可在设置中改为"单击预览、双击插入"）
- **快速调整条**：插入后立即微调（旋转/透明度/上下左右微调）

### 素材管理
- **分类**：分类标签条筛选、移动到分类、批量移动分类
- **标签**：自由标签、搜索同时匹配名称/分类/标签、标签管理器（重命名/删除）
- **排序**：最新/最旧/名称/最近使用/最常用/自定义（拖拽卡片排序）
- **多选**：Ctrl 单击多选 → 批量删除/移动分类
- **搜索**：实时过滤
- **去重提示**：收藏相似形状时提醒；内置完全重复检测可清理

### 素材库与分享
- **内置预设素材**：首次运行自动安装 15 个常用形状组合（卡片/箭头/流程框/图表元素等）
- **导出 PPTX**：把收藏库导出为 .pptx 分享
- **备份/恢复**：zip 打包备份、一键恢复（可跨电脑迁移）

### 样式进阶
- **主题色适配**：应用样式时可选把主色映射到当前文档主题色
- **批量应用到同类**：把样式应用到本页所有同类型形状

### PPT 小工具
- **批量修改文字样式**：选中形状/当前页/全部页面范围内统一字体/字号/加粗/斜体/颜色/对齐
- **一键更换字体**：原字体 → 新字体（自动列出文档中出现的字体）

### 界面
- Canva 风格：圆角卡片、阴影、悬停高亮、悬停删除、空状态引导、渐变标题栏
- **深色/浅色主题**切换、卡片密度（紧凑/标准/大图）、面板宽度可调

## 开发与运行

- VS2022 + Office/SharePoint 工作负载，按 F5 调试
- 存储：`%APPDATA%\PPTShapeLibrary\`（items 每收藏一目录；settings.json 设置；packs 素材包；logs 日志）
- 界面宿主：WebView2（Evergreen Runtime）。缺失/初始化失败自动降级到提示页（`uiBackend:"wpf"` 旧开关已废弃）
- 快捷键：无全局快捷键（工具栏操作）

## 文件结构

```
PPTShapeLibrary\            # VSTO 项目（PowerPoint 2016+ / M365 x64）
  宿主与面板
  ThisAddIn.cs            # 任务窗格创建 + WebView2 决策树/降级
  Ribbon1.cs              # 「形状库」选项卡（收藏/素材库/工具/设置）
  ILibraryPane.cs         # 面板统一接口（宿主对面板零感知）
  WebPaneHost.cs          # WebView2 宿主：虚拟域名映射/消息桥/失败信号
  FallbackPane.cs         # 极薄降级提示页（WebView2 不可用时的兜底）
  BridgeController.cs     # C#↔JS 消息路由器（协议翻译，不含业务逻辑）
  业务层
  ShapeExporter.cs        # 收藏/插入/应用样式/内置素材/导出 PPTX/微调/撤销
  ShapeStyleExtractor.cs  # 样式提取与应用
  ShapeLibraryService.cs  # 存储/分类/标签/去重/备份/缓存
  PackService.cs          # 素材包（.ppack.zip）导出/导入
  StylePreset.cs          # 样式数据结构
  Settings.cs / SettingsForm.cs / TagManagerForm.cs / InputDialog.cs
  PPTToolsForm.cs         # PPT 小工具（批量文字/换字体）
  Log.cs                  # 日志（>1MB 轮转）
  WebUI\                  # 前端 SPA（免构建，随程序集复制）
    index.html + css\app.css + js\bridge.js(通信) + js\app.js(应用)
```

UI 演进：A 阶段 WPF 面板 → B 阶段 WebView2 Web UI（业务层 C# 原样复用；
旧 WPF 面板已移除，降级路径为 FallbackPane 提示页）。
