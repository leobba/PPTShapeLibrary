# PPTShapeLibrary —— PowerPoint 素材库与 PPT 工具插件（VSTO）

对标 Canva/创客贴素材库的 PowerPoint 素材（形状/样式）收藏插件。

## 功能总览

### 收藏与插入
- **收藏选中形状**：整组收藏，插入时自动居中、超大自动缩放、插入后自动选中
- **收藏整页**：一键把当前页所有形状打包收藏
- **收藏样式**：完整还原渐变/圆角/旋转/阴影/三维/文字/发光等属性
- **收藏剪贴板**：不要求预选形状——复制后在 Ribbon 点「收藏剪贴板」即入库
- **收藏整稿**：一键把整份演示文稿每页非占位形状批量收藏（自动命名"第 N 页"）
- **点击插入/应用**：单击卡片插入（可在设置中改为"单击预览、双击插入"）
- **快速调整条**：插入后立即微调（旋转/透明度/上下左右微调），支持撤销上一步

### 素材管理
- **双级分类**：左侧一级分组栏（全部 / 我的 / 模板 / 素材 / 文字 / 背景 / 图例 + 可自定义组），顶部二级分类胶囊随所选组切换；组可自建/改名/删除（删除组内条目归「我的」），二级可新建
- **内容库（内置）**：程序生成并归类——背景（纯色/渐变）、文字（标题/副标题/数字/引用）、图例（流程/环形/对比）、素材（形状/线条/边框/图片）；「我的」收藏由 Ribbon 收藏动作与右键/批量"移动到分组"维护
- **图片素材**：库管理「导入图片…」把本地图片入库到 素材 › 图片；另有程序生成示意图（渐变/斜纹/圆点/网格）自动补装
- **跨文档自动刷新**：打开/新建/关闭演示文稿时素材库自动刷新（防抖 600ms）
- **分类**：分类标签条筛选、移动到分类、批量移动分类
- **标签**：自由标签、搜索同时匹配名称/分类/标签、标签管理器（重命名/删除）、标签下拉筛选
- **类型筛选**：全部 / 形状 / 样式 一键切换
- **排序**：最新/最旧/名称/最近使用/最常用/自定义（拖拽卡片排序）
- **多选**：Ctrl 单击多选 → 批量删除/移动分类/移动分组/批量重命名/导出素材包
- **回收站**：删除先入回收站（库管理→回收站 可恢复/彻底清除/清空）
- **搜索**：实时过滤
- **去重清理**：收藏相似形状时提醒；库管理里可扫描完全重复项并批量清理

### 素材库与分享
- **内置预设素材**：首次运行自动安装 15 个常用形状组合（卡片/箭头/流程框/图表元素等）
- **素材包**：选中收藏导出为 .ppack.zip（存于 packs 目录），可导入他人素材包；id 冲突自动重写
- **导出 PPTX**：把收藏库导出为 .pptx 分享
- **备份/恢复**：zip 打包备份、一键恢复（可跨电脑迁移）

### 样式进阶
- **主题色适配**：应用样式时可选把主色映射到当前文档主题色
- **批量应用到同类**：把样式应用到本页 / 整份演示文稿所有同类型形状

### 界面
- **布局**：右侧任务窗格只承载素材库（搜索/筛选/卡片/库管理）；所有工具与设置经 Ribbon 独立弹窗使用，不再占窗格
- **图标**：Material Symbols Outlined 本地图标字体——Ribbon 按钮由 RibbonIconFactory 运行时自绘字形(TTF)，Web 界面用图标字体(woff2)；涵盖素材库/文字样式/换字体/查找替换/清理/导出图片/瘦身/统一尺寸/图片网格/设置等
- **UI 组件精修**（纯 CSS 组件化质感层）：统一聚焦环/三级阴影/圆角 tokens、按钮渐变与按压微交互、卡片 hover 抬升+缩略图缩放、玻璃标题条与菜单、模态重阴影、滚动条美化、侧栏深浅主题专项配色
- Canva 风格：圆角卡片、阴影、悬停高亮、悬停删除、空状态引导、玻璃浮动条、统一 CSS 变量深浅双主题
- **深色/浅色主题**切换、卡片密度（紧凑/标准/大图）、**面板宽度自适应**（拖宽卡片自动变大、列数联动）

### 工具入口（Ribbon 图标按钮 → Web UI 无边框弹窗）
Ribbon「素材库」选项卡 →「工具」组：文字样式 / 一键换字体 / 查找替换 / 文字清理 / 导出图片 / 清理瘦身 / 统一尺寸 / 图片网格 / 设置（均带自绘图标）。
每个按钮点击**弹出独立 WebView2 窗口**：无系统标题栏（避免双标题），标题由窗口内 Web 标题条呈现（图标+标题+关闭钮，可拖动），整窗同一 Web UI/图标/深浅主题、布局自适应可滚动。
- 8 个工具与设置窗口复用共享服务层（TextToolService / FormatToolService）与素材库同一前端
- Ribbon「高级批处理」为原生 Hub（内含四项，各自打开 Web 工具窗口）

### PPT 工具能力清单
- **批量修改文字样式**：选中/当前页/全部页统一字体字号加粗斜体颜色对齐
- **一键更换字体**：自动列出文档字体，原 → 新（整稿或当前页）
- **查找/替换**：跨页保留格式替换
- **文字清理**：去多余空白、中文间西文标点转全角（按首字符格式回刷）
- **整稿导出 PNG/JPG**（可调分辨率）
- **清理瘦身**：删空白页 / 隐藏对象 / 页外对象、清空备注、移除形状动作链接（先扫描统计再执行，执行前有确认）
- **统一尺寸**：按宽/高/等比批量统一选中、当前页或整稿形状
- **图片网格插入**：多选图片 contain 等比排成整齐网格，按需翻页插入

### 学术/学习工具（Ribbon「学术/学习」组 → Web 弹窗，面向学生科研汇报）
- **页脚页码**：一键为全部页加页脚（姓名/课题组/日期）与页码
- **导出 PDF / 大纲**：整稿导出 PDF，或把全部文字整理为 .txt 大纲（做笔记/打印）
- **图表美化**：选中图表套用 4 套预设配色（学术蓝/清新绿/暖橙/莫兰迪）+ 统一字体
- **成套模板**：插入封面/目录/内容×N/结束成套结构页
- **学术模板素材**：首次启动自动装入「学术模板」分类（答辩封面/开题封面/研究要点页/结论页），素材库单击插入后改字即用

## 开发与运行

- VS2022 + Office/SharePoint 工作负载，按 F5 调试
- 存储：`%APPDATA%\PPTShapeLibrary\`（items 每收藏一目录；settings.json 设置；packs 素材包；logs 日志）
- 界面宿主：WebView2（Evergreen Runtime）。缺失/初始化失败自动降级到提示页（`uiBackend:"wpf"` 旧开关已废弃）
- 快捷键：无全局快捷键（工具栏操作）

### Web UI 前端（shadcn/ui · React + Tailwind v4）

界面已从免构建静态迁移为 **Vite + React + TypeScript + Tailwind CSS v4 + shadcn/ui** 工程：

```bash
cd PPTShapeLibrary/webui-src
npm install          # 首次：安装依赖
npm run dev          # 开发：热更新预览（无宿主时可浏览器看 UI，桥调用会失败仅 UI）
npm run build        # 生产：产出到 webui-src/dist
# 构建后部署到 WebUI/（VSTO 加载目录）：
#   PowerShell: 拷贝 webui-src/dist/* → PPTShapeLibrary/WebUI/（或运行 deploy 脚本）
```

- 源码在 `webui-src/src`；旧版纯静态前端备份于 `webui-legacy/`（可回退）
- C#↔JS 桥协议不变（`src/lib/bridge.ts` 与旧 bridge.js 一致）；`npm run build` 后重启 PowerPoint 生效
- 新增 shadcn 组件：`npx shadcn add <组件>`；主题 tokens 见 `src/index.css`（`.dark` 类切深色）
- 注意：`webui-src/node_modules` 与 `dist` 不入 git；`WebUI/` 下提交的是构建产物

## 文件结构

```
PPTShapeLibrary\            # VSTO 项目（PowerPoint 2016+ / M365 x64）
  宿主与面板
  ThisAddIn.cs            # 任务窗格创建 + WebView2 决策树/降级
  Ribbon1.cs              # 「素材库」选项卡（收藏/素材库/工具/设置）
  ILibraryPane.cs         # 面板统一接口（宿主对面板零感知）
  WebPaneHost.cs          # WebView2 宿主：虚拟域名映射/消息桥/失败信号
  ToolPopupHostForm.cs    # WebView2 工具/设置独立弹窗宿主（?view=toolX|settings）
  FallbackPane.cs         # 极薄降级提示页（WebView2 不可用时的兜底）
  BridgeController.cs     # C#↔JS 消息路由器（协议翻译，不含业务逻辑）
  业务层
  ShapeExporter.cs        # 收藏/插入/应用样式/内置素材/导出 PPTX/微调/撤销
  ShapeStyleExtractor.cs  # 样式提取与应用
  ShapeLibraryService.cs  # 存储/分类/标签/去重/备份/缓存
  PackService.cs          # 素材包（.ppack.zip）导出/导入
  TextToolService.cs      # 文本工具共享服务层（Web 与 WinForms 共用）
  FormatToolService.cs    # 版式批处理共享服务层（导出图片/清理瘦身/统一尺寸/图片网格）
  AcademicToolService.cs  # 学术工具服务层（页脚页码/PDF与大纲/图表美化/参考文献）
  StylePreset.cs          # 样式数据结构
  Settings.cs / SettingsForm.cs / TagManagerForm.cs / InputDialog.cs
  ToolsHubForm.cs          # 高级批处理 Hub（Ribbon「高级批处理」入口）
  TextStyleToolForm.cs     # 工具窗：文字样式（Ribbon 直达）
  FontReplaceToolForm.cs   # 工具窗：一键换字体（Ribbon 直达）
  PPTTextToolsForm.cs      # 工具窗：查找替换 / 文字清理（单模式直达）
  PPTFormatToolsForm.cs    # 工具窗：导出图片/清理瘦身/统一尺寸/图片网格（startTab+singleMode 直达）
  PPTToolsForm.cs          # 旧文字样式/换字体整窗（兼容保留）
  Log.cs                   # 日志（>1MB 轮转）
  WebUI\                  # 前端 SPA（免构建，随程序集复制）
    index.html + css\app.css + js\bridge.js(通信) + js\app.js(应用)
```

UI 演进：A 阶段 WPF 面板 → B 阶段 WebView2 Web UI（业务层 C# 原样复用；
旧 WPF 面板已移除，降级路径为 FallbackPane 提示页）。
