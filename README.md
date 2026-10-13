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
- **搜索**：实时过滤（✕/Esc 清空），匹配名称 / 分类
- **类型筛选**：全部 / 形状 / 样式 一键切换
- **排序**：最新/最旧/名称/最近使用/最常用/自定义（拖拽卡片排序）
- **多选**：Ctrl 单击多选 → 批量删除/移动分类/移动分组/批量重命名/导出素材包
- **回收站**：删除先入回收站（库管理→回收站 可恢复/彻底清除/清空）
- **最近使用横条**：「全部」视图且无筛选时顶部展示最近使用 Top6（可点插）
- **插入后调整条**：插入成功浮现旋转/透明度/八向微调 + 撤销/重做（10 秒自动隐藏）
- **去重清理**：收藏相似形状时提醒；库管理里可扫描完全重复项并批量清理
- **对话框统一**：全部 confirm/prompt/alert 均 shadcn 组件（无系统灰框）

### 素材库与分享
- **内置预设素材**：首次运行自动安装 15 个常用形状组合（卡片/箭头/流程框/图表元素等）
- **素材包（v1.1）**：选中收藏导出为 `.ppack.zip`（含 manifest.json + 条目目录；带 作者署名/CC BY 4.0 许可/来源 元数据，格式规范见 `docs/material-pack-spec.md`）；「库管理 → 导入素材包…」直接弹文件选择器（可多选），id 冲突自动重写，导入后显示数量/作者/许可/来源——点对点分享、将来社区合集流通皆可
- **导出 PPTX**：把收藏库导出为 .pptx 分享
- **备份/恢复**：zip 打包备份、一键恢复（可跨电脑迁移）
- **开源状态**：代码暂不开源（开发成熟后再开源）；素材包格式与分享链路已就绪

### 样式进阶
- **主题色适配**：应用样式时可选把主色映射到当前文档主题色
- **批量应用到同类**：把样式应用到本页 / 整份演示文稿所有同类型形状

### 界面
- **布局**：右侧任务窗格只承载素材库（搜索/筛选/卡片/库管理）；所有工具与设置经 Ribbon 独立弹窗使用，不再占窗格
- **图标**：Ribbon 按钮由 RibbonIconFactory 运行时自绘字形（本地 Material Symbols TTF，见 `RibbonFonts/`）；Web 界面用 **lucide-react** 矢量图标组件（不依赖图标字体，无字体缺失导致的方框问题）
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
- **一键更换字体**：无需指定原字体，直接指定新的**中文字体 / 英文字体**并覆盖原有字体设置（含主题字体、表格与组合形状内文字；整稿或当前页）。字体下拉列出**系统全部字体**（含用户级安装字体），可搜索、按"文稿/系统"分组
- **查找/替换**：跨页保留格式替换
- **文字清理**：去多余空白、中文间西文标点转全角（按首字符格式回刷）
- **整稿导出 PNG/JPG**（可调分辨率）
- **清理瘦身**：删空白页 / 隐藏对象 / 页外对象、清空备注、移除形状动作链接（先扫描统计再执行，执行前有确认）
- **统一尺寸**：按宽/高/等比批量统一选中、当前页或整稿形状
- **图片网格插入**：多选图片 contain 等比排成整齐网格，按需翻页插入

### 学术/学习工具（Ribbon「学术/学习」组 → Web 弹窗，面向学生科研汇报）
- **页脚页码**：一键为全部页加页脚（姓名/课题组/日期）与页码
- **导出 PDF / 大纲**：整稿导出 PDF；大纲文本可存文件或直接预览/复制（留空目录时）
- **图表美化**：选中图表套用 4 套预设配色（学术蓝/清新绿/暖橙/莫兰迪）+ 统一字体
- **成套模板**：插入封面/目录/内容×N/结束成套结构页（毕业答辩/开题报告/课程汇报/组会汇报）
- **学术模板素材**：首次启动自动装入「学术模板」分类（答辩封面/开题封面/研究要点页/结论页），素材库单击插入后改字即用

## 开发与运行

- VS2022 + Office/SharePoint 工作负载，按 F5 调试
- 存储：`%APPDATA%\PPTShapeLibrary\`（items 每收藏一目录；settings.json 设置；packs 素材包；logs 日志）
- 界面宿主：WebView2（Evergreen Runtime）。缺失/初始化失败自动降级到提示页（`uiBackend:"wpf"` 旧开关已废弃）
- 快捷键：无全局快捷键（工具栏操作）
- **整页插入不占用剪贴板**：走 `Slides.InsertFromFile`（前提是库文件设计唯一改名，见 `docs/PowerPoint-COM经验与坑.md` §13）；
  形状插入仍需 Copy→Paste，剪贴板被"活着但不响应"的程序独占时会**快速失败并给出可操作提示**（不再白等十几秒）
- **长任务不卡死（分片执行）**：整稿导出图片 / 一键换字体 / 收藏整稿 / 重建缩略图 / 统一尺寸 / 清理瘦身
  改为**逐片执行**（片间回到消息循环），长任务期间 PowerPoint 仍可响应（实测 ping 即时答复）、
  取消立刻生效、期间再点别的会得到"正在执行…"的明确提示。仍同步的（只有进度+取消）：查找替换 / 文字清理 /
  整页应用样式 / 图片网格，原因与设计守则见 `docs/PowerPoint-COM经验与坑.md` §9.7
- **整稿工具只作用于用户文稿**：PowerPoint 文档恢复可能把库条目变成活动文稿，此时若不管，
  "导出图片/换字体/清理瘦身"会作用到**用户收藏的素材文件**上。已加两层防护（窗口激活回弹 +
  `ActiveUserPresentation()` 绝不返回库条目），见同文档 §9.6
- **前端改动记得让 bin 同步**：插件实际加载的是 `bin\<Config>\WebUI\`（csproj 从 `WebUI\` 拷过去）。
  `npm run deploy` 现在会**同时镜像更新**源码树 `WebUI\` 与已存在的 `bin\Debug\WebUI`、`bin\Release\WebUI`，
  否则会出现"改了前端、PowerPoint 里还是旧界面"
- **「插件在 PPT 里消失」已根治**：根因是 `msbuild /t:Rebuild /p:Configuration=Release` 会跑 Office targets 的
  `VSTOClean → UnregisterOfficeAddin`，**把加载项注册删掉**（csproj 此前只挡住了"注册"没挡住"注销"）。
  已在 csproj 里对称重定义 `UnregisterOfficeAddin`（Release 置空）。详见 `docs/PowerPoint-COM经验与坑.md` §9。
  兜底仍在：冒烟脚本自动修复、`dev-register.ps1 -Machine`（HKLM）、安装器「修复加载项注册」按钮

### 自动化验证（改完代码按顺序跑）

```powershell
# 1) C#：0 错误 0 警告
msbuild PPTShapeLibrary.sln /t:Rebuild /p:Configuration=Debug /restore /nologo /v:m

# 2) 前端：typecheck → 桥调用失败处理门禁 → build → 产物自检 → 部署（任一环失败即中止）
cd PPTShapeLibrary\webui-src ; npm run build:all

# 3) 端到端冒烟：真机 PowerPoint + WebView2 + CDP 驱动真实前端
#    覆盖：数据加载 / 坏输入健壮性 / 设置夹紧 / 去重不误判 / 整页插入三位置+不碰剪贴板+性能门禁 /
#          形状插入 / 前端无 JS 异常；**跑完自动还原素材库 meta.json**（不会污染用户数据）
pwsh -File scripts\e2e-smoke.ps1

# 4) 整页插入保真度：守卫"整页插入可以不用剪贴板"的前提（设计唯一改名 Lib<id>-*）
python scripts\verify-insert-fidelity.py

# 5) 一键换字体的 COM 语义（中/英分别落到 NameFarEast / NameAscii）
python scripts\verify-font-switch.py
```

### Web UI 前端（shadcn/ui · React + Tailwind v4）

界面已从免构建静态迁移为 **Vite + React + TypeScript + Tailwind CSS v4 + shadcn/ui** 工程：

```bash
cd PPTShapeLibrary/webui-src
npm install          # 首次：安装依赖
npm run dev          # 开发：热更新预览（无宿主时可浏览器看 UI，桥调用会失败仅 UI）
npm run build        # 生产：产出到 webui-src/dist
npm run typecheck    # TypeScript 类型检查（build:all 已内置）
npm run check:promises  # 桥调用失败处理门禁：禁止"裸调用"（build:all 已内置）
npm run smoke        # 产物完整性自检（index.html 引用 + JS/CSS 内资源；失败即非零退出）
npm run deploy       # 镜像拷贝 dist → ../WebUI 与 bin\<Config>\WebUI（插件实际加载目录）
npm run build:all    # 一键：typecheck → check:promises → build → smoke → deploy（任一环失败即中止）
```

- 源码在 `webui-src/src`；迁移前的纯静态前端（`webui-legacy/`）**已删除**——它停留在旧协议（`replaceFont`/标签等类型早已从 C# 桥移除），当"回退路径"只会误导；需要时从 git 历史取回：`git checkout HEAD~1 -- PPTShapeLibrary/webui-legacy`
- C#↔JS 桥协议见 `src/lib/bridge.ts`（默认 30s 超时；模态对话框类请求不超时）
- 新增 shadcn 组件：`npx shadcn add <组件>`；主题 tokens 见 `src/index.css`（`.dark` 类切深色）
- 原生 confirm/prompt/alert 已全局替换为 shadcn（`src/lib/dialogs.tsx` 的 `dialogs.*` + `<DialogsHost/>`）
- 注意：`webui-src/node_modules` 与 `dist` 不入 git；`WebUI/` 下提交的是构建产物——**改前端后必须在同一次提交里跑 `npm run build:all`**，否则会出现 index.html 指向不存在 bundle 的空白面板

## 文件结构

```
PPTShapeLibrary\            # VSTO 项目（PowerPoint 2016+ / M365 x64）
  宿主与面板
  ThisAddIn.cs            # 任务窗格创建 + 事件订阅/退订 + WebView2 决策树/降级
  Ribbon1.cs              # 「素材库」选项卡（收藏/素材库/工具/设置，自绘图标）
  ILibraryPane.cs / WebPaneHost.cs / ToolPopupHostForm.cs / FallbackPane.cs
  BridgeController.cs     # C#↔JS 消息路由器（协议翻译，不含业务逻辑）
  业务层
  ShapeExporter.cs        # 收藏/插入/应用样式/内置素材/导出PPTX/微调/图片导入/预览生成
  ShapeStyleExtractor.cs / StylePreset.cs / ClipboardGuard.cs
  ShapeLibraryService.cs  # 存储/分类/去重/备份/缓存/回收站/孤儿清理/id白名单
  PackService.cs          # 素材包（.ppack.zip v1.1）导出/导入
  TextToolService.cs / FormatToolService.cs / AcademicToolService.cs
  Settings.cs / Log.cs / InputDialog.cs（收藏命名）
  ToolsHubForm.cs          # Ribbon「高级批处理」Hub（打开 Web 工具弹窗）
  RibbonIconFactory.cs     # Ribbon 图标：本地 TTF 字库按族名匹配 + 码点表
  WebUI\                  # 前端构建产物（React+shadcn，随程序集复制）
  webui-src\              # React 前端工程（Vite；node_modules/dist 不入库）
docs\                     # 设计/清单/素材包格式规范（material-pack-spec.md）
```

UI 演进：A 阶段 WPF 面板 → B 阶段 WebView2 Web UI（业务层 C# 原样复用；
旧 WPF 面板已移除，降级路径为 FallbackPane 提示页）。
