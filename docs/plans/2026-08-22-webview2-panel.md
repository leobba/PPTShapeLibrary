> ## ⚠️ 历史存档（2026-10 注）
> 本文是 2026-08 迁移到 WebView2 时的**设计计划**，部分结论已被后续实现推翻，阅读时请注意：
> - 文中把 **WPF 面板当作长期回退方案**、把 `uiBackend:"wpf"` 当作"随时可逆的总开关"——**已不成立**：
>   旧 WPF 面板已移除，现在的降级路径是 `FallbackPane` 提示页（见 README「界面宿主」）。
> - 文中"C3：不引入 React/Vue/Node 构建链"、图标用 Segoe Fluent Icons/MDL2 —— **已不成立**：
>   前端已是 Vite + React + TypeScript + Tailwind v4 + shadcn/ui，Web 图标用 lucide-react。
> - 文中前端 5s 超时、17 种上行类型等数字 **已过时**：现为默认 30s + 模态对话框类不超时。
> - 仍然有效且值得参考：C2 的 WebView2 版本要求、C4 的虚拟域映射（`app.pptshapes.local` / `data.pptshapes.local`）、
>   以及"桥消息同步在 UI 线程执行"这一约束（现已有分片执行方案，见 `PowerPoint-COM经验与坑.md` §9.7）。
# PPT 形状收藏插件 B 阶段：任务窗格迁移 WebView2（Web UI）开发方案

- 日期：2026-08-22
- 状态：开发中（生命周期：草案→已确认→开发中→已验收）
- 确认记录：2026-08-22 用户确认通过；开放问题 1 决策=先不捆绑 Fixed Version Runtime；开放问题 2 决策=P2 主题快捷按钮先实现、验收时定
- 仓库：E:\code\PPTShapeLibrary\
- 规模：完整

---

## 1. 背景与目标

### 1.1 目标

把「形状收藏库」任务窗格的 UI 层从 WPF（ElementHost 内嵌）整体替换为 WebView2 加载的纯 HTML/CSS/JS 单页应用，达到创客贴/Canva 级别的现代化视觉与交互质感；业务层 C# 零改动复用，并保留旧 WPF 面板作为一键降级路径。

### 1.2 背景

- 用户反馈现有面板"UI 很廉价"，期望 Canva 式形状库观感；A 阶段（Ribbon 化）已完成验证，本方案为 B 阶段。
- WPF 面板（`ShapeLibraryPaneWpf.xaml.cs`，726 行）视觉上限受限于手写 XAML 样式：圆角/阴影/动效/玻璃感实现成本高、迭代慢。
- WebView2 是 Microsoft 365 已内置的 Evergreen 运行时，用 Web 技术做 UI 可获得：CSS 变量主题系统、backdrop-filter 毛玻璃层叠、流畅 CSS 动效、成熟的组件化布局（Grid/Flex），且免 Node 构建链。

### 1.3 不做清单（明确排除）

| 不做 | 原因 |
|------|------|
| 不改业务层 C# 逻辑 | `ShapeExporter.cs` / `ShapeLibraryService.cs` / `ShapeStyleExtractor.cs` / `Log.cs` 原样复用，仅新增调用方 |
| 不引入 React/Vue/Node 构建链 | 免构建，原生 HTML/CSS/JS 直写，保证离线可用与低维护成本 |
| 不做跨进程拖拽插入（从面板拖形状到画布） | 涉及 OLE DragDrop 与 WebView2 跨进程协同，风险高收益低，另立后续任务 |
| 不重写收藏数据存储格式 | `%APPDATA%\PPTShapeLibrary\items\{id}\` 结构、settings.json 既有字段全部保持 |
| 不动 A 阶段成果 | `Ribbon1.cs` 及其调用的 `ThisAddIn.Do*()` 方法签名不变 |

---

## 2. 范围与全局约束

### 2.1 做

- 新增 WinForms 宿主控件 `WebViewPaneHost`（内嵌 `Microsoft.Web.WebView2.WinForms.WebView2`），替换 `ThisAddIn.cs` 中对 `WpfPaneHost` 的实例化。
- 新增前端 SPA：`WebUI\index.html + css\app.css + js\bridge.js + js\app.js`（随程序集输出复制）。
- 新增 C#↔JS 消息路由器 `BridgeController`（基于 `System.Web.Extensions` 已引用的 `JavaScriptSerializer`，零新增 JSON 依赖）。
- 全量对齐现有 13 项面板功能（见第 8 节功能清单）。
- WebView2 Runtime 缺失检测、初始化失败自动降级回 WPF 面板、手动降级开关。

### 2.2 不做

见 1.3。

### 2.3 硬性约束

| # | 约束 |
|---|------|
| C1 | .NET Framework 4.8 / VSTO 运行时 10.0.60915 / PowerPoint 2016+（Microsoft 365 x64），不得升级目标框架 |
| C2 | WebView2 SDK 采用 NuGet 包 `Microsoft.Web.WebView2` 固定版本 **1.0.2903.40**（如实施时需升级，只允许升到更高稳定版并全量回归） |
| C3 | 前端完全离线：不引任何 CDN/外链资源；图标字体用系统自带 `'Segoe Fluent Icons'`（Win11）/ `'Segoe MDL2 Assets'`（Win10 回退）；正文用 `'Segoe UI', 'Microsoft YaHei UI'` |
| C4 | 页面经 `SetVirtualHostNameToFolderMapping` 以虚拟域名 https 加载，禁止 file:// 协议 |
| C5 | 深/浅主题用 CSS 变量切换，切换过程不刷新页面、不重新导航 |
| C6 | 所有 JS→C# 调用必须异步消息往返（postMessage），C# 侧处理器一律 try/catch 并写 `Log.Error`，不允许未捕获异常冒泡到 Office 进程 |
| C7 | WebView2 用户数据文件夹固定为 `%APPDATA%\PPTShapeLibrary\webview2`，不使用默认路径（Office 进程默认路径可能不可写） |
| C8 | 构建命令必须在 Visual Studio 开发者命令行可重复执行：`msbuild PPTShapeLibrary.sln /t:Rebuild /p:Configuration=Debug /restore` |

---

## 3. 总体架构与技术选型

### 3.1 文字版架构图

```
┌─────────────────────────────────────────────────────────────────────┐
│ PowerPoint (Microsoft 365 x64, 主 STA 线程)                          │
│                                                                     │
│  Ribbon「形状库」(Ribbon1.cs, 不动)                                  │
│      │ ToggleTaskPane()/ShowTaskPane()/RefreshLibrary()/DoSettings()│
│      ▼                                                              │
│  ThisAddIn.cs (修改)                                                │
│      │ CustomTaskPanes.Add(<ILibraryPane>, "形状收藏库")             │
│      ▼                                                              │
│  ┌─ ILibraryPane ────────────────────────────────────────────────┐  │
│  │  RefreshLibrary() : void                                      │  │
│  │  ApplySettings(Settings s) : void                             │  │
│  ├──────────────────────────┬────────────────────────────────────┤  │
│  │  [新] WebPaneHost        │  [保留·降级路径] WpfPaneHost        │  │
│  │  (WinForms UserControl)  │  (ElementHost → ShapeLibraryPaneWpf)│  │
│  │   └─ WebView2 控件       │   └─ 冻结不再演进                   │  │
│  └──────────┬───────────────┴────────────────────────────────────┘  │
│             │ SetVirtualHostNameToFolderMapping ×2                  │
│             │   https://app.pptshapes.local  → <程序集目录>\WebUI    │
│             │   https://data.pptshapes.local → %APPDATA%\…\items\   │
│             ▼                                                       │
│  BridgeController (新增, C#)                                        │
│      │ WebMessageReceived ← window.chrome.webview.postMessage(json) │
│      │ PostWebMessageAsJson     → chrome.webview 'message' 监听      │
│      ▼  只做"协议翻译"，不含业务逻辑                                 │
│  业务层（原样复用）：ShapeExporter / ShapeLibraryService /           │
│                     ShapeStyleExtractor / Log                        │
└─────────────────────────────────────────────────────────────────────┘
         │ %APPDATA%\PPTShapeLibrary\
         ├── items\{id}\{meta.json, shape.pptx|style.json, preview.png}
         ├── settings.json
         └── webview2\            ← WebView2 用户数据文件夹 (约束 C7)
```

### 3.2 关键技术决策与备选取舍

| 决策点 | 选定方案 | 被否方案及原因 |
|--------|----------|----------------|
| C#→JS 下行通道 | `CoreWebView2.PostWebMessageAsJson(obj)`，JS 侧 `chrome.webview.addEventListener('message')` 接收 | `ExecuteScriptAsync("fn(...)")`：JSON 字符串拼接转义易错、无请求-响应配对、无法携带错误对象；仅保留用于 Debug 构建 DevTools 辅助 |
| JS→C# 上行通道 | `window.chrome.webview.postMessage(JSON字符串)` + `WebMessageReceived` | `AddHostObjectToScript`：COM 编组类型受限、同步调用存在重入死锁风险、协议不可日志化、调试困难。本项目消息均为短小 JSON，异步总线更稳 |
| 缩略图加载 | 第二个虚拟域名 `https://data.pptshapes.local/{id}/preview.png` 直接映射 items 目录，`<img>` 原生懒加载 | `getPreview` 消息回传 base64：每图一次消息往返、主线程序列化开销大、无法利用浏览器 HTTP 缓存。故协议中不设 getPreview 类型（缓存失效用 URL 查询参数 `?v=` 解决，见 4.3） |
| 过滤/排序/分页位置 | C# 侧 `BridgeController` 实现（逐条复制现 WPF 面板规则），前端只发意图 | 前端过滤：需把全量元数据推给 JS，两处排序逻辑日后漂移；C# 侧保持单一事实源，与现行为逐字节一致 |
| JSON 序列化 | `System.Web.Extensions.JavaScriptSerializer`（csproj 已引用，第 64 行） | Newtonsoft.Json/Newtonsoft 或 System.Text.Json：引入新依赖违背最小改动原则 |
| 确认/输入对话框 | 前端自绘模态框（删除确认、重命名、标签、移动分类） | 原生 `confirm()`/WinForms MessageBox：风格割裂，正是本次要消灭的"廉价感"来源之一 |
| 毛玻璃 | 浮动条/菜单用 `background: rgba(...)` + `backdrop-filter: blur(12px) saturate(140%)`，模糊对象是网页内卡片网格 | 如实说明：任务窗格内 backdrop-filter **只能模糊 WebView2 页面内部内容**，无法透视模糊其后的 PowerPoint 画布。"真毛玻璃到桌面"在宿主为窗口子控件的场景不可实现，本方案以半透明层叠+噪点纹理模拟同等观感 |

### 3.3 线程与生命周期约定

1. `WebViewPaneHost` 在 `ThisAddIn_Startup`（PowerPoint 主 STA 线程）中创建；`EnsureCoreWebView2Async` 的 await 续体经 WinForms SynchronizationContext 回到同一线程，所有 COM 互操作天然线程安全。
2. `WebMessageReceived` 事件同样在创建线程触发；每个 handler 外层统一 try/catch + `Log.Error(ex, "Bridge:"+type)`（约束 C6）。
3. 初始化失败（Runtime 缺失 / Environment 创建异常 / NavigationCompleted 失败）→ 记日志 + 提示框 + 自动切回 WPF 面板，PowerPoint 会话内不再重试 WebView2。
4. Debug 构建下按 F12 调 `CoreWebView2.OpenDevToolsWindow()`；Release 构建屏蔽。

---

## 4. 通信协议设计（v1）

### 4.1 信封格式

**上行（JS → C#）**，`window.chrome.webview.postMessage(JSON.stringify(msg))`：

```json
{ "type": "insert", "reqId": "r42", "payload": { "id": "a3f9c2e1" } }
```

**下行响应（C# → JS）**，与 reqId 配对：

```json
{ "type": "result", "reqId": "r42", "ok": true,  "data": { } }
{ "type": "result", "reqId": "r42", "ok": false, "error": "InsertShape failed: no open presentation" }
```

**下行推送（C# → JS，无 reqId）**：

```json
{ "type": "refresh", "payload": { } }
{ "type": "settingsChanged", "payload": { "theme": "dark", "...": "完整 settings 对象" } }
```

约定：reqId 由前端自增生成；前端维护 pending map 做超时兜底（5s 未收到 result 视为失败并提示）；未知 type 一律回 `{ok:false, error:"unknown type"}` 保证前向兼容。

### 4.2 消息类型总表

| type | 方向 | payload 关键字段 | C# 处理（全部为既有方法直调） |
|------|------|------------------|------------------------------|
| `ping` | ↑ | 无 | 链路自检，返回 ok |
| `init` | ↑ | 无 | LoadSettings + GetAllCategories + GetAllTags + LoadAllItems().Count |
| `list` | ↑ | filter, category, sort, offset, limit | 过滤/排序/切片规则见 4.3 |
| `insert` | ↑ | id | InsertShape(id) 成功后 TouchUsed(id)，返回 {kind} |
| `applyStyle` | ↑ | id | ApplyStyle(id) 成功后 TouchUsed(id) |
| `applyStyleSameType` | ↑ | id | ApplyStyleToSameType(id) 成功后 TouchUsed(id) |
| `adjust` | ↑ | action ∈ rotateL/rotateR/opacityUp/opacityDown/nudgeL/nudgeR/nudgeU/nudgeD | AdjustSelected(action) |
| `delete` | ↑ | ids: string[]（单个=长度1） | 循环 DeleteItem，返回 {deleted:n} |
| `rename` | ↑ | id, name | RenameItem |
| `setTags` | ↑ | id, tags（逗号分隔串） | SetTags |
| `setCategory` | ↑ | id, category | SetCategory |
| `batchMove` | ↑ | ids[], category | 循环 SetCategory，返回 {moved:n} |
| `reorder` | ↑ | orderedIds: string[] | 循环 SetSortOrder(id, index)，index 为数组下标 |
| `openFolder` | ↑ | id | Process.Start explorer ItemsDir/{id} |
| `getSettings` | ↑ | 无 | LoadSettings |
| `saveSettings` | ↑ | patch: {…合并字段} | 合并 LoadSettings 后 SaveSettings，广播 settingsChanged，返回最新 settings |
| `log` | ↑ | level(message/error), message | Log.Write / Log.Error |
| `refresh` | ↓ | 无 | 触发方：Ribbon 收藏成功后 ThisAddIn.RefreshLibrary()、备份恢复完成、内置库延迟安装完成 |
| `settingsChanged` | ↓ | settings | 触发方：设置对话框保存（DoSettings）、saveSettings 消息处理完 |

### 4.3 `list` 消息详述（核心消息）

请求：

```json
{ "type": "list", "reqId": "r7",
  "payload": { "filter": "圆形", "category": "未分类", "sort": "最近使用", "offset": 0, "limit": 200 } }
```

响应 data：

```json
{ "items": [
    { "id": "a3f9c2e1", "name": "蓝色圆形", "kind": "Shape",
      "category": "未分类", "tags": ["圆","基础"],
      "shapeCount": 1, "width": 120.0, "height": 120.0,
      "useCount": 3, "sortOrder": 0,
      "previewUrl": "https://data.pptshapes.local/a3f9c2e1/preview.png?v=1724300000" }
  ],
  "filteredTotal": 57, "total": 213, "offset": 0, "limit": 200 }
```

- 过滤：`name/category/tags` 三字段 OrdinalIgnoreCase 包含匹配；category ≠ "全部" 时再按分类等值过滤（空分类归一化为"未分类"，对应 `GetCategory()` 行为）。
- 排序六模式与现 WPF 面板一致：`最新`=CreatedAt 倒序 / `最旧`=CreatedAt 正序 / `名称`=Name 序 / `最近使用`=LastUsedAt 倒序 / `最常用`=UseCount 倒序再 LastUsedAt / `自定义`=SortOrder 升序再 CreatedAt 倒序。
- 分页语义：offset 为已渲染条数，limit 为增量页大小（取 settings.PageSize）；`previewUrl` 由 C# 拼装，`?v=` 取 meta 文件 LastModifiedUtc 的 Unix 秒——重收藏同名预览后浏览器缓存自动失效。

### 4.4 各消息请求/响应示例（节选高频三类）

插入成功：

```json
→ {"type":"insert","reqId":"r42","payload":{"id":"a3f9c2e1"}}
← {"type":"result","reqId":"r42","ok":true,"data":{"kind":"Shape"}}
```

批量删除：

```json
→ {"type":"delete","reqId":"r43","payload":{"ids":["a3f9","b7d1"]}}
← {"type":"result","reqId":"r43","ok":true,"data":{"deleted":2}}
```

保存设置（前端头部主题按钮）：

```json
→ {"type":"saveSettings","reqId":"r44","payload":{"patch":{"theme":"dark"}}}
← {"type":"result","reqId":"r44","ok":true,"data":{"settings":{"theme":"dark","paneWidth":300,"...":"..."}}}
↓ {"type":"settingsChanged","payload":{"theme":"dark","...":"..."}}
```

---

## 5. 前端页面结构设计

### 5.1 布局分区（窄边栏 260–360px 优先）

```
┌──────────────────────────────────┐
│ ① 头部（sticky, 玻璃浮动层）      │  搜索框(🔍占位"搜索 名称/分类/标签") + 排序下拉
├──────────────────────────────────┤
│ ② 分类胶囊横滚条  [全部][圆][…]   │  overflow-x auto，激活态实心 accent
├──────────────────────────────────┤
│ ③ 状态行：共 N 项 · 筛选中 x/y/M  │  右侧：主题切换按钮(日/月图标)
├──────────────────────────────────┤
│ ④ 卡片网格滚动区                  │  grid repeat(auto-fill,minmax(88px,1fr))
│   ┌────┐ ┌────┐ ┌────┐           │  卡片 = 缩略图 + 名称单行截断
│   │img │ │img │ │ …  │           │       + 左上类型徽章(蓝=形状 紫=样式)
│   └────┘ └────┘ └────┘           │       + hover 右上淡入删除钮
│   ⋮（IntersectionObserver 哨兵）  │
├──────────────────────────────────┤
│ ⑤ 底部浮动条（互斥出现，玻璃层）  │  批量操作条 / 快速调整条
└──────────────────────────────────┘
覆盖层：⑥ 右键菜单 ⑦ 大图预览弹层 ⑧ 输入模态框(重命名/标签/分类/删除确认) ⑨ Toast
```

### 5.2 组件清单

| 组件 | 文件内实现要点 |
|------|----------------|
| SearchBox | input + 内嵌搜索图标 + 占位文案；输入 150ms debounce 后发 `list`(offset 重置 0)；Esc 清空 |
| SortSelect | 原生 `<select>` 深度换肤（appearance:none + 自绘下拉箭头 E710）；六选项文案与现面板一致 |
| CategoryChips | 横向滚动容器；点击发 `list`；当前项 `.active` |
| StatusBar | "共 M 项"；筛选态显示 "筛选 x / y · 共 M"（对齐功能 13） |
| CardGrid/CardItem | `<article class="card">`；img `loading="lazy"` + onerror 显示占位图形；徽章 `Shape→形状(蓝)/Style→样式(紫)` |
| BatchBar | 已选 n 项 · 取消 · 删除 · 移动分类；n>0 时滑入 |
| AdjustBar | 旋转±15° / 透明度± / ←↑↓→ 微调 / 完成，八键各发 `adjust` |
| ContextMenu | 自绘绝对定位菜单，九项见功能 6；视口边缘翻转防溢出 |
| Lightbox | 大图预览：图片居中 + 名称 + 类型说明 + 「插入」主按钮 + Esc/遮罩关闭 |
| Modal | 单输入框模态（重命名/标签/移动分类/批量移动）与确认模态（删除）共用骨架 |
| Toast | 右上轻提示（操作失败原因），2.5s 自动消失 |
| EmptyState | 两种文案见功能 9；配 Segoe 图标装饰（空状态用 \uE8A5 收藏星轮廓） |

### 5.3 CSS 变量表（`:root[data-theme="light|dark"]` 两套实测色值）

| 变量 | 浅色 | 深色 |
|------|------|------|
| `--bg` | `#F5F7FA` | `#14171C` |
| `--bg-decor` | `radial-gradient(at top left, rgba(0,103,192,.05), transparent 55%)` | 同左，色值 `rgba(77,163,255,.06)` |
| `--surface` | `#FFFFFF` | `#1F242B` |
| `--surface-hover` | `#F0F4FA` | `#272D36` |
| `--text` | `#171B21` | `#E8EDF4` |
| `--text-sub` | `#66707E` | `#8A94A3` |
| `--border` | `#E1E7EF` | `#303845` |
| `--accent` | `#0067C0` | `#4DA3FF` |
| `--accent-hover` | `#005AA8` | `#63B1FF` |
| `--accent-soft` | `#E3F0FB` | `rgba(77,163,255,.14)` |
| `--badge-shape` | `#0067C0`（白字） | `#2E86DE`（白字） |
| `--badge-style` | `#7B1FA2`（白字） | `#A45BC6`（白字） |
| `--danger` | `#D13438` | `#E5555A` |
| `--shadow-card` | `0 1px 2px rgba(16,24,40,.05), 0 4px 14px rgba(16,24,40,.07)` | `0 1px 2px rgba(0,0,0,.45)` |
| `--glass-bar` | `rgba(255,255,255,.72)` + `blur(12px) saturate(150%)` | `rgba(20,23,28,.66)` 同滤镜 |
| `--radius-sm/md/lg` | `6px / 10px / 14px`（两套同值） | |
| `--motion-fast` | `120ms cubic-bezier(.2,.7,.3,1)`（hover/按压） | |

字体栈：正文 `"Segoe UI","Microsoft YaHei UI","微软雅黑"`；图标类 `.ico { font-family:'Segoe Fluent Icons','Segoe MDL2 Assets'; }`。本方案使用的字形码（均为 MDL2 公开码位）：`\uE721` 搜索、`\uE710` 下拉箭头、`\uE711` 关闭、`\uE74D` 删除、`\uE700` 刷新、`\uE713` 设置、`\uE8A5` 星形（空状态装饰）。

### 5.4 关键交互实现方式

| 交互 | 实现要点 |
|------|----------|
| 多选 | Ctrl+单击 toggle `.selected` 与选中集合（跨分页保留）；无修饰键单击按 ClickToInsert 分派；批量条实时显示集合大小 |
| 右键菜单 | 卡片 `contextmenu` preventDefault → 定位菜单（右/下边缘翻转到内侧）；click/Esc/滚动任一发生即关闭 |
| 滚动分页 | IntersectionObserver 监听网格底部哨兵（rootMargin 240px）；相交且 rendered<filteredTotal 时发 `list{offset:rendered}`，响应追加渲染 |
| 拖拽排序 | 仅 sort==="自定义" 时卡片 draggable=true；dragover 用中点比较计算落点索引并显示插入指示线；drop 组装 loaded 子集的新顺序数组发 `reorder`，乐观更新、失败回滚 |
| 单击行为 | settings.clickToInsert=true → 单击即 `insert`；false → 单击开 Lightbox，Lightbox 内「插入」按钮与卡片双击均发 `insert` |
| 反馈动画 | insert/applyStyle 结果 ok 后给源卡片加 `.flash` 类：400ms 描边光晕脉冲 keyframes；TouchUsed 由 C# 在处理器内完成，前端无需二次请求 |
| 主题切换 | `html.dataset.theme` 整体换变量；入口两处：头部按钮（发 saveSettings.patch.theme）与设置对话框（C# 广播 settingsChanged）；全程无导航/刷新（约束 C5） |

---

## 6. 文件清单与现有面板处置

### 6.1 新增文件

| 文件 | 内容 |
|------|------|
| `PPTShapeLibrary\WebPaneHost.cs` | WinForms UserControl：WebView2 初始化、Runtime 检测、用户数据文件夹（C7）、双虚拟域名映射、导航失败降级信号、Debug F12 入口；实现 `ILibraryPane` |
| `PPTShapeLibrary\BridgeController.cs` | 消息路由与处理器（第 4 节协议唯一实现处）；持有 `List` 缓存与过滤/排序逻辑；对外暴露 `RefreshLibrary()` / `ApplySettings(Settings)` 转发为下行推送 |
| `PPTShapeLibrary\ILibraryPane.cs` | 接口：`void RefreshLibrary(); void ApplySettings(Settings s);` —— 让 ThisAddIn 对新旧面板零感知切换 |
| `PPTShapeLibrary\WebUI\index.html` | SPA 入口，含 `<div id="app">` 与脚本引用 |
| `PPTShapeLibrary\WebUI\css\app.css` | 全部样式：变量表（5.3）、布局、组件、动效 |
| `PPTShapeLibrary\WebUI\js\bridge.js` | 通信层：reqId 管理、pending 超时、事件订阅 API |
| `PPTShapeLibrary\WebUI\js\app.js` | 应用层：状态树、渲染函数、交互绑定（5.2/5.4） |

### 6.2 修改文件

| 文件 | 改动 |
|------|------|
| `PPTShapeLibrary\PPTShapeLibrary.csproj` | ① 新增 `<PackageReference Include="Microsoft.Web.WebView2" Version="1.0.2903.40" />`（若 VS 对该旧式工程还原异常则退回 packages.config 方案，二选一落地）；② 新增 `<Content Include="WebUI\**\*" />` 且 `CopyToOutputDirectory=PreserveNewest`；③ Compile 增加 6.1 前三个 .cs；**保留** WPF 面板的 Page/Compile 条目 |
| `PPTShapeLibrary\ThisAddIn.cs` | `_paneControl/_paneHost` 字段类型改为 `ILibraryPane`；Startup 中按 6.3 决策树实例化；其余方法体一行不动 |
| `PPTShapeLibrary\Settings.cs` | 仅追加一个可选字段 `[DataMember(Name="uiBackend", EmitDefaultValue=false)] string UiBackend`（构造默认 `"webview2"`）。**说明**：这是配置模型上的加法扩展，旧 settings.json 无此字段时反序列化走默认值，完全向后兼容；四个业务类文件不受影响 |
| `README.md` | 追加 B 阶段架构简述与 uiBackend 开关说明 |

### 6.3 宿主选择决策树（ThisAddIn_Startup 内）

```
读取 settings.UiBackend
 ├─ == "wpf"                    → 直接使用 WPF 面板（手动开关）
 └─ == "webview2"(默认)
     ├─ GetAvailableBrowserVersionString() 抛异常（Runtime 缺失）
     │    → MessageBox 提示安装地址 https://developer.microsoft.com/microsoft-edge/webview2/
     │      + Log.Error("WebView2 runtime missing") → 本次会话使用 WPF 面板
     └─ 版本可得 → new WebPaneHost()
          └─ EnsureCoreWebView2Async / NavigationCompleted 任一步失败
               → Log.Error + Toast 提示 → 降级 WPF 面板
```

### 6.4 WPF 面板处置方式

- `ShapeLibraryPaneWpf.xaml(.cs)`（含 `InputDialogWpf`、`WpfPaneHost` 类）**保留编译、作为永久降级路径**，`WpfPaneHost` 补一行接口实现（声明 `: ILibraryPane`，方法体已有同名方法无需改逻辑）。
- 自本方案合入起代码冻结：不再向 WPF 面板添加任何新功能。
- B 阶段整体验收通过后，另立清理任务评估移除时机（不在本方案范围内）。

---

## 7. 任务分解（每步独立交付、可验证）

> 所有任务的最低共同判据：`msbuild PPTShapeLibrary.sln /t:Rebuild /p:Configuration=Debug /restore` 输出 **Build succeeded**，且启动 `powerpnt.exe` 打开任意演示文稿无插件报错弹窗。日志位置 `%APPDATA%\PPTShapeLibrary\logs\app.log`。

### 任务 1：引入 WebView2 SDK 与前端资源管道
- 文件：修改 `PPTShapeLibrary\PPTShapeLibrary.csproj`
- 内容：加 PackageReference（C2 版本）、WebUI Content 复制规则、新建空的 `WebUI\index.html/css/js` 占位文件纳入编译外复制。
- 完成判据：运行上述 msbuild 命令后，`bin\Debug\` 出现 `Microsoft.Web.WebView2.Core.dll`、`Microsoft.Web.WebView2.WinForms.dll`、`WebView2Loader.dll` 与 `WebUI\index.html` 四者。

### 任务 2：WebViewPaneHost 宿主骨架
- 文件：新建 `WebPaneHost.cs`、`ILibraryPane.cs`；临时接线 `ThisAddIn.cs`（直接 new WebPaneHost，暂无降级分支）
- 内容：环境创建（userDataFolder=C7）、双虚拟域名映射、加载 index.html 占位页、Debug F12。
- 完成判据：启动 PowerPoint 打开收藏窗格，显示占位页文字"WebView2 OK"；app.log 出现 `WebView2 runtime: <版本号>` 与 `Pane ready (WebView2)` 两行。

### 任务 3：BridgeController 只读链路（ping/init/list/getSettings/log）
- 文件：新建 `BridgeController.cs`
- 内容：信封解析、reqId 配对、`list` 的过滤/排序/切片（规则照抄 4.3，逐条对照现 `RefreshCore()` 实现）、`previewUrl` 拼装。
- 完成判据：窗格内按 F12 打开 DevTools，Console 执行
  `chrome.webview.postMessage(JSON.stringify({type:'list',reqId:'t1',payload:{filter:'',category:'全部',sort:'最新',offset:0,limit:50}}))`
  能观察到 `ok:true` 且 items 数与实际收藏数一致（对照状态栏旧面板计数）。

### 任务 4：前端骨架 + 主题系统 + 卡片网格渲染
- 文件：`WebUI\css\app.css`、`WebUI\js\app.js`、`index.html`
- 内容：5.1 五个分区布局、5.3 全套 CSS 变量双主题、CardGrid 渲染 `list` 结果、徽章/名称/懒加载缩略图/空状态两文案/状态行计数、settingsChanged→换 theme 不刷新。
- 完成判据：窗格显示真实卡片网格；设置对话框切换深色→界面整体变色且 DevTools Console 执行 `performance.getEntriesByType('navigation').length` 保持 1（未发生重新导航）。

### 任务 5：浏览类交互（搜索/分类/排序/分页/单击行为/Lightbox）
- 文件：`js/app.js`
- 内容：SearchBox debounce、CategoryChips、SortSelect 六模式、IntersectionObserver 分页、ClickToInsert 双模式、Lightbox 大图与双击插入。
- 完成判据：手工脚本——① 输入某标签关键字，300ms 内列表变化；② 点分类胶囊列表收敛；③ 切"最常用"首位为 UseCount 最大项；④ 收藏 >PageSize 时滚到底部自动追加一页；⑤ ClickToInsert=false 时单击出大图、双击插入成功。

### 任务 6：管理类交互（多选/批量/右键/模态框/delete/rename/setTags/setCategory/batchMove/openFolder/reorder）
- 文件：`js/app.js`、`BridgeController.cs` 补齐对应处理器
- 内容：Ctrl 多选、BatchBar、自绘右键菜单九项、三种输入模态+删除确认模态、自定义排序拖拽（HTML5 DnD + 插入指示线 + `reorder` 乐观更新）。
- 完成判据：① Ctrl+单击三张卡，批量条显示"已选 3 项"；② 批量移动到新分类后在胶囊条出现新分类；③ 右键九项逐一点击行为正确；④ "自定义"模式下拖拽两张卡后重启 PowerPoint 顺序保持（SetSortOrder 已落盘）。

### 任务 7：插入链路与快速调整条
- 文件：`js/app.js`、`BridgeController.cs` 补 insert/applyStyle/applyStyleSameType/adjust
- 内容：insert 成功→卡片 `.flash` 动画+AdjustBar 滑入；八键 adjust；样式类收藏走 applyStyle；失败→Toast 显示 error 字段。
- 完成判据：① 单击形状卡→画布出现形状、卡片闪一下、调整条出现；② 连点两次"旋转+"形状累计旋转 30°；③ 点"完成"调整条收起；④ app.log 中对应 meta 的 UseCount 递增（重新拉 list 可见 useCount+1）。

### 任务 8：降级链路与运行时检测
- 文件：`ThisAddIn.cs` 决策树（6.3）、`Settings.cs` 加 UiBackend 字段、`WebPaneHost.cs` 失败信号
- 内容：Runtime 缺失检测（try GetAvailableBrowserVersionString）、uiBackend="wpf" 手动开关、初始化失败自动降级。
- 完成判据：① 手工把 `%APPDATA%\PPTShapeLibrary\settings.json` 加 `"uiBackend":"wpf"` 重启→显示旧 WPF 面板，log 含 `uiBackend=wpf, legacy panel active`；② 删掉该字段重启恢复 Web 面板；③（有条件则在未装 Runtime 的干净虚拟机）缺失场景弹出含官方下载链接的提示框且旧面板正常可用。

### 任务 9：发布管道验证与文档
- 文件：`README.md`、必要时 csproj 发布属性微调
- 内容：ClickOnce 发布一次（生成 publish\），确认 `Application Files` 内包含 `WebUI\**` 与 `WebView2Loader.dll.deploy`；README 补架构图与开关说明。
- 完成判据：`msbuild /t:Publish` 后检查 publish\Application Files\PPTShapeLibrary_*\ 目录含上述文件；安装该 vsto 后窗格功能与 Debug 一致。

---

## 8. 功能清单（验收合同）

> P0 = 必须达成否则验收不通过；P1 = 本阶段一并交付；P2 = 可选项。验证列给出可直接执行的步骤。

- [ ] P0 宿主接入：任务窗格由 WebView2 控件（WebPaneHost 内嵌）承载，初始宽度沿用 settings.PaneWidth（现状仅设置对话框会持久化该值，拖拽边条不写回，行为保持不变）。验证：设置对话框把宽度改为 360 并保存→窗格立即变宽；重启 PowerPoint 后仍为 360。
- [ ] P0 运行时检测与降级：Runtime 缺失或初始化失败自动回退 WPF 面板；settings.json 写 `"uiBackend":"wpf"` 可手动回退。验证：任务 8 判据 ①②③。
- [ ] P0 虚拟域名离线加载：页面与缩略图分别经 `app.` / `data.` 虚拟域加载。验证：断网重启 PowerPoint 功能完好；DevTools Network 面板无 file:// 且无任何外网请求。
- [ ] P0 消息通道：第 4 节全部 17 种上行类型 + 2 种下行推送联调通过。验证：任务 3 判据 + 任务 6/7 全流程。
- [ ] P0 卡片网格：缩略图+名称（超长省略号）+类型徽章（形状=蓝/样式=紫）；单击行为由 ClickToInsert 决定（true=单击插入 / false=单击大图+双击插入）。验证：两种设置下分别操作，行为符合且另一行为不触发。
- [ ] P0 搜索：实时过滤（150ms debounce），匹配名称+分类+标签，带占位提示。验证：分别用名称片段、分类名、标签名搜索均命中同一目标卡。
- [ ] P0 分类条：横向胶囊含"全部"，点击筛选，空分类归一"未分类"。验证：给一个收藏设分类后胶囊条即时出现该分类。
- [ ] P0 排序：最新/最旧/名称/最近使用/最常用/自定义 六模式结果正确。验证：造 3 条 UseCount/CreatedAt 可区分的数据逐一核对顺序。
- [ ] P0 多选与批量：Ctrl+单击切换选中（跨分页保留）；底部条显示已选数，取消/删除（带确认）/移动分类三操作生效。验证：跨两页选 4 项→批量移动→新分类计数 +4。
- [ ] P0 右键菜单九项：插入到当前页/应用样式到选中形状/应用到本页同类形状/预览大图/移动到分类/编辑标签/重命名/删除/打开收藏文件夹。验证：逐项点击，行为与旧 WPF 面板一致（打开收藏夹应定位到 `%APPDATA%\PPTShapeLibrary\items\{id}`）。
- [ ] P0 分页：超过 PageSize 首屏只渲染前 PageSize 张，滚动近底部自动追加。验证：设 PageSize=50，准备 ≥110 条，首屏 DOM 卡片数=50，滚动三次后 ≥150。
- [ ] P0 快速调整条：插入成功出现，旋转±15°/透明度±/四向微调/完成 八键驱动 `AdjustSelected`。验证：任务 7 判据 ②。
- [ ] P0 空状态：无收藏显示引导文案（提示去工具栏收藏），筛选无结果显示"没有符合筛选条件的收藏"。验证：清空 items 目录看前者；搜不存在词看后者。
- [ ] P0 主题：跟随 settings.Theme，深浅切换不刷新页面（CSS 变量）。验证：任务 4 判据 navigation.length===1 + 两主题截图对比变量表色值。
- [ ] P0 操作反馈：插入/应用成功源卡片闪烁，且 UseCount 经 TouchUsed 递增。验证：任务 7 判据 ③。
- [ ] P0 拖拽排序：自定义模式下拖拽重排→SetSortOrder 落盘，其他排序模式下禁用拖拽。验证：任务 6 判据 ④ + 切到"最新"时拖拽不触发。
- [ ] P0 总数统计：状态行显示"共 M 项"，筛选态显示 x/y·M。验证：与 items 目录实际子目录数比对一致。
- [ ] P1 CardSize 生效：compact/standard/large 映射网格列宽 84/104/128px。验证：设置对话框切换三档，卡片目测变档。
- [ ] P1 前端诊断日志：未捕获 JS 异常经 `log` 消息写入 app.log。验证：DevTools 人为抛错一次，app.log 出现 `[webview]` 前缀记录。
- [ ] P1 DevTools 入口：Debug 构建 F12 打开 DevTools，Release 无效。验证：两配置各构建一次按键测试。
- [ ] P2 头部主题快捷按钮：不进设置对话框即可切深浅（仍走 saveSettings 持久化）。验证：点按钮后重启 PowerPoint 主题记忆保持。

统计：P0×17，P1×3，P2×1，共 21 项。

---

## 9. 整体验收标准

### 9.1 构建与部署

```powershell
cd E:\code\PPTShapeLibrary
msbuild PPTShapeLibrary.sln /t:Rebuild /p:Configuration=Debug /restore /nologo /v:m
# 预期：Build succeeded；bin\Debug\ 含 PPTShapeLibrary.dll/.vsto、Microsoft.Web.WebView2.Core.dll、
#       Microsoft.Web.WebView2.WinForms.dll、WebView2Loader.dll、WebUI\ 四件套
start powerpnt.exe   # 打开任意 pptx，Ribbon「形状库」→「收藏库」
```

### 9.2 功能验收

逐项勾验第 8 节 21 个复选框；其中 P0 17 项全部通过为验收通过底线，P1/P2 允许遗留但须记录。

### 9.3 性能要求

| 指标 | 标准 | 测法 |
|------|------|------|
| 首屏时间 | 从 `_taskPane.Visible=true` 到首批卡片可见 < 1000ms（200 条收藏规模） | app.log 中 `Pane ready` 时间戳 vs DevTools 渲染完成帧；连续冷启 3 次取中位 |
| 滚动流畅性 | 500 条渲染后滚动无可感知卡顿（无 >100ms 长帧） | DevTools Performance 录制 10s 滚动，Long Task 数 ≤2 |
| 主题切换 | < 100ms 完成、无白闪、无导航 | 同任务 4 判据 |
| 消息延迟 | 单次 list(200条) 往返 < 80ms | bridge.js 内置耗时打点经 log 消息输出 |

### 9.4 回归与兼容

- 旧 WPF 面板回归：`uiBackend:"wpf"` 下原 13 功能抽测（收藏/插入/主题/搜索）不劣化——降级路径必须是真保险。
- 数据兼容：B 阶段前后 items 目录、settings.json 二进制级结构不变；升级安装后老收藏全部可见。
- 离线验收：拔网线完整跑一遍 9.2 抽样（插入/搜索/主题/预览）。
- 多实例：同时开两个 PowerPoint 窗口进程，各自窗格互不干扰、无 UDF 锁冲突报错。

---

## 10. 风险与回退

| # | 风险 | 概率/影响 | 缓解与回退 |
|---|------|-----------|------------|
| R1 | 目标机 WebView2 Runtime 缺失（企业精简镜像、LTSC） | 中/高 | 启动期 `GetAvailableBrowserVersionString` 探测；缺失→官方下载链接提示框+WPF 面板兜底，功能无损仅视觉降级 |
| R2 | COM 互操作时序：幻灯片放映中/模态对话框打开时收到 insert 等消息，`ActiveWindow` 为放映窗或被拒 | 中/中 | 处理器 try/catch+Log；失败 result 回传前端 Toast"请退出当前放映后重试"；业务方法本身已有的防御保持不动 |
| R3 | WebView2Loader.dll（native）未随 ClickOnce 发布，装机后白屏 | 中/高 | 任务 9 强制发布产物校验；白屏场景 NavigationCompleted 失败自动降级 WPF 并留日志 |
| R4 | 多 PowerPoint 进程共享用户数据文件夹冲突 | 低/中 | 固定 UDF（C7）为 Evergreen 默认支持的共享形态；仍冲突时 log 明确指向 UDF，可在 settings 增设独立 UDF 字段（预留，不做） |
| R5 | 旧式工程 PackageReference 还原异常 | 低/低 | 退回 packages.config + HintPath 方案（6.2 已注明），不影响其余任务 |
| R6 | 性能不达标（首屏>1s） | 低/中 | 首屏只请求一页 list；缩略图 lazy+HTTP 缓存；必要时将 WebView2 环境预热提前到 Startup（当前已在 Startup 创建，天然预热） |

总体回退开关：`settings.json → "uiBackend":"wpf"` 一键回到现状，随时可逆。

---

## 11. 开放问题

1. ClickOnce 正式分发是否需要捆绑 Fixed Version WebView2 Runtime 以彻底消除 R1？（建议先靠检测+提示运行一轮，视目标用户反馈决定，不阻塞本方案）
2. P2 的头部主题快捷按钮若与设置对话框产生认知歧义（两处入口），是否改为仅跳转设置？（开发期可先做，验收时由产品拍板保留与否）
