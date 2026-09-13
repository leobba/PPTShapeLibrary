# PowerPoint COM 经验与踩坑记录（收藏/插入外观一致性）

> 记录 2026-09-13 这一轮排查的结论。**这些坑都是实际发生过的**，改 `ShapeExporter.cs` 前请先读一遍，
> 不要重复"逐形状刷属性"这类看起来合理、实际会踩 COM 语义坑的做法。

## 一、总体结论（重要）

**收藏/插入外观要一致，正确做法只有两条：**

1. **不要逐形状去刷属性**（填充、线条、文字、阴影…）。
   COM 的属性读取结果与"文件里实际怎么写"经常不一致（见下面第 3、4 条），刷出来的结果会**破坏原样**。
2. **让收藏文件自带源文稿的"设计/主题"**，并让插入时**必须导入它**：
   - 源母版的 **12 个主题色**整体复制到工作文稿主题（`CopyThemeColors`）；
   - 把工作文稿的 **设计/母版/版式改成唯一名字**（`MakeDesignUniqueForLibrary`，`Lib<id>-D/M/L`），
     这样目标文稿匹配不到同名设计，只能导入我们文件的设计。

## 二、已确认的机制（含证据）

### 1. 往"新建的空白文稿"里粘贴幻灯片，会丢失源设计
PowerPoint 按 **设计/主题名字** 匹配（源模板与空白文稿都叫「Office 主题」），于是**直接复用空白文稿的默认主题/母版/版式**。
实测（同一页）：

| 部件 | 源模板 | 粘贴到空白文稿后的产品 |
|---|---|---|
| `theme1.xml` | 4523 B | 8739 B（被换成默认 Office 主题） |
| `slideMaster1.xml` | 10266 B | 14108 B |
| 该页引用的版式 | 空白版式 1108 B | 带占位符的版式 4085 B |

→ 于是出现"占位符继承到线条 = 多出边框""主题色被重新解析 = 颜色 A→B"。

### 2. `Slides.InsertFromFile` **不搬版式绑定**（已弃用）
用它把页面插进临时文稿时，插入页会被重绑到目标文稿的默认版式 → 同样出现边框。
收藏/插入请用 **`Slide.Copy()` → `Slides.Paste()`**（复制幻灯片会连版式绑定一起带过去）。

### 3. "无填充"形状会被 COM 误报为"有实心填充" ⚠️
源文件里无填充写法是 `<a:noFill/>` **加** `<a14:hiddenFill>`（PowerPoint 的隐藏填充）。
此时 `Shape.Fill.Visible` / `Fill.Type` 会给出"实心填充"的假象。
**后果**：按此刷填充，会把无填充形状刷成**实心色块**（用户看到"收藏/插入后多出边框"）。
实测：第 1 页形状 1/2 为 `<a:noFill/>`，被我刷成 `<a:solidFill><a:srgbClr val="262626"/>`（深灰块）。

### 4. 渐变线条（`<a:ln><a:gradFill>`）不能被写"线条颜色"
很多模板用"渐隐到透明"的渐变线条，**视觉上等于没有边框**。
若按 `Line.Visible == msoTrue` 去写 `Line.ForeColor.RGB`，渐变会被**压成实线** → 凭空出现边框。

### 5. 剪贴板：**不要**回放 Office/OLE 专有格式 ⚠️
把 `GetData(fmt, true)` 取到的全部格式（含 `EnhancedMetafile`、`Object Descriptor`、Office 私有数据等）
塞进托管 `DataObject` 再 `SetDataObject` 回放，会在**封送阶段损坏堆**，抛
`FatalExecutionEngineError 0xc0000005`（**无法被 try/catch 捕获，PowerPoint 直接崩**）。
`ClipboardGuard` 因此只保护"文本/图片/文件"，**不保护 PowerPoint 形状数据**——
代价是：收藏/插入时用户剪贴板里的 PPT 形状会被内部复制覆盖（已知限制）。

### 6. 长期持有 `Presentation` 句柄有风险
曾做"库文稿句柄池"（复用已打开的库 pptx 省去每次约 700ms 的 `Presentations.Open`）。
跨多次操作复用 COM 句柄是崩溃高危区，出问题后已回退为「每次 Open → 用 → Close」。
（句柄池版本保留在分支 `backup/fidelity-work`，另有"复制源文件→只留目标页"的完整保真方案与实测数据。）

### 7. 任务窗格会被 PowerPoint 删除，必须探活
窗口关闭/迁移后 `_taskPane` 指向的对象已被删除，之后访问 `.Visible` 抛 `0x800A01A8`
（现象：**点「素材库」按钮打不开**）。`ThisAddIn.PaneIsAlive()` 负责探活并重建。

### 8. 其它已修的坑
- 整页存档曾固定 `ppSlideSizeOnScreen`(4:3)，把 16:9 模板压变形 → 现沿用**源页面尺寸**。
- 预览曾固定 640×480 → 现按页面真实宽高比导出（960×宽）。
- **不要**删除"未使用版式"（`CleanUnusedLayouts`）：其"哪些版式被使用"的判断不可靠，
  会删掉本页真正使用的版式 → PowerPoint 重绑到别的版式 → 外观走样。该设置项现已停用。

### 9. 「插件在 PPT 里消失」：PowerPoint 的加载项子系统初始化失败 ⚠️（2026-09-13 实战）
现象：**功能区没有「素材库」**，`app.log` 里连 `Startup begin` 都没有（代码一行没跑）；
但注册项、清单文件哈希、VSTO 信任项、禁用项列表全部正常。

**判定方法（可复现）**：用 VBScript 附加到正常启动的 PowerPoint 读加载项集合
（PowerShell 7 没有 `Marshal.GetActiveObject`，用 `cscript` + `GetObject(, "PowerPoint.Application")`）：
```
Application.COMAddIns : Invalid request.  Office AddInsX is not initialized.
```
→ 说明 **PowerPoint 的加载项子系统根本没初始化**（此时任何加载项都不加载，连系统自带的 OneNote 也不加载）。
对照实验：同样的探测在 **Word** 里返回 `COMAddIns.Count = 4`（正常）→ **PowerPoint 侧问题**。
再对照：**把我们的注册项整个删掉后依然报同样错误** → 与我们的插件无关。

**根因与修复**：注册项只写在**版本化路径**时会出现该问题：
```
HKCU\Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary   ← 只写这里 → AddInsX 未初始化
```
把注册项补写到**版本无关路径**（Office 文档记载的每用户 COM 加载项目录，系统自带 OneNote 加载项就在这里）后，
加载项子系统**立刻恢复**（`COMAddIns.Count` 由报错变为 `1`），插件正常加载（日志出现 `Startup begin`）：
```
HKCU\Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary        ← 关键，必须有
```
→ `scripts\dev-register.ps1` 现在**两处都写**；注销时两处一起删。

**顺带排除的干扰项**（查过、都不是原因）：
- `HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes`（`comaddins=1` 等）：
  8 月 14 日写入、非域策略，且插件当天白天仍能加载 → 属遥测范围，**不是**加载拦截。
- 程序集未强名称签名（`PublicKeyToken=null`）而清单声明了 token：项目从无强名称签名设置，之前同样能加载 → 无关。
- `HKCU\...\16.0\PowerPoint\AddInLoadTimes` 缓存、空的 `Resiliency` 键：删除后无效 → 无关。
- 任务管理器里的加载项禁用列表（`Resiliency\DisabledItems`）：为空 → 无关。

## 三、排查套路（推荐复用）

1. **别猜机制，读文件**：收藏条目是 `%APPDATA%\PPTShapeLibrary\items\<id>\shape.pptx`（本质是 zip）。
   把 `ppt/slides/slideN.xml` 与**源文稿同一页**做"归一化空白后逐字节 diff"，一眼看出哪处 XML 被改坏。
2. **同一文稿内对比**：让用户把"原始页 + 插入页"存在同一个 pptx 里，两页逐字节 diff——
   主题、版式、继承关系全部相同，差异**必然**来自我们写入的内容。
3. **确认代码真的在跑**：`bin\Debug\PPTShapeLibrary.dll` 的时间戳必须**早于**用户启动 PowerPoint 的时间；
   改完代码必须**完全退出并重开 PowerPoint**（历史上多次"修复无效"其实是在跑旧 DLL）。
4. **对照收藏条目与源**：线条/版式/主题色逐项比对（本仓库用过的检查脚本思路见下）：
   - 线条：`<a:ln>` 是 `noFill` / `solidFill` / `gradFill` / 缺失
   - 版式绑定：`ppt/slides/_rels/slide1.xml.rels` 的 `slideLayout` 指向，以及该版式的 `<p:sp>` 数量
   - 主题色：`ppt/theme/theme1.xml` 的 `clrScheme` 12 项与源是否一致
   - 设计名：是否带 `Lib<id>-` 前缀（说明唯一改名生效）
