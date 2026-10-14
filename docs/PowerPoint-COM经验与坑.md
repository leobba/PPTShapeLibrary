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

### 2. `Slides.InsertFromFile` 与剪贴板路径的等价性（2026-10 实测，结论已更新）
**旧结论（已作废）**："InsertFromFile 不搬版式绑定，已弃用"。
**新结论**：在**库文件的设计/母版/版式已唯一改名**（`MakeDesignUniqueForLibrary` → `Lib<id>-*`）的前提下，
`InsertFromFile` 与 `Slide.Copy() → Slides.Paste()` 的产物**逐项一致**（证据见第 13 节），
因此整页插入现在**只用 `InsertFromFile`**，完全不再碰剪贴板。
若哪天取消设计唯一改名，这个等价性会立刻失效（目标文稿按名复用默认设计 → 版式被重绑 → 外观走样）。

**保留的历史教训**：设计名相同（例如源模板与空白文稿都叫「Office 主题」）时，PowerPoint 会按名复用目标设计，
于是粘贴/插入的页面被重绑到目标默认版式 → 多出边框、颜色漂移。

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

**✅ 根因已查明并修复（2026-10）**：**一次 `msbuild /t:Rebuild /p:Configuration=Release` 就会删掉加载项注册。**

机制（读 VS 的 `Microsoft.VisualStudio.Tools.Office.targets` 得到）：
```
VSTOClean  →  UnregisterOfficeAddin  →  SetOffice2007AddInRegistration(Unregister="true")
```
`Clean`/`Rebuild` 都会跑 `VSTOClean`，而 `UnregisterOfficeAddin` **不看 `RegisterForVSTO`**，
无条件注销加载项。项目此前只在 csproj 里重定义了 `RegisterOfficeAddin`（Release 不注册），
**没有对称地挡住 `UnregisterOfficeAddin`** → Release 下"只注销、不注册" = 注册被删。

实测（修复前后各跑一遍，两处 HKCU 键逐次采样）：

| 操作 | 修复前 | 修复后 |
|---|---|---|
| `dev-register.ps1` | 版本无关=True 版本化=True | 同左 |
| `msbuild /t:Rebuild /p:Configuration=Release` | **版本无关=False** 版本化=True | 版本无关=True 版本化=True |
| `msbuild /t:Rebuild /p:Configuration=Debug` | 版本无关=True（Debug 会重新注册，所以看起来"恢复了"） | 同左 |
| `msbuild /t:Clean /p:Configuration=Release` | — | 版本无关=True 版本化=True |

**修复**：在 csproj 里同样重定义 `UnregisterOfficeAddin`，Release 下置空（Debug 保留原行为）。
**这解释了长期现象**：每次发版/Release 构建之后，用户机器上功能区就"没有素材库"了——
不是环境玄学，是构建脚本自己把注册删了。

**为什么之前查不出来**：我先后排除了 Debug 构建、正常启停、COM `Quit()`、强杀、无窗口自动化实例、
逐阶段采样，甚至用 3 秒间隔的监视器盯了 35 分钟——**都复现不了**，因为那些实验里没有一次 Release 构建。
最后用"二分法"逐条跑构建命令才定位到。教训：**"偶发/不明原因"的问题，先把所有可疑动作列出来逐个二分，
不要靠猜机制。**

**保留的兜底（纵深防御，仍然有用）**：
1. `scripts\e2e-smoke.ps1` 前置检查发现缺键**自动跑 `dev-register.ps1` 修复**，跑完再校验一次；
2. `scripts\dev-register.ps1 -Machine`（管理员）额外写 **HKLM**（版本无关 + 版本化两处）；
3. 安装器新增**「修复加载项注册」**按钮：普通用户直接重写 HKCU 两处键，可选提权写 HKLM（`--repair-machine` 子模式）；
4. 人工遇到"功能区没有素材库"时，先跑 1 或点 3，再重启 PowerPoint。

**顺带排除的干扰项**（查过、都不是原因）：
- `HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes`（`comaddins=1` 等）：
  8 月 14 日写入、非域策略，且插件当天白天仍能加载 → 属遥测范围，**不是**加载拦截。
- 程序集未强名称签名（`PublicKeyToken=null`）而清单声明了 token：项目从无强名称签名设置，之前同样能加载 → 无关。
- `HKCU\...\16.0\PowerPoint\AddInLoadTimes` 缓存、空的 `Resiliency` 键：删除后无效 → 无关。
- 任务管理器里的加载项禁用列表（`Resiliency\DisabledItems`）：为空 → 无关。

### 9.8 ⚠️ VSTO 的"已安装另一版本"到底记在哪（2026-10 彻底查清）

**这是本项目最坑的一条**：`VSTOInstaller /Install` 报 `AddInAlreadyInstalledException`
（"由于已经安装了另一版本的自定义项且不能从该位置升级"），表现是**插件突然不加载**、装了也没用。

我依次清掉下面这些都**无效**（VSTOInstaller 依旧报同一个错）：

| 位置 | 说明 |
|---|---|
| `HKCU\…\PowerPoint\Addins\PPTShapeLibrary`（两处） | 加载项注册本身 |
| `HKCU\…\PowerPoint\AddinsData\PPTShapeLibrary` | 加载计数缓存 |
| `HKCU\Software\Microsoft\VSTO\SolutionMetadata\<GUID>` | 有 `addInName=PPTShapeLibrary` 的那条 |
| `%LOCALAPPDATA%\Apps\2.0\…` 部署缓存 | 移走会被 VSTOInstaller 重建 |
| `Deployment\SideBySide\2.0\{Installations,Components,Assemblies,Categories,…}` | 全是空 |
| ARP 卸载项 | 删了也不够 |

**真正的判定来源是 `HKCU\Software\Microsoft\VSTA\Solutions\<GUID>`** —— 它记着
`ProductName=PPTShapeLibrary` + `Url` + `ApplicationID`（含版本号）。
**只要这条还在，VSTO 就认为"该产品已装过某版本"，任何位置都装不进去。** 删掉它立刻就能装。

实测证据：删掉 `VSTA\Solutions\8c0b7721-…` 后，`VSTOInstaller /Install` 从 `-400` 变成
"已成功安装 Microsoft Office 自定义项"。

**配套的另外两个坑**（同一轮排查发现）：
1. **ARP 里的开发版残留**（`UninstallString` 指向 `bin\Debug`）→ 对它跑
   `VSTOInstaller /Uninstall` 会**挂死**（实测卡满 120s 超时）。这种记录只能直接删键。
2. **卸载串里是正斜杠**（`file:///E:/…/bin/Debug/…`）→ 判"是否开发路径"时若按 `bin\Debug` 匹配会**漏判**
   （踩过），必须先统一分隔符。

**已内置的处理**：
- `scripts\dev-register.ps1`：注册前/注销时都会清 `VSTA\Solutions` 与本仓库的 VSTO 信任项、
  以及指向开发路径的 ARP 记录（`Clear-VstaRecords` / `Clear-DevArpRecords`）。
- 安装器 `--pending-update`（一键更新）：装新版前**无条件**清 VSTA 记录。
- **构建逃生开关**：VSTO 元数据存储损坏时 `msbuild` 会在注册/注销步骤抛
  `MSB4018 SetOffice2007AddInRegistration → ClearMetadataStore → FileNotFoundException`，
  用 `msbuild /p:VstoSkipRegistration=true` 跳过这两个步骤即可正常构建（注册交给 `dev-register.ps1`）。

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
## 四、2026-10 新增的实证结论（整页插入 / 跨文稿剪贴板 / E2E 验证）

### 9.5 长任务期间"取消"送不进去：`Application.DoEvents()` 对 WebView2 消息无效 ⚠️

**现象**：整稿操作（收藏整稿 / 重建缩略图 / 整稿导出图片）在 UI 线程上跑长循环，
期间前端点"取消"完全没反应。

**实测（关键数据）**：在循环里每 150ms 调一次 `Application.DoEvents()`，
一次 33 页的导出过程中 `DoEvents` 实际执行了 **6~7 次**，
但同期宿主收到的桥消息数为 **0**（`WebMessageReceived` 一次都没触发）。
→ 结论：**DoEvents 不能把 WebView2 的上行消息送进被占住的 UI 线程**，
"泵消息让取消生效"这条路是死的。

**采用的做法**：长任务开始时，宿主在 **127.0.0.1 上开一个临时端口**（裸 `TcpListener`，
随机端口 + 随机令牌，只认 `GET /cancel?t=…`），把 URL 随 `progress` 一起下发给前端，
前端 `fetch(url)` 完成取消；任务结束立刻关闭监听。
- 不用 `HttpListener`：绑定 URL 前缀通常需要管理员或 `netsh urlacl`；裸套接字普通用户即可绑定回环端口。
- 安全边界：仅回环 + 随机端口 + 随机令牌 + 只认 cancel + **只在长任务期间存活**。
- CSP 需要放行：`connect-src 'self' http://127.0.0.1:*`（实测不放行会被 CSP 直接拒掉，
  浏览器控制台报 `violates ... "connect-src 'self'"`）。
- 既然 DoEvents 无用，**代码里就不再泵消息**——顺带消除了"长任务中途被用户操作重入"的风险。

**通用教训**：Office 加载项里"让 UI 保持响应"的常规手段（DoEvents）在 WebView2 场景下
不保证消息送达；需要真正异步的信号通道时，回环 socket 是可靠且低权限要求的方案。

### 9.6 库条目会变成 `ActivePresentation`，整稿工具会作用到用户的素材文件上 ⚠️

**现象（实测日志）**：
```
[pane] WindowActivate ignored (library presentation) caption=[shape.pptx - 最后由用户保存]
export images: active=shape.pptx slides=1 path=…\PPTShapeLibrary\items\<id>\shape.pptx
```
于是"导出图片/换字体/统一尺寸/清理瘦身/页脚页码…"全部作用到**用户收藏的素材文件**上，
写操作会直接把收藏改坏，而用户以为在改自己的文稿。

**触发路径**：PowerPoint 的**文档恢复**——上次 PowerPoint 被强制结束时，
句柄池里还开着某个 `items\*\shape.pptx`，下次启动 PowerPoint 会把它当"未正常关闭的文档"
重新打开（**读写 + 有窗口**），于是它成了活动文稿。
（对照实验：`Presentations.Open(..., WithWindow:=False)` 不会产生窗口、也不会抢活动状态；
只有带窗口打开才会。）

**修复（两层）**：
1. `ThisAddIn.OnWindowActivate` 发现库文稿窗口激活时，除了不迁移窗格，还会
   `ActivateUserWindow()` 把焦点弹回用户文稿；
2. 服务层一律通过 `ShapeExporter.ActiveUserPresentation()` 取文稿——它**绝不返回库条目**，
   找不到用户文稿就返回 null（宁可报"没有打开的演示文稿"，也不去动素材文件）。

**顺带**：调试时不要用强杀结束 PowerPoint，否则下次启动就可能进入文档恢复状态。
`scripts\e2e-smoke.ps1` 里已内置"打开一个带窗口的库条目"来复现并回归这条防护。

### 9.7 让长任务不卡死 PowerPoint：**分片执行**（2026-10 落地，实测有效）

**问题**：整稿操作（导出图片/换字体/收藏整稿…）在 UI 线程上跑长循环，期间 PowerPoint 完全卡死。

**试过但无效**：在循环里 `Application.DoEvents()`——实测 33 页导出期间调用了 6~7 次，
PowerPoint 是否重绘不确定，而 WebView2 的上行消息（取消/ping）**一条都送不进来**（见 §9.5）。

**采用的做法：把长循环切成片，片与片之间回到消息循环。**
- 服务侧把长循环改造成作业对象：`XxxJob.Create(...)` / `Step(i)` / `Finish(cancelled)`；
  同步入口（Ribbon 等）保留为"循环跑完 job"，行为完全一致。
- 桥侧 `RunChunked(reqId, title, total, slice, finish)`：每片在 UI 线程上执行一小步，
  然后 `SynchronizationContext.Post` 投递下一片 → 线程回到消息循环。
- **协议不变**：`RunChunked` 不立即下发 result，全部完成（或取消）后异步发出，前端仍按 reqId 等待。

**实测效果（33 页文稿）**：
| 指标 | 结果 |
|---|---|
| 长任务期间 `ping` | **被即时答复**（比导出完成早约 2 秒）→ 证明宿主确实在响应 |
| 取消（普通桥消息） | 第 2 页即生效（"已取消，已导出 2 / 33 页"） |
| 取消（回环端口 fetch 兜底） | 同样生效 |
| 长任务中再发请求 | **被明确拒绝**并提示"正在执行「导出图片」，请等待完成或点击「取消」"（不插队、不静默丢失） |
| 进度 | 逐片上报（34 条），前端进度条可用 |

**已分片（7 个）**：整稿导出图片 / 一键换字体 / 收藏整稿 / 重建缩略图 / 统一尺寸 / 清理瘦身（执行）/ 清理扫描（干跑）。
**仍同步**（只有进度+取消，界面会卡）：查找替换 / 文字清理 / 整页应用样式 / 图片网格。
后几个要么是单次遍历很短的文本操作，要么本身会弹模态框，暂不需要分片。

**注意**：分片执行意味着长任务期间用户可以继续操作 PowerPoint。作业里因此都要
"每片重新确认目标是否还存在"（例如 `if (i + 1 > _pres.Slides.Count) return;`、
清理作业里 `slideIndex` 越界就跳过），不能假设切片之间文档没变。
破坏性作业（清理瘦身）额外守则：**读不到就保守处理**（`HasVisibleContent` 读异常一律判"有内容"，
宁可少删一页也不误删）。

### 10. 幻灯片插入的索引语义：`Slides.Paste(Index)` 与 `Slides.InsertFromFile(Index)` 完全不同 ⚠️

**实测（PowerPoint 16.0.20430，2 页文稿）：**

| 调用 | Index 含义 | 合法范围 | 传 `Count+1`（想"追加到末尾"） |
|---|---|---|---|
| `Slides.Paste(Index)` | 插到第 Index 页**之前** | 1 … Count+1 | ✅ 合法（3 页文稿传 3 = 追加） |
| `Slides.InsertFromFile(file, Index, s, e)` | 插到第 Index 页**之后** | 0 … Count | ❌ 抛 `Integer out of range. 3 is not in FileName's valid range of 0 to 2` |

实测记录（2 页文稿）：
```
Slides.Paste(3)              -> OK，新页落在第 3 页
InsertFromFile(idx=0)        -> OK，新页成为第 1 页
InsertFromFile(idx=1)        -> OK，新页成为第 2 页
InsertFromFile(idx=2=Count)  -> OK，新页追加到末尾
InsertFromFile(idx=3)        -> FAIL "not in FileName's valid range of 0 to 2"
```
**踩坑**：`InsertPageSnapshot` 的"免剪贴板兜底通道"曾对两种 API 用同一个 `Count+1`，
导致兜底在"插到末尾"时**必然失败**（等于兜底失效）；"插到当前页之后"也偏了一页。
现在 `Paste` 仍用 `Index+1`，`InsertFromFile` 用 `当前页序号 / Count`，并在 `InsertFromFile`
里另做一次夹紧。

### 11. 跨文稿 `Slide.Copy()` → `Paste()` 不可靠（同一进程、无第三方工具也复现）
**适用范围已缩小**：整页插入自 2026-10 起改走 `InsertFromFile`，已不使用剪贴板；
本条现在只对**形状级插入**（`ShapeRange.Copy() → Shapes.Paste()`）有效。

用 `Presentations.Open(..., WithWindow:=msoFalse)` 打开库文件后 `libPres.Slides(1).Copy()`，
再 `activePres.Slides.Paste(n)`：实测**两种索引都报**
`Clipboard is empty or contains data which may not be pasted here`，或返回空 SlideRange。
`Slide.Copy()` 是异步投递剪贴板，跨文稿时更容易丢。

应对（当前实现）：
1. 复制→粘贴之间**不读剪贴板**（读会加锁，反而更容易失败）；
2. 每次重试前 `Clipboard.Clear()`，等待递增，最多 3 次（`CopyPasteSlideWithRetry`）；
3. 仍失败就落到 `Slides.InsertFromFile` 兜底（索引见第 10 条）——代价是不搬版式绑定，
   所以它只是兜底，日志会标注 `(InsertFromFile fallback)`。

### 12. 端到端验证手法：用 CDP 驱动 WebView2 里的真实前端（推荐复用）

不用装测试框架，也不用模拟点击卡片，直接给 WebView2 开远程调试端口即可把
"C# 桥 + 前端 + PowerPoint 文档"整条链路跑一遍：

```powershell
# 1) 让 PowerPoint 继承这个环境变量后再启动（WebView2 原生支持）
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=9333"
Start-Process "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" -ArgumentList '"<某个测试文稿>"'
# 2) 用 UI Automation 点 Ribbon「素材库」→「打开素材库」（否则任务窗格不创建，也就没有 WebView2）
# 3) http://127.0.0.1:9333/json/list 拿到 page 的 webSocketDebuggerUrl，
#    Runtime.evaluate 里注入一个 helper 后即可调用任意桥消息：
#    __audit('insertPage', {id, position:'end'})  ->  拿到真实 result
```

要点/坑：
- Python `websocket-client` 连 CDP 必须 `suppress_origin=True`（Chromium 会 403 拒绝带 Origin 的连接）。
- 页面里 `PostWebMessageAsJson` 推来的 `e.data` **已经是对象**，不是字符串，解析要两种都兼容。
- `Runtime.evaluate` 里用 `awaitPromise` + `chrome.webview.addEventListener('message')` 按 `reqId` 配对。
- 顺便开 `Log.enable` / `Network.enable` 就能抓到 JS 未捕获异常与 404（例如缺 favicon）。
- 跑完记得**清理测试副作用**：收藏动作会真的往库里写条目、`TouchUsed` 会改 `useCount/lastUsedAt`、
  `reorder` 会重排全库 `sortOrder` —— 建议先备份 `%APPDATA%\PPTShapeLibrary\items`（逐字节还原 meta.json）。



### 13. 判决性实验：整页插入能否彻底不用剪贴板？（2026-10，结论：可以）

**背景**：整页插入原先"剪贴板为主、InsertFromFile 兜底"，实测在用户机上主路径要重试 3 次才失败，
整页插入耗时 **6.0~6.5 秒**，且失败后兜底又因索引写错（第 10 条）而必然报错。

**实验方法**（脚本化，可复跑）：
1. 目标文稿用 `Presentations.Add(0)` 造"默认 Office 主题"的干净文稿（正是第 1 条的名字冲突场景）；
2. 同一份库文件分别走两条路径插到目标末尾：A=`lib.Slides(1).Copy()` → `target.Slides.Paste()`；
   B=`target.Slides.InsertFromFile(lib, Count, 1, 1)`；
3. 两份产物都另存为 pptx，解包逐项比对：插入页 `slideN.xml`、`slideN.xml.rels` 的版式指向、
   被绑定版式的内容、`theme1.xml` 的 `clrScheme`、设计/母版名、`ppt/media/*`；
4. 再把两张插入页导出同尺寸 PNG（960×540）做像素差（Pillow）。

**结果（3 个不同整页快照，全部一致）**：

| 比对项 | 结果 |
|---|---|
| 版式绑定（`slideLayout` 指向） | **相同**（如 `slideLayout7.xml` / `slideLayout1.xml` / `slideLayout12.xml`） |
| 主题 `clrScheme` | 相同 |
| 设计/母版/版式名集合 | 相同（都是 `Lib<id>-M1` 等唯一名） |
| 媒体文件 | 数量与文件名相同（各 3 个） |
| 导出 PNG 像素平均差 | **0.0000 / 255，差异区域 = None（渲染完全一致）** |
| `slideN.xml` | 标签序列相似度 0.9994~0.9996；**唯一差异是 `<p14:creationId val="…">`**（PowerPoint 每次插入随机生成的 id，无语义/视觉影响） |

**结论与收益**：
- 整页插入删除剪贴板路径，改 `InsertFromFile` 单通道 →
  实测耗时 **6.5s → 0.06s（约 100×）**，且**用户剪贴板完全不被触碰**（用剪贴板哨兵文本验证）；
- 连带消失：库文稿句柄池在整页插入上的使用、剪贴板快照/还原、粘贴重试的 `Thread.Sleep`、
  `CollectClipboard` 误收残留数据的场景；
- **前提条件必须保持**：库文件设计唯一改名。测试若发现"插入后多出边框/颜色漂移"，
  先检查 `ppt/slideMasters/slideMaster1.xml` 的 `name` 是否带 `Lib<id>-` 前缀。

**给未来的通用教训**：把"某 API 不可靠"写进文档时，一定要写清**前提条件**。
第 2 条旧结论在"设计名不唯一"的时代是对的，但设计唯一改名之后它就成了误导——
本次差点因此保留一条又慢又不稳的剪贴板主路径。
