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

### 9.9 ⚠️ 两个"更新装不上"的真凶（2026-10 用户实测）

用户反馈：**关掉 PPT 后仍会弹更新；点更新后安装器说"请关闭 PPT"，可任务管理器里根本没有 PPT**。
排查后是**两个独立问题**，都已修。

#### ① 僵尸 PowerPoint 进程

用户关闭 PPT 窗口后进程没退干净：`MainWindowHandle == 0`（无任何窗口）、却有 89 个线程常驻、
存活 3 小时以上。而原判据是 `Process.GetProcessesByName("POWERPNT").Length > 0` —— 只看进程名：

- 插件的定时提醒照常弹（插件就在这个僵尸进程里跑着）
- 安装器 `--pending-update` 一直等它退出 → 等满 180 秒超时 → 用户只能手动关掉安装器

**修法**：判定"PPT 在用"必须**有主窗口**（`MainWindowHandle != 0`）。
无窗口的孤儿进程不锁文件、不拦安装，只在日志里提示（也不去杀它 —— 可能带着用户未保存的内容）。

#### ② ClickOnce 部署状态损坏 → VSTOInstaller 永远 -400

```
System.Deployment.Application.DeploymentException: 存储元数据"UpdateSkipTime"无效。
   在 Microsoft.VisualStudio.Tools.Applications.Deployment.ClickOnceAddInDeploymentManager.GetManifests(TimeSpan)
   在 ...ClickOnceAddInDeploymentManager.InstallAddIn()
```

（更早一次报的是 `ArgumentOutOfRangeException: entryValue`，来自 `MetadataStore.UpdateLastCheckedTime`。）

损坏点：`HKCU\Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0\PackageMetadata\{容器GUID}\ppts..vsto_<公钥token>_<散列>`
里的 `{订阅GUID}!UpdateSkipTime` 成了非法的 `0001/01/01`。

**只要这个状态坏了，任何位置都装不进去**（换目录、删 Addins 键、删 VSTA 记录全都无效），
而且**卸载也会卡死**（`/Uninstall` 卡满 120 秒超时）。

**修法**（已内置为安装器 `ClearBrokenClickOnceState`）：把 `SideBySide\2.0` 下
`PackageMetadata / Installations / Components / Marks / StateManager / Visibility /
VisibilityRoots / Categories / Assemblies` 里**只属于本产品**的条目删掉
（识别依据：公钥 token `a8a7103479c94f3f`、`ppts..` 前缀、或含产品名），
外加删掉 `%LOCALAPPDATA%\Apps\2.0\*\manifests` 里含本产品清单的缓存目录。
**只动本产品，不碰别的 ClickOnce 应用**（本机还装着"方方格子"FFCell）。

调用时机：
- `--pending-update`：装之前先清；**若清到了东西就跳过卸载直接重装**（坏状态下 `/Uninstall` 必卡 120 秒）
- `--pending-update`：装失败（非 0 退出码）时**再清一次并重试**一遍
- 界面「开始安装」与「修复加载项注册」：同样先清一次（正常机器上是空操作）

实测（同一台机器、故意注入坏值复现）：
```
21:10:31  检测到 ClickOnce 状态损坏，已清理 9 条 → 跳过卸载，直接重装
21:10:33  VSTOInstaller 退出码：0
21:10:33  ✓ 安装完成
```
耗时 **2.1 秒**（修复前：卡 120 秒卸载超时 + 安装 -400 失败）。

### 9.10 分发方式改版：放弃 ClickOnce，改用 `|vstolocal` 本地安装（2026-10）

**背景**：§9.8、§9.9 记的那些坑（`AddInAlreadyInstalledException`、"已安装另一版本"、
卸载卡 120 秒、`UpdateSkipTime` 元数据损坏导致 `VSTOInstaller` 永远 `-400`、
升级必须关闭 PowerPoint）**根源都在 ClickOnce**。既然每次都要跟它斗争，索性不用它。

**决定性实验**：`Manifest = file:///…\PPTShapeLibrary.vsto|vstolocal`（VS 按 F5 用的那种）
在**完全没有 VSTO 信任项（Inclusion）**的情况下能否加载？

结果：**能**。装到 `%LOCALAPPDATA%\PPTShapeLibrary\addin\<版本>\`，清空所有 Inclusion 项后
重启 PowerPoint：插件正常加载、`LoadBehavior` 保持 3、无 VSTO 错误事件。

**关键事实**：
- 部署清单（`.vsto`）与 `PPTShapeLibrary.dll.manifest` 里引用的都是**真实文件名**，
  `.deploy` 后缀只是 ClickOnce 下载时的约定 → `|vstolocal` 直接用真实名即可。
- 扁平布局（`codebase="PPTShapeLibrary.dll.manifest"`）即可，不需要 ClickOnce 那套
  `Application Files\PPTShapeLibrary_x_y_z_w\` 嵌套。
- 插件载荷只有 **20 个文件 / 2.45 MB**（Release 构建输出），加上 VSTO 离线包约 4 MB 压成单 exe。

**新旧对比**：

| | ClickOnce | `|vstolocal` 本地安装 |
|---|---|---|
| 安装 | `VSTOInstaller /Install` | 复制文件 + 写 2 个注册表键（**亚秒级**） |
| 升级 | 必须关 PowerPoint，`/Uninstall` + `/Install`（还可能卡 120 秒） | 装到**新版本目录** + 改注册表指向；**完全不用关 PowerPoint**（重启 PPT 生效） |
| 卸载 | ClickOnce 托管 | 删键 + 删目录（自身在待删目录里 → 派生 `cmd` 延迟删除） |
| 典型故障 | 已安装另一版本 / -400 / 卸载卡死 | 无（没有那套部署状态机） |

**实测**（同一台机器）：
```
装机耗时       < 1 秒
开着 PowerPoint 升级到新版本   0.4 秒，退出码 0；注册表已指向新目录；
                              PowerPoint 仍在运行、「素材库」仍可用；旧目录被占用→留待下次清理
卸载          注册/ARP/目录全清，素材库数据（92 条目）完好
```

**踩坑**：
1. `.vsto` 的版本号来自 MSBuild 的 **`ApplicationVersion`**（不是 `AssemblyInfo`），
   而且**增量构建不会重新生成已存在的 `.vsto`** → 打安装包前必须先删旧 `.vsto`，
   再用 `/p:ApplicationVersion=<版本>` 构建，否则打进包里的清单是旧版本号。
2. 升级时要装到**按版本分开的目录**（`addin\1.2.1.0\`）。若装到同一目录，
   运行中的 PowerPoint 锁着 DLL/清单，`File.Copy` 会失败 → 装到新目录就完全不冲突。
3. 卸载器自己住在待删目录里，直接 `Directory.Delete` 必然被拒 →
   派生 `cmd /c ping -n 3 127.0.0.1 >nul & rmdir /s /q "<dir>"` 等进程退出后再删。
4. `dist\` 同时放 setup exe 时，`Compress-Archive dist\*` 会把 4 MB 的 exe 也压进 zip
   （实测 4 MB → 16 MB）→ 打 zip 时要按文件名排除 setup/cer/sha256。
5. **`package-release.ps1` / `dev-register.ps1` 原来用「Manifest 含 `|vstolocal`」判定开发式注册并删除**。
   架构改版后正式安装也是 `|vstolocal` → 这条老判据会把用户的**正式注册**删掉，
   表现为"打完版 PowerPoint 里素材库消失"。已改为**按路径判定**（落在仓库 `bin\Debug|bin\Release`
   下才是开发式）。连带三个实测坑：
   - 注册表里的 `Manifest` 是 **URL 格式、全正斜杠**，判据要写 `/bin/debug/`
   - PowerShell `-like` 的通配符里 **`\` 是转义符**，`*\bin\debug\*` 实际匹配 `*bindebug*`，永远不中
     → 用 `.Contains()` 做字面量判断
   - 不能要求必须有 `|vstolocal`：误双击构建产物装成的 ClickOnce 注册（不带该后缀）同样要清

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

### 14. 判决性实验：形状插入能否彻底不用剪贴板？（2026-10，结论：**不能**，但可等待恢复）

**背景**：用户在真实环境报「剪贴板被占用」，冲突对象是 **Snipaste** 这类剪贴板工具。
报错原文（已复现）：

```
Shapes.Paste : Invalid request. Clipboard is empty or contains data which may not be pasted here.
```

**实验脚本**：`scripts\test\exp-insert-noclip.ps1`（可复跑；用 `hold-clipboard.ps1` 模拟占用）。
占用器用 Win32 `OpenClipboard()` 锁住剪贴板（实测：占用期间 `Clipboard.SetText` 抛"所请求的剪贴板操作失败"，释放后立刻成功 → 忠实复现 Snipaste 场景）。

**五条路径在"剪贴板被占用"时的实测结果**：

| 路径 | 做法 | 结果 |
|---|---|---|
| **A**（现状形状插入） | 跨文稿 `Shapes.Range().Copy()` → `target.Paste()` | ❌ **失败**（复现用户报错） |
| **B**（候选：插整页再搬） | `InsertFromFile` 插整页 → **同文稿内** `Copy`→`Paste` 搬形状 | ❌ **失败**（第一步成功且没碰剪贴板，第二步仍失败） |
| **C**（现有整页插入） | `Slides.InsertFromFile` | ✅ 成功 |
| **D**（整页 Duplicate） | `InsertFromFile` + `Slide.Duplicate()` − 删临时页 | ✅ 成功（`Duplicate` 不碰剪贴板） |
| **E**（重试策略） | `Clear` + 等待 + 重试，共 6 次 | ✅ **最终成功**（等到占用者释放的间隙，实测约 3s） |

**三条结论**：

1. **形状插入无法免剪贴板**。
   `InsertFromFile` 只能插**整页**，不能只插页里的几个形状；
   而任何"把形状从 A 放到 B"都要 `Copy`/`Paste`——
   **B 路径证明了：哪怕在**同一文稿内**，`Copy`→`Paste` 仍然依赖剪贴板**（这是本实验最有价值的否定性结果，
   它否掉了"先插整页再同文稿内搬运"这个看起来很聪明的绕行方案）。

2. **`Slide.Duplicate()` 是唯一不碰剪贴板的复制手段**，但它复制的是**整页**，
   对"把形状插到当前页"这个语义没有帮助（只能用于整页类操作，如 D）。

3. **剪贴板冲突是"可等待、可恢复"的，不是死局**。
   Snipaste 这类工具会周期性开合剪贴板（实测模拟占用 3s / 释放 1s）；
   重试策略能在释放的间隙成功（实测第 2 次尝试即恢复，总耗时约 3s）。
   所以正确方向是**优化等待/重试**，而不是消灭剪贴板。

**对现有代码的指引**：
- `ShapeExporter.CopyPasteWithRetry` 的方向是对的（清剪贴板 → 重试 → 快照还原），**保留**。
- 注意一个细节：`Clipboard.Clear()` 在剪贴板被占用时**自己也会抛异常**
  （"Requested Clipboard operation did not succeed."）。现有代码用 `try/catch` 吞掉是对的，
  但这意味着**第一次尝试很可能白费**，不要把 `Clear()` 的成功与否当作"剪贴板可用"的判据。
- 想进一步提升成功率：把固定 `Sleep` 改成**指数退避 + 更长上限**（Snipaste 占用窗口可能比现在的
  `40+30*attempt` ms 长得多；实测等约 3s 才恢复）。
- **不要**为了绕开剪贴板去改 `BackupTo`/收藏格式（形状库必须存 pptx 才能保真，见 §2/§13）。

### 15. 应用图标 + Ribbon 图标风格统一（2026-10）

#### 15.1 应用图标：`assets\app.ico`

之前 `<ApplicationIcon>` 一直是**空节点** → exe 用默认 .NET 图标。现在：

- 图标源：用户自制图（**实际是 WebP**，扩展名却写成了 `.jpeg`）
- 生成：`scripts\make-app-icon.py`（Pillow），产出 10 档尺寸的 `assets\app.ico`
- 接入：`installer\PPTShapeLibraryInstaller.csproj` 与 `PPTShapeLibrary.csproj` 的 `<ApplicationIcon>`

**三个踩过的坑**：
1. **GDI+ 读不了 WebP**：`Image.FromFile` 对 WebP 报 `Out of memory`（这个错误信息完全不提格式问题，
   极易误判成"文件损坏/太大"）。判断真实格式要看**文件头 magic bytes**：
   `52 49 46 46 .... 57 45 42 50` = `RIFF....WEBP`。解码用 **Pillow**（本机有）。
2. **要裁掉多余透明留白再统一边距**：原图 2048×2048，内容仅 1528×1472（四周留白 12.5%~14%）。
   不裁的话图标在实际显示尺寸里会**比别的图标小一圈**（留白也算进图标画布）。
   做法：按 alpha `getbbox()` 裁到内容 → 按 5% 边距重新居中成正方形 → 再逐尺寸 LANCZOS 缩放。
3. **别对已有有损图重新编码**：曾把源图另存为 lossless WebP，体积 244KB → 755KB。
   源图原字节复制即可。

#### 15.2 Ribbon 图标风格：换成 **Material Symbols Outlined**（线性）

**问题**：任务窗格 Web UI 用 **lucide**（线性描边），功能区却用 **Material Icons**（填充实心）+ 彩色，
两处风格割裂（用户反馈"侧边栏的更好看、和顶部不一样"）。

**做法**：功能区图标字体改为 **Material Symbols Outlined**，与 lucide 的线性语言一致。

**关键：完整字体 10.4MB 不能发货，用 fontTools 两步瘦身**：

```bash
# 1) 固化变量轴（Material Symbols 是可变字体：FILL/GRAD/opsz/wght；图标只用默认外观）
python -m fontTools.varLib.instancer MaterialSymbolsOutlined.ttf FILL=0.0 GRAD=0.0 opsz=24.0 wght=400.0 -o instanced.ttf
#    10.16MB → 1.32MB
# 2) 子集化：只保留用到的 23 个字形
python -m fontTools.subset instanced.ttf --unicodes=U+E262,U+E167,... \
       --output-file=MaterialSymbolsOutlined-subset.ttf --no-hinting --desubroutinize --layout-features=
#    1.32MB → 5.0KB
```

产物：`PPTShapeLibrary\RibbonFonts\MaterialSymbolsOutlined-subset.ttf`（5KB）+ 同名 `.codepoints`（23 行）。
MaterialIcons（348KB）降为**回退**（`RibbonIconFactory.EnsureLoaded` 里顺序已对调）。

**⚠️ 维护要点**：`RibbonIconFactory` 的图标名常量**新增字形时**，必须把新码点补进
subset 的 `.codepoints` 并**重新子集化**，否则该图标在 Outlined 里查不到会
**静默回退成填充风格** → 功能区出现风格不一致的图标（不报错，很难发现）。

**字号比例不用改**：绘制用的 `ss*0.92` 原本按 Material Icons 调的，实测 Outlined 的字形视觉大小
平均只差 1.2%（73.6% vs 74.5%），沿用即可。

**验证方式**（这类"看不出对错"的改动，靠程序判定）：
- 逐个字形渲染后统计 **ink 像素数**（>阈值即非空白），23/23 通过；
- 两个字体同参数渲染后比 **ink 包围盒占画布比例**，用于校准字号；
- exe/DLL 图标用 `Icon.ExtractAssociatedIcon` 提取后**采样中心像素 + 统计主色占比**
  （本例中心 `#2F9FFF`、蓝色占 37% → 确认嵌入的是自制图标而非默认图标）。
- ⚠️ 注意 `AppDomain.CurrentDomain.BaseDirectory`：在 PowerShell/独立测试器里加载 DLL 时它指向
  **宿主进程目录**，找不到 `RibbonFonts\`，会让图标工厂全部返回 null —— 这是**测试环境问题，不是代码问题**
  （日志里会出现 `Ribbon fonts missing: ... C:\WINDOWS\System32\WindowsPowerShell\v1.0\\RibbonFonts`）。
