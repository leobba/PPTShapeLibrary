# 一键打包正式版：发布 ClickOnce → 编译安装器 → 组装 dist 分发包
# 用法（PowerShell）：  cd E:\code\PPTShapeLibrary ;  .\scripts\package-release.ps1
param(
    [string]$Version = "1.0.1.0",
    [string]$Root = "E:\code\PPTShapeLibrary"
)

$ErrorActionPreference = "Stop"
$msb = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
$bootPath = "C:\Program Files (x86)\Microsoft SDKs\ClickOnce Bootstrapper\"

Write-Host "== 1/4 发布 ClickOnce（插件本体）==" -ForegroundColor Cyan

# 把 AssemblyInfo 的版本对齐到发版号。
# 为什么要做：插件内"检查更新"要拿当前版本和 GitHub 最新版比较，而
#   · ARP DisplayVersion / vsto 清单版本 都来自 AssemblyVersion（默认停在 1.0.1.0 不再动）
#   · ClickOnce 的 ApplicationVersion（如 1.0.5.0）**不会**反映到上面两处
# 于是插件会读到旧的 AssemblyVersion → 误报"有新版本"或显示错误的当前版本。
# 同时 dev-register.ps1 也读 DLL 的 FileVersion 写进注册表 ProductVersion，保持一致。
$asmInfo = "$Root\PPTShapeLibrary\Properties\AssemblyInfo.cs"
if (Test-Path $asmInfo) {
    $changed = $false
    $aiLines = Get-Content $asmInfo -Encoding UTF8
    for ($i = 0; $i -lt $aiLines.Count; $i++) {
        if ($aiLines[$i] -match '^\[assembly:\s*AssemblyVersion\("([\d\.]+)"\)\]') {
            if ($Matches[1] -ne $Version) { $aiLines[$i] = '[assembly: AssemblyVersion("' + $Version + '")]'; $changed = $true }
        }
        elseif ($aiLines[$i] -match '^\[assembly:\s*AssemblyFileVersion\("([\d\.]+)"\)\]') {
            if ($Matches[1] -ne $Version) { $aiLines[$i] = '[assembly: AssemblyFileVersion("' + $Version + '")]'; $changed = $true }
        }
    }
    if ($changed) {
        Set-Content $asmInfo -Value $aiLines -Encoding UTF8
        Write-Host "   AssemblyInfo 版本 → $Version（否则插件内读到的是旧版本，会误报有新版本）" -ForegroundColor Yellow
    } else {
        Write-Host "   AssemblyInfo 版本已是 $Version"
    }
}

& $msb "$Root\PPTShapeLibrary\PPTShapeLibrary.csproj" /t:Publish /p:Configuration=Release `
    /p:ApplicationVersion=$Version `
    "/p:PublishUrl=$Root\publish" "/p:InstallUrl=file:///$($Root.Replace('\','/'))/publish/" `
    /p:IsWebBootstrapper=true "/p:GenerateBootstrapperSdkPath=$bootPath" `
    /nologo /v:m
if ($LASTEXITCODE -ne 0) { throw "ClickOnce 发布失败" }

$appPublish = "$Root\PPTShapeLibrary\bin\Release\app.publish"
$publish = "$Root\publish"
# 先确认源存在再删旧产物：MSBuild 的 Clean/Rebuild 会删掉 app.publish，
# 旧写法会先删掉 publish\ 再在 Copy-Item 处报错，等于毁掉上一份可用产物。
if (-not (Test-Path $appPublish)) { throw "未找到 $appPublish（发布产物缺失，检查 /t:Publish 是否真的执行成功）" }
if (Test-Path $publish) { Remove-Item $publish -Recurse -Force }
Copy-Item $appPublish $publish -Recurse -Force
Write-Host "   产物 → $publish"

# 兜底自愈：构建期有可能把加载项注册成 <repo>\PPTShapeLibrary\bin\Release\...vsto|vstolocal
# （开发式注册），它会与正式安装抢注册表，导致再次安装时报"已安装另一版本"。
#
# ⚠️ **不能只看 "|vstolocal" 就删**！2026-10 起正式版也走 |vstolocal
#   （安装到 %LOCALAPPDATA%\PPTShapeLibrary\addin\<版本>\，见 installer/LocalInstall.cs）。
#   用"清单路径是否在仓库 bin\Debug|bin\Release 下"判定开发式，才不会把用户的正式安装删掉。
#   实测踩过：这条老判据在架构改版后把本机的**正式注册**清掉了（PowerPoint 里「素材库」消失）。
#
# 另：加载项注册可能同时写在"版本无关"与"16.0 版本化"两处，两处都要检查。
function Test-IsDevLocalRegistration([string]$manifest) {
    if ([string]::IsNullOrEmpty($manifest)) { return $false }
    # ⚠️ 三个坑都踩过：
    #  1. Manifest 是 URL 格式、**全正斜杠**（file:///E:/.../bin/Debug/...），判据要写 /bin/debug/。
    #  2. PowerShell 的 -like wildcard 里 `\` 是转义符，`*\bin\debug\*` 匹配不到东西。用 .Contains()。
    #  3. **不能要求必须有 "|vstolocal"**：误双击构建产物 bin\Debug\*.vsto 会装成
    #     ClickOnce 注册（不带 vstolocal），同样会跟正式安装抢注册表 → 也要清。
    # 判据就一条：清单路径落在仓库的 bin\Debug|bin\Release 下，即"开发式"。
    $m = $manifest.ToLowerInvariant()
    return ($m.Contains("/bin/debug/") -or $m.Contains("/bin/release/") -or
            $m.Contains("\bin\debug\") -or $m.Contains("\bin\release\"))
}

function Clear-DevLocalAddinRegistration([string]$label) {
    $keys = @(
        "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary",
        "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary"
    )
    foreach ($k in $keys) {
        if (-not (Test-Path $k)) { continue }
        $manifest = (Get-ItemProperty $k -ErrorAction SilentlyContinue).Manifest
        if (Test-IsDevLocalRegistration $manifest) {
            Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "   $label 已清理开发式注册（指向仓库构建输出）：$k" -ForegroundColor Yellow
        } else {
            Write-Host "   $label 保留现有注册（正式安装）：$k" -ForegroundColor Yellow
        }
    }
}
Clear-DevLocalAddinRegistration "1/4"

Write-Host "== 2/4 编译安装器（环境检测）==" -ForegroundColor Cyan
& $msb "$Root\installer\PPTShapeLibraryInstaller.csproj" /t:Rebuild /p:Configuration=Release /nologo /v:m
if ($LASTEXITCODE -ne 0) { throw "安装器编译失败" }

Write-Host "== 3/4 组装 dist 分发包 ==" -ForegroundColor Cyan
$dist = "$Root\dist"
# 只清"本脚本生成的内容"，保留用户自己放进来的 离线安装包\（文档要求用户预置 100~250MB
# 官方安装包，整目录删除会让用户每发一版都要重新下载）。
if (Test-Path $dist) {
    foreach ($name in @("App", "安装 PPTShapeLibrary.exe", "安装说明.txt")) {
        $p = Join-Path $dist $name
        if (Test-Path $p) { Remove-Item $p -Recurse -Force }
    }
}
New-Item -ItemType Directory -Path $dist -Force | Out-Null
New-Item -ItemType Directory -Path "$dist\App" | Out-Null
New-Item -ItemType Directory -Path "$dist\离线安装包" -Force | Out-Null
Copy-Item "$Root\installer\bin\Release\安装 PPTShapeLibrary.exe" $dist -Force
# 只放"部署清单 + 应用文件"：.vsto 的 codebase 是相对路径，安装源即本目录 → 拷到任何机器都能装。
# 不放 setup.exe（其中烤入了发布机上的绝对 URL，换机/换目录后不可用）。
Copy-Item "$publish\PPTShapeLibrary.vsto" "$dist\App" -Force
Copy-Item "$publish\Application Files" "$dist\App" -Recurse -Force

# 内置依赖包：VSTO Runtime 只有约 3MB（比 .NET 4.8 的 110MB 小两个数量级），
# 直接随分发包带上，用户装的时候点「安装本地包」即可，**完全不依赖网络**。
# 微软官方源在国内基本打不开，这一步是国内用户能不能装上的关键。
$offline = "$dist\离线安装包"
New-Item -ItemType Directory -Path $offline -Force | Out-Null
$hasVsto = @(Get-ChildItem $offline -File -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -like 'vstor*' -and $_.Name -match '\.(exe|zip)$' }).Count -gt 0
if (-not $hasVsto) {
    Write-Host "   离线包里没有 VSTO Runtime，自动取一份（走 GitHub 固定 tag + 加速站）…" -ForegroundColor Yellow
    & pwsh -NoProfile -File (Join-Path $PSScriptRoot "fetch-prerequisites.ps1") -OutDir $offline -Only vsto
    if ($LASTEXITCODE -ne 0) {
        Write-Host "   取包失败（不影响本次打包，但 dist 将不具备离线装 VSTO 的能力）" -ForegroundColor Yellow
    }
} else {
    Write-Host "   已内置 VSTO Runtime 离线包：" -ForegroundColor Green
    Get-ChildItem $offline -File | Where-Object { $_.Name -like 'vstor*' } |
        ForEach-Object { "     $($_.Name)  $([math]::Round($_.Length/1MB,2)) MB" }
}

# 拆掉"陷阱"：构建产物 bin\Release\PPTShapeLibrary.vsto 一旦被双击，会把加载项注册成
# …\bin\Release\…|vstolocal（开发式），与正式安装抢注册表并导致"已安装另一版本"报错。
# 删除前必须先把指向它的开发式注册清掉（两处键），否则注册表会指向一个已不存在的清单文件。
$trapVsto = "$Root\PPTShapeLibrary\bin\Release\PPTShapeLibrary.vsto"
if (Test-Path $trapVsto) {
    Clear-DevLocalAddinRegistration "3/4"
    Remove-Item $trapVsto -Force -ErrorAction SilentlyContinue
    Write-Host "   已移除构建产物 bin\Release\PPTShapeLibrary.vsto（避免误双击造成注册冲突）" -ForegroundColor Yellow
}

Write-Host "== 4/4 生成说明文件 ==" -ForegroundColor Cyan
$readme = @"
PPTShapeLibrary 安装说明（版本 $Version）
================================================

【怎么装】
1. 关闭 PowerPoint；
2. 双击「安装 PPTShapeLibrary.exe」；
3. 安装器会先检测运行环境（PowerPoint / .NET Framework 4.8 / VSTO Runtime / WebView2）：
   - 全部满足 → 点「开始安装」即可；
   - 缺少组件 → 点该项右侧「下载安装」按提示安装，装完回到窗口点「重新检测」；
   - 若「离线安装包」文件夹里放了官方安装包，按钮会变成「安装本地包」，可直接静默安装（需同意 UAC）。
4. 安装完成后打开 PowerPoint → 功能区出现「素材库」选项卡即成功。

【目录说明】
  安装 PPTShapeLibrary.exe   环境检测 + 安装引导程序（双击这个）
  App\                       插件本体：PPTShapeLibrary.vsto + Application Files\（请勿单独删改/移动）
  离线安装包\                可选：放入官方离线安装包后可离线安装
                             · ndp48-x86-x64-allos-enu.exe        (.NET Framework 4.8)
                             · vstor_redist.exe                   (VSTO Runtime 4.0)
                             · MicrosoftEdgeWebview2Setup.exe     (WebView2 Runtime)

【卸载 / 重装】
  控制面板 → 程序和功能 → 卸载「PPTShapeLibrary」；
  或再次运行本安装器，检测到已安装版本时点「卸载旧版本」。
  数据（素材库/设置/备份/日志）位于 %APPDATA%\PPTShapeLibrary，卸载不会删除。

【功能区没有「素材库」怎么办】
  1. 先确认已装 VSTO Runtime，再到 文件 → 选项 → 加载项 → 管理「COM 加载项」勾选 PPTShapeLibrary；
  2. 仍没有 → 打开本安装器点「修复加载项注册」（会重写当前用户的加载项注册）；
     也可以选择同时写入"所有用户"（会弹一次 UAC，写入 HKLM，更不容易被清理）；
  3. 重启 PowerPoint。若还不行，见 docs\PowerPoint-COM经验与坑.md 的『插件在 PPT 里消失』一节。

【常见问题】
  · 装了但 PowerPoint 里没有「素材库」：见上一节；另确认已装 VSTO Runtime 并重启 PowerPoint。
  · 面板空白/提示缺少 WebView2：安装 Microsoft Edge WebView2 Runtime。
  · 安装时提示"未知发布者"：当前使用临时证书签名，属正常，继续安装即可。
  · 报错「已安装另一版本的自定义项，不能从该位置升级」：本机已装过一份且这次换了位置。
    请在安装器里点「卸载旧版本」后重试（或控制面板卸载后重装）。
  · 请勿双击源码目录里的 bin\Release\PPTShapeLibrary.vsto —— 那是构建产物，
    双击会把它注册成开发式加载项、与正式安装冲突。安装只用本目录的安装器。
  · 开发者在 Release 构建后如果发现加载项"消失"：已在 csproj 里修掉（Release 构建不再注销加载项注册）；
    详见 docs\PowerPoint-COM经验与坑.md §9。
"@
[System.IO.File]::WriteAllText("$dist\安装说明.txt", $readme, (New-Object System.Text.UTF8Encoding($true)))

$offlineNote = @"
把官方离线安装包放到本文件夹，安装器会自动识别并显示「安装本地包」按钮：

  ndp48-x86-x64-allos-enu.exe      .NET Framework 4.8
      官方页： https://dotnet.microsoft.com/download/dotnet-framework/net48

  vstor_redist.exe                 VSTO Runtime 4.0（Office 加载项运行时）
      官方页： https://learn.microsoft.com/visualstudio/vsto/how-to-install-the-visual-studio-tools-for-office-runtime-redistributable

  MicrosoftEdgeWebview2Setup.exe   WebView2 Runtime
      官方页： https://developer.microsoft.com/microsoft-edge/webview2/#download

说明：
  · 微软下载中心的直链（fwlink / details.aspx）会不定期调整或下线，若链接失效请直接在上述官方页面里下载；
  · 只要把文件放进本文件夹，安装器就完全不依赖网络（推荐给内网/离线环境）。
"@
[System.IO.File]::WriteAllText("$dist\离线安装包\说明.txt", $offlineNote, (New-Object System.Text.UTF8Encoding($true)))

$size = [math]::Round(((Get-ChildItem $dist -Recurse -File | Measure-Object Length -Sum).Sum / 1MB), 2)
Write-Host "完成：$dist（$size MB）" -ForegroundColor Green
Get-ChildItem $dist | Select-Object Name
