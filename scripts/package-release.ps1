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

# 兜底自愈：构建期有可能把加载项注册成 <bin\Release>\...vsto|vstolocal（开发式注册），
# 它会与用户通过安装器安装的正式版抢注册表，导致再次安装时报"已安装另一版本"。
# 若发现是 vstolocal 就清掉；是正式安装（ClickOnce）则保留不动。
# 注意：加载项注册可能同时写在"版本无关"与"16.0 版本化"两处（scripts\dev-register.ps1 两处都写），
# 只清一处会留下残留继续抢注册表 —— 两处都要检查。
function Clear-DevLocalAddinRegistration([string]$label) {
    $keys = @(
        "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary",
        "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary"
    )
    foreach ($k in $keys) {
        if (-not (Test-Path $k)) { continue }
        $manifest = (Get-ItemProperty $k -ErrorAction SilentlyContinue).Manifest
        if ($manifest -like "*vstolocal*") {
            Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "   $label 已清理 vstolocal 开发式注册：$k" -ForegroundColor Yellow
        } else {
            Write-Host "   $label 检测到正式版注册，保留不动：$k" -ForegroundColor Yellow
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
