# 开发版加载项 注册 / 注销（无需启动 Visual Studio 的 F5）
#
# 用法：
#   .\scripts\dev-register.ps1                       # 注册 Debug 构建（bin\Debug）到 PowerPoint
#   .\scripts\dev-register.ps1 -Configuration Release # 注册 Release 构建
#   .\scripts\dev-register.ps1 -Unregister           # 注销开发版加载项
#
# 说明：
#  - 这正是 Visual Studio F5 所做的事（写入 HKCU 的 Addins 键，Manifest 带 |vstolocal 后缀）。
#  - 直接 msbuild 构建【不会】写注册项，所以卸载正式版后 PPT 里会"找不到素材库"，用本脚本补上。
#  - 注册后需【完全退出并重新打开 PowerPoint】才会加载。
#  - 与"安装器/ClickOnce 正式版"互斥（同一个注册表键），装正式版前先 -Unregister。

param(
    [switch]$Unregister,
    [string]$Configuration = "Debug",
    [string]$OfficeVersion = "16.0",
    [string]$Root = "E:\code\PPTShapeLibrary"
)

$key = "HKCU:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\Addins\PPTShapeLibrary"

if ($Unregister) {
    if (Test-Path $key) {
        Remove-Item $key -Recurse -Force
        Write-Host "已注销开发版加载项（$key）" -ForegroundColor Yellow
    } else {
        Write-Host "当前未注册开发版加载项" -ForegroundColor Yellow
    }
    return
}

$vsto = Join-Path $Root "PPTShapeLibrary\bin\$Configuration\PPTShapeLibrary.vsto"
if (-not (Test-Path $vsto)) {
    throw "找不到 $vsto —— 请先用 $Configuration 配置构建一次（或在 VS 中生成）"
}
$dll = Join-Path $Root "PPTShapeLibrary\bin\$Configuration\PPTShapeLibrary.dll"
if (-not (Test-Path $dll)) { throw "找不到 $dll —— 请先构建" }

if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
Set-ItemProperty -Path $key -Name "Description"  -Value "PPTShapeLibrary"
Set-ItemProperty -Path $key -Name "FriendlyName" -Value "PPTShapeLibrary"
Set-ItemProperty -Path $key -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $key -Name "Manifest"     -Value ("file:///" + ($vsto -replace '\\','/') + "|vstolocal")

Write-Host "已注册开发版加载项：$vsto" -ForegroundColor Green
Write-Host "请完全退出并重新打开 PowerPoint —— 功能区应出现「素材库」选项卡。" -ForegroundColor Green
Write-Host "若没有：文件 → 选项 → 加载项 → 管理「COM 加载项」→ 勾选 PPTShapeLibrary。" -ForegroundColor Gray
