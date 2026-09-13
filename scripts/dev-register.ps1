# 开发版加载项 注册 / 注销（无需启动 Visual Studio 的 F5）
#
# 用法：
#   pwsh -File scripts\dev-register.ps1                     # 注册 Debug 构建（bin\Debug）
#   pwsh -File scripts\dev-register.ps1 -Configuration Release
#   pwsh -File scripts\dev-register.ps1 -Unregister         # 注销开发版加载项
#
# 【重要】为什么写两个位置：
#   HKCU\Software\Microsoft\Office\PowerPoint\Addins\<名称>        ← "版本无关"路径（Office 文档记载的每用户
#                                                                    COM 加载项目录；系统自带 OneNote 加载项在此）
#   HKCU\Software\Microsoft\Office\16.0\PowerPoint\Addins\<名称>    ← 版本化路径（VSTO/ClickOnce 常用）
#   实测（2026-09-13）：只写版本化路径时，PowerPoint 报
#       "Application.COMAddIns : Invalid request.  Office AddInsX is not initialized."
#   结果【所有加载项都不加载】（连系统自带的也不加载），功能区自然没有「素材库」；
#   同时 Word 的加载项子系统完全正常 → 问题在 PowerPoint 侧。
#   把注册项补写到"版本无关"路径后，加载项子系统立刻恢复（读 COMAddIns 由报错变为 Count=1），
#   插件正常加载（日志出现 Startup begin）。
#   → 因此两处都写；注销时两处一起删。
#
# 其它要点：
#   - 直接 msbuild 构建【不会】写注册项（注册由 VS 的 F5 或本脚本完成）。
#   - LoadBehavior 必须是 REG_DWORD（类型错了 PowerPoint 会忽略该加载项）。
#   - 注册后需要【完全退出并重新打开 PowerPoint】才会加载。
#   - 与"安装器/ClickOnce 正式版"共用同一个键，装正式版前先 -Unregister。

param(
    [switch]$Unregister,
    [string]$Configuration = "Debug",
    [string]$OfficeVersion = "16.0",
    [string]$Root = "E:\code\PPTShapeLibrary"
)

$keyVersioned   = "HKCU:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\Addins\PPTShapeLibrary"
$keyVersionless = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary"

if ($Unregister) {
    foreach ($k in @($keyVersioned, $keyVersionless)) {
        if (Test-Path $k) { Remove-Item $k -Recurse -Force; Write-Host "已注销：$k" -ForegroundColor Yellow }
    }
    Write-Host "开发版加载项已注销。" -ForegroundColor Yellow
    return
}

$vsto = Join-Path $Root "PPTShapeLibrary\bin\$Configuration\PPTShapeLibrary.vsto"
$dll  = Join-Path $Root "PPTShapeLibrary\bin\$Configuration\PPTShapeLibrary.dll"
if (-not (Test-Path $vsto)) { throw "找不到 $vsto —— 请先用 $Configuration 配置构建一次" }
if (-not (Test-Path $dll))  { throw "找不到 $dll —— 请先构建" }

$manifest = "file:///" + ($vsto -replace '\\','/') + "|vstolocal"

foreach ($k in @($keyVersioned, $keyVersionless)) {
    if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
    Set-ItemProperty -Path $k -Name "Description"  -Value "PPTShapeLibrary"
    Set-ItemProperty -Path $k -Name "FriendlyName" -Value "PPTShapeLibrary"
    Set-ItemProperty -Path $k -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $k -Name "Manifest"     -Value $manifest
    Write-Host "已写入：$k"
}

Write-Host ""
Write-Host "开发版加载项已注册：$vsto" -ForegroundColor Green
Write-Host "请【完全退出】并重新打开 PowerPoint —— 功能区应出现「素材库」选项卡。" -ForegroundColor Green
Write-Host "若仍没有：文件 → 选项 → 加载项，看 PPTShapeLibrary 是否在「非活动应用程序加载项」里。" -ForegroundColor Gray
Write-Host "若连加载项列表都是空的，见 docs\PowerPoint-COM经验与坑.md 的『插件在 PPT 里消失』一节。" -ForegroundColor Gray
