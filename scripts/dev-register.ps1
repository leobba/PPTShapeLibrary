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
    [switch]$Force,                      # 允许 -Unregister 删除"正式版（ClickOnce）"注册
    [switch]$Machine,                    # 额外写入 HKLM（所有用户，需管理员）——见下方说明
    [string]$Configuration = "Debug",
    [string]$OfficeVersion = "16.0",
    [string]$Root = "E:\code\PPTShapeLibrary",
    # 手动指定"当前版本"（写进注册表 ProductVersion，插件内"检查更新"读它）。
    # 用途：发版后本地 DLL 版本会跟线上一样 → 看不到更新提示，没法测更新流程。
    # 这时用 -ProductVersion <旧版本> 注册，插件就会认为自己是旧版，从而看到新版本。
    # 不指定时按 DLL 的 FileVersion 写（正常开发用这个）。
    [string]$ProductVersion = ""
)

$ErrorActionPreference = "Stop"

$keyVersioned   = "HKCU:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\Addins\PPTShapeLibrary"
$keyVersionless = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary"
$machVersioned   = "HKLM:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\Addins\PPTShapeLibrary"
$machVersionless = "HKLM:\Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary"

# HKLM 说明：HKCU 的"版本无关"键在本机多次不明原因消失（功能区的「素材库」随之消失，
# 而插件此时根本没有代码在跑、无法自愈）。写一份到 HKLM（所有用户）可作为根治手段：
# 机器级注册不会被用户级加载项运行时改写。代价是需要管理员权限，且卸载要一起清。
function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ---------------------------------------------------------------------------
# 清理"会让 VSTO 拒绝安装/加载"的两类残留（实测踩过多次，都是它们造成
# VSTOInstaller 报 -400 / AddInAlreadyInstalledException：「已安装另一版本的自定义项」）
# ---------------------------------------------------------------------------

# ① VSTA\Solutions\<GUID>：这是 VSTO 判定"该产品已装过某版本"的**真正来源**。
#    实测把所有其它位置（Addins 键 / SolutionMetadata / ARP）都清了，只要它还在，
#    VSTOInstaller 依然拒绝安装 —— 必须删掉这条才能装上。
function Clear-VstaRecords {
    $sol = "HKCU:\Software\Microsoft\VSTA\Solutions"
    if (-not (Test-Path $sol)) { return }
    Get-ChildItem $sol -ErrorAction SilentlyContinue | ForEach-Object {
        $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        $hit = ($p.ProductName -like "*PPTShapeLibrary*") -or
               ($p.Url         -like "*PPTShapeLibrary*") -or
               ($p.ApplicationID -like "*PPTShapeLibrary*")
        if ($hit) {
            Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "已移除 VSTA 记录（『已安装另一版本』的判定来源）：$($_.PSChildName)" -ForegroundColor Yellow
        }
    }
}

# ② ARP 里指向**开发路径**的卸载项（bin\Debug / bin\Release / vstolocal）：
#    这种记录不是 VSTOInstaller 正规装的，用 /Uninstall 去卸会**挂死**（实测卡满 120s 超时）。
#    注意卸载串里是正斜杠（file:///E:/…/bin/Debug/…），必须先统一分隔符再匹配。
function Clear-DevArpRecords {
    $roots = @(
        @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Hive = "HKCU" },
        @{ Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Hive = "HKLM" }
    )
    foreach ($r in $roots) {
        if (-not (Test-Path $r.Path)) { continue }
        Get-ChildItem $r.Path -ErrorAction SilentlyContinue | ForEach-Object {
            $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($p.DisplayName -notlike "*PPTShapeLibrary*") { return }
            $norm = ("$($p.UninstallString)" -replace '/', '\')
            if ($norm -like "*bin\Debug*" -or $norm -like "*bin\Release*" -or $norm -like "*vstolocal*") {
                Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "已移除陈旧的开发版卸载项（$($r.Hive)）：$($_.PSChildName)" -ForegroundColor Yellow
            }
        }
    }
}

if ($Unregister) {
    foreach ($k in @($keyVersioned, $keyVersionless)) {
        if (-not (Test-Path $k)) { continue }
        # 安全闸：这个键与"安装器/ClickOnce 正式版"共用。无脑删除会把用户已装好的正式版
        # 注册表项删掉（ARP 记录还在）→ 插件从此不加载，只能重装。默认只清 vstolocal 开发式注册。
        $manifest = (Get-ItemProperty $k -ErrorAction SilentlyContinue).Manifest
        if ($manifest -notlike "*vstolocal*" -and -not $Force) {
            Write-Host "跳过（非开发式注册）：$k" -ForegroundColor Yellow
            Write-Host "  Manifest = $manifest" -ForegroundColor Gray
            Write-Host "  这看起来是正式版安装。确实要删除请加 -Force。" -ForegroundColor Gray
            continue
        }
        Remove-Item $k -Recurse -Force -ErrorAction Stop
        if (Test-Path $k) { throw "注销失败（键仍存在）：$k" }
        Write-Host "已注销：$k" -ForegroundColor Yellow
    }
    # 清掉 VSTO 的加载项缓存键（残留不影响功能，但会让"是否已注销"的判断变模糊）
    $dataKey = "HKCU:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\AddinsData\PPTShapeLibrary"
    if (Test-Path $dataKey) { Remove-Item $dataKey -Recurse -Force -ErrorAction SilentlyContinue }

    # 清掉 VSTO 的"信任/包含列表"里指向本仓库的条目（只认路径含 PPTShapeLibrary 的，绝不误删别的插件）
    $inclusionRoot = "HKCU:\Software\Microsoft\VSTO\Security\Inclusion"
    if (Test-Path $inclusionRoot) {
        Get-ChildItem $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $url = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Url
            if ($url -like "*PPTShapeLibrary*") {
                Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "已移除 VSTO 信任项：$url" -ForegroundColor Yellow
            }
        }
    }
    # 清掉会让 VSTO 拒绝安装/加载的两类残留（详见上面的函数注释）
    Clear-VstaRecords
    Clear-DevArpRecords

    Write-Host "开发版加载项已注销（注册表两处 + AddinsData 缓存 + 本仓库的 VSTO 信任项 + VSTA/ARP 残留）。" -ForegroundColor Yellow
    return
}

$vsto = Join-Path $Root "PPTShapeLibrary\bin\$Configuration\PPTShapeLibrary.vsto"
$dll  = Join-Path $Root "PPTShapeLibrary\bin\$Configuration\PPTShapeLibrary.dll"
if (-not (Test-Path $vsto)) { throw "找不到 $vsto —— 请先用 $Configuration 配置构建一次" }
if (-not (Test-Path $dll))  { throw "找不到 $dll —— 请先构建" }

$manifest = "file:///" + ($vsto -replace '\\','/') + "|vstolocal"

# 注册前先把"会让 VSTO 拒绝后续正式安装"的残留清掉：
# 否则开发版与正式版来回切换几次后，VSTO 会一直报「已安装另一版本的自定义项」，
# 且表现为"插件突然不加载"（本项目实测踩过多次）。
Clear-VstaRecords
Clear-DevArpRecords

foreach ($k in @($keyVersioned, $keyVersionless)) {
    if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
    Set-ItemProperty -Path $k -Name "Description"  -Value "PPTShapeLibrary"
    Set-ItemProperty -Path $k -Name "FriendlyName" -Value "PPTShapeLibrary"
    Set-ItemProperty -Path $k -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $k -Name "Manifest"     -Value $manifest
    # 插件内"检查更新"要读当前版本。开发模式（|vstolocal）下 ARP 里的版本可能陈旧（来自早期 vsto 清单），
    # 程序集版本也不随发版变 → 会把"当前版本"读成旧值（实测读到 1.0.1.0），于是误报/显示错误。
    # 所以这里把真实版本写进注册表，UpdateService.CurrentVersion() 会**优先**读它。
    $ver = $null
    if ($ProductVersion) {
        $ver = $ProductVersion
        Write-Host "   版本 → $ver（-ProductVersion 指定，用于测试更新流程）" -ForegroundColor Yellow
    } else {
        try { $ver = (Get-Item $dll).VersionInfo.FileVersion } catch { }
        if (-not $ver) {
            $asmInfo = Join-Path $Root "PPTShapeLibrary\Properties\AssemblyInfo.cs"
            if (Test-Path $asmInfo) {
                $mm = Select-String -Path $asmInfo -Pattern 'AssemblyVersion\("([\d\.]+)"\)' | Select-Object -First 1
                if ($mm) { $ver = $mm.Matches[0].Groups[1].Value }
            }
        }
    }
    if ($ver) {
        Set-ItemProperty -Path $k -Name "ProductVersion" -Value $ver
        Write-Host "   版本 → $ver"
    }
    Write-Host "已写入：$k"
}

# -Machine：额外写一份到 HKLM（所有用户）。用于根治"HKCU 版本无关键消失"。
if ($Machine) {
    if (-not (Test-Admin)) {
        Write-Host ""
        Write-Host "！ -Machine 需要管理员权限。请用管理员 PowerShell 重新执行：" -ForegroundColor Yellow
        Write-Host "    pwsh -File scripts\dev-register.ps1 -Machine" -ForegroundColor Yellow
        Write-Host "  （HKCU 两处键已写好，不影响当前用户使用；-Machine 只是加一道保险。）" -ForegroundColor Gray
    }
    else {
        foreach ($k in @($machVersioned, $machVersionless)) {
            if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
            Set-ItemProperty -Path $k -Name "Description"  -Value "PPTShapeLibrary"
            Set-ItemProperty -Path $k -Name "FriendlyName" -Value "PPTShapeLibrary"
            Set-ItemProperty -Path $k -Name "LoadBehavior" -Value 3 -Type DWord
            Set-ItemProperty -Path $k -Name "Manifest"     -Value $manifest
            Write-Host "已写入（所有用户）：$k" -ForegroundColor Green
        }
    }
}

# 最终自检：两处 HKCU 键必须在
$missingNow = @($keyVersioned, $keyVersionless | ForEach-Object { $_ }) | Where-Object { -not (Test-Path $_) }
if ($missingNow.Count -gt 0) {
    Write-Host "！ 注册后自检失败，以下键缺失：" -ForegroundColor Red
    $missingNow | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    exit 1
}

Write-Host ""
Write-Host "开发版加载项已注册：$vsto" -ForegroundColor Green
Write-Host "请【完全退出】并重新打开 PowerPoint —— 功能区应出现「素材库」选项卡。" -ForegroundColor Green
Write-Host "若仍没有：文件 → 选项 → 加载项，看 PPTShapeLibrary 是否在「非活动应用程序加载项」里。" -ForegroundColor Gray
Write-Host "若连加载项列表都是空的，见 docs\PowerPoint-COM经验与坑.md 的『插件在 PPT 里消失』一节。" -ForegroundColor Gray
Write-Host "要根治该键被清掉的问题：用管理员 PowerShell 执行 pwsh -File scripts\dev-register.ps1 -Machine" -ForegroundColor Gray
