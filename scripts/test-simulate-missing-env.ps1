<#
临时"隐藏"环境检测项，用于测试安装器在**一般电脑**（缺 VSTO / 缺 .NET）上的完整流程。
=====================================================================================
为什么用"隐藏"而不是"卸载"：
  · .NET Framework 4.8 在 Win10 1903+/Win11 是**操作系统组件**，根本没有卸载入口，也无法真正移除；
  · VSTO Runtime 卸载会连带影响机器上**其它 VSTO 插件**（例如方方格子/FFCell），代价太大。
安装器只读注册表判断是否存在，所以"临时把检测用的键挪走"就能真实复现"缺失"状态，
且**不动任何文件、随时可原样还原**（用 reg.exe 导出/导入，逐字节还原）。

用法（**需要管理员**，因为要改 HKLM）：
  pwsh -File scripts\test-simulate-missing-env.ps1 -Hide -Items vsto,net48
      → 备份并移除对应键，写出状态文件，然后**等待** %TEMP%\pptsl_restore.flag（最多 20 分钟）
      → 收到标记或超时后自动原样还原并退出

  pwsh -File scripts\test-simulate-missing-env.ps1 -Restore
      → 手动还原（如果上次异常中断）

参数：
  -Items  vsto / net48 / both（默认 both）
#>
param(
    [switch]$Hide,
    [switch]$Restore,
    [string]$Items = "both",
    [int]$TimeoutMinutes = 20
)

$ErrorActionPreference = "Stop"

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
}

$work = Join-Path $env:TEMP "pptsl_envhide"
$stateFile = Join-Path $work "state.json"
$flagFile = Join-Path $env:TEMP "pptsl_restore.flag"

# 目标键：安装器的检测路径（见 installer/InstallerForm.cs CheckVsto / CheckNet48）
function Get-Targets([string]$items) {
    $list = @()
    if ($items -eq "vsto" -or $items -eq "both") {
        $list += @{ Name = "VSTO(64)"; Reg64 = "HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"; View = 64 }
        $list += @{ Name = "VSTO(32)"; Reg64 = "HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"; View = 32 }
    }
    if ($items -eq "net48" -or $items -eq "both") {
        $list += @{ Name = "NET48"; Reg64 = "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"; View = 64 }
    }
    return $list
}

function Invoke-Reg([string[]]$regArgs) {
    $out = & reg.exe @regArgs 2>&1
    return @{ Code = $LASTEXITCODE; Out = ($out -join "`n") }
}

# 还原后自校验：读回注册表并与备份 .reg 里的 Version 比对（防"以为还原了其实没有"）
function Show-RestoreCheck($state) {
    Write-Host "还原校验：" -ForegroundColor Cyan
    foreach ($t in $state.targets) {
        $view = if ($t.View -eq 32) { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }
        $now = $null
        try {
            $k = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view).OpenSubKey('SOFTWARE\Microsoft\VSTO Runtime Setup\v4R')
            if ($k) { $now = $k.GetValue('Version'); $k.Close() }
        } catch { }
        $want = $null
        if (Test-Path $t.BackupFile) {
            $m = Select-String -Path $t.BackupFile -Pattern '"Version"="([^"]+)"' | Select-Object -First 1
            if ($m) { $want = $m.Matches[0].Groups[1].Value }
        }
        if ($null -eq $want) {
            Write-Host ("  {0}: 备份无 Version（跳过比对），当前={1}" -f $t.Name, $now)
        } elseif ($now -eq $want) {
            Write-Host ("  {0}: OK（Version={1}）" -f $t.Name, $now) -ForegroundColor Green
        } else {
            Write-Host ("  {0}: 不一致！期望 {1}，实际 {2}" -f $t.Name, $want, $now) -ForegroundColor Red
        }
    }
}

if ($Restore) {
    if (-not (Test-Admin)) { Write-Host "需要管理员权限。" -ForegroundColor Red; exit 1 }
    if (-not (Test-Path $stateFile)) { Write-Host "没有找到状态文件，无需还原：$stateFile" -ForegroundColor Yellow; exit 0 }
    $state = Get-Content $stateFile -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($t in $state.targets) {
        $f = if ($t.ImportFile -and (Test-Path $t.ImportFile)) { $t.ImportFile } else { $t.BackupFile }
        if (Test-Path $f) {
            $r = Invoke-Reg @("import", $f)
            Write-Host ("还原 {0}: {1}" -f $t.Name, ($(if ($r.Code -eq 0) { "OK" } else { "失败 " + $r.Out })))
        }
    }
    Show-RestoreCheck $state
    Remove-Item $stateFile -Force -ErrorAction SilentlyContinue
    Write-Host "已还原。" -ForegroundColor Green
    exit 0
}

if (-not $Hide) { Write-Host "请指定 -Hide 或 -Restore。"; exit 2 }
if (-not (Test-Admin)) {
    Write-Host "需要管理员权限（要改 HKLM）。请用管理员 PowerShell 运行本脚本。" -ForegroundColor Red
    exit 1
}

New-Item -ItemType Directory -Path $work -Force | Out-Null
Remove-Item $flagFile -Force -ErrorAction SilentlyContinue

$targets = @()
foreach ($t in (Get-Targets $Items)) {
    $backup = Join-Path $work (($t.Name -replace '[^\w]', '_') + ".reg")
    $args = @("export", $t.Reg64, $backup, "/y")
    if ($t.View -eq 32) { $args += "/reg:32" }
    $r = Invoke-Reg $args
    if ($r.Code -ne 0) {
        Write-Host ("跳过 {0}：导出失败（可能不存在）{1}" -f $t.Name, $r.Out) -ForegroundColor Yellow
        continue
    }
    # ⚠️ reg.exe 的 import **没有 /reg:32**：32 位视图的备份必须把路径改写为 WOW6432Node 才能导入，
    #    否则会（a）导入到 64 位视图把它污染，（b）32 位键永远回不来。这里在导出时就生成"可导入"副本。
    #    注意 PowerShell 单引号里 \ 是字面字符，用 .Replace 精确替换（不要用正则，\\ 会变成两个反斜杠）。
    $importFile = $backup
    if ($t.View -eq 32) {
        $txt = Get-Content $backup -Raw
        $txt = $txt.Replace('SOFTWARE\Microsoft\', 'SOFTWARE\WOW6432Node\Microsoft\')
        $importFile = Join-Path $work (($t.Name -replace '[^\w]', '_') + "_wow.reg")
        Set-Content $importFile -Value $txt -Encoding Unicode
    }
    $del = @("delete", $t.Reg64, "/f")
    if ($t.View -eq 32) { $del += "/reg:32" }
    $d = Invoke-Reg $del
    if ($d.Code -ne 0) {
        Write-Host ("删除 {0} 失败：{1}" -f $t.Name, $d.Out) -ForegroundColor Red
        continue
    }
    Write-Host ("已隐藏 {0}（备份 {1}）" -f $t.Name, $backup) -ForegroundColor Yellow
    $targets += @{ Name = $t.Name; BackupFile = $backup; ImportFile = $importFile; View = $t.View; RegPath = $t.Reg64 }
}

if ($targets.Count -eq 0) { Write-Host "没有任何键被隐藏。" -ForegroundColor Red; exit 1 }
@{ targets = $targets } | ConvertTo-Json -Depth 5 | Set-Content $stateFile -Encoding UTF8

Write-Host ""
Write-Host "现在安装器会把这些项判为**缺失**。测试完成后：" -ForegroundColor Cyan
Write-Host "  · 自动还原：创建标记文件（下面的命令会在超时/标记后自动还原）"
Write-Host ("  · 或手动：pwsh -File {0} -Restore" -f $PSCommandPath)
Write-Host ("等待标记文件（最多 {0} 分钟）：{1}" -f $TimeoutMinutes, $flagFile) -ForegroundColor Cyan

$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
while ((Get-Date) -lt $deadline) {
    if (Test-Path $flagFile) { break }
    Start-Sleep -Seconds 2
}

Write-Host "开始还原…" -ForegroundColor Cyan
$state = Get-Content $stateFile -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($t in $state.targets) {
    $f = if ($t.ImportFile -and (Test-Path $t.ImportFile)) { $t.ImportFile } else { $t.BackupFile }
    if (Test-Path $f) {
        $r = Invoke-Reg @("import", $f)
        Write-Host ("还原 {0}: {1}" -f $t.Name, ($(if ($r.Code -eq 0) { "OK" } else { "失败 " + $r.Out })))
    }
}
Show-RestoreCheck $state
Remove-Item $stateFile -Force -ErrorAction SilentlyContinue
Remove-Item $flagFile -Force -ErrorAction SilentlyContinue
Write-Host "已原样还原，可以关闭这个窗口。" -ForegroundColor Green
