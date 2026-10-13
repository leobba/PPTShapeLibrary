<#
PPTShapeLibrary 端到端冒烟（真机 PowerPoint + WebView2 + CDP）
================================================================
把"改完 ShapeExporter/桥之后必须跑一遍"的人工回归自动化：
  1) 前置检查（PowerPoint、插件两处注册表键）
  2) 备份素材库 meta.json（跑完逐字节还原，测试不会污染用户数据）
  3) 带 --remote-debugging-port 启动 PowerPoint + 测试文稿
  4) UI Automation 点 Ribbon「素材库」→「打开素材库」
  5) 调 scripts/e2e_assert.py 通过 CDP 驱动真实前端做断言
  6) **正常退出** PowerPoint（切勿强杀：实测强杀会丢掉"版本无关"注册表键 → 插件消失）
  7) 还原 meta.json，输出 PASS/FAIL

用法：
  pwsh -File scripts\e2e-smoke.ps1
  pwsh -File scripts\e2e-smoke.ps1 -Deck "D:\some.pptx" -Port 9333

退出码：0=通过；1=断言失败；2=前置条件不满足（未装 PowerPoint / 插件未注册 / 已有 PowerPoint 在运行）。
#>
param(
    [string]$Deck = "",
    [int]$Port = 9333,
    [switch]$Force,      # 允许在 PowerPoint 已运行时继续（会干扰用户当前会话，默认拒绝）
    [switch]$Installed   # 测"正式安装版"（ClickOnce）：只要求版本无关键，且**绝不**用 dev-register 改写注册
)

$ErrorActionPreference = "Stop"
$Root = Split-Path $PSScriptRoot -Parent
$Ppt = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
$Keys = @(
    "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PPTShapeLibrary",
    "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Addins\PPTShapeLibrary"
)
$Sentinel = "E2E-SENTINEL-" + ([guid]::NewGuid().ToString("N").Substring(0, 8))

function Fail($msg, $code = 2) { Write-Host "前置条件不满足：$msg" -ForegroundColor Red; exit $code }

# 注册表键采样：这个"版本无关"键曾因 Release 构建被注销而消失，逐阶段记录便于定位
function Show-Keys($phase) {
    $v = Test-Path $Keys[0]
    $k = Test-Path $Keys[1]
    Write-Host ("      [keys] {0,-30} 版本无关={1,-6} 版本化={2}" -f $phase, $v, $k) -ForegroundColor DarkGray
}

Write-Host "== 1/7 前置检查 ==" -ForegroundColor Cyan
if (-not (Test-Path $Ppt)) { Fail "未找到 PowerPoint：$Ppt" }
$running = @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
if ($running.Count -gt 0 -and -not $Force) {
    Fail "PowerPoint 正在运行（PID $($running[0].Id)）。请先关闭它，或用 -Force 强制（会打断当前会话）。"
}
if ($Installed) {
    # 正式安装（ClickOnce）只写"版本无关"键，Manifest 不带 |vstolocal。
    # 这种模式下绝不能跑 dev-register.ps1——那会把正式注册改写成 bin\Debug 的开发式注册。
    if (-not (Test-Path $Keys[0])) {
        Fail "未检测到正式安装的加载项注册（$($Keys[0])）。请先用安装器安装，或改用不带 -Installed 的开发模式。"
    }
    $mf = (Get-ItemProperty $Keys[0] -ErrorAction SilentlyContinue).Manifest
    if ($mf -like "*vstolocal*") {
        Fail "当前是**开发式**注册（$mf）。-Installed 只用于测试正式安装版；请去掉 -Installed。"
    }
    Write-Host "  正式安装版：$mf" -ForegroundColor Green
} else {
    $missing = @($Keys | Where-Object { -not (Test-Path $_) })
    if ($missing.Count -gt 0) {
        Write-Host "检测到注册表键缺失，自动修复（dev-register.ps1）：" -ForegroundColor Yellow
        $missing | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
        & pwsh -NoProfile -File (Join-Path $PSScriptRoot "dev-register.ps1") | Out-Null
        $still = @($Keys | Where-Object { -not (Test-Path $_) })
        if ($still.Count -gt 0) { Fail "注册表键修复失败，请手动执行 scripts\dev-register.ps1" }
        Write-Host "  已修复" -ForegroundColor Green
    }
    Write-Host "  开发模式：PowerPoint / 注册表两处键 均就绪" -ForegroundColor Green
}

Write-Host "== 2/7 备份素材库元数据 ==" -ForegroundColor Cyan
$AppData = Join-Path $env:APPDATA "PPTShapeLibrary"
$ItemsDir = Join-Path $AppData "items"
$Backup = Join-Path $env:TEMP ("pptsl_e2e_backup_" + (Get-Date -Format yyyyMMdd_HHmmss))
New-Item -ItemType Directory -Path $Backup -Force | Out-Null
$metaCount = 0
if (Test-Path $ItemsDir) {
    foreach ($d in Get-ChildItem $ItemsDir -Directory) {
        $m = Join-Path $d.FullName "meta.json"
        if (Test-Path $m) {
            $dest = Join-Path $Backup $d.Name
            New-Item -ItemType Directory -Path $dest -Force | Out-Null
            Copy-Item $m (Join-Path $dest "meta.json") -Force
            $metaCount++
        }
    }
}
Write-Host "  已备份 $metaCount 个 meta.json → $Backup" -ForegroundColor Green

Write-Host "== 3/7 启动 PowerPoint（远程调试端口 $Port）==" -ForegroundColor Cyan
if ([string]::IsNullOrWhiteSpace($Deck)) {
    $Deck = Join-Path $env:TEMP "pptsl_e2e_deck.pptx"
    # 每次重建：需要 ~30 页才能观察到"长任务进度/取消"（导出图片是逐页的）
    if (Test-Path $Deck) { Remove-Item $Deck -Force -ErrorAction SilentlyContinue }
    python -c @"
from pptx import Presentation
from pptx.util import Inches, Pt
p = Presentation(); p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
for i in range(1, 31):
    s = p.slides.add_slide(p.slide_layouts[6])
    tb = s.shapes.add_textbox(Inches(0.6), Inches(0.6), Inches(8), Inches(1.2))
    r = tb.text_frame.paragraphs[0].add_run(); r.text = 'E2E 第 %d 页' % i; r.font.size = Pt(28)
    sh = s.shapes.add_shape(1, Inches(0.6), Inches(2.2), Inches(3), Inches(1.6)); sh.fill.solid()
    s.shapes.add_textbox(Inches(4.5), Inches(2.2), Inches(6), Inches(2)).text_frame.text = '内容 %d' % i
p.save(r'$Deck')
"@
}
if (-not (Test-Path $Deck)) { Fail "测试文稿不可用：$Deck（需要 python-pptx）" }
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=$Port"
Start-Process -FilePath $Ppt -ArgumentList "`"$Deck`"" | Out-Null
# 剪贴板哨兵：整页插入不得触碰剪贴板。用**短命子进程**设置，避免本脚本自己成为剪贴板持有者
# （实测：脚本进程阻塞在同步子调用时，它持有的剪贴板会让"形状插入"的 Copy 失败）。
& pwsh -NoProfile -Command "Set-Clipboard -Value '$Sentinel'" | Out-Null
Start-Sleep -Seconds 14

Write-Host "== 4/7 打开任务窗格（UI Automation）==" -ForegroundColor Cyan
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]
function Find-ByName($r, $n) { $r.FindFirst($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $n))) }
$cls = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, "PPTFrameClass")
$win = $null; $tab = $null
for ($i = 0; $i -lt 15; $i++) {
    $win = $AE::RootElement.FindFirst($TS::Children, $cls)
    if ($win) { $tab = Find-ByName $win "素材库"; if ($tab) { break } }
    Start-Sleep -Seconds 2
}
if (-not $tab) {
    Write-Host "  找不到「素材库」选项卡（加载项未加载？）" -ForegroundColor Red
} else {
    $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
    Start-Sleep -Milliseconds 1200
    $btn = Find-ByName $win "打开素材库"
    if ($btn) { $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); Write-Host "  已点开素材库面板" -ForegroundColor Green }
    else { Write-Host "  找不到「打开素材库」按钮" -ForegroundColor Red }
}
Start-Sleep -Seconds 6

Write-Host "== 5/7 CDP 断言 ==" -ForegroundColor Cyan
$rc = 1
try {
    python (Join-Path $PSScriptRoot "e2e_assert.py") $Port $Sentinel
    $rc = $LASTEXITCODE
} catch {
    Write-Host "断言脚本异常：$($_.Exception.Message)" -ForegroundColor Red
    $rc = 1
}

Show-Keys "CDP 断言结束"
python -c @"
import sys, time
import win32com.client
try:
    app = win32com.client.Dispatch('PowerPoint.Application')
    for i in range(1, app.Presentations.Count + 1):
        try: app.Presentations(i).Saved = True
        except Exception: pass
    app.Quit(); time.sleep(4); print('  PowerPoint 已正常退出')
except Exception as e:
    print('  退出失败（请手动关闭 PowerPoint）：', e)
"@
Start-Sleep -Seconds 3
# COM Quit 有时不真正结束进程，残留实例会让下一次冒烟的前置检查失败 → 兜底清理
$left = @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
if ($left.Count -gt 0) {
    Write-Host "  仍有 $($left.Count) 个残留 PowerPoint 进程，强制结束（之后会重新校验注册表）" -ForegroundColor Yellow
    $left | ForEach-Object { try { Stop-Process -Id $_.Id -Force } catch { } }
    Start-Sleep -Seconds 2
}

Show-Keys "退出+清理之后"
$restored = 0
if (Test-Path $Backup) {
    foreach ($d in Get-ChildItem $Backup -Directory) {
        $src = Join-Path $d.FullName "meta.json"
        $dst = Join-Path (Join-Path $ItemsDir $d.Name) "meta.json"
        if ((Test-Path $src) -and (Test-Path $dst)) { Copy-Item $src $dst -Force; $restored++ }
    }
}
Write-Host "  已还原 $restored 个 meta.json" -ForegroundColor Green
# 兜底：若上面强杀过残留进程导致注册表键丢失，这里再修一次，避免留下"插件打不开"的环境。
# 注意：-Installed（正式安装版）模式下**不能**跑 dev-register，否则会把 ClickOnce 注册改写成开发式注册。
$lost = @()
if (-not $Installed) { $lost = @($Keys | Where-Object { -not (Test-Path $_) }) }
if ($lost.Count -gt 0) {
    Write-Host "警告：注册表键丢失，自动修复中：" -ForegroundColor Yellow
    $lost | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    & pwsh -NoProfile -File (Join-Path $PSScriptRoot "dev-register.ps1") | Out-Null
    foreach ($k in $Keys) {
        if (Test-Path $k) { Write-Host "  已修复：$k" -ForegroundColor Green }
        else { Write-Host "  修复失败（请手动执行 scripts\dev-register.ps1）：$k" -ForegroundColor Red }
    }
}

Write-Host ""
if ($rc -eq 0) { Write-Host "E2E SMOKE: PASS" -ForegroundColor Green } else { Write-Host "E2E SMOKE: FAIL" -ForegroundColor Red }
exit $rc
