# exp-insert-noclip.ps1 —— 判决性实验：形状插入能否彻底不用剪贴板？
#
# 【结论（2026-10 实测，已固化）】
#   形状插入**无法**免剪贴板；但剪贴板冲突是**可等待、可恢复**的。
#   详见 docs\PowerPoint-COM经验与坑.md §14。
#
# 各路径在"剪贴板被第三方（如 Snipaste）占用"时的表现：
#   A 跨文稿 Copy → Paste（现状形状插入）            → 失败（复现用户报错）
#   B InsertFromFile 插整页 + 同文稿内搬形状          → 失败（同文稿内搬形状也要剪贴板）
#   C 整页插入 InsertFromFile（现有做法）             → 成功
#   D InsertFromFile + Slide.Duplicate - 删临时页     → 成功（Duplicate 不碰剪贴板，但只复制整页）
#   E 重试策略（Clear + 等待 + 重试）                 → 最终成功（等占用者释放的间隙）
#
# 用法（PowerShell 5.1 / WindowsPowerShell，因为要用 COM + 正确的 STA）：
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test\exp-insert-noclip.ps1
param(
    [string]$Root = "E:\code\PPTShapeLibrary"
)

$ErrorActionPreference = "Continue"

# PowerPoint 枚举值（COM 不接受 bool）
$msoTrue = -1
$msoFalse = 0

Add-Type -Namespace PptslClip -Name Api -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool OpenClipboard(IntPtr hWndNewOwner);
[DllImport("user32.dll")] public static extern bool CloseClipboard();
'@

$outDir = Join-Path $Root "build\exp-insert-noclip"
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# ---------------------------------------------------------------- 准备：挑一个代表性库条目
$itemsDir = Join-Path $env:APPDATA "PPTShapeLibrary\items"
if (-not (Test-Path $itemsDir)) { throw "no library: $itemsDir" }
$cands = Get-ChildItem $itemsDir -Directory | ForEach-Object {
    $p = Join-Path $_.FullName "shape.pptx"
    if (Test-Path $p) { [pscustomobject]@{ Name = $_.Name; Path = $p; Size = (Get-Item $p).Length } }
} | Where-Object { $_.Size -ge 50KB } | Sort-Object Size
if (-not $cands) { throw "no shape.pptx found" }
$lib = $cands[[math]::Floor($cands.Count / 2)]
Write-Host "source item: $($lib.Name)  $([math]::Round($lib.Size/1KB)) KB   (of $($cands.Count) candidates)"
$libPath = $lib.Path

function New-Target($ppt, [string]$savePath) {
    $pres = $ppt.Presentations.Add($msoFalse)
    if ($pres.Slides.Count -eq 0) { $null = $pres.Slides.Add(1, 12) }   # 12 = ppLayoutBlank
    if ($savePath) { $pres.SaveAs($savePath) }
    return $pres
}

$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    Write-Host "PowerPoint $($ppt.Version)"

    # ============================================================ A
    Write-Host ""
    Write-Host "=== A: cross-doc Copy -> Paste (current shape insert) ==="
    $pres = New-Target $ppt (Join-Path $outDir "A-copypaste.pptx")
    try {
        $locked = [PptslClip.Api]::OpenClipboard([IntPtr]::Zero)
        Write-Host "  clipboard locked: $locked"
        $src = $ppt.Presentations.Open($libPath, $msoTrue, $msoFalse, $msoFalse)
        try {
            $s = $src.Slides.Item(1)
            $ids = @(); for ($i = 1; $i -le $s.Shapes.Count; $i++) { $ids += $i }
            $s.Shapes.Range($ids).Copy()
            Start-Sleep -Milliseconds 300
            $pasted = $pres.Slides.Item(1).Shapes.Paste()
            Write-Host "  RESULT: ok (shapes=$($pasted.Count))"
        } finally { $src.Close() }
    } catch {
        Write-Host "  RESULT: FAILED -> $($_.Exception.Message)"
    } finally {
        [PptslClip.Api]::CloseClipboard() | Out-Null
        $pres.Close()
    }

    # ============================================================ B
    Write-Host ""
    Write-Host "=== B: InsertFromFile page + SAME-DOC move shapes ==="
    $pres = New-Target $ppt (Join-Path $outDir "B-samedoc-move.pptx")
    try {
        $locked = [PptslClip.Api]::OpenClipboard([IntPtr]::Zero)
        $n0 = $pres.Slides.Count
        $pres.Slides.InsertFromFile($libPath, $n0, 1, 1)      # 不碰剪贴板
        $n1 = $pres.Slides.Count
        Write-Host ("  InsertFromFile ok (slides {0} -> {1})" -f $n0, $n1)
        $temp = $pres.Slides.Item($n1)
        $dest = $pres.Slides.Item(1)
        $ids = @(); for ($i = 1; $i -le $temp.Shapes.Count; $i++) { $ids += $i }
        $temp.Shapes.Range($ids).Copy()
        Start-Sleep -Milliseconds 300
        $pasted = $dest.Shapes.Paste()                         # 同文稿内仍然要剪贴板
        Write-Host "  RESULT: ok (moved=$($pasted.Count))"
    } catch {
        Write-Host "  RESULT: FAILED -> $($_.Exception.Message)"
    } finally {
        [PptslClip.Api]::CloseClipboard() | Out-Null
        $pres.Close()
    }

    # ============================================================ C
    Write-Host ""
    Write-Host "=== C: page insert via InsertFromFile (current) ==="
    $pres = New-Target $ppt (Join-Path $outDir "C-page-insert.pptx")
    try {
        $locked = [PptslClip.Api]::OpenClipboard([IntPtr]::Zero)
        $n = $pres.Slides.Count
        $pres.Slides.InsertFromFile($libPath, $n, 1, 1)
        Write-Host "  RESULT: ok (slides -> $($pres.Slides.Count))"
    } catch {
        Write-Host "  RESULT: FAILED -> $($_.Exception.Message)"
    } finally {
        [PptslClip.Api]::CloseClipboard() | Out-Null
        $pres.Close()
    }

    # ============================================================ D
    Write-Host ""
    Write-Host "=== D: InsertFromFile + Slide.Duplicate - delete temp ==="
    $pres = New-Target $ppt (Join-Path $outDir "D-duplicate.pptx")
    try {
        $locked = [PptslClip.Api]::OpenClipboard([IntPtr]::Zero)
        $n = $pres.Slides.Count
        $pres.Slides.InsertFromFile($libPath, $n, 1, 1)
        $ti = $pres.Slides.Count
        $null = $pres.Slides.Item($ti).Duplicate()             # Duplicate 不碰剪贴板
        $pres.Slides.Item($ti).Delete()
        Write-Host "  RESULT: ok (slides -> $($pres.Slides.Count))"
    } catch {
        Write-Host "  RESULT: FAILED -> $($_.Exception.Message)"
    } finally {
        [PptslClip.Api]::CloseClipboard() | Out-Null
        $pres.Close()
    }
}
finally {
    if ($ppt) { $ppt.Quit() }
}

# ================================================================ E（不需要 PowerPoint）
Write-Host ""
Write-Host "=== E: retry strategy vs PERIODIC clipboard holder (simulating Snipaste) ==="
Write-Host "    holder: lock 3s / release 1s, 3 rounds"
$job = Start-Job -ScriptBlock {
    Add-Type -Namespace H -Name C -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool OpenClipboard(IntPtr hWndNewOwner);
[DllImport("user32.dll")] public static extern bool CloseClipboard();
'@
    for ($i = 0; $i -lt 3; $i++) {
        [H.C]::OpenClipboard([IntPtr]::Zero) | Out-Null
        Start-Sleep -Seconds 3
        [H.C]::CloseClipboard() | Out-Null
        Start-Sleep -Seconds 1
    }
}
Start-Sleep -Milliseconds 400
Add-Type -AssemblyName System.Windows.Forms
$sw = [Diagnostics.Stopwatch]::StartNew()
$done = $false
for ($a = 1; $a -le 6; $a++) {
    Start-Sleep -Milliseconds 200
    try { [System.Windows.Forms.Clipboard]::Clear(); $c = "ok" } catch { $c = "blocked" }
    try {
        [System.Windows.Forms.Clipboard]::SetText("probe")
        Write-Host "  attempt ${a}: Clear=$c, write=OK  -> recovered"
        $done = $true
        break
    } catch {
        Write-Host "  attempt ${a}: Clear=$c, write=blocked"
    }
}
$sw.Stop()
Write-Host "  RESULT: recovered=$done in $([math]::Round($sw.Elapsed.TotalSeconds,1))s"
Get-Job | Remove-Job -Force

Write-Host ""
Write-Host "artifacts: $outDir"
