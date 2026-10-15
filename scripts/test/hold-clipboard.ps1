# hold-clipboard.ps1 —— 模拟 Snipaste 之类的剪贴板工具"占用剪贴板"
#
# 用途：复现"插入/收藏形状时报剪贴板被占用"的真实场景，用来验证
#       Copy/Paste 是否真的会被第三方剪贴板工具挡住，以及替代方案能否绕过。
#
# 原理：Snipaste 这类工具会调用 OpenClipboard() 并在一段时间内不释放，
#       此期间其它程序的 Copy()/Paste() 都会失败（ERROR_CLIPBOARD_NOT_AVAILABLE 等）。
#
# 用法：
#   powershell -File scripts\test\hold-clipboard.ps1 -Seconds 20
param(
    [int]$Seconds = 15,
    [switch]$Loop   # 反复占用-释放，更接近真实工具的"监控"行为
)

Add-Type -Namespace Pptsl -Name Clip -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool OpenClipboard(IntPtr hWndNewOwner);
[DllImport("user32.dll")] public static extern bool CloseClipboard();
[DllImport("user32.dll")] public static extern bool EmptyClipboard();
[DllImport("user32.dll")] public static extern IntPtr GetOpenClipboardWindow();
'@

function Hold-Once([int]$sec) {
    $ok = [Pptsl.Clip]::OpenClipboard([IntPtr]::Zero)
    if (-not $ok) {
        Write-Host "  [占用器] OpenClipboard 失败（剪贴板已被别的程序占用）" -ForegroundColor Yellow
        return $false
    }
    Write-Host "  [占用器] 已占用剪贴板 $sec 秒（模拟 Snipaste）…" -ForegroundColor Cyan
    Start-Sleep -Seconds $sec
    [Pptsl.Clip]::CloseClipboard() | Out-Null
    Write-Host "  [占用器] 已释放" -ForegroundColor Cyan
    return $true
}

if ($Loop) {
    # 真实剪贴板工具会周期性开合（监控剪贴板变化）。这里 5 秒占用 / 1 秒释放，循环若干次。
    for ($i = 1; $i -le 6; $i++) {
        Write-Host "  第 $i 轮"
        Hold-Once 5 | Out-Null
        Start-Sleep -Seconds 1
    }
} else {
    Hold-Once $Seconds | Out-Null
}
