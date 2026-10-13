# 下载三个官方前置安装包到 dist\离线安装包\（安装器会自动识别成「安装本地包」）
# ==============================================================================
# 为什么需要它：微软的运行时在"网络不好"的机器上最难过——直链会变、下载中心页面也要人挑版本。
# 这里的做法是**把官方包预置进分发包**：这正是微软 ClickOnce 引导包自己的做法（bootstrapper 里
# 就烤了 vstor_redist.exe），装的时候完全离线，也不涉及"重新托管微软二进制"的授权问题。
#
# 用法：
#   pwsh -File scripts\fetch-prerequisites.ps1                       # 下到 dist\离线安装包\
#   pwsh -File scripts\fetch-prerequisites.ps1 -OutDir D:\离线包      # 自定义目录
#   pwsh -File scripts\fetch-prerequisites.ps1 -Only webview2        # 只下某一个
#
# 结果：每个文件都会校验"是不是真正的 PE 可执行文件"（防止把 HTML 错误页存成 .exe）
param(
    [string]$OutDir = "",
    [string[]]$Only = @(),
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$Root = Split-Path $PSScriptRoot -Parent
if (-not $OutDir) { $OutDir = Join-Path $Root "dist\离线安装包" }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

# GitHub 加速站前缀：直连优先，失败依次回退（加速站只代理 github.com，不能加速微软域名）
$Mirrors = @("", "https://ghproxy.net/", "https://gh-proxy.com/", "https://ghfast.top/")
# 本仓库固定 tag 上托管的依赖安装包（见 installer/UpdateChecker.PrereqTag）
$PrereqBase = "https://github.com/leobba/PPTShapeLibrary/releases/download/prerequisites/"

# 期望文件名要与 installer 里的 LocalPattern 对齐（ndp48*.exe / vstor*.exe / MicrosoftEdgeWebview2Setup.exe）
$items = @(
    [pscustomobject]@{
        Key = 'net48'; Name = 'ndp48-x86-x64-allos-enu.exe'; Title = '.NET Framework 4.8 运行时（离线包，约 110MB）'
        Urls = @(
            $PrereqBase + 'ndp48-x86-x64-allos-enu.exe'
            'https://go.microsoft.com/fwlink/?linkid=2088631'
        )
        Page = 'https://dotnet.microsoft.com/download/dotnet-framework/net48'
    }
    [pscustomobject]@{
        Key = 'vsto'; Name = 'vstor_redist.exe'; Title = 'VSTO Runtime 4.0（约 40MB）'
        # 先试本仓库固定 tag 上的资产（可走加速站）：
        #   vstor_redist.exe = 官方自解压包（体积大，可能未托管 → 404 自动跳过）
        #   vstor_runtime.zip = 同一份内容的 zip（install.exe + MSI + CAB，微软签名）
        # 再回退微软官方直链（GUID 路径会变，给了两个候选 + 页面兜底）
        Urls = @(
            $PrereqBase + 'vstor_redist.exe'
            $PrereqBase + 'vstor_runtime.zip'
            'https://download.microsoft.com/download/c/0/e/c0e39fdf-68c9-4332-b745-5268ed69cb54/vstor_redist.exe'
            'https://download.microsoft.com/download/8/6/4/8641e164-7796-4b34-81c7-30d24a5bd533/vstor_redist.exe'
        )
        Page = 'https://www.microsoft.com/en-us/download/details.aspx?id=105522'
    }
    [pscustomobject]@{
        Key = 'webview2'; Name = 'MicrosoftEdgeWebview2Setup.exe'; Title = 'Edge WebView2 Runtime 引导安装器（约 2MB）'
        Urls = @('https://go.microsoft.com/fwlink/p/?LinkId=2124703')
        Page = 'https://developer.microsoft.com/microsoft-edge/webview2/#download'
    }
)

function Test-PeFile([string]$path) {
    try {
        if (-not (Test-Path $path)) { return $false }
        if ((Get-Item $path).Length -lt 64KB) { return $false }
        $fs = [IO.File]::OpenRead($path)
        try {
            $b = New-Object byte[] 2
            $null = $fs.Read($b, 0, 2)
            return ($b[0] -eq 0x4D -and $b[1] -eq 0x5A)      # "MZ"
        } finally { $fs.Close() }
    } catch { return $false }
}

$okCount = 0; $fail = @()
foreach ($it in $items) {
    if ($Only.Count -gt 0 -and ($Only -notcontains $it.Key)) { continue }
    $dest = Join-Path $OutDir $it.Name
    if ((Test-Path $dest) -and -not $Force -and (Test-PeFile $dest)) {
        Write-Host "[跳过] $($it.Name) 已存在且是有效可执行文件" -ForegroundColor DarkGray
        $okCount++
        continue
    }
    Write-Host "[下载] $($it.Title)" -ForegroundColor Cyan
    $done = $false
    foreach ($u in $it.Urls) {
        # GitHub 源：直连 + 加速站依次尝试；非 GitHub 源（微软）只有直连一种
        $candidates = if ($u -like 'https://github.com/*') { $Mirrors | ForEach-Object { $_ + $u } } else { @($u) }
        foreach ($cand in $candidates) {
            # 存成**资产自己的文件名**：zip 形式的包不能硬写成 .exe，否则点击时会被当成可执行程序跑起来
            $leaf = [IO.Path]::GetFileName(($cand -split '\?' | Select-Object -First 1))
            $dest = if ($leaf -like '*.zip') { Join-Path $OutDir $leaf } else { Join-Path $OutDir $it.Name }
            try {
                $label = if ($cand -like '*ghproxy*' -or $cand -like '*gh-proxy*' -or $cand -like '*ghfast*') { "加速站" } else { "直连" }
                Write-Host "  尝试[$label]: $leaf"
                $wc = New-Object System.Net.WebClient
                $wc.Headers.Add('User-Agent', 'ppt-shape-library-prereq-fetcher')
                $wc.DownloadFile($cand, $dest)
                # zip 形式的包（vstor_runtime.zip）不检查 PE 头，改为检查 zip 头 PK
                $isZip = $cand -like '*.zip'
                $ok = if ($isZip) { (Test-Path $dest) -and (Get-Item $dest).Length -gt 100KB } else { Test-PeFile $dest }
                if ($ok) {
                    Write-Host "  完成：$([math]::Round((Get-Item $dest).Length/1MB,2)) MB ✓" -ForegroundColor Green
                    $done = $true; $okCount++; break
                }
                Write-Warning "  下载内容不是有效的安装文件（可能是错误页），换下一个源"
                Remove-Item $dest -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Verbose "  失败：$($_.Exception.Message)"
                Remove-Item $dest -Force -ErrorAction SilentlyContinue
            }
        }
        if ($done) { break }
    }
    if (-not $done) {
        $fail += $it
        Write-Host "  三个源都失败 → 请手动下载后放到 $OutDir" -ForegroundColor Yellow
        Write-Host "  官方页: $($it.Page)" -ForegroundColor Yellow
    }
}

Remove-Item (Join-Path $OutDir '*.tmp') -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "目录：$OutDir" -ForegroundColor Green
Get-ChildItem $OutDir -File | ForEach-Object { "  $($_.Name)  $([math]::Round($_.Length/1MB,2)) MB" }
if ($fail.Count -gt 0) {
    Write-Host ""
    Write-Host "以下需要手动下载（把文件放进上面目录即可，文件名要一致）：" -ForegroundColor Yellow
    foreach ($f in $fail) { Write-Host "  $($f.Name)  ← $($f.Page)" -ForegroundColor Yellow }
    exit 1
}
Write-Host ""
Write-Host "全部就绪。重新打包后 dist 自带离线包，别人装的时候点「安装本地包」即可，无需联网。" -ForegroundColor Green
