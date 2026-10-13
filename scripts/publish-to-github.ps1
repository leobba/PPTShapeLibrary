# 幂等 GitHub 推送脚本
# ========================
# 目标：把本仓库（含全部历史）推到 GitHub，可重复执行不报错。
# 凭据：Classic PAT（scope: repo + workflow），放 .env 的 GITHUB_TOKEN（.env 已 gitignore）。
#
# 用法：
#   pwsh -File scripts\publish-to-github.ps1                    # 用 .env 里的 GITHUB_REPO
#   pwsh -File scripts\publish-to-github.ps1 -Repo leobba/PPTShapeLibrary
#   pwsh -File scripts\publish-to-github.ps1 -Proxy http://127.0.0.1:7890   # 网络受限时走代理
#
# 网络受限时的三种办法（本机实测 github.com:443 被墙、api.github.com 通）：
#   ① 开代理：-Proxy http://127.0.0.1:7890（或写进 .env 的 GIT_PROXY）
#   ② ssh.github.com:443（多数网络可达），需要先把公钥加到 GitHub 账号
#   ③ 只推 tag/分支不可行时，用 API 建 Release 上传资产（见 publish-release.ps1）
param(
    [string]$Repo = "",
    [string]$Proxy = "",
    [string]$Branch = "master",
    [switch]$SkipTags
)

$ErrorActionPreference = "Stop"
$Root = Split-Path $PSScriptRoot -Parent
$envFile = Join-Path $Root ".env"

function Read-DotEnv([string]$path) {
    $map = @{}
    if (-not (Test-Path $path)) { return $map }
    foreach ($line in (Get-Content $path)) {
        if ($line -match '^\s*#' -or $line -notmatch '=') { continue }
        $i = $line.IndexOf('=')
        $k = $line.Substring(0, $i).Trim()
        $v = $line.Substring($i + 1).Trim().Trim('"').Trim("'")
        if ($k) { $map[$k] = $v }
    }
    return $map
}

$env_ = Read-DotEnv $envFile
$token = $env:GITHUB_TOKEN
if (-not $token) { $token = $env_['GITHUB_TOKEN'] }
if (-not $token) { throw "未找到 GITHUB_TOKEN。请复制 .env.example 为 .env 并填入 Classic PAT（scope: repo + workflow）。" }
if (-not $Repo) { $Repo = $env_['GITHUB_REPO'] }
if (-not $Repo) { throw "未指定仓库。用 -Repo owner/name，或在 .env 里写 GITHUB_REPO=owner/name。" }
if (-not $Proxy -and $env_['GIT_PROXY']) { $Proxy = $env_['GIT_PROXY'] }

$headers = @{
    Authorization          = "Bearer $token"
    Accept                 = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
    "User-Agent"           = "ppt-shape-library-publisher"
}

Write-Host "== 1/5 校验凭据 ==" -ForegroundColor Cyan
$me = Invoke-RestMethod -Uri "https://api.github.com/user" -Headers $headers -TimeoutSec 30
Write-Host "  身份: $($me.login)" -ForegroundColor Green

Write-Host "== 2/5 确保仓库存在 ==" -ForegroundColor Cyan
$parts = $Repo.Split('/')
if ($parts.Count -ne 2) { throw "仓库格式应为 owner/name，当前：$Repo" }
$owner, $name = $parts[0], $parts[1]
$exists = $true
try { $null = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$name" -Headers $headers -TimeoutSec 30 }
catch {
    if ($_.Exception.Response.StatusCode.value__ -eq 404) { $exists = $false } else { throw }
}
if (-not $exists) {
    $body = @{
        name        = $name
        description = "PowerPoint 素材库与 PPT 工具插件（VSTO + WebView2 + React）"
        private     = $false
        has_issues  = $true
        has_wiki    = $false
        has_projects = $false
    } | ConvertTo-Json
    $null = Invoke-RestMethod -Uri "https://api.github.com/user/repos" -Method Post -Headers $headers `
        -Body $body -ContentType "application/json" -TimeoutSec 60
    Write-Host "  已创建公开仓库 $owner/$name" -ForegroundColor Green
} else {
    Write-Host "  仓库已存在：$owner/$name"
    $info = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$name" -Headers $headers -TimeoutSec 30
    if ($info.private) {
        Write-Warning "  该仓库是**私有**的 → 自动更新检查会静默失效（更新器不带凭据抓清单）。建议改成公开。"
    }
}

Write-Host "== 3/5 配置远端 ==" -ForegroundColor Cyan
$url = "https://github.com/$owner/$name.git"
Push-Location $Root
try {
    $existing = git remote get-url origin 2>$null
    if ($LASTEXITCODE -ne 0) { git remote add origin $url; Write-Host "  已添加 origin = $url" }
    elseif ($existing -ne $url) { git remote set-url origin $url; Write-Host "  已更新 origin → $url" }
    else { Write-Host "  origin 已是 $url" }

    # 用一次性带 token 的 URL 推送，token 不落进 .git/config
    $pushUrl = "https://x-access-token:$token@github.com/$owner/$name.git"
    $gitArgs = @()
    if ($Proxy) { $gitArgs += @("-c", "http.proxy=$Proxy") }
    if ($Proxy) { Write-Host "  使用代理：$Proxy" -ForegroundColor Yellow }

    Write-Host "== 4/5 推送 $Branch ==" -ForegroundColor Cyan
    & git @gitArgs push $pushUrl $Branch 2>&1 | ForEach-Object { "  $_" }
    if ($LASTEXITCODE -ne 0) {
        throw "推送失败。若报 'Could not connect to server' → 当前网络到 github.com:443 不通：" +
              "请用 -Proxy http://127.0.0.1:<端口>，或改用 ssh.github.com:443（需先把公钥加到 GitHub 账号）。"
    }

    if (-not $SkipTags) {
        Write-Host "== 5/5 推送标签 / 备份分支 ==" -ForegroundColor Cyan
        & git @gitArgs push $pushUrl --tags 2>&1 | ForEach-Object { "  $_" }
        foreach ($b in @("backup/fidelity-work")) {
            $has = git branch --list $b
            if ($has) { & git @gitArgs push $pushUrl $b 2>&1 | ForEach-Object { "  $_" } }
        }
    }
} finally {
    Pop-Location
}

Write-Host ""
Write-Host "完成：" -ForegroundColor Green
Write-Host "  仓库      https://github.com/$owner/$name"
Write-Host "  发布页    https://github.com/$owner/$name/releases"
Write-Host "  更新清单  https://github.com/$owner/$name/releases/latest/download/update.json"
Write-Host "  下一步    发版：pwsh -File scripts\publish-release.ps1 -Version 1.0.2.0"
