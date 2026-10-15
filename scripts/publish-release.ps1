# 一键发版：构建分发包 → 打 tag → 建 GitHub Release → 上传资产（含 update.json）
# ============================================================================
# 为什么不用 GitHub Actions 构建：本插件是 **VSTO** 工程，CI 的 windows runner 没有
# Office/SharePoint 工作负载（Microsoft.VisualStudio.Tools.Office.targets），也没装 Office，
# 构建不稳；所以这里用**本地构建 + API 上传**。日常发版只需要一个 token（放 .env）。
#
# 用法：
#   pwsh -File scripts\publish-release.ps1 -Version 1.0.2.0
#   pwsh -File scripts\publish-release.ps1 -Version 1.0.2.0 -Notes "修复 xxx"
#   pwsh -File scripts\publish-release.ps1 -Version 1.0.2.0 -Proxy http://127.0.0.1:7890 -SkipPush
#
# 产出（每个 Release 都上传，update.json 同名覆盖 → 更新器只需一个固定地址）：
#   PPTShapeLibrary-<ver>-dist.zip        完整分发包（安装器 + App\ + 离线安装包\）
#   PPTShapeLibrary-<ver>-dist.zip.sha256 校验值
#   update.json                           更新清单（固定地址，见 UpdateChecker.UpdateJsonUrl）
param(
    [Parameter(Mandatory = $true)][string]$Version,
    [string]$Repo = "",
    [string]$Notes = "",
    [string]$Proxy = "",
    [switch]$SkipPush,      # 不推 tag（只建 Release 并上传，用于网络受限时先试）
    [switch]$SkipBuild
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
if (-not $token) { throw "未找到 GITHUB_TOKEN（.env）" }
if (-not $Repo) { $Repo = $env_['GITHUB_REPO'] }
if (-not $Repo) { throw "未指定仓库（-Repo 或 .env 的 GITHUB_REPO）" }
if (-not $Proxy -and $env_['GIT_PROXY']) { $Proxy = $env_['GIT_PROXY'] }
$parts = $Repo.Split('/'); $owner = $parts[0]; $name = $parts[1]

$headers = @{
    Authorization          = "Bearer $token"
    Accept                 = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
    "User-Agent"           = "ppt-shape-library-publisher"
}
$tag = "v$Version"

Write-Host "== 1/6 构建分发包（Version=$Version）==" -ForegroundColor Cyan
if (-not $SkipBuild) {
    & pwsh -NoProfile -File (Join-Path $PSScriptRoot "package-release.ps1") -Version $Version -Root $Root
    if ($LASTEXITCODE -ne 0) { throw "打包失败" }
} else { Write-Host "  已跳过构建（-SkipBuild）" }

$dist = Join-Path $Root "dist"
if (-not (Test-Path (Join-Path $dist "安装 PPTShapeLibrary.exe"))) { throw "dist 不完整，请先打包" }

Write-Host "== 1.5/6 打单文件安装包（新分发方式：一个 exe）==" -ForegroundColor Cyan
if (-not $SkipBuild) {
    & pwsh -NoProfile -File (Join-Path $PSScriptRoot "build-setup.ps1") -Version $Version -Root $Root
    if ($LASTEXITCODE -ne 0) { throw "单文件安装包构建失败" }
} else { Write-Host "  已跳过构建（-SkipBuild）" }
$setup = Join-Path $dist "PPTShapeLibrary-Setup-$Version.exe"
$setupOk = Test-Path $setup
if ($setupOk) {
    $setupHash = (Get-FileHash $setup -Algorithm SHA256).Hash.ToLowerInvariant()
    Write-Host "  $(Split-Path $setup -Leaf)  $([math]::Round((Get-Item $setup).Length/1MB,2)) MB  sha256=$($setupHash.Substring(0,16))…" -ForegroundColor Green
} else {
    Write-Warning "  没有生成单文件安装包（$setup）；update.json 将不含 setupUrl，客户端会退回老的 zip 流程"
}

Write-Host "== 2/6 校验打包出的版本号一致 ==" -ForegroundColor Cyan
$vsto = Join-Path $dist "App\PPTShapeLibrary.vsto"
$appFiles = Join-Path $dist "App\Application Files"
$dirs = Get-ChildItem $appFiles -Directory | Where-Object { $_.Name -like 'PPTShapeLibrary_*' }
if ($dirs.Count -eq 0) { throw "未找到 Application Files\PPTShapeLibrary_*" }
$bundled = ($dirs[0].Name -replace '^PPTShapeLibrary_', '') -replace '_', '.'
Write-Host "  分发包内版本: $bundled"
if ($bundled -ne $Version) {
    Write-Warning "  分发包内版本($bundled) 与 -Version($Version) 不一致；更新器读的是**目录名**，请确认 package-release 是否用了同一个 -Version"
}

Write-Host "== 3/6 制作 zip 与 sha256 ==" -ForegroundColor Cyan
$zip = Join-Path $Root "PPTShapeLibrary-$Version-dist.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
# 注意：dist 里现在还有**单文件安装包**（PPTShapeLibrary-Setup-*.exe）和它的 .cer/.sha256。
# 它们本身就是独立资产，不能打进 zip —— 否则 zip 会被撑大一倍以上（实测 4MB → 16MB）。
$zipItems = Get-ChildItem -Path $dist -Force | Where-Object {
    $_.Name -notlike 'PPTShapeLibrary-Setup-*.exe' -and
    $_.Name -notlike 'PPTShapeLibrary-Setup-*.exe.sha256' -and
    $_.Name -notlike 'PPTShapeLibrary.cer'
}
Compress-Archive -Path ($zipItems | ForEach-Object { $_.FullName }) -DestinationPath $zip -CompressionLevel Optimal
$hash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLowerInvariant()
$shaFile = "$zip.sha256"
"$hash  $(Split-Path $zip -Leaf)" | Set-Content $shaFile -Encoding ASCII
$sizeMB = [math]::Round((Get-Item $zip).Length / 1MB, 2)
Write-Host "  $([IO.Path]::GetFileName($zip))  $sizeMB MB" -ForegroundColor Green
Write-Host "  sha256 = $hash"

Write-Host "== 4/6 生成 update.json ==" -ForegroundColor Cyan
$githubBase = "https://github.com/$owner/$name"
$updateJson = Join-Path $Root "update.json"
$manifest = [ordered]@{
    version     = $Version
    tag         = $tag
    publishedAt = (Get-Date).ToString("o")
    assetName   = (Split-Path $zip -Leaf)
    url         = "$githubBase/releases/download/$tag/$(Split-Path $zip -Leaf)"
    sha256Url   = "$githubBase/releases/download/$tag/$(Split-Path $shaFile -Leaf)"
    sha256      = $hash
    sizeBytes   = (Get-Item $zip).Length
    releaseUrl  = "$githubBase/releases/tag/$tag"
    notes       = if ($Notes) { $Notes } else { "见 Release 说明" }
}
# 新流程主资产：单文件安装包。新版客户端读 setupUrl → 下载 exe → 静默本地安装（**不用关 PowerPoint**）。
# 同时保留上面的 zip/url，让还在用老版 ClickOnce 流程的客户端仍能升级（向后兼容）。
if ($setupOk) {
    $manifest.setupName     = Split-Path $setup -Leaf
    $manifest.setupUrl      = "$githubBase/releases/download/$tag/$(Split-Path $setup -Leaf)"
    $manifest.setupSha256   = $setupHash
    $manifest.setupSizeBytes = (Get-Item $setup).Length
}
($manifest | ConvertTo-Json -Depth 5) | Set-Content $updateJson -Encoding UTF8

# 关键：所有资产的哈希必须**同源**。
# 实测踩过：上传失败后重跑脚本，zip 被重新压缩（字节变了）但另两个没重新生成 → 三者不一致，
# 用户下载后校验失败。所以每次上传前都**用当前文件重算**哈希并同步改写另两个文件。
$hash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLowerInvariant()
"$hash  $(Split-Path $zip -Leaf)" | Set-Content $shaFile -Encoding ASCII
$manifest.sha256 = $hash
$manifest.sizeBytes = (Get-Item $zip).Length
if ($setupOk) {
    $setupHash = (Get-FileHash $setup -Algorithm SHA256).Hash.ToLowerInvariant()
    $manifest.setupSha256 = $setupHash
    $manifest.setupSizeBytes = (Get-Item $setup).Length
}
($manifest | ConvertTo-Json -Depth 5) | Set-Content $updateJson -Encoding UTF8
$syncMsg = "  已同步 .sha256 与 update.json（zip=" + $hash
if ($setupOk) { $syncMsg += "，setup=" + $setupHash.Substring(0, 16) + "…" }
Write-Host ($syncMsg + "）")
Write-Host "  已生成 update.json（更新器的固定地址会取最新 Release 里的这一份）"

Write-Host "== 5/6 打 tag 并推送（可选）==" -ForegroundColor Cyan
Push-Location $Root
try {
    $hasTag = git tag --list $tag
    if (-not $hasTag) { git tag -a $tag -m "Release $Version" }
    if ($SkipPush) { Write-Host "  已跳过推送（-SkipPush）；本地 tag：$tag" }
    elseif ($Proxy) {
        $pushUrl = "https://x-access-token:$token@github.com/$owner/$name.git"
        git -c "http.proxy=$Proxy" push $pushUrl $tag 2>&1 | ForEach-Object { "  $_" }
    } else {
        git push origin $tag 2>&1 | ForEach-Object { "  $_" }
    }
} finally { Pop-Location }

Write-Host "== 6/6 建 Release 并上传资产 ==" -ForegroundColor Cyan
$releaseBody = if ($Notes) { $Notes } else { "自动发布 $Version" }
$payload = @{
    tag_name         = $tag
    name             = "PPTShapeLibrary $Version"
    body             = $releaseBody
    draft            = $false
    prerelease       = $false
} | ConvertTo-Json

$release = $null
# 按 tag 找 Release 必须重试：网络抖动时 GitHub 会返回 404，若直接当成"没有"就会**建出重复 Release**
# —— 实测踩过：同一个 tag 出现两个 Release，releases/latest 指到了缺 update.json 的副本上。
function Get-ReleaseByTag([string]$t) {
    for ($i = 1; $i -le 4; $i++) {
        try { return Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$name/releases/tags/$t" -Headers $headers -TimeoutSec 30 }
        catch { Start-Sleep -Seconds 3 }
    }
    return $null
}

$release = Get-ReleaseByTag $tag
if (-not $release) {
    $all = @()
    try { $all = @(Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$name/releases?per_page=100" -Headers $headers -TimeoutSec 30) } catch { }
    $same = @($all | Where-Object { $_.tag_name -eq $tag })
    if ($same.Count -gt 1) {
        Write-Warning "  发现 $($same.Count) 个同 tag 的 Release（重复），保留资产最多的一个，其余删除"
        $keep = $same | Sort-Object { $_.assets.Count } -Descending | Select-Object -First 1
        foreach ($dup in ($same | Where-Object { $_.id -ne $keep.id })) {
            try {
                Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$name/releases/$($dup.id)" -Method Delete -Headers $headers -TimeoutSec 30 | Out-Null
                Write-Host "    已删除重复 Release id=$($dup.id)（资产 $($dup.assets.Count) 个）" -ForegroundColor Yellow
            } catch { Write-Warning "    删除重复 Release 失败：$($_.Exception.Message)" }
        }
        $release = $keep
    } elseif ($same.Count -eq 1) {
        $release = $same[0]
    }
}
if ($release) {
    Write-Host "  Release 已存在（id=$($release.id)，资产 $($release.assets.Count) 个），改为上传/覆盖资产"
} else {
    try {
        $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$name/releases" -Method Post `
            -Headers $headers -Body $payload -ContentType "application/json" -TimeoutSec 60
        Write-Host "  已创建 Release: $($release.html_url)" -ForegroundColor Green
    } catch {
        Write-Warning "  创建 Release 失败（tag 可能还没推到远端）：$($_.Exception.Message)"
    }
}

# 关键：draft 状态的 Release **不会**被 releases/latest 返回（更新器的固定地址就是 latest/download/…），
# 于是用户端更新会静默失效。实测踩过：强制删掉带 Release 的 tag 再重建，GitHub 会把 Release 转成草稿。
if ($release -and ($release.draft -or $release.prerelease)) {
    try {
        $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$name/releases/$($release.id)" -Method Patch `
            -Headers $headers -Body (@{ draft = $false; prerelease = $false } | ConvertTo-Json) `
            -ContentType "application/json" -TimeoutSec 30
        Write-Host "  已把 Release 从 draft/prerelease 转为正式发布（否则 releases/latest 取不到）" -ForegroundColor Yellow
    } catch { Write-Warning "  取消草稿失败：$($_.Exception.Message)（更新检查会取不到清单，请到网页上手动 Publish release）" }
}

function Upload-Asset($releaseId, $path) {
    $fileName = Split-Path $path -Leaf
    # 已存在同名资产先删掉（幂等）
    $assets = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$name/releases/$releaseId/assets" -Headers $headers -TimeoutSec 30
    foreach ($a in $assets) { if ($a.name -eq $fileName) {
        Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$name/releases/assets/$($a.id)" -Method Delete -Headers $headers -TimeoutSec 30
        Write-Host "    已删除旧资产 $fileName" -ForegroundColor Yellow
    } }
    $uploadUrl = "https://uploads.github.com/repos/$owner/$name/releases/$releaseId/assets?name=$([uri]::EscapeDataString($fileName))"
    $bytes = [IO.File]::ReadAllBytes($path)
    $r = Invoke-RestMethod -Uri $uploadUrl -Method Post -Headers $headers -ContentType "application/octet-stream" -Body $bytes -TimeoutSec 600
    Write-Host "    已上传 $fileName（$([math]::Round($r.size/1MB,2)) MB）" -ForegroundColor Green
}

if ($release) {
    $assetsToUpload = @($zip, $shaFile, $updateJson)
    if ($setupOk) { $assetsToUpload += $setup; $assetsToUpload += "$setup.sha256" }
    foreach ($f in $assetsToUpload) { Upload-Asset $release.id $f }
} else {
    Write-Warning "  未拿到 Release 对象，跳过上传。请在网络可用时重跑本脚本（构建结果已就绪）。"
}

Write-Host ""
Write-Host "发版完成：$tag" -ForegroundColor Green
Write-Host "  发布页    $githubBase/releases/tag/$tag"
Write-Host "  更新清单  $githubBase/releases/latest/download/update.json"
if ($setupOk) {
    Write-Host "  单文件安装包（发给别人就发这一个）：" -ForegroundColor Green
    Write-Host "    $setup"
    Write-Host "    下载直链 $githubBase/releases/download/$tag/$(Split-Path $setup -Leaf)"
    Write-Host "  受信任证书（可选，消除『未知发布者』）：$dist\PPTShapeLibrary.cer"
} else {
    Write-Host "  提示      本次没有单文件安装包；dist\ 整个目录发给别人，双击「安装 PPTShapeLibrary.exe」"
}
