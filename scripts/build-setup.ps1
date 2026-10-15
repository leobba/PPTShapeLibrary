# build-setup.ps1 —— 打一个**单文件 exe** 安装包（方案 B：本地安装，不用 ClickOnce）
#
# 产物：dist\PPTShapeLibrary-Setup-<版本>.exe
#   · 内嵌插件载荷（Release 构建输出，真实文件名，扁平 .vsto 布局）
#   · 内嵌 VSTO Runtime 离线包（离线安装兜底）
#   · 自签代码签名（用户需信任证书才能看到发布者名；未信任则显示"未知发布者"）
#
# 安装行为（由安装器 LocalInstall 实现）：
#   解压到 %LOCALAPPDATA%\PPTShapeLibrary\addin\<版本>\
#   + 写 2 个加载项注册键（Manifest=file:///…|vstolocal）
#   + 写控制面板卸载项
#   **不需要 ClickOnce、不需要 VSTOInstaller、不需要关闭 PowerPoint**
#
# 用法：
#   pwsh -File scripts\build-setup.ps1 -Version 1.2.0.0
#   pwsh -File scripts\build-setup.ps1 -Version 1.2.0.0 -SkipSign
param(
    [Parameter(Mandatory = $true)][string]$Version,
    [string]$Root = "E:\code\PPTShapeLibrary",
    [switch]$SkipSign,
    [string]$TimestampUrl = "http://timestamp.digicert.com",
    [string]$CertThumbprint = "32BE9366A69E1D726F6AAAEE4C6664A094590B39",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
function Step($n, $t) { Write-Host "== $n $t ==" -ForegroundColor Cyan }
function Ok($m) { Write-Host "  $m" -ForegroundColor Green }
function Warn($m) { Write-Host "  $m" -ForegroundColor Yellow }

$msb = "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $msb)) { throw "找不到 MSBuild：$msb" }

$stage = Join-Path $Root "build\setup-stage"
$payload = Join-Path $Root "installer\payload.zip"
$outDir = Join-Path $Root "dist"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# ---------------------------------------------------------------- 1/7 版本对齐
Step "1/7" "把 AssemblyInfo 版本对齐到 $Version"
$asmInfo = Join-Path $Root "PPTShapeLibrary\Properties\AssemblyInfo.cs"
$changed = $false
$lines = Get-Content $asmInfo -Encoding UTF8
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match '^\[assembly:\s*AssemblyVersion\("([\d\.]+)"\)\]') {
        if ($Matches[1] -ne $Version) { $lines[$i] = '[assembly: AssemblyVersion("' + $Version + '")]'; $changed = $true }
    }
    elseif ($lines[$i] -match '^\[assembly:\s*AssemblyFileVersion\("([\d\.]+)"\)\]') {
        if ($Matches[1] -ne $Version) { $lines[$i] = '[assembly: AssemblyFileVersion("' + $Version + '")]'; $changed = $true }
    }
}
if ($changed) { Set-Content $asmInfo -Value $lines -Encoding UTF8; Ok "AssemblyInfo → $Version" } else { Ok "AssemblyInfo 已是 $Version" }

# ---------------------------------------------------------------- 2/7 构建插件
Step "2/7" "构建插件 Release（ApplicationVersion=$Version）"
# 关键：.vsto 的版本来自 ApplicationVersion 而不是 AssemblyInfo；
# 且增量构建**不会**重新生成已存在的 .vsto → 先删掉强制再生，否则会打进旧版本号。
foreach ($f in 'PPTShapeLibrary.vsto', 'PPTShapeLibrary.dll.manifest') {
    Remove-Item (Join-Path $Root "PPTShapeLibrary\bin\Release\$f") -Force -ErrorAction SilentlyContinue
}
if (-not $SkipBuild) {
    & $msb (Join-Path $Root "PPTShapeLibrary\PPTShapeLibrary.csproj") /t:Build /p:Configuration=Release `
        /p:ApplicationVersion=$Version /nologo /v:m
    if ($LASTEXITCODE -ne 0) { throw "插件构建失败" }
}
$relDir = Join-Path $Root "PPTShapeLibrary\bin\Release"
$vsto = Join-Path $relDir "PPTShapeLibrary.vsto"
if (-not (Test-Path $vsto)) { throw "没找到 $vsto（构建未生成部署清单）" }
# 取「依赖程序集 PPTShapeLibrary.dll」的版本（清单自身的 <assemblyIdentity> 恒为 version="1.0"，不能拿来判断）
$vstoVer = ([regex]::Match((Get-Content $vsto -Raw),
    'assemblyIdentity\s+name="PPTShapeLibrary\.dll"\s+version="([\d\.]+)"')).Groups[1].Value
Ok "bin\Release 就绪；.vsto 里插件版本=$vstoVer"
if ($vstoVer -ne $Version) { Warn "！ .vsto 版本($vstoVer) 与目标版本($Version) 不一致，请检查（可能增量构建没重生成）" }

# ---------------------------------------------------------------- 3/7 组装载荷
Step "3/7" "组装载荷（插件文件 + VSTO 离线包）"
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null

# 只收插件运行需要的文件；排除 app.publish（ClickOnce 产物）、pdb、xml 等
$excludeDir = @('app.publish')
$skipExt = @('.pdb', '.xml', '.cache', '.log', '.vshost.exe', '.vshost.pdb')
$copied = 0
Get-ChildItem $relDir -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($relDir.Length).TrimStart('\')
    if ($excludeDir | Where-Object { $rel -like "$_\*" -or $rel -eq $_ }) { return }
    if ($skipExt -contains $_.Extension.ToLower()) { return }
    if ($_.Name -like '*.vshost.*') { return }
    $dst = Join-Path $stage $rel
    New-Item -ItemType Directory -Path (Split-Path $dst -Parent) -Force | Out-Null
    Copy-Item $_.FullName $dst -Force
    $script:copied++
}
Ok "插件文件 $copied 个"
foreach ($need in 'PPTShapeLibrary.vsto', 'PPTShapeLibrary.dll', 'PPTShapeLibrary.dll.manifest', 'WebUI\index.html') {
    if (-not (Test-Path (Join-Path $stage $need))) { throw "载荷缺少必需文件：$need" }
}

# VSTO Runtime 离线包（缺就现拉）
$offlineSrc = Join-Path $Root "dist\离线安装包"
$offlineDst = Join-Path $stage "离线安装包"
New-Item -ItemType Directory -Path $offlineDst -Force | Out-Null
$vstor = Join-Path $offlineSrc "vstor_runtime.zip"
if (-not (Test-Path $vstor)) {
    Warn "dist\离线安装包\vstor_runtime.zip 不存在，尝试获取…"
    & pwsh -NoProfile -File (Join-Path $Root "scripts\fetch-prerequisites.ps1") -OutDir $offlineDst -Only vsto
    $vstor = Join-Path $offlineDst "vstor_runtime.zip"
}
if (Test-Path $vstor) { Copy-Item $vstor $offlineDst -Force; Ok "已内置 VSTO 离线包（$([math]::Round((Get-Item $vstor).Length/1KB)) KB）" }
else { Warn "！ 没有 VSTO 离线包：目标机缺 VSTO Runtime 时将只能联网下载" }

# ---------------------------------------------------------------- 4/7 打 zip
Step "4/7" "打成 payload.zip"
if (Test-Path $payload) { Remove-Item $payload -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($stage, $payload, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Ok "payload.zip = $([math]::Round((Get-Item $payload).Length/1KB)) KB"

# ---------------------------------------------------------------- 5/7 构建单 exe
Step "5/7" "构建安装器（载荷已内嵌）"
& $msb (Join-Path $Root "installer\PPTShapeLibraryInstaller.csproj") /t:Rebuild /p:Configuration=Release /nologo /v:m
if ($LASTEXITCODE -ne 0) { throw "安装器构建失败" }
$built = Get-ChildItem (Join-Path $Root "installer\bin\Release") -Filter '*.exe' | Select-Object -First 1
if (-not $built) { throw "没找到构建出的安装器 exe" }
Ok "$($built.Name) = $([math]::Round($built.Length/1MB, 2)) MB"

# 自检：确认载荷真的嵌进去了
$probe = Start-Process -FilePath $built.FullName -ArgumentList '--payload-info' -PassThru -Wait -NoNewWindow -RedirectStandardOutput (Join-Path $env:TEMP 'payload_info.txt')
$info = Get-Content (Join-Path $env:TEMP 'payload_info.txt') -Raw -ErrorAction SilentlyContinue
if ($info -notmatch 'HasPayload=True') { throw "安装器里没有内嵌载荷（--payload-info 返回 $($info -replace '\s+',' ')'）" }
Ok "已确认内嵌载荷"

$setupName = "PPTShapeLibrary-Setup-$Version.exe"
$setupPath = Join-Path $outDir $setupName
Copy-Item $built.FullName $setupPath -Force

# ---------------------------------------------------------------- 6/7 签名
Step "6/7" "代码签名（自签证书）"
$cerPath = Join-Path $outDir "PPTShapeLibrary.cer"
if ($SkipSign) {
    Warn "已跳过签名（-SkipSign）"
} else {
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $CertThumbprint } | Select-Object -First 1
    if (-not $cert) {
        Warn "！ 证书存储里找不到指纹 $CertThumbprint 的代码签名证书，跳过签名"
    } else {
        # 导出公钥证书，供用户导入「受信任的根证书颁发机构」后看到发布者名
        try {
            [IO.File]::WriteAllBytes($cerPath, $cert.Export([Security.Cryptography.X509Certificates.X509ContentType]::Cert))
            Ok "已导出公钥证书：$cerPath"
        } catch { Warn "导出 .cer 失败：$($_.Exception.Message)" }

        $signed = $false
        foreach ($ts in @($TimestampUrl, '')) {
            try {
                $p = @{ FilePath = $setupPath; Certificate = $cert }
                if ($ts) { $p.TimestampServer = $ts }
                Set-AuthenticodeSignature @p | Out-Null
                $sig = Get-AuthenticodeSignature $setupPath
                if ($sig.Status -eq 'Valid') { Ok "签名成功（时间戳=$(if ($ts) { $ts } else { '无' })）状态=$($sig.Status)"; $signed = $true; break }
                Warn "签名状态=$($sig.Status)，重试…"
            } catch { Warn "签名失败：$($_.Exception.Message)" }
        }
        if (-not $signed) { Warn "！ 签名未成功（exe 仍可用，只是会显示"未知发布者"）" }
    }
}

# ---------------------------------------------------------------- 7/7 校验与摘要
Step "7/7" "生成校验文件"
$hash = (Get-FileHash $setupPath -Algorithm SHA256).Hash.ToLower()
Set-Content -Path "$setupPath.sha256" -Value $hash -Encoding ASCII
Ok "sha256 = $hash"

$size = [math]::Round((Get-Item $setupPath).Length / 1MB, 2)
Write-Host ""
Write-Host "单文件安装包已生成" -ForegroundColor Green
Write-Host "  路径   $setupPath"
Write-Host "  体积   $size MB"
Write-Host "  签名   $((Get-AuthenticodeSignature $setupPath).Status)"
Write-Host "  证书   $cerPath （发给用户：先导入为受信任根证书，可消除"未知发布者"）"
Write-Host ""
Write-Host "安装（本机测试）：& '$setupPath'                       → 图形界面"
Write-Host "                  & '$setupPath' --install-local --quiet  → 静默"
Write-Host "卸载：            & '$setupPath' --uninstall"
