# L的工具箱 —— 安装 / 卸载（当前用户，无需管理员，不碰宏安全设置）
#
# 用法：
#   解压发布包后在本目录运行：
#     powershell -ExecutionPolicy Bypass -File install.ps1 -FromDir .
#   或联网一键安装（下载最新 release 的 zip）：
#     powershell -ExecutionPolicy Bypass -File install.ps1
#   卸载：
#     powershell -ExecutionPolicy Bypass -File uninstall.ps1
[CmdletBinding()]
param(
    [string]$FromDir,
    [string]$ZipUrl = "https://github.com/LLLin000/LToolbox/releases/latest/download/LToolbox-win64.zip",
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'

$LocalDir = Join-Path $env:LOCALAPPDATA 'LToolbox'
$AddinsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
$PpamDst = Join-Path $AddinsDir 'LToolbox.ppam'
$ExeDst = Join-Path $LocalDir 'LToolbox.exe'
$RegPath = 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\AddIns\LToolbox'
$LegacyNames = @('L的工具箱.ppam')

function Remove-Legacy {
    foreach ($n in $LegacyNames) {
        $p = Join-Path $AddinsDir $n
        if (Test-Path $p) { Remove-Item $p -Force; Write-Host "  删除旧包 $n" }
        $k = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\AddIns\$([IO.Path]::GetFileNameWithoutExtension($n))"
        if (Test-Path $k) { Remove-Item $k -Recurse -Force; Write-Host "  注销旧注册项 $k" }
    }
}

function Get-SourceDir {
    if ($Uninstall) { return $null }
    if ($FromDir) {
        if (-not (Test-Path $FromDir)) { throw "目录不存在: $FromDir" }
        return (Resolve-Path $FromDir).Path
    }
    if ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot 'LToolbox.exe'))) {
        return $PSScriptRoot
    }
    Write-Host "[*] 从 $ZipUrl 下载发布包 ..."
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("LToolbox_" + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp | Out-Null
    $zip = Join-Path $tmp 'LToolbox.zip'
    Invoke-WebRequest -Uri $ZipUrl -OutFile $zip -UseBasicParsing
    Expand-Archive -Path $zip -DestinationPath $tmp -Force
    return $tmp
}

if ($Uninstall) {
    Remove-Legacy
    if (Test-Path $RegPath) { Remove-Item $RegPath -Recurse -Force; Write-Host "[1/3] 已注销注册表项" }
    foreach ($p in @($PpamDst, $ExeDst)) {
        if (Test-Path $p) { Remove-Item $p -Force; Write-Host "[2/3] 已删除 $p" }
    }
    if (Test-Path $LocalDir) { Remove-Item $LocalDir -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Host "[3/3] 卸载完成。请重启 PowerPoint。"
    exit 0
}

$src = Get-SourceDir
$exeSrc = Join-Path $src 'LToolbox.exe'
$ppamSrc = Join-Path $src 'LToolbox.ppam'
foreach ($p in @($exeSrc, $ppamSrc)) {
    if (-not (Test-Path $p)) { throw "发布包缺少 $p" }
}
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    Write-Host "[x] PowerPoint 正在运行，加载项文件被锁定，复制会失败。"
    Write-Host "    请完全退出 PowerPoint（含后台窗口）后重新运行本脚本。"
    exit 1
}

New-Item -ItemType Directory -Path $LocalDir -Force | Out-Null
New-Item -ItemType Directory -Path $AddinsDir -Force | Out-Null
Remove-Legacy
Copy-Item $exeSrc $ExeDst -Force
Write-Host "[1/3] exe  -> $ExeDst"
Copy-Item $ppamSrc $PpamDst -Force
Write-Host "[2/3] ppam -> $PpamDst"

New-Item -Path $RegPath -Force | Out-Null
Set-ItemProperty -Path $RegPath -Name 'Path' -Value $PpamDst
Set-ItemProperty -Path $RegPath -Name 'Title' -Value 'L的工具箱'
Set-ItemProperty -Path $RegPath -Name 'Description' -Value '图片标定 / 比例尺 / 原图像素裁取'
Set-ItemProperty -Path $RegPath -Name 'AutoLoad' -Value 0xFFFFFFFF -Type DWord
Write-Host "[3/3] 已注册自动加载（HKCU，无需管理员）"

Write-Host ""
Write-Host "安装完成：完全关闭并重启 PowerPoint 后出现“L的工具箱”选项卡。"
Write-Host "若被安全策略拦截：文件 > 选项 > 信任中心 > 信任中心设置 > 加载项，勾选允许。"
