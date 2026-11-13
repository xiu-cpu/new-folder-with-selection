# install-shell-menu.ps1 — 实验性：把「新建文件夹并移入」装进 Windows 11 主右键菜单
#
# 当前状态：未启用（缺一次原生编译，详见 shim\NewFolderShellShim.cpp 头部）。
# 经典菜单入口仍由 install.ps1 负责，两者互不影响。
#
# 用法：
#   .\install-shell-menu.ps1                 安装（需要同目录有 NewFolderShellShim.dll）
#   .\install-shell-menu.ps1 -EnableDevMode  安装，并在需要时开启开发者模式（会提权一次）
#   .\install-shell-menu.ps1 -Uninstall      卸载（移除稀疏包，不留注册表/证书残余）
#
# 本机实测要点（Win11 build 26300）：
#   * 稀疏包用 Add-AppxPackage -Register -ExternalLocation 注册，未签名；
#     未开开发者模式时报 0x80073CFF（"需要开发者许可证或启用旁载"），所以必须先开。
#   * 开发者模式是机器级设置（HKLM\...\AppModelUnlock\AllowDevelopmentWithoutDevLicense），
#     卸载脚本不会替你关掉它——若安装前没开过，请按末尾提示手动还原。

param(
    [switch]$Uninstall,
    [switch]$EnableDevMode
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifest = Join-Path $here 'package\AppxManifest.xml'
$pkgName = 'NewFolderWithSelection.ShellExtension'
$shim = Join-Path $here 'NewFolderShellShim.dll'
$exe = Join-Path $here 'NewFolderWithSelection.exe'
$devKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock'

function Get-NfwsPackage {
    Get-AppxPackage -Name $pkgName -ErrorAction SilentlyContinue
}

$markerDir = Join-Path $env:LOCALAPPDATA 'NFWS-ShellMenu'
$marker = Join-Path $markerDir 'devmode-set-by-installer'

if ($Uninstall) {
    $p = Get-NfwsPackage
    if ($p) {
        Remove-AppxPackage -Package $p.PackageFullName
        Write-Host "已移除稀疏包：$($p.PackageFullName)"
    } else {
        Write-Host '没有找到已安装的稀疏包'
    }
    if (Test-Path $marker) {
        Write-Host '安装时开过开发者模式，正在还原（需要管理员确认一次）…'
        $script = @"
if ((Get-ItemProperty '$devKey' -Name AllowDevelopmentWithoutDevLicense -ErrorAction SilentlyContinue)) {
    Remove-ItemProperty '$devKey' -Name AllowDevelopmentWithoutDevLicense -Force
}
if (((Get-Item '$devKey' -ErrorAction SilentlyContinue).ValueCount -eq 0) -and ((Get-Item '$devKey' -ErrorAction SilentlyContinue).SubKeyCount -eq 0)) {
    Remove-Item '$devKey' -Force
}
"@
        $tmp = Join-Path $env:TEMP 'nfws-devmode-revert.ps1'
        Set-Content -LiteralPath $tmp -Value $script -Encoding UTF8
        Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $tmp) -Verb RunAs -Wait
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $marker -Force -ErrorAction SilentlyContinue
        Write-Host '开发者模式已还原'
    } else {
        Write-Host "开发者模式不是本次安装开启的，未改动（如确认不需要，可手动删 $devKey\AllowDevelopmentWithoutDevLicense）"
    }
    if ((Test-Path $markerDir) -and -not (Get-ChildItem $markerDir -Force)) { Remove-Item $markerDir -Force }
    Write-Host '经典菜单入口未受影响（由 uninstall.ps1 单独管理）'
    exit 0
}

# ---- 前置检查：缺任何一样都直接说清楚，不做降级 ----
if (-not (Test-Path $exe)) {
    Write-Host "缺少 $exe，先运行 build.ps1"
    exit 1
}
if (-not (Test-Path $manifest)) {
    Write-Host "缺少包定义 $manifest"
    exit 1
}
if (-not (Test-Path $shim)) {
    Write-Host '缺少原生 COM 适配 DLL：NewFolderShellShim.dll'
    Write-Host 'Win11 主菜单要求条目来自稀疏包，且包内 com:Class 的 Path 必须是原生 DLL'
    Write-Host '（代理宿主直接 LoadLibrary + DllGetClassObject，托管程序集放不进去）。'
    Write-Host '请先用 MSVC 或 MinGW 编译 shim\NewFolderShellShim.cpp，产物放到本目录后再运行。'
    Write-Host '没有任何东西被安装，机器的注册表/包/证书均未改动。'
    exit 2
}

$dev = (Get-ItemProperty -Path $devKey -Name AllowDevelopmentWithoutDevLicense -ErrorAction SilentlyContinue).AllowDevelopmentWithoutDevLicense
if ($dev -ne 1) {
    if (-not $EnableDevMode) {
        Write-Host '开发者模式未开启，稀疏包无法以松散方式注册（会报 0x80073CFF）。'
        Write-Host "重跑并加 -EnableDevMode（会提权一次，写入 $devKey\AllowDevelopmentWithoutDevLicense = 1）"
        exit 3
    }
    Write-Host '正在开启开发者模式（需要管理员确认一次）…'
    $script = @"
New-Item -Path '$devKey' -Force | Out-Null
New-ItemProperty -Path '$devKey' -Name AllowDevelopmentWithoutDevLicense -PropertyType DWord -Value 1 -Force | Out-Null
"@
    $tmp = Join-Path $env:TEMP 'nfws-devmode.ps1'
    Set-Content -LiteralPath $tmp -Value $script -Encoding UTF8
    $p = Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $tmp) -Verb RunAs -PassThru -Wait
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    if ($p.ExitCode -ne 0) { Write-Host "开启开发者模式失败（退出码 $($p.ExitCode)）"; exit 3 }
    New-Item -ItemType Directory -Force -Path $markerDir | Out-Null
    '本次安装开启了开发者模式' | Set-Content -LiteralPath $marker -Encoding UTF8
}

# ---- 注册稀疏包 ----
$existing = Get-NfwsPackage
if ($existing) { Remove-AppxPackage -Package $existing.PackageFullName }
try {
    Add-AppxPackage -Register $manifest -ExternalLocation $here
} catch {
    Write-Host "注册失败：$($_.Exception.Message)"
    Write-Host ('HRESULT: 0x{0:X8}' -f $_.Exception.HResult)
    exit 4
}

$after = Get-NfwsPackage
if (-not $after) { Write-Host '包注册后查不到，安装失败'; exit 4 }
Write-Host "已安装：$($after.PackageFullName)"
Write-Host "外部内容目录：$here"
Write-Host '重启 explorer 后，选中文件/文件夹 → 右键，主菜单里应出现「新建文件夹并移入 (&F)」'
Write-Host '卸载： .\install-shell-menu.ps1 -Uninstall'
