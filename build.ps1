# build.ps1 — 编译 NewFolderWithSelection.exe（无需 Visual Studio / .NET SDK）
# 使用系统自带的 .NET Framework 4.x csc.exe

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$csc  = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (-not (Test-Path $csc)) { throw "找不到 csc.exe（需要 .NET Framework 4.x）" }

$fw = Split-Path -Parent $csc
$src  = Join-Path $here 'NewFolderWithSelection.cs'
$out  = Join-Path $here 'NewFolderWithSelection.exe'
$mf   = Join-Path $here 'app.manifest'

if (-not (Test-Path $src)) { throw "找不到 $src" }

$args = @(
    '/nologo'
    '/target:winexe'
    '/optimize+'
    '/platform:anycpu'
    '/warn:4'
    ('/win32manifest:' + $mf)
    ('/out:' + $out)
    ('/r:' + (Join-Path $fw 'System.dll'))
    ('/r:' + (Join-Path $fw 'System.Core.dll'))
    ('/r:' + (Join-Path $fw 'System.Drawing.dll'))
    ('/r:' + (Join-Path $fw 'System.Windows.Forms.dll'))
    ('/r:' + (Join-Path $fw 'Microsoft.CSharp.dll'))
    $src
)

& $csc $args
if ($LASTEXITCODE -ne 0) { throw "编译失败，退出码 $LASTEXITCODE" }

$item = Get-Item $out
Write-Host ("编译成功：{0}  ({1:N0} 字节)" -f $item.FullName, $item.Length)
