# capture-window.ps1 — 真实截图（DPI 感知，150% 下也是原生像素）
# 用法：
#   powershell -File .\capture-window.ps1 -Dir "$env:TEMP\pics" -Count 5 -Out preview.png
# 抓完自检输入框区域：如果一行字都没画出来（抓屏时机太早 / 被别的窗口盖住），自动重抓。
param(
    [string]$Out = 'preview-light.png',
    [string]$Dir = '',
    [int]$Count = 5,
    [string]$Title = '新建文件夹并移入'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Cap {
    [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr ctx);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindow(IntPtr c, string n);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
}
"@
# -4 = DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
[void][Cap]::SetProcessDpiAwarenessContext([IntPtr](-4))

$toolDir = Split-Path -Parent $PSScriptRoot
$exe = Join-Path $toolDir 'NewFolderWithSelection.exe'

# 参数自己拼：powershell -File 传数组只会取到第一个元素
if ($Dir -eq '') { $Dir = Join-Path $env:TEMP 'opencode\preview-src' }
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
$files = @(Get-ChildItem -LiteralPath $Dir -File | Select-Object -First $Count | ForEach-Object { $_.FullName })
if ($files.Count -eq 0) { throw "目录里没有文件：$Dir" }
$argList = @($files | ForEach-Object { '"{0}"' -f $_ })

$env:NFS_QUIET = $null
$env:NFS_NAME = $null
$p = Start-Process -FilePath $exe -ArgumentList $argList -PassThru

$h = [IntPtr]::Zero
for ($i = 0; $i -lt 300; $i++) {
    $h = [Cap]::FindWindow([IntPtr]::Zero, $Title)
    if ($h -ne [IntPtr]::Zero -and [Cap]::IsWindowVisible($h)) { break }
    if ($p.HasExited) { throw "进程提前退出 exit=$($p.ExitCode)" }
    Start-Sleep -Milliseconds 25
}
if ($h -eq [IntPtr]::Zero -or -not [Cap]::IsWindowVisible($h)) {
    if (-not $p.HasExited) { $p.Kill() }
    throw "窗口没出现（h=$h title='$Title'）"
}

Start-Sleep -Milliseconds 500
[void][Cap]::SetForegroundWindow($h)
Start-Sleep -Milliseconds 400

$r = New-Object Cap+RECT
[void][Cap]::GetWindowRect($h, [ref]$r)
$w = $r.R - $r.L; $ht = $r.B - $r.T

# 输入框内部区域（逻辑 24..436 / 77..115，这里取内缩一圈的物理像素）
function Ink([System.Drawing.Bitmap]$bmp) {
    $x0 = [int]($w * 0.07); $x1 = [int]($w * 0.92)
    $y0 = [int]($ht * 0.40); $y1 = [int]($ht * 0.50)
    $n = 0
    for ($y = $y0; $y -lt $y1; $y++) {
        for ($x = $x0; $x -lt $x1; $x++) {
            $c = $bmp.GetPixel($x, $y)
            $lum = (0.2126 * $c.R + 0.7152 * $c.G + 0.0722 * $c.B) / 255
            if ($lum -lt 0.6 -or (($c.B - $c.R) -gt 40 -and $lum -lt 0.85)) { $n++ }
        }
    }
    return $n
}

$ink = 0
for ($try = 1; $try -le 12; $try++) {
    $bmp = New-Object System.Drawing.Bitmap $w, $ht
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.L, $r.T, 0, 0, (New-Object System.Drawing.Size $w, $ht))
    $ink = Ink $bmp
    if ($ink -ge 800) {
        $bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose()
        break
    }
    $g.Dispose(); $bmp.Dispose()
    [void][Cap]::SetForegroundWindow($h)
    Start-Sleep -Milliseconds 250
}
if ($ink -lt 800) {
    $p.Kill()
    throw "输入框区域没画出内容（ink=$ink），抓屏失败"
}
Write-Host "截图: $Out  ${w}x${ht}  ink=$ink  (文件 $($files.Count) 个)"

$p.CloseMainWindow() | Out-Null
Start-Sleep -Milliseconds 400
if (-not $p.HasExited) { $p.Kill() }
