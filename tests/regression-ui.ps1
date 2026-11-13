# regression-ui.ps1 — 真实窗口测试（默认跳过）
#
# 安全约定（2026-11-13 起强制）：
#   * 禁止 SendKeys / WScript.Shell.SendKeys / keybd_event / SetForegroundWindow + 模拟输入 /
#     剪贴板。曾出现过按键落到用户前台窗口（Trae 输入框）的事故。
#   * 所有按键只 PostMessage 到“校验过的”被测窗口句柄：句柄必须存在、属于被测进程、
#     标题完全匹配；校验不过就跳过该项并记为 SKIP，不做前台注入。
#   * 脚本会在桌面弹出真实窗口，所以默认不跑；要跑必须显式设 NFS_ALLOW_UI_TESTS=1。

if ($env:NFS_ALLOW_UI_TESTS -ne '1') {
    Write-Host 'SKIP regression-ui：这个脚本会在桌面弹出窗口；且前台注入已被禁用。'
    Write-Host '  需要实窗验证时：$env:NFS_ALLOW_UI_TESTS=1，再运行本脚本。'
    Write-Host 'PASS 0 / FAIL 0（跳过）'
    exit 0
}

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8

$dir  = Split-Path -Parent $PSScriptRoot
$exe  = Join-Path $dir 'NewFolderWithSelection.exe'
$qdir = Join-Path $env:TEMP 'NewFolderWithSelection'
$root = Join-Path $env:TEMP 'opencode\nfs-ui'
$TITLE = '新建文件夹并移入'

if (-not (Test-Path $exe)) { Write-Host "找不到 $exe"; exit 1 }

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public class W {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindow(IntPtr cls, string name);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string name);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern int GetWindowRgn(IntPtr h, IntPtr hrgn);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
    [DllImport("gdi32.dll")] public static extern IntPtr CreateRectRgn(int a, int b, int c, int d);
    [DllImport("gdi32.dll")] public static extern bool PtInRegion(IntPtr hrgn, int x, int y);
    [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr o);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
}
"@

$script:pass = 0
$script:fail = 0
$script:skip = 0
function Check($name, $cond, $detail) {
    if ($cond) { $script:pass++; Write-Host ("  PASS  {0}" -f $name) }
    else       { $script:fail++; Write-Host ("  FAIL  {0}   -> {1}" -f $name, $detail) }
}
function Skip($name, $why) {
    $script:skip++
    Write-Host ("  SKIP  {0}   -> {1}" -f $name, $why)
}
function Reset-All {
    Remove-Item -Recurse -Force $qdir -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $root -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $env:NFS_QUIET = $null; $env:NFS_NAME = $null; $env:NFS_NOMOVE = $null; $env:NFS_NOELEVATE = $null
}
function Wait-Window([int]$timeoutMs = 6000) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $timeoutMs) {
        $h = [W]::FindWindow([IntPtr]::Zero, $TITLE)
        if ($h -ne [IntPtr]::Zero -and [W]::IsWindowVisible($h)) { return @{ h = $h; ms = $sw.ElapsedMilliseconds } }
        Start-Sleep -Milliseconds 25
    }
    return @{ h = [IntPtr]::Zero; ms = -1 }
}

# 句柄校验：存在 + 属于被测进程 + 标题一致。任何一条不满足都不往里发消息。
function Test-Ours([IntPtr]$h, [int]$procId) {
    if ($h -eq [IntPtr]::Zero -or -not [W]::IsWindow($h)) { return $false }
    $owner = 0
    [void][W]::GetWindowThreadProcessId($h, [ref]$owner)
    if ($owner -ne $procId) { return $false }
    $sb = New-Object System.Text.StringBuilder 256
    [void][W]::GetWindowTextW($h, $sb, 256)
    return ($sb.ToString() -ceq $TITLE)
}

# 找到对话框里的真实 TextBox（WinForms 的类名形如 WindowsForms10.EDIT.app.0.xxx）
function Get-EditHandle([IntPtr]$parent) {
    $h = [IntPtr]::Zero
    for ($i = 0; $i -lt 32; $i++) {
        $h = [W]::FindWindowEx($parent, $h, $null, $null)
        if ($h -eq [IntPtr]::Zero) { return [IntPtr]::Zero }
        $cls = New-Object System.Text.StringBuilder 256
        [void][W]::GetClassNameW($h, $cls, 256)
        if ($cls.ToString() -like '*EDIT*') { return $h }
    }
    return [IntPtr]::Zero
}

# 只发消息，不抢焦点：WM_CHAR 直接投进输入框句柄
function Send-Text([IntPtr]$edit, [string]$text) {
    foreach ($ch in $text.ToCharArray()) {
        [void][W]::PostMessage($edit, 0x0102, [IntPtr][int]$ch, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 15
    }
}
# WM_KEYDOWN/WM_KEYUP 投给对话框（Form 的 ProcessCmdKey 在这个消息循环里生效）
function Send-Key([IntPtr]$h, [int]$vk) {
    [void][W]::PostMessage($h, 0x0100, [IntPtr]$vk, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 30
    [void][W]::PostMessage($h, 0x0101, [IntPtr]$vk, [IntPtr]::Zero)
}
function Close-Ours([IntPtr]$h) {
    if ($h -ne [IntPtr]::Zero) { [void][W]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) }
}

$VK_RETURN = 13
$VK_ESCAPE = 27

# =====================================================================
Write-Host "`n=== A. 窗口出现速度 + 几何 ==="
Reset-All
$src = Join-Path $root 'A'
New-Item -ItemType Directory -Force -Path $src | Out-Null
foreach ($n in '图一.png', '图二.png', '图三.png') { 'x' | Set-Content -LiteralPath (Join-Path $src $n) }

$lat = @()
$rect = $null; $rounded = $null; $regionType = $null
for ($i = 0; $i -lt 3; $i++) {
    Reset-All
    New-Item -ItemType Directory -Force -Path $src | Out-Null
    foreach ($n in '图一.png', '图二.png', '图三.png') { 'x' | Set-Content -LiteralPath (Join-Path $src $n) }
    $p = Start-Process -FilePath $exe -ArgumentList @(('"{0}"' -f (Join-Path $src '图一.png')), ('"{0}"' -f (Join-Path $src '图二.png')), ('"{0}"' -f (Join-Path $src '图三.png'))) -PassThru
    $w = Wait-Window
    if ($w.h -eq [IntPtr]::Zero) { Check '第 1 次窗口出现' $false 'timeout'; continue }
    $lat += $w.ms
    if ($i -eq 0) {
        $r = New-Object W+RECT
        [void][W]::GetWindowRect($w.h, [ref]$r)
        $rect = "$($r.R - $r.L)x$($r.B - $r.T)"
        $hr = [W]::CreateRectRgn(0, 0, 0, 0)
        $regionType = [W]::GetWindowRgn($w.h, $hr)
        $corner = [W]::PtInRegion($hr, 2, 2)
        $inside = [W]::PtInRegion($hr, 60, 60)
        [void][W]::DeleteObject($hr)
        $rounded = (-not $corner) -and $inside
    }
    Close-Ours $w.h
    $p.WaitForExit(5000) | Out-Null
    Start-Sleep -Milliseconds 300
}
$avg = if ($lat.Count) { [int](($lat | Measure-Object -Average).Average) } else { -1 }
Write-Host ("  窗口出现耗时: {0} ms（3 次：{1}）" -f $avg, ($lat -join ', '))
Check '窗口 1 秒内出现' ($avg -ge 0 -and $avg -lt 1000) "$avg ms"
Check '窗口尺寸合理（>=400x200 物理像素）' ($rect -match '^(\d+)x(\d+)$' -and [int]$Matches[1] -ge 400 -and [int]$Matches[2] -ge 200) $rect
Check '窗口是圆角（左上角在区域外、内部在区域内）' $rounded "regionType=$regionType"
Check 'Esc/关闭后没建文件夹' (@(Get-ChildItem -LiteralPath $src -Directory).Count -eq 0) 'dir created'

# =====================================================================
Write-Host "`n=== B. 输入框里打字 + 回车（PostMessage，不抢前台） ==="
Reset-All
$src = Join-Path $root 'B'
New-Item -ItemType Directory -Force -Path $src | Out-Null
foreach ($n in 'a.png', 'b.png') { 'x' | Set-Content -LiteralPath (Join-Path $src $n) }
$p = Start-Process -FilePath $exe -ArgumentList @(('"{0}"' -f (Join-Path $src 'a.png')), ('"{0}"' -f (Join-Path $src 'b.png'))) -PassThru
$w = Wait-Window
Check '窗口出现' ($w.h -ne [IntPtr]::Zero) 'timeout'
if ((Test-Ours $w.h $p.Id) -and ((Get-EditHandle $w.h) -ne [IntPtr]::Zero)) {
    $edit = Get-EditHandle $w.h
    Send-Text $edit 'myfolder'      # 默认名已全选，直接替换
    Start-Sleep -Milliseconds 200
    Send-Key $w.h $VK_RETURN
    $p.WaitForExit(15000) | Out-Null
    $dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
    Check '用输入的名字建文件夹' ($dirs.Count -eq 1 -and $dirs[0] -eq 'myfolder') ($dirs -join ' | ')
    Check '2 项都移进去了' (@(Get-ChildItem -LiteralPath (Join-Path $src 'myfolder') -File).Count -eq 2) 'count mismatch'
} else {
    Skip '输入框打字 + 回车' '句柄/标题/进程校验不通过，不发任何按键'
    Close-Ours $w.h
    $p.WaitForExit(5000) | Out-Null
}

# =====================================================================
Write-Host "`n=== C. Esc 取消 ==="
Reset-All
$src = Join-Path $root 'C'
New-Item -ItemType Directory -Force -Path $src | Out-Null
'x' | Set-Content -LiteralPath (Join-Path $src 'c.png')
$p = Start-Process -FilePath $exe -ArgumentList ('"{0}"' -f (Join-Path $src 'c.png')) -PassThru
$w = Wait-Window
if (Test-Ours $w.h $p.Id) {
    Send-Key $w.h $VK_ESCAPE
    $p.WaitForExit(15000) | Out-Null
    Check 'Esc 后没有建文件夹' (@(Get-ChildItem -LiteralPath $src -Directory).Count -eq 0) 'dir created'
    Check 'Esc 后文件还在原处' (Test-Path -LiteralPath (Join-Path $src 'c.png')) 'file gone'
} else {
    Skip 'Esc 取消' '句柄/标题/进程校验不通过，不发任何按键'
    Close-Ours $w.h
    $p.WaitForExit(5000) | Out-Null
}

# =====================================================================
Write-Host "`n=== D. 空名字：回车不关窗、不建文件夹 ==="
Reset-All
$src = Join-Path $root 'D'
New-Item -ItemType Directory -Force -Path $src | Out-Null
'x' | Set-Content -LiteralPath (Join-Path $src 'd.png')
$p = Start-Process -FilePath $exe -ArgumentList ('"{0}"' -f (Join-Path $src 'd.png')) -PassThru
$w = Wait-Window
if ((Test-Ours $w.h $p.Id) -and ((Get-EditHandle $w.h) -ne [IntPtr]::Zero)) {
    $edit = Get-EditHandle $w.h
    Send-Text $edit '   '
    Start-Sleep -Milliseconds 150
    Send-Key $w.h $VK_RETURN
    Start-Sleep -Milliseconds 600
    $stillOpen = [W]::IsWindow($w.h) -and [W]::IsWindowVisible($w.h)
    Check '空名字时窗口没关（还在等输入）' $stillOpen 'window closed'
    Send-Key $w.h $VK_ESCAPE
    $p.WaitForExit(15000) | Out-Null
    Check '空名字最终没建文件夹' (@(Get-ChildItem -LiteralPath $src -Directory).Count -eq 0) 'dir created'
} else {
    Skip '空名字不关窗' '句柄/标题/进程校验不通过，不发任何按键'
    Close-Ours $w.h
    $p.WaitForExit(5000) | Out-Null
}

# =====================================================================
Write-Host "`n=== E. 窗口开着时，后续项实时并入同一批 ==="
Reset-All
$src = Join-Path $root 'E'
New-Item -ItemType Directory -Force -Path $src | Out-Null
foreach ($n in 'e1.png', 'e2.png', 'e3.png', 'e4.png') { 'x' | Set-Content -LiteralPath (Join-Path $src $n) }
$p = Start-Process -FilePath $exe -ArgumentList ('"{0}"' -f (Join-Path $src 'e1.png')) -PassThru
$w = Wait-Window
Check '窗口出现' ($w.h -ne [IntPtr]::Zero) 'timeout'
# 模拟 Explorer 在窗口打开期间继续逐项调用
foreach ($n in 'e2.png', 'e3.png', 'e4.png') {
    Start-Process -FilePath $exe -ArgumentList ('"{0}"' -f (Join-Path $src $n)) -WindowStyle Hidden | Out-Null
    Start-Sleep -Milliseconds 150
}
Start-Sleep -Milliseconds 900
$log = ''
if (Test-Path (Join-Path $qdir 'last-run.log')) { $log = Get-Content (Join-Path $qdir 'last-run.log') -Raw -Encoding UTF8 }
Check '副标题计数实时涨到 4（窗口没关）' ($log -match '批内计数 4') 'log 里没有 批内计数 4'
if ((Test-Ours $w.h $p.Id) -and ((Get-EditHandle $w.h) -ne [IntPtr]::Zero)) {
    $edit = Get-EditHandle $w.h
    Send-Text $edit 'livebatch'
    Start-Sleep -Milliseconds 150
    Send-Key $w.h $VK_RETURN
    $p.WaitForExit(15000) | Out-Null
    $dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
    Check '4 项并进同一个文件夹' ($dirs.Count -eq 1 -and $dirs[0] -eq 'livebatch') ($dirs -join ' | ')
    Check '4 项都移入' (@(Get-ChildItem -LiteralPath (Join-Path $src 'livebatch') -File -ErrorAction SilentlyContinue).Count -eq 4) 'count mismatch'
} else {
    Skip '窗口开着时实时并入' '句柄/标题/进程校验不通过，不发任何按键'
    Close-Ours $w.h
    $p.WaitForExit(5000) | Out-Null
}

Write-Host "`n=== 汇总 ==="
Write-Host ("PASS {0} / FAIL {1} / SKIP {2}" -f $script:pass, $script:fail, $script:skip)
Remove-Item -Recurse -Force $root -ErrorAction SilentlyContinue
if ($script:fail -eq 0) { exit 0 } else { exit 1 }
