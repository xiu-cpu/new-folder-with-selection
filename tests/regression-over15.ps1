$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8

$dir  = Split-Path -Parent $PSScriptRoot
$tool = Join-Path $dir 'NewFolderWithSelection.exe'
$qdir = Join-Path $env:TEMP 'NewFolderWithSelection'
$test = Join-Path $env:TEMP 'opencode\verb-many'
$N    = 20

# 重新安装（命令里用 exe + %1）
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $dir 'install.ps1')
$cmd = (Get-ItemProperty 'HKCU:\Software\Classes\AllFilesystemObjects\shell\NewFolderWithSelection\command').'(Default)'
Write-Host "注册表命令: $cmd"

# 关掉上次残留的、指向测试目录的 Explorer 窗口（否则会拿到已失效的 Document）
$shell0 = New-Object -ComObject Shell.Application
foreach ($w in @($shell0.Windows())) {
    try { if ($w.Document.Folder.Self.Path -eq $test) { $w.Quit() } } catch {}
}

Remove-Item -Recurse -Force $qdir,$test -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $test | Out-Null
1..$N | ForEach-Object { "x" | Set-Content -LiteralPath (Join-Path $test ("v{0:d2}.txt" -f $_)) }
$all = @(Get-ChildItem -LiteralPath $test -File | ForEach-Object { $_.FullName })

# --- 真实 Explorer 窗口 + 全选 ---
Start-Process explorer.exe -ArgumentList "`"$test`""
$wsh   = New-Object -ComObject WScript.Shell
$shell = New-Object -ComObject Shell.Application
$win = $null
for ($i = 0; $i -lt 20; $i++) {
    Start-Sleep -Milliseconds 400
    foreach ($w in $shell.Windows()) { try { if ($w.Document.Folder.Self.Path -eq $test) { $win = $w; break } } catch {} }
    if ($win) { break }
}
if (-not $win) { Write-Host 'FAIL: 没找到 Explorer 窗口'; exit 1 }
$folder = $shell.Namespace($test)
$null = $wsh.AppActivate($win.HWND); Start-Sleep -Milliseconds 300
$sel = @()
for ($k = 1; $k -le 4; $k++) {
    # 用 Shell COM 直接设定选中（不依赖窗口焦点，比 SendKeys 稳）
    foreach ($p in $all) {
        try { $win.Document.SelectItem($folder.ParseName((Split-Path -Leaf $p)), 1) } catch {}
    }
    Start-Sleep -Milliseconds 600
    $sel = @($win.Document.SelectedItems() | ForEach-Object { $_.Path })
    if ($sel.Count -eq $N) { break }
    $wsh.SendKeys('^a'); Start-Sleep -Milliseconds 600
    $sel = @($win.Document.SelectedItems() | ForEach-Object { $_.Path })
    if ($sel.Count -eq $N) { break }
}
Write-Host "Explorer 选中: $($sel.Count) 项"
if ($sel.Count -ne $N) { Write-Host 'FAIL: 全选失败'; exit 1 }

# --- 用真实注册表菜单项触发，只触发前 15 项（模拟 Explorer 的 15 项上限） ---
# 先把 15 个菜单项都取出来再触发：工具会在几百毫秒内把文件移走，
# 边触发边 ParseName 会拿到 null（文件已经不在了）。
$env:NFS_NAME  = '真实菜单20项'
$env:NFS_QUIET = '1'
$folder = $shell.Namespace($test)
$verbs = @()
foreach ($p in ($all | Select-Object -First 15)) {
    $item = $folder.ParseName((Split-Path -Leaf $p))
    if (-not $item) { Write-Host "  取不到: $p"; continue }
    $verb = @($item.Verbs()) | Where-Object { ($_.Name -replace '&','') -like '*新建文件夹并移入*' } | Select-Object -First 1
    if ($verb) { $verbs += $verb } else { Write-Host "  菜单项缺失: $p" }
}
# 像 Explorer 那样连续触发（真实右键菜单不会隔 120ms 一项）
$triggered = 0
foreach ($verb in $verbs) {
    try { $verb.DoIt(); $triggered++ } catch { Write-Host "  触发失败: $($_.Exception.Message)" }
}
Write-Host "已触发 $triggered 次（只覆盖前 15 项）"
Start-Sleep -Seconds 8
$env:NFS_NAME = $null; $env:NFS_QUIET = $null

$target = Join-Path $test '真实菜单20项'
$moved  = @(Get-ChildItem -LiteralPath $target -File -ErrorAction SilentlyContinue)
$left   = @(Get-ChildItem -LiteralPath $test -File -ErrorAction SilentlyContinue)
$dirs   = @(Get-ChildItem -LiteralPath $test -Directory -ErrorAction SilentlyContinue | ForEach-Object Name)
Write-Host "文件夹: $($dirs -join ' | ')"
Write-Host "移入=$($moved.Count)  留在外面=$($left.Count)"
Write-Host '--- log ---'
if (Test-Path "$qdir\last-run.log") { Get-Content "$qdir\last-run.log" -Encoding UTF8 }
try { $win.Quit() } catch {}

if ($moved.Count -eq $N -and $left.Count -eq 0 -and $dirs.Count -eq 1) { Write-Host "PASS: 真实菜单项 + 20 项全移入同一个文件夹"; exit 0 }
Write-Host 'FAIL'
exit 1
