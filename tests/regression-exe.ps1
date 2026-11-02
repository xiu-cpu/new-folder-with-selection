# regression-exe.ps1 — NewFolderWithSelection.exe 功能回归
# 全部用 NFS_QUIET=1 跑，不弹任何窗口（不会打扰正在用电脑的人）

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8

$dir   = Split-Path -Parent $PSScriptRoot
$exe   = Join-Path $dir 'NewFolderWithSelection.exe'
$qdir  = Join-Path $env:TEMP 'NewFolderWithSelection'
$root  = Join-Path $env:TEMP 'opencode\nfs-exe'

if (-not (Test-Path $exe)) { Write-Host "找不到 $exe，先跑 build.ps1"; exit 1 }

$script:pass = 0
$script:fail = 0
function Check($name, $cond, $detail) {
    if ($cond) { $script:pass++; Write-Host ("  PASS  {0}" -f $name) }
    else       { $script:fail++; Write-Host ("  FAIL  {0}   -> {1}" -f $name, $detail) }
}
function Reset-All {
    Remove-Item -Recurse -Force $qdir -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $root) {
        & icacls $root /reset /T /C /Q 2>$null | Out-Null
        Remove-Item -Recurse -Force $root -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $env:NFS_QUIET = '1'
    $env:NFS_NOMOVE = $null
    $env:NFS_NAME = $null
    $env:NFS_NOELEVATE = $null
}
function Run-Tool {
    param([string[]]$paths, [string]$name)
    if ($name) { $env:NFS_NAME = $name } else { $env:NFS_NAME = $null }
    $argList = @()
    foreach ($p in $paths) { $argList += ('"{0}"' -f $p) }
    $proc = Start-Process -FilePath $exe -ArgumentList $argList -Wait -PassThru -WindowStyle Hidden
    $env:NFS_NAME = $null
    $out = ''
    $f = Join-Path $qdir 'last-output.txt'
    if (Test-Path -LiteralPath $f) { $out = Get-Content -LiteralPath $f -Raw -Encoding UTF8; Remove-Item -LiteralPath $f -Force }
    return @{ code = $proc.ExitCode; out = $out }
}
function Files-In($d) { @(Get-ChildItem -LiteralPath $d -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name }) }
function Esc([string]$s) { $s.Replace('%', '%25').Replace("`t", '%09').Replace("`r", '%0D').Replace("`n", '%0A') }

Write-Host "使用：$exe"

# =====================================================================
Write-Host "`n=== A. 一次传 4 项（中文/空格/子文件夹/无效项） ==="
Reset-All
$src = Join-Path $root 'A\源'
New-Item -ItemType Directory -Force -Path (Join-Path $src '子文件夹') | Out-Null
'x' | Set-Content -LiteralPath (Join-Path $src '我的 截图01.png')
'x' | Set-Content -LiteralPath (Join-Path $src 'my file 02.png')
'x' | Set-Content -LiteralPath (Join-Path $src '报告.pdf')
'x' | Set-Content -LiteralPath (Join-Path $src '子文件夹\内层.txt')
$r = Run-Tool @("$src\我的 截图01.png", "$src\my file 02.png", "$src\报告.pdf", "$src\子文件夹", "$src\不存在.xyz") '测试分组'
$dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
Check '只生成 1 个新文件夹' ($dirs.Count -eq 1 -and $dirs[0] -eq '测试分组') ($dirs -join ' | ')
Check '4 项全部移入' ((Files-In (Join-Path $src '测试分组')).Count -eq 4) ((Files-In (Join-Path $src '测试分组')) -join ', ')
Check '无效项不报错（exit 0）' ($r.code -eq 0) "exit=$($r.code) out=$($r.out)"
Check '子文件夹内部结构保留' (Test-Path -LiteralPath (Join-Path $src '测试分组\子文件夹\内层.txt')) 'missing'

# =====================================================================
Write-Host "`n=== B. 3 个独立进程（模拟 Explorer 逐项调用） ==="
Reset-All
$src = Join-Path $root 'B\源'
New-Item -ItemType Directory -Force -Path $src | Out-Null
foreach ($n in '甲 一.png', '乙二.png', '丙三.png') { 'x' | Set-Content -LiteralPath (Join-Path $src $n) }
$env:NFS_NAME = '并发汇总'
$procs = @()
foreach ($n in '甲 一.png', '乙二.png', '丙三.png') {
    $procs += Start-Process -FilePath $exe -ArgumentList ('"{0}"' -f (Join-Path $src $n)) -PassThru -WindowStyle Hidden
}
$procs | ForEach-Object { $_.WaitForExit(15000) | Out-Null }
$env:NFS_NAME = $null
$dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
Check '只生成 1 个文件夹（没被拆成多个）' ($dirs.Count -eq 1 -and $dirs[0] -eq '并发汇总') ($dirs -join ' | ')
Check '3 项全在同一个文件夹里' ((Files-In (Join-Path $src '并发汇总')).Count -eq 3) ((Files-In (Join-Path $src '并发汇总')) -join ', ')

# =====================================================================
Write-Host "`n=== C. 迟到调用不会另建文件夹（done 标记） ==="
$r = Run-Tool @("$src\甲 一.png") '并发汇总'
$dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
Check '迟到的调用被忽略（仍只有 1 个文件夹）' ($dirs.Count -eq 1) ($dirs -join ' | ')
Check '迟到调用 exit 0' ($r.code -eq 0) "exit=$($r.code)"

# =====================================================================
Write-Host "`n=== D. 单文件 / 单文件夹 / 首个是文件夹 ==="
Reset-All
$src = Join-Path $root 'D\源'
New-Item -ItemType Directory -Force -Path (Join-Path $src '空文件夹') | Out-Null
'x' | Set-Content -LiteralPath (Join-Path $src '报告.pdf')
$r = Run-Tool @("$src\报告.pdf") $null
$dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
Check '单文件用文件名做默认名' ($dirs -contains '报告') ($dirs -join ' | ')
Check '单文件已移入' ((Files-In (Join-Path $src '报告')).Count -eq 1) ''

Reset-All
$src = Join-Path $root 'D2\源'
New-Item -ItemType Directory -Force -Path (Join-Path $src '空文件夹') | Out-Null
$r = Run-Tool @("$src\空文件夹") $null
$dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
Check '单文件夹：父目录取对（不会移进自己）' ($dirs.Count -eq 1 -and $dirs[0] -eq '空文件夹 (2)') ($dirs -join ' | ')
Check '单文件夹已移入' (Test-Path -LiteralPath (Join-Path $src '空文件夹 (2)\空文件夹')) 'missing'

Reset-All
$src = Join-Path $root 'D3\源'
New-Item -ItemType Directory -Force -Path (Join-Path $src '一个目录') | Out-Null
'x' | Set-Content -LiteralPath (Join-Path $src '一个目录\内.txt')
'x' | Set-Content -LiteralPath (Join-Path $src '文件.txt')
$r = Run-Tool @("$src\一个目录", "$src\文件.txt") '混合选中'
$dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
Check '只生成 1 个文件夹（没在子目录里乱建）' ($dirs.Count -eq 1 -and $dirs[0] -eq '混合选中') ($dirs -join ' | ')
Check '文件夹和文件都移进去了' ((Files-In (Join-Path $src '混合选中')).Count -eq 2) ((Files-In (Join-Path $src '混合选中')) -join ', ')

# =====================================================================
Write-Host "`n=== E. 提权子进程模式（-JobFile） ==="
Reset-All
$src = Join-Path $root 'E\源'
New-Item -ItemType Directory -Force -Path $src | Out-Null
foreach ($n in 'j1.png', 'j2.png') { 'x' | Set-Content -LiteralPath (Join-Path $src $n) }
$target = Join-Path $src '提权建的'
$jobF = Join-Path $root 'E\job.txt'
$resF = Join-Path $root 'E\res.txt'
$lines = @((Esc $target), (Esc $resF), (Esc "$src\j1.png"), (Esc "$src\j2.png"))
[IO.File]::WriteAllLines($jobF, $lines, (New-Object Text.UTF8Encoding($false)))
$p = Start-Process -FilePath $exe -ArgumentList @('-JobFile', ('"{0}"' -f $jobF)) -Wait -PassThru -WindowStyle Hidden
$res = @()
if (Test-Path -LiteralPath $resF) { $res = @(Get-Content -LiteralPath $resF -Encoding UTF8) }
Check '子进程 exit 0' ($p.ExitCode -eq 0) "exit=$($p.ExitCode)"
Check '子进程写了结果文件' ($res.Count -eq 2) ($res -join ' / ')
Check '子进程结果全部成功' (($res | Where-Object { -not $_.StartsWith('1') }).Count -eq 0) ($res -join ' / ')
Check '子进程真的把文件移进去了' ((Files-In $target).Count -eq 2) ((Files-In $target) -join ', ')

# =====================================================================
Write-Host "`n=== F. 被占用的文件：失败要逐条报出 ==="
Reset-All
$src = Join-Path $root 'F\源'
New-Item -ItemType Directory -Force -Path $src | Out-Null
foreach ($n in 'locked.bin', 'free.txt') { 'x' | Set-Content -LiteralPath (Join-Path $src $n) }
$fs = [System.IO.File]::Open((Join-Path $src 'locked.bin'), 'Open', 'Read', 'None')
$r = Run-Tool @("$src\locked.bin", "$src\free.txt") '占用测试'
$fs.Dispose()
$moved = Files-In (Join-Path $src '占用测试')
Check '能移的照常移入' ($moved -contains 'free.txt') ($moved -join ', ')
Check '占用的没被移走' ($moved -notcontains 'locked.bin') ($moved -join ', ')
Check '失败原因被报出（含“占用/使用”）' ($r.out -match '占用|使用|being used') $r.out
$logTxt = ''
if (Test-Path (Join-Path $qdir 'last-run.log')) { $logTxt = Get-Content (Join-Path $qdir 'last-run.log') -Raw -Encoding UTF8 }
Check '日志里记录了失败项' ($logTxt -match '失败') 'no 失败 in log'

# =====================================================================
Write-Host "`n=== G. 只读目录：明确提示，且一个文件都不动 ==="
Reset-All
$src = Join-Path $root 'G\只读'
New-Item -ItemType Directory -Force -Path $src | Out-Null
'x' | Set-Content -LiteralPath (Join-Path $src 'a.txt')
& icacls $src /deny "$($env:USERNAME):(W)" | Out-Null
$env:NFS_NOELEVATE = '1'
$r = Run-Tool @("$src\a.txt") '只读测试'
$env:NFS_NOELEVATE = $null
& icacls $src /remove:d "$($env:USERNAME)" | Out-Null
$dirs = @(Get-ChildItem -LiteralPath $src -Directory -ErrorAction SilentlyContinue | ForEach-Object Name)
Check '没建出文件夹' ($dirs.Count -eq 0) ($dirs -join ' | ')
Check '文件原地未动' (Test-Path -LiteralPath (Join-Path $src 'a.txt')) 'file gone'
Check '给出了“无法新建”的提示' ($r.out -match '无法在该位置新建文件夹|只读') $r.out

# =====================================================================
Write-Host "`n=== H. 建议名（NFS_NOMOVE） ==="
Reset-All
$src = Join-Path $root 'H\源'
New-Item -ItemType Directory -Force -Path $src | Out-Null
foreach ($n in 'a.png', 'b.jpg', 'v.mp4', 'v2.mkv', 'a.pdf', '报告.pdf', 'z.zip') { 'x' | Set-Content -LiteralPath (Join-Path $src $n) }
function Sugg([string[]]$paths) {
    $env:NFS_NOMOVE = '1'
    $argList = @(); foreach ($p in $paths) { $argList += ('"{0}"' -f $p) }
    Start-Process -FilePath $exe -ArgumentList $argList -Wait -WindowStyle Hidden | Out-Null
    $env:NFS_NOMOVE = $null
    $s = ''
    $f = Join-Path $qdir 'last-output.txt'
    if (Test-Path $f) { $s = (Get-Content -LiteralPath $f -Encoding UTF8 | Where-Object { $_ } | Select-Object -Last 1); Remove-Item $f -Force }
    return $s
}
$s = Sugg @("$src\a.png", "$src\b.jpg");              Check '全图片 → 图片 时间戳' ($s -match '^图片 \d{8}-\d{6}$') $s
$s = Sugg @("$src\v.mp4", "$src\v2.mkv");             Check '全视频 → 视频 时间戳' ($s -match '^视频 \d{8}-\d{6}$') $s
$s = Sugg @("$src\a.pdf", "$src\报告.pdf");           Check '全文档 → 文档 时间戳' ($s -match '^文档 \d{8}-\d{6}$') $s
$s = Sugg @("$src\z.zip");                            Check '单文件 → 文件名' ($s -eq 'z') $s
$s = Sugg @("$src\a.png", "$src\v.mp4", "$src\a.pdf"); Check '混合 → 扩展名组合' ($s -match '^PNG\+MP4\+PDF 文件 \d{8}-\d{6}$') $s

# =====================================================================
Write-Host "`n=== I. 重名 / 非法字符 / 空名字 ==="
Reset-All
$src = Join-Path $root 'I\源'
New-Item -ItemType Directory -Force -Path $src | Out-Null
foreach ($n in 'a.png', 'b.png', 'c.png') { 'x' | Set-Content -LiteralPath (Join-Path $src $n) }
Run-Tool @("$src\a.png") '重名测试' | Out-Null
Run-Tool @("$src\b.png") '重名测试' | Out-Null
Run-Tool @("$src\c.png") '重名测试' | Out-Null
$dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
Check '重名自动追加 (2) (3)' ($dirs -contains '重名测试' -and $dirs -contains '重名测试 (2)' -and $dirs -contains '重名测试 (3)') ($dirs -join ' | ')

Reset-All
$src = Join-Path $root 'I2\源'
New-Item -ItemType Directory -Force -Path $src | Out-Null
'x' | Set-Content -LiteralPath (Join-Path $src 'd.png')
Run-Tool @("$src\d.png") ' a:b?c<d>e|  ' | Out-Null
$dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
Check '非法字符换下划线、首尾空格去掉' ($dirs.Count -eq 1 -and $dirs[0] -eq 'a_b_c_d_e_') ($dirs -join ' | ')

Reset-All
$src = Join-Path $root 'I3\源'
New-Item -ItemType Directory -Force -Path $src | Out-Null
'x' | Set-Content -LiteralPath (Join-Path $src 'e.png')
$r = Run-Tool @("$src\e.png") '   ...   '
$dirs = @(Get-ChildItem -LiteralPath $src -Directory | ForEach-Object Name)
Check '全空/全点名字：什么都不做' ($dirs.Count -eq 0 -and (Test-Path (Join-Path $src 'e.png'))) ($dirs -join ' | ')

# =====================================================================
Write-Host "`n=== J. 队列/锁残留 ==="
Check '跑完不残留队列文件' (-not (Test-Path (Join-Path $qdir 'queue.txt'))) 'queue.txt 还在'
Check '跑完不残留锁文件' (-not (Test-Path (Join-Path $qdir 'owner.lock'))) 'owner.lock 还在'

# =====================================================================
Write-Host "`n=== 汇总 ==="
Write-Host ("PASS {0} / FAIL {1}" -f $script:pass, $script:fail)
Remove-Item -Recurse -Force $root -ErrorAction SilentlyContinue
if ($script:fail -eq 0) { exit 0 } else { exit 1 }
