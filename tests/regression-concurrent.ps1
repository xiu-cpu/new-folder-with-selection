# regression-concurrent.ps1 — 真并发回归（不再靠串行拉起掩盖队列丢行）
#
# 做法：把 N 个实例放进同一个紧循环里同时启动（不等待），让它们真的在同一毫秒级时间里
# 往 queue.txt 追加。断言：
#   1) 只建出 1 个文件夹；
#   2) 源目录里一个文件都不剩（说明每一行都被 owner 收到并搬走）；
#   3) 文件夹里正好 N 项。
# 只要 FileMode.Append 不是原子追加、丢了任何一行，这个用例就会失败。
#
# 用法：
#   powershell -File tests\regression-concurrent.ps1                 # 默认 10 轮 x 15 项
#   powershell -File tests\regression-concurrent.ps1 -Rounds 3 -Items 15
#   powershell -File tests\regression-concurrent.ps1 -Exe <其它版本的 exe>   # 对比修复前后

param(
    [string]$Exe = '',
    [int]$Rounds = 10,
    [int]$Items = 15
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8

$dir  = Split-Path -Parent $PSScriptRoot
if (-not $Exe) { $Exe = Join-Path $dir 'NewFolderWithSelection.exe' }
if (-not (Test-Path -LiteralPath $Exe)) { Write-Host "找不到 $Exe，先跑 build.ps1"; exit 1 }

$qdir = Join-Path $env:TEMP 'NewFolderWithSelection'
$log  = Join-Path $qdir 'last-run.log'
$root = Join-Path $env:TEMP 'opencode\nfs-concurrent'

$script:pass = 0
$script:fail = 0
function Check($name, $cond, $detail) {
    if ($cond) { $script:pass++; Write-Host ("  PASS  {0}" -f $name) }
    else       { $script:fail++; Write-Host ("  FAIL  {0}   -> {1}" -f $name, $detail) }
}
function Esc([string]$s) { $s.Replace('%', '%25').Replace("`t", '%09').Replace("`r", '%0D').Replace("`n", '%0A') }

function Reset-All {
    Remove-Item -Recurse -Force $qdir -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $root) { Remove-Item -Recurse -Force $root -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $env:NFS_QUIET = '1'
    $env:NFS_NOMOVE = $null
    $env:NFS_NAME = $null
    $env:NFS_NOELEVATE = '1'
}

# 只读日志里 offset 之后新追加的部分（别的进程可能还在写，所以用 ReadWrite 共享）
function Read-LogFrom([long]$offset) {
    if (-not (Test-Path -LiteralPath $log)) { return '' }
    $len = (Get-Item -LiteralPath $log).Length
    if ($len -le $offset) { return '' }
    $fs = [IO.File]::Open($log, 'Open', 'Read', 'ReadWrite')
    try {
        $fs.Position = $offset
        $buf = New-Object byte[] ($len - $offset)
        $n = $fs.Read($buf, 0, $buf.Length)
        return [Text.Encoding]::UTF8.GetString($buf, 0, $n)
    } finally { $fs.Dispose() }
}

Write-Host "使用：$Exe"
Write-Host ("轮数：{0}   每轮实例：{1}" -f $Rounds, $Items)

# =====================================================================
Write-Host "`n=== A. $Items 个实例同毫秒级并发追加（$Rounds 轮） ==="
$lostRounds = 0
$lostItems  = 0
for ($r = 1; $r -le $Rounds; $r++) {
    Reset-All
    $src = Join-Path $root ("R{0:D2}\源" -f $r)
    New-Item -ItemType Directory -Force -Path $src | Out-Null
    $files = @()
    for ($i = 1; $i -le $Items; $i++) {
        $f = Join-Path $src ("项{0:D2}.txt" -f $i)
        'x' | Set-Content -LiteralPath $f
        $files += $f
    }

    $logOffset = 0
    if (Test-Path -LiteralPath $log) { $logOffset = (Get-Item -LiteralPath $log).Length }

    $env:NFS_QUIET = '1'
    $env:NFS_NAME  = '并发测试'

    # 先全部建好 ProcessStartInfo，再一口气 Start：抖动只剩 CreateProcess 本身
    $psis = @()
    foreach ($f in $files) {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName        = $Exe
        $psi.Arguments       = '"{0}"' -f $f
        $psi.UseShellExecute = $false
        $psi.WorkingDirectory = $dir
        $psis += $psi
    }
    $procs = @()
    foreach ($psi in $psis) { $procs += [System.Diagnostics.Process]::Start($psi) }
    foreach ($p in $procs) { $p.WaitForExit(30000) | Out-Null; $p.Dispose() }

    $env:NFS_NAME = $null

    $dirs  = @(Get-ChildItem -LiteralPath $src -Directory -ErrorAction SilentlyContinue)
    $left  = @(Get-ChildItem -LiteralPath $src -File -ErrorAction SilentlyContinue)
    $moved = @()
    if ($dirs.Count -ge 1) { $moved = @(Get-ChildItem -LiteralPath $dirs[0].FullName -File -ErrorAction SilentlyContinue) }

    $slice = Read-LogFrom $logOffset
    $q = -1
    $m = [regex]::Match($slice, '队列 (\d+) 项')
    if ($m.Success) { $q = [int]$m.Groups[1].Value }
    $appendFail = ($slice -match '队列追加失败')

    $ok = ($dirs.Count -eq 1) -and ($moved.Count -eq $Items) -and ($left.Count -eq 0)
    if (-not $ok) { $lostRounds++; $lostItems += ($Items - $moved.Count) }

    Write-Host ("  第 {0,2} 轮：文件夹={1} 移入={2}/{3} 源目录剩余={4} owner收到队列={5} 追加失败={6} -> {7}" -f `
        $r, $dirs.Count, $moved.Count, $Items, $left.Count, $q, $appendFail, $(if ($ok) { 'PASS' } else { 'FAIL' }))
    if (-not $ok) {
        Check ("第 {0} 轮：1 个文件夹 / {1} 项全移走" -f $r, $Items) $false `
            ("dirs={0} moved={1} left={2} queue={3}" -f $dirs.Count, $moved.Count, $left.Count, $q)
    } else { $script:pass++ }
}

$rate = 0
if ($Rounds -gt 0) { $rate = [Math]::Round(100.0 * $lostRounds / $Rounds, 1) }
Write-Host ("`n  汇总：{0}/{1} 轮通过，丢行轮次 {2}，累计丢项 {3}，丢行率 {4}%" -f `
    ($Rounds - $lostRounds), $Rounds, $lostRounds, $lostItems, $rate)
Check '并发追加没有丢项（全部轮次）' ($lostRounds -eq 0) ("{0} 轮里 {1} 轮丢项，共丢 {2} 项（{3}%）" -f $Rounds, $lostRounds, $lostItems, $rate)

# =====================================================================
Write-Host "`n=== B. 崩溃残留的旧队列行不会污染下一批（B2） ==="
Reset-All
$src = Join-Path $root 'B\源'
New-Item -ItemType Directory -Force -Path $src | Out-Null
'x' | Set-Content -LiteralPath (Join-Path $src '旧格式行.txt')      # 旧版本崩溃留下的无时间戳行
'x' | Set-Content -LiteralPath (Join-Path $src '陈旧时间戳行.txt')  # 本版本崩溃留下的旧时间戳行
'x' | Set-Content -LiteralPath (Join-Path $src '本次选中.txt')

# 伪造上一次崩溃留下的队列：两种格式各一行，都指向用户这次没选的文件
New-Item -ItemType Directory -Force -Path $qdir | Out-Null
$oldTicks = (Get-Date).ToUniversalTime().AddHours(-1).Ticks
$enc = New-Object Text.UTF8Encoding($false)
$stale = (Esc (Join-Path $src '旧格式行.txt')) + "`r`n" +
         ($oldTicks.ToString() + "`t" + (Esc (Join-Path $src '陈旧时间戳行.txt')) + "`r`n")
[IO.File]::WriteAllText((Join-Path $qdir 'queue.txt'), $stale, $enc)

$env:NFS_NAME = '陈旧行测试'
$p = Start-Process -FilePath $Exe -ArgumentList ('"{0}"' -f (Join-Path $src '本次选中.txt')) -Wait -PassThru -WindowStyle Hidden
$env:NFS_NAME = $null

Check '旧格式残留行没有被搬走' (Test-Path -LiteralPath (Join-Path $src '旧格式行.txt')) '旧格式行.txt 被搬走了'
Check '陈旧时间戳行没有被搬走' (Test-Path -LiteralPath (Join-Path $src '陈旧时间戳行.txt')) '陈旧时间戳行.txt 被搬走了'
Check '这次真的选中项照常移入' (Test-Path -LiteralPath (Join-Path $src '陈旧行测试\本次选中.txt')) '没移入'
Check '新文件夹里只有 1 项' (@(Get-ChildItem -LiteralPath (Join-Path $src '陈旧行测试') -File).Count -eq 1) '多搬了别的东西'

# =====================================================================
Write-Host "`n=== C. 队列/锁残留 ==="
Check '跑完不残留队列文件' (-not (Test-Path -LiteralPath (Join-Path $qdir 'queue.txt'))) 'queue.txt 还在'
Check '跑完不残留锁文件' (-not (Test-Path -LiteralPath (Join-Path $qdir 'owner.lock'))) 'owner.lock 还在'

# =====================================================================
Write-Host "`n=== 汇总 ==="
Write-Host ("PASS {0} / FAIL {1}" -f $script:pass, $script:fail)
Remove-Item -Recurse -Force $root -ErrorAction SilentlyContinue
if ($script:fail -eq 0) { exit 0 } else { exit 1 }
