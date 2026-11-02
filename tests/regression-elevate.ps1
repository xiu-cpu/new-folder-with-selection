# regression-elevate.ps1 — 真实提权链路（Program Files 下普通用户不可写）
# 需要本机 UAC 允许静默提权（ConsentPromptBehaviorAdmin=0）；否则会弹 UAC 框。

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8

$dir  = Split-Path -Parent $PSScriptRoot
$exe  = Join-Path $dir 'NewFolderWithSelection.exe'
$qdir = Join-Path $env:TEMP 'NewFolderWithSelection'
$src  = Join-Path $env:ProgramFiles 'NFS-Elevate-Test'
$name = '提权建的'
$tmp  = Join-Path $env:TEMP 'opencode'

$script:pass = 0
$script:fail = 0
function Check($name, $cond, $detail) {
    if ($cond) { $script:pass++; Write-Host ("  PASS  {0}" -f $name) }
    else       { $script:fail++; Write-Host ("  FAIL  {0}   -> {1}" -f $name, $detail) }
}

function Run-Elevated([string]$scriptPath) {
    $p = Start-Process -FilePath 'powershell' -Verb RunAs -Wait -PassThru -WindowStyle Hidden `
         -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $scriptPath))
    return $p.ExitCode
}

if (-not (Test-Path $exe)) { Write-Host "找不到 $exe"; exit 1 }

# 确认当前不是管理员（否则测不到提权路径）
$id  = [Security.Principal.WindowsIdentity]::GetCurrent()
$pr  = New-Object Security.Principal.WindowsPrincipal($id)
if ($pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host 'SKIP: 当前已是管理员，测不到提权路径'
    exit 0
}

New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$setup = Join-Path $tmp 'nfs-elev-setup.ps1'
@"
`$ErrorActionPreference = 'Stop'
if (Test-Path -LiteralPath '$src') { Remove-Item -LiteralPath '$src' -Recurse -Force }
New-Item -ItemType Directory -Force -Path '$src' | Out-Null
'x' | Set-Content -LiteralPath '$src\e1.txt'
'x' | Set-Content -LiteralPath '$src\e2.txt'
"@ | Set-Content -LiteralPath $setup -Encoding UTF8

$cleanup = Join-Path $tmp 'nfs-elev-cleanup.ps1'
@"
`$ErrorActionPreference = 'Stop'
if (Test-Path -LiteralPath '$src') {
    & icacls '$src' /reset /T /C /Q 2>`$null | Out-Null
    Remove-Item -LiteralPath '$src' -Recurse -Force
}
"@ | Set-Content -LiteralPath $cleanup -Encoding UTF8

Write-Host "准备（管理员）：$src"
if ((Run-Elevated $setup) -ne 0) { Write-Host 'FAIL: 准备失败'; exit 1 }

# 普通用户确认写不进去
$denied = $false
try { New-Item -ItemType Directory -Path (Join-Path $src 'probe') -ErrorAction Stop | Out-Null }
catch { $denied = $true }
Check '普通用户确实不能在目标目录新建文件夹（真的需要提权）' $denied 'probe succeeded'

Remove-Item -Recurse -Force $qdir -ErrorAction SilentlyContinue
$env:NFS_QUIET = '1'
$env:NFS_NAME  = $name
$env:NFS_NOELEVATE = $null
$sw = [Diagnostics.Stopwatch]::StartNew()
$p = Start-Process -FilePath $exe -ArgumentList @(('"{0}"' -f (Join-Path $src 'e1.txt')), ('"{0}"' -f (Join-Path $src 'e2.txt'))) -Wait -PassThru -WindowStyle Hidden
$elapsed = $sw.ElapsedMilliseconds
$env:NFS_QUIET = $null; $env:NFS_NAME = $null

$target = Join-Path $src $name
$moved  = @(Get-ChildItem -LiteralPath $target -File -ErrorAction SilentlyContinue)
$left   = @(Get-ChildItem -LiteralPath $src -File -ErrorAction SilentlyContinue)
$log    = ''
if (Test-Path (Join-Path $qdir 'last-run.log')) { $log = Get-Content (Join-Path $qdir 'last-run.log') -Raw -Encoding UTF8 }

Write-Host "耗时: $elapsed ms"
Write-Host "移入=$($moved.Count)  留在外面=$($left.Count)"
Write-Host '--- log ---'
if ($log) { Write-Host $log }

Check '提权后文件夹建出来了' (Test-Path -LiteralPath $target) $target
Check '2 项都移入（含被权限挡住的那项）' ($moved.Count -eq 2) (($moved | ForEach-Object Name) -join ', ')
Check '外面没有残留文件' ($left.Count -eq 0) (($left | ForEach-Object Name) -join ', ')
Check '主进程 exit 0' ($p.ExitCode -eq 0) "exit=$($p.ExitCode)"
Check '日志里有提权记录' ($log -match '提权') 'no elevate log'

Write-Host "`n=== 汇总 ==="
Write-Host ("PASS {0} / FAIL {1}" -f $script:pass, $script:fail)

Write-Host '清理（管理员）…'
[void](Run-Elevated $cleanup)

if ($script:fail -eq 0) { exit 0 } else { exit 1 }
