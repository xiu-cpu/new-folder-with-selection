param(
    [Parameter(Mandatory = $true)][string]$Path,
    [int]$Cols = 118
)

Add-Type -AssemblyName System.Drawing
$bmp = [System.Drawing.Bitmap]::FromFile($Path)
$W = $bmp.Width; $H = $bmp.Height
"=== $Path  ${W}x${H} ==="

$cw = $W / $Cols
$rows = [int][math]::Round($H / $cw)
$ch = $H / $rows

function Sym($c) {
    $r = [int]$c.R; $g = [int]$c.G; $b = [int]$c.B
    $lum = (0.2126 * $r + 0.7152 * $g + 0.0722 * $b) / 255.0
    if ((($b - $r) -gt 40) -and ($lum -lt 0.8)) { return 'B' }
    if ($lum -lt 0.25) { return '@' }
    if ($lum -lt 0.45) { return '#' }
    if ($lum -lt 0.65) { return '+' }
    if ($lum -lt 0.80) { return '.' }
    if ($lum -lt 0.93) { return ':' }
    if ($lum -lt 0.985) { return ',' }
    return ' '
}

$sb = New-Object System.Text.StringBuilder
for ($y = 0; $y -lt $rows; $y++) {
    $line = New-Object System.Text.StringBuilder
    for ($x = 0; $x -lt $Cols; $x++) {
        # 每格取 2x2 采样，取最“显著”的一个（避免细笔画被跳过）
        $best = $null; $bestLum = 9
        foreach ($dy in 0.25, 0.75) {
            foreach ($dx in 0.25, 0.75) {
                $px = [int](($x + $dx) * $cw); $py = [int](($y + $dy) * $ch)
                if ($px -ge $W) { $px = $W - 1 }
                if ($py -ge $H) { $py = $H - 1 }
                $c = $bmp.GetPixel($px, $py)
                $lum = (0.2126 * $c.R + 0.7152 * $c.G + 0.0722 * $c.B) / 255.0
                if ($lum -lt $bestLum) { $bestLum = $lum; $best = $c }
            }
        }
        [void]$line.Append((Sym $best))
    }
    '{0,3}|{1}' -f $y, $line.ToString()
}

''
'--- 区域检查 ---'

function Scan-Box {
    param([scriptblock]$Match, [string]$Name)
    $minX = $W; $minY = $H; $maxX = -1; $maxY = -1; $n = 0
    for ($y = 0; $y -lt $H; $y += 2) {
        for ($x = 0; $x -lt $W; $x += 2) {
            $c = $bmp.GetPixel($x, $y)
            if (& $Match $c) {
                $n++
                if ($x -lt $minX) { $minX = $x }
                if ($x -gt $maxX) { $maxX = $x }
                if ($y -lt $minY) { $minY = $y }
                if ($y -gt $maxY) { $maxY = $y }
            }
        }
    }
    if ($n -eq 0) { '{0,-14} 未找到' -f $Name; return }
    '{0,-14} 像素={1,-7} 范围=({2},{3})-({4},{5})  {6}x{7}' -f $Name, $n, $minX, $minY, $maxX, $maxY, ($maxX - $minX), ($maxY - $minY)
}

Scan-Box -Name '强调色块' -Match { param($c) (($c.B - $c.R) -gt 60) -and ($c.B -gt 140) -and ($c.R -lt 90) }
Scan-Box -Name '深色文字' -Match { param($c) (($c.R + $c.G + $c.B) -lt 330) }
Scan-Box -Name '白色区域' -Match { param($c) ($c.R -gt 250) -and ($c.G -gt 250) -and ($c.B -gt 250) }

$bmp.Dispose()
