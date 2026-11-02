# install.ps1 — 注册右键菜单项（仅当前用户，不需要管理员）
# 出现在"显示更多选项"(Shift+F10) 的经典菜单里，名称：新建文件夹并移入 (&F)
#
# 关键点：命令用 "%1"（Explorer 对每个选中项各调用一次），不能用 "%*"（实测为空）。
# 多选的汇总由程序内部的队列 + owner 锁完成。

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$exe  = Join-Path $here 'NewFolderWithSelection.exe'

if (-not (Test-Path $exe)) {
    Write-Host "没找到 $exe，先编译…"
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here 'build.ps1')
    if (-not (Test-Path $exe)) {
        Write-Host '编译失败，请手动运行同目录的 build.ps1（需要 .NET Framework 4.x，Windows 自带）'
        exit 1
    }
}

$key = 'HKCU:\Software\Classes\AllFilesystemObjects\shell\NewFolderWithSelection'
$cmd = '"{0}" "%1"' -f $exe

New-Item -ItemType Directory -Force -Path "$key\command" | Out-Null
Set-ItemProperty -Path $key -Name '(Default)' -Value '新建文件夹并移入 (&F)'
Set-ItemProperty -Path $key -Name 'Icon'      -Value 'imageres.dll,-1003'
Remove-ItemProperty -Path $key -Name 'MultiSelectModel' -ErrorAction SilentlyContinue
Set-ItemProperty -Path "$key\command" -Name '(Default)' -Value $cmd

# 清掉可能残留的队列/锁
Remove-Item -LiteralPath (Join-Path $env:TEMP 'NewFolderWithSelection') -Recurse -Force -ErrorAction SilentlyContinue

Write-Host '已安装：选中文件/文件夹 → 右键 → 显示更多选项(Shift+F10) → 新建文件夹并移入'
Write-Host "命令： $cmd"
Write-Host '卸载： 运行同目录 uninstall.ps1'
