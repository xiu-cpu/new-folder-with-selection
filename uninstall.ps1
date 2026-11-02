# uninstall.ps1 — 移除右键菜单项
Remove-Item -LiteralPath 'HKCU:\Software\Classes\AllFilesystemObjects\shell\NewFolderWithSelection' `
    -Recurse -Force -ErrorAction SilentlyContinue
Write-Host '已卸载（不删除工具本身）'
