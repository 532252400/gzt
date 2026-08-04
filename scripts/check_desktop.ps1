$desktop = [Environment]::GetFolderPath('Desktop')
Write-Host "=== 桌面文件 ==="
Get-ChildItem -Path $desktop | Where-Object { -not $_.PSIsContainer } | Sort-Object Name | ForEach-Object {
    $size = [math]::Round($_.Length / 1KB)
    Write-Host ("  " + $_.Name + "  [" + $_.LastWriteTime.ToString("MM-dd HH:mm") + "]  " + $size + "KB")
}
Write-Host ""
Write-Host "=== 桌面文件夹 ==="
Get-ChildItem -Path $desktop | Where-Object { $_.PSIsContainer } | Sort-Object Name | ForEach-Object {
    Write-Host ("  [DIR] " + $_.Name)
}
