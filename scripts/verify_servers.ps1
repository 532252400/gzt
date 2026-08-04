Write-Host "=== 验证服务 ==="
Start-Sleep -s 3
try {
    $r1 = Invoke-WebRequest -Uri "http://127.0.0.1:8932/" -UseBasicParsing -TimeoutSec 5
    Write-Host ("稳定版 8932: " + $r1.StatusCode + " OK")
} catch {
    Write-Host "稳定版 8932: 未运行"
}
try {
    $r2 = Invoke-WebRequest -Uri "http://127.0.0.1:8933/" -UseBasicParsing -TimeoutSec 5
    Write-Host ("测试版 8933: " + $r2.StatusCode + " OK")
} catch {
    Write-Host "测试版 8933: 未运行"
}
Write-Host "=== 开机启动项 ==="
$startup = [Environment]::GetFolderPath('Startup')
Get-ChildItem -Path $startup | ForEach-Object { Write-Host ("  " + $_.Name) }
