$startup = [Environment]::GetFolderPath('Startup')
$vbsPath = Join-Path $startup "FactoryWorkbench.vbs"
Write-Host "Starting VBS: $vbsPath"
Start-Process wscript.exe -ArgumentList "`"$vbsPath`""
Start-Sleep 5
$procs = Get-Process -Name python* -ErrorAction SilentlyContinue
if ($procs) {
    Write-Host "Python processes running: " + $procs.Count
    foreach ($p in $procs) { Write-Host ("  PID: " + $p.Id) }
} else {
    Write-Host "No Python processes found!"
}
try {
    $r1 = Invoke-WebRequest -Uri "http://127.0.0.1:8932/" -UseBasicParsing -TimeoutSec 5
    Write-Host ("Stable (8932): " + $r1.StatusCode)
} catch { Write-Host "Stable (8932): FAILED" }
try {
    $r2 = Invoke-WebRequest -Uri "http://127.0.0.1:8933/" -UseBasicParsing -TimeoutSec 5
    Write-Host ("Test (8933): " + $r2.StatusCode)
} catch { Write-Host "Test (8933): FAILED" }
