# Setup auto-start for both factory workbench servers
$batchFile = "C:\Users\pc\ZCodeProject\scripts\run_servers.cmd"

# Remove old tasks if exist
schtasks /delete /tn "工厂工作台-稳定版" /f *>$null
schtasks /delete /tn "工厂工作台-测试版" /f *>$null

# Create a single scheduled task that runs the batch file at logon
schtasks /create /tn "工厂工作台" /tr "$batchFile" /sc onlogon /rl limited /f

Write-Host "=== Scheduled Tasks ==="
schtasks /query /tn "工厂工作台" /v 2>*$null | Select-String "工厂工作台|TaskName|Task To Run"
Write-Host ""
Write-Host "Auto-start setup complete!"
Write-Host "Both servers will start automatically when you log in."

