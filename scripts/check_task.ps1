$task = Get-ScheduledTask -TaskName "FactoryWorkbench" -ErrorAction SilentlyContinue
if ($task) {
    Write-Host ("Task: " + $task.TaskName + " State: " + $task.State)
} else {
    Write-Host "Task not found, trying to create..."
    $batchFile = "C:\Users\pc\ZCodeProject\scripts\run_servers.cmd"
    $action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c start /B `"`" `"C:\Users\pc\ZCodeProject\scripts\run_servers.cmd`""
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    Register-ScheduledTask -TaskName "FactoryWorkbench" -Action $action -Trigger $trigger -Settings $settings -Force
    Write-Host "Task created!"
}
