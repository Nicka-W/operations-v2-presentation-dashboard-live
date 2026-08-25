<#
register_local_autopublish.ps1

Creates or refreshes the Windows scheduled task that runs
run_local_autopublish.pyw every minute.

Usage (run in an elevated PowerShell prompt):
    .\register_local_autopublish.ps1

Edit $TaskName and $PythonwPath below if your setup differs.
#>

$TaskName = "Operations V2 PPT Dashboard Auto Publish"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RunnerPath = Join-Path $ScriptDir "run_local_autopublish.pyw"

# Auto-detect pythonw.exe (no console window) on PATH, or fall back to
# the standard py launcher.
$PythonwPath = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $PythonwPath) {
    $PythonwPath = "pyw.exe"
}

Write-Host "Registering scheduled task '$TaskName'"
Write-Host "  Runner:  $RunnerPath"
Write-Host "  Pythonw: $PythonwPath"

$Action = New-ScheduledTaskAction -Execute $PythonwPath -Argument "`"$RunnerPath`""

$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Minutes 1) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

# Remove existing task with the same name, if any, then register fresh.
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "Auto-publishes Operations V2 dashboard data every minute."

Write-Host "Done. Check status with:"
Write-Host "  schtasks /Query /TN `"$TaskName`" /FO LIST /V"
