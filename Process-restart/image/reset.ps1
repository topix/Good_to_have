$process1 = Get-Process "EXAMPLE_PROCESS" -ErrorAction SilentlyContinue
$logFile = "c:\Utils\Process-restart.log"

If ($process1)
{
# Process 1 - Stop
    Write-Host -foreground "red" [$([DateTime]::Now)]": INFO - Shutting down Process 1" 
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Shutting down Process 1"
    Stop-Process -processname "EXAMPLE_PROCESS"
    Write-Host -foreground "red" [$([DateTime]::Now)]": INFO - Process 1 shutdown successfully" 
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Process 1 shutdown successfully"
    
# Subprocess 2 - Stop
    Write-Host -foreground "red" [$([DateTime]::Now)]": INFO - Restarting Subprocess 2 Service"
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Restarting Subprocess 2 Service"
    Stop-Service PosPayService -Force -WarningAction SilentlyContinue
    
# Subprocess 2 - Start
    Write-Host -foreground "green" [$([DateTime]::Now)]": INFO - Starting Subprocess 2 Service"
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Starting Subprocess 2 Service"
    Start-Service PosPayService -WarningAction SilentlyContinue
    Write-Host -foreground "green" [$([DateTime]::Now)]": INFO - Subprocess 2 Service Started"
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Subprocess 2 Service Started"
    
# Process 1 - Start
    Start-Sleep -s 1
    Write-Host -foreground "yellow" [$([DateTime]::Now)]": INFO - Starting Process 1"
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Starting Process 1"
    Start-Process "EXAMPLE_PROCESS.exe" -WorkingDirectory "C:\EXAMPLE_PROCESS\"
    Write-Host -foreground "green" [$([DateTime]::Now)]": INFO - Process 1 Started"
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Process 1 Started"
    Start-Sleep -s 5
}
Else
{
# Subprocess 2 - Stop
    Write-Host -foreground "red" [$([DateTime]::Now)]": INFO - Restarting Subprocess 2 Service"
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Restarting Subprocess 2 Service"
    Stop-Service PosPayService -Force -WarningAction SilentlyContinue
    
# Subprocess 2 - Start
    Write-Host -foreground "green" [$([DateTime]::Now)]": INFO - Starting Subprocess 2 Service"
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Starting Subprocess 2 Service"
    Start-Service PosPayService -WarningAction SilentlyContinue
    Write-Host -foreground "green" [$([DateTime]::Now)]": INFO - Subprocess 2 Service Started"
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Subprocess 2 Service Started"
    
# Process 1 - Start
    Start-Sleep -s 1
    Write-Host -foreground "yellow" [$([DateTime]::Now)]": INFO - Starting Process 1"
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Starting Process 1"
    Start-Process "EXAMPLE_PROCESS.exe" -WorkingDirectory "C:\EXAMPLE_PROCESS\"
    Write-Host -foreground "green" [$([DateTime]::Now)]": INFO - Process 1 Started"
    Add-Content $logFile [$([DateTime]::Now)]": INFO - Process 1 Started"
    Start-Sleep -s 5
}

If( (get-item $logFile).length -gt 100MB)
{
    Remove-Item $logFile
    Add-Content -Path $logFile -Value [$([DateTime]::Now)]": INFO - Log has been purged!"
}