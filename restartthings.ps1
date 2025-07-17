# PowerShell script to restart applications after restarting
# SkimSrv, RTTY SkimSrv, and CWSL_DIGI.

param([switch]$Verbose)

# -----------------------------------------------------------
# Configuration for this specific installation
# This is the only part of the script that should be edited

$clusterClient = $true   # Set to $true if you have a cluster client

# Installation path for cluster client
$clientPath = "C:\Program Files (x86)\DXLog.net\"

# Naming of executable
$clientExe = "DXLog.net.DXC.exe"

# Naming in task manager
$clientName = "DXLog.net.DXC*"

# End of configuration section
# -----------------------------------------------------------

try 
{

    if ($clusterClient)
    {
        if ($Verbose) { Write-Host "Stopping $clientExe..." }
        Stop-Process -Name $clientName -Force -ErrorAction SilentlyContinue 
        if ($Verbose) { Write-Host "Wait for OS process clean up..." }
        Start-Sleep -Seconds 4
        if ($Verbose) { Write-Host "Starting $clientExe..." }
        Start-Process -WorkingDirectory $clientPath -FilePath $clientExe -WindowStyle Minimized
        if ($Verbose) { Write-Host "$clientExe started." }
        Write-Host "Application restart complete at $(Get-Date -Format "HH:mm:ss")"
    }   
}
catch 
{
    Write-Error "An error occurred: $($_.Exception.Message)"
    exit 1
}
