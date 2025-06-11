# PowerShell script to update frequency calibration of CWSL_DIGI and two instances of SkimSrv
# Reads skew data from sm7iun.se/rbn/analytics and updates the ini files accordingly.

$dryRun = $false  # Set to $false to actually update the files

$callsign = "SM7IUN"  # Callsign to look for in the web page

# Configuration

$webUrl = "https://sm7iun.se/rbn/analytics"
$iniPath1 = $env:APPDATA + "\Afreet\Products\SkimSrv\"
$iniPath2 = $env:APPDATA + "\Afreet\Products\SkimSrv2\"
$configPath = "C:\CWSL_DIGI\"

$iniFile1 = "SkimSrv.ini"
$iniFile2 = "SkimSrv2.ini"
$configFile = "config.ini"

$iniFilePath1 = $iniPath1 + $iniFile1
$iniFilePath2 = $iniPath2 + $iniFile2
$configFilePath = $configPath + $configFile

$skimsrvPath1 = "C:\Program Files (x86)\Afreet\SkimSrv"
$skimsrvPath2 = "C:\Program Files (x86)\Afreet\SkimSrv2"
$cwslPath = "C:\CWSL_DIGI"

$skimsrvExe1 = "SkimSrv.exe"
$skimsrvExe2 = "SkimSrv2.exe"
$cwslExe = "CWSL_DIGI.exe"

$ProgressPreference = 'SilentlyContinue' # Show no progress bar

Write-Host "--------------------------------------------------"
Write-Host "Execution time is" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

try 
{
    # Parse SkimSrv ini file for current calibration factor
    # Both ini files should have the same calibration factor
    # Format of line 
    # FreqCalibration=1.00828283
    $iniContent = Get-Content $iniFilePath1 -Raw
    $iniMatch = [regex]::Match($iniContent, 'FreqCalibration=([01]\.\d+)?') 

    if ($iniMatch.Success) 
    {
        $inicalibration = [double]$iniMatch.Groups[1].Value
        Write-Host "Current SkimSrv calibration factor is: $inicalibration"
    }
    else 
    {
        Write-Error "Failed to read CWSL_DIGI calibration factor from $iniFile1"
        exit 1
    }

    # Parse CWSL_DIGI config file for current calibration factor
    # Format of line 
    # freqcalibration=1.00828283
    $configContent = Get-Content $configFilePath -Raw
    $configMatch = [regex]::Match($configContent, 'freqcalibration=([01]\.\d+)?')

    if ($configMatch.Success) 
    {
        $configCalibration = [double]$configMatch.Groups[1].Value
        Write-Host "Current CWSL_DIGI calibration factor is: $configCalibration"
    }
    else 
    {
        Write-Error "Failed to read calibration factor from $configFile"
        exit 1
    }

    # Parse web page for adjustment factor
    # The web page should contain a line with the desired callsign and statistics/analytics
    # Callsign has optional asterisk, followed by frequency error in ppm, spot count, and adjustment factor
    # Format of line
    #   SM7IUN*     +0.1   3999   1.000000099
    # Last updated 2025-06-09 00:16:23 UTC
    $webContent = Invoke-WebRequest -Uri $webUrl -UseBasicParsing -Headers @{"Cache-Control"="no-cache"}
    $webMatch = [regex]::Match($webContent.Content, $callsign + '\*? +[+-]\d\.\d+ +\d+ +([01]\.\d+)')
    $webTimeMatch = [regex]::Match($webContent.Content, 'Last updated +(20\d{2}-\d{1,2}-\d{1,2} +\d{1,2}:\d{2}:\d{2})')

    if (-not $webMatch.Success) 
    {
        Write-Host "No spots reported for $callsign. Skimmer may be down. Exiting."
        exit 0
    }

    if ($webTimeMatch.Success) 
    {
        $lastUpdated = $webTimeMatch.Groups[1].Value
        $newCalibration = [double]$webMatch.Groups[1].Value
        Write-Host "Adjustment factor from $webUrl at $lastUpdated from is: $newCalibration"
    }
    else 
    {
        Write-Error "Failed to find update time in $webUrl"
        exit 1
    }

    # Calculate new calibration factors
    # Round to 9 decimal places like in the web page which is an overkill of accuracy
    $skimSrvCalibration = [Math]::Round($newCalibration * $inicalibration, 9)
#    Write-Host "New calibration factor for SkimSrv: $skimSrvCalibration"
    $cwsldigiCalibration = [Math]::Round(1.0 / ($newCalibration * $inicalibration), 9)
#    Write-Host "New calibration factor for CWSL_DIGI: $cwsldigiCalibration"
    
    # Stop the applications
    Write-Host "Stopping $skimsrvExe1, $skimsrvExe2, and $cwslExe..."

    # Stop SkimSrv instances
    Stop-Process -Name "SkimSrv*" -Force -ErrorAction SilentlyContinue 
    
    # Close CWSL_DIGI process including subprocesses
    Stop-Process -Name "CWSL*" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "jt9*" -Force -ErrorAction SilentlyContinue

    # Wait a moment for cleanup
    Write-Host "Wait for OS process clean up..."
    Start-Sleep -Seconds 2

    if ($(Test-Path $iniFilePath1) -and (Test-Path $iniFilePath2)) 
    {
        # Read the ini files
        $fileContent1 = Get-Content $iniFilePath1 -Raw
        $fileContent2 = Get-Content $iniFilePath2 -Raw
        
        # Replace number in ini files
        #   FreqCalibration=1.00828283
        $replacementPattern = '(FreqCalibration=)\d\.\d+'
        $newContent1 = $fileContent1 -replace $replacementPattern, "`${1}$skimSrvCalibration"
        $newContent2 = $fileContent2 -replace $replacementPattern, "`${1}$skimSrvCalibration"

        if (-not $dryRun) 
        {
            $newContent1 | Set-Content $iniFilePath1
            $newContent2 | Set-Content $iniFilePath2
        } 
        else 
        {
            Write-Host "*** Dry run mode is on, not updating SkimSrv ini files."
        }

        # Announce success
        Write-Host "Successfully updated $iniFile1 ini with new calibration value: $skimSrvCalibration"
        Write-Host "Successfully updated $iniFile2 ini with new calibration value: $skimSrvCalibration"
    }
    else 
    {
        Write-Error "Ini file not found: $iniFilePath1"
        exit 1
    }

    if (Test-Path $configFilePath)
    {
        # Read the CWSL_DIGI config file
        $fileContent3 = Get-Content $configFilePath -Raw

        # Replace calibration factor in config file
        # freqcalibration=1.000000000
        $replacementPattern3 = '(freqcalibration=)\d\.\d+'
        $newContent3 = $fileContent3 -replace $replacementPattern3, "`${1}$cwsldigiCalibration"
        if (-not $dryRun) 
        {
            $newContent3 | Set-Content $configFilePath
        } 
        else 
        {
            Write-Host "*** Dry run mode is on, not updating CWSL_DIGI config file."
        }

        # Announce success
        Write-Host "Successfully updated $configFile with new calibration value: $cwsldigiCalibration"
    }
    else 
    {
        Write-Error "Ini file not found: $configFilePath"
        exit 1
    }
    
    # Start applications again
    Write-Host "Starting $skimsrvExe1..."
    Start-Process -WorkingDirectory $skimsrvPath1 -FilePath $skimsrvExe1 -WindowStyle Minimized
    Start-Sleep -Seconds 2

    Write-Host "Starting $skimsrvExe2..."
    Start-Process -WorkingDirectory $skimsrvPath2 -FilePath $skimsrvExe2 -WindowStyle Minimized

    Write-Host "Starting CWSL_DIGI..."
    Start-Process -WorkingDirectory $cwslPath -FilePath $cwslExe -WindowStyle Minimized

    Write-Host "Update complete. Applications restarted."
}
catch 
{
    Write-Error "An error occurred: $($_.Exception.Message)"
    exit 1
}
