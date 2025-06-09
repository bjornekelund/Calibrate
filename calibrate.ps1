# PowerShell script to update frequency calibration of CWSL_DIGI and two instances of SkimSrv
# Reads skew data from sm7iun.se/rbn/analytics and updates the ini files accordingly.

$dryrun = $false  # Set to $false to actually update the files

$ProgressPreference = 'SilentlyContinue' # Show no progress bar

$callsign = "SM7IUN"

$webUrl = "https://sm7iun.se/rbn/analytics"
$inipath1 = $env:APPDATA + "\Afreet\Products\SkimSrv\"
$inipath2 = $env:APPDATA + "\Afreet\Products\SkimSrv2\"
$inipath3 = "C:\CWSL_DIGI\"

$inifile1 = "SkimSrv.ini"
$inifile2 = "SkimSrv2.ini"
$inifile3 = "config.ini"

$inifilepath1 = $inipath1 + $inifile1
$inifilepath2 = $inipath2 + $inifile2
$inifilepath3 = $inipath3 + $inifile3

$exepath1 = "C:\Program Files (x86)\Afreet\SkimSrv"
$exepath2 = "C:\Program Files (x86)\Afreet\SkimSrv2"
$exepath3 = "C:\CWSL_DIGI"

$exefile1 = "SkimSrv.exe"
$exefile2 = "SkimSrv2.exe"
$exefile3 = "CWSL_DIGI.exe"

Write-Host "--------------------------------------------------"
Write-Host "Execution time is" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

try {
    # Parse SkimSrv ini file for current calibration factor
    # Both ini files should have the same calibration factor
    # Format of line 
    # FreqCalibration=1.00828283
	$inicontent = Get-Content $inifilepath1 -Raw
	$inimatch = [regex]::Match($inicontent, 'FreqCalibration=([01]\.\d+)?') 

    if ($inimatch.Success) 
    {
		$inicalibration = [double]$inimatch.Groups[1].Value
		Write-Host "Current SkimSrv calibration factor is: $inicalibration"
    }
    else 
    {
        Write-Error "Failed to read CWSL_DIGI calibration factor from $inifile1"
        exit 1
    }

    # Parse CWSL_DIGI config file for current calibration factor
    # Format of line 
    # freqcalibration=1.00828283
	$configcontent = Get-Content $inifilepath3 -Raw
	$configmatch = [regex]::Match($configcontent, 'freqcalibration=([01]\.\d+)?') 

    if ($configmatch.Success) 
    {
		$configcalibration = [double]$configmatch.Groups[1].Value
		Write-Host "Current CWSL_DIGI calibration factor is: $configcalibration"
    }
    else 
    {
        Write-Error "Failed to read calibration factor from $inifile3"
        exit 1
    }

    # Parse web page for adjustment factor
    # The web page should contain a line with the callsign and adjustment factor
    # Callsign has optional asterisk, followed by frequency error in ppm, spot count, and adjustment factor
    # Format of line
    #   SM7IUN*     +0.1   3999   1.000000099
    $webContent = Invoke-WebRequest -Uri $webUrl -UseBasicParsing
    $webmatch = [regex]::Match($webContent.Content, $callsign + '\*? +[+-]\d\.\d+ +\d+ +([01]\.\d+)')

    if ($webMatch.Success) 
    {
        $newCalibration = [double]$webmatch.Groups[1].Value
        Write-Host "Adjustment factor from $webUrl is: $newCalibration"
    }
    else 
    {
        Write-Error "Failed to find adjustment factor in $webUrl"
        exit 1
    }

    # Calculate new calibration factors
    $skimsrvcalibration = [Math]::Round($newCalibration * $inicalibration, 9)
#    Write-Host "New calibration factor for SkimSrv: $skimsrvcalibration"
    $cwsldigicalibration = [Math]::Round(1.0 / ($newCalibration * $inicalibration), 9)
#    Write-Host "New calibration factor for CWSL_DIGI: $cwsldigicalibration"
    
    # Stop the applications
    Write-Host "Stopping $exefile1, $exefile2, and $exefile3..."

    # Stop SkimSrv instances
    Stop-Process -Name "SkimSrv*" -Force -ErrorAction SilentlyContinue 
    
    # Close CWSL_DIGI process including subprocesses
    Stop-Process -Name "CWSL*" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "jt9*" -Force -ErrorAction SilentlyContinue

    # Wait a moment for cleanup
    Write-Host "Wait for OS process clean up..."
    Start-Sleep -Seconds 2

    if ($(Test-Path $inifilepath1) -and (Test-Path $inifilepath2)) 
    {
        # Read the ini files
        $fileContent1 = Get-Content $inifilepath1 -Raw
        $fileContent2 = Get-Content $inifilepath2 -Raw
        
        # Replace number in ini files
        #   FreqCalibration=1.00828283
        $replacementPattern = '(FreqCalibration=)\d\.\d+'
        $newContent1 = $fileContent1 -replace $replacementPattern, "`${1}$skimsrvcalibration"
        $newContent2 = $fileContent2 -replace $replacementPattern, "`${1}$skimsrvcalibration"
        if (-not $dryrun) {
            $newContent1 | Set-Content $inifilepath1
            $newContent2 | Set-Content $inifilepath2
        } else {
            Write-Host "*** Dry run mode is on, not updating SkimSrv ini files."
        }

        # Announce success
        Write-Host "Successfully updated $inifile1 ini with new calibration value: $skimsrvcalibration"
        Write-Host "Successfully updated $inifile2 ini with new calibration value: $skimsrvcalibration"
    }
    else 
    {
        Write-Error "Ini file not found: $inifilepath1"
        exit 1
    }

    if (Test-Path $inifilepath3) {
        # Read the CWSL_DIGI config file
        $fileContent3 = Get-Content $inifilepath3 -Raw

        # Replace calibration factor in config file
        # freqcalibration=1.000000000
        $replacementPattern3 = '(freqcalibration=)\d\.\d+'
        $newContent3 = $fileContent3 -replace $replacementPattern3, "`${1}$cwsldigicalibration"
        if (-not $dryrun) {
            $newContent3 | Set-Content $inifilepath3
        } else {
            Write-Host "*** Dry run mode is on, not updating CWSL_DIGI config file."
        }

        # Announce success
        Write-Host "Successfully updated $inifile3 with new calibration value: $cwsldigicalibration"
    }
    else 
    {
        Write-Error "Ini file not found: $inifilepath3"
        exit 1
    }
    
    # Start applications again
    Write-Host "Starting $exefile1..."
    Start-Process -WorkingDirectory $exepath1 -FilePath $exefile1 -WindowStyle Minimized
    Start-Sleep -Seconds 2

    Write-Host "Starting $exefile2..."
    Start-Process -WorkingDirectory $exepath2 -FilePath $exefile2 -WindowStyle Minimized

    Write-Host "Starting CWSL_DIGI..."
    Start-Process -WorkingDirectory $exepath3 -FilePath $exefile3 -WindowStyle Minimized

    Write-Host "Update complete. Applications restarted."
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    exit 1
}
