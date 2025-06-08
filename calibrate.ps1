# PowerShell script to parse web page, calculate, and update text file

# Configuration
$webUrl = "https://sm7iun.se/rbn/analytics/"
$inipath1 = $env:APPDATA + "\Afreet\Products\SkimSrv\"
$inipath2 = $env:APPDATA + "\Afreet\Products\SkimSrv2\"
$inipath3 = "C:\CWSL_DIGI\"

$inifile1 = "SkimSrv.ini"
$inifile2 = "SkimSrv2.ini"
$inifile3 = "config.ini"

$inifilepath1 = $inipath1 + $inifile1
$inifilepath2 = $inipath2 + $inifile2
$inifilepath3 = $inipath3 + $inifile3

Write-Host "--------------------------------------------------"
Write-Host "Now is" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

try {
    # Parse the web page for an adjustment factor and ini file
    # for current calibration factor.
    
    # Format of line 
    # FreqCalibration=1.00828283
    Write-Host "Fetching data from: $inifilepath1"
	$inicontent = Get-Content $inifilepath1 -Raw
	$inimatch = [regex]::Match($inicontent, 'FreqCalibration=(\d+\.\d+)?') 

    # Format of line
    #   SM7IUN      +0.1   3999   1.000000099
    Write-Host "Fetching data from: $webUrl"
    $webContent = Invoke-WebRequest -Uri $webUrl -UseBasicParsing
	$callsign = "SM7IUN"
    $webmatch = [regex]::Match($webContent.Content, $callsign + ' +[+-]\d\.\d+ +\d+ +(\d\.\d+)')
    
	Write-Host ""
	
    if ($webMatch.Success -and $inimatch.Success) 
    {
		$inicalibration = [double]$inimatch.Groups[1].Value
		Write-Host "Current calibration factor: $inicalibration"

        $newCalibration = [double]$webmatch.Groups[1].Value
        Write-Host "Adjustment factor: $newCalibration"

		Write-Host ""

        # Calculate new calibration factors
        $skimsrvcalibration = $newCalibration * $inicalibration
        Write-Host "New calibration factor for SkimSrv: $skimsrvcalibration"
		$cwsldigicalibration = 1 / $skimsrvcalibration
        Write-Host "New calibration factor for CWSL_DIGI: $cwsldigicalibration"
        
		Write-Host ""
    }
    else 
    {
        Write-Error "Failed to read web and/or current calibration factor from ini file."
    }

    # Stop the applications
    Write-Host "Stopping SkimSrv, SkimSrv2, and CWSL_DIGI"

    # Stop SkimSrv instances
    Stop-Process -Name "SkimSrv*" -Force -ErrorAction SilentlyContinue 
    
    # Close CWSL_DIGI process including subprocesses
    Stop-Process -Name "CWSL*" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "jt9*" -Force -ErrorAction SilentlyContinue

    # Wait a moment for cleanup
    Start-Sleep -Seconds 5

    if ($(Test-Path $inifilepath1) -and (Test-Path $inifilepath2)) 
    {
        # Read the ini files
#        Write-Host "Updating files: $inifile1 and $inifile2"
        $fileContent1 = Get-Content $inifilepath1 -Raw
        $fileContent2 = Get-Content $inifilepath2 -Raw
        
        # Replace number in ini files
        #   FreqCalibration=1.00828283
        $replacementPattern = '(FreqCalibration=)\d\.\d+'
        $newContent1 = $fileContent1 -replace $replacementPattern, "`${1}$skimsrvcalibration"
        $newContent2 = $fileContent2 -replace $replacementPattern, "`${1}$skimsrvcalibration"

        # Write back to files
        $newContent1 | Set-Content $inifilepath1
        $newContent2 | Set-Content $inifilepath2

        # Announce success
        Write-Host "Successfully updated SkimSrv1 ini with new calibration value: $skimsrvcalibration"
        Write-Host "Successfully updated SkimSrv2 ini with new calibration value: $skimsrvcalibration"
    }
    else 
    {
        Write-Error "Ini file not found: $inifilepath1"
    }

    if (Test-Path $inifilepath3) {
        # Read the config file
#        Write-Host "Updating file: $inifile3"
        $fileContent3 = Get-Content $inifilepath3 -Raw

        # Replace calibration factor in text file
        # freqcalibration=1.000000000
        $replacementPattern3 = '(freqcalibration=)\d\.\d+'
        $newContent3 = $fileContent3 -replace $replacementPattern3, "`${1}$cwsldigicalibration"

        # Write back to files
        $newContent3 | Set-Content $inifilepath3

        # Announce success
        Write-Host "Successfully updated CWSL_DIGI ini with new calibration value: $cwsldigicalibration"
    }
    else 
    {
        Write-Error "Ini file not found: $inifilepath3"
    }
    
    # Start applications again
    Write-Host "Starting SkimSrv..."
    Start-Process -WorkingDirectory "C:\Program Files (x86)\Afreet\SkimSrv" -FilePath "SkimSrv.exe" -WindowStyle Minimized
    Start-Sleep -Seconds 2

    Write-Host "Starting SkimSrv2..."
    Start-Process -WorkingDirectory "C:\Program Files (x86)\Afreet\SkimSrv2" -FilePath "SkimSrv2.exe" -WindowStyle Minimized

    Write-Host "Starting CWSL_DIGI..."
    Start-Process -WorkingDirectory "C:\CWSL_DIGI" -FilePath "CWSL_DIGI.exe" -WindowStyle Minimized

    Write-Host "Update complete. Applications restarted."
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}
