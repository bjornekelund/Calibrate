# PowerShell script to update frequency calibration of CWSL_DIGI and two instances of SkimSrv
# Reads skew data from sm7iun.se/rbn/analytics and updates the ini files accordingly.

param([switch]$DryRun, [switch]$Verbose)

# Configuration for this specific installation

$callsign = "SM7IUN"  # Callsign to look for in the web page

$skimsrv1 = $true  # Set to $true if you have SkimSrv installed
$skimsrv2 = $true  # Set to $true if you have two instances of SkimSrv installed
$rttyskimserv1 = $false  # Set to $true if one instance of RttySkimServ is installed
$rttyskimserv2 = $false  # Set to $true if you have instances of RttySkimServ installed
$cwsldigi = $true  # Set to $true if you are using CWSL_DIGI

$webUrl = "https://sm7iun.se/rbn/analytics"

$iniPath1 = $env:APPDATA + "\Afreet\Products\SkimSrv\"
$iniPath2 = $env:APPDATA + "\Afreet\Products\SkimSrv2\"
$iniPath3 = $env:APPDATA + "\Afreet\Products\RttySkimServ1\"
$iniPath4 = $env:APPDATA + "\Afreet\Products\RttySkimServ2\"
$configPath = "C:\CWSL_DIGI\"

$iniFile1 = "SkimSrv.ini"
$iniFile2 = "SkimSrv2.ini"
$iniFile3 = "RttySkimServ1.ini"
$iniFile4 = "RttySkimServ2.ini"
$configFile = "config.ini"

$skimsrvPath1 = "C:\Program Files (x86)\Afreet\SkimSrv\"
$skimsrvPath2 = "C:\Program Files (x86)\Afreet\SkimSrv2\"
$skimsrvPath3 = "C:\Program Files (x86)\Afreet\RttySkimServ1\"
$skimsrvPath4 = "C:\Program Files (x86)\Afreet\RttySkimServ2\"
$cwslPath = "C:\CWSL_DIGI"

$skimsrvExe1 = "SkimSrv.exe"
$skimsrvExe2 = "SkimSrv2.exe"
$skimsrvExe3 = "RttySkimServ1.exe"
$skimsrvExe4 = "RttySkimServ2.exe"
$cwslExe = "CWSL_DIGI.exe"

# End of configuration

if ($Verbose) { Write-Host "Verbose mode enabled" }
if ($DryRun) { Write-Host "Dry run mode enabled" }

$iniFilePath1 = $iniPath1 + $iniFile1
$iniFilePath2 = $iniPath2 + $iniFile2
$iniFilePath3 = $iniPath3 + $iniFile3
$iniFilePath4 = $iniPath4 + $iniFile4
$configFilePath = $configPath + $configFile

$ProgressPreference = 'SilentlyContinue' # Show no progress bar

Write-Host "--------------------------------------------------"
Write-Host "Execution time is" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

try 
{
    # Parse SkimSrv ini file for current calibration factor
    # If no SkimSrv ini file is found, try RttySkimServ ini file
    # Assume the found value to be for all SkimSrv and RttySkimSrv instances
    # Format of line 
    # FreqCalibration=1.00828283
    # FreqCalibration=1
    if ($skimsrv1 -and (Test-Path $iniFilePath1)) 
    {
        Write-Host "Reading SkimSrv ini file: $iniFilePath1"
        $iniContent = Get-Content $iniFilePath1 -Raw
        $usedIniFile = $iniFile1
    } 
    elseif ($rttyskimserv1 -and (Test-Path $iniFilePath3)) 
    {
        Write-Host "Reading RttySkimServ ini file: $iniFilePath3"
        $iniContent = Get-Content $iniFilePath3 -Raw
        $usedIniFile = $iniFile3
    } 
    else 
    {
        Write-Host "No skimmer ini file available. Exiting."
        exit 1
    }

    $iniMatch = [regex]::Match($iniContent, 'FreqCalibration=(0\.\d+|1(\.\d+)?)') 

    if ($iniMatch.Success) 
    {
        $inicalibration = [double]$iniMatch.Groups[1].Value
        Write-Host "Current skimmer calibration factor is: $inicalibration"
    }
    else 
    {
        Write-Host "Failed to to find a valid FreqCalibration= line in from $usedIniFile. Exiting."
        exit 1
    }

    # Parse CWSL_DIGI config file for current calibration factor
    # Format of line 
    # freqcalibration=1.00828283
    if ($cwsldigi -and (Test-Path $configFilePath)) 
    {
        $configContent = Get-Content $configFilePath -Raw
        $configMatch = [regex]::Match($configContent, 'freqcalibration=(0\.\d+|1(\.\d+)?)')

        if ($configMatch.Success) 
        {
            $configCalibration = [double]$configMatch.Groups[1].Value
            Write-Host "Current CWSL_DIGI calibration factor is: $configCalibration"
        }
        else 
        {
            Write-Host "Failed to find a valid freqcalibration= line in $configFile. Exiting"
            exit 1
        }
    }     

    # Scrape web page for adjustment factor
    # If the skimmer has produced enough spots, the web page contains a line with the desired statistics/analytics
    # The callsign has optional trailing asterisk, followed by frequency error in ppm, spot count, and adjustment factor
    # Also find the last updated time of the web page to confirm that the data is fresh
    # Make sure Invoke-WebRequest pushes through the cache
        # Format of line
    #   SM7IUN*     +0.1   3999   1.000000099
    # Last updated 2025-06-09 00:16:23 UTC
    $headers = @{'Cache-Control'='no-cache,no-store,must-revalidate';'Pragma'='no-cache';'Expires'='0'}
    $webContent = Invoke-WebRequest -Uri $webUrl -UseBasicParsing -Headers $headers
    $webMatch = [regex]::Match($webContent.Content, $callsign + '\*? +[+-]\d\.\d+ +\d+ +([01]\.\d+)')
    $webTimeMatch = [regex]::Match($webContent.Content, 'Last updated +(20\d{2}-\d{1,2}-\d{1,2} +\d{1,2}:\d{2}:\d{2})')

    if (-not $webMatch.Success) 
    {
        Write-Host "Not enough spots reported for $callsign. Skimmer may be down. Exiting."
        exit 0
    }

    if ($webTimeMatch.Success) 
    {
        $lastUpdated = $webTimeMatch.Groups[1].Value
        $webCalibration = [double]$webMatch.Groups[1].Value
        Write-Host "Absolute adjustment factor from $webUrl at $lastUpdated from is: $webCalibration"
        # Since there are statistical variations adjustment factor, only do a gradual adjustment
        $newCalibration = [Math]::Round([System.Math]::Pow($webCalibration, 0.6), 9)
        Write-Host "Used adjustment factor is: $newCalibration"
    }
    else 
    {
        Write-Host "Failed to find last update time in $webUrl. Exiting."
        exit 1
    }

    # Calculate new calibration factors
    # Round to 9 decimal places like in the web page which is an overkill of accuracy
    $skimSrvCalibration = [Math]::Round($newCalibration * $inicalibration, 9)
#    Write-Host "New calibration factor for SkimSrv: $skimSrvCalibration"
    $cwsldigiCalibration = [Math]::Round(1.0 / ($newCalibration * $inicalibration), 9)
#    Write-Host "New calibration factor for CWSL_DIGI: $cwsldigiCalibration"
    
    # Stop the applications
    Write-Host "Stopping skimmer processes..."

    # Stop SkimSrv instances
    Stop-Process -Name "SkimSrv*" -Force -ErrorAction SilentlyContinue 
    
    # Stop RttySkimServ instance if there is one
    Stop-Process -Name "RttySkimServ*" -Force -ErrorAction SilentlyContinue

    # Close CWSL_DIGI including subprocesses
    Stop-Process -Name "CWSL*" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "jt9*" -Force -ErrorAction SilentlyContinue

    # Wait a moment for cleanup
    Write-Host "Wait for OS process clean up..."
    Start-Sleep -Seconds 2

    if ($skimsrv1) 
    {
        if (Test-Path $iniFilePath1) 
        {
            # Read the ini files
            $fileContent1 = Get-Content $iniFilePath1 -Raw
            
            # Replace number in ini files
            #   FreqCalibration=1.00828283
            $replacementPattern = '(FreqCalibration=)\d\.\d+'
            $newContent1 = $fileContent1 -replace $replacementPattern, "`${1}$skimSrvCalibration"

            if (-not $DryRun) 
            {
                $newContent1 | Set-Content $iniFilePath1
                Write-Host "Successfully updated $iniFile1 ini with new calibration value: $skimSrvCalibration"
            } 
            else 
            {
                Write-Host "Did not update $iniFile1 ini with new calibration value: $skimSrvCalibration"
            }

        }
        else 
        {
            Write-Error "$iniFile1 file not found."
            exit 1
        }
    }

    if ($skimsrv2) 
    {
        if (Test-Path $iniFilePath2)
        {
            # Read the ini files
            $fileContent2 = Get-Content $iniFilePath2 -Raw
            
            # Replace number in ini files
            #   FreqCalibration=1.00828283
            $replacementPattern = '(FreqCalibration=)\d\.\d+'
            $newContent2 = $fileContent2 -replace $replacementPattern, "`${1}$skimSrvCalibration"

            if (-not $DryRun) 
            {
                $newContent2 | Set-Content $iniFilePath2
                Write-Host "Successfully updated $iniFile2 ini with new calibration value: $skimSrvCalibration"
            } 
            else 
            {
                Write-Host "Did not update $iniFile2 ini with new calibration value: $skimSrvCalibration"
            }
        }
        else 
        {
            Write-Error "$iniFile2 file not found."
            exit 1
        }
    }


    if ($rttyskimserv1)
    {
        if (Test-Path $iniFilePath3)
        {
            # Read the ini files
            $fileContent3 = Get-Content $iniFilePath3 -Raw
            
            # Replace number in ini files
            #   FreqCalibration=1.00828283
            $replacementPattern = '(FreqCalibration=)\d\.\d+'
            $newContent3 = $fileContent3 -replace $replacementPattern, "`${1}$skimSrvCalibration"

            if (-not $DryRun) 
            {
                $newContent3 | Set-Content $iniFilePath3
                Write-Host "Successfully updated $iniFile3 ini with new calibration value: $skimSrvCalibration"
            } 
            else 
            {
                Write-Host "Did not update $iniFile3 ini with new calibration value: $skimSrvCalibration"
            }
        }
        else 
        {
            Write-Host "$iniFile3 file not found."
        }
    }

    if ($rttyskimserv2)
    {
        if (Test-Path $iniFilePath4)
        {
            # Read the ini files
            $fileContent4 = Get-Content $iniFilePath4 -Raw
            
            # Replace number in ini files
            #   FreqCalibration=1.00828283
            $replacementPattern = '(FreqCalibration=)\d\.\d+'
            $newContent4 = $fileContent4 -replace $replacementPattern, "`${1}$skimSrvCalibration"

            if (-not $DryRun) 
            {
                $newContent4 | Set-Content $iniFilePath4
                Write-Host "Successfully updated $iniFile4 ini with new calibration value: $skimSrvCalibration"
            } 
            else 
            {
                Write-Host "Did not update $iniFile4 ini with new calibration value: $skimSrvCalibration"
            }
        }
        else 
        {
            Write-Host "$iniFile4 file not found."
        }
    }

    if ($cwsldigi) 
    {
        if (Test-Path $configFilePath )
        {
            # Read the CWSL_DIGI config file
            $fileContent3 = Get-Content $configFilePath -Raw

            # Replace calibration factor in config file
            # freqcalibration=1.000000000
            $replacementPattern3 = '(freqcalibration=)\d\.\d+'
            $newContent3 = $fileContent3 -replace $replacementPattern3, "`${1}$cwsldigiCalibration"
            if (-not $DryRun) 
            {
                $newContent3 | Set-Content $configFilePath
            Write-Host "Successfully updated $configFile with new calibration value: $cwsldigiCalibration"
            } 
            else 
            {
                Write-Host "Did not update $configFile with new calibration value: $cwsldigiCalibration"
            }
        }
        else 
        {
            Write-Error "$configFile not found."
            exit 1
        }
    }
    
    # Start applications again
    if ($skimsrv1)
    {
        Write-Host "Starting $skimsrvExe1..."
        Start-Process -WorkingDirectory $skimsrvPath1 -FilePath $skimsrvExe1 -WindowStyle Minimized
    }

    if ($skimsrv2)
    {
        Start-Sleep -Seconds 2
        Write-Host "Starting $skimsrvExe2..."
        Start-Process -WorkingDirectory $skimsrvPath2 -FilePath $skimsrvExe2 -WindowStyle Minimized
    }

    if ($rttyskimserv1)
    {
        Start-Sleep -Seconds 2
        Write-Host "Starting $skimsrvExe3..."
        Start-Process -WorkingDirectory $skimsrvPath3 -FilePath $skimsrvExe3 -WindowStyle Minimized
    }

    if ($rttyskimserv2)
    {
        Start-Sleep -Seconds 2
        Write-Host "Starting $skimsrvExe4..."
        Start-Process -WorkingDirectory $skimsrvPath4 -FilePath $skimsrvExe4 -WindowStyle Minimized
    }

    if ($cwsldigi)
    {
        Write-Host "Starting CWSL_DIGI..."
        Start-Process -WorkingDirectory $cwslPath -FilePath $cwslExe -WindowStyle Minimized
    }

    Write-Host "Update complete. Applications restarted."
}
catch 
{
    Write-Error "An error occurred: $($_.Exception.Message)"
    exit 1
}
