# PowerShell script to adjust frequency calibration of up to two instances 
# of SkimSrv, up to two instances of RTTY SkimSrv, and CWSL_DIGI.
# Uses skew data published at https://sm7iun.se/rbn/analytics

param([switch]$DryRun, [switch]$Verbose)

# -----------------------------------------------------------
# Configuration for this specific installation
# This is the only part of the script that should be edited

$callsign = ""  # Skimmer callsign. Empty means use the callsign of the first SkimSrv instance.

$skimsrv1 = $true        # Set to $true if you have SkimSrv installed
$skimsrv2 = $true        # Set to $true if you have two instances of SkimSrv installed
$rttyskimserv1 = $false  # Set to $true if one instance of RttySkimServ is installed
$rttyskimserv2 = $false  # Set to $true if you have two instances of RttySkimServ installed
$cwsldigi = $true        # Set to $true if you are using CWSL_DIGI

# Location of ini files and CWSL_DIGI config file
$iniPath1 = $env:APPDATA + "\Afreet\Products\SkimSrv\"
$iniPath2 = $env:APPDATA + "\Afreet\Products\SkimSrv2\"
$iniPath3 = $env:APPDATA + "\Afreet\Products\RttySkimServ1\"
$iniPath4 = $env:APPDATA + "\Afreet\Products\RttySkimServ2\"
$configPath = "C:\CWSL_DIGI\"

# Naming of ini and config files
$iniFile1 = "SkimSrv.ini"
$iniFile2 = "SkimSrv2.ini"
$iniFile3 = "RttySkimServ1.ini"
$iniFile4 = "RttySkimServ2.ini"
$configFile = "config.ini"

# Installation paths for SkimSrv, RttySkimSrv, and CWSL_DIGI
$skimsrvPath1 = "C:\Program Files (x86)\Afreet\SkimSrv\"
$skimsrvPath2 = "C:\Program Files (x86)\Afreet\SkimSrv2\"
$skimsrvPath3 = "C:\Program Files (x86)\Afreet\RttySkimServ1\"
$skimsrvPath4 = "C:\Program Files (x86)\Afreet\RttySkimServ2\"
$cwslPath = "C:\CWSL_DIGI\"

# Naming of executables
$skimsrvExe1 = "SkimSrv.exe"
$skimsrvExe2 = "SkimSrv2.exe"
$skimsrvExe3 = "RttySkimServ1.exe"
$skimsrvExe4 = "RttySkimServ2.exe"
$cwslExe = "CWSL_DIGI.exe"

# End of configuration section
# -----------------------------------------------------------

# Normally no parts below should require editing
# If you find that you need to edit something below, please let the author know

$webUrl = "https://sm7iun.se/rbn/analytics"

if ($Verbose) { Write-Host "Verbose mode enabled" }
if ($DryRun -and $Verbose) { Write-Host "Dry run mode enabled" }

$iniFilePath1 = $iniPath1 + $iniFile1
$iniFilePath2 = $iniPath2 + $iniFile2
$iniFilePath3 = $iniPath3 + $iniFile3
$iniFilePath4 = $iniPath4 + $iniFile4
$configFilePath = $configPath + $configFile

# Show now progress bars for web requests and file operations
$ProgressPreference = 'SilentlyContinue'

Write-Host "--------------------------------------------------"
Write-Host "Starting update at $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"

try 
{
    # Parse SkimSrv ini file for current calibration factor
    # If no SkimSrv ini file is found amd RttySkimSrv is used, try RttySkimServ's ini file
    # Assume the found factor to be the one used by all SkimSrv and RttySkimSrv instances
    # Format of line 
    # FreqCalibration=1.00828283
    # FreqCalibration=1
    if ($skimsrv1 -and (Test-Path $iniFilePath1)) 
    {
        if ($Verbose) { Write-Host "Reading SkimSrv ini file: $iniFilePath1" }
        $iniContent = Get-Content $iniFilePath1 -Raw
        $usedIniFile = $iniFile1
    } 
    elseif ($rttyskimserv1 -and (Test-Path $iniFilePath3)) 
    {
        if ($Verbose) { Write-Host "Reading RttySkimServ ini file: $iniFilePath3" }
        $iniContent = Get-Content $iniFilePath3 -Raw
        $usedIniFile = $iniFile3
    } 
    else 
    {
        Write-Host "No skimmer ini file available. Exiting."
        exit 1
    }

    $iniMatch = [regex]::Match($iniContent, '\sfreqcalibration=(0\.\d+|1(\.\d+)?)', 
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) 
    
    if ($callsign -eq "")
    {
        $iniCall = [regex]::Match($iniContent, '\scall=([A-Z0-9]+)', 
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        if ($iniCall.Success) {
            $callsign = $iniCall.Groups[1].Value
            Write-Host "Using callsign from ini file: $callsign"
        }
        else {
            Write-Host "Failed to to find a valid Call= line in from $usedIniFile. Exiting."
            exit 1
        }
    }


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
        $configMatch = [regex]::Match($configContent, '\sfreqcalibration=(0\.\d+|1(\.\d+)?)', 
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        if ($configMatch.Success) 
        {
            $configCalibration = [double]$configMatch.Groups[1].Value
            Write-Host "Current CWSL_DIGI calibration factor is: $configCalibration"
        }
        else 
        {
            Write-Host "Failed to find a valid freqcalibration= line in $configFile. Is it commented out? Exiting"
            exit 1
        }
    }     

    # Scrape web page for adjustment factor
    # If the skimmer has produced enough spots, the web page contains a line with the desired statistics/analytics
    # The callsign has optional trailing asterisk, followed by frequency error in ppm, spot count, and adjustment factor
    # Also find the update time of the web page to confirm that the data is fresh
    # Use header options to make sure Invoke-WebRequest pushes through the cache
        # Format of line
    #   SM7IUN*     +0.1   3999   1.000000099
    # Last updated 2025-06-09 00:16:23 UTC
    # Make sure to push through cache
    $headers = @{'Cache-Control' = 'no-cache, no-store, must-revalidate'; 'Pragma' = 'no-cache'; 'Expires' = '0'}
    # Get content from web page
    $webContent = Invoke-WebRequest -Uri $webUrl -UseBasicParsing -Headers $headers
    # Look for relevant line with skew data
    $webMatch = [regex]::Match($webContent.Content, $callsign + '\*? +[+-]\d\.\d+ +(\d+) +([01]\.\d+)')
    # Look for time stamp
    $webTimeMatch = [regex]::Match($webContent.Content, 'Last updated +(20\d{2}-\d{1,2}-\d{1,2} +\d{1,2}:\d{2}:\d{2})')

    if (-not $webMatch.Success) 
    {
        Write-Host "No spots reported for $callsign. Skimmer may be down. Exiting."
        exit 0
    }

    $spotCount = [int]$webMatch.Groups[1].Value

    if ($spotCount -lt 400) 
    {
        Write-Host "Only $spotCount spots reported for $callsign meaning skew estimate is too unreliable. Exiting."
        exit 0
    }
    
    if ($webTimeMatch.Success) 
    {
        $lastUpdated = $webTimeMatch.Groups[1].Value
        $webCalibration = [double]$webMatch.Groups[2].Value
        Write-Host "Reading skew data from $webUrl published at $lastUpdated UTC"
        $webskew = ($webCalibration - 1.0) * 1000000.0
        $absSkew = [Math]::Abs($webskew).ToString("F2")
        $direction = if ($webskew -gt 0) { "high" } else { "low" }        
        Write-Host "Suggested adjustment factor is $webCalibration meaning reports are on average $absSkew ppm too $direction"
        # Since there are statistical variations adjustment factor, do not compensate fully but do a gradual adjustment
        $newCalibration = [Math]::Round([System.Math]::Pow($webCalibration, 0.5), 9)
        $skewadjustment = ((1.0 - $newCalibration) * 1000000.0).ToString("F2")

        Write-Host "The moderated adjustment factor is $newCalibration corresponding to an adjustment of reports of $skewadjustment ppm"
    }
    else 
    {
        Write-Host "Failed to find last update time in $webUrl. Exiting."
        exit 1
    }

    # Calculate new calibration factors
    # Round to 9 decimal places like in the web page which is an overkill of accuracy
    $skimSrvCalibration = [Math]::Round($newCalibration * $inicalibration, 9)
    if ($Verbose) { Write-Host "New calibration factor for SkimSrv: $skimSrvCalibration" }
    $cwsldigiCalibration = [Math]::Round(1.0 / ($newCalibration * $inicalibration), 9)
    if ($Verbose) { Write-Host "New calibration factor for CWSL_DIGI: $cwsldigiCalibration" }
    
    # Stop the applications
    if ($Verbose) { Write-Host "Stopping skimmer processes..." }

    # Stop SkimSrv instances
    Stop-Process -Name "SkimSrv*" -Force -ErrorAction SilentlyContinue 
    
    # Stop RttySkimServ instance if there is one
    Stop-Process -Name "RttySkimServ*" -Force -ErrorAction SilentlyContinue

    # Close CWSL_DIGI including subprocesses
    Stop-Process -Name "CWSL*" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "jt9*" -Force -ErrorAction SilentlyContinue

    # Wait a moment for cleanup
    if ($Verbose) { Write-Host "Wait for OS process clean up..." }
    Start-Sleep -Seconds 4

    # Regular expression replacement pattern to update ini and config files
    # Case insensitive since CWSL_DIGI is
    $replacementPattern = [regex]::new("(freqcalibration=)(1(\.\d+)?|0\.\d+)", 
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    if ($skimsrv1) 
    {
        if (Test-Path $iniFilePath1) 
        {
            $newContent1 = (Get-Content $iniFilePath1 -Raw) -replace $replacementPattern, "`${1}$skimSrvCalibration"

            if (-not $DryRun) 
            {
                # Replace calibration factor with new value
                $newContent1 | Set-Content $iniFilePath1
                Write-Host "Successfully updated $iniFile1 with new calibration factor: $skimSrvCalibration"
            } 
            else 
            {
                Write-Host "Did not update $iniFile1 with new calibration factor: $skimSrvCalibration"
            }
        }
        else 
        {
            Write-Host "$iniFile1 not found. Exiting."
            exit 1
        }
    }

    if ($skimsrv2) 
    {
        if (Test-Path $iniFilePath2)
        {
            $newContent2 = (Get-Content $iniFilePath2 -Raw) -replace $replacementPattern, "`${1}$skimSrvCalibration"

            if (-not $DryRun) 
            {
                # Replace calibration factor with new value
                $newContent2 | Set-Content $iniFilePath2
                Write-Host "Successfully updated $iniFile2 with new calibration factor: $skimSrvCalibration"
            } 
            else 
            {
                Write-Host "Did not update $iniFile2 with new calibration factor: $skimSrvCalibration"
            }
        }
        else 
        {
            Write-Host "$iniFile2 not found. Exiting."
            exit 1
        }
    }

    if ($rttyskimserv1)
    {
        if (Test-Path $iniFilePath3)
        {
            $newContent3 = (Get-Content $iniFilePath3 -Raw) -replace $replacementPattern, "`${1}$skimSrvCalibration"

            if (-not $DryRun) 
            {
                # Replace calibration factor with new factor
                $newContent3 | Set-Content $iniFilePath3
                Write-Host "Successfully updated $iniFile3 with new calibration factor: $skimSrvCalibration"
            } 
            else 
            {
                Write-Host "Did not update $iniFile3 with new calibration factor: $skimSrvCalibration"
            }
        }
        else 
        {
            Write-Host "$iniFile3 not found. Exiting."
            exit 1
        }
    }

    if ($rttyskimserv2)
    {
        if (Test-Path $iniFilePath4)
        {
            $newContent4 = (Get-Content $iniFilePath4 -Raw) -replace $replacementPattern, "`${1}$skimSrvCalibration"

            if (-not $DryRun) 
            {
                # Replace calibration factor with new value
                $newContent4 | Set-Content $iniFilePath4
                Write-Host "Successfully updated $iniFile4 with new calibration factor: $skimSrvCalibration"
            } 
            else 
            {
                Write-Host "Did not update $iniFile4 with new calibration factor: $skimSrvCalibration"
            }
        }
        else 
        {
            Write-Host "$iniFile4 not found. Exiting."
            exit 1
        }
    }

    if ($cwsldigi) 
    {
        if (Test-Path $configFilePath )
        {
            $newContent3 = (Get-Content $configFilePath -Raw) -replace $replacementPattern, "`${1}$cwsldigiCalibration"

            if (-not $DryRun) 
            {
                # Replace calibration factor with new value
                $newContent3 | Set-Content $configFilePath
                Write-Host "Successfully updated $configFile with new calibration factor: $cwsldigiCalibration"
            } 
            else 
            {
                Write-Host "Did not update $configFile with new calibration factor: $cwsldigiCalibration"
            }
        }
        else 
        {
            Write-Host "$configFile not found. Exiting."
            exit 1
        }
    }
    
    # Start applications again
    if ($skimsrv1)
    {
        if ($Verbose) { Write-Host "Starting $skimsrvExe1..." }
        Start-Process -WorkingDirectory $skimsrvPath1 -FilePath $skimsrvExe1 -WindowStyle Minimized
    }

    if ($skimsrv2)
    {
        # Wait a moment to let UDP stream stabilize
        Start-Sleep -Seconds 3
        if ($Verbose) { Write-Host "Starting $skimsrvExe2..." }
        Start-Process -WorkingDirectory $skimsrvPath2 -FilePath $skimsrvExe2 -WindowStyle Minimized
    }

    if ($rttyskimserv1)
    {
        # Wait a moment to let UDP stream stabilize
        Start-Sleep -Seconds 3
        if ($Verbose) { Write-Host "Starting $skimsrvExe3..." }
        Start-Process -WorkingDirectory $skimsrvPath3 -FilePath $skimsrvExe3 -WindowStyle Minimized
    }

    if ($rttyskimserv2)
    {
        # Wait a moment to let UDP stream stabilize
        Start-Sleep -Seconds 3
        if ($Verbose) { Write-Host "Starting $skimsrvExe4..." }
        Start-Process -WorkingDirectory $skimsrvPath4 -FilePath $skimsrvExe4 -WindowStyle Minimized
    }

    if ($cwsldigi)
    {
        if ($Verbose) { Write-Host "Starting CWSL_DIGI..." }
        Start-Process -WorkingDirectory $cwslPath -FilePath $cwslExe -WindowStyle Minimized
    }

    if ($Verbose) { Write-Host "Applications restarted." }
    Write-Host "Update complete at $(Get-Date -Format "HH:mm:ss")"
}
catch 
{
    Write-Error "An error occurred: $($_.Exception.Message)"
    exit 1
}
