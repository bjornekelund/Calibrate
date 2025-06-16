# PowerShell script to back up frequency calibration of CWSL_DIGI and two instances of SkimSrv

# Configuration for this specific installation

$skimsrv1 = $true  # Set to $true if you have SkimSrv installed
$skimsrv2 = $true  # Set to $true if you have two instances of SkimSrv installed
$rttyskimserv = $false  # Set to $true if one instance of RttySkimServ is installed
$cwsldigi = $true  # Set to $true if you are using CWSL_DIGI

$iniPath1 = $env:APPDATA + "\Afreet\Products\SkimSrv\"
$iniPath2 = $env:APPDATA + "\Afreet\Products\SkimSrv2\"
$iniPath3 = $env:APPDATA + "\Afreet\Products\RttySkimServ\"
$configPath = "C:\CWSL_DIGI\"

$iniFile1 = "SkimSrv.ini"
$iniFile2 = "SkimSrv2.ini"
$iniFile3 = "RttySkimServ.ini"
$configFile = "config.ini"

$iniBackup1 = "SkimSrv_backup.ini"
$iniBackup2 = "SkimSrv2_backup.ini"
$iniBackup3 = "RttySkimServ_backup.ini"
$configBackup = "config_backup.ini"

# End of configuration

$iniFilePath1 = $iniPath1 + $iniFile1
$iniFilePath2 = $iniPath2 + $iniFile2
$iniFilePath3 = $iniPath3 + $iniFile3
$configFilePath = $configPath + $configFile

$iniBackupPath1 = $iniPath1 + $iniBackup1
$iniBackupPath2 = $iniPath2 + $iniBackup2
$iniBackupPath3 = $iniPath3 + $iniBackup3
$configBackupPath = $configPath + $configBackup

try 
{
    if ($skimsrv1) 
    {
        Copy-Item $iniFilePath1 $iniBackupPath1
        write-Host "Backup of $iniFile1 created at $iniBackup1"
    } 

    if ($skimsrv2) 
    {
        Copy-Item $iniFilePath2 $iniBackupPath2
        write-Host "Backup of $iniFile2 created at $iniBackup2"
    }   

    if ($rttyskimserv) 
    {
        Copy-Item $iniFilePath3 $iniBackupPath3
        write-Host "Backup of $iniFile3 created at $iniBackup3"
    }

    if ($cwsldigi) 
    {
        Copy-Item $configFilePath $configBackupPath
        write-Host "Backup of $configFile created at $configBackup"
    }
}
catch 
{
    Write-Error "An error occurred: $($_.Exception.Message)"
    exit 1
}
