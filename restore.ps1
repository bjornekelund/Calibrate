# PowerShell script to back up frequency calibration of CWSL_DIGI and two instances of SkimSrv

# Configuration for this specific installation

$skimsrv1 = $true  # Set to $true if you have SkimSrv installed
$skimsrv2 = $true  # Set to $true if you have two instances of SkimSrv installed
$rttyskimserv1 = $false  # Set to $true if one instance of RttySkimServ is installed
$rttyskimserv2 = $false  # Set to $true if one instance of RttySkimServ is installed
$cwsldigi = $true  # Set to $true if you are using CWSL_DIGI

$iniPath1 = $env:APPDATA + "\Afreet\Products\SkimSrv\"
$iniPath2 = $env:APPDATA + "\Afreet\Products\SkimSrv2\"
$iniPath3 = $env:APPDATA + "\Afreet\Products\RttySkimServ1\"
$iniPath3 = $env:APPDATA + "\Afreet\Products\RttySkimServ2\"
$configPath = "C:\CWSL_DIGI\"

$iniFile1 = "SkimSrv.ini"
$iniFile2 = "SkimSrv2.ini"
$iniFile3 = "RttySkimServ1.ini"
$iniFile4 = "RttySkimServ2.ini"
$configFile = "config.ini"

$iniBackup1 = "SkimSrv_backup.ini"
$iniBackup2 = "SkimSrv2_backup.ini"
$iniBackup3 = "RttySkimServ1_backup.ini"
$iniBackup4 = "RttySkimServ2_backup.ini"
$configBackup = "config_backup.ini"

# End of configuration

$iniFilePath1 = $iniPath1 + $iniFile1
$iniFilePath2 = $iniPath2 + $iniFile2
$iniFilePath3 = $iniPath3 + $iniFile3
$iniFilePath4 = $iniPath4 + $iniFile4
$configFilePath = $configPath + $configFile

$iniBackupPath1 = $iniPath1 + $iniBackup1
$iniBackupPath2 = $iniPath2 + $iniBackup2
$iniBackupPath3 = $iniPath3 + $iniBackup3
$iniBackupPath4 = $iniPath4 + $iniBackup4
$configBackupPath = $configPath + $configBackup

try 
{
    if ($skimsrv1) 
    {
        Copy-Item $iniBackupPath1 $iniFilePath1
        write-Host "Backup of $iniFile1 restored"
    } 

    if ($skimsrv2) 
    {
        Copy-Item $iniBackupPath2 $iniFilePath2
        write-Host "Backup of $iniFile2 restored"
    }   

    if ($rttyskimserv1) 
    {
        Copy-Item $iniBackupPath3 $iniFilePath3
        write-Host "Backup of $iniFile3 restored"
    }

    if ($rttyskimserv2) 
    {
        Copy-Item $iniBackupPath4 $iniFilePath4
        write-Host "Backup of $iniFile4 restored"
    }

    if ($cwsldigi) 
    {
        Copy-Item $configBackupPath $configFilePath
        write-Host "Backup of $configFile restored"
    }
}
catch 
{
    Write-Error "An error occurred: $($_.Exception.Message)"
    exit 1
}
