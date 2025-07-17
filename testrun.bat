@echo off
rem Batch file to run the calibration script daily and append output to a log file
rem Preferrably called from Windows Scheduler
powerShell -File calibrate.ps1 -dryrun
powerShell -File restartthings.ps1
exit /b 0
