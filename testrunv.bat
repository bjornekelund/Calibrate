@echo off
rem Batch file to run the calibration script daily and append output to a log file
rem Preferrably called from Windows Scheduler
powerShell -File calibrate.ps1 -dryrun -verbose
powerShell -File restartthings.ps1 -verbose
exit /b 0
