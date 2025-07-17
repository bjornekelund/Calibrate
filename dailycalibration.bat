@echo off
rem Batch file to run the calibration script daily and append output to a log file
rem Preferrably called from Windows Scheduler
rem Change the cd command to the directory where your calibrate.ps1 script is located
cd C:\Calibrate
powerShell -File calibrate.ps1 >> "calibrationlog.txt"
powerShell -File restartthings.ps1 >> "calibrationlog.txt"
exit /b 0
