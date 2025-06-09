@echo off
rem Batch file to run the calibration script daily and append output to a log file
rem Preferrably called from Windows Scheduler
cd C:\Calibrate
powerShell -File calibrate.ps1 >> "calibrationlog.txt"
exit /b 0
